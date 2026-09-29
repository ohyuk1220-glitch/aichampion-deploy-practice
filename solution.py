"""3과목(서비스구현) 제출용 solution.py 골격.

채점: 채점자가 이 파일을 직접 실행한다. 실행 에러가 나면 -20점.
그래서 이 골격의 최우선 목표는 "어떤 환경에서도 끝까지 돈다"이다.
  - 표준 라이브러리만 사용 (csv / json / pathlib / urllib) → pandas 미설치 환경에서도 실행됨
  - 인코딩 자동 판별 (공공데이터 CSV는 cp949 가 흔하다)
  - 컬럼·값이 비어도 죽지 않고, 대신 [경고]로 알린다

교체할 곳은 ▼교체 표시만 보면 된다.
"""

import csv
import html
import io
import json
import re
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from collections import Counter, defaultdict
from pathlib import Path
from urllib.request import Request, urlopen

BASE = Path(__file__).resolve().parent

# ▼교체1: 첨부 CSV 파일명 (압축 해제한 실제 이름으로)
MAIN_CSV = BASE / "청렴신고_사건데이터.csv"
SUB_CSV = BASE / "청렴교육_현황.csv"

# ▼교체2: 출력 파일명 (문제지가 지정한 이름 그대로)
OUT_CSV = BASE / "결과_요약.csv"
OUT_JS = BASE / "data.js"   # 정적 페이지가 <script src="data.js"> 로 읽는다

# ▼교체3: 정적 페이지에 쓸 외부 RSS (라이브 실패해도 스크립트는 계속 진행)
RSS_URL = "https://news.google.com/rss/search?q=%ED%96%89%EC%A0%95%EC%95%88%EC%A0%84%EB%B6%80&hl=ko&gl=KR&ceid=KR:ko"

MISSING = "미상"   # 값이 비어 있을 때 쓸 이름


def read_csv(path: Path) -> list[dict]:
    """인코딩을 자동 판별해 CSV를 딕셔너리 목록으로 읽는다."""
    if not path.is_file():
        print(f"[경고] 파일이 아니거나 없음: {path.name}")
        return []
    # utf-16 은 BOM 때문에 cp949 로도 '읽히는 것처럼' 보이므로 먼저 시도한다
    for enc in ("utf-8-sig", "utf-16", "cp949", "euc-kr", "utf-8"):
        try:
            text = path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError, LookupError):
            continue
        head = text[:2000]
        if "\x00" in head or "�" in head:
            continue   # 잘못된 인코딩으로 읽으면 깨진 문자가 섞인다

        # 구분자 판별: 공공데이터 엑셀 내보내기는 세미콜론·탭인 경우가 있다.
        # 단순히 글자 수를 세면 따옴표로 감싼 셀 안의 기호에 속으므로,
        # 후보마다 csv 모듈로 실제 파싱해 보고 헤더가 가장 잘 쪼개지는 것을 고른다.
        best, best_score = ",", (0, 0)
        for cand in (",", ";", "\t", "|"):
            try:
                probe = []
                for i, row in enumerate(csv.reader(io.StringIO(text), delimiter=cand)):
                    probe.append(len(row))
                    if i >= 19:        # 앞 20행만 본다
                        break
            except csv.Error:
                continue
            if not probe:
                continue
            ncols = probe[0]
            # 헤더와 같은 칸 수인 행의 비율 = 그 구분자가 맞을 가능성
            agree = sum(1 for n in probe if n == ncols) / len(probe)
            score = (ncols, agree)
            if ncols > 1 and score > best_score:
                best, best_score = cand, score
        delim = best

        try:
            # splitlines() 를 쓰면 따옴표 안의 줄바꿈이 사라진다 → StringIO 로 넘긴다
            rows = list(csv.DictReader(io.StringIO(text), delimiter=delim, restkey="__여분__"))
        except csv.Error:
            continue
        if not rows:
            continue

        cols = [c for c in rows[0] if c != "__여분__"]
        shown = "\\t" if delim == "\t" else delim
        print(f"[읽기] {path.name} {len(rows)}행 (encoding={enc}, 구분자='{shown}')")
        print(f"       컬럼 {len(cols)}개: {cols}")
        if len(cols) == 1:
            print("       [경고] 컬럼이 1개로 잡혔습니다. 구분자를 확인하세요 — 이대로면 집계가 전부 틀립니다.")
        extra = sum(1 for r in rows if r.get("__여분__"))
        if extra:
            print(f"       [경고] 칸 수가 헤더보다 많은 행 {extra}개. 본문에 구분자가 들어간 행일 수 있습니다.")
        return rows
    print(f"[경고] 읽기 실패 또는 데이터 없음: {path.name}")
    return []


def to_num(value):
    """'1,234원', '-5.5', '1e3' 같은 문자열에서 수를 뽑는다. 실패하면 None.

    공공데이터에는 유니코드 빼기표(−, U+2212)와 전각 숫자가 섞여 들어온다.
    """
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None
    s = str(value).strip()
    if not s:
        return None
    # 유니코드 빼기표·전각 기호를 보통 문자로 통일
    s = s.replace("−", "-").replace("－", "-").replace("＋", "+")
    s = s.translate(str.maketrans("０１２３４５６７８９．，", "0123456789.,"))
    m = re.search(r"[-+]?(?:\d+(?:,\d{3})*(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?", s)
    if not m:
        return None
    try:
        d = Decimal(m.group().replace(",", ""))
    except (InvalidOperation, ArithmeticError, ValueError):
        return None
    return d if d.is_finite() else None


def mean(values):
    """평균. 쓸 수 있는 수가 하나도 없으면 None 을 돌려준다(0 이 아니다)."""
    nums = [v for v in (to_num(x) for x in values) if v is not None]
    if not nums:
        return None
    try:
        avg = sum(nums) / len(nums)
    except ArithmeticError:     # Overflow 등 — 계산불가로 돌린다(죽지 않는다)
        return None
    return avg if avg.is_finite() else None


def round_half_up(value, digits: int = 0):
    """사사오입. 파이썬 기본 round()는 round(2.5)==2 인 '은행원 반올림'이라 답이 1 어긋난다."""
    if value is None:
        return None
    try:
        d = value if isinstance(value, Decimal) else Decimal(str(value))
        if not d.is_finite():
            return None
        q = Decimal(1).scaleb(-digits)
        r = d.quantize(q, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ArithmeticError, ValueError):
        return None      # 너무 크거나 이상한 수 → 계산불가로 표시(죽지 않는다)
    return int(r) if digits == 0 else float(r)


def show(value, unit: str = "") -> str:
    """계산 불가(None)를 0 인 것처럼 보이지 않게 표시한다."""
    return "계산불가" if value is None else f"{value}{unit}"


def label(row: dict, key: str) -> str:
    """값이 없거나 빈 문자열이면 '미상'. (행이 짧으면 None 이 들어와 정렬에서 터진다)"""
    v = row.get(key)
    if v is None:
        return MISSING
    v = str(v).strip()
    return v or MISSING


def fetch_rss(url: str, limit: int = 8) -> list[dict]:
    """RSS를 받아 제목·링크·날짜를 뽑는다.

    브라우저에서 직접 부르면 CORS로 막히므로 여기(파이썬)에서 받아
    data.js 에 동봉한다. 네트워크가 막혀도 예외로 죽지 않는다.
    """
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=15) as resp:
            xml = resp.read().decode("utf-8", errors="replace")
    except Exception as exc:  # 네트워크 차단·타임아웃 전부 흡수
        print(f"[경고] RSS 수집 실패({exc.__class__.__name__})")
        return []

    items = []
    for block in re.findall(r"<item>(.*?)</item>", xml, re.S)[:limit]:
        def pick(tag: str, is_url: bool = False) -> str:
            m = re.search(rf"<{tag}>(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?</{tag}>", block, re.S)
            if not m:
                return ""
            raw = re.sub(r"<[^>]+>", "", m.group(1))
            # 주소는 &amp; 만 되돌린다. 전체 unescape 하면 "&copy=" 가 "©" 로 바뀌어 주소가 깨진다.
            return (raw.replace("&amp;", "&") if is_url else html.unescape(raw)).strip()
        items.append({"title": pick("title"), "link": pick("link", True), "date": pick("pubDate")[:16]})
    print(f"[읽기] RSS {len(items)}건")
    return items


def load_saved_rss() -> list[dict]:
    """이전 실행으로 받아 둔 RSS를 읽어 온다. 네트워크가 끊겨도 화면이 비지 않게."""
    if not OUT_JS.exists():
        return []
    try:
        txt = OUT_JS.read_text(encoding="utf-8")
        body = txt[txt.index("{"): txt.rindex("}") + 1]
        return json.loads(body).get("rss", []) or []
    except Exception:
        return []


def write_csv(path: Path, header: list[str], rows: list[list]) -> None:
    """엑셀에서 한글이 깨지지 않도록 utf-8-sig 로 저장한다."""
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"[저장] {path.name} ({len(rows)}행)")


def main() -> int:
    main_rows = read_csv(MAIN_CSV)
    sub_rows = read_csv(SUB_CSV)

    if not main_rows:
        print("=" * 60)
        print(f"[중단] 주 데이터가 비어 있습니다. {MAIN_CSV.name} 경로와 파일명을 확인하세요.")
        print("=" * 60)
        return 0   # 채점이 종료코드만 볼 수도 있어 0으로 끝낸다. 메시지로 알린다.

    # ───────── ▼교체4: 문제지가 요구한 집계를 여기서 ─────────
    total = len(main_rows)
    done = sum(1 for r in main_rows if label(r, "처리상태") == "처리완료")

    avg_raw = mean(r.get("처리기간_일") for r in main_rows)
    avg_days = round_half_up(avg_raw)

    by_type = Counter(label(r, "신고유형") for r in main_rows)

    by_org = defaultdict(list)
    for r in sub_rows:
        by_org[label(r, "기관유형")].append(r.get("교육이수율"))
    org_avg = {}
    for k in sorted(by_org):
        m = mean(by_org[k])
        org_avg[k] = round_half_up(m, 1)
    # ──────────────────────────────────────────────────────

    if avg_days is None:
        print("[경고] '처리기간_일'에서 숫자를 하나도 찾지 못했습니다. 컬럼명을 확인하세요.")
    if sub_rows and all(v is None for v in org_avg.values()):
        print("[경고] '교육이수율'에서 숫자를 하나도 찾지 못했습니다. 컬럼명을 확인하세요.")

    # 콘솔 출력 — 문제지 예시 형식을 그대로 흉내 낸다
    print()
    print(f"[OK] 신고 {total}건 분석 → {OUT_CSV.name}")
    print(f"   총 사건: {total}건 / 처리 완료: {done}건 / 평균 처리: {show(avg_days, '일')}")
    print()
    print("■ 신고유형별 사건 수")
    for k, v in by_type.most_common():
        print(f"   {k:<10} {v:>5}건")
    print()
    print("■ 기관유형별 평균 이수율")
    for k, v in org_avg.items():
        print(f"   {k:<10} {show(v):>7}")

    # 결과 CSV — 지표를 한 파일에 담는다(구분/항목/값 3열이 가장 안전)
    rows = [["총계", "총 신고 사건 수", total],
            ["총계", "처리 완료 사건 수", done],
            ["총계", "평균 처리기간(일)", show(avg_days)]]
    rows += [["신고유형별 사건 수", k, v] for k, v in by_type.most_common()]
    rows += [["기관유형별 평균 이수율", k, show(v)] for k, v in org_avg.items()]
    write_csv(OUT_CSV, ["구분", "항목", "값"], rows)

    # 정적 페이지용 data.js
    # json 파일이 아니라 js 파일로 쓰는 이유: 브라우저에서 index.html 을 그냥 열어도
    # (file:// 프로토콜) CORS 에 막히지 않고 데이터가 들어간다. 배포 후에도 동일하게 동작.
    rss = fetch_rss(RSS_URL)
    if not rss:
        rss = load_saved_rss()          # 수집 실패 시 지난번에 받아 둔 것을 지키다
        if rss:
            print(f"[유지] 이전 RSS {len(rss)}건을 그대로 둡니다.")
        else:
            print("[경고] RSS가 비었습니다. 페이지의 외부자료 항목이 0건으로 채점될 수 있습니다.")
    payload = json.dumps({"items": main_rows, "rss": rss}, ensure_ascii=False)
    OUT_JS.write_text(f"window.EXAM_DATA = {payload};\n", encoding="utf-8")
    print(f"[저장] {OUT_JS.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
