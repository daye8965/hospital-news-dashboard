"""
news.csv 오염 데이터 정리 스크립트
"""
import csv, re, sys
from pathlib import Path

CSV_PATH = Path("docs/news.csv")

REMOVE_KEYWORDS = {"삼성병원", "연대병원", "서울중앙병원"}

REMOVE_PATTERN = re.compile(
    r"(삼성전자|삼성증권|삼성물산|삼성SDI|삼성SDS|삼성화재|삼성생명|삼성바이오로직스"
    r"|北축구|노동연대|성과급\s*논란|주가\s*급등|영업이익|매출액\s*발표"
    r"|부고|부음|별세|타계|빈소|발인)"
)

def title_key(t):
    return re.sub(r"[^\w가-힣]", "", t).lower()

# CSV 파일 없으면 조용히 종료
if not CSV_PATH.exists():
    print(f"파일 없음: {CSV_PATH} — 건너뜀")
    sys.exit(0)

# 파일 읽기
with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
    content = f.read().strip()

# 빈 파일이면 종료
if not content:
    print("CSV 파일이 비어있음 — 건너뜀")
    sys.exit(0)

rows, removed = [], 0
seen_titles = set()
fieldnames = None

with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f)
    fieldnames = reader.fieldnames

    # 헤더가 없으면 종료
    if not fieldnames:
        print("헤더 없음 — 건너뜀")
        sys.exit(0)

    for row in reader:
        keyword = row.get("검색어", "")
        title   = row.get("제목", "")
        summary = row.get("요약", "")
        tk = title_key(title)

        if keyword in REMOVE_KEYWORDS:
            removed += 1; continue
        if REMOVE_PATTERN.search(title + summary):
            removed += 1; continue
        if tk and tk in seen_titles:
            removed += 1; continue

        if tk: seen_titles.add(tk)

        # 구버전 행 보정: 발행일시 없으면 날짜 컬럼으로 채우기
        if not row.get("발행일시","").strip():
            row["발행일시"] = row.get("날짜","")
        # 병원그룹 없으면 검색어로 유추
        if not row.get("병원그룹","").strip():
            kw = row.get("검색어","")
            if kw in ("서울아산병원","아산병원","울산의대"):
                row["병원그룹"] = "아산"
            elif kw in ("삼성서울병원","성대의대","성균관의대"):
                row["병원그룹"] = "삼성"
            elif kw in ("세브란스병원","연세암병원","연세대병원"):
                row["병원그룹"] = "세브란스"
            elif kw in ("서울대병원","서울대학교병원","서울의대"):
                row["병원그룹"] = "서울대"
            elif kw in ("서울성모병원","카톨릭성모병원"):
                row["병원그룹"] = "성모"
        # 교수명 컬럼 없으면 빈값 추가
        if "교수명" not in row:
            row["교수명"] = ""

        rows.append(row)

# 신버전 컬럼 순서로 통일해서 저장
NEW_FIELDNAMES = ["날짜","병원그룹","검색어","매체","제목","교수명","요약","언론사원문","네이버링크","발행일시","수집일시"]

with open(CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=NEW_FIELDNAMES, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(rows)

print(f"정리 완료: {removed}건 제거 → 남은 데이터: {len(rows)}건")
