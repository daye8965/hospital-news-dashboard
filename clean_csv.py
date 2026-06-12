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
        rows.append(row)

# 덮어쓰기
with open(CSV_PATH, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(rows)

print(f"정리 완료: {removed}건 제거 → 남은 데이터: {len(rows)}건")
