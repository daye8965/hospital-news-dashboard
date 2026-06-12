import os, csv, html, re
import requests
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

# ── API 키 ────────────────────────────────────────────────────────────────────
CLIENT_ID     = os.environ.get("NAVER_CLIENT_ID",     "_eBIxCwm7W7VId2GzO2R")
CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "itD8WVlnrL")

# ── 병원별 검색 키워드 ────────────────────────────────────────────────────────
HOSPITAL_QUERIES = {
    "아산":    ["서울아산병원", "아산병원", "울산의대"],
    "삼성":    ["삼성서울병원", "성대의대", "성균관의대"],
    "세브란스": ["세브란스병원", "연세암병원", "연세대병원"],
    "서울대":  ["서울대병원", "서울대학교병원", "서울의대"],
    "성모":    ["서울성모병원", "카톨릭성모병원"],
}

EXCLUDE_KEYWORDS = {
    "부고", "부음", "별세", "타계", "빈소", "발인",
    "삼성전자", "삼성증권", "삼성물산", "삼성SDI", "삼성SDS",
    "삼성화재", "삼성생명", "삼성바이오",
    "서울중앙병원", "삼성병원", "연대병원",
}

CSV_PATH   = Path("docs/news.csv")
FIELDNAMES = ["날짜","병원그룹","검색어","매체","제목","교수명","요약","언론사원문","네이버링크","발행일시","수집일시"]

KST         = timezone(timedelta(hours=9))
today       = datetime.now(KST).date()
target_date = today - timedelta(days=1)
start_dt    = datetime(target_date.year, target_date.month, target_date.day, tzinfo=KST)
end_dt      = datetime(today.year,       today.month,       today.day,       tzinfo=KST)

# ── 매체 매핑 (도메인 → 언론사명) ─────────────────────────────────────────────
MEDIA_MAP = {
    # 종합일간지
    "chosun.com":"조선일보", "joongang.co.kr":"중앙일보", "joongang.joins.com":"중앙일보",
    "donga.com":"동아일보", "hani.co.kr":"한겨레", "khan.co.kr":"경향신문",
    "hankookilbo.com":"한국일보", "seoul.co.kr":"서울신문", "kmib.co.kr":"국민일보",
    "segye.com":"세계일보", "munhwa.com":"문화일보",
    # 경제지
    "hankyung.com":"한국경제", "mk.co.kr":"매일경제", "sedaily.com":"서울경제",
    "fnnews.com":"파이낸셜뉴스", "mt.co.kr":"머니투데이", "heraldcorp.com":"헤럴드경제",
    "asiae.co.kr":"아시아경제", "edaily.co.kr":"이데일리", "etnews.com":"전자신문",
    "bizwatch.co.kr":"비즈워치", "thebell.co.kr":"더벨",
    # 방송
    "kbs.co.kr":"KBS", "news.kbs.co.kr":"KBS", "imbc.com":"MBC", "imnews.imbc.com":"MBC",
    "news.sbs.co.kr":"SBS", "jtbc.co.kr":"JTBC", "ytn.co.kr":"YTN",
    "mbn.co.kr":"MBN", "tvchosun.com":"TV조선", "ichannela.com":"채널A",
    "ebs.co.kr":"EBS", "yonhapnewstv.co.kr":"연합뉴스TV",
    # 통신
    "yna.co.kr":"연합뉴스", "newsis.com":"뉴시스", "news1.kr":"뉴스1",
    "newspim.com":"뉴스핌",
    # 의료/건강 전문
    "healthchosun.com":"헬스조선", "kormedi.com":"코메디닷컴",
    "medicaltimes.com":"메디칼타임즈", "doctorsnews.co.kr":"의사신문",
    "rapportian.com":"라포르시안", "mdtoday.co.kr":"메디컬투데이",
    "newsmp.com":"메디컬포스트", "dailymedi.com":"데일리메디",
    "bosa.co.kr":"보건신문", "yakup.com":"약업신문", "kpanews.co.kr":"약사공론",
    "pharmnews.com":"팜뉴스", "medigatenews.com":"메디게이트뉴스",
    "청년의사":"청년의사", "docdocdoc.co.kr":"청년의사",
    "mdjournal.kr":"메디컬저널", "medipana.com":"메디파나뉴스",
    "hitnews.co.kr":"히트뉴스",
    # 일반 온라인
    "ohmynews.com":"오마이뉴스", "pressian.com":"프레시안",
    "mediatoday.co.kr":"미디어오늘", "sisajournal.com":"시사저널",
    "weekly.chosun.com":"주간조선", "news.joins.com":"중앙일보",
    "koreabiomed.com":"코리아바이오메드",
    # 지역지
    "imaeil.com":"매일신문", "kookje.co.kr":"국제신문",
    "busan.com":"부산일보", "ulsanpress.com":"울산매일",
    "joongdo.co.kr":"중도일보",
}

# 네이버 뉴스 링크일 때 originallink 없는 경우 처리용 도메인 블랙리스트
NAVER_DOMAINS = {"news.naver.com", "n.news.naver.com", "m.news.naver.com"}


def clean_text(t):
    return html.unescape(re.sub(r"<.*?>", "", t)).strip()

def extract_media(orig_url, naver_url=""):
    """언론사 원문 URL 우선, 없으면 빈 문자열 (네이버 도메인은 무의미)"""
    for url in [orig_url, naver_url]:
        if not url:
            continue
        try:
            domain = re.sub(r"https?://(www\.)?", "", url).split("/")[0].lower()
            if domain in NAVER_DOMAINS:
                continue
            for key, name in MEDIA_MAP.items():
                if key in domain:
                    return name
            # 매핑 없으면 도메인 두 번째 파트 (co.kr 앞)
            parts = domain.split(".")
            if len(parts) >= 2:
                # example.co.kr → example / example.com → example
                if parts[-1] in ("kr","jp","cn","uk") and len(parts) >= 3:
                    return parts[-3]
                return parts[-2]
            return domain
        except Exception:
            continue
    return ""

def extract_professor(title, summary):
    """제목+요약에서 아산병원 교수명 추출"""
    combined = title + " " + summary
    # 패턴: 이름(2~4자) + 교수 / 교수 + 이름
    patterns = [
        r"([가-힣]{2,4})\s*(?:서울아산병원|아산병원)?\s*(?:\w+과?\s*)?교수",
        r"(?:서울아산병원|아산병원)\s*(?:\w+과?\s*)?([가-힣]{2,4})\s*교수",
        r"교수\s+([가-힣]{2,4})(?:\s|,|·|$)",
    ]
    found = []
    for pat in patterns:
        for m in re.finditer(pat, combined):
            name = m.group(1).strip()
            if 2 <= len(name) <= 4 and name not in found:
                found.append(name)
    return ", ".join(found[:3])  # 최대 3명

def is_excluded(item):
    combined = clean_text(item.get("title","")) + clean_text(item.get("description",""))
    return any(kw in combined for kw in EXCLUDE_KEYWORDS)

def title_key(title):
    """제목 정규화: 특수문자 모두 제거 후 앞 30자로 중복 감지"""
    t = re.sub(r"[^\w가-힣a-zA-Z0-9]", "", title)
    return t.lower()[:30]

def collect_news(query):
    collected, now = [], datetime.now(KST)
    for start in range(1, 1001, 100):
        try:
            resp = requests.get(
                "https://openapi.naver.com/v1/search/news.json",
                headers={"X-Naver-Client-Id":CLIENT_ID,"X-Naver-Client-Secret":CLIENT_SECRET},
                params={"query":query,"display":100,"start":start,"sort":"date"},
                timeout=10,
            )
            resp.raise_for_status()
        except Exception as e:
            print(f"  API 오류: {e}"); break

        items = resp.json().get("items", [])
        if not items: break

        for item in items:
            pub_dt = parsedate_to_datetime(item["pubDate"]).astimezone(KST)
            if start_dt <= pub_dt < end_dt and not is_excluded(item):
                orig  = item.get("originallink","")
                naver = item.get("link","")
                title = clean_text(item.get("title",""))
                summ  = clean_text(item.get("description",""))
                collected.append({
                    "날짜":     str(target_date),
                    "검색어":   query,
                    "매체":     extract_media(orig, naver),
                    "제목":     title,
                    "교수명":   extract_professor(title, summ),
                    "요약":     summ,
                    "언론사원문": orig,
                    "네이버링크": naver,
                    "발행일시": pub_dt.strftime("%Y-%m-%d %H:%M"),
                    "수집일시": now.strftime("%Y-%m-%d %H:%M"),
                })

        last_pub = parsedate_to_datetime(items[-1]["pubDate"]).astimezone(KST)
        if last_pub < start_dt: break
    return collected

def collect_all():
    all_items = []
    # 전체 URL 중복 + 제목 유사 중복 모두 잡기
    seen_urls   = set()
    seen_titles = set()

    for group, queries in HOSPITAL_QUERIES.items():
        group_count = 0
        for query in queries:
            news = collect_news(query)
            added = 0
            for item in news:
                url_key = item["언론사원문"] or item["네이버링크"]
                t_key   = title_key(item["제목"])

                # URL 중복 또는 제목 90% 유사 (정규화 후 동일) 시 건너뜀
                if url_key and url_key in seen_urls:
                    continue
                if t_key and t_key in seen_titles:
                    continue

                if url_key: seen_urls.add(url_key)
                if t_key:   seen_titles.add(t_key)

                item["병원그룹"] = group
                all_items.append(item)
                added += 1
            group_count += added
            print(f"  [{group}] {query}: {len(news)}건 → 신규 {added}건")
        print(f"  ▶ [{group}] 합계: {group_count}건\n")
    return all_items

def save_to_csv(items):
    CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    file_exists = CSV_PATH.exists()
    with open(CSV_PATH, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        if not file_exists:
            writer.writeheader()
        writer.writerows(items)
    print(f"CSV 저장 완료: {len(items)}건")

if __name__ == "__main__":
    print(f"수집 기간: {start_dt} ~ {end_dt}\n")
    news = collect_all()
    print(f"\n총 {len(news)}건 (중복 제거 후)")
    if news:
        save_to_csv(news)
