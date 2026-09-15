# hospital-news-dashboard (운영)

서울아산병원 홍보팀 언론Unit의 **병원 뉴스 수집·대시보드 운영 버전**입니다.
네이버 뉴스 API로 주요 병원 기사를 매일 수집해 CSV에 쌓고, GitHub Pages로 대시보드를 서비스합니다.

- 공개 주소: https://daye8965.github.io/hospital-news-dashboard/
- 개발 버전: `../hospital-news-dashboard-dev` (별도 저장소 `hospital-news-dashboard-dev`)

## ⚠️ 이 저장소는 운영본입니다

**여기서 기능을 새로 만들지 마세요.** 신규 기능·실험은 dev 저장소에서 하고, 검증된 것만 이쪽으로 옮깁니다.
push하면 GitHub Pages로 즉시 배포되어 실제 업무에 쓰이는 화면이 바뀝니다.

## 구조

| 파일 | 역할 |
|---|---|
| `naver_to_csv.py` | 네이버 뉴스 API 수집. 병원별 검색어, 제외 키워드/매체, 매체명 정규화, 중복 제거 |
| `clean_csv.py` | 수집 전 CSV 오염 행 정리 |
| `generate_weekly_report.py` | 주간 보고서(xlsx/html) 생성 — 매주 월요일 |
| `docs/index.html` | 대시보드 본체 (단일 파일, CSV를 클라이언트에서 읽어 렌더링) |
| `docs/news.csv` | 수집 데이터. **사람이 직접 편집하지 말 것** — 워크플로가 계속 덮어씁니다 |
| `.github/workflows/naver_news.yml` | 수집 자동화 |

대시보드 탭: `ours`(우리병원) / `others`(타병원) / `compare`(병원 비교)

## 데이터

`docs/news.csv` 헤더 (순서 고정):

```
날짜,병원그룹,검색어,매체,제목,교수명,요약,언론사원문,네이버링크,발행일시,수집일시,기자명,기자이메일,지면,출입기자
```

- 뒤쪽 4개(`기자명`~`출입기자`)는 dev의 보강 기능이 채우는 칸입니다. 운영본에는 보강 스크립트가 없어 비어 있을 수 있습니다.
- 약 17,000행. 파일 전체를 읽지 말고 `head`/`grep`/`awk`로 필요한 부분만 보세요.

## 자동화

- 스케줄: `cron "20 19 * * *"`, `"17 23 * * *"` (UTC) — GitHub 예약 실행이 1~2시간 밀리는 것을 감안해 **앞당겨 설정**해 둔 값입니다. 시간을 조정할 땐 이 지연을 반드시 고려하세요.
- 매주 월요일 `cron "0 0 * * 1"`에는 수집 대신 주간 보고서를 만듭니다.
- `concurrency: news-data` — 같은 CSV를 건드리는 작업이 동시에 돌지 않게 묶여 있습니다.
- 커밋 충돌 시 **이번 수집 결과를 버리고 정상 종료**합니다. 빠진 날짜는 다음 실행의 백필이 다시 수집합니다 (`MAX_BACKFILL_DAYS = 7`).
- 자동 커밋 메시지: `자동 수집 YYYY-MM-DD` (작성자 `github-actions[bot]`)

## Secrets

`NAVER_CLIENT_ID`, `NAVER_CLIENT_SECRET`

## 작업 규칙

- 커밋 메시지는 기존 히스토리를 따릅니다 — 한글 서술형 또는 영문 명령형 한 줄.
- 워크플로가 만든 `자동 수집 …` 커밋과 섞이므로, push 전에 항상 `git pull --rebase`로 최신을 받으세요.
- 로컬에서 수집 스크립트를 돌리려면 네이버 API 키가 필요합니다. 키 없이 검증할 땐 기존 CSV로 대시보드(`docs/index.html`)만 확인하세요.
- 정적 페이지 확인: `python -m http.server 8765` 후 `http://localhost:8765/docs/`
