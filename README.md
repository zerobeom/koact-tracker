# KoAct ETF 구성종목 변동 추적기

삼성액티브자산운용 **투자종목정보(PDF) 엑셀**을 매 영업일 자동으로 받아와,
**전일 대비 편입·편출·추가매수·일부매도**를 웹페이지로 보여줍니다. 비중은 주가 등락만으로도
변하지만 **수량(주식 수)은 매니저가 실제로 매매해야만 바뀌므로**, 보유수량과 변동수량(+/−주)을
중심에 두고 비중은 보조 지표로 함께 표시합니다. 운용사 팩트시트(상위 종목만)와 달리 이 엑셀은
**전 종목**을 담고 있어, 전체 구성 변화를 빠짐없이 추적합니다.
한 화면에서 **여러 ETF를 골라** 볼 수 있고, 한 번 세팅하면 손대지 않아도 매일 갱신됩니다.

현재 추적 대상:
- **KoAct 미국나스닥성장기업액티브** (`0015B0`, 펀드ID `2ETFQ1`)
- **KoAct 코리아밸류업액티브** (`495230`, 펀드ID `2ETFP3`)

```
koact-tracker/
├─ index.html                  # 보여주는 사이트 (GitHub Pages)
├─ scripts/fetch_holdings.py   # 엑셀 다운로드 + 파싱 + 변동 계산 (모든 ETF)
├─ .github/workflows/update.yml# 매 영업일 자동 실행
├─ requirements.txt
└─ data/                       # 자동 생성/갱신되는 데이터(JSON)
   ├─ etfs.json                # 사이트가 읽는 ETF 목록
   ├─ us-nasdaq/               # ETF별 폴더 (slug 기준)
   │  ├─ latest.json
   │  ├─ dates.json
   │  └─ snapshots/YYYY-MM-DD.json   # 2026-06-26 한 건이 미리 들어있음(예시)
   └─ kr-valueup/
      └─ snapshots/                  # 첫 Action 실행 시 채워짐
```

**데이터 소스**
`https://www.samsungactive.co.kr/excel_pdf.do?fId={펀드ID}&gijunYMD=YYYYMMDD`
→ 구형 `.xls`(BIFF) 파일. 컬럼: 번호 · 종목명 · **ISIN** · 종목코드 · 수량 · 비중(%) · 평가금액 …
종목 식별은 **ISIN**을 키로 사용합니다(이름이 조금 바뀌어도 안전).

---

## 1. 저장소 만들기

GitHub에서 새 저장소 생성(예: `koact-tracker`, **Public** 권장) 후 이 폴더를 그대로 push:

```bash
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/<아이디>/koact-tracker.git
git push -u origin main
```

## 2. 자동 실행(Actions) 켜기

1. **Settings → Actions → General → Workflow permissions** 에서 **Read and write permissions** 선택 후 저장.
2. **Actions** 탭 → *Update KoAct holdings* → **Run workflow** 를 한 번 눌러 즉시 실행.
   → `data/`에 그날 구성종목이 추가됩니다.
3. 이후 매 영업일 18:00(KST) 자동 실행. (시간은 `update.yml`의 `cron`에서 변경)

## 3. 사이트 공개(Pages)

1. **Settings → Pages**.
2. **Source: Deploy from a branch**, **Branch: `main` / `(root)`** 선택 후 저장.
3. 1~2분 뒤 `https://<아이디>.github.io/koact-tracker/` 에서 열립니다.

> 예시로 `2026-06-26` 데이터가 이미 들어있어, Pages만 켜도 바로 실제 구성종목이 보입니다.
> 변동(편입/편출)은 둘째 날 데이터가 쌓이는 순간부터 표시됩니다.

---

## 로컬에서 직접 돌려보기

```bash
pip install -r requirements.txt
python scripts/fetch_holdings.py            # 오늘(KST) 기준, 없으면 직전 영업일로 자동 후퇴
python scripts/fetch_holdings.py 20260626   # 특정일 강제 수집(과거 채우기)
python scripts/fetch_holdings.py --debug    # 파싱 전 원본 표 확인
python -m http.server 8000   →   http://localhost:8000   # 사이트 미리보기
```

상장 이후 전체 이력을 원하면 과거 날짜를 하나씩 채우면 됩니다(주말/휴장일은 자동으로 건너뜀):

```bash
for d in 20260623 20260624 20260625 20260626; do python scripts/fetch_holdings.py $d; done
```

## 커스터마이즈

| 바꾸고 싶은 것 | 위치 |
|---|---|
| 추적 ETF 추가/변경 | `scripts/fetch_holdings.py` 상단의 `ETFS` 목록에 `{slug, fid, ticker, name}` 한 줄 추가 |
| 비중 증감 민감도 | 같은 파일의 `WEIGHT_EPS` (기본 0.05%p) |
| 파일 없을 때 후퇴 일수 | 같은 파일의 `LOOKBACK_DAYS` (기본 7일) |
| 자동 실행 시각 | `.github/workflows/update.yml`의 `cron` (UTC 기준, `0 9`=18시 KST) |
| 색/디자인 | `index.html` 상단 `:root` CSS 변수 |

> 다른 KoAct ETF의 `fid`는 해당 상품 페이지 주소 `etf/view.do?id=____` 의 값과 같습니다.
> `slug`은 폴더명으로 쓰이는 영문 식별자라 ETF마다 겹치지 않게 정하면 됩니다(예: `kr-dividend`).

---

## 알아둘 점

- 이 엑셀은 **설정현금액·원화현금** 행을 포함합니다. 스크립트가 이를 `현금`으로 표시하고
  편입/편출/비중 집계에서 제외하며, 사이트에서는 표 하단에 흐리게 보여줍니다.
- 비중 값이 소수(예: 0.07)로 들어오든 `7.00%`로 들어오든 자동으로 백분율(합 100%)로 정규화합니다.
- `.xls`(구형) 파싱에 `xlrd`가 필요합니다(이미 `requirements.txt`에 포함). 응답이 드물게 HTML 표로
  올 경우 `lxml`로 자동 폴백합니다.
- 휴장일·미공시일에는 자동으로 직전 영업일 파일까지 후퇴해 찾고, 그래도 없으면 건너뜁니다.
- 정보 제공용이며 투자 권유가 아닙니다.

데이터 출처: 삼성액티브자산운용.
