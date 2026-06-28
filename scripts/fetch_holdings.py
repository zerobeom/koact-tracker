#!/usr/bin/env python3
"""
KoAct ETF 일별 전체 구성종목 수집기 (다중 ETF 지원).

데이터 소스: 삼성액티브자산운용 '투자종목정보(PDF)' 엑셀 다운로드
    https://www.samsungactive.co.kr/excel_pdf.do?fId={펀드ID}&gijunYMD=YYYYMMDD

추적 대상은 아래 ETFS 목록에 추가만 하면 늘어납니다.

산출물(ETF별로 분리):
    data/etfs.json                      : 사이트가 읽는 ETF 목록
    data/{slug}/snapshots/YYYY-MM-DD.json
    data/{slug}/dates.json
    data/{slug}/latest.json

사용:
    python scripts/fetch_holdings.py            # 오늘(KST) 기준, 모든 ETF
    python scripts/fetch_holdings.py 20260626   # 특정일 강제 수집(과거 채우기)
    python scripts/fetch_holdings.py --debug     # 파싱 전 원본 표 출력
"""
from __future__ import annotations

import io
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

# ── 추적할 ETF 목록 ──────────────────────────────────────────────────────
# slug: 폴더/URL용 영문 식별자 / fid: 운용사 펀드ID / ticker: 거래소 단축코드(표시용)
ETFS = [
    {"slug": "us-nasdaq", "fid": "2ETFQ1", "ticker": "0015B0",
     "name": "KoAct 미국나스닥성장기업액티브"},
    {"slug": "kr-valueup", "fid": "2ETFP3", "ticker": "495230",
     "name": "KoAct 코리아밸류업액티브"},
]

URL = "https://www.samsungactive.co.kr/excel_pdf.do"
LOOKBACK_DAYS = 7                       # 해당일 파일이 없으면 며칠 전까지 후퇴 탐색

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
KST = timezone(timedelta(hours=9))


# ── 유틸 ────────────────────────────────────────────────────────────────
def today_kst() -> str:
    return datetime.now(KST).strftime("%Y%m%d")


def to_num(v):
    if v is None:
        return None
    try:
        s = str(v).replace(",", "").replace("%", "").strip()
        if s in ("", "-", "nan", "None"):
            return None
        return float(s)
    except (ValueError, TypeError):
        return None


def clean(v):
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def clean_ticker(code: str) -> str:
    """'MU US Equity' -> 'MU', '005930 KS Equity' -> '005930'. 현금/특수코드는 빈 문자열."""
    code = clean(code)
    if not code or code.startswith(("CASH", "KRD", "KRW")):
        return ""
    return code.split()[0]


# ── 다운로드 + 파싱 ─────────────────────────────────────────────────────
def download(date_yyyymmdd: str, fid: str) -> bytes:
    import requests
    r = requests.get(
        URL,
        params={"fId": fid, "gijunYMD": date_yyyymmdd},
        timeout=30,
        headers={"User-Agent": "Mozilla/5.0",
                 "Referer": "https://www.samsungactive.co.kr/"},
    )
    r.raise_for_status()
    return r.content


def read_table(content: bytes):
    """엑셀(.xls BIFF) 우선, 실패 시 HTML 표로 폴백. header 없이 원본 셀 그대로."""
    import pandas as pd
    try:
        return pd.read_excel(io.BytesIO(content), header=None, dtype=str, engine="xlrd")
    except Exception:
        pass
    try:
        tables = pd.read_html(io.BytesIO(content))
        if tables:
            return max(tables, key=len)
    except Exception:
        pass
    return None


def normalize(raw, debug: bool = False):
    """원본 표(header=None DataFrame) -> (기준일 'YYYY-MM-DD', holdings[list])."""
    if raw is None or len(raw) == 0:
        return None, None
    raw = raw.reset_index(drop=True)
    if debug:
        print(raw.head(6).to_string())

    hdr = None
    for i in range(min(12, len(raw))):
        cells = [str(x).strip() for x in raw.iloc[i].tolist()]
        if "종목명" in cells and any("ISIN" in c or "비중" in c for c in cells):
            hdr = i
            break
    if hdr is None:
        return None, None

    cols = [str(x).strip() for x in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:].reset_index(drop=True)
    body.columns = cols

    base_date = None
    for i in range(hdr):
        for x in raw.iloc[i].tolist():
            m = re.match(r"(\d{4})[/.\-](\d{2})[/.\-](\d{2})", str(x).strip())
            if m:
                base_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                break
        if base_date:
            break

    def col(*names):
        for n in names:
            for c in cols:
                if n in c:
                    return c
        return None

    cN = col("종목명")
    cI = col("ISIN")
    cC = col("종목코드", "코드")
    cQ = col("수량")
    cW = col("비중")
    cA = col("평가금액", "평가")

    holdings = []
    for _, r in body.iterrows():
        name = clean(r.get(cN)) if cN else ""
        if not name or name in ("번호", "종목명"):
            continue
        isin = clean(r.get(cI)) if cI else ""
        code = clean(r.get(cC)) if cC else ""
        is_cash = (isin.startswith(("CASH", "KRD", "KRW"))
                   or "현금" in name or "설정현금" in name)
        holdings.append({
            "isin": isin, "name": name, "code": code,
            "ticker": clean_ticker(code),
            "weight": to_num(r.get(cW)) if cW else None,
            "shares": to_num(r.get(cQ)) if cQ else None,
            "amount": to_num(r.get(cA)) if cA else None,
            "is_cash": is_cash,
            "key": isin or code or name,
        })

    holdings = [h for h in holdings if h["key"]]

    total = sum((h["weight"] or 0) for h in holdings)
    if 0 < total <= 3:
        for h in holdings:
            if h["weight"] is not None:
                h["weight"] *= 100
    for h in holdings:
        if h["weight"] is not None:
            h["weight"] = round(h["weight"], 4)
        if h["shares"] is not None:
            h["shares"] = int(round(h["shares"]))

    holdings.sort(key=lambda z: (z["weight"] or 0), reverse=True)
    return base_date, holdings


def fetch_latest_available(start_yyyymmdd: str, fid: str, debug: bool = False):
    d = datetime.strptime(start_yyyymmdd, "%Y%m%d")
    for _ in range(LOOKBACK_DAYS + 1):
        ds = d.strftime("%Y%m%d")
        try:
            content = download(ds, fid)
            base_date, holdings = normalize(read_table(content), debug=debug)
            if holdings:
                if not base_date:
                    base_date = f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"
                return base_date, holdings
            print(f"    - {ds}: 유효 데이터 없음, 하루 전으로")
        except Exception as e:
            print(f"    - {ds}: 다운로드/파싱 실패 ({e})")
        d -= timedelta(days=1)
    return None, None


# ── 변동 계산 (수량 중심) ────────────────────────────────────────────────
def diff(cur, prev):
    if not prev:
        return {"added": [], "removed": [], "bought": [], "sold": []}
    pmap = {h["key"]: h for h in prev}
    cmap = {h["key"]: h for h in cur}
    added, removed, bought, sold = [], [], [], []

    for k, h in cmap.items():
        if h.get("is_cash"):
            continue
        if k not in pmap:
            added.append({"name": h["name"], "ticker": h["ticker"],
                          "weight": h["weight"], "shares": h["shares"]})
            continue
        cs = h["shares"] or 0
        ps = pmap[k]["shares"] or 0
        ds = round(cs - ps, 4)
        if ds == 0:
            continue
        rec = {"name": h["name"], "ticker": h["ticker"],
               "shares": cs, "prev_shares": ps, "share_delta": ds,
               "weight": h["weight"], "prev_weight": pmap[k]["weight"],
               "weight_delta": round((h["weight"] or 0) - (pmap[k]["weight"] or 0), 4)}
        (bought if ds > 0 else sold).append(rec)

    for k, h in pmap.items():
        if h.get("is_cash"):
            continue
        if k not in cmap:
            removed.append({"name": h["name"], "ticker": h["ticker"],
                            "weight": h["weight"], "shares": h["shares"]})

    added.sort(key=lambda z: z["weight"] or 0, reverse=True)
    removed.sort(key=lambda z: z["weight"] or 0, reverse=True)
    bought.sort(key=lambda z: z["share_delta"], reverse=True)
    sold.sort(key=lambda z: z["share_delta"])
    return {"added": added, "removed": removed, "bought": bought, "sold": sold}


# ── 저장 ────────────────────────────────────────────────────────────────
def load_snapshot(snap_dir: Path, date_iso: str):
    f = snap_dir / f"{date_iso}.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))["holdings"]
    return None


def existing_dates(snap_dir: Path):
    return sorted(p.stem for p in snap_dir.glob("*.json"))


def process_etf(etf: dict, start: str, debug: bool = False):
    slug, fid = etf["slug"], etf["fid"]
    edir = DATA / slug
    snap_dir = edir / "snapshots"
    snap_dir.mkdir(parents=True, exist_ok=True)

    print(f"[{slug}] {etf['name']} ({etf['ticker']}) — 요청 기준일 {start}")
    date_iso, holdings = fetch_latest_available(start, fid, debug=debug)
    if not holdings:
        print(f"  [skip] 최근 영업일 데이터를 찾지 못함.")
        return

    prev_dates = [d for d in existing_dates(snap_dir) if d < date_iso]
    prev_date = prev_dates[-1] if prev_dates else None
    prev = load_snapshot(snap_dir, prev_date) if prev_date else None

    snap = {"date": date_iso, "ticker": etf["ticker"], "fund_id": fid,
            "name": etf["name"],
            "count": sum(1 for h in holdings if not h["is_cash"]),
            "holdings": holdings}
    (snap_dir / f"{date_iso}.json").write_text(
        json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")

    latest = dict(snap)
    latest["prev_date"] = prev_date
    latest["changes"] = diff(holdings, prev)
    (edir / "latest.json").write_text(
        json.dumps(latest, ensure_ascii=False, indent=2), encoding="utf-8")
    (edir / "dates.json").write_text(
        json.dumps(existing_dates(snap_dir), ensure_ascii=False, indent=2), encoding="utf-8")

    c = latest["changes"]
    print(f"  [ok] {date_iso} · {snap['count']}종목. 전일({prev_date}) 대비 "
          f"편입 {len(c['added'])} · 편출 {len(c['removed'])} · "
          f"추가매수 {len(c['bought'])} · 일부매도 {len(c['sold'])}")


def main():
    args = [a for a in sys.argv[1:] if a != "--debug"]
    debug = "--debug" in sys.argv
    start = args[0] if args else today_kst()

    DATA.mkdir(parents=True, exist_ok=True)
    # 사이트가 읽는 ETF 목록
    (DATA / "etfs.json").write_text(
        json.dumps([{k: e[k] for k in ("slug", "name", "ticker", "fid")} for e in ETFS],
                   ensure_ascii=False, indent=2), encoding="utf-8")

    for etf in ETFS:
        process_etf(etf, start, debug=debug)
    return 0


if __name__ == "__main__":
    sys.exit(main())
