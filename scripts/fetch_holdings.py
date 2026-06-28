#!/usr/bin/env python3
"""
KoAct 미국나스닥성장기업액티브 (0015B0) 일별 전체 구성종목 수집기.

데이터 소스: 삼성액티브자산운용 '투자종목정보(PDF)' 엑셀 다운로드
    https://www.samsungactive.co.kr/excel_pdf.do?fId=2ETFQ1&gijunYMD=YYYYMMDD
운용사 팩트시트(상위 종목만)와 달리 이 엑셀은 전 종목을 담고 있다.

산출물:
    data/snapshots/YYYY-MM-DD.json  : 그날 전체 구성종목 스냅샷
    data/dates.json                 : 보유한 스냅샷 날짜 목록
    data/latest.json                : 최신 스냅샷 + 전일 대비 변동(편입/편출/비중증감)

사용:
    python scripts/fetch_holdings.py            # 오늘(KST) 기준, 없으면 직전 영업일로 자동 후퇴
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

# ── 설정 ────────────────────────────────────────────────────────────────
FID = "2ETFQ1"                          # 운용사 펀드 ID (KoAct 미국나스닥성장기업액티브)
TICKER = "0015B0"                       # 거래소 단축코드 (표시용)
ETF_NAME = "KoAct 미국나스닥성장기업액티브"
URL = "https://www.samsungactive.co.kr/excel_pdf.do"
WEIGHT_EPS = 0.05                       # 이 %p 이상 변할 때만 '비중 증가/감소'로 기록
LOOKBACK_DAYS = 7                       # 해당일 파일이 없으면 며칠 전까지 후퇴 탐색

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
SNAP = DATA / "snapshots"
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
    """'MU US Equity' -> 'MU'. 현금/특수코드는 빈 문자열."""
    code = clean(code)
    if not code or code.startswith(("CASH", "KRD", "KRW")):
        return ""
    return code.split()[0]


# ── 다운로드 + 파싱 ─────────────────────────────────────────────────────
def download(date_yyyymmdd: str) -> bytes:
    import requests
    r = requests.get(
        URL,
        params={"fId": FID, "gijunYMD": date_yyyymmdd},
        timeout=30,
        headers={"User-Agent": "Mozilla/5.0",
                 "Referer": "https://www.samsungactive.co.kr/"},
    )
    r.raise_for_status()
    return r.content


def read_table(content: bytes):
    """엑셀(.xls BIFF) 우선, 실패 시 HTML 표로 폴백. header 없이 원본 셀 그대로."""
    import pandas as pd
    # 구형 .xls (CDFV2/BIFF8)
    try:
        return pd.read_excel(io.BytesIO(content), header=None, dtype=str, engine="xlrd")
    except Exception:
        pass
    # 일부 응답이 HTML 표일 경우
    try:
        tables = pd.read_html(io.BytesIO(content))
        if tables:
            return max(tables, key=len)
    except Exception:
        pass
    return None


def normalize(raw, debug: bool = False):
    """원본 표(header=None DataFrame) -> (기준일 'YYYY-MM-DD', holdings[list])."""
    import pandas as pd
    if raw is None or len(raw) == 0:
        return None, None
    raw = raw.reset_index(drop=True)
    if debug:
        print(raw.head(6).to_string())

    # 헤더 행 찾기 (종목명 + ISIN/비중 포함)
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

    # 헤더 위쪽에서 기준일 추출
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
            "isin": isin,
            "name": name,
            "code": code,
            "ticker": clean_ticker(code),
            "weight": to_num(r.get(cW)) if cW else None,
            "shares": to_num(r.get(cQ)) if cQ else None,
            "amount": to_num(r.get(cA)) if cA else None,
            "is_cash": is_cash,
            "key": isin or code or name,
        })

    holdings = [h for h in holdings if h["key"]]

    # 비중 정규화: 소수(합≈1)로 들어오면 ×100 해서 백분율(합≈100)로 통일
    total = sum((h["weight"] or 0) for h in holdings)
    if 0 < total <= 3:
        for h in holdings:
            if h["weight"] is not None:
                h["weight"] *= 100
    for h in holdings:
        if h["weight"] is not None:
            h["weight"] = round(h["weight"], 4)

    holdings.sort(key=lambda z: (z["weight"] or 0), reverse=True)
    return base_date, holdings


def fetch_latest_available(start_yyyymmdd: str, debug: bool = False):
    """start일부터 과거로 내려가며 첫 유효 파일을 찾는다. (기준일, holdings) 반환."""
    d = datetime.strptime(start_yyyymmdd, "%Y%m%d")
    for _ in range(LOOKBACK_DAYS + 1):
        ds = d.strftime("%Y%m%d")
        try:
            content = download(ds)
            base_date, holdings = normalize(read_table(content), debug=debug)
            if holdings:
                if not base_date:
                    base_date = f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"
                return base_date, holdings
            print(f"  - {ds}: 유효 데이터 없음, 하루 전으로")
        except Exception as e:
            print(f"  - {ds}: 다운로드/파싱 실패 ({e})")
        d -= timedelta(days=1)
    return None, None


# ── 변동 계산 ────────────────────────────────────────────────────────────
def diff(cur, prev):
    if not prev:
        return {"added": [], "removed": [], "increased": [], "decreased": []}
    pmap = {h["key"]: h for h in prev} if prev else {}
    cmap = {h["key"]: h for h in cur}
    added, removed, increased, decreased = [], [], [], []

    for k, h in cmap.items():
        if h.get("is_cash"):
            continue
        if k not in pmap:
            added.append({"name": h["name"], "ticker": h["ticker"], "weight": h["weight"]})
            continue
        cw = h["weight"] or 0
        pw = pmap[k]["weight"] or 0
        d = round(cw - pw, 4)
        if d >= WEIGHT_EPS:
            increased.append({"name": h["name"], "ticker": h["ticker"],
                              "weight": cw, "prev_weight": pw, "delta": d})
        elif d <= -WEIGHT_EPS:
            decreased.append({"name": h["name"], "ticker": h["ticker"],
                              "weight": cw, "prev_weight": pw, "delta": d})

    for k, h in pmap.items():
        if h.get("is_cash"):
            continue
        if k not in cmap:
            removed.append({"name": h["name"], "ticker": h["ticker"], "weight": h["weight"]})

    added.sort(key=lambda z: z["weight"] or 0, reverse=True)
    removed.sort(key=lambda z: z["weight"] or 0, reverse=True)
    increased.sort(key=lambda z: z["delta"], reverse=True)
    decreased.sort(key=lambda z: z["delta"])
    return {"added": added, "removed": removed,
            "increased": increased, "decreased": decreased}


# ── 저장 ────────────────────────────────────────────────────────────────
def load_snapshot(date_iso):
    f = SNAP / f"{date_iso}.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))["holdings"]
    return None


def existing_dates():
    return sorted(p.stem for p in SNAP.glob("*.json"))


def main():
    args = [a for a in sys.argv[1:] if a != "--debug"]
    debug = "--debug" in sys.argv
    start = args[0] if args else today_kst()

    SNAP.mkdir(parents=True, exist_ok=True)
    print(f"[fetch] {ETF_NAME} ({TICKER}) — 요청 기준일 {start}")

    date_iso, holdings = fetch_latest_available(start, debug=debug)
    if not holdings:
        print("[skip] 최근 영업일 데이터를 찾지 못함. 정상 종료.")
        return 0

    prev_dates = [d for d in existing_dates() if d < date_iso]
    prev_date = prev_dates[-1] if prev_dates else None
    prev = load_snapshot(prev_date) if prev_date else None

    snap = {"date": date_iso, "ticker": TICKER, "fund_id": FID, "name": ETF_NAME,
            "count": sum(1 for h in holdings if not h["is_cash"]),
            "holdings": holdings}
    (SNAP / f"{date_iso}.json").write_text(
        json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")

    latest = dict(snap)
    latest["prev_date"] = prev_date
    latest["changes"] = diff(holdings, prev)
    (DATA / "latest.json").write_text(
        json.dumps(latest, ensure_ascii=False, indent=2), encoding="utf-8")

    (DATA / "dates.json").write_text(
        json.dumps(existing_dates(), ensure_ascii=False, indent=2), encoding="utf-8")

    c = latest["changes"]
    print(f"[ok] {date_iso} · {snap['count']}종목 저장. 전일({prev_date}) 대비 "
          f"편입 {len(c['added'])} · 편출 {len(c['removed'])} · "
          f"비중↑ {len(c['increased'])} · 비중↓ {len(c['decreased'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
