import io
import requests
import pandas as pd
import yfinance as yf

START = "1990-06-01"
END   = "1991-04-30"
OUTPUT = "data/raw_stocks_gulf.csv"

STOCK_TICKERS = ["XOM", "CVX", "COP", "LMT", "RTX", "NOC"]

# EIA spot-price XLS — daily WTI back to 1986, no API key required.
# Used as fallback when FRED web endpoint is unavailable.
EIA_XLS_URL = "https://www.eia.gov/dnav/pet/xls/PET_PRI_SPT_S1_D.xls"

# ── 1. Stocks via yfinance ────────────────────────────────────────────────────
print("Fetching stocks from yfinance …")
frames = {}
for tkr in STOCK_TICKERS:
    df = yf.Ticker(tkr).history(start=START, end=END, auto_adjust=True)["Close"]
    df.index = df.index.tz_localize(None).normalize()
    frames[tkr] = df

stocks = pd.DataFrame(frames)
stocks.index.name = "Date"
print(f"  Stocks: {len(stocks)} rows, {stocks.index[0].date()} → {stocks.index[-1].date()}")

# ── 2. WTI crude oil spot price ───────────────────────────────────────────────
# Source priority:
#   1. FRED fredgraph.csv — vintage_date pins the series to data as of 1991-04-30,
#      avoiding any later revisions to the historical record.
#   2. EIA PET_PRI_SPT_S1_D.xls — official EIA WTI spot price, no API key needed.

FRED_URL = (
    "https://fred.stlouisfed.org/graph/fredgraph.csv"
    "?id=DCOILWTICO&vintage_date=1991-04-30"
)

def fetch_wti_fred(url, start, end, timeout=60):
    resp = requests.get(url, timeout=timeout)
    resp.raise_for_status()
    if resp.headers.get("Content-Type", "").startswith("text/html"):
        raise ValueError("FRED returned HTML instead of CSV (server-side error)")
    oil = pd.read_csv(io.StringIO(resp.text), parse_dates=["DATE"], index_col="DATE")
    oil.index.name = "Date"
    oil.columns = ["WTI_Crude"]
    oil = oil[oil["WTI_Crude"] != "."].copy()
    oil["WTI_Crude"] = oil["WTI_Crude"].astype(float)
    return oil.loc[start:end]

def fetch_wti_eia(url, start, end, timeout=60):
    resp = requests.get(url, timeout=timeout)
    resp.raise_for_status()
    xls  = pd.ExcelFile(io.BytesIO(resp.content))
    df   = xls.parse("Data 1", header=2, parse_dates=["Date"], index_col="Date")
    oil  = df.iloc[:, 0].rename("WTI_Crude").dropna()
    oil.index.name = "Date"
    oil.index = pd.to_datetime(oil.index)
    return oil.loc[start:end].to_frame()

print("Fetching WTI crude oil …")
oil = None
try:
    oil = fetch_wti_fred(FRED_URL, START, END)
    print(f"  WTI source: FRED fredgraph.csv  ({len(oil)} rows)")
except Exception as e:
    print(f"  FRED unavailable ({type(e).__name__}); falling back to EIA XLS …")
    oil = fetch_wti_eia(EIA_XLS_URL, START, END)
    print(f"  WTI source: EIA PET_PRI_SPT_S1_D.xls  ({len(oil)} rows)")

print(f"  Range: {oil.index[0].date()} → {oil.index[-1].date()}")

# ── 3. Merge on date (outer join so neither series clips the other) ───────────
merged = stocks.join(oil, how="outer").sort_index()
merged.index = merged.index.normalize()

# Keep only rows within our window
merged = merged.loc[START:END]

# Forward-fill WTI weekends/holidays to align with stock trading days,
# then restrict to days where at least one stock has a price
merged["WTI_Crude"] = merged["WTI_Crude"].ffill()
merged = merged.dropna(subset=STOCK_TICKERS, how="all")

print(f"\nMerged: {len(merged)} rows × {len(merged.columns)} columns")
print(merged.tail(5).to_string())

# ── 4. Save ───────────────────────────────────────────────────────────────────
merged.to_csv(OUTPUT)
print(f"\nSaved → {OUTPUT}")
