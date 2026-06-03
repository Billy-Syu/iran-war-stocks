import yfinance as yf
import pandas as pd

TICKERS = {
    "energy": ["XOM", "CVX", "COP"],
    "defense": ["LMT", "RTX", "NOC"],
    "commodities": ["CL=F"],
}

START = "2026-01-01"
END = "2026-06-03"
OUTPUT = "data/raw_stocks.csv"

all_tickers = [t for group in TICKERS.values() for t in group]

data = yf.download(all_tickers, start=START, end=END, auto_adjust=True)["Close"]

data.index.name = "Date"
data.to_csv(OUTPUT)

print(f"Saved {len(data)} rows × {len(data.columns)} tickers to {OUTPUT}")
print(data.tail())
