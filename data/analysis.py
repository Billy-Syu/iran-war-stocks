import pandas as pd

df = pd.read_csv("data/raw_stocks.csv", index_col="Date", parse_dates=True)

periods = {
    "before_war": ("2026-01-01", "2026-02-27"),
    "war_outbreak": ("2026-02-28", "2026-03-31"),
    "after_ceasefire": ("2026-04-08", "2026-06-02"),
}

def period_pct_change(df, start, end):
    mask = (df.index >= start) & (df.index <= end)
    sub = df[mask]
    if len(sub) < 2:
        return pd.Series({col: None for col in df.columns})
    return (sub.iloc[-1] / sub.iloc[0] - 1) * 100

rows = []
for name, (start, end) in periods.items():
    pct = period_pct_change(df, start, end)
    pct.name = name
    rows.append(pct)

results = pd.DataFrame(rows).T
results.index.name = "ticker"
results = results.round(4)

results.to_csv("data/analysis_results.csv")
print(results)
