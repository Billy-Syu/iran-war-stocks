import pandas as pd

df = pd.read_csv("data/raw_stocks_gulf.csv", index_col="Date", parse_dates=True)

periods = {
    "before_invasion": ("1990-06-01", "1990-08-01"),
    "war_period": ("1990-08-02", "1991-01-16"),
    "desert_storm": ("1991-01-17", "1991-02-28"),
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

results.to_csv("data/analysis_results_gulf.csv")
print(results)
