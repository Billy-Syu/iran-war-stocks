import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

df = pd.read_csv("data/raw_stocks.csv", index_col="Date", parse_dates=True)

WAR_DATE = pd.Timestamp("2026-02-28")

GROUPS = {
    "energy": {
        "tickers": ["XOM", "CVX", "COP"],
        "title": "Energy Stocks — Jan to Jun 2026",
        "outfile": "charts/energy_stocks.png",
    },
    "defense": {
        "tickers": ["LMT", "RTX", "NOC"],
        "title": "Defense Stocks — Jan to Jun 2026",
        "outfile": "charts/defense_stocks.png",
    },
}

for group in GROUPS.values():
    fig, ax = plt.subplots(figsize=(12, 6))

    for ticker in group["tickers"]:
        ax.plot(df.index, df[ticker], linewidth=1.8, label=ticker)

    ax.axvline(WAR_DATE, color="red", linestyle="--", linewidth=1.4)
    ax.text(
        WAR_DATE,
        ax.get_ylim()[1],
        "  US-Iran War Begins",
        color="red",
        fontsize=9,
        va="top",
        rotation=90,
    )

    ax.set_title(group["title"], fontsize=14)
    ax.set_xlabel("Date")
    ax.set_ylabel("Closing Price (USD)")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.legend()
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(group["outfile"], dpi=150)
    plt.close(fig)
    print(f"Saved {group['outfile']}")
