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

    ax2 = ax.twinx()
    ax2.plot(df.index, df["CL=F"], color="orange", linewidth=1.4,
             linestyle="--", label="CL=F (Crude Oil)", alpha=0.85)
    ax2.set_ylabel("Crude Oil Futures (USD)", color="orange")
    ax2.tick_params(axis="y", labelcolor="orange")

    ax.axvline(WAR_DATE, color="red", linestyle="--", linewidth=1.4)
    # place label just inside the top of the left axis after autoscale
    ax.draw(fig.canvas.get_renderer())
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

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left")

    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(group["outfile"], dpi=150)
    plt.close(fig)
    print(f"Saved {group['outfile']}")
