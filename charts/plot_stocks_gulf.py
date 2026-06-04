import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

df = pd.read_csv("data/raw_stocks_gulf.csv", index_col="Date", parse_dates=True)

EVENTS = [
    (pd.Timestamp("1990-08-02"), "Iraq Invades Kuwait",  "red"),
    (pd.Timestamp("1991-01-17"), "Desert Storm Begins",  "darkorange"),
    (pd.Timestamp("1991-02-28"), "Ceasefire",            "green"),
]

GROUPS = {
    "energy": {
        "tickers": ["XOM", "CVX", "COP"],
        "title":   "Energy Stocks — Gulf War 1990–1991",
        "outfile": "charts/energy_stocks_gulf.png",
    },
    "defense": {
        "tickers": ["LMT", "RTX", "NOC"],
        "title":   "Defense Stocks — Gulf War 1990–1991",
        "outfile": "charts/defense_stocks_gulf.png",
    },
}

for group in GROUPS.values():
    fig, ax = plt.subplots(figsize=(13, 6))

    for ticker in group["tickers"]:
        ax.plot(df.index, df[ticker], linewidth=1.8, label=ticker)

    # WTI on right axis
    ax2 = ax.twinx()
    ax2.plot(df.index, df["WTI_Crude"], color="orange", linewidth=1.4,
             linestyle="--", label="WTI Crude Oil", alpha=0.85)
    ax2.set_ylabel("WTI Crude Oil (USD/bbl)", color="orange")
    ax2.tick_params(axis="y", labelcolor="orange")

    # Draw once so get_ylim() reflects the autoscaled data range
    fig.canvas.draw()
    ymin, ymax = ax.get_ylim()

    # Vertical event lines — stagger label heights to avoid overlap
    label_positions = [0.97, 0.82, 0.67]
    for (date, label, color), label_frac in zip(EVENTS, label_positions):
        ax.axvline(date, color=color, linestyle="--", linewidth=1.3, alpha=0.85)
        ax.text(
            date,
            ymin + (ymax - ymin) * label_frac,
            f"  {label}",
            color=color,
            fontsize=8.5,
            va="top",
            rotation=90,
        )

    ax.set_title(group["title"], fontsize=14)
    ax.set_xlabel("Date")
    ax.set_ylabel("Closing Price (USD, split-adjusted)")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)

    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(group["outfile"], dpi=150)
    plt.close(fig)
    print(f"Saved {group['outfile']}")
