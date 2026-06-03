"""
Generate a Traditional Chinese financial analysis report from analysis_results.csv
using the OpenAI GPT-4o API, and save the output to reports/final_report.md.
"""

import os
import sys
import textwrap
import pandas as pd
from openai import OpenAI

CSV_PATH = "data/analysis_results.csv"
REPORT_PATH = "reports/final_report.md"

PERIODS = {
    "before_war": ("2026-01-01", "2026-02-27"),
    "war_outbreak": ("2026-02-28", "2026-03-31"),
    "after_ceasefire": ("2026-04-08", "2026-06-02"),
}

PERIOD_LABELS = {
    "before_war": "戰前（2026-01-01 至 2026-02-27）",
    "war_outbreak": "戰爭爆發期（2026-02-28 至 2026-03-31）",
    "after_ceasefire": "停火後（2026-04-08 至 2026-06-02）",
}

TICKER_LABELS = {
    "CL=F": "原油期貨（WTI）",
    "XOM":  "埃克森美孚（XOM，能源）",
    "CVX":  "雪佛龍（CVX，能源）",
    "COP":  "康菲石油（COP，能源）",
    "LMT":  "洛克希德馬丁（LMT，國防）",
    "RTX":  "雷神科技（RTX，國防）",
    "NOC":  "諾斯洛普格魯曼（NOC，國防）",
}

ENERGY_TICKERS = ["XOM", "CVX", "COP"]
DEFENSE_TICKERS = ["LMT", "RTX", "NOC"]
OIL_TICKER = "CL=F"


def build_data_summary(df: pd.DataFrame) -> str:
    lines = ["各標的各期間漲跌幅（%）：", ""]
    header = f"{'標的':<30}" + "".join(f"{PERIOD_LABELS[p]:<30}" for p in PERIODS)
    lines.append(header)
    lines.append("-" * (30 + 30 * len(PERIODS)))

    for ticker, row in df.iterrows():
        label = TICKER_LABELS.get(ticker, ticker)
        cells = "".join(f"{row[p]:>+.2f}%{'':<22}" for p in PERIODS)
        lines.append(f"{label:<30}{cells}")

    return "\n".join(lines)


def build_prompt(data_summary: str) -> str:
    return textwrap.dedent(f"""
        你是一位資深金融分析師，專長為地緣政治事件對資本市場的影響分析。
        請根據以下數據，以**繁體中文**撰寫一份完整的財務分析報告（Markdown 格式）。

        ## 背景
        本報告分析 2026 年美伊戰爭（US-Iran War）對能源股與國防股的影響。
        時間軸分為三個階段：
        - 戰前（2026-01-01 至 2026-02-27）：局勢緊張但尚未開戰
        - 戰爭爆發期（2026-02-28 至 2026-03-31）：軍事行動正式展開
        - 停火後（2026-04-08 至 2026-06-02）：美伊達成停火協議

        ## 數據
        {data_summary}

        ## 報告要求
        請包含以下四個章節，每章節至少 200 字，並以數據為依據提出具體論述：

        1. **能源股表現分析**
           - 分析 XOM、CVX、COP 在三個階段的漲跌幅
           - 說明戰爭如何影響能源股，以及能源股反應是否符合預期

        2. **國防股表現分析**
           - 分析 LMT、RTX、NOC 在三個階段的漲跌幅
           - 討論國防股在戰爭期間的意外走勢並提出可能解釋

        3. **原油價格與股票走勢的相關性**
           - 分析原油（CL=F）與能源股、國防股的關聯
           - 說明原油價格在各階段的驅動因素

        4. **關鍵洞察與結論**
           - 總結本次事件的市場啟示
           - 提出投資人面對地緣政治風險時的策略建議

        請在報告開頭加上標題、摘要，並在結尾標注資料來源與分析日期（{pd.Timestamp.today().strftime('%Y-%m-%d')}）。
    """).strip()


def main():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        sys.exit("Error: OPENAI_API_KEY environment variable is not set.")

    df = pd.read_csv(CSV_PATH, index_col="ticker")

    # Reorder columns to match period order
    df = df[list(PERIODS.keys())]

    data_summary = build_data_summary(df)
    prompt = build_prompt(data_summary)

    print("Calling OpenAI GPT-4o ...")
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "system",
                "content": (
                    "你是一位專業的財務分析師，擅長以繁體中文撰寫清晰、深入的金融分析報告。"
                    "報告格式為 Markdown，語言為繁體中文，論述須以數據為根據。"
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.4,
        max_tokens=4096,
    )

    report_text = response.choices[0].message.content

    os.makedirs("reports", exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report_text)

    print(f"Report saved to {REPORT_PATH}")


if __name__ == "__main__":
    main()
