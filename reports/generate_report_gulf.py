"""
Generate a Traditional Chinese financial analysis report for the 1990-91 Gulf War
from data/analysis_results_gulf.csv using the OpenAI GPT-4o API, and save the
output to reports/final_report_gulf.md.

The report includes a dedicated contrast section comparing Gulf War defence-stock
behaviour (rose) with the 2026 US-Iran War (fell).
"""

import os
import sys
import textwrap
import pandas as pd
from openai import OpenAI

CSV_PATH = "data/analysis_results_gulf.csv"
REPORT_PATH = "reports/final_report_gulf.md"

PERIODS = {
    "before_invasion": ("1990-06-01", "1990-08-01"),
    "war_period":      ("1990-08-02", "1991-01-16"),
    "desert_storm":    ("1991-01-17", "1991-02-28"),
}

PERIOD_LABELS = {
    "before_invasion": "入侵前（1990-06-01 至 1990-08-01）",
    "war_period":      "佔領科威特期（1990-08-02 至 1991-01-16）",
    "desert_storm":    "沙漠風暴作戰（1991-01-17 至 1991-02-28）",
}

TICKER_LABELS = {
    "WTI_Crude": "原油現貨（WTI）",
    "XOM":       "埃克森美孚（XOM，能源）",
    "CVX":       "雪佛龍（CVX，能源）",
    "COP":       "康菲石油（COP，能源）",
    "LMT":       "洛克希德馬丁（LMT，國防）",
    "RTX":       "雷神科技（RTX，國防）",
    "NOC":       "諾斯洛普格魯曼（NOC，國防）",
}

# 2026 US-Iran War reference data for the contrast section
IRAN_WAR_DATA = {
    "periods": {
        "before_war":      "2026-01-01 至 2026-02-27",
        "war_outbreak":    "2026-02-28 至 2026-03-31",
        "after_ceasefire": "2026-04-08 至 2026-06-02",
    },
    "returns": {
        "WTI": {"before_war": "+16.92%", "war_outbreak": "+42.33%", "after_ceasefire": "+0.25%"},
        "XOM": {"before_war": "+25.17%", "war_outbreak": "+10.01%", "after_ceasefire": "-3.61%"},
        "CVX": {"before_war": "+20.97%", "war_outbreak": "+9.12%",  "after_ceasefire": "-1.88%"},
        "COP": {"before_war": "+18.25%", "war_outbreak": "+11.64%", "after_ceasefire": "-5.97%"},
        "LMT": {"before_war": "+32.39%", "war_outbreak": "-10.69%", "after_ceasefire": "-17.77%"},
        "RTX": {"before_war": "+8.57%",  "war_outbreak": "-9.08%",  "after_ceasefire": "-14.00%"},
        "NOC": {"before_war": "+24.08%", "war_outbreak": "-11.17%", "after_ceasefire": "-21.60%"},
    },
}


def build_data_summary(df: pd.DataFrame) -> str:
    col_w = 32
    lines = ["各標的各期間漲跌幅（%）：", ""]
    header = f"{'標的':<{col_w}}" + "".join(
        f"{PERIOD_LABELS[p]:<{col_w}}" for p in PERIODS
    )
    lines.append(header)
    lines.append("-" * (col_w + col_w * len(PERIODS)))

    for ticker, row in df.iterrows():
        label = TICKER_LABELS.get(ticker, ticker)
        cells = "".join(f"{row[p]:>+.2f}%{'':<{col_w - 8}}" for p in PERIODS)
        lines.append(f"{label:<{col_w}}{cells}")

    return "\n".join(lines)


def build_iran_reference_block() -> str:
    d = IRAN_WAR_DATA
    periods = list(d["periods"].keys())
    col_w = 28
    lines = [
        "【參照】2026 年美伊戰爭各期間漲跌幅（%）：",
        "",
        f"{'標的':<{col_w}}" + "".join(
            f"{d['periods'][p]:<{col_w}}" for p in periods
        ),
        "-" * (col_w * (1 + len(periods))),
    ]
    label_map = {
        "WTI": "原油（WTI）",
        "XOM": "埃克森美孚（XOM，能源）",
        "CVX": "雪佛龍（CVX，能源）",
        "COP": "康菲石油（COP，能源）",
        "LMT": "洛克希德馬丁（LMT，國防）",
        "RTX": "雷神科技（RTX，國防）",
        "NOC": "諾斯洛普格魯曼（NOC，國防）",
    }
    for key, label in label_map.items():
        cells = "".join(
            f"{d['returns'][key][p]:<{col_w}}" for p in periods
        )
        lines.append(f"{label:<{col_w}}{cells}")
    return "\n".join(lines)


def build_prompt(gulf_summary: str, iran_block: str) -> str:
    return textwrap.dedent(f"""
        你是一位資深金融分析師，專長為地緣政治事件對資本市場的影響分析。
        請根據以下數據，以**繁體中文**撰寫一份完整的財務分析報告（Markdown 格式）。

        ## 背景：1990-91 年波灣戰爭
        本報告分析波灣戰爭（Gulf War）對能源股與國防股的影響。
        時間軸分為三個階段：
        - 入侵前（1990-06-01 至 1990-08-01）：中東局勢緊張，伊拉克兵力集結
        - 佔領科威特期（1990-08-02 至 1991-01-16）：伊拉克入侵科威特，國際制裁與軍事部署
        - 沙漠風暴作戰（1991-01-17 至 1991-02-28）：多國聯軍發動空襲與地面攻勢，伊拉克撤軍

        ## 波灣戰爭數據
        {gulf_summary}

        ## 對照：2026 年美伊戰爭數據
        {iran_block}

        ## 報告要求
        請包含以下五個章節，每章節至少 200 字，並以數據為依據提出具體論述：

        1. **能源股表現分析**
           - 分析 XOM、CVX、COP 在波灣戰爭三個階段的漲跌幅
           - 說明原油供應衝擊如何傳導至能源股，以及各股表現的異同

        2. **國防股表現分析**
           - 分析 LMT、RTX、NOC 在波灣戰爭三個階段的漲跌幅
           - 重點說明為何國防股在佔領科威特期與沙漠風暴期大幅上漲

        3. **WTI 原油與股票走勢的相關性**
           - 分析 WTI 原油現貨在三個階段的走勢及驅動因素
           - 說明原油與能源股、國防股之間的關聯程度及方向

        4. **與 2026 年美伊戰爭的對比分析：為何國防股走勢截然相反？**
           - 波灣戰爭：LMT 佔領期 +24.25%、NOC +30.36%，沙漠風暴期 NOC +37.17%
           - 2026 美伊戰爭：LMT 戰爭爆發期 -10.69%、NOC -11.17%，停火後持續下跌
           - 請從以下角度深入分析差異原因：
             a) 戰爭規模與持續時間的預期差異
             b) 市場提前定價（pre-pricing）vs. 事後修正
             c) 財政環境與國防預算政治背景
             d) 武器系統演變（精準武器 vs. 大規模地面部署）
             e) 其他你認為關鍵的結構性因素

        5. **關鍵洞察與結論**
           - 統整兩場戰爭的市場教訓
           - 提出投資人面對地緣政治風險時的策略建議
           - 說明哪些指標最能預測戰爭期間能源股與國防股的走向

        請在報告開頭加上標題與摘要，並在結尾標注資料來源與分析日期（{pd.Timestamp.today().strftime('%Y-%m-%d')}）。
    """).strip()


def main():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        sys.exit("Error: OPENAI_API_KEY environment variable is not set.")

    df = pd.read_csv(CSV_PATH, index_col="ticker")
    df = df[list(PERIODS.keys())]

    gulf_summary = build_data_summary(df)
    iran_block = build_iran_reference_block()
    prompt = build_prompt(gulf_summary, iran_block)

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
                    "比較分析部分需有深度，避免流於表面。"
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
