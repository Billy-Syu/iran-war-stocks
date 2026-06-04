# 戰爭與金融市場：Spec-Driven Multi-Agent 實作

**EMBA 生成式人工智慧導論｜期末實作報告｜Billy Chu 2026**

---

## 研究主題

以 **Spec-Driven Multi-Agent** 架構，自動化分析兩場中東戰爭對美國能源股與國防股的影響，並進行跨戰爭對比：

- 🏜 **1990 波灣戰爭**（1990/8/2 伊拉克入侵科威特 — 1991/2/28 停火）
- ⚔️ **2026 美伊戰爭**（2026/2/28 開打 — 2026/4/8 停火）

---

## 研究標的

### 股票（6 支）

| 代號 | 公司全名 | 類型 |
|------|---------|------|
| XOM | Exxon Mobil Corporation | 能源股 |
| CVX | Chevron Corporation | 能源股 |
| COP | ConocoPhillips | 能源股 |
| LMT | Lockheed Martin Corporation | 國防股 |
| RTX | RTX Corporation（原 Raytheon Technologies）| 國防股 |
| NOC | Northrop Grumman Corporation | 國防股 |

### 石油數據（2 個來源）

| 代號 | 全名 | 用於 |
|------|------|------|
| CL=F | NYMEX WTI Crude Oil Futures（紐約商業交易所 西德克薩斯中級原油期貨）| 2026 美伊戰爭 |
| DCOILWTICO | WTI Crude Oil Prices — FRED, Federal Reserve Bank of St. Louis | 1990 波灣戰爭 |

---

## Multi-Agent 架構

```
你（人類）決定題目與 Spec
        │
        ▼
Tab 1：主控 Agent（Claude Code）
├── 資料抓取：fetch_stocks.py / fetch_stocks_gulf.py
└── 視覺化：plot_stocks.py / plot_stocks_gulf.py
        │
        ├──────────────────────────┐
        ▼                          ▼
Tab 3：副 Agent A              Tab 4：副 Agent B
（Claude Code）                （Claude Code × GPT-4o API）
數據分析                        財經報告撰寫
analysis.py                    generate_report.py
analysis_gulf.py               generate_report_gulf.py
        │                          │
        └──────────────────────────┘
                    │
                    ▼
              Tab 2：Git
         每個 Agent 完成即 commit
```

---

## 專案結構

```
iran-war-stocks/
├── data/
│   ├── fetch_stocks.py          # 抓取美伊戰爭股價（yfinance）
│   ├── fetch_stocks_gulf.py     # 抓取波灣戰爭股價（yfinance + FRED）
│   ├── raw_stocks.csv           # 美伊戰爭原始數據
│   ├── raw_stocks_gulf.csv      # 波灣戰爭原始數據
│   ├── analysis.py              # 美伊戰爭三時期漲跌幅分析
│   ├── analysis_gulf.py         # 波灣戰爭三時期漲跌幅分析
│   ├── analysis_results.csv     # 美伊戰爭分析結果
│   └── analysis_results_gulf.csv # 波灣戰爭分析結果
├── charts/
│   ├── energy_stocks.png        # 美伊戰爭能源股走勢圖
│   ├── defense_stocks.png       # 美伊戰爭國防股走勢圖
│   ├── energy_stocks_gulf.png   # 波灣戰爭能源股走勢圖
│   └── defense_stocks_gulf.png  # 波灣戰爭國防股走勢圖
└── reports/
    ├── final_report.md          # 美伊戰爭財經分析報告（GPT-4o 產出）
    ├── final_report_gulf.md     # 波灣戰爭財經分析報告（GPT-4o 產出）
    ├── gulf_vs_iran_comparison.pptx  # 跨戰爭對比 PPT（主要報告）
    ├── iran_war_stocks_v2.pptx  # 美伊戰爭分析 PPT
    └── generate_ppt_comparison.py   # PPT 產生腳本
```

---

## Git Commit 記錄

| Commit | 說明 | Agent |
|--------|------|-------|
| `a03959a` | feat: fetch stock and crude oil data | 主控 Agent |
| `9d60b13` | feat: plot energy and defense stock charts | 主控 Agent |
| `ed252fa` | feat: add crude oil second y-axis | 主控 Agent |
| `bea7eb8` | feat: analysis of stock performance across 3 war periods | 副 Agent A |
| `d31e864` | feat: generate financial report via GPT-4o | 副 Agent B |
| `43af092` | chore: add .gitignore | — |
| `7756e7c` | feat: plot Gulf War stock charts | 主控 Agent |
| `73166f6` | feat: Gulf War three-period stock analysis | 副 Agent A |
| `09aed77` | feat: generate Gulf War financial report via GPT-4o | 副 Agent B |
| `71766ca` | feat: add Gulf War analysis and comparison | 主控 Agent |

---

## 關鍵發現

> **同樣六支股票，同樣是戰爭，國防股反應完全相反！**

| | 國防股（開戰後）| 能源股（開戰後）|
|--|--|--|
| 🏜 1990 波灣 | 📈 大漲 +24~37% | 先漲後整理 |
| ⚔️ 2026 美伊 | 📉 大跌 -10~22% | 📈 大漲 +8~10% |

### 兩個假說

- **假說 A（美軍表現）**：伊朗成功封鎖霍爾木茲海峽、飛彈穿透美軍防禦，市場對武器系統失去信心
- **假說 B（利多出盡）**：2026 年局勢早有預兆，市場提前定價，開打後反手賣出

---

## 核心反思

> **AI 能告訴你「發生了什麼」，但「為什麼」和「代表什麼意義」，仍需要人類的批判性思考。**

| AI Agent 能做的 | 人類不可取代的 |
|----------------|--------------|
| 大量數據收集與整理 | 決定要研究什麼問題 |
| 精確計算漲跌幅 | 設計有意義的 Spec 規格 |
| 自動產出圖表與報告 | 發現反直覺的結果 |
| 多 agent 平行作業 | 提出「為什麼」的假說 |
| 跨 AI 平台協作 | 為最終結論負責 |

---

## 技術環境

- **Shell**：macOS zsh（iTerm2 多 Tab）
- **主控 Agent**：Claude Code CLI
- **副 Agent A**：Claude Code CLI（Tab 3）
- **副 Agent B**：Claude Code × OpenAI GPT-4o API（Tab 4）
- **版本控制**：Git + GitHub
- **數據來源**：Yahoo Finance（yfinance）、FRED（DCOILWTICO）
- **分析工具**：Python、pandas、matplotlib、python-pptx
