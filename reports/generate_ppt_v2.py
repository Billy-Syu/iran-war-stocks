import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── palette ───────────────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1F, 0x35, 0x64)
MID_BLUE   = RGBColor(0x2E, 0x74, 0xB5)
LIGHT_BLUE = RGBColor(0xBD, 0xD7, 0xEE)
ORANGE     = RGBColor(0xED, 0x7D, 0x31)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
MID_GRAY   = RGBColor(0xCC, 0xCC, 0xCC)
DARK_GRAY  = RGBColor(0x70, 0x70, 0x70)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
RED        = RGBColor(0xC0, 0x00, 0x00)
GREEN      = RGBColor(0x37, 0x86, 0x30)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── primitives ────────────────────────────────────────────────────────────────

def new_prs():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs

def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def rect(slide, left, top, w, h, fill, line_color=None):
    s = slide.shapes.add_shape(1, left, top, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_color:
        s.line.color.rgb = line_color
    else:
        s.line.fill.background()
    return s

def oval(slide, left, top, w, h, fill, line_color=WHITE):
    s = slide.shapes.add_shape(9, left, top, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = line_color
    return s

def textbox(slide, text, left, top, w, h,
            size=14, bold=False, italic=False,
            color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(left, top, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text       = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def header_band(slide, title, subtitle=None):
    rect(slide, 0, 0, SLIDE_W, Inches(1.3), DARK_BLUE)
    textbox(slide, title,
            Inches(0.45), Inches(0.1), Inches(12.4), Inches(0.85),
            size=26, bold=True, color=WHITE)
    if subtitle:
        textbox(slide, subtitle,
                Inches(0.45), Inches(0.85), Inches(12.4), Inches(0.38),
                size=13, color=LIGHT_BLUE)

def bullet_rows(slide, items, left, top, w, row_h,
                size=13, color=DARK_BLUE, marker="▸", marker_color=None):
    mc = marker_color or MID_BLUE
    for i, text in enumerate(items):
        y = top + i * row_h
        textbox(slide, marker, left, y, Inches(0.3), row_h,
                size=size, bold=True, color=mc)
        textbox(slide, text, left + Inches(0.3), y, w - Inches(0.3), row_h,
                size=size, color=color)

def label_box(slide, title, body, left, top, w, h,
              fill=MID_BLUE, title_size=14, body_size=11, text_col=WHITE):
    rect(slide, left, top, w, h, fill, WHITE)
    txb = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.1),
                                   w - Inches(0.2), h - Inches(0.2))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(title_size)
    r.font.bold = True
    r.font.color.rgb = text_col
    if body:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = body
        r2.font.size = Pt(body_size)
        r2.font.color.rgb = text_col


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title
# ═══════════════════════════════════════════════════════════════════════════════
def slide_01_title(prs):
    slide = blank_slide(prs)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_BLUE)
    rect(slide, 0, 0, Inches(0.22), SLIDE_H, ORANGE)
    rect(slide, Inches(0.22), 0, SLIDE_W, Inches(0.07), MID_BLUE)

    textbox(slide, "美伊戰爭金融市場衝擊分析",
            Inches(0.6), Inches(1.7), Inches(12.1), Inches(1.1),
            size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "Spec-Driven Multi-Agent 實作報告",
            Inches(0.6), Inches(2.75), Inches(12.1), Inches(0.75),
            size=26, color=ORANGE, align=PP_ALIGN.CENTER)
    textbox(slide, "US–Iran War  ·  Impact on Energy & Defense Stocks",
            Inches(0.6), Inches(3.55), Inches(12.1), Inches(0.5),
            size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    rect(slide, Inches(3.5), Inches(4.25), Inches(6.3), Inches(0.04), MID_BLUE)

    textbox(slide, "生成式人工智慧導論  ·  Final Term Project  ·  2026-06-04",
            Inches(0.6), Inches(4.45), Inches(12.1), Inches(0.45),
            size=13, color=RGBColor(0x9D, 0xC3, 0xE6), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Research Topic & Motivation
# ═══════════════════════════════════════════════════════════════════════════════
def slide_02_motivation(prs):
    slide = blank_slide(prs)
    header_band(slide, "研究主題與動機", "Research Topic & Motivation")

    rect(slide, Inches(0.4), Inches(1.45), Inches(5.8), Inches(5.75), LIGHT_GRAY)
    textbox(slide, "研究主題",
            Inches(0.55), Inches(1.6), Inches(5.5), Inches(0.5),
            size=16, bold=True, color=DARK_BLUE)
    rect(slide, Inches(0.55), Inches(2.15), Inches(0.06), Inches(3.6), ORANGE)
    bullet_rows(slide, [
        "2026 年 2 月 28 日美伊戰爭爆發，\n地緣政治風險急遽升高",
        "能源股與國防股被視為戰爭受益標的，\n但市場反應是否符合直覺？",
        "原油期貨 (CL=F) 如何領先或滯後\n於股票市場？",
        "三個時間段：戰前、爆發期、停火後\n各有何不同的市場邏輯？",
    ], Inches(0.75), Inches(2.2), Inches(5.3),
       Inches(0.88), size=12, color=DARK_BLUE, marker="◆", marker_color=ORANGE)

    rect(slide, Inches(6.6), Inches(1.45), Inches(6.45), Inches(5.75),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "我們想回答的問題",
            Inches(6.75), Inches(1.6), Inches(6.1), Inches(0.5),
            size=16, bold=True, color=DARK_BLUE)
    rect(slide, Inches(6.75), Inches(2.15), Inches(0.06), Inches(3.6), MID_BLUE)
    bullet_rows(slide, [
        "戰爭爆發時，能源股真的會上漲嗎？",
        "國防股在戰爭期間的表現是否符合\n「戰爭概念股」預期？",
        "原油價格是市場的領先指標嗎？",
        "停火後資金如何重新配置？",
    ], Inches(6.95), Inches(2.2), Inches(5.9),
       Inches(0.88), size=12, color=DARK_BLUE, marker="？", marker_color=MID_BLUE)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — What is Spec-Driven
# ═══════════════════════════════════════════════════════════════════════════════
def slide_03_spec_driven(prs):
    slide = blank_slide(prs)
    header_band(slide, "什麼是 Spec-Driven 開發？",
                "Spec-Driven Development Applied to AI Agents")

    rect(slide, Inches(0.4), Inches(1.45), Inches(12.53), Inches(1.1),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide,
            "核心概念：先寫「規格書 (Spec)」，再讓 AI Agent 按規格執行 — "
            "人類定義 WHAT，Agent 決定 HOW",
            Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.85),
            size=14, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    steps = [
        ("1  定義規格",
         "明確寫出任務目標、\n輸入/輸出格式、品質標準\n(本研究：CSV 格式、時間區間、\n圖表樣式)",
         DARK_BLUE),
        ("2  Agent 執行",
         "Claude Code 解讀規格，\n自動分派子任務給\nAgent A / Agent B，\ngit commit 記錄每步驟",
         MID_BLUE),
        ("3  驗證與迭代",
         "人類審查輸出是否符合規格，\n不符則修改規格重跑 —\n而非修改 Agent 內部邏輯",
         ORANGE),
    ]
    bw, bh, by = Inches(3.8), Inches(3.3), Inches(2.75)
    for i, (title, body, color) in enumerate(steps):
        bx = Inches(0.4) + i * Inches(4.3)
        rect(slide, bx, by, bw, bh, color)
        textbox(slide, title, bx + Inches(0.15), by + Inches(0.15),
                bw - Inches(0.3), Inches(0.55), size=16, bold=True, color=WHITE)
        rect(slide, bx + Inches(0.15), by + Inches(0.72),
             bw - Inches(0.3), Inches(0.04), WHITE)
        textbox(slide, body, bx + Inches(0.15), by + Inches(0.85),
                bw - Inches(0.3), bh - Inches(1.0), size=12, color=WHITE)
        if i < 2:
            textbox(slide, "→", bx + bw + Inches(0.1), by + Inches(1.3),
                    Inches(0.4), Inches(0.6), size=24, bold=True,
                    color=MID_GRAY, align=PP_ALIGN.CENTER)

    textbox(slide,
            "優點：可重現、可審計、人類保有控制權；每次 git commit 都對應一條規格指令",
            Inches(0.4), Inches(6.35), Inches(12.53), Inches(0.45),
            size=11, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Multi-Agent Architecture
# ═══════════════════════════════════════════════════════════════════════════════
def slide_04_architecture(prs):
    slide = blank_slide(prs)
    header_band(slide, "多智能體工作流程架構",
                "Multi-Agent Workflow  ·  Claude Code  +  Agent A  +  GPT-4o Agent B")

    # git log ribbon
    rect(slide, Inches(11.6), Inches(1.35), Inches(1.73), Inches(6.15),
         RGBColor(0xF0, 0xF8, 0xE8))
    textbox(slide, "git log", Inches(11.65), Inches(1.45), Inches(1.6), Inches(0.4),
            size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    commits = [
        "fetch_stocks.py", "raw_stocks.csv",
        "analysis.py", "analysis_results.csv",
        "plot_stocks.py", "energy/defense .png",
        "generate_report.py", "final_report.md",
        "generate_ppt_v2.py", "iran_war_stocks_v2.pptx",
    ]
    for i, c in enumerate(commits):
        y = Inches(1.95) + i * Inches(0.41)
        rect(slide, Inches(11.75), y + Inches(0.07), Inches(0.12), Inches(0.24), GREEN)
        textbox(slide, c, Inches(11.95), y, Inches(1.3), Inches(0.38),
                size=9, color=DARK_GRAY)

    # Orchestrator box
    cx, cy, bw, bh = Inches(5.5), Inches(2.0), Inches(2.8), Inches(1.2)
    rect(slide, cx, cy, bw, bh, DARK_BLUE, WHITE)
    textbox(slide, "Claude Code", cx, cy + Inches(0.1), bw, Inches(0.5),
            size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "Orchestrator\n規格解讀 · 任務分派",
            cx, cy + Inches(0.55), bw, Inches(0.55),
            size=10, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    # Spec box above orchestrator
    rect(slide, cx + Inches(0.5), Inches(1.35), Inches(1.8), Inches(0.55), ORANGE)
    textbox(slide, "人類撰寫規格 (Spec)",
            cx + Inches(0.5), Inches(1.35), Inches(1.8), Inches(0.55),
            size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rect(slide, cx + Inches(1.35), Inches(1.9), Inches(0.06), Inches(0.1), MID_GRAY)

    # Arrows down to agents
    mid_x = cx + bw / 2
    rect(slide, mid_x - Inches(2.2), cy + bh, Inches(0.06), Inches(0.7), MID_GRAY)
    rect(slide, mid_x + Inches(2.14), cy + bh, Inches(0.06), Inches(0.7), MID_GRAY)

    # Agent A
    ax, ay = Inches(2.2), cy + bh + Inches(0.7)
    rect(slide, ax, ay, bw, Inches(1.55), MID_BLUE, WHITE)
    textbox(slide, "Agent A", ax, ay + Inches(0.08), bw, Inches(0.45),
            size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "數據抓取 & 圖表生成\nyfinance · matplotlib\n→ raw_stocks.csv · .png",
            ax, ay + Inches(0.5), bw, Inches(1.0),
            size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # Agent B
    bx2, by2 = Inches(8.8), ay
    rect(slide, bx2, by2, bw, Inches(1.55), RGBColor(0x10, 0x7C, 0x10), WHITE)
    textbox(slide, "GPT-4o  Agent B", bx2, by2 + Inches(0.08), bw, Inches(0.45),
            size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "報告撰寫 & 洞察分析\nanalysis.py · generate_report.py\n→ final_report.md · .pptx",
            bx2, by2 + Inches(0.5), bw, Inches(1.0),
            size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # Output arrows + box
    out_y = ay + Inches(1.55) + Inches(0.3)
    rect(slide, ax + bw / 2 - Inches(0.03), ay + Inches(1.55),
         Inches(0.06), Inches(0.3), MID_GRAY)
    rect(slide, bx2 + bw / 2 - Inches(0.03), ay + Inches(1.55),
         Inches(0.06), Inches(0.3), MID_GRAY)
    rect(slide, Inches(3.5), out_y, Inches(6.8), Inches(0.75), LIGHT_GRAY)
    textbox(slide, "iran_war_stocks_v2.pptx  ·  final_report.md",
            Inches(3.5), out_y, Inches(6.8), Inches(0.75),
            size=12, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # Data source
    label_box(slide, "Yahoo Finance",
              "XOM · CVX · COP\nLMT · RTX · NOC · CL=F",
              Inches(0.3), Inches(3.6), Inches(1.8), Inches(1.3),
              LIGHT_GRAY, title_size=11, body_size=9, text_col=DARK_BLUE)
    rect(slide, Inches(2.1), Inches(4.15), Inches(0.1), Inches(0.06), MID_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Event Background / Timeline
# ═══════════════════════════════════════════════════════════════════════════════
def slide_05_timeline(prs):
    slide = blank_slide(prs)
    header_band(slide, "事件背景：美伊戰爭時間軸", "US–Iran War Timeline  ·  2026")

    events = [
        ("2026-01-01", "觀察期開始\nBaseline",      MID_BLUE),
        ("2026-02-28", "戰爭爆發\nWar Outbreak",    RED),
        ("2026-04-08", "停火協議\nCeasefire",        ORANGE),
        ("2026-05-05", "軍事行動結束\nOp. Ended",    GREEN),
        ("2026-06-04", "資料截止\nData End",         DARK_GRAY),
    ]

    bar_top, bar_left, bar_w = Inches(3.9), Inches(0.9), Inches(11.5)
    n = len(events)
    rect(slide, bar_left, bar_top + Inches(0.2), bar_w, Inches(0.08), MID_GRAY)
    lx = [bar_left + i * (bar_w / (n - 1)) for i in range(n)]

    for i, (date, label, color) in enumerate(events):
        x = lx[i]
        oval(slide, x - Inches(0.14), bar_top + Inches(0.1),
             Inches(0.28), Inches(0.28), color)
        textbox(slide, date, x - Inches(0.9), bar_top - Inches(0.65),
                Inches(1.8), Inches(0.4), size=11, bold=True,
                color=DARK_BLUE, align=PP_ALIGN.CENTER)
        textbox(slide, label, x - Inches(1.0), bar_top + Inches(0.55),
                Inches(2.0), Inches(0.9), size=11, color=color, align=PP_ALIGN.CENTER)

    def period_span(label, x1, x2, color):
        mid = (x1 + x2) / 2
        sy = Inches(2.45)
        rect(slide, x1, sy + Inches(0.42), x2 - x1, Inches(0.05), color)
        rect(slide, x1, sy + Inches(0.18), Inches(0.05), Inches(0.24), color)
        rect(slide, x2 - Inches(0.05), sy + Inches(0.18), Inches(0.05), Inches(0.24), color)
        textbox(slide, label, mid - Inches(1.3), sy, Inches(2.6), Inches(0.4),
                size=11, italic=True, bold=True, color=color, align=PP_ALIGN.CENTER)

    period_span("戰前期  Before War",       lx[0], lx[1], MID_BLUE)
    period_span("戰爭爆發期  War Outbreak", lx[1], lx[2], RED)
    period_span("停火後  After Ceasefire",  lx[2], lx[4], GREEN)

    bullet_rows(slide, [
        "伊朗核武計畫談判破裂，美軍實施精準打擊德黑蘭核設施",
        "霍木茲海峽短暫封鎖，全球約 20% 石油運輸受阻",
        "4 月 8 日在聯合國斡旋下達成臨時停火，5 月 5 日正式結束敵對行動",
    ], Inches(0.5), Inches(5.6), Inches(12.5),
       Inches(0.55), size=11, color=DARK_BLUE, marker="•", marker_color=RED)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Energy Stocks Chart
# ═══════════════════════════════════════════════════════════════════════════════
def slide_06_energy(prs):
    slide = blank_slide(prs)
    header_band(slide, "能源股走勢分析",
                "Energy Stocks  ·  XOM · CVX · COP  +  WTI Crude Oil (right axis)")

    slide.shapes.add_picture("charts/energy_stocks.png",
                             Inches(0.3), Inches(1.38), Inches(9.3), Inches(5.95))

    rect(slide, Inches(9.75), Inches(1.38), Inches(3.3), Inches(5.95),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "關鍵發現", Inches(9.9), Inches(1.5), Inches(3.0), Inches(0.45),
            size=14, bold=True, color=DARK_BLUE)
    rect(slide, Inches(9.9), Inches(1.95), Inches(2.9), Inches(0.04), MID_BLUE)

    fy = Inches(2.1)
    for title, body, color in [
        ("戰前期\n+16%～+25%",
         "市場提前反應供應\n中斷預期，三股\n同步上揚", MID_BLUE),
        ("爆發期\n+9%～+12%",
         "原油飆升 +42%\n帶動能源股續漲\n（但漲幅縮窄）", ORANGE),
        ("停火後\n-2%～-6%",
         "供應恢復預期使\n能源股獲利了結\n回落", RED),
    ]:
        rect(slide, Inches(9.9), fy, Inches(2.9), Inches(0.04), color)
        textbox(slide, title, Inches(9.9), fy + Inches(0.08),
                Inches(2.9), Inches(0.5), size=12, bold=True, color=color)
        textbox(slide, body, Inches(9.9), fy + Inches(0.58),
                Inches(2.9), Inches(0.75), size=11, color=DARK_BLUE)
        fy += Inches(1.55)

    textbox(slide, "原油領先能源股\n約 1-2 個交易日",
            Inches(9.9), Inches(6.55), Inches(2.9), Inches(0.55),
            size=11, italic=True, color=DARK_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Defense Stocks Chart
# ═══════════════════════════════════════════════════════════════════════════════
def slide_07_defense(prs):
    slide = blank_slide(prs)
    header_band(slide, "國防股走勢分析",
                "Defense Stocks  ·  LMT · RTX · NOC  +  WTI Crude Oil (right axis)")

    slide.shapes.add_picture("charts/defense_stocks.png",
                             Inches(0.3), Inches(1.38), Inches(9.3), Inches(5.95))

    rect(slide, Inches(9.75), Inches(1.38), Inches(3.3), Inches(5.95),
         RGBColor(0xFF, 0xF0, 0xF0))
    textbox(slide, "關鍵發現", Inches(9.9), Inches(1.5), Inches(3.0), Inches(0.45),
            size=14, bold=True, color=DARK_BLUE)
    rect(slide, Inches(9.9), Inches(1.95), Inches(2.9), Inches(0.04), RED)

    fy = Inches(2.1)
    for title, body, color in [
        ("戰前期\n+8%～+32%",
         "市場預期戰爭提前\n買入，LMT 大漲 32%\n為三股最高", MID_BLUE),
        ("爆發期\n-9%～-11%",
         "出乎意料的下跌！\n市場判斷短期衝突\n不增加國防預算", RED),
        ("停火後\n-14%～-22%",
         "軍事需求預期降溫，\n資金轉向能源股，\n持續回落", RED),
    ]:
        rect(slide, Inches(9.9), fy, Inches(2.9), Inches(0.04), color)
        textbox(slide, title, Inches(9.9), fy + Inches(0.08),
                Inches(2.9), Inches(0.5), size=12, bold=True, color=color)
        textbox(slide, body, Inches(9.9), fy + Inches(0.58),
                Inches(2.9), Inches(0.75), size=11, color=DARK_BLUE)
        fy += Inches(1.55)

    rect(slide, Inches(9.9), Inches(6.45), Inches(2.9), Inches(0.72),
         RGBColor(0xFF, 0xE8, 0xE8))
    textbox(slide, "違反直覺：戰爭期間\n國防股反而下跌！",
            Inches(9.95), Inches(6.48), Inches(2.8), Inches(0.65),
            size=11, bold=True, color=RED, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Performance Table
# ═══════════════════════════════════════════════════════════════════════════════
def slide_08_table(prs, csv_path):
    slide = blank_slide(prs)
    header_band(slide, "各階段漲跌幅彙整",
                "Percentage Change by Period  ·  source: data/analysis_results.csv")

    df = pd.read_csv(csv_path)
    ENERGY = {"XOM", "CVX", "COP", "CL=F"}

    tbl = slide.shapes.add_table(
        len(df) + 1, 4,
        Inches(1.2), Inches(1.5), Inches(10.9), Inches(5.75)
    ).table

    for ci, w in enumerate([Inches(2.2), Inches(2.9), Inches(2.9), Inches(2.9)]):
        tbl.columns[ci].width = w

    def cell(r, c, text, bold=False, size=13, bg=None, fg=WHITE,
             align=PP_ALIGN.CENTER):
        cl = tbl.cell(r, c)
        cl.text = ""
        tf = cl.text_frame
        tf.paragraphs[0].alignment = align
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size  = Pt(size)
        run.font.bold  = bold
        run.font.color.rgb = fg
        if bg:
            cl.fill.solid()
            cl.fill.fore_color.rgb = bg

    for ci, h in enumerate(["股票代碼\nTicker",
                             "戰前期\nBefore War",
                             "戰爭爆發期\nWar Outbreak",
                             "停火後\nAfter Ceasefire"]):
        cell(0, ci, h, bold=True, bg=DARK_BLUE, size=12)

    energy_df  = df[df["ticker"].isin(ENERGY)]
    defense_df = df[~df["ticker"].isin(ENERGY)]
    for ri, (_, row) in enumerate(pd.concat([energy_df, defense_df]).iterrows()):
        tkr = row["ticker"]
        is_e = tkr in ENERGY
        bg = RGBColor(0xE8, 0xF0, 0xF8) if is_e else RGBColor(0xFF, 0xF0, 0xF0)
        fc = MID_BLUE if is_e else RED
        cell(ri + 1, 0, tkr, bold=True, bg=bg, fg=fc, size=14)
        for ci, col in enumerate(["before_war", "war_outbreak", "after_ceasefire"]):
            v = row[col]
            cell(ri + 1, ci + 1, f"{v:+.2f}%",
                 bg=bg, fg=GREEN if v >= 0 else RED, size=14)

    rect(slide, Inches(1.2), Inches(6.65), Inches(0.25), Inches(0.25),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "能源股 / CL=F", Inches(1.5), Inches(6.62),
            Inches(2.5), Inches(0.3), size=11, color=MID_BLUE)
    rect(slide, Inches(4.2), Inches(6.65), Inches(0.25), Inches(0.25),
         RGBColor(0xFF, 0xF0, 0xF0))
    textbox(slide, "國防股", Inches(4.5), Inches(6.62),
            Inches(1.5), Inches(0.3), size=11, color=RED)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Key Findings: Two-Hypothesis Debate
# ═══════════════════════════════════════════════════════════════════════════════
def slide_09_findings(prs):
    slide = blank_slide(prs)
    header_band(slide, "關鍵發現：國防股為何下跌？",
                "Key Findings  ·  Two Competing Hypotheses  ·  You Decide")

    # Observation banner
    rect(slide, Inches(0.4), Inches(1.4), Inches(12.53), Inches(0.72),
         RGBColor(0xFF, 0xE8, 0xE8))
    textbox(slide,
            "現象：戰爭爆發期間國防股 LMT / NOC / RTX 下跌 -9%～-11%，"
            "能源股同期上漲 +9%～+12% — 為什麼？",
            Inches(0.55), Inches(1.44), Inches(12.2), Inches(0.62),
            size=13, bold=True, color=RED, align=PP_ALIGN.CENTER)

    COL_A, COL_B = Inches(0.4), Inches(6.85)
    COL_W, BOX_H, TOP = Inches(6.15), Inches(4.3), Inches(2.22)

    # ── Hypothesis A ──────────────────────────────────────────────────────────
    rect(slide, COL_A, TOP, COL_W, BOX_H, RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, COL_A, TOP, COL_W, Inches(0.52), MID_BLUE)
    textbox(slide, "假說 A　美軍表現不如預期",
            COL_A + Inches(0.15), TOP + Inches(0.06),
            COL_W - Inches(0.3), Inches(0.42),
            size=14, bold=True, color=WHITE)
    bullet_rows(slide, [
        "伊朗成功短暫封鎖霍木茲海峽，\n美軍未能快速解除封鎖",
        "飛彈與無人機攻擊穿透美軍防空系統，\n顯示防空效能不足",
        "戰爭從 2/28 持續至 4/8 停火，長達 39 天，\n超出市場「閃電戰」預期",
        "美軍武器耗損高，補充採購雖增加\n但短期無法轉為獲利",
    ], COL_A + Inches(0.15), TOP + Inches(0.65),
       COL_W - Inches(0.2), Inches(0.83),
       size=11, color=DARK_BLUE, marker="▸", marker_color=MID_BLUE)
    textbox(slide,
            "→ 市場認為美國國防工業「能力受質疑」，\n   股價反映負面評估",
            COL_A + Inches(0.15), TOP + Inches(3.75),
            COL_W - Inches(0.3), Inches(0.48),
            size=11, bold=True, italic=True, color=MID_BLUE)

    # ── Hypothesis B ──────────────────────────────────────────────────────────
    rect(slide, COL_B, TOP, COL_W, BOX_H, RGBColor(0xFF, 0xF3, 0xE8))
    rect(slide, COL_B, TOP, COL_W, Inches(0.52), ORANGE)
    textbox(slide, "假說 B　買謠言、賣事實（Pre-priced）",
            COL_B + Inches(0.15), TOP + Inches(0.06),
            COL_W - Inches(0.3), Inches(0.42),
            size=14, bold=True, color=WHITE)
    bullet_rows(slide, [
        "LMT、NOC 在戰前期已大漲 32% / 24%，\n市場早已定價戰爭概念",
        "LMT 與 NOC 股價高點出現在 2/28 之前，\n宣戰當天即開始回落",
        "實際武器訂單從簽約到入帳獲利\n需要 3-5 年，無法反映短期股價",
        "短期衝突 (39 天) 不足以推動\n新一輪長期國防預算增加",
    ], COL_B + Inches(0.15), TOP + Inches(0.65),
       COL_W - Inches(0.2), Inches(0.83),
       size=11, color=DARK_BLUE, marker="▸", marker_color=ORANGE)
    textbox(slide,
            "→ 市場行為符合「預期已充分反映，\n   事件落地後獲利了結」的經典模式",
            COL_B + Inches(0.15), TOP + Inches(3.75),
            COL_W - Inches(0.3), Inches(0.48),
            size=11, bold=True, italic=True, color=ORANGE)

    # VS divider
    rect(slide, Inches(6.6), TOP + Inches(0.7),
         Inches(0.5), BOX_H - Inches(0.9), RGBColor(0xDD, 0xDD, 0xDD))
    textbox(slide, "VS", Inches(6.6), TOP + Inches(1.9),
            Inches(0.5), Inches(0.5),
            size=16, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # Bottom note
    rect(slide, Inches(0.4), Inches(6.65), Inches(12.53), Inches(0.68), LIGHT_GRAY)
    textbox(slide,
            "分析侷限：純價格分析無法區分兩種假說 — "
            "需結合基本面資料（訂單數據、武器耗損報告、國防預算修正）才能驗證。\n"
            "這正是 Spec-Driven Multi-Agent 分析的邊界："
            "Agent 能高效處理量化資料，但「為什麼」仍需人類判斷。",
            Inches(0.55), Inches(6.67), Inches(12.2), Inches(0.64),
            size=10, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — Reflection: AI 是工具，判斷是人的責任
# ═══════════════════════════════════════════════════════════════════════════════
def slide_10_reflection(prs):
    slide = blank_slide(prs)
    header_band(slide, "AI 是工具，判斷是人的責任",
                "Spec-Driven Multi-Agent 實作反思  ·  Lessons Learned")

    COL_W, COL_H, COL_TOP = Inches(6.0), Inches(4.55), Inches(1.42)

    # Left: AI can do
    rect(slide, Inches(0.4), COL_TOP, COL_W, COL_H, RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, Inches(0.4), COL_TOP, COL_W, Inches(0.52), MID_BLUE)
    textbox(slide, "AI Agent 能做的",
            Inches(0.55), COL_TOP + Inches(0.07), COL_W - Inches(0.3), Inches(0.4),
            size=15, bold=True, color=WHITE)
    bullet_rows(slide, [
        "收集大量數據（Yahoo Finance，7 支股票 × 104 天）",
        "自動計算漲跌幅、生成圖表與視覺化",
        "撰寫結構化報告，格式一致、可重現",
        "多 Agent 並行處理，效率遠高於人工",
        "全程 git 記錄，每步可追蹤、可審計",
    ], Inches(0.55), COL_TOP + Inches(0.65),
       COL_W - Inches(0.2), Inches(0.75),
       size=12, color=DARK_BLUE, marker="✓", marker_color=MID_BLUE)

    # Right: humans irreplaceable
    rect(slide, Inches(6.93), COL_TOP, COL_W, COL_H, RGBColor(0xFF, 0xF3, 0xE8))
    rect(slide, Inches(6.93), COL_TOP, COL_W, Inches(0.52), ORANGE)
    textbox(slide, "人類不可取代的",
            Inches(7.08), COL_TOP + Inches(0.07), COL_W - Inches(0.3), Inches(0.4),
            size=15, bold=True, color=WHITE)
    bullet_rows(slide, [
        "判斷數據背後的邏輯（假說 A vs 假說 B）",
        "理解地緣政治的複雜性與脈絡",
        "提出有意義的研究假說",
        "決定哪一種解釋更可信",
        "對結論承擔最後責任",
    ], Inches(7.08), COL_TOP + Inches(0.65),
       COL_W - Inches(0.2), Inches(0.75),
       size=12, color=DARK_BLUE, marker="★", marker_color=ORANGE)

    # Divider
    rect(slide, Inches(6.6), COL_TOP + Inches(0.6),
         Inches(0.33), COL_H - Inches(0.7), RGBColor(0xDD, 0xDD, 0xDD))
    textbox(slide, "+", Inches(6.6), COL_TOP + Inches(1.8),
            Inches(0.33), Inches(0.45),
            size=18, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # Conclusion bar
    conc_top = COL_TOP + COL_H + Inches(0.18)
    rect(slide, Inches(0.4), conc_top, Inches(12.53), Inches(1.55), DARK_BLUE)
    textbox(slide,
            "本次實作展示了 Spec-Driven Multi-Agent 的效率，"
            "但國防股假說 A vs 假說 B 的辯證說明：\n"
            "AI 能告訴你「發生了什麼」，"
            "但「為什麼」和「代表什麼意義」，"
            "仍需要人類的批判性思考。",
            Inches(0.6), conc_top + Inches(0.2),
            Inches(12.13), Inches(1.2),
            size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    prs = new_prs()

    slide_01_title(prs)
    slide_02_motivation(prs)
    slide_03_spec_driven(prs)
    slide_04_architecture(prs)
    slide_05_timeline(prs)
    slide_06_energy(prs)
    slide_07_defense(prs)
    slide_08_table(prs, "data/analysis_results.csv")
    slide_09_findings(prs)
    slide_10_reflection(prs)

    out = "reports/iran_war_stocks_v2.pptx"
    prs.save(out)
    print(f"Saved {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
