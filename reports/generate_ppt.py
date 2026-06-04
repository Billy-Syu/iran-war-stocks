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
GOLD       = RGBColor(0xFF, 0xC0, 0x00)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── low-level helpers ─────────────────────────────────────────────────────────

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
    s.fill.solid(); s.fill.fore_color.rgb = fill
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
    run.text = text
    run.font.size   = Pt(size)
    run.font.bold   = bold
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
    """items = list of strings; draws one row per item."""
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
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = title
    r.font.size = Pt(title_size); r.font.bold = True
    r.font.color.rgb = text_col
    if body:
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run(); r2.text = body
        r2.font.size = Pt(body_size); r2.font.color.rgb = text_col


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title
# ═══════════════════════════════════════════════════════════════════════════════
def slide_01_title(prs):
    slide = blank_slide(prs)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_BLUE)
    rect(slide, 0, 0, Inches(0.22), SLIDE_H, ORANGE)       # left accent

    # top decorative stripe
    rect(slide, Inches(0.22), Inches(0), SLIDE_W, Inches(0.07), MID_BLUE)

    textbox(slide,
            "美伊戰爭金融市場衝擊分析",
            Inches(0.6), Inches(1.7), Inches(12.1), Inches(1.1),
            size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    textbox(slide,
            "Spec-Driven Multi-Agent 實作報告",
            Inches(0.6), Inches(2.75), Inches(12.1), Inches(0.75),
            size=26, bold=False, color=ORANGE, align=PP_ALIGN.CENTER)

    textbox(slide,
            "US–Iran War  ·  Impact on Energy & Defense Stocks",
            Inches(0.6), Inches(3.55), Inches(12.1), Inches(0.5),
            size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    # divider
    rect(slide, Inches(3.5), Inches(4.25), Inches(6.3), Inches(0.04), MID_BLUE)

    textbox(slide,
            "生成式人工智慧導論  ·  Final Term Project  ·  2026-06-03",
            Inches(0.6), Inches(4.45), Inches(12.1), Inches(0.45),
            size=13, color=RGBColor(0x9D, 0xC3, 0xE6), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Research Topic & Motivation
# ═══════════════════════════════════════════════════════════════════════════════
def slide_02_motivation(prs):
    slide = blank_slide(prs)
    header_band(slide, "研究主題與動機", "Research Topic & Motivation")

    # left column — topic
    rect(slide, Inches(0.4), Inches(1.45), Inches(5.8), Inches(5.75), LIGHT_GRAY)
    textbox(slide, "研究主題", Inches(0.55), Inches(1.6), Inches(5.5), Inches(0.5),
            size=16, bold=True, color=DARK_BLUE)
    rect(slide, Inches(0.55), Inches(2.15), Inches(0.06), Inches(3.6), ORANGE)
    topics = [
        "2026 年 2 月 28 日美伊戰爭爆發，\n地緣政治風險急遽升高",
        "能源股與國防股被視為戰爭受益標的，\n但市場反應是否符合直覺？",
        "原油期貨 (CL=F) 如何領先或滯後\n於股票市場？",
        "三個時間段：戰前、爆發期、停火後\n各有何不同的市場邏輯？",
    ]
    bullet_rows(slide, topics, Inches(0.75), Inches(2.2), Inches(5.3),
                Inches(0.88), size=12, color=DARK_BLUE, marker="◆", marker_color=ORANGE)

    # right column — what we wanted to find
    rect(slide, Inches(6.6), Inches(1.45), Inches(6.45), Inches(5.75), RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "我們想回答的問題", Inches(6.75), Inches(1.6), Inches(6.1), Inches(0.5),
            size=16, bold=True, color=DARK_BLUE)
    rect(slide, Inches(6.75), Inches(2.15), Inches(0.06), Inches(3.6), MID_BLUE)
    questions = [
        "戰爭爆發時，能源股真的會上漲嗎？",
        "國防股在戰爭期間的表現是否符合\n「戰爭概念股」預期？",
        "原油價格是市場的領先指標嗎？",
        "停火後資金如何重新配置？",
    ]
    bullet_rows(slide, questions, Inches(6.95), Inches(2.2), Inches(5.9),
                Inches(0.88), size=12, color=DARK_BLUE, marker="？", marker_color=MID_BLUE)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — What is Spec-Driven
# ═══════════════════════════════════════════════════════════════════════════════
def slide_03_spec_driven(prs):
    slide = blank_slide(prs)
    header_band(slide, "什麼是 Spec-Driven 開發？",
                "Spec-Driven Development Applied to AI Agents")

    # concept definition box
    rect(slide, Inches(0.4), Inches(1.45), Inches(12.53), Inches(1.1),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide,
            "核心概念：先寫「規格書 (Spec)」，再讓 AI Agent 按規格執行 — "
            "人類定義 WHAT，Agent 決定 HOW",
            Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.85),
            size=14, bold=False, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # three-step flow
    steps = [
        ("1  定義規格", "明確寫出任務目標、\n輸入/輸出格式、品質標準\n(本研究：CSV 格式、時間區間、\n圖表樣式)", DARK_BLUE),
        ("2  Agent 執行", "Claude Code 解讀規格，\n自動分派子任務給\nAgent A / Agent B，\ngit commit 記錄每步驟", MID_BLUE),
        ("3  驗證與迭代", "人類審查輸出是否符合規格，\n不符則修改規格重跑 —\n而非修改 Agent 內部邏輯", ORANGE),
    ]

    bw = Inches(3.8)
    bh = Inches(3.3)
    by = Inches(2.75)
    for i, (title, body, color) in enumerate(steps):
        bx = Inches(0.4) + i * Inches(4.3)
        rect(slide, bx, by, bw, bh, color)
        textbox(slide, title, bx + Inches(0.15), by + Inches(0.15),
                bw - Inches(0.3), Inches(0.55),
                size=16, bold=True, color=WHITE)
        rect(slide, bx + Inches(0.15), by + Inches(0.72),
             bw - Inches(0.3), Inches(0.04), WHITE)
        textbox(slide, body, bx + Inches(0.15), by + Inches(0.85),
                bw - Inches(0.3), bh - Inches(1.0),
                size=12, color=WHITE, wrap=True)

        if i < 2:
            textbox(slide, "→", bx + bw + Inches(0.1), by + Inches(1.3),
                    Inches(0.4), Inches(0.6), size=24, bold=True,
                    color=MID_GRAY, align=PP_ALIGN.CENTER)

    # bottom note
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

    # git ribbon on right
    rect(slide, Inches(11.6), Inches(1.35), Inches(1.73), Inches(6.15),
         RGBColor(0xF0, 0xF8, 0xE8))
    textbox(slide, "git log", Inches(11.65), Inches(1.45), Inches(1.6), Inches(0.4),
            size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    commits = [
        "fetch_stocks.py",
        "raw_stocks.csv",
        "analysis.py",
        "analysis_results.csv",
        "plot_stocks.py",
        "energy/defense .png",
        "generate_report.py",
        "final_report.md",
        "generate_ppt.py",
        "iran_war_stocks.pptx",
    ]
    for i, c in enumerate(commits):
        y = Inches(1.95) + i * Inches(0.41)
        rect(slide, Inches(11.75), y + Inches(0.07),
             Inches(0.12), Inches(0.24), GREEN)
        textbox(slide, c, Inches(11.95), y, Inches(1.3), Inches(0.38),
                size=9, color=DARK_GRAY)

    # main flow area
    cx_orch = Inches(5.5)
    cy_orch = Inches(2.0)
    bw = Inches(2.8); bh = Inches(1.2)

    # orchestrator
    rect(slide, cx_orch, cy_orch, bw, bh, DARK_BLUE, WHITE)
    textbox(slide, "Claude Code", cx_orch, cy_orch + Inches(0.1), bw, Inches(0.5),
            size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "Orchestrator\n規格解讀 · 任務分派",
            cx_orch, cy_orch + Inches(0.55), bw, Inches(0.55),
            size=10, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    # user spec box (above)
    rect(slide, cx_orch + Inches(0.5), Inches(1.35), Inches(1.8), Inches(0.55),
         ORANGE)
    textbox(slide, "人類撰寫規格 (Spec)",
            cx_orch + Inches(0.5), Inches(1.35), Inches(1.8), Inches(0.55),
            size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # arrow spec→orch
    rect(slide, cx_orch + Inches(1.35), Inches(1.9), Inches(0.06), Inches(0.1), MID_GRAY)

    # arrows orch → agents
    mid_x = cx_orch + bw / 2
    rect(slide, mid_x - Inches(2.2), cy_orch + bh + Inches(0.0),
         Inches(0.06), Inches(0.7), MID_GRAY)
    rect(slide, mid_x + Inches(2.14), cy_orch + bh + Inches(0.0),
         Inches(0.06), Inches(0.7), MID_GRAY)

    # Agent A
    ax = Inches(2.2); ay = cy_orch + bh + Inches(0.7)
    rect(slide, ax, ay, bw, Inches(1.55), MID_BLUE, WHITE)
    textbox(slide, "Agent A", ax, ay + Inches(0.08), bw, Inches(0.45),
            size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "數據抓取 & 圖表生成\nyfinance · matplotlib\n→ raw_stocks.csv · .png",
            ax, ay + Inches(0.5), bw, Inches(1.0),
            size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # Agent B
    bx2 = Inches(8.8); by2 = ay
    rect(slide, bx2, by2, bw, Inches(1.55), RGBColor(0x10, 0x7C, 0x10), WHITE)
    textbox(slide, "GPT-4o  Agent B", bx2, by2 + Inches(0.08), bw, Inches(0.45),
            size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(slide, "報告撰寫 & 洞察分析\nanalysis.py · generate_report.py\n→ final_report.md · .pptx",
            bx2, by2 + Inches(0.5), bw, Inches(1.0),
            size=10, color=WHITE, align=PP_ALIGN.CENTER)

    # output arrow → merged result
    out_y = ay + Inches(1.55) + Inches(0.3)
    rect(slide, ax + bw / 2 - Inches(0.03), ay + Inches(1.55),
         Inches(0.06), Inches(0.3), MID_GRAY)
    rect(slide, bx2 + bw / 2 - Inches(0.03), ay + Inches(1.55),
         Inches(0.06), Inches(0.3), MID_GRAY)

    rect(slide, Inches(3.5), out_y, Inches(6.8), Inches(0.75),
         RGBColor(0xF2, 0xF2, 0xF2))
    textbox(slide, "iran_war_stocks.pptx  ·  final_report.md",
            Inches(3.5), out_y, Inches(6.8), Inches(0.75),
            size=12, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # data source
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
    header_band(slide, "事件背景：美伊戰爭時間軸",
                "US–Iran War Timeline  ·  2026")

    events = [
        ("2026-01-01", "觀察期開始\nBaseline", MID_BLUE),
        ("2026-02-28", "戰爭爆發\nWar Outbreak", RED),
        ("2026-04-08", "停火協議\nCeasefire", ORANGE),
        ("2026-05-05", "軍事行動結束\nOp. Ended", GREEN),
        ("2026-06-03", "資料截止\nData End", DARK_GRAY),
    ]

    bar_top  = Inches(3.9)
    bar_left = Inches(0.9)
    bar_w    = Inches(11.5)
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
                Inches(2.0), Inches(0.9), size=11,
                color=color, align=PP_ALIGN.CENTER)

    # period bracket labels
    def period_span(label, x1, x2, color):
        mid = (x1 + x2) / 2
        span_y = Inches(2.45)
        rect(slide, x1, span_y + Inches(0.42), x2 - x1, Inches(0.05), color)
        rect(slide, x1, span_y + Inches(0.18), Inches(0.05), Inches(0.24), color)
        rect(slide, x2 - Inches(0.05), span_y + Inches(0.18), Inches(0.05), Inches(0.24), color)
        textbox(slide, label, mid - Inches(1.3), span_y, Inches(2.6), Inches(0.4),
                size=11, italic=True, bold=True, color=color, align=PP_ALIGN.CENTER)

    period_span("戰前期  Before War",        lx[0], lx[1], MID_BLUE)
    period_span("戰爭爆發期  War Outbreak",  lx[1], lx[2], RED)
    period_span("停火後  After Ceasefire",   lx[2], lx[4], GREEN)

    # context bullets
    context = [
        "伊朗核武計畫談判破裂，美軍實施精準打擊德黑蘭核設施",
        "霍木茲海峽短暫封鎖，全球約 20% 石油運輸受阻",
        "4 月 8 日在聯合國斡旋下達成臨時停火，5 月 5 日正式結束敵對行動",
    ]
    bullet_rows(slide, context, Inches(0.5), Inches(5.6), Inches(12.5),
                Inches(0.55), size=11, color=DARK_BLUE,
                marker="•", marker_color=RED)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Energy Stocks Chart + findings
# ═══════════════════════════════════════════════════════════════════════════════
def slide_06_energy(prs):
    slide = blank_slide(prs)
    header_band(slide, "能源股走勢分析",
                "Energy Stocks  ·  XOM · CVX · COP  +  CL=F (right axis)")

    slide.shapes.add_picture("charts/energy_stocks.png",
                             Inches(0.3), Inches(1.38), Inches(9.3), Inches(5.95))

    # findings panel
    rect(slide, Inches(9.75), Inches(1.38), Inches(3.3), Inches(5.95),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "關鍵發現", Inches(9.9), Inches(1.5), Inches(3.0), Inches(0.45),
            size=14, bold=True, color=DARK_BLUE)
    rect(slide, Inches(9.9), Inches(1.95), Inches(2.9), Inches(0.04), MID_BLUE)

    findings = [
        ("戰前期\n+16%～+25%", "市場提前反應供應\n中斷預期，三股\n同步上揚", MID_BLUE),
        ("爆發期\n+9%～+12%", "原油飆升 +42%\n帶動能源股續漲\n（但漲幅縮窄）", ORANGE),
        ("停火後\n-2%～-6%", "供應恢復預期使\n能源股獲利了結\n回落", RED),
    ]

    fy = Inches(2.1)
    for title, body, color in findings:
        rect(slide, Inches(9.9), fy, Inches(2.9), Inches(0.04), color)
        textbox(slide, title, Inches(9.9), fy + Inches(0.08),
                Inches(2.9), Inches(0.5),
                size=12, bold=True, color=color)
        textbox(slide, body, Inches(9.9), fy + Inches(0.58),
                Inches(2.9), Inches(0.75),
                size=11, color=DARK_BLUE)
        fy += Inches(1.55)

    textbox(slide,
            "原油領先能源股\n約 1-2 個交易日",
            Inches(9.9), Inches(6.55), Inches(2.9), Inches(0.55),
            size=11, italic=True, color=DARK_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Defense Stocks Chart + findings
# ═══════════════════════════════════════════════════════════════════════════════
def slide_07_defense(prs):
    slide = blank_slide(prs)
    header_band(slide, "國防股走勢分析",
                "Defense Stocks  ·  LMT · RTX · NOC  +  CL=F (right axis)")

    slide.shapes.add_picture("charts/defense_stocks.png",
                             Inches(0.3), Inches(1.38), Inches(9.3), Inches(5.95))

    rect(slide, Inches(9.75), Inches(1.38), Inches(3.3), Inches(5.95),
         RGBColor(0xFF, 0xF0, 0xF0))
    textbox(slide, "關鍵發現", Inches(9.9), Inches(1.5), Inches(3.0), Inches(0.45),
            size=14, bold=True, color=DARK_BLUE)
    rect(slide, Inches(9.9), Inches(1.95), Inches(2.9), Inches(0.04), RED)

    findings = [
        ("戰前期\n+8%～+32%", "市場預期戰爭\n提前買入，尤其\nLMT 大漲 32%", MID_BLUE),
        ("爆發期\n-9%～-11%", "出乎意料的下跌！\n市場判斷短期衝突\n不增加國防預算", RED),
        ("停火後\n-14%～-22%", "軍事需求預期降溫，\n資金轉向能源股，\n持續回落", RED),
    ]

    fy = Inches(2.1)
    for title, body, color in findings:
        rect(slide, Inches(9.9), fy, Inches(2.9), Inches(0.04), color)
        textbox(slide, title, Inches(9.9), fy + Inches(0.08),
                Inches(2.9), Inches(0.5),
                size=12, bold=True, color=color)
        textbox(slide, body, Inches(9.9), fy + Inches(0.58),
                Inches(2.9), Inches(0.75),
                size=11, color=DARK_BLUE)
        fy += Inches(1.55)

    # surprise callout
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

    ENERGY_TICKERS  = {"XOM", "CVX", "COP", "CL=F"}
    col_labels = ["股票代碼\nTicker",
                  "戰前期\nBefore War",
                  "戰爭爆發期\nWar Outbreak",
                  "停火後\nAfter Ceasefire"]

    rows = len(df) + 2   # header + section sub-headers embedded
    tbl = slide.shapes.add_table(
        len(df) + 1, 4,
        Inches(1.2), Inches(1.5),
        Inches(10.9), Inches(5.75)
    ).table

    col_widths = [Inches(2.2), Inches(2.9), Inches(2.9), Inches(2.9)]
    for ci, w in enumerate(col_widths):
        tbl.columns[ci].width = w

    def cell(r, c, text, bold=False, size=13, bg=None, fg=WHITE,
             align=PP_ALIGN.CENTER):
        cl = tbl.cell(r, c)
        cl.text = ""
        tf = cl.text_frame
        tf.paragraphs[0].alignment = align
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(size); run.font.bold = bold
        run.font.color.rgb = fg
        if bg:
            cl.fill.solid(); cl.fill.fore_color.rgb = bg

    # header row
    for ci, h in enumerate(col_labels):
        cell(0, ci, h, bold=True, bg=DARK_BLUE, size=12)

    # data rows — group energy first, then defense
    energy_rows = df[df["ticker"].isin(ENERGY_TICKERS)]
    defense_rows = df[~df["ticker"].isin(ENERGY_TICKERS)]
    ordered = pd.concat([energy_rows, defense_rows])

    for ri, (_, row) in enumerate(ordered.iterrows()):
        ticker = row["ticker"]
        vals   = [row["before_war"], row["war_outbreak"], row["after_ceasefire"]]
        is_energy = ticker in ENERGY_TICKERS

        row_bg = RGBColor(0xE8, 0xF0, 0xF8) if is_energy else RGBColor(0xFF, 0xF0, 0xF0)
        label_col = MID_BLUE if is_energy else RED

        cell(ri + 1, 0, ticker, bold=True, bg=row_bg, fg=label_col, size=14)
        for ci, v in enumerate(vals):
            val_color = GREEN if v >= 0 else RED
            cell(ri + 1, ci + 1, f"{v:+.2f}%", bg=row_bg, fg=val_color, size=14)

    # legend
    rect(slide, Inches(1.2), Inches(6.65), Inches(0.25), Inches(0.25),
         RGBColor(0xE8, 0xF0, 0xF8))
    textbox(slide, "能源股 / CL=F", Inches(1.5), Inches(6.62),
            Inches(2.5), Inches(0.3), size=11, color=MID_BLUE)
    rect(slide, Inches(4.2), Inches(6.65), Inches(0.25), Inches(0.25),
         RGBColor(0xFF, 0xF0, 0xF0))
    textbox(slide, "國防股", Inches(4.5), Inches(6.62),
            Inches(1.5), Inches(0.3), size=11, color=RED)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Key Findings & Conclusions
# ═══════════════════════════════════════════════════════════════════════════════
def slide_09_findings(prs):
    slide = blank_slide(prs)
    header_band(slide, "\u95dc\u9375\u767c\u73fe\uff1a\u570b\u9632\u80a1\u70ba\u4f55\u4e0b\u8dcc\uff1f",
                "Key Findings  \u00b7  Two Competing Hypotheses  \u00b7  You Decide")

    # shared observation banner
    rect(slide, Inches(0.4), Inches(1.4), Inches(12.53), Inches(0.72),
         RGBColor(0xFF, 0xE8, 0xE8))
    textbox(slide,
            "\u73fe\u8c61\uff1a\u6230\u722d\u7206\u767c\u671f\u9593\u570b\u9632\u80a1 LMT / NOC / RTX \u4e0b\u8dcc -9%\uff5e-11%\uff0c\u80fd\u6e90\u80a1\u540c\u671f\u4e0a\u6f32 +9%\uff5e+12% \u2014 \u70ba\u4ec0\u9ebc\uff1f",
            Inches(0.55), Inches(1.44), Inches(12.2), Inches(0.62),
            size=13, bold=True, color=RED, align=PP_ALIGN.CENTER)

    COL_A = Inches(0.4)
    COL_B = Inches(6.85)
    COL_W = Inches(6.15)
    TOP   = Inches(2.22)
    BOX_H = Inches(4.3)

    # ── Hypothesis A (left) ───────────────────────────────────────────────────
    rect(slide, COL_A, TOP, COL_W, BOX_H, RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, COL_A, TOP, COL_W, Inches(0.52), MID_BLUE)
    textbox(slide, "\u5047\u8aaa A\u3000\u7f8e\u8ecd\u8868\u73fe\u4e0d\u5982\u9810\u671f",
            COL_A + Inches(0.15), TOP + Inches(0.06), COL_W - Inches(0.3), Inches(0.42),
            size=14, bold=True, color=WHITE)
    evidence_a = [
        "\u4f0a\u6717\u6210\u529f\u77ed\u66ab\u5c01\u9396\u970d\u6728\u8332\u6d77\u5ce1\uff0c\u7f8e\u8ecd\u672a\u80fd\u5feb\u901f\u89e3\u9664\u5c01\u9396",
        "\u98db\u5f48\u8207\u7121\u4eba\u6a5f\u653b\u64ca\u7a7f\u900f\u7f8e\u8ecd\u9632\u7a7a\u7cfb\u7d71\uff0c\u986f\u793a\u9632\u7a7a\u6548\u80fd\u4e0d\u8db3",
        "\u6230\u722d\u5f9e 2/28 \u6301\u7e8c\u81f3 4/8 \u505c\u706b\uff0c\u9577\u9054 39 \u5929\uff0c\u8d85\u51fa\u5e02\u5834\u300c\u9583\u96fb\u6230\u300d\u9810\u671f",
        "\u7f8e\u8ecd\u6b66\u5668\u8017\u640f\u9ad8\uff0c\u88dc\u5145\u63a1\u8cfc\u96d6\u5897\u52a0\u4f46\u77ed\u671f\u7121\u6cd5\u8f49\u70ba\u7372\u5229",
    ]
    bullet_rows(slide, evidence_a,
                COL_A + Inches(0.15), TOP + Inches(0.65),
                COL_W - Inches(0.2), Inches(0.83),
                size=11, color=DARK_BLUE, marker="\u25b8", marker_color=MID_BLUE)
    textbox(slide,
            "\u2192 \u5e02\u5834\u8a8d\u70ba\u7f8e\u570b\u570b\u9632\u5de5\u696d\u300c\u80fd\u529b\u53d7\u8cea\u7591\u300d\uff0c\u80a1\u50f9\u53cd\u6620\u8ca0\u9762\u8a55\u4f30",
            COL_A + Inches(0.15), TOP + Inches(3.75),
            COL_W - Inches(0.3), Inches(0.48),
            size=11, bold=True, italic=True, color=MID_BLUE)

    # ── Hypothesis B (right) ──────────────────────────────────────────────────
    rect(slide, COL_B, TOP, COL_W, BOX_H, RGBColor(0xFF, 0xF3, 0xE8))
    rect(slide, COL_B, TOP, COL_W, Inches(0.52), ORANGE)
    textbox(slide, "\u5047\u8aaa B\u3000\u8cb7\u8b20\u8a00\u3001\u8ce3\u4e8b\u5be6\uff08Pre-priced\uff09",
            COL_B + Inches(0.15), TOP + Inches(0.06), COL_W - Inches(0.3), Inches(0.42),
            size=14, bold=True, color=WHITE)
    evidence_b = [
        "LMT\u3001NOC \u5728\u6230\u524d\u671f\u5df2\u5927\u6f32 32% / 24%\uff0c\u5e02\u5834\u65e9\u5df2\u5b9a\u50f9\u6230\u722d\u6982\u5ff5",
        "LMT \u8207 NOC \u80a1\u50f9\u9ad8\u9ede\u51fa\u73fe\u5728 2/28 \u4e4b\u524d\uff0c\u5ba3\u6230\u7576\u5929\u5373\u958b\u59cb\u56de\u843d",
        "\u5be6\u969b\u6b66\u5668\u8a02\u55ae\u5f9e\u7c3d\u7d04\u5230\u5165\u5e33\u7372\u5229\u9700\u8981 3-5 \u5e74\uff0c\u7121\u6cd5\u53cd\u6620\u77ed\u671f\u80a1\u50f9",
        "\u77ed\u671f\u885d\u7a81 (39 \u5929) \u4e0d\u8db3\u4ee5\u63a8\u52d5\u65b0\u4e00\u8f2a\u9577\u671f\u570b\u9632\u9810\u7b97\u589e\u52a0",
    ]
    bullet_rows(slide, evidence_b,
                COL_B + Inches(0.15), TOP + Inches(0.65),
                COL_W - Inches(0.2), Inches(0.83),
                size=11, color=DARK_BLUE, marker="\u25b8", marker_color=ORANGE)
    textbox(slide,
            "\u2192 \u5e02\u5834\u884c\u70ba\u7b26\u5408\u300c\u9810\u671f\u5df2\u5145\u5206\u53cd\u6620\uff0c\u4e8b\u4ef6\u843d\u5730\u5f8c\u7372\u5229\u4e86\u7d50\u300d\u7684\u7d93\u5178\u6a21\u5f0f",
            COL_B + Inches(0.15), TOP + Inches(3.75),
            COL_W - Inches(0.3), Inches(0.48),
            size=11, bold=True, italic=True, color=ORANGE)

    # VS divider
    rect(slide, Inches(6.6), TOP + Inches(0.7), Inches(0.5), BOX_H - Inches(0.9),
         RGBColor(0xDD, 0xDD, 0xDD))
    textbox(slide, "VS", Inches(6.6), TOP + Inches(1.9), Inches(0.5), Inches(0.5),
            size=16, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # bottom note
    rect(slide, Inches(0.4), Inches(6.65), Inches(12.53), Inches(0.68),
         RGBColor(0xF2, 0xF2, 0xF2))
    textbox(slide,
            "\u5206\u6790\u5c40\u9650\uff1a\u7d14\u50f9\u683c\u5206\u6790\u7121\u6cd5\u5340\u5206\u5169\u7a2e\u5047\u8aaa \u2014 \u9700\u7d50\u5408\u57fa\u672c\u9762\u8cc7\u6599\uff08\u8a02\u55ae\u6578\u64da\u3001\u6b66\u5668\u8017\u640f\u5831\u544a\u3001\u570b\u9632\u9810\u7b97\u4fee\u6b63\uff09\u624d\u80fd\u9a57\u8b49\u3002\n"
            "\u9019\u6b63\u662f Spec-Driven Multi-Agent \u5206\u6790\u7684\u908a\u754c\uff1aAgent \u80fd\u9ad8\u6548\u8655\u7406\u91cf\u5316\u8cc7\u6599\uff0c\u4f46\u300c\u70ba\u4ec0\u9ebc\u300d\u4ecd\u9700\u4eba\u985e\u5224\u65b7\u3002",
            Inches(0.55), Inches(6.67), Inches(12.2), Inches(0.64),
            size=10, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)


def slide_10_reflection(prs):
    slide = blank_slide(prs)
    header_band(slide, "AI 是工具，判斷是人的責任",
                "Spec-Driven Multi-Agent 實作反思  ·  Lessons Learned")

    COL_W = Inches(6.0)
    COL_H = Inches(4.55)
    COL_TOP = Inches(1.42)

    # ── Left column: AI can do ────────────────────────────────────────────────
    rect(slide, Inches(0.4), COL_TOP, COL_W, COL_H, RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, Inches(0.4), COL_TOP, COL_W, Inches(0.52), MID_BLUE)
    textbox(slide, "AI Agent 能做的",
            Inches(0.55), COL_TOP + Inches(0.07), COL_W - Inches(0.3), Inches(0.4),
            size=15, bold=True, color=WHITE)

    ai_items = [
        "收集大量數據（Yahoo Finance，7 支股票 \xd7 104 天）",
        "自動計算漲跌幅、生成圖表與視覚化",
        "撰寫結構化報告，格式一致、可重現",
        "多 Agent 並行處理，效率遠高於人工",
        "全程 git 記錄，每步可追湯、可審計",
    ]
    bullet_rows(slide, ai_items,
                Inches(0.55), COL_TOP + Inches(0.65),
                COL_W - Inches(0.2), Inches(0.75),
                size=12, color=DARK_BLUE, marker="✓", marker_color=MID_BLUE)

    # ── Right column: humans irreplaceable ────────────────────────────────────
    rect(slide, Inches(6.93), COL_TOP, COL_W, COL_H, RGBColor(0xFF, 0xF3, 0xE8))
    rect(slide, Inches(6.93), COL_TOP, COL_W, Inches(0.52), ORANGE)
    textbox(slide, "人類不可取代的",
            Inches(7.08), COL_TOP + Inches(0.07), COL_W - Inches(0.3), Inches(0.4),
            size=15, bold=True, color=WHITE)

    human_items = [
        "判斷數據背後的邏輯（假說 A vs 假說 B）",
        "理解地緣政治的複雜性與胈絡",
        "提出有意義的研究假說",
        "决定哪一種解釋更可信",
        "對結論承擔最後責任",
    ]
    bullet_rows(slide, human_items,
                Inches(7.08), COL_TOP + Inches(0.65),
                COL_W - Inches(0.2), Inches(0.75),
                size=12, color=DARK_BLUE, marker="★", marker_color=ORANGE)

    # ── Divider + label ───────────────────────────────────────────────────────
    rect(slide, Inches(6.6), COL_TOP + Inches(0.6),
         Inches(0.33), COL_H - Inches(0.7), RGBColor(0xDD, 0xDD, 0xDD))
    textbox(slide, "+", Inches(6.6), COL_TOP + Inches(1.8),
            Inches(0.33), Inches(0.45),
            size=18, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # ── Bottom conclusion ─────────────────────────────────────────────────────
    CONC_TOP = COL_TOP + COL_H + Inches(0.18)
    rect(slide, Inches(0.4), CONC_TOP, Inches(12.53), Inches(1.55), DARK_BLUE)
    textbox(slide,
            "本次實作展示了 Spec-Driven Multi-Agent 的效率，"
            "但國防股假說 A vs 假說 B 的辩證說明：\n"
            "AI 能告訴你「發生了什麼」，"
            "但「為什麼」和「代表什麼意義」，"
            "仍需要人類的批判性思考。",
            Inches(0.6), CONC_TOP + Inches(0.2),
            Inches(12.13), Inches(1.2),
            size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


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

    out = "reports/iran_war_stocks.pptx"
    prs.save(out)
    print(f"Saved {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
