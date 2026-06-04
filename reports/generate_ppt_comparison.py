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
SAND       = RGBColor(0xFF, 0xF3, 0xE0)   # Gulf War warm tone
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
MID_GRAY   = RGBColor(0xCC, 0xCC, 0xCC)
DARK_GRAY  = RGBColor(0x70, 0x70, 0x70)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
RED        = RGBColor(0xC0, 0x00, 0x00)
GREEN      = RGBColor(0x37, 0x86, 0x30)
TEAL       = RGBColor(0x00, 0x70, 0x7F)   # 1990 accent
PURPLE     = RGBColor(0x70, 0x30, 0xA0)   # 2026 accent

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
    run.text        = text
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

def war_badge(slide, label, year, left, top, w=Inches(2.8), h=Inches(0.5), color=TEAL):
    rect(slide, left, top, w, h, color)
    textbox(slide, f"{label}  {year}",
            left, top, w, h,
            size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def bullet_rows(slide, items, left, top, w, row_h,
                size=12, color=DARK_BLUE, marker="▸", marker_color=None):
    mc = marker_color or MID_BLUE
    for i, text in enumerate(items):
        y = top + i * row_h
        textbox(slide, marker, left, y, Inches(0.3), row_h,
                size=size, bold=True, color=mc)
        textbox(slide, text, left + Inches(0.3), y, w - Inches(0.3), row_h,
                size=size, color=color)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title
# ═══════════════════════════════════════════════════════════════════════════════
def slide_01_title(prs):
    slide = blank_slide(prs)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_BLUE)

    # two accent bars — one per war
    rect(slide, 0, 0, Inches(0.22), SLIDE_H, TEAL)
    rect(slide, SLIDE_W - Inches(0.22), 0, Inches(0.22), SLIDE_H, PURPLE)

    rect(slide, Inches(0.22), 0, SLIDE_W - Inches(0.44), Inches(0.06), MID_BLUE)

    textbox(slide, "戰爭與金融市場",
            Inches(0.5), Inches(1.4), Inches(12.3), Inches(1.0),
            size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    textbox(slide, "1990 波灣戰爭  vs  2026 美伊戰爭",
            Inches(0.5), Inches(2.4), Inches(12.3), Inches(0.75),
            size=28, bold=False, color=ORANGE, align=PP_ALIGN.CENTER)

    textbox(slide, "Gulf War vs US–Iran War  ·  Energy & Defense Stocks  ·  35 Years Apart",
            Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.5),
            size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    rect(slide, Inches(3.5), Inches(3.9), Inches(6.3), Inches(0.04), MID_BLUE)

    # war badges
    war_badge(slide, "波灣戰爭", "Aug 1990 – Feb 1991",
              Inches(2.3), Inches(4.15), Inches(3.6), Inches(0.55), TEAL)
    textbox(slide, "vs", Inches(6.15), Inches(4.15), Inches(1.0), Inches(0.55),
            size=18, bold=True, color=MID_GRAY, align=PP_ALIGN.CENTER)
    war_badge(slide, "美伊戰爭", "Feb 2026 – May 2026",
              Inches(7.4), Inches(4.15), Inches(3.6), Inches(0.55), PURPLE)

    textbox(slide, "生成式人工智慧導論  ·  Final Term Project  ·  2026-06-04",
            Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.45),
            size=12, color=RGBColor(0x9D, 0xC3, 0xE6), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Research Targets
# ═══════════════════════════════════════════════════════════════════════════════
def slide_02_targets(prs):
    slide = blank_slide(prs)
    header_band(slide, "研究標的",
                "Research Targets  ·  6 Stocks  +  2 Oil Data Sources")

    # ── Energy section ────────────────────────────────────────────────────────
    rect(slide, Inches(0.4), Inches(1.42), Inches(5.95), Inches(2.5),
         RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, Inches(0.4), Inches(1.42), Inches(5.95), Inches(0.45), MID_BLUE)
    textbox(slide, "⚡  能源股  Energy Stocks",
            Inches(0.55), Inches(1.44), Inches(5.7), Inches(0.4),
            size=13, bold=True, color=WHITE)

    energy_stocks = [
        ("XOM", "ExxonMobil",      "全球最大上市石油公司，業務涵蓋勘探、煉油、化工"),
        ("CVX", "Chevron",         "美國第二大石油公司，主要生產地橫跨美洲與中東"),
        ("COP", "ConocoPhillips",  "美國最大獨立勘探生產商，不含煉油下游業務"),
    ]
    for i, (tkr, name, desc) in enumerate(energy_stocks):
        y = Inches(1.97) + i * Inches(0.62)
        rect(slide, Inches(0.5), y + Inches(0.06), Inches(0.55), Inches(0.42), MID_BLUE)
        textbox(slide, tkr, Inches(0.5), y + Inches(0.06), Inches(0.55), Inches(0.42),
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, name, Inches(1.15), y, Inches(1.5), Inches(0.38),
                size=11, bold=True, color=DARK_BLUE)
        textbox(slide, desc, Inches(1.15), y + Inches(0.36), Inches(4.95), Inches(0.3),
                size=9, color=DARK_GRAY)

    # ── Defense section ───────────────────────────────────────────────────────
    rect(slide, Inches(0.4), Inches(4.1), Inches(5.95), Inches(2.5),
         RGBColor(0xFF, 0xF0, 0xF0))
    rect(slide, Inches(0.4), Inches(4.1), Inches(5.95), Inches(0.45), RED)
    textbox(slide, "🛡  國防股  Defense Stocks",
            Inches(0.55), Inches(4.12), Inches(5.7), Inches(0.4),
            size=13, bold=True, color=WHITE)

    defense_stocks = [
        ("LMT", "Lockheed Martin",  "全球最大國防承包商，F-35、飛彈防禦系統"),
        ("RTX", "Raytheon Tech.",   "愛國者飛彈、導彈系統，2020 年由雷神與 UTC 合併"),
        ("NOC", "Northrop Grumman", "B-21 隱形轟炸機、太空與電子戰系統"),
    ]
    for i, (tkr, name, desc) in enumerate(defense_stocks):
        y = Inches(4.65) + i * Inches(0.62)
        rect(slide, Inches(0.5), y + Inches(0.06), Inches(0.55), Inches(0.42), RED)
        textbox(slide, tkr, Inches(0.5), y + Inches(0.06), Inches(0.55), Inches(0.42),
                size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        textbox(slide, name, Inches(1.15), y, Inches(1.6), Inches(0.38),
                size=11, bold=True, color=DARK_BLUE)
        textbox(slide, desc, Inches(1.15), y + Inches(0.36), Inches(4.95), Inches(0.3),
                size=9, color=DARK_GRAY)

    # ── Oil data sources (right column) ───────────────────────────────────────
    rect(slide, Inches(6.6), Inches(1.42), Inches(6.45), Inches(5.18),
         RGBColor(0xFF, 0xF8, 0xEC))
    rect(slide, Inches(6.6), Inches(1.42), Inches(6.45), Inches(0.45), ORANGE)
    textbox(slide, "🛢  原油數據來源  Oil Data Sources",
            Inches(6.75), Inches(1.44), Inches(6.2), Inches(0.4),
            size=13, bold=True, color=WHITE)

    oil_sources = [
        (TEAL,   "NYMEX  CL=F",
         "波灣戰爭（1990）數據來源",
         "FRED DCOILWTICO — WTI 每日現貨價",
         "美國能源資訊局 (EIA) 授權，\n免費公開，最早可追溯至 1986 年",
         "EIA  PET_PRI_SPT_S1_D.xls"),
        (PURPLE, "NYMEX  CL=F",
         "美伊戰爭（2026）數據來源",
         "Yahoo Finance — CL=F 連續期貨",
         "yfinance 套件抓取，\n數據可追溯至 ~2000 年後",
         "yfinance  CL=F"),
    ]
    for i, (accent, title, war_label, series, note, source) in enumerate(oil_sources):
        by = Inches(2.05) + i * Inches(2.4)
        rect(slide, Inches(6.7), by, Inches(0.08), Inches(2.15), accent)
        textbox(slide, war_label, Inches(6.9), by,
                Inches(6.0), Inches(0.38), size=10, bold=True, color=accent)
        textbox(slide, series, Inches(6.9), by + Inches(0.38),
                Inches(6.0), Inches(0.42), size=14, bold=True, color=DARK_BLUE)
        textbox(slide, note, Inches(6.9), by + Inches(0.82),
                Inches(6.0), Inches(0.65), size=10, color=DARK_GRAY)
        rect(slide, Inches(6.9), by + Inches(1.55), Inches(5.8), Inches(0.38),
             RGBColor(0xEE, 0xEE, 0xEE))
        textbox(slide, f"來源：{source}", Inches(7.0), by + Inches(1.57),
                Inches(5.6), Inches(0.35), size=9, color=DARK_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Two Wars Background Comparison
# ═══════════════════════════════════════════════════════════════════════════════
def slide_03_background(prs):
    slide = blank_slide(prs)
    header_band(slide, "兩場戰爭背景比較",
                "Background Comparison  ·  Gulf War 1990  vs  US–Iran War 2026")

    COL_H  = Inches(1.4)
    LEFT_X = Inches(0.4)
    MID_X  = Inches(4.55)
    RIG_X  = Inches(8.95)
    COL_W  = Inches(4.0)

    rows = [
        ("",        "項目",             "1990 波灣戰爭",                  "2026 美伊戰爭"),
        (TEAL,      "導火線",           "伊拉克入侵科威特\n（1990-08-02）",
                                        "美伊核武談判破裂，\n美軍精準打擊德黑蘭\n（2026-02-28）"),
        (PURPLE,    "主要交戰方",       "美國主導多國聯軍 vs 伊拉克",
                                        "美國 vs 伊朗"),
        (ORANGE,    "戰事持續",         "~7 個月（沙漠盾牌+沙漠風暴）\n停火：1991-02-28",
                                        "39 天\n停火：2026-04-08"),
        (MID_BLUE,  "霍木茲海峽",       "未受直接威脅",
                                        "短暫封鎖，全球 20%\n石油運輸受阻"),
        (GREEN,     "石油市場衝擊",     "WTI 從 $17 飆至 $41\n（+141%，沙漠風暴前後回落）",
                                        "WTI 從 $57 飆至 $94\n（+65%，停火後維持高位）"),
        (RED,       "國防股市場預期",   "突發事件，市場無預期\n→ 宣戰後國防股大漲",
                                        "地緣緊張已醞釀，市場提前定價\n→ 宣戰後國防股反跌"),
    ]

    row_ys = [Inches(1.35) + i * Inches(0.88) for i in range(len(rows))]

    for i, (color, label, gulf, iran) in enumerate(rows):
        y  = row_ys[i]
        rh = Inches(0.83) if i > 0 else Inches(0.42)

        header_row = (i == 0)
        bg = DARK_BLUE if header_row else (RGBColor(0xF5, 0xF5, 0xF5) if i % 2 == 0
                                           else WHITE)

        # label column
        rect(slide, LEFT_X, y, MID_X - LEFT_X - Inches(0.05), rh,
             DARK_BLUE if header_row else (color if not header_row else DARK_BLUE),
             None)
        textbox(slide, label,
                LEFT_X + Inches(0.1), y + Inches(0.05),
                MID_X - LEFT_X - Inches(0.25), rh - Inches(0.1),
                size=11 if not header_row else 12,
                bold=True, color=WHITE if not header_row else WHITE,
                align=PP_ALIGN.CENTER)

        # Gulf War column
        rect(slide, MID_X, y, COL_W, rh,
             DARK_BLUE if header_row else RGBColor(0xE8, 0xF4, 0xF0))
        textbox(slide, gulf,
                MID_X + Inches(0.1), y + Inches(0.05),
                COL_W - Inches(0.2), rh - Inches(0.1),
                size=10 if not header_row else 12,
                bold=header_row,
                color=WHITE if header_row else DARK_BLUE)

        # Iran War column
        rect(slide, RIG_X, y, COL_W, rh,
             DARK_BLUE if header_row else RGBColor(0xF0, 0xEA, 0xF8))
        textbox(slide, iran,
                RIG_X + Inches(0.1), y + Inches(0.05),
                COL_W - Inches(0.2), rh - Inches(0.1),
                size=10 if not header_row else 12,
                bold=header_row,
                color=WHITE if header_row else DARK_BLUE)

    # column badges
    war_badge(slide, "波灣戰爭", "1990–1991",
              MID_X, Inches(1.35), COL_W, Inches(0.42), TEAL)
    war_badge(slide, "美伊戰爭", "2026",
              RIG_X, Inches(1.35), COL_W, Inches(0.42), PURPLE)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Energy Stocks Chart Comparison
# ═══════════════════════════════════════════════════════════════════════════════
def slide_04_energy_charts(prs):
    slide = blank_slide(prs)
    header_band(slide, "能源股走勢對比",
                "Energy Stocks  ·  XOM · CVX · COP  ·  Split-adjusted closing prices + crude oil (right axis)")

    # Left badge
    war_badge(slide, "波灣戰爭  Gulf War",  "Jun 1990 – Apr 1991",
              Inches(0.4), Inches(1.38), Inches(6.2), Inches(0.42), TEAL)
    slide.shapes.add_picture("charts/energy_stocks_gulf.png",
                             Inches(0.4), Inches(1.85), Inches(6.2), Inches(5.45))

    # Right badge
    war_badge(slide, "美伊戰爭  US–Iran War", "Jan 2026 – Jun 2026",
              Inches(6.75), Inches(1.38), Inches(6.25), Inches(0.42), PURPLE)
    slide.shapes.add_picture("charts/energy_stocks.png",
                             Inches(6.75), Inches(1.85), Inches(6.25), Inches(5.45))

    # Divider
    rect(slide, Inches(6.58), Inches(1.38), Inches(0.04), Inches(5.95), MID_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Defense Stocks Chart Comparison
# ═══════════════════════════════════════════════════════════════════════════════
def slide_05_defense_charts(prs):
    slide = blank_slide(prs)
    header_band(slide, "國防股走勢對比",
                "Defense Stocks  ·  LMT · RTX · NOC  ·  Split-adjusted closing prices + crude oil (right axis)")

    war_badge(slide, "波灣戰爭  Gulf War",  "Jun 1990 – Apr 1991",
              Inches(0.4), Inches(1.38), Inches(6.2), Inches(0.42), TEAL)
    slide.shapes.add_picture("charts/defense_stocks_gulf.png",
                             Inches(0.4), Inches(1.85), Inches(6.2), Inches(5.45))

    war_badge(slide, "美伊戰爭  US–Iran War", "Jan 2026 – Jun 2026",
              Inches(6.75), Inches(1.38), Inches(6.25), Inches(0.42), PURPLE)
    slide.shapes.add_picture("charts/defense_stocks.png",
                             Inches(6.75), Inches(1.85), Inches(6.25), Inches(5.45))

    rect(slide, Inches(6.58), Inches(1.38), Inches(0.04), Inches(5.95), MID_GRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Performance Data Table (side-by-side)
# ═══════════════════════════════════════════════════════════════════════════════
def slide_06_table(prs, gulf_csv, iran_csv):
    slide = blank_slide(prs)
    header_band(slide, "各階段漲跌幅對比",
                "Percentage Change by Period  ·  Gulf War  vs  US–Iran War")

    gulf = pd.read_csv(gulf_csv).set_index("ticker")
    iran = pd.read_csv(iran_csv).set_index("ticker")

    # Unified ticker order: energy first, then defense, then oil
    ORDER   = ["XOM", "CVX", "COP", "LMT", "RTX", "NOC"]
    OIL_ROW = ("WTI_Crude", "CL=F")   # (gulf key, iran key)

    ENERGY  = {"XOM", "CVX", "COP"}

    # Column layout
    # [Ticker | Gulf P1 | Gulf P2 | Gulf P3 | divider | Iran P1 | Iran P2 | Iran P3]
    tbl = slide.shapes.add_table(
        len(ORDER) + 2, 8,   # +2: header + oil row
        Inches(0.3), Inches(1.42),
        Inches(12.73), Inches(5.85)
    ).table

    widths = [Inches(1.35),
              Inches(1.6), Inches(1.6), Inches(1.6),
              Inches(0.28),
              Inches(1.6), Inches(1.6), Inches(1.6)]
    for ci, w in enumerate(widths):
        tbl.columns[ci].width = w

    def cell(r, c, text, bold=False, size=11, bg=None, fg=WHITE,
             align=PP_ALIGN.CENTER):
        cl = tbl.cell(r, c)
        cl.text = ""
        tf = cl.text_frame
        tf.paragraphs[0].alignment = align
        run = tf.paragraphs[0].add_run()
        run.text           = text
        run.font.size      = Pt(size)
        run.font.bold      = bold
        run.font.color.rgb = fg
        if bg:
            cl.fill.solid()
            cl.fill.fore_color.rgb = bg

    # Header row
    cell(0, 0, "股票", bold=True, bg=DARK_BLUE, size=11)
    for ci, h in enumerate(["戰前期", "戰事中", "沙漠風暴"]):
        cell(0, ci + 1, h + "\n(Gulf)", bold=True, bg=TEAL, size=10)
    cell(0, 4, "", bg=DARK_BLUE)
    for ci, h in enumerate(["戰前期", "爆發期", "停火後"]):
        cell(0, ci + 5, h + "\n(Iran)", bold=True, bg=PURPLE, size=10)

    def val_cell(r, c, v, row_bg):
        fg = GREEN if v >= 0 else RED
        cell(r, c, f"{v:+.1f}%", bg=row_bg, fg=fg, size=11)

    # Data rows
    for ri, tkr in enumerate(ORDER):
        row_i   = ri + 1
        is_e    = tkr in ENERGY
        row_bg  = RGBColor(0xE8, 0xF0, 0xF8) if is_e else RGBColor(0xFF, 0xF0, 0xF0)
        label_c = MID_BLUE if is_e else RED
        cell(row_i, 0, tkr, bold=True, bg=row_bg, fg=label_c, size=12)

        g = gulf.loc[tkr]
        val_cell(row_i, 1, g["before_invasion"], row_bg)
        val_cell(row_i, 2, g["war_period"],      row_bg)
        val_cell(row_i, 3, g["desert_storm"],    row_bg)

        cell(row_i, 4, "", bg=MID_GRAY)

        n = iran.loc[tkr]
        val_cell(row_i, 5, n["before_war"],      row_bg)
        val_cell(row_i, 6, n["war_outbreak"],    row_bg)
        val_cell(row_i, 7, n["after_ceasefire"], row_bg)

    # Oil row
    oil_bg = RGBColor(0xFF, 0xF3, 0xE0)
    oil_ri = len(ORDER) + 1
    cell(oil_ri, 0, "原油", bold=True, bg=oil_bg, fg=ORANGE, size=12)
    g_oil = gulf.loc[OIL_ROW[0]]
    val_cell(oil_ri, 1, g_oil["before_invasion"], oil_bg)
    val_cell(oil_ri, 2, g_oil["war_period"],      oil_bg)
    val_cell(oil_ri, 3, g_oil["desert_storm"],    oil_bg)
    cell(oil_ri, 4, "", bg=MID_GRAY)
    n_oil = iran.loc[OIL_ROW[1]]
    val_cell(oil_ri, 5, n_oil["before_war"],      oil_bg)
    val_cell(oil_ri, 6, n_oil["war_outbreak"],    oil_bg)
    val_cell(oil_ri, 7, n_oil["after_ceasefire"], oil_bg)

    # Legend
    war_badge(slide, "波灣戰爭  Gulf War", "P1=戰前 / P2=戰事 / P3=沙漠風暴",
              Inches(0.3), Inches(7.1), Inches(5.5), Inches(0.35), TEAL)
    war_badge(slide, "美伊戰爭  US–Iran War", "P1=戰前 / P2=爆發 / P3=停火後",
              Inches(6.5), Inches(7.1), Inches(5.8), Inches(0.35), PURPLE)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Key Finding: Defense Stock Paradox
# ═══════════════════════════════════════════════════════════════════════════════
def slide_07_paradox(prs):
    slide = blank_slide(prs)
    header_band(slide, "核心謎題：國防股為何反應完全相反？",
                "Defense Stock Paradox  ·  1990 Rose  vs  2026 Fell  ·  Two Hypotheses")

    # Fact row
    rect(slide, Inches(0.4), Inches(1.4), Inches(5.95), Inches(0.72),
         RGBColor(0xE3, 0xF4, 0xF0))
    textbox(slide, "1990 波灣戰爭：國防股戰事中大漲  LMT +24%  NOC +30%",
            Inches(0.55), Inches(1.44), Inches(5.8), Inches(0.62),
            size=12, bold=True, color=TEAL, align=PP_ALIGN.CENTER)

    rect(slide, Inches(7.0), Inches(1.4), Inches(5.95), Inches(0.72),
         RGBColor(0xF3, 0xEA, 0xF8))
    textbox(slide, "2026 美伊戰爭：國防股戰事中下跌  LMT -11%  NOC -11%",
            Inches(7.15), Inches(1.44), Inches(5.8), Inches(0.62),
            size=12, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)

    rect(slide, Inches(6.25), Inches(1.4), Inches(0.83), Inches(0.72),
         RGBColor(0xFF, 0xE8, 0xE8))
    textbox(slide, "相反！", Inches(6.25), Inches(1.48), Inches(0.83), Inches(0.56),
            size=11, bold=True, color=RED, align=PP_ALIGN.CENTER)

    # Two hypotheses
    COL_W, BOX_H, TOP = Inches(6.1), Inches(3.9), Inches(2.28)
    COL_A, COL_B = Inches(0.4), Inches(6.83)

    rect(slide, COL_A, TOP, COL_W, BOX_H, RGBColor(0xE8, 0xF0, 0xF8))
    rect(slide, COL_A, TOP, COL_W, Inches(0.5), MID_BLUE)
    textbox(slide, "假說 A　市場預期差異（Surprise Effect）",
            COL_A + Inches(0.15), TOP + Inches(0.06),
            COL_W - Inches(0.3), Inches(0.4),
            size=13, bold=True, color=WHITE)
    bullet_rows(slide, [
        "1990：伊拉克入侵科威特為突發事件，\n市場毫無預期 → 宣戰後國防股大漲",
        "2026：地緣緊張醞釀數月，\nLMT/NOC 已在戰前大漲 32%/24%",
        "股價高點出現在 2026-02-28 之前，\n宣戰即成獲利了結訊號",
        "「買謠言、賣事實」\n(Buy the Rumor, Sell the News)",
    ], COL_A + Inches(0.15), TOP + Inches(0.6),
       COL_W - Inches(0.2), Inches(0.74),
       size=11, color=DARK_BLUE, marker="▸", marker_color=MID_BLUE)

    rect(slide, COL_B, TOP, COL_W, BOX_H, RGBColor(0xFF, 0xF3, 0xE8))
    rect(slide, COL_B, TOP, COL_W, Inches(0.5), ORANGE)
    textbox(slide, "假說 B　戰爭性質差異（War Character）",
            COL_B + Inches(0.15), TOP + Inches(0.06),
            COL_W - Inches(0.3), Inches(0.4),
            size=13, bold=True, color=WHITE)
    bullet_rows(slide, [
        "1990：沙漠風暴展現美軍壓倒性優勢，\n愛國者飛彈、精準轟炸大幅提升國防股評價",
        "2026：伊朗封鎖霍木茲海峽、\n飛彈穿透防空，顯示美軍防禦出現漏洞",
        "市場下修對美國國防工業的\n長期競爭力評估",
        "短期衝突 (39 天) 不足以驅動\n新一輪長期國防預算增加",
    ], COL_B + Inches(0.15), TOP + Inches(0.6),
       COL_W - Inches(0.2), Inches(0.74),
       size=11, color=DARK_BLUE, marker="▸", marker_color=ORANGE)

    # VS
    rect(slide, Inches(6.58), TOP + Inches(0.6),
         Inches(0.5), BOX_H - Inches(0.7), RGBColor(0xDD, 0xDD, 0xDD))
    textbox(slide, "VS", Inches(6.58), TOP + Inches(1.6), Inches(0.5), Inches(0.5),
            size=15, bold=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # Note
    rect(slide, Inches(0.4), Inches(6.35), Inches(12.53), Inches(0.95), LIGHT_GRAY)
    textbox(slide,
            "兩種假說都有數據支持，且可能同時成立。"
            "要明確區分，需要基本面資料：武器訂單數據、實際戰場損耗報告、國會國防預算修正案。\n"
            "這是純量化分析的邊界 — 也是人類判斷不可取代之處。",
            Inches(0.55), Inches(6.38), Inches(12.2), Inches(0.88),
            size=10, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Conclusion: Three Findings Across 35 Years
# ═══════════════════════════════════════════════════════════════════════════════
def slide_08_conclusion(prs):
    slide = blank_slide(prs)
    header_band(slide, "跨越 35 年的三個發現",
                "Three Findings Across 35 Years  ·  Gulf War 1990  &  US–Iran War 2026")

    findings = [
        (ORANGE, "①",
         "能源股：中東衝突的一致性受益者",
         "Energy Stocks Consistently Benefit from Middle-East Conflicts",
         "兩場戰爭中，能源股在戰前期與爆發期均呈正報酬。"
         "原油供應中斷預期是核心驅動力。"
         "差異在於幅度：1990 年 WTI 漲幅 +141%（供應實際中斷）；"
         "2026 年 +65%（短暫封鎖後恢復）。"
         "停火後能源股均回落，反映供應恢復預期。"),
        (RED, "②",
         "國防股：反應方向取決於「戰爭是否被預期」",
         "Defense Stocks: Direction Depends on Whether the War Was Pre-Priced",
         "1990 年突發入侵 → 無前置定價 → 宣戰後大漲（LMT+24%, NOC+30%）。\n"
         "2026 年緊張醞釀數月 → 充分定價 → 宣戰後反跌（LMT-11%, NOC-11%）。\n"
         "投資啟示：評估國防股，必須先問「市場已定價多少？」而非「戰爭會發生嗎？」"),
        (TEAL, "③",
         "原油是地緣衝突的最快領先指標",
         "Crude Oil Is the Fastest Leading Indicator of Geopolitical Conflict",
         "1990 年：伊拉克入侵當天 WTI 即跳漲，能源股滯後 1-2 日。\n"
         "2026 年：WTI 爆發期大漲 +42%，為所有資產中反應最劇烈的。\n"
         "操作啟示：監測原油期貨的異常波動，可作為能源股進出場的領先訊號。"),
    ]

    for i, (color, num, zh_title, en_title, body) in enumerate(findings):
        top = Inches(1.45) + i * Inches(1.95)
        rect(slide, Inches(0.4), top, Inches(12.53), Inches(1.88),
             RGBColor(0xF8, 0xF8, 0xF8))
        rect(slide, Inches(0.4), top, Inches(0.45), Inches(1.88), color)
        textbox(slide, num, Inches(0.4), top + Inches(0.62),
                Inches(0.45), Inches(0.6),
                size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        textbox(slide, zh_title,
                Inches(1.0), top + Inches(0.08),
                Inches(11.75), Inches(0.45),
                size=14, bold=True, color=color)
        textbox(slide, en_title,
                Inches(1.0), top + Inches(0.48),
                Inches(11.75), Inches(0.3),
                size=10, italic=True, color=DARK_GRAY)
        textbox(slide, body,
                Inches(1.0), top + Inches(0.8),
                Inches(11.75), Inches(1.0),
                size=10, color=DARK_BLUE)

    # Footer
    rect(slide, Inches(0.4), Inches(7.1), Inches(12.53), Inches(0.3), DARK_BLUE)
    textbox(slide,
            "數據期間：Gulf War 1990-06-01 ~ 1991-04-30  ·  "
            "US–Iran War 2026-01-01 ~ 2026-06-04  ·  "
            "數據來源：Yahoo Finance · EIA · FRED",
            Inches(0.5), Inches(7.1), Inches(12.3), Inches(0.3),
            size=9, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    prs = new_prs()

    slide_01_title(prs)
    slide_02_targets(prs)
    slide_03_background(prs)
    slide_04_energy_charts(prs)
    slide_05_defense_charts(prs)
    slide_06_table(prs, "data/analysis_results_gulf.csv", "data/analysis_results.csv")
    slide_07_paradox(prs)
    slide_08_conclusion(prs)

    out = "reports/gulf_vs_iran_comparison.pptx"
    prs.save(out)
    print(f"Saved {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
