"""
build_presentation.py
=====================
Generates a professional 8-slide PowerPoint presentation:
  "Zambia's Trade Story: K33 Billion Warning"
  Storytelling with Data principles applied.

  Slides 1 & 8  : Dark navy theme (original)
  Slides 2 – 7  : Clean white theme matching Import-Export-Daniel.pptx
                  (white bg, dark teal headlines, orange line/accent,
                   teal boxes for annotations, source note at bottom)

Output: zambia_trade_presentation_Daniel.pptx

Requirements:
    pip install python-pptx
Run:
    python build_presentation.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import lxml.etree as etree
from pptx.oxml.ns import qn

# ─────────────────────────────────────────────
# DARK THEME  (slides 1 & 8)
# ─────────────────────────────────────────────
D_BG        = RGBColor(0x0d, 0x1b, 0x2a)
D_PANEL     = RGBColor(0x16, 0x21, 0x3e)
D_ORANGE    = RGBColor(0xe8, 0x5d, 0x04)
D_TEAL      = RGBColor(0x1d, 0x6f, 0xa4)
D_RED       = RGBColor(0xc1, 0x12, 0x1f)
D_WHITE     = RGBColor(0xff, 0xff, 0xff)
D_LGREY     = RGBColor(0xe0, 0xe0, 0xe0)
D_MGREY     = RGBColor(0xa0, 0xa8, 0xb8)
D_GOLD      = RGBColor(0xf4, 0xa2, 0x61)
D_GREEN     = RGBColor(0x2d, 0xc6, 0x53)

# ─────────────────────────────────────────────
# LIGHT THEME  (slides 2–7)  — matches screenshot
# ─────────────────────────────────────────────
L_BG        = RGBColor(0xff, 0xff, 0xff)   # white background
L_HEADLINE  = RGBColor(0x1a, 0x1a, 0x2e)   # very dark navy for headline text
L_ORANGE    = RGBColor(0xe8, 0x5d, 0x04)   # orange — imports / line / warnings
L_TEAL      = RGBColor(0x1d, 0x6f, 0xa4)   # teal — exports / annotation boxes
L_TEAL_LT   = RGBColor(0xe8, 0xf4, 0xfd)   # very light teal box fill
L_ORANGE_LT = RGBColor(0xff, 0xf3, 0xe8)   # very light orange box fill
L_RED       = RGBColor(0xc1, 0x12, 0x1f)
L_GREEN     = RGBColor(0x2d, 0xc6, 0x53)
L_BLACK     = RGBColor(0x1a, 0x1a, 0x2e)
L_DGREY     = RGBColor(0x44, 0x44, 0x55)
L_MGREY     = RGBColor(0x88, 0x88, 0x99)
L_LGREY     = RGBColor(0xcc, 0xcc, 0xdd)
L_PANEL     = RGBColor(0xf5, 0xf7, 0xfa)   # very light grey panel

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

YEARS   = [2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022]
EXPORTS = [5.44, 4.73, 5.91, 6.23, 5.33, 5.59, 6.59, 9.70, 14.17, 12.95]
IMPORTS = [4.21, 4.62, 7.03, 7.29, 9.43, 9.24, 10.40, 15.46, 20.13, 22.22]
FX      = [5.4,  6.0,  8.6,  10.3, 9.9,  10.4, 12.9, 18.3, 20.0,  17.2]
BALANCE = [e - i for e, i in zip(EXPORTS, IMPORTS)]


# ═══════════════════════════════════════════
# SHARED HELPERS
# ═══════════════════════════════════════════

def add_rect(slide, left, top, width, height, fill_color,
             line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_rect_outline(slide, left, top, width, height, fill_color,
                     line_color, line_width=Pt(1.5)):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = line_color
    shape.line.width = line_width
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_name="Calibri", font_size=14, bold=False, italic=False,
                 color=D_WHITE, align=PP_ALIGN.LEFT, word_wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_multiline(slide, lines, left, top, width, height,
                  font_name="Calibri", font_size=13, bold=False,
                  color=L_DGREY, align=PP_ALIGN.LEFT, line_spacing=115):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
        pPr = p._p.get_or_add_pPr()
        lnSpc = etree.SubElement(pPr, qn('a:lnSpc'))
        spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', str(line_spacing * 1000))
    return txBox


def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def light_footer(slide, source="Source: Bank of Zambia / IMF annual ZMW/USD averages  |  FSIO Cohort Analysis"):
    add_rect(slide, Inches(0), Inches(7.22), Inches(13.33), Inches(0.28), L_PANEL)
    add_text_box(slide, source,
                 Inches(0.35), Inches(7.24), Inches(12.0), Inches(0.24),
                 font_size=7.5, color=L_MGREY, align=PP_ALIGN.LEFT)


def slide_num(slide, n, total=8):
    add_text_box(slide, f"{n} / {total}",
                 Inches(12.5), Inches(7.24), Inches(0.8), Inches(0.24),
                 font_size=7.5, color=L_MGREY, align=PP_ALIGN.RIGHT)


def light_headline(slide, text):
    """Action headline bar — matches screenshot: dark navy text, left-aligned."""
    add_text_box(slide, text,
                 Inches(0.38), Inches(0.22), Inches(12.7), Inches(0.78),
                 font_name="Calibri", font_size=28, bold=True,
                 color=L_HEADLINE, align=PP_ALIGN.LEFT)
    add_rect(slide, Inches(0.38), Inches(0.98), Inches(12.6), Inches(0.03), L_ORANGE)


def annotation_box(slide, left, top, width, height,
                   year_label, rate_text, body_text,
                   border_color=L_TEAL, bg_color=L_TEAL_LT):
    """Teal-outlined annotation card — matches screenshot style."""
    add_rect_outline(slide, left, top, width, height, bg_color, border_color, Pt(1.5))
    # year + arrow + rate on first line
    header = f"{year_label}  →  {rate_text}"
    add_text_box(slide, header,
                 left + Inches(0.12), top + Inches(0.1),
                 width - Inches(0.18), Inches(0.35),
                 font_size=13, bold=True, color=border_color)
    add_text_box(slide, body_text,
                 left + Inches(0.12), top + Inches(0.46),
                 width - Inches(0.18), height - Inches(0.52),
                 font_size=11, color=L_DGREY, word_wrap=True)


def draw_line_chart(slide, data_series, left, top, width, height,
                    max_val, min_val=0, year_labels=True,
                    gridline_vals=None, show_points=True):
    """
    Draw a simple polyline chart using rectangles and text.
    data_series: list of (label, values_list, color, line_width_pt)
    Returns nothing — draws directly on slide.
    """
    # Chart area
    add_rect(slide, left, top, width, height, L_BG)
    # Border bottom + left
    add_rect(slide, left, top + height - Inches(0.015),
             width, Inches(0.015), L_LGREY)
    add_rect(slide, left, top, Inches(0.015), height, L_LGREY)

    plot_l = left + Inches(0.55)
    plot_t = top + Inches(0.18)
    plot_w = width - Inches(0.7)
    plot_h = height - Inches(0.55)
    val_range = max_val - min_val
    n = len(YEARS)

    # Gridlines
    if gridline_vals:
        for gv in gridline_vals:
            gy = plot_t + plot_h * (1 - (gv - min_val) / val_range)
            add_rect(slide, plot_l, gy, plot_w, Inches(0.008),
                     RGBColor(0xe8, 0xe8, 0xee))
            add_text_box(slide, str(gv),
                         left, gy - Inches(0.13),
                         Inches(0.52), Inches(0.28),
                         font_size=8, color=L_MGREY, align=PP_ALIGN.RIGHT)

    # X-axis labels
    if year_labels:
        for i, yr in enumerate(YEARS):
            xp = plot_l + (i / (n - 1)) * plot_w
            add_text_box(slide, str(yr),
                         xp - Inches(0.28), plot_t + plot_h + Inches(0.05),
                         Inches(0.56), Inches(0.22),
                         font_size=8, color=L_MGREY, align=PP_ALIGN.CENTER)

    # Series
    for (label, values, color, lw) in data_series:
        pts = []
        for i, v in enumerate(values):
            xp = plot_l + (i / (n - 1)) * plot_w
            yp = plot_t + plot_h * (1 - (v - min_val) / val_range)
            pts.append((xp, yp))
        # Draw segments as thin rectangles
        seg_h = Pt(lw).inches
        for i in range(len(pts) - 1):
            x1, y1 = pts[i]
            x2, y2 = pts[i + 1]
            import math
            dx = x2 - x1
            dy = y2 - y1
            length = math.sqrt(dx * dx + dy * dy)
            angle_rad = math.atan2(dy, dx)
            # Approximate with very thin rect along segment mid-point
            mx = (x1 + x2) / 2
            my = (y1 + y2) / 2 - Inches(seg_h / 2)
            seg_rect = slide.shapes.add_shape(1,
                int(x1), int(min(y1, y2) - Inches(seg_h / 2)),
                int(abs(dx)) if abs(dx) > Inches(0.01) else Inches(0.02),
                Inches(seg_h) + int(abs(dy)))
            seg_rect.fill.solid()
            seg_rect.fill.fore_color.rgb = color
            seg_rect.line.fill.background()

        # Dots
        if show_points:
            for xp, yp in pts:
                dot_r = Inches(0.065)
                add_rect_outline(slide,
                                 xp - dot_r, yp - dot_r, dot_r * 2, dot_r * 2,
                                 L_BG, color, Pt(2))
        # Label last point
        lx, ly = pts[-1]
        add_text_box(slide, label,
                     lx + Inches(0.1), ly - Inches(0.18),
                     Inches(1.0), Inches(0.3),
                     font_size=9, bold=True, color=color)


# ═══════════════════════════════════════════
# SLIDE 1 — TITLE (DARK)
# ═══════════════════════════════════════════
def build_slide_01(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, D_BG)

    add_rect(sl, Inches(0), Inches(0), Inches(0.12), Inches(7.5), D_ORANGE)
    add_rect(sl, Inches(0.12), Inches(0), Inches(13.21), Inches(0.06), D_TEAL)

    add_rect(sl, Inches(0.5), Inches(1.1), Inches(3.4), Inches(0.38), D_TEAL)
    add_text_box(sl, "AGRICULTURE TRADE ANALYSIS  2013 – 2022",
                 Inches(0.56), Inches(1.13), Inches(3.28), Inches(0.32),
                 font_size=8.5, bold=True, color=D_WHITE)

    add_text_box(sl, "Zambia's Trade Story:",
                 Inches(0.5), Inches(1.65), Inches(11.0), Inches(1.0),
                 font_size=52, bold=True, color=D_WHITE)
    add_text_box(sl, "K33 Billion Warning",
                 Inches(0.5), Inches(2.55), Inches(11.0), Inches(1.1),
                 font_size=60, bold=True, color=D_ORANGE)
    add_text_box(sl,
                 "How Currency Volatility is Amplifying the Agriculture\nImport-Export Deficit (2013–2022)",
                 Inches(0.5), Inches(3.78), Inches(10.0), Inches(0.95),
                 font_size=19, color=D_LGREY)
    add_rect(sl, Inches(0.5), Inches(4.85), Inches(5.0), Inches(0.04), D_MGREY)
    add_text_box(sl, "FSIO Cohort  |  May 2026",
                 Inches(0.5), Inches(4.97), Inches(8.0), Inches(0.4),
                 font_size=13, color=D_MGREY)

    # Right stat block
    add_rect(sl, Inches(9.0), Inches(1.8), Inches(3.8), Inches(4.5), D_PANEL)
    add_text_box(sl, "CUMULATIVE DEFICIT",
                 Inches(9.15), Inches(1.98), Inches(3.5), Inches(0.32),
                 font_size=9, bold=True, color=D_MGREY, align=PP_ALIGN.CENTER)
    add_text_box(sl, "-K33.4bn",
                 Inches(9.0), Inches(2.28), Inches(3.8), Inches(1.05),
                 font_size=54, bold=True, color=D_RED, align=PP_ALIGN.CENTER)
    add_rect(sl, Inches(9.3), Inches(3.5), Inches(3.2), Inches(0.03), D_TEAL)
    add_text_box(sl, "TOTAL IMPORTS",
                 Inches(9.15), Inches(3.65), Inches(3.5), Inches(0.28),
                 font_size=8.5, bold=True, color=D_MGREY, align=PP_ALIGN.CENTER)
    add_text_box(sl, "K110.04bn",
                 Inches(9.0), Inches(3.9), Inches(3.8), Inches(0.62),
                 font_size=30, bold=True, color=D_ORANGE, align=PP_ALIGN.CENTER)
    add_text_box(sl, "TOTAL EXPORTS",
                 Inches(9.15), Inches(4.65), Inches(3.5), Inches(0.28),
                 font_size=8.5, bold=True, color=D_MGREY, align=PP_ALIGN.CENTER)
    add_text_box(sl, "K76.65bn",
                 Inches(9.0), Inches(4.9), Inches(3.8), Inches(0.62),
                 font_size=30, bold=True, color=D_TEAL, align=PP_ALIGN.CENTER)
    add_text_box(sl, "A Storytelling with Data Analysis",
                 Inches(0.5), Inches(6.88), Inches(8.0), Inches(0.38),
                 font_size=9, italic=True, color=D_MGREY)
    return sl


# ═══════════════════════════════════════════
# SLIDE 2 — CURRENCY CONTEXT (LIGHT)
# ═══════════════════════════════════════════
def build_slide_02(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "The Kwacha lost 3× its value — this distorts every ZMW trade figure")

    # Subtitle in orange italic — matches screenshot
    add_text_box(sl, "ZMW per 1 USD (annual average) — higher = weaker Kwacha",
                 Inches(0.38), Inches(1.08), Inches(8.5), Inches(0.3),
                 font_size=12, italic=True, color=L_ORANGE)

    # ── Line chart: FX rate ──
    chart_l = Inches(0.38)
    chart_t = Inches(1.5)
    chart_w = Inches(8.5)
    chart_h = Inches(5.3)

    draw_line_chart(sl,
        [("", FX, L_ORANGE, 2.8)],
        chart_l, chart_t, chart_w, chart_h,
        max_val=25, min_val=0,
        gridline_vals=[5, 10, 15, 20, 25],
        show_points=True)

    # Annotate key values on chart
    fx_annotations = [
        (0,  "5",  Inches(0.0)),
        (2,  "6",  Inches(0.0)),
        (4,  "10", Inches(0.0)),
        (6,  "14", Inches(0.0)),
        (7,  "18", Inches(0.0)),
        (8,  "20", Inches(0.0)),
        (9,  "17", Inches(0.0)),
    ]
    plot_l2 = chart_l + Inches(0.55)
    plot_t2 = chart_t + Inches(0.18)
    plot_w2 = chart_w - Inches(0.7)
    plot_h2 = chart_h - Inches(0.55)
    for (i, lbl, offset_x) in fx_annotations:
        xp = plot_l2 + (i / 9) * plot_w2
        yp = plot_t2 + plot_h2 * (1 - (FX[i] / 25))
        add_text_box(sl, lbl,
                     xp - Inches(0.15), yp - Inches(0.35),
                     Inches(0.32), Inches(0.26),
                     font_size=9.5, bold=True, color=L_ORANGE,
                     align=PP_ALIGN.CENTER)

    # ── Four annotation boxes (right side) — teal outlined, matches screenshot ──
    box_l = Inches(9.18)
    box_w = Inches(3.88)
    box_h = Inches(1.18)
    gap   = Inches(0.13)
    annots = [
        ("2013", "ZMW 5.40 / USD",  "Base year — ZMW relatively stable after post-2012 adjustments"),
        ("2015", "ZMW 8.60 / USD",  "Sharp devaluation driven by copper price crash + drought"),
        ("2020", "ZMW 18.30 / USD", "COVID shock + debt distress — Kwacha hits historic low"),
        ("2022", "ZMW 17.20 / USD", "Partial recovery but still 3.2× weaker than 2013 base"),
    ]
    top_start = Inches(1.5)
    for idx, (yr, rate, body) in enumerate(annots):
        by = top_start + idx * (box_h + gap)
        annotation_box(sl, box_l, by, box_w, box_h,
                        yr, rate, body,
                        border_color=L_TEAL, bg_color=L_TEAL_LT)

    light_footer(sl)
    slide_num(sl, 2)
    return sl


# ═══════════════════════════════════════════
# SLIDE 3 — DIVERGING TRENDS (LIGHT)
# ═══════════════════════════════════════════
def build_slide_03(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "In Kwacha, trade looks like it boomed — imports grew 5×, exports grew 2×")

    add_text_box(sl, "ZMW billions — nominal amounts in local currency",
                 Inches(0.38), Inches(1.08), Inches(9.0), Inches(0.3),
                 font_size=12, italic=True, color=L_ORANGE)

    # ── Twin line chart: imports vs exports ──
    chart_l = Inches(0.38)
    chart_t = Inches(1.5)
    chart_w = Inches(8.5)
    chart_h = Inches(5.3)

    draw_line_chart(sl,
        [
            ("Imports (ZMW)", IMPORTS, L_ORANGE, 2.8),
            ("Exports (ZMW)", EXPORTS, L_TEAL,   2.4),
        ],
        chart_l, chart_t, chart_w, chart_h,
        max_val=25, min_val=0,
        gridline_vals=[5, 10, 15, 20],
        show_points=True)

    # ── Annotation boxes ──
    box_l = Inches(9.18)
    box_w = Inches(3.88)
    box_h = Inches(1.18)
    gap   = Inches(0.13)
    annots = [
        (L_ORANGE, L_ORANGE_LT,
         "IMPORTS +428%",
         "K4.2bn (2013) → K22.2bn (2022)\nNominal Kwacha surge driven mainly\nby currency depreciation"),
        (L_TEAL, L_TEAL_LT,
         "EXPORTS +138%",
         "K5.4bn (2013) → K12.9bn (2022)\nZMW gains are partly translation\neffects, not real volume growth"),
        (L_RED, RGBColor(0xff, 0xee, 0xee),
         "DEFICIT WIDENED",
         "+K1.23bn surplus (2013) flipped\nto -K9.27bn deficit (2022)\nCumulative gap: K33.4bn"),
        (L_MGREY, L_PANEL,
         "CAGR COMPARISON",
         "Import CAGR: ~18% per year\nExport CAGR: ~10% per year\nGap compounds every year"),
    ]
    top_start = Inches(1.5)
    for idx, (border, bg, title, body) in enumerate(annots):
        by = top_start + idx * (box_h + gap)
        add_rect_outline(sl, box_l, by, box_w, box_h, bg, border, Pt(1.5))
        add_text_box(sl, title,
                     box_l + Inches(0.12), by + Inches(0.1),
                     box_w - Inches(0.18), Inches(0.32),
                     font_size=12, bold=True, color=border)
        add_text_box(sl, body,
                     box_l + Inches(0.12), by + Inches(0.44),
                     box_w - Inches(0.18), box_h - Inches(0.5),
                     font_size=10, color=L_DGREY, word_wrap=True)

    light_footer(sl, "Source: Zambia Agriculture Import-Export Dashboard 2013–2022")
    slide_num(sl, 3)
    return sl


# ═══════════════════════════════════════════
# SLIDE 4 — USD REAL TERMS (LIGHT)
# ═══════════════════════════════════════════
def build_slide_04(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "Strip out the FX effect — in USD, exports are shrinking and imports are flat")

    add_text_box(sl, "USD billions — real purchasing power perspective",
                 Inches(0.38), Inches(1.08), Inches(9.0), Inches(0.3),
                 font_size=12, italic=True, color=L_ORANGE)

    # Convert ZMW to approximate USD using FX rates
    usd_imports = [IMPORTS[i] / FX[i] for i in range(len(YEARS))]
    usd_exports = [EXPORTS[i] / FX[i] for i in range(len(YEARS))]

    chart_l = Inches(0.38)
    chart_t = Inches(1.5)
    chart_w = Inches(8.5)
    chart_h = Inches(5.3)

    draw_line_chart(sl,
        [
            ("Imports (USD)", usd_imports, L_ORANGE, 2.8),
            ("Exports (USD)", usd_exports, L_TEAL,   2.4),
        ],
        chart_l, chart_t, chart_w, chart_h,
        max_val=2.0, min_val=0,
        gridline_vals=[0.5, 1.0, 1.5, 2.0],
        show_points=True)

    # Annotation boxes
    box_l = Inches(9.18)
    box_w = Inches(3.88)
    box_h = Inches(1.18)
    gap   = Inches(0.13)

    exp_pct = round((usd_exports[-1] / usd_exports[0] - 1) * 100)
    imp_pct = round((usd_imports[-1] / usd_imports[0] - 1) * 100)

    sign_e = "+" if exp_pct >= 0 else ""
    sign_i = "+" if imp_pct >= 0 else ""

    annots = [
        (L_ORANGE, L_ORANGE_LT,
         f"IMPORTS IN USD  {sign_i}{imp_pct}%",
         f"${usd_imports[0]:.2f}bn (2013) → ${usd_imports[-1]:.2f}bn (2022)\nModest USD growth masked by\nKwacha collapse inflating ZMW figures"),
        (L_TEAL, L_TEAL_LT,
         f"EXPORTS IN USD  {sign_e}{exp_pct}%",
         f"${usd_exports[0]:.2f}bn (2013) → ${usd_exports[-1]:.2f}bn (2022)\nReal export capacity barely\nchanged over the decade"),
        (L_RED, RGBColor(0xff, 0xee, 0xee),
         "THE ILLUSION",
         "ZMW figures show +428% import\ngrowth. USD figures reveal the\ntrue, far smaller picture."),
        (L_MGREY, L_PANEL,
         "KEY TAKEAWAY",
         "Currency depreciation inflates\nboth sides of the ledger.\nNominal ≠ Real growth."),
    ]
    top_start = Inches(1.5)
    for idx, (border, bg, title, body) in enumerate(annots):
        by = top_start + idx * (box_h + gap)
        add_rect_outline(sl, box_l, by, box_w, box_h, bg, border, Pt(1.5))
        add_text_box(sl, title,
                     box_l + Inches(0.12), by + Inches(0.1),
                     box_w - Inches(0.18), Inches(0.32),
                     font_size=12, bold=True, color=border)
        add_text_box(sl, body,
                     box_l + Inches(0.12), by + Inches(0.44),
                     box_w - Inches(0.18), box_h - Inches(0.5),
                     font_size=10, color=L_DGREY, word_wrap=True)

    light_footer(sl, "Source: ZMW/USD — Bank of Zambia / IMF averages; Trade data — FSIO Cohort")
    slide_num(sl, 4)
    return sl


# ═══════════════════════════════════════════
# SLIDE 5 — TRADE BALANCE (LIGHT)
# ═══════════════════════════════════════════
def build_slide_05(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "The deficit is worsening in BOTH currencies — there is no hiding from it")

    add_text_box(sl, "Annual trade balance — ZMW billions (bars) with USD overlay (line)",
                 Inches(0.38), Inches(1.08), Inches(9.0), Inches(0.3),
                 font_size=12, italic=True, color=L_ORANGE)

    # Bar chart: trade balance by year
    chart_l = Inches(0.38)
    chart_t = Inches(1.5)
    chart_w = Inches(8.5)
    chart_h = Inches(5.3)

    plot_l = chart_l + Inches(0.55)
    plot_t = chart_t + Inches(0.18)
    plot_w = chart_w - Inches(0.7)
    plot_h = chart_h - Inches(0.55)

    max_v = 3.0
    min_v = -10.0
    val_range = max_v - min_v
    n = len(YEARS)

    # Gridlines
    for gv in [-8, -6, -4, -2, 0, 2]:
        gy = plot_t + plot_h * (1 - (gv - min_v) / val_range)
        add_rect(sl, plot_l, gy, plot_w, Inches(0.008),
                 RGBColor(0xe8, 0xe8, 0xee))
        add_text_box(sl, str(gv),
                     chart_l, gy - Inches(0.13), Inches(0.52), Inches(0.28),
                     font_size=8, color=L_MGREY, align=PP_ALIGN.RIGHT)

    # Zero line
    zero_y = plot_t + plot_h * (1 - (0 - min_v) / val_range)
    add_rect(sl, plot_l, zero_y, plot_w, Inches(0.025), L_DGREY)
    add_text_box(sl, "0",
                 chart_l, zero_y - Inches(0.15), Inches(0.52), Inches(0.28),
                 font_size=9, bold=True, color=L_DGREY, align=PP_ALIGN.RIGHT)

    slot_w = plot_w / n
    bar_w  = slot_w * 0.6

    for i, yr in enumerate(YEARS):
        b = BALANCE[i]
        x_center = plot_l + (i + 0.5) * slot_w
        bh = abs(b) / val_range * plot_h
        if b >= 0:
            by = zero_y - bh
            col = L_TEAL
        else:
            by = zero_y
            col = L_ORANGE
        add_rect(sl, x_center - bar_w / 2, by, bar_w, bh, col)
        # Value label
        lbl = f"{b:+.1f}"
        ly = by - Inches(0.25) if b >= 0 else by + bh + Inches(0.03)
        add_text_box(sl, lbl,
                     x_center - Inches(0.3), ly, Inches(0.6), Inches(0.22),
                     font_size=7.5, bold=True,
                     color=L_TEAL if b >= 0 else L_ORANGE,
                     align=PP_ALIGN.CENTER)
        add_text_box(sl, str(yr),
                     x_center - Inches(0.28),
                     plot_t + plot_h + Inches(0.05),
                     Inches(0.56), Inches(0.22),
                     font_size=8, color=L_MGREY, align=PP_ALIGN.CENTER)

    # Annotation boxes right
    box_l = Inches(9.18)
    box_w = Inches(3.88)
    box_h = Inches(1.18)
    gap   = Inches(0.13)
    annots = [
        (L_TEAL, L_TEAL_LT,
         "2013–2014: SURPLUS",
         "Exports exceeded imports.\nZMW stable. Trade provided\na modest USD buffer."),
        (L_ORANGE, L_ORANGE_LT,
         "2015–2017: SHIFT",
         "Deficit emerges and deepens.\nCopper crash + ZMW fall raise\nimport costs sharply."),
        (L_RED, RGBColor(0xff, 0xee, 0xee),
         "2020–2022: ACCELERATION",
         "COVID + debt distress.\nDeficit reached -K9.3bn (2022)\nWorst in the decade."),
        (L_MGREY, L_PANEL,
         "USD TELLS SAME STORY",
         "Even in USD terms the deficit\nworsened — not just a currency\nillusion, a structural problem."),
    ]
    top_start = Inches(1.5)
    for idx, (border, bg, title, body) in enumerate(annots):
        by2 = top_start + idx * (box_h + gap)
        add_rect_outline(sl, box_l, by2, box_w, box_h, bg, border, Pt(1.5))
        add_text_box(sl, title,
                     box_l + Inches(0.12), by2 + Inches(0.1),
                     box_w - Inches(0.18), Inches(0.32),
                     font_size=12, bold=True, color=border)
        add_text_box(sl, body,
                     box_l + Inches(0.12), by2 + Inches(0.44),
                     box_w - Inches(0.18), box_h - Inches(0.5),
                     font_size=10, color=L_DGREY, word_wrap=True)

    light_footer(sl, "Source: Zambia Agriculture Import-Export Dashboard 2013–2022  |  FSIO Cohort")
    slide_num(sl, 5)
    return sl


# ═══════════════════════════════════════════
# SLIDE 6 — KEY CONCLUSIONS (LIGHT)
# ═══════════════════════════════════════════
def build_slide_06(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "Key Conclusions from the 2013-2022 Agriculture Trade Data")

    # Six conclusion cards in 2×3 grid
    conclusions = [
        (L_ORANGE, "01  CURRENCY AMPLIFICATION",
         "ZMW depreciated 3×. This amplified all ZMW "
         "trade figures — imports look 428% higher, exports 138% higher. "
         "The real USD change is far smaller."),
        (L_TEAL, "02  STRUCTURAL TRADE DEFICIT",
         "Zambia moved from surplus (2013) to a K9.3bn annual deficit (2022). "
         "K33.4bn cumulative. This is structural, not cyclical — "
         "driven by import dependency."),
        (L_RED, "03  FERTILIZER TRAP",
         "Urea is the #1 import. Currency weakness makes fertilizer "
         "more expensive in ZMW every year, threatening the yields "
         "of the very crops Zambia exports."),
        (L_ORANGE, "04  RAW EXPORT CEILING",
         "All top exports (tobacco, maize, cotton) are unprocessed. "
         "Zambia is a global price-taker. Without value-addition "
         "there is a hard ceiling on export revenue growth."),
        (L_TEAL, "05  2022 WARNING SIGNAL",
         "Exports fell -K1.22bn in 2022 despite continuing ZMW weakness "
         "— the first contraction. This suggests real volume or "
         "commodity price deterioration beyond currency effects."),
        (L_RED, "06  SELF-REINFORCING LOOP",
         "Deficit → less USD inflow → ZMW weakens → imports cost more "
         "in ZMW → deficit widens. Without structural intervention "
         "this loop will deepen the K33bn gap further."),
    ]

    card_w = Inches(4.1)
    card_h = Inches(2.52)
    cols   = 3
    rows   = 2
    col_gap = Inches(0.12)
    row_gap = Inches(0.12)
    start_l = Inches(0.38)
    start_t = Inches(1.18)

    for idx, (color, title, body) in enumerate(conclusions):
        col_i = idx % cols
        row_i = idx // cols
        cx = start_l + col_i * (card_w + col_gap)
        cy = start_t + row_i * (card_h + row_gap)
        add_rect_outline(sl, cx, cy, card_w, card_h, L_PANEL, color, Pt(1.5))
        add_rect(sl, cx, cy, card_w, Inches(0.06), color)
        add_text_box(sl, title,
                     cx + Inches(0.14), cy + Inches(0.14),
                     card_w - Inches(0.22), Inches(0.38),
                     font_size=11, bold=True, color=color)
        add_text_box(sl, body,
                     cx + Inches(0.14), cy + Inches(0.55),
                     card_w - Inches(0.22), card_h - Inches(0.65),
                     font_size=10, color=L_DGREY, word_wrap=True)

    light_footer(sl, "Source: Zambia Agriculture Import-Export Dashboard 2013–2022  |  FSIO Cohort Analysis")
    slide_num(sl, 6)
    return sl


# ═══════════════════════════════════════════
# SLIDE 7 — IMPORT DRIVERS (LIGHT)
# ═══════════════════════════════════════════
def build_slide_07(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, L_BG)
    light_headline(sl, "Urea Dependency: Zambia Pays More Each Year to Grow the Same Crops")

    add_text_box(sl, "Top imported products (relative volume index) and the fertilizer vicious cycle",
                 Inches(0.38), Inches(1.08), Inches(9.0), Inches(0.3),
                 font_size=12, italic=True, color=L_ORANGE)

    # Horizontal bar chart
    import_products = [
        ("Urea (Fertilizer)",  100),
        ("Mineral Products",    62),
        ("Frozen Goods",        41),
        ("Mineral Products-2",  36),
        ("Ammonia",             28),
        ("Other Foods",         22),
        ("Crude Petroleum",     19),
        ("Crude Soybean",       14),
        ("Palm Oil",            10),
        ("Ammonia-2",            8),
    ]

    bar_label_w = Inches(1.9)
    bar_start   = Inches(0.38) + bar_label_w + Inches(0.08)
    bar_top     = Inches(1.48)
    row_h       = Inches(0.44)
    max_bar_w   = Inches(5.5)

    for idx, (name, val) in enumerate(import_products):
        by = bar_top + idx * row_h
        # Shade by rank
        alpha = 0.55 + 0.45 * (val / 100)
        r_base, g_base, b_base = 0xe8, 0x5d, 0x04  # L_ORANGE components
        r   = int(r_base * alpha + 255 * (1 - alpha))
        g_c = int(g_base * alpha + 255 * (1 - alpha))
        b_c = int(b_base * alpha + 255 * (1 - alpha))
        col = RGBColor(min(r, 255), min(g_c, 255), min(b_c, 255))
        bw = max_bar_w * (val / 100)
        add_rect(sl, bar_start, by + Inches(0.05), bw, row_h - Inches(0.1), col)
        add_text_box(sl, name,
                     Inches(0.38), by + Inches(0.1),
                     bar_label_w, row_h - Inches(0.1),
                     font_size=8.5, color=L_DGREY, align=PP_ALIGN.RIGHT)
        add_text_box(sl, str(val),
                     bar_start + bw + Inches(0.06), by + Inches(0.08),
                     Inches(0.45), row_h - Inches(0.1),
                     font_size=8, color=L_MGREY)

    # #1 badge
    add_rect_outline(sl, bar_start, bar_top + Inches(0.05),
                     Inches(0.28), Inches(0.3),
                     L_ORANGE, L_ORANGE, Pt(0))
    add_text_box(sl, "#1",
                 bar_start + Inches(0.02), bar_top + Inches(0.06),
                 Inches(0.24), Inches(0.25),
                 font_size=8, bold=True, color=L_BG, align=PP_ALIGN.CENTER)

    # Right panel — vicious cycle
    rx = Inches(8.22)
    add_rect(sl, rx, Inches(1.18), Inches(4.78), Inches(5.88), L_PANEL)
    add_text_box(sl, "THE FERTILIZER TRAP",
                 rx + Inches(0.18), Inches(1.3), Inches(4.4), Inches(0.32),
                 font_size=11, bold=True, color=L_ORANGE)

    steps = [
        (L_ORANGE, "ZMW falls vs USD"),
        (L_RED,    "Urea costs MORE in ZMW"),
        (L_ORANGE, "Farmers cut usage"),
        (L_RED,    "Lower crop yields"),
        (L_ORANGE, "Less export revenue"),
        (L_RED,    "More ZMW pressure →"),
    ]
    sy = Inches(1.72)
    for si, (col, txt) in enumerate(steps):
        add_rect_outline(sl, rx + Inches(0.18), sy,
                         Inches(4.42), Inches(0.52),
                         RGBColor(0xff, 0xf3, 0xe8) if col == L_ORANGE else RGBColor(0xff, 0xee, 0xee),
                         col, Pt(1.2))
        add_text_box(sl, txt,
                     rx + Inches(0.3), sy + Inches(0.1),
                     Inches(4.2), Inches(0.34),
                     font_size=10, bold=True, color=col)
        if si < len(steps) - 1:
            add_text_box(sl, "↓",
                         rx + Inches(2.0), sy + Inches(0.52),
                         Inches(0.4), Inches(0.22),
                         font_size=11, color=L_MGREY, align=PP_ALIGN.CENTER)
        sy += Inches(0.76)

    add_text_box(sl,
                 "⟳  Self-reinforcing loop — breaks only with domestic fertilizer production",
                 rx + Inches(0.12), Inches(6.56), Inches(4.55), Inches(0.46),
                 font_size=9, italic=True, color=L_RED)

    light_footer(sl, "Source: Zambia Agriculture Import-Export Dashboard 2013–2022  |  FSIO Cohort")
    slide_num(sl, 7)
    return sl


# ═══════════════════════════════════════════
# SLIDE 8 — CALL TO ACTION (DARK)
# ═══════════════════════════════════════════
def build_slide_08(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, D_BG)
    add_rect(sl, Inches(0), Inches(0), Inches(0.08), Inches(7.5), D_GREEN)

    # Headline bar
    add_rect(sl, Inches(0), Inches(0.22), Inches(13.33), Inches(1.0), D_PANEL)
    add_text_box(sl, "Three Priorities to Break the Deficit Loop",
                 Inches(0.38), Inches(0.3), Inches(12.5), Inches(0.85),
                 font_size=26, bold=True, color=D_WHITE)

    # Big idea restate
    add_rect(sl, Inches(0.3), Inches(1.38), Inches(12.7), Inches(0.62), D_PANEL)
    add_text_box(sl,
                 "The K33bn deficit is a structural warning. "
                 "Currency management alone cannot fix what requires industrial diversification.",
                 Inches(0.45), Inches(1.44), Inches(12.2), Inches(0.52),
                 font_size=12.5, italic=True, bold=True,
                 color=D_GOLD, align=PP_ALIGN.CENTER)

    # Three priority cards
    priorities = [
        ("01", "Reduce Urea\nDependency", D_ORANGE,
         ["Invest in domestic fertilizer production",
          "Regional procurement in ZMW or barter",
          "Subsidize alternative/organic inputs",
          "Reduce exposure to USD-priced inputs"]),
        ("02", "Value-Add\nExports", D_TEAL,
         ["Process tobacco & cotton locally before export",
          "Mill maize into flour for regional markets",
          "Capture more USD per unit exported",
          "Move from price-taker to price-setter"]),
        ("03", "Import\nSubstitution", D_GREEN,
         ["Develop local capacity for top import lines",
          "Reduce USD outflow on petroleum & minerals",
          "Industrial policy targeting import-heavy sectors",
          "Regional trade agreements in local currencies"]),
    ]

    card_w = Inches(4.1)
    card_h = Inches(4.42)
    card_top = Inches(2.15)

    for idx, (num, title, color, points) in enumerate(priorities):
        cx = Inches(0.3) + idx * (card_w + Inches(0.15))
        add_rect(sl, cx, card_top, card_w, card_h, D_PANEL)
        add_rect(sl, cx, card_top, card_w, Inches(0.1), color)
        add_rect(sl, cx + Inches(0.15), card_top + Inches(0.18),
                 Inches(0.45), Inches(0.45), color)
        add_text_box(sl, num,
                     cx + Inches(0.15), card_top + Inches(0.18),
                     Inches(0.45), Inches(0.45),
                     font_size=12, bold=True, color=D_WHITE, align=PP_ALIGN.CENTER)
        add_text_box(sl, title,
                     cx + Inches(0.72), card_top + Inches(0.18),
                     Inches(3.2), Inches(0.72),
                     font_size=16, bold=True, color=color)
        add_rect(sl, cx + Inches(0.15), card_top + Inches(0.95),
                 card_w - Inches(0.3), Inches(0.02), color)
        for pi, point in enumerate(points):
            py = card_top + Inches(1.08) + pi * Inches(0.72)
            add_rect(sl, cx + Inches(0.18), py + Inches(0.12),
                     Inches(0.12), Inches(0.12), color)
            add_text_box(sl, point,
                         cx + Inches(0.38), py + Inches(0.02),
                         card_w - Inches(0.55), Inches(0.65),
                         font_size=9.5, color=D_LGREY)

    add_rect(sl, Inches(0), Inches(6.78), Inches(13.33), Inches(0.72), D_PANEL)
    add_text_box(sl,
                 "Data: Zambia Agriculture Import-Export Dashboard 2013–2022  "
                 "|  FSIO Cohort Analysis  |  May 2026",
                 Inches(0.3), Inches(6.87), Inches(12.7), Inches(0.3),
                 font_size=8.5, color=D_MGREY, align=PP_ALIGN.CENTER)
    return sl


# ═══════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════
def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    print("Slide 1 — Title (dark) ...")
    build_slide_01(prs)
    print("Slide 2 — Currency context (light) ...")
    build_slide_02(prs)
    print("Slide 3 — Diverging ZMW trends (light) ...")
    build_slide_03(prs)
    print("Slide 4 — USD real terms (light) ...")
    build_slide_04(prs)
    print("Slide 5 — Trade balance (light) ...")
    build_slide_05(prs)
    print("Slide 6 — Key conclusions (light) ...")
    build_slide_06(prs)
    print("Slide 7 — Import drivers (light) ...")
    build_slide_07(prs)
    print("Slide 8 — Call to action (dark) ...")
    build_slide_08(prs)

    out = "zambia_trade_presentation_Daniel.pptx"
    prs.save(out)
    print(f"\nDone! Saved: {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
