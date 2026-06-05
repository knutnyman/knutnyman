from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette ──────────────────────────────────────────────────────────────────
ABBOTT_BLUE   = RGBColor(0x00, 0x47, 0x8A)
LIGHT_BLUE    = RGBColor(0xE8, 0xF2, 0xFB)
MID_BLUE      = RGBColor(0x4A, 0x90, 0xC4)
DARK_GREY     = RGBColor(0x2D, 0x2D, 0x2D)
MID_GREY      = RGBColor(0x5A, 0x5A, 0x5A)
LIGHT_GREY    = RGBColor(0xF4, 0xF4, 0xF4)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_GREEN  = RGBColor(0x00, 0x8A, 0x4B)
ACCENT_ORANGE = RGBColor(0xE8, 0x7D, 0x1E)
ACCENT_RED    = RGBColor(0xC0, 0x39, 0x2B)
LIGHT_RED     = RGBColor(0xFB, 0xEB, 0xE8)
LIGHT_GREEN   = RGBColor(0xE8, 0xF5, 0xEE)
LIGHT_ORANGE  = RGBColor(0xFD, 0xF3, 0xE5)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ── Helpers ───────────────────────────────────────────────────────────────────
def add_rect(slide, l, t, w, h, fill=None, line=None, line_w=Pt(0.5)):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             font_size=12, bold=False, color=DARK_GREY,
             align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def add_lines(slide, lines, l, t, w, h, font_size=10,
              color=DARK_GREY, bold_first=False):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = (bold_first and i == 0)
        run.font.color.rgb = color
    return txb

def header_bar(slide, title, subtitle=None):
    add_rect(slide, 0, 0, 13.33, 1.1, fill=ABBOTT_BLUE)
    add_text(slide, title, 0.3, 0.1, 10, 0.55,
             font_size=26, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, subtitle, 0.3, 0.65, 10, 0.4,
                 font_size=13, color=RGBColor(0xBF, 0xD9, 0xF2))
    add_rect(slide, 0, 1.1, 13.33, 0.04, fill=ACCENT_RED)

def footer(slide, text="Abbott Diagnostics — China VBP Analysis"):
    add_rect(slide, 0, 7.25, 13.33, 0.25, fill=LIGHT_GREY)
    add_text(slide, text, 0.3, 7.27, 12, 0.2, font_size=8, color=MID_GREY)

def table_slide(slide, headers, rows, l, t, w, col_widths=None,
                header_bg=ABBOTT_BLUE, row_bg1=WHITE, row_bg2=LIGHT_GREY,
                font_size=10, header_font=11, row_h=0.38):
    n_cols = len(headers)
    if col_widths is None:
        col_widths = [w / n_cols] * n_cols
    x = l
    for hdr, cw in zip(headers, col_widths):
        add_rect(slide, x, t, cw, row_h, fill=header_bg)
        add_text(slide, hdr, x+0.07, t+0.06, cw-0.14, row_h-0.1,
                 font_size=header_font, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        x += cw
    for r, row in enumerate(rows):
        y = t + row_h * (r + 1)
        bg = row_bg1 if r % 2 == 0 else row_bg2
        x = l
        for i, (cell, cw) in enumerate(zip(row, col_widths)):
            add_rect(slide, x, y, cw, row_h, fill=bg,
                     line=RGBColor(0xCC, 0xCC, 0xCC))
            add_text(slide, str(cell), x+0.07, y+0.05, cw-0.14, row_h-0.08,
                     font_size=font_size,
                     align=PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT)
            x += cw

def timeline_event(slide, x, y, date, title, body, color=ABBOTT_BLUE, dot_color=None):
    if dot_color is None:
        dot_color = color
    # dot
    dot = slide.shapes.add_shape(9, Inches(x-0.12), Inches(y+0.04), Inches(0.22), Inches(0.22))
    dot.fill.solid()
    dot.fill.fore_color.rgb = dot_color
    dot.line.fill.background()
    add_rect(slide, x+0.14, y+0.12, 3.6, 0.02, fill=RGBColor(0xCC, 0xCC, 0xCC))
    add_text(slide, date, x+0.2, y-0.02, 3.5, 0.22,
             font_size=9, bold=True, color=color)
    add_text(slide, title, x+0.2, y+0.17, 3.5, 0.22,
             font_size=10, bold=True, color=DARK_GREY)
    add_text(slide, body, x+0.2, y+0.38, 3.5, 0.5,
             font_size=9, color=MID_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, 13.33, 7.5, fill=DARK_GREY)
add_rect(slide, 0, 0, 4.5, 7.5, fill=ABBOTT_BLUE)
add_rect(slide, 4.5, 0, 0.06, 7.5, fill=ACCENT_RED)

add_text(slide, "CHINA", 0.4, 1.2, 3.8, 1.0,
         font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "VBP", 0.4, 2.1, 3.8, 0.9,
         font_size=52, bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER)
add_text(slide, "IMPACT", 0.4, 3.0, 3.8, 0.9,
         font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_text(slide, "Abbott Diagnostics", 5.0, 1.5, 7.9, 0.7,
         font_size=30, bold=True, color=WHITE)
add_text(slide, "Volume-Based Procurement\nDeep Dive & Headwind Assessment", 5.0, 2.3, 7.9, 1.0,
         font_size=20, color=RGBColor(0xBF, 0xD9, 0xF2))
add_text(slide, "How big is the hit?  What has cleared?  What remains?  When does it end?",
         5.0, 3.4, 7.9, 0.5, font_size=12, color=RGBColor(0x99, 0xAA, 0xBB), italic=True)
add_text(slide, "June 2026", 5.0, 6.5, 7.9, 0.4,
         font_size=12, color=MID_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — WHAT IS VBP? (PLAIN ENGLISH)
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "What Is VBP? — Plain English",
           "China's government forcing hospitals to buy lab test supplies at much lower prices via bulk tenders")
footer(slide)

add_rect(slide, 0.35, 1.25, 12.6, 1.0, fill=LIGHT_BLUE)
add_text(slide, "The Simple Version",
         0.55, 1.3, 4, 0.3, font_size=12, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "China's government (NHSA) organizes massive group-buying events for hospital supplies. "
    "A coalition of 20–28 provinces pools together and says: 'We will commit to buying X million tests per year — "
    "but only if you cut your price by 60–85%. Whoever offers the lowest price wins the contract.' "
    "This is called Volume-Based Procurement (VBP / 集采 in Chinese).",
    0.55, 1.6, 12.1, 0.6, font_size=11, color=DARK_GREY)

# Three columns: How it Works / Who Runs It / Why It Hurts MNCs
col_defs = [
    ("⚙️  How It Works", ABBOTT_BLUE, [
        "Provinces pool together (20–28 at a time)",
        "They tender for specific test reagents",
        "  (e.g. 'thyroid tests' or 'chemistry panels')",
        "Winning company gets ~80% of volume",
        "  across all hospitals in those provinces",
        "Price is fixed for 2–3 year contract",
        "Non-winners get only the leftover 20%",
        "Instruments are NOT included — just reagents",
        "  (the 'ink cartridge' part of the business)",
    ]),
    ("🏛️  Who Runs It", MID_BLUE, [
        "NOT a single national body",
        "7 lead provinces designated by NHSA:",
        "  Jiangxi, Anhui, Guangdong, Zhejiang,",
        "  Fujian, Henan, Hebei",
        "Each province leads a multi-province",
        "  alliance for specific categories",
        "Anhui pioneered this in 2021 with",
        "  immunoassay; other provinces followed",
        "Goal: 80% of IVD spending through VBP",
    ]),
    ("💥  Why It Hurts Multinationals", ACCENT_RED, [
        "MNCs (Abbott, Roche, Siemens) had the",
        "  highest pre-VBP prices → biggest cuts",
        "Domestic companies have lower cost",
        "  bases → can win at lower prices",
        "Scoring criteria favor domestic makers",
        "  (implicit government preference)",
        "Double whammy: price cuts AND volumes",
        "  often go to local rivals, not MNCs",
        "Abbott won ZERO of 13 lots it bid in",
        "  the key 23-province chemistry round",
    ]),
]
for i, (title, color, bullets) in enumerate(col_defs):
    x = 0.35 + i * 4.18
    add_rect(slide, x, 2.38, 4.0, 0.35, fill=color)
    add_text(slide, title, x+0.12, 2.42, 3.8, 0.28,
             font_size=12, bold=True, color=WHITE)
    add_rect(slide, x, 2.73, 4.0, 2.6, fill=LIGHT_GREY)
    add_lines(slide, bullets, x+0.15, 2.82, 3.7, 2.4, font_size=10)

# Bottom: DRG compound
add_rect(slide, 0.35, 5.43, 12.6, 1.4, fill=RGBColor(0xFF, 0xF0, 0xE8))
add_text(slide, "⚠️  The Compound Effect: DRG/DIP Reform",
         0.55, 5.49, 6, 0.3, font_size=12, bold=True, color=ACCENT_ORANGE)
add_text(slide,
    "At the same time as VBP is cutting per-test prices, China is also rolling out 'DRG/DIP' hospital payment reform. "
    "This caps what hospitals get paid per patient admission — so hospitals have an incentive to order fewer (expensive) tests per patient. "
    "Result: VBP cuts the price Abbott gets per test, while DRG/DIP cuts the number of tests ordered. "
    "Both hit at once — that's why the China Core Lab revenue declined 15–30% per quarter in 2025.",
    0.55, 5.82, 12.1, 0.93, font_size=11, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — HOW BIG IS ABBOTT'S CHINA BUSINESS?
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "How Big Is Abbott's China Diagnostics Business?",
           "Abbott doesn't disclose China diagnostics separately — here's the best available estimate")
footer(slide)

# Left: Known facts
add_rect(slide, 0.35, 1.25, 6.15, 5.85, fill=LIGHT_GREY)
add_text(slide, "What Abbott Does Disclose",
         0.55, 1.32, 5.5, 0.3, font_size=13, bold=True, color=ABBOTT_BLUE)

known_rows = [
    ["Total Abbott China Revenue", "All segments", "$2.25B", "$2.11B", "$1.87B"],
    ["  of Total Company Sales",   "",            "5.5%",   "5.0%",   "4.3%"],
    ["China VBP headwind",         "Core Lab only", "—",   "~$200M",  "~$400–500M"],
    ["VBP as % organic Dx growth", "",              "—",     "—",     "~750 bps drag"],
]

add_text(slide, "Metric", 0.55, 1.68, 3.0, 0.25, font_size=9, bold=True, color=MID_GREY)
add_text(slide, "2023",   3.6,  1.68, 0.95, 0.25, font_size=9, bold=True, color=MID_GREY, align=PP_ALIGN.CENTER)
add_text(slide, "2024",   4.55, 1.68, 0.95, 0.25, font_size=9, bold=True, color=MID_GREY, align=PP_ALIGN.CENTER)
add_text(slide, "2025",   5.5,  1.68, 0.9,  0.25, font_size=9, bold=True, color=MID_GREY, align=PP_ALIGN.CENTER)

add_rect(slide, 0.45, 1.93, 5.95, 0.02, fill=RGBColor(0xCC, 0xCC, 0xCC))

for r, (metric, note, v23, v24, v25) in enumerate(known_rows):
    y = 2.0 + r * 0.42
    bg = WHITE if r % 2 == 0 else LIGHT_GREY
    add_rect(slide, 0.45, y, 5.95, 0.4, fill=bg)
    add_text(slide, metric, 0.55, y+0.08, 3.0, 0.28, font_size=10, bold=(note==""))
    add_text(slide, v23, 3.6, y+0.08, 0.9, 0.28, font_size=10, align=PP_ALIGN.CENTER)
    add_text(slide, v24, 4.55, y+0.08, 0.9, 0.28, font_size=10, align=PP_ALIGN.CENTER)
    add_text(slide, v25, 5.5, y+0.08, 0.9, 0.28, font_size=10, align=PP_ALIGN.CENTER)

add_text(slide, "China diagnostics revenue is NOT disclosed separately.",
         0.55, 3.72, 5.7, 0.25, font_size=9, italic=True, color=ACCENT_ORANGE)

add_text(slide, "China Diagnostics Revenue — Triangulated Estimate",
         0.55, 4.05, 5.7, 0.28, font_size=11, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "Abbott disclosed a '$400–500M annual VBP headwind.' "
    "If that represents a ~50–70% price cut (consistent with disclosed price reductions), "
    "the pre-VBP China Core Lab reagent base was approximately $600M–$1B. "
    "The post-VBP steady-state base is estimated at $350–500M — "
    "permanently reset, will not recover to prior levels.",
    0.55, 4.38, 5.65, 1.6, font_size=10, color=DARK_GREY)

# Right: China products + how they're affected
add_text(slide, "What Abbott Sells in China — & VBP Exposure",
         6.65, 1.32, 6.4, 0.3, font_size=13, bold=True, color=ABBOTT_BLUE)

prod_rows = [
    ("Core Lab (Immunoassay)", "Alinity i / ARCHITECT", ACCENT_RED, "🔴 HIGH", "Tumor markers, thyroid, hormones, infection serology — all went through VBP"),
    ("Core Lab (Chemistry)", "Alinity c / ARCHITECT", ACCENT_RED, "🔴 HIGH", "Liver, kidney, lipids, glucose — 23-province round; Abbott won 0 of 13 lots"),
    ("Molecular Diagnostics", "m2000 / Alinity m", ACCENT_ORANGE, "🟡 MOD", "HPV went through Dec 2023 round; HIV/HCV smaller share in China"),
    ("Rapid Diagnostics", "BinaxNOW / Panbio", ACCENT_GREEN, "🟢 LOW", "Different channel (consumer/POC) — less VBP exposure; Q1 2026 down 10% on weak flu season"),
    ("Point of Care", "i-STAT", ACCENT_GREEN, "🟢 LOW", "Hospital bedside; smaller China revenue base; POCT not yet in alliance-scale VBP"),
    ("Fertility / Cancer markers", "ARCHITECT / Alinity i", ACCENT_ORANGE, "🟡 NEXT", "Pending — but Abbott has SMALL market share here. Management: 'not significant'"),
]

y = 1.68
for cat, platform, color, severity, note in prod_rows:
    add_rect(slide, 6.6, y, 0.75, 0.62, fill=color)
    add_text(slide, severity.split()[0], 6.62, y+0.18, 0.72, 0.28,
             font_size=12, align=PP_ALIGN.CENTER, color=WHITE)
    add_rect(slide, 7.35, y, 5.95, 0.62, fill=LIGHT_GREY if (prod_rows.index((cat, platform, color, severity, note)) % 2 == 0) else WHITE)
    add_text(slide, cat, 7.5, y+0.03, 5.7, 0.22, font_size=10, bold=True, color=DARK_GREY)
    add_text(slide, f"{platform}  ·  {note}", 7.5, y+0.27, 5.65, 0.32, font_size=8.5, color=MID_GREY)
    y += 0.65

add_rect(slide, 6.6, y+0.05, 6.7, 0.62, fill=LIGHT_BLUE)
add_text(slide, "💡  Abbott's own statement (Q1 2026):", 6.78, y+0.1, 5.5, 0.22,
         font_size=10, bold=True, color=ABBOTT_BLUE)
add_text(slide, '"About 80% of our portfolio has gone through VBP in China."  — CEO Robert Ford, April 2026',
         6.78, y+0.32, 6.3, 0.28, font_size=10, italic=True, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — DISCLOSURE TIMELINE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "How We Got Here — Abbott's VBP Disclosure Timeline",
           "From a quiet miss in Q3 2024 to a >$1B headwind admission by Q3 2025")
footer(slide)

# Horizontal timeline bar
add_rect(slide, 0.5, 3.55, 12.3, 0.06, fill=ABBOTT_BLUE)

events = [
    (0.5, "Q3 2024\nOct 2024", "First mention", "Core Lab missed expectations. 'VBP in China' blamed. Had expected impact in Q1 2024 — it arrived later.", ACCENT_ORANGE),
    (2.9, "Q4 2024\nJan 2025", "First dollar figure", "FIRST TIME Abbott puts a number on it: '$400–500M annual headwind from China VBP.' Stock fell.", ACCENT_ORANGE),
    (5.3, "Q2 2025\nJul 2025", "Guidance cut\nStock –8%", "Full-year 2025 diagnostic headwind = ~$700M (VBP + COVID decline). VBP = 750 bps drag on organic growth. Analysts cut targets.", ACCENT_RED),
    (7.7, "Q3 2025\nOct 2025", "Escalation\n'>$1B headwind'", "'Over $1 billion' combined headwind (VBP + COVID) for full-year 2025. 'Full lapping expected in 2026.' Dx segment –7.8% organic.", ACCENT_RED),
    (10.1, "Q4 2025\nJan 2026", "'Bulk of impact\nbehind us'", "Ford: 'We've gone through the bulk of our VBP impact.' China guided to 'single-digit decline' in 2026. No new pricing shock.", ACCENT_GREEN),
]

for x_pos, date, title, body, dot_color in events:
    # Dot
    dot = slide.shapes.add_shape(9, Inches(x_pos+0.98), Inches(3.42), Inches(0.25), Inches(0.25))
    dot.fill.solid()
    dot.fill.fore_color.rgb = dot_color
    dot.line.fill.background()
    # Date above line
    add_text(slide, date, x_pos+0.75, 2.85, 2.2, 0.5,
             font_size=9, bold=True, color=ABBOTT_BLUE, align=PP_ALIGN.CENTER)
    # Title + body below line
    add_text(slide, title, x_pos+0.75, 3.72, 2.2, 0.35,
             font_size=10, bold=True, color=dot_color, align=PP_ALIGN.CENTER)
    add_text(slide, body, x_pos+0.75, 4.1, 2.2, 1.2,
             font_size=8.5, color=MID_GREY, align=PP_ALIGN.CENTER)

# Q1 2026 milestone at end
add_rect(slide, 0.35, 5.45, 12.6, 1.8, fill=LIGHT_GREEN)
add_text(slide, "✅  Q1 2026 (April 2026) — The Inflection Point",
         0.55, 5.52, 7, 0.3, font_size=13, bold=True, color=ACCENT_GREEN)
add_text(slide,
    "China Core Lab revenue was FLAT year-over-year in Q1 2026 — the first quarter without a double-digit decline "
    "after five consecutive quarters of –15% to –30%. CEO Ford confirmed: 'About 80% of our portfolio has gone "
    "through VBP.' He said he is 'not expecting big growth — all I need for it is to be pretty stable.' "
    "Overall Core Lab grew +3% in Q1 2026; Diagnostics segment +1.8% comparable overall.",
    0.55, 5.88, 12.1, 1.3, font_size=11, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — VBP CATEGORIES: WHAT HAS CLEARED, WHAT'S COMING
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "VBP Rollout by Category — What Has Cleared vs. What's Next",
           "The worst categories for Abbott are largely done; remaining waves hit where Abbott has small share")
footer(slide)

add_text(slide, "Category-by-Category Status (as of mid-2026)", 0.4, 1.25, 12, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)

cat_rows = [
    ("Clinical Chemistry\n(Biochemistry)", "✅ DONE", ACCENT_GREEN, "68–95% price cuts", "23-prov. Jiangxi 2024 + follow-on 2025",
     "HIGH — Abbott won 0 of 13 lots in key round.\nDomestic rivals swept contracts."),
    ("Immunoassay / CLIA\n(Tumor markers, thyroid,\nhormones, infection)", "✅ DONE", ACCENT_GREEN, "75–85% price cuts", "Anhui 2021, 25-prov. Dec 2023,\n28-prov. Anhui 2024–25",
     "HIGH — Abbott won some lots in 28-province\nround (Snibe, Mindray also won)."),
    ("HPV / Molecular\n(specific assays)", "✅ DONE", ACCENT_GREEN, "Significant cuts", "Dec 2023 national round",
     "MODERATE — smaller revenue base."),
    ("Coagulation", "⚠️ PARTIAL", ACCENT_ORANGE, "Provincial only so far", "Anhui 2023 provincial;\nnot yet alliance-scale",
     "MODERATE — not yet at multi-province scale."),
    ("Hematology", "⚠️ ONGOING", ACCENT_ORANGE, "Selective provinces", "No full alliance round",
     "LOW — Sysmex/Mindray dominate.\nAbbott Alinity h is new."),
    ("Fertility / Sex Hormones\n& Cancer Markers", "🔜 NEXT", MID_BLUE, "TBD", "Anticipated but not announced",
     "SMALL — Abbott has low market share here.\nManagement: 'not significant for us.'"),
    ("POCT / Point of Care", "⏳ PENDING", MID_GREY, "Not yet", "No multi-province alliance",
     "MODERATE — but POCT format is different\nfrom hospital core lab."),
    ("High-complexity Molecular\n(NGS, liquid biopsy)", "🚫 EXCLUDED", ACCENT_GREEN, "Excluded from\nreimbursement 2025", "12 liquid biopsy assays excluded",
     "N/A — these are excluded from VBP scope."),
]

headers = ["Test Category", "Status", "Avg. Price Cut", "Key Round(s)", "Abbott Exposure"]
col_w = [2.6, 1.3, 1.35, 2.5, 5.2]
x = 0.35
for hdr, cw in zip(headers, col_w):
    add_rect(slide, x, 1.62, cw, 0.35, fill=ABBOTT_BLUE)
    add_text(slide, hdr, x+0.07, 1.66, cw-0.14, 0.28,
             font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    x += cw

for r, (cat, status, status_col, cut, rounds, exposure) in enumerate(cat_rows):
    y = 1.97 + r * 0.55
    bg = WHITE if r % 2 == 0 else LIGHT_GREY
    vals = [cat, status, cut, rounds, exposure]
    x = 0.35
    for i, (v, cw) in enumerate(zip(vals, col_w)):
        if i == 1:
            sc = status_col
            fill = RGBColor(min(sc[0] + 180, 255), min(sc[1] + 180, 255), min(sc[2] + 180, 255))
        else:
            fill = bg
        add_rect(slide, x, y, cw, 0.53, fill=fill,
                 line=RGBColor(0xCC, 0xCC, 0xCC))
        c = status_col if i == 1 else DARK_GREY
        b = (i == 1)
        add_text(slide, v, x+0.07, y+0.05, cw-0.14, 0.44,
                 font_size=8.5, color=c, bold=b)
        x += cw

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — PEER COMPARISON
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Abbott vs. Peers — China VBP Headwind Comparison",
           "All major lab diagnostics companies were hit — Abbott's was the biggest in absolute dollars")
footer(slide)

add_text(slide, "Confirmed Peer Disclosures", 0.4, 1.25, 12, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)

peer_data = [
    ("Abbott\nDiagnostics", ABBOTT_BLUE,
     "$400–500M", "FY2025", "China Core Lab: –15% to –30% each quarter of 2025",
     "~Single-digit decline\n(flat to slightly down)", "Q1 2026 = FLAT ✅\n'Bulk of impact behind us'",
     "H1–H2 2026"),
    ("Roche\nDiagnostics", MID_BLUE,
     "~CHF 500M+\n(CHF 136M/quarter\nat peak Q1 2025)", "FY2025", "Asia-Pacific –15%; China –23% in FY2025",
     "'Diminished but\ncontinuing headwinds'", "Q1 2026 CHF ~60M/quarter\n(down from CHF 136M peak)",
     "H2 2026 into 2027"),
    ("Danaher /\nBeckman Coulter", RGBColor(0x6B, 0x4C, 0xA0),
     "$150M", "FY2025", "Most quantified disclosure; $50M/yr initially then accelerated",
     "$75–100M\n(~50% reduction)", "Most advanced in\nthe base-effect cycle",
     "H1 2026 (mostly done)"),
    ("Siemens\nHealthineers", ACCENT_ORANGE,
     "~€100–120M\n(estimated)", "FY2025 (Sep yr-end)", "China Dx 'low double-digit' decline",
     "'Especially pronounced\nH1 FY2026'", "Behind Abbott &\nDanaher in the cycle",
     "H2 FY2026 (Apr–Sep)"),
]

col_labels = ["Company", "Peak Annual\nHeadwind", "Headwind Year", "What Happened", "2026 Guidance", "Current Status", "Base Effect\nRelief Expected"]
col_w = [1.5, 1.7, 1.3, 2.4, 1.9, 2.1, 1.8]

x = 0.35
for lbl, cw in zip(col_labels, col_w):
    add_rect(slide, x, 1.62, cw, 0.45, fill=ABBOTT_BLUE)
    add_text(slide, lbl, x+0.07, 1.65, cw-0.14, 0.4,
             font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    x += cw

for r, (name, color, headwind, yr, what, guide, status, relief) in enumerate(peer_data):
    y = 2.07 + r * 1.1
    bg = RGBColor(min(color[0]+185, 255), min(color[1]+185, 255), min(color[2]+185, 255))
    vals = [name, headwind, yr, what, guide, status, relief]
    x = 0.35
    for i, (v, cw) in enumerate(zip(vals, col_w)):
        cell_bg = bg if i == 0 else (WHITE if r % 2 == 0 else LIGHT_GREY)
        add_rect(slide, x, y, cw, 1.06, fill=cell_bg,
                 line=RGBColor(0xCC, 0xCC, 0xCC))
        c = WHITE if i == 0 else DARK_GREY
        add_text(slide, v, x+0.07, y+0.08, cw-0.14, 0.92,
                 font_size=9, bold=(i == 0), color=c)
        x += cw

add_rect(slide, 0.35, 6.5, 12.6, 0.68, fill=LIGHT_BLUE)
add_text(slide, "Key Takeaway from Peer Comparison:",
         0.55, 6.54, 4, 0.28, font_size=11, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "Abbott's headwind ($400–500M) is the largest in absolute dollars — consistent with having had a larger China Core Lab share. "
    "Danaher is furthest through the cycle. Roche and Siemens are still absorbing. "
    "The industry-wide pattern is clear: 2025 = peak pain, 2026 = transition, 2027 = clean comps (if no new VBP waves).",
    0.55, 6.82, 12.1, 0.33, font_size=10, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — THREE BUCKETS OF HEADWIND
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Breaking Down the Remaining Headwind — Three Buckets",
           "Not all of the China pain is the same — some is over, some is structural, some is still incoming")
footer(slide)

add_text(slide,
    "To assess remaining risk, split the China headwind into three separate components. "
    "They have different timelines and different magnitudes.",
    0.4, 1.25, 12.5, 0.45, font_size=11.5, color=DARK_GREY)

buckets = [
    (ACCENT_RED, "BUCKET 1", "Price Reset Shock",
     "✅ LARGELY COMPLETE",
     [
         "This is the initial 60–85% price cut on reagents when VBP hits",
         "~80% of Abbott's China Core Lab portfolio has been repriced",
         "Remaining ~20% (fertility, cancer markers) = Abbott has small",
         "  share → management says 'not significant'",
         "The prices WON'T recover — this is permanent",
         "But it stops being a YoY headwind once lapped",
     ],
     "$400–500M at peak;\nestimated $50–100M\nresidual from remaining\n20% of portfolio",
     "Largely done"),
    (ACCENT_ORANGE, "BUCKET 2", "Base Effect (Comp Normalization)",
     "🔄 IN PROGRESS — resolves through 2026",
     [
         "Even after prices reset, the YoY comparison stays negative",
         "  until you're comparing against a period that ALREADY had the",
         "  new low prices in the base",
         "2025 = every quarter was –15% to –30% China Core Lab",
         "Q1 2026 = first flat quarter (lapping Q1 2025's bad numbers)",
         "H2 2026 = best comp relief (lapping worst quarters of 2025)",
         "This is a mathematical effect — requires no business recovery",
     ],
     "Turns from headwind\nto neutral/tailwind\nas 2026 progresses",
     "Resolves H2 2026"),
    (ACCENT_RED, "BUCKET 3", "Structural / Volume Loss",
     "⚠️ ONGOING — not fully resolved",
     [
         "Abbott lost market share in the 23-province Jiangxi round",
         "  (won 0 of 13 lots) — that volume went to domestic rivals",
         "DRG/DIP payment reform continues to suppress test volumes",
         "  in Chinese hospitals regardless of VBP pricing",
         "Domestic companies (Mindray, Snibe, Maccura) are targeting",
         "  IVD share from ~10% → >20% — VBP accelerated this",
         "Abbott 'got the price cut with no volume offset'",
         "Management guides 'stable' — not 'recovering'",
     ],
     "Ongoing; no recovery\nexpected. Structural\nnew normal for China\nCore Lab.",
     "Permanent / ongoing"),
]

for i, (color, label, title, status, bullets, impact, timing) in enumerate(buckets):
    x = 0.35 + i * 4.24
    add_rect(slide, x, 1.78, 4.0, 0.35, fill=color)
    add_text(slide, f"{label}: {title}", x+0.12, 1.82, 3.8, 0.28,
             font_size=11, bold=True, color=WHITE)
    # Status bar
    stat_col = ACCENT_GREEN if "COMPLETE" in status else (ACCENT_ORANGE if "PROGRESS" in status else ACCENT_RED)
    add_rect(slide, x, 2.13, 4.0, 0.25, fill=stat_col)
    add_text(slide, status, x+0.12, 2.15, 3.8, 0.22,
             font_size=9, bold=True, color=WHITE)
    add_rect(slide, x, 2.38, 4.0, 2.85, fill=LIGHT_GREY)
    add_lines(slide, bullets, x+0.15, 2.45, 3.7, 2.6, font_size=9.5)

    add_rect(slide, x, 5.23, 4.0, 0.28, fill=RGBColor(0xDD, 0xDD, 0xDD))
    add_text(slide, f"Impact: {impact}", x+0.12, 5.25, 3.8, 0.25,
             font_size=8.5, bold=True, color=DARK_GREY)
    add_rect(slide, x, 5.51, 4.0, 0.25, fill=color)
    add_text(slide, f"Timeline: {timing}", x+0.12, 5.53, 3.8, 0.22,
             font_size=8.5, bold=True, color=WHITE)

add_rect(slide, 0.35, 5.85, 12.6, 1.38, fill=LIGHT_BLUE)
add_text(slide, "Bottom Line on Remaining Headwind:",
         0.55, 5.91, 5, 0.28, font_size=12, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "The acute shock (Bucket 1) is mostly done. The YoY comparisons get easier mechanically through H2 2026 (Bucket 2). "
    "What's permanent is the lower price level itself and some share loss to domestic rivals (Bucket 3). "
    "Management is correctly guiding 'stable' rather than 'recovery.' A true growth story in China Core Lab "
    "would require Abbott to win back volume on newly priced reagents — which is possible but not assumed in guidance.",
    0.55, 6.24, 12.1, 0.95, font_size=11, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — MONITORING SIGNALS & FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "How to Monitor the Remaining Headwind",
           "What to watch to assess whether the VBP recovery is on track — or if new risks emerge")
footer(slide)

# Left: monitoring signals table
add_text(slide, "Key Signals to Track", 0.4, 1.25, 6.2, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)

signals = [
    ("China Core Lab YoY growth",
     "Is the base effect working through? Q1 2026 = flat. Watch Q2–Q4 2026 for first positive prints.",
     "Q1 2026: FLAT ✅ — first non-negative quarter"),
    ("New VBP round announcements\n(Chinese NHSA / provincial)",
     "New categories entering procurement = new headwinds. Coagulation at scale, hematology, POCT are the next logical targets.",
     "Coagulation, POCT, hematology not yet at alliance scale"),
    ("Domestic rival disclosures\n(Mindray, Snibe)",
     "Mindray is targeting IVD share from ~10% → 20%+. Their China IVD revenue growth tells you how much share is flowing away from MNCs.",
     "Mindray IVD growing double-digits; structural risk ongoing"),
    ("Roche / Siemens China\nDx commentary",
     "Best read-across on pace of normalization. Roche confirmed CHF headwind dropping from ~CHF 136M/qtr to ~CHF 60M/qtr by Q1 2026.",
     "Roche: still 'diminished but continuing' in 2026"),
    ("Abbott Alinity i China\nlocal manufacturing ramp",
     "Abbott launched locally made Alinity i instruments in China (Hangzhou factory) in Oct 2024. Success here could lower cost basis and improve price competitiveness in future VBP rounds.",
     "Launched Oct 2024; early; medium-term strategic bet"),
    ("DRG/DIP policy updates",
     "If China relaxes hospital payment caps (DRG/DIP), that would boost test volumes without any new VBP needed.",
     "No relaxation signaled as of mid-2026"),
]

y = 1.62
for sig, why, current in signals:
    add_rect(slide, 0.35, y, 2.5, 0.72, fill=LIGHT_BLUE)
    add_text(slide, sig, 0.5, y+0.06, 2.3, 0.62, font_size=9, bold=True, color=ABBOTT_BLUE)
    add_rect(slide, 2.85, y, 3.5, 0.72, fill=LIGHT_GREY)
    add_text(slide, why, 2.98, y+0.06, 3.3, 0.62, font_size=8.5, color=DARK_GREY)
    add_rect(slide, 6.35, y, 0.12, 0.72, fill=ABBOTT_BLUE)
    add_text(slide, current, 6.5, y+0.06, 3.3, 0.62, font_size=8.5, color=MID_GREY, italic=True)
    y += 0.77

# Right: Expected quarterly comp pattern
add_text(slide, "Expected Comp Pattern — China Core Lab", 9.95, 1.25, 3.2, 0.3,
         font_size=11, bold=True, color=ABBOTT_BLUE)

comp_rows_data = [
    ("Q1 2025", "–20% to –25%", ACCENT_RED,   "Peak pain"),
    ("Q2 2025", "–15% to –20%", ACCENT_RED,   "Still deep negative"),
    ("Q3 2025", "–15% to –20%", ACCENT_RED,   "Continued decline"),
    ("Q4 2025", "–10% to –15%", ACCENT_ORANGE,"Early improvement"),
    ("Q1 2026", "~Flat ✅",      ACCENT_GREEN, "First comp relief (confirmed)"),
    ("Q2 2026", "+1% to +3%E",   ACCENT_GREEN, "Easier comp from Q2 2025"),
    ("Q3 2026", "+3% to +5%E",   ACCENT_GREEN, "Lapping worst Q3 2025"),
    ("Q4 2026", "+3% to +5%E",   ACCENT_GREEN, "Lapping worst Q4 2025"),
]
for r, (period, growth, color, note) in enumerate(comp_rows_data):
    y = 1.62 + r * 0.57
    add_rect(slide, 9.9, y, 1.2, 0.54, fill=LIGHT_GREY, line=RGBColor(0xCC,0xCC,0xCC))
    add_text(slide, period, 9.97, y+0.12, 1.1, 0.3, font_size=9, bold=True, color=DARK_GREY)
    light = RGBColor(min(color[0]+180, 255), min(color[1]+180, 255), min(color[2]+180, 255))
    add_rect(slide, 11.1, y, 1.6, 0.54, fill=light, line=RGBColor(0xCC,0xCC,0xCC))
    add_text(slide, growth, 11.17, y+0.12, 1.5, 0.3,
             font_size=9, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_rect(slide, 12.7, y, 0.3, 0.54, fill=color)

add_text(slide, "E = estimated; actual depends on new VBP rounds & respiratory season", 9.9, 7.1, 3.3, 0.2,
         font_size=7.5, color=MID_GREY, italic=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — BOTTOM LINE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, 13.33, 7.5, fill=DARK_GREY)
add_rect(slide, 0, 1.1, 13.33, 0.06, fill=ACCENT_RED)
add_rect(slide, 0, 6.3, 13.33, 0.06, fill=ACCENT_RED)

add_text(slide, "Bottom Line", 0.4, 0.2, 12.5, 0.75,
         font_size=36, bold=True, color=WHITE)

conclusions = [
    (ACCENT_GREEN,  "1",
     "The acute shock is over.",
     "~80% of China Core Lab reagents are now priced at post-VBP levels. "
     "Q1 2026 was the first quarter of flat China Core Lab revenue after 5 quarters of –15% to –30% declines. "
     "The incremental remaining rounds (fertility, cancer markers) hit where Abbott has small share — 'not significant.'"),
    (ACCENT_ORANGE, "2",
     "Prices will not recover — this is permanent.",
     "VBP prices don't bounce back. The $400–500M annual headwind represents a permanent reduction "
     "in the pricing of China Core Lab reagents. Management correctly guides 'stable' not 'recovery.' "
     "True China growth would require winning volume back from domestic rivals at the new lower prices."),
    (ACCENT_ORANGE, "3",
     "The 2026 story is a base-effect story, not a business recovery.",
     "H2 2026 will look great on a YoY basis because H2 2025 was terrible. "
     "That's mathematics, not a fundamental improvement. Clean, interpretable comps won't "
     "exist until 2027 — when all four quarters of 2026 are in the base."),
    (ACCENT_RED,    "4",
     "The structural headwind (volume loss, domestic share gain) continues.",
     "Abbott explicitly got 'the price cut with no volume offset.' Domestic rivals (Mindray, Snibe, Maccura) "
     "won VBP contracts and are growing rapidly. DRG/DIP hospital payment reform continues to suppress "
     "test volumes. These are ongoing, not temporary."),
    (MID_BLUE,      "5",
     "Watch for new VBP category rollouts — the main residual risk.",
     "Coagulation at multi-province scale, hematology alliances, and eventually POCT are the next targets. "
     "If any of these hit in a category where Abbott has significant share, the headwind cycle starts again. "
     "Roche and Siemens confirm the same risk in their own forward guidance language."),
]

for i, (color, num, title, body) in enumerate(conclusions):
    y = 1.25 + i * 0.98
    add_rect(slide, 0.3, y, 0.55, 0.88, fill=color)
    add_text(slide, num, 0.3, y+0.2, 0.55, 0.5,
             font_size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, 0.85, y, 12.15, 0.88, fill=RGBColor(0x3A, 0x3A, 0x3A))
    add_text(slide, title, 1.02, y+0.04, 11.7, 0.32,
             font_size=11, bold=True, color=color)
    add_text(slide, body, 1.02, y+0.38, 11.7, 0.48,
             font_size=9.5, color=RGBColor(0xCC, 0xCC, 0xCC))

prs.save("/home/user/knutnyman/Abbott_China_VBP_Analysis.pptx")
print("Deck 2 saved.")
