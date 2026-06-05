from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
ABBOTT_BLUE   = RGBColor(0x00, 0x47, 0x8A)   # Abbott brand blue
LIGHT_BLUE    = RGBColor(0xE8, 0xF2, 0xFB)
MID_BLUE      = RGBColor(0x4A, 0x90, 0xC4)
DARK_GREY     = RGBColor(0x2D, 0x2D, 0x2D)
MID_GREY      = RGBColor(0x5A, 0x5A, 0x5A)
LIGHT_GREY    = RGBColor(0xF4, 0xF4, 0xF4)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_GREEN  = RGBColor(0x00, 0x8A, 0x4B)
ACCENT_ORANGE = RGBColor(0xE8, 0x7D, 0x1E)
ACCENT_RED    = RGBColor(0xC0, 0x39, 0x2B)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]   # completely blank

# ── Helper functions ──────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None, line_w=Pt(0)):
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
             font_size=12, bold=False, color=DARK_GREY, align=PP_ALIGN.LEFT,
             italic=False, wrap=True):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def add_text_lines(slide, lines, l, t, w, h,
                   font_size=11, bold_first=False, color=DARK_GREY,
                   line_spacing=None):
    """lines = list of strings; first may be bolded"""
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
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
    # bottom accent line
    add_rect(slide, 0, 1.1, 13.33, 0.04, fill=ACCENT_ORANGE)

def footer(slide, text="Abbott Diagnostics — Confidential"):
    add_rect(slide, 0, 7.25, 13.33, 0.25, fill=LIGHT_GREY)
    add_text(slide, text, 0.3, 7.27, 12, 0.2,
             font_size=8, color=MID_GREY)

def card(slide, l, t, w, h, title, bullets, title_bg=ABBOTT_BLUE,
         title_color=WHITE, body_bg=LIGHT_GREY, bullet_size=10,
         title_size=12):
    add_rect(slide, l, t, w, 0.38, fill=title_bg)
    add_text(slide, title, l+0.1, t+0.04, w-0.2, 0.3,
             font_size=title_size, bold=True, color=title_color)
    add_rect(slide, l, t+0.38, w, h-0.38, fill=body_bg)
    add_text_lines(slide, bullets, l+0.12, t+0.44, w-0.24, h-0.5,
                   font_size=bullet_size)

def table_slide(slide, headers, rows, l, t, w, col_widths=None,
                header_bg=ABBOTT_BLUE, row_bg1=WHITE, row_bg2=LIGHT_GREY,
                font_size=10, header_font=11):
    n_cols = len(headers)
    if col_widths is None:
        col_widths = [w / n_cols] * n_cols
    row_h = 0.36
    # header row
    x = l
    for i, (hdr, cw) in enumerate(zip(headers, col_widths)):
        add_rect(slide, x, t, cw, row_h, fill=header_bg)
        add_text(slide, hdr, x+0.06, t+0.06, cw-0.12, row_h-0.08,
                 font_size=header_font, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        x += cw
    # data rows
    for r, row in enumerate(rows):
        y = t + row_h * (r + 1)
        bg = row_bg1 if r % 2 == 0 else row_bg2
        x = l
        for i, (cell, cw) in enumerate(zip(row, col_widths)):
            add_rect(slide, x, y, cw, row_h, fill=bg,
                     line=RGBColor(0xD0, 0xD0, 0xD0), line_w=Pt(0.5))
            add_text(slide, str(cell), x+0.06, y+0.05, cw-0.12, row_h-0.08,
                     font_size=font_size, align=PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT)
            x += cw

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, 13.33, 7.5, fill=ABBOTT_BLUE)
add_rect(slide, 0, 4.8, 13.33, 0.08, fill=ACCENT_ORANGE)
add_text(slide, "Abbott Diagnostics", 1, 1.2, 11, 1.4,
         font_size=48, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "Full Business & Market Mapping", 1, 2.7, 11, 0.7,
         font_size=28, color=RGBColor(0xBF, 0xD9, 0xF2), align=PP_ALIGN.CENTER)
add_text(slide, "Products  ·  Customers  ·  Competitors  ·  Financials", 1, 3.4, 11, 0.5,
         font_size=16, color=RGBColor(0x9B, 0xC4, 0xE8), align=PP_ALIGN.CENTER)
add_text(slide, "June 2026", 1, 5.2, 11, 0.4,
         font_size=13, color=RGBColor(0x9B, 0xC4, 0xE8), align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — WHAT IS ABBOTT DIAGNOSTICS?
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "What Is Abbott's Diagnostics Business?",
           "One of the world's top 2 diagnostics companies — everything from hospital lab machines to rapid home COVID tests")
footer(slide)

add_text(slide,
    "Abbott makes the tools that doctors and labs use to run medical tests. "
    "Think of it as the company behind the machine that reads your blood test at the hospital, "
    "the COVID rapid test at the pharmacy, and the device in the ER that checks your heart in 15 minutes.",
    0.4, 1.25, 12.5, 0.65, font_size=13, color=DARK_GREY)

# Four segment boxes
segs = [
    ("🏥  Core Laboratory", "$5.2B revenue", "Big hospital lab machines\n(chemistry, immunoassay, blood testing)", ABBOTT_BLUE),
    ("🧬  Molecular", "$0.5B revenue", "PCR / DNA-based tests\n(HIV viral load, STIs, HPV, COVID PCR)", MID_BLUE),
    ("⚡  Point of Care", "$0.6B revenue", "Handheld bedside devices\n(blood gas, troponin, concussion tests)", ACCENT_GREEN),
    ("📦  Rapid Diagnostics", "$3.0B revenue*", "Rapid strip tests for mass use\n(BinaxNOW COVID, flu, HIV, malaria)", ACCENT_ORANGE),
]

for i, (title, rev, desc, col) in enumerate(segs):
    x = 0.35 + i * 3.17
    add_rect(slide, x, 2.05, 3.0, 0.6, fill=col)
    add_text(slide, title, x+0.12, 2.1, 2.8, 0.3,
             font_size=13, bold=True, color=WHITE)
    add_text(slide, rev, x+0.12, 2.4, 2.8, 0.22,
             font_size=11, color=WHITE)
    add_rect(slide, x, 2.65, 3.0, 1.35, fill=LIGHT_GREY)
    add_text(slide, desc, x+0.15, 2.72, 2.7, 1.2, font_size=11, color=DARK_GREY)

add_text(slide, "* Peak COVID ~$8.5B in 2022; normalized back to ~$3B by 2024", 0.4, 4.08, 12, 0.3,
         font_size=9, color=MID_GREY, italic=True)

# Key stats row
add_rect(slide, 0.35, 4.45, 12.6, 0.04, fill=RGBColor(0xCC, 0xCC, 0xCC))
stats = [
    ("#2 Globally", "Behind only Roche"),
    ("$9.3B", "FY2024 Diagnostics Revenue"),
    ("22%", "of Total Abbott Sales"),
    ("~60%", "Revenue from International Markets"),
    ("4 Factories", "in China alone"),
]
for i, (val, lbl) in enumerate(stats):
    x = 0.4 + i * 2.52
    add_text(slide, val, x, 4.6, 2.4, 0.45,
             font_size=18, bold=True, color=ABBOTT_BLUE, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x, 5.05, 2.4, 0.3,
             font_size=9, color=MID_GREY, align=PP_ALIGN.CENTER)

add_rect(slide, 0.35, 5.42, 12.6, 1.6, fill=LIGHT_BLUE)
add_text(slide, "🔑  The Big Story", 0.55, 5.5, 4, 0.3, font_size=12, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "Abbott was the biggest beneficiary of the COVID testing boom ($8–10B at peak in 2021–22) "
    "and has spent the last 3 years 'normalizing' back to its base business. "
    "Its non-COVID diagnostics business has actually grown steadily throughout — it was just hidden by the COVID collapse. "
    "The March 2026 acquisition of Exact Sciences (Cologuard cancer test, $21B) is the next big chapter.",
    0.55, 5.82, 12.2, 1.1, font_size=11, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — FINANCIAL SNAPSHOT
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Revenue Snapshot: 2022 – 2025",
           "COVID inflated then collapsed — the underlying business is ~$8.6B and growing")
footer(slide)

# Revenue table
headers = ["Segment", "FY 2022", "FY 2023", "FY 2024", "FY 2025", "Trend"]
rows = [
    ["Core Laboratory",   "$5.1B", "$5.2B", "$5.2B", "~$5.2B",  "➡ Steady +3–5%/yr"],
    ["Rapid Diagnostics", "$8.6B", "$3.7B", "$3.0B", "~$2.5B",  "↘ COVID unwinding"],
    ["Molecular",         "$1.4B", "$0.6B", "$0.5B", "~$0.5B",  "➡ Stabilized"],
    ["Point of Care",     "$0.5B", "$0.6B", "$0.6B", "~$0.6B",  "↗ Steady growth"],
    ["TOTAL",             "$16.5B","$10.0B","$9.3B", "$8.9B",   ""],
    ["  of which: COVID", "$8.4B", "$1.6B", "$0.7B", "$0.3B",   "↘ Nearly gone"],
    ["  ex-COVID base",   "~$8.1B","~$8.4B","~$8.6B","~$8.6B", "↗ +6% since 2022"],
]
col_widths = [3.0, 1.5, 1.5, 1.5, 1.5, 3.8]
table_slide(slide, headers, rows, 0.35, 1.25, 12.6,
            col_widths=col_widths, font_size=10, header_font=11)

# Callout boxes below
callouts = [
    (ACCENT_ORANGE, "The COVID Illusion",
     "Revenue peaked at $16.5B in 2022 because COVID tests were "
     "$8B+ of sales. That money is now essentially gone ($300M left). "
     "Headline decline looks alarming — but the real business never stopped growing."),
    (ACCENT_GREEN, "Core Lab is the Engine",
     "The Core Laboratory segment ($5.2B) is the most important and "
     "most stable part of the business. It grows 3–5% per year organically "
     "and generates high recurring revenue from reagents (think ink cartridges for lab machines)."),
    (ABBOTT_BLUE, "What's Next: Exact Sciences",
     "Abbott bought Exact Sciences in March 2026 for $21B. This adds "
     "~$3B in cancer diagnostics revenue (Cologuard colon cancer test, "
     "Oncotype DX breast cancer test). Total diagnostics will cross $12B+ in 2026."),
]
for i, (col, title, text) in enumerate(callouts):
    x = 0.35 + i * 4.26
    add_rect(slide, x, 5.45, 4.0, 0.35, fill=col)
    add_text(slide, title, x+0.12, 5.49, 3.8, 0.28,
             font_size=11, bold=True, color=WHITE)
    add_rect(slide, x, 5.8, 4.0, 1.45, fill=LIGHT_GREY)
    add_text(slide, text, x+0.12, 5.86, 3.78, 1.3, font_size=9.5, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — CORE LABORATORY
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Core Laboratory  —  $5.2B",
           "The workhorse: big automated machines that process thousands of tests a day in hospital labs")
footer(slide)

# Left: what it does
add_rect(slide, 0.35, 1.25, 5.5, 5.7, fill=LIGHT_GREY)
add_text(slide, "What It Does", 0.55, 1.32, 5, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "Hospitals run hundreds of different blood tests every day — checking "
    "for diabetes, kidney disease, heart attacks, infections, hormones, and more. "
    "Abbott sells the automated analyzer machines and the chemical reagents (like "
    "ink for a printer) that run these tests. The machines are often placed for free "
    "or at low cost; Abbott makes its money on reagent consumables for years afterward.",
    0.55, 1.65, 5.15, 1.35, font_size=10.5, color=DARK_GREY)

add_text(slide, "Key Products (Alinity Platform)", 0.55, 3.05, 5.0, 0.28,
         font_size=12, bold=True, color=ABBOTT_BLUE)
products = [
    "Alinity ci  —  Chemistry + immunoassay combo (~200 tests, ~1,550/hr)",
    "Alinity i    —  Immunoassay only (cardiac, thyroid, hormones, tumor markers)",
    "Alinity c   —  Clinical chemistry (liver, kidney, lipids, glucose)",
    "Alinity h   —  Hematology / CBC (FDA cleared Aug 2023 — re-entry product)",
    "Alinity s   —  Blood bank serology (HIV, HCV, HBV — for donated blood)",
    "AlinIQ      —  Software/informatics layer — auto-verifies results, manages QC",
    "ARCHITECT  —  Legacy platform; still widely installed globally",
]
add_text_lines(slide, products, 0.55, 3.38, 5.15, 2.8, font_size=10)

add_text(slide, "Who Buys It", 0.55, 6.25, 5.0, 0.28,
         font_size=12, bold=True, color=ABBOTT_BLUE)
add_text_lines(slide, [
    "Hospital core labs (primary — IDNs like HCA, Ascension, CommonSpirit)",
    "Reference labs (Quest, LabCorp — large, price-sensitive, multi-vendor)",
    "Blood banks / plasma centers (for donor screening)",
], 0.55, 6.55, 5.15, 0.9, font_size=10)

# Right: competitors table
add_text(slide, "Who Abbott Competes Against", 5.95, 1.32, 7.0, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)
comp_headers = ["Category", "Key Rivals", "Abbott's Edge / Gap"]
comp_rows = [
    ["Immunoassay",
     "Roche (cobas e 801)\nSiemens (Atellica IM)\nBeckman (DxI 800)",
     "hs-Troponin I is the\nflagship cardiac assay;\nstrong clinical data"],
    ["Clinical Chemistry",
     "Roche (cobas 8000)\nSiemens (Atellica CH)\nBeckman (DxC 660i)",
     "Siemens has higher\nthroughput (1,800/hr);\nAbbott ~1,550/hr"],
    ["Integrated Platform",
     "Roche, Siemens Atellica\nSolution, Beckman,\nQuidelOrtho VITROS 5600",
     "Alinity ci competes\ndirectly; AlinIQ\nsoftware stickiness"],
    ["Hematology",
     "Sysmex (market leader\nglobally), Beckman,\nSiemens ADVIA",
     "Alinity h is new (2023);\nSysmex dominates;\nAbbott is the challenger"],
    ["Blood Screening",
     "Grifols (>80% US NAT)\nRoche, QuidelOrtho",
     "Abbott serology = ok;\nNAT = new (Alinity n)\nin development"],
]
table_slide(slide, comp_headers, comp_rows, 5.9, 1.65, 7.05,
            col_widths=[2.2, 2.45, 2.4], font_size=9, header_font=10)

add_rect(slide, 5.9, 5.72, 7.05, 1.18, fill=LIGHT_BLUE)
add_text(slide, "💡  Abbott's Biggest Differentiator", 6.1, 5.78, 6.5, 0.25,
         font_size=11, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "The Alinity platform covers chemistry, immunoassay, hematology, molecular, POC, and blood screening "
    "all under one software/UI. No rival does all of that on one harmonized system. "
    "This makes it easier to sell enterprise-wide contracts to hospital networks.",
    6.1, 6.06, 6.7, 0.78, font_size=10, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — MOLECULAR
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Molecular Diagnostics  —  ~$520M",
           "DNA/RNA-based tests: HIV viral load, STIs, HPV, cancer genetics — high complexity, high value")
footer(slide)

# Product cards
mol_cards = [
    ("Alinity m  (Flagship)", [
        "Fully automated PCR system",
        "Up to 1,080 tests per 24 hours",
        "True random-access — no batching needed",
        "75 STAT (urgent) tests per 8-hour shift",
        "One box replaces 5 older instruments",
        "Menu: HIV, HCV, HBV, STIs, HPV, COVID,",
        "  CMV/EBV/BKV (transplant), Flu/RSV",
    ]),
    ("m2000  (Legacy, Still Selling)", [
        "Older PCR platform — still widely used",
        "WHO-prequalified for HIV viral load",
        "Critical for PEPFAR HIV programs in Africa",
        "Generates sticky reagent revenue in",
        "  low/middle-income countries",
        "Slower being phased out for Alinity m",
    ]),
    ("Vysis FISH  (Cancer Genetics)", [
        "Not PCR — uses fluorescent DNA probes",
        "PathVysion: HER-2 in breast/gastric cancer",
        "  (determines who gets Herceptin)",
        "ALK probe: lung cancer targeted therapy",
        "UroVysion: bladder cancer detection",
        "Used in cancer centers / pathology labs",
    ]),
]
for i, (title, bullets) in enumerate(mol_cards):
    x = 0.35 + i * 4.18
    card(slide, x, 1.25, 3.95, 4.25, title, bullets,
         bullet_size=10, title_size=12)

# Key wins / gaps
add_rect(slide, 0.35, 5.6, 6.15, 1.65, fill=LIGHT_BLUE)
add_text(slide, "🏆  Key Competitive Wins", 0.55, 5.67, 5.5, 0.3,
         font_size=12, bold=True, color=ACCENT_GREEN)
add_text_lines(slide, [
    "✓  Only FDA-cleared test detecting CT, NG, TV & MG (4 STIs) simultaneously",
    "✓  HPV test for cervical cancer screening — FDA cleared Oct 2023 (attacks Hologic's turf)",
    "✓  CMV/EBV/BKV transplant virology from one plasma sample — reduces complexity",
    "✓  WHO prequalified m2000 = entrenched in PEPFAR/Global Fund HIV programs worldwide",
], 0.55, 6.0, 5.8, 1.2, font_size=10)

add_rect(slide, 6.6, 5.6, 6.38, 1.65, fill=RGBColor(0xFF, 0xF3, 0xE8))
add_text(slide, "⚠️  Gaps to Know", 6.8, 5.67, 5.8, 0.3,
         font_size=12, bold=True, color=ACCENT_ORANGE)
add_text_lines(slide, [
    "✗  No syndromic multiplex panels (bioMérieux BioFire FilmArray dominates — 132 pathogens)",
    "✗  No NGS/next-gen sequencing oncology offering (Roche Foundation Medicine dominates)",
    "✗  Blood donor NAT is Grifols territory — Abbott's new Alinity n is targeting this but still in ramp",
    "→  Exact Sciences (closed March 2026) fills the oncology molecular gap",
], 6.8, 6.0, 6.0, 1.2, font_size=10)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — POINT OF CARE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Point of Care  —  ~$590M",
           "Handheld & bedside testing: results in minutes without sending a sample to the central lab")
footer(slide)

add_text(slide,
    "Point of Care means testing happens at the patient's bedside — in the ICU, ER, operating room, "
    "or doctor's office. The key selling point: results in 2–15 minutes instead of hours.",
    0.4, 1.25, 12.6, 0.5, font_size=12, color=DARK_GREY)

# Products left column
add_text(slide, "Key Products", 0.4, 1.82, 5.5, 0.3, font_size=13, bold=True, color=ABBOTT_BLUE)

poc_products = [
    ("i-STAT Alinity  (Core Device)", "Handheld, wireless, touchscreen. Tests include blood gas, electrolytes, kidney/liver chemistry, cardiac troponin, coagulation, hemoglobin. One device, one cartridge — results in 2–10 min. Same cartridges as old i-STAT 1, so hospitals don't need to re-stock."),
    ("i-STAT hs-TnI  (Heart Attack Test)", "High-sensitivity troponin I — the test that rules in or out a heart attack at the bedside. FIRST whole-blood POC hs-troponin cleared in the US. Enables '0/1 hour' ER rule-out protocols without sending to central lab."),
    ("i-STAT TBI  (Concussion Test)", "GFAP + UCH-L1 biomarkers from whole blood — results in 15 min. FIRST FDA-cleared whole-blood bedside concussion test. Co-developed with the US Army. Cleared April 2024."),
    ("Piccolo Xpress  (Doctor's Office)", "Small benchtop analyzer for physician offices. 31 chemistry tests (liver, kidney, lipids, HbA1c) in 12 minutes. CLIA-waived = no lab license needed."),
    ("ID NOW  (Rapid Molecular)", "Isothermal amplification — faster than PCR. 2–13 min for COVID, Flu, RSV, Strep A. CLIA-waived. In physician offices, urgent care, pharmacies."),
]
y = 2.18
for title, desc in poc_products:
    add_rect(slide, 0.4, y, 5.7, 0.22, fill=ABBOTT_BLUE)
    add_text(slide, title, 0.55, y+0.02, 5.5, 0.2, font_size=10, bold=True, color=WHITE)
    add_text(slide, desc, 0.55, y+0.24, 5.55, 0.52, font_size=9, color=DARK_GREY)
    y += 0.78

# Right: customers + competitors
add_text(slide, "Who Buys It", 6.35, 1.82, 4, 0.3, font_size=13, bold=True, color=ABBOTT_BLUE)
customers = [
    ("Hospital ICUs", "Blood gas, electrolytes, lactate — 24/7 monitoring"),
    ("Emergency Departments", "Heart attack triage (hs-TnI), sepsis (lactate), concussion (TBI)"),
    ("Operating Rooms", "Coagulation during cardiac surgery — heparin dosing"),
    ("Physician Offices", "Piccolo chemistry panels, ID NOW respiratory tests"),
    ("Urgent Care", "ID NOW — same-visit prescription decision"),
    ("Military / Field", "i-STAT + TBI designed for forward surgical teams"),
]
y = 2.18
for setting, use in customers:
    add_rect(slide, 6.35, y, 6.6, 0.22, fill=LIGHT_BLUE)
    add_text(slide, f"📍 {setting}", 6.5, y+0.02, 3, 0.2, font_size=10, bold=True, color=ABBOTT_BLUE)
    add_text(slide, use, 6.5, y+0.24, 6.3, 0.22, font_size=9, color=DARK_GREY)
    y += 0.5

add_text(slide, "Top Competitors", 6.35, 5.25, 4, 0.3, font_size=13, bold=True, color=ABBOTT_BLUE)
comp_rows = [
    ["Blood Gas", "Siemens epoc (room-temp cartridges ✓), Radiometer ABL90"],
    ["Cardiac POC", "QuidelOrtho TriageTrue hs-TnI, Radiometer AQT90"],
    ["Rapid Molecular", "Cepheid GeneXpert (broader menu), Roche cobas Liat"],
    ["Office Chemistry", "Roche cobas b 101, Siemens DCA Vantage"],
]
table_slide(slide, ["Category", "Key Rivals"], comp_rows, 6.35, 5.58, 6.6,
            col_widths=[2.2, 4.4], font_size=9, header_font=10)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — RAPID DIAGNOSTICS
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Rapid Diagnostics  —  ~$2.5B (was $8.6B at COVID peak)",
           "Strip tests, NAAT rapid tests, cardiometabolic POC, and HIV/malaria tests for developing countries")
footer(slide)

add_text(slide,
    "This segment came primarily from Abbott's $5.3B acquisition of Alere in 2017. "
    "It covers everything from the COVID test you took at home to HIV rapid tests distributed by the Gates Foundation in Africa.",
    0.4, 1.25, 12.5, 0.45, font_size=11.5, color=DARK_GREY)

# Three columns
cols = [
    ("US / Developed Markets", ABBOTT_BLUE, [
        ("BinaxNOW  (Lateral Flow)", "COVID-19, Flu A+B, COVID+Flu combo, Strep A, RSV\nProfessional-use + OTC self-test\nBecame a household name during COVID"),
        ("ID NOW  (Rapid Molecular)", "2–13 min; COVID, Flu, RSV, Strep A\nISLA/LAMP tech — faster than PCR\nCLIA-waived for physician offices"),
        ("Afinion 2  (Cardiometabolic)", "HbA1c (diabetes), lipid panel, CRP, kidney\n~3 min per cartridge\nPhysician offices, diabetes clinics"),
        ("Toxicology  (Alere Legacy)", "iCup / iScreen — 3 to 12-panel urine drug screens\nPain clinics, occupational health, courts"),
    ]),
    ("Developing / Global Markets", MID_BLUE, [
        ("Panbio  (International Brand)", "Same as BinaxNOW but for non-US markets\n200M+ COVID tests shipped to 120 countries\nDengue, HIV Self Test, Malaria Ag"),
        ("Abbott-Bioline  (LMIC HIV/Malaria)", "HIV 1/2 rapid antibody tests\nWHO prequalified — required for\nGlobal Fund / PEPFAR procurement"),
        ("m-PIMA  (HIV Viral Load POC)", "WORLD'S FIRST WHO-prequalified\nPOC HIV viral load test\nBattery-powered — for rural African clinics"),
        ("Cholestech LDX", "Lipid + glucose; 5-min result\nOlder product but large installed base"),
    ]),
    ("Who Buys", ACCENT_GREEN, [
        ("Physician Offices (US)", "Largest US volume buyer\nBinaxNOW, ID NOW, Afinion, Toxicology"),
        ("Urgent Care & Retail Clinics", "ID NOW for same-visit prescribing\nMinuteClinic, independent pharmacies"),
        ("Retail Pharmacy (OTC)", "BinaxNOW at CVS, Walgreens, Walmart"),
        ("Global Health Programs", "PEPFAR, Global Fund, UNICEF, USAID\nBuy Panbio & Bioline in bulk\nWHO prequalification is the entry ticket"),
        ("Governments (stockpiling)", "US HHS/BARDA, EU governments bought\nhundreds of millions during COVID"),
    ]),
]

for i, (title, color, items) in enumerate(cols):
    x = 0.35 + i * 4.3
    add_rect(slide, x, 1.78, 4.1, 0.35, fill=color)
    add_text(slide, title, x+0.12, 1.83, 3.9, 0.28,
             font_size=12, bold=True, color=WHITE)
    y = 2.18
    for item_title, item_desc in items:
        add_rect(slide, x, y, 4.1, 0.2, fill=RGBColor(0xDD, 0xEA, 0xF5) if color == ABBOTT_BLUE else
                 RGBColor(0xD5, 0xE8, 0xF5) if color == MID_BLUE else RGBColor(0xD5, 0xEE, 0xE3))
        add_text(slide, item_title, x+0.1, y+0.01, 3.9, 0.18,
                 font_size=9.5, bold=True, color=color)
        add_text(slide, item_desc, x+0.1, y+0.21, 3.9, 0.5,
                 font_size=8.5, color=DARK_GREY)
        y += 0.75

# Bottom: key risks
add_rect(slide, 0.35, 6.55, 12.6, 0.7, fill=RGBColor(0xFF, 0xF0, 0xE8))
add_text(slide, "⚠️  Key Risks in This Segment", 0.55, 6.6, 4, 0.25,
         font_size=11, bold=True, color=ACCENT_ORANGE)
add_text(slide,
    "ID NOW sensitivity (~85–91% for COVID vs PCR's 95–99%) was controversial early on; "
    "Cepheid GeneXpert has a broader menu. "
    "Abbott-Bioline malaria test sensitivity questioned in 2025 publication — could affect WHO/Global Fund procurement. "
    "Roche acquired LumiraDx (July 2024) — now a stronger multi-analyte cardiometabolic POC rival.",
    0.55, 6.85, 12.2, 0.38, font_size=9.5, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — COMPETITIVE LANDSCAPE
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "The Competitive Landscape",
           "Abbott is #2 globally in IVD — Roche leads, Danaher/Siemens close behind")
footer(slide)

add_text(slide, "Global IVD Market  (~$82B in 2024, growing ~5–6%/yr)", 0.4, 1.25, 12, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)

comp_headers = ["Rank", "Company", "IVD Revenue", "Key Strengths", "Where They Beat Abbott"]
comp_rows = [
    ["#1", "Roche Diagnostics",     "~$14B",  "Everything — immunoassay, molecular, tissue pathology, cobas platform", "Higher throughput immunoassay; Foundation Medicine NGS; Tissue/pathology"],
    ["#2", "Abbott Diagnostics",    "$9.3B",  "Core Lab breadth (Alinity), POC leadership (i-STAT), Rapid Dx scale", "— (this is Abbott)"],
    ["#3", "Danaher\n(Beckman/Cepheid/Radiometer)", "~$9B", "Beckman clinical chemistry/IA, Cepheid molecular POC, Radiometer critical care", "Cepheid GeneXpert menu at POC; Sysmex hematology"],
    ["#4", "Siemens Healthineers",  "~$6–7B", "Atellica platform, high throughput (1,800 tests/hr chemistry)", "Throughput; imaging + lab integration"],
    ["#5", "Thermo Fisher",         "~$5–6B", "Molecular, mass spec, research-grade clinical", "Oncology molecular (pre-Exact Sciences)"],
    ["#6", "bioMérieux",            "~$4B",   "Blood culture, VITEK microbiology, BioFire syndromic panels", "Syndromic panels (BioFire) — Abbott has none"],
    ["#7", "Sysmex",                "~$3B",   "Hematology — dominant globally",                              "Hematology (Abbott's Alinity h is a re-entry)"],
    ["#8", "BD",                    "~$3B",   "Microbiology, molecular (BD MAX), Veritor rapid tests",        "Veritor reader vs BinaxNOW; BD MAX molecular panels"],
]
table_slide(slide, comp_headers, comp_rows, 0.35, 1.6, 12.6,
            col_widths=[0.6, 2.5, 1.4, 4.3, 3.8],
            font_size=9, header_font=10)

# Abbott's moat box
add_rect(slide, 0.35, 6.45, 7.5, 0.85, fill=LIGHT_BLUE)
add_text(slide, "Abbott's Structural Advantage: Breadth", 0.55, 6.5, 7, 0.28,
         font_size=11, bold=True, color=ABBOTT_BLUE)
add_text(slide,
    "No competitor matches Abbott's combination of Core Lab + POC + Rapid + Molecular + (now) Oncology "
    "all under one commercial umbrella. This enables enterprise-wide selling to hospital networks "
    "and creates significant switching costs through integrated informatics (AlinIQ).",
    0.55, 6.78, 7.1, 0.48, font_size=9.5, color=DARK_GREY)

add_rect(slide, 7.95, 6.45, 4.95, 0.85, fill=RGBColor(0xFF, 0xF3, 0xE8))
add_text(slide, "Abbott's Biggest Gap (Pre-Exact Sciences)", 8.1, 6.5, 4.7, 0.28,
         font_size=11, bold=True, color=ACCENT_ORANGE)
add_text(slide,
    "No next-generation sequencing (NGS) oncology offering. Roche (Foundation Medicine) "
    "and Thermo Fisher dominate cancer genomics. The Exact Sciences deal addresses this directly "
    "by adding Cologuard, Oncotype DX, and CancerGuard liquid biopsy.",
    8.1, 6.78, 4.7, 0.48, font_size=9.5, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — EXACT SCIENCES: THE NEXT CHAPTER
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "The Next Chapter: Exact Sciences Acquisition",
           "Closed March 23, 2026 · $21B deal · Adds ~$3B in fast-growing cancer diagnostics")
footer(slide)

add_text(slide,
    "Abbott had one major strategic gap: it had no meaningful oncology diagnostics business. "
    "Roche had Foundation Medicine (tumor profiling), Thermo Fisher had Ion Torrent (NGS). "
    "The Exact Sciences acquisition fills that hole and transforms Abbott's diagnostics mix.",
    0.4, 1.25, 12.5, 0.55, font_size=12, color=DARK_GREY)

# Three product boxes
exact_products = [
    ("Cologuard",
     "Colorectal Cancer Screening",
     ABBOTT_BLUE,
     [
         "Non-invasive stool DNA test for colon cancer",
         "Patient mails in a sample — no colonoscopy prep",
         "FDA approved; covered by Medicare/insurance",
         "~10M Americans screened so far; enormous TAM",
         "Revenue: $1B+ and growing high-teens annually",
         "Main competition: colonoscopy (procedure-based),",
         "  Guardant Shield (liquid biopsy — emerging rival)",
     ]),
    ("Oncotype DX",
     "Breast & Prostate Cancer Precision",
     ACCENT_GREEN,
     [
         "Genomic test on tumor tissue",
         "Tells oncologists whether a breast cancer patient",
         "  needs chemotherapy or can skip it",
         "Changes treatment decisions in ~30% of cases",
         "Reimbursed in US and many EU countries",
         "~$600–800M revenue, growing steadily",
         "Used in: breast cancer (primary), prostate, colon",
     ]),
    ("CancerGuard",
     "Multi-Cancer Early Detection",
     ACCENT_ORANGE,
     [
         "Liquid biopsy — a simple blood test",
         "Screens for signals of 50+ cancer types",
         "  from a single blood draw",
         "Launched 2025 — very early stage revenue",
         "The 'holy grail' of diagnostics if it works at scale",
         "Competes with: Grail (Galleri test), Guardant,",
         "  Foundation Medicine (Roche)",
         "Market could be $10B+ if MCED becomes standard of care",
     ]),
]

for i, (title, subtitle, color, bullets) in enumerate(exact_products):
    x = 0.35 + i * 4.24
    add_rect(slide, x, 1.9, 4.0, 0.52, fill=color)
    add_text(slide, title, x+0.15, 1.93, 3.8, 0.3,
             font_size=15, bold=True, color=WHITE)
    add_text(slide, subtitle, x+0.15, 2.27, 3.8, 0.22,
             font_size=10, color=WHITE, italic=True)
    add_rect(slide, x, 2.42, 4.0, 3.2, fill=LIGHT_GREY)
    add_text_lines(slide, bullets, x+0.18, 2.52, 3.7, 3.0, font_size=10)

# Impact callout
add_rect(slide, 0.35, 5.72, 12.6, 1.55, fill=LIGHT_BLUE)
add_text(slide, "What This Changes for Abbott", 0.55, 5.78, 8, 0.3,
         font_size=13, bold=True, color=ABBOTT_BLUE)
impact_items = [
    ("Revenue", "Total Diagnostics crosses $12B+ in 2026 — closes the gap with Roche faster than organic growth could"),
    ("Growth Rate", "Exact Sciences grew at ~15–18%/yr. Adds a high-growth engine alongside stable Core Lab"),
    ("Oncology Gap", "Abbott now has a credible cancer diagnostics portfolio — the #1 strategic hole is plugged"),
    ("Reporting Change", "New segment 'Cancer Diagnostics' added in Q1 2026; Rapid+Molecular+POC consolidated into one line"),
]
for i, (label, text) in enumerate(impact_items):
    x = 0.55 + (i % 2) * 6.3
    y = 6.12 + (i // 2) * 0.5
    add_text(slide, f"▸  {label}: ", x, y, 1.4, 0.45, font_size=10, bold=True, color=ABBOTT_BLUE)
    add_text(slide, text, x+1.35, y, 4.7, 0.45, font_size=10, color=DARK_GREY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — SUMMARY / TAKEAWAYS
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
header_bar(slide, "Key Takeaways",
           "What to remember about Abbott's diagnostics business")
footer(slide)

takeaways = [
    (ABBOTT_BLUE, "1", "Core Lab is the stable foundation",
     "The ~$5.2B Core Laboratory business (hospital analyzers and reagents) grows 3–5% per year organically "
     "and is relatively predictable. It is Abbott's highest-quality recurring revenue stream — "
     "think of it like a printer business: place the machine, sell the ink for 5–7 years."),
    (ACCENT_ORANGE, "2", "COVID inflated and deflated — the base business never stopped growing",
     "Revenue looked like it collapsed from $16.5B to $8.9B — but that's almost entirely COVID test revenue disappearing. "
     "Strip out COVID and the underlying business grew ~6% over 2022–2025. "
     "The story is now about the ex-COVID trajectory, not the headline number."),
    (ACCENT_GREEN, "3", "China VBP was the big 2024–2025 headwind — largely behind us",
     "China's government forced 60–85% price cuts on lab reagents through 'volume-based procurement.' "
     "This created a $400–500M annual drag on Core Lab. The worst is over (Q1 2026 China was flat). "
     "Prices won't recover, but the year-over-year comparison gets easier through H2 2026."),
    (MID_BLUE, "4", "Exact Sciences transforms the growth profile",
     "The $21B Exact Sciences acquisition (March 2026) adds Cologuard (colon cancer screening), "
     "Oncotype DX (breast cancer precision medicine), and CancerGuard (liquid biopsy). "
     "This fills Abbott's biggest strategic gap and pushes total diagnostics to $12B+ with a faster growth rate."),
    (RGBColor(0x6B, 0x4C, 0xA0), "5", "Abbott is a true platform company — breadth is the moat",
     "No competitor matches Abbott's ability to serve one hospital across Core Lab, POC (i-STAT), "
     "Rapid testing (BinaxNOW, ID NOW), Molecular (Alinity m), and now Oncology. "
     "The Alinity ecosystem + AlinIQ informatics creates significant switching costs."),
]

y = 1.3
for col, num, title, text in takeaways:
    add_rect(slide, 0.35, y, 0.55, 0.78, fill=col)
    add_text(slide, num, 0.35, y+0.18, 0.55, 0.5,
             font_size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, 0.9, y, 12.05, 0.78, fill=LIGHT_GREY)
    add_text(slide, title, 1.05, y+0.04, 11.5, 0.28,
             font_size=12, bold=True, color=col)
    add_text(slide, text, 1.05, y+0.34, 11.5, 0.42, font_size=10, color=DARK_GREY)
    y += 0.88

prs.save("/home/user/knutnyman/Abbott_Diagnostics_Overview.pptx")
print("Deck 1 saved.")
