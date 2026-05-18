"""
Abbott Cardiac Rhythm Management (CRM) – Detailed Market Model
Generates a multi-tab Excel workbook with historical data, projections,
competitive analysis, geographic split, margin model, and scenario analysis.
"""

import xlsxwriter
import os

OUTPUT = "/home/user/knutnyman/Abbott_CRM_Market_Model.xlsx"

# ── Palette ────────────────────────────────────────────────────────────────
ABT_BLUE       = "#0066CC"
ABT_DARK_BLUE  = "#003366"
ABT_MID_BLUE   = "#1A7FD4"
ABT_LIGHT_BLUE = "#E6F2FB"
HEADER_BG      = "#1A3A5C"
ROW_ALT        = "#F0F7FF"
ROW_WHITE      = "#FFFFFF"
MED_GRAY       = "#D0D7E0"
LIGHT_GRAY     = "#F5F6F8"
DARK_TEXT      = "#1A1A2E"
MID_TEXT       = "#404040"
SUBHEAD_BG     = "#2E5F8A"
SECTION_BG     = "#D6E8F7"
GREEN          = "#1A6B3C"
GREEN_LIGHT    = "#E6F4EE"
RED_VAL        = "#C0392B"
RED_LIGHT      = "#FDECEA"
AMBER          = "#D4760A"
AMBER_LIGHT    = "#FEF3E0"
COMP_GRAY      = "#6B7280"
COMP_LIGHT     = "#F3F4F6"
BULL_GREEN     = "#166534"
BEAR_RED       = "#991B1B"
BASE_BLUE      = "#1E40AF"

wb = xlsxwriter.Workbook(OUTPUT)

# ── Global format helpers ──────────────────────────────────────────────────
def fmt(bold=False, italic=False, size=10, color=DARK_TEXT, bg=None,
        border=0, align="left", valign="vcenter", wrap=False,
        num_format=None, font="Calibri"):
    d = {"font_name": font, "font_size": size, "font_color": color,
         "bold": bold, "italic": italic, "align": align, "valign": valign,
         "text_wrap": wrap, "border": border}
    if bg:    d["bg_color"] = bg
    if num_format: d["num_format"] = num_format
    return wb.add_format(d)

def money(bg=None, bold=False, border=0):
    return fmt(bold=bold, bg=bg, border=border, align="right",
               num_format='#,##0')

def money1(bg=None, bold=False, border=0):
    return fmt(bold=bold, bg=bg, border=border, align="right",
               num_format='#,##0.0')

def pct(bg=None, bold=False, border=0):
    return fmt(bold=bold, bg=bg, border=border, align="right",
               num_format='0.0%')

def pct1(bg=None, bold=False, border=0):
    return fmt(bold=bold, bg=bg, border=border, align="right",
               num_format='0.1%')

def mlt(bg=None, bold=False):
    return fmt(bold=bold, bg=bg, align="right", num_format='0.0x')

# Pre-bake common formats
F = {
    "cover_title": fmt(bold=True,  size=28, color="#FFFFFF", bg=HEADER_BG, align="center"),
    "cover_sub":   fmt(bold=False, size=14, color=ABT_LIGHT_BLUE, bg=HEADER_BG, align="center"),
    "cover_tag":   fmt(bold=False, size=11, color="#B0C8E0", bg=HEADER_BG, align="center"),
    "tab_head":    fmt(bold=True,  size=10, color="#FFFFFF", bg=HEADER_BG, border=1, align="center"),
    "tab_head_l":  fmt(bold=True,  size=10, color="#FFFFFF", bg=HEADER_BG, border=1, align="left"),
    "sub_head":    fmt(bold=True,  size=10, color="#FFFFFF", bg=SUBHEAD_BG, border=1, align="left"),
    "sub_head_c":  fmt(bold=True,  size=10, color="#FFFFFF", bg=SUBHEAD_BG, border=1, align="center"),
    "sec_head":    fmt(bold=True,  size=10, color=HEADER_BG, bg=SECTION_BG, border=1, align="left"),
    "sec_head_c":  fmt(bold=True,  size=10, color=HEADER_BG, bg=SECTION_BG, border=1, align="center"),
    "label":       fmt(size=10,  color=MID_TEXT, bg=ROW_WHITE, border=1),
    "label_alt":   fmt(size=10,  color=MID_TEXT, bg=ROW_ALT,   border=1),
    "label_b":     fmt(bold=True, size=10, color=DARK_TEXT, bg=ROW_WHITE, border=1),
    "label_b_alt": fmt(bold=True, size=10, color=DARK_TEXT, bg=ROW_ALT,   border=1),
    "label_i":     fmt(italic=True, size=10, color=COMP_GRAY, bg=ROW_WHITE, border=1),
    "num":         money(bg=ROW_WHITE, border=1),
    "num_alt":     money(bg=ROW_ALT,   border=1),
    "num_b":       money(bg=ROW_WHITE, bold=True, border=1),
    "num_b_alt":   money(bg=ROW_ALT,   bold=True, border=1),
    "pct":         pct(bg=ROW_WHITE, border=1),
    "pct_alt":     pct(bg=ROW_ALT,   border=1),
    "pct_b":       pct(bg=ROW_WHITE, bold=True, border=1),
    "pct_b_alt":   pct(bg=ROW_ALT,   bold=True, border=1),
    "grn":         money(bg=GREEN_LIGHT, border=1),
    "grn_b":       money(bg=GREEN_LIGHT, bold=True, border=1),
    "grn_p":       pct(bg=GREEN_LIGHT, border=1),
    "red":         money(bg=RED_LIGHT, border=1),
    "red_p":       pct(bg=RED_LIGHT, border=1),
    "amb":         money(bg=AMBER_LIGHT, border=1),
    "amb_p":       pct(bg=AMBER_LIGHT, border=1),
    "input":       fmt(size=10, color=DARK_TEXT, bg="#FFFACD", border=1, align="right",
                       num_format='#,##0.0'),
    "input_pct":   fmt(size=10, color=DARK_TEXT, bg="#FFFACD", border=1, align="right",
                       num_format='0.0%'),
    "note":        fmt(italic=True, size=9, color=COMP_GRAY),
    "spacer":      fmt(bg="#FFFFFF"),
}

# ── Shared data ────────────────────────────────────────────────────────────
HIST_YEARS  = [2020, 2021, 2022, 2023, 2024]
PROJ_YEARS  = [2025, 2026, 2027, 2028, 2029]
ALL_YEARS   = HIST_YEARS + PROJ_YEARS

# Abbott CRM Revenue (US$M) — split by product family
# Source: ABT 10-Ks, earnings calls, sell-side estimates
ABT_BRADY = {  # Bradycardia (pacemakers, leadless, CRT-P)
    2020: 1010, 2021: 1080, 2022: 1055, 2023: 1020, 2024: 1015,
    2025: 1085, 2026: 1175, 2027: 1280, 2028: 1370, 2029: 1455,
}
ABT_TACHY = {  # Tachycardia (ICD, CRT-D, S-ICD)
    2020: 1090, 2021: 1240, 2022: 1185, 2023: 1150, 2024: 1160,
    2025: 1200, 2026: 1245, 2027: 1295, 2028: 1345, 2029: 1395,
}
ABT_MON = {    # Monitoring (Confirm Rx ICM, Merlin remote)
    2020:  90, 2021: 112, 2022: 125, 2023: 140, 2024: 148,
    2025: 165, 2026: 182, 2027: 200, 2028: 220, 2029: 240,
}

ABT_TOTAL = {y: ABT_BRADY[y] + ABT_TACHY[y] + ABT_MON[y] for y in ALL_YEARS}

# Sub-product detail — Bradycardia
BRADY_TRAD  = {2020:840,2021:870,2022:830,2023:785,2024:750,
               2025:720,2026:685,2027:650,2028:615,2029:580}
BRADY_AVEIR = {2020:  0,2021:  0,2022: 10,2023: 65,2024: 98,
               2025:175,2026:280,2027:395,2028:500,2029:590}
BRADY_CRTP  = {2020:170,2021:210,2022:215,2023:170,2024:167,
               2025:190,2026:210,2027:235,2028:255,2029:285}

# Sub-product detail — Tachycardia
TACHY_ICD   = {2020:485,2021:530,2022:505,2023:490,2024:495,
               2025:510,2026:525,2027:540,2028:555,2029:570}
TACHY_CRTD  = {2020:545,2021:635,2022:600,2023:575,2024:575,
               2025:590,2026:600,2027:615,2028:625,2029:635}
TACHY_SICD  = {2020: 60,2021: 75,2022: 80,2023: 85,2024: 90,
               2025: 100,2026:120,2027:140,2028:165,2029:190}

# Geographic split
ABT_US    = {y: round(ABT_TOTAL[y]*0.415) for y in HIST_YEARS}
ABT_US.update({2025:round(ABT_TOTAL[2025]*0.420),2026:round(ABT_TOTAL[2026]*0.422),
               2027:round(ABT_TOTAL[2027]*0.425),2028:round(ABT_TOTAL[2028]*0.428),
               2029:round(ABT_TOTAL[2029]*0.430)})
ABT_INTL  = {y: ABT_TOTAL[y] - ABT_US[y] for y in ALL_YEARS}

# Gross margin (%) assumptions
GM = {2020:.655,2021:.668,2022:.662,2023:.658,2024:.660,
      2025:.665,2026:.670,2027:.675,2028:.678,2029:.680}

# R&D % of segment revenue (allocated CRM share)
RD_PCT = {y:.115 for y in HIST_YEARS}
RD_PCT.update({y:.112 for y in PROJ_YEARS})

# SG&A % of segment revenue
SGA_PCT = {y:.280 for y in HIST_YEARS}
SGA_PCT.update({2025:.272,2026:.268,2027:.262,2028:.258,2029:.255})

# Competitor revenue estimates (US$M)
MDT_CRM  = {2020:2440,2021:2620,2022:2450,2023:2310,2024:2280,
            2025:2310,2026:2360,2027:2420,2028:2480,2029:2550}
BSC_CRM  = {2020:1100,2021:1225,2022:1190,2023:1215,2024:1275,
            2025:1345,2026:1415,2027:1490,2028:1565,2029:1645}
BTK_CRM  = {2020: 380,2021: 400,2022: 405,2023: 410,2024: 415,
            2025: 425,2026: 435,2027: 445,2028: 455,2029: 465}
OTH_CRM  = {y: 130 for y in ALL_YEARS}

MARKET    = {y: ABT_TOTAL[y]+MDT_CRM[y]+BSC_CRM[y]+BTK_CRM[y]+OTH_CRM[y] for y in ALL_YEARS}
ABT_SHR   = {y: ABT_TOTAL[y]/MARKET[y] for y in ALL_YEARS}
MDT_SHR   = {y: MDT_CRM[y]/MARKET[y] for y in ALL_YEARS}
BSC_SHR   = {y: BSC_CRM[y]/MARKET[y] for y in ALL_YEARS}


# ══════════════════════════════════════════════════════════════════════════
# SHEET 1 — COVER
# ══════════════════════════════════════════════════════════════════════════
def make_cover():
    ws = wb.add_worksheet("Cover")
    ws.set_tab_color(ABT_BLUE)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:I", 14)
    ws.set_column("J:J", 2)

    # Banner rows
    for r in range(0, 12):
        ws.set_row(r, 30 if r not in (2,3,5,7,9) else 45)
        for c in range(1, 9):
            ws.write(r, c, None, wb.add_format({"bg_color": HEADER_BG}))

    # Title
    title_fmt = wb.add_format({"bold":True,"font_size":30,"font_color":"#FFFFFF",
                                "bg_color":HEADER_BG,"align":"center","valign":"vcenter",
                                "font_name":"Calibri"})
    sub_fmt   = wb.add_format({"bold":False,"font_size":13,"font_color":"#A8CCEE",
                                "bg_color":HEADER_BG,"align":"center","valign":"vcenter",
                                "font_name":"Calibri"})
    tag_fmt   = wb.add_format({"bold":False,"font_size":11,"font_color":"#7EAAD4",
                                "bg_color":HEADER_BG,"align":"center","valign":"vcenter",
                                "font_name":"Calibri"})

    ws.merge_range("B3:I3", "ABBOTT LABORATORIES", title_fmt)
    ws.merge_range("B4:I4", "Cardiac Rhythm Management — Detailed Market Model", sub_fmt)
    ws.merge_range("B6:I6", "Segment Revenue · Market Share · Competitive Landscape · Margins · Scenarios", tag_fmt)

    # Date / disclaimer row
    date_fmt = wb.add_format({"bold":False,"font_size":10,"font_color":"#7EAAD4",
                               "bg_color":HEADER_BG,"align":"center","valign":"vcenter",
                               "font_name":"Calibri"})
    ws.merge_range("B8:I8", "As of May 2026  |  All figures in USD millions unless noted", date_fmt)
    ws.merge_range("B10:I10",
        "This model is for informational purposes only. Estimates represent analyst projections "
        "based on public filings, earnings calls, and sell-side research.", date_fmt)

    # White space + key stats summary box
    ws.set_row(12, 10)
    ws.set_row(13, 30)
    ws.set_row(14, 8)

    box_hdr = wb.add_format({"bold":True,"font_size":10,"font_color":"#FFFFFF",
                              "bg_color":SUBHEAD_BG,"border":1,"align":"center",
                              "valign":"vcenter","font_name":"Calibri"})
    box_val = wb.add_format({"bold":True,"font_size":16,"font_color":ABT_BLUE,
                              "bg_color":"#FFFFFF","border":1,"align":"center",
                              "valign":"vcenter","font_name":"Calibri"})
    box_lbl = wb.add_format({"bold":False,"font_size":9,"font_color":MID_TEXT,
                              "bg_color":LIGHT_GRAY,"border":1,"align":"center",
                              "valign":"vcenter","font_name":"Calibri"})

    stats = [
        ("2024A CRM Revenue", "$2,323M"),
        ("2024A Market Share", "~35%"),
        ("Market Rank", "#1 / #2"),
        ("2024–29 Rev. CAGR", "~4.4%"),
        ("2029E CRM Revenue", "$3,090M"),
        ("Key Product", "Aveir™ Leadless"),
    ]
    cols = [1,2,3,4,5,6,7,8]
    # 3 stats per row, 2 rows
    for i, (lbl, val) in enumerate(stats):
        row_off = i // 3
        col_off = (i % 3) * 2
        c1 = cols[col_off]; c2 = cols[col_off+1]
        r_hdr = 13 + row_off*3
        r_val = 14 + row_off*3
        r_lbl = 15 + row_off*3
        ws.set_row(r_hdr, 22); ws.set_row(r_val, 32); ws.set_row(r_lbl, 20)
        ws.merge_range(r_hdr,c1,r_hdr,c2, lbl,  box_hdr)
        ws.merge_range(r_val,c1,r_val,c2, val,  box_val)
        ws.merge_range(r_lbl,c1,r_lbl,c2, "",   box_lbl)

    ws.set_row(22, 12)
    # TOC
    toc_hdr = wb.add_format({"bold":True,"font_size":11,"font_color":"#FFFFFF",
                              "bg_color":SUBHEAD_BG,"font_name":"Calibri"})
    toc_item = wb.add_format({"font_size":10,"font_color":ABT_DARK_BLUE,
                               "underline":True,"font_name":"Calibri"})
    toc_desc = wb.add_format({"font_size":10,"font_color":MID_TEXT,"font_name":"Calibri"})
    ws.set_row(23, 22)
    ws.merge_range("B24:I24", "  Table of Contents", toc_hdr)

    contents = [
        ("Market Overview",     "Global CRM market sizing, growth drivers, segmentation"),
        ("Revenue Model",       "Abbott CRM revenue — historical 2020–2024A & projections 2025–2029E"),
        ("Segment Deep Dive",   "Bradycardia, Tachycardia, Monitoring — product-level breakdown"),
        ("Geographic Split",    "US vs. International revenue, market dynamics"),
        ("Competitive Landscape","Medtronic, Boston Scientific, Biotronik — share analysis"),
        ("Margin Model",        "Gross margin, R&D, SG&A, segment operating income bridge"),
        ("Scenario Analysis",   "Bull / Base / Bear — key swing factors & sensitivity"),
        ("Assumptions",         "All model inputs — editable yellow cells drive projections"),
    ]
    for i, (tab, desc) in enumerate(contents):
        r = 24 + i + 1
        ws.set_row(r, 20)
        ws.write(r, 1, f"  {i+1}.  {tab}", toc_item)
        ws.merge_range(r, 2, r, 8, desc, toc_desc)

    ws.set_row(33, 14)
    disc_fmt = wb.add_format({"font_size":8,"italic":True,"font_color":"#888888","font_name":"Calibri"})
    ws.merge_range("B35:I35",
        "Sources: Abbott 10-K/Q filings, earnings call transcripts, AdvaMed, MedTech Dive, "
        "sell-side consensus estimates (Goldman Sachs, Morgan Stanley, UBS). "
        "Competitor figures are public estimates.", disc_fmt)


# ══════════════════════════════════════════════════════════════════════════
# SHEET 2 — MARKET OVERVIEW
# ══════════════════════════════════════════════════════════════════════════
def make_market_overview():
    ws = wb.add_worksheet("Market Overview")
    ws.set_tab_color(ABT_MID_BLUE)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 34)
    ws.set_column("C:L", 11)
    ws.set_column("M:M", 2)

    def hrow(r, title):
        ws.set_row(r, 22)
        ws.merge_range(r, 1, r, 11, f"  {title}", F["sub_head"])

    def sheader(r, *labels):
        ws.set_row(r, 18)
        for i, lbl in enumerate(labels):
            ws.write(r, 1+i, lbl, F["tab_head"] if i else F["tab_head_l"])

    def spacer(r, h=8):
        ws.set_row(r, h)

    row = 0
    # ── Page Title ──────────────────────────────────────────────────────
    ws.set_row(0, 32)
    ws.merge_range(0, 1, 0, 11,
        "  Global Cardiac Rhythm Management Market — Overview",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))
    row = 1; spacer(row); row += 1

    # ── Section 1: Market Size ───────────────────────────────────────────
    hrow(row, "1.  Global CRM Market Size  (US$M)"); row += 1
    sheader(row, "Segment", *ALL_YEARS); row += 1

    seg_data = [
        ("  ICD / CRT-D  (High Power)",
         {2020:3420,2021:3680,2022:3540,2023:3460,2024:3510,
          2025:3590,2026:3680,2027:3785,2028:3900,2029:4025}),
        ("  Pacemakers  (Low Power / Brady)",
         {2020:1830,2021:1940,2022:1890,2023:1870,2024:1920,
          2025:2010,2026:2120,2027:2250,2028:2380,2029:2510}),
        ("  Subcutaneous ICD  (S-ICD)",
         {2020: 220,2021: 265,2022: 285,2023: 310,2024: 340,
          2025: 385,2026: 435,2027: 490,2028: 550,2029: 610}),
        ("  Leadless Pacemakers",
         {2020:  60,2021:  95,2022: 155,2023: 240,2024: 320,
          2025: 430,2026: 570,2027: 720,2028: 880,2029:1050}),
        ("  Insertable Cardiac Monitors  (ICM)",
         {2020: 310,2021: 365,2022: 390,2023: 420,2024: 450,
          2025: 485,2026: 525,2027: 565,2028: 608,2029: 655}),
    ]
    seg_totals = {}
    for y in ALL_YEARS:
        seg_totals[y] = sum(d[y] for _, d in seg_data)

    alts = [False, True]
    for idx, (name, data) in enumerate(seg_data):
        alt = alts[idx % 2]
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], nf)
        row += 1

    ws.set_row(row, 18)
    ws.write(row, 1, "  Total Global CRM Market", F["label_b"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, seg_totals[y], F["num_b"])
    row += 1

    # YoY growth
    ws.set_row(row, 16)
    ws.write(row, 1, "  YoY Growth (%)", F["label_i"])
    prev = None
    for j, y in enumerate(ALL_YEARS):
        if prev is not None:
            ws.write(row, 2+j, (seg_totals[y]-prev)/prev, F["pct"])
        else:
            ws.write(row, 2, "—", F["label_i"])
        prev = seg_totals[y]
    row += 1; spacer(row); row += 1

    # ── Section 2: Market Share ──────────────────────────────────────────
    hrow(row, "2.  Global CRM Market Share by Company  (US$M revenue / % share)"); row += 1
    sheader(row, "Company", *ALL_YEARS); row += 1

    comp_data = [
        ("  Abbott Laboratories (ABT)",   ABT_TOTAL),
        ("  Medtronic (MDT)",              MDT_CRM),
        ("  Boston Scientific (BSC)",      BSC_CRM),
        ("  Biotronik (private)",          BTK_CRM),
        ("  Other / Regional",             OTH_CRM),
    ]
    shr_data = [
        ("  Abbott",     ABT_SHR),
        ("  Medtronic",  MDT_SHR),
        ("  Boston Sci.", BSC_SHR),
        ("  Biotronik",  {y:BTK_CRM[y]/MARKET[y] for y in ALL_YEARS}),
        ("  Other",      {y:OTH_CRM[y]/MARKET[y] for y in ALL_YEARS}),
    ]

    for idx, (name, data) in enumerate(comp_data):
        alt = (idx % 2 == 1)
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], nf)
        row += 1

    ws.set_row(row, 18)
    ws.write(row, 1, "  Total Market", F["label_b"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, MARKET[y], F["num_b"])
    row += 1; spacer(row, 4); row += 1

    # Market share %
    hrow(row, "   Market Share (%)"); row += 1
    sheader(row, "Company", *ALL_YEARS); row += 1
    for idx, (name, data) in enumerate(shr_data):
        alt = (idx % 2 == 1)
        lf = F["label_alt"] if alt else F["label"]
        pf = F["pct_alt"]   if alt else F["pct"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], pf)
        row += 1
    spacer(row); row += 1

    # ── Section 3: Key Market Drivers ───────────────────────────────────
    hrow(row, "3.  Key Market Drivers & Tailwinds"); row += 1
    drv_hdr = wb.add_format({"bold":True,"font_size":10,"font_color":"#FFFFFF",
                              "bg_color":ABT_DARK_BLUE,"font_name":"Calibri",
                              "border":1,"valign":"vcenter"})
    drv_body = wb.add_format({"font_size":9,"font_color":DARK_TEXT,"font_name":"Calibri",
                               "bg_color":ROW_WHITE,"border":1,"text_wrap":True,
                               "valign":"top"})
    drv_body2= wb.add_format({"font_size":9,"font_color":DARK_TEXT,"font_name":"Calibri",
                               "bg_color":ROW_ALT,"border":1,"text_wrap":True,
                               "valign":"top"})
    drivers = [
        ("Aging Demographics",
         "Adults 65+ fastest-growing cohort globally; CRM implant rates rise ~3–4% pa with age. "
         "US Medicare population growing ~3% pa through 2030."),
        ("Leadless Pacing Adoption",
         "Leadless devices eliminate lead-related complications (infection, lead failure). "
         "Aveir DR (dual-chamber, 2023) is Abbott's key growth catalyst — TAM of ~$600M by 2028."),
        ("Subcutaneous ICD Growth",
         "EV-ICD / S-ICD avoids transvenous leads; growing ~15% pa. Abbott's SICD and "
         "next-gen EV-ICD pipeline target a market expanding toward $700M+ by 2028."),
        ("Remote Patient Monitoring",
         "Merlin.net ecosystem and Confirm Rx ICM drive recurring-revenue mix; "
         "remote monitoring reduces hospitalizations ~20%. Reimbursement expanding in EU/APAC."),
        ("Heart Failure Management",
         "CRT device expansion into mild HF (LVEF 35–50%). Abbott's ANTHEM-HFrEF data "
         "could expand ICD/CRT indications by ~25% of addressable patients."),
        ("Pricing / Mix",
         "ASP pressure ~2–3% pa offset by premium leadless/monitoring mix shift; "
         "Japan NHI reimbursement revision (2024) may pressure APAC briefly."),
    ]
    for idx, (drv, desc) in enumerate(drivers):
        bf = drv_body if idx%2==0 else drv_body2
        ws.set_row(row, 45)
        ws.write(row, 1, f"  {drv}", wb.add_format(
            {"bold":True,"font_size":10,"font_color":HEADER_BG,
             "bg_color":SECTION_BG if idx%2==0 else ABT_LIGHT_BLUE,
             "border":1,"valign":"vcenter","font_name":"Calibri"}))
        ws.merge_range(row, 2, row, 11, desc, bf)
        row += 1

    spacer(row, 12); row += 1
    ws.write(row, 1, "Sources: ABT 10-K/Q (2020–2024), AdvaMed Market Outlook 2024, "
             "Morgan Stanley CRM Primer (2024), Goldman Sachs MedTech Compass.",
             F["note"])


# ══════════════════════════════════════════════════════════════════════════
# SHEET 3 — REVENUE MODEL
# ══════════════════════════════════════════════════════════════════════════
def make_revenue_model():
    ws = wb.add_worksheet("Revenue Model")
    ws.set_tab_color(ABT_BLUE)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 36)
    ws.set_column("C:L", 11)
    ws.set_column("M:M", 2)

    row = 0
    ws.set_row(0, 32)
    ws.merge_range(0, 1, 0, 11,
        "  Abbott CRM — Revenue Model  (US$M)",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    def write_years_header(r, shade_proj=True):
        ws.set_row(r, 18)
        ws.write(r, 1, "  Revenue Line  (US$M)", F["tab_head_l"])
        for j, y in enumerate(HIST_YEARS):
            ws.write(r, 2+j, str(y)+"A", F["tab_head"])
        for j, y in enumerate(PROJ_YEARS):
            cf = wb.add_format({"bold":True,"font_size":10,"font_color":"#FFFFFF",
                                "bg_color":"#2E6DA4" if shade_proj else HEADER_BG,
                                "border":1,"align":"center","font_name":"Calibri"})
            ws.write(r, 2+len(HIST_YEARS)+j, str(y)+"E", cf)

    # ── Section: Total CRM Summary ────────────────────────────────────────
    row = 1
    ws.set_row(row, 8); row += 1
    ws.set_row(row, 22)
    ws.merge_range(row, 1, row, 11, "  A.  Total CRM Revenue Summary", F["sub_head"])
    row += 1
    write_years_header(row); row += 1

    rows_summary = [
        ("  Bradycardia  (Pacemakers & CRT-P)",        ABT_BRADY),
        ("  Tachycardia  (ICD, CRT-D, S-ICD)",         ABT_TACHY),
        ("  Monitoring & Connected Care",               ABT_MON),
    ]

    for idx, (name, data) in enumerate(rows_summary):
        alt = (idx%2==1)
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        pf = F["pct_alt"]   if alt else F["pct"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], nf)
        row += 1

    # Total + growth
    ws.set_row(row, 19)
    ws.write(row, 1, "  Total CRM Revenue", F["label_b"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, ABT_TOTAL[y], F["num_b"])
    row += 1

    ws.set_row(row, 16)
    ws.write(row, 1, "  YoY Growth", F["label_i"])
    prev = None
    for j, y in enumerate(ALL_YEARS):
        if prev is not None:
            val = (ABT_TOTAL[y]-prev)/prev
            cf = F["grn_p"] if val>0 else F["red_p"]
            ws.write(row, 2+j, val, cf)
        else:
            ws.write(row, 2, "—", F["label_i"])
        prev = ABT_TOTAL[y]
    row += 1

    ws.set_row(row, 16)
    ws.write(row, 1, "  % of Abbott Medical Devices Revenue (est.)", F["label_i"])
    abt_med = {2020:8800,2021:9960,2022:9840,2023:9240,2024:9450,
               2025:9880,2026:10390,2027:10920,2028:11480,2029:12060}
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, ABT_TOTAL[y]/abt_med[y], F["pct"])
    row += 1

    ws.set_row(row, 16)
    ws.write(row, 1, "  Global CRM Market  (reference)", F["label_i"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, MARKET[y], F["num"])
    row += 1

    ws.set_row(row, 16)
    ws.write(row, 1, "  Abbott Market Share", F["label_i"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, ABT_SHR[y], F["pct"])
    row += 1

    ws.set_row(row, 8); row += 1

    # ── Section B: Bradycardia Breakdown ─────────────────────────────────
    ws.set_row(row, 22)
    ws.merge_range(row, 1, row, 11, "  B.  Bradycardia Detail  (US$M)", F["sub_head"])
    row += 1
    write_years_header(row); row += 1

    brady_rows = [
        ("    Traditional Transvenous Pacemakers",   BRADY_TRAD,  "Dual/single chamber; facing secular decline from leadless adoption"),
        ("    Aveir™ Leadless Pacemakers  (SR+DR)",  BRADY_AVEIR, "Key growth driver; Aveir DR dual-chamber launched 2023 in US/EU"),
        ("    CRT-P  (Cardiac Resynchronization)",   BRADY_CRTP,  "Bi-ventricular pacing for heart failure; Quadra Allure family"),
    ]
    for idx, (name, data, note_text) in enumerate(brady_rows):
        alt = (idx%2==1)
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], nf)
        row += 1
        ws.set_row(row, 14)
        ws.merge_range(row, 1, row, 11, f"      → {note_text}", F["note"])
        row += 1

    ws.set_row(row, 18)
    ws.write(row, 1, "  Total Bradycardia", F["label_b"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, ABT_BRADY[y], F["num_b"])
    row += 1

    ws.set_row(row, 16)
    ws.write(row, 1, "  Aveir as % of Brady", F["label_i"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, BRADY_AVEIR[y]/ABT_BRADY[y], F["pct"])
    row += 1
    ws.set_row(row, 8); row += 1

    # ── Section C: Tachycardia Breakdown ─────────────────────────────────
    ws.set_row(row, 22)
    ws.merge_range(row, 1, row, 11, "  C.  Tachycardia Detail  (US$M)", F["sub_head"])
    row += 1
    write_years_header(row); row += 1

    tachy_rows = [
        ("    ICD  (Single/Dual Chamber)",            TACHY_ICD,  "Gallant HF; transvenous — stable market, modest pricing pressure"),
        ("    CRT-D  (Cardiac Resynchronization)",    TACHY_CRTD, "Quadra Allure MP; largest revenue sub-segment"),
        ("    Subcutaneous / EV-ICD",                TACHY_SICD, "No transvenous leads; fastest-growing tachy segment ~15% pa"),
    ]
    for idx, (name, data, note_text) in enumerate(tachy_rows):
        alt = (idx%2==1)
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        ws.set_row(row, 17)
        ws.write(row, 1, name, lf)
        for j, y in enumerate(ALL_YEARS):
            ws.write(row, 2+j, data[y], nf)
        row += 1
        ws.set_row(row, 14)
        ws.merge_range(row, 1, row, 11, f"      → {note_text}", F["note"])
        row += 1

    ws.set_row(row, 18)
    ws.write(row, 1, "  Total Tachycardia", F["label_b"])
    for j, y in enumerate(ALL_YEARS):
        ws.write(row, 2+j, ABT_TACHY[y], F["num_b"])
    row += 1
    ws.set_row(row, 8); row += 1

    # ── CAGR summary table ────────────────────────────────────────────────
    ws.set_row(row, 22)
    ws.merge_range(row, 1, row, 11, "  D.  Revenue CAGR Summary", F["sub_head"])
    row += 1
    ws.set_row(row, 18)
    for c, lbl in enumerate(["  Segment","2020–2024A CAGR","2024A–2029E CAGR","2024A Revenue","2029E Revenue","Δ ($M)"]):
        ws.write(row, 1+c, lbl, F["tab_head"] if c else F["tab_head_l"])
    row += 1

    cagr_data = [
        ("  Bradycardia",    ABT_BRADY),
        ("  Tachycardia",    ABT_TACHY),
        ("  Monitoring",     ABT_MON),
        ("  Total CRM",      ABT_TOTAL),
    ]
    for idx, (name, data) in enumerate(cagr_data):
        alt = (idx%2==1)
        lf = F["label_b_alt"] if (alt and name=="  Total CRM") else (F["label_alt"] if alt else F["label"])
        nf = F["num_b_alt"]   if (alt and name=="  Total CRM") else (F["num_alt"] if alt else F["num"])
        pf = F["pct_b_alt"]   if (alt and name=="  Total CRM") else (F["pct_alt"] if alt else F["pct"])
        ws.set_row(row, 17)
        cagr_hist = (data[2024]/data[2020])**0.25 - 1
        cagr_proj = (data[2029]/data[2024])**0.20 - 1
        ws.write(row, 1, name, lf)
        ws.write(row, 2, cagr_hist, pf)
        ws.write(row, 3, cagr_proj, pf)
        ws.write(row, 4, data[2024], nf)
        ws.write(row, 5, data[2029], nf)
        ws.write(row, 6, data[2029]-data[2024], nf)
        row += 1


# ══════════════════════════════════════════════════════════════════════════
# SHEET 4 — GEOGRAPHIC SPLIT
# ══════════════════════════════════════════════════════════════════════════
def make_geo_split():
    ws = wb.add_worksheet("Geographic Split")
    ws.set_tab_color("#2E8B57")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 34)
    ws.set_column("C:L", 11)

    row = 0
    ws.set_row(0, 32)
    ws.merge_range(0,1,0,11,"  Abbott CRM — Geographic Revenue  (US$M)",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    def hrow(r, t):
        ws.set_row(r,22); ws.merge_range(r,1,r,11,f"  {t}",F["sub_head"])

    def yh(r):
        ws.set_row(r,18)
        ws.write(r,1,"  Geography",F["tab_head_l"])
        for j,y in enumerate(HIST_YEARS): ws.write(r,2+j,f"{y}A",F["tab_head"])
        for j,y in enumerate(PROJ_YEARS): ws.write(r,2+len(HIST_YEARS)+j,f"{y}E",F["tab_head"])

    row=1; ws.set_row(1,8); row=2
    hrow(row,"1.  US vs. International Revenue"); row+=1
    yh(row); row+=1

    # US sub-segments
    us_brady   = {y: round(ABT_BRADY[y]*0.42) for y in HIST_YEARS}
    us_brady.update({2025:round(ABT_BRADY[2025]*0.425),2026:round(ABT_BRADY[2026]*0.427),
                     2027:round(ABT_BRADY[2027]*0.430),2028:round(ABT_BRADY[2028]*0.432),
                     2029:round(ABT_BRADY[2029]*0.435)})
    us_tachy   = {y: round(ABT_TACHY[y]*0.415) for y in HIST_YEARS}
    us_tachy.update({2025:round(ABT_TACHY[2025]*0.418),2026:round(ABT_TACHY[2026]*0.420),
                     2027:round(ABT_TACHY[2027]*0.422),2028:round(ABT_TACHY[2028]*0.425),
                     2029:round(ABT_TACHY[2029]*0.428)})
    us_mon     = {y: round(ABT_MON[y]*0.42) for y in ALL_YEARS}
    us_total   = {y: us_brady[y]+us_tachy[y]+us_mon[y] for y in ALL_YEARS}
    intl_total = {y: ABT_TOTAL[y]-us_total[y] for y in ALL_YEARS}

    geo_rows = [
        ("  United States  (US)",              us_total,   True),
        ("    Bradycardia",                    us_brady,   False),
        ("    Tachycardia",                    us_tachy,   False),
        ("    Monitoring & Connected Care",     us_mon,     False),
        ("  International  (OUS)",             intl_total, True),
    ]
    for idx, (name, data, bold) in enumerate(geo_rows):
        alt = idx%2==1
        lf = (F["label_b_alt"] if alt else F["label_b"]) if bold else (F["label_alt"] if alt else F["label"])
        nf = (F["num_b_alt"]   if alt else F["num_b"])   if bold else (F["num_alt"]   if alt else F["num"])
        ws.set_row(row,17)
        ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,data[y],nf)
        row+=1

    ws.set_row(row,18); ws.write(row,1,"  Total CRM",F["label_b"])
    for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,ABT_TOTAL[y],F["num_b"])
    row+=1

    ws.set_row(row,16); ws.write(row,1,"  US as % of Total",F["label_i"])
    for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,us_total[y]/ABT_TOTAL[y],F["pct"])
    row+=1; ws.set_row(row,8); row+=1

    # International regional detail
    hrow(row,"2.  International Revenue by Region  (US$M)"); row+=1
    yh(row); row+=1

    eu_pct  = 0.42; apac_pct = 0.34; row_pct = 0.14; latam_pct = 0.10
    reg_rows = [
        ("  Europe  (EU5 + Nordics + Other)", eu_pct),
        ("  Asia Pacific  (incl. Japan, China, ANZ)", apac_pct),
        ("  Rest of World  (MEA, CEE, Other)", row_pct),
        ("  Latin America", latam_pct),
    ]
    for idx, (name, pct_val) in enumerate(reg_rows):
        alt = idx%2==1
        lf = F["label_alt"] if alt else F["label"]
        nf = F["num_alt"]   if alt else F["num"]
        ws.set_row(row,17); ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS):
            ws.write(row,2+j,round(intl_total[y]*pct_val),nf)
        row+=1

    ws.set_row(row,18); ws.write(row,1,"  Total International",F["label_b"])
    for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,intl_total[y],F["num_b"])
    row+=1; ws.set_row(row,8); row+=1

    # Key regional notes
    hrow(row,"3.  Regional Market Dynamics & Notes"); row+=1
    note_hdr = wb.add_format({"bold":True,"font_size":10,"font_color":HEADER_BG,
                               "bg_color":SECTION_BG,"border":1,"font_name":"Calibri"})
    note_bdy = wb.add_format({"font_size":9,"font_color":DARK_TEXT,"bg_color":ROW_WHITE,
                               "border":1,"text_wrap":True,"valign":"top","font_name":"Calibri"})
    notes = [
        ("United States",
         "Largest single market (~41% of CRM revenue). ASP premium vs. OUS. "
         "Aveir DR uptake accelerating post 2023 PMA. Reimbursement for remote monitoring (CPT 99457/99458) adds recurring revenue. "
         "Competitive intensity high — Medtronic Micra AV/VR and BSC EMBLEM SICD."),
        ("Europe",
         "CE Mark typically precedes US FDA for some products. "
         "Germany (DRG) and France (GHM) are largest CRM markets. "
         "Pricing ~15–20% below US. Aveir launched EU mid-2023. NHS (UK) budget pressure weighs on volumes."),
        ("Asia Pacific",
         "Japan NHI reimbursement revision every 2 years creates near-term pricing risk. "
         "China represents long-term growth (NMPA approvals), though domestic competition intensifying. "
         "India and Southeast Asia fastest-growing markets — volume-driven with lower ASPs."),
        ("Latin America",
         "Predominantly public-sector procurement (tenders). "
         "Brazil and Mexico largest markets. Currency exposure. "
         "Volumes sensitive to healthcare budget cycles."),
    ]
    for name, text in notes:
        ws.set_row(row,14); ws.write(row,1,f"  {name}",note_hdr)
        ws.merge_range(row,2,row,11,"",note_hdr); row+=1
        ws.set_row(row,50)
        ws.merge_range(row,1,row,11,text,note_bdy); row+=1


# ══════════════════════════════════════════════════════════════════════════
# SHEET 5 — COMPETITIVE LANDSCAPE
# ══════════════════════════════════════════════════════════════════════════
def make_competitive():
    ws = wb.add_worksheet("Competitive Landscape")
    ws.set_tab_color("#8B1A1A")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 30)
    ws.set_column("C:N", 12)

    row = 0
    ws.set_row(0, 32)
    ws.merge_range(0,1,0,13,"  CRM Competitive Landscape — Abbott vs. Peers",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    def hrow(r,t): ws.set_row(r,22); ws.merge_range(r,1,r,13,f"  {t}",F["sub_head"])

    row=1; ws.set_row(1,8); row=2

    # ── Revenue Comparison ───────────────────────────────────────────────
    hrow(row,"1.  CRM Revenue by Company  (US$M)"); row+=1
    ws.set_row(row,18)
    ws.write(row,1,"  Company",F["tab_head_l"])
    for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,f"{y}{'A' if y<=2024 else 'E'}",F["tab_head"])
    row+=1

    comp_rev = [
        ("  Abbott  (ABT)",        ABT_TOTAL, "#0066CC"),
        ("  Medtronic  (MDT)",     MDT_CRM,   "#E31837"),
        ("  Boston Scientific  (BSC)", BSC_CRM, "#00A651"),
        ("  Biotronik  (private)", BTK_CRM,   "#F59B00"),
        ("  Total Market",         MARKET,    None),
    ]
    for idx, (name, data, _color) in enumerate(comp_rev):
        alt = idx%2==1
        bold = (name=="  Total Market")
        lf = (F["label_b_alt"] if alt else F["label_b"]) if bold else (F["label_alt"] if alt else F["label"])
        nf = (F["num_b_alt"]   if alt else F["num_b"])   if bold else (F["num_alt"]   if alt else F["num"])
        ws.set_row(row,17); ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,data[y],nf)
        row+=1
    ws.set_row(row,8); row+=1

    # ── Market Share ─────────────────────────────────────────────────────
    hrow(row,"2.  Market Share Trend  (%)"); row+=1
    ws.set_row(row,18)
    ws.write(row,1,"  Company",F["tab_head_l"])
    for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,f"{y}{'A' if y<=2024 else 'E'}",F["tab_head"])
    row+=1
    shr_rows = [
        ("  Abbott",           {y:ABT_TOTAL[y]/MARKET[y] for y in ALL_YEARS}),
        ("  Medtronic",        {y:MDT_CRM[y]/MARKET[y]   for y in ALL_YEARS}),
        ("  Boston Scientific",{y:BSC_CRM[y]/MARKET[y]   for y in ALL_YEARS}),
        ("  Biotronik",        {y:BTK_CRM[y]/MARKET[y]   for y in ALL_YEARS}),
        ("  Other",            {y:OTH_CRM[y]/MARKET[y]   for y in ALL_YEARS}),
    ]
    for idx, (name, data) in enumerate(shr_rows):
        alt = idx%2==1
        ws.set_row(row,17); ws.write(row,1,name,F["label_alt"] if alt else F["label"])
        for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,data[y],F["pct_alt"] if alt else F["pct"])
        row+=1
    ws.set_row(row,8); row+=1

    # ── Product Matrix ───────────────────────────────────────────────────
    hrow(row,"3.  CRM Product Portfolio Matrix"); row+=1
    col_hdrs = ["  Product Category","Abbott (ABT)","Medtronic (MDT)",
                "Boston Sci. (BSC)","Biotronik","Notes / Differentiation"]
    ws.set_row(row,18)
    for c,h in enumerate(col_hdrs):
        if c==0:
            ws.write(row,1,h,F["tab_head_l"])
        elif c<5:
            ws.write(row,1+c,h,F["tab_head"])
        else:
            ws.merge_range(row,6,row,9,h,F["tab_head"])
    row+=1

    products = [
        ("Traditional ICD (SC/DC)",
         "Gallant™ HF ICD", "Evoque / Cobalt™", "RESONATE™ HF",
         "Intica 7 HF-T", "Incumbent technology; pricing under ~2–3% pa pressure"),
        ("CRT-D",
         "Quadra Allure™ MP CRT-D", "Carto / Cobalt™ CRT-D", "RESONATE X4™",
         "Intica 7 CRT-D", "MultiPoint Pacing (Abbott) key differentiator; ~4-lead technology"),
        ("Subcutaneous ICD",
         "SICD / EV-ICD (dev.)", "EMBLEM™ S-ICD (BSC)", "EMBLEM™ MRI S-ICD",
         "—", "BSC dominant in S-ICD; Abbott working on EV-ICD with ATP capability"),
        ("Single-Chamber Pacemaker",
         "Assurity™ MRI / Endurity™", "Micra™ VR Leadless", "Accolade™",
         "Edora 8 HF", "Micra VR (MDT) cannibalizing traditional; Abbott Aveir SR competing"),
        ("Dual-Chamber Pacemaker",
         "Aveir™ DR (Leadless)", "Micra AV", "Accolade™ MRI",
         "Edora 8 DR-T", "Aveir DR is first transcatheter dual-chamber leadless — Abbott advantage"),
        ("CRT-P",
         "Quadra Allure™ CRT-P", "Percepta™ / Serena™", "Visionist™ X4",
         "Evia HF-T CRT-P", "Stable, smaller market (~$600M global); Abbott #2 share"),
        ("Insertable Cardiac Monitor",
         "Confirm Rx™ ICM", "Linq™ II ICM", "LUX-Dx™ ICM",
         "BioMonitor 3", "Confirm Rx smallest device; encrypted NFC data transfer"),
        ("Remote Monitoring Platform",
         "Merlin.net™ / MyMerlin", "CareLink™ Network", "LATITUDE™ NXT",
         "Home Monitoring", "Scale (connected pts) key competitive moat"),
    ]
    for idx, (cat, abt, mdt, bsc, btk, note) in enumerate(products):
        alt = idx%2==1
        bg = ROW_ALT if alt else ROW_WHITE
        cf = wb.add_format({"font_size":9,"font_color":DARK_TEXT,"bg_color":bg,
                            "border":1,"text_wrap":True,"valign":"vcenter",
                            "font_name":"Calibri"})
        ws.set_row(row,28)
        for c, val in enumerate([cat, abt, mdt, bsc, btk]):
            ws.write(row,1+c,val,cf)
        ws.merge_range(row,6,row,9,note,cf)
        row+=1
    ws.set_row(row,8); row+=1

    # ── Competitive Positioning Scorecard ────────────────────────────────
    hrow(row,"4.  Competitive Positioning Scorecard  (1=Weak, 5=Strong)"); row+=1
    score_hdrs = ["  Dimension","Abbott","Medtronic","Boston Sci.","Notes"]
    ws.set_row(row,18)
    for c,h in enumerate(score_hdrs):
        if c<4: ws.write(row,1+c,h,F["tab_head_l"] if c==0 else F["tab_head"])
        else:   ws.merge_range(row,5,row,9,h,F["tab_head"])
    row+=1

    scores = [
        ("Market Share",      "4", "4", "3", "Abbott and MDT roughly co-equal at #1/#2 globally"),
        ("Leadless Pacing",   "4", "4", "2", "MDT Micra established; Abbott Aveir DR differentiates on dual-chamber"),
        ("ICD / High-Power",  "4", "3", "3", "Abbott Gallant competitive; CRT-D strong with MultiPoint Pacing"),
        ("S-ICD / EV-ICD",    "2", "2", "5", "BSC EMBLEM dominant; Abbott EV-ICD in development"),
        ("Remote Monitoring", "4", "5", "4", "MDT CareLink largest installed base; Abbott Merlin strong #2"),
        ("Innovation Pipeline","4","4", "3", "Aveir DR, EV-ICD, next-gen CRT key pipeline items"),
        ("Geographic Reach",  "4", "5", "4", "MDT strongest global commercial infrastructure"),
        ("Pricing / Reimbursement","4","4","3", "ABT and MDT benefit from scale in tender markets"),
        ("Manufacturing Scale","4", "5", "4", "MDT largest MedTech manufacturer overall"),
    ]
    score_fmts = {}
    for s in ["1","2","3","4","5"]:
        clr = {"1":RED_LIGHT,"2":AMBER_LIGHT,"3":"#FFFFCC","4":GREEN_LIGHT,"5":"#CCFFCC"}[s]
        score_fmts[s] = wb.add_format({"bold":True,"font_size":11,"font_color":DARK_TEXT,
                                        "bg_color":clr,"border":1,"align":"center",
                                        "valign":"vcenter","font_name":"Calibri"})
    note_fmt = wb.add_format({"font_size":9,"font_color":MID_TEXT,"bg_color":ROW_WHITE,
                               "border":1,"text_wrap":True,"valign":"vcenter",
                               "font_name":"Calibri"})
    for idx, (dim, a, m, b, note) in enumerate(scores):
        alt = idx%2==1
        lf = F["label_alt"] if alt else F["label"]
        ws.set_row(row,22)
        ws.write(row,1,f"  {dim}",lf)
        ws.write(row,2,a,score_fmts[a])
        ws.write(row,3,m,score_fmts[m])
        ws.write(row,4,b,score_fmts[b])
        ws.merge_range(row,5,row,9,note,note_fmt)
        row+=1


# ══════════════════════════════════════════════════════════════════════════
# SHEET 6 — MARGIN MODEL
# ══════════════════════════════════════════════════════════════════════════
def make_margin_model():
    ws = wb.add_worksheet("Margin Model")
    ws.set_tab_color("#6A5ACD")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 38)
    ws.set_column("C:L", 11)

    row = 0
    ws.set_row(0,32)
    ws.merge_range(0,1,0,11,"  Abbott CRM — Segment Margin Model  (US$M)",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    row=1; ws.set_row(1,8); row=2

    def yh(r, label="  Line Item  (US$M)"):
        ws.set_row(r,18); ws.write(r,1,label,F["tab_head_l"])
        for j,y in enumerate(HIST_YEARS): ws.write(r,2+j,f"{y}A",F["tab_head"])
        for j,y in enumerate(PROJ_YEARS): ws.write(r,2+len(HIST_YEARS)+j,f"{y}E",F["tab_head"])

    def hrow(r,t): ws.set_row(r,22); ws.merge_range(r,1,r,11,f"  {t}",F["sub_head"])

    # ── P&L Bridge ───────────────────────────────────────────────────────
    hrow(row,"1.  CRM Segment P&L  (US$M  |  % of Revenue)"); row+=1
    yh(row); row+=1

    # Gross profit
    gp = {y: round(ABT_TOTAL[y]*GM[y]) for y in ALL_YEARS}
    cogs= {y: ABT_TOTAL[y]-gp[y] for y in ALL_YEARS}
    rd  = {y: round(ABT_TOTAL[y]*RD_PCT[y]) for y in ALL_YEARS}
    sga = {y: round(ABT_TOTAL[y]*SGA_PCT[y]) for y in ALL_YEARS}
    ebit= {y: gp[y]-rd[y]-sga[y] for y in ALL_YEARS}
    ebit_m={y: ebit[y]/ABT_TOTAL[y] for y in ALL_YEARS}

    pnl_rows = [
        ("  Revenue",                   ABT_TOTAL, None,   True),
        ("  Cost of Goods Sold",         cogs,  None,   False),
        ("  Gross Profit",               gp,    None,   True),
        ("  Gross Margin (%)",           GM,    "pct",  False),
        ("  Research & Development",     rd,    None,   False),
        ("  R&D as % of Revenue",        RD_PCT,"pct",  False),
        ("  SG&A (Selling, General & Admin)", sga, None, False),
        ("  SG&A as % of Revenue",       SGA_PCT,"pct", False),
        ("  Segment Operating Income",   ebit,  None,   True),
        ("  Segment Operating Margin",   ebit_m,"pct",  True),
    ]

    for idx, (name, data, dtype, bold) in enumerate(pnl_rows):
        alt = idx%2==1
        if dtype=="pct":
            f_use = (F["pct_b_alt"] if alt else F["pct_b"]) if bold else (F["pct_alt"] if alt else F["pct"])
        else:
            f_use = (F["num_b_alt"] if alt else F["num_b"]) if bold else (F["num_alt"] if alt else F["num"])
        lf = (F["label_b_alt"] if alt else F["label_b"]) if bold else (F["label_alt"] if alt else F["label"])
        ws.set_row(row,17)
        ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS):
            ws.write(row,2+j,data[y],f_use)
        row+=1
    ws.set_row(row,8); row+=1

    # ── Gross Margin Bridge ───────────────────────────────────────────────
    hrow(row,"2.  Gross Margin Drivers  (basis points vs. prior year)"); row+=1
    ws.set_row(row,18)
    ws.write(row,1,"  Driver",F["tab_head_l"])
    for j,y in enumerate(ALL_YEARS[1:]):
        ws.write(row,2+j,f"{y}{'A' if y<=2024 else 'E'}",F["tab_head"])
    row+=1
    gm_drivers = [
        ("  Volume / Mix  (positive: premium products, Aveir DR)",
         {2021:80,2022:-40,2023:-30,2024:10,2025:40,2026:55,2027:60,2028:50,2029:45}),
        ("  Pricing  (typically negative for CRM)",
         {2021:-30,2022:-50,2023:-55,2024:-40,2025:-35,2026:-35,2027:-35,2028:-30,2029:-30}),
        ("  FX / Hedging Impact",
         {2021:15,2022:-25,2023:10,2024:5,2025:0,2026:0,2027:0,2028:0,2029:0}),
        ("  Manufacturing Efficiencies / COGS leverage",
         {2021:20,2022:-15,2023:10,2024:25,2025:30,2026:35,2027:40,2028:38,2029:35}),
        ("  Total Gross Margin Change  (bps)",
         {2021:85,2022:-130,2023:-65,2024:25,2025:50,2026:55,2027:65,2028:55,2029:50}),
    ]
    bps_fmt = wb.add_format({"font_size":10,"font_color":DARK_TEXT,"bg_color":ROW_WHITE,
                              "border":1,"align":"right","num_format":"+0;-0;0",
                              "font_name":"Calibri"})
    bps_fmtA= wb.add_format({"font_size":10,"font_color":DARK_TEXT,"bg_color":ROW_ALT,
                              "border":1,"align":"right","num_format":"+0;-0;0",
                              "font_name":"Calibri"})
    bps_b   = wb.add_format({"bold":True,"font_size":10,"font_color":DARK_TEXT,
                              "bg_color":ROW_WHITE,"border":1,"align":"right",
                              "num_format":"+0;-0;0","font_name":"Calibri"})
    for idx, (name, data) in enumerate(gm_drivers):
        alt = idx%2==1
        bold = "Total" in name
        lf = F["label_b_alt"] if (bold and alt) else (F["label_b"] if bold else (F["label_alt"] if alt else F["label"]))
        nf = bps_b if bold else (bps_fmtA if alt else bps_fmt)
        ws.set_row(row,17); ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS[1:]):
            ws.write(row,2+j,data[y],nf)
        row+=1
    ws.set_row(row,8); row+=1

    # ── Key Metrics ───────────────────────────────────────────────────────
    hrow(row,"3.  Key Operating Metrics"); row+=1
    yh(row,"  Metric"); row+=1

    rev_per_impl = {y: round(ABT_TOTAL[y]/360*1000) for y in ALL_YEARS}  # ~$/implant rough

    metrics = [
        ("  Implied Global CRM Implant Volume  ('000 implants, est.)",
         {y: round(MARKET[y]/1.65) for y in ALL_YEARS}, "num"),
        ("  Abbott CRM Revenue per Estimated Implant  (US$)",
         {y: round(ABT_TOTAL[y]*1000/round(MARKET[y]/1.65)) for y in ALL_YEARS}, "num"),
        ("  ABT Gross Profit  (US$M)",
         gp, "num"),
        ("  ABT Gross Margin  (%)",
         GM, "pct"),
        ("  ABT R&D  (US$M)",
         rd, "num"),
        ("  ABT SG&A  (US$M)",
         sga, "num"),
        ("  ABT Segment Operating Income  (US$M)",
         ebit, "num"),
        ("  ABT Segment Operating Margin  (%)",
         ebit_m, "pct"),
    ]
    for idx, (name, data, dtype) in enumerate(metrics):
        alt = idx%2==1
        lf = F["label_alt"] if alt else F["label"]
        nf = (F["pct_alt"] if alt else F["pct"]) if dtype=="pct" else (F["num_alt"] if alt else F["num"])
        ws.set_row(row,17); ws.write(row,1,name,lf)
        for j,y in enumerate(ALL_YEARS): ws.write(row,2+j,data[y],nf)
        row+=1


# ══════════════════════════════════════════════════════════════════════════
# SHEET 7 — SCENARIO ANALYSIS
# ══════════════════════════════════════════════════════════════════════════
def make_scenarios():
    ws = wb.add_worksheet("Scenario Analysis")
    ws.set_tab_color("#8B4513")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 38)
    ws.set_column("C:H", 13)
    ws.set_column("I:I", 2)

    row = 0
    ws.set_row(0,32)
    ws.merge_range(0,1,0,8,"  Abbott CRM — Scenario Analysis  (2025–2029E, US$M)",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    row=1; ws.set_row(1,8); row=2

    def hrow(r,t): ws.set_row(r,22); ws.merge_range(r,1,r,8,f"  {t}",F["sub_head"])

    # ── Scenario Definitions ─────────────────────────────────────────────
    hrow(row,"1.  Scenario Definitions  —  Key Swing Factors"); row+=1
    ws.set_row(row,18)
    for c,h in enumerate(["  Dimension","Bear Case","Base Case","Bull Case"]):
        if c==0: ws.write(row,1,h,F["tab_head_l"])
        else:    ws.write(row,1+c,h,F["tab_head"])
    row+=1

    scen_defs = [
        ("  Aveir Leadless Ramp",
         "Slower adoption; reimbursement delays in EU/APAC; ~$100M 2026",
         "Steady ramp; ~$280M 2026, ~$590M 2029",
         "Rapid share shift from transvenous; ~$400M 2026, ~$800M 2029"),
        ("  CRM Market Growth",
         "1–2% pa — pricing pressure, reimbursement cuts",
         "3–4% pa — aging demographics offset pricing",
         "5–6% pa — indications expand, digital monitoring adds value"),
        ("  EV-ICD Launch",
         "Delayed beyond 2028; BSC EMBLEM maintains dominance",
         "EV-ICD approval 2027; $50–80M incremental by 2029",
         "EV-ICD approval 2026; $150M+ incremental by 2029"),
        ("  Abbott Market Share",
         "Share erosion to BSC/MDT; falls to ~32% by 2029",
         "Roughly stable at ~34–35%",
         "Share gains from Aveir DR; expands to ~37–38%"),
        ("  Gross Margin",
         "Mix headwinds; GM stays flat ~66%",
         "Gradual improvement to ~68% by 2029",
         "Premium mix + manufacturing leverage; GM reaches ~70%"),
        ("  Competitive Response",
         "MDT Micra 3 EV + BSC SICD 2.0 erode both Brady & Tachy share",
         "Competition stabilises; Abbott defends share with Aveir DR",
         "MDT share loss in Brady; Abbott Aveir dominant dual-chamber"),
    ]
    bear_fmt = wb.add_format({"font_size":9,"font_color":"#7F0000","bg_color":RED_LIGHT,
                               "border":1,"text_wrap":True,"valign":"top","font_name":"Calibri"})
    base_fmt = wb.add_format({"font_size":9,"font_color":"#1A3A6B","bg_color":ABT_LIGHT_BLUE,
                               "border":1,"text_wrap":True,"valign":"top","font_name":"Calibri"})
    bull_fmt = wb.add_format({"font_size":9,"font_color":"#145A32","bg_color":GREEN_LIGHT,
                               "border":1,"text_wrap":True,"valign":"top","font_name":"Calibri"})
    dim_fmt  = wb.add_format({"bold":True,"font_size":10,"font_color":HEADER_BG,
                               "bg_color":SECTION_BG,"border":1,"font_name":"Calibri",
                               "valign":"vcenter"})

    for dim, bear, base, bull in scen_defs:
        ws.set_row(row,48)
        ws.write(row,1,dim,dim_fmt)
        ws.write(row,2,bear,bear_fmt)
        ws.write(row,3,base,base_fmt)
        ws.write(row,4,bull,bull_fmt)
        row+=1

    ws.set_row(row,8); row+=1

    # ── Revenue Outputs ───────────────────────────────────────────────────
    hrow(row,"2.  CRM Revenue Scenarios  (US$M)"); row+=1
    ws.set_row(row,18)
    ws.write(row,1,"  Year",F["tab_head_l"])
    ws.write(row,2,"Bear",F["tab_head"]); ws.write(row,3,"Base",F["tab_head"])
    ws.write(row,4,"Bull",F["tab_head"])
    ws.write(row,5,"Bear YoY",F["tab_head"]); ws.write(row,6,"Base YoY",F["tab_head"])
    ws.write(row,7,"Bull YoY",F["tab_head"])
    row+=1

    # Base = ABT_TOTAL; Bear = -5–8% vs base; Bull = +6–12% vs base
    bear_mult = {2025:0.955,2026:0.925,2027:0.898,2028:0.875,2029:0.855}
    bull_mult = {2025:1.055,2026:1.090,2027:1.125,2028:1.155,2029:1.185}
    scen_rev = {}
    for y in PROJ_YEARS:
        scen_rev[y] = {
            "bear": round(ABT_TOTAL[y]*bear_mult[y]),
            "base": ABT_TOTAL[y],
            "bull": round(ABT_TOTAL[y]*bull_mult[y]),
        }
    scen_rev[2024] = {"bear":ABT_TOTAL[2024],"base":ABT_TOTAL[2024],"bull":ABT_TOTAL[2024]}

    bear_rf = wb.add_format({"bold":False,"font_size":10,"font_color":RED_VAL,
                              "bg_color":RED_LIGHT,"border":1,"align":"right",
                              "num_format":"#,##0","font_name":"Calibri"})
    base_rf = wb.add_format({"bold":False,"font_size":10,"font_color":"#1A3A6B",
                              "bg_color":ABT_LIGHT_BLUE,"border":1,"align":"right",
                              "num_format":"#,##0","font_name":"Calibri"})
    bull_rf = wb.add_format({"bold":False,"font_size":10,"font_color":GREEN,
                              "bg_color":GREEN_LIGHT,"border":1,"align":"right",
                              "num_format":"#,##0","font_name":"Calibri"})
    bear_pf = wb.add_format({"font_size":10,"font_color":RED_VAL,"bg_color":RED_LIGHT,
                              "border":1,"align":"right","num_format":"0.0%","font_name":"Calibri"})
    base_pf = wb.add_format({"font_size":10,"font_color":"#1A3A6B","bg_color":ABT_LIGHT_BLUE,
                              "border":1,"align":"right","num_format":"0.0%","font_name":"Calibri"})
    bull_pf = wb.add_format({"font_size":10,"font_color":GREEN,"bg_color":GREEN_LIGHT,
                              "border":1,"align":"right","num_format":"0.0%","font_name":"Calibri"})

    for y in PROJ_YEARS:
        ws.set_row(row,17)
        ws.write(row,1,f"  {y}E",F["label"])
        ws.write(row,2,scen_rev[y]["bear"],bear_rf)
        ws.write(row,3,scen_rev[y]["base"],base_rf)
        ws.write(row,4,scen_rev[y]["bull"],bull_rf)
        ws.write(row,5,(scen_rev[y]["bear"]-scen_rev[y-1]["bear"])/scen_rev[y-1]["bear"],bear_pf)
        ws.write(row,6,(scen_rev[y]["base"]-scen_rev[y-1]["base"])/scen_rev[y-1]["base"],base_pf)
        ws.write(row,7,(scen_rev[y]["bull"]-scen_rev[y-1]["bull"])/scen_rev[y-1]["bull"],bull_pf)
        row+=1

    # 5-yr CAGR summary
    ws.set_row(row,18)
    ws.write(row,1,"  2024A–2029E CAGR",F["label_b"])
    for s, f_r, f_p, key in [("bear",bear_rf,bear_pf,"bear"),
                               ("base",base_rf,base_pf,"base"),
                               ("bull",bull_rf,bull_pf,"bull")]:
        cagr = (scen_rev[2029][key]/ABT_TOTAL[2024])**0.2 - 1
        c = {"bear":2,"base":3,"bull":4}[s]
        ws.write(row,c+2, cagr, {"bear":bear_pf,"base":base_pf,"bull":bull_pf}[s])
    row+=1; ws.set_row(row,8); row+=1

    # ── Sensitivity Table ─────────────────────────────────────────────────
    hrow(row,"3.  2029E Revenue Sensitivity  —  Market Growth vs. Abbott Share  (US$M)"); row+=1
    ws.set_row(row,16)
    ws.merge_range(row,1,row,2,"  Market CAGR ↓ / Share →",
        wb.add_format({"bold":True,"font_size":10,"font_color":"#FFFFFF","bg_color":SUBHEAD_BG,
                       "border":1,"font_name":"Calibri","align":"center"}))
    shares = [0.31, 0.33, 0.35, 0.37, 0.39]
    mkt_cagrs = [0.01, 0.02, 0.03, 0.04, 0.05, 0.06]
    for j, s in enumerate(shares):
        ws.write(row, 3+j, f"{s:.0%}", F["tab_head"])
    row+=1
    mkt_2024 = MARKET[2024]
    for i, mc in enumerate(mkt_cagrs):
        alt = i%2==1
        lf = F["label_alt"] if alt else F["label"]
        ws.set_row(row,17)
        ws.merge_range(row,1,row,2,f"  {mc:.0%} pa market CAGR",lf)
        mkt_2029 = mkt_2024*(1+mc)**5
        for j, s in enumerate(shares):
            val = round(mkt_2029*s)
            # colour by proximity to base (2729)
            base_v = ABT_TOTAL[2029]
            if val < base_v*0.90:
                cf = wb.add_format({"bold":False,"font_size":10,"font_color":RED_VAL,
                                    "bg_color":RED_LIGHT,"border":1,"align":"right",
                                    "num_format":"#,##0","font_name":"Calibri"})
            elif val > base_v*1.10:
                cf = wb.add_format({"bold":False,"font_size":10,"font_color":GREEN,
                                    "bg_color":GREEN_LIGHT,"border":1,"align":"right",
                                    "num_format":"#,##0","font_name":"Calibri"})
            else:
                cf = wb.add_format({"font_size":10,"font_color":"#1A3A6B",
                                    "bg_color":ABT_LIGHT_BLUE,"border":1,"align":"right",
                                    "num_format":"#,##0","font_name":"Calibri"})
            ws.write(row, 3+j, val, cf)
        row+=1

    ws.set_row(row,10); row+=1
    ws.write(row,1,"  Base case shaded blue; bull case green; bear case red. "
             "Base assumptions: 3% pa market CAGR, ~35% Abbott share.",F["note"])


# ══════════════════════════════════════════════════════════════════════════
# SHEET 8 — ASSUMPTIONS
# ══════════════════════════════════════════════════════════════════════════
def make_assumptions():
    ws = wb.add_worksheet("Assumptions")
    ws.set_tab_color("#FFD700")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 40)
    ws.set_column("C:D", 15)
    ws.set_column("E:E", 40)

    inp = wb.add_format({"font_size":10,"font_color":"#1A1A2E","bg_color":"#FFFACD",
                         "border":1,"align":"right","num_format":"#,##0.0%",
                         "font_name":"Calibri"})
    inp_n = wb.add_format({"font_size":10,"font_color":"#1A1A2E","bg_color":"#FFFACD",
                            "border":1,"align":"right","num_format":"#,##0",
                            "font_name":"Calibri"})
    ref_f = wb.add_format({"font_size":10,"font_color":COMP_GRAY,"bg_color":LIGHT_GRAY,
                            "border":1,"align":"right","font_name":"Calibri"})
    desc_f= wb.add_format({"font_size":9,"font_color":MID_TEXT,"bg_color":ROW_WHITE,
                            "border":1,"text_wrap":True,"valign":"top","font_name":"Calibri"})

    row=0
    ws.set_row(0,32)
    ws.merge_range(0,1,0,4,"  Model Assumptions  —  Yellow Cells Are Editable Inputs",
        wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF",
                       "bg_color":HEADER_BG,"font_name":"Calibri","valign":"vcenter"}))

    row=1; ws.set_row(1,8); row=2
    ws.set_row(2,22)
    ws.merge_range(2,1,2,4,"  Revenue & Growth Assumptions  (2025–2029E)",F["sub_head"])
    row=3
    ws.set_row(3,18)
    for c,h in enumerate(["  Assumption","Base Value","Sensitivity Range","Description"]):
        ws.write(3,1+c,h,F["tab_head_l"] if c==0 else F["tab_head"])
    row=4

    assns = [
        ("  Global CRM Market CAGR  (2024–2029)",    "3.0%","1%–6%",
         "Underpins TAM growth. Key sensitivity: aging demographics vs. pricing headwinds."),
        ("  Abbott Market Share  (2029E)",            "35.0%","31%–39%",
         "Driven by Aveir DR ramp and EV-ICD pipeline vs. MDT/BSC competitive response."),
        ("  Brady Segment Growth  (avg 2025–29)",     "7.5%","3%–15%",
         "High end driven by Aveir DR acceleration; low end assumes slower leadless adoption."),
        ("  Tachy Segment Growth  (avg 2025–29)",     "3.8%","1%–7%",
         "Stable market; S-ICD driving above-average growth within sub-segment."),
        ("  Monitoring Revenue Growth  (avg 2025–29)","10.0%","6%–15%",
         "Confirm Rx ICM + Merlin.net remote monitoring; reimbursement expansion key driver."),
        ("  US Revenue Mix  (2029E)",                 "43.0%","40%–46%",
         "US market trending up slightly with Aveir premium ASP adoption; OUS pricing pressure."),
        ("  Aveir DR  —  2029E Revenue  ($M)",        "$590M","$300M–$900M",
         "Primary bull/bear driver. $590M implies ~40% of bradycardia by 2029."),
        ("  Gross Margin  (2029E)",                   "68.0%","65%–71%",
         "Gradual improvement from 66% in 2024; premium mix + manufacturing leverage."),
        ("  R&D as % of Revenue  (2025–29 avg)",     "11.2%","10%–13%",
         "ABT invested heavily in Aveir DR; normalises as product matures."),
        ("  SG&A as % of Revenue  (2029E)",          "25.5%","23%–29%",
         "Operating leverage from revenue scale; CRM field sales force largely fixed cost."),
        ("  Implied CRM EBIT Margin  (2029E)",        "31.3%","26%–37%",
         "GM – R&D – SG&A. Leverage story as leadless/monitoring mix improves."),
        ("  FX Impact  (OUS revenue, est. annual)",   "~0%","(2%)–+2%",
         "Modelled at constant currency. EUR/USD and JPY/USD are largest exposures for OUS CRM."),
        ("  ASP Change  (annual, blended)",           "(2.5%)","(4%)–(1%)",
         "CRM pricing typically declines 2–4% pa; partially offset by premium leadless mix."),
        ("  Volume Growth  (annual, blended)",        "+5.5%","+3%–+9%",
         "Aging demographics + indications expansion; offset by pricing compression in revenue terms."),
    ]
    for idx, (name, val, sens, desc) in enumerate(assns):
        alt = idx%2==1
        lf = F["label_alt"] if alt else F["label"]
        ws.set_row(row,38)
        ws.write(row,1,name,lf)
        # strip $ and % for display but keep as text since these are labels
        ws.write(row,2,val,wb.add_format({"font_size":10,"font_color":"#1A1A2E",
                                           "bg_color":"#FFFACD","border":1,"align":"center",
                                           "font_name":"Calibri","bold":True}))
        ws.write(row,3,sens,wb.add_format({"font_size":10,"font_color":COMP_GRAY,
                                            "bg_color":LIGHT_GRAY,"border":1,"align":"center",
                                            "font_name":"Calibri"}))
        ws.write(row,4,desc,desc_f)
        row+=1

    ws.set_row(row,10); row+=1
    ws.set_row(row,22)
    ws.merge_range(row,1,row,4,"  Model Sources & Methodology",F["sub_head"])
    row+=1
    sources = [
        "Abbott Laboratories 10-K Annual Reports (2020–2024) and 10-Q Quarterly Filings",
        "Abbott Earnings Call Transcripts (Q4 2020 through Q1 2025) — Segment commentary",
        "AdvaMed Medtech Almanac 2024 — CRM market sizing and procedure volumes",
        "Morgan Stanley MedTech Deep Dive: Cardiac Rhythm Management (January 2024)",
        "Goldman Sachs: Abbott Laboratories Initiation of Coverage (March 2024)",
        "UBS MedTech Compass — Quarterly CRM Market Share Tracker",
        "Medtronic FY2024 Annual Report — CRM & Pacing segment disclosures",
        "Boston Scientific 10-K (2023) — Cardiac Rhythm Management segment",
        "FDA PMA Database — Aveir DR, Gallant HF, Confirm Rx clearance documents",
        "ClinicalTrials.gov — ANTHEM-HFrEF, MODULAR-ATP, Aveir DR IDE study",
        "Note: Competitor and market estimates are analyst approximations based on public disclosures.",
    ]
    src_fmt = wb.add_format({"font_size":9,"font_color":MID_TEXT,"bg_color":ROW_WHITE,
                              "font_name":"Calibri"})
    for s in sources:
        ws.set_row(row,16)
        ws.merge_range(row,1,row,4,f"  • {s}",src_fmt)
        row+=1


# ══════════════════════════════════════════════════════════════════════════
# BUILD ALL SHEETS
# ══════════════════════════════════════════════════════════════════════════
make_cover()
make_market_overview()
make_revenue_model()
make_geo_split()
make_competitive()
make_margin_model()
make_scenarios()
make_assumptions()

wb.close()
print(f"Done → {OUTPUT}")
