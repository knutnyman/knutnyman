#!/usr/bin/env python3
"""Build an Excel workbook summarizing Abbott Laboratories (ABT) full-year
guidance as issued at each quarterly earnings report, 2020 -> Q1 2026.

Data compiled from Abbott press releases (PR Newswire / abbott.mediaroom.com)
and SEC 8-K exhibits, cross-checked against secondary financial sources.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from collections import OrderedDict

# ---------------------------------------------------------------------------
# Source data: one record per quarterly report (chronological).
# Fields:
#  0 date  1 quarter  2 FY  3 action
#  4 adjEPS text  5 adj_low  6 adj_high  7 adj_mid
#  8 GAAP EPS text  9 sales-growth headline  10 sales-growth basis
#  11 COVID-testing-sales assumption  12 commentary  13 source
# For "at least $X" EPS guidance, adj_low = X and adj_high = None.
# ---------------------------------------------------------------------------
COLS = [
    "Report Date", "Quarter Reported", "FY Guided", "Action",
    "Adj. EPS Guidance", "Adj. EPS Low", "Adj. EPS High", "Adj. EPS Midpoint",
    "GAAP EPS Guidance", "Sales Growth Guidance", "Sales Growth Basis",
    "COVID-19 Testing Sales Assumption", "Key Commentary", "Source",
]

ROWS = [
    # ---------------- FY2020 ----------------
    ["2020-01-22", "Q4 2019", "FY2020", "Initial",
     "$3.55 – $3.65", 3.55, 3.65, 3.60,
     "$2.35 – $2.45", "7.0% – 8.0%", "Total-company organic",
     "n/a (pre-COVID)",
     "Initial FY2020 outlook issued pre-COVID; double-digit EPS growth at midpoint.",
     "https://www.sec.gov/Archives/edgar/data/0000001800/000110465920005717/tm205456d1_ex99-1.htm"],

    ["2020-04-16", "Q1 2020", "FY2020", "Withdrawn",
     "Withdrawn / suspended", None, None, None,
     "Withdrawn / suspended", "Withdrawn", "Withdrawn",
     "Uncertain",
     "Suspended all FY2020 guidance due to COVID-19 uncertainty; no new numbers provided.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2020-results-301041836.html"],

    ["2020-07-16", "Q2 2020", "FY2020", "Reinstated (floor)",
     "At least $3.25", 3.25, None, None,
     "At least $2.00", "Not provided (EPS floor only)", "Not quantified",
     "Ramping",
     "Guidance reinstated as an 'at least' EPS floor given COVID uncertainty; no sales-growth % given; ~$1.25 of specified items.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2020-results-exceeds-analysts-expectations-301094737.html"],

    ["2020-10-21", "Q3 2020", "FY2020", "Raised",
     "At least $3.55", 3.55, None, None,
     "At least $2.35", "Not provided (EPS floor only)", "Not quantified",
     "Strong (BinaxNOW / ID NOW / Panbio)",
     "Raised EPS floor from 'at least $3.25' on strong COVID diagnostics demand; no sales-growth %. FY2020 actual landed at $3.65 adj.",
     "https://abbott.mediaroom.com/2020-10-21-Abbott-Reports-Third-Quarter-2020-Results-Achieves-Strong-Double-Digit-Earnings-Growth-and-Raises-Guidance"],

    # ---------------- FY2021 ----------------
    ["2021-01-27", "Q4 2020", "FY2021", "Initial",
     "At least $5.00", 5.00, None, None,
     "At least $3.74", "~Double-digit (qualitative; no % issued)", "Not quantified",
     "~$2.4B in Q4'20",
     "Initial FY2021 outlook: EPS growth of more than 35%. Headline was 'strong double-digit growth' but no numeric full-year sales range was issued (~3.5% favorable FX noted).",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-2020-results-issues-strong-double-digit-growth-forecast-for-2021-301216123.html"],

    ["2021-04-20", "Q1 2021", "FY2021", "Maintained",
     "At least $5.00", 5.00, None, None,
     "At least $3.74", "Not provided (qualitative)", "Not quantified",
     "Strong",
     "Guidance maintained; no numeric full-year sales-growth range issued.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2021-results-301272437.html"],

    ["2021-07-22", "Q2 2021", "FY2021", "Lowered",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$2.75 – $2.95", "Low double digits", "Base business, ex-COVID testing",
     "Sharp decline",
     "EPS lowered on falling COVID-test demand; full-year base-business organic framed as 'low double digits' (first numerically labeled at the Jun 1, 2021 off-cycle update).",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2021-results-301339347.html"],

    ["2021-10-20", "Q3 2021", "FY2021", "Raised",
     "$5.00 – $5.10", 5.00, 5.10, 5.05,
     "$3.55 – $3.65", "Low double digits", "Base business, ex-COVID testing",
     "Renewed (Delta wave)",
     "EPS raised sharply (~38.4% growth at midpoint) on renewed COVID-testing demand plus base-business strength; base organic 'low double digits' maintained.",
     "https://abbott.mediaroom.com/2021-10-20-Abbott-Reports-Third-Quarter-2021-Results-Achieves-Strong-Double-Digit-Earnings-Growth-and-Raises-Guidance"],

    # ---------------- FY2022 ----------------
    ["2022-01-26", "Q4 2021", "FY2022", "Initial",
     "At least $4.70", 4.70, None, None,
     "At least $3.43", "High single digits", "Base business, ex-COVID testing",
     "~$2.5B (initial)",
     "Initial FY2022 outlook; base-business organic 'high single digits'. Initial COVID-testing sales forecast ~$2.5B, to be updated quarterly.",
     "https://www.prnewswire.com/news-releases/abbott-reports-strong-fourth-quarter-2021-results-issues-2022-forecast-301468554.html"],

    ["2022-04-20", "Q1 2022", "FY2022", "Maintained",
     "At least $4.70", 4.70, None, None,
     "At least $3.35", "High single digits", "Base business, ex-COVID testing",
     "Raised to ~$4.5B",
     "Adjusted EPS maintained; base-business organic 'high single digits'. COVID-testing sales forecast raised to ~$4.5B (Q1 COVID test sales $3.3B).",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2022-results-301528820.html"],

    ["2022-07-20", "Q2 2022", "FY2022", "Raised",
     "At least $4.90", 4.90, None, None,
     "At least $3.50", "Low double digits", "Base business, ex-COVID testing",
     "Raised to $6.1B",
     "EPS raised (>= $4.70 -> >= $4.90); base-business organic raised to 'low double digits'; COVID-testing sales raised to $6.1B.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2022-results-and-raises-full-year-eps-guidance-301590017.html"],

    ["2022-10-19", "Q3 2022", "FY2022", "Raised",
     "$5.17 – $5.23", 5.17, 5.23, 5.20,
     "$3.75 – $3.81", "Low double digits", "Base business, ex-COVID testing",
     "Raised to ~$7.8B",
     "EPS raised again to a defined range; base-business organic 'low double digits' maintained. FY2022 actual landed at $5.34 adj.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2022-results-and-raises-full-year-eps-guidance-301653429.html"],

    # ---------------- FY2023 ----------------
    ["2023-01-25", "Q4 2022", "FY2023", "Initial",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.05 – $3.25", "High single digits", "Base business, ex-COVID testing",
     "~$2.0B",
     "Initial FY2023 outlook; base-business organic 'high single digits' ex-COVID. Assumed ~$2.0B of COVID-testing sales.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2022-results-issues-2023-financial-outlook-301730346.html"],

    ["2023-04-19", "Q1 2023", "FY2023", "Maintained",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.05 – $3.25", "≈High single digits (raised outlook)", "Base business, ex-COVID testing",
     "Lowered to ~$1.5B",
     "Adjusted EPS maintained; underlying base-business outlook raised (~high single digits), offset by lower COVID. Q1 base organic +10.0%.",
     "https://abbott.mediaroom.com/2023-04-19-Abbott-Reports-First-Quarter-2023-Results-Increases-Outlook-For-Underlying-Base-Business"],

    ["2023-07-20", "Q2 2023", "FY2023", "Maintained",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.02 – $3.22", "Low double digits", "Base business, ex-COVID testing",
     "Lowered to ~$1.3B",
     "Adjusted EPS maintained; base-business outlook raised to 'low double digits', offset by lower COVID. Q2 base organic +11.5%.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2023-results-increases-outlook-for-underlying-base-business-301882027.html"],

    ["2023-10-18", "Q3 2023", "FY2023", "Narrowed (midpoint raised)",
     "$4.42 – $4.46", 4.42, 4.46, 4.44,
     "$3.14 – $3.18", "Low double digits", "Base business, ex-COVID testing",
     "~$1.6B (implied)",
     "EPS range narrowed and midpoint raised; base-business organic 'low double digits' maintained.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2023-results-and-raises-midpoint-of-full-year-eps-guidance-range-301960446.html"],

    # ---------------- FY2024 ----------------
    ["2024-01-24", "Q4 2023", "FY2024", "Initial",
     "$4.50 – $4.70", 4.50, 4.70, 4.60,
     "$3.20 – $3.40", "8.0% – 10.0%", "Organic (COVID immaterial)",
     "Minimal",
     "Initial FY2024 outlook; headline reverts to a single organic range now that COVID is immaterial.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2023-results-issues-2024-financial-outlook-302043275.html"],

    ["2024-04-17", "Q1 2024", "FY2024", "Raised midpoint",
     "$4.55 – $4.70", 4.55, 4.70, 4.625,
     "$3.25 – $3.40", "8.5% – 10.0%", "Organic (COVID immaterial)",
     "Minimal",
     "Raised midpoint of all ranges; narrowed organic range by raising the low end.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2024-results-and-raises-midpoint-of-full-year-guidance-ranges-302119393.html"],

    ["2024-07-18", "Q2 2024", "FY2024", "Raised",
     "$4.61 – $4.71", 4.61, 4.71, 4.66,
     "$3.30 – $3.40", "9.5% – 10.0%", "Organic (COVID immaterial)",
     "Minimal",
     "Raised full-year guidance again; organic range narrowed upward.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2024-results-and-raises-full-year-guidance-302200584.html"],

    ["2024-10-16", "Q3 2024", "FY2024", "Raised midpoint / narrowed",
     "$4.64 – $4.70", 4.64, 4.70, 4.67,
     "$3.34 – $3.40", "9.5% – 10.0%", "Organic (COVID immaterial)",
     "Minimal",
     "Raised EPS midpoint / narrowed; organic range maintained. FY actual landed at the upper end.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2024-results-and-raises-midpoint-of-full-year-eps-guidance-range-302277755.html"],

    # ---------------- FY2025 ----------------
    ["2025-01-22", "Q4 2024", "FY2025", "Initial",
     "$5.05 – $5.25", 5.05, 5.25, 5.15,
     "Not provided (discontinued)", "7.5% – 8.5%", "Total-company organic",
     "Negligible",
     "Initial FY2025 outlook; double-digit EPS growth at midpoint; adj. operating margin 23.5%-24.0%. GAAP EPS guidance discontinued.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2024-results-issues-2025-financial-outlook-302357321.html"],

    ["2025-04-16", "Q1 2025", "FY2025", "Reaffirmed",
     "$5.05 – $5.25", 5.05, 5.25, 5.15,
     "Not provided (discontinued)", "7.5% – 8.5%", "Total-company organic",
     "Negligible",
     "Reaffirmed all FY2025 guidance.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2025-results-and-reaffirms-full-year-guidance-302430212.html"],

    ["2025-07-17", "Q2 2025", "FY2025", "Narrowed",
     "$5.10 – $5.20", 5.10, 5.20, 5.15,
     "Not provided (discontinued)", "7.5% – 8.0% (6.0% – 7.0% incl. COVID)", "Organic, ex-COVID",
     "Negligible",
     "Narrowed adj. EPS range (midpoint unchanged); narrowed organic range.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2025-results-302507875.html"],

    ["2025-10-15", "Q3 2025", "FY2025", "Reaffirmed / narrowed",
     "$5.12 – $5.18", 5.12, 5.18, 5.15,
     "Not provided (discontinued)", "7.5% – 8.0% (6.0% – 7.0% incl. COVID)", "Organic, ex-COVID",
     "Negligible",
     "Reaffirmed midpoint, narrowed adj. EPS range; reaffirmed organic.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2025-results-and-reaffirms-full-year-guidance-302584746.html"],

    # ---------------- FY2026 ----------------
    ["2026-01-22", "Q4 2025", "FY2026", "Initial",
     "$5.55 – $5.80", 5.55, 5.80, 5.675,
     "Not provided (discontinued)", "6.5% – 7.5%", "Total-company organic",
     "n/a",
     "Initial FY2026 outlook; ~10% EPS growth at midpoint. Announced acquisition of Exact Sciences (not yet closed at guide).",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2025-results-issues-2026-financial-outlook-302668032.html"],

    ["2026-04-16", "Q1 2026", "FY2026", "Lowered (acquisition dilution)",
     "$5.38 – $5.58", 5.38, 5.58, 5.48,
     "Not provided (discontinued)", "6.5% – 7.5%", "Comparable (adj. for M&A & FX)",
     "n/a",
     "EPS lowered to reflect ~$0.20 dilution from Exact Sciences (closed Mar 23, 2026); top-line range maintained, now stated on a 'comparable' basis. Not an operational guide-down.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2026-results-updates-guidance-to-reflect-acquisition-of-exact-sciences-302744652.html"],
]

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------
NAVY = "1F3864"
ALT = "EEF3FB"
WHITE = "FFFFFF"

ACTION_FILL = {
    "Initial": "BDD7EE",
    "Raised": "C6EFCE",
    "Raised midpoint": "C6EFCE",
    "Raised midpoint / narrowed": "C6EFCE",
    "Maintained": "FFF2CC",
    "Reaffirmed": "FFF2CC",
    "Reaffirmed / narrowed": "FFF2CC",
    "Narrowed (midpoint raised)": "FFF2CC",
    "Narrowed": "FFF2CC",
    "Reinstated (floor)": "FCE4D6",
    "Lowered": "FFC7CE",
    "Lowered (acquisition dilution)": "FFC7CE",
    "Withdrawn": "F2F2F2",
}

# fill for the Sales Growth Basis column, so the revenue dimension reads at a glance
BASIS_FILL = {
    "Total-company organic": "C6EFCE",
    "Organic (COVID immaterial)": "C6EFCE",
    "Organic, ex-COVID": "D9EAD3",
    "Base business, ex-COVID testing": "FFF2CC",
    "Comparable (adj. for M&A & FX)": "BDD7EE",
    "Not quantified": "F2F2F2",
    "Withdrawn": "F2F2F2",
}

thin = Side(style="thin", color="B4C6E7")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEADER_FONT = Font(name="Calibri", bold=True, color=WHITE, size=11)
TITLE_FONT = Font(name="Calibri", bold=True, color=NAVY, size=16)
SUB_FONT = Font(name="Calibri", italic=True, color="404040", size=10)
CELL_FONT = Font(name="Calibri", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_header_row(ws, row_idx, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.font = HEADER_FONT
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = CENTER
        cell.border = BORDER


# ---------------------------------------------------------------------------
# Sheet 1: Guidance Timeline (chronological)
# ---------------------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Guidance Timeline"

ncol = len(COLS)
last_col = get_column_letter(ncol)
ws.merge_cells(f"A1:{last_col}1")
ws["A1"] = "Abbott Laboratories (NYSE: ABT) — Full-Year Guidance by Quarter"
ws["A1"].font = TITLE_FONT
ws.merge_cells(f"A2:{last_col}2")
ws["A2"] = ("How Abbott's full-year outlook (EPS and sales growth) was set and revised at each quarterly earnings report, "
            "Q4 2019 report (initial FY2020) through Q1 2026.  Compiled from Abbott press releases & SEC 8-K exhibits.")
ws["A2"].font = SUB_FONT
ws.row_dimensions[1].height = 24

header_row = 4
for j, name in enumerate(COLS, start=1):
    ws.cell(row=header_row, column=j, value=name)
style_header_row(ws, header_row, ncol)
ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

r = header_row + 1
for rec in ROWS:
    fy = rec[2]
    for j, val in enumerate(rec, start=1):
        cell = ws.cell(row=r, column=j, value=val)
        cell.font = CELL_FONT
        cell.border = BORDER
        cell.alignment = WRAP
        if j in (2, 3, 6, 7, 8):
            cell.alignment = CENTER
        if j in (6, 7, 8) and isinstance(val, float):
            cell.number_format = '$#,##0.00'
    # zebra by fiscal-year block
    block_fill = ALT if fy in ("FY2020", "FY2022", "FY2024", "FY2026") else WHITE
    for j in range(1, ncol + 1):
        ws.cell(row=r, column=j).fill = PatternFill("solid", fgColor=block_fill)
    # color the Action cell (col 4) and Sales Growth Basis cell (col 11)
    ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=ACTION_FILL.get(rec[3], WHITE))
    ws.cell(row=r, column=4).alignment = CENTER
    ws.cell(row=r, column=11).fill = PatternFill("solid", fgColor=BASIS_FILL.get(rec[10], WHITE))
    ws.cell(row=r, column=11).alignment = CENTER
    r += 1

widths = [12, 13, 9, 20, 16, 9, 9, 11, 18, 22, 22, 20, 44, 30]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
for rr in range(header_row + 1, r):
    ws.row_dimensions[rr].height = 60

# ---------------------------------------------------------------------------
# Sheet 2: By Fiscal Year (how guidance for each FY evolved) — revenue growth featured
# ---------------------------------------------------------------------------
ws2 = wb.create_sheet("By Fiscal Year")
ws2.merge_cells("A1:I1")
ws2["A1"] = "Evolution of Guidance — grouped by the fiscal year being guided"
ws2["A1"].font = TITLE_FONT
ws2.merge_cells("A2:I2")
ws2["A2"] = ("Read across each block to see how the outlook for a given fiscal year changed quarter to quarter. "
             "Sales-growth columns show the headline figure and the basis it was measured on.")
ws2["A2"].font = SUB_FONT

cols2 = ["FY Guided", "Set At (Report)", "Report Date", "Action",
         "Adj. EPS Guidance", "Adj. EPS Midpoint",
         "Sales Growth Guidance", "Sales Growth Basis", "GAAP EPS Guidance"]
hr2 = 4
for j, name in enumerate(cols2, start=1):
    ws2.cell(row=hr2, column=j, value=name)
style_header_row(ws2, hr2, len(cols2))
ws2.freeze_panes = ws2.cell(row=hr2 + 1, column=1)

groups = OrderedDict()
for rec in ROWS:
    groups.setdefault(rec[2], []).append(rec)

r2 = hr2 + 1
fy_colors = {"FY2020": "E2EFDA", "FY2021": "DDEBF7", "FY2022": "FCE4D6",
             "FY2023": "FFF2CC", "FY2024": "E2EFDA", "FY2025": "DDEBF7", "FY2026": "FCE4D6"}
for fy, recs in groups.items():
    start_r = r2
    for rec in recs:
        out = [rec[2], rec[1], rec[0], rec[3], rec[4], rec[7], rec[9], rec[10], rec[8]]
        for j, val in enumerate(out, start=1):
            cell = ws2.cell(row=r2, column=j, value=val)
            cell.font = CELL_FONT
            cell.border = BORDER
            cell.alignment = WRAP
            if j in (1, 2, 3, 6, 7):
                cell.alignment = CENTER
            if j == 6 and isinstance(val, float):
                cell.number_format = '$#,##0.00'
            cell.fill = PatternFill("solid", fgColor=fy_colors.get(fy, WHITE))
        # emphasize action + the sales-growth columns
        ws2.cell(row=r2, column=4).fill = PatternFill("solid", fgColor=ACTION_FILL.get(rec[3], WHITE))
        ws2.cell(row=r2, column=4).alignment = CENTER
        ws2.cell(row=r2, column=8).fill = PatternFill("solid", fgColor=BASIS_FILL.get(rec[10], WHITE))
        ws2.cell(row=r2, column=8).alignment = CENTER
        r2 += 1
    if r2 - start_r > 1:
        ws2.merge_cells(start_row=start_r, start_column=1, end_row=r2 - 1, end_column=1)
        ws2.cell(row=start_r, column=1).alignment = Alignment(horizontal="center", vertical="center")

widths2 = [10, 16, 12, 24, 16, 12, 26, 28, 20]
for i, w in enumerate(widths2, start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w
for rr in range(hr2 + 1, r2):
    ws2.row_dimensions[rr].height = 32

# ---------------------------------------------------------------------------
# Sheet: Growth to Guidance (how much growth is needed to hit guidance/consensus)
# ---------------------------------------------------------------------------
ws5 = wb.create_sheet("Growth to Guidance")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # yellow = user input
CALC_FILL = PatternFill("solid", fgColor="F2F7FF")    # light = computed
SECT_FILL = PatternFill("solid", fgColor=NAVY)
EPS_FMT = '$#,##0.00'
M_FMT = '#,##0'
PCT_FMT = '0.0%'


def w(ws, r, c, val, fmt=None, bold=False, fill=None, align=None,
      border=True, font_color=None, italic=False, size=10):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Calibri", bold=bold, italic=italic,
                     size=size, color=font_color or "000000")
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    else:
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    if border:
        cell.border = BORDER
    return cell


def section(ws, r, text, span=5):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(bold=True, color=WHITE, size=11)
    c.fill = SECT_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")


# ---- Title -------------------------------------------------------------
ws5.merge_cells("A1:N1")
ws5["A1"] = "Growth Needed to Reach Guidance & Consensus"
ws5["A1"].font = TITLE_FONT
ws5.merge_cells("A2:N2")
ws5["A2"] = ("Yellow cells = inputs you populate (consensus). White/blue cells compute automatically in Excel. "
             "FY2026 is the live year (Q1 reported); the lower table backtests every year and accepts consensus EPS.")
ws5["A2"].font = SUB_FONT

# =======================================================================
# BLOCK 1a — FY2026 ADJUSTED EPS calculator
# =======================================================================
section(ws5, 4, "FY2026  ·  Adjusted EPS — growth still needed to reach guidance & consensus", span=5)
hdr = ["Metric", "Guidance Low", "Guidance Mid", "Guidance High", "Consensus (input)"]
for j, t in enumerate(hdr, start=1):
    w(ws5, 5, j, t, bold=True, fill=PatternFill("solid", fgColor="D9E1F2"),
      align="center")

# row 6 full-year target ($)
w(ws5, 6, 1, "Full-year target — adjusted EPS ($)")
w(ws5, 6, 2, 5.38, EPS_FMT, align="center")
w(ws5, 6, 3, 5.48, EPS_FMT, align="center")
w(ws5, 6, 4, 5.58, EPS_FMT, align="center")
w(ws5, 6, 5, None, EPS_FMT, align="center", fill=INPUT_FILL)   # consensus EPS input
# row 7 prior-year actual
w(ws5, 7, 1, "Prior-year actual — FY2025 ($)")
w(ws5, 7, 2, 5.15, EPS_FMT, align="center")
for j in (3, 4, 5):
    w(ws5, 7, j, "=$B$7", EPS_FMT, align="center")
# row 8 implied full-year growth
w(ws5, 8, 1, "Implied full-year EPS growth vs FY2025")
for j, col in zip((2, 3, 4, 5), "BCDE"):
    w(ws5, 8, j, f"={col}6/$B$7-1", PCT_FMT, align="center", fill=CALC_FILL)
# row 9 YTD actual
w(ws5, 9, 1, "YTD actual — Q1'26 ($)")
w(ws5, 9, 2, 1.15, EPS_FMT, align="center")
for j in (3, 4, 5):
    w(ws5, 9, j, "=$B$9", EPS_FMT, align="center")
# row 10 prior-year YTD
w(ws5, 10, 1, "Prior-year YTD — Q1'25 ($)")
w(ws5, 10, 2, 1.09, EPS_FMT, align="center")
for j in (3, 4, 5):
    w(ws5, 10, j, "=$B$10", EPS_FMT, align="center")
# row 11 YTD growth achieved
w(ws5, 11, 1, "YTD growth achieved (Q1'26 vs Q1'25)")
for j in (2, 3, 4, 5):
    w(ws5, 11, j, "=$B$9/$B$10-1", PCT_FMT, align="center", fill=CALC_FILL)
# row 12 remaining EPS still needed
w(ws5, 12, 1, "Remaining (Q2–Q4) EPS still needed ($)")
for j, col in zip((2, 3, 4, 5), "BCDE"):
    w(ws5, 12, j, f"={col}6-$B$9", EPS_FMT, align="center", fill=CALC_FILL)
# row 13 prior-year remaining
w(ws5, 13, 1, "Prior-year remaining — Q2–Q4'25 ($)")
for j in (2, 3, 4, 5):
    w(ws5, 13, j, 4.06 if j == 2 else "=$B$13", EPS_FMT, align="center")
# row 14 required remaining growth
w(ws5, 14, 1, "Required Q2–Q4 EPS growth to hit target", bold=True)
for j, col in zip((2, 3, 4, 5), "BCDE"):
    w(ws5, 14, j, f"={col}12/$B$13-1", PCT_FMT, align="center", bold=True,
      fill=PatternFill("solid", fgColor="E2EFDA"))

# =======================================================================
# BLOCK 1b — FY2026 SALES (comparable growth) calculator
# =======================================================================
section(ws5, 16, "FY2026  ·  Sales (comparable growth) — growth still needed to reach guidance & consensus", span=5)
for j, t in enumerate(hdr, start=1):
    w(ws5, 17, j, t, bold=True, fill=PatternFill("solid", fgColor="D9E1F2"), align="center")
# row 18 full-year comparable growth target
w(ws5, 18, 1, "Full-year comparable sales-growth target")
w(ws5, 18, 2, 0.065, PCT_FMT, align="center")
w(ws5, 18, 3, 0.070, PCT_FMT, align="center")
w(ws5, 18, 4, 0.075, PCT_FMT, align="center")
w(ws5, 18, 5, None, PCT_FMT, align="center", fill=INPUT_FILL)   # consensus growth input
# row 19 YTD comparable achieved
w(ws5, 19, 1, "YTD (Q1'26) comparable growth achieved")
w(ws5, 19, 2, 0.037, PCT_FMT, align="center")
for j in (3, 4, 5):
    w(ws5, 19, j, "=$B$19", PCT_FMT, align="center")
# row 20 required remaining comparable growth (weighted by prior-yr quarterly sales)
w(ws5, 20, 1, "Required Q2–Q4 comparable growth to hit target", bold=True)
for j, col in zip((2, 3, 4, 5), "BCDE"):
    w(ws5, 20, j, f"=({col}18*$H$18-$B$19*$H$19)/$H$20", PCT_FMT, align="center",
      bold=True, fill=PatternFill("solid", fgColor="E2EFDA"))
# helper weights to the right (cols G/H)
w(ws5, 17, 7, "Prior-yr sales weights ($M)", bold=True, fill=PatternFill("solid", fgColor="D9E1F2"))
ws5.merge_cells("G17:H17")
w(ws5, 18, 7, "FY2025 full-year"); w(ws5, 18, 8, 44328, M_FMT, align="center")
w(ws5, 19, 7, "Q1'25"); w(ws5, 19, 8, 10358, M_FMT, align="center")
w(ws5, 20, 7, "Q2–Q4'25"); w(ws5, 20, 8, 33970, M_FMT, align="center")
w(ws5, 21, 1,
  "Note: the required-remaining-growth row weights quarters by prior-year reported sales (approximation; Abbott's "
  "'comparable' base also adjusts for the Exact Sciences acquisition & FX). EPS calc above is exact (uses actual quarterly EPS).",
  italic=True, size=9, border=False)
ws5.merge_cells("A21:H21")

# =======================================================================
# BLOCK 2 — All fiscal years: guidance-implied growth vs actual + consensus
# =======================================================================
section(ws5, 24, "All fiscal years  ·  Guidance-implied growth vs actual results (enter consensus in the yellow column)", span=14)
b2h = ["FY", "Prior-Yr Net Sales ($M)", "Actual Net Sales ($M)", "Actual Reported Sales Growth",
       "Actual Total Organic %", "Actual Base ex-COVID Organic %", "Prior-Yr Adj EPS",
       "Actual Adj EPS", "Actual EPS Growth", "Initial EPS Guidance", "Final EPS Guidance",
       "Actual − Initial Guid ($)", "Consensus Adj EPS (input)", "Actual − Consensus ($)"]
for j, t in enumerate(b2h, start=1):
    w(ws5, 25, j, t, bold=True, fill=PatternFill("solid", fgColor="D9E1F2"),
      align="center", size=9)

# FY, priorSales, actualSales, totOrg, baseOrg, priorEPS, actualEPS, initGuid, finalGuid
B2 = [
    ["FY2020", 31904, 34608, None,  None, 3.24, 3.65, 3.60, 3.55],
    ["FY2021", 34608, 43075, 0.229, 0.137, 3.65, 5.21, 5.00, 5.05],
    ["FY2022", 43075, 43653, 0.064, None, 5.21, 5.34, 4.70, 5.20],
    ["FY2023", 43653, 40109, None,  0.116, 5.34, 4.44, 4.40, 4.44],
    ["FY2024", 40109, 41950, 0.071, 0.096, 4.44, 4.67, 4.60, 4.67],
    ["FY2025", 41950, 44328, 0.055, 0.067, 4.67, 5.15, 5.15, 5.15],
    ["FY2026", 44328, None,  None,  None, 5.15, None, 5.675, 5.48],
]
r = 26
for fy, psales, asales, torg, borg, peps, aeps, ig, fg in B2:
    w(ws5, r, 1, fy, bold=True, align="center")
    w(ws5, r, 2, psales, M_FMT, align="center")
    w(ws5, r, 3, asales, M_FMT, align="center")
    w(ws5, r, 4, (f"=C{r}/B{r}-1" if asales else None), PCT_FMT, align="center", fill=CALC_FILL)
    w(ws5, r, 5, torg, PCT_FMT, align="center")
    w(ws5, r, 6, borg, PCT_FMT, align="center")
    w(ws5, r, 7, peps, EPS_FMT, align="center")
    w(ws5, r, 8, aeps, EPS_FMT, align="center")
    w(ws5, r, 9, (f"=H{r}/G{r}-1" if aeps else None), PCT_FMT, align="center", fill=CALC_FILL)
    w(ws5, r, 10, ig, EPS_FMT, align="center")
    w(ws5, r, 11, fg, EPS_FMT, align="center")
    w(ws5, r, 12, (f"=H{r}-J{r}" if aeps else None), EPS_FMT, align="center", fill=CALC_FILL)
    w(ws5, r, 13, None, EPS_FMT, align="center", fill=INPUT_FILL)        # consensus input
    w(ws5, r, 14, f"=IF(AND(H{r}<>\"\",M{r}<>\"\"),H{r}-M{r},\"\")", EPS_FMT, align="center", fill=CALC_FILL)
    r += 1
# FY2026 in-progress note
w(ws5, r, 1,
  "FY2026 in progress: YTD Q1'26 comparable sales growth +3.7%, Q1'26 adjusted EPS $1.15 (vs $1.09). "
  "Initial EPS guidance $5.55–$5.80 (mid $5.675); current $5.38–$5.58 (mid $5.48) after Exact Sciences dilution.",
  italic=True, size=9, border=False)
ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)

# column widths
gw = [40, 12, 12, 13, 12, 12, 12, 12, 12, 12, 12, 13, 14, 13]
for i, ww in enumerate(gw, start=1):
    ws5.column_dimensions[get_column_letter(i)].width = ww
ws5.freeze_panes = "A3"

# ---------------------------------------------------------------------------
# Sheet 3: Notes & Sources
# ---------------------------------------------------------------------------
ws3 = wb.create_sheet("Notes & Sources")
ws3.column_dimensions["A"].width = 4
ws3.column_dimensions["B"].width = 120
ws3["B1"] = "Notes, Methodology & Sources"
ws3["B1"].font = TITLE_FONT

notes = [
    "",
    "WHAT THIS SHOWS",
    "Each row is the full-year guidance Abbott provided at a given quarterly earnings report. 'Action' flags whether that update",
    "set initial guidance or raised / lowered / maintained / narrowed / withdrew it versus the prior update for the same fiscal year.",
    "",
    "READING THE SALES-GROWTH COLUMNS",
    "Abbott's headline top-line metric changed with the COVID cycle, so 'Sales Growth Basis' tells you what each figure measures:",
    "• Total-company organic — a clean organic % for the whole company (FY2020 initial; FY2025; FY2026 initial).",
    "• Base business, ex-COVID testing — Abbott's preferred metric during 2021-2023; it strips out volatile COVID-test sales.",
    "  In those years Abbott usually gave a DESCRIPTOR ('high single digits' / 'low double digits') rather than a numeric range,",
    "  and put hard numbers only on EPS and on the COVID-testing-sales dollar forecast.",
    "• Organic (COVID immaterial) — from FY2024, with COVID negligible, Abbott returned to a single organic range.",
    "• Comparable (adj. for M&A & FX) — FY2026 Q1 relabeling that neutralizes the Exact Sciences acquisition and FX.",
    "• Not quantified — no sales-growth figure was issued (e.g., the 2020 EPS-floor-only reinstatements; FY2021 setting).",
    "",
    "NOTABLE SALES-GROWTH PATTERN (2022 & 2023)",
    "Both years followed the same arc: base-business organic set at 'high single digits', raised to 'low double digits' by mid-year",
    "as base-business strength offset declining COVID-testing revenue — while adjusted EPS was held roughly flat each time.",
    "",
    "KEY STRUCTURAL POINTS",
    "• 2020-2022 were dominated by COVID-19 testing. Abbott withdrew FY2020 guidance entirely in Apr 2020, reinstated it as an",
    "  'at least' EPS floor mid-year, then raised it. FY2021-2022 swung sharply with COVID-test demand (FY2021 EPS was cut in",
    "  Q2 2021 as testing fell, then raised in Q3 2021 during the Delta wave).",
    "• 'At least $X' EPS entries are floor guidance (common 2020-2022); numeric Low/High columns leave the High blank for these.",
    "• GAAP EPS guidance was discontinued starting with the FY2025 outlook (Jan 2025): Abbott states it cannot reliably forecast",
    "  special items (restructuring, impairments, acquisition charges). 'Not provided (discontinued)' reflects that policy change.",
    "• FY2021 had the least numeric sales guidance: at the Jan 2021 setting Abbott issued NO full-year sales-growth %; the first",
    "  numeric base-business label ('low double digits') came at an off-cycle update on Jun 1, 2021.",
    "• Q1 2026: adjusted EPS was lowered purely to reflect ~$0.20 of dilution from the Exact Sciences acquisition (closed Mar 23,",
    "  2026); the top-line growth range was maintained and restated on a 'comparable' basis. Not an operational guide-down.",
    "",
    "DATA CONFIDENCE",
    "• Adjusted-EPS figures are the highest-confidence numbers and were corroborated across at least two sources per quarter.",
    "• Minor uncertainty: the initial FY2022 GAAP floor (cited as ~$3.43 vs $3.40); the exact COVID-testing-sales assumption at",
    "  Q3 2022 (~$7.8B) and the implied FY2023 figure at Q3 2023; and the precise wording of the Q1 2023 base-business descriptor",
    "  ('high single digits' vs 'at least high single digits').",
    "",
    "GROWTH TO GUIDANCE TAB",
    "A calculator showing how much growth is still needed to reach guidance and consensus. Yellow cells are inputs (consensus);",
    "everything else recomputes in Excel. The top blocks track the live FY2026 year: given Q1'26 actuals, they show the Q2-Q4",
    "growth required to land at the low/mid/high of guidance (and at a consensus you enter). The lower table backtests every",
    "fiscal year (guidance-implied growth vs actual results) and accepts a consensus adj. EPS per year.",
    "Actuals used (Abbott reported): net sales ($M) 2019: 31,904; 2020: 34,608; 2021: 43,075; 2022: 43,653; 2023: 40,109;",
    "2024: 41,950; 2025: 44,328. Adjusted EPS 2019: 3.24; 2020: 3.65; 2021: 5.21; 2022: 5.34; 2023: 4.44; 2024: 4.67; 2025: 5.15.",
    "FY2025 quarterly adj. EPS: 1.09 / 1.26 / 1.30 / 1.50. Q1'26: net sales 11,164; adj. EPS 1.15; comparable growth +3.7%.",
    "Caveats: the FY2026 sales 'required remaining growth' weights quarters by prior-year REPORTED sales (Abbott's comparable",
    "base also nets out the Exact Sciences acquisition and FX, which I don't have at quarterly granularity) - treat as indicative;",
    "the EPS calculator is exact. FY2020 total-company organic % and FY2022/FY2023 some organic splits were not cleanly",
    "confirmed and are left blank rather than estimated.",
    "",
    "PRIMARY SOURCES",
    "Abbott investor news (abbott.mediaroom.com) and PR Newswire press releases, plus the corresponding SEC 8-K exhibit 99.1 for",
    "each quarter. Per-row source links are in the 'Source' column of the Guidance Timeline sheet. Secondary cross-checks included",
    "Nasdaq, StockTitan, RTTNews, MedTech Dive, 360Dx, Motley Fool / Seeking Alpha transcripts and Investing.com.",
    "",
    "Compiled 2026-06-05. Figures are guidance as issued at the time of each report, not actuals. For an auditable record, confirm",
    "each figure against the linked SEC 8-K / press release.",
]
for i, line in enumerate(notes, start=2):
    cell = ws3.cell(row=i, column=2, value=line)
    if line.isupper() and line.strip():
        cell.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    else:
        cell.font = Font(name="Calibri", size=10)
    cell.alignment = Alignment(wrap_text=True, vertical="top")

# legends
lr = len(notes) + 4
ws3.cell(row=lr, column=2, value="ACTION COLOR LEGEND").font = Font(bold=True, color=NAVY, size=11)
legend = [("Initial", "Initial outlook for a fiscal year"),
          ("Raised", "Guidance raised vs prior update"),
          ("Maintained", "Held or reaffirmed around the same level"),
          ("Narrowed (midpoint raised)", "Range tightened, midpoint up"),
          ("Reinstated (floor)", "Guidance restored as an 'at least' floor"),
          ("Lowered", "Guidance reduced vs prior update"),
          ("Withdrawn", "Guidance suspended")]
row_cursor = lr + 1
for label, desc in legend:
    ws3.cell(row=row_cursor, column=1).fill = PatternFill("solid", fgColor=ACTION_FILL.get(label, "F2F2F2"))
    ws3.cell(row=row_cursor, column=1).border = BORDER
    ws3.cell(row=row_cursor, column=2, value=f"{label} — {desc}").font = Font(name="Calibri", size=10)
    row_cursor += 1

row_cursor += 1
ws3.cell(row=row_cursor, column=2, value="SALES GROWTH BASIS COLOR LEGEND").font = Font(bold=True, color=NAVY, size=11)
row_cursor += 1
basis_legend = [("Total-company organic", "Clean organic % for the whole company"),
                ("Organic (COVID immaterial)", "Single organic range, COVID negligible (FY2024)"),
                ("Organic, ex-COVID", "Organic excluding residual COVID testing"),
                ("Base business, ex-COVID testing", "Underlying growth ex-COVID (often a descriptor, 2021-2023)"),
                ("Comparable (adj. for M&A & FX)", "Adjusted for the Exact Sciences acquisition & FX (FY2026)"),
                ("Not quantified", "No sales-growth figure issued")]
for label, desc in basis_legend:
    ws3.cell(row=row_cursor, column=1).fill = PatternFill("solid", fgColor=BASIS_FILL.get(label, "F2F2F2"))
    ws3.cell(row=row_cursor, column=1).border = BORDER
    ws3.cell(row=row_cursor, column=2, value=f"{label} — {desc}").font = Font(name="Calibri", size=10)
    row_cursor += 1

# ---------------------------------------------------------------------------
# Sheet 4: Trends (charts of the guidance trajectory)
# ---------------------------------------------------------------------------
ws4 = wb.create_sheet("Trends")
ws4.merge_cells("A1:F1")
ws4["A1"] = "Guidance Trajectory — adjusted EPS and organic sales-growth midpoints over time"
ws4["A1"].font = TITLE_FONT
ws4.merge_cells("A2:F2")
ws4["A2"] = ("EPS line uses the guidance midpoint, or the floor for 'at least $X' guidance. "
            "Organic line shows the midpoint only for quarters where Abbott gave a numeric range (FY2020 initial, FY2024 onward).")
ws4["A2"].font = SUB_FONT

# helper data block (hidden-ish, to the side / below)
dh = 4  # data header row
ws4.cell(row=dh, column=1, value="Report")
ws4.cell(row=dh, column=2, value="Adj. EPS guide point ($)")
ws4.cell(row=dh, column=3, value="Organic growth midpoint (%)")
for c in range(1, 4):
    ws4.cell(row=dh, column=c).font = Font(bold=True, size=9, color=NAVY)

# numeric organic midpoints aligned to ROWS (None where Abbott gave no numeric range)
ORG_MID = [7.5, None, None, None, None, None, None, None, None, None, None, None,
           None, None, None, None, 9.0, 9.25, 9.75, 9.75, 8.0, 8.0, 7.75, 7.75, 7.0, 7.0]

rr = dh + 1
for rec, omid in zip(ROWS, ORG_MID):
    label = f"{rec[1]} ({rec[2][2:]})"            # e.g. "Q3 2024 (2024)"
    eps_pt = rec[7] if rec[7] is not None else rec[5]   # midpoint, else floor
    ws4.cell(row=rr, column=1, value=label)
    ws4.cell(row=rr, column=2, value=eps_pt)
    ws4.cell(row=rr, column=3, value=omid)
    rr += 1
last_data = rr - 1

# Chart 1 — adjusted EPS guidance trajectory
c1 = LineChart()
c1.title = "Abbott adjusted EPS guidance over time"
c1.style = 12
c1.y_axis.title = "Adjusted EPS ($)"
c1.x_axis.title = "Earnings report"
c1.height = 8.5
c1.width = 22
data1 = Reference(ws4, min_col=2, min_row=dh, max_row=last_data)
cats = Reference(ws4, min_col=1, min_row=dh + 1, max_row=last_data)
c1.add_data(data1, titles_from_data=True)
c1.set_categories(cats)
c1.series[0].smooth = False
c1.series[0].graphicalProperties.line.width = 28000
ws4.add_chart(c1, "E4")

# Chart 2 — organic sales-growth guidance trajectory
c2 = LineChart()
c2.title = "Abbott organic / comparable sales-growth guidance midpoint (where numeric)"
c2.style = 13
c2.y_axis.title = "Organic growth midpoint (%)"
c2.x_axis.title = "Earnings report"
c2.height = 8.5
c2.width = 22
data2 = Reference(ws4, min_col=3, min_row=dh, max_row=last_data)
c2.add_data(data2, titles_from_data=True)
c2.set_categories(cats)
c2.series[0].graphicalProperties.line.width = 28000
ws4.add_chart(c2, "E22")

ws4.column_dimensions["A"].width = 16
ws4.column_dimensions["B"].width = 20
ws4.column_dimensions["C"].width = 24

# final tab order
order = ["Guidance Timeline", "By Fiscal Year", "Growth to Guidance", "Trends", "Notes & Sources"]
wb._sheets.sort(key=lambda s: order.index(s.title))

out_path = "/home/user/knutnyman/Abbott_Guidance_By_Quarter_2020-2026.xlsx"
wb.save(out_path)
print("Saved:", out_path)
print("Rows:", len(ROWS), "| Sheets:", wb.sheetnames)
