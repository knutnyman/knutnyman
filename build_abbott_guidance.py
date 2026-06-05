#!/usr/bin/env python3
"""Build an Excel workbook summarizing Abbott Laboratories (ABT) full-year
guidance as issued at each quarterly earnings report, 2020 -> Q1 2026.

Data compiled from Abbott press releases (PR Newswire / abbott.mediaroom.com)
and SEC 8-K exhibits, cross-checked against secondary financial sources.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

out_path = "/home/user/knutnyman/Abbott_Guidance_By_Quarter_2020-2026.xlsx"
wb.save(out_path)
print("Saved:", out_path)
print("Rows:", len(ROWS), "| Sheets:", wb.sheetnames)
