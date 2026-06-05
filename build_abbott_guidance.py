#!/usr/bin/env python3
"""Build an Excel workbook summarizing Abbott Laboratories (ABT) full-year
guidance as issued at each quarterly earnings report, 2020 -> Q1 2026.

Data compiled from Abbott press releases (PR Newswire / abbott.mediaroom.com)
and SEC 8-K exhibits, cross-checked against secondary financial sources.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Source data: one record per quarterly report (chronological).
# adj_low / adj_high: numeric ends of the adjusted-EPS range (None if open/NA).
#   For "at least $X" guidance, adj_low = X and adj_high = None.
# ---------------------------------------------------------------------------
COLS = [
    "Report Date", "Quarter Reported", "FY Guided", "Action",
    "Adj. EPS Guidance", "Adj. EPS Low", "Adj. EPS High", "Adj. EPS Midpoint",
    "GAAP EPS Guidance", "Organic / Comparable Sales Growth",
    "COVID-19 Testing Sales Assumption", "Key Commentary", "Source",
]

ROWS = [
    # ---------------- FY2020 ----------------
    ["2020-01-22", "Q4 2019", "FY2020", "Initial",
     "$3.55 – $3.65", 3.55, 3.65, 3.60,
     "$2.35 – $2.45", "7.0% – 8.0%", "n/a (pre-COVID)",
     "Initial FY2020 outlook issued pre-COVID; double-digit EPS growth at midpoint.",
     "https://www.sec.gov/Archives/edgar/data/0000001800/000110465920005717/tm205456d1_ex99-1.htm"],

    ["2020-04-16", "Q1 2020", "FY2020", "Withdrawn",
     "Withdrawn / suspended", None, None, None,
     "Withdrawn / suspended", "Withdrawn", "Uncertain",
     "Suspended all FY2020 guidance due to COVID-19 uncertainty; no new numbers provided.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2020-results-301041836.html"],

    ["2020-07-16", "Q2 2020", "FY2020", "Reinstated (floor)",
     "At least $3.25", 3.25, None, None,
     "At least $2.00", "Not provided (qualitative)", "Ramping",
     "Guidance reinstated as an 'at least' floor given COVID uncertainty; ~$1.25 of specified items.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2020-results-exceeds-analysts-expectations-301094737.html"],

    ["2020-10-21", "Q3 2020", "FY2020", "Raised",
     "At least $3.55", 3.55, None, None,
     "At least $2.35", "Not provided (qualitative)", "Strong (BinaxNOW / ID NOW / Panbio)",
     "Raised from 'at least $3.25' on strong COVID diagnostics demand. FY2020 actual landed at $3.65 adj.",
     "https://abbott.mediaroom.com/2020-10-21-Abbott-Reports-Third-Quarter-2020-Results-Achieves-Strong-Double-Digit-Earnings-Growth-and-Raises-Guidance"],

    # ---------------- FY2021 ----------------
    ["2021-01-27", "Q4 2020", "FY2021", "Initial",
     "At least $5.00", 5.00, None, None,
     "At least $3.74", "Double-digit (qualitative)", "~$2.4B in Q4'20",
     "Initial FY2021 outlook: EPS growth of more than 35% vs prior year on COVID-testing boom.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-2020-results-issues-strong-double-digit-growth-forecast-for-2021-301216123.html"],

    ["2021-04-20", "Q1 2021", "FY2021", "Maintained",
     "At least $5.00", 5.00, None, None,
     "At least $3.74", "Not provided (qualitative)", "Strong",
     "Guidance maintained unchanged ('growth of more than 35%').",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2021-results-301272437.html"],

    ["2021-07-22", "Q2 2021", "FY2021", "Lowered",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$2.75 – $2.95", "Not provided (qualitative)", "Sharp decline",
     "Lowered on sharp drop in COVID-19 testing demand. (Abbott pre-announced a cut to 'at least $4.30' in mid-June 2021.)",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2021-results-301339347.html"],

    ["2021-10-20", "Q3 2021", "FY2021", "Raised",
     "$5.00 – $5.10", 5.00, 5.10, 5.05,
     "$3.55 – $3.65", "Not provided (qualitative)", "Renewed (Delta wave)",
     "Raised sharply (~38.4% growth at midpoint) on renewed COVID-testing demand plus base-business strength.",
     "https://abbott.mediaroom.com/2021-10-20-Abbott-Reports-Third-Quarter-2021-Results-Achieves-Strong-Double-Digit-Earnings-Growth-and-Raises-Guidance"],

    # ---------------- FY2022 ----------------
    ["2022-01-26", "Q4 2021", "FY2022", "Initial",
     "At least $4.70", 4.70, None, None,
     "At least $3.43", "Not stated as single number", "~$2.5B (initial)",
     "Initial FY2022 outlook with an initial COVID-testing sales forecast of ~$2.5B, to be updated quarterly.",
     "https://www.prnewswire.com/news-releases/abbott-reports-strong-fourth-quarter-2021-results-issues-2022-forecast-301468554.html"],

    ["2022-04-20", "Q1 2022", "FY2022", "Maintained",
     "At least $4.70", 4.70, None, None,
     "At least $3.35", "Not stated as single number", "Raised to ~$4.5B",
     "Adjusted EPS maintained; COVID-testing sales forecast raised to ~$4.5B (Q1 COVID test sales $3.3B).",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2022-results-301528820.html"],

    ["2022-07-20", "Q2 2022", "FY2022", "Raised",
     "At least $4.90", 4.90, None, None,
     "At least $3.50", "Low double digits (base business)", "Raised to $6.1B",
     "EPS raised (>= $4.70 -> >= $4.90); COVID-testing sales raised to $6.1B; strong base business.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2022-results-and-raises-full-year-eps-guidance-301590017.html"],

    ["2022-10-19", "Q3 2022", "FY2022", "Raised",
     "$5.17 – $5.23", 5.17, 5.23, 5.20,
     "$3.75 – $3.81", "Double-digit base business", "~$6.1–6.4B (implied)",
     "EPS raised again to a defined range. FY2022 actual landed at $5.34 adj.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2022-results-and-raises-full-year-eps-guidance-301653429.html"],

    # ---------------- FY2023 ----------------
    ["2023-01-25", "Q4 2022", "FY2023", "Initial",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.05 – $3.25", "High-single digits (ex-COVID)", "~$2.0B",
     "Initial FY2023 outlook; assumed ~$2.0B of COVID-testing sales for the year.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2022-results-issues-2023-financial-outlook-301730346.html"],

    ["2023-04-19", "Q1 2023", "FY2023", "Maintained",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.05 – $3.25", "At least high-single digits (ex-COVID)", "Lowered to ~$1.5B",
     "Adjusted EPS maintained; higher base-business outlook offset by lower COVID. Q1 base organic +10.0%.",
     "https://abbott.mediaroom.com/2023-04-19-Abbott-Reports-First-Quarter-2023-Results-Increases-Outlook-For-Underlying-Base-Business"],

    ["2023-07-20", "Q2 2023", "FY2023", "Maintained",
     "$4.30 – $4.50", 4.30, 4.50, 4.40,
     "$3.02 – $3.22", "Low double digits (ex-COVID)", "Lowered to ~$1.3B",
     "Adjusted EPS maintained; base-business outlook raised again, offset by lower COVID. Q2 base organic +11.5%.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2023-results-increases-outlook-for-underlying-base-business-301882027.html"],

    ["2023-10-18", "Q3 2023", "FY2023", "Narrowed (midpoint raised)",
     "$4.42 – $4.46", 4.42, 4.46, 4.44,
     "$3.14 – $3.18", "Low double digits (ex-COVID)", "~$1.6B (implied)",
     "EPS range narrowed and midpoint raised; base-business growth maintained.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2023-results-and-raises-midpoint-of-full-year-eps-guidance-range-301960446.html"],

    # ---------------- FY2024 ----------------
    ["2024-01-24", "Q4 2023", "FY2024", "Initial",
     "$4.50 – $4.70", 4.50, 4.70, 4.60,
     "$3.20 – $3.40", "8.0% – 10.0% (ex-COVID)", "Minimal",
     "Initial FY2024 outlook.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2023-results-issues-2024-financial-outlook-302043275.html"],

    ["2024-04-17", "Q1 2024", "FY2024", "Raised midpoint",
     "$4.55 – $4.70", 4.55, 4.70, 4.625,
     "$3.25 – $3.40", "8.5% – 10.0% (ex-COVID)", "Minimal",
     "Raised midpoint of all ranges; narrowed organic range by raising the low end.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2024-results-and-raises-midpoint-of-full-year-guidance-ranges-302119393.html"],

    ["2024-07-18", "Q2 2024", "FY2024", "Raised",
     "$4.61 – $4.71", 4.61, 4.71, 4.66,
     "$3.30 – $3.40", "9.5% – 10.0% (ex-COVID)", "Minimal",
     "Raised full-year guidance again.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2024-results-and-raises-full-year-guidance-302200584.html"],

    ["2024-10-16", "Q3 2024", "FY2024", "Raised midpoint / narrowed",
     "$4.64 – $4.70", 4.64, 4.70, 4.67,
     "$3.34 – $3.40", "9.5% – 10.0% (ex-COVID)", "Minimal",
     "Raised midpoint / narrowed EPS; organic range maintained.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2024-results-and-raises-midpoint-of-full-year-eps-guidance-range-302277755.html"],

    # ---------------- FY2025 ----------------
    ["2025-01-22", "Q4 2024", "FY2025", "Initial",
     "$5.05 – $5.25", 5.05, 5.25, 5.15,
     "Not provided (discontinued)", "7.5% – 8.5%", "Negligible",
     "Initial FY2025 outlook; double-digit EPS growth at midpoint; adj. operating margin 23.5%-24.0%. GAAP EPS guidance discontinued.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2024-results-issues-2025-financial-outlook-302357321.html"],

    ["2025-04-16", "Q1 2025", "FY2025", "Reaffirmed",
     "$5.05 – $5.25", 5.05, 5.25, 5.15,
     "Not provided (discontinued)", "7.5% – 8.5%", "Negligible",
     "Reaffirmed all FY2025 guidance.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2025-results-and-reaffirms-full-year-guidance-302430212.html"],

    ["2025-07-17", "Q2 2025", "FY2025", "Narrowed",
     "$5.10 – $5.20", 5.10, 5.20, 5.15,
     "Not provided (discontinued)", "7.5% – 8.0% ex-COVID (6.0% – 7.0% incl.)", "Negligible",
     "Narrowed adj. EPS range (midpoint unchanged); narrowed organic range.",
     "https://www.prnewswire.com/news-releases/abbott-reports-second-quarter-2025-results-302507875.html"],

    ["2025-10-15", "Q3 2025", "FY2025", "Reaffirmed / narrowed",
     "$5.12 – $5.18", 5.12, 5.18, 5.15,
     "Not provided (discontinued)", "7.5% – 8.0% ex-COVID (6.0% – 7.0% incl.)", "Negligible",
     "Reaffirmed midpoint, narrowed adj. EPS range; reaffirmed organic.",
     "https://www.prnewswire.com/news-releases/abbott-reports-third-quarter-2025-results-and-reaffirms-full-year-guidance-302584746.html"],

    # ---------------- FY2026 ----------------
    ["2026-01-22", "Q4 2025", "FY2026", "Initial",
     "$5.55 – $5.80", 5.55, 5.80, 5.675,
     "Not provided (discontinued)", "6.5% – 7.5%", "n/a",
     "Initial FY2026 outlook; ~10% EPS growth at midpoint. Announced acquisition of Exact Sciences.",
     "https://www.prnewswire.com/news-releases/abbott-reports-fourth-quarter-and-full-year-2025-results-issues-2026-financial-outlook-302668032.html"],

    ["2026-04-16", "Q1 2026", "FY2026", "Lowered (acquisition dilution)",
     "$5.38 – $5.58", 5.38, 5.58, 5.48,
     "Not provided (discontinued)", "6.5% – 7.5% (comparable basis)", "n/a",
     "EPS lowered to reflect ~$0.20 dilution from Exact Sciences (closed Mar 23, 2026); top-line range maintained, now stated on a 'comparable' basis. Not an operational guide-down.",
     "https://www.prnewswire.com/news-releases/abbott-reports-first-quarter-2026-results-updates-guidance-to-reflect-acquisition-of-exact-sciences-302744652.html"],
]

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------
NAVY = "1F3864"
LIGHT = "D9E1F2"
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

ws.merge_cells("A1:M1")
ws["A1"] = "Abbott Laboratories (NYSE: ABT) — Full-Year Guidance by Quarter"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A2:M2")
ws["A2"] = ("How Abbott's full-year outlook was set and revised at each quarterly earnings report, "
            "Q4 2019 report (initial FY2020) through Q1 2026.  Compiled from Abbott press releases & SEC 8-K exhibits.")
ws["A2"].font = SUB_FONT
ws.row_dimensions[1].height = 24

header_row = 4
for j, name in enumerate(COLS, start=1):
    ws.cell(row=header_row, column=j, value=name)
style_header_row(ws, header_row, len(COLS))
ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

r = header_row + 1
prev_fy = None
for rec in ROWS:
    fy = rec[2]
    band = (["F4F8FF", "FFFFFF"])  # subtle alternating by FY group
    for j, val in enumerate(rec, start=1):
        cell = ws.cell(row=r, column=j, value=val)
        cell.font = CELL_FONT
        cell.border = BORDER
        cell.alignment = WRAP
        if j in (2, 3, 6, 7, 8):  # quarter, FY, numeric EPS
            cell.alignment = CENTER
        if j in (6, 7, 8) and isinstance(val, float):
            cell.number_format = '$#,##0.00'
    # color the Action cell (col 4)
    action = rec[3]
    fill = ACTION_FILL.get(action, WHITE)
    ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=fill)
    ws.cell(row=r, column=4).alignment = CENTER
    # zebra by fiscal-year block
    block_fill = ALT if (["FY2020", "FY2022", "FY2024", "FY2026"].count(fy)) else WHITE
    for j in (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13):
        if ws.cell(row=r, column=j).fill.fgColor.rgb in (None, "00000000"):
            ws.cell(row=r, column=j).fill = PatternFill("solid", fgColor=block_fill)
    r += 1
    prev_fy = fy

# column widths
widths = [12, 13, 9, 20, 16, 9, 9, 11, 18, 26, 22, 46, 30]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
for rr in range(header_row + 1, r):
    ws.row_dimensions[rr].height = 58

# ---------------------------------------------------------------------------
# Sheet 2: By Fiscal Year (how guidance for each FY evolved)
# ---------------------------------------------------------------------------
ws2 = wb.create_sheet("By Fiscal Year")
ws2.merge_cells("A1:H1")
ws2["A1"] = "Evolution of Guidance — grouped by the fiscal year being guided"
ws2["A1"].font = TITLE_FONT
ws2.merge_cells("A2:H2")
ws2["A2"] = "Read across each block to see how the outlook for a given fiscal year changed quarter to quarter."
ws2["A2"].font = SUB_FONT

cols2 = ["FY Guided", "Set At (Report)", "Report Date", "Action",
         "Adj. EPS Guidance", "Adj. EPS Midpoint", "GAAP EPS Guidance", "Organic / Comparable Sales Growth"]
hr2 = 4
for j, name in enumerate(cols2, start=1):
    ws2.cell(row=hr2, column=j, value=name)
style_header_row(ws2, hr2, len(cols2))
ws2.freeze_panes = ws2.cell(row=hr2 + 1, column=1)

# group rows by FY preserving order
from collections import OrderedDict
groups = OrderedDict()
for rec in ROWS:
    groups.setdefault(rec[2], []).append(rec)

r2 = hr2 + 1
fy_colors = {"FY2020": "E2EFDA", "FY2021": "DDEBF7", "FY2022": "FCE4D6",
             "FY2023": "FFF2CC", "FY2024": "E2EFDA", "FY2025": "DDEBF7", "FY2026": "FCE4D6"}
for fy, recs in groups.items():
    start_r = r2
    for rec in recs:
        out = [rec[2], rec[1], rec[0], rec[3], rec[4], rec[7], rec[8], rec[9]]
        for j, val in enumerate(out, start=1):
            cell = ws2.cell(row=r2, column=j, value=val)
            cell.font = CELL_FONT
            cell.border = BORDER
            cell.alignment = WRAP
            if j in (1, 2, 3, 6):
                cell.alignment = CENTER
            if j == 6 and isinstance(val, float):
                cell.number_format = '$#,##0.00'
            cell.fill = PatternFill("solid", fgColor=fy_colors.get(fy, WHITE))
        ws2.cell(row=r2, column=4).fill = PatternFill("solid", fgColor=ACTION_FILL.get(rec[3], WHITE))
        ws2.cell(row=r2, column=4).alignment = CENTER
        r2 += 1
    # merge FY label cell down the block
    if r2 - start_r > 1:
        ws2.merge_cells(start_row=start_r, start_column=1, end_row=r2 - 1, end_column=1)
        ws2.cell(row=start_r, column=1).alignment = Alignment(horizontal="center", vertical="center")

widths2 = [10, 16, 12, 24, 16, 12, 20, 34]
for i, w in enumerate(widths2, start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w
for rr in range(hr2 + 1, r2):
    ws2.row_dimensions[rr].height = 30

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
    "KEY STRUCTURAL POINTS",
    "• 2020-2022 were dominated by COVID-19 testing. Abbott withdrew FY2020 guidance entirely in Apr 2020, reinstated it as an",
    "  'at least' floor mid-year, then raised it. FY2021-2022 guidance swung sharply with COVID-test demand (e.g., FY2021 was",
    "  cut in Q2 2021 as testing fell, then raised in Q3 2021 during the Delta wave).",
    "• 'At least $X' entries are floor guidance (common in 2020-2022); numeric Low/High columns leave the High blank for these.",
    "• GAAP EPS guidance was discontinued starting with the FY2025 outlook (Jan 2025): Abbott states it cannot reliably forecast",
    "  special items (restructuring, impairments, acquisition charges). 'Not provided (discontinued)' reflects that policy change,",
    "  not missing data.",
    "• Organic sales-growth guidance was given as an explicit % range only at the initial FY2020 outlook and from FY2024 onward.",
    "  In 2020-2023 Abbott mostly guided on EPS and described sales qualitatively (e.g., 'double-digit base business ex-COVID').",
    "  From FY2024 the headline organic figure excludes COVID testing; FY2025 gave both ex-COVID and all-in ranges.",
    "• Q1 2026: adjusted EPS was lowered purely to reflect ~$0.20 of dilution from the Exact Sciences acquisition (closed Mar 23,",
    "  2026); the top-line growth range was maintained and restated on a 'comparable' basis. This is not an operational guide-down.",
    "",
    "DATA CONFIDENCE",
    "• Adjusted-EPS figures are the highest-confidence numbers and were corroborated across at least two sources per quarter.",
    "• A few items carry minor uncertainty: the initial FY2022 GAAP floor (cited as ~$3.43 vs $3.40) and the exact restated",
    "  COVID-testing-sales assumptions at Q3 2022 / Q3 2023 (only the quarter's actual COVID sales were confirmed).",
    "",
    "REPORTING-DATE NOTES",
    "• FY2021 initial guidance was issued Jan 27, 2021 (Q4 2020 report). FY2021 Q2 cut was first signaled off-cycle in mid-June 2021,",
    "  then formalized as the $4.30-$4.50 range at the Jul 22, 2021 report.",
    "",
    "PRIMARY SOURCES",
    "Abbott investor news (abbott.mediaroom.com) and PR Newswire press releases, plus the corresponding SEC 8-K exhibit 99.1 for",
    "each quarter. Per-row source links are in the 'Source' column of the Guidance Timeline sheet. Secondary cross-checks included",
    "Nasdaq, StockTitan, RTTNews, MedTech Dive, 360Dx and Investing.com.",
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

# legend on notes sheet
lr = len(notes) + 4
ws3.cell(row=lr, column=2, value="ACTION COLOR LEGEND").font = Font(bold=True, color=NAVY, size=11)
legend = [("Initial", "Initial outlook for a fiscal year"),
          ("Raised", "Guidance raised vs prior update"),
          ("Maintained / Reaffirmed / Narrowed", "Held or tightened around the same midpoint"),
          ("Reinstated (floor)", "Guidance restored as an 'at least' floor"),
          ("Lowered", "Guidance reduced vs prior update"),
          ("Withdrawn", "Guidance suspended")]
for k, (label, desc) in enumerate(legend, start=lr + 1):
    ws3.cell(row=k, column=1).fill = PatternFill(
        "solid", fgColor=ACTION_FILL.get(label.split(" /")[0], "F2F2F2"))
    ws3.cell(row=k, column=1).border = BORDER
    c = ws3.cell(row=k, column=2, value=f"{label} — {desc}")
    c.font = Font(name="Calibri", size=10)

out_path = "/home/user/knutnyman/Abbott_Guidance_By_Quarter_2020-2026.xlsx"
wb.save(out_path)
print("Saved:", out_path)
print("Rows:", len(ROWS), "| Sheets:", wb.sheetnames)
