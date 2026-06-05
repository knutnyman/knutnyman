#!/usr/bin/env python3
"""Build the Abbott Nutrition & EPD precedent-transaction analysis as a formatted,
formula-driven Excel workbook."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

wb = openpyxl.Workbook()

# ---- palette / reusable styles ---------------------------------------------
NAVY   = "1F3864"
BLUE   = "2E5496"
LBLUE  = "D9E1F2"
GREY   = "F2F2F2"
GOLD   = "FFF2CC"
GREEN  = "E2EFDA"
WHITE  = "FFFFFF"

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
title_font  = Font(name="Calibri", size=16, bold=True, color=WHITE)
sub_font    = Font(name="Calibri", size=10, italic=True, color="595959")
hdr_font    = Font(name="Calibri", size=10, bold=True, color=WHITE)
secn_font   = Font(name="Calibri", size=11, bold=True, color=NAVY)
bold        = Font(name="Calibri", size=10, bold=True)
reg         = Font(name="Calibri", size=10)
input_font  = Font(name="Calibri", size=10, color="0000CC")  # blue = input cell
hdr_fill    = PatternFill("solid", fgColor=BLUE)
navy_fill   = PatternFill("solid", fgColor=NAVY)
lblue_fill  = PatternFill("solid", fgColor=LBLUE)
grey_fill   = PatternFill("solid", fgColor=GREY)
gold_fill   = PatternFill("solid", fgColor=GOLD)
green_fill  = PatternFill("solid", fgColor=GREEN)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
right  = Alignment(horizontal="right", vertical="center")

USD = '#,##0.0,,"bn"'      # value stored in $m, shown as $bn
USD0 = '#,##0'
MULT = '0.0"x"'
PCT = '0.0%'

def hrow(ws, r, c0, headers, fill=hdr_fill, font=hdr_font):
    for i, h in enumerate(headers):
        cell = ws.cell(row=r, column=c0+i, value=h)
        cell.font = font; cell.fill = fill; cell.alignment = center; cell.border = border

def put(ws, r, c, v, font=reg, fill=None, fmt=None, align=None, bd=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    cell.alignment = align or left
    if bd: cell.border = border
    return cell

def banner(ws, ncols, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title); c.font = title_font; c.fill = navy_fill
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(row=2, column=1, value=subtitle); s.font = sub_font
    s.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ============================================================ INPUTS sheet
ws = wb.active; ws.title = "Inputs & Assumptions"
ws.sheet_view.showGridLines = False
banner(ws, 6, "Abbott — Nutrition & EPD Divestiture  |  Inputs & Assumptions",
        "Blue cells are inputs — change them and the Valuation/Summary tabs recalc. FY2024 reported segment basis. EBITDA is estimated (see notes).")
put(ws, 4, 1, "Blue figures are drivers. $ figures in US$ millions unless noted.", sub_font, bd=False)

put(ws, 6, 1, "Target sizing (FY2024)", secn_font, bd=False)
hrow(ws, 7, 1, ["Metric", "Nutrition", "EPD", "Source / note"])
rows = [
    ("FY24 net sales ($m)",            8410, 5190, "Abbott FY24 segment reporting (~$8.4bn / ~$5.2bn)"),
    ("FY24 segment operating margin",  0.179, 0.237, "Disclosed: Nutrition 17.9%, EPD 23.7%"),
    ("Assumed D&A add-back (% sales)", 0.0375, 0.0275, "Estimate — Nutrition more capital-intensive"),
]
r = 8
for label, n, e, note in rows:
    put(ws, r, 1, label, bold)
    fmt = PCT if label != "FY24 net sales ($m)" else USD0
    put(ws, r, 2, n, input_font, fmt=fmt, align=right)
    put(ws, r, 3, e, input_font, fmt=fmt, align=right)
    put(ws, r, 4, note, reg)
    r += 1
# derived operating earnings & EBITDA (formulas)
put(ws, r, 1, "Implied segment operating earnings ($m)", bold, grey_fill)
put(ws, r, 2, "=B8*B9", reg, grey_fill, USD0, right)
put(ws, r, 3, "=C8*C9", reg, grey_fill, USD0, right)
put(ws, r, 4, "= sales x operating margin", sub_font, grey_fill); op_row = r; r += 1
put(ws, r, 1, "Estimated EBITDA ($m)", bold, green_fill)
put(ws, r, 2, "=B8*(B9+B10)", bold, green_fill, USD0, right)
put(ws, r, 3, "=C8*(C9+C10)", bold, green_fill, USD0, right)
put(ws, r, 4, "= operating earnings + D&A add-back", sub_font, green_fill); ebitda_row = r; r += 1
put(ws, r, 1, "Implied EBITDA margin", reg, green_fill)
put(ws, r, 2, f"=B{ebitda_row}/B8", reg, green_fill, PCT, right)
put(ws, r, 3, f"=C{ebitda_row}/C8", reg, green_fill, PCT, right)
put(ws, r, 4, "", reg, green_fill); r += 2

# selected multiple ranges
put(ws, r, 1, "Selected multiple ranges (judgement)", secn_font, bd=False); r += 1
hrow(ws, r, 1, ["Multiple", "Nutrition low", "Nutrition high", "EPD low", "EPD high"]); r += 1
mult_start = r
put(ws, r, 1, "EV / Sales", bold)
for c, v in zip((2,3,4,5), (2.0, 3.0, 2.5, 3.5)):
    put(ws, r, c, v, input_font, fmt=MULT, align=right)
r += 1
put(ws, r, 1, "EV / EBITDA", bold)
for c, v in zip((2,3,4,5), (11.0, 15.0, 11.0, 14.0)):
    put(ws, r, c, v, input_font, fmt=MULT, align=right)
r += 2
put(ws, r, 1, "Rationale", secn_font, bd=False); r += 1
notes = [
    "Nutrition: peak infant-formula comps (Reckitt/MJN 17.4x, Nestlé/Pfizer ~20x) discounted for blended adult+pediatric mix and the NEC litigation overhang.",
    "EPD: premium end of branded-generics range (STADA 13.4x / Krka ~11.5x) given +9.2% growth and 23.7% margin; above mature Western generics (Hikma/Richter ~6x).",
    "EBITDA is ESTIMATED — Abbott discloses segment operating margin, not EBITDA. Refine with carve-out quality-of-earnings financials.",
]
for n in notes:
    put(ws, r, 1, "•  " + n, reg, bd=False); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 28; r += 1

ws.cell(row=ebitda_row, column=1).comment = Comment(
    "Abbott does not publish EBITDA by segment. Estimated by adding assumed D&A to disclosed segment operating earnings.", "Analysis")
widths = [40, 14, 14, 14, 14, 14]
for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w

INP = "'Inputs & Assumptions'"
EB_N, EB_E = f"{INP}!B{ebitda_row}", f"{INP}!C{ebitda_row}"
SAL_N, SAL_E = f"{INP}!B8", f"{INP}!C8"
MS = mult_start  # EV/Sales row ; EV/EBITDA = MS+1

# ============================================================ NUTRITION COMPS
def comps_sheet(title, banner_sub, data):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    banner(ws, 7, title, banner_sub)
    hrow(ws, 4, 1, ["Year", "Acquirer", "Target", "EV ($bn)", "EV/Sales", "EV/EBITDA", "Notes"])
    r = 5
    for row in data:
        put(ws, r, 1, row[0], reg, align=center)
        put(ws, r, 2, row[1], reg)
        put(ws, r, 3, row[2], bold)
        put(ws, r, 4, row[3], reg, fmt='#,##0.0', align=right)
        put(ws, r, 5, row[4], reg, fmt=MULT, align=right)
        put(ws, r, 6, row[5], reg, fmt=MULT, align=right)
        put(ws, r, 7, row[6], reg)
        if r % 2 == 0:
            for c in range(1, 8): ws.cell(row=r, column=c).fill = grey_fill
        r += 1
    # stats row (median / mean of available multiples)
    first, last = 5, r-1
    put(ws, r, 1, "Median", bold, lblue_fill); put(ws, r, 2, "", reg, lblue_fill); put(ws, r, 3, "", reg, lblue_fill)
    put(ws, r, 4, f"=MEDIAN(D{first}:D{last})", bold, lblue_fill, '#,##0.0', right)
    put(ws, r, 5, f"=MEDIAN(E{first}:E{last})", bold, lblue_fill, MULT, right)
    put(ws, r, 6, f"=MEDIAN(F{first}:F{last})", bold, lblue_fill, MULT, right)
    put(ws, r, 7, "", reg, lblue_fill); r += 1
    put(ws, r, 1, "Mean", bold, lblue_fill); put(ws, r, 2, "", reg, lblue_fill); put(ws, r, 3, "", reg, lblue_fill)
    put(ws, r, 4, f"=AVERAGE(D{first}:D{last})", bold, lblue_fill, '#,##0.0', right)
    put(ws, r, 5, f"=AVERAGE(E{first}:E{last})", bold, lblue_fill, MULT, right)
    put(ws, r, 6, f"=AVERAGE(F{first}:F{last})", bold, lblue_fill, MULT, right)
    put(ws, r, 7, "", reg, lblue_fill)
    for i, w in enumerate([8, 22, 26, 11, 12, 12, 50], 1): ws.column_dimensions[get_column_letter(i)].width = w
    return ws

nutrition = [
    (2007, "Danone", "Numico", 16.8, 4.5, 22.0, "Baby + medical nutrition; peak strategic multiple"),
    (2007, "Nestlé", "Gerber (from Novartis)", 5.5, 2.8, 20.0, "EBITDA mult. est.; made Nestlé #1 baby food"),
    (2012, "Nestlé", "Pfizer Nutrition", 11.85, 5.0, 22.0, "~$2.1bn sales / ~$0.5bn EBITDA (2011); EM (China) entry"),
    (2017, "Danone", "WhiteWave", 12.5, 2.7, 20.0, "Organic/plant-based (not infant); lower-multiple bookend"),
    (2017, "Reckitt Benckiser", "Mead Johnson", 17.9, 4.4, 17.4, "Pure-play infant formula; closest single comp (~14x w/ synergies)"),
    (2021, "Nestlé", "The Bountiful Company", 5.75, 2.7, 17.0, "Vitamins/supplements; adult-nutrition-adjacent"),
]
comps_sheet("Nutrition Comps",
            "Precedent transactions — infant + adult/medical nutrition. Multiples as reported by public sources; treat as directional.",
            nutrition)

epd = [
    (2014, "Sun Pharma", "Ranbaxy", 4.0, 2.2, None, "Incl. debt; low-margin target (EBITDA mult. n/m); India/EM"),
    (2015, "Teva", "Allergan Generics (Actavis)", 40.5, 3.6, 11.5, "EBITDA mult. est.; global generics scale"),
    (2016, "Mylan", "Meda", 9.9, 3.3, 10.5, "8.9x w/ synergies; ~10–11x reported; branded/specialty + EM/OTC"),
    (2017, "Bain / Cinven", "STADA", 5.7, 2.4, 13.4, "EUR 5.32bn; PE take-private; branded generics + consumer health"),
    (2018, "Advent", "Zentiva (Sanofi EU gen.)", 2.1, 2.5, 10.3, "EUR ~1.9bn; EU/EM branded-generics carve-out"),
    (2025, "CapVest", "STADA (majority stake)", 10.0, 3.0, 11.3, "Reported ~EUR 10bn; latest large branded-generics print"),
]
ws_epd = comps_sheet("EPD Comps",
            "Precedent transactions — branded / emerging-market generics. Listed reads for context: Krka ~11.5x, Hikma ~6.3x, Richter ~6x EV/EBITDA.",
            epd)
# EPD has a None EBITDA mult -> ensure MEDIAN/AVERAGE skip blank: rewrite that cell blank
# (already None -> openpyxl writes empty; formulas ignore text/empty)

# ============================================================ VALUATION
ws = wb.create_sheet("Valuation")
ws.sheet_view.showGridLines = False
banner(ws, 6, "Valuation — Implied Enterprise Value", "Driven by Inputs & Assumptions tab. $ in US$ millions; EV summary in $bn.")
def val_block(ws, r, name, sal, eb, ms_low_col, ms_high_col, me_low_col, me_high_col):
    put(ws, r, 1, name, secn_font, bd=False); r += 1
    hrow(ws, r, 1, ["Method", "Multiple low", "Multiple high", "EV low ($m)", "EV high ($m)", "EV midpoint ($bn)"]); r += 1
    # EV/Sales
    put(ws, r, 1, "EV / Sales", bold)
    put(ws, r, 2, f"={INP}!{ms_low_col}{MS}", reg, fmt=MULT, align=right)
    put(ws, r, 3, f"={INP}!{ms_high_col}{MS}", reg, fmt=MULT, align=right)
    put(ws, r, 4, f"={sal}*B{r}", reg, fmt=USD0, align=right)
    put(ws, r, 5, f"={sal}*C{r}", reg, fmt=USD0, align=right)
    put(ws, r, 6, f"=AVERAGE(D{r}:E{r})/1000", reg, fmt='#,##0.0"bn"', align=right)
    sales_r = r; r += 1
    # EV/EBITDA
    put(ws, r, 1, "EV / EBITDA", bold)
    put(ws, r, 2, f"={INP}!{me_low_col}{MS+1}", reg, fmt=MULT, align=right)
    put(ws, r, 3, f"={INP}!{me_high_col}{MS+1}", reg, fmt=MULT, align=right)
    put(ws, r, 4, f"={eb}*B{r}", reg, fmt=USD0, align=right)
    put(ws, r, 5, f"={eb}*C{r}", reg, fmt=USD0, align=right)
    put(ws, r, 6, f"=AVERAGE(D{r}:E{r})/1000", reg, fmt='#,##0.0"bn"', align=right)
    eb_r = r; r += 1
    # blended
    put(ws, r, 1, "Blended indicative EV", bold, gold_fill)
    put(ws, r, 2, "", reg, gold_fill); put(ws, r, 3, "", reg, gold_fill)
    put(ws, r, 4, f"=MIN(D{sales_r}:E{eb_r})", bold, gold_fill, USD0, right)
    put(ws, r, 5, f"=MAX(D{sales_r}:E{eb_r})", bold, gold_fill, USD0, right)
    put(ws, r, 6, f"=AVERAGE(D{r}:E{r})/1000", bold, gold_fill, '#,##0.0"bn"', right)
    blend_r = r
    return r + 2, blend_r

r = 4
r, nut_blend = val_block(ws, r, "Nutrition", SAL_N, EB_N, "B", "C", "B", "C")
r, epd_blend = val_block(ws, r, "Established Pharmaceuticals (EPD)", SAL_E, EB_E, "D", "E", "D", "E")
# combined
put(ws, r, 1, "Combined (Nutrition + EPD)", secn_font, bd=False); r += 1
hrow(ws, r, 1, ["", "", "", "EV low ($m)", "EV high ($m)", "EV midpoint ($bn)"]); r += 1
put(ws, r, 1, "Total indicative enterprise value", bold, green_fill)
put(ws, r, 2, "", reg, green_fill); put(ws, r, 3, "", reg, green_fill)
put(ws, r, 4, f"=D{nut_blend}+D{epd_blend}", bold, green_fill, USD0, right)
put(ws, r, 5, f"=E{nut_blend}+E{epd_blend}", bold, green_fill, USD0, right)
put(ws, r, 6, f"=AVERAGE(D{r}:E{r})/1000", bold, green_fill, '#,##0.0"bn"', right)
for i, w in enumerate([34, 14, 14, 15, 15, 18], 1): ws.column_dimensions[get_column_letter(i)].width = w

# ============================================================ SUMMARY (first tab)
ws = wb.create_sheet("Summary", 0)
ws.sheet_view.showGridLines = False
banner(ws, 6, "Abbott — Precedent Transaction Analysis", "Potential divestiture of the Nutrition and Established Pharmaceuticals (EPD) businesses  •  FY2024 basis  •  Indicative / discussion draft")
put(ws, 4, 1, "Headline indicative enterprise value (EV) — see Valuation tab for build, Inputs tab for assumptions.", sub_font, bd=False)
hrow(ws, 6, 1, ["Business", "FY24 Sales ($bn)", "Est. EBITDA ($bn)", "EV/Sales", "EV/EBITDA", "Indicative EV ($bn)"])
VAL = "Valuation"
def srow(r, name, sal_ref, eb_ref, evs_lo, evs_hi, eve_lo, eve_hi, lo_ref, hi_ref, fill=None):
    put(ws, r, 1, name, bold, fill)
    put(ws, r, 2, f"={sal_ref}/1000", reg, fill, '#,##0.0', right)
    put(ws, r, 3, f"={eb_ref}/1000", reg, fill, '#,##0.0', right)
    put(ws, r, 4, evs_lo and f'=TEXT({INP}!{evs_lo},"0.0")&"–"&TEXT({INP}!{evs_hi},"0.0")&"x"' or "", reg, fill, align=center)
    put(ws, r, 5, eve_lo and f'=TEXT({INP}!{eve_lo},"0.0")&"–"&TEXT({INP}!{eve_hi},"0.0")&"x"' or "", reg, fill, align=center)
    put(ws, r, 6, f'=TEXT({VAL}!{lo_ref}/1000,"0")&"–"&TEXT({VAL}!{hi_ref}/1000,"0")&"bn"', bold, fill, align=center)

srow(7, "Nutrition", SAL_N, EB_N, f"B{MS}", f"C{MS}", f"B{MS+1}", f"C{MS+1}", f"D{nut_blend}", f"E{nut_blend}")
srow(8, "Established Pharmaceuticals (EPD)", SAL_E, EB_E, f"D{MS}", f"E{MS}", f"D{MS+1}", f"E{MS+1}", f"D{epd_blend}", f"E{epd_blend}", grey_fill)
# combined row
put(ws, 9, 1, "Combined", Font(name="Calibri", size=11, bold=True, color=NAVY), green_fill)
put(ws, 9, 2, f"=({SAL_N}+{SAL_E})/1000", bold, green_fill, '#,##0.0', right)
put(ws, 9, 3, f"=({EB_N}+{EB_E})/1000", bold, green_fill, '#,##0.0', right)
put(ws, 9, 4, "—", reg, green_fill, align=center)
put(ws, 9, 5, "—", reg, green_fill, align=center)
put(ws, 9, 6, f'=TEXT(({VAL}!D{nut_blend}+{VAL}!D{epd_blend})/1000,"0")&"–"&TEXT(({VAL}!E{nut_blend}+{VAL}!E{epd_blend})/1000,"0")&"bn"', bold, green_fill, align=center)

put(ws, 11, 1, "Key points", secn_font, bd=False)
pts = [
    "EPD is the cleaner asset — high-growth (+9.2%), high-margin (23.7%) emerging-market branded generics; sits at the premium end of the branded-generics deal range.",
    "Nutrition is two-speed: attractive, growing adult/medical nutrition (Ensure, Glucerna) vs. a structurally challenged infant-formula half (Similac) carrying an NEC product-liability litigation overhang.",
    "Peak infant-formula comps (Reckitt/MJN 17.4x, Nestlé/Pfizer ~20x) are discounted for Nutrition's blended mix and litigation risk.",
    "Sum-of-the-parts optionality: selling adult nutrition at consumer-health multiples while ring-fencing infant formula could exceed the blended range.",
    "EBITDA is ESTIMATED (Abbott discloses segment operating margin, not EBITDA). Indicative only — not investment advice; refine with carve-out QoE financials.",
]
rr = 12
for p in pts:
    put(ws, rr, 1, "•  " + p, reg, bd=False); ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
    ws.row_dimensions[rr].height = 30; rr += 1
put(ws, rr+1, 1, "Tabs:  Summary  ·  Inputs & Assumptions (drivers)  ·  Nutrition Comps  ·  EPD Comps  ·  Valuation (build)", sub_font, bd=False)
for i, w in enumerate([38, 16, 17, 14, 14, 20], 1): ws.column_dimensions[get_column_letter(i)].width = w

# order tabs
wb.move_sheet("Summary", -wb.sheetnames.index("Summary"))
out = "/home/user/knutnyman/abbott-precedent-transaction-analysis/Abbott_Nutrition_and_EPD_Precedent_Transaction_Analysis.xlsx"
wb.save(out)
print("saved", out)
print("sheets:", wb.sheetnames)
