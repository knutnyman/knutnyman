#!/usr/bin/env python3
"""Build the Abbott Nutrition & EPD precedent-transaction analysis as a formatted,
formula-driven Excel workbook, with full sourcing / estimate flags on every figure."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

wb = openpyxl.Workbook()

# ---- palette / styles -------------------------------------------------------
NAVY="1F3864"; BLUE="2E5496"; LBLUE="D9E1F2"; GREY="F2F2F2"; GOLD="FFF2CC"; GREEN="E2EFDA"; WHITE="FFFFFF"
thin=Side(style="thin", color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
title_font=Font(name="Calibri",size=16,bold=True,color=WHITE)
sub_font=Font(name="Calibri",size=10,italic=True,color="595959")
hdr_font=Font(name="Calibri",size=10,bold=True,color=WHITE)
secn_font=Font(name="Calibri",size=11,bold=True,color=NAVY)
bold=Font(name="Calibri",size=10,bold=True); reg=Font(name="Calibri",size=10)
input_font=Font(name="Calibri",size=10,color="0000CC")
link_font=Font(name="Calibri",size=9,color="0563C1",underline="single")
navy_fill=PatternFill("solid",fgColor=NAVY); hdr_fill=PatternFill("solid",fgColor=BLUE)
lblue_fill=PatternFill("solid",fgColor=LBLUE); grey_fill=PatternFill("solid",fgColor=GREY)
gold_fill=PatternFill("solid",fgColor=GOLD); green_fill=PatternFill("solid",fgColor=GREEN)
center=Alignment(horizontal="center",vertical="center",wrap_text=True)
left=Alignment(horizontal="left",vertical="center",wrap_text=True)
right=Alignment(horizontal="right",vertical="center")
USD0='#,##0'; MULT='0.0"x"'; PCT='0.0%'

def hrow(ws,r,c0,headers,fill=hdr_fill,font=hdr_font):
    for i,h in enumerate(headers):
        c=ws.cell(row=r,column=c0+i,value=h); c.font=font; c.fill=fill; c.alignment=center; c.border=border

def put(ws,r,c,v,font=reg,fill=None,fmt=None,align=None,bd=True):
    cell=ws.cell(row=r,column=c,value=v); cell.font=font
    if fill: cell.fill=fill
    if fmt: cell.number_format=fmt
    cell.alignment=align or left
    if bd: cell.border=border
    return cell

def banner(ws,ncols,title,subtitle):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
    c=ws.cell(row=1,column=1,value=title); c.font=title_font; c.fill=navy_fill
    c.alignment=Alignment(horizontal="left",vertical="center",indent=1); ws.row_dimensions[1].height=30
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncols)
    s=ws.cell(row=2,column=1,value=subtitle); s.font=sub_font
    s.alignment=Alignment(horizontal="left",vertical="center",indent=1)

# ============================================================ INPUTS sheet
ws=wb.active; ws.title="Inputs & Assumptions"; ws.sheet_view.showGridLines=False
banner(ws,6,"Abbott — Nutrition & EPD Divestiture  |  Inputs & Assumptions",
       "Blue = input cells (change to recalc). FY2024 reported basis. EBITDA is ESTIMATED — see 'Sources & Methodology' tab for the basis of every figure.")
put(ws,4,1,"Blue figures are drivers. $ figures in US$ millions unless noted. Ref column shows source/estimate basis (full detail on Sources tab).",sub_font,bd=False)
put(ws,6,1,"Target sizing (FY2024)",secn_font,bd=False)
hrow(ws,7,1,["Metric","Nutrition","EPD","Basis / source (see Sources tab)"])
rows=[
 ("FY24 net sales ($m)",8410,5190,"REPORTED — Abbott FY24 segment reporting [S1, S2]"),
 ("FY24 segment operating margin",0.179,0.237,"REPORTED — Abbott (Nutrition 17.9%, EPD 23.7%) [S3]"),
 ("Assumed D&A add-back (% sales)",0.0375,0.0275,"ESTIMATE — benchmarked to group D&A ~7.7% of sales (amort. concentrated in MedDev/Dx) [S4]"),
]
r=8
for label,n,e,note in rows:
    put(ws,r,1,label,bold)
    fmt=PCT if label!="FY24 net sales ($m)" else USD0
    put(ws,r,2,n,input_font,fmt=fmt,align=right); put(ws,r,3,e,input_font,fmt=fmt,align=right)
    put(ws,r,4,note,reg); r+=1
put(ws,r,1,"Implied segment operating earnings ($m)",bold,grey_fill)
put(ws,r,2,"=B8*B9",reg,grey_fill,USD0,right); put(ws,r,3,"=C8*C9",reg,grey_fill,USD0,right)
put(ws,r,4,"DERIVED = sales x operating margin",sub_font,grey_fill); op_row=r; r+=1
put(ws,r,1,"Estimated EBITDA ($m)",bold,green_fill)
put(ws,r,2,"=B8*(B9+B10)",bold,green_fill,USD0,right); put(ws,r,3,"=C8*(C9+C10)",bold,green_fill,USD0,right)
put(ws,r,4,"ESTIMATE = operating earnings + D&A add-back",sub_font,green_fill); ebitda_row=r; r+=1
put(ws,r,1,"Implied EBITDA margin",reg,green_fill)
put(ws,r,2,f"=B{ebitda_row}/B8",reg,green_fill,PCT,right); put(ws,r,3,f"=C{ebitda_row}/C8",reg,green_fill,PCT,right)
put(ws,r,4,"DERIVED",sub_font,green_fill); r+=2
# ---- carve-out / standalone adjustments ----
put(ws,r,1,"Carve-out / standalone adjustments",secn_font,bd=False); r+=1
hrow(ws,r,1,["Adjustment","Nutrition","EPD","Basis / source (see Sources tab)"]); r+=1
put(ws,r,1,"Standalone cost / dis-synergy (% of sales)",bold)
put(ws,r,2,0.015,input_font,fmt=PCT,align=right); put(ws,r,3,0.020,input_font,fmt=PCT,align=right)
put(ws,r,4,"ESTIMATE — corporate functions a standalone unit must rebuild; typical carve-out 1-3% of sales [S20]",reg); stdcost_row=r; r+=1
put(ws,r,1,"Standalone-adjusted EBITDA ($m)",bold,green_fill)
put(ws,r,2,f"=B{ebitda_row}-B{stdcost_row}*B8",bold,green_fill,USD0,right)
put(ws,r,3,f"=C{ebitda_row}-C{stdcost_row}*C8",bold,green_fill,USD0,right)
put(ws,r,4,"ESTIMATE = estimated EBITDA - standalone cost (drives EV/EBITDA). PE buyer bears this; a strategic may offset via synergies.",sub_font,green_fill); stand_row=r; r+=1
put(ws,r,1,"NEC litigation risk deduction ($m, applied to Nutrition EV)",bold)
put(ws,r,2,1500,input_font,fmt=USD0,align=right); put(ws,r,3,0,input_font,fmt=USD0,align=right)
put(ws,r,4,"ESTIMATE / ILLUSTRATIVE — buyer indemnity ask; not a booked Abbott provision. Verdicts $58-495m, many on appeal [S19, S21]",reg); nec_row=r; r+=2
put(ws,r,1,"Selected multiple ranges (judgement — see rationale below)",secn_font,bd=False); r+=1
hrow(ws,r,1,["Multiple","Nutrition low","Nutrition high","EPD low","EPD high"]); r+=1
mult_start=r
put(ws,r,1,"EV / Sales",bold)
for c,v in zip((2,3,4,5),(2.0,3.0,2.5,3.5)): put(ws,r,c,v,input_font,fmt=MULT,align=right)
r+=1
put(ws,r,1,"EV / EBITDA",bold)
for c,v in zip((2,3,4,5),(11.0,15.0,11.0,14.0)): put(ws,r,c,v,input_font,fmt=MULT,align=right)
r+=2
put(ws,r,1,"Rationale for selected ranges",secn_font,bd=False); r+=1
for n in [
 "Nutrition: peak infant-formula comps (Reckitt/MJN 17.4x, Nestlé/Pfizer ~20-24x) DISCOUNTED for blended adult+pediatric mix and the NEC litigation overhang [S5].",
 "EPD: premium end of branded-generics range (STADA 13.4x / 11.3x; Krka ~11.5x) given +9.2% growth and 23.7% margin; above mature Western generics (Hikma/Richter ~6x) [S6].",
 "EBITDA is ESTIMATED — Abbott discloses segment operating margin, not segment EBITDA. Refine with carve-out quality-of-earnings financials [S4].",
]:
    put(ws,r,1,"•  "+n,reg,bd=False); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
    ws.row_dimensions[r].height=28; r+=1
ws.cell(row=ebitda_row,column=1).comment=Comment(
 "ESTIMATE. Abbott does not publish EBITDA by segment. = disclosed segment operating earnings + assumed D&A add-back. See Sources tab note S4.","Analysis")
for i,w in enumerate([42,13,13,52],1): ws.column_dimensions[get_column_letter(i)].width=w
INP="'Inputs & Assumptions'"
# EV/EBITDA is applied to the STANDALONE-adjusted EBITDA (post dis-synergy)
EB_N,EB_E=f"{INP}!B{stand_row}",f"{INP}!C{stand_row}"
EBPRE_N,EBPRE_E=f"{INP}!B{ebitda_row}",f"{INP}!C{ebitda_row}"
SAL_N,SAL_E=f"{INP}!B8",f"{INP}!C8"
NEC_N=f"{INP}!B{nec_row}"
MS=mult_start
ME=mult_start+1

# ============================================================ COMPS
def comps_sheet(title,banner_sub,data):
    ws=wb.create_sheet(title); ws.sheet_view.showGridLines=False
    banner(ws,9,title,banner_sub)
    hrow(ws,4,1,["Year","Acquirer","Target","EV ($bn)","EV/Sales","EV/EBITDA","Mult. basis","Source ref","Notes"])
    r=5
    for row in data:
        yr,acq,tgt,ev,evs,eve,basis,srcref,notes=row
        put(ws,r,1,yr,reg,align=center); put(ws,r,2,acq,reg); put(ws,r,3,tgt,bold)
        put(ws,r,4,ev,reg,fmt='#,##0.0',align=right)
        put(ws,r,5,evs,reg,fmt=MULT,align=right) if evs is not None else put(ws,r,5,"n/d",reg,align=center)
        put(ws,r,6,eve,reg,fmt=MULT,align=right) if eve is not None else put(ws,r,6,"n/d",reg,align=center)
        put(ws,r,7,basis,reg); put(ws,r,8,srcref,reg,align=center); put(ws,r,9,notes,reg)
        if r%2==0:
            for c in range(1,10): ws.cell(row=r,column=c).fill=grey_fill
        r+=1
    first,last=5,r-1
    for stat,fn in (("Median","MEDIAN"),("Mean","AVERAGE")):
        put(ws,r,1,stat,bold,lblue_fill)
        for c in (2,3): put(ws,r,c,"",reg,lblue_fill)
        put(ws,r,4,f"={fn}(D{first}:D{last})",bold,lblue_fill,'#,##0.0',right)
        put(ws,r,5,f"={fn}(E{first}:E{last})",bold,lblue_fill,MULT,right)
        put(ws,r,6,f"={fn}(F{first}:F{last})",bold,lblue_fill,MULT,right)
        for c in (7,8,9): put(ws,r,c,"",reg,lblue_fill)
        r+=1
    put(ws,r,1,"n/d = not disclosed / not reliably sourced (excluded from median & mean).",sub_font,bd=False)
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=9)
    for i,w in enumerate([7,18,24,10,11,12,30,11,42],1): ws.column_dimensions[get_column_letter(i)].width=w
    return ws

nutrition=[
 (2007,"Danone","Numico",16.8,4.7,22.0,"EV/S derived (÷FY06 sales €2.6bn); EBITDA per press","S7","Baby + medical nutrition; peak strategic multiple"),
 (2007,"Nestlé","Gerber (Novartis)",5.5,2.8,15.7,"EV/S derived; EBITDA reported (15.7x FY07e)","S8","Made Nestlé #1 in baby food"),
 (2012,"Nestlé","Pfizer Nutrition",11.85,5.0,23.7,"EV/S derived (4.9x FY12e / 5.6x FY11); EBITDA derived ($0.5bn '11); press ~20x fwd","S9","EM (China) entry"),
 (2017,"Danone","WhiteWave",12.5,2.7,20.0,"EV/S reported (FY17); EBITDA reported (~19.8-20x FY16e)","S10","Organic/plant-based; lower-multiple bookend (not infant)"),
 (2017,"Reckitt Benckiser","Mead Johnson",17.9,4.4,17.4,"EV/S reported (FY15); EBITDA reported (FY16)","S11","Pure-play infant formula; closest single comp (~14x w/ synergies)"),
 (2021,"Nestlé","The Bountiful Company",5.75,2.9,None,"EV/S derived (÷~$2.0bn sales); EBITDA n/d","S12","Vitamins/supplements; adult-nutrition-adjacent"),
]
comps_sheet("Nutrition Comps",
 "Precedent transactions — infant + adult/medical nutrition. 'Mult. basis' flags Reported vs Derived vs estimate. Full citations on Sources tab.",nutrition)

epd=[
 (2014,"Sun Pharma","Ranbaxy",4.0,2.2,None,"EV/S derived (÷~$1.8bn sales); EBITDA n/d (low margin)","S13","India/EM generics consolidation"),
 (2015,"Teva","Allergan Generics (Actavis)",40.5,None,None,"EV reported; multiples n/d (analyst ~12-13x w/ synergies)","S14","Global generics scale; widely viewed as full price"),
 (2016,"Mylan","Meda",9.9,4.0,10.5,"EV/S derived (~4x FY15 sales); EBITDA reported (8.9x w/ syn; ~10-11x rep.)","S15","Branded/specialty + EM/OTC reach"),
 (2017,"Bain / Cinven","STADA",5.7,2.5,13.4,"EV/S derived (÷FY16 sales €2.14bn); EBITDA reported (FY16)","S16","PE take-private; branded generics + consumer health"),
 (2018,"Advent","Zentiva (Sanofi EU gen.)",2.1,2.5,10.3,"EV/S derived (approx); EBITDA reported (10.25x)","S17","EU/EM branded-generics carve-out"),
 (2025,"CapVest","STADA (majority stake)",10.0,2.4,11.3,"EV/S derived (÷FY24 sales ~€4.1bn); EBITDA reported (11.29x)","S18","Latest large branded-generics print; re-rate vs 2017"),
]
comps_sheet("EPD Comps",
 "Precedent transactions — branded / EM generics. Listed reads for context: Krka ~11.5x, Hikma ~6.3x, Richter ~6x EV/EBITDA [S6].",epd)

# ============================================================ VALUATION
ws=wb.create_sheet("Valuation"); ws.sheet_view.showGridLines=False
banner(ws,6,"Valuation — Implied Enterprise Value","Driven by Inputs & Assumptions tab. $ in US$m; EV midpoint shown in $bn.")
def block(ws,r,name,sal,eb,ms_lo,ms_hi,me_lo,me_hi,nec=None):
    put(ws,r,1,name,secn_font,bd=False); r+=1
    hrow(ws,r,1,["Method","Multiple low","Multiple high","EV low ($m)","EV high ($m)","EV midpoint ($bn)"]); r+=1
    put(ws,r,1,"EV / Sales (reported sales)",bold)
    put(ws,r,2,f"={INP}!{ms_lo}{MS}",reg,fmt=MULT,align=right); put(ws,r,3,f"={INP}!{ms_hi}{MS}",reg,fmt=MULT,align=right)
    put(ws,r,4,f"={sal}*B{r}",reg,fmt=USD0,align=right); put(ws,r,5,f"={sal}*C{r}",reg,fmt=USD0,align=right)
    put(ws,r,6,f"=AVERAGE(D{r}:E{r})/1000",reg,fmt='#,##0.0"bn"',align=right); sr=r; r+=1
    put(ws,r,1,"EV / EBITDA (standalone EBITDA)",bold)
    put(ws,r,2,f"={INP}!{me_lo}{ME}",reg,fmt=MULT,align=right); put(ws,r,3,f"={INP}!{me_hi}{ME}",reg,fmt=MULT,align=right)
    put(ws,r,4,f"={eb}*B{r}",reg,fmt=USD0,align=right); put(ws,r,5,f"={eb}*C{r}",reg,fmt=USD0,align=right)
    put(ws,r,6,f"=AVERAGE(D{r}:E{r})/1000",reg,fmt='#,##0.0"bn"',align=right); er=r; r+=1
    put(ws,r,1,"Blended indicative EV (pre-litigation)",bold,gold_fill)
    for c in (2,3): put(ws,r,c,"",reg,gold_fill)
    put(ws,r,4,f"=MIN(D{sr}:E{er})",bold,gold_fill,USD0,right); put(ws,r,5,f"=MAX(D{sr}:E{er})",bold,gold_fill,USD0,right)
    put(ws,r,6,f"=AVERAGE(D{r}:E{r})/1000",bold,gold_fill,'#,##0.0"bn"',right); blend=r; r+=1
    if nec:
        put(ws,r,1,"less: NEC litigation risk deduction",reg)
        for c in (2,3): put(ws,r,c,"",reg)
        put(ws,r,4,f"=-{nec}",reg,fmt=USD0,align=right); put(ws,r,5,f"=-{nec}",reg,fmt=USD0,align=right)
        put(ws,r,6,"",reg); r+=1
        put(ws,r,1,"Adjusted indicative EV",bold,green_fill)
        for c in (2,3): put(ws,r,c,"",reg,green_fill)
        put(ws,r,4,f"=D{blend}-{nec}",bold,green_fill,USD0,right); put(ws,r,5,f"=E{blend}-{nec}",bold,green_fill,USD0,right)
        put(ws,r,6,f"=AVERAGE(D{r}:E{r})/1000",bold,green_fill,'#,##0.0"bn"',right); final=r
    else:
        final=blend
    return r+2,final
r=4
r,nut_b=block(ws,r,"Nutrition",SAL_N,EB_N,"B","C","B","C",nec=NEC_N)
r,epd_b=block(ws,r,"Established Pharmaceuticals (EPD)",SAL_E,EB_E,"D","E","D","E")
put(ws,r,1,"Combined (Nutrition adjusted + EPD)",secn_font,bd=False); r+=1
hrow(ws,r,1,["","","","EV low ($m)","EV high ($m)","EV midpoint ($bn)"]); r+=1
put(ws,r,1,"Total indicative enterprise value",bold,green_fill)
for c in (2,3): put(ws,r,c,"",reg,green_fill)
put(ws,r,4,f"=D{nut_b}+D{epd_b}",bold,green_fill,USD0,right); put(ws,r,5,f"=E{nut_b}+E{epd_b}",bold,green_fill,USD0,right)
put(ws,r,6,f"=AVERAGE(D{r}:E{r})/1000",bold,green_fill,'#,##0.0"bn"',right)
for i,w in enumerate([34,14,14,15,15,18],1): ws.column_dimensions[get_column_letter(i)].width=w

# ============================================================ SOURCES & METHODOLOGY
ws=wb.create_sheet("Sources & Methodology"); ws.sheet_view.showGridLines=False
banner(ws,5,"Sources & Methodology","Every figure classified REPORTED (from a primary/press source), DERIVED (calculated), or ESTIMATE (analyst judgement). Click links to verify.")
put(ws,4,1,"A. Target financials & EBITDA estimate",secn_font,bd=False)
hrow(ws,5,1,["Ref","Item / figure","Value used","Type","Source"])
src=[
 ("S1","Abbott total net sales FY2024","$41,950m","REPORTED","Abbott FY2024 Form 10-K (FY ended 31-Dec-2024)","https://www.webull.com/news/12352164595459072"),
 ("S2","Nutrition ~$8.41bn; EPD ~$5.19bn; Dx ~$9.34bn FY24 segment sales","$8.41bn / $5.19bn","REPORTED","Abbott segment reporting (Bullfincher; Statista)","https://bullfincher.io/companies/abbott-laboratories/revenue-by-segment"),
 ("S3","Segment operating margin: Nutrition 17.9%, EPD 23.7%; growth +5.9% / +9.2%","17.9% / 23.7%","REPORTED","Abbott FY2024 10-K MD&A (segment results)","https://www.webull.com/news/12352164595459072"),
 ("S4","Group D&A FY24: Depreciation ~$1,340m + Amort. of intangibles ~$1,878m = ~$3,218m (~7.7% of sales). Amortization concentrated in Med Devices/Diagnostics (acquired intangibles), so Nutrition/EPD add-backs (depreciation-led) assumed below group: 3.75% / 2.75% of sales.","D&A add-back 3.75% / 2.75%","ESTIMATE","Abbott FY2024 cash-flow disclosure (per 10-Q/10-K); segment split is analyst judgement","https://www.sec.gov/Archives/edgar/data/0000001800/000162828024044602/abt-20240930.htm"),
 ("—","Implied segment operating earnings (Nutrition ~$1.50bn; EPD ~$1.23bn)","sales x margin","DERIVED","Calculation (S2 x S3)",""),
 ("—","Estimated segment EBITDA: Nutrition ~$1.7-1.9bn; EPD ~$1.3-1.45bn","op earnings + D&A","ESTIMATE","Calculation (operating earnings + S4 add-back)",""),
]
r=6
for ref,item,val,typ,source,url in src:
    put(ws,r,1,ref,bold,align=center); put(ws,r,2,item,reg); put(ws,r,3,val,reg,align=center)
    tf=PatternFill("solid",fgColor={"REPORTED":GREEN,"DERIVED":LBLUE,"ESTIMATE":GOLD}.get(typ,WHITE))
    put(ws,r,4,typ,bold,tf,align=center)
    if url: c=put(ws,r,5,source,link_font); c.hyperlink=url
    else: put(ws,r,5,source,reg)
    ws.row_dimensions[r].height=42; r+=1
r+=1
put(ws,r,1,"B. Precedent transaction sources (multiples)",secn_font,bd=False); r+=1
hrow(ws,r,1,["Ref","Transaction","Key reported figure","Type","Source"]); r+=1
src2=[
 ("S5","Comp discount rationale (Nutrition)","Peak comps discounted","ESTIMATE","Analyst judgement — see Nutrition Comps & RB/MJN coverage",""),
 ("S6","Listed branded-generics reads: Krka ~11.5x, Hikma ~6.3x, Richter ~6x; Zentiva 10.25x, STADA 11.29x","Context multiples","REPORTED","InterCapital — European pharma generics M&A","https://inter.capital/recent-mas-in-the-european-pharma-generics-industry/"),
 ("S7","Danone / Numico (2007)","€12.3bn EV; ~22x EBITDA; FY06 sales €2.6bn (18.4% EBITA margin)","REPORTED/DERIVED","Danone analyst presentation (Jul-2007); 22x cited in RB/MJN coverage","http://media.corporate-ir.net/media_files/irol/95/95168/presentations/company/2007/090707_Numico_Prez_Analystes_v16.pdf"),
 ("S8","Nestlé / Gerber (2007)","$5.5bn EV; 15.7x FY07e EBITDA; FY07e sales $1.95bn","REPORTED/DERIVED","Washington Post (Apr-2007); Nestlé/Novartis releases","http://www.washingtonpost.com/wp-dyn/content/article/2007/04/12/AR2007041200372.html"),
 ("S9","Nestlé / Pfizer Nutrition (2012)","$11.85bn EV; FY11 sales $2.1bn, EBITDA $0.5bn (=23.7x); press ~20x fwd","REPORTED/DERIVED","Proactive Investors; Pfizer 8-K (SEC)","https://www.proactiveinvestors.com/companies/news/85559/pfizer-sells-nutrition-business-to-nestl-for-1185-bln-28063.html"),
 ("S10","Danone / WhiteWave (2016/17)","$12.5bn EV; 2.7x FY17 sales; ~19.8-20x FY16e EBITDA","REPORTED","FoodNavigator-USA","https://www.foodnavigator-usa.com/Article/2016/07/07/Danone-to-acquire-WhiteWave-Foods-in-12.5bn-deal/"),
 ("S11","Reckitt / Mead Johnson (2017)","$17.9bn EV; 4.4x FY15 sales; 17.4x FY16 EBITDA (~14x w/ syn.)","REPORTED","Bocconi BSIC; Mead Johnson 8-K (SEC)","https://bsic.it/go-big-go-home-reckitt-benckiser-acquire-mead-johnson-17-9bn/"),
 ("S12","Nestlé / The Bountiful Company (2021)","$5.75bn EV; ~$2.0bn sales (=~2.9x); EBITDA n/d","REPORTED/DERIVED","Nestlé press (Aug-2021)",""),
 ("S13","Sun Pharma / Ranbaxy (2014)","~$4.0bn EV incl. debt; ~$1.8bn sales (=~2.2x); EBITDA n/d","REPORTED/DERIVED","Sun Pharma / Ranbaxy deal disclosures",""),
 ("S14","Teva / Allergan Generics (Actavis, 2015)","$40.5bn EV (closed ~$39.5bn); multiples n/d (analyst ~12-13x)","REPORTED","SEC Form 425 (Allergan); analyst commentary","https://www.sec.gov/Archives/edgar/data/0000850693/000119312515086653/d889402dex991.htm"),
 ("S15","Mylan / Meda (2016)","$9.9bn EV; 8.9x adj. EBITDA w/ synergies; FY15 sales ~$2.3bn (=~4.0x)","REPORTED/DERIVED","Mylan 8-K (SEC)","https://www.sec.gov/Archives/edgar/data/0001623613/000119312516766807/d263574dex991.htm"),
 ("S16","Bain/Cinven / STADA (2017)","€5.32bn EV; 13.4x FY16 EBITDA; FY16 sales €2.14bn (=~2.5x)","REPORTED/DERIVED","Bocconi BSIC","https://bsic.it/going-one-going-twice-sold-cinven-bain-capital-acquire-stada-e5-3bn-intense-bidding-war/"),
 ("S17","Advent / Zentiva (2018)","~€1.9bn EV; ~10.25x EBITDA; ~2.5x sales","REPORTED/DERIVED","InterCapital","https://inter.capital/recent-mas-in-the-european-pharma-generics-industry/"),
 ("S18","CapVest / STADA majority (2025)","~€10bn EV reported; ~11.29x EBITDA; FY24 sales ~€4.1bn (=~2.4x)","REPORTED/DERIVED","Bain Capital/Cinven press; InterCapital","https://www.businesswire.com/news/home/20250901671846/en/CapVest-to-Acquire-Majority-Stake-in-STADA-from-Bain-Capital-and-Cinven"),
 ("S19","NEC infant-formula litigation overhang","$495m (2024), $58-60m (2025-26) verdicts; ~683 active cases (Apr-25)","REPORTED","Reuters/Yahoo; Top Class Actions","https://finance.yahoo.com/sectors/healthcare/articles/abbott-laboratories-contest-damages-awarded-111141543.html"),
 ("S20","Standalone cost / dis-synergy assumption","1.5% (Nutrition) / 2.0% (EPD) of sales","ESTIMATE","Analyst judgement — typical carve-out dis-synergy 1-3% of revenue (corporate finance/IT/HR/legal/regulatory rebuilt; net of TSA). PE buyer bears; strategic may offset via synergies.",""),
 ("S21","NEC litigation risk deduction (Nutrition EV)","$1,500m (illustrative, editable)","ESTIMATE","Illustrative buyer indemnity/escrow ask — NOT a booked Abbott provision. Placeholder given unresolved litigation; Abbott has also won defense verdicts and is appealing [S19].",""),
 ("S22","Stranded costs at RemainCo (Abbott retained group)","Not deducted from target EV","ESTIMATE","Costs left behind that don't transfer with the unit affect Abbott's RemainCo, not the target's standalone EV; flagged for the seller's net-proceeds / dis-synergy analysis.",""),
]
for ref,item,val,typ,source,url in src2:
    put(ws,r,1,ref,bold,align=center); put(ws,r,2,item,reg); put(ws,r,3,val,reg)
    tf=PatternFill("solid",fgColor={"REPORTED":GREEN,"DERIVED":LBLUE,"ESTIMATE":GOLD,"REPORTED/DERIVED":LBLUE}.get(typ,WHITE))
    put(ws,r,4,typ,bold,tf,align=center)
    if url: c=put(ws,r,5,source,link_font); c.hyperlink=url
    else: put(ws,r,5,source,reg)
    ws.row_dimensions[r].height=40; r+=1
r+=1
put(ws,r,1,"Legend:  REPORTED = from a primary filing or press source  ·  DERIVED = calculated (EV ÷ sales/EBITDA)  ·  ESTIMATE = analyst judgement.  Multiples across sources may use different EBITDA definitions / years; treat as directional. Not investment advice.",sub_font,bd=False)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5); ws.row_dimensions[r].height=42
for i,w in enumerate([7,40,40,16,52],1): ws.column_dimensions[get_column_letter(i)].width=w

# ============================================================ SUMMARY (front)
ws=wb.create_sheet("Summary",0); ws.sheet_view.showGridLines=False
banner(ws,6,"Abbott — Precedent Transaction Analysis","Potential divestiture of Nutrition and Established Pharmaceuticals (EPD)  •  FY2024 basis  •  Indicative / discussion draft")
put(ws,4,1,"Headline indicative EV. Build = Valuation tab; assumptions = Inputs tab; every figure sourced/flagged on Sources & Methodology tab.",sub_font,bd=False)
hrow(ws,6,1,["Business","FY24 Sales ($bn)","Standalone EBITDA ($bn)","EV/Sales","EV/EBITDA","Indicative EV ($bn)"])
VAL="Valuation"
def srow(r,name,sal,eb,evs_lo,evs_hi,eve_lo,eve_hi,lo,hi,fill=None):
    put(ws,r,1,name,bold,fill)
    put(ws,r,2,f"={sal}/1000",reg,fill,'#,##0.0',right); put(ws,r,3,f"={eb}/1000",reg,fill,'#,##0.0',right)
    put(ws,r,4,f'=TEXT({INP}!{evs_lo},"0.0")&"-"&TEXT({INP}!{evs_hi},"0.0")&"x"',reg,fill,align=center)
    put(ws,r,5,f'=TEXT({INP}!{eve_lo},"0.0")&"-"&TEXT({INP}!{eve_hi},"0.0")&"x"',reg,fill,align=center)
    put(ws,r,6,f'=TEXT({VAL}!{lo}/1000,"0")&"-"&TEXT({VAL}!{hi}/1000,"0")&"bn"',bold,fill,align=center)
srow(7,"Nutrition",SAL_N,EB_N,f"B{MS}",f"C{MS}",f"B{MS+1}",f"C{MS+1}",f"D{nut_b}",f"E{nut_b}")
srow(8,"Established Pharmaceuticals (EPD)",SAL_E,EB_E,f"D{MS}",f"E{MS}",f"D{MS+1}",f"E{MS+1}",f"D{epd_b}",f"E{epd_b}",grey_fill)
put(ws,9,1,"Combined",Font(name="Calibri",size=11,bold=True,color=NAVY),green_fill)
put(ws,9,2,f"=({SAL_N}+{SAL_E})/1000",bold,green_fill,'#,##0.0',right); put(ws,9,3,f"=({EB_N}+{EB_E})/1000",bold,green_fill,'#,##0.0',right)
put(ws,9,4,"—",reg,green_fill,align=center); put(ws,9,5,"—",reg,green_fill,align=center)
put(ws,9,6,f'=TEXT(({VAL}!D{nut_b}+{VAL}!D{epd_b})/1000,"0")&"-"&TEXT(({VAL}!E{nut_b}+{VAL}!E{epd_b})/1000,"0")&"bn"',bold,green_fill,align=center)
put(ws,11,1,"Key points",secn_font,bd=False)
pts=[
 "EPD: high-growth (+9.2%), high-margin (23.7%) EM branded generics — premium end of the branded-generics deal range.",
 "Nutrition: two-speed — growing adult/medical nutrition (Ensure, Glucerna) vs. structurally challenged infant formula (Similac) with an NEC litigation overhang [S19].",
 "Peak infant-formula comps (Reckitt/MJN 17.4x; Nestlé/Pfizer ~20-24x) discounted for Nutrition's blend and litigation risk.",
 "Sum-of-the-parts optionality: adult nutrition at consumer-health multiples while ring-fencing infant formula could exceed the blended range.",
 "Carve-out adjustments applied: standalone/dis-synergy cost 1.5% (Nutrition) / 2.0% (EPD) of sales reduces EBITDA; an illustrative $1.5bn NEC risk deduction is netted off Nutrition EV. Both are editable on the Inputs tab [S20, S21]. Stranded costs sit at Abbott RemainCo, not the target [S22].",
 "EBITDA is ESTIMATED (Abbott discloses segment margin, not EBITDA). Every figure's basis is on the Sources & Methodology tab. See Sensitivity tab for the swing factors.",
]
rr=12
for p in pts:
    put(ws,rr,1,"•  "+p,reg,bd=False); ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=6)
    ws.row_dimensions[rr].height=30; rr+=1
put(ws,rr+1,1,"Indicative EV is post standalone/dis-synergy costs and (for Nutrition) post the illustrative NEC deduction.",sub_font,bd=False)
put(ws,rr+2,1,"Tabs:  Summary · Inputs & Assumptions · Nutrition Comps · EPD Comps · Valuation · Sensitivity · Sources & Methodology",sub_font,bd=False)
for i,w in enumerate([40,16,18,14,14,20],1): ws.column_dimensions[get_column_letter(i)].width=w

# ============================================================ SENSITIVITY
ws=wb.create_sheet("Sensitivity")  # placed after Valuation below via move
ws.sheet_view.showGridLines=False
banner(ws,7,"Sensitivity Analysis","Live grids driven by Inputs tab. Shows how indicative EV swings with the key carve-out estimates. $bn.")
N_SAL=f"{INP}!B8"; N_OPM=f"{INP}!B9"; N_DA=f"{INP}!B10"
E_SAL=f"{INP}!C8"; E_OPM=f"{INP}!C9"; E_DA=f"{INP}!C10"
N_MEmid=f"((({INP}!B{ME})+({INP}!C{ME}))/2)"   # Nutrition EV/EBITDA midpoint
E_MEmid=f"((({INP}!D{ME})+({INP}!E{ME}))/2)"   # EPD EV/EBITDA midpoint
NECc=f"{INP}!B{nec_row}"

# --- Table A: Nutrition adjusted EV ($bn): EV/EBITDA multiple (rows) x standalone cost % (cols), net of NEC
put(ws,4,1,"A.  Nutrition adjusted EV ($bn)  —  EV/EBITDA multiple  ×  standalone cost % of sales  (net of NEC deduction on Inputs tab)",secn_font,bd=False)
costs=[0.010,0.015,0.020,0.025]
mults=[10,11,12,13,14,15]
put(ws,5,1,"EV/EBITDA \\ Std cost %",bold,hdr_fill); ws.cell(5,1).font=hdr_font; ws.cell(5,1).alignment=center; ws.cell(5,1).border=border
for j,cst in enumerate(costs):
    put(ws,5,2+j,cst,hdr_font,hdr_fill,fmt=PCT,align=center)
for i,m in enumerate(mults):
    rr=6+i
    put(ws,rr,1,m,bold,lblue_fill,fmt=MULT,align=center)
    for j,cst in enumerate(costs):
        # (preEBITDA - cost*sales)*mult - NEC, all /1000 -> $bn
        f=f"=(({N_SAL}*({N_OPM}+{N_DA}))-{cst}*{N_SAL})*{m}/1000-{NECc}/1000"
        put(ws,rr,2+j,f,reg,fmt='#,##0.0',align=center)

# --- Table B: Combined adjusted EV ($bn): NEC deduction (rows) x standalone cost % both (cols), at selected EV/EBITDA midpoints
base=13
put(ws,base,1,"B.  Combined adjusted EV ($bn)  —  NEC deduction $bn (rows)  ×  standalone cost % applied to both (cols),  at selected EV/EBITDA midpoints",secn_font,bd=False)
necs=[0,500,1000,1500,2000,2500]
put(ws,base+1,1,"NEC $m \\ Std cost %",hdr_font,hdr_fill,align=center)
for j,cst in enumerate(costs):
    put(ws,base+1,2+j,cst,hdr_font,hdr_fill,fmt=PCT,align=center)
for i,nv in enumerate(necs):
    rr=base+2+i
    put(ws,rr,1,nv,bold,lblue_fill,fmt=USD0,align=center)
    for j,cst in enumerate(costs):
        nut=f"(({N_SAL}*({N_OPM}+{N_DA}))-{cst}*{N_SAL})*{N_MEmid}-{nv}"
        epd=f"(({E_SAL}*({E_OPM}+{E_DA}))-{cst}*{E_SAL})*{E_MEmid}"
        put(ws,rr,2+j,f"=({nut}+{epd})/1000",reg,fmt='#,##0.0',align=center)
put(ws,base+2+len(necs)+1,1,"Tables use the EV/EBITDA method (the method affected by standalone costs); EV/Sales is unaffected. Shaded cells nearest the base case (1.5%/2.0% cost, $1.5bn NEC) are the headline. Estimates — not investment advice.",sub_font,bd=False)
ws.merge_cells(start_row=base+2+len(necs)+1,start_column=1,end_row=base+2+len(necs)+1,end_column=7)
ws.row_dimensions[base+2+len(necs)+1].height=42
ws.column_dimensions['A'].width=22
for col in "BCDEFG": ws.column_dimensions[col].width=12

wb.move_sheet("Sensitivity", offset=-1)  # place before Sources & Methodology
out="/home/user/knutnyman/abbott-precedent-transaction-analysis/Abbott_Nutrition_and_EPD_Precedent_Transaction_Analysis.xlsx"
wb.save(out)
print("saved",out); print("sheets:",wb.sheetnames)
