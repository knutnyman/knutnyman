#!/usr/bin/env python3
"""Generate the CGM LCD/NCD scenario analysis PowerPoint deck."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY = RGBColor(0x1F, 0x33, 0x5E)
ACCENT = RGBColor(0x2E, 0x74, 0xB5)
LIGHT = RGBColor(0xEE, 0xF2, 0xF8)
GREY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
AMBER = RGBColor(0xB8, 0x86, 0x0B)
LGREEN = RGBColor(0xE3, 0xF1, 0xE4)
LRED = RGBColor(0xFB, 0xE7, 0xE5)
LAMBER = RGBColor(0xFB, 0xF3, 0xDD)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def box(s, l, t, w, h, fill=None, line=None, line_w=None):
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w or 1)
    return shp

def text(s, l, t, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.0, wrap=True):
    tb = s.shapes.add_textbox(l, t, w, h); tf = tb.text_frame
    tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2); tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    first = True
    for item in runs:
        # item: (text, size, color, bold, italic, bullet?, align?)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = item.get('align', align)
        p.space_after = Pt(item.get('space_after', space_after))
        p.space_before = Pt(item.get('space_before', 0))
        p.line_spacing = item.get('line_spacing', line_spacing)
        segs = item['segs'] if 'segs' in item else [(item['t'], item.get('size',14),
                 item.get('color',NAVY), item.get('bold',False), item.get('italic',False))]
        if item.get('bullet'):
            r0 = p.add_run(); r0.text = '▪  '
            r0.font.size = Pt(item.get('size',14)); r0.font.color.rgb = item.get('bcolor', ACCENT)
            r0.font.bold = True
        for seg in segs:
            r = p.add_run(); r.text = seg[0]
            r.font.size = Pt(seg[1]); r.font.color.rgb = seg[2]
            r.font.bold = seg[3]; r.font.italic = seg[4]
            r.font.name = 'Calibri'
    return tb

def header(s, title, kicker=None, num=None):
    box(s, 0, 0, SW, Inches(1.0), fill=NAVY)
    box(s, 0, Inches(1.0), SW, Pt(3), fill=ACCENT)
    text(s, Inches(0.45), Inches(0.12), Inches(11.8), Inches(0.8),
         [{'t': title, 'size': 26, 'color': WHITE, 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}],
         anchor=MSO_ANCHOR.MIDDLE)
    if kicker:
        text(s, Inches(0.47), Inches(0.66), Inches(11.8), Inches(0.3),
             [{'t': kicker, 'size': 11.5, 'color': RGBColor(0xBF,0xD3,0xEC), 'italic': True}])
    if num:
        text(s, Inches(12.5), Inches(0.12), Inches(0.7), Inches(0.8),
             [{'t': num, 'size': 12, 'color': RGBColor(0xBF,0xD3,0xEC), 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}],
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def footer(s, n):
    text(s, Inches(0.45), Inches(7.12), Inches(8), Inches(0.3),
         [{'t': 'CGM Medicare LCD / NCD Scenario Analysis  |  Confidential — Draft for Discussion', 'size': 8, 'color': GREY}])
    text(s, Inches(12.4), Inches(7.12), Inches(0.7), Inches(0.3),
         [{'t': str(n), 'size': 9, 'color': GREY}], align=PP_ALIGN.RIGHT)

def chip(s, l, t, w, label, color, fill):
    b = box(s, l, t, w, Inches(0.34), fill=fill, line=color, line_w=1)
    tf = b.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = color
    return b

slide_no = 0

# ---------- SLIDE 1: TITLE ----------
s = slide()
box(s, 0, 0, SW, SH, fill=NAVY)
box(s, 0, Inches(4.55), SW, Pt(3.5), fill=ACCENT)
text(s, Inches(0.7), Inches(1.5), Inches(12), Inches(1.2),
     [{'t': 'Continuous Glucose Monitors', 'size': 46, 'color': WHITE, 'bold': True}])
text(s, Inches(0.7), Inches(2.6), Inches(12), Inches(1.0),
     [{'t': 'Scenario Analysis of a New Medicare LCD / NCD', 'size': 30, 'color': RGBColor(0x9F,0xC2,0xE8)}])
text(s, Inches(0.72), Inches(3.6), Inches(12), Inches(0.7),
     [{'t': 'Probability-Weighted Outcomes  ·  Precedent Base Rates  ·  Investment Implications',
       'size': 16, 'color': RGBColor(0xBF,0xD3,0xEC), 'italic': True}])
text(s, Inches(0.72), Inches(4.85), Inches(12), Inches(1.6),
     [{'t': 'Prepared for: Investment Diligence — Internal', 'size': 13, 'color': WHITE, 'space_after': 6},
      {'t': 'Date: June 10, 2026', 'size': 13, 'color': WHITE, 'space_after': 6},
      {'t': 'Horizon: Coverage & payment events through YE2028', 'size': 13, 'color': WHITE, 'space_after': 6},
      {'segs': [('Bottom line: ', 13, RGBColor(0x9F,0xC2,0xE8), True, False),
                ('Coverage is expanding while payment is being cut — two opposing tracks, often conflated.',
                 13, WHITE, False, True)]}])
slide_no += 1

# ---------- SLIDE 2: THE BIG IDEA — TWO TRACKS ----------
s = slide(); slide_no += 1
header(s, 'The Central Idea: Two Opposing Tracks', 'The most common analytical error is conflating coverage with payment', 'Thesis')
# Track A box
box(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(4.9), fill=LGREEN, line=GREEN, line_w=2)
text(s, Inches(0.7), Inches(1.55), Inches(5.6), Inches(0.6),
     [{'t': 'TRACK A — ELIGIBILITY  ▲', 'size': 19, 'color': GREEN, 'bold': True}])
text(s, Inches(0.7), Inches(2.15), Inches(5.6), Inches(4.0),
     [{'segs':[('Direction:  ',13,NAVY,True,False),('Expansionary',13,GREEN,True,False)],'space_after':9,'bullet':True,'bcolor':GREEN},
      {'segs':[('Instrument:  ',13,NAVY,True,False),('DME MAC LCD L33822 revision',13,GREY,False,False)],'space_after':9,'bullet':True,'bcolor':GREEN},
      {'segs':[('Status:  ',13,NAVY,True,False),('Anticipated, not yet filed',13,AMBER,True,False)],'space_after':9,'bullet':True,'bcolor':GREEN},
      {'segs':[('Key event:  ',13,NAVY,True,False),('Draft LCD for non-insulin T2D, exp. H2 2026',13,GREY,False,False)],'space_after':9,'bullet':True,'bcolor':GREEN,'line_spacing':1.0},
      {'segs':[('Effect:  ',13,NAVY,True,False),('Adds up to ~12M eligible lives',13,GREEN,True,False)],'space_after':9,'bullet':True,'bcolor':GREEN}])
# Track B box
box(s, Inches(6.85), Inches(1.4), Inches(6.0), Inches(4.9), fill=LRED, line=RED, line_w=2)
text(s, Inches(7.05), Inches(1.55), Inches(5.6), Inches(0.6),
     [{'t': 'TRACK B — PAYMENT  ▼', 'size': 19, 'color': RED, 'bold': True}])
text(s, Inches(7.05), Inches(2.15), Inches(5.6), Inches(4.0),
     [{'segs':[('Direction:  ',13,NAVY,True,False),('Restrictive',13,RED,True,False)],'space_after':9,'bullet':True,'bcolor':RED},
      {'segs':[('Instrument:  ',13,NAVY,True,False),('Competitive bidding + monthly rental',13,GREY,False,False)],'space_after':9,'bullet':True,'bcolor':RED},
      {'segs':[('Status:  ',13,NAVY,True,False),('FINALIZED Nov 28, 2025 — now law',13,RED,True,False)],'space_after':9,'bullet':True,'bcolor':RED},
      {'segs':[('Key event:  ',13,NAVY,True,False),('Contracts effective ≤ Jan 1, 2028',13,GREY,False,False)],'space_after':9,'bullet':True,'bcolor':RED},
      {'segs':[('Effect:  ',13,NAVY,True,False),('Cuts unit reimbursement ~15–40%',13,RED,True,False)],'space_after':9,'bullet':True,'bcolor':RED}])
text(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.7),
     [{'t': 'Model Medicare CGM revenue = VOLUME (coverage scenario) × PRICE (almost certainly lower from 2028). No scenario leaves 2028 unit price above today.',
       'size': 13, 'color': NAVY, 'bold': True, 'align': PP_ALIGN.CENTER}], align=PP_ALIGN.CENTER)
footer(s, slide_no)

# ---------- SLIDE 3: SCENARIO SUMMARY ----------
s = slide(); slide_no += 1
header(s, 'Probability-Weighted Scenario Set', 'Coverage-instrument outcome over a ~24–36 month horizon', 'Summary')
rows = [
    ('S1', 'Conditioned LCD expansion to non-insulin T2D (A1c / hypo-med / structured)', '35%', 'MODAL', GREEN, LGREEN),
    ('S2', 'Broad LCD expansion to all non-insulin T2D', '20%', 'Bull', GREEN, LGREEN),
    ('S3', 'Status quo persists (expansion slips past horizon)', '25%', 'Neutral*', AMBER, LAMBER),
    ('S4', 'Restriction-led: bidding bites + documentation / prior-auth', '15%', 'Bear', RED, LRED),
    ('S5', 'NCD / CED tail (national action; restrictive trapdoor)', '5%', 'Binary', RED, LRED),
]
y = Inches(1.45)
# header row
box(s, Inches(0.5), y, Inches(12.33), Inches(0.45), fill=NAVY)
for x, w, lab in [(0.5,0.7,'#'), (1.2,7.3,'Scenario'), (8.5,1.3,'Prob.'), (9.8,3.03,'Net effect')]:
    text(s, Inches(x), y, Inches(w), Inches(0.45),
         [{'t': lab, 'size': 12, 'color': WHITE, 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
y = Emu(y + Inches(0.45))
for tag, name, prob, eff, color, fill in rows:
    rh = Inches(0.78)
    box(s, Inches(0.5), y, Inches(12.33), rh, fill=fill, line=WHITE, line_w=1)
    text(s, Inches(0.5), y, Inches(0.7), rh, [{'t': tag, 'size': 15, 'color': NAVY, 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.25), y, Inches(7.2), rh, [{'t': name, 'size': 12.5, 'color': NAVY, 'anchor': MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(8.5), y, Inches(1.3), rh, [{'t': prob, 'size': 18, 'color': color, 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(9.8), y, Inches(3.0), rh, [{'t': eff, 'size': 13, 'color': color, 'bold': True, 'anchor': MSO_ANCHOR.MIDDLE}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    y = Emu(y + rh + Pt(2))
text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5),
     [{'segs':[('Some expansion (S1+S2) ≈ 55%   ·   Do-no-harm-to-volume (S1+S2+S3) ≈ 80%   ·   Materially adverse (S4 + restrictive S5) ≈ 18%.   ',12.5,NAVY,True,False),
               ('*S3 still neutral-to-negative because the price cut lands regardless.',12.5,GREY,False,True)]}])
footer(s, slide_no)

# ---------- SLIDE 4: HISTORY TIMELINE ----------
s = slide(); slide_no += 1
header(s, 'How We Got Here', 'Coverage has expanded via the lightest instruments; payment only now turned restrictive', 'Context')
events = [
    ('2017', 'Ruling 1682-R: therapeutic CGM = DME', GREEN),
    ('2021', 'Fingerstick 4×/day prerequisite removed', GREEN),
    ('2023', 'Expansion: any insulin / problematic hypo', GREEN),
    ('Nov 2025', 'FINAL: bidding + monthly rental', RED),
    ('H2 2026', 'Expected draft LCD: non-insulin T2D', AMBER),
    ('Jan 2028', 'Bidding contracts effective', RED),
]
n = len(events)
cstart = Inches(1.25); cend = Inches(12.08); cy = Inches(3.7)
bw = Inches(1.86); halfbw = Inches(0.93)
box(s, cstart, cy, Emu(cend - cstart), Pt(4), fill=ACCENT)
step = (cend - cstart) / (n - 1)
for i, (yr, lab, color) in enumerate(events):
    cx = Emu(cstart + step * i)
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Emu(cx - Inches(0.16)), Emu(cy - Inches(0.13)), Inches(0.32), Inches(0.32))
    dot.fill.solid(); dot.fill.fore_color.rgb = color; dot.line.color.rgb = WHITE; dot.line.width = Pt(2)
    dot.shadow.inherit = False
    above = (i % 2 == 0)
    ty = Inches(1.7) if above else Inches(4.25)
    bh = Inches(1.7)
    bx = Emu(cx - halfbw)
    box(s, bx, ty, bw, bh, fill=LIGHT, line=color, line_w=1.5)
    text(s, bx, Emu(ty+Inches(0.08)), bw, Inches(0.4),
         [{'t': yr, 'size': 15, 'color': color, 'bold': True, 'align': PP_ALIGN.CENTER}], align=PP_ALIGN.CENTER)
    text(s, Emu(bx+Inches(0.08)), Emu(ty+Inches(0.5)), Emu(bw-Inches(0.16)), Inches(1.1),
         [{'t': lab, 'size': 11, 'color': NAVY, 'align': PP_ALIGN.CENTER, 'line_spacing':1.0}], align=PP_ALIGN.CENTER)
text(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.6),
     [{'segs':[('Pattern:  ',13,NAVY,True,False),
               ('every coverage change rode a sub-NCD instrument (ruling, payment rule, LCD). The next move is overwhelmingly likely to be an LCD, not an NCD.',13,GREY,False,True)]}])
footer(s, slide_no)

# ---------- SLIDE 5: MARKET AT STAKE ----------
s = slide(); slide_no += 1
header(s, 'What Is at Stake', 'Medicare CGM spend grew ~12× in five years; the non-insulin TAM is the prize', 'Market')
# Left: spend stats
box(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(2.6), fill=LIGHT, line=ACCENT, line_w=1.5)
text(s, Inches(0.7), Inches(1.5), Inches(5.6), Inches(0.4), [{'t':'MEDICARE CGM SPEND','size':14,'color':ACCENT,'bold':True}])
text(s, Inches(0.7), Inches(1.95), Inches(5.6), Inches(2.0),
     [{'segs':[('$109M ',24,GREY,True,False),('(2018)   →   ',16,GREY,False,False),('$1.3B ',24,NAVY,True,False),('(2023)',16,GREY,False,False)],'space_after':10},
      {'segs':[('OIG (Nov 2025): paid ',12,NAVY,False,False),('69% above supplier cost',12,RED,True,False),(' — $377M/yr',12,NAVY,False,False)],'space_after':7,'bullet':True},
      {'segs':[('25.2% improper-payment rate',12,RED,True,False),(' on glucose monitors (2024)',12,NAVY,False,False)],'space_after':7,'bullet':True},
      {'segs':[('A4239 supply allowance ',12,NAVY,False,False),('~$268/mo',12,NAVY,True,False),(' vs. OTC ~$89–99',12,GREY,False,False)],'bullet':True}])
# Left bottom: adoption
box(s, Inches(0.5), Inches(4.2), Inches(6.0), Inches(2.1), fill=LGREEN, line=GREEN, line_w=1.5)
text(s, Inches(0.7), Inches(4.3), Inches(5.6), Inches(0.4), [{'t':'ADOPTION INFLECTED POST-2023','size':14,'color':GREEN,'bold':True}])
text(s, Inches(0.7), Inches(4.75), Inches(5.6), Inches(1.5),
     [{'segs':[('CGM use, MA Type 2 insulin users:  ',12,NAVY,False,False),('1.4% → 17.2%',14,GREEN,True,False)],'space_after':8,'bullet':True},
      {'segs':[('Jan 2021 → 2023, accelerating with the April 2023 LCD expansion',11.5,GREY,False,True)],'space_after':8},
      {'segs':[('GLP-1s a modest tailwind, not a substitute',12,NAVY,False,False)],'bullet':True}])
# Right: TAM ladder
box(s, Inches(6.85), Inches(1.4), Inches(6.0), Inches(4.9), fill=WHITE, line=NAVY, line_w=1.5)
text(s, Inches(7.05), Inches(1.5), Inches(5.6), Inches(0.4), [{'t':'THE TAM LADDER','size':14,'color':NAVY,'bold':True}])
ladder = [
    ('Type 1 + intensive insulin', 'Covered 2017', GREEN),
    ('Basal-only insulin (T2)', 'Covered 2023', GREEN),
    ('Non-insulin T2 w/ problematic hypo', 'Covered 2023', GREEN),
    ('Non-insulin T2 (no hypo) — ~12M', 'THE TARGET', AMBER),
    ('Prediabetes / wellness — ~31M', 'OTC only', RED),
]
yy = Inches(2.0)
for name, status, color in ladder:
    box(s, Inches(7.05), yy, Inches(5.6), Inches(0.72), fill=LIGHT, line=color, line_w=1)
    text(s, Inches(7.2), yy, Inches(3.9), Inches(0.72), [{'t':name,'size':12.5,'color':NAVY,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(11.0), yy, Inches(1.55), Inches(0.72), [{'t':status,'size':11,'color':color,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy = Emu(yy + Inches(0.83))
footer(s, slide_no)

# ---------- SLIDE 6: CLINICAL EVIDENCE ----------
s = slide(); slide_no += 1
header(s, 'The Evidence That Drives the Decision', 'Borderline-and-strengthening for structured populations; not a slam-dunk for blanket coverage', 'Clinical')
# table
cols = [('Population',3.5),('Key evidence',4.0),('Effect',2.8),('Strength',1.8)]
data = [
    ('Intensive insulin', 'DIAMOND, GOLD, HypoDE', 'Large benefit', 'Settled (A)', GREEN),
    ('Older adults (T1D)', 'WISDM (n=203, ≥60y)', '−1.9% time <70', 'Strong', GREEN),
    ('Basal insulin T2D', 'MOBILE (JAMA 2021)', '−0.4% A1c', 'Adequate', GREEN),
    ('Non-insulin T2D', 'CONNECT (Jun 2026); IMMEDIATE', 'CONNECT −0.9%; pooled ~−0.3%', 'Contested', AMBER),
    ('Prediabetes / wellness', 'No supportive RCT', 'Null / minimal', 'Weak', RED),
]
x = Inches(0.5); y = Inches(1.45)
box(s, x, y, Inches(12.1), Inches(0.42), fill=NAVY)
cx = x
for lab, w in cols:
    text(s, cx, y, Inches(w), Inches(0.42), [{'t':lab,'size':12,'color':WHITE,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
    cx = Emu(cx + Inches(w))
y = Emu(y + Inches(0.42))
for row in data:
    name, ev, eff, strg, color = row
    rh = Inches(0.62)
    box(s, x, y, Inches(12.1), rh, fill=LIGHT, line=WHITE, line_w=1)
    vals = [(name,3.5,NAVY,False),(ev,4.0,GREY,False),(eff,2.8,NAVY,False),(strg,1.8,color,True)]
    cx = x
    for v, w, c, b in vals:
        al = PP_ALIGN.CENTER if w==1.8 else PP_ALIGN.LEFT
        text(s, Emu(cx+Inches(0.08)), y, Inches(w), rh, [{'t':v,'size':11.5,'color':c,'bold':b,'anchor':MSO_ANCHOR.MIDDLE}], align=al, anchor=MSO_ANCHOR.MIDDLE)
        cx = Emu(cx + Inches(w))
    y = Emu(y + rh + Pt(1.5))
# CONNECT callout + skeptic
box(s, Inches(0.5), Inches(5.45), Inches(6.0), Inches(1.55), fill=LGREEN, line=GREEN, line_w=1.5)
text(s, Inches(0.65), Inches(5.52), Inches(5.7), Inches(1.45),
     [{'t':'PIVOTAL: CONNECT (Jun 2026)','size':12.5,'color':GREEN,'bold':True,'space_after':4},
      {'segs':[('n=283 non-insulin T2D, 22 sites: ',11,NAVY,False,False),('−0.9% A1c',11,GREEN,True,False),(' vs. routine care. "First Level A evidence." But open-label, industry-funded, 26 wks, A1c not hard outcomes.',11,NAVY,False,False)]}])
box(s, Inches(6.85), Inches(5.45), Inches(5.75), Inches(1.55), fill=LRED, line=RED, line_w=1.5)
text(s, Inches(7.0), Inches(5.52), Inches(5.5), Inches(1.45),
     [{'t':'SKEPTIC CASE','size':12.5,'color':RED,'bold':True,'space_after':4},
      {'segs':[('Pooled effects shrink to ~−0.28% (Richardson 2024); durability unproven; hard-outcome data uncontrolled; NICE & USPSTF decline non-insulin. ',11,NAVY,False,False),('→ favors a conditioned expansion.',11,RED,True,False)]}])
footer(s, slide_no)

# ---------- SLIDE 7: PRECEDENTS ----------
s = slide(); slide_no += 1
header(s, 'Precedent Base Rates: Bull vs. Bear', 'Analogous Medicare decisions bracket the plausible outcomes', 'Precedent')
# Bull column
box(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(2.45), fill=LGREEN, line=GREEN, line_w=2)
text(s, Inches(0.68), Inches(1.5), Inches(5.6), Inches(0.45), [{'t':'BULL — CPAP analog (2008)','size':15,'color':GREEN,'bold':True}])
text(s, Inches(0.68), Inches(1.98), Inches(5.6), Inches(1.8),
     [{'segs':[('Lowered diagnostic barrier (home sleep test) → ',11.5,NAVY,False,False),('decade-plus volume boom',11.5,GREEN,True,False),('.',11.5,NAVY,False,False)],'space_after':6,'bullet':True,'bcolor':GREEN},
      {'segs':[('The 2023 CGM LCD already ran this play (dropped fingerstick). Non-insulin would extend it.',11.5,NAVY,False,False)],'space_after':6,'bullet':True,'bcolor':GREEN},
      {'segs':[('Catch: ',11.5,RED,True,False),('a resupply boom invites the OIG/audit scrutiny CGM now faces.',11.5,NAVY,False,True)],'bullet':True,'bcolor':GREEN}])
# Bear column
box(s, Inches(6.85), Inches(1.4), Inches(6.0), Inches(2.45), fill=LRED, line=RED, line_w=2)
text(s, Inches(7.03), Inches(1.5), Inches(5.6), Inches(0.45), [{'t':'BEAR — Power mobility analog','size':15,'color':RED,'bold':True}])
text(s, Inches(7.03), Inches(1.98), Inches(5.6), Inches(1.8),
     [{'segs':[('Documentation + prior auth + competitive bidding → ',11.5,NAVY,False,False),('spend ~$964M→$190M (~80%)',11.5,RED,True,False)],'space_after':6,'bullet':True,'bcolor':RED},
      {'segs':[('No formal "non-coverage" decision needed — attrition does the work.',11.5,NAVY,False,False)],'space_after':6,'bullet':True,'bcolor':RED},
      {'segs':[('CGM is already 3 of 4 steps in: growing spend, fraud takedowns, bidding finalized.',11.5,NAVY,False,True)],'bullet':True,'bcolor':RED}])
# Other precedents strip
text(s, Inches(0.5), Inches(4.05), Inches(12), Inches(0.35), [{'t':'OTHER INSTRUCTIVE PRECEDENTS','size':13,'color':NAVY,'bold':True}])
mini = [
    ('Test strips (CB)', 'Price −79%', RED),
    ('TENS / CLBP (CED)', 'CED = de facto non-coverage', RED),
    ('Seat elevation NCD', 'Advocacy widened final', GREEN),
    ('Insulin pump NCD', 'Restrictive criteria stick 20+ yrs', AMBER),
    ('Eversense I-CGM', 'Coverage ≠ revenue', AMBER),
]
mw = Inches(2.36); mx = Inches(0.5)
for name, note, color in mini:
    box(s, mx, Inches(4.45), mw, Inches(1.25), fill=LIGHT, line=color, line_w=1.5)
    text(s, Emu(mx+Inches(0.1)), Inches(4.55), Emu(mw-Inches(0.2)), Inches(0.5), [{'t':name,'size':11.5,'color':NAVY,'bold':True,'align':PP_ALIGN.CENTER,'line_spacing':0.95}], align=PP_ALIGN.CENTER)
    text(s, Emu(mx+Inches(0.1)), Inches(5.15), Emu(mw-Inches(0.2)), Inches(0.5), [{'t':note,'size':10.5,'color':color,'bold':True,'align':PP_ALIGN.CENTER,'line_spacing':0.95}], align=PP_ALIGN.CENTER)
    mx = Emu(mx + mw + Inches(0.08))
text(s, Inches(0.5), Inches(5.95), Inches(12.2), Inches(1.0),
     [{'segs':[('Process reality:  ',12.5,NAVY,True,False),
               ('NCDs are rare (~4/yr) and slowing (CMS coverage staff −22% since Dec 2024). LCDs finalize within 365 days and flex both ways. Comment volume & advocacy demonstrably move CMS.',12.5,GREY,False,False)]}])
footer(s, slide_no)

# ---------- SLIDE 8: SCENARIO DETAIL S1/S2 ----------
def scen_detail_slide(items, title, kicker, tag_label):
    s = slide()
    header(s, title, kicker, tag_label)
    colw = Inches(6.0)
    xs = [Inches(0.5), Inches(6.85)]
    for (tag, name, prob, color, fill, mech, tline, impact), x in zip(items, xs):
        box(s, x, Inches(1.4), colw, Inches(5.55), fill=fill, line=color, line_w=2)
        text(s, Emu(x+Inches(0.18)), Inches(1.5), Emu(colw-Inches(0.36)), Inches(0.5),
             [{'segs':[(f'{tag}  ',20,color,True,False),(f'—  {prob}',20,NAVY,True,False)]}])
        text(s, Emu(x+Inches(0.18)), Inches(2.0), Emu(colw-Inches(0.36)), Inches(0.6),
             [{'t':name,'size':14,'color':NAVY,'bold':True,'line_spacing':1.0}])
        text(s, Emu(x+Inches(0.18)), Inches(2.75), Emu(colw-Inches(0.36)), Inches(4.1),
             [{'segs':[('Mechanism:  ',11.5,color,True,False),(mech,11.5,NAVY,False,False)],'space_after':9,'line_spacing':1.0},
              {'segs':[('Timeline:  ',11.5,color,True,False),(tline,11.5,NAVY,False,False)],'space_after':9,'line_spacing':1.0},
              {'segs':[('Impact:  ',11.5,color,True,False),(impact,11.5,NAVY,False,False)],'line_spacing':1.0}])
    footer(s, slide_no)
    return s

slide_no += 1
scen_detail_slide(
    [('S1','Conditioned LCD expansion to non-insulin T2D','35%  MODAL',GREEN,LGREEN,
      'DME MACs add non-insulin T2D but gated by an A1c threshold, hypoglycemia-risk meds, or a structured / time-limited (professional) CGM allowance — the outcome the evidence and hedged ADA-2026 language best support.',
      'Draft LCD H2 2026 → comment + open meeting → final mid-2027 → effective ~Q3–Q4 2027, landing alongside the Jan 2028 price cut.',
      'Positive but margin-compressed. A gate may capture 30–60% of the ~12M, phased in as documentation throttles uptake. Net revenue grows; realized uptake lags the headline TAM (the Eversense lesson).'),
     ('S2','Broad LCD expansion to all non-insulin T2D','20%',GREEN,LGREEN,
      'Minimal gating (diagnosis + visit cadence), mirroring how the 2023 final LCD came out broader than its draft. Driven by advocacy/comment volume, MAHA wearable enthusiasm, and lobbying.',
      'Draft H2 2026 → final broadened by comments → effective 2027. Fastest large-TAM unlock.',
      'Strongly positive on volume — approaches the full ~12M over time and de-risks ~6.5M commercial lives. Still subject to the price cut, but volume dominates. The manufacturers’ base case.')],
    'Scenarios in Detail (1 of 2): The Expansion Cases', 'S1 + S2 ≈ 55% probability of some expansion', 'S1–S2')

slide_no += 1
scen_detail_slide(
    [('S3','Status quo persists','25%',AMBER,LAMBER,
      'No expansion LCD finalizes in-horizon — delayed, withdrawn, or stuck. Drivers: depleted coverage staff, MAHA fiscal hawkishness, an evidence base critics call thin, CMS bandwidth on bidding.',
      'No draft, or a draft that does not finalize by early 2028. Eligibility frozen at 2023 criteria.',
      'Neutral on the thesis but NOT overall — the price cut still lands in 2028. Mild negative to flat; likely a "TAM-unlock-delayed" de-rating.'),
     ('S4','Restriction-led','15%',RED,LRED,
      'Beyond bidding, integrity pressure (OIG 69% overpayment, DOJ CGM fraud, 25% improper rate) prompts documentation tightening, prior authorization, or resupply audits — the PAP / power-mobility playbook. Expansion stalls.',
      'Prior auth / enhanced docs 2026–27; bidding SPAs at the low end (−30–45%); resupply audits expand.',
      'Negative. Volume throttled by friction and prior auth; price cut hard. The realistic downside is multi-year category attrition, not "non-coverage" — the highest-conviction bear path.')],
    'Scenarios in Detail (2 of 2): Stall & Restriction', 'S3 delay vs. S4 the highest-conviction bear case', 'S3–S4')

# ---------- SLIDE 10: S5 + PAYMENT CERTAINTY ----------
s = slide(); slide_no += 1
header(s, 'The Tail (S5) and the Certainty (Payment)', 'A low-odds binary, and the cross-cutting price cut present in every scenario', 'S5 + B')
box(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(5.55), fill=LRED, line=RED, line_w=2)
text(s, Inches(0.68), Inches(1.5), Inches(5.6), Inches(0.5), [{'segs':[('S5  ',20,RED,True,False),('—  NCD / CED tail   5%',20,NAVY,True,False)]}])
text(s, Inches(0.68), Inches(2.1), Inches(5.6), Inches(4.7),
     [{'segs':[('Mechanism:  ',12,RED,True,False),('CMS opens a National Coverage Analysis — either to lock in expansion (seat-elevation style) or, adversely, to impose Coverage-with-Evidence-Development.',12,NAVY,False,False)],'space_after':10,'line_spacing':1.0},
      {'segs':[('Why it matters:  ',12,RED,True,False),('CED without funded trials = de facto non-coverage (the TENS precedent). Restrictive NCD criteria are sticky for 20+ years (the insulin-pump lesson).',12,NAVY,False,False)],'space_after':10,'line_spacing':1.0},
      {'segs':[('Odds:  ',12,RED,True,False),('Low — ~4 NCAs/yr, depleted staff, a 1–2 yr pre-opening queue.',12,NAVY,False,False)],'space_after':10,'line_spacing':1.0},
      {'segs':[('Impact:  ',12,RED,True,False),('Binary, high-variance — a tail to hedge, not a base case.',12,NAVY,True,False)],'line_spacing':1.0}])
box(s, Inches(6.85), Inches(1.4), Inches(6.0), Inches(5.55), fill=LIGHT, line=NAVY, line_w=2)
text(s, Inches(7.03), Inches(1.5), Inches(5.6), Inches(0.5), [{'t':'PAYMENT: THE CROSS-CUTTING CERTAINTY','size':15,'color':NAVY,'bold':True}])
text(s, Inches(7.03), Inches(2.1), Inches(5.6), Inches(4.7),
     [{'segs':[('Magnitude:  ',12,ACCENT,True,False),('CB cuts DMEPOS 15–45%; test strips fell 79%. Plan for a ~20–35% CGM SPA cut.',12,NAVY,False,False)],'space_after':10,'bullet':True,'line_spacing':1.0},
      {'segs':[('Structure:  ',12,ACCENT,True,False),('Monthly rental + ~10 national suppliers compresses the distributor layer.',12,NAVY,False,False)],'space_after':10,'bullet':True,'line_spacing':1.0},
      {'segs':[('Wildcard:  ',12,ACCENT,True,False),('DIABETES Act (S.4037) would exempt CGMs 5 yrs — ~30–40% upside option, not a base case.',12,NAVY,False,False)],'space_after':10,'bullet':True,'line_spacing':1.0},
      {'segs':[('OTC gravity:  ',12,ACCENT,True,False),('~$89–99/mo cash anchors reimbursement down over time, bidding or not.',12,NAVY,False,False)],'space_after':10,'bullet':True,'line_spacing':1.0},
      {'segs':[('No scenario leaves 2028 unit price above today.',12.5,RED,True,False)],'line_spacing':1.0}])
footer(s, slide_no)

# ---------- SLIDE 11: SIGNPOSTS ----------
s = slide(); slide_no += 1
header(s, 'Signposts: What Would Change Our Mind', 'Track these to update probabilities in near-real time', 'Monitoring')
sign = [
    ('Draft LCD posted H2 2026 with an A1c / med gate', 'Raise S1 → ~45%', GREEN),
    ('Draft LCD posted with no gating criteria', 'Raise S2 → ~30%+', GREEN),
    ('No draft by YE2026 + integrity-focused messaging', 'Shift to S3 / S4', AMBER),
    ('CGM added to Required Prior Auth List / new TPE wave', 'Raise S4 materially', RED),
    ('An NCA tracking sheet for CGM appears', 'Re-underwrite S5 as binary', RED),
    ('DIABETES Act advances out of committee', 'Soften the price-cut assumption', GREEN),
    ('Bidding SPAs come in <$200/mo', 'Deepen price compression', RED),
    ('ADA / AACE upgrade non-insulin CGM to clear (A)', 'Raise S1 + S2 jointly', GREEN),
]
y = Inches(1.5)
box(s, Inches(0.5), y, Inches(12.33), Inches(0.42), fill=NAVY)
text(s, Inches(0.65), y, Inches(7.5), Inches(0.42), [{'t':'If you observe…','size':12.5,'color':WHITE,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(8.3), y, Inches(4.4), Inches(0.42), [{'t':'Update toward…','size':12.5,'color':WHITE,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
y = Emu(y + Inches(0.42))
for obs, upd, color in sign:
    rh = Inches(0.6)
    box(s, Inches(0.5), y, Inches(12.33), rh, fill=LIGHT, line=WHITE, line_w=1)
    text(s, Inches(0.65), y, Inches(7.5), rh, [{'t':obs,'size':12.5,'color':NAVY,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(8.3), y, Inches(4.4), rh, [{'t':upd,'size':12.5,'color':color,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], anchor=MSO_ANCHOR.MIDDLE)
    y = Emu(y + rh + Pt(1.5))
footer(s, slide_no)

# ---------- SLIDE 12: PRIMARY DILIGENCE ----------
s = slide(); slide_no += 1
header(s, 'Primary Diligence: Who to Speak With', 'Edge = non-public read on draft-LCD timing, gating criteria, and the bidding price cut', 'Diligence')
tiers = [
    ('TIER 1 — Closest to the decision', NAVY, LIGHT, [
        'Current & former DME MAC Medical Directors (Noridian: Caldarella — endocrinologist, Ballyamanda, Mamuya; CGS: Lalla) — author the LCD',
        'Former CMS Coverage & Analysis Group leadership (e.g., ex-CAG Director T. Syrek Jensen) — NCD-vs-LCD routing, CED risk',
        'Former CMS DMEPOS / competitive-bidding payment officials — realistic SPA discount range, rental mechanics',
        'DME MAC reconsideration / medical-policy staff — is a request pending? comment calendar',
    ]),
    ('TIER 2 — Expert intermediaries & advocates', ACCENT, LIGHT, [
        'Reimbursement attorneys & consultancies (Epstein Becker Green, DLA Piper, Applied Policy, Avalere, ADVI) — base rates, timing',
        'ADA, Breakthrough T1D, ADCES, Endocrine Society policy staff — reconsideration filed? MAC receptiveness',
        'CGM trialists (MOBILE / CONNECT / WISDM investigators) — will evidence clear "reasonable & necessary"?',
    ]),
    ('TIER 3 — Market & channel reality checks', GREY, LIGHT, [
        'DME distributors / suppliers (Byram, Edgepark/Cardinal, US MED) — bid behavior, margin at $200–270/mo, consolidation',
        'Manufacturer market-access / gov-affairs leads (DexCom, Abbott) — timeline beyond earnings-call language',
        'MA & commercial payer medical directors / PBM leads — covering non-insulin ahead of FFS? UM posture',
        'Prescribers & diabetes educators — documentation burden; how a gate throttles real-world uptake',
    ]),
]
y = Inches(1.4)
for title_t, color, fill, items in tiers:
    n_items = len(items)
    bh = Inches(0.42) + Inches(0.34) * n_items + Inches(0.12)
    box(s, Inches(0.5), y, Inches(12.33), bh, fill=fill, line=color, line_w=1.5)
    text(s, Inches(0.65), Emu(y+Inches(0.06)), Inches(12), Inches(0.35), [{'t':title_t,'size':13,'color':color,'bold':True}])
    iy = Emu(y + Inches(0.44))
    for it in items:
        text(s, Inches(0.75), iy, Inches(12.0), Inches(0.34),
             [{'t':it,'size':10.5,'color':NAVY,'bullet':True,'bcolor':color,'line_spacing':0.95}])
        iy = Emu(iy + Inches(0.34))
    y = Emu(y + bh + Inches(0.1))
footer(s, slide_no)

# ---------- SLIDE 13: SEQUENCING + COI ----------
s = slide(); slide_no += 1
header(s, 'Diligence Sequencing & Conflict Discipline', 'A four-step path to de-risk the two questions that matter most', 'Plan')
steps = [
    ('1. CALIBRATE', '2–3 reimbursement consultancies / attorneys to set the base case and the calendar — fast, broad, non-confidential.', ACCENT),
    ('2. GO DEEP', 'A former DME MAC medical director + a former CMS payment official — the two people who most directly de-risk timing and the SPA magnitude.', NAVY),
    ('3. REALITY-CHECK', 'Distributors + MA/PBM medical directors to ground the TAM-to-revenue conversion and the real price.', GREEN),
    ('4. CONFIRM THESIS', 'Advocacy + academic calls to test whether evidence and the advocacy push are strong enough to force the modal S1/S2 vs. an S3 delay.', AMBER),
]
y = Inches(1.55)
for label, body, color in steps:
    box(s, Inches(0.5), y, Inches(2.7), Inches(1.05), fill=color, line=None)
    text(s, Inches(0.6), y, Inches(2.5), Inches(1.05), [{'t':label,'size':15,'color':WHITE,'bold':True,'anchor':MSO_ANCHOR.MIDDLE}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, Inches(3.3), y, Inches(9.53), Inches(1.05), fill=LIGHT, line=color, line_w=1.5)
    text(s, Inches(3.5), y, Inches(9.2), Inches(1.05), [{'t':body,'size':12.5,'color':NAVY,'anchor':MSO_ANCHOR.MIDDLE,'line_spacing':1.0}], anchor=MSO_ANCHOR.MIDDLE)
    y = Emu(y + Inches(1.2))
box(s, Inches(0.5), Inches(6.4), Inches(12.33), Inches(0.7), fill=LAMBER, line=AMBER, line_w=1.5)
text(s, Inches(0.65), Inches(6.4), Inches(12.0), Inches(0.7),
     [{'segs':[('Conflict discipline:  ',12.5,AMBER,True,False),
               ('manufacturer, advocacy & consultancy sources are long the expansion. Weight former-government and independent-distributor sources most for timing and the price cut, where the incentive to spin is lowest.',12.5,NAVY,False,False)],'anchor':MSO_ANCHOR.MIDDLE,'line_spacing':1.0}],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, slide_no)

# ---------- SLIDE 14: CLOSING / KEY TAKEAWAYS ----------
s = slide(); slide_no += 1
box(s, 0, 0, SW, SH, fill=NAVY)
box(s, Inches(0.7), Inches(0.6), Inches(11.9), Pt(3), fill=ACCENT)
text(s, Inches(0.7), Inches(0.75), Inches(12), Inches(0.8), [{'t':'Bottom Line','size':32,'color':WHITE,'bold':True}])
takeaways = [
    ('Two tracks, not one.', 'Coverage is expanding (eligibility) while payment is being cut (price). Most market commentary conflates them.'),
    ('The next instrument is an LCD, not an NCD.', 'Every CGM change since 2017 used sub-NCD instruments; NCDs are rare and slowing.'),
    ('Some expansion is the modal outcome (~55%) — most likely conditioned.', 'The evidence supports a structured / A1c-gated non-insulin expansion over blanket coverage.'),
    ('The payment cut is effectively certain.', 'Competitive bidding + monthly rental is finalized, effective ≤ Jan 1, 2028; only magnitude and the DIABETES Act delay are open.'),
    ('Bull = CPAP; Bear = power mobility.', 'CGM sits between a barrier-lowering volume boom and a documentation/bidding attrition of category spend.'),
    ('Edge is in primary diligence.', 'Former DME MAC medical directors and CMS payment officials can de-risk the two questions that matter: timing and the price cut.'),
]
y = Inches(1.75)
for head, body in takeaways:
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.75), Emu(y+Inches(0.07)), Inches(0.22), Inches(0.22))
    dot.fill.solid(); dot.fill.fore_color.rgb = ACCENT; dot.line.fill.background(); dot.shadow.inherit=False
    text(s, Inches(1.2), y, Inches(11.4), Inches(0.85),
         [{'segs':[(head+'  ',15,WHITE,True,False),(body,13.5,RGBColor(0xBF,0xD3,0xEC),False,False)],'line_spacing':1.0}])
    y = Emu(y + Inches(0.88))
prs.save('/home/user/knutnyman/cgm-lcd-ncd-analysis/CGM_LCD_NCD_Scenario_Analysis.pptx')
print('PowerPoint saved with', slide_no, 'slides.')
