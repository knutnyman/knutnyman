#!/usr/bin/env python3
"""Generate the CGM LCD/NCD scenario analysis Word document."""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x33, 0x5E)
ACCENT = RGBColor(0x2E, 0x74, 0xB5)
GREY = RGBColor(0x59, 0x59, 0x59)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
AMBER = RGBColor(0xB8, 0x86, 0x0B)

doc = Document()

# ---- base styles ----
normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.08

def set_cell_bg(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def h1(text):
    p = doc.add_heading(level=1)
    r = p.add_run(text)
    r.font.color.rgb = NAVY
    r.font.size = Pt(16)
    r.font.name = 'Calibri'
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    return p

def h2(text):
    p = doc.add_heading(level=2)
    r = p.add_run(text)
    r.font.color.rgb = ACCENT
    r.font.size = Pt(12.5)
    r.font.name = 'Calibri'
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    return p

def h3(text):
    p = doc.add_heading(level=3)
    r = p.add_run(text)
    r.font.color.rgb = NAVY
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.name = 'Calibri'
    return p

def para(text=None, runs=None, bold=False, italic=False, color=None, size=None, align=None, after=6):
    p = doc.add_paragraph()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    if runs:
        for seg in runs:
            r = p.add_run(seg[0])
            r.font.bold = seg[1].get('bold', False)
            r.font.italic = seg[1].get('italic', False)
            if seg[1].get('color'): r.font.color.rgb = seg[1]['color']
            if seg[1].get('size'): r.font.size = Pt(seg[1]['size'])
    elif text is not None:
        r = p.add_run(text)
        r.font.bold = bold
        r.font.italic = italic
        if color: r.font.color.rgb = color
        if size: r.font.size = Pt(size)
    return p

def bullet(text, level=0, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
    p.paragraph_format.space_after = Pt(3)
    if bold_lead:
        r = p.add_run(bold_lead)
        r.font.bold = True
        p.add_run(text)
    else:
        p.add_run(text)
    return p

def numbered(text, bold_lead=None):
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.space_after = Pt(3)
    if bold_lead:
        r = p.add_run(bold_lead); r.font.bold = True
        p.add_run(text)
    else:
        p.add_run(text)
    return p

def table(headers, rows, widths=None, header_bg='1F335E', font_size=9, zebra=True):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = ''
        run = hdr[i].paragraphs[0].add_run(htext)
        run.font.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_bg(hdr[i], header_bg)
    for ridx, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ''
            para0 = cells[i].paragraphs[0]
            # allow tuple (text, color) for emphasis
            if isinstance(val, tuple):
                run = para0.add_run(val[0]); run.font.color.rgb = val[1]; run.font.bold = True
            else:
                run = para0.add_run(str(val))
            run.font.size = Pt(font_size)
            if zebra and ridx % 2 == 1:
                set_cell_bg(cells[i], 'EEF2F8')
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def hr():
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '2E74B5')
    pbdr.append(bottom); pPr.append(pbdr)

# =====================================================================
# COVER
# =====================================================================
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = title.add_run('Continuous Glucose Monitors')
r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = NAVY
sub = doc.add_paragraph()
r = sub.add_run('Scenario Analysis of a New Medicare LCD / NCD')
r.font.size = Pt(18); r.font.color.rgb = ACCENT
doc.add_paragraph()
r = doc.add_paragraph().add_run('Probability-Weighted Outcomes, Precedent Analysis & Investment Implications')
r.font.size = Pt(13); r.font.italic = True; r.font.color.rgb = GREY
doc.add_paragraph()
hr()
meta = doc.add_paragraph()
for label, val in [('Prepared for: ', 'Investment Diligence — Internal'),
                   ('Date: ', 'June 10, 2026'),
                   ('Classification: ', 'Confidential — Draft for Discussion'),
                   ('Horizon: ', 'Coverage & payment events through YE2028')]:
    rr = meta.add_run(label); rr.font.bold = True; rr.font.size = Pt(10.5); rr.font.color.rgb = NAVY
    rr2 = meta.add_run(val + '\n'); rr2.font.size = Pt(10.5)
hr()
note = doc.add_paragraph()
r = note.add_run('Methodology note: ')
r.font.bold = True; r.font.size = Pt(9); r.font.color.rgb = GREY
r = note.add_run('Synthesized from a five-stream evidence sweep (CMS coverage database, Federal '
    'Register, OIG/GAO, peer-reviewed clinical literature, manufacturer SEC filings and earnings calls, and '
    'health-policy trade press) with adversarial verification of load-bearing claims. Probabilities are '
    'analyst judgments anchored to base rates from analogous Medicare coverage decisions; they are estimates, '
    'not forecasts, and should be stress-tested through the primary diligence plan in Section 9.')
r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GREY

doc.add_page_break()

# =====================================================================
# EXEC SUMMARY
# =====================================================================
h1('1. Executive Summary')

para(runs=[('The central finding of this analysis is that CGM Medicare policy is moving on ', {}),
           ('two opposing tracks at once', {'bold': True}),
           (', and conflating them is the most common analytical error in the market:', {})])

table(
    ['', 'Track A — Eligibility (Coverage)', 'Track B — Payment (Price)'],
    [
     ['Direction', ('Expansionary ▲', GREEN), ('Restrictive ▼', RED)],
     ['Instrument', 'DME MAC LCD L33822 revision (or NCD)', 'DMEPOS Competitive Bidding + monthly rental'],
     ['Status (Jun 2026)', 'Anticipated, not yet proposed', 'FINALIZED Nov 28, 2025 — now law'],
     ['Key event', 'Draft LCD for non-insulin T2D, expected H2 2026', 'Contracts effective ≤ Jan 1, 2028'],
     ['Volume effect', 'Adds up to ~12M eligible lives', 'Neutral to volume'],
     ['Price effect', 'Neutral to price', 'Cuts unit reimbursement ~15–40%'],
    ],
    widths=[1.3, 2.8, 2.8], font_size=9)

para(runs=[('Investors fixated on the expansion story (TAM unlock to non-insulin Type 2 diabetes) are '
    'under-weighting the ', {}),
    ('already-finalized payment restructuring', {'bold': True}),
    (', which is the single highest-confidence event in this analysis. Conversely, the expansion itself is '
     'probable but ', {}),
    ('not yet a filed proposal', {'bold': True}),
    (' and carries real timing and scope risk.', {})])

h2('1.1 Headline Conclusions')
numbered('the next material coverage instrument is overwhelmingly likely to be a DME MAC LCD revision, '
    'not an NCD. Every CGM coverage change since 2017 (the 2017 Ruling, the 2021 fingerstick removal, the '
    '2023 expansion) has travelled through sub-NCD instruments. NCDs are rare (~4/year program-wide) and the '
    'coverage staff is down 22% since December 2024.', bold_lead='Instrument: ')
numbered('an LCD expansion of eligibility to non-insulin Type 2 diabetes is the modal outcome, most '
    'likely in a conditioned/structured form (A1c threshold, hypoglycemia-risk medications, or '
    'time-limited/professional CGM) rather than blanket coverage. We weight some form of expansion at ~55% '
    'within the horizon.', bold_lead='Coverage: ')
numbered('a payment cut is effectively certain. Competitive bidding + reclassification to monthly rental was '
    'finalized in November 2025; the only live question is magnitude and whether Congress (DIABETES Act, '
    'S.4037) delays it. This compresses per-patient economics regardless of the coverage outcome.',
    bold_lead='Payment: ')
numbered('the bull case is a CPAP-style volume boom (lower qualification barrier → decade of growth); the '
    'bear case is a power-mobility-style attrition (documentation audits + prior authorization + competitive '
    'bidding cutting category spend 70–80% over a decade). Both precedents are real and CGM sits between them.',
    bold_lead='Net: ')

h2('1.2 Probability-Weighted Scenario Set (coverage-instrument outcome, ~24–36 month horizon)')
table(
    ['#', 'Scenario', 'Prob.', 'Net effect on manufacturers'],
    [
     ['S1', 'Conditioned LCD expansion to non-insulin T2D (A1c / hypo-med / structured)', ('35%', NAVY), ('Positive — volume up, price down', GREEN)],
     ['S2', 'Broad LCD expansion to all non-insulin T2D', ('20%', NAVY), ('Strongly positive on volume', GREEN)],
     ['S3', 'Status quo persists through horizon (expansion slips past 2027)', ('25%', NAVY), ('Neutral; price still cut', AMBER)],
     ['S4', 'Restriction-led: bidding bites + documentation/PA tightening, no expansion', ('15%', NAVY), ('Negative', RED)],
     ['S5', 'NCD / CED tail (national action; restrictive trapdoor possible)', ('5%', NAVY), ('Binary — large up or down', RED)],
    ],
    widths=[0.4, 4.0, 0.7, 2.6], font_size=9)
para('Expansion of some kind (S1+S2) ≈ 55%. A do-no-harm-to-volume outcome (S1+S2+S3) ≈ 80%. '
     'Materially adverse coverage actions (S4+restrictive leg of S5) ≈ ~18%. The payment cut (Track B) '
     'is treated as a near-certain cross-cutting modifier present in all five scenarios.', italic=True, size=9.5, color=GREY)

doc.add_page_break()

# =====================================================================
# 2. REGULATORY HISTORY
# =====================================================================
h1('2. How We Got Here: The Coverage & Payment History')
para('Understanding the precedent trajectory inside CGM itself is the strongest predictor of the next move. '
     'The pattern is unambiguous: coverage has expanded through the lightest-weight instruments available, '
     'each step lowering the qualification barrier, while payment policy has only now turned restrictive.')

table(
    ['Date', 'Action', 'Significance'],
    [
     ['Jan 12, 2017', 'CMS Ruling 1682-R: "therapeutic" (non-adjunctive) CGM classified as DME', 'Created the Medicare benefit; codes K0554 / K0553'],
     ['Jul 18, 2021', 'LCD L33822 removes ≥4×/day fingerstick prerequisite', 'CMS: no evidence fingerstick frequency predicts CGM benefit'],
     ['Dec 28, 2021', 'DMEPOS final rule CMS-1738-F: adjunctive CGMs also DME', 'Effective Feb 28, 2022; broadens device scope'],
     ['Jan 1, 2023', 'Code transition: A4239 (supply) / E2103 (receiver) replace K-codes', 'Current billing architecture'],
     ['Apr 16, 2023', 'LCD expansion: ANY insulin (basal-only incl.) OR non-insulin w/ problematic hypoglycemia', '~1.5M basal-only beneficiaries newly eligible'],
     ['Nov 2025', 'OIG OEI-04-23-00430: payments 69% above supplier cost ($377M/yr)', 'Program-integrity trigger for payment cut'],
     ['Nov 28, 2025', 'FINAL rule: CGMs/pumps → competitive bidding + monthly rental', 'Contracts effective ≤ Jan 1, 2028; bid limit $272.69/mo'],
     ['H2 2026 (exp.)', 'Anticipated draft LCD: non-insulin Type 2 diabetes', 'Not yet filed as of Jun 2026'],
    ],
    widths=[1.0, 3.1, 2.9], font_size=8.5)

h2('2.1 Current Coverage Criteria (LCD L33822, effective April 16, 2023)')
para('CGM is covered when the beneficiary has diabetes, has sufficient training, the device is used per FDA '
     'indications, there is a qualifying practitioner visit within 6 months (and every 6 months thereafter), '
     'AND at least one of:')
bullet('the beneficiary is insulin-treated — with no minimum injection frequency, so basal-only qualifies; OR',
       bold_lead='Pathway A: ')
bullet('the beneficiary has a history of problematic hypoglycemia — documented as recurrent (>1) Level 2 events '
       '(<54 mg/dL) despite ≥2 treatment adjustments, OR one Level 3 event (<54 mg/dL) requiring third-party '
       'assistance.', bold_lead='Pathway B: ')
para('The clear gap, and the target of the next expansion, is the non-insulin Type 2 patient who does NOT '
     'have documented problematic hypoglycemia — the ~12M-life population manufacturers are pursuing.', italic=True)

h2('2.2 The Payment Restructuring (finalized November 28, 2025)')
bullet('CGMs and insulin pumps reclassified to the "frequent and substantial servicing" category, paid as a '
       'bundled monthly rental (device + all supplies) rather than purchase-receiver + monthly supply allowance.',
       bold_lead='Mechanism: ')
bullet('Class II CGMs enter the next Competitive Bidding round via a national "Remote Item Delivery" model; '
       '~10 national contract suppliers expected.', bold_lead='Competition: ')
bullet('Monthly bid limit $272.69 (2025 fees: A4239 $267.92 + E2103 $286.03 ÷ 60), indexed forward. Single '
       'payment amounts are set by competitive bids, historically 15–45% below fee schedule across DMEPOS.',
       bold_lead='Price: ')
bullet('Bid window late summer/fall 2026 → SPAs late 2027 → contracts effective no later than Jan 1, 2028; '
       '6-month beneficiary transition.', bold_lead='Timeline: ')
bullet('Up to 3 months of rental billable in advance; the 5-year replacement rule is eliminated.',
       bold_lead='Mechanics: ')

doc.add_page_break()

# =====================================================================
# 3. MARKET CONTEXT
# =====================================================================
h1('3. Market Context: What Is at Stake')

h2('3.1 Medicare CGM Spend & Utilization')
table(
    ['Metric', 'Value', 'Source / note'],
    [
     ['Medicare Part B CGM spend, 2018', '$109M', 'OIG OEI-04-23-00430 (Nov 2025)'],
     ['Medicare Part B CGM spend, 2023', '$1.3B', '≈12× growth in 5 years'],
     ['OIG overpayment vs. supplier acquisition cost', '$377M (69%) / yr', 'Jul 2022–Jun 2023 window'],
     ['Glucose-monitor improper payment rate (2024)', '25.2% ($278.5M)', 'CMS FFS supplemental data'],
     ['Medicare beneficiaries on CGM (approx.)', '~2M', 'Supplier estimate — low confidence'],
     ['CGM use among MA Type 2 insulin users', '1.4% (Jan-21) → 17.2% (2023)', 'JMCP 2026 — adoption inflected post-2023'],
     ['A4239 monthly supply allowance (2025)', '$267.92', 'CMS; ~$273 for 2026'],
    ],
    widths=[3.1, 1.9, 2.0], font_size=8.5)

h2('3.2 Eligible Populations (the TAM ladder)')
table(
    ['Population segment', 'Approx. size (US)', 'Medicare coverage status'],
    [
     ['Type 1 + intensive insulin', 'Covered base', 'Covered since 2017'],
     ['Basal-only insulin (Type 2)', '~1.5M Medicare added', 'Covered since Apr 2023'],
     ['Non-insulin T2D w/ problematic hypoglycemia', 'Subset', 'Covered since Apr 2023'],
     ['Non-insulin Type 2 (no hypoglycemia)', '~25M US; ~half Medicare-age', 'NOT covered — the expansion target'],
     ['Prediabetes / wellness', '~31M (65+) prediabetes', 'Not covered; OTC channel only'],
    ],
    widths=[3.2, 1.9, 1.9], font_size=8.5)
para('Manufacturer framing: Dexcom publicly sizes the non-insulin Type 2 Medicare opportunity at ~12M lives '
     'and frames CMS coverage as "the largest single driver" of full coverage; its May 2026 Investor Day '
     'modeled covered lives rising from ~15M to ~30M if expansion proceeds.', size=9.5, color=GREY, italic=True)

h2('3.3 Competitive & Channel Landscape')
table(
    ['Manufacturer', 'CGM revenue (latest FY)', 'Global CGM share (2024)', 'Medicare exposure note'],
    [
     ['Abbott (FreeStyle Libre)', '>$7.5B (FY2025)', '~57%', 'Pharmacy-channel strength; OTC Lingo/Libre Rio'],
     ['DexCom (G6/G7/Stelo)', '$4.66B (FY2025)', '~35%', 'Historically overweight in Medicare DME channel'],
     ['Medtronic (Guardian/Simplera)', '$2.76B diabetes (FY25)', '~7%', 'MiniMed spin-off announced May 2025'],
     ['Senseonics (Eversense 365)', '~$35M (FY2025)', '<1%', 'Implantable; coverage ≠ revenue inflection'],
    ],
    widths=[2.2, 1.9, 1.4, 2.2], font_size=8.5)
bullet('Most commercial access runs through the pharmacy benefit; Medicare runs through DME suppliers/'
       'distributors. DexCom exited direct Medicare billing in 2020 (distributors: Byram, Cardinal/Edgepark).',
       bold_lead='Channel: ')
bullet('OTC CGMs (Stelo ~$89–99/mo; Lingo $89/2-pack; Libre Rio) sit ~60–65% below the Medicare A4239 '
       'allowance — a price benchmark OIG explicitly used to argue Medicare overpays. The OTC cash price is '
       'now a gravitational anchor pulling reimbursement down.', bold_lead='OTC: ')
bullet('GLP-1s appear to be a modest tailwind, not a substitute: Abbott and DexCom both report sustained or '
       'higher CGM utilization among GLP-1 users (self-interested but directionally consistent).',
       bold_lead='GLP-1: ')

doc.add_page_break()

# =====================================================================
# 4. CLINICAL EVIDENCE
# =====================================================================
h1('4. The Clinical Evidence That Drives the Decision')
para('A coverage expansion to non-insulin Type 2 diabetes must clear CMS\'s "reasonable and necessary" bar. '
     'On the evidence as of mid-2026, the case is borderline-and-strengthening for selected/structured '
     'populations, but not a slam-dunk for blanket coverage.')

h2('4.1 Evidence by Population')
table(
    ['Population', 'Key evidence', 'Effect size', 'Strength'],
    [
     ['Intensive insulin (T1D/T2D)', 'DIAMOND, GOLD, HypoDE', 'Large TIR / A1c / hypo benefit', ('Settled (A)', GREEN)],
     ['Older adults (T1D)', 'WISDM (JAMA 2020, n=203, ≥60y)', '−1.9% time <70 mg/dL', ('Strong', GREEN)],
     ['Basal insulin T2D', 'MOBILE (JAMA 2021, n=175)', '−0.4% A1c (basis of 2023 expansion)', ('Adequate', GREEN)],
     ['Non-insulin T2D', 'CONNECT (Jun 2026, n=283); IMMEDIATE; Wada', 'CONNECT −0.9% A1c; pooled ~−0.3%', ('Contested', AMBER)],
     ['Prediabetes / wellness', 'No supportive RCT', 'Null / minimal', ('Weak', RED)],
    ],
    widths=[1.9, 2.6, 1.9, 1.0], font_size=8.5)

h2('4.2 The Pivotal New Data Point: CONNECT (June 2026)')
para(runs=[('Dexcom\'s CONNECT RCT (n=283 non-insulin Type 2 adults, 22 US primary-care sites, G7 vs. routine '
    'care, 26 weeks) reported a ', {}),
    ('~0.9-point greater A1c reduction', {'bold': True}),
    (' and ~5 hrs/day more time-in-range, across metformin/GLP-1/SGLT2 combinations. Baseline A1c was 8.8% '
     '(31% ≥9%). Marketed as the "first Level A evidence" in non-insulin T2D. ', {}),
    ('Caveats: ', {'bold': True}),
    ('industry-sponsored, open-label, A1c (not hard outcomes), 26 weeks.', {})])

h2('4.3 Guideline Posture')
bullet('Grade (B) for CGM in non-insulin T2D on hypoglycemia-causing therapies and where it aids management — '
       'but text explicitly cautions benefit is "not consistently" shown outside structured programs.',
       bold_lead='ADA 2026: ')
bullet('Grade A for intensive insulin; Grade B for less-intensive insulin T2D; did not broadly endorse '
       'non-insulin T2D.', bold_lead='AACE 2021: ')
bullet('Restricts CGM/flash to insulin-treated T2D meeting specific criteria — does NOT cover non-insulin '
       'T2D. A credible "responsible payer" anchor for skeptics.', bold_lead='NICE (UK): ')

h2('4.4 What a Skeptic at a MAC Contractor Advisory Meeting Would Say')
bullet('Pooled non-insulin A1c effects are small and shrinking — a 25-RCT behavior-change meta-analysis '
       '(Richardson 2024) found only −0.28% with non-significant effects on time-above-range, BMI, and weight.')
bullet('Positive trials (CONNECT, IMMEDIATE, Wada) are open-label with education co-interventions — benefit '
       'may reflect coaching, not the device.')
bullet('Durability is unproven (trials ≤26 weeks; MOBILE shows benefit fades after discontinuation → indefinite cost).')
bullet('Hard-outcome (hospitalization) evidence is uncontrolled claims data (Optum, RELIEF) subject to '
       'regression to the mean and selection.')
bullet('Industry funding pervades the positive signal; NICE and USPSTF decline non-insulin/wellness endorsement.')
para('Implication: the evidence best supports a CONDITIONED expansion (structured programs, higher baseline '
     'A1c, hypoglycemia-risk medications, or time-limited/professional CGM) rather than indefinite personal '
     'CGM for every non-insulin patient — which is precisely why S1 outweighs S2.', bold=True, color=NAVY)

doc.add_page_break()

# =====================================================================
# 5. PRECEDENT ANALYSIS
# =====================================================================
h1('5. Precedent Analysis: What Analogous Decisions Teach Us')
para('No coverage decision happens in a vacuum. The following Medicare precedents bracket the plausible '
     'outcomes for CGM and supply concrete base rates for direction, timeline, and market impact.')

h2('5.1 Precedent Map')
table(
    ['Precedent', 'What happened', 'Market impact', 'Lesson for CGM'],
    [
     ['CPAP NCD (2008)', 'Lowered diagnostic barrier (home sleep test) + 12-wk adherence trial', 'Decade-plus volume boom; PSG spend +39%', 'BULL case: barrier-lowering → growth (but invites supply audits)'],
     ['Power mobility (2006–18)', 'Documentation + prior auth + competitive bidding', 'Spend collapsed ~$964M→$190M (~80%)', 'BEAR case: attrition without any "non-coverage" decision'],
     ['Test strips (CB 2011–15)', 'Competitive bidding', 'Price −79% ($1.6B→$0.3B)', 'Closest supply analog: bidding can cut CGM unit price hard'],
     ['TENS for CLBP (2012)', 'Coverage-with-Evidence-Development; no studies funded', 'De facto national non-coverage', 'CED is a trap door if CMS demands trials'],
     ['Seat elevation NCD (2023)', 'Advocacy-driven expansion; final broader than proposed', 'New benefit; comments moved scope', 'Advocacy + comment volume move CMS'],
     ['Insulin pump NCD 280.14', 'C-peptide gate (1999), minor 2004 loosening', 'Sticky 20+ yrs; constrains AID in Medicare', 'Restrictive criteria in an NCD are very hard to reverse'],
     ['Eversense I-CGM LCD', 'MAC parallel coverage; criteria converged to 2023 rules', 'Doubled eligible pop., little revenue inflection', 'Coverage ≠ revenue; product economics dominate'],
    ],
    widths=[1.5, 2.4, 1.9, 2.2], font_size=8)

h2('5.2 Process Base Rates (the timing reality)')
bullet('Statute: proposed decision within 6 months of opening (9 with technology assessment/MEDCAC); final '
       'within 60 days of comment close — i.e., ~9–12 months once opened.', bold_lead='NCD clock: ')
bullet('GAO (Sept 2025): CMS met NCA timeframes 83% of the time (44/53 over 12.5 yrs) — but only ~4 NCAs '
       'complete per year, and the coverage group lost 22% of staff since Dec 2024. NCDs are rare and '
       'getting slower.', bold_lead='NCD throughput: ')
bullet('Post-2019 (21st Century Cures): ≥45-day comment period + open meeting; MAC must finalize or retire a '
       'proposed LCD within 365 days. Faster and more flexible than an NCD in both directions.',
       bold_lead='LCD clock: ')
bullet('CMS removed 6 NCDs effective Jan 2021 and retired the amyloid-PET NCD in 2023, devolving to MACs — '
       'the structural current favors the LCD-governed status quo for CGM.', bold_lead='Trend: ')
bullet('Comment volume and advocacy demonstrably move CMS (seat-elevation final broader than proposed; the '
       '2023 CGM LCD finalized faster and broader than the draft; the 2025 skin-substitute LCDs were '
       'finalized then withdrawn under pressure).', bold_lead='Politics: ')

h2('5.3 The Two Reference Trajectories')
para(runs=[('BULL (CPAP analog): ', {'bold': True, 'color': GREEN}),
           ('the 2023 CGM LCD already executed the same barrier-lowering move CPAP did in 2008 (dropping the '
            'fingerstick/MDI prerequisite). A further drop to non-insulin T2D could replicate a multi-year '
            'volume ramp. The catch CPAP also teaches: a volume boom in a resupply item invites the exact '
            'OIG/audit scrutiny CGM is now under.', {})])
para(runs=[('BEAR (power-mobility analog): ', {'bold': True, 'color': RED}),
           ('rapidly growing resupply spend + fraud headlines (the June 2025 DOJ takedown featured CGM '
            'telemarketing schemes) + competitive bidding is precisely the cocktail that cut power-wheelchair '
            'spend ~80% over a decade — without any formal non-coverage decision. CGM is already three of '
            'four steps into this template.', {})])

doc.add_page_break()

# =====================================================================
# 6. SCENARIOS
# =====================================================================
h1('6. The Scenarios in Detail')
para('Each scenario below specifies the triggering mechanism, the most probable timeline, the leading '
     'indicators to watch, and the revenue/utilization consequence. Probabilities are for the coverage-'
     'instrument outcome over a ~24–36 month horizon; the Track B payment cut is assumed present in all.')

def scenario_block(tag, name, prob, color, mechanism, timeline, indicators, impact):
    h2(f'{tag} — {name}  ({prob})')
    para(runs=[('Mechanism: ', {'bold': True, 'color': color}), (mechanism, {})], after=3)
    para(runs=[('Most-likely path & timeline: ', {'bold': True, 'color': color}), (timeline, {})], after=3)
    para(runs=[('Revenue / utilization impact: ', {'bold': True, 'color': color}), (impact, {})], after=3)
    p = doc.add_paragraph(); r = p.add_run('Leading indicators: '); r.font.bold = True; r.font.color.rgb = color
    for ind in indicators:
        bullet(ind)

scenario_block('S1', 'Conditioned LCD Expansion to Non-Insulin Type 2 Diabetes', '35% — MODAL',
    GREEN,
    'DME MACs publish a draft LCD revising L33822 to add non-insulin Type 2 patients, but gated by '
    'qualifying conditions — an A1c threshold (e.g., ≥8%), use of hypoglycemia-risk medications '
    '(sulfonylureas), enrollment in a structured/DSMES program, or a time-limited/professional CGM allowance. '
    'This is the outcome the CONNECT/MOBILE evidence and the hedged ADA-2026 language best support, and it '
    'lets CMS expand access while containing budget exposure and answering skeptics.',
    'Draft LCD H2 2026 → ≥45-day comment + open meeting → final mid-2027 → effective ~Q3–Q4 2027. Runs in '
    'parallel with competitive-bidding implementation (Jan 2028), so the volume unlock and the price cut '
    'land almost simultaneously.',
    ['Draft LCD DL33822 posted to the Medicare Coverage Database',
     'DME MAC open-meeting notice referencing glucose monitors / non-insulin criteria',
     'CONNECT full publication + ADA/AACE guideline upgrade to a clearer (B)/(A)',
     'Manufacturer earnings-call language shifting from "expect a proposal" to "proposal posted"'],
    'Positive but margin-compressed. Eligible lives rise materially (a structured gate might capture 30–60% '
    'of the ~12M, phased in over years as documentation requirements throttle uptake). Per-patient revenue '
    'falls on the rental/bidding cut. Net manufacturer revenue still grows, but ASP erosion and '
    'documentation friction mean realized uptake lags the headline TAM — the Eversense lesson that coverage '
    '≠ revenue applies.')

scenario_block('S2', 'Broad LCD Expansion to All Non-Insulin Type 2 Diabetes', '20%',
    GREEN,
    'DME MACs expand to non-insulin Type 2 with minimal gating (diagnosis + visit cadence), mirroring how the '
    '2023 final LCD dropped the proposed injection-frequency language and came out broader than the draft. '
    'Driven by strong advocacy/comment volume (ADA, Breakthrough T1D, ADCES), MAHA-administration rhetorical '
    'enthusiasm for wearables, and manufacturer lobbying.',
    'Draft H2 2026 → final broadened in response to comments → effective 2027. Fastest large-TAM unlock.',
    ['High comment volume skewing pro-expansion (seat-elevation precedent: ~2,130 comments)',
     'Explicit MAHA / HHS endorsement of CGM access in rulemaking preambles',
     'A draft LCD whose criteria lack an A1c or medication gate'],
    'Strongly positive on volume — approaches the full ~12M Medicare opportunity over time and de-risks the '
    'parallel ~6.5M commercial lives manufacturers cite. Still subject to the Track B price cut, but volume '
    'dominates. This is the manufacturers\' base case and the market\'s implicit expectation.')

scenario_block('S3', 'Status Quo Persists — Expansion Slips Past the Horizon', '25%',
    AMBER,
    'No expansion LCD is finalized within the horizon. The draft is delayed, withdrawn, or stuck in comment '
    'resolution. Drivers: a 22%-depleted coverage workforce, MAHA fiscal hawkishness colliding with its own '
    'pro-wearable rhetoric, an evidence base critics still call thin for blanket coverage, and CMS bandwidth '
    'consumed by competitive-bidding implementation.',
    'Draft posts but does not finalize by YE2027/early-2028; or no draft posts at all in 2026. Eligibility '
    'frozen at the 2023 criteria.',
    ['Calendar slips past H2 2026 with no draft LCD',
     'CMS public statements emphasizing budget neutrality / program integrity over expansion',
     'Coverage-staff attrition or leadership vacuum at CAG/MAC medical-director level'],
    'Neutral on the expansion thesis — but NOT neutral overall, because the Track B price cut still lands in '
    'Jan 2028. Net effect is mild negative to flat: same covered population at a lower unit price. The market '
    'would likely re-rate manufacturers down on a "TAM unlock delayed" narrative.')

scenario_block('S4', 'Restriction-Led — Bidding Bites and Documentation/Prior-Auth Tightens', '15%',
    RED,
    'Beyond the finalized competitive bidding, program-integrity pressure (OIG 69%-overpayment finding, DOJ '
    'CGM-fraud takedowns, 25% improper-payment rate) prompts the DME MACs to tighten documentation, add prior '
    'authorization to the CGM Required Prior Authorization List, or audit resupply/continued-use compliance — '
    'the PAP-supplies and power-mobility playbook. Expansion stalls in parallel.',
    'Prior-authorization or enhanced documentation rolled out 2026–2027; competitive-bidding single payment '
    'amounts come in at the low end (−30% to −45%); resupply audits expand.',
    ['CGM added to the Required Prior Authorization List or a TPE (Targeted Probe & Educate) wave',
     'New OIG/GAO work-plan items on CGM resupply or continued-use documentation',
     'Competitive-bidding SPAs (late 2027) materially below the $272.69 limit',
     'Supplier consolidation / distributor exits as margins compress'],
    'Negative. Volume growth throttled by documentation friction and prior auth (power-mobility showed 80% of '
    'claims failing requirements once scrutiny arrived); unit price cut hard by bidding. The realistic downside '
    'is not "non-coverage" but multi-year category-spend attrition — the highest-conviction bear path.')

scenario_block('S5', 'National Action — NCD / Coverage-with-Evidence-Development Tail', '5%',
    RED,
    'CMS opens a National Coverage Analysis to standardize CGM policy nationally — either to lock in expansion '
    '(advocacy-driven, seat-elevation style) or, in the adverse leg, to impose Coverage-with-Evidence-'
    'Development for marginal non-insulin populations. CED without funded trials equals de facto non-coverage '
    '(the TENS precedent).',
    'NCA opened 2026–2027 → ~9–12 months to final once opened, but a 1–2 year pre-opening queue is typical. '
    'Low probability given ~4 NCAs/year throughput and a depleted coverage staff.',
    ['An NCA tracking sheet for CGM appears in the Medicare Coverage Database',
     'A formal NCD request filed by a manufacturer, society, or CMS-internal initiation',
     'Any CMS signal of demanding outcome trials (CED language) for non-insulin CGM'],
    'Binary. The expansion leg would standardize and durably enlarge coverage (very positive); the CED/'
    'restriction leg would freeze or reverse non-insulin coverage nationally and be hard to undo (very '
    'negative — the insulin-pump-NCD stickiness lesson). Low odds, high variance — a tail to hedge, not a base case.')

doc.add_page_break()

# =====================================================================
# 7. CROSS-CUTTING: PAYMENT
# =====================================================================
h1('7. The Cross-Cutting Certainty: Payment Compression')
para('Whatever happens on coverage, the payment restructuring is finalized and will reshape per-patient '
     'economics. This is the most under-appreciated element of the CGM Medicare story.')
bullet('Competitive bidding historically cuts DMEPOS prices 15–45%; the closest supply analog (diabetes test '
       'strips) fell 79%. A 20–35% CGM single-payment-amount cut is a reasonable planning range.',
       bold_lead='Magnitude: ')
bullet('Monthly rental + ~10 national contract suppliers compresses the distributor layer and may strand '
       'smaller suppliers — a channel-consolidation event independent of manufacturer ASP.',
       bold_lead='Structure: ')
bullet('The DIABETES Act (S.4037, Mar 2026) would exempt CGMs/pumps from competitive bidding for 5 years. '
       'Bipartisan Diabetes Caucus support (Shaheen, Collins, DeGette, Bilirakis) is real but passage is '
       'uncertain; treat a delay as a ~30–40% probability upside option, not a base case.',
       bold_lead='Wildcard: ')
bullet('OTC CGMs at ~$89–99/mo have established a public price anchor far below Medicare\'s ~$268, which OIG '
       'weaponized. Over time this anchor pulls reimbursement toward cash-pay levels regardless of bidding.',
       bold_lead='OTC gravity: ')
para('Synthesis: model CGM Medicare revenue as VOLUME (driven by the coverage scenario) × PRICE (almost '
     'certainly lower from 2028). The bull coverage cases (S1/S2) are partially offset by Track B; the bear '
     'cases (S3/S4) compound with it. There is no scenario in which 2028 per-unit Medicare price is higher '
     'than today.', bold=True, color=NAVY)

# =====================================================================
# 8. WHAT WOULD CHANGE OUR MIND
# =====================================================================
h1('8. Signposts & Falsification')
para('Track these to update the probabilities in near-real time:')
table(
    ['If you observe…', 'Update toward…'],
    [
     ['Draft LCD DL33822 posted H2 2026 with an A1c/medication gate', 'S1 (raise to ~45%)'],
     ['Draft LCD posted with no gating criteria', 'S2 (raise to ~30%+)'],
     ['No draft LCD by YE2026 + integrity-focused CMS messaging', 'S3 / S4'],
     ['CGM added to Required Prior Authorization List or new TPE wave', 'S4 (raise materially)'],
     ['An NCA tracking sheet for CGM appears', 'S5 (re-underwrite as binary)'],
     ['DIABETES Act advances out of committee', 'Track B relief — soften price-cut assumption'],
     ['Competitive-bidding SPAs come in <$200/mo', 'Deepen price-compression assumption across all scenarios'],
     ['ADA/AACE upgrade non-insulin CGM to clear (A)', 'Raise S1+S2 jointly'],
    ],
    widths=[4.3, 2.7], font_size=9)

doc.add_page_break()

# =====================================================================
# 9. PRIMARY DILIGENCE
# =====================================================================
h1('9. Primary Diligence: Who to Speak With')
para('The single biggest information edge here is non-public read on (a) whether and when the DME MACs will '
     'post a draft non-insulin LCD, (b) the likely gating criteria, and (c) how aggressively competitive-'
     'bidding SPAs will cut price. Prioritize sources closest to the DME MAC LCD process and CMS coverage/'
     'payment policy. Below, sources are tiered by value and paired with the specific questions only they '
     'can answer.')

h2('9.1 Tier 1 — Closest to the Decision (highest priority)')
table(
    ['Who', 'Why they have edge', 'What to ask'],
    [
     ['Current & former DME MAC Medical Directors (Noridian: Caldarella — an endocrinologist, Ballyamanda, '
      'Mamuya; CGS: Lalla)', 'Author and vote the Glucose Monitors LCD; know the internal evidence threshold '
      'and timeline', 'Likelihood/timing of a non-insulin draft LCD; what gating criteria they would require; '
      'how CONNECT changes their view; prior-auth appetite'],
     ['Former CMS Coverage & Analysis Group leadership (e.g., Tamara Syrek Jensen, ex-CAG Director, now '
      'Rubrum Advising)', 'Defined modern CED; knows NCD-vs-LCD routing logic and CMS internal posture',
      'Would CMS ever nationalize CGM via NCD? CED risk for non-insulin? How staff cuts affect timelines'],
     ['Former CMS DMEPOS / competitive-bidding payment officials (CMM / DME pricing group alumni)',
      'Built the bidding + monthly-rental architecture', 'Realistic SPA discount range; rental mechanics; '
      'how inherent-reasonableness interacts; DIABETES Act odds'],
     ['DME MAC LCD reconsideration / contractor medical policy staff', 'Process gatekeepers for draft LCDs',
      'Is a reconsideration request pending? Valid/invalid status? Comment-period calendar'],
    ],
    widths=[2.3, 2.3, 2.4], font_size=8)

h2('9.2 Tier 2 — Expert Intermediaries & Advocates')
table(
    ['Who', 'Why', 'What to ask'],
    [
     ['Health-policy / reimbursement attorneys (Epstein Becker Green, DLA Piper, King & Spalding DME teams)',
      'Advise manufacturers on exactly this LCD/CB process', 'Base rates for draft-LCD finalization; client '
      'read on timing; litigation risk on the CB rule'],
     ['Reimbursement consultancies (Applied Policy, Rubrum, ADVI, Avalere)', 'Track the tracking sheets; model SPAs',
      'Their probability/timeline on expansion; SPA modeling; congressional handicapping'],
     ['ADA, Breakthrough T1D, ADCES, Endocrine Society policy staff', 'File the comment letters; meet CMS/MACs',
      'Have they filed a reconsideration request? Read on MAC receptiveness; advocacy plan for non-insulin'],
     ['Academic endocrinologists & CGM trialists (MOBILE/CONNECT/WISDM investigators; e.g., Martens, Aronson, Pratley)',
      'Define the evidence CMS weighs', 'Will the evidence clear "reasonable & necessary" for blanket vs. '
      'structured coverage? Durability and non-insulin effect-size honesty'],
    ],
    widths=[2.3, 2.3, 2.4], font_size=8)

h2('9.3 Tier 3 — Market & Channel Reality Checks')
table(
    ['Who', 'Why', 'What to ask'],
    [
     ['DME distributors / suppliers (Byram, Edgepark/Cardinal, ADS, US MED)', 'Live the channel; will bid in CB',
      'Expected bid behavior; margin at $200–270/mo rental; consolidation; documentation-denial rates'],
     ['Manufacturer market-access / government-affairs leaders (DexCom, Abbott)', 'Drive the lobbying & evidence strategy',
      'Their LCD timeline expectations beyond earnings-call language; gating they would accept; CB mitigation'],
     ['Medicare Advantage & commercial payer medical directors / PBM formulary leads', 'Leading indicator of coverage drift',
      'Are they covering non-insulin T2D ahead of FFS? Utilization-management posture; OTC substitution'],
     ['Practicing primary-care & endocrinology prescribers + diabetes educators', 'Determine real-world uptake vs. TAM',
      'Documentation burden; how a structured-program gate would throttle prescribing; patient demand'],
    ],
    widths=[2.3, 2.3, 2.4], font_size=8)

h2('9.4 Recommended Sequencing')
numbered('Start with 2–3 reimbursement consultancies/attorneys to calibrate the base case and the calendar '
    '(fast, broad, non-confidential).', bold_lead='Calibrate: ')
numbered('Then a former DME MAC medical director and a former CMS payment official — the two people who can '
    'most directly de-risk the coverage timeline and the SPA magnitude.', bold_lead='Go deep: ')
numbered('Cross-check with distributors and MA/PBM medical directors to ground the TAM-to-revenue conversion '
    'and price reality.', bold_lead='Reality-check: ')
numbered('Use advocacy and academic calls to confirm whether the evidence and the advocacy push are strong '
    'enough to force the modal S1/S2 outcome vs. the S3 delay.', bold_lead='Confirm thesis: ')
para('Conflict-of-interest discipline: manufacturer, advocacy, and consultancy sources are directionally '
     'long the expansion. Weight former government and independent-distributor sources most heavily for the '
     'timing and the price cut, where the incentive to spin is lowest.', italic=True, size=9.5, color=GREY)

doc.add_page_break()

# =====================================================================
# 10. APPENDIX
# =====================================================================
h1('Appendix A. Key Facts & Confidence')
table(
    ['Fact', 'Value', 'Conf.'],
    [
     ['Competitive bidding + monthly rental finalized', 'Nov 28, 2025', ('High', GREEN)],
     ['CB contracts effective', '≤ Jan 1, 2028', ('High', GREEN)],
     ['CGM monthly bid limit', '$272.69', ('High', GREEN)],
     ['A4239 monthly supply allowance (2025)', '$267.92', ('High', GREEN)],
     ['2023 LCD expansion (any insulin / hypo)', 'Apr 16, 2023', ('High', GREEN)],
     ['OIG overpayment vs. supplier cost', '$377M / 69% per yr', ('High', GREEN)],
     ['Medicare Part B CGM spend 2018→2023', '$109M → $1.3B', ('High', GREEN)],
     ['CONNECT non-insulin T2D A1c benefit', '~−0.9% vs. routine', ('High', GREEN)],
     ['Non-insulin T2D Medicare opportunity (mfr. est.)', '~12M lives', ('Medium', AMBER)],
     ['Draft non-insulin LCD as of Jun 2026', 'NOT yet posted', ('Med-High', AMBER)],
     ['Dexcom-expected expansion proposal', 'H1/H2 2026 (mfr. guidance)', ('Medium', AMBER)],
     ['NCAs completed per year (program-wide)', '~4', ('High', GREEN)],
     ['CMS coverage staff change since Dec 2024', '−22%', ('High', GREEN)],
    ],
    widths=[3.8, 2.0, 1.1], font_size=8.5)

h1('Appendix B. Source Base')
para('This analysis was built from primary and authoritative secondary sources, including:', size=9.5)
for s in [
    'CMS: Ruling 1682-R; LCD L33822 & Policy Article A52464; DMEPOS final rules CMS-1738-F (2021) and '
    'CMS-1828-F (Nov 2025); Competitive Bidding Program fact sheets; DMEPOS fee schedule.',
    'OIG: OEI-04-23-00430 (Nov 2025) CGM payment report; FFS improper-payment supplemental data.',
    'GAO-25-107623 (Sept 2025) on NCD timelines and CMS coverage staffing.',
    'Clinical: MOBILE (JAMA 2021); CONNECT (ADA 2026); WISDM (JAMA 2020); IMMEDIATE (DOM 2023); Wada (BMJ '
    'Open DRC 2020); Richardson behavior-change meta-analysis (IJBNPA 2024); Optum (Garg 2024) & RELIEF '
    '(Diabetes Care 2021) real-world studies.',
    'Guidelines: ADA Standards of Care 2026; AACE 2021 Advanced Diabetes Technology; NICE NG28/QS209.',
    'Precedents: CMS NCDs 280.14 (insulin pump), 240.2 (oxygen), 240.4 (CPAP), 160.27 (TENS/CLBP), 280.16 '
    '(seat elevation); MolDX foundation LCDs; Eversense I-CGM LCDs; power-mobility & test-strip CB history.',
    'Markets: DexCom, Abbott, Medtronic, Senseonics SEC filings & FY2023–FY2025 earnings calls; Mordor '
    'Intelligence CGM share; MedTech Dive, Modern Healthcare, Applied Policy, AAHomecare, DLA Piper analyses.',
    'Policy/stakeholder: DIABETES Act S.4037 (2026); Diabetes Caucus letters; AdvaMed/ADA/Breakthrough T1D/'
    'ADCES/Endocrine Society positions; DOJ June 2025 fraud takedown; OpenSecrets lobbying data.',
]:
    bullet(s, )
    doc.paragraphs[-1].runs[0].font.size = Pt(9)

para('')
discl = doc.add_paragraph()
r = discl.add_run('Disclaimer: ')
r.font.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = GREY
r = discl.add_run('This document is an internal diligence aid. Probabilities are analyst judgments, not '
    'forecasts. Several quantitative figures derive from manufacturer estimates or single analyst sources '
    'and are flagged accordingly; verify load-bearing assumptions through the Section 9 primary diligence '
    'plan before relying on them for an investment decision. Not investment advice.')
r.font.size = Pt(8.5); r.font.italic = True; r.font.color.rgb = GREY

doc.save('/home/user/knutnyman/cgm-lcd-ncd-analysis/CGM_LCD_NCD_Scenario_Analysis.docx')
print('Word document saved.')
