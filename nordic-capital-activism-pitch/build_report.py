#!/usr/bin/env python3
"""Builds the Word report: Nordic Capital Engaged Equities strategy pitch."""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x0B, 0x25, 0x45)
GOLD = RGBColor(0xC4, 0x9A, 0x3B)
GREY = RGBColor(0x59, 0x5959 >> 8 & 0xFF, 0x59)

doc = Document()

# ---------- base styles ----------
style = doc.styles["Normal"]
style.font.name = "Georgia"
style.font.size = Pt(10.5)
style.paragraph_format.space_after = Pt(8)
style.paragraph_format.line_spacing = 1.15

for name, size, color, space_before in [
    ("Heading 1", 16, NAVY, 18),
    ("Heading 2", 13, NAVY, 14),
    ("Heading 3", 11.5, GOLD, 10),
]:
    h = doc.styles[name]
    h.font.name = "Georgia"
    h.font.size = Pt(size)
    h.font.color.rgb = color
    h.font.bold = True
    h.paragraph_format.space_before = Pt(space_before)
    h.paragraph_format.space_after = Pt(6)


def p(text, bold=False, italic=False, size=None, color=None, align=None, style_name=None):
    par = doc.add_paragraph(style=style_name)
    run = par.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    if align:
        par.alignment = align
    return par


def bullets(items, level=0):
    for it in items:
        par = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        if isinstance(it, tuple):
            lead, rest = it
            r = par.add_run(lead)
            r.bold = True
            par.add_run(rest)
        else:
            par.add_run(it)
        par.paragraph_format.space_after = Pt(4)


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0]
    for j, htext in enumerate(headers):
        cell = hdr.cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(htext)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "0B2545")
        cell._tc.get_or_add_tcPr().append(shading)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.rows[i + 1].cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            run.font.size = Pt(9.5)
            if i % 2 == 1:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "F2F4F8")
                cell._tc.get_or_add_tcPr().append(shading)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Inches(w)
    doc.add_paragraph()
    return t


# ---------- cover ----------
for _ in range(6):
    doc.add_paragraph()
p("STRICTLY PRIVATE & CONFIDENTIAL", bold=True, size=10, color=GOLD,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p("Nordic Capital Engaged Equities", bold=True, size=28, color=NAVY,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p("A Constructive Activism Strategy for US and European Public Markets",
  size=14, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
p("Proposal to launch a USD 1.5 billion engaged public-equities vehicle",
  italic=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
for _ in range(8):
    doc.add_paragraph()
p("Prepared for: Managing Partner & Executive Committee, Nordic Capital",
  size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
p("June 2026", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
p("Illustrative discussion material. Company references are based solely on public "
  "information and do not constitute investment recommendations.",
  italic=True, size=8.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_page_break()

# ---------- 1. Executive summary ----------
doc.add_heading("1. Executive Summary", level=1)
p("We propose that Nordic Capital launch a dedicated public-markets strategy — "
  "“Nordic Capital Engaged Equities” — a USD 1.5 billion (hard cap USD 2.0 billion) "
  "vehicle pursuing constructive activism in listed US and European mid-caps. The strategy "
  "applies the firm's private equity value-creation playbook — operational improvement, "
  "portfolio simplification, disciplined capital allocation and strategic M&A — to a "
  "concentrated book of 6–8 public companies in the firm's core sectors, where Nordic "
  "Capital takes 2–10% ownership stakes, secures board-level influence, and drives a "
  "2–4 year value plan. Europe (including the Nordic home market) is expected to carry "
  "50–60% of the book and the United States 40–50%, with allocation following sector "
  "opportunity rather than geographic quota.")
bullets([
    ("Activism is at record levels and increasingly M&A-driven. ", "2025 was a record year "
     "for global shareholder campaigns, and an unusually high share of recent campaigns carry "
     "M&A demands — terrain where a private equity sponsor has structural advantages no "
     "hedge-fund activist can match. The US remains the deepest activism market; Europe is "
     "the fastest-growing and least crowded."),
    ("Two regimes, one playbook. ", "In Europe — and above all in the Nordic home market — "
     "shareholder-elected nomination committees, low EGM thresholds and ownership "
     "transparency let a 4–8% holder win board influence quietly. In the US, the universal "
     "proxy card has structurally lowered the cost of board challenges, making credible, "
     "well-resourced engagement effective at 2–6% stakes. The same value plans work in "
     "both; only the influence mechanics differ."),
    ("The sector lens travels; the competition doesn't. ", "Nordic Capital's underwriting "
     "edge in healthcare & medtech, technology & payments, financial services and "
     "industrials applies identically to a Philadelphia medtech and a Stockholm one. "
     "Operationally credible, sponsor-backed engagement capital is scarce on both sides of "
     "the Atlantic — Cevian dominates constructive activism in Europe but has no US book, "
     "while US activists lack operating benches and European legitimacy."),
    ("Nordic Capital is uniquely positioned. ", "Three decades of operating credibility, "
     ">EUR 31bn AUM, a New York presence and a growing North American investment record in "
     "healthcare and technology, and an unrivalled bench of chairs, CEOs and operating "
     "advisors to propose as directors — the single hardest resource for any activist to "
     "assemble."),
    ("Strategic synergies compound the core P&L. ", "The public strategy becomes a pipeline "
     "for take-privates and structured PIPEs, a catalyst for carve-outs the buyout funds can "
     "acquire, and an extension of LP relationships into a new product line — diversifying "
     "fee streams beyond the flagship buyout franchise."),
])
p("We recommend a phased launch: internal approval and conflicts framework in H2 2026, a "
  "founding team of 11–13 led by a partner-level Head of Engaged Equities — with US "
  "situations and campaign strategy led by Knut Nyman, joining from Elliott Investment "
  "Management — a EUR 150–200m GP/employee seed commitment, first close of ~USD 750m in "
  "mid-2027, and initial deployment timed to the 2028 board cycles in both regions. Target "
  "net returns of 15%+ over a cycle, with management economics accretive to the platform "
  "from year two.")
doc.add_page_break()

# ---------- 2. Why now ----------
doc.add_heading("2. The Opportunity: Why Public-Market Activism, Why Now", level=1)
doc.add_heading("2.1 The valuation gap has become a strategic opening", level=2)
p("On both sides of the Atlantic, a large cohort of quality mid-cap franchises in Nordic "
  "Capital's core sectors trades at 30–50% discounts to intrinsic value on sum-of-the-parts "
  "or normalised-margin bases. In Europe, the discount is structural: persistent "
  "under-valuation versus US peers, compounded in mid-caps by a decade of sell-side "
  "consolidation and MiFID II research unbundling. In the US, the post-2021 de-rating of "
  "medtech, payments and software left a generation of fallen-angel quality compounders at "
  "private-market discounts even as indices set records — and the activism toolkit there "
  "has never been stronger. In both regions, public-market multiples for these companies "
  "sit below the entry multiples Nordic Capital's buyout funds routinely pay in competitive "
  "private processes — the public market is offering the firm's core hunting ground at a "
  "discount, without an auction.")
doc.add_heading("2.2 Activism has gone mainstream — and M&A-centric", level=2)
bullets([
    "2025 set a record for global activist campaigns, with activity broadening across "
    "geographies, market caps and campaign types.",
    "In late 2025, roughly six in ten campaigns carried an M&A thesis (break-ups, "
    "divestitures, sale processes) — the highest share in five years. M&A-driven activism "
    "plays directly to a private equity sponsor's strengths: valuing businesses, running "
    "separations, and standing ready as a buyer or financing partner.",
    "Universal owners (index funds, pension institutions) increasingly support well-argued "
    "engagement agendas on both continents, lowering the vote-getting cost of campaigns.",
    "Competition is mis-shaped, not absent: US activism is crowded at the mega-cap end but "
    "thin in operationally complex mid-caps; in Europe, Cevian (constructive, large-cap "
    "drift) and episodic US visitors leave the EUR 2–10bn segment structurally underserved. "
    "Operationally credible, sponsor-backed engagement capital is scarce everywhere.",
])
doc.add_heading("2.3 Two engagement regimes, one playbook", level=2)
p("The strategy runs the same underwriting and the same value plans in both regions; only "
  "the influence mechanics differ:")
table(
    ["Dimension", "Europe / Nordics (home field)", "United States"],
    [
        ["Path to the board",
         "Shareholder-elected nomination committees (Nordics: top 3–4 holders as of Aug/Sep "
         "records) and constructive sphere engagement; a 4–8% stake routinely wins direct "
         "influence without a proxy fight.",
         "Universal proxy card (2022) structurally lowered the cost of board challenges; "
         "most campaigns now settle for seats. Credible slates + 2–6% stakes suffice."],
        ["Escalation tools",
         "Low EGM thresholds (10% in Sweden); ownership transparency (public registers) "
         "enables precise coalition mapping.",
         "13D platform, precision proxy contests, settlement negotiation; deep advisor "
         "ecosystem (solicitors, banks, law firms)."],
        ["Culture & cost",
         "Consensus boards; private argumentation usually sufficient; low campaign cost; "
         "brand-compatible conduct.",
         "Faster, more transactional; settlements common within 6–12 months; higher "
         "campaign cost, higher liquidity."],
        ["Our edge",
         "Domestic legitimacy of a 30-year Stockholm-rooted owner; the region's best "
         "director bench; Cevian-validated model with no second practitioner.",
         "Sector operating credibility US activists lack; PE-grade separation expertise "
         "for the M&A-heavy campaign mix; team hired from the US activist front line."],
    ],
    widths=[1.3, 2.8, 2.7],
)
p("Proof of concept exists in both regimes: Cevian's two-decade, board-seat-driven record "
  "in Europe (low-to-mid-teens net over cycles), and the post-universal-proxy wave of US "
  "settlements showing that well-resourced engagement now converts to board influence at "
  "historically small stakes. What does not yet exist is a practitioner combining PE "
  "operating depth with a transatlantic mandate. Nordic Capital can be that firm, with "
  "advantages neither Cevian nor the US funds have: a buyout balance sheet behind the "
  "engagement, and operating resources measured in hundreds of professionals.")
doc.add_page_break()

# ---------- 3. Why Nordic Capital ----------
doc.add_heading("3. Why Nordic Capital Wins", level=1)
bullets([
    ("Operating credibility, not financial engineering. ", "Three decades and >EUR 31bn of "
     "AUM built on operational transformation in healthcare, technology & payments, financial "
     "services and industrial & business services. When Nordic Capital tells a board the "
     "margin gap to best-in-class is 400bps, it speaks as an owner that has closed that gap "
     "repeatedly in private portfolios."),
    ("The network is the moat. ", "An unmatched bench of Nordic chairs, former CEOs and "
     "operating advisors who can be proposed as board candidates — the single hardest "
     "resource for any activist to assemble, and the one that decides outcomes in a "
     "nomination-committee system."),
    ("Legitimacy in Europe, credibility in the US. ", "European institutions, family "
     "spheres and governments engage very differently with a Stockholm-rooted owner of 30 "
     "years than with a hedge-fund activist — lower friction means cheaper, faster, quieter "
     "wins. In the US, the firm's New York presence, North American healthcare and "
     "technology investments, and a campaign team hired from the US activist front line "
     "give it standing that European engagement funds have never built."),
    ("Sector depth = underwriting edge, in both regions. ", "The firm's diligence "
     "pattern-recognition in medtech, software, payments and specialty financials applies "
     "identically to listed mid-caps in Boston, Amsterdam and Stockholm — much of the "
     "universe comprises businesses Nordic Capital has studied, bid on, or competed "
     "against."),
    ("Two-way pipeline with the buyout franchise. ", "Public engagement surfaces "
     "take-private and PIPE opportunities for the flagship funds; conversely, activism that "
     "catalyses carve-outs creates proprietary deal flow. (Subject to the information-barrier "
     "framework in Section 8 — the synergy operates at the level of market intelligence and "
     "brand, with strict legal separation of MNPI.)"),
    ("LP demand for a second product. ", "Existing LPs seeking Nordic exposure with "
     "liquidity, plus dedicated activist/engagement allocators, give the raise a warm "
     "starting book. Evolution II's four-month, hard-cap close demonstrates the franchise's "
     "fundraising velocity."),
])

# ---------- 4. Strategy ----------
doc.add_heading("4. Investment Strategy", level=1)
doc.add_heading("4.1 Strategy statement", level=2)
p("Acquire 2–10% positions in 6–8 listed US and European companies with USD 2–15bn market "
  "capitalisation, in Nordic Capital's core sectors, where a clearly identifiable gap "
  "exists between market price and intrinsic value, and where that gap can be closed "
  "within 2–4 years through actions the company itself can take: operational improvement, "
  "portfolio separation or divestiture, capital-allocation reform, or strategic "
  "transactions. Influence is exercised through each market's governance machinery — "
  "nomination committees, board representation and sphere engagement in Europe; "
  "universal-proxy-backed private engagement and negotiated board refreshment in the US — "
  "with public pressure reserved as an escalation, not a default.", )
doc.add_heading("4.2 The engagement playbook", level=2)
table(
    ["Phase", "Months", "Actions"],
    [
        ["1. Underwrite", "0–4",
         "PE-grade diligence: full operational benchmark, SOTP, downside case. Thesis must "
         "survive an internal “red team” and clear a 30%+ upside to base-case "
         "intrinsic value with limited structural downside."],
        ["2. Accumulate", "3–9",
         "Build 2–10% via open market and blocks, timed to the board calendar: Q3 ownership "
         "snapshots for Nordic nomination committees; nomination windows for US slates. "
         "Disclose at thresholds; first private meeting with chair before crossing 5%."],
        ["3. Engage", "6–18",
         "Present value plan privately to board and anchor shareholders. Europe: secure "
         "nomination-committee seat; propose 1–2 directors from the operating network. US: "
         "negotiate board refreshment under credible universal-proxy optionality. Align "
         "with index funds and domestic institutions."],
        ["4. Execute", "12–36",
         "Board-level sponsorship of the plan: margin program, separation/divestiture, "
         "balance-sheet reset, management upgrades where needed. Quarterly milestones "
         "tracked against the underwrite."],
        ["5. Realise", "24–48",
         "Exit into re-rating, strategic sale, or — where compelling and conflict-cleared — "
         "a take-private led by Nordic Capital's buyout funds with the position rolled or "
         "cashed at a premium."],
    ],
    widths=[1.2, 0.9, 4.7],
)
doc.add_heading("4.3 Three thesis archetypes", level=2)
bullets([
    ("Margin-gap operators (est. 40–50% of book). ", "Quality franchises running 300–600bps "
     "below achievable margins due to execution, supply chain or quality-system issues. "
     "Value plan: operational program with board-level accountability."),
    ("Conglomerate & SOTP discounts (est. 30–40%). ", "Multi-division companies where "
     "separation, divestiture or monetisation of hidden assets (stakes, real estate, "
     "infrastructure) unlocks 30–60% upside. Plays directly to record M&A-activism demand "
     "and to Nordic Capital's carve-out expertise."),
    ("Capital-allocation & strategic-catalyst situations (est. 15–25%). ", "Lazy balance "
     "sheets, value-destructive M&A patterns, or companies that should initiate or accept "
     "strategic processes. Often overlaps with the first two archetypes."),
])
doc.add_heading("4.4 What we will not do", level=2)
bullets([
    "No hostile public campaigns as an opening move; escalation requires Investment "
    "Committee approval and a reputational review.",
    "No positions inside controlled spheres where the controller is unwilling to engage "
    "(screened out at underwrite).",
    "No shorting of engagement targets; no activist short selling under the Nordic Capital "
    "brand.",
    "No more than two concurrent campaigns in their public-escalation phase.",
])
doc.add_page_break()

# ---------- 5. Focus areas ----------
doc.add_heading("5. Focus Areas", level=1)
p("Sector focus mirrors the firm's private equity franchises, applied across both regions; "
  "geography follows the sector opportunity rather than a quota:")
table(
    ["Sector", "Indicative weight", "Why", "Illustrative listed universe (US & EU)"],
    [
        ["Healthcare & MedTech", "30–40%",
         "Nordic Capital's deepest franchise; on both continents, quality medtech "
         "franchises sit in margin or execution repair at deep discounts.",
         "Elekta, Smith+Nephew, Philips, Getinge, GN Store Nord / Baxter, Zimmer Biomet, "
         "Teleflex, Henry Schein"],
        ["Industrials & Business Services", "25–35%",
         "Largest segment of the listed universe; rich in conglomerate discounts, "
         "separations and self-help stories.",
         "Orkla, SKF, Dometic, ISS, Valmet / mid-cap US multi-industry and services "
         "names in separation or margin repair"],
        ["Technology, Software & Payments", "15–25%",
         "Post-2021 de-ratings left quality software, IT-services and payments assets at "
         "private-market discounts; consolidation logic is strong.",
         "Tieto, Nexi-linked European payment assets / Global Payments, Fiserv-class "
         "de-rated US payments and vertical software"],
        ["Financial Services", "10–20%",
         "Specialty finance, insurance and savings platforms with capital-return and "
         "consolidation angles.",
         "Storebrand, European specialty lenders / US specialty insurers and savings "
         "platforms"],
    ],
    widths=[1.4, 0.9, 2.2, 2.3],
)
p("Geographic spread: Europe 50–60% of the book — of which the Nordics 20–30% (home-field "
  "mechanics, deepest networks), UK/Benelux/DACH the balance — and the United States "
  "40–50% (deepest and most liquid activism market; universal proxy regime). Single-name "
  "limits apply uniformly; FX hedged to the fund's USD base.")

# ---------- 6. Target universe ----------
doc.add_heading("6. Target Universe & Screening", level=1)
bullets([
    ("Universe: ", "US and European listed companies in the firm's four core sectors; "
     "~1,300 names in the USD 2–15bn band after liquidity screens (roughly 55% US, "
     "45% Europe)."),
    ("Screen 1 — Quality: ", "defensible market position, structurally sound gross "
     "margins, identifiable best-in-class peer set. (~300 names survive.)"),
    ("Screen 2 — Value gap: ", ">30% upside to intrinsic value on normalised margins or "
     "SOTP. (~80–100 names.)"),
    ("Screen 3 — Path to influence: ", "register structure permitting an influential "
     "stake — nomination-committee relevance in Europe, credible universal-proxy "
     "optionality in the US; no blocking holder hostile to change; no poison-pill or "
     "staggered-board structure that defeats the thesis timeline. (~40 names.)"),
    ("Screen 4 — Conflict check: ", "no MNPI contamination from current PE processes; "
     "cleared by compliance before any accumulation. (Live pipeline: 12–16 names, of "
     "which 6–8 funded.)"),
])
p("A USD 1.5bn book across 6–8 core positions implies USD 180–280m per position — i.e., "
  "4–10% of a USD 3–6bn European mid-cap, or 2–5% of a USD 5–15bn US company, both "
  "sufficient for board-level influence under the respective regimes.")
doc.add_page_break()

# ---------- 7. Illustrative pipeline ----------
doc.add_heading("7. Illustrative Pipeline (Public Information Only)", level=1)
p("The following six situations illustrate the strategy's archetypes using publicly "
  "available information as of early/mid 2026. They are illustrations for discussion, not "
  "recommendations, and would each require full underwriting and conflict clearance.",
  italic=True)

doc.add_heading("7.1 Elekta (Sweden — MedTech) — Margin-gap operator", level=2)
bullets([
    "Global #2 in radiation oncology, a structurally growing oligopoly, yet the stock has "
    "de-rated dramatically; third-party fair-value estimates have placed the discount at "
    "~50%+ at points in 2026. FY2025/26 adjusted EBIT margin of ~12% versus "
    "mid-to-high-teens at Siemens Healthineers' Varian — a gap attributable to supply "
    "chain, order execution and mix, not market position.",
    "Value plan: board-sponsored margin program, order-to-revenue execution reset, "
    "service-revenue acceleration; strategic optionality given medtech consolidation "
    "logic. Home-field engagement via the Swedish nomination-committee system.",
])
doc.add_heading("7.2 Smith+Nephew (UK — MedTech) — Margin gap & structural options", level=2)
bullets([
    "Quality orthopaedics/sports-medicine/wound franchise that has lagged peers on margin "
    "and execution for years; an established engagement situation (Cevian on the register "
    "since 2024) validating the value gap, with the turnaround only partly delivered.",
    "Value plan: hold the board to peer-level margins; evaluate structural options "
    "including separation of underperforming franchises and a US listing review — a "
    "live debate where a second credible engaged owner can be decisive.",
])
doc.add_heading("7.3 Philips (Netherlands — HealthTech) — Recovery & SOTP", level=2)
bullets([
    "Global healthtech franchise emerging from the Respironics recall era (US litigation "
    "settled), with margins still well below pre-crisis levels and peers; anchor "
    "shareholder (Exor) already on the register demonstrates openness to engaged owners.",
    "Value plan: margin normalisation with board accountability; portfolio review of "
    "Personal Health versus the healthcare core — a separation long argued by analysts "
    "that would force the SOTP into the price.",
])
doc.add_heading("7.4 Baxter International (US — MedTech) — Completing the simplification", level=2)
bullets([
    "Multi-year portfolio simplification already in motion (BioPharma Solutions sold; "
    "Vantive kidney-care divested to Carlyle in early 2025), yet the equity has continued "
    "to de-rate amid execution stumbles and a CEO transition — a fallen angel trading far "
    "below medtech peer multiples.",
    "Value plan: finish the job — complete the separation program, reset the cost base of "
    "the remaining core, refresh the board, and apply divestiture proceeds to a "
    "disciplined capital-return framework. Classic universal-proxy-era US engagement.",
])
doc.add_heading("7.5 Global Payments (US — Payments) — Capital allocation & credibility reset",
                level=2)
bullets([
    "Scaled merchant-acquiring franchise trading at a deeply depressed multiple after the "
    "poorly received 2025 Worldpay/Issuer Solutions restructuring; activist interest has "
    "been publicly reported, underscoring the value gap.",
    "Value plan: integration milestones with board-level accountability, divestiture of "
    "non-core assets, binding capital-return commitments post-deleveraging, and board "
    "refreshment with payments operating expertise — squarely within Nordic Capital's "
    "payments franchise (Nets/Nexi heritage).",
])
doc.add_heading("7.6 Zimmer Biomet (US — MedTech) — Capital-allocation catalyst", level=2)
bullets([
    "Orthopaedics leader compounding modestly but trading at a structurally depressed "
    "multiple, with investor confidence eroded by guidance resets and M&A choices.",
    "Value plan: capital-allocation framework (organic investment and buybacks over "
    "dilutive M&A), portfolio pruning of subscale adjacencies, and margin program toward "
    "peer levels.",
])
p("Watchlist (further candidates surviving Screens 1–3): Europe — Getinge, GN Store Nord, "
  "Orkla, Tieto, Demant, Securitas (self-help largely delivered — a model of the re-rating "
  "we underwrite), ISS, Valmet, Storebrand. US — Teleflex (separation announced), Henry "
  "Schein, Fiserv, mid-cap specialty insurers.", italic=True)
doc.add_page_break()

# ---------- 8. Fund structure ----------
doc.add_heading("8. Vehicle, Size & Terms", level=1)
table(
    ["Parameter", "Recommendation", "Rationale"],
    [
        ["Target size", "USD 1.5bn (hard cap USD 2.0bn)",
         "Sized to the USD 2–15bn transatlantic mid-cap band: large enough for 6–8 "
         "influential stakes, small enough to stay in the under-fished segment and exit "
         "without market impact."],
        ["Structure", "Luxembourg RAIF (AIFMD), evergreen with 3-year soft lock, "
         "quarterly liquidity thereafter with 25% investor-level gates",
         "Engagement horizons of 2–4 years require patient capital; evergreen structure "
         "matches Cevian-model precedent and LP expectations for the asset class."],
        ["GP commitment", "EUR 150–200m (balance sheet + partners/employees)",
         "Anchors alignment; signals conviction to LPs and to target boards."],
        ["Management fee", "1.25–1.50% on NAV", "Mid-cap engaged-equity market standard."],
        ["Performance fee", "17.5% over the higher of a 50/50 S&P 500 / MSCI Europe "
         "blend + 200bps and a 5% absolute hurdle; 3-year crystallisation; high-water mark",
         "Pays only for engagement alpha, not beta; long crystallisation matches the "
         "campaign cycle."],
        ["Concentration limits", "Max 25% NAV per position at cost; max 35% at market; "
         "max 2 concurrent public-escalation campaigns",
         "Concentration is the strategy, but single-name risk is capped."],
        ["Side capacity", "Co-investment sleeves for oversized situations and "
         "take-private bridges",
         "Lets the fund punch above its weight on the best ideas without breaching limits."],
        ["Target returns", "15%+ net IRR over a cycle; expected realised volatility "
         "below index due to engagement catalysts and entry discounts",
         "Consistent with the realised economics of the leading Nordic engagement "
         "practitioner over two decades."],
    ],
    widths=[1.3, 2.6, 2.9],
)

doc.add_heading("8.1 Market exposure: hedged or unhedged?", level=2)
p("A deliberate design choice with real economics attached. Three models were considered:")
table(
    ["Model", "Mechanics", "Assessment"],
    [
        ["Long-only, unhedged (Cevian model)",
         "Net exposure ~100%; returns = market beta + engagement alpha.",
         "Simple, cheap, matches LP expectations for engaged equity. Full drawdown "
         "participation in risk-off years; campaign mid-life is exposed."],
        ["Long with tactical beta overlay (recommended)",
         "Net exposure 80–100% by default; IC may hedge up to ~30% of NAV notionally via "
         "index futures/options under defined stress triggers; FX systematically hedged "
         "to USD.",
         "Keeps the equity risk premium and the catalyst optionality, but caps tail "
         "drawdowns cheaply at portfolio level — practical for a 6–8 name book where "
         "idiosyncratic risk dominates anyway."],
        ["Market-neutral activist (hedge-fund model)",
         "Permanent index/sector shorts isolate campaign alpha; net exposure ~0–30%.",
         "Rejected: 150–300bps annual hedge drag compounds against multi-year campaign "
         "horizons; sector shorts in concentrated books add basis risk, not protection; "
         "and a structural short book sits awkwardly with a constructive, board-seat "
         "brand."],
    ],
    widths=[1.7, 2.6, 2.5],
)
bullets([
    ("Recommendation: ", "run the book predominantly unhedged (net 80–100%), with a "
     "tactical index-hedge overlay at Investment Committee discretion — pre-defined "
     "triggers (drawdown thresholds, spread of book value-gap vs market), capped at ~30% "
     "of NAV, implemented in liquid index futures and options only."),
    ("Never single-name shorts. ", "No shorting of engagement targets or their peers — "
     "the conflict, disclosure and reputational costs outweigh any hedging value; this "
     "is also an explicit LP-facing commitment (Section 4.4)."),
    ("Currency: ", "USD base; SEK/EUR/GBP/CHF/DKK exposures systematically hedged at the "
     "share-class level so LP returns reflect engagement outcomes, not FX."),
    ("Why this is the right trade: ", "the strategy's alpha is idiosyncratic and "
     "catalyst-driven; entry discounts of 30%+ are the primary downside protection. "
     "Paying a permanent hedge premium to neutralise the beta LPs are deliberately "
     "buying would lower net returns ~2–3% p.a. for protection the evergreen structure "
     "(3-year lock, gates) already provides against forced selling."),
])

# ---------- 9. Team ----------
doc.add_heading("9. Team & Organisation", level=1)
p("A deliberately small, senior team of 11–13 at launch, split between Stockholm and New "
  "York, structured for PE-grade underwriting and board-level engagement in both regimes:")
table(
    ["Role", "#", "Profile"],
    [
        ["Head of Engaged Equities (Partner)", "1",
         "Senior hire with 15+ years in engaged/activist investing (Cevian, European "
         "activist franchise, or top-tier event-driven background) plus board credibility. "
         "Sits on Nordic Capital's Executive Committee; co-PM with the Deputy. Based "
         "Stockholm."],
        ["Deputy PM / Head of Research", "1",
         "Internal transfer or hire; owns underwriting standards and the red-team "
         "process."],
        ["Principal — US Situations & Campaign Strategy: Knut Nyman", "1",
         "Founding hire from Elliott Investment Management. Owns the US side of the book: "
         "idea generation and underwriting for US names, the universal-proxy campaign "
         "playbook, and the US advisor network. Based New York. See 9.1 below."],
        ["Sector Principals", "2",
         "Healthcare/MedTech and Industrials/Technology; principal-level, each capable "
         "of leading a campaign end-to-end. One based in each region."],
        ["Investment professionals", "4–5",
         "Associate/VP level; modelling, register analytics, peer benchmarking. Mix of "
         "PE, equity research and special-situations backgrounds; split across offices."],
        ["Head of Stewardship & Governance", "1",
         "Drives nomination-committee work in Europe, director sourcing from the "
         "operating-advisor bench, proxy strategy and engagement documentation."],
        ["Dedicated Compliance Officer", "1",
         "Owns the information barrier with the PE business; reports to Group GC, not to "
         "the PM."],
        ["COO/IR (shared with platform)", "0.5",
         "Fund operations, LP reporting, co-investment execution."],
    ],
    widths=[2.1, 0.4, 4.3],
)
doc.add_heading("9.1 Knut Nyman — Principal, US Situations & Campaign Strategy", level=2)
p("The US half of the mandate requires someone who has run the modern US activism "
  "playbook from inside the most effective practitioner of it. Knut Nyman joins from "
  "Elliott Investment Management as a founding member of the investment team, with the "
  "following mandate:")
bullets([
    ("US idea generation and underwriting. ", "Owns the US screen (Section 6) and "
     "presents US situations to the Investment Committee; leads diligence on US names "
     "alongside the sector principals."),
    ("Campaign strategy. ", "Designs and runs the US engagement mechanics: 13D strategy "
     "and disclosure sequencing, universal-proxy slate construction, settlement "
     "negotiation, and coordination of proxy solicitors, banks, litigation counsel and "
     "communications advisors."),
    ("Bridge between regimes. ", "Translates the firm's constructive European model into "
     "US board dynamics — and brings US-style campaign rigour (precision materials, "
     "shareholder-base analytics, settlement optionality) to the European book."),
    ("Progression. ", "Co-leads one to two campaigns per year from launch, with campaign "
     "lead responsibility from the second fund year and a defined path to Partner; "
     "participates in management-company carry from day one."),
])
doc.add_heading("9.2 Governance & conflicts framework (the make-or-break design issue)", level=2)
bullets([
    ("Information barrier: ", "the public-equities team sits outside the PE deal-flow "
     "perimeter; no access to live PE process information; restricted-list screening before "
     "any accumulation; wall-crossing only via documented GC approval."),
    ("Take-private protocol: ", "if a portfolio position becomes a buyout-fund target, the "
     "public fund's participation/exit is governed by a pre-agreed, LPAC-disclosed protocol "
     "(independent pricing, optional roll or cash-out, conflicted-party abstentions)."),
    ("Reputation committee: ", "any public escalation, media engagement or EGM requisition "
     "requires sign-off from a committee including the Managing Partner — the firm's "
     "30-year brand is the strategy's key asset and is governed accordingly."),
    ("Investment Committee: ", "Head of Engaged Equities, Deputy PM, Principal — US "
     "Situations (for US names), one PE Partner (conflict-cleared per situation), "
     "Managing Partner ex-officio."),
])
doc.add_page_break()

# ---------- 10. Economics & roadmap ----------
doc.add_heading("10. Economics for the Platform", level=1)
bullets([
    "At USD 1.5bn and 1.375% blended management fee: ~USD 20m annual fee revenue against a "
    "~USD 12–14m steady-state cost base — accretive from year two even before performance "
    "fees.",
    "At 15% gross portfolio return, performance fees of ~USD 30–35m per crystallisation "
    "cycle at target size.",
    "Strategic P&L is larger than the direct P&L: one incremental take-private sourced "
    "through the public strategy, or one carve-out acquired by the flagship funds, likely "
    "exceeds the vehicle's annual fee income in value to the platform.",
])
doc.add_heading("11. Risks & Mitigants", level=1)
table(
    ["Risk", "Mitigant"],
    [
        ["Reputational spillover to the PE franchise from a contentious campaign",
         "Constructive-first doctrine; reputation committee veto on escalation; no hostile "
         "opening moves; Cevian's record shows the model rarely requires public conflict "
         "in the Nordics."],
        ["Conflicts / MNPI between public and private strategies",
         "Hard information barrier, dedicated compliance officer, restricted lists, "
         "LPAC-disclosed take-private protocol (Section 9.1)."],
        ["Sphere control blocks change (dual-class shares, foundations)",
         "Screen 3 excludes hostile-controller situations ex ante; where spheres are "
         "present, engage them as allies first."],
        ["Liquidity in European mid-caps (SEK/NOK/DKK/EUR lines thinner than US)",
         "Position sizing tied to ADV; US names carry the larger tickets; evergreen "
         "structure with gates; co-invest sleeves absorb size; exits often via strategic "
         "events rather than market sales."],
        ["Key-person dependence on the Head of Engaged Equities",
         "Co-PM structure, institutional IC process, equity in the management company."],
        ["Cycle risk: a Nordic downturn delays re-rating",
         "Entry discounts of 30%+ provide margin of safety; engagement catalysts are "
         "largely cycle-independent (separations, governance, capital allocation)."],
        ["Talent competition (Active Value Partners et al. hiring the same bench)",
         "Move decisively in 2026; Nordic Capital's platform economics and brand beat "
         "start-up offers for the right principal-level talent."],
    ],
    widths=[3.0, 3.8],
)
doc.add_heading("12. Implementation Roadmap", level=1)
table(
    ["Phase", "Timing", "Milestones"],
    [
        ["0. Decision & design", "Jul–Sep 2026",
         "Executive Committee approval; conflicts framework drafted with GC; regulatory "
         "scoping (Lux RAIF / AIFM passport); search mandate for Head of Engaged Equities."],
        ["1. Build", "Oct 2026 – Mar 2027",
         "Head + Deputy hired; Knut Nyman and the New York pod on board; core team of 7; "
         "seed commitment EUR 150–200m closed; pipeline of 12–16 underwritten names; LP "
         "pre-marketing under existing relationships."],
        ["2. First close & deploy", "Apr–Sep 2027",
         "First close ~USD 750m; 2–3 positions accumulated against the board calendar: "
         "Q3 2027 ownership snapshots for 2028 Nordic nomination committees, and US "
         "nomination windows for the 2028 proxy season."],
        ["3. Scale", "Oct 2027 – Dec 2028",
         "Final close USD 1.5–2.0bn; 6–8 positions across both regions; first "
         "nomination-committee seats, settlements and director appointments in the 2028 "
         "board cycles; first realisation events 2029."],
    ],
    widths=[1.5, 1.4, 3.9],
)

doc.add_page_break()

# ---------- Appendix A: precedents ----------
doc.add_heading("Appendix A: Precedent Campaigns — The Evidence Base", level=1)
p("Three bodies of precedent inform the design of this strategy: Cevian's two-decade "
  "record (the constructive model works, and its failure modes are identifiable), EQT "
  "Public Value (how a PE sponsor's public-markets entry fails without engagement "
  "mechanics), and KKR–Henry Schein (sponsor activism validated in the US, in our core "
  "sector). Outcomes and return indications below are based on public reporting and are "
  "approximate.")

doc.add_heading("A.1 Cevian Capital — the constructive model at scale", level=2)
p("Founded 2002 (Gardell/Förberg, early backing from Carl Icahn); peak AUM ~EUR 15bn+; "
  "partners currently on nine boards in six countries; reported net returns in the "
  "low-to-mid teens annualised over the fund's life. Its disclosed campaign record:")
table(
    ["Company", "Period", "Campaign & outcome", "Result"],
    [
        ["Skandia (SE)", "2002–06",
         "Governance overhaul, pushed sale; acquired by Old Mutual 2006.", "Win"],
        ["Volvo Group (SE)", "2006–17",
         "Portfolio focus (Aero disposal), cost discipline, governance; 8.2% stake sold "
         "to Geely 2017 for ~USD 3.9bn — reported among its most profitable positions.",
         "Win"],
        ["Tieto (FI)", "2009–c.17",
         "Board seat, restructuring and focus; steady re-rating plus dividends.",
         "Modest win"],
        ["Danske Bank (DK)", "2011–c.15",
         "Post-crisis efficiency and capital agenda; stock recovered strongly into exit.",
         "Win"],
        ["Cookson / Vesuvius & Alent (UK)", "2011–c.16",
         "Drove the 2012 demerger; Alent acquired by Platform Specialty 2015 at a "
         "premium; Vesuvius board seat.", "Win"],
        ["Bilfinger (DE)", "2011–c.20",
         "Services transformation; instead repeated profit warnings and CEO churn; "
         "roughly a decade for little value.", "Failure"],
        ["ThyssenKrupp (DE)", "2013–c.22",
         "18% stake; conglomerate break-up thesis. Elevator sale (EUR 17.2bn, 2020) "
         "vindicated the SOTP, but balance-sheet fragility and stakeholder politics "
         "destroyed equity value en route; exited at a significant reported loss.",
         "Failure"],
        ["RSA Insurance (UK)", "2013–21",
         "Focus, disposals, capital discipline, then sale advocacy; Intact/Tryg takeover "
         "2021 at ~50% premium.", "Win"],
        ["Metso (FI)", "2013–20",
         "Supported Valmet demerger, then Metso Outotec merger and Neles separation.",
         "Win"],
        ["ABB (CH/SE)", "2015–25",
         "Power Grids exit (sold to Hitachi 2020), decentralisation, buybacks, leadership "
         "change; long-held position finally sold down in 2025 after a strong multi-year "
         "re-rating.", "Win (slow)"],
        ["Wolseley / Ferguson (UK→US)", "2016–c.21",
         "US focus, name change, listing migration to NYSE; strong re-rating.", "Win"],
        ["Autoliv (SE)", "2017–c.21",
         "Supported the Veoneer spin (2018); Veoneer acquired 2021 at a large premium.",
         "Win"],
        ["Ericsson (SE)", "2017–c.23",
         "Chair change (Leten), cost and portfolio focus; stock roughly doubled into "
         "2021, then compliance issues gave much back; reported net positive but bumpy.",
         "Mixed+"],
        ["Panalpina (CH)", "2018–19",
         "~12%; pushed sale; DSV all-share takeover 2019 at a substantial premium within "
         "~18 months.", "Win (fast)"],
        ["Nordea (FI/SE)", "2019–c.24",
         "Cost discipline and capital returns; large buybacks and re-rating followed.",
         "Win"],
        ["Pearson (UK)", "2020–",
         "Backed digital turnaround; re-rated under new CEO; Apollo approaches (2022) "
         "rejected above-market; Cevian has since lifted its stake to ~18% — its largest "
         "disclosed percentage holding.", "Win, ongoing"],
        ["Vodafone (UK)", "2021–c.23",
         "Consolidation and portfolio agenda; change came too slowly; reported exit "
         "around flat-to-negative.", "Failure"],
        ["Aviva (UK)", "2021–c.24",
         "Demanded ~GBP 5bn capital return and cost cuts; delivered, stock re-rated.",
         "Win"],
        ["UBS (CH)", "2023–25",
         "~1.3% re-rating bet after the Credit Suisse rescue; stock up sharply; reported "
         "~2x on early tranches at 2025 sales.", "Win"],
        ["Baloise (CH)", "2023–25",
         "9.4%; strategy reset and removal of vote caps; Helvetia merger announced April "
         "2025 (completed December 2025); Cevian sold to Patria at ~CHF 184 vs ~CHF 130 "
         "entry (~+40%).", "Win"],
        ["Smith+Nephew (UK)", "2024–",
         "Entered ~5% in 2024; conviction has grown with the campaign — stake lifted "
         "above 10% by mid-2026 (>USD 1.3bn position). Margin recovery and structural "
         "options including a US listing review; shares up since disclosure.", "Ongoing"],
        ["Akzo Nobel (NL)", "c.2024–",
         "Coatings margin gap vs US peers and consolidation thesis; stake doubled to "
         "~10.2% in late 2025 as part of a portfolio reshape funded by the ABB exit.",
         "Ongoing"],
        ["SIG Group (CH)", "2025–",
         "New position in the Swiss food-packaging maker disclosed late 2025; "
         "de-rated quality franchise, self-help thesis.", "Ongoing"],
    ],
    widths=[1.5, 0.9, 3.6, 0.8],
)
p("Read-across: roughly three clear failures in ~23 disclosed campaigns, with the wins "
  "concentrated where a transaction or separation crystallised value (Skandia, RSA, "
  "Panalpina, Alent, Autoliv/Veoneer, Baloise) — supporting this strategy's M&A-archetype "
  "weighting. The failures share a fingerprint our screens are built to exclude: fragile "
  "balance sheets, stakeholder/political complexity, and catalysts outside shareholders' "
  "control (ThyssenKrupp, Bilfinger, Vodafone).", italic=True)

doc.add_heading("A.2 EQT Public Value — the cautionary precedent", level=2)
bullets([
    ("What it was: ", "EQT's attempt to apply its toolbox to listed Nordic mid-caps, "
     "launched 2018–19. Disclosed holdings included Securitas, BHG Group, Storebrand, "
     "BioGaia, AFRY and Storytel — where it built ~10% in September 2021 and became the "
     "largest shareholder."),
    ("What happened: ", "Fundraising never reached scale, the strategy lacked a "
     "board-seat-driven engagement playbook and senior activist leadership, and the fund "
     "was ultimately liquidated — its entire Storytel holding sold in August 2024, after "
     "the 2021-vintage growth entry had been crushed in the 2022 de-rating."),
    ("Lessons designed into this proposal: ", "scale and seed from day one (EUR 150–200m "
     "GP commitment, USD 1.5bn target); a partner-level activist hire and US campaign "
     "leadership rather than PE generalists; stakes sized for influence (nomination "
     "committees / universal proxy), not passive minority positions; entry discipline "
     "anchored on ≥30% upside to intrinsic value, not growth momentum; and an evergreen "
     "structure so positions are never liquidated on a fund clock at the bottom."),
])

doc.add_heading("A.3 Sponsors entering public-market activism", level=2)
p("Governance commentators now track “occasional activists” — sponsors and strategics "
  "using activist tactics — as a defining trend of the 2025–26 cycle. The pattern: a "
  "modest minority stake plus operating credibility converts into board influence within "
  "months, often alongside or after a conventional activist.")
bullets([
    ("KKR – Henry Schein (2025) — the template. ", "Ananym Capital ran a conventional "
     "campaign in late 2024 (board refresh, cost cuts, succession, portfolio review). In "
     "January 2025 Henry Schein announced a USD 250m strategic investment from KKR with "
     "two board seats (Dan Daniel, ex-Danaher; Max Lin, KKR Healthcare); completed May "
     "2025 with KKR at ~12% — the largest non-index holder, with clearance to build to "
     "14.9%. Sponsor activism, in our core healthcare sector, with the activist publicly "
     "supportive."),
    ("KKR – US Foods (2020). ", "USD 500m PIPE at the COVID trough with a board seat; "
     "patient sponsor capital entered alongside activist pressure (Sachem Head later won "
     "seats) and exited into the recovery with strong reported returns."),
    ("Silver Lake – Expedia / Twitter / Airbnb (2020). ", "Crisis PIPEs with board "
     "representation — demonstrating sponsors can underwrite public minority positions "
     "at speed and govern through the board rather than control."),
    ("Apollo – Western Digital (2023). ", "USD 900m convertible preferred while Elliott "
     "publicly pushed the flash/HDD separation; the split completed in 2025 (SanDisk "
     "spin). Sponsor structured capital riding an activist catalyst."),
    ("Warburg Pincus – ESR Group (2022–25). ", "Long-held minority stake in the listed "
     "Asian logistics platform converted into co-leading the ~USD 7bn take-private — "
     "the toehold-to-buyout path this strategy's take-private protocol formalises."),
    ("Read-across: ", "operating credibility plus board seats is the winning currency of "
     "the universal-proxy era, sponsors are arriving now, and the toehold-to-influence "
     "(and occasionally to-control) path is repeatable. Nordic Capital would be early "
     "among European sponsors — but not first in the world, which de-risks the LP "
     "conversation."),
])

doc.add_heading("A.4 Other relevant signals", level=2)
bullets([
    ("Triton – Caverion (FI, 2022–23): ", "sponsor stake-building in a listed Nordic "
     "services company escalated into a competing tender battle (Bain consortium vs "
     "Triton), completed 2023 at a large premium to the undisturbed price — evidence "
     "that public stakes create take-private optionality at premium outcomes."),
    ("Active Value Partners (2026): ", "ex-Cevian partners launched a Nordic/European "
     "engagement strategy, confirming institutional LP appetite — and starting the clock "
     "on the competitive window this proposal addresses."),
])
doc.add_page_break()

doc.add_heading("Appendix B: Disclaimer", level=1)
p("This document is illustrative discussion material prepared for internal strategy "
  "purposes. All references to listed companies are based exclusively on publicly available "
  "information and do not constitute research, investment advice, or a recommendation or "
  "solicitation to buy or sell any security. Figures marked as estimates are directional "
  "and unaudited. Any launch of the strategy described would be subject to legal, "
  "regulatory, tax and conflicts review.", italic=True, size=9, color=GREY)

doc.save("/home/user/knutnyman/nordic-capital-activism-pitch/Nordic_Capital_Engaged_Equities_Pitch.docx")
print("Report saved.")
