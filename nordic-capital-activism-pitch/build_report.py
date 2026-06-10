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
p("A Constructive Activism Strategy for the Nordic Public Markets",
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
  "vehicle pursuing constructive activism in listed Nordic mid-caps. The strategy applies "
  "the firm's private equity value-creation playbook — operational improvement, portfolio "
  "simplification, disciplined capital allocation and strategic M&A — to a concentrated book "
  "of 6–8 public companies where Nordic Capital takes 4–10% ownership stakes, secures "
  "nomination-committee and board influence, and drives a 2–4 year value plan.")
bullets([
    ("The market gap is real. ", "Cevian Capital has demonstrated for two decades that "
     "board-seat-driven engagement in Nordic and European equities generates durable alpha, "
     "yet dedicated Nordic activist capital remains scarce relative to the opportunity. The "
     "2026 launch of Active Value Partners by former Cevian partners confirms growing "
     "institutional appetite for the space — and the window to establish the category-defining "
     "domestic franchise is open now."),
    ("Activism is at record levels and increasingly M&A-driven. ", "2025 was a record year "
     "for global shareholder campaigns, and an unusually high share of recent campaigns carry "
     "M&A demands — terrain where a private equity sponsor has structural advantages no "
     "hedge-fund activist can match."),
    ("Nordic governance is structurally activist-friendly. ", "Shareholder-elected nomination "
     "committees, low EGM thresholds (10% in Sweden), exceptional ownership transparency and a "
     "consensus-driven board culture mean a 5% holder can achieve influence that requires 15%+ "
     "and a proxy fight elsewhere."),
    ("Nordic Capital is uniquely positioned. ", "Three decades of operating credibility, deep "
     "sector franchises in healthcare, technology & payments, financial services and "
     "industrials, and an unrivalled network of Nordic chairs, CEOs and operating advisors. "
     "Critically, Nordic Capital will be received as a constructive domestic owner — not a "
     "foreign raider — which materially lowers the cost and friction of every campaign."),
    ("Strategic synergies compound the core P&L. ", "The public strategy becomes a pipeline "
     "for take-privates and structured PIPEs, a catalyst for carve-outs the buyout funds can "
     "acquire, and an extension of LP relationships into a new product line — diversifying "
     "fee streams beyond the flagship buyout franchise."),
])
p("We recommend a phased launch: internal approval and conflicts framework in H2 2026, a "
  "founding team of 10–12 led by a partner-level Head of Engaged Equities, a EUR 150–200m "
  "GP/employee seed commitment, first close of ~USD 750m in mid-2027 and stake-building "
  "timed to the Q3 ownership snapshots that determine 2028 AGM nomination committees. "
  "Target net returns of 15%+ over a cycle, with management economics accretive to the "
  "platform from year two.")
doc.add_page_break()

# ---------- 2. Why now ----------
doc.add_heading("2. The Opportunity: Why Public-Market Activism, Why Now", level=1)
doc.add_heading("2.1 The valuation gap has become a strategic opening", level=2)
p("European equities continue to trade at a persistent discount to US peers, and Nordic "
  "mid-caps — under-covered after a decade of sell-side consolidation and MiFID II research "
  "unbundling — trade at a further discount to European large-caps despite superior "
  "governance and profitability. A meaningful cohort of quality Nordic franchises trades at "
  "30–50% discounts to intrinsic value on sum-of-the-parts or normalised-margin bases. "
  "Public-market multiples for Nordic mid-caps sit well below the entry multiples Nordic "
  "Capital's buyout funds routinely pay in competitive private processes — in effect, the "
  "public market is offering the firm's core hunting ground at a discount, without an "
  "auction.")
doc.add_heading("2.2 Activism has gone mainstream — and M&A-centric", level=2)
bullets([
    "2025 set a record for global activist campaigns, with activity broadening across "
    "geographies, market caps and campaign types.",
    "In late 2025, roughly six in ten campaigns carried an M&A thesis (break-ups, "
    "divestitures, sale processes) — the highest share in five years. M&A-driven activism "
    "plays directly to a private equity sponsor's strengths: valuing businesses, running "
    "separations, and standing ready as a buyer or financing partner.",
    "Universal owners (index funds, Nordic institutions, AP funds) increasingly support "
    "well-argued engagement agendas, lowering the vote-getting cost of campaigns.",
    "Dedicated Nordic engagement capital remains thin: Cevian (pan-European, increasingly "
    "large-cap), a handful of small Nordic engagement funds, and episodic visits from US "
    "activists with limited local standing. The mid-cap segment (EUR 1–10bn) is "
    "structurally underserved.",
])
doc.add_heading("2.3 Why the Nordics are the best activism jurisdiction in the world", level=2)
p("The Nordic governance model was effectively designed for engaged ownership:")
table(
    ["Feature", "Mechanism", "Implication for the strategy"],
    [
        ["Nomination committees",
         "Board nominations are controlled by committees composed of the largest "
         "shareholders (typically top 3–4 as of Aug/Sep ownership records), not by the board "
         "itself.",
         "A 4–8% stake in a mid-cap routinely earns a nomination-committee seat — direct "
         "influence over board composition without a proxy fight."],
        ["Low EGM threshold",
         "10% of shares can requisition an extraordinary general meeting in Sweden (similar "
         "thresholds across the region).",
         "Credible escalation path that rarely needs to be used."],
        ["Ownership transparency",
         "Public shareholder registers (e.g., Euroclear Sweden) updated frequently.",
         "Precise coalition mapping; no guessing who holds the register."],
        ["Consensus culture",
         "Boards and ownership spheres respond to well-researched private argumentation; "
         "public hostility is rare and usually unnecessary.",
         "Lower campaign cost, faster outcomes, brand-compatible conduct."],
        ["Dual-class shares & spheres",
         "Founding families, foundations and investment companies (Wallenberg/Investor, "
         "Industrivärden, Lundberg, Maersk family, Danish foundations) anchor many issuers.",
         "Screens out uninvestable names early; conversely, sphere-free mid-caps are "
         "exceptionally open to engagement."],
    ],
    widths=[1.5, 2.7, 2.6],
)
p("Cevian's two-decade record — board seats across the region, low-to-mid-teens net returns "
  "over cycles, and outcomes such as the value created around Danske Bank, ABB, Ericsson and "
  "numerous Nordic industrials — is the proof of concept. The constructive, board-led model "
  "works in this region better than anywhere else. What the region lacks is a second "
  "institutional-scale practitioner. Nordic Capital can be that firm, with advantages Cevian "
  "never had: a buyout balance sheet behind the engagement, and operating resources measured "
  "in hundreds of professionals.")
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
    ("Domestic legitimacy. ", "Nordic institutions, family spheres and governments engage "
     "very differently with a Stockholm-rooted owner of 30 years than with a New York or "
     "London activist. Lower friction means cheaper, faster, quieter wins — and access to "
     "situations foreign activists cannot touch."),
    ("Sector depth = underwriting edge. ", "The firm's diligence pattern-recognition in "
     "medtech, software, payments and specialty financials applies directly to the listed "
     "mid-cap universe, much of which comprises businesses Nordic Capital has studied, "
     "bid on, or competed against."),
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
p("Acquire 4–10% positions in 6–8 listed Nordic companies with EUR 1–10bn market "
  "capitalisation, where a clearly identifiable gap exists between market price and "
  "intrinsic value, and where that gap can be closed within 2–4 years through actions the "
  "company itself can take: operational improvement, portfolio separation or divestiture, "
  "capital-allocation reform, or strategic transactions. Influence is exercised through the "
  "Nordic governance machinery — nomination-committee participation, board representation, "
  "and private engagement with management, boards and anchor owners — with public pressure "
  "reserved as an escalation, not a default.", )
doc.add_heading("4.2 The engagement playbook", level=2)
table(
    ["Phase", "Months", "Actions"],
    [
        ["1. Underwrite", "0–4",
         "PE-grade diligence: full operational benchmark, SOTP, downside case. Thesis must "
         "survive an internal “red team” and clear a 30%+ upside to base-case "
         "intrinsic value with limited structural downside."],
        ["2. Accumulate", "3–9",
         "Build 4–10% via open market and blocks, timed against Q3 ownership snapshots that "
         "seat nomination committees. Disclose at thresholds; first private meeting with "
         "chair before crossing 5%."],
        ["3. Engage", "6–18",
         "Present value plan privately to board and anchor shareholders. Secure "
         "nomination-committee seat at the next cycle; propose 1–2 directors from the "
         "operating network. Align with domestic institutions and AP funds."],
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
p("Sector focus mirrors the firm's private equity franchises, weighted to where the listed "
  "Nordic mid-cap universe offers depth:")
table(
    ["Sector", "Indicative weight", "Why", "Illustrative listed universe"],
    [
        ["Healthcare & MedTech", "30–40%",
         "Nordic Capital's deepest franchise; the region hosts a dense cluster of "
         "global-niche medtech and pharma-services leaders, several in margin or "
         "execution repair.",
         "Elekta, Getinge, Demant, GN Store Nord, Ambu, Coloplast-adjacent suppliers"],
        ["Industrials & Business Services", "25–35%",
         "Largest segment of the listed universe; rich in conglomerate discounts and "
         "self-help stories.",
         "Orkla, SKF, Dometic, Thule, ISS, Securitas, Valmet, Huhtamäki"],
        ["Technology, Software & Payments", "15–25%",
         "Post-2021 de-ratings left quality Nordic software and IT-services assets at "
         "private-market discounts; consolidation logic is strong.",
         "Tieto, Crayon-adjacent, Nexi-linked Nordic payment assets, Kahoot-class "
         "de-rated SaaS"],
        ["Financial Services", "10–20%",
         "Specialty finance, insurance and savings platforms with capital-return and "
         "consolidation angles.",
         "Storebrand, Sampo-adjacent, specialty lenders"],
    ],
    widths=[1.5, 1.0, 2.4, 1.9],
)
p("Geographic spread: Sweden ~40–50% (deepest market, most activist-friendly mechanics), "
  "Denmark ~20–25%, Finland ~15–20%, Norway ~10–15% (sphere- and state-heavy, more "
  "selective).")

# ---------- 6. Target universe ----------
doc.add_heading("6. Target Universe & Screening", level=1)
bullets([
    ("Universe: ", "~1,400 Nordic-listed companies; ~150 in the EUR 1–10bn sweet spot "
     "after liquidity screens."),
    ("Screen 1 — Quality: ", "defensible market position, structurally sound gross "
     "margins, identifiable best-in-class peer set. (~60 names survive.)"),
    ("Screen 2 — Value gap: ", ">30% upside to intrinsic value on normalised margins or "
     "SOTP. (~25–30 names.)"),
    ("Screen 3 — Path to influence: ", "free-float and register structure permitting a "
     "4–10% stake with nomination-committee relevance; no blocking sphere hostile to "
     "change. (~15–20 names.)"),
    ("Screen 4 — Conflict check: ", "no MNPI contamination from current PE processes; "
     "cleared by compliance before any accumulation. (Live pipeline: 10–14 names, of "
     "which 6–8 funded.)"),
])
p("A EUR 1.5bn book across 6–8 core positions implies EUR 180–280m per position — i.e., "
  "4–10% of companies with EUR 2–6bn market caps, precisely the band where nomination-"
  "committee seats are won and where dedicated competition is thinnest.")
doc.add_page_break()

# ---------- 7. Illustrative pipeline ----------
doc.add_heading("7. Illustrative Pipeline (Public Information Only)", level=1)
p("The following five situations illustrate the strategy's archetypes using publicly "
  "available information as of early/mid 2026. They are illustrations for discussion, not "
  "recommendations, and would each require full underwriting and conflict clearance.",
  italic=True)

doc.add_heading("7.1 Elekta (Sweden — MedTech) — Margin-gap operator", level=2)
bullets([
    "Global #2 in radiation oncology, a structurally growing oligopoly, yet the stock has "
    "de-rated dramatically; third-party fair-value estimates have placed the discount at "
    "~50%+ at points in 2026.",
    "FY2025/26 adjusted EBIT margin of ~12% versus mid-to-high-teens at Siemens "
    "Healthineers' Varian — a 400–600bps gap attributable to supply chain, order execution "
    "and mix, not market position.",
    "Value plan: board-sponsored margin program, order-to-revenue execution reset, "
    "service-revenue acceleration; strategic optionality given consolidation logic in "
    "medtech. Healthcare is Nordic Capital's home turf — credibility here is immediate.",
])
doc.add_heading("7.2 Getinge (Sweden — MedTech) — Margin recovery & portfolio focus", level=2)
bullets([
    "Profitability is recovering (profit margin ~6.5% in 2025 vs ~4.7% in 2024) as "
    "quality-remediation costs roll off, but remains far below medtech peer levels.",
    "Three loosely related business areas (Acute Care Therapies, Life Science, Surgical "
    "Workflows) invite portfolio review; a Life Science separation is a credible SOTP "
    "catalyst.",
    "Value plan: complete the quality agenda with board-level accountability, then drive "
    "portfolio focus and margin normalisation toward double-digit EBITA.",
])
doc.add_heading("7.3 GN Store Nord (Denmark — Hearing & Audio) — Separation thesis", level=2)
bullets([
    "Long-standing two-business structure (hearing care and enterprise/consumer audio) with "
    "limited synergy; separation has been publicly debated for years while leverage from "
    "the Steelseries era suppressed action.",
    "Value plan: deleverage milestones, then formal separation review; both businesses have "
    "natural strategic and sponsor buyers. Classic M&A-activism archetype in a "
    "foundation-light register that is unusually open for a Danish large-cap.",
])
doc.add_heading("7.4 Tieto (Finland — Software & IT Services) — Completing the break-up", level=2)
bullets([
    "Tietoevry has already validated the separation path: Banking demerger toward a "
    "Nasdaq Helsinki listing announced, Tech Services sold to Agilitas for EUR 300m "
    "(closed September 2025), proceeds applied to debt.",
    "Value plan: support and accelerate completion of the value-realisation program, "
    "enforce capital-return discipline on proceeds, and position the remaining Nordic "
    "software core (Industry/Care) for consolidation — as buyer or seller.",
])
doc.add_heading("7.5 Orkla (Norway — Branded Consumer / Holding) — Conglomerate discount", level=2)
bullets([
    "Self-declared transition to an “industrial investment company” with ~12 "
    "portfolio companies, a ~42.6% stake in Jotun (paints — a world-class hidden asset), "
    "hydropower assets and net cash optionality; the equity persistently trades below "
    "credible SOTP.",
    "Value plan: accelerate announced portfolio monetisations, establish a capital-return "
    "framework tied to disposal proceeds, and push for transparent per-asset reporting that "
    "forces the SOTP into the price.",
])
p("Watchlist (further candidates surviving Screens 1–3): Securitas (self-help largely "
  "delivered — instructive as a model of the re-rating we underwrite), Demant, Ambu, "
  "Dometic, Thule, Huhtamäki, Valmet, Storebrand, ISS.", italic=True)
doc.add_page_break()

# ---------- 8. Fund structure ----------
doc.add_heading("8. Vehicle, Size & Terms", level=1)
table(
    ["Parameter", "Recommendation", "Rationale"],
    [
        ["Target size", "USD 1.5bn (hard cap USD 2.0bn)",
         "Sized to the EUR 1–10bn mid-cap band: large enough for 6–8 influential stakes, "
         "small enough to stay in the under-fished segment and exit without market impact."],
        ["Structure", "Luxembourg RAIF (AIFMD), evergreen with 3-year soft lock, "
         "quarterly liquidity thereafter with 25% investor-level gates",
         "Engagement horizons of 2–4 years require patient capital; evergreen structure "
         "matches Cevian-model precedent and LP expectations for the asset class."],
        ["GP commitment", "EUR 150–200m (balance sheet + partners/employees)",
         "Anchors alignment; signals conviction to LPs and to target boards."],
        ["Management fee", "1.25–1.50% on NAV", "Mid-cap engaged-equity market standard."],
        ["Performance fee", "17.5% over the higher of MSCI Nordic + 200bps and a 5% "
         "absolute hurdle; 3-year crystallisation; high-water mark",
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

# ---------- 9. Team ----------
doc.add_heading("9. Team & Organisation", level=1)
p("A deliberately small, senior team of 10–12 at launch, structured for PE-grade "
  "underwriting and board-level engagement:")
table(
    ["Role", "#", "Profile"],
    [
        ["Head of Engaged Equities (Partner)", "1",
         "Senior hire with 15+ years in engaged/activist investing (Cevian, European "
         "activist franchise, or top-tier event-driven background) plus Nordic board "
         "credibility. Sits on Nordic Capital's Executive Committee; co-PM with the "
         "Deputy."],
        ["Deputy PM / Head of Research", "1",
         "Internal transfer or hire; owns underwriting standards and the red-team "
         "process."],
        ["Sector Principals", "2",
         "Healthcare/MedTech and Industrials/Technology; principal-level, each capable "
         "of leading a campaign end-to-end."],
        ["Investment professionals", "4–5",
         "Associate/VP level; modelling, register analytics, peer benchmarking. Mix of "
         "PE, equity research and special-situations backgrounds."],
        ["Head of Stewardship & Governance", "1",
         "Drives nomination-committee work, director sourcing from the operating-advisor "
         "bench, proxy strategy and engagement documentation."],
        ["Dedicated Compliance Officer", "1",
         "Owns the information barrier with the PE business; reports to Group GC, not to "
         "the PM."],
        ["COO/IR (shared with platform)", "0.5",
         "Fund operations, LP reporting, co-investment execution."],
    ],
    widths=[2.1, 0.4, 4.3],
)
doc.add_heading("9.1 Governance & conflicts framework (the make-or-break design issue)", level=2)
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
    ("Investment Committee: ", "Head of Engaged Equities, Deputy PM, one PE Partner "
     "(conflict-cleared per situation), Managing Partner ex-officio."),
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
        ["Liquidity in SEK/NOK/DKK mid-caps",
         "Position sizing tied to ADV; evergreen structure with gates; co-invest sleeves "
         "absorb size; exits often via strategic events rather than market sales."],
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
         "Head + Deputy hired; core team of 6; seed commitment EUR 150–200m closed; "
         "pipeline of 10–14 underwritten names; LP pre-marketing under existing "
         "relationships."],
        ["2. First close & deploy", "Apr–Sep 2027",
         "First close ~USD 750m; 2–3 positions accumulated against the Q3 2027 ownership "
         "snapshots that seat 2028 nomination committees."],
        ["3. Scale", "Oct 2027 – Dec 2028",
         "Final close USD 1.5–2.0bn; 6–8 positions; first nomination-committee seats and "
         "director appointments at the 2028 AGM season; first realisation events 2029."],
    ],
    widths=[1.5, 1.4, 3.9],
)

doc.add_heading("Appendix: Disclaimer", level=1)
p("This document is illustrative discussion material prepared for internal strategy "
  "purposes. All references to listed companies are based exclusively on publicly available "
  "information and do not constitute research, investment advice, or a recommendation or "
  "solicitation to buy or sell any security. Figures marked as estimates are directional "
  "and unaudited. Any launch of the strategy described would be subject to legal, "
  "regulatory, tax and conflicts review.", italic=True, size=9, color=GREY)

doc.save("/home/user/knutnyman/nordic-capital-activism-pitch/Nordic_Capital_Engaged_Equities_Pitch.docx")
print("Report saved.")
