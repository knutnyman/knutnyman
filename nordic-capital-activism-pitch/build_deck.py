#!/usr/bin/env python3
"""Builds the PowerPoint deck: Nordic Capital Engaged Equities strategy pitch."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

NAVY = RGBColor(0x0B, 0x25, 0x45)
NAVY2 = RGBColor(0x13, 0x3A, 0x63)
GOLD = RGBColor(0xC4, 0x9A, 0x3B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0x55, 0x5F, 0x6E)
LIGHT = RGBColor(0xF2, 0xF4, 0xF8)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def add_rect(slide, x, y, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, runs, size=14, color=GREY, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.0, font="Georgia"):
    """runs: str, or list of (text, dict-of-overrides) paragraphs."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(runs, str):
        runs = [(runs, {})]
    for i, (text, ov) in enumerate(runs):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ov.get("align", align)
        para.line_spacing = ov.get("line_spacing", line_spacing)
        para.space_after = Pt(ov.get("space_after", 4))
        r = para.add_run()
        r.text = text
        f = r.font
        f.name = ov.get("font", font)
        f.size = Pt(ov.get("size", size))
        f.bold = ov.get("bold", bold)
        f.italic = ov.get("italic", False)
        f.color.rgb = ov.get("color", color)
    return tb


def bullet_box(slide, x, y, w, h, items, size=13, gap=8, color=RGBColor(0x2B, 0x33, 0x40)):
    """items: list of str or (bold_lead, rest) tuples."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.space_after = Pt(gap)
        para.line_spacing = 1.05
        b = para.add_run()
        b.text = "▪  "
        b.font.size = Pt(size)
        b.font.color.rgb = GOLD
        b.font.name = "Georgia"
        if isinstance(it, tuple):
            lead, rest = it
            r1 = para.add_run()
            r1.text = lead
            r1.font.bold = True
            r1.font.size = Pt(size)
            r1.font.color.rgb = NAVY
            r1.font.name = "Georgia"
            r2 = para.add_run()
            r2.text = rest
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
            r2.font.name = "Georgia"
        else:
            r = para.add_run()
            r.text = it
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.name = "Georgia"
    return tb


def new_slide(title, kicker=None):
    slide = prs.slides.add_slide(BLANK)
    add_rect(slide, 0, 0, SW, Inches(0.95), NAVY)
    add_rect(slide, 0, Inches(0.95), SW, Pt(3), GOLD)
    if kicker:
        add_text(slide, Inches(0.55), Inches(0.08), Inches(11), Inches(0.3),
                 kicker.upper(), size=10, color=GOLD, bold=True)
        ty = Inches(0.34)
    else:
        ty = Inches(0.18)
    add_text(slide, Inches(0.55), ty, Inches(12.2), Inches(0.6), title,
             size=24, color=WHITE, bold=True)
    add_text(slide, Inches(0.55), Inches(7.12), Inches(12.2), Inches(0.3),
             "Strictly private & confidential — illustrative discussion material",
             size=8, color=RGBColor(0x9A, 0xA4, 0xB2))
    return slide


def card(slide, x, y, w, h, title, body, title_size=13, body_size=11.5):
    add_rect(slide, x, y, w, h, LIGHT)
    add_rect(slide, x, y, Pt(3.5), h, GOLD)
    add_text(slide, x + Inches(0.18), y + Inches(0.10), w - Inches(0.3), Inches(0.45),
             title, size=title_size, color=NAVY, bold=True)
    add_text(slide, x + Inches(0.18), y + Inches(0.52), w - Inches(0.3), h - Inches(0.6),
             body, size=body_size, color=GREY, line_spacing=1.05)


def stat(slide, x, y, w, value, label, value_size=30):
    add_text(slide, x, y, w, Inches(0.55), value, size=value_size, color=GOLD, bold=True,
             align=PP_ALIGN.CENTER)
    add_text(slide, x, y + Inches(0.58), w, Inches(0.65), label, size=11, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.0)


def table_slide(slide, x, y, w, headers, rows, col_widths, row_h=0.5, size=11):
    from pptx.util import Inches as In
    t = slide.shapes.add_table(len(rows) + 1, len(headers), x, y, w,
                               In(row_h * (len(rows) + 1))).table
    for j, cw in enumerate(col_widths):
        t.columns[j].width = In(cw)
    for j, htext in enumerate(headers):
        cell = t.cell(0, j)
        cell.text = htext
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        para = cell.text_frame.paragraphs[0]
        para.runs[0].font.size = Pt(size)
        para.runs[0].font.bold = True
        para.runs[0].font.color.rgb = WHITE
        para.runs[0].font.name = "Georgia"
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i + 1, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT if i % 2 else WHITE
            for para in cell.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(size - 0.5)
                    r.font.color.rgb = RGBColor(0x2B, 0x33, 0x40)
                    r.font.name = "Georgia"
    return t


# ============ 1. Title ============
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, SW, SH, NAVY)
add_rect(slide, 0, Inches(3.62), SW, Pt(3), GOLD)
add_text(slide, Inches(1), Inches(1.45), Inches(11.3), Inches(0.4),
         "STRICTLY PRIVATE & CONFIDENTIAL", size=12, color=GOLD, bold=True,
         align=PP_ALIGN.CENTER)
add_text(slide, Inches(1), Inches(2.0), Inches(11.3), Inches(1.0),
         "Nordic Capital Engaged Equities", size=44, color=WHITE, bold=True,
         align=PP_ALIGN.CENTER)
add_text(slide, Inches(1), Inches(3.0), Inches(11.3), Inches(0.5),
         "A constructive activism strategy for the Nordic public markets",
         size=19, color=RGBColor(0xC9, 0xD3, 0xE0), align=PP_ALIGN.CENTER)
add_text(slide, Inches(1), Inches(4.05), Inches(11.3), Inches(0.4),
         "Proposal for a USD 1.5bn engaged public-equities vehicle",
         size=14, color=GOLD, align=PP_ALIGN.CENTER, bold=True)
add_text(slide, Inches(1), Inches(5.9), Inches(11.3), Inches(0.8),
         [("Prepared for the Managing Partner & Executive Committee", {"align": PP_ALIGN.CENTER, "size": 12, "color": RGBColor(0xC9, 0xD3, 0xE0)}),
          ("June 2026  |  Company references are based on public information only", {"align": PP_ALIGN.CENTER, "size": 10, "color": RGBColor(0x8A, 0x96, 0xA6)})])

# ============ 2. Executive summary ============
slide = new_slide("Executive summary", "The proposal")
bullet_box(slide, Inches(0.55), Inches(1.25), Inches(7.4), Inches(5.6), [
    ("Launch “Nordic Capital Engaged Equities”: ", "a USD 1.5bn (cap USD 2.0bn) "
     "constructive-activism fund taking 4–10% stakes in 6–8 listed Nordic mid-caps "
     "(EUR 1–10bn market cap)."),
    ("Apply the firm's PE playbook in public markets: ", "operational improvement, "
     "portfolio separation, capital-allocation reform — executed via nomination "
     "committees and board seats, not public fights."),
    ("The window is open: ", "record global activism, an M&A-heavy campaign mix that "
     "favours a PE sponsor, persistent 30–50% value gaps in under-covered Nordic "
     "mid-caps — and only one established practitioner (Cevian) in the region."),
    ("Nordic Capital's edge is structural: ", "operating credibility, the region's best "
     "director bench, domestic legitimacy, and sector depth in healthcare, tech & "
     "payments, financial services and industrials."),
    ("Platform synergies compound: ", "take-private and PIPE pipeline for the buyout "
     "funds, carve-out deal flow, and a new fee stream — accretive from year two."),
], size=13.5, gap=12)
add_rect(slide, Inches(8.35), Inches(1.25), Inches(4.4), Inches(5.6), NAVY)
stat(slide, Inches(8.5), Inches(1.55), Inches(4.1), "USD 1.5bn", "Target size (hard cap USD 2.0bn)")
stat(slide, Inches(8.5), Inches(2.85), Inches(4.1), "6–8", "Concentrated core positions, 4–10% stakes")
stat(slide, Inches(8.5), Inches(4.15), Inches(4.1), "15%+", "Target net IRR over a cycle")
stat(slide, Inches(8.5), Inches(5.45), Inches(4.1), "H2 2027", "First close & initial deployment")

# ============ 3. Why now ============
slide = new_slide("Why now: the opening for a second Nordic engagement franchise", "Opportunity")
card(slide, Inches(0.55), Inches(1.3), Inches(6.1), Inches(1.7), "Record activism, M&A-centric",
     "2025 set a record for global activist campaigns; ~6 in 10 late-2025 campaigns carried "
     "M&A demands (break-ups, divestitures, sale processes) — terrain where a private equity "
     "sponsor has advantages no hedge-fund activist can match.")
card(slide, Inches(6.85), Inches(1.3), Inches(6.0), Inches(1.7), "Persistent Nordic value gaps",
     "Nordic mid-caps are structurally under-covered post-MiFID II; quality franchises trade "
     "at 30–50% discounts to intrinsic value on SOTP / normalised margins — below the entry "
     "multiples our buyout funds pay in auctions.")
card(slide, Inches(0.55), Inches(3.2), Inches(6.1), Inches(1.7), "Thin dedicated competition",
     "Cevian is the only institutional-scale practitioner and has migrated toward pan-European "
     "large-caps. The EUR 1–10bn mid-cap band is underserved — and the 2026 launch of Active "
     "Value Partners (ex-Cevian) confirms institutional appetite for the space.")
card(slide, Inches(6.85), Inches(3.2), Inches(6.0), Inches(1.7), "Supportive shareholder base",
     "Nordic institutions, AP funds and universal owners increasingly back well-argued "
     "engagement agendas — lowering the cost of building winning coalitions.")
add_rect(slide, Inches(0.55), Inches(5.15), Inches(12.3), Inches(1.55), NAVY)
add_text(slide, Inches(0.85), Inches(5.38), Inches(11.8), Inches(1.2),
         [("The strategic point", {"size": 12, "color": GOLD, "bold": True}),
          ("The category-defining domestic engagement franchise in the Nordics has not yet been "
           "built. Nordic Capital can build it — with advantages Cevian never had: a buyout "
           "platform behind the engagement and an operating bench measured in hundreds of "
           "professionals.", {"size": 14, "color": WHITE, "line_spacing": 1.1})])

# ============ 4. Nordic governance ============
slide = new_slide("The Nordics: the most activist-friendly governance model in the world", "Opportunity")
table_slide(slide, Inches(0.55), Inches(1.35), Inches(12.3),
            ["Feature", "Mechanism", "What it means for us"],
            [
                ["Nomination committees",
                 "Board nominations controlled by the largest shareholders (top 3–4 as of "
                 "Aug/Sep records) — not by the board",
                 "A 4–8% stake routinely earns direct influence over board composition — no "
                 "proxy fight needed"],
                ["Low EGM threshold",
                 "10% of shares can call an extraordinary general meeting (Sweden; similar "
                 "across the region)",
                 "Credible escalation path that rarely needs to be used"],
                ["Ownership transparency",
                 "Public shareholder registers (e.g. Euroclear Sweden), frequently updated",
                 "Precise coalition mapping before and during campaigns"],
                ["Consensus board culture",
                 "Boards and ownership spheres respond to well-researched private engagement",
                 "Lower campaign cost, faster outcomes, brand-compatible conduct"],
                ["Spheres & dual-class shares",
                 "Families, foundations and investment companies anchor many issuers",
                 "Screened ex ante; sphere-free mid-caps are exceptionally open to engagement"],
            ],
            [2.3, 5.2, 4.8], row_h=0.78, size=11.5)
add_text(slide, Inches(0.55), Inches(5.75), Inches(12.3), Inches(1.0),
         [("Proof of concept: Cevian's two decades of board-seat-driven engagement across the "
           "region — and returns in the low-to-mid teens net over cycles — demonstrate the "
           "model. Timing discipline: stakes must be built before the Q3 ownership snapshots "
           "that seat the following spring's nomination committees.",
           {"size": 12.5, "color": GREY, "italic": True, "line_spacing": 1.15})])

# ============ 5. Why Nordic Capital ============
slide = new_slide("Why Nordic Capital wins", "Right to win")
cards = [
    ("Operating credibility", "30 years and >EUR 31bn AUM built on operational "
     "transformation. When we tell a board the margin gap is 400bps, we speak as an owner "
     "that has closed it repeatedly in private portfolios."),
    ("The director bench", "An unmatched network of Nordic chairs, former CEOs and "
     "operating advisors to propose as board candidates — the resource that decides "
     "outcomes in a nomination-committee system."),
    ("Domestic legitimacy", "Institutions, family spheres and governments engage "
     "differently with a Stockholm-rooted owner than with a New York or London activist. "
     "Lower friction = cheaper, faster, quieter wins."),
    ("Sector underwriting edge", "Healthcare, tech & payments, financial services, "
     "industrials — the listed mid-cap universe is full of businesses we have studied, "
     "bid on or competed against."),
    ("Two-way PE pipeline", "Engagement surfaces take-privates and PIPEs for the flagship "
     "funds; activist-catalysed carve-outs create proprietary deal flow (strict info "
     "barriers; see governance slide)."),
    ("LP demand", "Existing LPs want Nordic exposure with liquidity; Evolution II's "
     "four-month hard-cap close shows the franchise's fundraising velocity."),
]
for i, (t, b) in enumerate(cards):
    x = Inches(0.55 + (i % 3) * 4.18)
    y = Inches(1.3 + (i // 3) * 2.75)
    card(slide, x, y, Inches(3.95), Inches(2.55), t, b, body_size=11.5)

# ============ 6. Strategy ============
slide = new_slide("Investment strategy: constructive activism, board-led", "Strategy")
add_text(slide, Inches(0.55), Inches(1.2), Inches(12.3), Inches(0.85),
         [("Take 4–10% positions in 6–8 listed Nordic companies (EUR 1–10bn market cap) with "
           "≥30% upside to intrinsic value, and close the gap in 2–4 years through actions the "
           "company itself can take — driven from the nomination committee and the boardroom.",
           {"size": 14.5, "color": NAVY, "bold": True, "line_spacing": 1.15})])
steps = [
    ("1. Underwrite", "M 0–4", "PE-grade diligence; operational benchmark; SOTP; red-team; "
     "≥30% upside with limited structural downside"),
    ("2. Accumulate", "M 3–9", "4–10% via market + blocks, timed to Q3 ownership snapshots; "
     "chair meeting before crossing 5%"),
    ("3. Engage", "M 6–18", "Private value plan to board & anchors; nomination-committee "
     "seat; propose 1–2 directors from our bench"),
    ("4. Execute", "M 12–36", "Margin program / separation / balance-sheet reset; quarterly "
     "milestones vs the underwrite"),
    ("5. Realise", "M 24–48", "Re-rating, strategic sale — or conflict-cleared take-private "
     "by our buyout funds"),
]
for i, (t, m, b) in enumerate(steps):
    x = Inches(0.55 + i * 2.52)
    add_rect(slide, x, Inches(2.35), Inches(2.32), Inches(2.9), NAVY if i % 2 == 0 else NAVY2)
    add_text(slide, x + Inches(0.12), Inches(2.5), Inches(2.1), Inches(0.4), t,
             size=13.5, color=GOLD, bold=True)
    add_text(slide, x + Inches(0.12), Inches(2.92), Inches(2.1), Inches(0.3), m,
             size=10, color=RGBColor(0x9A, 0xA4, 0xB2), bold=True)
    add_text(slide, x + Inches(0.12), Inches(3.25), Inches(2.1), Inches(1.9), b,
             size=10.5, color=WHITE, line_spacing=1.1)
bullet_box(slide, Inches(0.55), Inches(5.55), Inches(12.3), Inches(1.4), [
    ("Doctrine: ", "constructive-first — public escalation needs IC + reputation-committee "
     "sign-off; never a hostile opening move; no shorting of engagement targets; max 2 "
     "concurrent public-phase campaigns."),
], size=12.5)

# ============ 7. Thesis archetypes ============
slide = new_slide("Three thesis archetypes", "Strategy")
arch = [
    ("Margin-gap operators", "40–50% of book",
     "Quality franchises running 300–600bps below achievable margins (execution, supply "
     "chain, quality systems — not market position). Value plan: board-sponsored "
     "operational program with hard milestones.", "Examples: Elekta, Getinge"),
    ("Conglomerate / SOTP discounts", "30–40% of book",
     "Multi-division companies where separation, divestiture or monetisation of hidden "
     "assets unlocks 30–60% upside. Plays to the record M&A-activism mix and our carve-out "
     "expertise.", "Examples: GN Store Nord, Orkla"),
    ("Capital allocation & strategic catalysts", "15–25% of book",
     "Lazy balance sheets, value-destructive M&A patterns, or companies that should start "
     "or accept strategic processes; often overlaps the other archetypes.",
     "Example: Tieto (completing its break-up)"),
]
for i, (t, w, b, ex) in enumerate(arch):
    x = Inches(0.55 + i * 4.18)
    add_rect(slide, x, Inches(1.35), Inches(3.95), Inches(4.7), LIGHT)
    add_rect(slide, x, Inches(1.35), Inches(3.95), Pt(4), GOLD)
    add_text(slide, x + Inches(0.2), Inches(1.55), Inches(3.55), Inches(0.75), t,
             size=15.5, color=NAVY, bold=True)
    add_text(slide, x + Inches(0.2), Inches(2.3), Inches(3.55), Inches(0.35), w,
             size=11.5, color=GOLD, bold=True)
    add_text(slide, x + Inches(0.2), Inches(2.7), Inches(3.55), Inches(2.4), b,
             size=12, color=GREY, line_spacing=1.15)
    add_text(slide, x + Inches(0.2), Inches(5.35), Inches(3.55), Inches(0.6), ex,
             size=11, color=NAVY, bold=True)
add_text(slide, Inches(0.55), Inches(6.35), Inches(12.3), Inches(0.5),
         [("Sector focus mirrors the PE franchise: Healthcare & MedTech 30–40% · Industrials & "
           "Business Services 25–35% · Tech, Software & Payments 15–25% · Financial Services "
           "10–20%. Geography: SE ~40–50%, DK ~20–25%, FI ~15–20%, NO ~10–15%.",
           {"size": 11.5, "color": GREY, "italic": True})])

# ============ 8. Funnel ============
slide = new_slide("Target universe: a disciplined funnel", "Universe")
funnel = [
    ("~1,400", "Nordic-listed companies", 12.3),
    ("~150", "EUR 1–10bn market cap, post liquidity screens", 10.2),
    ("~60", "Quality screen: defensible position, sound gross margins, clear peer set", 8.1),
    ("~25–30", "Value gap ≥30% on normalised margins or SOTP", 6.0),
    ("~15–20", "Path to influence: register open, no hostile blocking sphere", 4.2),
    ("6–8", "Funded positions (after MNPI / conflicts clearance)", 2.6),
]
y = 1.3
for i, (n, label, w) in enumerate(funnel):
    x = Inches(0.55 + (12.3 - w) / 2)
    shade = [NAVY, NAVY2, RGBColor(0x1E, 0x4D, 0x7B), RGBColor(0x2E, 0x60, 0x91),
             RGBColor(0x44, 0x74, 0xA4), GOLD][i]
    add_rect(slide, x, Inches(y), Inches(w), Inches(0.72), shade)
    add_text(slide, x + Inches(0.25), Inches(y + 0.08), Inches(1.5), Inches(0.5), n,
             size=18, color=WHITE if i < 5 else NAVY, bold=True)
    add_text(slide, x + Inches(1.8), Inches(y + 0.14), Inches(w - 2.0), Inches(0.5), label,
             size=12.5, color=WHITE if i < 5 else NAVY)
    y += 0.88
add_text(slide, Inches(0.55), Inches(6.7), Inches(12.3), Inches(0.5),
         [("EUR 1.5bn across 6–8 cores = EUR 180–280m per position = 4–10% of EUR 2–6bn "
           "companies — exactly the band where nomination-committee seats are won and "
           "dedicated competition is thinnest.", {"size": 12.5, "color": NAVY, "bold": True})])

# ============ 9–11. Pipeline ============
slide = new_slide("Illustrative pipeline (1/2) — public information only", "Pipeline")
pipe1 = [
    ("Elekta  ·  SE  ·  MedTech", "Margin-gap operator",
     "Global #2 in radiation oncology in a growing oligopoly; heavily de-rated, with "
     "third-party fair-value estimates implying ~50%+ discounts during 2026. FY25/26 adj. "
     "EBIT margin ~12% vs mid-to-high teens at Siemens Healthineers/Varian.",
     "Board-sponsored margin program; order-to-revenue execution reset; service "
     "acceleration; strategic optionality. Healthcare is our home turf."),
    ("Getinge  ·  SE  ·  MedTech", "Margin recovery & portfolio focus",
     "Profitability recovering (profit margin ~6.5% in 2025 vs ~4.7% in 2024) as quality "
     "costs roll off, but far below peers; three loosely related business areas invite "
     "portfolio review.",
     "Complete the quality agenda with board accountability; drive Life Science separation "
     "review; normalise margins toward double-digit EBITA."),
    ("GN Store Nord  ·  DK  ·  Hearing & Audio", "Separation thesis",
     "Two-business structure (hearing care; enterprise/consumer audio) with limited "
     "synergy; separation long debated while leverage suppressed action; register "
     "unusually open for a Danish large-cap.",
     "Deleverage milestones, then a formal separation review — both units have natural "
     "strategic and sponsor buyers."),
]
y = 1.3
for name, arch_t, sit, plan in pipe1:
    add_rect(slide, Inches(0.55), Inches(y), Inches(12.3), Inches(1.78), LIGHT)
    add_rect(slide, Inches(0.55), Inches(y), Pt(4), Inches(1.78), GOLD)
    add_text(slide, Inches(0.8), Inches(y + 0.08), Inches(5.2), Inches(0.35), name,
             size=13.5, color=NAVY, bold=True)
    add_text(slide, Inches(8.6), Inches(y + 0.08), Inches(4.1), Inches(0.35), arch_t,
             size=11, color=GOLD, bold=True, align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.8), Inches(y + 0.45), Inches(6.6), Inches(1.3),
             [("Situation: " + sit, {"size": 10.5, "color": GREY, "line_spacing": 1.06})])
    add_text(slide, Inches(7.6), Inches(y + 0.45), Inches(5.1), Inches(1.3),
             [("Value plan: " + plan, {"size": 10.5, "color": RGBColor(0x2B, 0x33, 0x40),
                                       "line_spacing": 1.06})])
    y += 1.92

slide = new_slide("Illustrative pipeline (2/2) — public information only", "Pipeline")
pipe2 = [
    ("Tieto  ·  FI  ·  Software & IT Services", "Completing the break-up",
     "Tietoevry has validated the separation path: Banking demerger toward a Helsinki "
     "listing announced; Tech Services sold to Agilitas for EUR 300m (closed Sep 2025) with "
     "proceeds to debt.",
     "Support and accelerate value realisation; enforce capital-return discipline; position "
     "the remaining Nordic software core for consolidation — as buyer or seller."),
    ("Orkla  ·  NO  ·  Branded Consumer / Holding", "Conglomerate discount",
     "Self-declared industrial investment company with ~12 portfolio companies, a ~42.6% "
     "Jotun stake (world-class hidden asset) and hydropower; equity persistently below "
     "credible SOTP.",
     "Accelerate announced monetisations; capital-return framework tied to disposal "
     "proceeds; transparent per-asset reporting that forces the SOTP into the price."),
]
y = 1.3
for name, arch_t, sit, plan in pipe2:
    add_rect(slide, Inches(0.55), Inches(y), Inches(12.3), Inches(1.78), LIGHT)
    add_rect(slide, Inches(0.55), Inches(y), Pt(4), Inches(1.78), GOLD)
    add_text(slide, Inches(0.8), Inches(y + 0.08), Inches(5.2), Inches(0.35), name,
             size=13.5, color=NAVY, bold=True)
    add_text(slide, Inches(8.6), Inches(y + 0.08), Inches(4.1), Inches(0.35), arch_t,
             size=11, color=GOLD, bold=True, align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(0.8), Inches(y + 0.45), Inches(6.6), Inches(1.3),
             [("Situation: " + sit, {"size": 10.5, "color": GREY, "line_spacing": 1.06})])
    add_text(slide, Inches(7.6), Inches(y + 0.45), Inches(5.1), Inches(1.3),
             [("Value plan: " + plan, {"size": 10.5, "color": RGBColor(0x2B, 0x33, 0x40),
                                       "line_spacing": 1.06})])
    y += 1.92
add_rect(slide, Inches(0.55), Inches(5.25), Inches(12.3), Inches(1.45), NAVY)
add_text(slide, Inches(0.85), Inches(5.45), Inches(11.8), Inches(1.1),
         [("Watchlist (survives screens 1–3)", {"size": 12, "color": GOLD, "bold": True}),
          ("Securitas (self-help largely delivered — a model of the re-rating we underwrite) · "
           "Demant · Ambu · Dometic · Thule · Huhtamäki · Valmet · Storebrand · ISS",
           {"size": 13, "color": WHITE, "line_spacing": 1.15})])

# ============ 12. Fund terms ============
slide = new_slide("Vehicle, size & terms", "Structure")
table_slide(slide, Inches(0.55), Inches(1.3), Inches(12.3),
            ["Parameter", "Recommendation"],
            [
                ["Target size", "USD 1.5bn target; USD 2.0bn hard cap — sized to the EUR 1–10bn "
                 "mid-cap band where influence is cheap and competition thin"],
                ["Structure", "Luxembourg RAIF (AIFMD); evergreen with 3-year soft lock, then "
                 "quarterly liquidity with 25% investor-level gates"],
                ["GP commitment", "EUR 150–200m from balance sheet + partners/employees"],
                ["Fees", "1.25–1.50% management; 17.5% performance over max(MSCI Nordic + "
                 "200bps, 5% absolute); 3-year crystallisation; high-water mark"],
                ["Concentration", "Max 25% NAV per position at cost (35% at market); max 2 "
                 "concurrent public-escalation campaigns"],
                ["Co-invest", "Sleeves for oversized situations and take-private bridges"],
                ["Target returns", "15%+ net IRR over a cycle; catalyst-driven, entry "
                 "discounts ≥30% provide downside margin"],
            ],
            [2.4, 9.9], row_h=0.68, size=12)

# ============ 13. Team ============
slide = new_slide("Team & organisation: small, senior, conflict-proofed", "Team")
add_rect(slide, Inches(4.3), Inches(1.25), Inches(4.7), Inches(0.75), NAVY)
add_text(slide, Inches(4.3), Inches(1.36), Inches(4.7), Inches(0.5),
         "Head of Engaged Equities (Partner)\nEx-Cevian / top activist franchise · ExCo member",
         size=11.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
boxes = [
    ("Deputy PM / Head of Research", "Underwriting standards · red-team owner"),
    ("Sector Principal — Healthcare/MedTech", "Campaign lead"),
    ("Sector Principal — Industrials/Tech", "Campaign lead"),
    ("4–5 Investment professionals", "Modelling · register analytics · benchmarking"),
    ("Head of Stewardship & Governance", "Nomination committees · director sourcing · proxy"),
    ("Dedicated Compliance Officer", "Owns the info barrier · reports to Group GC"),
]
for i, (t, b) in enumerate(boxes):
    x = Inches(0.55 + (i % 3) * 4.18)
    y = Inches(2.35 + (i // 3) * 1.25)
    add_rect(slide, x, y, Inches(3.95), Inches(1.05), LIGHT)
    add_rect(slide, x, y, Pt(3.5), Inches(1.05), GOLD)
    add_text(slide, x + Inches(0.15), y + Inches(0.08), Inches(3.7), Inches(0.4), t,
             size=11.5, color=NAVY, bold=True)
    add_text(slide, x + Inches(0.15), y + Inches(0.52), Inches(3.7), Inches(0.45), b,
             size=10, color=GREY)
bullet_box(slide, Inches(0.55), Inches(5.05), Inches(12.3), Inches(1.8), [
    ("Conflicts framework (make-or-break): ", "hard information barrier vs the PE business; "
     "restricted-list screening before any accumulation; wall-crossing only via documented "
     "GC approval."),
    ("Take-private protocol: ", "pre-agreed, LPAC-disclosed terms (independent pricing, "
     "roll-or-cash optionality, conflicted-party abstentions) if a position becomes a "
     "buyout target."),
    ("Reputation committee: ", "Managing Partner sign-off required for any public "
     "escalation, media engagement or EGM requisition."),
], size=11.5, gap=6)

# ============ 14. Economics & risks ============
slide = new_slide("Platform economics, risks & mitigants", "Economics & risk")
add_text(slide, Inches(0.55), Inches(1.2), Inches(12.3), Inches(0.35),
         "Economics for the platform", size=14, color=NAVY, bold=True)
bullet_box(slide, Inches(0.55), Inches(1.6), Inches(12.3), Inches(1.2), [
    "~USD 20m annual management fees at target size vs ~USD 12–14m steady-state costs — "
    "accretive from year two; ~USD 30–35m performance fees per cycle at 15% gross.",
    "Strategic P&L is bigger: one incremental take-private or carve-out sourced through "
    "the strategy likely exceeds the vehicle's annual fee income in platform value.",
], size=12, gap=5)
table_slide(slide, Inches(0.55), Inches(2.85), Inches(12.3),
            ["Risk", "Mitigant"],
            [
                ["Reputational spillover to the PE brand",
                 "Constructive-first doctrine; reputation-committee veto; no hostile openings"],
                ["Conflicts / MNPI vs private strategies",
                 "Hard info barrier; dedicated compliance; LPAC-disclosed take-private protocol"],
                ["Sphere control blocks change",
                 "Screened out ex ante; engage spheres as allies where present"],
                ["Mid-cap liquidity (SEK/NOK/DKK)",
                 "ADV-tied sizing; evergreen + gates; co-invest sleeves; event-driven exits"],
                ["Key-person dependence",
                 "Co-PM structure; institutional IC; management-company equity"],
                ["Talent competition (Active Value Partners et al.)",
                 "Move in 2026; platform economics and brand beat start-up offers"],
            ],
            [4.6, 7.7], row_h=0.55, size=11)

# ============ 15. Roadmap ============
slide = new_slide("Implementation roadmap", "Execution")
phases = [
    ("Phase 0 — Decide & design", "Jul–Sep 2026",
     "ExCo approval · conflicts framework with GC · regulatory scoping (Lux RAIF / AIFM "
     "passport) · launch Head search"),
    ("Phase 1 — Build", "Oct 2026 – Mar 2027",
     "Head + Deputy hired · core team of 6 · EUR 150–200m seed closed · 10–14 underwritten "
     "names · LP pre-marketing"),
    ("Phase 2 — First close & deploy", "Apr–Sep 2027",
     "First close ~USD 750m · 2–3 positions built against Q3-2027 ownership snapshots "
     "(seats 2028 nomination committees)"),
    ("Phase 3 — Scale", "Oct 2027 – Dec 2028",
     "Final close USD 1.5–2.0bn · 6–8 positions · first board seats at the 2028 AGM "
     "season · first realisations 2029"),
]
for i, (t, when, b) in enumerate(phases):
    x = Inches(0.55 + i * 3.14)
    add_rect(slide, x, Inches(1.5), Inches(2.95), Inches(3.6), NAVY if i % 2 == 0 else NAVY2)
    add_rect(slide, x, Inches(1.5), Inches(2.95), Pt(4), GOLD)
    add_text(slide, x + Inches(0.15), Inches(1.7), Inches(2.65), Inches(0.8), t,
             size=13.5, color=GOLD, bold=True)
    add_text(slide, x + Inches(0.15), Inches(2.5), Inches(2.65), Inches(0.35), when,
             size=11, color=RGBColor(0x9A, 0xA4, 0xB2), bold=True)
    add_text(slide, x + Inches(0.15), Inches(2.9), Inches(2.65), Inches(2.1), b,
             size=11, color=WHITE, line_spacing=1.15)
add_rect(slide, Inches(0.55), Inches(5.5), Inches(12.3), Inches(1.2), LIGHT)
add_text(slide, Inches(0.85), Inches(5.68), Inches(11.8), Inches(0.9),
         [("The ask", {"size": 12, "color": GOLD, "bold": True}),
          ("Executive Committee mandate to proceed with Phase 0: conflicts framework, "
           "regulatory scoping and the Head of Engaged Equities search — decision point for "
           "full launch in September 2026.", {"size": 13.5, "color": NAVY, "bold": True})])

# ============ 16. Closing ============
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, SW, SH, NAVY)
add_rect(slide, 0, Inches(3.1), SW, Pt(3), GOLD)
add_text(slide, Inches(1.2), Inches(2.0), Inches(11), Inches(1.0),
         "The Nordic engagement franchise has not been built yet.",
         size=30, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(1.2), Inches(3.35), Inches(11), Inches(0.6),
         "Nordic Capital is the firm to build it.", size=22, color=GOLD, bold=True,
         align=PP_ALIGN.CENTER)
add_text(slide, Inches(1.2), Inches(6.5), Inches(11), Inches(0.6),
         "Illustrative discussion material · Company references based on public information "
         "only · Not investment advice", size=9, color=RGBColor(0x8A, 0x96, 0xA6),
         align=PP_ALIGN.CENTER)

prs.save("/home/user/knutnyman/nordic-capital-activism-pitch/Nordic_Capital_Engaged_Equities_Deck.pptx")
print("Deck saved with", len(prs.slides.__iter__.__self__._sldIdLst), "slides.")
