#!/usr/bin/env python3
"""Generate an Excel workbook of notable funds founded by Stanford GSB alumni.

Run: python3 build_gsb_funds_xlsx.py
Output: stanford-gsb-fund-founders.xlsx
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Data -------------------------------------------------------------------
# Columns:
# Fund | Category | GSB Founder(s) | GSB Credential | Other Co-Founders (non-GSB)
# | Founded | HQ | Strategy / Focus | AUM / Fund Size | AUM As-Of | Status
# | Notable Investments / Notes | Source

HEADERS = [
    "Fund", "Category", "GSB Founder(s)", "GSB Credential",
    "Other Co-Founders (non-GSB)", "Founded", "HQ", "Strategy / Focus",
    "AUM / Fund Size (approx.)", "AUM As-Of", "Status",
    "Notable Investments / Notes", "Source",
]

FUNDS = [
    ["Khosla Ventures", "Venture Capital", "Vinod Khosla", "MBA 1980", "—",
     2004, "Menlo Park, CA", "Early & growth stage; AI, clean energy, deep tech, digital health, fintech",
     "$16–18B AUM", "2025–26", "Active",
     "Co-founded Sun Microsystems; early backer of OpenAI, Stripe, DoorDash, Instacart, Square",
     "https://en.wikipedia.org/wiki/Khosla_Ventures"],

    ["Thoma Bravo", "Private Equity", "Orlando Bravo", "MBA + Stanford Law JD", "—",
     2008, "Chicago / San Francisco, CA", "Enterprise software & tech-enabled services buyouts",
     ">$130B AUM", "2024–25", "Active",
     "\"King of SaaS\"; rebranded Thoma Cressey to Thoma Bravo in 2008; one of the largest software PE firms",
     "https://en.wikipedia.org/wiki/Orlando_Bravo"],

    ["Lightspeed Venture Partners", "Venture Capital", "Ravi Mhatre; Barry Eggers", "Both MBA",
     "Christopher Schaepe; Peter Nieh", 2000, "Menlo Park, CA",
     "Multi-stage, global; enterprise, consumer, fintech, crypto, health",
     "~$25B+ AUM", "2024–25", "Active",
     "Spun out of Weiss, Peck & Greer; backed Snap, Nutanix, Affirm, Rubrik, Mistral",
     "https://en.wikipedia.org/wiki/Lightspeed_Venture_Partners"],

    ["Kaszek Ventures", "Venture Capital", "Hernán Kazah; Nicolás Szekasy", "Both MBA (Szekasy 1991)", "—",
     2011, "Latin America (Cayman-registered)", "Latin America's leading early-stage VC",
     "$2B+ raised across funds", "2024", "Active",
     "Both built MercadoLibre (COO/CFO); seed backer of Nubank; ~16 LatAm unicorns",
     "https://www.americasquarterly.org/article/still-betting-big-argentine-venture-capitalists-hernan-kazah-and-nicolas-szekasy/"],

    ["Future Ventures", "Venture Capital", "Steve Jurvetson", "MBA (+ BS/MS EE)", "Maryanna Saenko",
     2018, "Menlo Park, CA", "Frontier/deep tech: space, AI, synthetic biology, transportation",
     "~$400M across funds", "2023", "Active",
     "Jurvetson was the 'J' in DFJ; early SpaceX & Tesla investor/board member",
     "https://en.wikipedia.org/wiki/Steve_Jurvetson"],

    ["Aspect Ventures / Acrew Capital", "Venture Capital", "Theresia Gouw", "MBA",
     "Jennifer Fonstad (Harvard MBA)", 2014, "San Francisco / Palo Alto, CA",
     "Early-stage; cybersecurity, fintech, enterprise & consumer software",
     "Aspect ~$350M; Acrew adds several $B", "2024", "Active (as Acrew)",
     "Aspect 2014; Gouw later co-founded Acrew (2019); backed ForeScout, Exabeam, Cato Networks",
     "https://en.wikipedia.org/wiki/Theresia_Gouw"],

    ["Fifth Wall", "Venture Capital", "Brendan Wallace", "MBA 2010", "Brad Greiwe",
     2016, "Los Angeles / New York", "Largest PropTech-focused VC; real-estate & built-world tech, climate",
     "~$3B+ AUM", "2024", "Active",
     "Backed by major real-estate strategics; largest dedicated PropTech fund",
     "https://en.wikipedia.org/wiki/Brendan_F._Wallace"],

    ["Costanoa Ventures", "Venture Capital", "Greg Sands", "MBA", "—",
     2012, "Palo Alto, CA", "Early-stage B2B; applied AI, data, security, fintech, vertical SaaS",
     "~$1B+ AUM across funds", "2024", "Active",
     "Sands was Netscape's first PM (coined the name 'Netscape'); ex-Sutter Hill MD",
     "https://www.costanoavc.com/team/investors/greg-sands/"],

    ["Baseline Ventures", "Venture Capital (Seed)", "Steve Anderson", "MBA", "—",
     2006, "San Francisco, CA", "Seed stage; deliberately small, solo-GP",
     "~$300M+ across funds", "2023", "Active",
     "FIRST seed investor in Instagram; early Twitter, Heroku, SocialCode",
     "https://en.wikipedia.org/wiki/Baseline_Ventures"],

    ["Homebrew", "Venture Capital (Seed)", "Hunter Walk", "MBA", "Satya Patel (UPenn)",
     2013, "San Francisco, CA", "Seed; moved to evergreen/permanent-capital model",
     "~$300M raised; now evergreen", "2023", "Active",
     "Ex-YouTube/Google product lead; co-launched Screendoor to back diverse fund managers",
     "https://conferences.law.stanford.edu/vcs2020/speakers/hunter-walk/"],

    ["Precursor Ventures", "Venture Capital (Pre-seed)", "Charles Hudson", "MBA", "—",
     2015, "San Francisco, CA", "Pre-seed / first institutional check specialist",
     "~$250M+ AUM", "2024", "Active",
     "Among the most active pre-seed firms; Hudson also lectures at Stanford GSB; ex-SoftTech VC",
     "https://precursorvc.com/team_member/charles-hudson/"],

    ["Renegade Partners", "Venture Capital", "Renata Quintini", "MBA + LL.M.", "Roseanne Wincek",
     2019, "San Francisco Bay Area, CA", "'Supercritical' Series A/B stage",
     "~$250M+ AUM", "2023", "Active",
     "Quintini helped build Felicis and was a Lux Capital partner; backed Warby Parker, Planet, Cruise",
     "https://www.renegadepartners.com/renata-quintini-renegade"],

    ["Freestyle Capital", "Venture Capital (Seed)", "Jenny Lefcourt (GP)", "Attended GSB (left for startup)",
     "Founders Dave Samuel & Josh Felser", 2009, "San Francisco, CA", "Seed-stage software",
     "~$600M+ AUM", "2023", "Active",
     "Lefcourt is GP (joined 2015), a GSB attendee rather than degree-holder; co-founded WeddingChannel",
     "https://freestyle.vc/team/jenny-lefcourt/"],

    ["The 20|20 Fund", "Student/Community Fund", "Steph Mui & GSB Class of 2020", "MBA 2020", "—",
     2020, "Stanford, CA", "Student-led fund democratizing startup investing",
     "$1.5M inaugural fund", "2020", "Active",
     "Invested in 16 portfolio companies founded by GSB classmates",
     "https://poetsandquants.com/2023/04/10/these-stanford-mbas-are-opening-doors-to-greater-startup-investment-heres-how/"],

    ["Reaction", "Venture Capital (Community)", "Dan Matthies & Enrico Carbone", "Stanford Exec. Program (SEP)", "—",
     2021, "Global", "Global, mission-driven ('transform 1 billion lives in 10 years')",
     "Pooled SEP-alumni fund", "2023", "Active",
     "Founded by 150+ Stanford SEP/alumni across 45 countries",
     "https://poetsandquantsforexecs.com/news/stanford-born-vc-funds-goal-transform-1-billion-lives-in-10-years/"],

    # --- Hedge Funds ---
    ["Viking Global Investors", "Hedge Fund", "Andreas Halvorsen", "MBA 1990",
     "David Ott; Brian Olson", 1999, "Greenwich, CT", "Long/short & long-only global equities (Tiger Cub)",
     "~$40B+ AUM", "2024", "Active",
     "Ex-Tiger Management director of equities; one of the largest, most respected Tiger Cubs",
     "https://en.wikipedia.org/wiki/Ole_Andreas_Halvorsen"],

    ["Farallon Capital Management", "Hedge Fund", "Tom Steyer", "MBA 1983 (Arjay Miller Scholar)", "—",
     1986, "San Francisco, CA", "Multi-strategy / absolute return; merger arb, credit, distressed, real estate",
     "~$39B AUM", "2024", "Active (Steyer left 2012)",
     "Pioneered West Coast multi-strategy; Steyer later climate activist & presidential candidate",
     "https://en.wikipedia.org/wiki/Tom_Steyer"],

    ["Blue Ridge Capital", "Hedge Fund", "John A. Griffin", "MBA 1990", "—",
     1996, "New York, NY", "Long/short equity (Tiger Cub)",
     "~$6B at peak", "2017 (peak)", "Closed (wound down 2017)",
     "Ex-president of Tiger Management; 15%+ annualized over ~2 decades; Halvorsen's GSB classmate",
     "https://en.wikipedia.org/wiki/Blue_Ridge_Capital"],

    ["Passport Capital", "Hedge Fund", "John H. Burbank III", "MBA 1992", "—",
     2000, "San Francisco, CA", "Global macro + high-conviction fundamental equity",
     "~$5B at peak", "~2015 (peak)", "Largely wound down (became family office ~2018)",
     "Seeded with ~$800K; +36% first year; known for pre-2008 subprime short",
     "https://thehedgefundjournal.com/passport-capital/"],

    ["Light Street Capital", "Hedge Fund", "Glen Kacher", "MBA 1998", "—",
     2010, "Palo Alto, CA", "Technology long/short & long-only + private growth (Tiger Grandcub)",
     "~$2B+ AUM", "2024", "Active",
     "Came via Integral Capital & Tiger Management; repeatedly top-performing tech hedge fund",
     "https://www.cnbc.com/glen-kacher-light-street-capital/"],
]

NON_QUALIFIERS = [
    ["Bond Capital", "Mary Meeker", "Cornell (Johnson) MBA — NOT GSB",
     "Founded 2019, $1.25B debut growth fund; often wrongly assumed GSB"],
    ["Floodgate", "Mike Maples Jr.", "Harvard MBA (Stanford undergrad only) — NOT GSB",
     "Founded 2006; Stanford engineering BS but MBA is from HBS"],
    ["Aspect Ventures (co-founder)", "Jennifer Fonstad", "Harvard MBA — NOT GSB",
     "Co-founded Aspect with Theresia Gouw (who IS GSB)"],
    ["Homebrew (co-founder)", "Satya Patel", "University of Pennsylvania — NOT GSB",
     "Co-founded Homebrew with Hunter Walk (who IS GSB)"],
    ["Sequoia Capital", "Roelof Botha", "GSB MBA 2000 — but LEADS, did not FOUND",
     "Sequoia founded by Don Valentine in 1972; Botha is a GSB alum but not a fund founder"],
    ["Shumway Capital Partners", "Chris Shumway", "Harvard MBA — NOT GSB",
     "Founded 2001, grew to $9B+; prominent Tiger Cub but HBS, not Stanford"],
    ["Lone Pine Capital", "Stephen Mandel", "Harvard MBA — NOT GSB", "Tiger Cub"],
    ["Maverick Capital", "Lee Ainslie", "UNC MBA — NOT GSB", "Tiger Cub"],
    ["Tiger Global Management", "Chase Coleman", "Williams (no MBA) — NOT GSB", "Tiger Cub"],
    ["Coatue Management", "Philippe Laffont", "MIT — NOT GSB", "Tiger Cub"],
]

# --- Styling ----------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
CAT_FILLS = {
    "Hedge Fund": "FCE4D6",
    "Private Equity": "E2EFDA",
    "Student/Community Fund": "FFF2CC",
    "Venture Capital (Community)": "FFF2CC",
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{ws.max_row}"


def build():
    wb = Workbook()

    # Sheet 1: Funds
    ws = wb.active
    ws.title = "Funds"
    ws.append(HEADERS)
    for row in FUNDS:
        ws.append(row)
    style_header(ws, len(HEADERS))

    widths = [26, 22, 24, 22, 26, 9, 24, 38, 26, 12, 26, 46, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    cat_col = 2
    for r in range(2, ws.max_row + 1):
        cat = ws.cell(row=r, column=cat_col).value
        fill_hex = CAT_FILLS.get(cat)
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = WRAP_TOP
            cell.border = BORDER
            if fill_hex:
                cell.fill = PatternFill("solid", fgColor=fill_hex)
        ws.cell(row=r, column=1).font = Font(bold=True)

    # Sheet 2: Appendix - Non-Qualifiers
    ws2 = wb.create_sheet("Non-Qualifiers")
    nq_headers = ["Fund", "Person", "Why excluded (school)", "Notes"]
    ws2.append(nq_headers)
    for row in NON_QUALIFIERS:
        ws2.append(row)
    style_header(ws2, len(nq_headers))
    for i, w in enumerate([28, 22, 40, 50], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for r in range(2, ws2.max_row + 1):
        for c in range(1, len(nq_headers) + 1):
            cell = ws2.cell(row=r, column=c)
            cell.alignment = WRAP_TOP
            cell.border = BORDER
        ws2.cell(row=r, column=1).font = Font(bold=True)

    # Sheet 3: About / Methodology
    ws3 = wb.create_sheet("About")
    about = [
        ["Notable Fund Launches Founded by Stanford GSB Alumni"],
        [""],
        ["Scope: VC, growth, private-equity, and hedge funds whose FOUNDERS earned a"],
        ["degree from (or attended) the Stanford Graduate School of Business — distinct"],
        ["from Stanford undergrad, engineering, or law programs."],
        [""],
        ["Compiled: June 2026."],
        [""],
        ["Methodology:"],
        [" - Each founder verified to have a GSB affiliation (MBA unless noted)."],
        [" - Mixed founding teams: only the GSB founder(s) are credited; others flagged."],
        [" - AUM / fund-size figures are APPROXIMATE, drawn from public reporting, and"],
        ["   labelled with an 'As-Of' year. They fluctuate and should be treated as"],
        ["   indicative, not audited."],
        [" - See the 'Non-Qualifiers' tab for commonly-assumed names that are NOT GSB."],
        [""],
        ["Tabs:"],
        [" - Funds: main list (20 funds, sortable/filterable)."],
        [" - Non-Qualifiers: excluded names with the reason."],
        [""],
        ["Color key (Funds tab, Category column):"],
        ["   Orange = Hedge Fund   Green = Private Equity   Yellow = Student/Community"],
    ]
    for row in about:
        ws3.append(row)
    ws3["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws3.column_dimensions["A"].width = 78
    for r in (9, 17, 21):
        ws3.cell(row=r, column=1).font = Font(bold=True)

    out = "stanford-gsb-fund-founders.xlsx"
    wb.save(out)
    print(f"Wrote {out} with {len(FUNDS)} funds across {len(wb.sheetnames)} sheets.")


if __name__ == "__main__":
    build()
