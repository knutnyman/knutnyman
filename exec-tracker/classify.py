"""Title → seniority label + division + sort rank."""

_SENIORITY = [
    ("chief executive", "CEO", 1),
    ("ceo", "CEO", 1),
    ("chairman", "Chairman", 2),
    ("vice chairman", "Vice Chairman", 3),
    ("president", "President", 4),
    ("chief operating", "COO", 5),
    ("coo", "COO", 5),
    ("chief financial", "CFO", 6),
    ("cfo", "CFO", 6),
    ("chief accounting", "CAO / Controller", 7),
    ("chief legal", "CLO / General Counsel", 8),
    ("general counsel", "CLO / General Counsel", 8),
    ("chief technology", "CTO", 9),
    ("cto", "CTO", 9),
    ("chief information", "CIO", 10),
    ("cio", "CIO", 10),
    ("chief marketing", "CMO", 11),
    ("chief commercial", "Chief Commercial Officer", 12),
    ("chief revenue", "CRO", 13),
    ("cro", "CRO", 13),
    ("chief human", "CHRO", 14),
    ("chief people", "CHRO", 14),
    ("chief strategy", "CSO", 15),
    ("chief scientific", "Chief Scientific Officer", 16),
    ("chief medical", "Chief Medical Officer", 17),
    ("chief product", "CPO", 18),
    ("chief supply", "Chief Supply Chain Officer", 19),
    ("chief compliance", "Chief Compliance Officer", 20),
    ("executive vice president", "EVP", 30),
    ("evp", "EVP", 30),
    ("senior vice president", "SVP", 40),
    ("svp", "SVP", 40),
    ("vice president", "VP", 50),
]

_DIVISION = [
    # Finance
    (["chief financial", "cfo", "financ", "accounting", "controller", "treasurer", "tax", "audit", "investor relation"], "Finance"),
    # Technology
    (["technology", "information technology", "digital", " data ", "engineering", "software", "infrastructure", "cybersecurity", "cyber", "cto", "cio"], "Technology"),
    # Legal & Compliance
    (["legal", "counsel", "compliance", "regulatory", "governance", "secretary"], "Legal & Compliance"),
    # Marketing & Comms
    (["marketing", "communications", "brand", "public relation", "pr,", "media", "cmo"], "Marketing & Comms"),
    # Sales & Revenue
    (["sales", "revenue", " cro", "commercial", "customer success", "account"], "Sales & Revenue"),
    # Strategy & BD
    (["strategy", "business development", "corporate development", "m&a", "mergers", "acquisition", "planning"], "Strategy & BD"),
    # Human Resources
    (["human resources", "people", "talent", " hr ", "chro", "workforce", "culture", "learning"], "Human Resources"),
    # Operations
    (["operations", "manufacturing", "supply chain", "logistics", "procurement", "quality", "facilities"], "Operations"),
    # R&D / Science
    (["research", "scientific", "science", "innovation", "r&d", "intellectual property"], "R&D / Science"),
    # Medical Affairs
    (["medical", "clinical", "regulatory affairs", "pharmacovigil", "drug safety"], "Medical Affairs"),
    # International
    (["international", "global", "latin america", "europe", "asia", "emea", "apac"], "International"),
]


def classify(title: str) -> tuple[str, str, int]:
    """Returns (seniority_label, division, sort_rank)."""
    tl = title.lower()

    # Seniority
    label, rank = "SVP", 40
    for kw, lbl, r in _SENIORITY:
        if kw in tl:
            label, rank = lbl, r
            break

    # Division
    division = "Executive Leadership" if rank <= 20 else "Other"
    for keywords, div in _DIVISION:
        if any(kw in tl for kw in keywords):
            division = div
            break

    return label, division, rank


def linkedin_search_url(name: str, company: str) -> str:
    from urllib.parse import quote
    q = quote(f'"{name}" {company}')
    return f"https://www.linkedin.com/search/results/people/?keywords={q}"
