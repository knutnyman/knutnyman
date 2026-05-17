"""
SEC EDGAR parser for executive officer data.

Sources mined:
  - 10-K "Executive Officers of the Registrant" section (broadest coverage)
  - DEF 14A Summary Compensation Table (Named Executive Officers)

Cross-year comparison surfaces who has departed.
"""

import re
import time
import requests
from typing import Optional
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Former Executive Tracker contact@research-firm.com"}
_RATE_LIMIT_SECS = 0.12  # SEC rate limit: max 10 req/s


def _get(url: str, timeout: int = 20) -> Optional[requests.Response]:
    try:
        time.sleep(_RATE_LIMIT_SECS)
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r
    except Exception:
        return None


# ── Company lookup ──────────────────────────────────────────────────────────

def get_cik_and_name(ticker: str) -> Optional[tuple[str, str]]:
    """Returns (zero-padded CIK, company name) for a ticker, or None."""
    r = _get("https://www.sec.gov/files/company_tickers.json")
    if not r:
        return None
    for item in r.json().values():
        if item["ticker"].upper() == ticker.upper():
            return str(item["cik_str"]).zfill(10), item["name"]
    return None


# ── Filing index ────────────────────────────────────────────────────────────

def get_filings(cik: str, form_type: str, max_results: int = 12) -> list[dict]:
    """Return list of {year, date, accession, primary_doc} for form_type."""
    r = _get(f"https://data.sec.gov/submissions/CIK{cik}.json")
    if not r:
        return []

    data = r.json()
    results = _extract_filings(data["filings"]["recent"], form_type, max_results)

    if len(results) < max_results:
        for file_info in data.get("filings", {}).get("files", []):
            if len(results) >= max_results:
                break
            r2 = _get(f"https://data.sec.gov/submissions/{file_info['name']}")
            if r2:
                results += _extract_filings(r2.json(), form_type, max_results - len(results))

    return sorted(results, key=lambda x: x["year"], reverse=True)[:max_results]


def _extract_filings(recent: dict, form_type: str, limit: int) -> list[dict]:
    forms = recent.get("form", [])
    dates = recent.get("filingDate", [])
    accessions = recent.get("accessionNumber", [])
    primary_docs = recent.get("primaryDocument", [])

    out = []
    for i, form in enumerate(forms):
        if form == form_type:
            out.append({
                "year": int(dates[i][:4]),
                "date": dates[i],
                "accession": accessions[i].replace("-", ""),
                "primary_doc": primary_docs[i] if i < len(primary_docs) else None,
            })
            if len(out) >= limit:
                break
    return out


# ── Document download ───────────────────────────────────────────────────────

def fetch_filing_html(cik: str, filing: dict) -> Optional[str]:
    cik_int = int(cik)
    acc = filing["accession"]
    primary = filing.get("primary_doc")

    if primary:
        r = _get(f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc}/{primary}")
        if r:
            return r.text

    # Fall back: parse the index page to find the first .htm document
    idx = _get(f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc}/{acc}-index.htm")
    if not idx:
        return None
    soup = BeautifulSoup(idx.text, "lxml")
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if href.lower().endswith((".htm", ".html")) and "index" not in href.lower():
            url = f"https://www.sec.gov{href}" if href.startswith("/") else href
            r = _get(url)
            if r:
                return r.text
    return None


# ── Parsing: executive names + titles ──────────────────────────────────────

_SENIOR_MARKERS = [
    "chief ", "president", "executive vice president", "evp", "evp,",
    "senior vice president", "svp", "svp,", "general counsel",
    "ceo", "cfo", "cto", "coo", "cmo", "clo", "chro", "cso",
]

_NAME_RE = re.compile(
    r"\b([A-Z][a-z]+(?:\s+[A-Z]\.?)?\s+[A-Z][a-z]+(?: [A-Z][a-z]+)?)\b"
)

_EXEC_PROSE_RE = re.compile(
    r"([A-Z][a-z]+(?: [A-Z]\.?)? [A-Z][a-z]+(?: [A-Z][a-z]+)?)"
    r"(?:,\s*age \d+[,\.]?)?\s*"
    r"(?:has served as|serves as|is the|was|currently)\s+"
    r"(?:our\s+)?"
    r"((?:Chief|Executive Vice President|Senior Vice President|SVP|EVP|President|General Counsel|Vice Chairman)"
    r"[^\.]{0,80})",
    re.IGNORECASE,
)


def is_senior(title: str) -> bool:
    tl = title.lower()
    return any(m in tl for m in _SENIOR_MARKERS)


def _clean_name(raw: str) -> str:
    # Remove Jr., III, age info, leading/trailing junk
    name = re.sub(r",?\s*(Jr\.|Sr\.|III|II|IV)\b", "", raw)
    name = re.sub(r"\s{2,}", " ", name).strip()
    return name


def parse_executives(html: str) -> list[dict]:
    """Return [{name, title}] found in the HTML of one filing."""
    soup = BeautifulSoup(html, "lxml")
    results: list[dict] = []
    seen: set[str] = set()

    def add(name: str, title: str):
        n = _clean_name(name)
        key = n.lower()
        if key in seen or len(n) < 5 or not is_senior(title):
            return
        seen.add(key)
        results.append({"name": n, "title": title.strip()})

    # ── Strategy 1: tables with "name and principal position" or exec officer ──
    for table in soup.find_all("table"):
        header_text = " ".join(
            th.get_text(" ", strip=True).lower()
            for th in table.find_all(["th", "td"])[:12]
        )
        if not (
            ("principal position" in header_text or "executive officer" in header_text)
            and "name" in header_text
        ):
            continue

        for row in table.find_all("tr"):
            cells = row.find_all(["td", "th"])
            if not cells:
                continue
            raw = cells[0].get_text("\n", strip=True)
            lines = [l.strip() for l in raw.split("\n") if l.strip() and len(l.strip()) > 2]
            if not lines:
                continue

            # Skip header row
            if any(h in lines[0].lower() for h in ["name", "principal", "position", "salary", "compensation"]):
                continue

            name, title = None, None
            for i, line in enumerate(lines):
                words = line.split()
                # Heuristic: name = 2–5 title-cased words, no digits or $
                if (
                    2 <= len(words) <= 5
                    and not any(c.isdigit() or c in "$%(),." for c in line)
                    and sum(1 for w in words if w and w[0].isupper()) >= max(1, len(words) - 1)
                ):
                    candidate_title = lines[i + 1] if i + 1 < len(lines) else ""
                    # Also try second cell as title source
                    if not candidate_title and len(cells) > 1:
                        candidate_title = cells[1].get_text(" ", strip=True)
                    name, title = line, candidate_title
                    break

            if name and title:
                add(name, title)
            elif name and len(cells) > 1:
                add(name, cells[1].get_text(" ", strip=True))

    # ── Strategy 2: "Executive Officers of the Registrant" prose section ──
    exec_header = None
    for tag in soup.find_all(True):
        text = tag.get_text(" ", strip=True)
        if re.search(r"executive officers? of the registrant", text, re.I) and len(text) < 120:
            exec_header = tag
            break

    if exec_header:
        chunk_parts = []
        for sib in exec_header.next_siblings:
            if not hasattr(sib, "name"):
                continue
            txt = sib.get_text(" ", strip=True)
            # Stop at the next major section header
            if len(txt) < 120 and re.search(r"^(item\s+\d|part\s+[ivx])", txt, re.I):
                break
            chunk_parts.append(txt)
            if len(chunk_parts) > 80:
                break
        chunk = " ".join(chunk_parts)

        # Pattern A: "Name, age N, [has served as / is / serves as] Title"
        for m in _EXEC_PROSE_RE.finditer(chunk):
            add(m.group(1), m.group(2))

        # Pattern B: Bolded/formatted name followed by title on next line (10-K tables)
        for bold in exec_header.find_all_next(["b", "strong"], limit=60):
            name_cand = bold.get_text(" ", strip=True)
            if not _NAME_RE.match(name_cand):
                continue
            nxt = bold.find_next(string=True)
            title_cand = nxt.strip() if nxt else ""
            if not title_cand:
                parent_next = bold.parent.find_next_sibling()
                if parent_next:
                    title_cand = parent_next.get_text(" ", strip=True)
            add(name_cand, title_cand)

    return results


# ── Cross-year diffing ──────────────────────────────────────────────────────

def find_formers(
    yearly: dict[int, list[dict]],
    current_year: int,
    lookback_years: int = 10,
) -> list[dict]:
    """
    yearly: {year: [{name, title}]}
    Returns former executives — appeared in past filings, absent in most recent.
    """
    all_people: dict[str, dict] = {}

    for year, execs in yearly.items():
        for e in execs:
            key = e["name"].lower()
            if key not in all_people:
                all_people[key] = {"name": e["name"], "years": [], "titles": {}}
            all_people[key]["years"].append(year)
            all_people[key]["titles"][year] = e["title"]

    current_keys = {e["name"].lower() for e in yearly.get(current_year, [])}

    formers = []
    for key, data in all_people.items():
        if key in current_keys:
            continue
        years = sorted(data["years"])
        last_year = max(years)
        first_year = min(years)
        latest_title = data["titles"][last_year]
        formers.append({
            "name": data["name"],
            "title": latest_title,
            "first_seen": first_year,
            "last_seen": last_year,
            "departed_after": last_year,
        })

    return formers
