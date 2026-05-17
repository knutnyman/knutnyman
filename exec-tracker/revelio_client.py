"""
Revelio Labs workforce intelligence API client.

Revelio has ex-executive data sourced from LinkedIn profiles — great for
SVP+ coverage beyond what SEC proxy statements list.

Get API access at: https://www.reveliolabs.com
Docs: https://api.reveliolabs.com/docs
"""

import requests
from typing import Optional

_BASE = "https://api.reveliolabs.com/v1"


def fetch_former_executives(
    company_name: str,
    api_key: str,
    min_seniority: str = "svp",  # "svp" | "evp" | "c_suite"
    limit: int = 200,
) -> list[dict]:
    """
    Returns list of former senior employees from Revelio Labs.
    Each record: {name, title, division, start_date, end_date, current_employer}
    """
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    # Revelio uses a /people/search endpoint with workforce filters
    payload = {
        "employer": company_name,
        "employment_status": "past",
        "min_seniority": min_seniority,
        "limit": limit,
        "fields": ["name", "title", "function", "start_date", "end_date", "current_employer"],
    }

    try:
        r = requests.post(f"{_BASE}/people/search", json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        data = r.json()
        return _normalize(data.get("results", []))
    except requests.HTTPError as e:
        raise RuntimeError(f"Revelio API error {e.response.status_code}: {e.response.text[:200]}") from e
    except Exception as e:
        raise RuntimeError(f"Revelio request failed: {e}") from e


def _normalize(raw: list[dict]) -> list[dict]:
    """Normalize Revelio response to our internal schema."""
    out = []
    for p in raw:
        out.append({
            "name": p.get("name", ""),
            "title": p.get("title", ""),
            "division": p.get("function", ""),
            "first_seen": _year(p.get("start_date")),
            "last_seen": _year(p.get("end_date")),
            "departed_after": _year(p.get("end_date")),
            "current_employer": p.get("current_employer", {}).get("name", ""),
            "source": "Revelio Labs",
        })
    return out


def _year(date_str: Optional[str]) -> Optional[int]:
    if not date_str:
        return None
    try:
        return int(str(date_str)[:4])
    except ValueError:
        return None
