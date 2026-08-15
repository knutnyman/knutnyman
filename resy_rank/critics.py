"""Critic list loading and fuzzy name matching.

`data/critics.csv` is hand-maintained: venue_name,source,tier,url. Names there
rarely match the Resy spelling exactly ("Cote" vs "COTE Korean Steakhouse"), so
matching is fuzzy — and every row that fails to match is printed rather than
silently dropped, because a missed match quietly removes up to a third of a
venue's composite score.
"""

from __future__ import annotations

import csv
import logging
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from rapidfuzz import fuzz, process

log = logging.getLogger(__name__)

REQUIRED_COLUMNS = {"venue_name", "source"}


@dataclass(frozen=True)
class CriticRow:
    venue_name: str
    source: str
    tier: str | None
    url: str | None
    lineno: int


@dataclass
class CriticMatch:
    row: CriticRow
    venue_id: int
    venue_name: str
    score: float


@dataclass
class MatchReport:
    matched: list[CriticMatch]
    unmatched: list[tuple[CriticRow, str, float]]  # row, closest name, its score
    unknown_sources: list[CriticRow]

    @property
    def ok(self) -> bool:
        return not self.unmatched and not self.unknown_sources


def load_critics_csv(path: Path) -> list[CriticRow]:
    if not path.exists():
        raise FileNotFoundError(f"critic file not found: {path}")

    rows: list[CriticRow] = []
    with path.open(newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise ValueError(f"{path} is empty — expected a header row")
        missing = REQUIRED_COLUMNS - set(reader.fieldnames)
        if missing:
            raise ValueError(f"{path} is missing column(s): {', '.join(sorted(missing))}")

        for lineno, raw in enumerate(reader, start=2):
            venue_name = (raw.get("venue_name") or "").strip()
            source = (raw.get("source") or "").strip().lower()
            if not venue_name or not source:
                log.warning("%s:%d — blank venue_name or source, skipping", path.name, lineno)
                continue
            rows.append(
                CriticRow(
                    venue_name=venue_name,
                    source=source,
                    tier=(raw.get("tier") or "").strip() or None,
                    url=(raw.get("url") or "").strip() or None,
                    lineno=lineno,
                )
            )
    return rows


def match_critics(
    rows: list[CriticRow],
    venues: list[sqlite3.Row],
    *,
    threshold: int,
    known_sources: set[str],
) -> MatchReport:
    """Fuzzy-match critic rows to venues. Anything below threshold is reported."""
    matched: list[CriticMatch] = []
    unmatched: list[tuple[CriticRow, str, float]] = []
    unknown_sources: list[CriticRow] = []

    choices = {row["venue_id"]: row["name"] for row in venues}
    if not choices:
        return MatchReport([], [(r, "", 0.0) for r in rows], [])

    for row in rows:
        if row.source not in known_sources:
            unknown_sources.append(row)
            # Still try to match it, so the report shows one problem, not two.

        # WRatio handles the common cases here: partial names ("Cote" vs
        # "COTE Korean Steakhouse") and reordered words.
        best = process.extractOne(
            row.venue_name, choices, scorer=fuzz.WRatio, processor=lambda s: s.lower()
        )
        if best is None:
            unmatched.append((row, "", 0.0))
            continue
        name, score, venue_id = best
        if score < threshold:
            unmatched.append((row, name, score))
            continue
        matched.append(CriticMatch(row=row, venue_id=venue_id, venue_name=name, score=score))

    return MatchReport(matched=matched, unmatched=unmatched, unknown_sources=unknown_sources)


def save_flags(conn: sqlite3.Connection, matches: list[CriticMatch]) -> int:
    """Upsert matched flags. One flag per (venue, source) — re-running is safe."""
    with conn:
        conn.executemany(
            """
            INSERT INTO critic_flags (venue_id, source, tier, url)
            VALUES (?, ?, ?, ?)
            ON CONFLICT (venue_id, source) DO UPDATE SET
                tier = excluded.tier,
                url  = excluded.url
            """,
            [(m.venue_id, m.row.source, m.row.tier, m.row.url) for m in matches],
        )
    return len(matches)


def flags_by_venue(conn: sqlite3.Connection) -> dict[int, list[str]]:
    """venue_id -> list of critic sources, for scoring."""
    result: dict[int, list[str]] = {}
    for row in conn.execute("SELECT venue_id, source FROM critic_flags"):
        result.setdefault(row["venue_id"], []).append(row["source"])
    return result
