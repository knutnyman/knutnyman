"""SQLite schema, migrations, and connection helpers.

Migrations are a plain ordered list keyed off `PRAGMA user_version`. Adding a
schema change means appending a new entry — never editing an existing one.
"""

from __future__ import annotations

import csv
import logging
import sqlite3
from dataclasses import dataclass
from pathlib import Path

log = logging.getLogger(__name__)

# Ordered DDL. Index i in this list migrates user_version i -> i+1.
MIGRATIONS: list[str] = [
    # ── 0 -> 1: initial schema ──────────────────────────────────────────────
    """
    CREATE TABLE venues (
        -- Local surrogate key. The Resy id is unknown until `resolve` runs,
        -- so it cannot serve as the primary key at insert time.
        venue_id        INTEGER PRIMARY KEY AUTOINCREMENT,
        name            TEXT    NOT NULL,
        neighborhood    TEXT    NOT NULL DEFAULT '',
        cuisine         TEXT    NOT NULL DEFAULT '',
        address         TEXT    NOT NULL DEFAULT '',
        resy_url        TEXT,
        resy_venue_id   INTEGER UNIQUE,
        google_place_id TEXT,
        lat             REAL,
        lng             REAL,
        created_at      TEXT    NOT NULL DEFAULT (datetime('now')),
        UNIQUE (name, neighborhood)
    );

    CREATE TABLE google_meta (
        venue_id          INTEGER PRIMARY KEY REFERENCES venues(venue_id) ON DELETE CASCADE,
        rating            REAL,
        user_rating_count INTEGER,
        price_level       TEXT,
        website_uri       TEXT,
        fetched_at        TEXT NOT NULL
    );

    CREATE TABLE critic_flags (
        venue_id INTEGER NOT NULL REFERENCES venues(venue_id) ON DELETE CASCADE,
        source   TEXT    NOT NULL,
        tier     TEXT,
        url      TEXT,
        PRIMARY KEY (venue_id, source)
    );

    CREATE TABLE scans (
        scan_id     INTEGER PRIMARY KEY AUTOINCREMENT,
        venue_id    INTEGER NOT NULL REFERENCES venues(venue_id) ON DELETE CASCADE,
        scan_ts     TEXT    NOT NULL,
        target_date TEXT    NOT NULL,
        party_size  INTEGER NOT NULL,
        -- Raw /4/find payload, kept so historical scans can be re-parsed if
        -- the undocumented response shape shifts under us.
        raw_json    TEXT
    );

    CREATE TABLE slots (
        scan_id      INTEGER NOT NULL REFERENCES scans(scan_id) ON DELETE CASCADE,
        venue_id     INTEGER NOT NULL REFERENCES venues(venue_id) ON DELETE CASCADE,
        target_date  TEXT    NOT NULL,
        slot_time    TEXT    NOT NULL,
        service_type TEXT,
        slot_token   TEXT,
        is_prime     INTEGER NOT NULL DEFAULT 0
    );

    CREATE INDEX idx_slots_venue_date ON slots (venue_id, target_date);
    CREATE INDEX idx_scans_venue_ts   ON scans (venue_id, scan_ts);

    -- Backs the global daily request cap, which has to survive across runs.
    CREATE TABLE request_log (
        day     TEXT    NOT NULL,
        service TEXT    NOT NULL,
        count   INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY (day, service)
    );
    """,
]

SCHEMA_VERSION = len(MIGRATIONS)


def connect(db_path: Path) -> sqlite3.Connection:
    """Open a connection with the pragmas this tool assumes everywhere."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def migrate(conn: sqlite3.Connection) -> tuple[int, int]:
    """Apply pending migrations. Returns (version_before, version_after)."""
    before = conn.execute("PRAGMA user_version").fetchone()[0]
    if before > SCHEMA_VERSION:
        raise RuntimeError(
            f"Database schema version {before} is newer than this build "
            f"({SCHEMA_VERSION}). Upgrade resy-rank or start a fresh DB."
        )
    for version in range(before, SCHEMA_VERSION):
        log.debug("applying migration %d -> %d", version, version + 1)
        # DDL and the version bump go in one explicit transaction — SQLite
        # rolls back DDL too, so a failed migration leaves no partial schema.
        conn.executescript(
            "BEGIN;\n"
            f"{MIGRATIONS[version]}\n"
            f"PRAGMA user_version = {version + 1};\n"
            "COMMIT;"
        )
    return before, SCHEMA_VERSION


def init_db(db_path: Path) -> tuple[int, int]:
    conn = connect(db_path)
    try:
        return migrate(conn)
    finally:
        conn.close()


# ── venues.csv loading ──────────────────────────────────────────────────────

VENUE_CSV_COLUMNS = {
    "name",
    "neighborhood",
    "cuisine",
    "address",
    "resy_url",
    "resy_venue_id",
    "google_place_id",
}


@dataclass
class LoadResult:
    inserted: int = 0
    updated: int = 0
    skipped: int = 0

    @property
    def total(self) -> int:
        return self.inserted + self.updated + self.skipped


def load_venues_csv(conn: sqlite3.Connection, csv_path: Path) -> LoadResult:
    """Upsert `venues.csv` into the venues table, keyed on (name, neighborhood).

    Re-running is safe: existing rows have their descriptive columns refreshed
    from the CSV, but ids resolved later (resy_venue_id, google_place_id, and
    coordinates) are only filled in when the CSV actually supplies a value.
    """
    if not csv_path.exists():
        raise FileNotFoundError(f"venue file not found: {csv_path}")

    result = LoadResult()
    with csv_path.open(newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise ValueError(f"{csv_path} is empty — expected a header row")
        unknown = set(reader.fieldnames) - VENUE_CSV_COLUMNS
        if unknown:
            log.warning("ignoring unknown column(s) in %s: %s", csv_path, ", ".join(sorted(unknown)))
        if "name" not in reader.fieldnames:
            raise ValueError(f"{csv_path} must have a 'name' column")

        with conn:
            for lineno, row in enumerate(reader, start=2):
                name = (row.get("name") or "").strip()
                if not name:
                    log.warning("%s:%d — blank name, skipping", csv_path.name, lineno)
                    result.skipped += 1
                    continue

                neighborhood = (row.get("neighborhood") or "").strip()
                existing = conn.execute(
                    "SELECT venue_id FROM venues WHERE name = ? AND neighborhood = ?",
                    (name, neighborhood),
                ).fetchone()

                fields = {
                    "cuisine": (row.get("cuisine") or "").strip(),
                    "address": (row.get("address") or "").strip(),
                    "resy_url": (row.get("resy_url") or "").strip() or None,
                    "resy_venue_id": _maybe_int(row.get("resy_venue_id"), csv_path.name, lineno),
                    "google_place_id": (row.get("google_place_id") or "").strip() or None,
                }

                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO venues
                            (name, neighborhood, cuisine, address, resy_url,
                             resy_venue_id, google_place_id)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            name,
                            neighborhood,
                            fields["cuisine"],
                            fields["address"],
                            fields["resy_url"],
                            fields["resy_venue_id"],
                            fields["google_place_id"],
                        ),
                    )
                    result.inserted += 1
                else:
                    conn.execute(
                        """
                        UPDATE venues SET
                            cuisine         = ?,
                            address         = ?,
                            resy_url        = COALESCE(?, resy_url),
                            resy_venue_id   = COALESCE(?, resy_venue_id),
                            google_place_id = COALESCE(?, google_place_id)
                        WHERE venue_id = ?
                        """,
                        (
                            fields["cuisine"],
                            fields["address"],
                            fields["resy_url"],
                            fields["resy_venue_id"],
                            fields["google_place_id"],
                            existing["venue_id"],
                        ),
                    )
                    result.updated += 1
    return result


def _maybe_int(raw: str | None, source: str, lineno: int) -> int | None:
    value = (raw or "").strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        log.warning("%s:%d — ignoring non-numeric resy_venue_id %r", source, lineno, value)
        return None


def venue_counts(conn: sqlite3.Connection) -> dict[str, int]:
    """Small summary used by `init` and (later) `resolve` to report progress."""
    row = conn.execute(
        """
        SELECT COUNT(*)                                            AS total,
               SUM(resy_venue_id IS NULL)                          AS missing_resy,
               SUM(google_place_id IS NULL)                        AS missing_place
        FROM venues
        """
    ).fetchone()
    return {
        "total": row["total"] or 0,
        "missing_resy": row["missing_resy"] or 0,
        "missing_place": row["missing_place"] or 0,
    }
