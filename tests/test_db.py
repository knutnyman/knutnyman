"""Milestone 1 coverage: migrations are idempotent, venue loading upserts."""

from __future__ import annotations

import sqlite3

import pytest

from resy_rank import db


@pytest.fixture()
def conn(tmp_path):
    path = tmp_path / "test.db"
    db.init_db(path)
    c = db.connect(path)
    yield c
    c.close()


def write_csv(tmp_path, text: str):
    path = tmp_path / "venues.csv"
    path.write_text(text, encoding="utf-8")
    return path


def test_migrate_is_idempotent(tmp_path):
    path = tmp_path / "test.db"
    assert db.init_db(path) == (0, db.SCHEMA_VERSION)
    assert db.init_db(path) == (db.SCHEMA_VERSION, db.SCHEMA_VERSION)


def test_rejects_newer_schema(tmp_path):
    path = tmp_path / "test.db"
    db.init_db(path)
    c = db.connect(path)
    c.executescript(f"PRAGMA user_version = {db.SCHEMA_VERSION + 5};")
    with pytest.raises(RuntimeError, match="newer than this build"):
        db.migrate(c)
    c.close()


def test_load_inserts_then_updates(conn, tmp_path):
    csv_path = write_csv(
        tmp_path,
        "name,neighborhood,cuisine,address,resy_url,resy_venue_id,google_place_id\n"
        "Lilia,Williamsburg,Italian,567 Union Ave,https://resy.com/x,,\n",
    )
    result = db.load_venues_csv(conn, csv_path)
    assert (result.inserted, result.updated) == (1, 0)

    # Same natural key, refreshed cuisine — an update, not a duplicate.
    csv_path.write_text(
        "name,neighborhood,cuisine,address,resy_url,resy_venue_id,google_place_id\n"
        "Lilia,Williamsburg,Roman Italian,567 Union Ave,https://resy.com/x,,\n",
        encoding="utf-8",
    )
    result = db.load_venues_csv(conn, csv_path)
    assert (result.inserted, result.updated) == (0, 1)
    row = conn.execute("SELECT COUNT(*) n, MIN(cuisine) c FROM venues").fetchone()
    assert (row["n"], row["c"]) == (1, "Roman Italian")


def test_load_does_not_clobber_resolved_ids(conn, tmp_path):
    """Ids filled in by `resolve` survive a re-run of `init`."""
    csv_path = write_csv(
        tmp_path,
        "name,neighborhood,cuisine,address,resy_url,resy_venue_id,google_place_id\n"
        "Misi,Williamsburg,Italian,329 Kent Ave,,,\n",
    )
    db.load_venues_csv(conn, csv_path)
    with conn:
        conn.execute(
            "UPDATE venues SET resy_venue_id = 42, google_place_id = 'abc' WHERE name = 'Misi'"
        )

    db.load_venues_csv(conn, csv_path)  # CSV still has blanks for both ids
    row = conn.execute("SELECT resy_venue_id, google_place_id FROM venues").fetchone()
    assert (row["resy_venue_id"], row["google_place_id"]) == (42, "abc")


def test_blank_and_bad_rows_are_skipped_not_fatal(conn, tmp_path):
    csv_path = write_csv(
        tmp_path,
        "name,neighborhood,resy_venue_id\n"
        ",Nowhere,1\n"
        "Cote,Flatiron,not-a-number\n",
    )
    result = db.load_venues_csv(conn, csv_path)
    assert (result.inserted, result.skipped) == (1, 1)
    row = conn.execute("SELECT name, resy_venue_id FROM venues").fetchone()
    assert (row["name"], row["resy_venue_id"]) == ("Cote", None)


def test_header_required(conn, tmp_path):
    csv_path = write_csv(tmp_path, "")
    with pytest.raises(ValueError, match="expected a header row"):
        db.load_venues_csv(conn, csv_path)

    csv_path = write_csv(tmp_path, "restaurant,neighborhood\nLilia,Williamsburg\n")
    with pytest.raises(ValueError, match="must have a 'name' column"):
        db.load_venues_csv(conn, csv_path)


def test_venue_counts(conn, tmp_path):
    csv_path = write_csv(
        tmp_path,
        "name,neighborhood,resy_venue_id,google_place_id\n"
        "A,X,1,place-a\n"
        "B,Y,,\n",
    )
    db.load_venues_csv(conn, csv_path)
    assert db.venue_counts(conn) == {"total": 2, "missing_resy": 1, "missing_place": 1}


def test_foreign_keys_enforced(conn):
    with pytest.raises(sqlite3.IntegrityError):
        with conn:
            conn.execute(
                "INSERT INTO scans (venue_id, scan_ts, target_date, party_size) "
                "VALUES (999, '2026-08-14T12:00:00', '2026-09-12', 2)"
            )
