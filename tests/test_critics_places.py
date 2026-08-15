"""Critic matching and Places caching/budgeting. No network."""

from __future__ import annotations

import asyncio
from datetime import datetime, timedelta, timezone

import httpx
import pytest

from resy_rank import critics, db, places
from resy_rank.config import Settings

KNOWN = {"michelin_star", "michelin_bib", "eater_38", "nyt", "infatuation"}


@pytest.fixture()
def conn(tmp_path):
    path = tmp_path / "t.db"
    db.init_db(path)
    c = db.connect(path)
    with c:
        c.execute("INSERT INTO venues (venue_id, name, address) VALUES (1, 'Cote', '16 W 22nd')")
        c.execute("INSERT INTO venues (venue_id, name) VALUES (2, 'Le Bernardin')")
        c.execute("INSERT INTO venues (venue_id, name) VALUES (3, 'Lilia')")
    yield c
    c.close()


def write_critics(tmp_path, body: str):
    path = tmp_path / "critics.csv"
    path.write_text("venue_name,source,tier,url\n" + body, encoding="utf-8")
    return path


# ── critics ─────────────────────────────────────────────────────────────────


def test_matches_suffix_and_typo(conn, tmp_path):
    path = write_critics(
        tmp_path,
        "COTE Korean Steakhouse,michelin_star,1,http://x\n"
        "Le Bernadin,michelin_star,3,http://x\n",  # deliberate typo
    )
    report = critics.match_critics(
        critics.load_critics_csv(path), db.fetch_venues(conn), threshold=88, known_sources=KNOWN
    )
    assert {m.venue_name for m in report.matched} == {"Cote", "Le Bernardin"}
    assert not report.unmatched


def test_reports_unmatched_with_closest_candidate(conn, tmp_path):
    path = write_critics(tmp_path, "Totally Fake Restaurant,eater_38,,http://x\n")
    report = critics.match_critics(
        critics.load_critics_csv(path), db.fetch_venues(conn), threshold=88, known_sources=KNOWN
    )
    assert not report.matched
    assert len(report.unmatched) == 1
    row, closest, score = report.unmatched[0]
    assert row.venue_name == "Totally Fake Restaurant"
    assert closest  # a suggestion is offered so the CSV can be fixed
    assert score < 88


def test_unknown_source_is_flagged(conn, tmp_path):
    path = write_critics(tmp_path, "Lilia,zagat,,http://x\n")
    report = critics.match_critics(
        critics.load_critics_csv(path), db.fetch_venues(conn), threshold=88, known_sources=KNOWN
    )
    assert [r.source for r in report.unknown_sources] == ["zagat"]
    assert not report.ok


def test_threshold_is_respected(conn, tmp_path):
    path = write_critics(tmp_path, "Lilias,eater_38,,http://x\n")
    rows = critics.load_critics_csv(path)
    venues = db.fetch_venues(conn)
    assert critics.match_critics(rows, venues, threshold=99, known_sources=KNOWN).unmatched
    assert critics.match_critics(rows, venues, threshold=80, known_sources=KNOWN).matched


def test_save_flags_is_idempotent(conn, tmp_path):
    path = write_critics(tmp_path, "Lilia,eater_38,,http://x\n")
    report = critics.match_critics(
        critics.load_critics_csv(path), db.fetch_venues(conn), threshold=88, known_sources=KNOWN
    )
    critics.save_flags(conn, report.matched)
    critics.save_flags(conn, report.matched)
    assert conn.execute("SELECT COUNT(*) c FROM critic_flags").fetchone()["c"] == 1
    assert critics.flags_by_venue(conn) == {3: ["eater_38"]}


def test_blank_rows_skipped_and_bad_header_rejected(conn, tmp_path):
    path = write_critics(tmp_path, ",eater_38,,\nLilia,,,\n")
    assert critics.load_critics_csv(path) == []

    bad = tmp_path / "bad.csv"
    bad.write_text("name,source\nLilia,nyt\n", encoding="utf-8")
    with pytest.raises(ValueError, match="missing column"):
        critics.load_critics_csv(bad)


# ── places: caching ─────────────────────────────────────────────────────────


def test_is_stale():
    now = datetime(2026, 8, 15, tzinfo=timezone.utc)
    assert places.is_stale(None, 180, now) is True
    assert places.is_stale("garbage", 180, now) is True
    assert places.is_stale((now - timedelta(days=179)).isoformat(), 180, now) is False
    assert places.is_stale((now - timedelta(days=181)).isoformat(), 180, now) is True
    # A naive timestamp is treated as UTC rather than crashing.
    assert places.is_stale((now - timedelta(days=1)).replace(tzinfo=None).isoformat(), 180, now) is False


def test_only_stale_venues_are_refetched(conn):
    fresh = datetime.now(timezone.utc).isoformat()
    with conn:
        conn.execute(
            "INSERT INTO google_meta (venue_id, rating, user_rating_count, fetched_at) "
            "VALUES (1, 4.5, 100, ?)",
            (fresh,),
        )
    pending = places.venues_needing_enrichment(conn, 180)
    assert {r["venue_id"] for r in pending} == {2, 3}  # venue 1 is cached and fresh

    # --refresh overrides the cache entirely.
    assert len(places.venues_needing_enrichment(conn, 180, refresh=True)) == 3


def test_save_details_upserts(conn):
    details = places.PlaceDetails(
        place_id="abc", rating=4.6, user_rating_count=1200, price_level="PRICE_LEVEL_EXPENSIVE",
        lat=40.7, lng=-74.0,
    )
    places.save_details(conn, 1, details)
    places.save_details(conn, 1, places.PlaceDetails(place_id="abc", rating=4.9, user_rating_count=1300))
    row = conn.execute("SELECT rating, user_rating_count FROM google_meta WHERE venue_id=1").fetchone()
    assert (row["rating"], row["user_rating_count"]) == (4.9, 1300)
    assert conn.execute("SELECT COUNT(*) c FROM google_meta").fetchone()["c"] == 1
    venue = conn.execute("SELECT google_place_id, lat FROM venues WHERE venue_id=1").fetchone()
    assert venue["google_place_id"] == "abc" and venue["lat"] == 40.7


# ── places: parsing ─────────────────────────────────────────────────────────


def test_parse_details_full():
    parsed = places.parse_details(
        {
            "id": "xyz",
            "displayName": {"text": "Lilia", "languageCode": "en"},
            "rating": 4.7,
            "userRatingCount": 3100,
            "priceLevel": "PRICE_LEVEL_EXPENSIVE",
            "location": {"latitude": 40.71, "longitude": -73.95},
            "websiteUri": "https://lilianewyork.com",
        }
    )
    assert parsed.display_name == "Lilia"
    assert (parsed.rating, parsed.user_rating_count) == (4.7, 3100)
    assert (parsed.lat, parsed.lng) == (40.71, -73.95)


def test_parse_details_tolerates_missing_fields():
    """A place with no ratings yet must parse, not explode."""
    parsed = places.parse_details({"id": "xyz"})
    assert parsed.place_id == "xyz"
    assert parsed.rating is None and parsed.user_rating_count is None


@pytest.mark.parametrize("payload", [None, {}, {"id": ""}, {"id": 5}, "string", []])
def test_parse_details_rejects_unusable(payload):
    assert places.parse_details(payload) is None


@pytest.mark.parametrize(
    "payload,expected",
    [
        ({"places": [{"id": "abc"}]}, "abc"),
        ({"places": []}, None),
        ({"places": "nope"}, None),
        ({}, None),
        (None, None),
        ({"places": [{"noid": 1}]}, None),
    ],
)
def test_parse_search_text(payload, expected):
    assert places.parse_search_text(payload) == expected


# ── places: monthly budget ──────────────────────────────────────────────────


def test_monthly_budget_caps_enterprise_calls(conn):
    async def go(cap, n):
        budget = places.MonthlyBudget(conn, cap)
        for _ in range(n):
            await budget.consume(places.SERVICE_DETAILS)

    asyncio.run(go(3, 3))
    assert places.MonthlyBudget(conn, 3).remaining(places.SERVICE_DETAILS) == 0
    with pytest.raises(places.MonthlyCapReached, match="free tier"):
        asyncio.run(go(3, 1))


def test_monthly_budget_is_per_service(conn):
    async def go():
        budget = places.MonthlyBudget(conn, 2)
        await budget.consume(places.SERVICE_DETAILS)
        await budget.consume(places.SERVICE_TEXT)
        return budget

    budget = asyncio.run(go())
    assert budget.used(places.SERVICE_DETAILS) == 1
    assert budget.used(places.SERVICE_TEXT) == 1


def test_default_cap_is_the_free_tier():
    """Guard against a default that silently costs money."""
    assert Settings(_env_file=None).monthly_enrichment_cap == 1000


def test_dry_run_places_makes_no_calls_and_needs_no_key():
    intents: list[str] = []

    async def go():
        async with places.PlacesClient(
            Settings(_env_file=None), dry_run=True, on_intent=intents.append
        ) as client:
            assert await client.resolve_place_id("Lilia", "567 Union Ave") is None
            assert await client.fetch_details("abc") is None
            return client

    client = asyncio.run(go())
    assert client.billable_total == 0
    assert len(intents) == 2
    assert any("Enterprise" in i or places.SERVICE_DETAILS in i for i in intents)


def test_billable_calls_are_counted():
    settings = Settings(_env_file=None, GOOGLE_PLACES_API_KEY="k")

    def handler(request: httpx.Request) -> httpx.Response:
        if "searchText" in str(request.url):
            return httpx.Response(200, json={"places": [{"id": "abc"}]})
        return httpx.Response(200, json={"id": "abc", "rating": 4.6, "userRatingCount": 100})

    async def go():
        client = places.PlacesClient(settings)
        client._client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
        await client.resolve_place_id("Lilia")
        await client.fetch_details("abc")
        await client._client.aclose()
        return client

    client = asyncio.run(go())
    assert client.calls[places.SERVICE_TEXT] == 1
    assert client.calls[places.SERVICE_DETAILS] == 1
    assert client.billable_total == 2


def test_live_places_requires_api_key():
    async def go():
        async with places.PlacesClient(Settings(_env_file=None)):
            pass

    with pytest.raises(RuntimeError, match="GOOGLE_PLACES_API_KEY"):
        asyncio.run(go())
