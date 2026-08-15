"""Milestone 2: defensive parsing, politeness, and the read-only guarantee.

No network. The /4/find fixtures below are hand-written to match the shapes
seen in the wild, plus the malformed variants the parser must survive.
"""

from __future__ import annotations

import asyncio
from datetime import time

import httpx
import pytest

from resy_rank import db, resy
from resy_rank.config import Settings
from resy_rank.resy import (
    DailyBudget,
    DailyCapReached,
    Politeness,
    PrimeWindow,
    ResyClient,
    ScanOutcome,
    parse_find_response,
    parse_venue_search,
    persist_scan,
)

# 2026-09-12 is a Saturday, so 18:30-20:30 slots on it are prime.
SATURDAY = "2026-09-12"
MONDAY = "2026-09-14"

PRIME = PrimeWindow(days=frozenset({3, 4, 5}), start=time(18, 30), end=time(20, 30))


def find_payload(slots: list[dict]) -> dict:
    return {"results": {"venues": [{"venue": {"id": {"resy": 42}}, "slots": slots}]}}


def slot(start: str, type_: str = "Dining Room", token: str | None = "tok") -> dict:
    config = {"type": type_}
    if token is not None:
        config["token"] = token
    return {"config": config, "date": {"start": start, "end": start}}


@pytest.fixture()
def settings():
    return Settings(_env_file=None, RESY_API_KEY="k", RESY_AUTH_TOKEN="t")


# ── parsing ─────────────────────────────────────────────────────────────────


def test_parses_time_service_and_token():
    payload = find_payload([slot(f"{SATURDAY} 19:00:00", "Dining Room", "abc123")])
    slots = parse_find_response(payload, SATURDAY, PRIME)
    assert len(slots) == 1
    assert slots[0].slot_time == "19:00"
    assert slots[0].service_type == "Dining Room"
    assert slots[0].slot_token == "abc123"
    assert slots[0].slot_date == SATURDAY


def test_accepts_iso_timestamps():
    payload = find_payload([slot(f"{SATURDAY}T19:30:00")])
    assert parse_find_response(payload, SATURDAY, PRIME)[0].slot_time == "19:30"


def test_falls_back_to_config_id_when_no_token():
    payload = find_payload([{"config": {"type": "Bar", "id": 998877}, "date": {"start": f"{SATURDAY} 18:00:00"}}])
    assert parse_find_response(payload, SATURDAY, PRIME)[0].slot_token == "998877"


def test_prime_classification():
    payload = find_payload(
        [
            slot(f"{SATURDAY} 18:00:00"),  # too early
            slot(f"{SATURDAY} 18:30:00"),  # boundary, inclusive
            slot(f"{SATURDAY} 20:30:00"),  # boundary, inclusive
            slot(f"{SATURDAY} 21:00:00"),  # too late
        ]
    )
    by_time = {s.slot_time: s.is_prime for s in parse_find_response(payload, SATURDAY, PRIME)}
    assert by_time == {"18:00": False, "18:30": True, "20:30": True, "21:00": False}


def test_monday_is_never_prime():
    payload = find_payload([slot(f"{MONDAY} 19:00:00")])
    assert parse_find_response(payload, MONDAY, PRIME)[0].is_prime is False


@pytest.mark.parametrize(
    "payload",
    [
        {},
        {"results": None},
        {"results": {}},
        {"results": {"venues": None}},
        {"results": {"venues": "nope"}},
        {"results": {"venues": [None, 7, "x"]}},
        {"results": {"venues": [{"slots": None}]}},
        {"results": {"venues": [{"slots": "nope"}]}},
        {"results": {"venues": [{"slots": [None, 3]}]}},
        {"results": {"venues": [{}]}},
        [],
        None,
        "unexpected string body",
    ],
)
def test_malformed_payloads_return_empty_not_crash(payload):
    assert parse_find_response(payload, SATURDAY, PRIME) == []


def test_skips_bad_slots_but_keeps_good_ones():
    payload = find_payload(
        [
            {"config": "not-a-dict", "date": {"start": f"{SATURDAY} 19:00:00"}},
            {"date": {"start": "garbage"}},
            {"config": {}, "date": "not-a-dict"},
            slot(f"{SATURDAY} 20:00:00", "Bar", "keepme"),
        ]
    )
    slots = parse_find_response(payload, SATURDAY, PRIME)
    # The first slot has an unusable config but a valid time, so it survives
    # with no service type; the two with no usable time are dropped.
    assert [(s.slot_time, s.slot_token) for s in slots] == [("19:00", None), ("20:00", "keepme")]


def test_deduplicates_identical_slots():
    payload = find_payload([slot(f"{SATURDAY} 19:00:00"), slot(f"{SATURDAY} 19:00:00")])
    assert len(parse_find_response(payload, SATURDAY, PRIME)) == 1


def test_same_time_different_service_both_kept():
    payload = find_payload(
        [slot(f"{SATURDAY} 19:00:00", "Dining Room", "a"), slot(f"{SATURDAY} 19:00:00", "Bar", "b")]
    )
    assert len(parse_find_response(payload, SATURDAY, PRIME)) == 2


def test_non_string_service_type_is_coerced():
    payload = find_payload([{"config": {"type": 7}, "date": {"start": f"{SATURDAY} 19:00:00"}}])
    assert parse_find_response(payload, SATURDAY, PRIME)[0].service_type == "7"


def test_venue_search_parsing_handles_both_shapes():
    payload = {
        "search": {
            "hits": [
                {"venue": {"id": {"resy": 123}, "name": "Lilia", "location": {"locality": "Brooklyn"}}},
                {"id": {"resy": "456"}, "name": "Misi", "location": {"locality": "Brooklyn"}},
                {"id": {"resy": None}, "name": "Broken"},
                "not-a-dict",
            ]
        }
    }
    hits = parse_venue_search(payload)
    assert [(h["resy_venue_id"], h["name"]) for h in hits] == [(123, "Lilia"), (456, "Misi")]


@pytest.mark.parametrize("payload", [{}, {"search": None}, {"search": {"hits": None}}, None])
def test_venue_search_malformed(payload):
    assert parse_venue_search(payload) == []


# ── read-only guarantee ─────────────────────────────────────────────────────


def test_booking_endpoints_are_not_allowlisted():
    for url in (
        "https://api.resy.com/3/book",
        "https://api.resy.com/3/details",
        "https://api.resy.com/3/cancel",
    ):
        assert url not in resy._ALLOWED_ENDPOINTS


def test_request_to_unlisted_endpoint_raises(settings):
    async def go():
        client = ResyClient(settings, dry_run=True)
        with pytest.raises(RuntimeError, match="read-only"):
            await client._request("POST", "https://api.resy.com/3/book")

    asyncio.run(go())


# ── politeness ──────────────────────────────────────────────────────────────


def test_politeness_spaces_requests_and_caps_concurrency():
    async def go():
        politeness = Politeness(concurrency=3, delay_seconds=0.05)
        in_flight = 0
        peak = 0
        starts: list[float] = []

        async def worker():
            nonlocal in_flight, peak
            async with politeness.slot():
                starts.append(asyncio.get_running_loop().time())
                in_flight += 1
                peak = max(peak, in_flight)
                await asyncio.sleep(0.01)
                in_flight -= 1

        await asyncio.gather(*(worker() for _ in range(6)))
        return peak, starts

    peak, starts = asyncio.run(go())
    assert peak <= 3
    gaps = [b - a for a, b in zip(starts, starts[1:])]
    assert all(gap >= 0.04 for gap in gaps), gaps


def test_daily_budget_enforced_and_persisted(tmp_path):
    path = tmp_path / "t.db"
    db.init_db(path)
    conn = db.connect(path)

    async def go(cap: int, n: int):
        budget = DailyBudget(conn, cap)
        for _ in range(n):
            await budget.consume("resy")

    asyncio.run(go(3, 3))
    assert DailyBudget(conn, 3).used("resy") == 3
    assert DailyBudget(conn, 3).remaining("resy") == 0

    with pytest.raises(DailyCapReached):
        asyncio.run(go(3, 1))

    # A brand new budget object (i.e. a separate run) sees the same tally.
    conn.close()
    conn2 = db.connect(path)
    assert DailyBudget(conn2, 3).remaining("resy") == 0
    conn2.close()


# ── retries ─────────────────────────────────────────────────────────────────


def _client_with_transport(settings, handler) -> ResyClient:
    client = ResyClient(settings)
    client._client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return client


@pytest.fixture()
def no_backoff_sleep(monkeypatch):
    """Collapse backoff waits so retry tests run instantly."""
    real_sleep = asyncio.sleep
    waits: list[float] = []

    async def fake_sleep(seconds: float = 0, *args, **kwargs):
        waits.append(seconds)
        await real_sleep(0)

    monkeypatch.setattr(asyncio, "sleep", fake_sleep)
    return waits


def test_retries_on_429_then_succeeds(settings, no_backoff_sleep):
    calls = {"n": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        calls["n"] += 1
        if calls["n"] < 3:
            return httpx.Response(429, headers={"Retry-After": "0"})
        return httpx.Response(200, json=find_payload([slot(f"{SATURDAY} 19:00:00")]))

    async def go():
        client = _client_with_transport(settings, handler)
        outcome = await client.find(
            venue_id=1, resy_venue_id=42, target_date=SATURDAY, party_size=2
        )
        await client._client.aclose()
        return outcome

    outcome = asyncio.run(go())
    assert calls["n"] == 3
    assert outcome.ok and len(outcome.slots) == 1


def test_expired_token_raises_immediately(settings):
    calls = {"n": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        calls["n"] += 1
        return httpx.Response(401)

    async def go():
        client = _client_with_transport(settings, handler)
        with pytest.raises(resy.ResyUnauthorized, match="expired"):
            await client.find(venue_id=1, resy_venue_id=42, target_date=SATURDAY, party_size=2)
        await client._client.aclose()

    asyncio.run(go())
    assert calls["n"] == 1  # no retry storm against a dead token


def test_venue_failure_is_isolated_not_fatal(no_backoff_sleep):
    # Zero spacing delay so the only sleeps recorded are the retry backoffs.
    settings = Settings(
        _env_file=None, RESY_API_KEY="k", RESY_AUTH_TOKEN="t", REQUEST_DELAY_SECONDS=0
    )

    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500)

    async def go():
        client = _client_with_transport(settings, handler)
        outcome = await client.find(
            venue_id=1, resy_venue_id=42, target_date=SATURDAY, party_size=2
        )
        await client._client.aclose()
        return outcome

    outcome = asyncio.run(go())
    assert not outcome.ok
    assert "failed after" in outcome.error

    # Backoff grows exponentially: 2s, 4s, 8s, 16s (plus <0.5s of jitter).
    backoffs = [w for w in no_backoff_sleep if w >= 2]
    assert len(backoffs) == 4
    for expected, actual in zip([2, 4, 8, 16], backoffs):
        assert expected <= actual < expected + 0.5


def test_dry_run_makes_no_requests(settings):
    intents: list[str] = []

    async def go():
        async with ResyClient(settings, dry_run=True, on_intent=intents.append) as client:
            assert client._client is None  # no HTTP client is even constructed
            outcome = await client.find(
                venue_id=1, resy_venue_id=42, target_date=SATURDAY, party_size=2
            )
        return outcome

    outcome = asyncio.run(go())
    assert outcome.slots == [] and outcome.raw_json is None
    assert len(intents) == 1 and "/4/find" in intents[0]


def test_dry_run_needs_no_credentials():
    bare = Settings(_env_file=None)

    async def go():
        async with ResyClient(bare, dry_run=True, on_intent=lambda _: None) as client:
            await client.find(venue_id=1, resy_venue_id=42, target_date=SATURDAY, party_size=2)

    asyncio.run(go())  # must not raise about missing credentials


def test_live_run_requires_credentials():
    bare = Settings(_env_file=None)

    async def go():
        async with ResyClient(bare):
            pass

    with pytest.raises(RuntimeError, match="RESY_API_KEY"):
        asyncio.run(go())


# ── persistence ─────────────────────────────────────────────────────────────


def test_persist_scan_writes_raw_blob_and_slots(tmp_path):
    path = tmp_path / "t.db"
    db.init_db(path)
    conn = db.connect(path)
    with conn:
        conn.execute("INSERT INTO venues (name, resy_venue_id) VALUES ('Lilia', 42)")

    payload = find_payload([slot(f"{SATURDAY} 19:00:00"), slot(f"{SATURDAY} 21:00:00", "Bar", "b")])
    outcome = ScanOutcome(
        venue_id=1,
        resy_venue_id=42,
        target_date=SATURDAY,
        party_size=2,
        slots=parse_find_response(payload, SATURDAY, PRIME),
        raw_json='{"results":{}}',
    )
    scan_id = persist_scan(conn, outcome, "2026-08-14T12:00:00+00:00")

    rows = conn.execute(
        "SELECT slot_time, service_type, slot_token, is_prime FROM slots WHERE scan_id = ? "
        "ORDER BY slot_time",
        (scan_id,),
    ).fetchall()
    assert [tuple(r) for r in rows] == [
        ("19:00", "Dining Room", "tok", 1),
        ("21:00", "Bar", "b", 0),
    ]
    raw = conn.execute("SELECT raw_json FROM scans WHERE scan_id = ?", (scan_id,)).fetchone()
    assert raw["raw_json"] == '{"results":{}}'
    conn.close()


def geo_payload(venues: list[dict]) -> dict:
    return {"results": {"venues": venues}}


def geo_venue(vid: int, name: str, slots: list[dict], neighborhood: str = "Nolita") -> dict:
    return {
        "venue": {
            "id": {"resy": vid},
            "name": name,
            "neighborhood": neighborhood,
            "location": {"latitude": 40.72, "longitude": -73.99},
        },
        "slots": slots,
    }


def test_parse_find_venues_groups_by_venue():
    payload = geo_payload(
        [
            geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00")]),
            geo_venue(2, "Misi", [slot(f"{SATURDAY} 18:00:00"), slot(f"{SATURDAY} 20:00:00")]),
        ]
    )
    venues = resy.parse_find_venues(payload, SATURDAY, PRIME)
    assert [(v.resy_venue_id, v.name, len(v.slots)) for v in venues] == [
        (1, "Lilia", 1),
        (2, "Misi", 2),
    ]
    assert venues[0].lat == 40.72 and venues[0].neighborhood == "Nolita"


def test_parse_find_venues_drops_entries_with_no_id():
    payload = geo_payload(
        [{"venue": {"name": "Nameless"}, "slots": []}, geo_venue(9, "Real", [])]
    )
    venues = resy.parse_find_venues(payload, SATURDAY, PRIME)
    assert [v.resy_venue_id for v in venues] == [9]


def test_parse_find_venues_accepts_string_ids():
    payload = geo_payload([{"venue": {"id": {"resy": "77"}, "name": "X"}, "slots": []}])
    assert resy.parse_find_venues(payload, SATURDAY, PRIME)[0].resy_venue_id == 77


def test_merge_geo_results_unions_slots_across_anchors():
    a = resy.parse_find_venues(
        geo_payload([geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00", "Dining Room", "a")])]),
        SATURDAY,
        PRIME,
    )
    b = resy.parse_find_venues(
        geo_payload(
            [
                geo_venue(1, "Lilia", [slot(f"{SATURDAY} 20:00:00", "Bar", "b")]),
                geo_venue(2, "Misi", []),
            ]
        ),
        SATURDAY,
        PRIME,
    )
    merged = {v.resy_venue_id: v for v in resy.merge_geo_results([a, b])}
    assert len(merged) == 2
    assert [s.slot_time for s in merged[1].slots] == ["19:00", "20:00"]


def test_merge_geo_results_does_not_duplicate_identical_slots():
    payload = geo_payload([geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00")])])
    a = resy.parse_find_venues(payload, SATURDAY, PRIME)
    b = resy.parse_find_venues(payload, SATURDAY, PRIME)
    merged = resy.merge_geo_results([a, b])
    assert len(merged) == 1 and len(merged[0].slots) == 1


def _geo_db(tmp_path):
    path = tmp_path / "geo.db"
    db.init_db(path)
    return db.connect(path)


def test_geo_sweep_discovers_new_venues(tmp_path):
    conn = _geo_db(tmp_path)
    venues = resy.parse_find_venues(
        geo_payload([geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00")])]), SATURDAY, PRIME
    )
    counts = resy.persist_geo_sweep(
        conn, target_date=SATURDAY, party_size=2, venues=venues, scan_ts="2026-08-14T12:00:00+00:00"
    )
    assert counts == {"available": 1, "unavailable": 0}
    row = conn.execute("SELECT name, resy_venue_id, in_geo_scope FROM venues").fetchone()
    assert (row["name"], row["resy_venue_id"], row["in_geo_scope"]) == ("Lilia", 1, 1)
    conn.close()


def test_geo_sweep_links_seed_row_instead_of_duplicating(tmp_path):
    """A sweep hit resolves a seed venue for free — no extra lookup call."""
    conn = _geo_db(tmp_path)
    with conn:
        conn.execute("INSERT INTO venues (name, neighborhood) VALUES ('Lilia', 'Williamsburg')")

    venues = resy.parse_find_venues(
        geo_payload([geo_venue(1, "Lilia", [])]), SATURDAY, PRIME
    )
    resy.persist_geo_sweep(
        conn, target_date=SATURDAY, party_size=2, venues=venues, scan_ts="2026-08-14T12:00:00+00:00"
    )
    rows = conn.execute("SELECT name, neighborhood, resy_venue_id FROM venues").fetchall()
    assert len(rows) == 1
    assert (rows[0]["neighborhood"], rows[0]["resy_venue_id"]) == ("Williamsburg", 1)
    conn.close()


def test_geo_sweep_records_absence_for_in_scope_venues(tmp_path):
    """The absent venues are the scarcity signal — they must get scan rows."""
    conn = _geo_db(tmp_path)
    first = resy.parse_find_venues(
        geo_payload([geo_venue(1, "Lilia", []), geo_venue(2, "Misi", [])]), SATURDAY, PRIME
    )
    resy.persist_geo_sweep(
        conn, target_date=SATURDAY, party_size=2, venues=first, scan_ts="2026-08-14T12:00:00+00:00"
    )

    # Next sweep: only Lilia has availability. Misi must still be observed.
    second = resy.parse_find_venues(
        geo_payload([geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00")])]), SATURDAY, PRIME
    )
    counts = resy.persist_geo_sweep(
        conn, target_date="2026-09-19", party_size=2, venues=second,
        scan_ts="2026-08-15T12:00:00+00:00",
    )
    assert counts == {"available": 1, "unavailable": 1}

    misi = conn.execute(
        "SELECT COUNT(*) c FROM scans WHERE venue_id = 2 AND target_date = '2026-09-19'"
    ).fetchone()
    assert misi["c"] == 1  # observed, with no slots — i.e. booked out
    assert conn.execute("SELECT COUNT(*) c FROM slots WHERE venue_id = 2").fetchone()["c"] == 0
    conn.close()


def test_geo_sweep_does_not_mark_out_of_scope_venues_absent(tmp_path):
    """A seed venue never seen by a sweep is unknown, not booked out."""
    conn = _geo_db(tmp_path)
    with conn:
        conn.execute("INSERT INTO venues (name, resy_venue_id) VALUES ('Brooklyn Spot', 999)")

    venues = resy.parse_find_venues(geo_payload([geo_venue(1, "Lilia", [])]), SATURDAY, PRIME)
    counts = resy.persist_geo_sweep(
        conn, target_date=SATURDAY, party_size=2, venues=venues, scan_ts="2026-08-14T12:00:00+00:00"
    )
    assert counts["unavailable"] == 0
    assert conn.execute("SELECT COUNT(*) c FROM scans WHERE venue_id = 1").fetchone()["c"] == 0
    conn.close()


def test_geo_sweep_stores_only_the_venue_slice_of_the_payload(tmp_path):
    """Not the whole city blob repeated per venue."""
    conn = _geo_db(tmp_path)
    venues = resy.parse_find_venues(
        geo_payload(
            [geo_venue(1, "Lilia", [slot(f"{SATURDAY} 19:00:00")]), geo_venue(2, "Misi", [])]
        ),
        SATURDAY,
        PRIME,
    )
    resy.persist_geo_sweep(
        conn, target_date=SATURDAY, party_size=2, venues=venues, scan_ts="2026-08-14T12:00:00+00:00"
    )
    for row in conn.execute("SELECT raw_json FROM scans"):
        assert "Lilia" not in row["raw_json"] or "Misi" not in row["raw_json"]
    conn.close()


def test_geo_pagination_stops_when_no_new_venues(settings):
    """The page parameter is undocumented; a repeating page must not loop forever."""
    calls = {"n": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        calls["n"] += 1
        return httpx.Response(200, json=geo_payload([geo_venue(1, "Lilia", [])]))

    async def go():
        client = _client_with_transport(settings, handler)
        found = await client.find_geo(lat=40.7, long=-74.0, target_date=SATURDAY, party_size=2)
        await client._client.aclose()
        return found

    found = asyncio.run(go())
    assert len(found) == 1
    assert calls["n"] == 2  # page 1 discovers, page 2 repeats and stops the loop


def test_scan_with_no_availability_still_records_an_observation(tmp_path):
    """A booked-out venue must leave a scan row — that's the scarcity signal."""
    path = tmp_path / "t.db"
    db.init_db(path)
    conn = db.connect(path)
    with conn:
        conn.execute("INSERT INTO venues (name, resy_venue_id) VALUES ('Carbone', 7)")

    outcome = ScanOutcome(
        venue_id=1, resy_venue_id=7, target_date=SATURDAY, party_size=2, slots=[], raw_json="{}"
    )
    persist_scan(conn, outcome, "2026-08-14T12:00:00+00:00")
    assert conn.execute("SELECT COUNT(*) c FROM scans").fetchone()["c"] == 1
    assert conn.execute("SELECT COUNT(*) c FROM slots").fetchone()["c"] == 0
    conn.close()
