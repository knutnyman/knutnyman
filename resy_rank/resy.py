"""Resy client — strictly read-only.

Resy has no public API. Requests carry the public web api_key:

    Authorization: ResyAPI api_key="<RESY_API_KEY>"

and optionally a personal session token (X-Resy-Auth-Token), which
availability lookups appear not to need.

/4/find is a POST taking a JSON body — verified against a captured browser
request, not the GET the shape of this endpoint is often documented as.

Two safety properties this module is built to hold:

1. **Read-only.** Every outbound URL is checked against `_ALLOWED_ENDPOINTS`
   before the request goes out. Booking endpoints (/3/details, /3/book) are
   not in it, so an accidental call raises rather than reserves a table.
2. **Polite.** Concurrency is capped, requests are spaced by a configurable
   delay, 429/5xx get exponential backoff, and a global daily cap persisted in
   SQLite stops the day's scanning even across separate runs. This is a slow
   background scanner and must never behave like a drop-time sniper.
"""

from __future__ import annotations

import asyncio
import json
import logging
import random
import sqlite3
import time as _time
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from datetime import date as date_cls, datetime, time, timezone
from typing import Any, Iterable

import httpx

from resy_rank.config import Settings

log = logging.getLogger(__name__)

RESY_API_BASE = "https://api.resy.com"
FIND_URL = f"{RESY_API_BASE}/4/find"
VENUE_SEARCH_URL = f"{RESY_API_BASE}/3/venuesearch/search"

# Read-only allowlist. Adding a booking endpoint here would be a bug, not a
# feature — this tool must never book, hold, modify, or cancel a reservation.
_ALLOWED_ENDPOINTS: frozenset[str] = frozenset({FIND_URL, VENUE_SEARCH_URL})

_RETRY_STATUS = frozenset({429, 500, 502, 503, 504})


class DailyCapReached(RuntimeError):
    """Raised when the configured daily request cap is exhausted."""


class ResyUnauthorized(RuntimeError):
    """Raised on 401/403 — the browser-session token has almost certainly expired."""


# ── Parsing ─────────────────────────────────────────────────────────────────
# The /4/find response shape is undocumented and changes without notice.
# Everything below assumes nothing: any missing or wrong-typed key means "skip
# this venue/slot and log it", never a crash. The raw payload is persisted by
# the caller so historical scans can be re-parsed if the shape shifts.


@dataclass(frozen=True)
class ParsedSlot:
    slot_time: str  # "HH:MM", local to the venue
    service_type: str | None
    slot_token: str | None
    is_prime: bool
    slot_date: str | None = None  # "YYYY-MM-DD" when the payload carried one


@dataclass
class PrimeWindow:
    """Which (day, time) combinations count as prime for the scarcity index."""

    days: frozenset[int]
    start: time
    end: time

    @classmethod
    def from_settings(cls, settings: Settings) -> "PrimeWindow":
        return cls(
            days=frozenset(settings.prime_days),
            start=settings.prime_start,
            end=settings.prime_end,
        )

    def covers(self, day: date_cls, slot_time: time) -> bool:
        return day.weekday() in self.days and self.start <= slot_time <= self.end


def _parse_timestamp(value: Any) -> tuple[str | None, str | None]:
    """Return (YYYY-MM-DD, HH:MM) from Resy's loosely-specified timestamps.

    Seen in the wild: "2026-09-12 19:00:00" and ISO-8601 with a "T". Anything
    else returns (None, None) and the caller skips the slot.
    """
    if not isinstance(value, str):
        return None, None
    text = value.strip().replace("T", " ")
    if not text:
        return None, None
    parts = text.split(" ")
    day_part = parts[0]
    time_part = parts[1] if len(parts) > 1 else ""
    try:
        day = date_cls.fromisoformat(day_part).isoformat()
    except ValueError:
        return None, None
    hhmm = time_part[:5]
    if len(hhmm) != 5 or hhmm[2] != ":" or not (hhmm[:2] + hhmm[3:]).isdigit():
        return day, None
    try:
        time.fromisoformat(hhmm)
    except ValueError:
        return day, None
    return day, hhmm


def _iter_venue_entries(payload: Any) -> Iterable[tuple[dict, int]]:
    """Yield (venue_entry, index) from a /4/find payload, tolerating bad shapes."""
    results = payload.get("results") if isinstance(payload, dict) else None
    if not isinstance(results, dict):
        log.debug("find payload has no dict 'results' key; nothing to parse")
        return

    venues = results.get("venues")
    if not isinstance(venues, list):
        log.debug("find payload 'results.venues' is %s, not a list", type(venues).__name__)
        return

    for index, venue in enumerate(venues):
        if not isinstance(venue, dict):
            log.warning("skipping venue #%d: expected dict, got %s", index, type(venue).__name__)
            continue
        yield venue, index


def _iter_slot_dicts(payload: Any) -> Iterable[tuple[dict, str]]:
    """Yield (slot_dict, provenance) pairs across every venue in the payload."""
    for venue, index in _iter_venue_entries(payload):
        for slot in _venue_slot_dicts(venue, index):
            yield slot, f"venue#{index}"


def _venue_slot_dicts(venue: dict, index: int) -> Iterable[dict]:
    slots = venue.get("slots")
    if slots is None:
        return  # a venue with no availability legitimately has no slots
    if not isinstance(slots, list):
        log.warning("skipping venue #%d: 'slots' is %s, not a list", index, type(slots).__name__)
        return
    for slot in slots:
        if not isinstance(slot, dict):
            log.warning("skipping malformed slot in venue #%d", index)
            continue
        yield slot


def _coerce_resy_id(raw: Any) -> int | None:
    if isinstance(raw, dict):
        raw = raw.get("resy")
    if isinstance(raw, bool):
        return None
    if isinstance(raw, int):
        return raw
    if isinstance(raw, str) and raw.strip().isdigit():
        return int(raw.strip())
    return None


def _coerce_float(raw: Any) -> float | None:
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    if isinstance(raw, str):
        try:
            return float(raw)
        except ValueError:
            return None
    return None


def _extract_identity(entry: dict) -> dict[str, Any]:
    """Pull venue identity out of a /4/find entry, whatever the nesting."""
    venue = entry.get("venue") if isinstance(entry.get("venue"), dict) else entry
    location = venue.get("location") if isinstance(venue.get("location"), dict) else {}

    name = venue.get("name")
    if not isinstance(name, str):
        name = ""

    neighborhood = venue.get("neighborhood")
    if not isinstance(neighborhood, str):
        neighborhood = location.get("neighborhood") if isinstance(location.get("neighborhood"), str) else ""

    lat = _coerce_float(location.get("latitude"))
    lng = _coerce_float(location.get("longitude"))
    if lat is None:
        lat = _coerce_float(venue.get("latitude"))
    if lng is None:
        lng = _coerce_float(venue.get("longitude"))

    return {
        "resy_venue_id": _coerce_resy_id(venue.get("id")),
        "name": name.strip(),
        "neighborhood": (neighborhood or "").strip(),
        "lat": lat,
        "lng": lng,
    }


@dataclass
class VenueSlots:
    """One venue's availability, as returned by a geo (city-wide) find."""

    resy_venue_id: int | None
    name: str
    neighborhood: str
    lat: float | None
    lng: float | None
    slots: list[ParsedSlot]
    raw: dict | None = None


def parse_find_venues(
    payload: Any,
    target_date: str,
    prime: PrimeWindow,
) -> list[VenueSlots]:
    """Group a /4/find payload by venue — the shape a geo sweep returns.

    Venues without a usable Resy id are dropped with a log line: without an id
    there is nothing stable to key observations against.
    """
    venues: list[VenueSlots] = []
    for entry, index in _iter_venue_entries(payload):
        identity = _extract_identity(entry)
        if identity["resy_venue_id"] is None:
            log.warning("skipping venue #%d: no usable Resy id", index)
            continue
        slots = _parse_slots(_venue_slot_dicts(entry, index), target_date, prime, f"venue#{index}")
        venues.append(
            VenueSlots(
                resy_venue_id=identity["resy_venue_id"],
                name=identity["name"],
                neighborhood=identity["neighborhood"],
                lat=identity["lat"],
                lng=identity["lng"],
                slots=slots,
                raw=entry,
            )
        )
    return venues


def parse_find_response(
    payload: Any,
    target_date: str,
    prime: PrimeWindow,
) -> list[ParsedSlot]:
    """Extract per-slot time, service type, and token from a /4/find payload.

    Flattens across venues, which is what a single-venue targeted find wants.
    """
    return _parse_slots(
        (slot for slot, _ in _iter_slot_dicts(payload)), target_date, prime, "find"
    )


def _parse_slots(
    slot_dicts: Iterable[dict],
    target_date: str,
    prime: PrimeWindow,
    where: str,
) -> list[ParsedSlot]:
    parsed: list[ParsedSlot] = []
    seen: set[tuple[str, str | None, str | None]] = set()

    for slot in slot_dicts:
        config = slot.get("config")
        config = config if isinstance(config, dict) else {}
        date_info = slot.get("date")
        date_info = date_info if isinstance(date_info, dict) else {}

        slot_date, slot_time = _parse_timestamp(date_info.get("start"))
        if slot_time is None:
            log.warning("%s: slot has no usable start time (%r), skipping", where, date_info.get("start"))
            continue

        service_type = config.get("type")
        if service_type is not None and not isinstance(service_type, str):
            service_type = str(service_type)

        # The booking token lives under different keys across shapes; take the
        # first that is actually a string. It is stored, never used to book.
        slot_token = None
        for key in ("token", "id"):
            candidate = config.get(key)
            if isinstance(candidate, str) and candidate:
                slot_token = candidate
                break
            if isinstance(candidate, int):
                slot_token = str(candidate)
                break

        effective_date = slot_date or target_date
        try:
            day = date_cls.fromisoformat(effective_date)
            is_prime = prime.covers(day, time.fromisoformat(slot_time))
        except ValueError:
            log.warning("%s: could not classify prime for %r", where, effective_date)
            is_prime = False

        key = (slot_time, service_type, slot_token)
        if key in seen:
            continue
        seen.add(key)
        parsed.append(
            ParsedSlot(
                slot_time=slot_time,
                service_type=service_type,
                slot_token=slot_token,
                is_prime=is_prime,
                slot_date=slot_date,
            )
        )

    parsed.sort(key=lambda s: (s.slot_time, s.service_type or ""))
    return parsed


def parse_venue_search(payload: Any) -> list[dict[str, Any]]:
    """Extract (resy_venue_id, name, locality) triples from /3/venuesearch/search."""
    hits: list[dict[str, Any]] = []
    search = payload.get("search") if isinstance(payload, dict) else None
    raw_hits = search.get("hits") if isinstance(search, dict) else None
    if not isinstance(raw_hits, list):
        log.debug("venuesearch payload had no 'search.hits' list")
        return hits

    for hit in raw_hits:
        if not isinstance(hit, dict):
            continue
        # Shape varies: sometimes the venue is nested, sometimes it is the hit.
        venue = hit.get("venue") if isinstance(hit.get("venue"), dict) else hit
        raw_id = venue.get("id")
        if isinstance(raw_id, dict):
            raw_id = raw_id.get("resy")
        if isinstance(raw_id, str) and raw_id.isdigit():
            raw_id = int(raw_id)
        if not isinstance(raw_id, int):
            log.warning("skipping venuesearch hit with unusable id %r", venue.get("id"))
            continue

        location = venue.get("location")
        location = location if isinstance(location, dict) else {}
        hits.append(
            {
                "resy_venue_id": raw_id,
                "name": venue.get("name") if isinstance(venue.get("name"), str) else "",
                "locality": location.get("locality") if isinstance(location.get("locality"), str) else "",
                "neighborhood": venue.get("neighborhood") if isinstance(venue.get("neighborhood"), str) else "",
            }
        )
    return hits


# ── Politeness ──────────────────────────────────────────────────────────────


class Politeness:
    """Caps concurrency and enforces a minimum gap between *all* requests."""

    def __init__(self, concurrency: int, delay_seconds: float) -> None:
        self._sem = asyncio.Semaphore(concurrency)
        self._spacing = asyncio.Lock()
        self._delay = delay_seconds
        self._next_at = 0.0

    @asynccontextmanager
    async def slot(self):
        async with self._sem:
            async with self._spacing:
                now = _time.monotonic()
                wait = self._next_at - now
                if wait > 0:
                    await asyncio.sleep(wait)
                self._next_at = max(now, self._next_at) + self._delay
            yield


class DailyBudget:
    """Global daily request cap, persisted so it survives across runs."""

    def __init__(self, conn: sqlite3.Connection, cap: int) -> None:
        self._conn = conn
        self._cap = cap
        self._lock = asyncio.Lock()

    @staticmethod
    def _today() -> str:
        return datetime.now(timezone.utc).date().isoformat()

    def used(self, service: str = "resy") -> int:
        row = self._conn.execute(
            "SELECT count FROM request_log WHERE day = ? AND service = ?",
            (self._today(), service),
        ).fetchone()
        return row["count"] if row else 0

    def remaining(self, service: str = "resy") -> int:
        return max(0, self._cap - self.used(service))

    async def consume(self, service: str = "resy", n: int = 1) -> None:
        async with self._lock:
            day = self._today()
            with self._conn:
                self._conn.execute(
                    "INSERT INTO request_log (day, service, count) VALUES (?, ?, 0) "
                    "ON CONFLICT (day, service) DO NOTHING",
                    (day, service),
                )
                row = self._conn.execute(
                    "SELECT count FROM request_log WHERE day = ? AND service = ?",
                    (day, service),
                ).fetchone()
                if row["count"] + n > self._cap:
                    raise DailyCapReached(
                        f"daily request cap of {self._cap} reached for '{service}' "
                        f"({row['count']} used today) — raise DAILY_REQUEST_CAP or wait for UTC midnight"
                    )
                self._conn.execute(
                    "UPDATE request_log SET count = count + ? WHERE day = ? AND service = ?",
                    (n, day, service),
                )


# ── Client ──────────────────────────────────────────────────────────────────


@dataclass
class ScanOutcome:
    """One venue/date observation, ready to persist."""

    venue_id: int
    resy_venue_id: int
    target_date: str
    party_size: int
    slots: list[ParsedSlot] = field(default_factory=list)
    raw_json: str | None = None
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


class ResyClient:
    """Async, rate-limited, read-only Resy client."""

    def __init__(
        self,
        settings: Settings,
        *,
        budget: DailyBudget | None = None,
        dry_run: bool = False,
        on_intent=None,
    ) -> None:
        self.settings = settings
        self.budget = budget
        self.dry_run = dry_run
        self._on_intent = on_intent or (lambda msg: log.info("dry-run: %s", msg))
        self._politeness = Politeness(settings.max_concurrency, settings.request_delay_seconds)
        self._client: httpx.AsyncClient | None = None
        self.request_count = 0
        # Counted so a sweep can tell "nothing was available" apart from
        # "nothing got through" — both otherwise look like zero venues.
        self.failed_requests = 0
        self.prime = PrimeWindow.from_settings(settings)

    # -- lifecycle --------------------------------------------------------
    async def __aenter__(self) -> "ResyClient":
        if not self.dry_run:
            api_key, token = self.settings.require_resy_credentials()
            headers = {
                "Authorization": f'ResyAPI api_key="{api_key}"',
                "Accept": "application/json, text/plain, */*",
                "Content-Type": "application/json",
                "Origin": "https://resy.com",
                "Referer": "https://resy.com/",
                "User-Agent": "resy-rank/0.1 (personal read-only research tool)",
            }
            # The browser's /4/find request carries only the api_key, so the
            # user token appears to be optional for availability lookups. It is
            # sent when present and simply omitted when not; a 401/403 will say
            # plainly if this endpoint turns out to want it after all.
            if token:
                headers["X-Resy-Auth-Token"] = token
            self._client = httpx.AsyncClient(
                headers=headers,
                timeout=self.settings.request_timeout_seconds,
            )
        return self

    async def __aexit__(self, *exc_info) -> None:
        if self._client is not None:
            await self._client.aclose()
            self._client = None

    # -- transport --------------------------------------------------------
    async def _request(
        self,
        method: str,
        url: str,
        *,
        params: dict | None = None,
        json_body: dict | None = None,
    ) -> Any:
        if url not in _ALLOWED_ENDPOINTS:
            raise RuntimeError(
                f"refusing to call non-allowlisted endpoint {url!r}; resy-rank is read-only"
            )

        if self.dry_run:
            detail = f"{method} {url}"
            if params:
                detail += f" params={params}"
            if json_body:
                detail += f" json={json_body}"
            self._on_intent(detail)
            return None

        assert self._client is not None, "use `async with ResyClient(...)`"

        last_exc: Exception | None = None
        for attempt in range(self.settings.max_retries + 1):
            if self.budget is not None:
                await self.budget.consume("resy")
            async with self._politeness.slot():
                self.request_count += 1
                try:
                    response = await self._client.request(
                        method, url, params=params, json=json_body
                    )
                except httpx.HTTPError as exc:
                    last_exc = exc
                    log.warning("%s %s failed: %s", method, url, exc)
                    response = None

            if response is not None:
                if response.status_code in (401, 403):
                    raise ResyUnauthorized(
                        f"Resy returned {response.status_code} — your RESY_AUTH_TOKEN has "
                        "probably expired. Grab a fresh one from your browser session."
                    )
                if response.status_code not in _RETRY_STATUS:
                    response.raise_for_status()
                    try:
                        return response.json()
                    except ValueError as exc:
                        raise RuntimeError(
                            f"{url} returned non-JSON body ({response.status_code})"
                        ) from exc
                last_exc = httpx.HTTPStatusError(
                    f"{response.status_code} from {url}", request=response.request, response=response
                )

            if attempt == self.settings.max_retries:
                break

            delay = self.settings.backoff_base_seconds ** (attempt + 1)
            if response is not None:
                retry_after = response.headers.get("Retry-After")
                if retry_after:
                    try:
                        delay = max(delay, float(retry_after))
                    except ValueError:
                        pass
            delay += random.uniform(0, 0.5)  # jitter, so retries don't sync up
            log.warning(
                "retry %d/%d for %s in %.1fs",
                attempt + 1,
                self.settings.max_retries,
                url,
                delay,
            )
            await asyncio.sleep(delay)

        raise RuntimeError(f"{url} failed after {self.settings.max_retries} retries") from last_exc

    # -- endpoints --------------------------------------------------------
    async def find(
        self,
        *,
        venue_id: int,
        resy_venue_id: int,
        target_date: str,
        party_size: int,
        lat: float | None = None,
        long: float | None = None,
    ) -> ScanOutcome:
        """GET /4/find for one venue on one date. Never raises on bad payloads."""
        # A venue-targeted find sends lat/long as 0 — the venue_id is the
        # filter, and the geo anchor is ignored. This mirrors the request the
        # Resy web app makes.
        body = {
            "day": target_date,
            "lat": 0,
            "long": 0,
            "party_size": party_size,
            "venue_id": resy_venue_id,
        }
        outcome = ScanOutcome(
            venue_id=venue_id,
            resy_venue_id=resy_venue_id,
            target_date=target_date,
            party_size=party_size,
        )
        try:
            payload = await self._request("POST", FIND_URL, json_body=body)
        except (DailyCapReached, ResyUnauthorized):
            raise  # fatal for the whole run, not just this venue
        except Exception as exc:
            self.failed_requests += 1
            outcome.error = str(exc)
            log.warning("find failed for venue %s on %s: %s", resy_venue_id, target_date, exc)
            return outcome

        if payload is None:  # dry-run
            return outcome

        outcome.raw_json = json.dumps(payload, separators=(",", ":"))
        try:
            outcome.slots = parse_find_response(payload, target_date, self.prime)
        except Exception as exc:  # parser bug — keep the raw blob, keep scanning
            outcome.error = f"parse error: {exc}"
            log.exception("parse failed for venue %s on %s", resy_venue_id, target_date)
        return outcome

    async def find_geo(
        self,
        *,
        lat: float,
        long: float,
        target_date: str,
        party_size: int,
    ) -> list[VenueSlots]:
        """GET /4/find with no venue_id — every venue with availability nearby.

        This is the sweep primitive: one call covers a neighborhood instead of
        one call per venue. Pagination is best-effort, since the parameters are
        undocumented; it stops as soon as a page adds no new venue ids.
        """
        collected: dict[int, VenueSlots] = {}
        for page in range(1, self.settings.geo_max_pages + 1):
            body = {
                "day": target_date,
                "lat": lat,
                "long": long,
                "party_size": party_size,
                "per_page": self.settings.geo_per_page,
                "page": page,
            }
            try:
                payload = await self._request("POST", FIND_URL, json_body=body)
            except (DailyCapReached, ResyUnauthorized):
                raise
            except Exception as exc:
                self.failed_requests += 1
                log.warning("geo find failed at (%s, %s) page %d: %s", lat, long, page, exc)
                break

            if payload is None:  # dry-run
                break

            venues = parse_find_venues(payload, target_date, self.prime)
            fresh = [v for v in venues if v.resy_venue_id not in collected]
            for venue in fresh:
                collected[venue.resy_venue_id] = venue

            # Stop when the page is empty or adds nothing new — the safest
            # reading of an endpoint that may ignore `page` entirely.
            if not venues or not fresh:
                break
        return list(collected.values())

    async def search_venues(
        self, query: str, *, lat: float | None = None, long: float | None = None
    ) -> list[dict[str, Any]]:
        """POST /3/venuesearch/search — resolves a name to a Resy venue id."""
        body = {
            "query": query,
            "geo": {
                "latitude": lat if lat is not None else self.settings.default_lat,
                "longitude": long if long is not None else self.settings.default_long,
                "radius": 35000,
            },
            "per_page": 5,
        }
        payload = await self._request("POST", VENUE_SEARCH_URL, json_body=body)
        if payload is None:  # dry-run
            return []
        return parse_venue_search(payload)


# ── Persistence ─────────────────────────────────────────────────────────────


def upsert_geo_venue(conn: sqlite3.Connection, venue: VenueSlots, now: str) -> int:
    """Insert or link a venue seen in a geo sweep. Returns the local venue_id.

    A sweep hit also resolves seed rows for free: if venues.csv has a row with
    a matching name and no Resy id yet, the id is filled in here rather than
    costing a separate venuesearch call.
    """
    row = conn.execute(
        "SELECT venue_id FROM venues WHERE resy_venue_id = ?", (venue.resy_venue_id,)
    ).fetchone()

    if row is None and venue.name:
        row = conn.execute(
            "SELECT venue_id FROM venues WHERE resy_venue_id IS NULL AND LOWER(name) = ?",
            (venue.name.lower(),),
        ).fetchone()
        if row is not None:
            conn.execute(
                "UPDATE venues SET resy_venue_id = ? WHERE venue_id = ?",
                (venue.resy_venue_id, row["venue_id"]),
            )

    if row is None:
        cursor = conn.execute(
            "INSERT INTO venues (name, neighborhood, resy_venue_id, lat, lng, "
            "in_geo_scope, last_seen_at) VALUES (?, ?, ?, ?, ?, 1, ?)",
            (
                venue.name or f"resy:{venue.resy_venue_id}",
                venue.neighborhood,
                venue.resy_venue_id,
                venue.lat,
                venue.lng,
                now,
            ),
        )
        return cursor.lastrowid

    conn.execute(
        "UPDATE venues SET in_geo_scope = 1, last_seen_at = ?, "
        "lat = COALESCE(lat, ?), lng = COALESCE(lng, ?), "
        "neighborhood = CASE WHEN neighborhood = '' THEN ? ELSE neighborhood END "
        "WHERE venue_id = ?",
        (now, venue.lat, venue.lng, venue.neighborhood, row["venue_id"]),
    )
    return row["venue_id"]


def persist_geo_sweep(
    conn: sqlite3.Connection,
    *,
    target_date: str,
    party_size: int,
    venues: list[VenueSlots],
    scan_ts: str,
) -> dict[str, int]:
    """Record one date's sweep: availability for venues seen, absence for the rest.

    The absence rows are the point. A venue that is in scope but missing from
    the sweep was observed and had nothing — which is exactly the signal the
    scarcity index reads.
    """
    seen_ids: set[int] = set()
    with conn:
        for venue in venues:
            venue_id = upsert_geo_venue(conn, venue, scan_ts)
            seen_ids.add(venue_id)
            cursor = conn.execute(
                "INSERT INTO scans (venue_id, scan_ts, target_date, party_size, raw_json) "
                "VALUES (?, ?, ?, ?, ?)",
                (
                    venue_id,
                    scan_ts,
                    target_date,
                    party_size,
                    # Only this venue's slice of the payload, not the whole
                    # city blob repeated per venue.
                    json.dumps(venue.raw, separators=(",", ":")) if venue.raw is not None else None,
                ),
            )
            conn.executemany(
                "INSERT INTO slots (scan_id, venue_id, target_date, slot_time, "
                "service_type, slot_token, is_prime) VALUES (?, ?, ?, ?, ?, ?, ?)",
                [
                    (
                        cursor.lastrowid,
                        venue_id,
                        target_date,
                        slot.slot_time,
                        slot.service_type,
                        slot.slot_token,
                        int(slot.is_prime),
                    )
                    for slot in venue.slots
                ],
            )

        absent = [
            row["venue_id"]
            for row in conn.execute("SELECT venue_id FROM venues WHERE in_geo_scope = 1")
            if row["venue_id"] not in seen_ids
        ]
        conn.executemany(
            "INSERT INTO scans (venue_id, scan_ts, target_date, party_size, raw_json) "
            "VALUES (?, ?, ?, ?, NULL)",
            [(venue_id, scan_ts, target_date, party_size) for venue_id in absent],
        )

    return {"available": len(seen_ids), "unavailable": len(absent)}


def merge_geo_results(batches: Iterable[list[VenueSlots]]) -> list[VenueSlots]:
    """Union venue results across anchor points, deduplicating slots."""
    merged: dict[int, VenueSlots] = {}
    for batch in batches:
        for venue in batch:
            existing = merged.get(venue.resy_venue_id)
            if existing is None:
                merged[venue.resy_venue_id] = venue
                continue
            known = {(s.slot_time, s.service_type, s.slot_token) for s in existing.slots}
            for slot in venue.slots:
                if (slot.slot_time, slot.service_type, slot.slot_token) not in known:
                    existing.slots.append(slot)
            existing.slots.sort(key=lambda s: (s.slot_time, s.service_type or ""))
    return list(merged.values())


def persist_scan(conn: sqlite3.Connection, outcome: ScanOutcome, scan_ts: str) -> int:
    """Write one scan plus its slots. Returns the scan_id."""
    with conn:
        cursor = conn.execute(
            "INSERT INTO scans (venue_id, scan_ts, target_date, party_size, raw_json) "
            "VALUES (?, ?, ?, ?, ?)",
            (
                outcome.venue_id,
                scan_ts,
                outcome.target_date,
                outcome.party_size,
                outcome.raw_json,
            ),
        )
        scan_id = cursor.lastrowid
        conn.executemany(
            "INSERT INTO slots "
            "(scan_id, venue_id, target_date, slot_time, service_type, slot_token, is_prime) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            [
                (
                    scan_id,
                    outcome.venue_id,
                    outcome.target_date,
                    slot.slot_time,
                    slot.service_type,
                    slot.slot_token,
                    int(slot.is_prime),
                )
                for slot in outcome.slots
            ],
        )
    return scan_id
