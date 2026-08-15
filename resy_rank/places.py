"""Google Places API (New) enrichment.

Cost model, and why this module is written the way it is:

Since March 2025 Google bills Places per-SKU with a per-SKU monthly free tier —
10,000 calls for Essentials, 5,000 for Pro, 1,000 for Enterprise — and the free
allowances do NOT pool across SKUs. The SKU is chosen by the *field mask*, not
by a billing setting: asking for any Enterprise-tier field promotes the whole
call to Enterprise pricing.

`rating`, `userRatingCount`, and `priceLevel` are Enterprise fields. So the
ranking signal we actually want is the expensive one, which makes caching a
cost control and not merely a speed optimization. Two consequences:

  * Resolution (name -> place_id) is kept in a separate, cheap call with a
    minimal field mask, so it draws on its own free allowance rather than
    burning the Enterprise one.
  * Enterprise calls are capped per calendar month (default 1,000 = the free
    tier) so a large venue universe spreads across months at zero cost instead
    of silently running up a bill.

Every result is cached permanently in SQLite and only re-fetched when
`--refresh` is passed or the row is older than GOOGLE_CACHE_DAYS.
"""

from __future__ import annotations

import asyncio
import logging
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx

from resy_rank.config import Settings

log = logging.getLogger(__name__)

SEARCH_TEXT_URL = "https://places.googleapis.com/v1/places:searchText"
DETAILS_URL_TEMPLATE = "https://places.googleapis.com/v1/places/{place_id}"

# Minimal mask — id and name only. Keeps resolution off the Enterprise SKU.
SEARCH_FIELD_MASK = "places.id,places.displayName,places.formattedAddress"

# Enterprise SKU: `rating`, `userRatingCount`, and `priceLevel` are what
# promote this call. Requesting them is the entire point of the call, so the
# defence is the cache and the monthly cap, not a smaller mask.
DETAILS_FIELD_MASK = "id,displayName,rating,userRatingCount,priceLevel,location,websiteUri"

SERVICE_TEXT = "places_text"
SERVICE_DETAILS = "places_enterprise"


class MonthlyCapReached(RuntimeError):
    """Raised when the configured monthly Enterprise-SKU cap is exhausted."""


class MonthlyBudget:
    """Calendar-month cap on billable calls, persisted across runs."""

    def __init__(self, conn: sqlite3.Connection, cap: int) -> None:
        self._conn = conn
        self._cap = cap
        self._lock = asyncio.Lock()

    @staticmethod
    def _today() -> str:
        return datetime.now(timezone.utc).date().isoformat()

    def used(self, service: str) -> int:
        row = self._conn.execute(
            "SELECT COALESCE(SUM(count), 0) AS n FROM request_log "
            "WHERE service = ? AND day LIKE ?",
            (service, f"{self._today()[:7]}%"),
        ).fetchone()
        return row["n"] or 0

    def remaining(self, service: str) -> int:
        return max(0, self._cap - self.used(service))

    async def consume(self, service: str, n: int = 1) -> None:
        async with self._lock:
            if self.used(service) + n > self._cap:
                raise MonthlyCapReached(
                    f"monthly cap of {self._cap} reached for '{service}'. This cap exists "
                    "to keep you inside Google's free tier — raise MONTHLY_ENRICHMENT_CAP "
                    "only if you intend to be billed."
                )
            day = self._today()
            with self._conn:
                self._conn.execute(
                    "INSERT INTO request_log (day, service, count) VALUES (?, ?, 0) "
                    "ON CONFLICT (day, service) DO NOTHING",
                    (day, service),
                )
                self._conn.execute(
                    "UPDATE request_log SET count = count + ? WHERE day = ? AND service = ?",
                    (n, day, service),
                )


@dataclass
class PlaceDetails:
    place_id: str
    display_name: str | None = None
    rating: float | None = None
    user_rating_count: int | None = None
    price_level: str | None = None
    lat: float | None = None
    lng: float | None = None
    website_uri: str | None = None


def _as_float(raw: Any) -> float | None:
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    return None


def _as_int(raw: Any) -> int | None:
    if isinstance(raw, bool):
        return None
    if isinstance(raw, int):
        return raw
    if isinstance(raw, float):
        return int(raw)
    return None


def _display_name(raw: Any) -> str | None:
    """displayName is {"text": ..., "languageCode": ...}, but tolerate a string."""
    if isinstance(raw, dict):
        text = raw.get("text")
        return text if isinstance(text, str) else None
    return raw if isinstance(raw, str) else None


def parse_details(payload: Any) -> PlaceDetails | None:
    """Parse a places/{id} response defensively — missing fields are fine."""
    if not isinstance(payload, dict):
        log.warning("places details payload was %s, not a dict", type(payload).__name__)
        return None
    place_id = payload.get("id")
    if not isinstance(place_id, str) or not place_id:
        log.warning("places details payload had no usable id")
        return None

    location = payload.get("location") if isinstance(payload.get("location"), dict) else {}
    price_level = payload.get("priceLevel")

    return PlaceDetails(
        place_id=place_id,
        display_name=_display_name(payload.get("displayName")),
        rating=_as_float(payload.get("rating")),
        user_rating_count=_as_int(payload.get("userRatingCount")),
        price_level=price_level if isinstance(price_level, str) else None,
        lat=_as_float(location.get("latitude")),
        lng=_as_float(location.get("longitude")),
        website_uri=payload.get("websiteUri") if isinstance(payload.get("websiteUri"), str) else None,
    )


def parse_search_text(payload: Any) -> str | None:
    """Return the first place id from a searchText response, or None."""
    if not isinstance(payload, dict):
        return None
    places = payload.get("places")
    if not isinstance(places, list) or not places:
        return None
    first = places[0]
    if not isinstance(first, dict):
        return None
    place_id = first.get("id")
    return place_id if isinstance(place_id, str) and place_id else None


class PlacesClient:
    """Async Google Places client with per-SKU budgets and a billable counter."""

    def __init__(
        self,
        settings: Settings,
        *,
        budget: MonthlyBudget | None = None,
        dry_run: bool = False,
        on_intent=None,
    ) -> None:
        self.settings = settings
        self.budget = budget
        self.dry_run = dry_run
        self._on_intent = on_intent or (lambda msg: log.info("dry-run: %s", msg))
        self._client: httpx.AsyncClient | None = None
        # Running tally of billable calls, reported at the end of every run.
        self.calls: dict[str, int] = {SERVICE_TEXT: 0, SERVICE_DETAILS: 0}

    async def __aenter__(self) -> "PlacesClient":
        if not self.dry_run:
            api_key = self.settings.require_google_key()
            self._client = httpx.AsyncClient(
                headers={"X-Goog-Api-Key": api_key, "Content-Type": "application/json"},
                timeout=self.settings.request_timeout_seconds,
            )
        return self

    async def __aexit__(self, *exc_info) -> None:
        if self._client is not None:
            await self._client.aclose()
            self._client = None

    @property
    def billable_total(self) -> int:
        return sum(self.calls.values())

    async def _call(
        self,
        method: str,
        url: str,
        *,
        service: str,
        field_mask: str,
        json_body: dict | None = None,
    ) -> Any:
        if self.dry_run:
            self._on_intent(f"{method} {url} mask={field_mask} ({service} SKU)")
            return None

        assert self._client is not None, "use `async with PlacesClient(...)`"
        if self.budget is not None and service == SERVICE_DETAILS:
            await self.budget.consume(service)
        elif self.budget is not None:
            # Text search has its own, larger free allowance; still counted.
            await self.budget.consume(service)

        self.calls[service] = self.calls.get(service, 0) + 1
        response = await self._client.request(
            method, url, headers={"X-Goog-FieldMask": field_mask}, json=json_body
        )
        if response.status_code == 403:
            raise RuntimeError(
                "Google Places returned 403 — check that the Places API (New) is enabled "
                "for this key and that billing is set up on the project."
            )
        response.raise_for_status()
        return response.json()

    async def resolve_place_id(self, name: str, address: str = "") -> str | None:
        """Name (+ address) -> place_id. Cheap SKU: id and name only."""
        query = f"{name} {address}".strip()
        payload = await self._call(
            "POST",
            SEARCH_TEXT_URL,
            service=SERVICE_TEXT,
            field_mask=SEARCH_FIELD_MASK,
            json_body={"textQuery": query, "maxResultCount": 1},
        )
        if payload is None:
            return None
        place_id = parse_search_text(payload)
        if place_id is None:
            log.warning("no Google Places match for %r", query)
        return place_id

    async def fetch_details(self, place_id: str) -> PlaceDetails | None:
        """place_id -> rating and friends. Enterprise SKU: this is the billable one."""
        payload = await self._call(
            "GET",
            DETAILS_URL_TEMPLATE.format(place_id=place_id),
            service=SERVICE_DETAILS,
            field_mask=DETAILS_FIELD_MASK,
        )
        if payload is None:
            return None
        return parse_details(payload)


# ── Cache ───────────────────────────────────────────────────────────────────


def is_stale(fetched_at: str | None, cache_days: int, now: datetime | None = None) -> bool:
    """True when a cached row is missing or older than the cache window."""
    if not fetched_at:
        return True
    try:
        stamp = datetime.fromisoformat(fetched_at)
    except ValueError:
        return True
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    reference = now or datetime.now(timezone.utc)
    return (reference - stamp) > timedelta(days=cache_days)


def venues_needing_enrichment(
    conn: sqlite3.Connection, cache_days: int, *, refresh: bool = False
) -> list[sqlite3.Row]:
    """Venues whose Google data is missing or stale (all of them if refreshing)."""
    rows = conn.execute(
        """
        SELECT v.venue_id, v.name, v.address, v.google_place_id, g.fetched_at
        FROM venues v
        LEFT JOIN google_meta g ON g.venue_id = v.venue_id
        ORDER BY v.name
        """
    ).fetchall()
    if refresh:
        return rows
    return [r for r in rows if is_stale(r["fetched_at"], cache_days)]


def save_details(conn: sqlite3.Connection, venue_id: int, details: PlaceDetails) -> None:
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with conn:
        conn.execute(
            "UPDATE venues SET google_place_id = ?, lat = COALESCE(lat, ?), "
            "lng = COALESCE(lng, ?) WHERE venue_id = ?",
            (details.place_id, details.lat, details.lng, venue_id),
        )
        conn.execute(
            """
            INSERT INTO google_meta
                (venue_id, rating, user_rating_count, price_level, website_uri, fetched_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT (venue_id) DO UPDATE SET
                rating            = excluded.rating,
                user_rating_count = excluded.user_rating_count,
                price_level       = excluded.price_level,
                website_uri       = excluded.website_uri,
                fetched_at        = excluded.fetched_at
            """,
            (
                venue_id,
                details.rating,
                details.user_rating_count,
                details.price_level,
                details.website_uri,
                now,
            ),
        )


def save_place_id(conn: sqlite3.Connection, venue_id: int, place_id: str) -> None:
    with conn:
        conn.execute(
            "UPDATE venues SET google_place_id = ? WHERE venue_id = ?", (place_id, venue_id)
        )
