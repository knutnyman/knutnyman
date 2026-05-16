"""Resy API client using the unofficial REST API."""

import json
import logging
from dataclasses import dataclass, field
from typing import Optional

import requests

log = logging.getLogger(__name__)

RESY_API_KEY = "VbWk7s3L4KiK5fzlO7JD3Q5EYolJI7n5"
BASE_URL = "https://api.resy.com"


@dataclass
class Slot:
    config_id: str
    date: str
    time_start: str
    time_end: str
    party_size: int
    type: str = ""

    @property
    def time_label(self) -> str:
        return self.time_start[:5]  # "HH:MM"


@dataclass
class ResyClient:
    email: str
    password: str
    _token: Optional[str] = field(default=None, repr=False)
    _payment_method_id: Optional[int] = field(default=None, repr=False)
    _session: requests.Session = field(default_factory=requests.Session, repr=False)

    def _base_headers(self) -> dict:
        headers = {
            "Authorization": f'ResyAPI api_key="{RESY_API_KEY}"',
            "Origin": "https://resy.com",
            "Referer": "https://resy.com/",
            "User-Agent": "Mozilla/5.0",
            "Content-Type": "application/x-www-form-urlencoded",
        }
        if self._token:
            headers["X-Resy-Auth-Token"] = self._token
        return headers

    def login(self) -> None:
        resp = self._session.post(
            f"{BASE_URL}/3/auth/password",
            data={"email": self.email, "password": self.password},
            headers=self._base_headers(),
        )
        resp.raise_for_status()
        body = resp.json()
        self._token = body["token"]
        self._payment_method_id = body.get("payment_method_id")
        log.info("Logged in to Resy as %s", self.email)

    def search_venue(self, query: str, lat: float = 40.7128, lon: float = -74.0060) -> list[dict]:
        """Return a list of venues matching the search query."""
        resp = self._session.get(
            f"{BASE_URL}/3/venue/search",
            params={"query": query, "geo[lat]": lat, "geo[lon]": lon},
            headers=self._base_headers(),
        )
        resp.raise_for_status()
        body = resp.json()
        return body.get("search", {}).get("hits", [])

    def find_slots(
        self,
        venue_id: int,
        date: str,
        party_size: int,
        lat: float = 0,
        lon: float = 0,
    ) -> list[Slot]:
        """Return available reservation slots for a venue on a given date."""
        resp = self._session.get(
            f"{BASE_URL}/4/find",
            params={
                "lat": lat,
                "long": lon,
                "day": date,
                "party_size": party_size,
                "venue_id": venue_id,
            },
            headers=self._base_headers(),
        )
        resp.raise_for_status()
        body = resp.json()

        slots: list[Slot] = []
        for venue in body.get("results", {}).get("venues", []):
            for slot_group in venue.get("slots", []):
                config = slot_group.get("config", {})
                date_info = slot_group.get("date", {})
                slots.append(
                    Slot(
                        config_id=config.get("id", ""),
                        date=date_info.get("start", "")[:10],
                        time_start=date_info.get("start", "")[11:16],
                        time_end=date_info.get("end", "")[11:16],
                        party_size=party_size,
                        type=config.get("type", ""),
                    )
                )
        return slots

    def get_book_token(self, slot: Slot) -> tuple[str, dict]:
        """Fetch the book token and payment details needed to complete a booking."""
        resp = self._session.post(
            f"{BASE_URL}/3/details",
            data={
                "config_id": slot.config_id,
                "day": slot.date,
                "party_size": slot.party_size,
            },
            headers=self._base_headers(),
        )
        resp.raise_for_status()
        body = resp.json()
        book_token = body["book_token"]["value"]
        payment_method = body.get("user", {}).get("payment_methods", [{}])[0]
        return book_token, payment_method

    def book(self, slot: Slot) -> dict:
        """Book a slot. Returns the confirmation response body."""
        book_token, payment_method = self.get_book_token(slot)
        resp = self._session.post(
            f"{BASE_URL}/3/book",
            data={
                "book_token": book_token,
                "struct_payment_method": json.dumps(payment_method),
                "source_id": "resy.com-venue-details",
            },
            headers=self._base_headers(),
        )
        resp.raise_for_status()
        return resp.json()
