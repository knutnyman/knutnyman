"""Polling loop that watches for a matching slot and books it."""

import logging
import time
from dataclasses import dataclass

from notifier import notify
from resy_client import ResyClient, Slot

log = logging.getLogger(__name__)


@dataclass
class BookingTarget:
    venue_id: int
    venue_name: str
    date: str           # YYYY-MM-DD
    party_size: int
    earliest_time: str  # "HH:MM"
    latest_time: str    # "HH:MM"


def _slot_in_window(slot: Slot, target: BookingTarget) -> bool:
    return target.earliest_time <= slot.time_start <= target.latest_time


def run(
    client: ResyClient,
    target: BookingTarget,
    poll_interval: int = 30,
    dry_run: bool = False,
) -> None:
    """Poll until a matching slot is found, then book it."""
    log.info(
        "Watching %s on %s for party of %d between %s and %s (every %ds)",
        target.venue_name,
        target.date,
        target.party_size,
        target.earliest_time,
        target.latest_time,
        poll_interval,
    )

    attempt = 0
    while True:
        attempt += 1
        try:
            slots = client.find_slots(target.venue_id, target.date, target.party_size)
            matching = [s for s in slots if _slot_in_window(s, target)]

            if not matching:
                log.info("[%d] No matching slots yet. Retrying in %ds…", attempt, poll_interval)
            else:
                slot = matching[0]
                log.info("[%d] Found slot at %s!", attempt, slot.time_label)

                if dry_run:
                    notify(
                        "Resy slot found (dry run)",
                        f"{target.venue_name} — {slot.date} at {slot.time_label} "
                        f"for {target.party_size} ({slot.type})\nNot booked (dry-run mode).",
                    )
                    return

                confirmation = client.book(slot)
                resy_token = confirmation.get("resy_token", "N/A")
                notify(
                    f"Reservation booked at {target.venue_name}!",
                    f"Date: {slot.date}\n"
                    f"Time: {slot.time_label}\n"
                    f"Party size: {target.party_size}\n"
                    f"Type: {slot.type}\n"
                    f"Resy token: {resy_token}",
                )
                return

        except Exception as exc:
            log.warning("[%d] Error during poll: %s", attempt, exc)

        time.sleep(poll_interval)
