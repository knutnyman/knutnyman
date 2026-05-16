"""Polling loop that watches for a matching slot and books it."""

import logging
import time
from dataclasses import dataclass, field

from notifier import notify
from resy_client import ResyClient, Slot

log = logging.getLogger(__name__)


@dataclass
class BookingTarget:
    venue_id: int
    venue_name: str
    dates: list[str]    # one or more YYYY-MM-DD dates to check each poll cycle
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
    """Poll until a matching slot is found across any of the target dates, then book it."""
    date_summary = ", ".join(target.dates) if len(target.dates) <= 5 else f"{target.dates[0]} … {target.dates[-1]} ({len(target.dates)} dates)"
    log.info(
        "Watching %s on [%s] for party of %d between %s and %s (every %ds)",
        target.venue_name,
        date_summary,
        target.party_size,
        target.earliest_time,
        target.latest_time,
        poll_interval,
    )

    attempt = 0
    while True:
        attempt += 1
        try:
            found_slot: Slot | None = None
            for date in target.dates:
                slots = client.find_slots(target.venue_id, date, target.party_size)
                matching = [s for s in slots if _slot_in_window(s, target)]
                if matching:
                    found_slot = matching[0]
                    break

            if found_slot is None:
                log.info("[%d] No matching slots on any date yet. Retrying in %ds…", attempt, poll_interval)
            else:
                slot = found_slot
                log.info("[%d] Found slot on %s at %s!", attempt, slot.date, slot.time_label)

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
