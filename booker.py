"""Polling loop that watches for a matching slot and books it."""

import logging
import time
from dataclasses import dataclass
from datetime import datetime, time as dtime

from notifier import notify
from resy_client import ResyClient, Slot

log = logging.getLogger(__name__)

DEFAULT_RELEASE_TIMES = [dtime(0, 0), dtime(9, 0)]  # midnight and 9am


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


def _secs_from_midnight(t: dtime) -> int:
    return t.hour * 3600 + t.minute * 60 + t.second


def _near_release_time(release_times: list[dtime], window_secs: int) -> bool:
    """Return True if the current clock time is within window_secs of any release time."""
    now_secs = _secs_from_midnight(datetime.now().time())
    for rt in release_times:
        rt_secs = _secs_from_midnight(rt)
        diff = abs(now_secs - rt_secs)
        diff = min(diff, 86400 - diff)  # handle midnight wraparound
        if diff <= window_secs:
            return True
    return False


def run(
    client: ResyClient,
    target: BookingTarget,
    poll_interval: int = 30,
    fast_interval: int = 1,
    release_times: list[dtime] | None = None,
    fast_window_secs: int = 120,
    dry_run: bool = False,
) -> None:
    """Poll until a matching slot is found across any of the target dates, then book it.

    Automatically switches to fast_interval polling within fast_window_secs of any
    release_time, then falls back to poll_interval between windows.
    """
    if release_times is None:
        release_times = DEFAULT_RELEASE_TIMES

    date_summary = (
        ", ".join(target.dates)
        if len(target.dates) <= 5
        else f"{target.dates[0]} … {target.dates[-1]} ({len(target.dates)} dates)"
    )
    release_labels = ", ".join(t.strftime("%H:%M") for t in release_times)
    log.info(
        "Watching %s | dates: [%s] | party: %d | window: %s–%s",
        target.venue_name, date_summary, target.party_size,
        target.earliest_time, target.latest_time,
    )
    log.info(
        "Polling every %ds normally, every %ds within %ds of release times [%s]",
        poll_interval, fast_interval, fast_window_secs, release_labels,
    )

    attempt = 0
    in_fast_mode = False

    while True:
        attempt += 1

        # Determine polling speed and log transitions
        fast_now = _near_release_time(release_times, fast_window_secs)
        if fast_now and not in_fast_mode:
            log.info("*** Entering fast-poll mode (1s) — release window open ***")
            in_fast_mode = True
        elif not fast_now and in_fast_mode:
            log.info("Release window closed — returning to slow poll (%ds)", poll_interval)
            in_fast_mode = False

        current_interval = fast_interval if in_fast_mode else poll_interval

        try:
            found_slot: Slot | None = None
            for date in target.dates:
                slots = client.find_slots(target.venue_id, date, target.party_size)
                matching = [s for s in slots if _slot_in_window(s, target)]
                if matching:
                    found_slot = matching[0]
                    break

            if found_slot is None:
                log.info("[%d] No slots yet (%ds poll)…", attempt, current_interval)
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

        time.sleep(current_interval)
