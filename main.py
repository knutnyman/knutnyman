#!/usr/bin/env python3
"""CLI entry point for the Resy auto-booker."""

import argparse
import logging
import os
import sys
from datetime import date, timedelta

from dotenv import load_dotenv

from booker import BookingTarget, run
from resy_client import ResyClient

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Automatically book a Resy reservation when a slot opens up.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Search for a venue and print its ID
  python main.py search --query "Carbone"

  # Book a specific date
  python main.py book --venue-id 1234 --venue-name "Carbone" \\
      --date 2026-06-15 --party-size 2 --earliest 19:00 --latest 21:00

  # Book any Friday, Saturday, or Sunday in the next 60 days
  python main.py book --venue-id 1234 --venue-name "Carbone" \\
      --days fri sat sun --party-size 2 --earliest 19:00 --latest 21:00

  # Narrow the window (next 4 weekends only)
  python main.py book --venue-id 1234 --venue-name "Carbone" \\
      --days sat sun --from-date 2026-06-01 --to-date 2026-06-30 \\
      --party-size 2 --earliest 19:00 --latest 21:00

  # Same but don't actually book (just notify)
  python main.py book ... --dry-run
""",
    )
    sub = p.add_subparsers(dest="command", required=True)

    # ── search ──────────────────────────────────────────────────────────────
    s = sub.add_parser("search", help="Search for a venue and print its ID")
    s.add_argument("--query", required=True, help="Restaurant name to search for")
    s.add_argument("--lat", type=float, default=40.7128, help="Latitude for geo search (default: NYC)")
    s.add_argument("--lon", type=float, default=-74.0060, help="Longitude for geo search (default: NYC)")

    # ── book ─────────────────────────────────────────────────────────────────
    b = sub.add_parser("book", help="Poll and book a reservation")
    b.add_argument("--venue-id", type=int, required=True, help="Resy venue ID (from 'search')")
    b.add_argument("--venue-name", required=True, help="Human-readable venue name (for notifications)")
    b.add_argument("--party-size", type=int, required=True, help="Number of guests")
    b.add_argument("--earliest", default="00:00", help="Earliest acceptable time slot (HH:MM, default 00:00)")
    b.add_argument("--latest", default="23:59", help="Latest acceptable time slot (HH:MM, default 23:59)")
    b.add_argument("--interval", type=int, default=30, help="Polling interval in seconds (default 30)")
    b.add_argument("--dry-run", action="store_true", help="Find a slot but don't actually book it")

    date_group = b.add_mutually_exclusive_group(required=True)
    date_group.add_argument("--date", help="Specific date to target (YYYY-MM-DD)")
    date_group.add_argument(
        "--days",
        nargs="+",
        choices=["mon", "tue", "wed", "thu", "fri", "sat", "sun"],
        metavar="DAY",
        help="Days of week to target (e.g. --days fri sat sun). Searches within --from-date..--to-date.",
    )
    b.add_argument(
        "--from-date",
        default=str(date.today()),
        help="Start of search window when using --days (YYYY-MM-DD, default: today)",
    )
    b.add_argument(
        "--to-date",
        default=str(date.today() + timedelta(days=60)),
        help="End of search window when using --days (YYYY-MM-DD, default: 60 days from today)",
    )

    return p.parse_args()


def require_env(*names: str) -> None:
    missing = [n for n in names if not os.getenv(n)]
    if missing:
        log.error("Missing required environment variables: %s", ", ".join(missing))
        log.error("Copy .env.example to .env and fill in your credentials.")
        sys.exit(1)


def make_client() -> ResyClient:
    require_env("RESY_EMAIL", "RESY_PASSWORD")
    client = ResyClient(email=os.environ["RESY_EMAIL"], password=os.environ["RESY_PASSWORD"])
    client.login()
    return client


def cmd_search(args: argparse.Namespace) -> None:
    client = make_client()
    hits = client.search_venue(args.query, lat=args.lat, lon=args.lon)
    if not hits:
        print("No venues found.")
        return
    print(f"{'ID':<10} {'Name':<40} {'Location'}")
    print("-" * 70)
    for hit in hits:
        venue = hit.get("venue", hit)  # shape differs by API version
        vid = venue.get("id", {}).get("resy", "?")
        name = venue.get("name", "?")
        location = venue.get("location", {}).get("locality", "?")
        print(f"{vid:<10} {name:<40} {location}")


_DAY_NAMES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


def _build_dates(args: argparse.Namespace) -> list[str]:
    if args.date:
        return [args.date]

    target_weekdays = {_DAY_NAMES.index(d) for d in args.days}
    start = date.fromisoformat(args.from_date)
    end = date.fromisoformat(args.to_date)
    if end < start:
        log.error("--to-date must be on or after --from-date")
        sys.exit(1)

    dates = []
    current = start
    while current <= end:
        if current.weekday() in target_weekdays:
            dates.append(str(current))
        current += timedelta(days=1)

    if not dates:
        log.error("No dates match --days %s in range %s..%s", args.days, start, end)
        sys.exit(1)

    log.info("Targeting %d date(s): %s%s", len(dates), ", ".join(dates[:5]), " …" if len(dates) > 5 else "")
    return dates


def cmd_book(args: argparse.Namespace) -> None:
    client = make_client()
    target = BookingTarget(
        venue_id=args.venue_id,
        venue_name=args.venue_name,
        dates=_build_dates(args),
        party_size=args.party_size,
        earliest_time=args.earliest,
        latest_time=args.latest,
    )
    run(client, target, poll_interval=args.interval, dry_run=args.dry_run)


def main() -> None:
    args = parse_args()
    if args.command == "search":
        cmd_search(args)
    elif args.command == "book":
        cmd_book(args)


if __name__ == "__main__":
    main()
