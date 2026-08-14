"""Typer entry point for resy-rank.

Read-only by construction: no command in this tool books, holds, modifies, or
cancels a reservation.
"""

from __future__ import annotations

import asyncio
import logging
from datetime import date as date_cls, datetime, timedelta, timezone
from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console
from rich.table import Table

from resy_rank import db
from resy_rank.config import get_settings
from resy_rank.resy import (
    DailyBudget,
    DailyCapReached,
    ResyClient,
    ResyUnauthorized,
    ScanOutcome,
    persist_scan,
)

app = typer.Typer(
    add_completion=False,
    no_args_is_help=True,
    help="Scan Resy for availability and rank the bookable options. Never books.",
)
console = Console()


class Context:
    """Shared state hung off the Typer context."""

    def __init__(self, dry_run: bool = False) -> None:
        self.dry_run = dry_run
        self.settings = get_settings()


@app.callback()
def main(
    ctx: typer.Context,
    dry_run: Annotated[
        bool,
        typer.Option("--dry-run", help="Hit nothing over the network; print intended requests."),
    ] = False,
    verbose: Annotated[bool, typer.Option("--verbose", "-v", help="Debug logging.")] = False,
) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    ctx.obj = Context(dry_run=dry_run)


@app.command()
def init(
    ctx: typer.Context,
    venues: Annotated[
        Optional[Path],
        typer.Option("--venues", help="Venue CSV to load (default: data/venues.csv)."),
    ] = None,
) -> None:
    """Create the database and load the venue universe from venues.csv."""
    settings: "Context" = ctx.obj
    cfg = settings.settings
    venues_csv = venues or cfg.venues_csv

    if settings.dry_run:
        console.print("[yellow]dry-run:[/yellow] would create/migrate database and load venues")
        console.print(f"  database : {cfg.db_path}")
        console.print(f"  venues   : {venues_csv}")
        return

    before, after = db.init_db(cfg.db_path)
    if before == after:
        console.print(f"Database already at schema v{after}: [dim]{cfg.db_path}[/dim]")
    else:
        console.print(
            f"Database migrated v{before} → v{after}: [dim]{cfg.db_path}[/dim]"
        )

    conn = db.connect(cfg.db_path)
    try:
        result = db.load_venues_csv(conn, venues_csv)
        counts = db.venue_counts(conn)
    finally:
        conn.close()

    table = Table(title=f"Loaded {venues_csv.name}", title_justify="left")
    table.add_column("", style="dim")
    table.add_column("", justify="right")
    table.add_row("inserted", str(result.inserted))
    table.add_row("updated", str(result.updated))
    if result.skipped:
        table.add_row("skipped", f"[yellow]{result.skipped}[/yellow]")
    table.add_row("venues in DB", str(counts["total"]))
    console.print(table)

    if counts["missing_resy"] or counts["missing_place"]:
        console.print(
            f"[dim]{counts['missing_resy']} venue(s) need a Resy id, "
            f"{counts['missing_place']} need a Google place id — "
            f"run [/dim]resy-rank resolve[dim] next.[/dim]"
        )


@app.command()
def scan(
    ctx: typer.Context,
    target_date: Annotated[
        str, typer.Option("--date", help="Target date, YYYY-MM-DD (start of range with --days-ahead).")
    ],
    party: Annotated[int, typer.Option("--party", help="Party size.")] = 2,
    days_ahead: Annotated[
        int,
        typer.Option("--days-ahead", help="Also scan this many dates after --date."),
    ] = 0,
    venue: Annotated[
        Optional[str],
        typer.Option("--venue", help="Only scan venues whose name contains this (for a single-venue test)."),
    ] = None,
) -> None:
    """Poll Resy availability and persist scans and slots. Never books."""
    context: "Context" = ctx.obj
    cfg = context.settings

    try:
        start = date_cls.fromisoformat(target_date)
    except ValueError:
        raise typer.BadParameter(f"--date must be YYYY-MM-DD, got {target_date!r}")
    if days_ahead < 0:
        raise typer.BadParameter("--days-ahead cannot be negative")
    dates = [(start + timedelta(days=i)).isoformat() for i in range(days_ahead + 1)]

    if not cfg.db_path.exists():
        console.print(f"[red]No database at {cfg.db_path}[/red] — run [bold]resy-rank init[/bold] first.")
        raise typer.Exit(1)

    conn = db.connect(cfg.db_path)
    try:
        venues = db.fetch_venues(conn, name_filter=venue, require_resy_id=True)
        if not venues:
            unresolved = db.fetch_venues(conn, name_filter=venue)
            if unresolved:
                console.print(
                    f"[yellow]{len(unresolved)} matching venue(s) have no Resy id yet[/yellow] — "
                    "run [bold]resy-rank resolve[/bold] first."
                )
            else:
                console.print("[yellow]No venues match.[/yellow]")
            raise typer.Exit(1)

        planned = len(venues) * len(dates)
        console.print(
            f"Scanning [bold]{len(venues)}[/bold] venue(s) × [bold]{len(dates)}[/bold] date(s) "
            f"= {planned} request(s), party of {party}"
        )

        if context.dry_run:
            budget = None
        else:
            budget = DailyBudget(conn, cfg.daily_request_cap)
            remaining = budget.remaining("resy")
            console.print(
                f"[dim]daily cap: {remaining} of {cfg.daily_request_cap} request(s) left today; "
                f"≤{cfg.max_concurrency} concurrent, {cfg.request_delay_seconds}s apart[/dim]"
            )
            if planned > remaining:
                console.print(
                    f"[red]That exceeds today's remaining budget ({remaining}).[/red] "
                    "Narrow the scan or raise DAILY_REQUEST_CAP."
                )
                raise typer.Exit(1)

        outcomes = asyncio.run(_run_scan(context, conn, venues, dates, party, budget))
    finally:
        conn.close()

    _report_scan(outcomes, dry_run=context.dry_run)


async def _run_scan(
    context: "Context",
    conn,
    venues,
    dates: list[str],
    party: int,
    budget: DailyBudget | None,
) -> list[ScanOutcome]:
    cfg = context.settings
    outcomes: list[ScanOutcome] = []

    async with ResyClient(
        cfg,
        budget=budget,
        dry_run=context.dry_run,
        on_intent=lambda msg: console.print(f"[yellow]dry-run:[/yellow] {msg}"),
    ) as client:
        tasks = [
            client.find(
                venue_id=row["venue_id"],
                resy_venue_id=row["resy_venue_id"],
                target_date=day,
                party_size=party,
                lat=row["lat"],
                long=row["lng"],
            )
            for row in venues
            for day in dates
        ]
        # The client's own semaphore caps in-flight requests; gathering all
        # coroutines just queues them behind it.
        try:
            outcomes = await asyncio.gather(*tasks)
        except (DailyCapReached, ResyUnauthorized) as exc:
            console.print(f"[red]Stopped: {exc}[/red]")
            raise typer.Exit(1)

    if not context.dry_run:
        scan_ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
        for outcome in outcomes:
            if outcome.ok or outcome.raw_json:
                persist_scan(conn, outcome, scan_ts)
    return list(outcomes)


def _report_scan(outcomes: list[ScanOutcome], *, dry_run: bool) -> None:
    if dry_run:
        console.print("[yellow]dry-run:[/yellow] nothing was requested or written.")
        return

    failed = [o for o in outcomes if not o.ok]
    with_slots = [o for o in outcomes if o.slots]

    table = Table(title="Scan results", title_justify="left")
    table.add_column("Venue", justify="right", style="dim")
    table.add_column("Date")
    table.add_column("Slots", justify="right")
    table.add_column("Prime", justify="right")
    table.add_column("Times")

    for outcome in sorted(outcomes, key=lambda o: (o.target_date, o.resy_venue_id)):
        if not outcome.ok:
            table.add_row(
                str(outcome.resy_venue_id),
                outcome.target_date,
                "[red]err[/red]",
                "",
                f"[red]{outcome.error[:60]}[/red]",
            )
            continue
        times = ", ".join(dict.fromkeys(s.slot_time for s in outcome.slots)) or "[dim]none[/dim]"
        prime_count = sum(1 for s in outcome.slots if s.is_prime)
        table.add_row(
            str(outcome.resy_venue_id),
            outcome.target_date,
            str(len(outcome.slots)),
            str(prime_count) if prime_count else "[dim]0[/dim]",
            times if len(times) < 70 else times[:67] + "…",
        )

    console.print(table)
    console.print(
        f"[dim]{len(outcomes)} observation(s) recorded; {len(with_slots)} with availability; "
        f"{len(failed)} failed.[/dim]"
    )


if __name__ == "__main__":
    app()
