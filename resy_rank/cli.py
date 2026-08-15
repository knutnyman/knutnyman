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

from resy_rank import critics, db, places, scoring
from resy_rank.config import get_settings
from resy_rank.resy import (
    DailyBudget,
    DailyCapReached,
    ResyClient,
    ResyUnauthorized,
    ScanOutcome,
    merge_geo_results,
    persist_geo_sweep,
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


def _run_async(coro):
    """Run a coroutine, turning setup failures into a message not a traceback."""
    try:
        return asyncio.run(coro)
    except RuntimeError as exc:
        console.print(f"[red]{exc}[/red]")
        raise typer.Exit(1) from None


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
def resolve(
    ctx: typer.Context,
    refresh: Annotated[
        bool, typer.Option("--refresh", help="Re-fetch Google data even if cached and fresh.")
    ] = False,
    limit: Annotated[
        Optional[int],
        typer.Option("--limit", help="Stop after enriching this many venues (spend control)."),
    ] = None,
    skip_resy: Annotated[
        bool, typer.Option("--skip-resy", help="Only do Google enrichment, no Resy lookups.")
    ] = False,
) -> None:
    """Fill in missing Resy venue ids and Google Places data. Cached, capped, read-only."""
    context: "Context" = ctx.obj
    cfg = context.settings

    if not cfg.db_path.exists():
        console.print(f"[red]No database at {cfg.db_path}[/red] — run [bold]resy-rank init[/bold] first.")
        raise typer.Exit(1)

    conn = db.connect(cfg.db_path)
    try:
        _run_async(_run_resolve(context, conn, refresh=refresh, limit=limit, skip_resy=skip_resy))
    finally:
        conn.close()


async def _run_resolve(
    context: "Context", conn, *, refresh: bool, limit: int | None, skip_resy: bool
) -> None:
    cfg = context.settings
    intent = lambda msg: console.print(f"[yellow]dry-run:[/yellow] {msg}")  # noqa: E731

    # ── 1. Resy venue ids ───────────────────────────────────────────────────
    if not skip_resy:
        missing = [r for r in db.fetch_venues(conn) if r["resy_venue_id"] is None]
        if missing:
            console.print(f"Resolving [bold]{len(missing)}[/bold] Resy venue id(s)…")
            resy_budget = None if context.dry_run else DailyBudget(conn, cfg.daily_request_cap)
            async with ResyClient(
                cfg, budget=resy_budget, dry_run=context.dry_run, on_intent=intent
            ) as client:
                for row in missing:
                    try:
                        hits = await client.search_venues(row["name"])
                    except (DailyCapReached, ResyUnauthorized) as exc:
                        console.print(f"[red]Stopped: {exc}[/red]")
                        raise typer.Exit(1)
                    except Exception as exc:
                        console.print(f"  [red]![/red] {row['name']}: {exc}")
                        continue
                    if not hits:
                        if not context.dry_run:
                            console.print(f"  [yellow]?[/yellow] {row['name']}: no Resy match")
                        continue
                    best = hits[0]
                    with conn:
                        conn.execute(
                            "UPDATE venues SET resy_venue_id = ? WHERE venue_id = ?",
                            (best["resy_venue_id"], row["venue_id"]),
                        )
                    console.print(
                        f"  [green]✓[/green] {row['name']} → {best['resy_venue_id']} "
                        f"[dim]({best['name']})[/dim]"
                    )
        else:
            console.print("[dim]All venues already have a Resy id.[/dim]")

    # ── 2. Google Places ────────────────────────────────────────────────────
    pending = places.venues_needing_enrichment(conn, cfg.google_cache_days, refresh=refresh)
    if not pending:
        console.print("[dim]Google data is fresh for every venue — nothing to fetch.[/dim]")
        return

    if limit is not None:
        pending = pending[:limit]

    monthly = None if context.dry_run else places.MonthlyBudget(conn, cfg.monthly_enrichment_cap)
    if monthly is not None:
        left = monthly.remaining(places.SERVICE_DETAILS)
        console.print(
            f"Enriching [bold]{len(pending)}[/bold] venue(s). "
            f"[dim]Enterprise-SKU budget: {left} of {cfg.monthly_enrichment_cap} left this month.[/dim]"
        )
        if len(pending) > left:
            console.print(
                f"[yellow]Only {left} will be fetched this month[/yellow] — the rest stay "
                "queued so this run costs nothing. Re-run next month, or raise "
                "MONTHLY_ENRICHMENT_CAP to pay for the remainder now."
            )
    else:
        console.print(f"Enriching [bold]{len(pending)}[/bold] venue(s).")

    enriched = failed = 0
    async with places.PlacesClient(
        cfg, budget=monthly, dry_run=context.dry_run, on_intent=intent
    ) as client:
        for row in pending:
            try:
                place_id = row["google_place_id"]
                if not place_id or refresh:
                    place_id = await client.resolve_place_id(row["name"], row["address"] or "")
                    if place_id and not context.dry_run:
                        places.save_place_id(conn, row["venue_id"], place_id)
                if not place_id:
                    failed += 1
                    continue

                details = await client.fetch_details(place_id)
                if details is None:
                    if not context.dry_run:
                        failed += 1
                    continue
                places.save_details(conn, row["venue_id"], details)
                enriched += 1
                console.print(
                    f"  [green]✓[/green] {row['name']}: {details.rating} "
                    f"({details.user_rating_count} ratings)"
                )
            except places.MonthlyCapReached as exc:
                console.print(f"[yellow]Stopping: {exc}[/yellow]")
                break
            except Exception as exc:
                failed += 1
                console.print(f"  [red]![/red] {row['name']}: {exc}")

    if context.dry_run:
        console.print("[yellow]dry-run:[/yellow] nothing was requested, written, or billed.")
        return

    console.print(
        f"Enriched {enriched}, failed {failed}. "
        f"[dim]Billable calls this run: {client.calls[places.SERVICE_TEXT]} text search, "
        f"{client.calls[places.SERVICE_DETAILS]} Enterprise details "
        f"({client.billable_total} total).[/dim]"
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
    geo: Annotated[
        bool,
        typer.Option(
            "--geo",
            help="Sweep whole neighborhoods instead of polling venue by venue. "
            "Far fewer requests; discovers venues automatically.",
        ),
    ] = False,
) -> None:
    """Poll Resy availability and persist scans and slots. Never books."""
    context: "Context" = ctx.obj
    cfg = context.settings

    if geo and venue:
        raise typer.BadParameter("--geo sweeps by area, so it cannot be combined with --venue")

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
        if geo:
            _scan_geo(context, conn, dates, party)
            return

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

        outcomes = _run_async(_run_scan(context, conn, venues, dates, party, budget))
    finally:
        conn.close()

    _report_scan(outcomes, dry_run=context.dry_run)


def _scan_geo(context: "Context", conn, dates: list[str], party: int) -> None:
    """City-wide sweep: one request per anchor per date instead of per venue."""
    cfg = context.settings
    anchors = cfg.geo_anchors
    planned = len(anchors) * len(dates)

    console.print(
        f"Geo sweep: [bold]{len(anchors)}[/bold] anchor(s) × [bold]{len(dates)}[/bold] date(s) "
        f"≈ {planned} request(s) (more only if pagination kicks in), party of {party}"
    )

    budget = None
    if not context.dry_run:
        budget = DailyBudget(conn, cfg.daily_request_cap)
        remaining = budget.remaining("resy")
        console.print(
            f"[dim]daily cap: {remaining} of {cfg.daily_request_cap} request(s) left today[/dim]"
        )
        if planned > remaining:
            console.print(
                f"[red]That exceeds today's remaining budget ({remaining}).[/red] "
                "Scan fewer dates or raise DAILY_REQUEST_CAP."
            )
            raise typer.Exit(1)

    summary, failures = _run_async(_run_geo_scan(context, conn, dates, party, budget))

    if context.dry_run:
        console.print("[yellow]dry-run:[/yellow] nothing was requested or written.")
        return

    table = Table(title="Geo sweep", title_justify="left")
    table.add_column("Date")
    table.add_column("Venues w/ availability", justify="right")
    table.add_column("Observed unavailable", justify="right")
    table.add_column("Prime slots", justify="right")
    for row in summary:
        table.add_row(
            row["date"], str(row["available"]), str(row["unavailable"]), str(row["prime"])
        )
    console.print(table)

    if failures:
        total_planned = len(cfg.geo_anchors) * len(dates)
        console.print(
            f"[red]{failures} of {total_planned} request(s) failed.[/red] "
            + (
                "Every request failed, so this sweep recorded nothing — the zeros above "
                "mean 'no data', not 'no availability'. Check your network and "
                "RESY_API_KEY, then re-run."
                if failures >= total_planned
                else "The venues behind those requests were not observed and are absent "
                "from the counts above."
            )
        )

    counts = db.venue_counts(conn)
    console.print(
        f"[dim]{counts['total']} venue(s) known; {counts['missing_place']} still need Google "
        f"enrichment — run [/dim]resy-rank resolve[dim].[/dim]"
    )


async def _run_geo_scan(
    context: "Context", conn, dates: list[str], party: int, budget: DailyBudget | None
) -> tuple[list[dict], int]:
    cfg = context.settings
    summary: list[dict] = []
    failures = 0

    async with ResyClient(
        cfg,
        budget=budget,
        dry_run=context.dry_run,
        on_intent=lambda msg: console.print(f"[yellow]dry-run:[/yellow] {msg}"),
    ) as client:
        for day in dates:
            try:
                batches = await asyncio.gather(
                    *(
                        client.find_geo(lat=lat, long=lng, target_date=day, party_size=party)
                        for lat, lng in cfg.geo_anchors
                    )
                )
            except (DailyCapReached, ResyUnauthorized) as exc:
                console.print(f"[red]Stopped: {exc}[/red]")
                raise typer.Exit(1)

            if context.dry_run:
                continue

            venues = merge_geo_results(batches)
            scan_ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
            counts = persist_geo_sweep(
                conn, target_date=day, party_size=party, venues=venues, scan_ts=scan_ts
            )
            counts["date"] = day
            counts["prime"] = sum(1 for v in venues for s in v.slots if s.is_prime)
            summary.append(counts)
        failures = client.failed_requests
    return summary, failures


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


@app.command(name="critics")
def critics_cmd(
    ctx: typer.Context,
    path: Annotated[
        Optional[Path], typer.Option("--file", help="Critic CSV (default: data/critics.csv).")
    ] = None,
    threshold: Annotated[
        Optional[int], typer.Option("--threshold", help="Fuzzy match score 0-100.")
    ] = None,
) -> None:
    """Load critic lists and fuzzy-match them to venues. Prints every miss."""
    context: "Context" = ctx.obj
    cfg = context.settings
    csv_path = path or cfg.critics_csv
    cutoff = threshold if threshold is not None else cfg.critic_match_threshold

    conn = db.connect(cfg.db_path)
    try:
        rows = critics.load_critics_csv(csv_path)
        venues = db.fetch_venues(conn)
        report = critics.match_critics(
            rows, venues, threshold=cutoff, known_sources=set(cfg.critic_points)
        )

        if not context.dry_run:
            critics.save_flags(conn, report.matched)

        console.print(
            f"Matched [green]{len(report.matched)}[/green] of {len(rows)} critic row(s) "
            f"at threshold {cutoff}."
        )

        if report.unknown_sources:
            console.print(
                f"\n[yellow]Unknown source(s)[/yellow] — these score 0; known sources are "
                f"{', '.join(sorted(cfg.critic_points))}:"
            )
            for row in report.unknown_sources:
                console.print(f"  {csv_path.name}:{row.lineno}  {row.venue_name} → {row.source}")

        if report.unmatched:
            table = Table(title="\nUnmatched rows — fix these in the CSV", title_justify="left")
            table.add_column("Line", justify="right", style="dim")
            table.add_column("Critic name")
            table.add_column("Source")
            table.add_column("Closest venue")
            table.add_column("Score", justify="right")
            for row, closest, score in report.unmatched:
                table.add_row(
                    str(row.lineno),
                    row.venue_name,
                    row.source,
                    closest or "[dim]—[/dim]",
                    f"{score:.0f}",
                )
            console.print(table)

            if len(report.unmatched) > len(report.matched):
                console.print(
                    "\n[dim]Most rows are unmatched because those venues aren't in the "
                    "database yet. Critic lists cover the whole city; your venue table "
                    "only fills up as sweeps discover it. Run [/dim]"
                    "resy-rank scan --geo[dim] first, then re-run this — it is safe to "
                    "re-run any number of times.[/dim]"
                )

        if context.dry_run:
            console.print("[yellow]dry-run:[/yellow] no flags were written.")
    finally:
        conn.close()


@app.command()
def rank(
    ctx: typer.Context,
    target_date: Annotated[str, typer.Option("--date", help="Date to rank, YYYY-MM-DD.")],
    party: Annotated[int, typer.Option("--party", help="Party size.")] = 2,
    at_time: Annotated[str, typer.Option("--time", help="Centre of the window, HH:MM.")] = "19:30",
    window: Annotated[
        int, typer.Option("--window", help="Half-width of the window, in minutes.")
    ] = 60,
    limit: Annotated[int, typer.Option("--limit", help="Rows to show.")] = 25,
) -> None:
    """Rank the venues that are actually bookable in your window. The main command."""
    context: "Context" = ctx.obj
    cfg = context.settings

    try:
        centre = datetime.strptime(at_time, "%H:%M")
    except ValueError:
        raise typer.BadParameter(f"--time must be HH:MM, got {at_time!r}")
    try:
        date_cls.fromisoformat(target_date)
    except ValueError:
        raise typer.BadParameter(f"--date must be YYYY-MM-DD, got {target_date!r}")

    start = (centre - timedelta(minutes=window)).strftime("%H:%M")
    end = (centre + timedelta(minutes=window)).strftime("%H:%M")

    if not cfg.db_path.exists():
        console.print(f"[red]No database at {cfg.db_path}[/red] — run [bold]resy-rank init[/bold] first.")
        raise typer.Exit(1)

    conn = db.connect(cfg.db_path)
    try:
        available = db.available_in_window(
            conn,
            target_date=target_date,
            party_size=party,
            start_time=start,
            end_time=end,
        )
        if not available:
            console.print(
                f"[yellow]Nothing bookable[/yellow] for {target_date}, party of {party}, "
                f"between {start} and {end}."
            )
            console.print(
                f"[dim]If you haven't scanned that date yet: "
                f"resy-rank scan --date {target_date} --party {party} --geo[/dim]"
            )
            raise typer.Exit(0)

        # Scores are normalized across the whole universe, then filtered — so a
        # venue's rating score means the same thing regardless of the window.
        scores = scoring.score_universe(conn, cfg)
    finally:
        conn.close()

    rows = [
        (scores[venue_id], slots)
        for venue_id, slots in available.items()
        if venue_id in scores
    ]
    rows.sort(key=lambda pair: (pair[0].composite is None, -(pair[0].composite or 0)))

    table = Table(
        title=f"{target_date} · party of {party} · {start}–{end}",
        title_justify="left",
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Venue")
    table.add_column("Neighborhood", style="dim")
    table.add_column("Times")
    table.add_column("Score", justify="right")
    table.add_column("Rating", justify="right")
    table.add_column("Critic", justify="right")
    table.add_column("Scarce", justify="right")

    for index, (score, slots) in enumerate(rows[:limit], start=1):
        times = ", ".join(dict.fromkeys(time for time, _ in slots))
        table.add_row(
            str(index),
            score.name,
            score.neighborhood or "—",
            times if len(times) <= 34 else times[:31] + "…",
            f"[bold]{score.composite:.1f}[/bold]" if score.composite is not None else "[dim]n/a[/dim]",
            f"{score.rating_score:.0f}" if score.rating_score is not None else "[dim]n/a[/dim]",
            f"{score.critic_score:.0f}" if score.critic_score is not None else "[dim]n/a[/dim]",
            score.scarcity_label if score.scarcity_score is not None else f"[dim]{score.scarcity_label}[/dim]",
        )

    console.print(table)

    partial = sum(
        1
        for score, _ in rows[:limit]
        if score.rating_score is None or score.scarcity_score is None
    )
    console.print(
        f"[dim]{len(rows)} venue(s) bookable in this window. "
        f"Weights: rating {cfg.weight_rating}, critic {cfg.weight_critic}, "
        f"scarcity {cfg.weight_scarcity}.[/dim]"
    )
    if partial:
        console.print(
            f"[dim]{partial} row(s) are scored on partial data — a missing component's "
            f"weight is redistributed, not counted as zero.[/dim]"
        )


@app.command()
def backfill(
    ctx: typer.Context,
    days: Annotated[int, typer.Option("--days", help="How many days ahead to cover.")] = 30,
    party: Annotated[int, typer.Option("--party", help="Party size.")] = 2,
    all_days: Annotated[
        bool,
        typer.Option(
            "--all-days",
            help="Scan every date in range, not just the ones that feed the scarcity index.",
        ),
    ] = False,
) -> None:
    """Sweep a date range to bootstrap the scarcity index."""
    context: "Context" = ctx.obj
    cfg = context.settings

    today = date_cls.today()
    candidates = [today + timedelta(days=offset) for offset in range(1, days + 1)]

    if all_days:
        dates = [d.isoformat() for d in candidates]
    else:
        # Only dates that can actually produce a scarcity observation: a prime
        # weekday, seen inside the 14-28 day horizon. Scanning the rest costs
        # requests and contributes nothing to the index.
        dates = [
            d.isoformat()
            for d in candidates
            if d.weekday() in cfg.prime_days
            and cfg.scarcity_horizon_min_days <= (d - today).days <= cfg.scarcity_horizon_max_days
        ]

    if not dates:
        console.print(
            f"[yellow]No qualifying dates in the next {days} day(s).[/yellow] "
            f"The scarcity index reads prime days ({', '.join(_day_names(cfg.prime_days))}) "
            f"observed {cfg.scarcity_horizon_min_days}-{cfg.scarcity_horizon_max_days} days out — "
            f"try --days {cfg.scarcity_horizon_max_days} or --all-days."
        )
        raise typer.Exit(0)

    console.print(
        f"Backfilling [bold]{len(dates)}[/bold] date(s): {dates[0]} … {dates[-1]}"
        + ("" if all_days else " [dim](prime dates inside the scarcity horizon)[/dim]")
    )

    if not cfg.db_path.exists():
        console.print(f"[red]No database at {cfg.db_path}[/red] — run [bold]resy-rank init[/bold] first.")
        raise typer.Exit(1)

    conn = db.connect(cfg.db_path)
    try:
        _scan_geo(context, conn, dates, party)
    finally:
        conn.close()


_DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def _day_names(days: list[int]) -> list[str]:
    return [_DAY_NAMES[d] for d in sorted(days)]


if __name__ == "__main__":
    app()
