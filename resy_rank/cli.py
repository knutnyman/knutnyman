"""Typer entry point for resy-rank.

Read-only by construction: no command in this tool books, holds, modifies, or
cancels a reservation.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console
from rich.table import Table

from resy_rank import db
from resy_rank.config import get_settings

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


if __name__ == "__main__":
    app()
