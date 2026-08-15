# resy-rank

A personal, **read-only** research tool. It scans Resy for availability at times
worth eating at, then ranks only the bookable options by a composite quality
score.

It never books, holds, modifies, or cancels a reservation — there is no code
path to Resy's booking endpoints — and it deliberately does not fire at
reservation-drop time.

## Setup

```bash
uv sync
cp .env.example .env   # fill in RESY_API_KEY / RESY_AUTH_TOKEN
uv run resy-rank init
```

## Workflow

```bash
uv run resy-rank init                       # create DB, load data/venues.csv
uv run resy-rank critics                    # match data/critics.csv to venues
uv run resy-rank scan --date 2026-09-05 --geo   # sweep availability
uv run resy-rank resolve                    # Google enrichment (free tier by default)
uv run resy-rank rank --date 2026-09-05 --party 2 --time 19:30 --window 60
uv run resy-rank backfill --days 30         # bootstrap the scarcity index
```

`--dry-run` is global, hits nothing over the network, needs no credentials, and
spends nothing — use it to review the exact request plan first.

## Two scanning modes

**Geo sweep (`--geo`)** — `/4/find` without a `venue_id` returns every venue
with availability near a point, which is the same query resy.com makes when you
browse. Nine anchor points cover Manhattan in **9 requests per date** instead of
one request per venue (~2,000). It also discovers venues automatically, so
`venues.csv` is a seed list rather than the whole universe, and a sweep hit
fills in a seed row's Resy id for free.

**Targeted (`--venue`)** — one request per venue per date. Use it for the
handful of places you care about that are *always* booked out, since those never
appear in a sweep's results and would otherwise stay invisible.

A venue that is in sweep scope but absent from a sweep's results is recorded as
an observation with zero slots. That absence is the scarcity signal, so it is
stored deliberately rather than skipped.

## Cost

**Resy: free**, and the request volume is the thing actually worth controlling.
Concurrency is capped at 3, requests are spaced `REQUEST_DELAY_SECONDS` apart
globally, 429/5xx get exponential backoff (2s → 16s, jittered, honoring
`Retry-After`), and a daily cap persisted in SQLite holds across separate runs.
A 401/403 aborts immediately rather than retrying a dead token. Outbound URLs
are checked against an allowlist of exactly two endpoints before any request
leaves; Resy's booking endpoints are absent, so an accidental call raises
instead of reserving a table.

**Google Places: $0 by default.** Since March 2025, Google bills per SKU with a
per-SKU monthly free tier (10K Essentials / 5K Pro / 1K Enterprise) that does
not pool. Requesting `rating` forces the Enterprise tier, so:

- resolution (name → place_id) uses a separate minimal field mask, drawing on
  its own free allowance;
- enrichment is capped at `MONTHLY_ENRICHMENT_CAP` (default 1,000 = the free
  tier), so a large universe spreads across months at no cost instead of
  silently running up a bill;
- results are cached permanently and only re-fetched with `--refresh` or after
  `GOOGLE_CACHE_DAYS` (180).

Every run reports its billable call count. For a ~2,000-venue Manhattan
universe that's about $25 one-time if you lift the cap, or $0 spread over two
months. Verify current rates in your own billing console — Google's pricing
pages moved recently.

To skip Google entirely, set `WEIGHT_RATING=0`; the composite reweights itself.

## Scoring

Three independent 0–100 signals, weighted (defaults 0.35 / 0.35 / 0.30):

1. **Bayesian-adjusted Google rating** — `(v·R + m·C) / (v + m)`, where `C` is
   the mean rating across the venue universe (computed, never hardcoded) and
   `m` is the prior weight (default 300). Min-max normalized across the
   universe afterwards, because raw ratings compress into 4.3–4.7 and are
   useless sorted directly. A 5.0 from twelve people loses to a 4.7 from eight
   thousand — that's the point.
2. **Critic score** — points per flag from `data/critics.csv`, stacking but
   saturating at 100. Defaults: michelin_star 100, michelin_bib 70, eater_38 70,
   nyt 50, infatuation 40.
3. **Scarcity index** — the share of *prime* observations (Thu–Sat, 18:30–20:30)
   that found nothing bookable, over the trailing 30 days of your own scan
   history, counting only observations made 14–28 days ahead of the date. Below
   `SCARCITY_MIN_OBSERVATIONS` (20) it returns `None`.

**Missing components are redistributed, not zeroed.** A component that cannot be
computed has its weight spread across the survivors in proportion. Dropping
scarcity from the default weights leaves rating and critic splitting the
composite 50/50, so a venue with too little history is not silently punished for
being new to your dataset. Zero means "bad"; `None` means "unknown"; conflating
them would produce a confident-looking ranking that is wrong.

The `rank` table shows all three components alongside the composite, so any
placement can be explained.

Every weight and constant lives in `.env` / `resy_rank/config.py`.

## Data files

- `data/venues.csv` — seed venues (hand-maintained). Geo sweeps add to it.
- `data/critics.csv` — `venue_name,source,tier,url`. Names are fuzzy-matched
  with rapidfuzz, and every unmatched row is printed with its closest candidate,
  since a silent miss removes up to a third of a venue's score.

> **The shipped CSVs are scaffolding, not data.** The venue list and especially
> the Michelin tiers were written from memory and are not verified. Replace them
> with your own before trusting a ranking.

## Tests

```bash
uv run pytest
```

126 tests, no network. `tests/test_scoring.py` pins the Bayesian shrinkage and
the weight-redistribution rules against hand-computed cases.

## Note on the rest of this repo

`main.py`, `booker.py`, `resy_client.py`, and `notifier.py` are a separate,
earlier auto-booking tool. `resy_rank/` shares no code with them.
