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

## Commands

| Command | Status |
| --- | --- |
| `init` — create the DB, load `data/venues.csv` | ✅ milestone 1 |
| `scan` — poll Resy, persist scans + slots | milestone 2 |
| `resolve` — fill in Resy and Google Place ids | milestone 3 |
| `rank` — the main command | milestone 5 |
| `backfill` — bootstrap the scarcity index | milestone 6 |

`--dry-run` is global and hits nothing over the network.

## Scoring

Three independent 0–100 signals, weighted (defaults 0.35 / 0.35 / 0.30):

1. **Bayesian-adjusted Google rating** — `(v·R + m·C) / (v + m)`, where `C` is
   the mean rating across the venue universe (computed, never hardcoded) and
   `m` is the prior weight. Min-max normalized across the universe afterwards,
   because raw ratings compress into 4.3–4.7 and are useless sorted directly.
2. **Critic score** — points per flag from `data/critics.csv`, stacking but
   saturating at 100.
3. **Scarcity index** — share of *prime* slots (Thu–Sat, 18:30–20:30) that were
   already gone when observed 14–28 days out, over the trailing 30 days of your
   own scan history. Below the minimum observation count it scores `None`, and
   its weight is redistributed across the other two rather than counted as zero.

Every weight and constant lives in `.env` / `resy_rank/config.py`.

## Data files

- `data/venues.csv` — the venue universe (hand-maintained).
- `data/critics.csv` — `venue_name,source,tier,url`; sources are
  `michelin_star`, `michelin_bib`, `eater_38`, `infatuation`, `nyt`.

## Note on the rest of this repo

`main.py`, `booker.py`, `resy_client.py`, and `notifier.py` are a separate,
earlier auto-booking tool. `resy_rank/` shares nothing with them.
