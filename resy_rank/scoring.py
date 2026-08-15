"""Composite ranking math.

Three independent signals, each normalized to 0-100, then weighted:

  a) Bayesian-adjusted Google rating — shrinks thin-sample ratings toward the
     universe mean, then min-max normalizes to undo the 4.3-4.7 compression
     that makes raw ratings useless when sorted directly.
  b) Critic score — points per flag, stacking but saturating at a cap.
  c) Scarcity index — how often a venue's prime slots were already gone when
     observed two to four weeks out, from your own scan history.

A component that cannot be computed returns None, and its weight is
redistributed across the components that *can* be — never treated as zero. The
distinction matters: zero says "this venue is bad", None says "we don't know",
and conflating them would push every unknown venue to the bottom of the table
while looking like a real ranking.

Everything here is a pure function of its arguments so the math is testable
without a database or a network.
"""

from __future__ import annotations

import logging
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone

from resy_rank.config import Settings

log = logging.getLogger(__name__)

RATING = "rating"
CRITIC = "critic"
SCARCITY = "scarcity"


# ── a) Bayesian-adjusted rating ─────────────────────────────────────────────


def bayesian_adjust(
    rating: float, rating_count: int, prior_mean: float, prior_weight: float
) -> float:
    """(v*R + m*C) / (v + m).

    A venue with few ratings is pulled toward the universe mean C; one with
    many keeps its own rating R. `m` is how many average ratings a venue is
    charged before its own score counts at full strength.
    """
    v = max(0, rating_count)
    denominator = v + prior_weight
    if denominator <= 0:
        return prior_mean
    return (v * rating + prior_weight * prior_mean) / denominator


def universe_mean_rating(ratings: list[float]) -> float | None:
    """C — computed from the venue universe, never hardcoded."""
    if not ratings:
        return None
    return sum(ratings) / len(ratings)


def minmax_normalize(values: dict[int, float]) -> dict[int, float]:
    """Spread values across 0-100.

    When every value is identical the range is zero and there is no meaningful
    spread, so everything ties at 50 — a neutral middle rather than a
    misleading 0 or 100.
    """
    if not values:
        return {}
    lo = min(values.values())
    hi = max(values.values())
    if hi == lo:
        return {key: 50.0 for key in values}
    span = hi - lo
    return {key: (value - lo) / span * 100.0 for key, value in values.items()}


def rating_scores(
    venue_ratings: dict[int, tuple[float, int]], prior_weight: float
) -> dict[int, float]:
    """venue_id -> (rating, count) mapped to 0-100 Bayesian-adjusted scores.

    Venues with no Google rating must simply be absent from the input; they
    score None (handled by the caller) rather than zero.
    """
    if not venue_ratings:
        return {}
    prior_mean = universe_mean_rating([r for r, _ in venue_ratings.values()])
    if prior_mean is None:
        return {}
    adjusted = {
        venue_id: bayesian_adjust(rating, count, prior_mean, prior_weight)
        for venue_id, (rating, count) in venue_ratings.items()
    }
    return minmax_normalize(adjusted)


# ── b) Critic score ─────────────────────────────────────────────────────────


def critic_score(sources: list[str], points: dict[str, float], cap: float) -> float:
    """Sum points per flag, saturating at `cap`. Unknown sources score nothing."""
    total = 0.0
    for source in set(sources):
        value = points.get(source)
        if value is None:
            log.debug("unknown critic source %r scores 0", source)
            continue
        total += value
    return min(total, cap)


# ── c) Scarcity index ───────────────────────────────────────────────────────


def scarcity_score(
    unavailable_observations: int, total_observations: int, min_observations: int
) -> float | None:
    """Share of prime observations that found nothing available, as 0-100.

    Returns None below the minimum observation count. A venue seen three times
    and booked out all three is not "100% scarce" in any useful sense — it is
    unmeasured, and saying so is more honest than inventing a number.
    """
    if total_observations < min_observations or total_observations <= 0:
        return None
    share = unavailable_observations / total_observations
    return max(0.0, min(1.0, share)) * 100.0


# ── Composite ───────────────────────────────────────────────────────────────


def composite_score(
    components: dict[str, float | None], weights: dict[str, float]
) -> float | None:
    """Weighted mean over the components that exist, renormalized.

    This is the redistribution rule: a missing component's weight is spread
    across the survivors in proportion to their own weights, so dropping
    scarcity (0.30) from the default 0.35/0.35/0.30 leaves rating and critic
    splitting the composite 50/50 rather than the whole score collapsing.

    Returns None when nothing at all is known.
    """
    available = {
        name: value
        for name, value in components.items()
        if value is not None and weights.get(name, 0.0) > 0
    }
    if not available:
        return None
    total_weight = sum(weights[name] for name in available)
    if total_weight <= 0:
        return None
    return sum(value * weights[name] for name, value in available.items()) / total_weight


# ── Assembly ────────────────────────────────────────────────────────────────


@dataclass
class VenueScore:
    venue_id: int
    name: str
    neighborhood: str
    composite: float | None
    rating_score: float | None
    critic_score: float | None
    scarcity_score: float | None
    # Diagnostics, so a rank can always be explained.
    raw_rating: float | None = None
    rating_count: int | None = None
    critic_sources: tuple[str, ...] = ()
    scarcity_observations: int = 0

    @property
    def scarcity_label(self) -> str:
        if self.scarcity_score is not None:
            return f"{self.scarcity_score:.0f}"
        return f"n/a ({self.scarcity_observations} obs)"


def _sqlite_prime_days(prime_days: list[int]) -> list[str]:
    """Python weekday (Mon=0) -> SQLite strftime('%w') (Sun=0)."""
    return [str((day + 1) % 7) for day in prime_days]


def scarcity_observations(
    conn: sqlite3.Connection, settings: Settings, now: datetime | None = None
) -> dict[int, tuple[int, int]]:
    """venue_id -> (unavailable_observations, total_observations).

    An observation is one scan of one venue for one prime date, made 14-28 days
    ahead of that date, within the trailing lookback window. It counts as
    "unavailable" when the scan recorded no prime slot for that venue.

    Note this is observation-level, not slot-level: the unit is "was there
    anything bookable in the prime window when we looked", which is robust to
    venues that publish different numbers of seatings.
    """
    reference = now or datetime.now(timezone.utc)
    cutoff = (reference - timedelta(days=settings.scarcity_lookback_days)).isoformat(
        timespec="seconds"
    )
    day_placeholders = ",".join("?" * len(settings.prime_days))

    sql = f"""
        SELECT s.venue_id                                   AS venue_id,
               COUNT(*)                                     AS total,
               SUM(
                   CASE WHEN NOT EXISTS (
                       SELECT 1 FROM slots sl
                       WHERE sl.scan_id = s.scan_id AND sl.is_prime = 1
                   ) THEN 1 ELSE 0 END
               )                                            AS unavailable
        FROM scans s
        WHERE s.scan_ts >= ?
          AND strftime('%w', s.target_date) IN ({day_placeholders})
          AND (julianday(s.target_date) - julianday(date(s.scan_ts)))
              BETWEEN ? AND ?
        GROUP BY s.venue_id
    """
    params = [
        cutoff,
        *_sqlite_prime_days(settings.prime_days),
        settings.scarcity_horizon_min_days,
        settings.scarcity_horizon_max_days,
    ]
    return {
        row["venue_id"]: (row["unavailable"] or 0, row["total"] or 0)
        for row in conn.execute(sql, params)
    }


def score_universe(
    conn: sqlite3.Connection, settings: Settings, now: datetime | None = None
) -> dict[int, VenueScore]:
    """Score every venue. Normalization is across the whole universe by design."""
    from resy_rank.critics import flags_by_venue

    venues = conn.execute(
        """
        SELECT v.venue_id, v.name, v.neighborhood, g.rating, g.user_rating_count
        FROM venues v
        LEFT JOIN google_meta g ON g.venue_id = v.venue_id
        """
    ).fetchall()

    venue_ratings = {
        row["venue_id"]: (row["rating"], row["user_rating_count"] or 0)
        for row in venues
        if row["rating"] is not None
    }
    normalized_ratings = rating_scores(venue_ratings, settings.bayesian_prior_weight)

    flags = flags_by_venue(conn)
    observations = scarcity_observations(conn, settings, now=now)

    weights = {
        RATING: settings.weight_rating,
        CRITIC: settings.weight_critic,
        SCARCITY: settings.weight_scarcity,
    }

    scores: dict[int, VenueScore] = {}
    for row in venues:
        venue_id = row["venue_id"]
        sources = flags.get(venue_id, [])
        unavailable, total = observations.get(venue_id, (0, 0))

        components: dict[str, float | None] = {
            RATING: normalized_ratings.get(venue_id),
            # No critic flags is a real, informative zero — the venue is on no
            # list — unlike a missing rating, which is simply unknown.
            CRITIC: critic_score(sources, settings.critic_points, settings.critic_score_cap),
            SCARCITY: scarcity_score(unavailable, total, settings.scarcity_min_observations),
        }

        scores[venue_id] = VenueScore(
            venue_id=venue_id,
            name=row["name"],
            neighborhood=row["neighborhood"] or "",
            composite=composite_score(components, weights),
            rating_score=components[RATING],
            critic_score=components[CRITIC],
            scarcity_score=components[SCARCITY],
            raw_rating=row["rating"],
            rating_count=row["user_rating_count"],
            critic_sources=tuple(sorted(sources)),
            scarcity_observations=total,
        )
    return scores
