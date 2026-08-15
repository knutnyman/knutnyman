"""Scoring math, verified against hand-computed cases.

The two things that most need pinning down: the Bayesian shrinkage, and the
weight redistribution when a component is None.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from resy_rank import db
from resy_rank.config import Settings
from resy_rank.scoring import (
    CRITIC,
    RATING,
    SCARCITY,
    bayesian_adjust,
    composite_score,
    critic_score,
    minmax_normalize,
    rating_scores,
    scarcity_observations,
    scarcity_score,
    score_universe,
    universe_mean_rating,
)

WEIGHTS = {RATING: 0.35, CRITIC: 0.35, SCARCITY: 0.30}
POINTS = {
    "michelin_star": 100.0,
    "michelin_bib": 70.0,
    "eater_38": 70.0,
    "nyt": 50.0,
    "infatuation": 40.0,
}


# ── Bayesian shrinkage ──────────────────────────────────────────────────────


def test_bayesian_hand_computed():
    # v=100, R=4.9, m=300, C=4.5
    #   (100*4.9 + 300*4.5) / 400 = (490 + 1350) / 400 = 1840/400 = 4.6
    assert bayesian_adjust(4.9, 100, prior_mean=4.5, prior_weight=300) == pytest.approx(4.6)


def test_bayesian_thin_sample_pulled_hard_toward_prior():
    # v=10, R=5.0, m=300, C=4.5
    #   (10*5.0 + 300*4.5) / 310 = (50 + 1350)/310 = 1400/310 = 4.516129...
    assert bayesian_adjust(5.0, 10, 4.5, 300) == pytest.approx(4.5161290, abs=1e-6)


def test_bayesian_large_sample_keeps_own_rating():
    # v=30000, R=4.8, m=300, C=4.5 -> (144000 + 1350)/30300 = 4.797029...
    assert bayesian_adjust(4.8, 30000, 4.5, 300) == pytest.approx(4.7970297, abs=1e-6)


def test_bayesian_zero_count_is_exactly_the_prior():
    assert bayesian_adjust(5.0, 0, 4.5, 300) == 4.5


def test_bayesian_a_perfect_five_from_twelve_people_loses_to_a_solid_four_seven():
    """The whole reason shrinkage exists."""
    hyped = bayesian_adjust(5.0, 12, prior_mean=4.5, prior_weight=300)
    established = bayesian_adjust(4.7, 8000, prior_mean=4.5, prior_weight=300)
    assert established > hyped


def test_prior_weight_controls_shrinkage_strength():
    weak = bayesian_adjust(5.0, 100, 4.5, prior_weight=10)
    strong = bayesian_adjust(5.0, 100, 4.5, prior_weight=1000)
    assert weak > strong  # a smaller m lets the venue's own rating dominate


def test_universe_mean_is_computed_not_assumed():
    assert universe_mean_rating([4.4, 4.6, 4.8]) == pytest.approx(4.6)
    assert universe_mean_rating([]) is None


# ── Min-max normalization ───────────────────────────────────────────────────


def test_minmax_spreads_to_full_range():
    out = minmax_normalize({1: 4.5, 2: 4.6, 3: 4.7})
    assert out == {1: pytest.approx(0.0), 2: pytest.approx(50.0), 3: pytest.approx(100.0)}


def test_minmax_all_identical_ties_at_fifty():
    assert minmax_normalize({1: 4.5, 2: 4.5}) == {1: 50.0, 2: 50.0}


def test_minmax_empty():
    assert minmax_normalize({}) == {}


def test_rating_scores_end_to_end_hand_computed():
    # C = mean(4.9, 4.5, 4.3) = 4.5666...; m = 300
    #   A: (1000*4.9 + 300*4.5666...)/1300 = (4900 + 1370)/1300 = 4.8231
    #   B: (  50*4.5 + 300*4.5666...)/350  = (225  + 1370)/350  = 4.5571
    #   C: (5000*4.3 + 300*4.5666...)/5300 = (21500+ 1370)/5300 = 4.3151
    scores = rating_scores({1: (4.9, 1000), 2: (4.5, 50), 3: (4.3, 5000)}, prior_weight=300)
    assert scores[1] == pytest.approx(100.0)  # highest adjusted
    assert scores[3] == pytest.approx(0.0)  # lowest adjusted
    # B sits between, at (4.5571 - 4.3151) / (4.8231 - 4.3151) * 100
    assert scores[2] == pytest.approx(47.64, abs=0.1)


def test_rating_scores_empty_universe():
    assert rating_scores({}, prior_weight=300) == {}


# ── Critic score ────────────────────────────────────────────────────────────


def test_critic_points_stack():
    assert critic_score(["nyt", "infatuation"], POINTS, cap=100) == 90.0


def test_critic_saturates_at_cap():
    assert critic_score(["michelin_star", "eater_38", "nyt"], POINTS, cap=100) == 100.0


def test_critic_single_flag():
    assert critic_score(["eater_38"], POINTS, cap=100) == 70.0


def test_critic_no_flags_is_a_real_zero():
    assert critic_score([], POINTS, cap=100) == 0.0


def test_critic_unknown_source_scores_nothing():
    assert critic_score(["zagat"], POINTS, cap=100) == 0.0
    assert critic_score(["nyt", "zagat"], POINTS, cap=100) == 50.0


def test_critic_duplicate_flags_count_once():
    assert critic_score(["nyt", "nyt"], POINTS, cap=100) == 50.0


# ── Scarcity ────────────────────────────────────────────────────────────────


def test_scarcity_share_hand_computed():
    # 18 of 24 observations found nothing -> 75.0
    assert scarcity_score(18, 24, min_observations=20) == pytest.approx(75.0)


def test_scarcity_below_minimum_is_none_not_zero():
    """A venue seen 3 times and booked out thrice is unmeasured, not 100% scarce."""
    assert scarcity_score(3, 3, min_observations=20) is None


def test_scarcity_exactly_at_minimum_scores():
    assert scarcity_score(10, 20, min_observations=20) == pytest.approx(50.0)


def test_scarcity_never_unavailable_is_zero():
    assert scarcity_score(0, 30, min_observations=20) == 0.0


def test_scarcity_always_unavailable_is_one_hundred():
    assert scarcity_score(30, 30, min_observations=20) == 100.0


def test_scarcity_zero_observations():
    assert scarcity_score(0, 0, min_observations=0) is None


# ── Composite and weight redistribution ─────────────────────────────────────


def test_composite_all_three_hand_computed():
    # 80*0.35 + 60*0.35 + 40*0.30 = 28 + 21 + 12 = 61, over total weight 1.0
    got = composite_score({RATING: 80.0, CRITIC: 60.0, SCARCITY: 40.0}, WEIGHTS)
    assert got == pytest.approx(61.0)


def test_composite_redistributes_when_scarcity_is_none():
    """The headline case: 0.30 is spread over the survivors, not counted as 0."""
    got = composite_score({RATING: 80.0, CRITIC: 60.0, SCARCITY: None}, WEIGHTS)
    # (80*0.35 + 60*0.35) / 0.70 = 49 / 0.70 = 70.0
    assert got == pytest.approx(70.0)
    # Treating the missing component as zero would have given 49.0 — the whole
    # point of redistribution is that these differ.
    assert got != pytest.approx(61.0 - 12.0)


def test_missing_scarcity_never_penalizes_relative_to_a_zero():
    """A venue with unknown scarcity must not rank below its own zero-scarcity self."""
    unknown = composite_score({RATING: 80.0, CRITIC: 60.0, SCARCITY: None}, WEIGHTS)
    measured_zero = composite_score({RATING: 80.0, CRITIC: 60.0, SCARCITY: 0.0}, WEIGHTS)
    assert unknown > measured_zero


def test_composite_redistribution_preserves_a_uniform_score():
    """If every present component is 70, the composite is 70 whatever is missing."""
    for components in (
        {RATING: 70.0, CRITIC: 70.0, SCARCITY: 70.0},
        {RATING: 70.0, CRITIC: 70.0, SCARCITY: None},
        {RATING: None, CRITIC: 70.0, SCARCITY: 70.0},
        {RATING: 70.0, CRITIC: None, SCARCITY: None},
    ):
        assert composite_score(components, WEIGHTS) == pytest.approx(70.0)


def test_composite_with_only_scarcity():
    assert composite_score({RATING: None, CRITIC: None, SCARCITY: 42.0}, WEIGHTS) == 42.0


def test_composite_all_none():
    assert composite_score({RATING: None, CRITIC: None, SCARCITY: None}, WEIGHTS) is None


def test_composite_ignores_zero_weighted_components():
    """Setting WEIGHT_RATING=0 must actually remove the signal, not divide by it."""
    weights = {RATING: 0.0, CRITIC: 0.5, SCARCITY: 0.5}
    got = composite_score({RATING: 100.0, CRITIC: 60.0, SCARCITY: 40.0}, weights)
    assert got == pytest.approx(50.0)


def test_composite_all_weights_zero_is_none():
    weights = {RATING: 0.0, CRITIC: 0.0, SCARCITY: 0.0}
    assert composite_score({RATING: 80.0, CRITIC: 60.0, SCARCITY: 40.0}, weights) is None


def test_composite_unnormalized_weights_still_work():
    """Weights need not sum to 1 — they are renormalized either way."""
    weights = {RATING: 7.0, CRITIC: 7.0, SCARCITY: 6.0}
    assert composite_score(
        {RATING: 80.0, CRITIC: 60.0, SCARCITY: 40.0}, weights
    ) == pytest.approx(61.0)


# ── Scarcity from real scan history ─────────────────────────────────────────


@pytest.fixture()
def scored_db(tmp_path):
    path = tmp_path / "score.db"
    db.init_db(path)
    conn = db.connect(path)
    with conn:
        conn.execute("INSERT INTO venues (venue_id, name, neighborhood) VALUES (1, 'Scarce', 'WV')")
        conn.execute("INSERT INTO venues (venue_id, name, neighborhood) VALUES (2, 'Open', 'LES')")
    yield conn
    conn.close()


def _add_scan(conn, venue_id, scan_ts, target_date, *, prime_slot: bool):
    cur = conn.execute(
        "INSERT INTO scans (venue_id, scan_ts, target_date, party_size) VALUES (?, ?, ?, 2)",
        (venue_id, scan_ts, target_date),
    )
    if prime_slot:
        conn.execute(
            "INSERT INTO slots (scan_id, venue_id, target_date, slot_time, is_prime) "
            "VALUES (?, ?, ?, '19:00', 1)",
            (cur.lastrowid, venue_id, target_date),
        )


def test_scarcity_observations_counts_only_prime_dates_in_horizon(scored_db):
    settings = Settings(_env_file=None)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    scan_ts = now.isoformat(timespec="seconds")

    # 2026-09-05 is a Saturday, 21 days out — prime date, inside the 14-28 window.
    _add_scan(scored_db, 1, scan_ts, "2026-09-05", prime_slot=False)
    # 2026-09-07 is a Monday — not a prime date, must be ignored.
    _add_scan(scored_db, 1, scan_ts, "2026-09-07", prime_slot=False)
    # 2026-08-20 is a Thursday but only 5 days out — outside the horizon.
    _add_scan(scored_db, 1, scan_ts, "2026-08-20", prime_slot=False)
    # 2026-10-30 is a Friday but 76 days out — outside the horizon.
    _add_scan(scored_db, 1, scan_ts, "2026-10-30", prime_slot=False)
    scored_db.commit()

    observations = scarcity_observations(scored_db, settings, now=now)
    assert observations[1] == (1, 1)  # exactly the one qualifying observation


def test_scarcity_observations_ignores_stale_scans(scored_db):
    settings = Settings(_env_file=None)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    old = (now - timedelta(days=45)).isoformat(timespec="seconds")
    _add_scan(scored_db, 1, old, "2026-07-18", prime_slot=False)
    scored_db.commit()
    assert scarcity_observations(scored_db, settings, now=now) == {}


def test_scarcity_observations_availability_marks_observation_as_available(scored_db):
    settings = Settings(_env_file=None)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    scan_ts = now.isoformat(timespec="seconds")
    _add_scan(scored_db, 2, scan_ts, "2026-09-05", prime_slot=True)
    _add_scan(scored_db, 2, scan_ts, "2026-09-04", prime_slot=False)
    scored_db.commit()
    assert scarcity_observations(scored_db, settings, now=now)[2] == (1, 2)


def test_non_prime_slots_do_not_count_as_prime_availability(scored_db):
    """A venue with only a 22:00 table was still unavailable in the prime window."""
    settings = Settings(_env_file=None)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    cur = scored_db.execute(
        "INSERT INTO scans (venue_id, scan_ts, target_date, party_size) "
        "VALUES (1, ?, '2026-09-05', 2)",
        (now.isoformat(timespec="seconds"),),
    )
    scored_db.execute(
        "INSERT INTO slots (scan_id, venue_id, target_date, slot_time, is_prime) "
        "VALUES (?, 1, '2026-09-05', '22:00', 0)",
        (cur.lastrowid,),
    )
    scored_db.commit()
    assert scarcity_observations(scored_db, settings, now=now)[1] == (1, 1)


def test_score_universe_wires_everything_together(scored_db):
    settings = Settings(_env_file=None, SCARCITY_MIN_OBSERVATIONS=1)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    with scored_db:
        scored_db.execute(
            "INSERT INTO google_meta (venue_id, rating, user_rating_count, fetched_at) "
            "VALUES (1, 4.8, 2000, '2026-08-01T00:00:00+00:00')"
        )
        scored_db.execute(
            "INSERT INTO google_meta (venue_id, rating, user_rating_count, fetched_at) "
            "VALUES (2, 4.4, 2000, '2026-08-01T00:00:00+00:00')"
        )
        scored_db.execute(
            "INSERT INTO critic_flags (venue_id, source, tier) VALUES (1, 'michelin_star', '1')"
        )
    _add_scan(scored_db, 1, now.isoformat(timespec="seconds"), "2026-09-05", prime_slot=False)
    _add_scan(scored_db, 2, now.isoformat(timespec="seconds"), "2026-09-05", prime_slot=True)
    scored_db.commit()

    scores = score_universe(scored_db, settings, now=now)

    scarce, open_venue = scores[1], scores[2]
    assert scarce.rating_score == 100.0 and open_venue.rating_score == 0.0
    assert scarce.critic_score == 100.0 and open_venue.critic_score == 0.0
    assert scarce.scarcity_score == 100.0 and open_venue.scarcity_score == 0.0
    assert scarce.composite == pytest.approx(100.0)
    assert open_venue.composite == pytest.approx(0.0)
    assert scarce.critic_sources == ("michelin_star",)


def test_score_universe_venue_with_no_google_data_redistributes(scored_db):
    """A venue with no rating must not be pushed to the bottom by a phantom zero."""
    settings = Settings(_env_file=None)
    now = datetime(2026, 8, 15, 12, 0, tzinfo=timezone.utc)
    with scored_db:
        scored_db.execute(
            "INSERT INTO critic_flags (venue_id, source, tier) VALUES (1, 'michelin_star', '1')"
        )

    scores = score_universe(scored_db, settings, now=now)
    assert scores[1].rating_score is None
    assert scores[1].scarcity_score is None
    # Only the critic component survives, so the composite is exactly it.
    assert scores[1].composite == pytest.approx(100.0)
    assert scores[2].composite == pytest.approx(0.0)
