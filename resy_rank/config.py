"""Settings for resy-rank.

Every tunable — rate limits, scoring weights, the Bayesian prior, critic
points, the scarcity thresholds — lives here and is overridable from `.env`,
so tuning the ranking never means editing code.
"""

from __future__ import annotations

from datetime import time
from functools import lru_cache
from pathlib import Path

from pydantic import Field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

PROJECT_ROOT = Path(__file__).resolve().parent.parent

_WEEKDAYS = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}


class Settings(BaseSettings):
    # extra="ignore" so unrelated vars in a shared .env don't fail startup.
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )

    # ── Credentials ─────────────────────────────────────────────────────────
    # Pulled from a logged-in browser session. Optional at import time so the
    # offline commands (init) work on a fresh checkout with no .env.
    resy_api_key: str | None = None
    resy_auth_token: str | None = None
    google_places_api_key: str | None = None

    # ── Paths ───────────────────────────────────────────────────────────────
    db_path: Path = PROJECT_ROOT / "data" / "resy_rank.db"
    venues_csv: Path = PROJECT_ROOT / "data" / "venues.csv"
    critics_csv: Path = PROJECT_ROOT / "data" / "critics.csv"

    # ── Politeness / rate limiting ──────────────────────────────────────────
    # These are hard requirements, not suggestions. This tool is a slow
    # background scanner; it must never behave like a drop-time sniper.
    max_concurrency: int = Field(default=3, ge=1, le=3)
    request_delay_seconds: float = Field(default=1.5, ge=0.0)
    daily_request_cap: int = Field(default=750, ge=1)
    max_retries: int = Field(default=4, ge=0)
    backoff_base_seconds: float = Field(default=2.0, gt=1.0)
    request_timeout_seconds: float = Field(default=20.0, gt=0)

    # ── Resy request defaults ───────────────────────────────────────────────
    # /4/find wants a geo anchor even when filtering to a single venue_id.
    default_lat: float = 40.7128
    default_long: float = -74.0060
    default_party_size: int = Field(default=2, ge=1)

    # ── Geo sweep ───────────────────────────────────────────────────────────
    # /4/find without a venue_id returns every venue with availability near a
    # point — the same query resy.com makes when you browse. One sweep covers
    # a city in a few dozen requests instead of one-per-venue, which is both
    # ~40x cheaper and far more polite than enumerating venue ids.
    geo_anchors: list[tuple[float, float]] = Field(
        default_factory=lambda: [
            (40.7075, -74.0113),  # Financial District / Battery
            (40.7185, -73.9950),  # Chinatown / Lower East Side
            (40.7265, -73.9830),  # East Village
            (40.7350, -74.0030),  # West Village
            (40.7440, -73.9900),  # Chelsea / Flatiron
            (40.7549, -73.9840),  # Midtown
            (40.7736, -73.9566),  # Upper East Side
            (40.7870, -73.9754),  # Upper West Side
            (40.8116, -73.9465),  # Harlem
        ]
    )
    geo_per_page: int = Field(default=50, ge=1, le=100)
    geo_max_pages: int = Field(default=10, ge=1)

    # ── Google Places ───────────────────────────────────────────────────────
    google_cache_days: int = Field(default=180, ge=1)
    # Enrichment is capped per calendar month so it stays inside Google's free
    # tier by default (1,000 Enterprise-SKU calls/month since March 2025).
    # Raise this only if you actually want to be billed.
    monthly_enrichment_cap: int = Field(default=1000, ge=0)

    # ── Critic matching ─────────────────────────────────────────────────────
    critic_match_threshold: int = Field(default=88, ge=0, le=100)
    critic_points: dict[str, float] = Field(
        default_factory=lambda: {
            "michelin_star": 100.0,
            "michelin_bib": 70.0,
            "eater_38": 70.0,
            "nyt": 50.0,
            "infatuation": 40.0,
        }
    )
    critic_score_cap: float = Field(default=100.0, gt=0)

    # ── Scoring ─────────────────────────────────────────────────────────────
    # Bayesian prior weight `m`: how many "average" reviews a venue is charged
    # before its own rating counts fully. The prior mean `C` is computed from
    # the venue universe at scoring time, never hardcoded.
    bayesian_prior_weight: float = Field(default=300.0, gt=0)

    weight_rating: float = Field(default=0.35, ge=0)
    weight_critic: float = Field(default=0.35, ge=0)
    weight_scarcity: float = Field(default=0.30, ge=0)

    # ── Scarcity index ──────────────────────────────────────────────────────
    # Prime = Thu-Sat, 18:30-20:30. Scarcity only counts observations made
    # 14-28 days out, where a full book-out is signal rather than noise.
    prime_days: list[int] = Field(default_factory=lambda: [3, 4, 5])
    prime_start: time = time(18, 30)
    prime_end: time = time(20, 30)
    scarcity_lookback_days: int = Field(default=30, ge=1)
    scarcity_horizon_min_days: int = Field(default=14, ge=0)
    scarcity_horizon_max_days: int = Field(default=28, ge=1)
    scarcity_min_observations: int = Field(default=20, ge=1)

    @field_validator("geo_anchors", mode="before")
    @classmethod
    def _parse_anchors(cls, v: object) -> object:
        """Accept "lat,long;lat,long" from .env as well as a list of pairs."""
        if isinstance(v, str):
            anchors = []
            for chunk in v.split(";"):
                chunk = chunk.strip()
                if not chunk:
                    continue
                lat, _, lng = chunk.partition(",")
                anchors.append((float(lat), float(lng)))
            return anchors
        return v

    @field_validator("prime_days", mode="before")
    @classmethod
    def _parse_prime_days(cls, v: object) -> object:
        """Accept either [3,4,5] or a friendlier "thu,fri,sat" from .env."""
        if isinstance(v, str):
            parts = [p.strip().lower() for p in v.split(",") if p.strip()]
            return [int(p) if p.isdigit() else _WEEKDAYS[p[:3]] for p in parts]
        return v

    @field_validator("db_path", "venues_csv", "critics_csv")
    @classmethod
    def _absolutize(cls, v: Path) -> Path:
        return v if v.is_absolute() else (PROJECT_ROOT / v).resolve()

    @model_validator(mode="after")
    def _check_consistency(self) -> "Settings":
        total = self.weight_rating + self.weight_critic + self.weight_scarcity
        if total <= 0:
            raise ValueError("at least one scoring weight must be greater than zero")
        if self.prime_start >= self.prime_end:
            raise ValueError("prime_start must be earlier than prime_end")
        if self.scarcity_horizon_min_days >= self.scarcity_horizon_max_days:
            raise ValueError(
                "scarcity_horizon_min_days must be less than scarcity_horizon_max_days"
            )
        return self

    @property
    def weight_total(self) -> float:
        return self.weight_rating + self.weight_critic + self.weight_scarcity

    def require_resy_credentials(self) -> tuple[str, str]:
        """Return (api_key, auth_token), or raise with a fixable message."""
        missing = [
            name
            for name, value in (
                ("RESY_API_KEY", self.resy_api_key),
                ("RESY_AUTH_TOKEN", self.resy_auth_token),
            )
            if not value
        ]
        if missing:
            raise RuntimeError(
                f"Missing {', '.join(missing)} in .env — copy .env.example and paste "
                "the values from your logged-in Resy browser session."
            )
        return self.resy_api_key, self.resy_auth_token  # type: ignore[return-value]

    def require_google_key(self) -> str:
        if not self.google_places_api_key:
            raise RuntimeError("Missing GOOGLE_PLACES_API_KEY in .env")
        return self.google_places_api_key


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
