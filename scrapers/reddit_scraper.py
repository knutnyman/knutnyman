"""Fetch Dexcom-related posts from Reddit using the public JSON feed.

No API credentials required — uses Reddit's unauthenticated .json endpoints.
Rate limit: ~1 req/sec (enforced by a small delay between requests).
"""

import logging
import os
import time
from datetime import datetime, timezone

import requests

log = logging.getLogger(__name__)

_DEFAULT_SUBREDDITS = ["dexcom", "diabetes", "diabetes_t1", "diabetes_t2", "CGM"]
_DEXCOM_KEYWORDS = {"dexcom", "g6", "g7", "cgm", "continuous glucose", "g5", "g4"}
_HEADERS = {"User-Agent": "dexcom-sentiment-tracker/1.0 (personal research)"}


def _fetch_subreddit(sub: str, limit: int = 100) -> list[dict]:
    url = f"https://www.reddit.com/r/{sub}/new.json"
    try:
        resp = requests.get(url, headers=_HEADERS, params={"limit": limit}, timeout=10)
        resp.raise_for_status()
        return resp.json()["data"]["children"]
    except Exception as exc:
        log.warning("r/%s fetch failed: %s", sub, exc)
        return []


def fetch_posts(hours_back: int = 24) -> list[dict]:
    subreddits = [
        s.strip()
        for s in os.getenv("DEXCOM_SUBREDDITS", ",".join(_DEFAULT_SUBREDDITS)).split(",")
        if s.strip()
    ]
    cutoff = datetime.now(timezone.utc).timestamp() - hours_back * 3600
    posts: list[dict] = []

    for sub in subreddits:
        children = _fetch_subreddit(sub)
        count = 0
        for child in children:
            d = child["data"]
            if d.get("created_utc", 0) < cutoff:
                continue
            text = f"{d.get('title', '')} {d.get('selftext', '')}"
            # filter broader subs to Dexcom-relevant posts only
            if sub.lower() not in ("dexcom", "cgm"):
                if not any(kw in text.lower() for kw in _DEXCOM_KEYWORDS):
                    continue
            posts.append({
                "id": f"reddit_{d['id']}",
                "source": f"reddit/r/{sub}",
                "title": d.get("title", ""),
                "body": d.get("selftext", "")[:500],
                "text": text[:1200],
                "url": f"https://reddit.com{d.get('permalink', '')}",
                "score": d.get("score", 0),
                "num_comments": d.get("num_comments", 0),
                "created_utc": d.get("created_utc"),
            })
            count += 1
        log.info("r/%s: %d relevant posts", sub, count)
        time.sleep(1)  # respect Reddit's unauthenticated rate limit

    log.info("Reddit total: %d posts", len(posts))
    return posts
