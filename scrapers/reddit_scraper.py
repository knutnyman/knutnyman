"""Fetch Dexcom-related posts from Reddit via PRAW."""

import logging
import os
from datetime import datetime, timezone

import praw

log = logging.getLogger(__name__)

_DEFAULT_SUBREDDITS = ["dexcom", "diabetes", "diabetes_t1", "diabetes_t2", "CGM"]
_DEXCOM_KEYWORDS = {"dexcom", "g6", "g7", "cgm", "continuous glucose", "g5", "g4"}


def fetch_posts(hours_back: int = 24) -> list[dict]:
    client_id = os.getenv("REDDIT_CLIENT_ID")
    client_secret = os.getenv("REDDIT_CLIENT_SECRET")
    if not all([client_id, client_secret]):
        log.warning("REDDIT_CLIENT_ID/REDDIT_CLIENT_SECRET not set — skipping Reddit")
        return []

    reddit = praw.Reddit(
        client_id=client_id,
        client_secret=client_secret,
        user_agent=os.getenv("REDDIT_USER_AGENT", "dexcom-sentiment-tracker/1.0"),
    )

    subreddits = [
        s.strip()
        for s in os.getenv("DEXCOM_SUBREDDITS", ",".join(_DEFAULT_SUBREDDITS)).split(",")
        if s.strip()
    ]
    cutoff = datetime.now(timezone.utc).timestamp() - hours_back * 3600
    posts: list[dict] = []

    for sub_name in subreddits:
        try:
            subreddit = reddit.subreddit(sub_name)
            for submission in subreddit.new(limit=200):
                if submission.created_utc < cutoff:
                    continue
                text = f"{submission.title} {submission.selftext}"
                # for broad diabetes subs, only include Dexcom-relevant posts
                if sub_name.lower() not in ("dexcom", "cgm"):
                    if not any(kw in text.lower() for kw in _DEXCOM_KEYWORDS):
                        continue
                posts.append({
                    "id": f"reddit_{submission.id}",
                    "source": f"reddit/r/{sub_name}",
                    "title": submission.title,
                    "body": submission.selftext[:500],
                    "text": text[:1200],
                    "url": f"https://reddit.com{submission.permalink}",
                    "score": submission.score,
                    "num_comments": submission.num_comments,
                    "created_utc": submission.created_utc,
                })
            log.info("r/%s: %d relevant posts", sub_name, sum(1 for p in posts if f"r/{sub_name}" in p["source"]))
        except Exception as exc:
            log.warning("r/%s failed: %s", sub_name, exc)

    log.info("Reddit total: %d posts", len(posts))
    return posts
