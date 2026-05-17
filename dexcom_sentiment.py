#!/usr/bin/env python3
"""Dexcom community sentiment tracker.

Pulls posts from Reddit (via PRAW) and Facebook groups (via Selenium),
scores them with VADER, and sends an HTML email digest.

Usage:
    python dexcom_sentiment.py              # last 24 hours
    SENTIMENT_HOURS_BACK=48 python dexcom_sentiment.py

Required env vars (see .env.example):
    Reddit:   REDDIT_CLIENT_ID, REDDIT_CLIENT_SECRET
    Facebook: FACEBOOK_EMAIL, FACEBOOK_PASSWORD, FACEBOOK_GROUP_URLS
    Email:    NOTIFY_EMAIL_TO, NOTIFY_SMTP_HOST, NOTIFY_SMTP_USER, NOTIFY_SMTP_PASS
"""

import logging
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv

import digest_builder
import sentiment_analyzer
from scrapers.facebook_scraper import fetch_posts as fb_fetch
from scrapers.reddit_scraper import fetch_posts as reddit_fetch

load_dotenv()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(name)s — %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def _send_digest(subject: str, html: str, plain: str) -> None:
    to = os.getenv("NOTIFY_EMAIL_TO")
    host = os.getenv("NOTIFY_SMTP_HOST")
    user = os.getenv("NOTIFY_SMTP_USER")
    password = os.getenv("NOTIFY_SMTP_PASS")
    port = int(os.getenv("NOTIFY_SMTP_PORT", "587"))

    if not all([to, host, user, password]):
        log.info("Email not configured — printing digest to stdout instead")
        print(plain)
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = to
    msg.attach(MIMEText(plain, "plain"))
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP(host, port) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(user, password)
            smtp.sendmail(user, to, msg.as_string())
        log.info("Digest emailed to %s", to)
    except Exception as exc:
        log.warning("Email failed (%s) — printing digest to stdout", exc)
        print(plain)


def main() -> None:
    hours_back = int(os.getenv("SENTIMENT_HOURS_BACK", "24"))
    log.info("Collecting posts from the last %d hours...", hours_back)

    posts: list[dict] = []
    posts.extend(reddit_fetch(hours_back=hours_back))
    posts.extend(fb_fetch())

    if not posts:
        log.warning("No posts collected — verify credentials in .env")
        return

    log.info("Running VADER sentiment analysis on %d posts...", len(posts))
    posts = sentiment_analyzer.analyze(posts)

    positive = sum(1 for p in posts if p["sentiment_label"] == "positive")
    negative = sum(1 for p in posts if p["sentiment_label"] == "negative")
    log.info("Results: %d positive / %d neutral / %d negative", positive,
             len(posts) - positive - negative, negative)

    subject = digest_builder.build_subject(posts)
    html = digest_builder.build_html(posts)
    plain = digest_builder.build_plain(posts)
    _send_digest(subject, html, plain)


if __name__ == "__main__":
    main()
