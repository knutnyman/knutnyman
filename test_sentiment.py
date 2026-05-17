#!/usr/bin/env python3
"""Smoke test — runs the full analyze → digest pipeline on mock posts.

No API credentials needed.  Prints the plain-text digest and exits.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import sentiment_analyzer
import digest_builder

MOCK_POSTS = [
    {
        "id": "t1", "source": "reddit/r/dexcom",
        "title": "G7 is a game changer — best CGM I've ever used",
        "body": "Accuracy is incredible, warmup is only 30 min, loving it.",
        "text": "G7 is a game changer — best CGM I've ever used. Accuracy is incredible, warmup is only 30 min, loving it.",
        "url": "https://reddit.com/r/dexcom/comments/abc", "score": 142, "num_comments": 34, "created_utc": None,
    },
    {
        "id": "t2", "source": "reddit/r/dexcom",
        "title": "Dexcom sensor fell off after one day — so frustrated",
        "body": "Third time this month. The adhesive is terrible, skin prep didn't help.",
        "text": "Dexcom sensor fell off after one day — so frustrated. Third time this month. The adhesive is terrible, skin prep didn't help.",
        "url": "https://reddit.com/r/dexcom/comments/def", "score": 89, "num_comments": 21, "created_utc": None,
    },
    {
        "id": "t3", "source": "reddit/r/diabetes_t1",
        "title": "Anyone else getting ??? errors constantly on their G6?",
        "body": "Two sensors in a row both died after day 2 with ??? errors. Called support, they replaced them but still.",
        "text": "Anyone else getting ??? errors constantly on their G6? Two sensors in a row both died after day 2. Called support, they replaced them.",
        "url": "https://reddit.com/r/diabetes_t1/comments/ghi", "score": 56, "num_comments": 47, "created_utc": None,
    },
    {
        "id": "t4", "source": "reddit/r/CGM",
        "title": "Dexcom G7 vs Libre 3 — my honest 6-month comparison",
        "body": "Both are solid options. G7 has better alerts, Libre 3 costs less.",
        "text": "Dexcom G7 vs Libre 3 — my honest 6-month comparison. Both are solid options. G7 has better alerts, Libre 3 costs less.",
        "url": "https://reddit.com/r/CGM/comments/jkl", "score": 201, "num_comments": 88, "created_utc": None,
    },
    {
        "id": "t5", "source": "facebook/dexcomusers",
        "title": "Love my Dexcom, changed my life managing Type 1!",
        "body": "Never felt so in control of my numbers. Highly recommend to anyone on the fence.",
        "text": "Love my Dexcom, changed my life managing Type 1! Never felt so in control. Highly recommend to anyone on the fence.",
        "url": "https://www.facebook.com/groups/dexcomusers", "score": 0, "num_comments": 0, "created_utc": None,
    },
    {
        "id": "t6", "source": "facebook/dexcomusers",
        "title": "Customer service is a nightmare — been on hold for 2 hours",
        "body": "Sensor failure and they keep transferring me. Absolutely terrible experience.",
        "text": "Customer service is a nightmare — been on hold for 2 hours. Sensor failure and they keep transferring me. Absolutely terrible experience.",
        "url": "https://www.facebook.com/groups/dexcomusers", "score": 0, "num_comments": 0, "created_utc": None,
    },
    {
        "id": "t7", "source": "reddit/r/dexcom",
        "title": "G7 direct to Apple Watch — works okay I guess",
        "body": "It's fine. Connection drops occasionally but mostly reliable.",
        "text": "G7 direct to Apple Watch — works okay I guess. Connection drops occasionally but mostly reliable.",
        "url": "https://reddit.com/r/dexcom/comments/mno", "score": 33, "num_comments": 12, "created_utc": None,
    },
    {
        "id": "t8", "source": "reddit/r/diabetes_t2",
        "title": "New to Dexcom G7 — tips for first-time users?",
        "body": "Just got prescribed it, haven't started yet. Any advice appreciated!",
        "text": "New to Dexcom G7 — tips for first-time users? Just got prescribed it, haven't started yet. Any advice appreciated!",
        "url": "https://reddit.com/r/diabetes_t2/comments/pqr", "score": 18, "num_comments": 29, "created_utc": None,
    },
]


def main():
    print("Running sentiment analysis on mock Dexcom posts...\n")

    posts = sentiment_analyzer.analyze(MOCK_POSTS)

    print("Per-post scores:")
    print(f"  {'Label':<10} {'Score':>6}  Title")
    print("  " + "-" * 65)
    for p in sorted(posts, key=lambda x: x["sentiment_compound"], reverse=True):
        print(f"  {p['sentiment_label']:<10} {p['sentiment_compound']:>+6.3f}  {p['title'][:55]}")

    print()
    plain = digest_builder.build_plain(posts)
    print(plain)

    html = digest_builder.build_html(posts)
    out = "/tmp/dexcom_digest_preview.html"
    with open(out, "w") as f:
        f.write(html)
    print(f"\nHTML digest saved to: {out}")


if __name__ == "__main__":
    main()
