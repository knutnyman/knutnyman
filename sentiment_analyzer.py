"""VADER-based sentiment analysis for social media posts."""

from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

_vader = SentimentIntensityAnalyzer()


def analyze(posts: list[dict]) -> list[dict]:
    for post in posts:
        scores = _vader.polarity_scores(post["text"])
        compound = scores["compound"]
        post["sentiment_compound"] = compound
        post["sentiment_pos"] = scores["pos"]
        post["sentiment_neg"] = scores["neg"]
        post["sentiment_neu"] = scores["neu"]
        if compound >= 0.05:
            post["sentiment_label"] = "positive"
        elif compound <= -0.05:
            post["sentiment_label"] = "negative"
        else:
            post["sentiment_label"] = "neutral"
    return posts
