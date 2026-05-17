"""Build the HTML + plain-text email digest from analyzed posts."""

import re
from collections import Counter
from datetime import datetime


def build_subject(posts: list[dict]) -> str:
    if not posts:
        return "Dexcom Sentiment Digest — No data collected"
    total = len(posts)
    positive = sum(1 for p in posts if p["sentiment_label"] == "positive")
    avg = sum(p["sentiment_compound"] for p in posts) / total
    mood = "Positive" if avg >= 0.05 else "Negative" if avg <= -0.05 else "Neutral"
    pct = int(positive / total * 100)
    return f"Dexcom Sentiment Digest — {mood} ({pct}% positive across {total} posts)"


def build_html(posts: list[dict]) -> str:
    if not posts:
        return "<p>No posts collected. Check credentials and try again.</p>"

    total = len(posts)
    positive = [p for p in posts if p["sentiment_label"] == "positive"]
    negative = [p for p in posts if p["sentiment_label"] == "negative"]
    neutral = [p for p in posts if p["sentiment_label"] == "neutral"]
    avg = sum(p["sentiment_compound"] for p in posts) / total

    mood_color = "#43a047" if avg >= 0.05 else "#e53935" if avg <= -0.05 else "#fb8c00"
    date_str = datetime.now().strftime("%B %d, %Y")
    source_counts = Counter(p["source"] for p in posts)
    keywords = _top_keywords(posts)
    top_pos = sorted(posts, key=lambda p: p["sentiment_compound"], reverse=True)[:3]
    top_neg = sorted(posts, key=lambda p: p["sentiment_compound"])[:3]

    sources_html = "".join(
        f"<li><strong>{s}</strong>: {c} posts</li>" for s, c in source_counts.most_common()
    )
    keywords_html = "".join(
        f"<span style='display:inline-block;background:#e3f2fd;color:#1565c0;"
        f"padding:3px 10px;border-radius:12px;margin:3px 2px;font-size:12px;'>"
        f"{w} <b>({c})</b></span>"
        for w, c in keywords
    )

    def _card(post: dict, border: str) -> str:
        link = f" &bull; <a href='{post['url']}' style='color:#1565c0;'>view</a>" if post["url"] else ""
        return (
            f"<div style='border-left:4px solid {border};padding:8px 14px;"
            f"margin:8px 0;background:#fafafa;border-radius:0 6px 6px 0;'>"
            f"<div style='font-size:14px;font-weight:600;color:#212121;'>{post['title'][:90]}</div>"
            f"<div style='font-size:12px;color:#757575;margin-top:4px;'>"
            f"{post['source']} &bull; score: <b>{post['sentiment_compound']:+.2f}</b>{link}"
            f"</div></div>"
        )

    pos_cards = "".join(_card(p, "#43a047") for p in top_pos)
    neg_cards = "".join(_card(p, "#e53935") for p in top_neg)

    return f"""<!DOCTYPE html>
<html>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
             max-width:680px;margin:0 auto;padding:20px;color:#212121;background:#fff;">

  <h2 style="margin:0 0 4px;color:#1a237e;">Dexcom Community Sentiment</h2>
  <p style="margin:0 0 20px;color:#757575;font-size:13px;">{date_str}</p>

  <!-- Stats bar -->
  <div style="display:table;width:100%;border-spacing:8px;margin-bottom:20px;">
    <div style="display:table-row;">
      <div style="display:table-cell;text-align:center;background:#f5f5f5;padding:16px;
                  border-radius:8px;width:20%;">
        <div style="font-size:28px;font-weight:700;color:{mood_color};">{avg:+.2f}</div>
        <div style="font-size:11px;color:#9e9e9e;margin-top:2px;">avg score</div>
      </div>
      <div style="display:table-cell;text-align:center;background:#f5f5f5;padding:16px;
                  border-radius:8px;width:20%;">
        <div style="font-size:28px;font-weight:700;">{total}</div>
        <div style="font-size:11px;color:#9e9e9e;margin-top:2px;">posts</div>
      </div>
      <div style="display:table-cell;text-align:center;background:#e8f5e9;padding:16px;
                  border-radius:8px;width:20%;">
        <div style="font-size:28px;font-weight:700;color:#43a047;">{len(positive)}</div>
        <div style="font-size:11px;color:#9e9e9e;margin-top:2px;">positive</div>
      </div>
      <div style="display:table-cell;text-align:center;background:#fff8e1;padding:16px;
                  border-radius:8px;width:20%;">
        <div style="font-size:28px;font-weight:700;color:#fb8c00;">{len(neutral)}</div>
        <div style="font-size:11px;color:#9e9e9e;margin-top:2px;">neutral</div>
      </div>
      <div style="display:table-cell;text-align:center;background:#ffebee;padding:16px;
                  border-radius:8px;width:20%;">
        <div style="font-size:28px;font-weight:700;color:#e53935;">{len(negative)}</div>
        <div style="font-size:11px;color:#9e9e9e;margin-top:2px;">negative</div>
      </div>
    </div>
  </div>

  <h3 style="color:#424242;border-bottom:1px solid #e0e0e0;padding-bottom:6px;">Sources</h3>
  <ul style="font-size:13px;color:#424242;">{sources_html}</ul>

  <h3 style="color:#424242;border-bottom:1px solid #e0e0e0;padding-bottom:6px;">
    Top Keywords
  </h3>
  <div style="margin:8px 0 20px;">{keywords_html}</div>

  <h3 style="color:#43a047;border-bottom:1px solid #e0e0e0;padding-bottom:6px;">
    Most Positive Posts
  </h3>
  {pos_cards}

  <h3 style="color:#e53935;border-bottom:1px solid #e0e0e0;padding-bottom:6px;margin-top:20px;">
    Most Negative Posts
  </h3>
  {neg_cards}

  <p style="font-size:11px;color:#bdbdbd;margin-top:32px;border-top:1px solid #eeeeee;
            padding-top:12px;">
    Dexcom Sentiment Tracker &bull; {date_str} &bull;
    Data from Reddit &amp; Facebook community groups
  </p>
</body>
</html>"""


def build_plain(posts: list[dict]) -> str:
    if not posts:
        return "No posts collected. Check credentials and try again."

    total = len(posts)
    positive = sum(1 for p in posts if p["sentiment_label"] == "positive")
    negative = sum(1 for p in posts if p["sentiment_label"] == "negative")
    avg = sum(p["sentiment_compound"] for p in posts) / total
    date_str = datetime.now().strftime("%B %d, %Y")

    lines = [
        f"Dexcom Community Sentiment — {date_str}",
        "=" * 52,
        f"Posts analyzed : {total}",
        f"Avg score      : {avg:+.2f}",
        f"Positive       : {positive}",
        f"Neutral        : {total - positive - negative}",
        f"Negative       : {negative}",
        "",
        "--- MOST POSITIVE ---",
    ]
    for p in sorted(posts, key=lambda x: x["sentiment_compound"], reverse=True)[:3]:
        lines += [f"  [{p['sentiment_compound']:+.2f}] {p['title'][:75]}", f"         {p['url']}"]
    lines += ["", "--- MOST NEGATIVE ---"]
    for p in sorted(posts, key=lambda x: x["sentiment_compound"])[:3]:
        lines += [f"  [{p['sentiment_compound']:+.2f}] {p['title'][:75]}", f"         {p['url']}"]
    return "\n".join(lines)


_STOPWORDS = {
    "the", "a", "an", "is", "it", "i", "my", "me", "of", "to", "in",
    "and", "that", "was", "for", "on", "are", "with", "as", "at", "this",
    "be", "have", "from", "or", "had", "has", "its", "but", "not", "do",
    "im", "ive", "just", "get", "got", "can", "use", "using", "one", "so",
    "if", "by", "no", "up", "been", "they", "we", "he", "she", "you",
}


def _top_keywords(posts: list[dict], n: int = 12) -> list[tuple[str, int]]:
    counts: Counter = Counter()
    for post in posts:
        words = re.findall(r"\b[a-z]{3,}\b", post["text"].lower())
        counts.update(w for w in words if w not in _STOPWORDS)
    return counts.most_common(n)
