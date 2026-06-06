"""AI music-production backend.

Serves the static studio UI and exposes a single /api/generate endpoint that
asks Claude to compose a short, grid-based arrangement (chords + melody + drums)
which the browser renders to sound and lets you edit.

The API key never reaches the browser — all Claude calls happen here.
"""

import json
import os

import anthropic
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory

load_dotenv()

MODEL = os.environ.get("STUDIO_MODEL", "claude-opus-4-8")
STEPS_PER_BAR = 16  # 16th-note grid

app = Flask(__name__, static_folder="static", static_url_path="")

# JSON schema the model must fill in. Everything is quantized to a 16th-note
# grid so the front-end sequencer can play and edit it without re-interpreting
# free timing. `start`/`length` are in grid steps; pitches are MIDI note numbers.
ARRANGEMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "bpm": {"type": "integer"},
        "key": {"type": "string"},
        "chords": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "start": {"type": "integer"},
                    "length": {"type": "integer"},
                    "notes": {"type": "array", "items": {"type": "integer"}},
                },
                "required": ["name", "start", "length", "notes"],
                "additionalProperties": False,
            },
        },
        "melody": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "pitch": {"type": "integer"},
                    "start": {"type": "integer"},
                    "length": {"type": "integer"},
                },
                "required": ["pitch", "start", "length"],
                "additionalProperties": False,
            },
        },
        "drums": {
            "type": "object",
            "properties": {
                "kick": {"type": "array", "items": {"type": "integer"}},
                "snare": {"type": "array", "items": {"type": "integer"}},
                "hat": {"type": "array", "items": {"type": "integer"}},
            },
            "required": ["kick", "snare", "hat"],
            "additionalProperties": False,
        },
    },
    "required": ["title", "bpm", "key", "chords", "melody", "drums"],
    "additionalProperties": False,
}

SYSTEM_PROMPT = """You are an expert music producer and composer. You write \
short, musically coherent loops as structured data that a step sequencer plays \
back.

Rules:
- Everything is quantized to a 16th-note grid. One bar = 16 steps.
- `start` and `length` are integer grid steps. Pitches are MIDI note numbers \
(middle C = 60).
- Chords: voice 3-4 notes each, in a sensible register (roughly MIDI 48-67). \
Lay out a progression that fills the requested number of bars.
- Melody: a singable lead line, mostly in the requested key, register roughly \
MIDI 67-84. Use rests (gaps in `start`) — do not fill every step. Make it \
rhythmically interesting and locked to the chords.
- Drums: each of kick/snare/hat is an array of 0/1, exactly one entry per grid \
step for the whole loop (bars * 16 entries). Put kicks on strong beats, snares \
on the backbeat (steps 4 and 12 of each bar), hats as the groove suggests.
- Choose a `bpm` that fits the style if the user didn't pin one.
- Keep it tasteful and idiomatic for the requested genre/vibe."""


def build_user_prompt(vibe: str, bars: int, bpm) -> str:
    total_steps = bars * STEPS_PER_BAR
    lines = [
        f"Compose a {bars}-bar loop ({total_steps} grid steps total).",
        f"Vibe / style: {vibe.strip() or 'anything you think sounds great'}.",
    ]
    if bpm:
        lines.append(f"Target tempo: {bpm} BPM.")
    lines.append(
        "Each drum array must have exactly "
        f"{total_steps} entries (one per step)."
    )
    return "\n".join(lines)


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/generate", methods=["POST"])
def generate():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return (
            jsonify(
                error="ANTHROPIC_API_KEY is not set. Copy .env.example to .env "
                "and add your key, then restart the server."
            ),
            503,
        )

    data = request.get_json(silent=True) or {}
    vibe = str(data.get("prompt", ""))
    bars = max(1, min(int(data.get("bars", 4) or 4), 8))
    bpm = data.get("bpm")

    client = anthropic.Anthropic()
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=16000,
            thinking={"type": "adaptive"},
            system=SYSTEM_PROMPT,
            output_config={
                "format": {"type": "json_schema", "schema": ARRANGEMENT_SCHEMA}
            },
            messages=[{"role": "user", "content": build_user_prompt(vibe, bars, bpm)}],
        )
    except anthropic.APIStatusError as exc:
        return jsonify(error=f"Claude API error ({exc.status_code}): {exc.message}"), 502
    except anthropic.APIConnectionError:
        return jsonify(error="Could not reach the Claude API. Check your connection."), 502

    text = next((b.text for b in response.content if b.type == "text"), "")
    try:
        arrangement = json.loads(text)
    except json.JSONDecodeError:
        return jsonify(error="Model returned malformed JSON. Try again."), 502

    arrangement["bars"] = bars
    arrangement["stepsPerBar"] = STEPS_PER_BAR
    return jsonify(arrangement)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    print(f"\n  🎹  Studio running at http://127.0.0.1:{port}\n")
    app.run(host="127.0.0.1", port=port, debug=True)
