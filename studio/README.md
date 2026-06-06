# 🎹 AI Studio

An AI-enabled music **production** tool. Describe a vibe, and Claude composes a
short, editable loop — chords, a melody, and a drum pattern — that you can play
in the browser, tweak note-by-note, and export as a MIDI file to drop into any
DAW (Ableton, Logic, FL Studio, etc.).

This is a starting point, not a finished DAW — it's built to grow.

## What it does

- **Generate** — type a style ("dreamy lo-fi with jazzy chords, 75 BPM") and
  Claude returns a grid-quantized arrangement.
- **Play** — a Web Audio engine sequences chords, melody, and drums in the
  browser (no plugins).
- **Edit** — click the piano roll to add/remove melody notes; click the drum
  grid to toggle hits; change the tempo live.
- **Export MIDI** — download a `.mid` of the whole loop and finish it in your DAW.

The Claude API key stays on the server — the browser never sees it.

## Architecture

```
browser (Web Audio + piano roll + drum grid)  ──fetch──►  Flask backend
        ▲                                                      │
        └──────────── arrangement JSON ◄──── Claude API (claude-opus-4-8)
```

- `app.py` — Flask server + `/api/generate` (asks Claude for a structured
  arrangement via JSON-schema structured outputs).
- `static/` — the UI and audio engine (vanilla JS, no build step).

The arrangement is a 16th-note grid: one bar = 16 steps. Melody/chords are
`{start, length}` in steps; drums are `0/1` arrays, one entry per step.

## Setup

```bash
cd studio
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env       # then edit .env and add your ANTHROPIC_API_KEY
python app.py
```

Open http://127.0.0.1:5001 and click **Generate**.

Get an API key at https://console.anthropic.com/.

## Ideas to build next

- Drag to resize/move melody notes; velocity per note
- A bassline track and a second melodic layer
- Per-track instrument selection / better synths (or sample playback)
- "Regenerate just the drums / just the melody" with the rest locked
- Natural-language edits ("make the chords darker", "add a turnaround in bar 4")
- Save/load projects; export stems as audio (offline render)
