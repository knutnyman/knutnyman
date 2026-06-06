// AI Studio — front-end audio engine, sequencer, editor, MIDI export.
// The arrangement is a 16th-note grid: melody/chords are {start,length} in
// steps, drums are 0/1 arrays one-per-step. Claude fills it; you edit it here.

const $ = (id) => document.getElementById(id);

const state = {
  arr: null,        // current arrangement
  ctx: null,        // AudioContext
  master: null,     // master gain
  playing: false,
  currentStep: 0,
  nextNoteTime: 0,
  schedulerId: null,
  scheduled: [],    // [{step, time}] for the moving playhead
};

const LOOKAHEAD = 25;          // ms between scheduler ticks
const SCHEDULE_AHEAD = 0.12;   // seconds of audio scheduled in advance

// ---------- helpers ----------
const totalSteps = () => state.arr.bars * state.arr.stepsPerBar;
const bpm = () => Math.max(40, Math.min(220, Number($("bpm").value) || 120));
const stepDur = () => 60 / bpm() / 4; // one 16th note in seconds
const midiToFreq = (m) => 440 * Math.pow(2, (m - 69) / 12);

function setStatus(msg, isError = false) {
  const el = $("status");
  el.textContent = msg;
  el.classList.toggle("error", isError);
}

// ---------- audio engine ----------
function ensureCtx() {
  if (!state.ctx) {
    state.ctx = new (window.AudioContext || window.webkitAudioContext)();
    state.master = state.ctx.createGain();
    state.master.gain.value = 0.9;
    state.master.connect(state.ctx.destination);
  }
  if (state.ctx.state === "suspended") state.ctx.resume();
}

// A simple ADSR-ish voice: oscillator -> lowpass -> gain -> master.
function voice(midi, time, dur, { type = "sawtooth", gain = 0.25, cutoff = 4000 } = {}) {
  const ctx = state.ctx;
  const osc = ctx.createOscillator();
  const filt = ctx.createBiquadFilter();
  const amp = ctx.createGain();

  osc.type = type;
  osc.frequency.value = midiToFreq(midi);
  filt.type = "lowpass";
  filt.frequency.value = cutoff;

  const a = 0.008, r = Math.min(0.25, dur * 0.6);
  amp.gain.setValueAtTime(0, time);
  amp.gain.linearRampToValueAtTime(gain, time + a);
  amp.gain.setValueAtTime(gain, time + Math.max(a, dur - r));
  amp.gain.linearRampToValueAtTime(0, time + dur);

  osc.connect(filt).connect(amp).connect(state.master);
  osc.start(time);
  osc.stop(time + dur + 0.02);
}

function kick(time) {
  const ctx = state.ctx;
  const osc = ctx.createOscillator();
  const amp = ctx.createGain();
  osc.frequency.setValueAtTime(150, time);
  osc.frequency.exponentialRampToValueAtTime(45, time + 0.12);
  amp.gain.setValueAtTime(0.9, time);
  amp.gain.exponentialRampToValueAtTime(0.001, time + 0.3);
  osc.connect(amp).connect(state.master);
  osc.start(time);
  osc.stop(time + 0.32);
}

function noiseBurst(time, dur, { cutoff, hp = false, gain = 0.4 }) {
  const ctx = state.ctx;
  const n = Math.floor(ctx.sampleRate * dur);
  const buf = ctx.createBuffer(1, n, ctx.sampleRate);
  const d = buf.getChannelData(0);
  for (let i = 0; i < n; i++) d[i] = Math.random() * 2 - 1;
  const src = ctx.createBufferSource();
  src.buffer = buf;
  const filt = ctx.createBiquadFilter();
  filt.type = hp ? "highpass" : "lowpass";
  filt.frequency.value = cutoff;
  const amp = ctx.createGain();
  amp.gain.setValueAtTime(gain, time);
  amp.gain.exponentialRampToValueAtTime(0.001, time + dur);
  src.connect(filt).connect(amp).connect(state.master);
  src.start(time);
  src.stop(time + dur + 0.02);
}

const snare = (t) => noiseBurst(t, 0.18, { cutoff: 2000, gain: 0.5 });
const hat = (t) => noiseBurst(t, 0.05, { cutoff: 7000, hp: true, gain: 0.25 });

// ---------- sequencer ----------
function scheduleStep(step, time) {
  const arr = state.arr;
  for (const c of arr.chords) {
    if (c.start === step) {
      for (const note of c.notes) {
        voice(note, time, c.length * stepDur(), { type: "triangle", gain: 0.12, cutoff: 2600 });
      }
    }
  }
  for (const m of arr.melody) {
    if (m.start === step) {
      voice(m.pitch, time, m.length * stepDur(), { type: "square", gain: 0.18, cutoff: 5000 });
    }
  }
  if (arr.drums.kick[step]) kick(time);
  if (arr.drums.snare[step]) snare(time);
  if (arr.drums.hat[step]) hat(time);
}

function scheduler() {
  const total = totalSteps();
  while (state.nextNoteTime < state.ctx.currentTime + SCHEDULE_AHEAD) {
    scheduleStep(state.currentStep, state.nextNoteTime);
    state.scheduled.push({ step: state.currentStep, time: state.nextNoteTime });
    state.nextNoteTime += stepDur();
    state.currentStep = (state.currentStep + 1) % total;
  }
}

function play() {
  if (!state.arr || state.playing) return;
  ensureCtx();
  state.playing = true;
  state.currentStep = 0;
  state.nextNoteTime = state.ctx.currentTime + 0.06;
  state.scheduled = [];
  state.schedulerId = setInterval(scheduler, LOOKAHEAD);
  requestAnimationFrame(drawLoop);
  $("play").disabled = true;
  $("stop").disabled = false;
}

function stop() {
  state.playing = false;
  clearInterval(state.schedulerId);
  state.schedulerId = null;
  state.scheduled = [];
  $("play").disabled = false;
  $("stop").disabled = true;
  clearPlayingMarkers();
  drawPianoRoll();
}

// current playing step, derived from audio clock for tight visual sync
function currentPlayheadStep() {
  if (!state.playing || !state.scheduled.length) return -1;
  const now = state.ctx.currentTime;
  let active = state.scheduled[0];
  for (const s of state.scheduled) {
    if (s.time <= now) active = s;
  }
  state.scheduled = state.scheduled.filter((s) => s.time > now - 0.5);
  return active.step;
}

function drawLoop() {
  if (!state.playing) return;
  const step = currentPlayheadStep();
  drawPianoRoll(step);
  markDrumStep(step);
  requestAnimationFrame(drawLoop);
}

// ---------- piano roll ----------
const CELL_W = 24, CELL_H = 14;

function pitchRange() {
  const pitches = state.arr.melody.map((m) => m.pitch);
  if (!pitches.length) return { lo: 60, hi: 84 };
  return { lo: Math.min(...pitches) - 2, hi: Math.max(...pitches) + 2 };
}

function drawPianoRoll(playStep = -1) {
  const canvas = $("pianoroll");
  const { lo, hi } = pitchRange();
  const rows = hi - lo + 1;
  const cols = totalSteps();
  canvas.width = cols * CELL_W;
  canvas.height = rows * CELL_H;
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#1d212b";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // rows (black-key shading)
  for (let r = 0; r < rows; r++) {
    const midi = hi - r;
    const isBlack = [1, 3, 6, 8, 10].includes(((midi % 12) + 12) % 12);
    if (isBlack) {
      ctx.fillStyle = "#191d26";
      ctx.fillRect(0, r * CELL_H, canvas.width, CELL_H);
    }
  }
  // grid lines
  for (let c = 0; c <= cols; c++) {
    ctx.strokeStyle = c % 16 === 0 ? "#46506a" : c % 4 === 0 ? "#39415440" : "#2c334030";
    ctx.lineWidth = c % 16 === 0 ? 2 : 1;
    ctx.beginPath();
    ctx.moveTo(c * CELL_W, 0);
    ctx.lineTo(c * CELL_W, canvas.height);
    ctx.stroke();
  }
  // notes
  for (const m of state.arr.melody) {
    const r = hi - m.pitch;
    if (r < 0 || r >= rows) continue;
    ctx.fillStyle = "#ff8a5b";
    ctx.fillRect(m.start * CELL_W + 1, r * CELL_H + 1, m.length * CELL_W - 2, CELL_H - 2);
  }
  // playhead
  if (playStep >= 0) {
    ctx.fillStyle = "#5bd6ff33";
    ctx.fillRect(playStep * CELL_W, 0, CELL_W, canvas.height);
  }
}

function onRollClick(e) {
  const canvas = $("pianoroll");
  const rect = canvas.getBoundingClientRect();
  const { lo, hi } = pitchRange();
  const col = Math.floor((e.clientX - rect.left) / CELL_W);
  const row = Math.floor((e.clientY - rect.top) / CELL_H);
  const pitch = hi - row;
  if (col < 0 || col >= totalSteps()) return;

  // toggle: remove a note covering this cell, else add a 1-step note
  const idx = state.arr.melody.findIndex(
    (m) => m.pitch === pitch && col >= m.start && col < m.start + m.length
  );
  if (idx >= 0) {
    state.arr.melody.splice(idx, 1);
  } else {
    state.arr.melody.push({ pitch, start: col, length: 1 });
    if (state.ctx) voice(pitch, state.ctx.currentTime, 0.2, { type: "square", gain: 0.18, cutoff: 5000 });
  }
  drawPianoRoll();
}

// ---------- chords + drums UI ----------
function renderChords() {
  const wrap = $("chord-list");
  wrap.innerHTML = "";
  for (const c of state.arr.chords) {
    const el = document.createElement("div");
    el.className = "chord";
    el.innerHTML = `<b>${c.name}</b><small>bar ${Math.floor(c.start / 16) + 1}</small>`;
    wrap.appendChild(el);
  }
}

function renderDrums() {
  const wrap = $("drums");
  wrap.innerHTML = "";
  for (const name of ["kick", "snare", "hat"]) {
    const row = document.createElement("div");
    row.className = "drum-row";
    row.innerHTML = `<div class="drum-label">${name}</div>`;
    const steps = document.createElement("div");
    steps.className = "steps";
    state.arr.drums[name].forEach((on, i) => {
      const cell = document.createElement("div");
      cell.className = "step" + (on ? " on" : "") + (i % 4 === 0 ? " beat" : "");
      cell.dataset.row = name;
      cell.dataset.i = i;
      cell.addEventListener("click", () => {
        state.arr.drums[name][i] = state.arr.drums[name][i] ? 0 : 1;
        cell.classList.toggle("on");
      });
      steps.appendChild(cell);
    });
    row.appendChild(steps);
    wrap.appendChild(row);
  }
}

function markDrumStep(step) {
  clearPlayingMarkers();
  if (step < 0) return;
  document.querySelectorAll(`.step[data-i="${step}"]`).forEach((c) => c.classList.add("playing"));
}
function clearPlayingMarkers() {
  document.querySelectorAll(".step.playing").forEach((c) => c.classList.remove("playing"));
}

// ---------- render everything ----------
function loadArrangement(arr) {
  state.arr = arr;
  $("bpm").value = arr.bpm || 120;
  $("info").textContent = `${arr.title || "Untitled"} · ${arr.key || ""} · ${arr.bars} bars`;
  renderChords();
  renderDrums();
  drawPianoRoll();
  ["play", "export"].forEach((id) => ($(id).disabled = false));
}

// ---------- MIDI export ----------
function exportMidi() {
  const TPQ = 96, perStep = TPQ / 4; // 16th note
  const events = []; // {tick, type:'on'|'off', ch, note, vel}

  const add = (start, length, note, ch) => {
    events.push({ tick: start * perStep, on: true, ch, note });
    events.push({ tick: (start + length) * perStep, on: false, ch, note });
  };
  for (const c of state.arr.chords) for (const n of c.notes) add(c.start, c.length, n, 0);
  for (const m of state.arr.melody) add(m.start, m.length, m.pitch, 1);
  const drumMap = { kick: 36, snare: 38, hat: 42 };
  for (const [name, note] of Object.entries(drumMap)) {
    state.arr.drums[name].forEach((on, i) => { if (on) add(i, 1, note, 9); });
  }
  events.sort((a, b) => a.tick - b.tick || (a.on === b.on ? 0 : a.on ? 1 : -1));

  // build single track
  const track = [];
  // tempo meta
  const usPerBeat = Math.round(60000000 / bpm());
  pushVar(track, 0);
  track.push(0xff, 0x51, 0x03, (usPerBeat >> 16) & 255, (usPerBeat >> 8) & 255, usPerBeat & 255);

  let last = 0;
  for (const e of events) {
    pushVar(track, e.tick - last);
    last = e.tick;
    const status = (e.on ? 0x90 : 0x80) | e.ch;
    track.push(status, e.note & 127, e.on ? 90 : 0);
  }
  pushVar(track, 0);
  track.push(0xff, 0x2f, 0x00); // end of track

  const bytes = [];
  pushStr(bytes, "MThd");
  push32(bytes, 6);
  push16(bytes, 0);        // format 0
  push16(bytes, 1);        // 1 track
  push16(bytes, TPQ);
  pushStr(bytes, "MTrk");
  push32(bytes, track.length);
  for (const b of track) bytes.push(b);

  const blob = new Blob([new Uint8Array(bytes)], { type: "audio/midi" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = (state.arr.title || "studio-loop").replace(/[^\w-]+/g, "_") + ".mid";
  a.click();
  URL.revokeObjectURL(a.href);
}

function pushVar(arr, value) {
  let buffer = value & 0x7f;
  while ((value >>= 7)) { buffer <<= 8; buffer |= (value & 0x7f) | 0x80; }
  while (true) { arr.push(buffer & 255); if (buffer & 0x80) buffer >>= 8; else break; }
}
const push16 = (a, v) => a.push((v >> 8) & 255, v & 255);
const push32 = (a, v) => a.push((v >> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255);
const pushStr = (a, s) => { for (const ch of s) a.push(ch.charCodeAt(0)); };

// ---------- generate (call backend) ----------
async function generate() {
  const btn = $("generate");
  btn.disabled = true;
  setStatus("Composing with Claude… (this can take ~10–20s)");
  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt: $("prompt").value,
        bars: Number($("bars").value),
        bpm: Number($("bpm").value) || undefined,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Generation failed");
    if (state.playing) stop();
    loadArrangement(data);
    setStatus(`Loaded “${data.title}”. Hit Play, edit the notes, or export MIDI.`);
  } catch (err) {
    setStatus(err.message, true);
  } finally {
    btn.disabled = false;
  }
}

// ---------- wire up ----------
$("generate").addEventListener("click", generate);
$("play").addEventListener("click", play);
$("stop").addEventListener("click", stop);
$("export").addEventListener("click", exportMidi);
$("pianoroll").addEventListener("click", onRollClick);
$("bpm").addEventListener("change", () => { if (state.arr) $("info"); });
