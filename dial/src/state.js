/**
 * state.js — the query string becomes what the dial SAYS.
 *
 * ⭐ THE DIAL SHOWS WHAT MATTERS IN THAT STATE, and nothing else. A face that prints the
 *    same five numbers all day is a readout; a face whose centre changes with the shift is
 *    an instrument. So the centre carries one big fact per verdict:
 *
 *      rest   the clock            — the floor is asleep; the day just done is on the ring
 *      clear  the clock            — open, queue empty, nothing to act on
 *      busy   the oldest wait      — the number that costs money when it grows
 *      late   the oldest wait      — same fact, louder, because it crossed the 3h line
 *      stale  time since last sync — the pipeline, not the orders, is the problem
 *
 * ⚠ BLANK, NEVER A REASSURING ZERO. A missing parameter renders an em dash. "0 waiting" on
 *   a table nobody could read is a reassuring label on a dangerous state, which this
 *   codebase already rules is a bug (__SparkData's pubNumBlank exists for the same reason).
 */
'use strict';

const { fmtMins, fmtClock } = require('./draw');

const VERDICTS = ['rest', 'clear', 'busy', 'late', 'stale'];

function num(v, fallback) {
  if (v === undefined || v === null || v === '') return fallback;
  const n = Number(v);
  return isFinite(n) ? n : fallback;
}

/** "12,0,3,..." (24 of them) -> number[24]. Anything unreadable becomes zeros. */
function parseHours(raw) {
  const out = new Array(24).fill(0);
  if (!raw) return out;
  const parts = String(raw).split(',');
  for (let i = 0; i < 24 && i < parts.length; i++) {
    const n = Number(parts[i]);
    out[i] = isFinite(n) && n > 0 ? n : 0;
  }
  return out;
}

/** "1425" -> minutes of day. Falls back to the host clock, which is only ever a preview. */
function parseClock(raw) {
  const s = String(raw == null ? '' : raw).replace(/\D/g, '');
  if (s.length === 3 || s.length === 4) {
    const h = Number(s.slice(0, s.length - 2));
    const m = Number(s.slice(-2));
    if (h >= 0 && h < 24 && m >= 0 && m < 60) return h * 60 + m;
  }
  const d = new Date();
  return d.getHours() * 60 + d.getMinutes();
}

const dash = (n, f) => (n == null || !isFinite(n) ? '—' : f(n));

/**
 * @param {object} q  raw query params (strings)
 * @returns {object}  the render state consumed by drawDial
 */
function buildState(q) {
  q = q || {};
  const verdict = VERDICTS.indexOf(q.s) >= 0 ? q.s : 'clear';
  const nowMin  = parseClock(q.t);

  const oldest    = num(q.o, null);
  const toGrab    = num(q.g, null);
  const received  = num(q.r, null);
  const shipped   = num(q.p, null);
  const untilOpen = num(q.u, null);
  const syncMin   = num(q.y, null);
  const pastLine  = num(q.l, null);
  const hours     = parseHours(q.h);

  // ⚠ The day's pips are OPTIONAL because they overlap the F1:H1 curve — both answer
  //   "when did the day happen". Which one should own that is a live question; this makes
  //   it something you can look at instead of argue about. `np=1` = no pips.
  const st = { verdict, nowMin, oldestMin: oldest || 0, hours, showPips: q.np !== '1' };

  if (verdict === 'rest') {
    st.big     = fmtClock(nowMin);
    st.caption = 'RESTING';
    st.flank = [
      { value: dash(shipped,  String), label: 'shipped',    tone: 'quiet' },
      { value: dash(received, String), label: 'received',   tone: 'quiet' },
      { value: dash(untilOpen, fmtMins), label: 'until open', tone: 'accent' }
    ];
  } else if (verdict === 'stale') {
    // ⚠ The queue is still shown. A dead pipeline does not make the floor's work vanish —
    //   it means nobody can TRUST the count, which is why the caption says so plainly.
    st.big     = dash(syncMin, fmtMins);
    st.caption = 'SINCE LAST SYNC';
    // ⚠ NO "STALE / PIPELINE" ROW. It was a word sitting where numbers go, and the face
    //   already says STALE twice over — in red, and in the caption. The flank's job is to
    //   keep the floor's own figures visible; the fault is the FACE's message.
    st.flank = [
      { value: dash(toGrab,   String), label: 'to grab',   tone: 'accent' },
      { value: dash(received, String), label: 'received',  tone: 'quiet'  },
      { value: dash(shipped,  String), label: 'out today', tone: 'quiet'  }
    ];
  } else if (verdict === 'clear') {
    st.big     = fmtClock(nowMin);
    st.caption = 'ALL CLEAR';
    st.flank = [
      { value: dash(toGrab,   String), label: 'to grab',   tone: 'quiet' },
      { value: dash(received, String), label: 'received',  tone: 'quiet' },
      { value: dash(shipped,  String), label: 'out today', tone: 'accent' }
    ];
  } else {
    // busy · late — the hands own the face, so the wait moves to the flank and LEADS.
    // ⭐ The wedge and its number then sit side by side: the shape and the figure it is the
    //   shape of. That adjacency is the whole argument for drawing a duration at all.
    st.big     = '';
    st.caption = '';
    const third = (verdict === 'late' && pastLine != null && pastLine > 0)
      ? { value: String(pastLine), label: 'past the line', tone: 'accent' }
      : { value: dash(shipped, String), label: 'out today', tone: 'quiet' };
    st.flank = [
      { value: dash(oldest, fmtMins),  label: 'oldest waiting', tone: 'accent' },
      { value: dash(toGrab,   String), label: 'to grab',        tone: 'quiet'  },
      third
    ];
  }

  return st;
}

module.exports = { buildState, VERDICTS, parseHours, parseClock };
