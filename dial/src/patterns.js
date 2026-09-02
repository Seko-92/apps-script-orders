/**
 * patterns.js — what the flip-disc field DOES. One function per idea.
 *
 * ⭐ EACH PATTERN IS DRAWN AT GRID RESOLUTION AND THRESHOLDED INTO DISCS. That means a
 *   pattern can use ordinary canvas primitives — arcs, lines, rects — and still come out as
 *   a mechanical field of discs. Ten lines buys a real mechanism instead of a hand-authored
 *   bitmap, which is the same trick that gave the letters real Oswald letterforms.
 *
 * ⚠ THE GRID IS SMALL AND THAT IS THE DESIGN CONSTRAINT, not an obstacle. 280x121 at a 4px
 *   pitch is 70 x 30 discs. Anything needing fine detail will sample to mush — measured on
 *   "MOTOR SERVICE" at 4px, which came out a solid bar. Draw BIG shapes.
 */
'use strict';

const TAU = Math.PI * 2;

/** ⭐ THE FOUR-STROKE. This is an automotive parts business and the medium is a mechanism —
 *  a piston stroking in its bore is the most on-brand thing this banner could possibly do,
 *  and it is legible at any size because it is one big moving shape. */
function piston(c, cols, rows, phase) {
  const cx = cols * 0.42;
  const bore = Math.min(22, cols * 0.30);
  const top = 3, bottom = rows - 9;
  const crankY = rows - 5, crankR = 4.2;

  const ang = phase * TAU;
  const pinX = cx + Math.sin(ang) * crankR;
  const pinY = crankY - Math.cos(ang) * crankR;
  const stroke = (bottom - top);
  const headY = top + (0.5 - Math.cos(ang) * 0.5) * stroke * 0.72;

  c.strokeStyle = '#fff'; c.fillStyle = '#fff';

  c.lineWidth = 1.9;                                   // the bore walls
  c.beginPath();
  c.moveTo(cx - bore / 2, top - 2); c.lineTo(cx - bore / 2, bottom);
  c.moveTo(cx + bore / 2, top - 2); c.lineTo(cx + bore / 2, bottom);
  c.stroke();

  c.fillRect(cx - bore / 2, headY, bore, 4.2);          // the piston crown
  c.fillRect(cx - bore / 2, headY + 5.4, bore, 1.2);    // one ring

  c.lineWidth = 2.2;                                    // the connecting rod
  c.beginPath(); c.moveTo(cx, headY + 4); c.lineTo(pinX, pinY); c.stroke();

  c.lineWidth = 1.9;                                    // the crank circle and its pin
  c.beginPath(); c.arc(cx, crankY, crankR, 0, TAU); c.stroke();
  c.beginPath(); c.arc(pinX, pinY, 1.6, 0, TAU); c.fill();
}

/** ⭐ THE TICKER. What a departure board does between announcements: text walks across the
 *  field. ⚠ It must WRAP seamlessly or the loop shows a seam — so the string is drawn twice,
 *  a full period apart, and the window slides exactly one period. */
function ticker(c, cols, rows, phase, text, strip) {
  // ⚠ THE FIRST CUT DREW THE STRING TWICE, `measureText` apart, AND THE COPIES OVERLAPPED —
  //   it rendered "HOUSTOSTON". measureText and the rasterised glyph run are not the same
  //   width once you are drawing into a 70-unit grid. So the strip is rendered ONCE, its
  //   real pixel width IS the period, and the loop blits it twice. Seamless by construction
  //   rather than by getting a measurement right.
  const w = strip.width;
  const x = -((phase * w) % w);
  c.drawImage(strip, x, 0);
  c.drawImage(strip, x + w, 0);
}

/** Renders the ticker's text once. Its width is the loop period. */
function tickerStrip(createCanvas, cols, rows, text) {
  const probe = createCanvas(8, 8).getContext('2d');
  const size = Math.max(6, Math.round(rows * 0.46));
  probe.font = '600 ' + size + 'px Oswald';
  const tw = Math.ceil(probe.measureText(text).width);
  const gap = Math.round(cols * 0.55);                  // clear space between repeats
  const cv = createCanvas(tw + gap, rows);
  const c = cv.getContext('2d');
  c.clearRect(0, 0, cv.width, rows);
  c.fillStyle = '#fff';
  c.font = '600 ' + size + 'px Oswald';
  c.textBaseline = 'middle';
  c.fillText(text, 0, rows / 2);
  return cv;
}

/** ⭐ THE REFRESH. The Solari signature in discs: a wavefront sweeps the field, every disc
 *  it passes flips, and behind it the mark is standing. Then it sweeps back to clear. */
function refresh(cols, rows, phase, mark) {
  const field = [];
  // Diagonal wavefront position, out and back across the loop.
  // ⚠ SWEEP FAST, HOLD LONG. The first cut spread the sweep across the whole loop, which
  //   left the panel EMPTY for the frames before the wavefront entered — a dead board, not
  //   an idling one. The wipe now takes a quarter of the loop at each end and the mark
  //   stands for the half in between.
  let p;
  if (phase < 0.25)      p = phase / 0.25;          // wipe on
  else if (phase < 0.78) p = 1;                     // hold, mark standing
  else                   p = 1 - (phase - 0.78) / 0.22;  // wipe off
  const head = p * (cols + rows * 0.8) * 1.15 - rows * 0.4;
  for (let y = 0; y < rows; y++) {
    const row = [];
    for (let x = 0; x < cols; x++) {
      const d = head - (x + y * 0.8);
      // A band of churn AT the wavefront, settled behind it, dark ahead of it.
      // ⚠ A WIDER, DENSER CHURN BAND. The first cut used ±3 and a 1-in-3 fill, which left
      //   two of every four frames nearly empty — the loop read as dead rather than idle.
      if (d > 5) row.push(mark[y][x]);
      else if (d > -7) {
        const n = (x * 5 + y * 11 + Math.floor(phase * 137)) % 5;
        row.push(n < 2 || (d > 1 && mark[y][x]));       // settling INTO the mark, not onto black
      } else row.push(false);
    }
    field.push(row);
  }
  return field;
}

/** ⭐ THE BELT. Warehouse-native: parts travelling right to left, the way the floor works.
 *  Blocks of different lengths so it never reads as a repeating tile. */
function belt(c, cols, rows, phase) {
  c.fillStyle = '#fff';
  const lane = [rows * 0.30, rows * 0.58];
  const sizes = [7, 4, 11, 5, 8, 3, 9, 6];
  const period = cols * 1.6;
  lane.forEach((y, li) => {
    const speed = li ? 1 : 0.72;                      // two lanes, different speeds
    for (let i = 0; i < sizes.length; i++) {
      const w = sizes[(i + li * 3) % sizes.length];
      let x = (i * period / sizes.length) - phase * period * speed;
      x = ((x % period) + period) % period - w;
      c.fillRect(x, y, w, 3);
    }
  });
  c.fillStyle = '#fff';                                // the belt line itself
  c.fillRect(0, rows - 3, cols, 1);
}

/** ⭐ THE NIGHT. Sparse points drifting — the quietest of the set, and the closest in
 *  feeling to the night dial this banner already wore. Deterministic, so it loops exactly. */
function night(c, cols, rows, phase) {
  c.fillStyle = '#fff';
  // ⚠ 0.018 of the grid was ~38 points in 2100 — indistinguishable from a dead panel.
  const N = Math.round(cols * rows * 0.055);
  for (let i = 0; i < N; i++) {
    // ⚠ TWO INDEPENDENT HASHES. Using (i*37)%rows for y walked a regular diagonal and the
    //   field showed visible stripes instead of scatter — obvious at 2x, invisible in the
    //   arithmetic.
    const h1 = ((i * 2654435761) % 10007) / 10007;
    const h2 = ((i * 40503 + 17) % 10009) / 10009;
    const sp = 0.25 + (i % 5) * 0.12;
    const x = (h1 * cols + phase * cols * sp) % cols;
    const y = Math.floor(h2 * rows);
    // a slow twinkle, out of phase per point, so the field is never static
    const tw = Math.sin((phase * 2 + h1 * 7) * Math.PI * 2);
    if (tw > -0.2) c.fillRect(Math.floor(x), y, 1, 1);
  }
}

module.exports = { piston, ticker, tickerStrip, refresh, belt, night };
