/**
 * ambient.js — a banner slot as a flip-disc field that carries NO DATA.
 *
 * ⚠⚠ THIS DELETES A BUG CLASS RATHER THAN MANAGING ONE. Every flash chased on 2026-09-02
 *    traced to the same mechanism: a number changed -> the `=IMAGE()` URL changed -> Sheets
 *    refetched -> the cell sat EMPTY while it loaded. That is structural. Coarsening the
 *    URL only makes it blink less often.
 *
 *    A loop with no data in it has a URL that never changes, so it is fetched ONCE and
 *    never again. **It cannot flash — there is nothing to refetch.**
 *
 * ⭐ IT LOOPS FOREVER, unlike the settle GIF which plays once. A settle is an EVENT and must
 *   not repeat; ambience is a STATE and must not stop.
 *
 * ⚠⚠ WHAT MAY NEVER BE COVERED: F2 (Pick ID for Shipping, the F2:G2 merge) and H2 (Pick ID
 *    for Adjustment). A floating image swallows CLICKS as well as pixels, so covering those
 *    breaks the accountability gate on printing and picking — a floor outage, not a
 *    cosmetic bug. Row 1 is entirely safe; in row 2 only columns A-E are.
 */
'use strict';

const { createCanvas } = require('@napi-rs/canvas');
const M = require('./board-matrix');

/** Two crossing swells. Continuous in t, so the loop seam is invisible. */
let WIDE = false;
/** ⚠ THE TUNING DOES NOT CARRY ACROSS ASPECT RATIOS. At 30 disc rows (280x121) a 4.4-cycle
 *  swell reads as drifting blobs; at 14 rows (876x56) the same numbers thin into horizontal
 *  streaks, because there is no vertical room for a crest to curve. A wide, short field
 *  needs FEWER cycles so each blob is wider than it is tall. Found by rendering both, not
 *  by reasoning. */
function setWide(v) { WIDE = !!v; }

function wave(x, y, t, cols, rows) {
  const u = x / cols, v = y / rows;
  const fx = WIDE ? 2.0 : 4.4, fy = WIDE ? 0.9 : 1.8;
  const gx = WIDE ? 0.8 : 1.7, gy = WIDE ? 1.3 : 2.6;
  const a = Math.sin((u * fx + v * fy - t * 0.85) * Math.PI);
  const b = Math.sin((u * gx - v * gy + t * 0.55) * Math.PI);
  return (a * 0.62 + b * 0.38) * 0.5 + 0.5;
}

/** ⚠ ONLY THE CRESTS LIGHT. The first pass used 0.62 and lit half the grid — a solid white
 *  block, not a wave. A flip-dot field idles MOSTLY DARK; the beauty is the few discs on. */
const CREST = 0.80;

/** Sample text into the disc grid once. ⚠ Font size is in DISCS, not pixels. */
function markField(cols, rows, lines) {
  const off = createCanvas(cols, rows);
  const c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows);
  c.fillStyle = '#fff';
  c.textAlign = 'center';
  c.textBaseline = 'middle';
  const n = lines.length;
  lines.forEach((ln, i) => {
    c.font = ln.weight + ' ' + ln.size + 'px Oswald';
    c.fillText(ln.text, cols / 2 + (ln.dx ? ln.dx / M.PITCH : 0), rows / 2 + (ln.dy || 0));
  });
  const px = c.getImageData(0, 0, cols, rows).data;
  const f = [];
  for (let y = 0; y < rows; y++) {
    const row = [];
    for (let x = 0; x < cols; x++) row.push(px[(y * cols + x) * 4 + 3] > 110);
    f.push(row);
  }
  return f;
}

/** How assembled the mark is at this point in the loop. Held long — it is the payoff. */
function gatherAt(phase, mode) {
  if (mode !== 'mark') return 0;
  if (phase < 0.26)      return 0;
  if (phase < 0.40)      return (phase - 0.26) / 0.14;
  if (phase < 0.80)      return 1;
  if (phase < 0.94)      return 1 - (phase - 0.80) / 0.14;
  return 0;
}

function composeAmbient(mode, phase, cols, rows, mark) {
  const gather = gatherAt(phase, mode);
  const field = [];
  for (let y = 0; y < rows; y++) {
    const row = [];
    for (let x = 0; x < cols; x++) {
      const w = wave(x, y, phase * 2, cols, rows);
      // ⚠ A PER-DISC THRESHOLD, never a global fade. A disc is one side or the other —
      //   opacity is impossible on a flip-dot. Each crosses at its own point, so the mark
      //   assembles disc by disc the way the real mechanism does.
      const jitter = ((x * 7 + y * 13) % 17) / 17;
      row.push(gather > jitter ? mark[y][x] : w > CREST);
    }
    field.push(row);
  }
  return field;
}

/** The sentinel colour punched into window areas, later mapped to the GIF's transparent
 *  index. Chosen because it appears nowhere in the board's palette. */
const WINDOW = '#ff00ff';

function drawAmbient(ctx, o) {
  const s = o.scale || 1, W = o.w, H = o.h;
  const g = ctx.createLinearGradient(0, 0, 0, H * s);
  g.addColorStop(0, '#26221c');
  g.addColorStop(0.14, '#141210');
  g.addColorStop(1, '#100e0c');
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, W * s, H * s);

  const cols = Math.floor(W / M.PITCH), rows = Math.floor(H / M.PITCH);
  M.drawMatrix(ctx, { scale: s, w: W, h: H,
                      field: composeAmbient(o.mode, o.phase, cols, rows, o.mark) });

  // The dial's own right-edge seam, so the block dissolves into the flat band beside it
  // instead of reading as a sticker pasted on.
  if (o.seam !== false) {
    const seam = ctx.createLinearGradient((W - 26) * s, 0, W * s, 0);
    seam.addColorStop(0, 'rgba(26,26,26,0)');
    seam.addColorStop(1, '#1a1a1a');
    ctx.fillStyle = seam;
    ctx.fillRect((W - 26) * s, 0, 26 * s, H * s);
  }

  // Windows: punched last so nothing draws over a live cell.
  (o.windows || []).forEach((win) => {
    ctx.fillStyle = WINDOW;
    ctx.fillRect(win.x * s, win.y * s, win.w * s, win.h * s);
    ctx.strokeStyle = '#0a0908';
    ctx.lineWidth = Math.max(1, 1.2 * s);
    ctx.strokeRect(win.x * s - 0.6 * s, win.y * s - 0.6 * s,
                   win.w * s + 1.2 * s, win.h * s + 1.2 * s);
  });
}

function gridSize(w, h) {
  return { cols: Math.floor(w / M.PITCH), rows: Math.floor(h / M.PITCH) };
}

module.exports = { wave, setWide, CREST, markField, composeAmbient, drawAmbient, gridSize, WINDOW };
