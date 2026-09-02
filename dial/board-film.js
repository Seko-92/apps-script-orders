/**
 * board-film.js — render the board so a human can judge it BEFORE it goes near the sheet.
 *
 * ⚠⚠ THIS FILE IS THE PROCESS RULE. Every other surface here renders headlessly; the SHEET
 *    cannot, so the user's eyes are its first render. The 2026-08-31 round that skipped
 *    this step was rejected wholesale. Build here, look, THEN point it at the sheet.
 *
 * ⭐ IT COMPOSITES OVER MOCK LIVE CELLS ON PURPOSE. The strip's two windows are holes; a
 *   render of the strip ALONE would show them as empty and prove nothing. What has to be
 *   judged is the strip sitting ON the band with E1's pulse and the F1:H1 curve showing
 *   through — so this draws the band, the cells, and then the board on top, in that order.
 *
 * Outputs to renders/:  board-still.png · board-flip.png (contact sheet) · board-band.png
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const T = require('./src/board-terminal');
const B = require('./src/board');

const OUT = path.join(__dirname, 'renders');
const BAND = '#1a1a1a';
const SCALE = Number(process.env.SCALE || 2);      // 2x so the flap edges are judgeable
const S = (v) => Math.round(v * SCALE);

registerFonts();

// Two real states, so the render shows an actual transition rather than a contrived one.
const REST = T.buildBoard({ s: 'rest', t: '1914', p: '30', r: '8', u: '295', o: '95', g: '6' });
const BUSY = T.buildBoard({ s: 'busy',  t: '1032', o: '29', g: '12', p: '14' });

/** Mock of what the LIVE CELLS underneath look like, so the windows can be judged. */
function drawLiveCells(ctx) {
  // E1 — the system pulse. Real text, real position: it starts at D1's right edge.
  ctx.save();
  ctx.font = `400 ${S(10)}px Oswald`;
  ctx.textBaseline = 'middle';
  // D1 — the headline, uncovered in v1. Two lines, exactly as the sheet renders it.
  ctx.fillStyle = '#f0ece0'; ctx.font = `400 ${S(11)}px Oswald`;
  ctx.fillText('the floor is asleep', S(T.ANCHOR.w + 12), S(21));
  ctx.fillStyle = '#8b8578'; ctx.font = `400 ${S(9)}px Oswald`;
  ctx.fillText('waiting: 2', S(T.ANCHOR.w + 12), S(37));
  ctx.font = `400 ${S(10)}px Oswald`;
  ctx.fillStyle = '#b9b4a8';
  // ⚠ The real E1 leads with ⊙ and a status lamp. Oswald has neither glyph, so writing them
  //   here renders TOFU and the mock would be lying about the sheet — the same tofu class
  //   that hit ⛶ and 📻 on the warehouse tablet. Draw the lamp, then the text.
  const lx = T.ANCHOR.w + T.COL_D + 18;
  ctx.beginPath(); ctx.arc(S(lx), S(28), S(3), 0, Math.PI * 2);
  ctx.fillStyle = '#7e8894'; ctx.fill();
  ctx.fillStyle = '#b9b4a8';
  ctx.fillText('RESTING · 7:14 PM · 8h 50m ago', S(lx + 9), S(28));
  ctx.restore();

  // F1:H1 — the day curve. Bars, the way the real SPARKLINE renders.
  const x0 = T.ANCHOR.w + T.COL_D + T.COL_E + 14;
  const bars = [2,1,3,2,5,8,12,17,14,9,13,7,4,2,6,3,1,2];
  ctx.save();
  ctx.fillStyle = '#7e8894';
  bars.forEach((v, i) => {
    const h = 2 + v * 1.9, w = 6;
    ctx.fillRect(S(x0 + i * 8.4), S(46 - h), S(w), S(h));
  });
  ctx.restore();
}

/** The full banner: band, live cells, then the two board images on top. */
function banner(t, from, to, label) {
  const W = T.ANCHOR.w + T.STRIP.w, H = 121;
  const cv = createCanvas(S(W), S(H));
  const ctx = cv.getContext('2d');

  ctx.fillStyle = BAND; ctx.fillRect(0, 0, cv.width, cv.height);
  ctx.fillStyle = '#fff8e7';                                  // row 2, the cream band
  ctx.fillRect(0, S(56), cv.width, S(65));
  drawLiveCells(ctx);

  // ⚠⚠ EACH BOARD IMAGE IS RENDERED TO ITS OWN CANVAS AND THEN COMPOSITED, because that
  //    is exactly what Sheets does with a floating image — and because drawStrip CUTS its
  //    windows with clearRect. Drawn straight onto this canvas, that clear would erase the
  //    band and the live cells underneath, and the render would show two white holes while
  //    claiming to prove transparency. The first cut of this file did precisely that.
  const a = createCanvas(S(T.ANCHOR.w), S(T.ANCHOR.h));
  T.drawAnchor(a.getContext('2d'), from, to, t, SCALE);
  ctx.drawImage(a, 0, 0);
  // ⚠ v1 draws the ANCHOR ONLY. D1, E1 and F1:H1 stay uncovered — that is the design, not
  //   an omission, so the mock must show them live and unframed or it flatters the build.

  if (label) {
    ctx.font = `600 ${S(7)}px Oswald`;
    ctx.fillStyle = '#ffd400';
    ctx.textBaseline = 'top';
    ctx.fillText(label, S(6), S(H - 12));
  }
  return cv;
}

function write(name, cv) {
  fs.writeFileSync(path.join(OUT, name), cv.toBuffer('image/png'));
  console.log('  renders/' + name);
}

// ── 1 · the settled board, true size, on the real band ───────────────────────
write('board-band.png', banner(99, REST, REST, null));

// ── 2 · the flip, sampled across the cascade ─────────────────────────────────
const dur = B.loopSeconds(REST.word, BUSY.word);
const shots = [0, 0.10, 0.22, 0.34, 0.46, 0.62, 0.84, dur - B.HOLD_S + 0.05];
const gap = 14;
const sheet = createCanvas(S(T.ANCHOR.w) + gap * 2, (S(121) + gap) * shots.length + gap);
const sctx = sheet.getContext('2d');
sctx.fillStyle = '#2b2b2b'; sctx.fillRect(0, 0, sheet.width, sheet.height);
shots.forEach((t, i) => {
  const one = createCanvas(S(T.ANCHOR.w), S(121));
  const c = one.getContext('2d');
  c.fillStyle = BAND; c.fillRect(0, 0, one.width, one.height);
  T.drawAnchor(c, REST, BUSY, t, SCALE);
  c.font = `600 ${S(6)}px Oswald`;
  c.fillStyle = '#ffd400'; c.textBaseline = 'top';
  c.fillText(`t=${t.toFixed(2)}s`, S(4), S(112));
  sctx.drawImage(one, gap, gap + i * (S(121) + gap));
});
write('board-flip.png', sheet);

// ── 3 · the strip alone at TRUE SIZE (scale 1) — the honest legibility test ──
const one = createCanvas(T.ANCHOR.w + T.STRIP.w, 121);
const octx = one.getContext('2d');
octx.fillStyle = BAND; octx.fillRect(0, 0, one.width, one.height);
octx.fillStyle = '#fff8e7'; octx.fillRect(0, 56, one.width, 65);
(function trueSizeCells() {
  octx.font = '400 10px Oswald'; octx.textBaseline = 'middle';
  octx.fillStyle = '#f0ece0'; octx.font = '400 11px Oswald';
  octx.fillText('the floor is asleep', T.ANCHOR.w + 12, 21);
  octx.fillStyle = '#8b8578'; octx.font = '400 9px Oswald';
  octx.fillText('waiting: 2', T.ANCHOR.w + 12, 37);
  octx.font = '400 10px Oswald';
  const lx = T.ANCHOR.w + T.COL_D + 18;
  octx.beginPath(); octx.arc(lx, 28, 3, 0, Math.PI * 2); octx.fillStyle = '#7e8894'; octx.fill();
  octx.fillStyle = '#b9b4a8';
  octx.fillText('RESTING · 7:14 PM · 8h 50m ago', lx + 9, 28);
  const x0 = T.ANCHOR.w + T.COL_D + T.COL_E + 14;
  octx.fillStyle = '#7e8894';
  [2,1,3,2,5,8,12,17,14,9,13,7,4,2,6,3,1,2].forEach((v, i) => {
    const h = 2 + v * 1.9; octx.fillRect(x0 + i * 8.4, 46 - h, 6, h);
  });
})();
const a1 = createCanvas(T.ANCHOR.w, T.ANCHOR.h);
T.drawAnchor(a1.getContext('2d'), REST, REST, 99, 1);
octx.drawImage(a1, 0, 0);
write('board-truesize.png', one);

console.log(`\nloop = ${dur.toFixed(2)}s  (cap is 8s — trap 7)`);
