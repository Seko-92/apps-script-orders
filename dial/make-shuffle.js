/**
 * make-shuffle.js — every pattern in ONE loop, with the board wiping between them.
 *
 * ⭐⭐ A SHUFFLE COSTS NOTHING IF IT IS ONE GIF. The alternative — swapping a different image
 *    periodically — needs a trigger, a swap path and Apps Script runtime forever. Sequencing
 *    the patterns inside a single file makes the variety free: still one insert, still never
 *    refetched, still no flash.
 *
 * ⭐ THE TRANSITION IS THE MEDIUM'S OWN. A hard cut between patterns reads as a glitch; a
 *   diagonal wavefront that wipes one display off and the next on is exactly what a real
 *   flip-disc board does when it changes what it shows.
 *
 * ⚠⚠ FILE SIZE IS THE ONLY REAL CONSTRAINT, and it is not about bandwidth — insertImage
 *    STORES the image in the spreadsheet, so a fat GIF is weight every viewer carries on
 *    every load, forever. Measured here rather than guessed.
 */
'use strict';
const fs = require('fs'), path = require('path');
const { GIFEncoder, quantize, applyPalette } = require('gifenc');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A = require('./src/ambient');
const P = require('./src/patterns');
const M = require('./src/board-matrix');
registerFonts();

const E = process.env;
const W = +(E.W||280), H = +(E.H||121);
const { cols, rows } = A.gridSize(W, H);
const mark = A.markField(cols, rows, [{ text: 'HQ', size: 19, weight: '600' }]);
const TSTRIP = P.tickerStrip(createCanvas, cols, rows, 'HQ MOTOR SERVICE · HOUSTON · ');

function fieldFromDraw(fn) {
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows); fn(c);
  const px = c.getImageData(0, 0, cols, rows).data;
  const f = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) r.push(px[(y * cols + x) * 4 + 3] > 110);
    f.push(r);
  }
  return f;
}

const PATTERNS = {
  piston: (ph) => fieldFromDraw((c) => P.piston(c, cols, rows, ph)),
  ticker: (ph) => fieldFromDraw((c) => P.ticker(c, cols, rows, ph, null, TSTRIP)),
  mark:   (ph) => A.composeAmbient('mark', ph, cols, rows, mark),
  belt:   (ph) => fieldFromDraw((c) => P.belt(c, cols, rows, ph)),
  night:  (ph) => fieldFromDraw((c) => P.night(c, cols, rows, ph))
};

/** Wipe from field a to field b: a diagonal wavefront with churn at its head. */
function wipe(a, b, t) {
  const head = t * (cols + rows * 0.8) * 1.15 - rows * 0.4;
  const out = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) {
      const d = head - (x + y * 0.8);
      if (d > 5) r.push(b[y][x]);
      else if (d > -5) r.push(((x * 5 + y * 11 + Math.floor(t * 211)) % 4) < 2);
      else r.push(a[y][x]);
    }
    out.push(r);
  }
  return out;
}

function build(opts) {
  const { fps, hold, wipeSecs, colours, names } = opts;
  const holdF = Math.round(hold * fps), wipeF = Math.round(wipeSecs * fps);
  const gif = GIFEncoder();
  const cv = createCanvas(W, H), ctx = cv.getContext('2d');
  let pal, total = 0;

  const frame = (field) => {
    const g = ctx.createLinearGradient(0, 0, 0, H);
    g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
    ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
    M.drawMatrix(ctx, { scale: 1, w: W, h: H, field });
    const seam = ctx.createLinearGradient(W - 26, 0, W, 0);
    seam.addColorStop(0, 'rgba(26,26,26,0)'); seam.addColorStop(1, '#1a1a1a');
    ctx.fillStyle = seam; ctx.fillRect(W - 26, 0, 26, H);
    const { data } = ctx.getImageData(0, 0, W, H);
    if (!pal) pal = quantize(data, colours, { format: 'rgb565' });
    gif.writeFrame(applyPalette(data, pal, 'rgb565'), W, H, {
      palette: total === 0 ? pal : undefined,
      delay: Math.round(1000 / fps),
      repeat: total === 0 ? 0 : undefined
    });
    total++;
  };

  names.forEach((name, i) => {
    const fn = PATTERNS[name];
    for (let f = 0; f < holdF; f++) frame(fn(f / holdF));
    const next = PATTERNS[names[(i + 1) % names.length]];
    const a = fn(1), b = next(0);
    for (let f = 0; f < wipeF; f++) frame(wipe(a, b, f / wipeF));
  });

  gif.finish();
  return { buf: Buffer.from(gif.bytes()), frames: total,
           secs: (holdF + wipeF) * names.length / fps };
}

const all = ['piston', 'ticker', 'mark', 'belt', 'night'];

// ⭐ THE CHOSEN SETTING, and every number in it was measured rather than picked.
//   · 8 COLOURS — indistinguishable from 32 side by side, because the board is a few flat
//     tones and one soft gradient. It nearly halves the file for nothing.
//   · 8 fps — 5 was small but the ticker and the piston stepped visibly. 8 is the floor
//     where a scroll still reads as motion rather than as frames.
//   · 6s a pattern — 8s cost 175 KB more for a loop nobody watches end to end anyway.
const CFG = { fps: +(E.FPS||8), hold: +(E.HOLD||6), wipeSecs: +(E.WIPE||1.2),
              colours: +(E.COLOURS||8), names: E.NAMES ? E.NAMES.split(',') : all };
const r = build(CFG);
fs.writeFileSync(path.join(__dirname, 'renders', (E.OUT||'shuffle.gif')), r.buf);
console.log('  shuffle.gif  ' + (r.buf.length / 1024).toFixed(0) + ' KB  ' +
            r.frames + ' frames  ' + r.secs.toFixed(0) + 's loop  ' +
            CFG.names.length + ' patterns');
