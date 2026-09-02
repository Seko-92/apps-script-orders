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
const { execFileSync } = require('child_process');
const { loadImage } = require('@napi-rs/canvas');
const A = require('./src/ambient');
const P = require('./src/patterns');
const M = require('./src/board-matrix');
registerFonts();

const E = process.env;
const W = +(E.W||280), H = +(E.H||121);
const { cols, rows } = A.gridSize(W, H);
/**
 * ⭐⭐ THE REAL MARK, NOT A TYPESET STAND-IN. The brand moment used to be the letters "HQ" set
 *    in Oswald — the wordmark, not the logo. This rasterises fav-google.svg and thresholds it
 *    into the disc field, so the board shows the actual roundel: double ring, H|Q, and the
 *    rule between them.
 *
 * ⭐ RESPONSIVE LOCKUP, because one composition cannot serve 2:1 and 16:1. The block takes the
 *   roundel alone at nearly full height, where H|Q is genuinely legible. The strip takes the
 *   roundel plus MOTOR SERVICE beside it — the logo's own structure, mark then descriptor. The
 *   roundel already says HQ, so repeating it would spend letters for nothing, and squashing
 *   the 3:1 lockup into 16:1 was tried and destroys both halves.
 *
 * ⚠ ALPHA DECIDES THE SHAPE, HUE DECIDES THE DISC. The art is black shapes AND yellow shapes
 *   on transparency, so alpha gives their union; a luminance threshold drops the yellow arc
 *   entirely. Hue is sampled per DISC, not per source pixel — a disc is one object and can
 *   only be one colour.
 * ⚠ ACCENT IS THE LOGO'S OWN #ffdc00, warmer than BRAND.yellow #ffd400. The mark wears its
 *   own colour, not the sheet's approximation of it.
 */
async function brandMark(cols, rows, colour) {
  const tmp = path.join(__dirname, 'renders', '_mark.png');
  execFileSync('rsvg-convert', ['-w', '900', '-o', tmp,
                                path.join(__dirname, 'logo', 'fav-google.svg')]);
  const img = await loadImage(tmp);
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows);
  const d = rows * (cols / rows > 6 ? 0.98 : 0.94);
  if (cols / rows > 6) {
    const text = 'MOTOR SERVICE', gap = rows * 0.30;
    let size = Math.round(rows * 0.95);
    while (size > 5) { c.font = '600 ' + size + 'px Oswald';
      if (d + gap + c.measureText(text).width <= cols * 0.92) break; size--; }
    const x0 = (cols - (d + gap + c.measureText(text).width)) / 2;
    c.drawImage(img, x0, (rows - d) / 2, d, d);
    c.fillStyle = '#fff'; c.textBaseline = 'middle'; c.textAlign = 'left';
    c.fillText(text, x0 + d + gap, rows / 2);
  } else {
    c.drawImage(img, (cols - d) / 2, (rows - d) / 2, d, d);
  }
  const px = c.getImageData(0, 0, cols, rows).data, f = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) {
      const i = (y * cols + x) * 4;
      if (px[i + 3] <= 90) { r.push(0); continue; }
      if (!colour) { r.push(1); continue; }
      r.push((px[i] > 150 && px[i + 1] > 110 && px[i + 2] < 120) ? 2 : 1);
    }
    f.push(r);
  }
  try { fs.unlinkSync(tmp); } catch (e) {}
  return f;
}
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
  inlinefour:(ph) => fieldFromDraw((c) => P.inlineFour(c, cols, rows, ph)),
  wave:   (ph) => fieldFromDraw((c) => P.wave(c, cols, rows, ph)),
  aisle:  (ph) => fieldFromDraw((c) => P.aisle(c, cols, rows, ph)),
  sweep:  (ph) => fieldFromDraw((c) => P.sweep(c, cols, rows, ph)),
  pendulum:(ph) => fieldFromDraw((c) => P.pendulum(c, cols, rows, ph)),
  moire:  (ph) => P.moire(cols, rows, ph),
  liquid: (ph) => P.liquid(cols, rows, ph),
  ripple: (ph) => P.ripple(cols, rows, ph),
  ticker: (ph) => fieldFromDraw((c) => P.ticker(c, cols, rows, ph, null, TSTRIP)),
  mark:   (ph) => A.composeAmbient('mark', ph, cols, rows, mark),
  belt:   (ph) => fieldFromDraw((c) => P.belt(c, cols, rows, ph)),
  night:  (ph) => fieldFromDraw((c) => P.night(c, cols, rows, ph))
};

/**
 * ⭐⭐ THE BLOOM. b grows out of the centre over a, on an ELLIPSE matched to the canvas aspect —
 *   so one call reads as a circle opening on the 2:1 block and a band spreading outward on the
 *   16:1 strip, instead of a circle that would spend the whole transition reaching the ends.
 * ⚠ A HARD EDGE ON PURPOSE. The wipe below softens its front with churn, and churn is exactly
 *   what cost `refresh` +278 KB in a set. The disc grid quantises this edge anyway.
 */
function bloom(a, b, t) {
  const cx = (cols - 1) / 2, cy = (rows - 1) / 2, out = [];
  for (let y = 0; y < rows; y++) {
    const ny = cy ? (y - cy) / cy : 0, row = [];
    for (let x = 0; x < cols; x++) {
      const nx = cx ? (x - cx) / cx : 0;
      row.push(Math.sqrt(nx * nx + ny * ny) / Math.SQRT2 <= t ? b[y][x] : a[y][x]);
    }
    out.push(row);
  }
  return out;
}

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

function build(opts, mark) {
  const { fps, hold, wipeSecs, colours, names, convSecs, markSecs } = opts;
  const holdF = Math.round(hold * fps), wipeF = Math.round(wipeSecs * fps);
  const convF = Math.round(convSecs * fps), markHoldF = Math.round(markSecs * fps);
  const gif = GIFEncoder();
  const cv = createCanvas(W, H), ctx = cv.getContext('2d');
  let pal, total = 0;

  /** Draw a field and hand back its pixels — used both to emit frames and to build the palette. */
  const paint = (field) => {
    const g = ctx.createLinearGradient(0, 0, 0, H);
    g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
    ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
    M.drawMatrix(ctx, { scale: 1, w: W, h: H, field: field });
    const seam = ctx.createLinearGradient(W - 26, 0, W, 0);
    seam.addColorStop(0, 'rgba(26,26,26,0)'); seam.addColorStop(1, '#1a1a1a');
    ctx.fillStyle = seam; ctx.fillRect(W - 26, 0, 26, H);
    return ctx.getImageData(0, 0, W, H).data;
  };

  // ⚠⚠ THE PALETTE MUST BE SAMPLED FROM A FRAME THAT CONTAINS THE MARK. gifenc quantises ONE
  //    sample and every later frame is mapped into it — and frame 0 is a pattern, which has no
  //    yellow in it at all. So the accent was being snapped to the nearest pale tone for the
  //    whole loop and the brand colour silently never appeared, at ANY palette size. Raising
  //    the colour count could not fix it; sampling the right pixels is the fix.
  //    Caught by extracting frame 60 and looking — the encoder reported nothing wrong.
  {
    const d0 = paint(PATTERNS[names[0]](0));
    if (mark) {
      const dm = paint(mark);
      const both = new Uint8Array(d0.length + dm.length);
      both.set(d0, 0); both.set(dm, d0.length);
      pal = quantize(both, colours, { format: 'rgb565' });
    } else {
      pal = quantize(d0, colours, { format: 'rgb565' });
    }
  }

  const frame = (field) => {
    const g = ctx.createLinearGradient(0, 0, 0, H);
    g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
    ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
    M.drawMatrix(ctx, { scale: 1, w: W, h: H, field });
    const seam = ctx.createLinearGradient(W - 26, 0, W, 0);
    seam.addColorStop(0, 'rgba(26,26,26,0)'); seam.addColorStop(1, '#1a1a1a');
    ctx.fillStyle = seam; ctx.fillRect(W - 26, 0, 26, H);
    const data = ctx.getImageData(0, 0, W, H).data;
    gif.writeFrame(applyPalette(data, pal, 'rgb565'), W, H, {
      palette: total === 0 ? pal : undefined,
      delay: Math.round(1000 / fps),
      repeat: total === 0 ? 0 : undefined
    });
    total++;
  };

  // ⭐⭐ THE MARK IS THE PUNCTUATION, NOT A PATTERN. Its own slot in the rotation shows it once
  //    per loop and costs a slot; BETWEEN every pattern it shows once per pattern and costs
  //    none — so the brand is what the board keeps returning to, not one more thing it cycles
  //    past. markSecs = 0 restores the plain wipe exactly.
  names.forEach((name, i) => {
    const fn = PATTERNS[name];
    for (let f = 0; f < holdF; f++) frame(fn(f / holdF));
    const next = PATTERNS[names[(i + 1) % names.length]];
    const a = fn(1), b = next(0);
    if (markHoldF > 0 && mark) {
      for (let f = 0; f < convF; f++) frame(bloom(a, mark, f / convF));
      for (let f = 0; f < markHoldF; f++) frame(mark);
      for (let f = 0; f < convF; f++) frame(bloom(mark, b, f / convF));
    } else {
      for (let f = 0; f < wipeF; f++) frame(wipe(a, b, f / wipeF));
    }
  });

  gif.finish();
  return { buf: Buffer.from(gif.bytes()), frames: total,
           secs: (holdF + (markHoldF > 0 ? convF * 2 + markHoldF : wipeF)) * names.length / fps };
}

const all = ['piston', 'ticker', 'mark', 'belt', 'night'];

// ⭐ THE CHOSEN SETTING, and every number in it was measured rather than picked.
//   · 8 COLOURS — indistinguishable from 32 side by side, because the board is a few flat
//     tones and one soft gradient. It nearly halves the file for nothing.
//   · 8 fps — 5 was small but the ticker and the piston stepped visibly. 8 is the floor
//     where a scroll still reads as motion rather than as frames.
//   · 6s a pattern — 8s cost 175 KB more for a loop nobody watches end to end anyway.
const CFG = { fps: +(E.FPS||8), hold: +(E.HOLD||6), wipeSecs: +(E.WIPE||1.2),
              convSecs: +(E.CONV||0.7), markSecs: +(E.MARK||0),
              colours: +(E.COLOURS||8), names: E.NAMES ? E.NAMES.split(',') : all };
(async () => {
  const mk = CFG.markSecs > 0 ? await brandMark(cols, rows, E.MONO !== '1') : null;
  const r = build(CFG, mk);
  fs.writeFileSync(path.join(__dirname, 'renders', (E.OUT||'shuffle.gif')), r.buf);
  console.log('  ' + (E.OUT||'shuffle.gif') + '  ' + (r.buf.length / 1024).toFixed(0) + ' KB  ' +
              r.frames + ' frames  ' + r.secs.toFixed(0) + 's loop  ' +
              CFG.names.length + ' patterns' + (mk ? '  + the mark' : ''));
})();
