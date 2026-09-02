/**
 * render.js — fonts in, PNG out. The only file that touches @napi-rs/canvas directly.
 *
 * ⚠ THE FONT LIVES IN THE REPO, NOT ON THE HOST. The VPS has no Oswald installed
 *   (fc-list -> 0 faces, verified 2026-08-31) and nobody will remember to install it after
 *   the next rebuild. Six static instances ship beside this file so the container is
 *   self-contained — which is the whole reason this is a container and not a host install.
 *
 * ⚠ STATIC INSTANCES, NOT THE VARIABLE FONT. Oswald[wght].ttf registers as a SINGLE 400
 *   face in @napi-rs/canvas and every weight measures identically — the design's 250-weight
 *   clock and 500-weight labels would have silently rendered the same. Measured, not
 *   assumed: 200..700 now come back 94.84 / 96.06 / 98.12 / 103.78 / 106.72 / 109.06 px on
 *   the same string.
 */
'use strict';

const path = require('path');
const { createCanvas, GlobalFonts } = require('@napi-rs/canvas');
const { drawDial, WIDTH, HEIGHT } = require('./draw');
const { buildState } = require('./state');
const { drawBoardFace, buildFigures, drawAnchorFigures } = require('./board-terminal');
const { encodeSettle, scramble } = require('./gif');

/**
 * ⭐ THE FACES. `dial` is the default and is BYTE-IDENTICAL to what shipped 2026-09-02 —
 *   an unknown or absent `face` renders it, so no existing URL can change behaviour and
 *   rollback is removing one query parameter.
 *
 * ⚠ Both faces are 280x121. A face of a different size would need MASTHEAD.dialW/dialH to
 *   change with it, and setupMasthead asserts those against the live column widths — so a
 *   new size is a sheet change, not just a drawing.
 */
const FACES = {
  dial:  (ctx, q, s) => drawDial(ctx, buildState(q), s),
  board: (ctx, q, s) => drawBoardFace(ctx, q, s)
};

const WEIGHTS = [200, 300, 400, 500, 600, 700];
let fontsReady = false;

/**
 * ⚠⚠ ONLY MULTIPLES OF 100. @napi-rs/canvas REPORTS the ExtraLight face as weight **250**
 *    (GlobalFonts.families says so) and then FAILS TO MATCH IT when you ask for 250 —
 *    ctx.font accepts the string, the matcher finds nothing, and every glyph renders as
 *    .notdef. It is silent: measureText("9:57") comes back 600px instead of 29px, and the
 *    PNG is four empty rectangles where the clock should be.
 *
 *    That shipped into the first render of this dial and was caught only by LOOKING at the
 *    picture — same shape as the ⛶ and 📻 tofu the Floor Board hit on the real tablet.
 *    Hence the probe below: a weight that cannot draw a digit must say so at boot.
 */
function assertFontWeights() {
  const { createCanvas } = require('@napi-rs/canvas');
  const ctx = createCanvas(8, 8).getContext('2d');
  const bad = [];
  for (const w of WEIGHTS) {
    ctx.font = `${w} 20px Oswald`;
    const px = ctx.measureText('0').width;
    // A 20px Oswald digit is ~10px. Anything past 25 is the .notdef box.
    if (!(px > 3 && px < 25)) bad.push(`${w} (measured ${px.toFixed(1)}px)`);
  }
  if (bad.length) {
    console.error('⚠⚠ Oswald weights that do NOT resolve — text will render as tofu: ' + bad.join(', '));
  }
  return bad;
}

function registerFonts() {
  if (fontsReady) return;
  const dir = path.join(__dirname, 'fonts');
  let ok = 0;
  for (const w of WEIGHTS) {
    const file = path.join(dir, `Oswald-${w}.ttf`);
    try { GlobalFonts.registerFromPath(file, 'Oswald'); ok++; }
    catch (e) { console.error('font register failed:', file, e.message); }
  }
  // ⚠ LOUD, NOT SILENT. A missing font falls back to a system sans and the dial still
  //   renders — it just stops being the design. Silent degradation is how a broken face
  //   ships looking merely "a bit off".
  if (ok !== WEIGHTS.length) {
    console.error(`⚠ Oswald: ${ok}/${WEIGHTS.length} weights registered — the dial will not match the design.`);
  }
  fontsReady = true;
  assertFontWeights();
}

/**
 * @param {object} query  raw ?s=&t=&o=... params
 * @param {number} [scale] 1 = the sheet's own 280x121. 2 = the same drawing, twice the px.
 * @returns {Buffer} PNG
 */
function renderPng(query, scale) {
  registerFonts();
  const s = scale || 1;
  const canvas = createCanvas(Math.round(WIDTH * s), Math.round(HEIGHT * s));
  const ctx = canvas.getContext('2d');
  const draw = FACES[(query && query.face) || 'dial'] || FACES.dial;
  draw(ctx, query || {}, s);
  return canvas.toBuffer('image/png');
}

/**
 * The board, animated: it settles once and holds.
 *
 * ⚠ ONLY the board face animates. The dial is a clock and an arc — a settle cascade means
 *   nothing there, and `=IMAGE()` cannot play a GIF anyway, so this is reachable only from
 *   the floating-image path.
 */
function renderGif(query, scale) {
  registerFonts();
  const s = scale || 1;
  const to = buildFigures(query || {});
  const from = {
    verdict: to.verdict,
    a: Object.assign({}, to.a, { value: scramble(to.a.value, 8, DRUM) }),
    b: Object.assign({}, to.b, { value: scramble(to.b.value, 5, DRUM) })
  };
  // The longest cell journey plus the deliberate stillness at the end. Capped so a slow
  // settle can never push the file past what a banner should cost.
  const dur = Math.min(6, Math.max(
    B_loop(from.a.value, to.a.value), B_loop(from.b.value, to.b.value)));
  return encodeSettle(
    (ctx, f, t2, t, sc) => drawAnchorFigures(ctx, f, t2, t, sc),
    from, to,
    { width: WIDTH, height: HEIGHT, scale: s, duration: dur });
}
const DRUM = require('./board').DIGITS;
function B_loop(a, b) { return require('./board').loopSeconds(a, b, DRUM); }

module.exports = { renderPng, renderGif, registerFonts, assertFontWeights,
                   WIDTH, HEIGHT, WEIGHTS, FACES };
