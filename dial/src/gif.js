/**
 * gif.js — the board, animated. Frames in, one GIF out.
 *
 * ⚠⚠ A FLOATING GIF LOOPS FOREVER BY DEFAULT, and that is fatal to this design. A Solari
 *    board's whole character is STILLNESS punctuated by a settle; a cascade replaying every
 *    four seconds is a screensaver, and it would be the first thing anyone asks to turn off.
 *
 *    So the GIF is written with **repeat = -1 — play ONCE and hold the final frame.** The
 *    board settles when it is hung, then stands still until it is replaced. That is exactly
 *    how a real board behaves and it needs no trigger, no polling and no second image.
 *
 * ⭐ AND THE ANIMATION IS A FULL SETTLE FROM SCRAMBLED, not a diff from the previous value.
 *   Two reasons, and the second is the important one:
 *     1 · A one-digit change (30 -> 31) moves a single flap, which is not worth watching.
 *         A real board re-syncs the whole row, and that cascade IS the thing.
 *     2 · ⚠⚠ IT KEEPS THE SERVER STATELESS. A diff would need the PREVIOUS value, which
 *         means either the server remembers (forfeiting the entire reason /dial can be an
 *         open route — a caller can only render numbers they supplied) or the caller passes
 *         it. Settling from scrambled needs neither. The URL still carries only what the
 *         face draws.
 */
'use strict';

const { GIFEncoder, quantize, applyPalette } = require('gifenc');
const { createCanvas } = require('@napi-rs/canvas');
const B = require('./board');

const FPS   = 12;
const DELAY = Math.round(1000 / FPS);

/**
 * Where each cell STARTS so the row settles rather than jumps.
 * ⚠ Distances differ per cell on purpose — equal spins land every flap on the same frame,
 *   which reads as a slot machine. Uneven travel is what makes it look mechanical.
 */
function scramble(to, spin, abc) {
  const A = abc || B.ALPHABET;
  let out = '';
  for (let i = 0; i < to.length; i++) {
    const back = (spin || 7) + (i % 4) * 3;
    out += A[(B.idxOf(to[i], A) - back + A.length * 4) % A.length];
  }
  return out;
}

/**
 * @param draw  (ctx, from, to, t, scale) — any board face
 * @param from  scrambled start · @param to  the settled value
 * @returns {Buffer} GIF
 */
function encodeSettle(draw, from, to, opts) {
  const o = opts || {};
  const w = o.width, h = o.height, scale = o.scale || 1;
  const dur = o.duration;
  const frames = Math.max(2, Math.ceil(dur * FPS));
  const gif = GIFEncoder();
  const cv = createCanvas(Math.round(w * scale), Math.round(h * scale));
  const ctx = cv.getContext('2d');

  for (let f = 0; f < frames; f++) {
    const t = (f / (frames - 1)) * dur;
    draw(ctx, from, to, t, scale);
    const { data } = ctx.getImageData(0, 0, cv.width, cv.height);
    // ⚠ The palette is built from the FIRST frame and reused. The board is a fixed set of
    //   flat tones — housing, two tile gradients, cream, one yellow — so a per-frame
    //   palette would cost time and could make identical pixels dither differently between
    //   frames, which shows up as a crawling texture on a still board.
    if (f === 0) gif.__pal = quantize(data, 64, { format: 'rgb565' });
    const idx = applyPalette(data, gif.__pal, 'rgb565');
    gif.writeFrame(idx, cv.width, cv.height, {
      palette: f === 0 ? gif.__pal : undefined,
      delay: DELAY,
      // ⚠ repeat is written on the FIRST frame only. -1 = play once and stop.
      repeat: f === 0 ? -1 : undefined
    });
  }
  gif.finish();
  return Buffer.from(gif.bytes());
}

module.exports = { encodeSettle, scramble, FPS };
