/**
 * board-matrix.js — THE MATRIX. The plan's second face: a flip-DISC grid.
 *
 * A flip-dot display is a field of discs, each black on one side and pale on the other,
 * flipped by a magnet. Its signature is that EVERYTHING is made of the same discs — letters,
 * numbers and the day curve alike — which is the entire argument for this face: it can draw
 * the curve in the same material as the text, where the Terminal can only put a number
 * beside it.
 *
 * ⚠ TEXT IS SAMPLED, NOT HAND-AUTHORED. The composition is drawn once into an offscreen
 *   canvas at GRID resolution and its alpha is thresholded into discs. That gives real
 *   Oswald letterforms for ~10 lines of code, and it means the type scale is chosen in
 *   grid units — the honest constraint of the medium, rather than a bitmap font to maintain.
 *
 * ⚠ FIXED LIGHT, same rule as the flaps: an ON disc catches light from above and casts a
 *   hairline shadow below. There is no sweep.
 */
'use strict';
const { createCanvas } = require('@napi-rs/canvas');
const B = require('./board');

const PITCH = 4.0;          // px between disc centres — the plan's number
const R     = 1.62;         // disc radius

const ON_HI  = '#f4f0e4', ON_LO = '#cdc7b6';
const OFF_HI = '#26221c', OFF_LO = '#15120f';

/**
 * @param compose  (c, cols, rows) — draw the composition at GRID resolution
 * @returns a cols x rows boolean field
 */
function sample(w, h, compose) {
  const cols = Math.floor(w / PITCH), rows = Math.floor(h / PITCH);
  const off = createCanvas(cols, rows);
  const c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows);
  compose(c, cols, rows);
  const px = c.getImageData(0, 0, cols, rows).data;
  const field = [];
  for (let y = 0; y < rows; y++) {
    const row = [];
    // ⚠ Threshold on ALPHA, not luminance. The composition is drawn in a single opaque ink
    //   on a transparent ground, so alpha is the shape and luminance would also pick up
    //   antialiasing as half-lit discs — which reads as a blurry display, not a crisp one.
    for (let x = 0; x < cols; x++) row.push(px[(y * cols + x) * 4 + 3] > 110);
    field.push(row);
  }
  return field;
}

function drawMatrix(ctx, o) {
  const s = o.scale || 1;
  const S = (v) => v * s;
  // A caller may hand in a ready-made field (the ambient loop computes its own per frame)
  // or a compose function to sample. Sampling text every frame would be wasteful and, worse,
  // would re-threshold antialiasing identically each time for no gain.
  const field = o.field || sample(o.w, o.h, o.compose);
  const rows = field.length, cols = field[0].length;
  const x0 = (o.w - cols * PITCH) / 2 + PITCH / 2;
  const y0 = (o.h - rows * PITCH) / 2 + PITCH / 2;

  for (let y = 0; y < rows; y++) {
    for (let x = 0; x < cols; x++) {
      const on = field[y][x];
      const cx = S(x0 + x * PITCH), cy = S(y0 + y * PITCH), r = S(R);
      // the hairline shadow an ON disc throws — this is most of what makes the grid
      // read as physical objects rather than printed dots
      if (on) {
        ctx.beginPath();
        ctx.arc(cx, cy + S(0.5), r, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(0,0,0,0.55)';
        ctx.fill();
      }
      const g = ctx.createLinearGradient(0, cy - r, 0, cy + r);
      g.addColorStop(0, on ? ON_HI : OFF_HI);
      g.addColorStop(1, on ? ON_LO : OFF_LO);
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.fillStyle = g;
      ctx.fill();
    }
  }
  return { cols, rows };
}

module.exports = { drawMatrix, sample, PITCH, R };
