/**
 * board-anatomy.js — ONE CELL, at 8x, across a whole flip.
 *
 * ⚠ The contact sheet judges the CASCADE; it cannot judge the MECHANIC — at true size a
 *   flap is 27px and a wrong hinge looks the same as a right one. This is the only render
 *   that can show whether the card actually rotates about the seam, whether the shadow
 *   travels, and whether the settle bounces. The plan's move 5 ("never a rectangle
 *   squashing in place") is only checkable here.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const B = require('./src/board');

registerFonts();
const Z = 8, W = 27, H = 40, PAD = 10;
const phases = [0, 0.12, 0.25, 0.38, 0.5, 0.62, 0.75, 0.88, 1];

const cv = createCanvas((W * Z + PAD) * phases.length + PAD, H * Z + PAD * 2 + 26);
const ctx = cv.getContext('2d');
ctx.fillStyle = '#2b2b2b'; ctx.fillRect(0, 0, cv.width, cv.height);

phases.forEach((ph, i) => {
  const x = PAD + i * (W * Z + PAD);
  ctx.save();
  ctx.translate(x, PAD);
  B.drawFlapCell(ctx, {
    scale: Z, x: 0, y: 0, w: W, h: H, fontPx: 25, weight: 500,
    cur: 'R', next: 'S', phase: ph, moving: ph > 0 && ph < 1, baseline: 0.5, radius: 2
  });
  ctx.restore();
  ctx.font = '600 13px Oswald';
  ctx.fillStyle = '#ffd400';
  ctx.textBaseline = 'top';
  ctx.fillText(ph === 0 ? 'settled' : ph === 1 ? 'landed' : 'φ ' + ph.toFixed(2),
               x, PAD + H * Z + 6);
});

fs.writeFileSync(path.join(__dirname, 'renders', 'board-anatomy.png'), cv.toBuffer('image/png'));
console.log('renders/board-anatomy.png — R -> S at 8x');
