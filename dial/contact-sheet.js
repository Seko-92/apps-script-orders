#!/usr/bin/env node
/**
 * contact-sheet.js — every state at TRUE SIZE, on the real banner ground, in one image.
 *
 * ⚠ A 4x render flatters everything. The dial ships at 280x121 next to a black band, and
 *   that is the only size whose verdict counts. This sheet exists so the true size is the
 *   default thing anyone looks at, not an afterthought.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { WIDTH, HEIGHT, drawDial } = require('./src/draw');
const { buildState } = require('./src/state');

const DAY = [0,0,0,0,0,0,0,0,1,6,14,19,11,16,22,13,7,2,0,0,0,0,0,0];
const upTo = (hh) => DAY.map((n, i) => (i <= hh ? n : 0)).join(',');

const CASES = [
  ['rest · 9:57 PM',        { s:'rest',  t:'2157', p:'88', r:'91', u:'663', h:DAY.join(',') }],
  ['rest · 6:40 AM',        { s:'rest',  t:'0640', p:'88', r:'91', u:'140', h:DAY.join(',') }],
  ['clear · just opened',   { s:'clear', t:'0912', g:'0',  r:'2',  p:'0',  h:upTo(9)  }],
  ['clear · midday',        { s:'clear', t:'1305', g:'0',  r:'21', p:'21', h:upTo(13) }],
  ['busy · morning',        { s:'busy',  t:'1042', g:'7',  r:'9',  p:'4',  o:'64',  h:upTo(10) }],
  ['busy · afternoon',      { s:'busy',  t:'1414', g:'12', r:'14', p:'14', o:'96',  h:upTo(14) }],
  ['late · past the line',  { s:'late',  t:'1414', g:'12', r:'14', p:'14', o:'192', l:'3', h:upTo(14) }],
  ['late · bad day',        { s:'late',  t:'1630', g:'19', r:'21', p:'9',  o:'412', l:'7', h:upTo(16) }],
  ['stale · no signal',     { s:'stale', t:'1120', g:'5',  p:'3',  y:'73',  h:upTo(11) }],
  ['missing params',        { s:'busy',  t:'1414' }]
];

const SCALE = Number(process.env.SCALE || 2);   // 2 = readable on a laptop, still 1:1 art
const BANNER = 560;                              // how much black band to show beside it
const PADX = 26, PADY = 30, LABEL = 22, GAPY = 20;
const W = (PADX * 2 + WIDTH + BANNER) * SCALE;
const H = (PADY * 2 + CASES.length * (HEIGHT + LABEL + GAPY)) * SCALE;

registerFonts();
const canvas = createCanvas(W, H);
const ctx = canvas.getContext('2d');
ctx.fillStyle = '#0b0b0b'; ctx.fillRect(0, 0, W, H);

let y = PADY;
for (const [label, q] of CASES) {
  ctx.fillStyle = '#7d776c';
  ctx.font = `500 ${10 * SCALE}px Oswald`;
  ctx.letterSpacing = `${1.4 * SCALE}px`;
  ctx.fillText(label.toUpperCase(), PADX * SCALE, (y + 11) * SCALE);
  ctx.letterSpacing = '0px';
  y += LABEL;

  // the banner row 1 the dial actually sits in
  ctx.fillStyle = '#1a1a1a';
  ctx.fillRect(PADX * SCALE, y * SCALE, (WIDTH + BANNER) * SCALE, HEIGHT * SCALE);

  // ⚠ DRAWN STRAIGHT IN, not decoded from a PNG. `new Image(); img.src = buffer` resolves
  //   ASYNCHRONOUSLY in @napi-rs/canvas, so the first version of this sheet composited ten
  //   empty rectangles and looked like the renderer was broken. It was the harness.
  ctx.save();
  ctx.translate(PADX * SCALE, y * SCALE);
  drawDial(ctx, buildState(q), SCALE);
  ctx.restore();
  y += HEIGHT + GAPY;
}

const out = path.join(__dirname, 'renders', 'contact-sheet.png');
fs.writeFileSync(out, canvas.toBuffer('image/png'));
console.log(`${CASES.length} states at ${SCALE}x preview (art is 1:1 at ${WIDTH}x${HEIGHT}) -> ${out}`);
