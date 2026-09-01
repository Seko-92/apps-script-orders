#!/usr/bin/env node
/**
 * minute-pair.js — the honest counterweight to shift.gif.
 *
 * ⚠ THE FILM COMPRESSES 9 HOURS INTO 14 SECONDS, which flatters it. This renders two
 *   frames ONE REAL MINUTE APART, at true size, so the other half of the truth is on the
 *   page too: minute to minute, almost nothing moves. The banner is a CLOCK, not an
 *   animation — it reads as alive because you catch it in different positions, the same
 *   way a wall clock does.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { drawDial, WIDTH, HEIGHT } = require('./src/draw');
const { buildState } = require('./src/state');

const day = [0,0,0,0,0,0,0,0,1,6,14,19,11,16,22,13,0,0,0,0,0,0,0,0].join(',');
const at = (t, o) => ({ s:'busy', t, o:String(o), g:'12', r:'31', p:'19', h:day });
const PAIRS = [
  ['2:14 PM', at('1414', 96)],
  ['2:15 PM · one minute later', at('1415', 97)],
  ['2:44 PM · thirty minutes later', at('1444', 126)],
  ['3:14 PM · one hour later', at('1514', 156)]
];

const S = 2, PAD = 18, LAB = 20, GAP = 16;
registerFonts();
const canvas = createCanvas((WIDTH + PAD * 2) * S,
                            (PAD * 2 + PAIRS.length * (LAB + HEIGHT + GAP)) * S);
const ctx = canvas.getContext('2d');
ctx.fillStyle = '#0b0b0b'; ctx.fillRect(0, 0, canvas.width, canvas.height);
let y = PAD;
for (const [label, q] of PAIRS) {
  ctx.fillStyle = '#7d776c'; ctx.font = `500 ${9.5 * S}px Oswald`;
  ctx.letterSpacing = `${1.3 * S}px`;
  ctx.fillText(label.toUpperCase(), PAD * S, (y + 10) * S);
  ctx.letterSpacing = '0px';
  y += LAB;
  ctx.save(); ctx.translate(PAD * S, y * S); drawDial(ctx, buildState(q), S); ctx.restore();
  y += HEIGHT + GAP;
}
const out = path.join(__dirname, 'renders', 'minute-pair@2x.png');
fs.writeFileSync(out, canvas.toBuffer('image/png'));
console.log('-> ' + out);
