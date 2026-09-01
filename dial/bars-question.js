#!/usr/bin/env node
/**
 * bars-question.js — two questions about F1:H1, drawn side by side.
 *
 * 1 · WHAT THE BARS COUNT. __SparkData!A1:X1 is COUNTIFS over the Activity Log's TIMESTAMP
 *     column with NO event filter, and StatusService writes "one event per row". So one
 *     12-line order flipping PENDING -> PREPARING -> SHIPPED writes ~36 rows, while five
 *     separate 1-line orders write 15. The curve is dominated by ORDER SIZE and status
 *     churn, not by throughput. It is labelled "the day"; it is measuring how much the log
 *     wrote. Adding an event criterion makes it what everyone already reads it as.
 *
 * 2 · WHO OWNS THE DAY'S SHAPE. The dial's pips and the curve both answer "when did the
 *     day happen". The curve can carry MAGNITUDE; a radial tick at 103px cannot. So if one
 *     of them goes, it is the pips — and the dial gets its clutter back as breathing room.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { drawDial, WIDTH, HEIGHT } = require('./src/draw');
const { buildState } = require('./src/state');

// The SAME day, counted two ways.
//   logRows : every Activity Log row — RECEIVED + PREPARING + SHIPPED + NOTE + PRINTED...
//   shipped : SHIPPED events only — lines that actually left the building
// The 11am spike is one 14-line order being worked; it dwarfs the 3pm hour that shipped
// more. That inversion is the whole point.
const HOURS   = [8,9,10,11,12,13,14,15,16,17];
const LOGROWS = [2,9,17,46,21,14,19,23,11,4];
const SHIPPED = [0,1,4, 6,  5, 4, 7,12, 6,2];

const S = 2, BAND = '#1a1a1a', YEL = '#ffd400', LAB = '#8a8f98';
registerFonts();

function curve(ctx, x, y, w, h, vals, colour, title, sub) {
  ctx.fillStyle = BAND; ctx.fillRect(x * S, y * S, w * S, h * S);
  const max = Math.max(1, ...vals), bw = (w - 30) / vals.length;
  ctx.fillStyle = colour;
  vals.forEach((v, i) => {
    const bh = v > 0 ? Math.max(2, (v / max) * (h - 34)) : 0;
    if (bh) ctx.fillRect((x + 15 + i * bw) * S, (y + h - 12 - bh) * S, (bw - 3) * S, bh * S);
  });
  ctx.fillStyle = LAB; ctx.font = `500 ${8 * S}px Oswald`; ctx.letterSpacing = `${1.3 * S}px`;
  ctx.fillText(title, (x + 15) * S, (y + 15) * S);
  ctx.letterSpacing = '0px';
  ctx.font = `400 ${8 * S}px Oswald`; ctx.fillStyle = '#6b6genuine'.slice(0,7);
  ctx.fillStyle = '#6b665e';
  ctx.fillText(sub, (x + 15) * S, (y + h - 3) * S);
}

const day = [0,0,0,0,0,0,0,0,2,9,17,46,21,14,19,23,11,4,0,0,0,0,0,0].join(',');
const busy = (np) => ({ s:'busy', t:'1530', g:'8', r:'34', p:'27', o:'118', h:day, np });
const rest = (np) => ({ s:'rest', t:'2140', p:'41', r:'44', u:'680', h:day, np });

const ROW = HEIGHT, GAPY = 26, PADX = 22, PADY = 22, LABH = 19;
const CW = 337;                                  // F1:H1's real width
const blocks = [
  ['A · the dial keeps its pips',       busy(undefined), rest(undefined)],
  ['B · the curve owns the day',        busy('1'),       rest('1')]
];
const W = (PADX * 2 + WIDTH + 24 + CW) * S;
const H = (PADY * 2 + 2 * (LABH + ROW * 2 + 10 + GAPY) + LABH + 96) * S;
const canvas = createCanvas(W, H);
const ctx = canvas.getContext('2d');
ctx.fillStyle = '#0b0b0b'; ctx.fillRect(0, 0, W, H);

let y = PADY;
const label = (t, yy) => {
  ctx.fillStyle = '#7d776c'; ctx.font = `500 ${10 * S}px Oswald`;
  ctx.letterSpacing = `${1.5 * S}px`;
  ctx.fillText(t.toUpperCase(), PADX * S, (yy + 11) * S);
  ctx.letterSpacing = '0px';
};

for (const [title, qBusy, qRest] of blocks) {
  label(title, y); y += LABH;
  for (const q of [qBusy, qRest]) {
    ctx.save(); ctx.translate(PADX * S, y * S); drawDial(ctx, buildState(q), S); ctx.restore();
    curve(ctx, PADX + WIDTH + 24, y, CW, HEIGHT,
          q.s === 'rest' ? SHIPPED : SHIPPED, q.s === 'rest' ? '#7e8894' : YEL,
          q.s === 'rest' ? 'YESTERDAY' : 'TODAY', q.s === 'rest' ? '41 out' : '27 out');
    y += ROW + 5;
  }
  y += GAPY;
}

label('What the bars are counting', y); y += LABH;
curve(ctx, PADX, y, CW, 96, LOGROWS, '#c0553a', 'AS BUILT · ACTIVITY LOG ROWS',
      'the 11am spike is ONE 14-line order being worked');
curve(ctx, PADX + CW + 24, y, CW, 96, SHIPPED, YEL, 'ONE CRITERION ADDED · SHIPPED ONLY',
      '3pm is the hour that actually shipped most');

const out = path.join(__dirname, 'renders', 'bars@2x.png');
fs.writeFileSync(out, canvas.toBuffer('image/png'));
console.log('-> ' + out);
