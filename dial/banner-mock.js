#!/usr/bin/env node
/**
 * banner-mock.js — the dial IN THE BANNER, at true column widths.
 *
 * ⚠ A dial judged on its own is a dial judged in the wrong place. It ships merged into
 *   A1:C2 with the headline, the pulse and the day curve beside it and the eBay label
 *   underneath, and the only question that matters is whether those five things read as
 *   ONE banner. So this draws the whole row at the widths applyBrandTheme actually pins:
 *
 *     A 107 · B 70 · C 103 | D 232 | E 307 | F 130 · G 100 · H 107   = 1156
 *     row 1  56px   row 2  65px    (the dial merges A1:C2 = 280 x 121)
 *
 * Everything except the dial is an APPROXIMATION of the live cells — the sparkline, the
 * pulse lamp and the eBay wordmark are drawn here, not read from Sheets. The dial is the
 * real renderer.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const { drawDial, WIDTH, HEIGHT, fmtMins, fmtClock } = require('./src/draw');
const { buildState } = require('./src/state');

const COL = { A:107, B:70, C:103, D:232, E:307, F:130, G:100, H:107 };
const DIALW = COL.A + COL.B + COL.C;            // 280
const TOTAL = Object.values(COL).reduce((a, b) => a + b, 0);   // 1156
const R1 = 56, R2 = 65;
const BAND = '#1a1a1a', CREAM = '#fff8e7', QUIET = '#e8e8e8', YEL = '#ffd400';
const REST_ACCENT = '#7e8894';

const DAY = [0,0,0,0,0,0,0,0,1,6,14,19,11,16,22,13,7,2,0,0,0,0,0,0];
const upTo = (hh) => DAY.map((n, i) => (i <= hh ? n : 0));

/** D1 — mirrors the cellStats formula in _setSystemPulseBannerFormulas, branch for branch. */
function headlineA(q) {
  const g = Number(q.g || 0), r = Number(q.r || 0), p = Number(q.p || 0);
  switch (q.s) {
    case 'rest':  return 'opens in ' + fmtMins(Number(q.u)) + (g > 0 ? ' · ' + g + ' waiting' : '');
    case 'late':  return ['oldest ' + fmtMins(Number(q.o)), g + ' still waiting'];
    case 'stale': return 'last seen ' + fmtMins(Number(q.y));
    case 'busy':  return [g + ' to grab', r + ' in · ' + p + ' out'];
    default:      return r + ' in · ' + p + ' out';
  }
}

/**
 * OPTION B — D1 stops repeating the dial.
 *
 * ⚠⚠ THE DIAL TAKES OVER D1's JOB and the artifact could not show it. Drawn at true width
 *    with real content, BUSY reads "12 to grab / 14 out today" on the dial and
 *    "12 to grab / 14 in · 14 out" in D1, 200px apart. STALE says the same duration THREE
 *    times — dial, D1 and the pulse. D1's headline was written to complement a face that
 *    only ever said a state WORD; it was never meant to sit beside an instrument.
 *
 * So B gives D1 the one thing nothing else in the banner shows: WHICH TABLE the work is
 * in. __SparkData already holds A17 ebayPending and A18 directPending, and no surface in
 * row 1 has ever carried the split.
 */
function headlineB(q) {
  const eb = Number(q.eb || 0), di = Number(q.di || 0);
  const split = 'eBay ' + eb + ' · Direct ' + di;
  switch (q.s) {
    case 'rest':  return ['the floor is asleep', 'yesterday · ' + q.p + ' out'];
    case 'late':  return [Number(q.l || 0) + ' past the 3h line', split];
    case 'stale': return ['pipeline quiet', 'nothing logged since ' + clock12(q.ls || q.t)];
    case 'busy':  return ['picking', split];
    default:      return ['all caught up', split];
  }
}
const headline = (q) => (process.env.VARIANT === 'B' ? headlineB(q) : headlineA(q));
/** E1 — the pulse. Keeps its "h:mm AM/PM" substring; ActivityLog.js regex-parses it. */
function pulse(q) {
  if (q.s === 'rest')  return ['#9aa3ad', 'RESTING · ' + clock12(q.t) + ' · ' + fmtMins(Number(q.y || 8)) + ' ago'];
  if (q.s === 'stale') return ['#ff6b6b', 'STALE · '   + clock12(q.t) + ' · ' + fmtMins(Number(q.y || 73)) + ' ago'];
  return ['#7ec98a', 'ALIVE · ' + clock12(q.t) + ' · ' + fmtMins(Number(q.y || 3)) + ' ago'];
}
function clock12(t) {
  const m = (Number(String(t).slice(0, -2)) * 60) + Number(String(t).slice(-2));
  return fmtClock(m) + (m < 720 ? ' AM' : ' PM');
}

function drawBanner(ctx, q, S) {
  const s = (v) => v * S;
  ctx.fillStyle = BAND; ctx.fillRect(0, 0, s(TOTAL), s(R1));
  ctx.fillStyle = CREAM; ctx.fillRect(s(DIALW), s(R1), s(TOTAL - DIALW), s(R2));

  // the dial, merged across both rows
  ctx.save(); ctx.translate(0, 0); drawDial(ctx, buildState(q), S); ctx.restore();

  // D1 — the headline
  const hl = headline(q);
  const lines = Array.isArray(hl) ? hl : [hl];
  ctx.fillStyle = QUIET; ctx.font = `400 ${s(12.5)}px Oswald`;
  lines.forEach((ln, i) => ctx.fillText(ln, s(DIALW + 16), s(lines.length === 1 ? 33 : 24 + i * 16)));

  // E1 — the pulse lamp + line
  const [lamp, txt] = pulse(q);
  const ex = DIALW + COL.D;
  ctx.fillStyle = lamp;
  ctx.beginPath(); ctx.arc(s(ex + 14), s(28), s(4.5), 0, Math.PI * 2); ctx.fill();
  ctx.fillStyle = QUIET; ctx.font = `400 ${s(11.5)}px Oswald`;
  ctx.fillText(txt, s(ex + 25), s(32));

  // F1:H1 — the day curve (SPARKLINE column chart)
  const fx = ex + COL.E, fw = COL.F + COL.G + COL.H;
  const vals = q.s === 'rest' ? DAY : upTo(Math.floor(Number(String(q.t).slice(0, -2))));
  const max = Math.max(1, ...vals);
  const bw = (fw - 28) / 24;
  ctx.fillStyle = q.s === 'rest' ? REST_ACCENT : YEL;
  vals.forEach((v, i) => {
    const h = Math.max(v > 0 ? 1.5 : 0, (v / max) * 36);
    if (h > 0) ctx.fillRect(s(fx + 14 + i * bw), s(48 - h), s(bw - 1.6), s(h));
  });
  ctx.fillStyle = '#8a8f98'; ctx.font = `500 ${s(8)}px Oswald`;
  ctx.letterSpacing = `${s(1.3)}px`;
  ctx.fillText((q.s === 'rest' ? 'YESTERDAY' : 'TODAY'), s(fx + 14), s(15));
  ctx.letterSpacing = '0px';

  // row 2 — the eBay label, then the two pick-ID cells
  ctx.fillStyle = '#111'; ctx.font = `700 ${s(26)}px "Noto Sans"`;
  const ebay = [['e','#e53238'],['b','#0064d2'],['a','#f5af02'],['y','#86b817']];
  let lx = DIALW + 18;
  for (const [ch, col] of ebay) { ctx.fillStyle = col; ctx.fillText(ch, s(lx), s(R1 + 44)); lx += ctx.measureText(ch).width / S; }

  const px = DIALW + COL.D + COL.E;
  const cells = [[px, COL.F + COL.G, 'PICK ID · SHIPPING', 'Shipping - Yassin 1'],
                 [px + COL.F + COL.G, COL.H, 'ADJUSTMENT', 'AShamma 2']];
  for (const [x, w, lab, val] of cells) {
    ctx.strokeStyle = '#e8dfc8'; ctx.lineWidth = Math.max(1, s(1));
    ctx.beginPath(); ctx.moveTo(s(x), s(R1)); ctx.lineTo(s(x), s(R1 + R2)); ctx.stroke();
    ctx.fillStyle = '#9a9280'; ctx.font = `600 ${s(7.5)}px Oswald`;
    ctx.letterSpacing = `${s(1.2)}px`;
    ctx.fillText(lab, s(x + 10), s(R1 + 22)); ctx.letterSpacing = '0px';
    ctx.fillStyle = '#1a1a1a'; ctx.font = `400 ${s(11)}px Oswald`;
    ctx.fillText(val, s(x + 10), s(R1 + 42));
  }
}

const CASES = [
  ['REST · 9:57 PM · the floor is asleep', { s:'rest',  t:'2157', p:'88', r:'91', u:'663', g:'0', eb:'0', di:'0', h:DAY.join(',') }],
  ['CLEAR · 1:05 PM · queue empty',        { s:'clear', t:'1305', g:'0',  r:'21', p:'21', eb:'0', di:'0', h:upTo(13).join(',') }],
  ['BUSY · 2:14 PM · 12 to grab',          { s:'busy',  t:'1414', g:'12', r:'14', p:'14', o:'96',  eb:'9', di:'3', h:upTo(14).join(',') }],
  ['LATE · 2:14 PM · 3 past the 3h line',  { s:'late',  t:'1414', g:'12', r:'14', p:'14', o:'192', l:'3', eb:'9', di:'3', h:upTo(14).join(',') }],
  ['STALE · 11:20 AM · pipeline quiet',    { s:'stale', t:'1120', g:'5',  r:'2', p:'3', y:'73', ls:'1007', eb:'4', di:'1', h:upTo(11).join(',') }]
];

const S = Number(process.env.SCALE || 1);
registerFonts();
const PAD = 22, LAB = 20, GAP = 26;
const W = (TOTAL + PAD * 2) * S;
const H = (PAD * 2 + CASES.length * (LAB + R1 + R2 + GAP)) * S;
const canvas = createCanvas(W, H);
const ctx = canvas.getContext('2d');
ctx.fillStyle = '#0b0b0b'; ctx.fillRect(0, 0, W, H);

let y = PAD;
for (const [label, q] of CASES) {
  ctx.fillStyle = '#7d776c'; ctx.font = `500 ${10 * S}px Oswald`;
  ctx.letterSpacing = `${1.4 * S}px`;
  ctx.fillText(label, PAD * S, (y + 11) * S);
  ctx.letterSpacing = '0px';
  y += LAB;
  ctx.save(); ctx.translate(PAD * S, y * S); drawBanner(ctx, q, S); ctx.restore();
  y += R1 + R2 + GAP;
}
const V = process.env.VARIANT === 'B' ? '-b' : '';
const out = path.join(__dirname, 'renders', `banner${V}${S === 1 ? '' : '@' + S + 'x'}.png`);
fs.writeFileSync(out, canvas.toBuffer('image/png'));
console.log(`banner mock ${TOTAL}px wide at ${S}x -> ${out}`);
