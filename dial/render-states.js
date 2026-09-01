#!/usr/bin/env node
/**
 * render-states.js — PHASE 1. Writes the real states to PNG so a human can look.
 *
 *   node render-states.js            1x + 2x into renders/
 *   SCALE=3 node render-states.js    a big one for reading detail
 *
 * ⚠⚠ THIS IS THE STEP THE 2026-08-31 ROUND SKIPPED, and skipping it is why that round was
 *    rejected wholesale. The sheet cannot be rendered headlessly; these files are the only
 *    preview that exists before a formula lands on the live banner.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { renderPng } = require('./src/render');

const OUT = path.join(__dirname, 'renders');
fs.mkdirSync(OUT, { recursive: true });

// A real shift, shaped like the live Activity Log: quiet overnight, ramp from 9, peak
// late morning, tail off after 4.
const DAY = [0,0,0,0,0,0,0,0,1,6,14,19,11,16,22,13,7,2,0,0,0,0,0,0];
const h = (a) => a.join(',');
/** The day SO FAR. ⚠ A fixture with counts on hours that have not happened yet is not a
 *  day, it is a bug report — and it hid a real one until the future-hour guard went in. */
const upTo = (hh) => h(DAY.map((n, i) => (i <= hh ? n : 0)));

const CASES = [
  { name: 'rest-night',   q: { s:'rest',  t:'2157', p:'88', r:'91', u:'663', h:h(DAY)  } },
  { name: 'rest-dawn',    q: { s:'rest',  t:'0640', p:'88', r:'91', u:'140', h:h(DAY)  } },
  { name: 'clear-open',   q: { s:'clear', t:'0912', g:'0',  r:'2',  p:'0',  h:upTo(9)  } },
  { name: 'clear-midday', q: { s:'clear', t:'1305', g:'0',  r:'21', p:'21', h:upTo(13) } },
  { name: 'busy-morning', q: { s:'busy',  t:'1042', g:'7',  r:'9',  p:'4',  o:'64',  h:upTo(10) } },
  { name: 'busy-after',   q: { s:'busy',  t:'1414', g:'12', r:'14', p:'14', o:'96',  h:upTo(14) } },
  { name: 'late',         q: { s:'late',  t:'1414', g:'12', r:'14', p:'14', o:'192', l:'3', h:upTo(14) } },
  { name: 'late-bad',     q: { s:'late',  t:'1630', g:'19', r:'21', p:'9',  o:'412', l:'7', h:upTo(16) } },
  { name: 'stale',        q: { s:'stale', t:'1120', g:'5',  p:'3',  y:'73',  h:upTo(11) } },
  { name: 'empty-args',   q: { s:'busy',  t:'1414' } },
  { name: 'cold-start',   q: {} }
];

const scales = process.env.SCALE ? [Number(process.env.SCALE)] : [1, 2];
let total = 0;
for (const c of CASES) {
  for (const s of scales) {
    const buf = renderPng(c.q, s);
    const file = path.join(OUT, `${c.name}${s === 1 ? '' : '@' + s + 'x'}.png`);
    fs.writeFileSync(file, buf);
    total += buf.length;
    if (s === 1) console.log(`  ${c.name.padEnd(14)} ${String(buf.length).padStart(6)}b`);
  }
}
console.log(`\n${CASES.length} states x ${scales.join('/')}x -> ${OUT}  (${(total/1024).toFixed(1)} KB total)`);
