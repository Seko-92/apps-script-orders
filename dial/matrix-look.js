'use strict';
const fs = require('fs'), path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const M = require('./src/board-matrix');
const B = require('./src/board');
registerFonts();

const SC = 2, S = (v) => Math.round(v * SC);
const BARS = [2,1,3,2,5,8,12,17,14,9,13,7,4,2,6,3,1,2,4,6,9,11,7,3];

function housing(ctx, w, h) {
  const g = ctx.createLinearGradient(0, 0, 0, S(h));
  g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
  ctx.fillStyle = g; ctx.fillRect(0, 0, S(w), S(h));
}
/** discs confined to a sub-rect, so the housing keeps its engraved labels */
function grid(ctx, x, y, w, h, compose) {
  const p = createCanvas(S(w), S(h));
  const pc = p.getContext('2d');
  M.drawMatrix(pc, { scale: SC, w, h, compose });
  ctx.drawImage(p, S(x), S(y));
}

const cv = createCanvas(S(1156), S(121) + S(56) + S(74));
const ctx = cv.getContext('2d');
ctx.fillStyle = '#2b2b2b'; ctx.fillRect(0, 0, cv.width, cv.height);
const cap = (t, y) => { ctx.font = `600 ${S(7)}px Oswald`; ctx.fillStyle = '#ffd400';
                        ctx.textBaseline = 'bottom'; ctx.fillText(t, 0, y - S(4)); };

// ── anchor: engraved labels in the housing, values in discs ──────────────────
const a = createCanvas(S(280), S(121)), ac = a.getContext('2d');
housing(ac, 280, 121);
B.drawEngraved(ac, { scale: SC, x: 12, y: 14, text: 'TO GRAB', fontPx: 8, weight: 500, ink: '#6b6459' });
B.drawEngraved(ac, { scale: SC, x: 148, y: 14, text: 'OUT TODAY', fontPx: 8, weight: 500, ink: '#6b6459' });
grid(ac, 10, 24, 128, 62, (c, cols, rows) => {
  c.fillStyle = '#fff'; c.font = '600 14px Oswald';
  c.textAlign = 'right'; c.textBaseline = 'alphabetic';
  c.fillText('6', cols - 1, rows - 1);
});
grid(ac, 146, 24, 128, 62, (c, cols, rows) => {
  c.fillStyle = '#fff'; c.font = '600 14px Oswald';
  c.textAlign = 'right'; c.textBaseline = 'alphabetic';
  c.fillText('30', cols - 1, rows - 1);
});
ctx.save(); ctx.strokeStyle = '#2e2a24'; ctx.lineWidth = Math.max(1, S(0.8));
ctx.beginPath(); ctx.roundRect(S(12), S(94), S(256), S(19), S(2)); ctx.stroke(); ctx.restore();
B.drawEngraved(ac, { scale: SC, x: 140, y: 100.5, text: 'HQ MOTOR SERVICE', fontPx: 8.5,
                     weight: 600, align: 'center', ink: '#7c7466' });
B.drawEngraved(ac, { scale: SC, x: 140, y: 109, text: 'HOUSTON', fontPx: 6.5,
                     weight: 400, align: 'center', ink: '#585144' });
ctx.drawImage(a, 0, S(20)); cap('THE MATRIX · anchor A1:C2 — engraved labels, values in discs', S(20));

// ── strip: the curve across the full width, in the same discs ────────────────
const st = createCanvas(S(876), S(56)), sc2 = st.getContext('2d');
housing(sc2, 876, 56);
B.drawEngraved(sc2, { scale: SC, x: 10, y: 9, text: 'THE DAY', fontPx: 7, weight: 500, ink: '#585144' });
grid(sc2, 8, 15, 860, 36, (c, cols, rows) => {
  c.fillStyle = '#fff';
  const max = Math.max(...BARS);
  const w = Math.max(2, Math.floor((cols - 2) / BARS.length) - 1);
  const step = (cols - 2) / BARS.length;
  BARS.forEach((v, i) => {
    const h = Math.max(1, Math.round((v / max) * (rows - 1)));
    c.fillRect(Math.round(1 + i * step), rows - h, w, h);
  });
});
ctx.drawImage(st, S(280), S(20) + S(121) + S(28));
cap('THE MATRIX · strip D1:H1 — the day curve drawn in the SAME discs as the text',
    S(20) + S(121) + S(28));

fs.writeFileSync(path.join(__dirname, 'renders', 'matrix-look.png'), cv.toBuffer('image/png'));
console.log('renders/matrix-look.png');
