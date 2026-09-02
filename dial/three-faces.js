'use strict';
const fs = require('fs'), path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const T = require('./src/board-terminal');
const M = require('./src/board-matrix');
const B = require('./src/board');
registerFonts();

const Q = { s: 'rest', t: '1914', p: '30', r: '8', u: '295', o: '95', g: '6' };
const SC = 2, S = (v) => Math.round(v * SC), W = 1156, H = 121;

function liveCells(ctx) {                     // D1, E1, F1:H1 — uncovered in every option
  ctx.textBaseline = 'middle';
  ctx.fillStyle = '#f0ece0'; ctx.font = `400 ${S(11)}px Oswald`;
  ctx.fillText('the floor is asleep', S(292), S(21));
  ctx.fillStyle = '#8b8578'; ctx.font = `400 ${S(9)}px Oswald`;
  ctx.fillText('waiting: 2', S(292), S(37));
  const lx = 280 + 232 + 18;
  ctx.beginPath(); ctx.arc(S(lx), S(28), S(3), 0, Math.PI * 2);
  ctx.fillStyle = '#7e8894'; ctx.fill();
  ctx.fillStyle = '#b9b4a8'; ctx.font = `400 ${S(10)}px Oswald`;
  ctx.fillText('RESTING · 7:14 PM · 8h 50m ago', S(lx + 9), S(28));
  const x0 = 280 + 232 + 307 + 14;
  ctx.fillStyle = '#7e8894';
  [2,1,3,2,5,8,12,17,14,9,13,7,4,2,6,3,1,2].forEach((v, i) => {
    const h = 2 + v * 1.9; ctx.fillRect(S(x0 + i * 8.4), S(46 - h), S(6), S(h));
  });
}

function plate(c) {
  c.save(); c.strokeStyle = '#2e2a24'; c.lineWidth = Math.max(1, S(0.8));
  c.beginPath(); c.roundRect(S(12), S(94), S(256), S(19), S(2)); c.stroke(); c.restore();
  B.drawEngraved(c, { scale: SC, x: 140, y: 100.5, text: 'HQ MOTOR SERVICE',
                      fontPx: 8.5, weight: 600, align: 'center', ink: '#7c7466' });
  B.drawEngraved(c, { scale: SC, x: 140, y: 109, text: 'HOUSTON',
                      fontPx: 6.5, weight: 400, align: 'center', ink: '#585144' });
}

const faces = [
  ['A · TERMINAL · the STATE     — flaps say RESTING (D1 and E1 already say it)',
   (c) => { const b = T.buildBoard(Q); T.drawAnchor(c, b, b, 99, SC); }],
  ['B · TERMINAL · the FIGURES   — flaps say what nothing else on row 1 says',
   (c) => { const b = T.buildFigures(Q); T.drawAnchorFigures(c, b, b, 99, SC); }],
  ['C · MATRIX   · the FIGURES   — same, in flip-discs',
   (c) => {
     const g = c.createLinearGradient(0, 0, 0, S(121));
     g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
     c.fillStyle = g; c.fillRect(0, 0, S(280), S(121));
     B.drawEngraved(c, { scale: SC, x: 12, y: 14, text: 'TO GRAB', fontPx: 8, weight: 500, ink: '#6b6459' });
     B.drawEngraved(c, { scale: SC, x: 148, y: 14, text: 'OUT TODAY', fontPx: 8, weight: 500, ink: '#6b6459' });
     [[10, '6'], [146, '30']].forEach(([x, txt]) => {
       const p = createCanvas(S(128), S(62));
       M.drawMatrix(p.getContext('2d'), { scale: SC, w: 128, h: 62, compose: (cc, cols, rows) => {
         cc.fillStyle = '#fff'; cc.font = '600 14px Oswald';
         cc.textAlign = 'right'; cc.textBaseline = 'alphabetic';
         cc.fillText(txt, cols - 1, rows - 1);
       }});
       c.drawImage(p, S(x), S(24));
     });
     plate(c);
   }]
];

const cv = createCanvas(S(W), (S(H) + S(26)) * faces.length + S(10));
const ctx = cv.getContext('2d');
ctx.fillStyle = '#2b2b2b'; ctx.fillRect(0, 0, cv.width, cv.height);

faces.forEach(([label, draw], i) => {
  const y = S(22) + i * (S(H) + S(26));
  const band = createCanvas(S(W), S(H)), bc = band.getContext('2d');
  bc.fillStyle = '#1a1a1a'; bc.fillRect(0, 0, band.width, band.height);
  bc.fillStyle = '#fff8e7'; bc.fillRect(0, S(56), band.width, S(65));
  liveCells(bc);
  const a = createCanvas(S(280), S(121));
  draw(a.getContext('2d'));
  bc.drawImage(a, 0, 0);
  ctx.drawImage(band, 0, y);
  ctx.font = `600 ${S(7.5)}px Oswald`; ctx.fillStyle = '#ffd400'; ctx.textBaseline = 'bottom';
  ctx.fillText(label, 0, y - S(5));
});

fs.writeFileSync(path.join(__dirname, 'renders', 'three-faces.png'), cv.toBuffer('image/png'));
console.log('renders/three-faces.png');
