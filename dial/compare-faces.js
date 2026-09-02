/** Two faces, true size, on the real band with D1/E1/F1:H1 live and uncovered. */
'use strict';
const fs = require('fs'), path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const T = require('./src/board-terminal');
registerFonts();

const Q = { s: 'rest', t: '1914', p: '30', r: '8', u: '295', o: '95', g: '6' };
const W = 1156, H = 121, SC = 2, S = (v) => Math.round(v * SC);

function cells(ctx) {
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

const cv = createCanvas(S(W), S(H) * 2 + S(34));
const ctx = cv.getContext('2d');
ctx.fillStyle = '#2b2b2b'; ctx.fillRect(0, 0, cv.width, cv.height);

[['A · THE STATE  — says RESTING, which D1 and E1 already say',
  (c) => { const b = T.buildBoard(Q); T.drawAnchor(c, b, b, 99, SC); }],
 ['B · THE FIGURES — says what nothing else on row 1 says',
  (c) => { const b = T.buildFigures(Q); T.drawAnchorFigures(c, b, b, 99, SC); }]
].forEach(([label, draw], i) => {
  const y = S(17) + i * (S(H) + S(17));
  const band = createCanvas(S(W), S(H)), bc = band.getContext('2d');
  bc.fillStyle = '#1a1a1a'; bc.fillRect(0, 0, band.width, band.height);
  bc.fillStyle = '#fff8e7'; bc.fillRect(0, S(56), band.width, S(65));
  cells(bc);
  const a = createCanvas(S(280), S(121));
  draw(a.getContext('2d'));
  bc.drawImage(a, 0, 0);
  ctx.drawImage(band, 0, y);
  ctx.font = `600 ${S(7)}px Oswald`; ctx.fillStyle = '#ffd400'; ctx.textBaseline = 'bottom';
  ctx.fillText(label, 0, y - S(4));
});
fs.writeFileSync(path.join(__dirname, 'renders', 'compare-faces.png'), cv.toBuffer('image/png'));
console.log('renders/compare-faces.png');
