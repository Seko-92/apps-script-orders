'use strict';
const fs = require('fs'), path = require('path');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A = require('./src/ambient');
const P = require('./src/patterns');
const M = require('./src/board-matrix');
registerFonts();

const W = 280, H = 121, SC = 2;
const { cols, rows } = A.gridSize(W, H);
const mark = A.markField(cols, rows, [{ text: 'HQ', size: 19, weight: '600' }]);
const TSTRIP = P.tickerStrip(createCanvas, cols, rows, 'HQ MOTOR SERVICE · HOUSTON · ');

function fieldFromDraw(fn) {
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows);
  fn(c);
  const px = c.getImageData(0, 0, cols, rows).data;
  const f = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) r.push(px[(y * cols + x) * 4 + 3] > 110);
    f.push(r);
  }
  return f;
}

const ideas = [
  ['THE FOUR-STROKE  · a piston in its bore', (ph) => fieldFromDraw((c) => P.piston(c, cols, rows, ph))],
  ['THE TICKER       · text walks the field', (ph) => fieldFromDraw((c) => P.ticker(c, cols, rows, ph, null, TSTRIP))],
  ['THE REFRESH      · a wavefront flips the board', (ph) => P.refresh(cols, rows, ph, mark)],
  ['THE BELT         · parts travelling', (ph) => fieldFromDraw((c) => P.belt(c, cols, rows, ph))],
  ['THE NIGHT        · sparse drift', (ph) => fieldFromDraw((c) => P.night(c, cols, rows, ph))]
];

const phases = [0.05, 0.28, 0.5, 0.72];
const PAD = 12;
const cv = createCanvas((W * SC + PAD) * phases.length + PAD,
                        (H * SC + PAD + 20) * ideas.length + PAD);
const ctx = cv.getContext('2d');
ctx.fillStyle = '#2b2b2b'; ctx.fillRect(0, 0, cv.width, cv.height);

ideas.forEach(([label, fieldAt], i) => {
  const y = PAD + 16 + i * (H * SC + PAD + 20);
  ctx.font = '600 13px Oswald'; ctx.fillStyle = '#ffd400'; ctx.textBaseline = 'bottom';
  ctx.fillText(label, PAD, y - 4);
  phases.forEach((ph, j) => {
    const p = createCanvas(W * SC, H * SC), pc = p.getContext('2d');
    const g = pc.createLinearGradient(0, 0, 0, H * SC);
    g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
    pc.fillStyle = g; pc.fillRect(0, 0, p.width, p.height);
    M.drawMatrix(pc, { scale: SC, w: W, h: H, field: fieldAt(ph) });
    ctx.drawImage(p, PAD + j * (W * SC + PAD), y);
  });
});
fs.writeFileSync(path.join(__dirname, 'renders', 'idea-sheet.png'), cv.toBuffer('image/png'));
console.log('renders/idea-sheet.png');
