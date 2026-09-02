'use strict';
const fs = require('fs'), path = require('path');
const { GIFEncoder, quantize, applyPalette } = require('gifenc');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A = require('./src/ambient');
const P = require('./src/patterns');
const M = require('./src/board-matrix');
registerFonts();

const W = 280, H = 121, FPS = 8, SECS = 12, FRAMES = FPS * SECS;
const { cols, rows } = A.gridSize(W, H);
const mark = A.markField(cols, rows, [{ text: 'HQ', size: 19, weight: '600' }]);
const TSTRIP = P.tickerStrip(createCanvas, cols, rows, 'HQ MOTOR SERVICE · HOUSTON · ');

function fieldFromDraw(fn) {
  const off = createCanvas(cols, rows), c = off.getContext('2d');
  c.clearRect(0, 0, cols, rows); fn(c);
  const px = c.getImageData(0, 0, cols, rows).data;
  const f = [];
  for (let y = 0; y < rows; y++) {
    const r = [];
    for (let x = 0; x < cols; x++) r.push(px[(y * cols + x) * 4 + 3] > 110);
    f.push(r);
  }
  return f;
}

const ideas = {
  piston:  (ph) => fieldFromDraw((c) => P.piston(c, cols, rows, ph)),
  ticker:  (ph) => fieldFromDraw((c) => P.ticker(c, cols, rows, ph, null, TSTRIP)),
  refresh: (ph) => P.refresh(cols, rows, ph, mark),
  belt:    (ph) => fieldFromDraw((c) => P.belt(c, cols, rows, ph)),
  night:   (ph) => fieldFromDraw((c) => P.night(c, cols, rows, ph))
};

for (const [name, fieldAt] of Object.entries(ideas)) {
  const gif = GIFEncoder();
  const cv = createCanvas(W, H), ctx = cv.getContext('2d');
  let pal;
  for (let f = 0; f < FRAMES; f++) {
    const g = ctx.createLinearGradient(0, 0, 0, H);
    g.addColorStop(0, '#26221c'); g.addColorStop(0.14, '#141210'); g.addColorStop(1, '#100e0c');
    ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
    M.drawMatrix(ctx, { scale: 1, w: W, h: H, field: fieldAt(f / FRAMES) });
    const seam = ctx.createLinearGradient(W - 26, 0, W, 0);
    seam.addColorStop(0, 'rgba(26,26,26,0)'); seam.addColorStop(1, '#1a1a1a');
    ctx.fillStyle = seam; ctx.fillRect(W - 26, 0, 26, H);
    const { data } = ctx.getImageData(0, 0, W, H);
    if (f === 0) pal = quantize(data, 32, { format: 'rgb565' });
    gif.writeFrame(applyPalette(data, pal, 'rgb565'), W, H, {
      palette: f === 0 ? pal : undefined,
      delay: Math.round(1000 / FPS),
      repeat: f === 0 ? 0 : undefined
    });
  }
  gif.finish();
  const b = Buffer.from(gif.bytes());
  fs.writeFileSync(path.join(__dirname, 'renders', 'idea-' + name + '.gif'), b);
  console.log('  idea-' + name + '.gif  ' + (b.length / 1024).toFixed(0) + ' KB');
}
