/** Renders the ambient candidates as looping GIFs. */
'use strict';
const fs = require('fs'), path = require('path');
const { GIFEncoder, quantize, applyPalette } = require('gifenc');
const { createCanvas } = require('@napi-rs/canvas');
const { registerFonts } = require('./src/render');
const A = require('./src/ambient');
registerFonts();

// ⚠ 12s at 8fps, not 6s at 10. The user asked for longer, and a slower frame rate suits a
//   field that drifts — 8fps reads as mechanical, which is what a flip-dot IS. It also
//   keeps the file down: 96 frames at 8fps costs less than 120 at 10 for the same duration.
const FPS = 8, SECS = 12, FRAMES = FPS * SECS;

function loop(o) {
  const gif = GIFEncoder();
  const scale = o.scale || 1;
  const cv = createCanvas(o.w * scale, o.h * scale);
  const ctx = cv.getContext('2d');
  const { cols, rows } = A.gridSize(o.w, o.h);
  const mark = A.markField(cols, rows, o.lines);
  A.setWide(!!o.wide);
  let pal, tIdx = -1;

  for (let f = 0; f < FRAMES; f++) {
    A.drawAmbient(ctx, { w: o.w, h: o.h, scale, mode: o.mode,
                         phase: f / FRAMES, mark, windows: o.windows, seam: o.seam });
    const { data } = ctx.getImageData(0, 0, cv.width, cv.height);
    if (f === 0) {
      pal = quantize(data, 32, { format: 'rgb565' });
      if (o.windows && o.windows.length) {
        // find the palette entry closest to the sentinel and make IT the transparent index
        let best = 0, bd = Infinity;
        pal.forEach((c, i) => {
          const d = Math.abs(c[0] - 255) + Math.abs(c[1] - 0) + Math.abs(c[2] - 255);
          if (d < bd) { bd = d; best = i; }
        });
        tIdx = best;
      }
    }
    gif.writeFrame(applyPalette(data, pal, 'rgb565'), cv.width, cv.height, {
      palette: f === 0 ? pal : undefined,
      delay: Math.round(1000 / FPS),
      repeat: f === 0 ? 0 : undefined,          // 0 = forever. Ambience must not stop.
      transparent: tIdx >= 0,
      transparentIndex: tIdx >= 0 ? tIdx : undefined
    });
  }
  gif.finish();
  return Buffer.from(gif.bytes());
}

// Column geometry, from BrandTheme's own setColumnWidth calls.
const A_C = 280, D = 232, E = 307, FH = 337;

const variants = [
  // 1 · the anchor alone. D1/E1/F1:H1 stay live and uncovered.
  { name: 'anchor', w: 280, h: 121, mode: 'mark',
    lines: [{ text: 'HQ', size: 19, weight: '600' }] },

  // 2 · the strip that turns it into an L across all of row 1. Figures gone.
  //   ⚠ The wave frequency is HALVED for this aspect. At 14 disc rows a 4.4-cycle swell
  //     thins into horizontal streaks; wide-and-short needs wider blobs to still read as
  //     a field. Same drawing, different proportion — the tuning does not carry over.
  { name: 'strip', w: 876, h: 56, mode: 'mark', seam: false, wide: true,
    lines: [{ text: 'HQ MOTOR SERVICE', size: 11, weight: '600' }] },

  // 3 · the same strip, but the pulse and the day curve show THROUGH it.
  { name: 'strip-win', w: 876, h: 56, mode: 'mark', seam: false, wide: true,
    lines: [{ text: 'HQ', size: 12, weight: '600', dx: -300 }],
    windows: [ { x: 236, y: 6, w: 299, h: 44 }, { x: 543, y: 6, w: 329, h: 44 } ] }
];

for (const v of variants) {
  const g = loop(v);
  fs.writeFileSync(path.join(__dirname, 'renders', 'amb-' + v.name + '.gif'), g);
  console.log('  amb-' + v.name + '.gif  ' + v.w + 'x' + v.h + '  ' +
              (g.length / 1024).toFixed(0) + ' KB' + (v.windows ? '  (windowed)' : ''));
}
console.log('\n' + FRAMES + ' frames @' + FPS + 'fps = ' + SECS + 's, loops forever');
