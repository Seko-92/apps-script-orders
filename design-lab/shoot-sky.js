/**
 * shoot-sky.js — row 2 becomes THE DAY. 24 hour-lit skies, 826x65, for A2:E2.
 *
 * The eBay logo is COMPOSITED IN rather than layered over: A2 can hold one =IMAGE(), so a
 * separate logo cell would mean splitting the merge and moving Schema.cellEmployeeId.
 * LOGO=0 renders the same set without it — a one-flag switch, since retiring a channel
 * badge is a brand decision, not a design one.
 */
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const path = require('path'), fs = require('fs');
const VER = process.env.VER || 'v2';
const LOGO = process.env.LOGO !== '0';

(async () => {
  const OUT = path.join(__dirname, 'renders', 'mast');
  fs.mkdirSync(OUT, { recursive: true });
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 826, height: 65 }, deviceScaleFactor: 3 });
  await p.goto('file://' + path.join(__dirname, 'sky.html'), { waitUntil: 'networkidle' });
  await p.waitForSelector('html[data-ready="1"]');
  await p.evaluate(() => document.fonts.ready);

  const strip = [];
  for (let h = 0; h < 24; h++) {
    await p.evaluate(([hh, lg]) => window.renderSky(hh, lg), [h, LOGO]);
    const f = path.join(OUT, `sky-h${String(h).padStart(2, '0')}-${VER}.png`);
    await p.screenshot({ path: f });
    strip.push(f);
  }
  await b.close();
  execFileSync('convert', [...strip, '-append', path.join(OUT, '_sky-day.png')]);
  console.log(`24 skies (logo ${LOGO ? 'in' : 'out'}) → ${OUT}`);
})();
