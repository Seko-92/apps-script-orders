/**
 * shoot-hours.js — the money shot: one face across all 24 Houston hours, stacked.
 *
 * The masthead cannot MOVE (a floating image scrolls off the frozen banner; =IMAGE()
 * shows a GIF's first frame only). But it can be LIT — and light is the one thing that
 * reads correctly at ~1 frame per minute, because a day changes slowly anyway.
 */
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const path = require('path'), fs = require('fs');

(async () => {
  const OUT = path.join(__dirname, 'renders', 'mast', '_hours');
  fs.rmSync(OUT, { recursive: true, force: true });
  fs.mkdirSync(OUT, { recursive: true });

  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 287, height: 56 }, deviceScaleFactor: 3 });
  await p.goto('file://' + path.join(__dirname, 'masthead.html'), { waitUntil: 'networkidle' });
  await p.waitForSelector('html[data-ready="1"]');
  await p.evaluate(() => document.fonts.ready);

  const frames = [];
  for (let h = 0; h < 24; h++) {
    // show the state that would REALLY be showing at that hour, so the strip reads as a
    // day rather than as a swatch chart
    const state = (h < 9 || h >= 17) ? 'rest' : 'busy';
    await p.evaluate(([s, t, hh]) => window.render(s, t, hh), [state, 0.5, h]);
    const f = path.join(OUT, String(h).padStart(2, '0') + '.png');
    await p.screenshot({ path: f });
    frames.push(f);
  }
  await b.close();

  execFileSync('convert', [...frames, '-append', path.join(OUT, '_day.png')]);
  console.log('24-hour strip → renders/mast/_hours/_day.png');
})();
