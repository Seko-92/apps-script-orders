// Batch the dial-size variants side by side. The `clamp(10px,1.95vh,16.5px)`
// cap was tuned when the dial lived INSIDE a panel on the old board; it now
// owns the whole screen. Render the candidates, let the eye pick in seconds.
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const html = fs.readFileSync(path.join(__dirname, '..', 'wall.html'), 'utf8');
const MOCK = require('./mock-tick.js');
const OUT  = path.join(__dirname, 'renders');

const VARIANTS = [
  ['a-current', 'clamp(10px, 1.95vh, 16.5px)', 'as ported — 396px dial'],
  ['b-medium',  'clamp(10px, 2.40vh, 24px)',   'medium — 576px dial'],
  ['c-large',   'clamp(10px, 2.90vh, 30px)',   'large — 720px dial']
];

function quietTick() {
  const t = JSON.parse(JSON.stringify(MOCK));
  t.cockpit.ebayGrab = 0; t.cockpit.directGrab = 0;
  t.openOrders = []; t.openOrdersBy = {}; t.openOrdersTotal = 0;
  return t;
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();

  for (const [tag, size, label] of VARIANTS) {
    for (const [sn, w, h] of [['wall', 1920, 1080], ['tablet-landscape', 1280, 800]]) {
      const ctx = await browser.newContext({ viewport: { width: w, height: h }, timezoneId: 'America/Chicago' });
      const page = await ctx.newPage();
      await page.clock.install({ time: new Date('2026-08-13T22:15:00-05:00') });
      await page.route('http://hqlab.test/**', route => {
        if (route.request().url().includes('/api/board')) {
          return route.fulfill({ contentType: 'application/json',
            body: JSON.stringify(Object.assign({ ok: true }, quietTick())) });
        }
        return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
      });
      await page.route(/aladhan\.com/, r => r.abort());
      await page.addInitScript(fs2 => {
        document.addEventListener('DOMContentLoaded', () => {
          const s = document.createElement('style');
          s.textContent = `.dial{font-size:${fs2} !important;}`;
          document.head.appendChild(s);
        });
      }, size);
      await page.goto('http://hqlab.test/wall', { waitUntil: 'load' });
      await page.clock.runFor(4000);
      await page.waitForTimeout(900);
      const box = await page.evaluate(() => {
        const b = document.querySelector('.dial').getBoundingClientRect();
        const s = document.querySelector('.rest-side').getBoundingClientRect();
        return { d: Math.round(b.width), top: Math.round(b.top), bottom: Math.round(b.bottom),
                 sideLeft: Math.round(s.left), sideRight: Math.round(s.right),
                 H: window.innerHeight, W: window.innerWidth,
                 scrollH: document.documentElement.scrollHeight };
      });
      const fits = box.bottom <= box.H + 1 && box.top >= -1 && box.scrollH <= box.H + 1;
      console.log(`${fits ? '✓' : '✗'} ${tag} @ ${sn}: dial ${box.d}px  (${box.top}..${box.bottom} of ${box.H})`
                + `  left flank ends x=${box.sideRight}   ${label}`);
      await page.screenshot({ path: path.join(OUT, `dialscale-${tag}-${sn}.png`) });
      await ctx.close();
    }
  }
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
