// Shoot the LAB pages: tool (portrait ebay/direct + landscape) + wall.
'use strict';
const path = require('path');
const { chromium } = require('playwright');
const SHOTS = [
  { file: 'lab-tool.html', q: '',            vp: [800, 1280],  touch: true,  out: 'lab-tool-portrait-ebay.png' },
  { file: 'lab-tool.html', q: '?tab=direct', vp: [800, 1280],  touch: true,  out: 'lab-tool-portrait-direct.png' },
  { file: 'lab-tool.html', q: '',            vp: [1280, 800],  touch: true,  out: 'lab-tool-landscape.png' },
  { file: 'lab-wall.html', q: '',            vp: [1920, 1080], touch: false, out: 'lab-wall.png' }
];
(async () => {
  const browser = await chromium.launch();
  for (const s of SHOTS) {
    const ctx = await browser.newContext({ viewport: { width: s.vp[0], height: s.vp[1] },
      hasTouch: s.touch, timezoneId: 'America/Chicago' });
    const page = await ctx.newPage();
    const errs = [];
    page.on('pageerror', e => errs.push(e.message));
    await page.goto('file://' + path.join(__dirname, s.file) + s.q, { waitUntil: 'load' });
    await page.waitForTimeout(1800);
    await page.screenshot({ path: 'renders/' + s.out });
    console.log('✓', s.out, errs.length ? 'ERRS: ' + errs.join(' | ') : '');
    await ctx.close();
  }
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
