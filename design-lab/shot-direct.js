// One-off: tablet-portrait with the DIRECT tab active (the accordion view).
'use strict';
const fs = require('fs'); const path = require('path');
const { chromium } = require('playwright');
const MOCK = require('./mock-tick.js');
const html = fs.readFileSync(path.join(__dirname, '..', 'FloorBoard.html'), 'utf8');
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 800, height: 1280 }, hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick') res = Object.assign({ ok: true }, MOCK);
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'), null, { timeout: 20000 });
  await page.waitForTimeout(2000);
  await page.tap('.ct[data-ch="DIRECT"]');
  await page.waitForTimeout(1200);
  await page.screenshot({ path: 'renders/board-tablet-portrait-direct.png' });
  console.log('✓ DIRECT tab shot saved');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
