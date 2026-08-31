// Render the PRODUCTION wall.html against the mock tick.
'use strict';
const fs = require('fs'); const path = require('path');
const { chromium } = require('playwright');
const MOCK = require('./mock-tick.js');
const html = fs.readFileSync(path.join(__dirname, '..', 'wall.html'), 'utf8');
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1920, height: 1080 }, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  await page.route('http://hqlab.test/**', route => {
    if (route.request().url().includes('/api/board')) {
      return route.fulfill({ contentType: 'application/json',
        body: JSON.stringify(Object.assign({ ok: true }, MOCK)) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com/, r => r.abort());
  await page.goto('http://hqlab.test/wall', { waitUntil: 'load' });
  await page.waitForTimeout(2500);
  await page.screenshot({ path: 'renders/wall-prod.png' });
  const state = await page.evaluate(() => ({
    grab: document.getElementById('grabNum').textContent,
    live: document.getElementById('liveText').textContent,
    ebayRows: document.querySelectorAll('#listE .wrow').length,
    directRows: document.querySelectorAll('#listD .wrow').length,
    bands: document.querySelectorAll('.wband').length,
    feed: document.querySelectorAll('#feedList li').length,
    beacon: document.getElementById('beacon').classList.contains('show'),
    rest: document.getElementById('restveil').classList.contains('show')
  }));
  console.log('state:', JSON.stringify(state));
  console.log(errs.length ? 'ERRORS: ' + errs.join(' | ') : '✓ no page errors');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
