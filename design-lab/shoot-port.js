// Render the PORTED FloorBoard.html — five states, page-error strict.
'use strict';
const fs = require('fs'); const path = require('path');
const { chromium } = require('playwright');
const MOCK = require('./mock-tick.js');
const html = fs.readFileSync(path.join(__dirname, '..', 'FloorBoard.html'), 'utf8');

const SHOTS = [
  { name: 'port-portrait-ebay',   vp: [800, 1280], act: null },
  { name: 'port-portrait-direct', vp: [800, 1280], act: async p => p.tap('.ct[data-ch="DIRECT"]') },
  { name: 'port-portrait-today',  vp: [800, 1280], act: async p => p.tap('#todayBar') },
  { name: 'port-landscape',       vp: [1280, 800], act: null },
  { name: 'port-phone',           vp: [390, 844],  act: null }
];
(async () => {
  const browser = await chromium.launch();
  for (const s of SHOTS) {
    const ctx = await browser.newContext({ viewport: { width: s.vp[0], height: s.vp[1] },
      hasTouch: true, timezoneId: 'America/Chicago' });
    const page = await ctx.newPage();
    const errs = [];
    page.on('pageerror', e => errs.push(e.message));
    page.on('console', m => { if (m.type() === 'error' && !/net::ERR_FAILED/.test(m.text())) errs.push('console: ' + m.text()); });
    await page.route('http://hqlab.test/**', route => {
      if (route.request().url().includes('/api/board')) {
        const body = JSON.parse(route.request().postData() || '{}');
        let res = { ok: false, message: 'unknown ' + body.action };
        if (body.action === 'boardTick') res = Object.assign({ ok: true }, MOCK);
        if (body.action === 'boardStatus') res = { ok: true };
        return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
      }
      return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
    });
    await page.route(/aladhan\.com/, r => r.abort());
    await page.goto('http://hqlab.test/', { waitUntil: 'load' });
    await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
      null, { timeout: 15000 }).catch(() => errs.push('never left booting'));
    await page.waitForTimeout(1500);
    if (s.act) { await s.act(page); await page.waitForTimeout(900); }
    const st = await page.evaluate(() => ({
      ch: document.body.className.match(/ch-\w+/) ? document.body.className.match(/ch-\w+/)[0] : '',
      rows: document.querySelectorAll('.pick-row').length,
      grab: document.getElementById('grabNum').textContent
    }));
    await page.screenshot({ path: 'renders/' + s.name + '.png' });
    console.log('✓', s.name.padEnd(22), st.ch, 'rows=' + st.rows, 'grab=' + st.grab,
                errs.length ? ' ⚠ ' + errs.join(' | ').slice(0, 300) : '');
    await ctx.close();
  }
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
