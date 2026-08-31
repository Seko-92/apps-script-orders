// Truthful render of the missing-line door. No fake clock — CSS transitions run
// on the compositor, so a driven clock would catch the drawer mid-slide.
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const OUT = path.join(__dirname, 'renders');
let refuseNext = false;

(async () => {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  for (const vp of [{ n: 'tablet-land', w: 1280, h: 800 }, { n: 'tablet-port', w: 800, h: 1280 }]) {
    const ctx = await browser.newContext({ viewport: { width: vp.w, height: vp.h }, hasTouch: true, timezoneId: 'America/Chicago' });
    const page = await ctx.newPage();
    await page.route('http://hqlab.test/**', route => {
      const url = route.request().url();
      if (url.includes('/api/board')) {
        const b = JSON.parse(route.request().postData() || '{}');
        let res = { ok: false };
        if (b.action === 'boardTick') { const t = JSON.parse(JSON.stringify(MOCK)); t.picker = 'Yassin · 1'; t.pickers = ['Shipping - Yassin 1']; res = Object.assign({ ok: true }, t); }
        if (b.action === 'boardRadio') res = { ok: true, nowPlaying: '' };
        if (b.action === 'boardMissingLine') res = refuseNext
          ? { ok: false, message: "That exact line already exists — row 5 is 212498 on 'Missing #: 05-15052-93025'. If you need more units, raise the qty on that row instead of adding a second one." }
          : { ok: true, message: '✅ Added', warnings: [] };
        return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
      }
      return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
    });
    await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
    refuseNext = false;
    await page.goto('http://hqlab.test/', { waitUntil: 'load' });
    await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'), null, { timeout: 20000 }).catch(() => {});
    await page.waitForTimeout(1200);

    await page.click('#menuBtn'); await page.waitForTimeout(150);
    await page.click('#missingBtn'); await page.waitForTimeout(700);   // let the slide settle
    await page.fill('#rlOrder', '05-15052-93025');
    await page.fill('#rlSku', '212498');
    await page.waitForTimeout(150);
    // assert it is SETTLED before shooting — a mid-transition shot is not a picture
    const t = await page.evaluate(() => getComputedStyle(document.getElementById('drw')).transform);
    console.log(vp.n + ' drawer transform: ' + t);
    await page.screenshot({ path: path.join(OUT, 'missingline-' + vp.n + '.png') });

    // The refusal state — the one people actually read.
    // ⚠ DRIVE THE REAL PATH. Calling rlPaint() directly bypasses rlRead(), so it
    //    repaints from a stale rlState and renders PLACEHOLDERS — which looked
    //    exactly like a data-loss bug in code that was fine. Submit for real and
    //    let the (refusing) server produce the state.
    refuseNext = true;
    await page.click('#rlGo');
    await page.waitForTimeout(600);
    await page.screenshot({ path: path.join(OUT, 'missingline-refusal-' + vp.n + '.png') });
    await ctx.close();
  }
  await browser.close();
  console.log('rendered → design-lab/renders/');
})();
