// =====================================================================================
// shoot-countline.js — the count line, BEFORE and AFTER the 2026-08-20 floor report.
//
// Shoots the two rows the picker was looking at, in the two states they named:
//   · counted and AGREES        — "should be 9"  beside  "counted 9"
//   · counted, fixed in Zoho    — "should be 4"  beside  "✓ zoho 1"
//
// Run against both boards to get a real comparison:
//   BOARD_FILE=/tmp/old-board.html TAG=before node shoot-countline.js
//   node shoot-countline.js
// =====================================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const TAG   = process.env.TAG || 'after';
const MOCK  = require('./mock-tick.js');
const OUT   = path.join(__dirname, 'renders');

(async () => {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 800, height: 1280 },
    hasTouch: true, timezoneId: 'America/Chicago', deviceScaleFactor: 2 });
  const page = await ctx.newPage();

  let savedLeft = {};
  const tick = () => {
    const t = JSON.parse(JSON.stringify(MOCK));
    (t.openOrders || []).forEach(r => {
      if (Object.prototype.hasOwnProperty.call(savedLeft, r.sku)) r.left = Number(savedLeft[r.sku]);
    });
    return t;
  };
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const b = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (b.action === 'boardTick')   res = Object.assign({ ok: true }, tick());
      if (b.action === 'boardLeft')   { savedLeft[b.sku] = b.count; res = { ok: true }; }
      /* ⚠ THE STUB MUST RETURN before/after/delta. The receipt renders
         `num(res.after)` — the number the SERVER confirmed, not the one we sent,
         which is the right design. A stub answering a bare {ok:true} therefore
         rendered "zoho ← 0" and looked exactly like the board pushing zero stock
         to Zoho. It was the harness. */
      if (b.action === 'boardAdjust') res = { ok: true, before: 4, after: 1, delta: -3, adjustment_id: 'T-1' };
      if (b.action === 'boardStatus') res = { ok: true };
      if (b.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 });
  await page.waitForTimeout(1200);

  const rowFor = (sku) => page.evaluateHandle((s) =>
    [...document.querySelectorAll('.pick-row')].find(r => (r.textContent || '').indexOf(s) !== -1), sku);

  // 1 · count a shelf that AGREES (194244, hand 9)
  await page.evaluate(() => [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf('194244') !== -1).querySelector('.pc-btn').click());
  await page.waitForTimeout(400);
  await page.keyboard.press('9'); await page.keyboard.press('Enter');
  await page.waitForTimeout(1600);

  // 2 · fix a disagreeing shelf in Zoho (165447, hand 4, counted 1)
  await page.evaluate(() => [...document.querySelectorAll('.pick-row')]
    .find(r => (r.textContent || '').indexOf('165447') !== -1).querySelector('.pc-adj').click());
  /* ⚠ WAIT FOR THE PREFILL, DO NOT GUESS AT IT. A 450ms wait raced openNumPad
     and Enter fired against the pad's default 0 — which rendered a receipt
     reading "zoho ← 0" and looked exactly like a product bug pushing zero stock
     to Zoho. It is not: probed directly, the real path sends target 1. A render
     that lies is worse than no render, so this waits for the value to land. */
  await page.waitForFunction(() =>
    document.getElementById('npValue').textContent === '1', null, { timeout: 5000 });
  await page.keyboard.press('Enter');
  await page.waitForTimeout(1500);

  for (const [sku, label] of [['194244', 'agrees'], ['165447', 'fixed']]) {
    const h = await rowFor(sku);
    const el = h.asElement();
    if (el) await el.screenshot({ path: path.join(OUT, `countline-${TAG}-${label}.png`) });
    const txt = await page.evaluate((s) => {
      const li = [...document.querySelectorAll('.pick-row')].find(r => (r.textContent || '').indexOf(s) !== -1);
      const vis = e => e && getComputedStyle(e).display !== 'none';
      const sys = li.querySelector('.pc-sys');
      return { sysText: sys ? sys.textContent : null, sysVisible: vis(sys),
               line: (li.querySelector('.pick-count') || {}).innerText.replace(/\s+/g, ' ').trim() };
    }, sku);
    console.log(`  ${TAG}/${label}: pc-sys=${JSON.stringify(txt.sysText)} visible=${txt.sysVisible}`);
    console.log(`     line reads: "${txt.line}"`);
  }

  await browser.close();
  console.log(`\nrenders → countline-${TAG}-agrees.png · countline-${TAG}-fixed.png`);
})();
