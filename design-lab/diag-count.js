// ============================================================================
// SHELF-COUNT MODEL TEST  —  rewritten for the 2026-08-13 ruling
//
// FLOOR FEEDBACK #5: the board must NOT predict what a shelf will hold. HAND on
// the sheet is already correct and near-live (Zoho's 2-min stock push for
// DIRECT/manual/prep rows; live eBay GetItem at order arrival for automated
// eBay rows), so the board's job is to READ it and hand it to the picker.
//
// This supersedes the 2026-08-12 model, which subtracted the line's qty once
// the row went PREPARING. That was itself a fix for a worse bug — every correct
// ×1 shelf scored "+1 vs system" — but it was still arithmetic layered on a
// number that was already right.
//
// THE ASSERTIONS:
//   PENDING    → reference is `hand`
//   PREPARING  → reference is STILL `hand`   ← the ruling, and the regression
//                that matters: it must NOT move when the row is picked
//   COUNT survives the pick (the 2026-08-12 fix must not regress)
//   a correct shelf raises NO deviance in either state
//   the Zoho suggestion is EXACTLY what was counted — proven by opening the
//   numpad and reading the value, not by inferring from attributes
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK  = require('./mock-tick.js');
const OUT   = path.join(__dirname, 'renders');

let fails = [];
const check = (name, got, want) => {
  const ok = String(got) === String(want);
  console.log(`   ${ok ? '✓' : '✗'} ${name}  →  ${JSON.stringify(got)}${ok ? '' : `   (expected ${JSON.stringify(want)})`}`);
  if (!ok) fails.push(`${name}: got ${got}, expected ${want}`);
};

(async () => {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 800, height: 1280 },
    hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));

  let statusCalls = [];
  /* ⚠ THE STUB HAS TO BEHAVE LIKE THE SERVER, not just answer ok. saveCount has
     NO optimistic render — it calls pollSoon() and waits for the sheet — so a
     stub that returns {ok:true} and keeps serving the ORIGINAL tick would leave
     the row looking uncounted forever, and every assertion about the
     already-counted states would be meaningless. So boardLeft is remembered and
     folded into every later boardTick, exactly as the round trip really works. */
  let savedLeft = {};
  const tickWithSavedCounts = () => {
    const t = JSON.parse(JSON.stringify(MOCK));
    (t.openOrders || []).forEach(r => {
      if (Object.prototype.hasOwnProperty.call(savedLeft, r.sku)) {
        const v = savedLeft[r.sku];
        if (v === '' || v === null) delete r.left; else r.left = Number(v);
      }
    });
    return t;
  };
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick')   res = Object.assign({ ok: true }, tickWithSavedCounts());
      if (body.action === 'boardStatus') { statusCalls.push(body); res = { ok: true }; }
      if (body.action === 'boardLeft')   { savedLeft[body.sku] = body.count; res = { ok: true }; }
      /* ⚠ THE STUB MUST RETURN before/after/delta. The receipt renders
         `num(res.after)` — the number the SERVER confirmed, not the one we sent,
         which is the right design. A stub answering a bare {ok:true} therefore
         rendered "zoho ← 0" and looked exactly like the board pushing zero stock
         to Zoho. It was the harness. */
      if (body.action === 'boardAdjust') res = { ok: true, before: 4, after: 1, delta: -3, adjustment_id: 'T-1' };
      if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());

  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'),
    null, { timeout: 20000 });
  await page.waitForTimeout(1200);

  // A-14 / 194244 — hand 9, qty 1, no prior count. The clean case.
  const read = () => page.evaluate(() => {
    const rows = [...document.querySelectorAll('.pick-row')];
    const li = rows.find(r => (r.textContent || '').indexOf('194244') !== -1);
    if (!li) return null;
    const q = s => { const e = li.querySelector(s); return e ? e.textContent.trim() : null; };
    const btn = li.querySelector('.pc-btn');
    return {
      sys: q('.pc-sys'),
      hasCountBtn: !!btn,
      expectedAttr: btn ? btn.getAttribute('data-expected') : null,
      isPrep: !!li.querySelector('.pick-status.prep'),
      hasPickBtn: !!li.querySelector('.pick-do'),
      deviance: li.classList.contains('deviance')
    };
  });

  console.log('\n── BEFORE THE PICK (PENDING) — hand 9, qty 1 ───────────────');
  const before = await read();
  check('reference shown', before.sys, 'should be 9');
  check('count button present', before.hasCountBtn, true);
  check('numpad target', before.expectedAttr, '9');
  check('pick button present', before.hasPickBtn, true);
  check('no false deviance', before.deviance, false);

  console.log('\n── TAP ✓ PICK — the reference must NOT move ────────────────');
  await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('194244') !== -1);
    li.querySelector('.pick-do').click();
  });
  await page.waitForTimeout(600);
  const after = await read();
  check('flipped to PREPARING', after.isPrep, true);
  check('COUNT SURVIVED the pick', after.hasCountBtn, true);
  // ⭐ THE RULING. Picking a row changes nothing about what is on hand — Zoho
  // only drops stock_on_hand on SHIP, and eBay's quantitySold moved when the
  // buyer paid. So the number the picker is asked to match is CONSTANT.
  check('reference UNCHANGED by the pick', after.sys, before.sys);
  check('reference is still hand', after.sys, 'should be 9');
  check('numpad target unchanged', after.expectedAttr, '9');
  check('still no false deviance', after.deviance, false);
  check('boardStatus fired once', statusCalls.length, 1);

  // B-30 / 165447 — hand 4, qty 2, left=1 already recorded → a genuine −3.
  console.log('\n── A REAL DEVIANCE (B-30: hand 4, qty 2, counted 1) ─────────');
  const dev = await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('165447') !== -1);
    const adj = li.querySelector('.pc-adj');
    return {
      chip: (li.querySelector('.pc-dev') || {}).textContent,
      sys: (li.querySelector('.pc-sys') || {}).textContent || null,
      flagged: li.classList.contains('deviance'),
      counted: adj && adj.getAttribute('data-counted'),
      hasAdj: !!adj,
      // these fed the retired projection and must be gone
      pulled: adj && adj.getAttribute('data-pulled'),
      otherPrep: adj && adj.getAttribute('data-otherprep')
    };
  });
  check('deviance chip wording', (dev.chip || '').includes('vs shelf'), true);
  check('deviance size', dev.chip, '-3 vs shelf');
  check('"should be" GONE once counted', dev.sys, null);
  check('row flagged', dev.flagged, true);
  check('fix button offered', dev.hasAdj, true);
  check('zoho: counted', dev.counted, '1');
  check('retired data-pulled is GONE', dev.pulled, null);
  check('retired data-otherprep is GONE', dev.otherPrep, null);

  // ⭐ END-TO-END: open the numpad and read what it actually pre-fills. The
  // suggestion must be the COUNT, with nothing added back.
  console.log('\n── THE ZOHO SUGGESTION (numpad, opened for real) ───────────');
  await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('165447') !== -1);
    li.querySelector('.pc-adj').click();
  });
  await page.waitForTimeout(500);
  const pad = await page.evaluate(() => ({
    open: !document.getElementById('numPad').classList.contains('hidden'),
    value: (document.getElementById('npValue') || {}).textContent,
    ctx: (document.getElementById('npContext') || {}).textContent || ''
  }));
  check('numpad opened', pad.open, true);
  check('pre-filled with the COUNT, nothing added', pad.value, '1');
  check('context does not talk about pulled units',
        /already pulled|pulled elsewhere|yours|other/i.test(pad.ctx), false);
  console.log(`   context: ${JSON.stringify(pad.ctx.replace(/\s+/g, ' ').trim())}`);

  /* ═══ 2026-08-20 FLOOR REPORT — "should be N" must go once the shelf is
     counted. The picker named two states, correct and fixed, and the SECOND is
     the one that was actually broken. ═══ */

  console.log('\n── FIXED: push the count to Zoho, then look at the row ──────');
  // The numpad is still open on 165447 from the section above, pre-filled with
  // the count. Confirming it is the real fix path.
  await page.keyboard.press('Enter');
  await page.waitForTimeout(1400);
  const fixed = await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('165447') !== -1);
    return {
      sys: (li.querySelector('.pc-sys') || {}).textContent || null,
      receipt: (li.querySelector('.pc-zoho') || {}).textContent || null,
      counted: (li.querySelector('.pc-btn') || {}).textContent || null,
      flagged: li.classList.contains('deviance'),
      dev: (li.querySelector('.pc-dev') || {}).textContent || null
    };
  });
  check('a receipt now stands on the row', !!fixed.receipt, true);
  // ⚠⚠ THE ONE THAT WAS BROKEN. A fixed row LOSES `.deviance` (that class
  // carries `&& !adj`), so the CSS rule which had been hiding this line stopped
  // applying and "should be 4" came BACK — re-asserting the stale figure right
  // next to the correction that supersedes it. The row read as still wrong.
  check('"should be" STAYS GONE after the fix', fixed.sys, null);
  check('...and the row is no longer flagged red', fixed.flagged, false);
  check('...and the deviance chip has stood down', fixed.dev, null);
  check('the count itself is still stated', fixed.counted, 'counted 1');

  console.log('\n── AGREES: count a shelf that matches (194244, hand 9) ─────');
  await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('194244') !== -1);
    li.querySelector('.pc-btn').click();
  });
  await page.waitForTimeout(400);
  await page.keyboard.press('9');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(1600);
  const agree = await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('194244') !== -1);
    return {
      sys: (li.querySelector('.pc-sys') || {}).textContent || null,
      counted: (li.querySelector('.pc-btn') || {}).textContent || null,
      flagged: li.classList.contains('deviance'),
      dev: (li.querySelector('.pc-dev') || {}).textContent || null,
      hasAdj: !!li.querySelector('.pc-adj')
    };
  });
  check('the count is recorded', agree.counted, 'counted 9');
  // The picker's own words: it said the same number twice.
  check('"should be 9" GONE beside "counted 9"', agree.sys, null);
  check('no deviance on a shelf that agreed', agree.flagged, false);
  check('no deviance chip', agree.dev, null);
  check('no fix button on a shelf that agreed', agree.hasAdj, false);

  console.log('\n── AND THE REFERENCE IS STILL ONE TAP AWAY ─────────────────');
  // Nothing may be LOST by hiding it: `counted N` IS the count button, and the
  // numpad it opens leads with the on-hand figure.
  await page.evaluate(() => {
    const li = [...document.querySelectorAll('.pick-row')]
      .find(r => (r.textContent || '').indexOf('194244') !== -1);
    li.querySelector('.pc-btn').click();
  });
  await page.waitForTimeout(450);
  const back = await page.evaluate(() => ({
    open: !document.getElementById('numPad').classList.contains('hidden'),
    ctx: (document.getElementById('npContext') || {}).textContent || ''
  }));
  check('numpad reopens from "counted N"', back.open, true);
  check('and it still states on hand', /On hand is\s*9/.test(back.ctx.replace(/\s+/g, ' ')), true);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(250);

  await page.screenshot({ path: path.join(OUT, 'count-model.png') });
  if (errs.length) { console.log('\n⚠ JS: ' + errs.join(' | ')); fails.push('js errors'); }
  await browser.close();

  console.log('\n' + '='.repeat(60));
  if (fails.length) { console.log('✗ ' + fails.join('\n✗ ')); process.exit(1); }
  console.log('✓ HAND is read, never projected: the reference is constant across');
  console.log('  the pick, and the Zoho suggestion is exactly what was counted.');
})().catch(e => { console.error('CRASH', e); process.exit(1); });
