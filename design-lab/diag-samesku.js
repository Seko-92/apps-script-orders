// ============================================================================
// ONE ORDER, THE SAME SKU TWICE — does ✓ Pick ASK before marking both?
// (diag-samesku.js, 2026-10-01; fixture shape copied from diag-dupsku.js)
// The server flips every open row matching {orderId, sku}. Until picking one line
// alone is worth a server change, the board must show both lines and ask.
//
// Floor report 2026-09-30, with screenshots: a line "C-42 · 164979 · Miguel"
// hung under SO-25985 though it belongs to SO-25792. Expanding and collapsing
// an order put it right.
//
// SO-25792 carries SKU 164979 twice — once from kit 157644 (×1) and once loose
// (×2). Rows were keyed channel|order|sku, so both got the SAME key. The keyed
// reconcile indexes the screen by key, so on the next repaint two rows matched
// ONE node and the other node was orphaned: never reused, never removed, pushed
// down the list as everything else was placed in front of it.
//
// Checked on the SECOND paint on purpose — the first paint builds every row
// fresh and looks perfect; the orphan only appears when the list is reconciled.
//
// Usage: node diag-dupsku.js
// Before/after:  git show HEAD:FloorBoard.html > /tmp/board-before.html
//                BOARD_FILE=/tmp/board-before.html node diag-dupsku.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');

const failures = [];
const CALLS = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}

const row = (so, sku, loc, qty, note, status) => ({
  channel: 'DIRECT', orderId: so, sku, qty, location: loc, status: status || 'PENDING',
  note: note || '', isKit: false, hand: 30
});

function fixture() {
  const t = JSON.parse(JSON.stringify(MOCK));
  const rows = [
    row('SO-25792', '173403', 'C-12', 1, 'Miguel'),
    row('SO-25792', '164979', 'C-42', 1, '↳ from KIT-157644 · Miguel'),   // ← the twin
    row('SO-25792', '164979', 'C-42', 2, 'Miguel'),                        // ← the twin
    row('SO-25982', '197430', 'C-60', 1, ''),
    row('SO-25982', '197430', 'C-60', 1, '', 'PREPARING'),                 // twin already picked
  ];
  t.openOrders = rows;
  t.openOrdersTotal = rows.length;
  t.openOrdersBy = { EBAY: 0, DIRECT: rows.length, AMAZON: 0 };
  t.kits = [];
  return t;
}

async function boot(browser, tick) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const ctx = await browser.newContext({ viewport: { width: 1340, height: 930 }, hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  await page.route('http://hqlab.test/**', route => {
    const url = route.request().url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(route.request().postData() || '{}');
      let res = { ok: false };
      if (body.action === 'boardTick')  res = Object.assign({ ok: true }, tick);
      if (body.action === 'boardRadio') res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardStatus') { CALLS.push(body.sku); res = { ok: true }; }
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan\.com|open-meteo\.com/, r => r.abort());
  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.waitForFunction(() => !document.getElementById('board').classList.contains('booting'), null, { timeout: 20000 })
    .catch(() => errs.push('board never left booting'));
  await page.waitForTimeout(1200);
  return { page, ctx, errs };
}


const pickBtn = (page, so, sku, n) => page.evaluate(([so, sku, n]) => {
  const bs = [...document.querySelectorAll('.pick-do')].filter(b => b.getAttribute('data-order') === so && b.getAttribute('data-sku') === sku);
  if (!bs[n || 0]) return false; bs[n || 0].click(); return true;
}, [so, sku, n]);
const drawer = page => page.evaluate(() => ({
  open: typeof drwOpen !== 'undefined' && !!drwOpen,
  title: document.getElementById('drwTitle').textContent,
  lines: [...document.querySelectorAll('#drwBody .same-line')].map(l => [...l.children].map(c => c.textContent.trim()).join(' '))
}));
// ⚠ fails SOFT: against an old board the buttons do not exist, and a throw here
//   would hide every later section (the choosePicker lesson).
const tap = (page, sel) => page.evaluate(sel => { const e = document.querySelector(sel); if (e) e.click(); return !!e; }, sel);

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);
  const { page, ctx, errs } = await boot(browser, fixture());
  await page.evaluate(() => { document.querySelectorAll('.pick-head').forEach(h => { if (h.classList.contains('shut') || /▸/.test(h.textContent)) h.click(); }); });
  await page.waitForTimeout(600);

  // A · the twin asks, and writes nothing
  check('found the ✓ Pick on a twin line', await pickBtn(page, 'SO-25792', '164979', 1), true);
  await page.waitForTimeout(300);
  let d = await drawer(page);
  check('the question names the count', d.title, 'This part is on 2 lines');
  check('both lines are shown, the kit part named', d.lines, ['C-42 ×1 kit part · 157644', 'C-42 ×2 on its own']);
  check('nothing was sent to the sheet yet', CALLS.length, 0);

  // B · "Not yet" leaves everything as it was
  await tap(page, '#sameNo'); await page.waitForTimeout(300);
  check('"Not yet" sends nothing', CALLS.length, 0);

  // C · "Mark all" sends ONE write for the SKU
  await pickBtn(page, 'SO-25792', '164979', 0); await page.waitForTimeout(300);
  await tap(page, '#sameYes'); await page.waitForTimeout(500);
  check('"Mark all" sends one write for that SKU', CALLS, ['164979']);

  // D · a single line does not ask
  CALLS.length = 0;
  await pickBtn(page, 'SO-25792', '173403', 0); await page.waitForTimeout(400);
  d = await drawer(page);
  check('a single line picks straight away', CALLS, ['173403']);
  check('...with no question', d.open, false);

  // E · a twin whose other line is already picked does not ask
  CALLS.length = 0;
  await pickBtn(page, 'SO-25982', '197430', 0); await page.waitForTimeout(400);
  check('an already-picked twin changes nothing, so no question', CALLS, ['197430']);

  check('no page errors', errs, []);
  await ctx.close(); await browser.close();
  console.log('\n' + (failures.length ? `✗ ${failures.length} FAILED` : '✓ ALL PASSED'));
  process.exit(failures.length ? 1 : 0);
})();
