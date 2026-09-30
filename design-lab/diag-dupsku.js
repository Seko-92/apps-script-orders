// ============================================================================
// ONE ORDER, THE SAME SKU TWICE — does every row stay under its own order?
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
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}

const row = (so, sku, loc, qty, note, status) => ({
  channel: 'DIRECT', orderId: so, sku, qty, location: loc, status: status || 'PREPARING',
  note: note || '', isKit: false, hand: 30
});

function fixture() {
  const t = JSON.parse(JSON.stringify(MOCK));
  const rows = [
    row('SO-25792', '173403', 'C-12', 1, 'Miguel'),
    row('SO-25792', '164979', 'C-42', 1, '↳ from KIT-157644 · Miguel'),   // ← the twin
    row('SO-25792', '164979', 'C-42', 2, 'Miguel'),                        // ← the twin
    row('SO-25792', '167067', 'C-84', 4, 'Miguel'),
    row('SO-25792', '158904', 'D-41', 1, 'Miguel'),
    row('SO-25982', '197430', 'C-60', 1, ''),
    row('SO-25982', '174843', 'H-5', 1, ''),
    row('SO-25985', '164529', 'B-59', 1, '', 'PENDING'),
    row('SO-25985', '164556', 'J-2', 2, '', 'PENDING')
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

/** Walk the list in DOM order: which order's band is each live row sitting under? */
const readLayout = page => page.evaluate(() => {
  const out = [];
  let band = null;
  document.querySelectorAll('.pick-head, .pick-row').forEach(el => {
    if (el.classList.contains('exit')) return;
    const k = el.getAttribute('data-key') || '';
    if (el.classList.contains('pick-head')) { band = k.split('|')[1]; return; }
    const parts = k.split('|');
    out.push({ order: String(parts[1] || '').replace(/#\d+$/, ''), sku: parts[2], under: band, key: k });
  });
  return out;
});

(async () => {
  const browser = await chromium.launch();
  console.log(`\nBOARD: ${BOARD}\n`);
  const { page, ctx, errs } = await boot(browser, fixture());

  // Open every DIRECT order so all rows are on screen.
  await page.evaluate(() => {
    document.querySelectorAll('.pick-head').forEach(h => {
      if (h.classList.contains('shut') || /▸/.test(h.textContent)) h.click();
    });
  });
  await page.waitForTimeout(600);

  // THE SECOND PAINT — where the orphan is born. Force two more reconciles.
  for (let n = 0; n < 2; n++) {
    await page.evaluate(() => { if (typeof pollSoon === 'function') pollSoon(); else if (typeof poll === 'function') poll(); });
    await page.waitForTimeout(1500);
  }

  const lay = await readLayout(page);
  const misplaced = lay.filter(r => r.under && r.under !== r.order);
  const twins = lay.filter(r => r.sku === '164979');

  check('both 164979 lines are on screen', twins.length, 2);
  // ⭐ THE HEADLINE — the floor's screenshot.
  check('no row sits under another order\'s band', misplaced.map(r => r.sku + ' under ' + r.under), []);
  check('the twins sit under SO-25792', [...new Set(twins.map(r => r.under))], ['SO-25792']);
  check('every row key is unique', lay.length === new Set(lay.map(r => r.key)).size, true);
  check('row count matches the data', lay.length, 9);
  check('no page errors', errs, []);
  await ctx.close();
  await browser.close();

  console.log('\n' + (failures.length ? `✗ ${failures.length} FAILED` : '✓ ALL PASSED'));
  process.exit(failures.length ? 1 : 0);
})();
