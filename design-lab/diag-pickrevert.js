// ============================================================================
// PICKED ON THE FLOOR, REVERTED IN THE SHEET (floor test 2026-10-02)
// ✓ Pick on the tablet, then the sheet set back to PENDING within seconds. No
// tick ever said PREPARING, so the optimistic override waited out its ~2-min
// TTL. Fix: a tick PUBLISHED AFTER the server confirmed the write (server Date
// header) retires the override whatever it says.
//   A · a stale tick (published before the confirm) still cannot snap the row
//       back — the original reason the override exists         [regression net]
//   B · a tick published after the confirm that says PENDING shows PENDING
//       immediately                                             [FAILS on HEAD]
//   C · with no Date header, behaviour is the old TTL one       [regression net]
// Usage: node diag-pickrevert.js   ·   BOARD_FILE=/tmp/old.html for before/after
// ============================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const ORDER = '24-15021-77421', SKU = '194244';
const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}
async function run(withDate) {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = []; page.on('pageerror', e => errs.push(e.message));
  const st = { status: 'PENDING', publishedAt: Date.now() - 30000, confirmAt: 0 };
  await page.route('http://hqlab.test/**', async route => {
    const req = route.request(), url = req.url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(req.postData() || '{}');
      let res = { ok: true };
      if (body.action === 'boardTick') {
        res = JSON.parse(JSON.stringify(MOCK)); res.picker = 'Shipping - Yassin 1';
        res.openOrders.forEach(r => { if (r.orderId === ORDER && r.sku === SKU) r.status = st.status; });
        res._publishedAt = new Date(st.publishedAt).toISOString();
      }
      if (body.action === 'boardStatus') { st.confirmAt = Date.now(); res = { ok: true, count: 1 }; }
      const headers = withDate ? { date: new Date(st.confirmAt || Date.now()).toUTCString() } : {};
      return route.fulfill({ contentType: 'application/json', headers, body: JSON.stringify(res) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.goto('http://hqlab.test/');
  await page.waitForFunction(() => window.lastTick, null, { timeout: 15000 });
  const rowState = () => page.evaluate(([o, s]) => {
    const rows = applyPickOverrides((lastTick.openOrders || []).slice(), Date.parse(lastTick._publishedAt || '') || 0);
    const r = rows.find(x => x.orderId === o && x.sku === s); return r && r.status;
  }, [ORDER, SKU]);
  // tap ✓ Pick (server confirms; sheet still PENDING on the published copy)
  await page.evaluate(([o, s]) => markPicked(o, s, null), [ORDER, SKU]);
  await page.waitForTimeout(600);
  // A: a tick published BEFORE the confirm, still PENDING
  st.publishedAt = st.confirmAt - 5000;
  await page.evaluate(() => pollSoon()); await page.waitForTimeout(900);
  const a = await rowState();
  // B: sheet set back to PENDING, published 5s AFTER the confirm
  st.publishedAt = st.confirmAt + 5000;
  await page.evaluate(() => pollSoon()); await page.waitForTimeout(900);
  const b = await rowState();
  await browser.close();
  return { a, b, errs };
}
(async () => {
  console.log('WITH a server Date header (production)');
  let r = await run(true);
  check('A · stale tick does not snap the pick back', r.a, 'PREPARING');
  check('B · post-confirm tick saying PENDING wins at once', r.b, 'PENDING');
  check('no page errors', r.errs, []);
  console.log('WITHOUT a Date header');
  r = await run(false);
  check('C · falls back to the TTL (still PREPARING)', r.b, 'PREPARING');
  console.log(failures.length ? `\n✗ ${failures.length} failure(s)` : '\n✓ all green');
  process.exit(failures.length ? 1 : 0);
})();
