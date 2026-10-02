// ============================================================================
// WHICH SCREEN IS THIS (2026-10-02)
//   A · every request carries device {id,name,role}; a touch device defaults
//       to role "floor"
//   B · the ⋯ menu names this screen and counts screens online from _devices
//   C · renaming + choosing a role persists across a reload, keeps the same id,
//       and the next request carries the new name
//   D · the "Screens" drawer lists every device from the tick
//   E · self-update: a changed page version reloads only when idle
// Usage: node diag-screens.js
// ============================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const MOCK = require('./mock-tick.js');
const failures = [];
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failures.push(label);
}
(async () => {
  const html = fs.readFileSync(BOARD, 'utf8');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  const errs = []; page.on('pageerror', e => errs.push(e.message));
  const seen = []; let etag = 'v1'; let loads = 0;
  await page.route('http://hqlab.test/**', async route => {
    const req = route.request(), url = req.url();
    if (url.includes('/api/board')) {
      const body = JSON.parse(req.postData() || '{}'); seen.push(body);
      let res = { ok: true };
      if (body.action === 'boardTick') {
        res = JSON.parse(JSON.stringify(MOCK)); res.picker = 'Shipping - Yassin 1';
        res._devices = [{ id: (body.device || {}).id, name: (body.device || {}).name, role: 'floor', ageSec: 3 },
                        { id: 'w1', name: 'Wall', role: 'wall', ageSec: 20 },
                        { id: 'o1', name: 'Office PC', role: 'office', ageSec: 900 }];
      }
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    if (req.method() === 'HEAD') return route.fulfill({ status: 200, headers: { etag }, body: '' });
    loads++;
    return route.fulfill({ contentType: 'text/html; charset=utf-8', headers: { etag }, body: html });
  });
  await page.goto('http://hqlab.test/');
  await page.waitForFunction(() => window.lastTick, null, { timeout: 15000 });

  console.log('A · an unnamed screen describes itself and counts as the warehouse');
  const t0 = seen.find(b => b.action === 'boardTick');
  check('device present', !!(t0 && t0.device && t0.device.id), true);
  check('role = warehouse (floor), no naming needed', t0.device.role, 'floor');
  check('name = what it physically is', t0.device.name, 'Linux · Chrome');

  console.log('B · menu: the floor line, and no naming nag');
  check('menu says This is my device · off', await page.textContent('#screenMeLbl'), 'This is my device · off');
  check('no amber cue anywhere', await page.evaluate(() => document.querySelectorAll('.needs-name').length), 0);
  check('floor line (no floorLast → QUIET; office PC silent 15 min → 1 on)', await page.textContent('#screensOnlineLbl'), 'Floor · QUIET · 1 on');

  console.log('B2 · This is my device');
  await page.click('#menuBtn'); await page.click('#screenMe');
  await page.waitForTimeout(400);
  const tm1 = seen.filter(b => b.action === 'boardTick').pop();
  check('toggle on → role remote, sent at once', [await page.evaluate(() => HQ_DEVICE.role), tm1.device.role], ['remote', 'remote']);
  await page.reload(); await page.waitForFunction(() => window.lastTick, null, { timeout: 15000 });
  check('survives a reload', await page.evaluate(() => HQ_DEVICE.role), 'remote');
  check('label says on', await page.textContent('#screenMeLbl'), 'This is my device · on');
  await page.click('#menuBtn'); await page.click('#screenMe');
  check('toggle off → floor again', await page.evaluate(() => HQ_DEVICE.role), 'floor');

  console.log('B3 · migration');
  await page.evaluate(() => localStorage.setItem('hqDevice', JSON.stringify({ id: 'dold1', name: 'Unnamed · Linux · Chrome', role: 'unset', named: false })));
  await page.reload(); await page.waitForFunction(() => window.lastTick, null, { timeout: 15000 });
  let tm = seen.filter(b => b.action === 'boardTick').pop();
  check('this morning\'s "Unnamed/unset" → plain label, warehouse', [tm.device.id, tm.device.name, tm.device.role], ['dold1', 'Linux · Chrome', 'floor']);
  await page.evaluate(() => localStorage.setItem('hqDevice', JSON.stringify({ id: 'dold2', name: 'Warehouse tablet', role: 'floor', named: true })));
  await page.reload(); await page.waitForFunction(() => window.lastTick, null, { timeout: 15000 });
  tm = seen.filter(b => b.action === 'boardTick').pop();
  check('a screen someone named keeps its name', [tm.device.name, tm.device.role], ['Warehouse tablet', 'floor']);

  console.log('C · same id across reloads');
  const tl = seen.filter(b => b.action === 'boardTick').pop();
  check('id stable', tl.device.id, 'dold2');

  console.log('D · screens drawer');
  await page.click('#menuBtn'); await page.click('#screensOnline');
  const txt = await page.textContent('#drwBody');
  check('floor summary + all three screens', ['QUIET', 'Wall', 'Office PC', '(this one)', 'warehouse'].every(s => txt.includes(s)), true);
  await page.click('#drwX');

  console.log('E · self-update');
  const before = loads;
  etag = 'v2';
  await page.evaluate(() => { selfVersionCheck(); });
  await page.waitForTimeout(300);
  await page.evaluate(() => { lastHumanMs = Date.now(); selfUpdateIfIdle(); });
  await page.waitForTimeout(500);
  check('busy (touched recently) → no reload', loads, before);
  await page.evaluate(() => { lastHumanMs = Date.now() - 120000; selfUpdateIfIdle(); });
  await page.waitForTimeout(1500);
  check('idle → reloads', loads, before + 1);

  check('no page errors', errs, []);
  await browser.close();
  console.log(failures.length ? `\n✗ ${failures.length} failure(s)` : '\n✓ all green');
  process.exit(failures.length ? 1 : 0);
})();
