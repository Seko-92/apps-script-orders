// Loads the REAL unpacked extension into Chromium and lets it poll the LIVE board
// (boardTick is read-only). Asserts the worker registers, reaches the server and
// paints a badge state. Run: node probe-alerts-extension.js
'use strict';
const path = require('path'), os = require('os'), fs = require('fs');
const { chromium } = require('playwright');
const EXT = path.join(__dirname, '..', 'deploy', 'hq-alerts-extension');
(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'hqext-'));
  const ctx = await chromium.launchPersistentContext(dir, { channel: 'chromium', headless: true,
    args: ['--disable-extensions-except=' + EXT, '--load-extension=' + EXT] });
  let [sw] = ctx.serviceWorkers();
  if (!sw) sw = await ctx.waitForEvent('serviceworker', { timeout: 15000 });
  const errs = []; sw.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });
  let st = {};
  for (let i = 0; i < 30 && !st.lastOk && !(st.fails >= 1); i++) {
    await new Promise(r => setTimeout(r, 1000));
    st = await sw.evaluate(() => chrome.storage.local.get(null));
  }
  const badge = await sw.evaluate(async () => ({ text: await chrome.action.getBadgeText({}), title: await chrome.action.getTitle({}) }));
  const alarms = await sw.evaluate(() => chrome.alarms.getAll());
  const id = sw.url().split('/')[2];
  const page = await ctx.newPage(); await page.goto('chrome-extension://' + id + '/popup.html'); await page.waitForTimeout(1500);
  await page.screenshot({ path: path.join(__dirname, 'renders', 'alerts-extension-popup.png') });
  const popup = await page.evaluate(() => document.getElementById('big').textContent + ' | ' + document.getElementById('sub').textContent);
  console.log(JSON.stringify({ worker: sw.url(), lastOk: !!st.lastOk, fails: st.fails, primed: st.primed,
    openHolds: st.openHolds, seenOrders: Object.keys((st.seen || {}).orders || {}).length, badge, alarms, popup, errs }, null, 1));
  await ctx.close();
  process.exit(st.lastOk && !errs.length ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
