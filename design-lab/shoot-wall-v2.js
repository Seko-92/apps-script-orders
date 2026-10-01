// /wall v2 (2026-10-02) — renders the PRODUCTION wall.html against the LIVE tick
// (renders/live-tick.json, fetched with curl) plus two variants, at two screen
// sizes, and asserts the page NEVER scrolls and no list overflows its column.
'use strict';
const fs = require('fs'); const path = require('path');
const { chromium } = require('playwright');
const html = fs.readFileSync(process.env.WALL_FILE || path.join(__dirname, '..', 'wall.html'), 'utf8');
const LIVE = JSON.parse(fs.readFileSync(path.join(__dirname, 'renders/live-tick.json'), 'utf8'));
const clone = o => JSON.parse(JSON.stringify(o));

function busy() {
  const t = clone(LIVE);
  t.openOrders.unshift(
    { channel: 'EBAY', orderId: '04-15301-22817', sku: '167517', qty: 2, location: 'B-12', status: 'PENDING', note: 'Buyer Note: Please Blind Drop shipped this order no INVOICES/ PAPER WORK INSIDE THE BOX !', hand: 9 },
    { channel: 'EBAY', orderId: '11-15238-66804', sku: '173079', qty: 1, location: 'E-37', status: 'PREPARING', note: '', hand: 4 },
    { channel: 'AMAZON', orderId: 'AMZ-114-3958271-0472616', sku: '163962', qty: 2, location: 'E-15', status: 'PENDING', note: '', hand: 23 });
  for (let i = 0; i < 14; i++) t.openOrders.push({ channel: 'EBAY', orderId: '2' + i + '-1530' + i + '-1188' + i, sku: String(160000 + i * 37), qty: 1 + (i % 3), location: String.fromCharCode(65 + i) + '-' + (10 + i), status: 'PENDING', note: i === 3 ? 'call before shipping' : '' });
  t.openOrdersBy = { EBAY: 18, DIRECT: 16, AMAZON: 1 };
  t.cockpit.orderAgeMin = Object.assign({}, t.cockpit.orderAgeMin, { '04-15301-22817': 205 });
  t.cockpit.ebayGrab = 16; t.cockpit.amazonGrab = 1;
  return t;
}
function noHold(t) { t = clone(t); t.held = []; return t; }

const SCEN = [
  ['live-1509',  LIVE,          '2026-10-02T15:09:00-05:00'],
  ['busy-1510',  busy(),        '2026-10-02T15:10:00-05:00'],
  ['calm-1100',  noHold(LIVE),  '2026-10-02T11:00:00-05:00'],
  ['night-2010', noHold(LIVE),  '2026-10-02T20:10:00-05:00'],
];
const VIEWS = [[1920, 1080], [1366, 768]];

(async () => {
  const browser = await chromium.launch();
  let fail = 0;
  for (const [vw, vh] of VIEWS) for (const [name, tick, at] of SCEN) {
    const ctx = await browser.newContext({ viewport: { width: vw, height: vh }, timezoneId: 'America/Chicago' });
    const page = await ctx.newPage();
    await page.clock.setFixedTime(new Date(at));
    const errs = [];
    page.on('pageerror', e => errs.push(e.message));
    await page.route('http://hqlab.test/**', route => route.request().url().includes('/api/board')
      ? route.fulfill({ contentType: 'application/json', body: JSON.stringify(tick) })
      : route.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.route(/aladhan\.com/, r => r.abort());
    await page.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
    await page.goto('http://hqlab.test/wall', { waitUntil: 'load' });
    await page.waitForTimeout(1500);
    const st = await page.evaluate(() => {
      const d = document.documentElement;
      const over = [...document.querySelectorAll('.chan ul, #ntList, .rail, .tick')]
        .filter(e => getComputedStyle(e).display !== 'none' && (e.scrollHeight > e.clientHeight + 1 || e.scrollWidth > e.clientWidth + 1))
        .map(e => (e.id || e.className) + '[' + e.scrollWidth + '/' + e.clientWidth + ',' + e.scrollHeight + '/' + e.clientHeight + ']');
      return {
        pageScroll: d.scrollHeight > innerHeight + 1 || d.scrollWidth > innerWidth + 1,
        over,
        att: [...document.querySelectorAll('#attList .al-t, #attList .att-calm')].map(e => e.textContent.trim().slice(0, 40)),
        notes: document.querySelectorAll('#ntList .nt').length,
        slim: [...document.querySelectorAll('.chan.slim')].map(e => e.id),
        more: [...document.querySelectorAll('.wmore')].map(e => e.textContent),
        tick: document.querySelectorAll('.tick .tev').length,
        rest: document.getElementById('restveil').classList.contains('show'),
      };
    });
    const bad = st.pageScroll || st.over.length || errs.length;
    if (bad) fail++;
    console.log((bad ? 'FAIL ' : 'ok   ') + vw + 'x' + vh + ' ' + name + ' ' + JSON.stringify(st) + (errs.length ? ' ERR ' + errs.join(' | ') : ''));
    await page.screenshot({ path: `renders/wall-v2-${name}-${vw}.png` });
    await ctx.close();
  }
  await browser.close();
  console.log(fail ? fail + ' FAIL' : 'all ok');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
