// /alerts (2026-10-02) — drives the REAL alerts.html in Chromium with a recording
// Notification stub and a scripted sequence of ticks. Asserts what is announced,
// what is not, and the tab title. Run: node test-alerts.js   (ALERTS_FILE= to point elsewhere)
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const html = fs.readFileSync(process.env.ALERTS_FILE || path.join(__dirname, '..', 'alerts.html'), 'utf8');
let pass = 0, fail = 0;
const ok = (n, c, g) => { if (c) pass++; else { fail++; console.log('FAIL', n, '→', JSON.stringify(g)); } };

const row = (ch, id, sku, loc, st, extra) => Object.assign({ channel: ch, orderId: id, sku, qty: 1, location: loc, status: st || 'PENDING', note: '' }, extra || {});
const base = { cockpit: { orderAgeMin: {} }, openOrders: [row('EBAY', '01-1-1', '100', 'L-67')], held: [], customers: {} };
const T = o => Object.assign(JSON.parse(JSON.stringify(base)), o);

async function run(at, ticks, storage, permSeq) {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ timezoneId: 'America/Chicago' });
  const page = await ctx.newPage();
  await page.clock.install({ time: new Date(at) });
  let i = 0;
  await page.addInitScript(({ st, p0 }) => {
    window.__sent = [];
    window.Notification = function (title, o) { window.__sent.push({ title, body: o.body, sticky: !!o.requireInteraction, tag: o.tag }); this.close = function () {}; };
    window.Notification.permission = p0 || 'granted';
    window.Notification.requestPermission = () => Promise.resolve('granted');
    if (st) localStorage.setItem('hqAlertsSeen', st);
  }, { st: storage || null, p0: permSeq ? permSeq[0] : null });
  await page.route('http://hqlab.test/**', route => {
    if (route.request().url().includes('/api/board')) {
      const t = ticks[Math.min(i, ticks.length - 1)]; i++;
      return route.fulfill({ contentType: 'application/json', body: JSON.stringify(t) });
    }
    return route.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/fonts\./, r => r.abort());
  await page.goto('http://hqlab.test/alerts');
  const out = [];
  for (let k = 0; k < ticks.length; k++) {
    if (permSeq && permSeq[k]) await page.evaluate(p => { window.Notification.permission = p; }, permSeq[k]);
    await page.clock.runFor(k === 0 ? 500 : 21000);
    await page.waitForTimeout(150);
    out.push(await page.evaluate(() => ({ sent: window.__sent.slice(), title: document.title })));
  }
  const st = await page.evaluate(() => localStorage.getItem('hqAlertsSeen'));
  const errs = [];
  await browser.close();
  return { out, st };
}

(async () => {
  const WORK = '2026-10-05T10:00:00-05:00', NIGHT = '2026-10-05T19:30:00-05:00';
  const ebay11 = []; for (let n = 0; n < 11; n++) ebay11.push(row('EBAY', '2' + n + '-9-9', String(200 + n), String.fromCharCode(70 - (n % 5)) + '-' + (n + 3)));
  const direct = row('DIRECT', 'SO-26007', '155853', 'H-21', 'PENDING', { qty: 3 });
  const amz = row('AMAZON', 'AMZ-114-1', '163962', 'E-15', 'PENDING', { note: 'ship by Fri 10/9' });

  // A–D: seeding, burst + Direct + Amazon, hold, hold acknowledged
  const hold = { orderId: '16-15228-99781', channel: 'EBAY', acked: false, shipped: true, urgent: true, items: [{ sku: '160245', qty: 1, loc: 'L-111' }] };
  const r = await run(WORK, [
    T({}),
    T({ openOrders: base.openOrders.concat(ebay11, [direct, amz]), customers: { 'SO-26007': 'Camden Fresno' } }),
    T({ openOrders: base.openOrders.concat(ebay11, [direct, amz]), held: [hold] }),
    T({ openOrders: base.openOrders.concat(ebay11, [direct, amz]), held: [hold] }),
    T({ openOrders: base.openOrders.concat(ebay11, [direct, amz]), held: [Object.assign({}, hold, { acked: true })] }),
  ]);
  ok('A first tick announces nothing', r.out[0].sent.length === 0, r.out[0].sent);
  const s1 = r.out[1].sent;
  ok('B eBay burst is ONE notification', s1.filter(x => /eBay/.test(x.title)).length === 1, s1.map(x => x.title));
  ok('B burst title counts 11', s1.some(x => x.title === '11 new eBay orders'), s1.map(x => x.title));
  ok('B burst closes by itself', s1.filter(x => /eBay/.test(x.title)).every(x => !x.sticky), s1);
  ok('B first walk in aisle order', /First walk B-7 · B-12 · C-6/.test((s1.find(x => /eBay/.test(x.title)) || {}).body || ''), s1);
  const d = s1.find(x => /Direct/.test(x.title)) || {};
  ok('C Direct has its own sticky notification', d.title === 'New Direct order · SO-26007' && d.sticky, d);
  ok('C Direct names shelf and customer', /H-21 · 155853 ×3/.test(d.body) && /Camden Fresno/.test(d.body), d.body);
  const a = s1.find(x => /Amazon/.test(x.title)) || {};
  ok('C Amazon sticky with ship-by', a.sticky && /ship by Fri 10\/9/.test(a.body), a);
  const s2 = r.out[2].sent.slice(s1.length);
  ok('D hold announced once, sticky', s2.length === 1 && s2[0].sticky && /label already bought/.test(s2[0].title), s2);
  ok('D hold body names shelf', /L-111/.test((s2[0] || {}).body || ''), s2);
  ok('D tab title names the hold', r.out[2].title === '⚠ HOLD · 16-15228-99781', r.out[2].title);
  ok('D same hold next tick → nothing new', r.out[3].sent.length === r.out[2].sent.length, r.out[3].sent.length);
  ok('E acknowledged → no notification', r.out[4].sent.length === r.out[3].sent.length, r.out[4].sent.length);
  ok('E acknowledged → title clears the hold', !/HOLD/.test(r.out[4].title), r.out[4].title);

  // F: off hours — orders silent, holds still speak
  const n = await run(NIGHT, [T({}), T({ openOrders: base.openOrders.concat([direct]), held: [hold] })]);
  const ns = n.out[1].sent;
  ok('F off hours: no new-order notification', !ns.some(x => /order/i.test(x.title)), ns.map(x => x.title));
  ok('F off hours: the hold still notifies', ns.some(x => /Hold/.test(x.title)), ns.map(x => x.title));

  // G: an old order scrolling into view is not an arrival
  const g = await run(WORK, [T({}), T({ openOrders: base.openOrders.concat([row('EBAY', '33-3-3', '300', 'A-1')]), cockpit: { orderAgeMin: { '33-3-3': 95 } } })]);
  ok('G 95-minute-old order is not news', g.out[1].sent.length === 0, g.out[1].sent);

  // H: reload — remembered ids are not announced again, even after the first tick
  const h1 = await run(WORK, [T({}), T({ openOrders: base.openOrders.concat([direct]) })]);
  const h2 = await run(WORK, [T({ openOrders: base.openOrders.concat([direct]) }), T({ openOrders: base.openOrders.concat([direct]) })], h1.st);
  ok('H reload does not re-announce', h2.out[1].sent.length === 0, h2.out[1].sent);

  // I: a PREPARING order appearing (picked elsewhere first) is not an arrival
  const i2 = await run(WORK, [T({}), T({ openOrders: base.openOrders.concat([row('EBAY', '44-4-4', '400', 'A-1', 'PREPARING')]) })]);
  ok('I a PREPARING-only order is not news', i2.out[1].sent.length === 0, i2.out[1].sent);

  // J: a hold that lands while notifications are OFF is announced once they are turned on,
  //    and nothing is logged as sent before that
  const j = await run(WORK, [T({}), T({ held: [hold] }), T({ held: [hold] })], null, ['default', 'default', 'granted']);
  ok('J nothing shown while off', j.out[1].sent.length === 0, j.out[1].sent);
  ok('J hold announced once allowed', j.out[2].sent.length === 1 && /Hold/.test(j.out[2].sent[0].title), j.out[2].sent);

  console.log(pass + ' pass, ' + fail + ' fail');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
