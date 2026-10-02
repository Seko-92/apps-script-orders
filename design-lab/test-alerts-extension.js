// HQ Alerts Chrome extension (2026-10-02) — runs the REAL background.js + core.js
// in a VM with stubbed chrome.* APIs and a scripted sequence of ticks. Mirrors the
// /alerts tab's scenarios (test-alerts.js) so the two surfaces are proven to agree,
// plus what only the extension has: the badge, restarts, and a broken server.
// Run: node test-alerts-extension.js
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const DIR = path.join(__dirname, '..', 'deploy', 'hq-alerts-extension');
const bg = fs.readFileSync(path.join(DIR, 'background.js'), 'utf8');
const core = fs.readFileSync(path.join(DIR, 'core.js'), 'utf8');
let pass = 0, fail = 0;
const ok = (n, c, g) => { if (c) pass++; else { fail++; console.log('FAIL', n, '→', JSON.stringify(g)); } };

const row = (ch, id, sku, loc, st, extra) => Object.assign({ channel: ch, orderId: id, sku, qty: 1, location: loc, status: st || 'PENDING', note: '' }, extra || {});
const base = { cockpit: { orderAgeMin: {} }, openOrders: [row('EBAY', '01-1-1', '100', 'L-67')], held: [], customers: {} };
const T = o => Object.assign(JSON.parse(JSON.stringify(base)), o);

// One extension "install" on one PC. `store` survives a browser restart.
function boot(at, store) {
  const S = { store: store || {}, sent: [], badge: {}, alarms: [], clock: new Date(at).getTime(), tick: null, httpFail: false,
              notifyFails: false, listeners: {} };
  const on = k => ({ addListener: f => { S.listeners[k] = f; } });
  const RealDate = Date;
  class FakeDate extends RealDate { constructor(...a) { super(...(a.length ? a : [S.clock])); } static now() { return S.clock; } }
  const chrome = {
    runtime: { onInstalled: on('installed'), onStartup: on('startup'), onMessage: on('message'), lastError: null },
    alarms: { create: (n, o) => S.alarms.push([n, o]), onAlarm: on('alarm') },
    storage: { local: {
      get: keys => Promise.resolve(keys.reduce((o, k) => { if (k in S.store) o[k] = JSON.parse(JSON.stringify(S.store[k])); return o; }, {})),
      set: obj => { Object.assign(S.store, JSON.parse(JSON.stringify(obj))); return Promise.resolve(); } } },
    notifications: {
      create: (id, o, cb) => { if (S.notifyFails) { chrome.runtime.lastError = { message: 'denied' }; cb(); chrome.runtime.lastError = null; return; }
                              S.sent.push({ id, title: o.title, body: o.message, sticky: o.requireInteraction, icon: o.iconUrl }); cb(id); },
      onClicked: on('click'), clear: () => {} },
    action: { setBadgeText: o => { S.badge.text = o.text; }, setBadgeBackgroundColor: o => { S.badge.bg = o.color; },
              setBadgeTextColor: () => {}, setTitle: o => { S.badge.title = o.title; } },
    tabs: { query: (q, cb) => cb([]), create: o => { S.opened = o.url; }, update: () => {} }, windows: { update: () => {} }
  };
  const ctx = {
    chrome, Date: FakeDate, Promise, JSON, String, Math, parseInt, parseFloat, isFinite, Object, Array, setTimeout, clearTimeout,
    AbortController, Error, console,
    fetch: () => S.httpFail ? Promise.reject(new Error('down'))
                            : Promise.resolve({ ok: true, json: () => Promise.resolve(JSON.parse(JSON.stringify(S.tick))) }),
    importScripts: () => vm.runInContext(core, ctx)
  };
  ctx.self = ctx; vm.createContext(ctx);
  vm.runInContext(bg, ctx);
  S.ctx = ctx;
  return S;
}
const flush = () => new Promise(r => setTimeout(r, 30));
async function step(S, tick, ms) { S.clock += ms || 30000; S.tick = tick; S.listeners.alarm({ name: 'poll' }); await flush(); }
async function start(S, tick, kind) { S.tick = tick; S.listeners[kind || 'installed'](); await flush(); }

(async () => {
  const WORK = '2026-10-05T10:00:00-05:00', NIGHT = '2026-10-05T19:30:00-05:00';
  const ebay11 = []; for (let n = 0; n < 11; n++) ebay11.push(row('EBAY', '2' + n + '-9-9', String(200 + n), String.fromCharCode(70 - (n % 5)) + '-' + (n + 3)));
  const direct = row('DIRECT', 'SO-26007', '155853', 'H-21', 'PENDING', { qty: 3 });
  const amz = row('AMAZON', 'AMZ-114-1', '163962', 'E-15', 'PENDING', { note: 'ship by Fri 10/9' });
  const hold = { orderId: '16-15228-99781', channel: 'EBAY', acked: false, shipped: true, urgent: true, items: [{ sku: '160245', qty: 1, loc: 'L-111' }] };
  const busy = base.openOrders.concat(ebay11, [direct, amz]);

  // A–E: install seeds silently, burst + Direct + Amazon, hold, hold acknowledged
  let S = boot(WORK);
  await start(S, T({}));
  ok('A install announces nothing', S.sent.length === 0, S.sent);
  ok('A alarm every 30 s', S.alarms.length === 1 && S.alarms[0][1].periodInMinutes === 0.5, S.alarms);
  ok('A badge clear', S.badge.text === '' && /all clear/.test(S.badge.title), S.badge);
  await step(S, T({ openOrders: busy, customers: { 'SO-26007': 'Camden Fresno' } }));
  const s1 = S.sent.slice();
  ok('B eBay burst is ONE notification', s1.filter(x => /eBay/.test(x.title)).length === 1, s1.map(x => x.title));
  ok('B burst title counts 11', s1.some(x => x.title === '11 new eBay orders'), s1.map(x => x.title));
  ok('B burst closes by itself', s1.filter(x => /eBay/.test(x.title)).every(x => !x.sticky), s1);
  ok('B first walk in aisle order', /First walk B-7 · B-12 · C-6/.test((s1.find(x => /eBay/.test(x.title)) || {}).body || ''), s1);
  const d = s1.find(x => /Direct/.test(x.title)) || {};
  ok('C Direct sticky, names shelf and customer', d.sticky && /H-21 · 155853 ×3/.test(d.body) && /Camden Fresno/.test(d.body), d);
  const a = s1.find(x => /Amazon/.test(x.title)) || {};
  ok('C Amazon sticky with ship-by', a.sticky && /ship by Fri 10\/9/.test(a.body), a);
  ok('C badge counts the 13 new orders', S.badge.text === '13', S.badge);
  await step(S, T({ openOrders: busy, held: [hold] }));
  const s2 = S.sent.slice(s1.length);
  ok('D hold announced once, sticky, red icon', s2.length === 1 && s2[0].sticky && /label already bought/.test(s2[0].title) && /hold/.test(s2[0].icon), s2);
  ok('D badge turns red with the hold', S.badge.bg === '#d32f2f' && S.badge.text === '1' && /16-15228-99781/.test(S.badge.title), S.badge);
  await step(S, T({ openOrders: busy, held: [hold] }));
  ok('D same hold next poll → nothing new', S.sent.length === s1.length + 1, S.sent.length);
  await step(S, T({ openOrders: busy, held: [Object.assign({}, hold, { acked: true })] }));
  ok('E acknowledged → no notification', S.sent.length === s1.length + 1, S.sent.length);
  ok('E acknowledged → badge falls back to the new count', S.badge.text === '13', S.badge);
  // popup opened → new orders count as seen
  await new Promise(r => S.listeners.message({ cmd: 'seen' }, null, r));
  ok('E popup clears the new-order badge', S.badge.text === '', S.badge);
  ok('E last-sent log kept', (S.store.log || []).length === 4 && /Hold/.test(S.store.log[0].x), S.store.log);

  // F: off hours — orders silent, holds still speak
  S = boot(NIGHT); await start(S, T({}));
  await step(S, T({ openOrders: base.openOrders.concat([direct]), held: [hold] }));
  ok('F off hours: no new-order notification', !S.sent.some(x => /order/i.test(x.title)), S.sent.map(x => x.title));
  ok('F off hours: the hold still notifies', S.sent.some(x => /Hold/.test(x.title)), S.sent.map(x => x.title));

  // G: an old order scrolling into view is not an arrival
  S = boot(WORK); await start(S, T({}));
  await step(S, T({ openOrders: base.openOrders.concat([row('EBAY', '33-3-3', '300', 'A-1')]), cockpit: { orderAgeMin: { '33-3-3': 95 } } }));
  ok('G 95-minute-old order is not news', S.sent.length === 0, S.sent);

  // H: browser restart — remembered ids are not re-announced; open orders at start are not news
  S = boot(WORK); await start(S, T({}));
  await step(S, T({ openOrders: base.openOrders.concat([direct]) }));
  const H = boot(WORK, S.store); await start(H, T({ openOrders: base.openOrders.concat([direct, amz]) }), 'startup');
  await step(H, T({ openOrders: base.openOrders.concat([direct, amz]) }));
  ok('H restart does not re-announce, and orders open at start are not news', H.sent.length === 0, H.sent);
  // …but the worker going to sleep and waking (no onStartup) stays primed
  await step(H, T({ openOrders: base.openOrders.concat([direct, amz, row('DIRECT', 'SO-26008', '1', 'A-2')]) }));
  ok('H worker wake keeps announcing', H.sent.length === 1 && /SO-26008/.test(H.sent[0].title), H.sent);

  // I: a PREPARING-only order is not an arrival
  S = boot(WORK); await start(S, T({}));
  await step(S, T({ openOrders: base.openOrders.concat([row('EBAY', '44-4-4', '400', 'A-1', 'PREPARING')]) }));
  ok('I a PREPARING-only order is not news', S.sent.length === 0, S.sent);

  // J: Chrome refuses the notification → hold not remembered, announced when it can be
  S = boot(WORK); await start(S, T({}));
  S.notifyFails = true; await step(S, T({ held: [hold] }));
  ok('J nothing logged while Chrome refuses', !(S.store.log || []).length && !S.store.seen.holds[hold.orderId], S.store);
  S.notifyFails = false; await step(S, T({ held: [hold] }));
  ok('J hold announced once Chrome allows it', S.sent.length === 1 && /Hold/.test(S.sent[0].title), S.sent);

  // K: server unreachable twice → grey "!" badge; recovers on the next good poll
  S = boot(WORK); await start(S, T({}));
  S.httpFail = true; await step(S, T({})); await step(S, T({}));
  ok('K two failures → offline badge', S.badge.text === '!' && /cannot reach/.test(S.badge.title), S.badge);
  S.httpFail = false; await step(S, T({}));
  ok('K recovers → badge clear', S.badge.text === '' && S.store.fails === 0, S.badge);

  // L: a tick with no cockpit is a failure, never "all clear"
  S = boot(WORK); await start(S, T({}));
  await step(S, { status: 'success', added: 0 }); await step(S, { status: 'success', added: 0 });
  ok('L a success-shaped non-tick reads as offline', S.badge.text === '!', S.badge);

  // M: clicking a notification opens the board
  S.listeners.click('hq-SO-26007'); await flush();
  ok('M click opens the board', S.opened === 'https://hq.yassinqurabi.com/', S.opened);

  console.log(pass + ' pass, ' + fail + ' fail');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
