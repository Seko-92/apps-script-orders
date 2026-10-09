// Event films (2026-10-09) — drives the REAL Sidebar.html with fake ticks and checks which
// film each change queues. Renders frames to renders/films-*.png for the eye.
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const SRC = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");
let pass = 0, fail = 0;
const ok = (n, c, got) => { if (c) { pass++; console.log('PASS ' + n); } else { fail++; console.log('FAIL ' + n + (got !== undefined ? ' → got ' + JSON.stringify(got) : '')); } };

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 900 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar'); await p.waitForTimeout(1500);
  await p.evaluate(() => {
    window._hqsShopOpen = () => true;
    try { dismissBootSplash(); } catch (e) {}
    window.T = function (over) {
      const base = { cockpit: { shippedToday: 10, receivedToday: 14, ebayGrab: 2, directGrab: 0, amazonGrab: 0, orderAgeMin: {}, timeline: [] },
                     openOrders: [], customers: {}, _publishedAt: new Date().toISOString() };
      const t = Object.assign(base, over || {}); t.cockpit = Object.assign(base.cockpit, (over || {}).cockpit || {});
      return t;
    };
    window.feed = function (t) {
      HQFilm._stages.forEach(s => { s.q = []; s.cur = null; });
      try { HQFilm.noteTick(t.cockpit, t); } catch (e) { return 'ERR ' + e; }
      return HQFilm._stages[0].q.map(j => j.kind + '|' + j.detail);
    };
  });
  const row = (id, ch, st, loc, note) => ({ orderId: id, channel: ch, status: st, location: loc, sku: '1', qty: 1, note: note || '' });
  const E1 = row('24-14979-87359', 'EBAY', 'PENDING', 'C-31');

  // A · cold start announces nothing
  let q = await p.evaluate(e1 => feed(T({ openOrders: [e1], cockpit: { orderAgeMin: { '24-14979-87359': 3 } } })), E1);
  ok('cold start: nothing queued', q.length === 0, q);

  // B · one new eBay order
  q = await p.evaluate(([e1, e2]) => feed(T({ openOrders: [e1, e2], cockpit: { orderAgeMin: { '24-14979-87359': 4, '07-15276-84722': 1 } } })),
                       [E1, row('07-15276-84722', 'EBAY', 'PENDING', 'L-159')]);
  ok('one new eBay order → land', q.length === 1 && /^land\|07-15276-84722 · L-159/.test(q[0]), q);

  // C · the same order again is not news
  q = await p.evaluate(([e1, e2]) => feed(T({ openOrders: [e1, e2], cockpit: { orderAgeMin: { '24-14979-87359': 5, '07-15276-84722': 2 } } })),
                       [E1, row('07-15276-84722', 'EBAY', 'PENDING', 'L-159')]);
  ok('seen order: nothing queued', q.length === 0, q);

  // D · a batch of three → ONE film
  q = await p.evaluate(() => feed(T({ openOrders: ['a', 'b', 'c'].map((x, i) => ({ orderId: '1' + i + '-1-' + x, channel: 'EBAY', status: 'PENDING', location: 'F-' + i })),
    cockpit: { orderAgeMin: { '10-1-a': 1, '11-1-b': 1, '12-1-c': 1 } } })));
  ok('eBay batch → one landMany film', q.length === 1 && q[0].startsWith('landMany|3 eBay orders landed'), q);

  // E · an old order scrolling into view is not an arrival
  q = await p.evaluate(() => feed(T({ openOrders: [{ orderId: '99-old', channel: 'EBAY', status: 'PENDING', location: 'A-1' }], cockpit: { orderAgeMin: { '99-old': 300 } } })));
  ok('order older than 20 min: not announced', q.length === 0, q);

  // F · Direct + Amazon landings
  q = await p.evaluate(() => feed(T({ customers: { 'SO-26200': 'Triple M Equipment' },
    openOrders: [{ orderId: 'SO-26200', channel: 'DIRECT', status: 'PENDING', location: 'E-12' },
                 { orderId: 'AMZ-114-4402117', channel: 'AMAZON', status: 'PENDING', location: 'F-40', note: 'ship by 10/12' }],
    cockpit: { orderAgeMin: { 'SO-26200': 2, 'AMZ-114-4402117': 2 } } })));
  ok('Direct order → landD with customer', q.some(x => /^landD\|SO-26200 · Triple M · E-12/.test(x)), q);
  ok('Amazon order → landA with ship-by', q.some(x => /^landA\|AMZ-114-44021… · ship by 10\/12/.test(x)), q);

  // G · box ready: 2-line order goes all-picked; a 1-line order never does
  await p.evaluate(() => feed(T({ openOrders: [{ orderId: 'SO-1', channel: 'DIRECT', status: 'PENDING' }, { orderId: 'SO-1', channel: 'DIRECT', status: 'PREPARING' },
                                               { orderId: 'SO-2', channel: 'DIRECT', status: 'PENDING' }] })));
  q = await p.evaluate(() => feed(T({ openOrders: [{ orderId: 'SO-1', channel: 'DIRECT', status: 'PREPARING' }, { orderId: 'SO-1', channel: 'DIRECT', status: 'PREPARING' },
                                                   { orderId: 'SO-2', channel: 'DIRECT', status: 'PREPARING' }] })));
  ok('2-line order all picked → ready', q.some(x => x === 'ready|SO-1 · all 2 picked'), q);
  ok('1-line order picked → no ready film', !q.some(x => /SO-2/.test(x)), q);

  // H · shipped names the newest SHIPPED order
  q = await p.evaluate(() => feed(T({ cockpit: { shippedToday: 12, ebayGrab: 1, timeline: [
    { event: 'SHIPPED', orderId: '08-15272-44190', hourFraction: 10.5 }, { event: 'SHIPPED', orderId: '08-15272-44190', hourFraction: 10.5 },
    { event: 'SHIPPED', orderId: '01-old', hourFraction: 9 }] } })));
  ok('shippedToday rises → ship film for the newest order', q.length === 1 && q[0] === 'ship|08-15272-44190 · 2 lines', q);

  // I · the floor clears → caught up (wins over ship in the same tick)
  q = await p.evaluate(() => feed(T({ cockpit: { shippedToday: 13, ebayGrab: 0, directGrab: 0, amazonGrab: 0 } })));
  ok('to-grab reaches 0 → caught film', q.length === 1 && q[0] === 'caught|13 out today', q);

  // J · switched off → nothing
  q = await p.evaluate(() => { HQFilm.settings(false, true); const r = feed(T({ cockpit: { shippedToday: 20, ebayGrab: 3 } })); HQFilm.settings(true, true); return r; });
  ok('films off: nothing queued', q.length === 0, q);

  // K · frames, true size
  await p.evaluate(() => HQFilm._stages.forEach(s => { s.q = []; s.cur = null; }));
  const shoot = async (kind, detail, t, name) => {
    await p.evaluate(([k, d]) => HQFilm.play(k, d, k), [kind, detail]);
    await p.waitForTimeout(t);
    const bx = await p.locator('#opsCockpit').boundingBox();
    await p.screenshot({ path: path.join(__dirname, 'renders', 'films-' + name + '.png'), clip: { x: 0, y: bx.y, width: 310, height: bx.height } });
    await p.evaluate(() => HQFilm._stages.forEach(s => { s.q = []; s.cur = null; }));
    await p.waitForTimeout(150);
  };
  await shoot('land', '07-15276-84722 · L-159 C-58', 1700, 'land');
  await shoot('landMany', '5 eBay orders landed|F-28 C-31 E-27 H-35 B-45', 2000, 'batch');
  await shoot('landD', 'SO-26200 · Triple M · E-12', 2300, 'direct');
  await shoot('landA', 'AMZ-114-44021… · ship by 10/12', 2300, 'amazon');
  await shoot('ready', 'SO-26164 · all 10 picked', 2500, 'ready');
  await shoot('ship', '08-15272-44190 · 2 lines', 1500, 'ship');
  await shoot('caught', '43 out today', 2700, 'caught');

  // L · condensed cockpit stops drawing
  const condensed = await p.evaluate(async () => {
    document.body.classList.add('cockpit-condensed'); HQFilm.play('ship', 'x', 'Shipped');
    await new Promise(r => setTimeout(r, 400));
    const s = HQFilm._stages[0]; const r = { cur: !!s.cur, drawn: s.drawn };
    document.body.classList.remove('cockpit-condensed'); return r;
  });
  ok('folded cockpit: nothing drawing', !condensed.cur && !condensed.drawn, condensed);

  // M · the resting panel plays without waking
  const rest = await p.evaluate(async () => {
    _rpShow(); await new Promise(r => setTimeout(r, 1600));
    HQFilm.play('ship', '08-15272 · 2 lines', 'Shipped');
    await new Promise(r => setTimeout(r, 1500));
    return { shown: document.getElementById('restPanel').classList.contains('rp-show'), drawn: HQFilm._stages[1].drawn };
  });
  ok('resting panel: film plays, panel stays at rest', rest.shown && rest.drawn, rest);
  await p.screenshot({ path: path.join(__dirname, 'renders', 'films-rest.png') });

  ok('no page errors', errs.length === 0, errs.slice(0, 3));
  console.log((fail ? '✗ ' : '✓ all · ') + pass + ' passed' + (fail ? ', ' + fail + ' failed' : ''));
  await b.close(); process.exit(fail ? 1 : 0);
})();
