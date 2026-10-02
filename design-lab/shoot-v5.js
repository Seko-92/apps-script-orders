// shoot-v5.js — the sidebar BEFORE and AFTER the v5 "one soul with the sheet" pass, same fake
// data: the head (cockpit), the card list with an OPEN card beside a section band, and the
// resting panel. Writes renders/v5-{before,after}-{head,cards,rest}.png and a side-by-side
// renders/v5-compare-*.png.   BEFORE=/path/old.html node shoot-v5.js
const { chromium } = require('playwright'); const path = require('path'); const fs = require('fs');
const OUT = path.join(__dirname, 'renders');
const WHEN = '2026-10-02T15:43:00-05:00';
const SRCS = { before: process.env.BEFORE, after: path.join(__dirname, '..', 'Sidebar.html') };
const tick = {
  cockpit: { shippedToday: 44, receivedToday: 60, oldestPendingMinutes: null, pendingCount: 18,
    ebayPending: 1, directPending: 17, amazonPending: 0, prepQueueCount: 15, zohoPending: 3, ebayGrab: 1, directGrab: 17,
    pastRedlineCount: 0, lastSyncMinutes: 1, floorLast: { at: Date.parse(WHEN) - 4 * 60000, event: 'PREPARING', orderId: 'SO-26018' }, floorToday: 23,
    timeline: [10.7, 11.2, 12.5, 13.1, 14.2, 15.4].map(h => ({ hourFraction: h, event: 'SHIPPED' })) },
  lastSync: '3:41 PM', picker: 'Shipping - Hatem 21332', api: null,
  alerts: { paidShipping: { count: 0, rows: [] }, lowStock: { count: 8 }, outOfStock: 145, needPhotos: 467 },
  paceCar: { projection: 60 }, openOrdersTotal: 18,
  rest: { cantBuild: { kits: 30, units: 125 }, ripple: [{ sku: '195306', sole: 30 }, { sku: '173817', sole: 19 }] },
  _devices: [{ id: 'a', role: 'floor', ageSec: 20 }]
};
(async () => {
  const b = await chromium.launch();
  for (const [tag, src] of Object.entries(SRCS)) {
    const html = fs.readFileSync(src, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");
    const page = await b.newPage({ viewport: { width: 310, height: 860 }, deviceScaleFactor: 2 });
    await page.addInitScript(({ T, whenMs }) => {
      const Real = Date; function Fake(...a) { return a.length ? new Real(...a) : new Real(whenMs); }
      Fake.prototype = Real.prototype; Fake.now = () => whenMs; Fake.parse = Real.parse; Fake.UTC = Real.UTC; window.Date = Fake;
      const DATA = { getSidebarTick: T, getCurrentPicker: T.picker, getActionableAlerts: T.alerts, getDashboardSnapshot: T.cockpit,
                     getLastSyncFromSheet: T.lastSync, getDisplayUrls: { board: 'x', wall: 'y', hosted: true } };
      const mk = (s, f) => new Proxy({}, { get(_, k) { if (k === 'withSuccessHandler') return g => mk(g, f);
        if (k === 'withFailureHandler') return g => mk(s, g);
        return () => { const v = Object.prototype.hasOwnProperty.call(DATA, k) ? DATA[k] : null; if (s) setTimeout(() => s(v), 0); }; } });
      window.google = { script: { run: mk(null, null), host: { close() {}, setHeight() {} }, url: { getLocation(f) { f({ parameter: {} }); } } } };
    }, { T: tick, whenMs: Date.parse(WHEN) });
    await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('http://hq.test/sidebar'); await page.waitForTimeout(2600);
    await page.evaluate(T => { try { _sidebarPaint(T.cockpit, T.lastSync, null, T.alerts, T.picker, T); } catch (e) {} }, tick);
    await page.waitForTimeout(600);
    await page.screenshot({ path: path.join(OUT, `v5-${tag}-head.png`), clip: { x: 0, y: 0, width: 310, height: 520 } });
    // the list: everything collapsed, top of the panel
    const setOpen = (id, open) => page.evaluate(([id, open]) => {
      const c = document.querySelector('.card[data-id="' + id + '"]');
      if (c && c.classList.contains('collapsed') === open) c.querySelector('.card-header').click(); }, [id, open]);
    await setOpen('alerts', false); await page.waitForTimeout(600);
    await page.evaluate(() => { const m = document.getElementById('modules'); if (m) m.scrollTop = 0; });
    await page.waitForTimeout(300);
    await page.screenshot({ path: path.join(OUT, `v5-${tag}-list.png`) });
    // an OPEN card right above the next section band — the confusion the owner saw
    await setOpen('displays', true); await page.waitForTimeout(700);
    await page.evaluate(() => { const m = document.getElementById('modules'); const c = document.querySelector('.card[data-id="displays"]');
      if (m && c) m.scrollTop = c.offsetTop - m.offsetTop - 150; });
    await page.waitForTimeout(400);
    await page.screenshot({ path: path.join(OUT, `v5-${tag}-cards.png`) });
    await setOpen('displays', false); await page.waitForTimeout(500);
    await page.evaluate(() => { try { _rpShow(); } catch (e) {} }); await page.waitForTimeout(2800);
    await page.screenshot({ path: path.join(OUT, `v5-${tag}-rest.png`) });
    await page.close();
  }
  await b.close();
  console.log('done');
})();
