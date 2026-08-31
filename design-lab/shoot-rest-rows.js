// =====================================================================================
// shoot-rest-rows.js — a TRUTHFUL picture of the resting panel's day state, and of the
// two dark rows the server half now feeds.
//
// ⚠⚠ WHY THIS EXISTS SEPARATELY FROM shoot-rest-panel.js. That harness installs
// page.clock to assert MECHANICS on a driven clock — and page.clock does NOT drive CSS
// transitions, which run on the compositor clock. So every screenshot it takes catches
// the panel mid-fade and is worthless as a picture (rest-day.png is a brown smear).
//
// The split this file implements: fake the WALL CLOCK the panel reads (a Date override,
// so _rpNow believes it is a Thursday afternoon in Houston), leave the ANIMATION clock
// completely real, then drive _rpShow() directly and wait out the real fade.
//
// Usage: node shoot-rest-rows.js
// =====================================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');

const ROOT = path.join(__dirname, '..');
const SIDEBAR = path.join(ROOT, 'Sidebar.html');
const OUT = path.join(__dirname, 'renders');
const WHEN = '2026-08-13T14:30:00-05:00';      // Thu 2:30 PM Houston — mid shift

const tick = (rest) => ({
  cockpit: {
    shippedToday: 14, receivedToday: 23, oldestPendingMinutes: 212, pendingCount: 9,
    ebayPending: 6, directPending: 3, prepQueueCount: 4, zohoPending: 65,
    // ⚠ 'On the floor' reads ebayGrab/directGrab (strictly PENDING), NOT the
    //   *Pending fields above, which also count PREPARING. Omitting them drops
    //   the row entirely — a fixture gap that looks exactly like a code bug.
    ebayGrab: 6, directGrab: 3,
    pastRedlineCount: 3, lastSyncMinutes: 4,
    timeline: [9.25, 10.5, 11.1, 13.75, 14.2].map(h => ({ hour: h, event: 'SHIPPED' }))
  },
  lastSync: '2:26 PM', picker: 'Yassin · 1',
  api: null, alerts: { paidShipping: { count: 0, rows: [] } },
  openOrdersTotal: 24,
  rest
});

(async () => {
  let html = fs.readFileSync(SIDEBAR, 'utf8');
  html = html.replace("'<?!= boardApiUrl ?>'", "''");
  if (/<\?/.test(html)) throw new Error('unresolved scriptlet left in the page');

  const browser = await chromium.launch();
  const shots = [];

  for (const [tag, rest] of [
    ['with',    { cantBuild: { kits: 26, units: 117 }, ripple: [{ sku: '195306', sole: 7 }, { sku: '173817', sole: 2 }] }],
    ['without', null]
  ]) {
    const page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
    const errs = []; page.on('pageerror', e => errs.push(String(e)));

    await page.addInitScript(({ tickJson, whenMs }) => {
      // ⚠ THE WALL CLOCK ONLY. Date is what _rpNow reads to decide day vs night and
      // to place the jewel; rAF, transitions and setTimeout stay on the real clock so
      // the fade actually runs and the screenshot lands on a settled panel.
      const Real = Date;
      function Fake(...a) { return a.length ? new Real(...a) : new Real(whenMs); }
      Fake.prototype = Real.prototype;
      Fake.now = () => whenMs; Fake.parse = Real.parse; Fake.UTC = Real.UTC;
      window.Date = Fake;

      const T = JSON.parse(tickJson);
      const DATA = { getSidebarTick: T, getCurrentPicker: T.picker,
                     getActionableAlerts: T.alerts, getLatestApiMetrics: null,
                     getLastSyncFromSheet: T.lastSync, getDashboardSnapshot: T.cockpit };
      function makeRunner(succ, failcb) {
        return new Proxy({}, { get(_, prop) {
          if (prop === 'withSuccessHandler') return (f) => makeRunner(f, failcb);
          if (prop === 'withFailureHandler') return (f) => makeRunner(succ, f);
          return () => { const v = Object.prototype.hasOwnProperty.call(DATA, prop) ? DATA[prop] : null;
                         if (succ) setTimeout(() => succ(v), 0); };
        }});
      }
      window.google = { script: { run: makeRunner(null, null), host: { close(){}, setHeight(){} },
                                  url: { getLocation(f){ f({ parameter: {} }); } } } };
    }, { tickJson: JSON.stringify(tick(rest)), whenMs: new Date(WHEN).getTime() });

    // ⚠ A REAL ORIGIN, and charset=utf-8 — setContent leaves an opaque origin where
    // localStorage throws, and windows-1252 turns every '·' into mojibake.
    await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('http://hq.test/sidebar', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(1200);                       // boot + first tick, real time

    await page.evaluate(() => { _rpShow(); });
    await page.waitForTimeout(2600);                       // the real ~1.4s staged fade, with margin

    const rows = await page.evaluate(() => Array.from(
      document.querySelectorAll('#restPanel .rp-row')).map(r => ({
        label: (r.querySelector('span') || {}).textContent,
        value: (r.querySelector('b') || {}).textContent,
        // ⚠ the warn class is on the VALUE <b>, not on the row
        warn: !!r.querySelector('b.rp-warn')
      })));
    const opacity = await page.evaluate(() =>
      parseFloat(getComputedStyle(document.getElementById('restPanel')).opacity));

    console.log(`\n═══ rest ${tag === 'with' ? 'PRESENT' : 'ABSENT'} ═══`);
    console.log(`  panel opacity: ${opacity}  (1 = the fade finished, so this is a real picture)`);
    rows.forEach(r => console.log(`  ${r.warn ? '⚠ ' : '  '}${r.label} · ${r.value}`));
    if (errs.length) console.log('  page errors: ' + errs.slice(0, 3).join(' | '));

    const p = path.join(OUT, `rest-rows-${tag}.png`);
    await page.screenshot({ path: p });
    shots.push(p);
    await page.close();
  }

  await browser.close();
  console.log('\nrenders → ' + shots.join('\n           '));
})();
