/**
 * shoot-cards.js — render EVERY sidebar card individually, expanded, at real width.
 *
 * ⚠ Element screenshots, not viewport clips: #modules scrolls inside a fixed header and
 *   footer, so a clip box only ever captures the visible slice of a tall card — which is
 *   exactly how a clipped card reads as "fine" in a render.
 *
 * ⚠ Cards are force-expanded and <details> forced open, because a harness that cannot see
 *   a surface passes vacuously — the blindness that let the command palette survive two
 *   emoji sweeps.
 *
 *   node shoot-cards.js            all cards
 *   node shoot-cards.js missing-line sheet-protect      only these
 */
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');

const SRC  = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");
const ONLY = process.argv.slice(2);
const OUT  = path.join(__dirname, 'renders', 'cards');

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 3000 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  p.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  await p.addInitScript(() => {
    const D = {
      getDisplayUrls: { board: 'https://hq.yassinqurabi.com/', wall: 'https://hq.yassinqurabi.com/wall', hosted: true },
      getSidebarTick: {
        cockpit: { shippedToday: 12, receivedToday: 18, oldestPendingMinutes: 47, pendingCount: 9,
                   ebayPending: 6, directPending: 3, zohoPending: 2, prepQueueCount: 4,
                   lastSyncMinutes: 4, timeline: [], pastRedlineCount: 1 },
        lastSync: 4, picker: 'Shipping - Yassin 1',
        alerts: { paidShipping: { count: 2 }, intl: { count: 1 }, lowStock: { count: 5 },
                  notFound: { count: 0 }, outOfStock: { count: 107 }, queueSize: { count: 4 },
                  newFromZoho: { count: 2 }, openCases: { count: 0 }, needPhotos: { count: 460 },
                  priceDrift: { count: 3 }, kitPriceDrift: { count: 7 } },
        api: { worstPct: 61, tradingApi: { pct: 61, used: 3011, limit: 5000, label: 'Trading API', endpoint: 'shared' },
               fulfillment: { pct: 1, used: 140, limit: 100000, label: 'Fulfillment' },
               feed: { pct: 0, used: 20, limit: 100000, label: 'Inventory Feed' },
               analytics: { pct: 2, used: 40, limit: 2000, label: 'Analytics' },
               tradingMonitorAvailable: true }
      },
      getCurrentPicker: 'Shipping - Yassin 1',
      getN8nSheetsAccountState: { ok: true, owner: true, isSet: true, isNone: false,
                                  value: 'hq-n8n-sheets@highqualitymotor.com', key: 'N8N_SHEETS_ACCOUNT' }
    };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return () => { const v = Object.prototype.hasOwnProperty.call(D, k) ? D[k] : null; if (su) setTimeout(() => su(v), 0); };
    }});
    window.google = { script: { run: mk(null, null), host: { close(){}, setHeight(){} },
                      url: { getLocation(f) { f({ parameter: {} }); } } } };
  });

  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);

  await p.evaluate(() => {
    document.querySelectorAll('.card.collapsed').forEach(c => c.classList.remove('collapsed'));
    document.querySelectorAll('#modules details').forEach(d => d.open = true);
    document.querySelectorAll('.card-body').forEach(b => { b.style.transition = 'none'; });
  });
  await p.waitForTimeout(500);
  if (typeof p.evaluate === 'function') { try { await p.evaluate(() => lockPanelLoad && lockPanelLoad()); } catch (e) {} }
  await p.waitForTimeout(400);

  const ids = await p.evaluate(() => [...document.querySelectorAll('.card')].map(c => c.dataset.id));
  const want = ONLY.length ? ids.filter(i => ONLY.includes(i)) : ids;

  for (const id of want) {
    const info = await p.evaluate(i => {
      const c = document.querySelector(`.card[data-id="${i}"]`);
      const bd = c.querySelector('.card-body');
      return { h: Math.round(c.getBoundingClientRect().height),
               clipped: bd ? bd.scrollHeight > bd.clientHeight + 2 : false,
               over: bd ? bd.scrollHeight - bd.clientHeight : 0 };
    }, id);
    await p.locator(`.card[data-id="${id}"]`).screenshot({ path: path.join(OUT, id + '.png') });
    console.log('  ' + (info.clipped ? '❌' : '  ') + ' ' + id.padEnd(18) + String(info.h).padStart(5) + 'px'
                + (info.clipped ? '   CLIPPED by ' + info.over + 'px' : ''));
  }
  if (errs.length) console.log('\n  ⚠ page errors: ' + errs.slice(0, 4).join(' | '));
  console.log('\n  wrote design-lab/renders/cards/*.png');
  await b.close();
})();
