// Truthful render of the Missing / Replacement sidebar card, expanded.
// Boot copied from check-sidebar.js — the server stub matters: without
// getDisplayUrls the panel renders wrong for reasons unrelated to this card.
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const SRC = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 900 }, deviceScaleFactor: 2 });
  await p.addInitScript(() => {
    const D = { getDisplayUrls: { board: 'https://hq.yassinqurabi.com/', wall: 'https://hq.yassinqurabi.com/wall', hosted: true },
                getSidebarTick: null, getCurrentPicker: '', getActionableAlerts: null,
                previewReplacementFromSidebar: { ok: true, kind: 'missing', label: 'MISSING',
                  salesOrder: 'Missing #: 05-15052-93025', originalOrder: '05-15052-93025',
                  originalWhere: 'in the Activity Log', sku: '212498', qty: 1, note: '',
                  location: 'E-57', hand: 7, warnings: [] } };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return () => { const v = Object.prototype.hasOwnProperty.call(D, k) ? D[k] : null; if (su) setTimeout(() => su(v), 0); }; } });
    window.google = { script: { run: mk(null, null), host: { close() {}, setHeight() {} }, url: { getLocation(f) { f({ parameter: {} }); } } } };
  });
  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2400);

  // Expand the card and drive the REAL path — never call the paint fn directly.
  await p.evaluate(() => {
    const c = document.querySelector('.card[data-id="missing-line"]');
    c.classList.remove('collapsed');
    c.scrollIntoView();
    document.getElementById('mlOrder').value = '05-15052-93025';
    document.getElementById('mlSku').value = '212498';
  });
  await p.waitForTimeout(300);
  await p.evaluate(() => previewMissingLine());
  await p.waitForTimeout(600);
  const box = await p.evaluate(() => {
    const c = document.querySelector('.card[data-id="missing-line"]').getBoundingClientRect();
    return { x: 0, y: Math.max(0, c.top - 8), width: 310, height: Math.min(700, c.height + 16) };
  });
  await p.screenshot({ path: path.join(__dirname, 'renders', 'missingcard.png'), clip: box });
  console.log('rendered → design-lab/renders/missingcard.png');
  await b.close();
})();
