// shoot-floor-pulse.js — renders the sidebar cockpit with the Floor line in its three states.
const { chromium } = require('playwright'); const path = require('path'); const fs = require('fs');
const HTML = fs.readFileSync(path.join(__dirname, '..', 'Sidebar.html'), 'utf8').replace("'<?!= boardApiUrl ?>'", "''");
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 700 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar'); await p.waitForTimeout(2200);
  await p.evaluate(() => { const s = document.getElementById('bootSplash'); if (s) s.remove(); });
  const cases = [
    ['active', { floorLast: { at: Date.now() - 4 * 60000, event: 'PREPARING', orderId: '24-15008-33107', sku: '166500', picker: 'Hatem · 21332', screen: 'Android tablet · Chrome' }, floorToday: 23 },
      [{ role: 'floor', ageSec: 5 }]],
    ['quiet', { floorLast: { at: Date.now() - 70 * 60000 }, floorToday: 9 }, [{ role: 'floor', ageSec: 5 }]],
    ['off', { floorToday: 0 }, [{ role: 'remote', ageSec: 5 }]]];
  for (const [n, c, d] of cases) {
    await p.evaluate(([c, d]) => _paintFloorPulse(c, d), [c, d]);
    const box = await p.evaluate(() => { const r = document.querySelector('.ops-cockpit, #opsCockpit, .cockpit') ; const q = document.getElementById('floorPulse').getBoundingClientRect(); return { y: Math.max(0, q.y - 120), t: document.getElementById('floorPulse').textContent, w: q.width, sw: document.getElementById('floorPulse').scrollWidth }; });
    console.log(n, JSON.stringify(box));
    await p.screenshot({ path: path.join(__dirname, 'renders', 'floor-pulse-' + n + '.png'), clip: { x: 0, y: box.y, width: 310, height: 160 } });
  }
  console.log('errors:', errs.slice(0, 3));
  await b.close();
})();
