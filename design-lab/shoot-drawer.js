// shoot-drawer.js — the sidebar's Sheets drawer (2026-10-05), driven + rendered.
//   node shoot-drawer.js   → renders/drawer-*.png + assertions
// The server is a stub holding the LIVE tab list (from the 2026-10-05 export), and it
// applies open/hide/star/tidy/new/done to that model, so every click is checked
// against what the server was asked to do AND what the drawer then shows.
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const HTML = fs.readFileSync(path.join(__dirname, '..', 'Sidebar.html'), 'utf8').replace("'<?!= boardApiUrl ?>'", "''");
const OUT = path.join(__dirname, 'renders');

let pass = 0, fail = 0;
const t = (n, got, want) => { const ok = JSON.stringify(got) === JSON.stringify(want); ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + n + (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want))); };

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 900 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));
  await p.addInitScript(() => {
    // Live tab state 2026-10-05 (system tabs left out on purpose — the server drops them).
    const S = {
      floor: ['All orders', 'Location Update', 'Out of Stock', 'Prep Queue', 'Heads & Cranks - Temp', 'ABDUL TEMP'],
      auto: false,
      tabs: { 'All orders': 0, 'Location Update': 0, 'Out of Stock': 0, 'Prep Queue': 0, 'Supplies': 1, 'KPI History': 1,
              'Order Archive': 1, 'Kit Health': 0, 'Kit Registry': 1, 'Price Audit': 1, 'Activity Log': 1,
              'Heads & Cranks - Temp': 0, 'ABDUL TEMP': 0, 'Seals - Temp': 1, 'L Location': 1 },
      work: ['Heads & Cranks - Temp', 'ABDUL TEMP', 'Seals - Temp', 'L Location'],
      calls: []
    };
    window.__S = S;
    const F = [['Orders', ['All orders', 'Prep Queue']], ['Inventory', ['Out of Stock', 'Supplies', 'Location Update']],
               ['Kits', ['Kit Health', 'Kit Registry']], ['Prices', ['Price Audit']], ['Reports & logs', ['KPI History', 'Order Archive', 'Activity Log']]];
    const row = n => ({ name: n, hidden: !!S.tabs[n], floor: S.floor.includes(n), active: false, by: n === 'ABDUL TEMP' ? 'abdul@x.com' : '', at: n === 'ABDUL TEMP' ? Date.now() - 3 * 864e5 : 0 });
    const impl = {
      getSheetDrawer: () => ({ ok: true, autoTidy: S.auto, floor: S.floor,
        folders: F.map(([l, ts]) => ({ key: l, label: l, tabs: ts.map(row) })), work: S.work.filter(n => n in S.tabs).map(row) }),
      openSheetFromDrawer: n => { S.tabs[n] = 0; return { ok: true, message: 'Opened ' + n }; },
      hideSheetFromDrawer: n => { S.tabs[n] = 1; return { ok: true, message: 'Put away ' + n }; },
      setFloorTab: (n, on) => { if (on) S.floor.push(n); else S.floor = S.floor.filter(x => x !== n); return { ok: true }; },
      setAutoTidy: on => { S.auto = on; return { ok: true }; },
      tidySheets: () => { let k = 0; for (const n in S.tabs) if (!S.tabs[n] && !S.floor.includes(n)) { S.tabs[n] = 1; k++; } return { ok: true, message: 'Put away ' + k + ' tabs' }; },
      createWorkSheet: (n, pin) => { S.tabs[n] = 0; S.work.push(n); if (pin) S.floor.push(n); return { ok: true, message: 'Made "' + n + '"' }; },
      finishWorkSheet: n => { delete S.tabs[n]; return { ok: true, message: 'Saved to the archive and removed "' + n + '"', url: 'https://docs.google.com/x' }; }
    };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return (...a) => { S.calls.push(k); const v = impl[k] ? impl[k](...a) : (k === 'getDisplayUrls' ? { board: '', wall: '', hosted: false } : null); if (su) setTimeout(() => su(v), k === 'setFloorTab' ? 1500 : 30); };
    }});
    window.google = { script: { run: mk(null, null), host: { close() {}, setHeight() {} } } };
    window.confirm = () => true;
  });
  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);

  const btn = await p.$('#drawerBtn');
  t('A1 folder button is in the top bar, drawn as a mark', await p.$eval('#drawerBtn use', u => u.getAttribute('href')).catch(() => null), '#m-tabs');
  await btn.click(); await p.waitForTimeout(300);
  const view = () => p.evaluate(() => ({
    rows: [...document.querySelectorAll('#sdList .sd-row')].map(r => ({ n: r.querySelector('.sd-name').textContent, star: r.querySelector('.sd-star').textContent, hid: r.classList.contains('hidden'), eye: !!r.querySelector('.sd-act svg'), done: [...r.querySelectorAll('.sd-act')].some(x => x.textContent === 'Done') })),
    folders: [...document.querySelectorAll('.sd-folder span:first-child')].map(e => e.textContent), sub: document.getElementById('sdSub').textContent,
    msg: document.getElementById('sdMsg').textContent, auto: document.getElementById('sdAuto').checked }));
  let v = await view();
  t('A2 drawer opens with every folder', v.folders, ['Orders', 'Inventory', 'Kits', 'Prices', 'Reports & logs', 'Work in progress']);
  t('A3 stars survive the emoji cleaner', v.rows.find(r => r.n === 'All orders').star, '★');
  t('A4 temp tabs land in Work in progress with Done', v.rows.filter(r => r.done).map(r => r.n), ['Heads & Cranks - Temp', 'ABDUL TEMP', 'Seals - Temp', 'L Location']);
  t('A5 open non-floor tab (Kit Health) offers put-away', v.rows.find(r => r.n === 'Kit Health').eye, true);
  t('A6 floor tab offers no put-away', v.rows.find(r => r.n === 'All orders').eye, false);
  await p.screenshot({ path: path.join(OUT, 'drawer-1-open.png') });

  await p.click('.sd-name:text-is("Price Audit")'); await p.waitForTimeout(250);
  t('B1 clicking a hidden tab opens it and closes the drawer', [await p.evaluate(() => window.__S.tabs['Price Audit']), await p.$eval('#sheetDrawer', e => e.classList.contains('active'))], [0, false]);

  await btn.click(); await p.waitForTimeout(250);
  await p.click('#sdTidy'); await p.waitForTimeout(300);
  v = await view();
  t('C1 Tidy hides Kit Health + Price Audit, keeps floor', [v.rows.find(r => r.n === 'Kit Health').hid, v.rows.find(r => r.n === 'Price Audit').hid, v.rows.find(r => r.n === 'ABDUL TEMP').hid], [true, true, false]);
  t('C2 Tidy reports what it did', v.msg, 'Put away 2 tabs');

  await p.click('.sd-row:has(.sd-name:text-is("Supplies")) .sd-star'); await p.waitForTimeout(60);
  t('D1 ☆ → ★ shows INSTANTLY (server answers 1.5 s later)', (await view()).rows.find(r => r.n === 'Supplies').star, '★');
  await p.click('.sd-row:has(.sd-name:text-is("Kit Registry")) .sd-star'); await p.waitForTimeout(3300);
  t('D1b two quick stars are both saved (queued, not lost)', await p.evaluate(() => ['Supplies','Kit Registry'].every(n => window.__S.floor.includes(n))), true);
  await p.check('#sdAuto'); await p.waitForTimeout(250);
  t('D2 auto-tidy switch reaches the server', await p.evaluate(() => window.__S.auto), true);

  await p.fill('#sdNewName', 'Gasket count'); await p.click('.sd-new .sd-act'); await p.waitForTimeout(300);
  t('E1 + New makes the tab, pinned to the floor', await p.evaluate(() => [window.__S.tabs['Gasket count'], window.__S.floor.includes('Gasket count')]), [0, true]);
  await btn.click(); await p.waitForTimeout(250);
  await p.click('.sd-row:has(.sd-name:text-is("Seals - Temp")) .sd-act:text-is("Done")'); await p.waitForTimeout(300);
  v = await view();
  t('E2 Done removes the tab and links the archived copy', [v.rows.some(r => r.n === 'Seals - Temp'), /open the copy/.test(v.msg)], [false, true]);
  await p.screenshot({ path: path.join(OUT, 'drawer-2-after.png') });
  await p.keyboard.press('Escape'); await p.waitForTimeout(150);
  t('F1 Esc key closes it', await p.$eval('#sheetDrawer', e => e.classList.contains('active')), false);
  await btn.click(); await p.waitForTimeout(250);
  await p.click('#sheetDrawer .palette-hint'); await p.waitForTimeout(150);
  t('F1b the ESC chip closes it', await p.$eval('#sheetDrawer', e => e.classList.contains('active')), false);
  await btn.click(); await p.waitForTimeout(250);
  await p.evaluate(() => window.dispatchEvent(new Event('blur'))); await p.waitForTimeout(100);
  t('F1c focus leaving the panel (clicked the sheet) closes it', await p.$eval('#sheetDrawer', e => e.classList.contains('active')), false);
  t('F2 no page errors', errs, []);
  await b.close();
  console.log('\n' + pass + ' passed · ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})();
