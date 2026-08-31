// =====================================================================================
// shoot-palette.js — the command palette, the LAST emoji surface inside the sidebar.
//
// It needed its own shot because the palette is invisible to every existing harness:
// it renders into #cmdList only once opened, and the _hqMarks pass walks #modules once
// at page load, so nothing had ever seen these icons.
//
// Shoots BOTH THEMES because .cmd-icon has four states (light/dark × rest/hover) and
// the marks are drawn with `stroke: currentColor` — the point is that one declaration
// is correct in all four, which a single-theme shot cannot show.
//
// Usage: node shoot-palette.js
// =====================================================================================
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(__dirname, 'renders');

(async () => {
  let html = fs.readFileSync(path.join(ROOT, 'Sidebar.html'), 'utf8');
  html = html.replace("'<?!= boardApiUrl ?>'", "''");
  if (/<\?/.test(html)) throw new Error('unresolved scriptlet left in the page');

  const browser = await chromium.launch();
  let bad = 0;

  for (const theme of ['light', 'dark']) {
    const page = await browser.newPage({ viewport: { width: 310, height: 700 }, deviceScaleFactor: 2 });
    const errs = []; page.on('pageerror', e => errs.push(String(e)));
    await page.addInitScript(() => {
      const DATA = { getSidebarTick: { cockpit: {}, alerts: null, api: null, picker: '', lastSync: '' } };
      function mk(succ, f) { return new Proxy({}, { get(_, p) {
        if (p === 'withSuccessHandler') return (x) => mk(x, f);
        if (p === 'withFailureHandler') return (x) => mk(succ, x);
        return () => { const v = Object.prototype.hasOwnProperty.call(DATA, p) ? DATA[p] : null;
                       if (succ) setTimeout(() => succ(v), 0); }; }}); }
      window.google = { script: { run: mk(null, null), host: { close(){}, setHeight(){} },
                                  url: { getLocation(f){ f({ parameter: {} }); } } } };
    });
    await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
    await page.goto('http://hq.test/sidebar', { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(900);

    if (theme === 'dark') await page.evaluate(() => document.body.classList.add('dark-theme'));
    await page.evaluate(() => { openPalette(); });
    await page.waitForTimeout(400);

    const r = await page.evaluate(() => {
      const rows = Array.from(document.querySelectorAll('#cmdList .cmd'));
      const emojiRe = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/u;
      return {
        count: rows.length,
        // every icon must resolve to a REAL symbol, and none may render empty
        unresolved: rows.map(x => {
          const u = x.querySelector('.cmd-icon use');
          if (!u) return 'NO <use>';
          const id = (u.getAttribute('href') || '').slice(1);
          return document.getElementById(id) ? null : ('dangling #' + id);
        }).filter(Boolean),
        empty: rows.filter(x => { const b = x.querySelector('.cmd-icon svg').getBoundingClientRect();
                                  return b.width < 4 || b.height < 4; }).length,
        emoji: rows.map(x => x.textContent).filter(t => emojiRe.test(t)),
        ink: rows.slice(0, 3).map(x => getComputedStyle(x.querySelector('.cmd-icon svg')).stroke)
      };
    });

    console.log(`\n═══ ${theme} ═══`);
    console.log(`  rows: ${r.count}`);
    console.log(`  every <use> resolves to a real <symbol>: ${r.unresolved.length === 0 ? 'yes' : r.unresolved.join(', ')}`);
    console.log(`  marks that render empty: ${r.empty}`);
    console.log(`  EMOJI LEFT IN THE PALETTE: ${r.emoji.length}`);
    console.log(`  stroke resolves to real ink (first 3): ${r.ink.join(' · ')}`);
    if (errs.length) console.log('  page errors: ' + errs.slice(0, 3).join(' | '));
    if (r.unresolved.length || r.empty || r.emoji.length || errs.length) bad++;

    await page.screenshot({ path: path.join(OUT, `palette-${theme}.png`) });
    await page.close();
  }

  await browser.close();
  console.log('\n' + (bad ? `✗ ${bad} theme(s) failed` : '✓ both themes clean'));
  process.exit(bad ? 1 : 0);
})();
