/**
 * shoot-lockpanel.js — LOOK at the Sheet Protection card, in both account states.
 *
 * ⚠ The card is COLLAPSED by default and its state line only paints after
 *   lockPanelLoad() answers, so nothing in check-sidebar ever sees this. A harness
 *   that cannot see a surface passes vacuously — that blindness is why the command
 *   palette survived two emoji sweeps.
 */
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');

const SRC  = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

const STATES = {
  unset: { ok: true, owner: true, isSet: false, isNone: false, value: '', key: 'N8N_SHEETS_ACCOUNT' },
  set:   { ok: true, owner: true, isSet: true,  isNone: false, value: 'hq-n8n-sheets@highqualitymotor.com', key: 'N8N_SHEETS_ACCOUNT' },
  none:  { ok: true, owner: true, isSet: true,  isNone: true,  value: 'none', key: 'N8N_SHEETS_ACCOUNT' }
};

(async () => {
  const b = await chromium.launch();
  for (const [tag, st] of Object.entries(STATES)) {
    const p = await b.newPage({ viewport: { width: 310, height: 2600 }, deviceScaleFactor: 2 });
    const errs = []; p.on('pageerror', e => errs.push(String(e)));
    await p.addInitScript(state => {
      const D = { getDisplayUrls: { board: 'https://hq.yassinqurabi.com/', wall: 'https://hq.yassinqurabi.com/wall', hosted: true },
                  getSidebarTick: null, getCurrentPicker: '', getActionableAlerts: null,
                  getN8nSheetsAccountState: state };
      const mk = (su, fa) => new Proxy({}, { get(_, k) {
        if (k === 'withSuccessHandler') return f => mk(f, fa);
        if (k === 'withFailureHandler') return f => mk(su, f);
        return () => { const v = Object.prototype.hasOwnProperty.call(D, k) ? D[k] : null; if (su) setTimeout(() => su(v), 0); };
      }});
      window.google = { script: { run: mk(null, null), host: { close(){}, setHeight(){} },
                        url: { getLocation(f) { f({ parameter: {} }); } } } };
    }, st);
    await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
    await p.goto('http://hq.test/sidebar');
    await p.waitForTimeout(1800);

    // open the card the way a person does — through its own header handler
    await p.evaluate(() => {
      const h = document.querySelector('.card[data-id="sheet-protect"] .card-header');
      h.click();
      document.querySelector('.card[data-id="sheet-protect"]')
        .scrollIntoView({ block: 'start' });
    });
    await p.waitForTimeout(700);

    // ⚠ ELEMENT screenshot, not a viewport clip. #modules scrolls independently inside a
    //   fixed header/footer, so a clip box can only ever capture the visible slice of a
    //   tall card — which is exactly how a clipped card reads as "fine" in a render.
    const out = path.join(__dirname, 'renders', 'lockpanel-' + tag + '.png');
    await p.locator('.card[data-id="sheet-protect"]').screenshot({ path: out });
    const said = await p.evaluate(() => (document.getElementById('lockAcctResult') || {}).textContent || '');
    const val  = await p.evaluate(() => (document.getElementById('lockAcct') || {}).value || '');
    console.log('  ' + tag.padEnd(6) + ' field=' + JSON.stringify(val));
    console.log('         says =' + JSON.stringify(said.trim().slice(0, 110)));
    if (errs.length) console.log('         ⚠ page errors: ' + errs.slice(0, 2).join(' | '));
    await p.close();
  }
  await b.close();
  console.log('\n  wrote design-lab/renders/lockpanel-{unset,set,none}.png');
})();
