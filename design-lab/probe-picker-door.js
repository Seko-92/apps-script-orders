/**
 * probe-picker-door.js — the sidebar's Pick ID door, driven for real.
 *
 * WHY IT EXISTS. The pick-ID cells moved into the hidden columns on 2026-08-31, which
 * removed the only desk-side way to set the shift's picker. preparePrintSheet REFUSES
 * without one, so printing from a desk was impossible until this control existed. The
 * floor kept working only because the Floor Board has its own drawer.
 *
 * Run:   node probe-picker-door.js
 * HEAD:  SIDEBAR_SRC=/tmp/old.html node probe-picker-door.js   (should fail loudly)
 */
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');

const SRC  = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

const PICKERS = ['Shipping - AShamma 12343', 'Shipping - Hatem 21332',
                 'Shipping - YAwiss 1', 'Shipping - Turkmani 43122'];

let fail = 0;
const ok = (n, c, g) => { c ? console.log('  ✓ ' + n)
  : (fail++, console.log('  ✗ ' + n + (g !== undefined ? '  → ' + JSON.stringify(g) : ''))); };
const section = t => console.log('\n' + t);

async function boot(opts) {
  opts = opts || {};
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 1400 }, deviceScaleFactor: 2 });
  const errs = [];
  p.on('pageerror', e => errs.push(String(e)));
  p.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  await p.addInitScript(([pickers, okReply]) => {
    window.__calls = [];                       // every server call, in order
    const tick = {
      cockpit: { ebayPending: 0, directPending: 0, prepQueueCount: 0 },
      lastSync: '', api: null, alerts: null, picker: '', pickers: pickers, rest: null
    };
    const D = {
      getDisplayUrls: { board: '', wall: '', hosted: true },
      getSidebarTick: tick, getCurrentPicker: '', getActionableAlerts: null,
      setSidebarPicker: okReply ? { ok: true, picker: 'x', message: '' }
                                : { ok: false, picker: '', message: 'Not a known picker.' }
    };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return (...a) => {
        window.__calls.push({ fn: k, args: a });
        const v = Object.prototype.hasOwnProperty.call(D, k) ? D[k] : null;
        if (k === 'setSidebarPicker' && v && v.ok) v.picker = a[0];
        if (su) setTimeout(() => su(v), 0);
      };
    }});
    window.google = { script: { run: mk(null, null), host: { close(){}, setHeight(){} },
                                url: { getLocation(f){ f({ parameter:{} }); } } } };
  }, [opts.pickers === undefined ? PICKERS : opts.pickers, opts.okReply !== false]);

  await p.route('http://hq.test/**', r =>
    r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);
  return { b, p, errs };
}

(async () => {
  // ── A · the banner is a door ────────────────────────────────────────────────
  section('A · the banner is a control, not a label');
  {
    const { b, p, errs } = await boot();
    ok('A1 it has a click handler', await p.evaluate(() =>
      !!document.getElementById('pickerBanner').getAttribute('onclick')));
    ok('A2 and reads as clickable', await p.evaluate(() =>
      getComputedStyle(document.getElementById('pickerBanner')).cursor === 'pointer'));
    ok('A3 unset copy invites the tap', await p.evaluate(() =>
      /tap to choose/i.test(document.getElementById('pickerValue').textContent)));
    ok('A4 the overlay starts closed', await p.evaluate(() =>
      !document.getElementById('pickerPick').classList.contains('active')));
    ok('A5 no page errors', errs.length === 0, errs);
    await b.close();
  }

  // ── B · it lists the pickers, labelled not raw ──────────────────────────────
  section('B · the list arrives on the tick and renders readably');
  {
    const { b, p } = await boot();
    await p.evaluate(() => openPickerPick());
    await p.waitForTimeout(180);
    const rows = await p.evaluate(() =>
      [...document.querySelectorAll('#pkBody .pk-row')].map(r => ({
        label: r.textContent.replace('✓','').trim(), raw: r.getAttribute('data-raw') })));
    ok('B1 four rows', rows.length === 4, rows.length);
    ok('B2 ⚠ labelled WITHOUT the "Shipping - " prefix', rows[0].label === 'AShamma 12343', rows[0].label);
    ok('B3 ⚠⚠ but the RAW option is what gets sent — the server matches with strict indexOf',
       rows[0].raw === 'Shipping - AShamma 12343', rows[0].raw);
    ok('B4 no server call was made just by opening',
       await p.evaluate(() => window.__calls.filter(c => c.fn === 'setSidebarPicker').length) === 0);
    await p.screenshot({ path: path.join(__dirname,'renders','mast','_pickerdoor.png') });
    await b.close();
  }

  // ── C · choosing ────────────────────────────────────────────────────────────
  section('C · choosing sends the raw value and holds the name');
  {
    const { b, p } = await boot();
    await p.evaluate(() => openPickerPick());
    await p.waitForTimeout(150);
    await p.evaluate(() => document.querySelectorAll('#pkBody .pk-row')[2].click());
    await p.waitForTimeout(400);
    const calls = await p.evaluate(() => window.__calls.filter(c => c.fn === 'setSidebarPicker'));
    ok('C1 exactly one write', calls.length === 1, calls.length);
    ok('C2 ⚠ with the RAW option, verbatim', calls[0] && calls[0].args[0] === 'Shipping - YAwiss 1',
       calls[0] && calls[0].args[0]);
    ok('C3 the overlay closed', await p.evaluate(() =>
      !document.getElementById('pickerPick').classList.contains('active')));
    ok('C4 ⭐ the banner holds the new name optimistically', await p.evaluate(() =>
      /YAwiss/.test(document.getElementById('pickerValue').textContent)),
      await p.evaluate(() => document.getElementById('pickerValue').textContent));
    ok('C5 ...and shows it as SET, not unset', await p.evaluate(() =>
      document.getElementById('pickerBanner').classList.contains('picker-set')));
    await b.close();
  }

  // ── D · ⚠⚠ the listener-stacking regression ─────────────────────────────────
  section('D · ⚠⚠ four opens, one tap, ONE write (the 2026-08-13 board bug)');
  {
    const { b, p } = await boot();
    for (let i = 0; i < 4; i++) {
      await p.evaluate(() => openPickerPick());
      await p.waitForTimeout(80);
      await p.evaluate(() => closePickerPick());
      await p.waitForTimeout(60);
    }
    await p.evaluate(() => openPickerPick());
    await p.waitForTimeout(120);
    await p.evaluate(() => document.querySelectorAll('#pkBody .pk-row')[0].click());
    await p.waitForTimeout(400);
    const n = await p.evaluate(() => window.__calls.filter(c => c.fn === 'setSidebarPicker').length);
    ok('D1 ⚠⚠ exactly ONE write after four opens — onclick ASSIGNMENT, not addEventListener',
       n === 1, n);
    await b.close();
  }

  // ── E · a refusal must not leave the optimistic name standing ───────────────
  section('E · ⚠ a refusal retires the override immediately');
  {
    const { b, p } = await boot({ okReply: false });
    await p.evaluate(() => openPickerPick());
    await p.waitForTimeout(150);
    await p.evaluate(() => document.querySelectorAll('#pkBody .pk-row')[0].click());
    await p.waitForTimeout(600);
    ok('E1 the banner does NOT keep showing the rejected name', await p.evaluate(() =>
      !/AShamma/.test(document.getElementById('pickerValue').textContent)),
      await p.evaluate(() => document.getElementById('pickerValue').textContent));
    ok('E2 and the refusal is on the status bar', await p.evaluate(() =>
      document.getElementById('status').className === 'error'));
    ok('E3 ⚠ the override is cleared, not left to its 5-minute TTL',
       await p.evaluate(() => window._pkOverride === null || window._pkOverride === undefined));
    await b.close();
  }

  // ── F · an empty list explains WHICH failure it is ──────────────────────────
  section('F · ⚠ an empty list says which failure it is');
  {
    const { b, p } = await boot({ pickers: [] });
    await p.evaluate(() => openPickerPick());
    await p.waitForTimeout(150);
    const t = await p.evaluate(() => document.getElementById('pkBody').textContent);
    ok('F1 it names the not-yet-loaded case', /30-second tick|wait a moment/i.test(t), t.slice(0,90));
    ok('F2 ⚠ and the genuinely-empty case, which needs the opposite response',
       /no options matching/i.test(t), t.slice(0,90));
    ok('F3 no rows are rendered', await p.evaluate(() =>
      document.querySelectorAll('#pkBody .pk-row').length) === 0);
    await b.close();
  }

  console.log('\n' + (fail ? '✗ ' : '✅ ') + 'probe-picker-door: ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})();
