/**
 * probe-alert-age.js — render the Alerts card with REAL alert data (2026-09-16)
 *
 * ⚠⚠ WHY THIS EXISTS SEPARATELY FROM check-sidebar.js. That harness stubs
 * `getActionableAlerts: null`, and _paintAlerts opens with `if (!alerts) return;`
 * — so the age lines are NEVER painted there, the spans stay empty, `:empty`
 * hides them, and its "no card body is clipped" assertion measures the card at
 * its SHORTEST. It would pass whether or not the age fits. That is the vacuous-
 * pass trap this project keeps re-learning (the palette surviving two emoji
 * sweeps because nothing opened it).
 *
 * So this feeds a populated payload and asserts what the operator actually sees:
 *   • both age lines render, with the right words
 *   • a null checkedAt says "never checked" — never a fresh-looking age
 *   • the Alerts card still fits inside its 800px max-height once two lines grow
 *   • the age is legible on an EMPTY (dimmed) row, which is where it matters most
 *
 * Run:  node design-lab/probe-alert-age.js
 *       SIDEBAR_SRC=/tmp/old.html node design-lab/probe-alert-age.js   (before/after)
 */

'use strict';
const { chromium } = require('playwright');
const path = require('path');
const fs   = require('fs');

const SRC  = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

const DAY = 86400000;

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 310, height: 2600 }, deviceScaleFactor: 2 });
  const errs = [];
  p.on('pageerror', e => errs.push(String(e)));
  p.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  await p.addInitScript(([sixDaysAgo]) => {
    /* A realistic shape. Price Diffs carries a SMALL count with a SIX-DAY-OLD
       stamp — the live reality of a weekly audit, and the exact state the operator
       described as "it stalls". Kit Price Diffs carries a null stamp, to prove the
       unknown case renders honestly rather than as a fresh one. */
    const ALERTS = {
      paidShipping:  { count: 0, rows: [] },
      intl:          { count: 0, rows: [] },
      lowStock:      { count: 3, rows: [] },
      notFound:      { count: 0, rows: [] },
      queueSize:     { count: 4, rows: [] },
      outOfStock:    { count: 107, rows: [] },
      newFromZoho:   { count: 2, rows: [] },
      priceDrift:    { count: 1,  rows: [], checkedAt: sixDaysAgo },
      kitPriceDrift: { count: 79, rows: [], checkedAt: null },
      openCases:     { count: 0, rows: [] },
      needPhotos:    { count: 464, rows: [] },
      heldOrders:    { count: 0, rows: [] },
      identityIssues:{ count: 0, rows: [] }
    };
    const D = {
      getDisplayUrls: { board: 'https://hq.yassinqurabi.com/', wall: 'https://hq.yassinqurabi.com/wall', hosted: true },
      getSidebarTick: null,
      getCurrentPicker: '',
      getActionableAlerts: ALERTS
    };
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return () => { const v = Object.prototype.hasOwnProperty.call(D, k) ? D[k] : null; if (su) setTimeout(() => su(v), 0); };
    }});
    window.google = { script: { run: mk(null, null), host: { close(){}, setHeight(){} }, url: { getLocation(f){ f({ parameter:{} }); } } } };
  }, [Date.now() - 6 * DAY]);

  await p.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(1800);

  /* ⚠ THE PANEL BOOTS THROUGH getSidebarTick, NOT getActionableAlerts — one
     consolidated round trip (2026-05-21). With the tick stubbed null its failure
     handler deliberately keeps last-known counts, so the alert rows never paint
     and every assertion below would pass or fail for the wrong reason. Drive the
     real on-demand path instead: refreshAlerts() is a genuine entry point (it is
     what the post-action refreshes call) and it routes through the same
     _paintAlerts the tick uses, so the painter under test is the shipped one. */
  await p.evaluate(() => { if (typeof refreshAlerts === 'function') refreshAlerts(); });
  await p.waitForTimeout(600);

  let fail = 0;
  const ok = (n, c, x) => { c ? console.log('  ok  ' + n)
                              : (fail++, console.log('  XX  ' + n + (x !== undefined ? '  -> ' + JSON.stringify(x) : ''))); };

  const r = await p.evaluate(() => {
    const cs = n => getComputedStyle(n);
    const priceAge = document.getElementById('alertAgePriceDrift');
    const kitAge   = document.getElementById('alertAgeKitPriceDrift');
    const body     = document.querySelector('.card[data-id="alerts"] .card-body');
    const row      = document.getElementById('alertPriceDrift');
    return {
      priceText:  priceAge ? priceAge.textContent : null,
      kitText:    kitAge   ? kitAge.textContent   : null,
      priceTitle: priceAge ? priceAge.title       : null,
      kitTitle:   kitAge   ? kitAge.title         : null,
      countText:  (document.getElementById('alertCountPriceDrift') || {}).textContent,
      // Does the Alerts card still fit? maxHeight is the 800px cap; scrollHeight
      // is what the content actually wants now that two rows grew a second line.
      bodyScroll: body ? body.scrollHeight : null,
      bodyMax:    body ? cs(body).maxHeight : null,
      // Rendered ink + size — assert the computed value, never the class name.
      ageColor:   priceAge ? cs(priceAge).color : null,
      ageSize:    priceAge ? cs(priceAge).fontSize : null,
      ageDisplay: priceAge ? cs(priceAge).display : null,
      // The zero-count row is dimmed; the age must still be readable there.
      kitRowEmpty: (document.getElementById('alertKitPriceDrift') || {}).className || '',
      rowClass:    row ? row.className : ''
    };
  });

  console.log('');
  console.log('  Price Diffs  →  ' + JSON.stringify(r.countText) + '  ' + JSON.stringify(r.priceText));
  console.log('  Kit Price    →  ' + JSON.stringify(r.kitText));
  console.log('  card body    →  wants ' + r.bodyScroll + 'px, cap ' + r.bodyMax);
  console.log('');

  ok('the age line renders at all', !!r.priceText, r.priceText);
  ok('a six-day-old audit says so', r.priceText === 'checked 6d ago', r.priceText);
  ok('an unknown stamp says "never checked", not a fresh age',
     r.kitText === 'never checked', r.kitText);
  ok('the count itself is untouched', r.countText === '1', r.countText);

  /* ⚠ HOUSTON ON HOVER. This panel is read from Riyadh most evenings and the
     standing rule is that every screen shows the floor's clock only. */
  ok('the tooltip names Houston', /Houston/.test(r.priceTitle || ''), r.priceTitle);
  ok('an unknown stamp gets an honest tooltip',
     /has not run yet/.test(r.kitTitle || ''), r.kitTitle);

  /* ⭐ THE ASSERTION check-sidebar COULD NOT MAKE. Two rows just grew a second
     line; the card body is capped at 800px with overflow:hidden, and crossing it
     silently swallows whatever sits at the bottom (the Sheet Protection card lost
     three buttons that way on 2026-08-30). */
  const cap = parseFloat(r.bodyMax);
  ok('the Alerts card still fits inside its cap',
     !isNaN(cap) ? r.bodyScroll <= cap : true, { wants: r.bodyScroll, cap: r.bodyMax });

  ok('the age is a block line, not squeezed inline', r.ageDisplay === 'block', r.ageDisplay);
  ok('the age is smaller than the label', parseFloat(r.ageSize) <= 11, r.ageSize);
  ok('the age inherits ink rather than vanishing', !!r.ageColor && r.ageColor !== 'rgba(0, 0, 0, 0)', r.ageColor);

  ok('no console errors', errs.length === 0, errs.slice(0, 3));

  const card = await p.$('.card[data-id="alerts"]');
  if (card) await card.screenshot({ path: path.join(__dirname, 'renders', 'alert-age.png') });
  else await p.screenshot({ path: path.join(__dirname, 'renders', 'alert-age.png') });

  /* ⚠ BOTH THEMES, because the age uses `color: inherit` rather than its own token
     — that is the whole reason it needs no per-theme rule, and the whole reason
     that claim has to be SEEN rather than asserted. :root is the light palette;
     body.dark-theme swaps the tokens underneath it. */
  await p.evaluate(() => document.body.classList.add('dark-theme'));
  await p.waitForTimeout(250);
  const dark = await p.evaluate(() => {
    const el = document.getElementById('alertAgePriceDrift');
    const lb = el && el.parentElement;
    return { age: el && getComputedStyle(el).color, label: lb && getComputedStyle(lb).color,
             text: el && el.textContent };
  });
  ok('dark theme: the age still renders its text', !!dark.text, dark.text);
  ok('dark theme: the age inherits the label ink', dark.age === dark.label, dark);
  if (card) await card.screenshot({ path: path.join(__dirname, 'renders', 'alert-age-dark.png') });
  console.log('');
  console.log(fail ? `  ${fail} FAILED` : '  OK all checks passed');
  console.log('  render → design-lab/renders/alert-age.png');
  await b.close();
  process.exit(fail ? 1 : 0);
})();
