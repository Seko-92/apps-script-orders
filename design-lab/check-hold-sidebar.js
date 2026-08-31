/**
 * check-hold-sidebar.js — the sheet's door to acknowledging a hold.
 *
 * ⚠ THE POINT: the button is DELIBERATELY absent when nothing is held (this
 * panel's grammar — a healthy state shows nothing, not a disabled control). So
 * "I can't find it" is the expected experience on a quiet day. What must be
 * proven is that it APPEARS the moment something is held, and says enough that
 * nobody has to be told how to use it.
 */
'use strict';
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const SRC = process.env.SIDEBAR_SRC || path.join(__dirname, '..', 'Sidebar.html');
const HTML = fs.readFileSync(SRC, 'utf8').replace("'<?!= boardApiUrl ?>'", "''");

let pass = 0, fail = 0;
const ok = (n, c, x) => { c ? (pass++, console.log('  ✓ ' + n))
                            : (fail++, console.log('  ✗ ' + n + (x !== undefined ? '  → ' + JSON.stringify(x) : ''))); };

const alerts = held => ({
  paidShipping:{count:2,rows:[]}, intl:{count:0,rows:[]}, lowStock:{count:1,rows:[]},
  notFound:{count:0,rows:[]}, queueSize:{count:3,rows:[]}, outOfStock:{count:107,rows:[]},
  newFromZoho:{count:0,rows:[]}, priceDrift:{count:0,rows:[]}, kitPriceDrift:{count:0,rows:[]},
  openCases:{count:0,rows:[]}, needPhotos:{count:460,rows:[]},
  heldOrders:{count:held, rows:[8,9,10].slice(0,held)}
});

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport:{width:310,height:2600}, deviceScaleFactor:2 });
  const errs = []; p.on('pageerror', e => errs.push(String(e)));

  await p.addInitScript(() => {
    window.__held = 0;
    const mk = (su, fa) => new Proxy({}, { get(_, k) {
      if (k === 'withSuccessHandler') return f => mk(f, fa);
      if (k === 'withFailureHandler') return f => mk(su, f);
      return (...args) => {
        let v = null;
        if (k === 'getDisplayUrls') v = { board:'https://hq.yassinqurabi.com/', wall:'https://hq.yassinqurabi.com/wall', hosted:true };
        if (k === 'getActionableAlerts') v = window.__mkAlerts(window.__held);
        if (k === 'acknowledgeSelectedHold') { window.__ackCalled = (window.__ackCalled||0)+1;
                                               v = { ok:true, orders:1, rows:3, message:'✅ Acknowledged 1 hold · 3 rows stamped.' }; }
        if (su) setTimeout(() => su(v), 0);
      };
    }});
    window.google = { script:{ run: mk(null,null), host:{close(){},setHeight(){}},
                               url:{ getLocation(f){ f({parameter:{}}); } } } };
  });
  await p.addInitScript(`window.__mkAlerts = ${alerts.toString()}`);
  await p.route('http://hq.test/**', r => r.fulfill({ contentType:'text/html; charset=utf-8', body: HTML }));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2200);

  const read = () => p.evaluate(() => {
    const act = document.getElementById('alertHoldAct');
    const row = document.getElementById('alertHeldOrders');
    const btn = act && act.querySelector('button');
    return {
      actExists:  !!act,
      actVisible: !!(act && act.offsetParent !== null),
      btnEnabled: !!(btn && !btn.disabled && btn.offsetParent !== null),
      rowEmpty:   !!(row && row.classList.contains('empty')),
      count:      (document.getElementById('alertCountHeldOrders')||{}).textContent || '',
      label:      row ? row.textContent.replace(/\s+/g,' ').trim() : '',
      btnText:    btn ? btn.textContent.replace(/\s+/g,' ').trim() : '',
      hint:       act ? (act.querySelector('.empty-hint')||{}).textContent || '' : '',
      firstAlert: (document.querySelector('.card[data-id="alerts"] .alert-row')||{}).id || '',
      hero:       !!document.querySelector('.card[data-id="alerts"].has-attention'),
      ackCalls:   window.__ackCalled || 0
    };
  });

  console.log('A · nothing held — the button is there and TAPPABLE, like every other one');
  {
    /* ⚠ IT WENT THROUGH THREE SHAPES. Rendered-only-when-held could not be
       LEARNED (which is why it was not found the first time it was wanted).
       Always-there-but-changing-volume was better, and still wrong for the
       reason the user gave last: a control whose APPEARANCE depends on polled
       state can be WRONG, because the count lags by up to a minute and can
       stall. So it is now stateless — the truth is checked fresh at tap time. */
    const r = await read();
    ok('present on a quiet day', r.actVisible, r.actVisible);
    ok('  … and tappable, not disabled', r.btnEnabled, r.btnEnabled);
    ok('the alert row is dimmed', r.rowEmpty, r.rowEmpty);
  }

  console.log('\nB · two holds land — the button appears');
  {
    await p.evaluate(() => { window.__held = 2; refreshAlerts(); });
    await p.waitForTimeout(600);
    const r = await read();
    ok('the button is unchanged by the count — it is a door, not a signal',
       r.btnEnabled && r.actVisible, [r.btnEnabled, r.actVisible]);
    ok('the label never changes: ' + JSON.stringify(r.btnText),
       r.btnText, '✔ Acknowledge hold');
    ok('and the hint warns before the label, not after',
       /do not buy a label yet/i.test(r.hint), r.hint);
    // Copy changed when the single-hold shortcut landed — it now leads with the
    // easy case (just tap) and mentions the selection only for the ambiguous one.
    ok('and says how to use it without being told',
       /just tap/i.test(r.hint) && /select a row/i.test(r.hint) && /Pick ID/i.test(r.hint), r.hint);
    ok('and is honest that it does not LIFT the hold',
       /does not lift the hold/i.test(r.hint), r.hint);
    ok('the count reads 2', r.count, '2');
    ok('the row is no longer dimmed', !r.rowEmpty, r.rowEmpty);
    ok('holds LEAD the alerts card', r.firstAlert, 'alertHeldOrders');
    // ⭐ THE BUTTON LIVES IN FULFILLMENT NOW, above Print Pick List — a hold
    // exists to stop a label being bought, so the control belongs physically
    // between the picker and the button that starts that.
    const place = await p.evaluate(() => {
      const act = document.getElementById('alertHoldAct');
      const card = act && act.closest('.card');
      const body = act && act.parentElement;
      const kids = body ? Array.prototype.map.call(body.children, n => n.id || n.className) : [];
      return { card: card ? card.dataset.id : null, idx: kids.indexOf('alertHoldAct'),
               printIdx: kids.findIndex(c => /btn-primary/.test(String(c))),
               open: card ? !card.classList.contains('collapsed') : null };
    });
    ok('it sits in the FULFILLMENT card', place.card, place.card);
    ok('above Print Pick List', place.idx < place.printIdx || place.printIdx === -1
       ? true : false, place);
    ok('and the card opened itself for it', place.open, place.open);
    ok('the card pulses (hero alert)', r.hero, r.hero);
  }

  console.log('\nC · tapping it calls the server exactly once');
  {
    await p.click('#alertHoldAct button');
    await p.waitForTimeout(700);
    const r = await read();
    ok('one call to acknowledgeSelectedHold', r.ackCalls, 1);
  }

  console.log('\n⚠ E · a FRESH tick beats a STALE alerts cache');
  {
    // The real 2026-08-21 report: getActionableAlerts lives in the 900s half, so
    // its heldOrders count can be up to 15 minutes behind — while the SAME tick
    // carries a `held` array rebuilt every publish. A hold is the only row in
    // this card with a deadline, so it must read the fresh one.
    await p.evaluate(() => {
      window.__held = 0;                       // the cache still says nothing is held
      _sidebarPaint({}, '', null, window.__mkAlerts(0),
                    '', { held: [ { orderId: '24-1', acked: false },
                                  { orderId: '24-2', acked: false },
                                  { orderId: '24-3', acked: true } ] });
    });
    await p.waitForTimeout(400);
    const r = await read();
    ok('the stale 0 is overridden by the live tick', r.count, '2');
    ok('  … acknowledged ones are not counted', r.count !== '3', r.count);
    ok('  … and the alert row shows the fresh number', r.count, '2');
  }
  {
    // ⚠ THE LIVE-FALLBACK TIER carries no `held`, so the cached count must still
    // be honoured rather than silently zeroed.
    await p.evaluate(() => {
      _sidebarPaint({}, '', null, window.__mkAlerts(4), '', { /* no held key */ });
    });
    await p.waitForTimeout(400);
    ok('with no held key, the cached count still shows', (await read()).count, '4');
  }

  console.log('\nD · once answered, it puts itself away');
  {
    await p.evaluate(() => { window.__held = 0;
      _sidebarPaint({}, '', null, window.__mkAlerts(0), '', { held: [] }); });
    await p.waitForTimeout(600);
    const r = await read();
    ok('the button is still THERE and still tappable', r.actVisible && r.btnEnabled,
       [r.actVisible, r.btnEnabled]);
    ok('row dimmed again', r.rowEmpty, r.rowEmpty);
  }

  ok('no console errors', errs.length === 0, errs.slice(0,2));
  console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
  await b.close();
  process.exit(fail ? 1 : 0);
})();
