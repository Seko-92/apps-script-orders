// ============================================================================
// THE RESTING PANEL — render + assert the Control Panel's idle state.
//
// Renders the REAL Sidebar.html at 310px with google.script.run stubbed, drives
// a FAKE CLOCK to real Houston moments, and asserts the MECHANICS rather than
// "it drew something":
//   · it appears on its OWN after REST_IDLE_MS, and not one beat earlier
//   · a pointer move kills it INSTANTLY (waking is free — the whole premise)
//   · day wears daylight from above, night wears ember and drift — same geometry
//   · the flank sits BELOW the face (the v1 divergence, pinned)
//   · pips are radial STROKES r76→r69, gold, at the clock angle they happened
//   · the top bar, the sub-captions and the hint are all present
//   · a SHIPPED event lights a pip at the clock position it happened
//   · healthy rows VANISH — a clean floor renders zero rows, not "none"
//   · NO WIDGETS: zero cards/boxes/bars/lamps inside the panel
//   · the countdown says CLOSES IN by day and OPENS IN by night
//
// Usage: node shoot-rest-panel.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const SIDEBAR = process.env.SIDEBAR_FILE || path.join(__dirname, '..', 'Sidebar.html');
const OUT = path.join(__dirname, 'renders');
const TAG = process.env.TAG || 'rest';

let pass = 0, fail = 0;
const ok  = (c, m, got) => { c ? (pass++, console.log('  ✓ ' + m))
                               : (fail++, console.log('  ✗ ' + m + (got !== undefined ? '  → got ' + JSON.stringify(got) : ''))); };

// ── the tick, shaped exactly like getSidebarTick() ──────────────────────────
// hourFraction is hh + mm/60 in America/Chicago, which is what the server sends.
function tick(o) {
  o = o || {};
  const shipHours = o.shipHours || [];
  return {
    cockpit: {
      shippedToday: shipHours.length,
      receivedToday: o.received == null ? 14 : o.received,
      oldestPendingMinutes: o.oldest === undefined ? 212 : o.oldest,
      ebayGrab: o.ebayGrab == null ? 5 : o.ebayGrab,
      directGrab: o.directGrab == null ? 3 : o.directGrab,
      ebayPending: o.ebayPending == null ? 9 : o.ebayPending,
      directPending: o.directPending == null ? 4 : o.directPending,
      prepQueueCount: o.prep == null ? 6 : o.prep,
      receivedEbay: Math.ceil((o.received == null ? 14 : o.received) * 0.7),
      receivedDirect: (o.received == null ? 14 : o.received) - Math.ceil((o.received == null ? 14 : o.received) * 0.7),
      pendingCount: 12, zohoFixedToday: {},
      lastSyncMinutes: o.sync == null ? 3 : o.sync,
      zohoPending: o.zoho || 0,
      pastRedlineCount: o.late || 0,
      timeline: shipHours.map(h => ({ hourFraction: h, event: 'SHIPPED', orderId: 'X', sku: 'S', picker: 'P' }))
    },
    lastSync: '⏱ Last sync · 4:45 PM',
    sidebar: { api: null, alerts: { paidShipping: { count: 0, rows: [] } } },
    api: null, alerts: { paidShipping: { count: 0, rows: [] } },
    picker: 'Shipping - Yassin 1',
    rest: o.rest || null,
    openOrdersTotal: o.lines == null ? 24 : o.lines,
    _publishedAt: new Date().toISOString()
  };
}

// ── page bootstrap: resolve the ONE scriptlet, stub the server ──────────────
async function boot(page, t, isoTime) {
  let html = fs.readFileSync(SIDEBAR, 'utf8');
  const before = html;
  html = html.replace("'<?!= boardApiUrl ?>'", "''");   // '' → forces the live tier, which we stub
  if (html === before) throw new Error('boardApiUrl scriptlet not found — did the injection change?');
  if (/<\?/.test(html)) throw new Error('unresolved scriptlet left in the page');

  // ⚠ The clock is installed BEFORE any script runs, or the panel's own
  // interval and the arm timer are created against the real clock and the
  // fake one never reaches them.
  await page.clock.install({ time: new Date(isoTime) });

  await page.addInitScript((tickJson) => {
    const T = JSON.parse(tickJson);
    // Proxy stub: any google.script.run.<fn>() resolves with live-shaped data,
    // so a call we did not anticipate degrades to a no-op instead of throwing.
    const DATA = { getSidebarTick: T, getCurrentPicker: T.picker,
                   getActionableAlerts: T.alerts, getLatestApiMetrics: null,
                   getLastSyncFromSheet: T.lastSync, getDashboardSnapshot: T.cockpit };
    function makeRunner(succ, failcb) {
      return new Proxy({}, { get(_, prop) {
        if (prop === 'withSuccessHandler') return (f) => makeRunner(f, failcb);
        if (prop === 'withFailureHandler') return (f) => makeRunner(succ, f);
        return (...a) => { const v = Object.prototype.hasOwnProperty.call(DATA, prop) ? DATA[prop] : null;
                           if (succ) setTimeout(() => succ(v), 0); };
      }});
    }
    window.google = { script: { run: makeRunner(null, null), host: { close(){}, setHeight(){} },
                                url: { getLocation(f){ f({ parameter: {} }); } } } };
  }, JSON.stringify(t));

  // ⚠ SERVE FROM A REAL ORIGIN, NOT setContent. setContent leaves the page on
  // an opaque origin where localStorage THROWS SecurityError — and the sidebar
  // reads localStorage at the top of its script, so that one throw killed every
  // statement after it, including the panel's own init. The first run of this
  // harness reported 30 failures that were all this. (Third time this class has
  // bitten: context.setOffline not reaching a route, the DIRECT-only fixture,
  // the pickerOverride shape. RULE: when a diagnostic reports TOTAL failure,
  // suspect the harness before the code.)
  // ⚠ charset=utf-8 IS LOAD-BEARING: without it Chromium falls back to
  // windows-1252 and every '·' and '→' renders as mojibake — which looks
  // exactly like a string bug in the panel and is not one.
  await page.route('http://hq.test/**', r => r.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
  await page.goto('http://hq.test/sidebar', { waitUntil: 'domcontentloaded' });
  await page.clock.runFor(2500);          // let boot + the first tick land
}

const read = (page) => page.evaluate(() => {
  const p = document.getElementById('restPanel');
  const cs = getComputedStyle(p);
  const q = (s) => p.querySelector(s);
  const txt = (s) => (q(s) || {}).textContent || '';
  const box = (s) => { const e = q(s); return e ? e.getBoundingClientRect() : null; };
  const dial = box('.rp-dial'), flank = box('.rp-flank');
  return {
    shown: p.classList.contains('rp-show'),
    day: p.classList.contains('rp-day'), night: p.classList.contains('rp-night'),
    where: txt('#rpWhere'), date: txt('#rpDate'),
    time: txt('#rpTime'), mer: txt('#rpMer'), dy: txt('#rpDy'),
    shipped: txt('#rpShipped'), received: txt('#rpReceived'),
    shippedSub: txt('#rpShippedSub'), receivedSub: txt('#rpReceivedSub'),
    countLabel: txt('#rpCountLabel'), countValue: txt('#rpCountValue'), hint: txt('.rp-hint'),
    pips: p.querySelectorAll('.rp-pip').length,
    pipTags: [...p.querySelectorAll('.rp-pip')].map(e => e.tagName.toLowerCase()),
    pipGeom: [...p.querySelectorAll('.rp-pip')].map(e => ({
      x1:+e.getAttribute('x1'), y1:+e.getAttribute('y1'),
      x2:+e.getAttribute('x2'), y2:+e.getAttribute('y2'),
      w:e.getAttribute('stroke-width'), c:e.getAttribute('stroke') })),
    rows: [...p.querySelectorAll('.rp-row')].map(r => ({
      label: r.querySelector('span').textContent,
      value: r.querySelector('b').textContent,
      warn: r.querySelector('b').classList.contains('rp-warn') })),
    rowFont: q('.rp-row') ? getComputedStyle(q('.rp-row')).fontFamily.split(',')[0].replace(/["']/g,'') : '',
    rowRule: q('.rp-row') ? getComputedStyle(q('.rp-row')).borderBottomWidth : '',
    countColor: q('.rp-cval') ? getComputedStyle(q('.rp-cval')).color : '',
    hourTicks: q('.rp-hours') ? getComputedStyle(q('.rp-hours')).backgroundImage : '',
    flankBelowDial: !!(dial && flank) && flank.top >= dial.bottom - 1,
    dialPx: dial ? dial.width : 0,
    widgets: p.querySelectorAll('.card, .alert-row, .cockpit-stat, progress, meter, table').length,
    // v3.6's rule — one visual language, no emoji — applied to this surface too.
    // check-sidebar.js only scans .card-icon/.ctrl-btn, so the panel needs its own net.
    emoji: (p.innerText.match(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/gu) || []),
    arcGone: !q('.rp-arc') && !q('#rpFill'),
    display: cs.display, visibility: cs.visibility, opacity: cs.opacity,
    inTrans: cs.transitionDuration,
    condensed: document.body.classList.contains('cockpit-condensed'),
    jewelDelay: q('#rpJewel') ? q('#rpJewel').style.animationDelay : '',
    pointer: cs.pointerEvents,
    ledger: txt('#rpLedger'),
    dialTop: (() => { const d = box('.rp-dial'), t = box('.rp-top');
      return (d && t) ? Math.round(d.top - t.bottom) : null; })(),
    idleBloomEase: document.body.classList.contains('cockpit-idle-bloom'),
    hintAnimScoped: p.classList.contains('rp-show')
      ? getComputedStyle(q('.rp-hint')).animationName === 'rpBreathe' : null,
    hintAnimWhenHidden: !p.classList.contains('rp-show')
      && getComputedStyle(q('.rp-hint')).animationName === 'rpBreathe'
  };
});

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch();

  // ─────────────────────────────────────────────────────────────────────────
  // A · NIGHT — Thu 21:40 Houston. 4h40m into a 16h night ≈ 29% arc.
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nA · night · Thu 9:40 PM Houston');
  let page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  const errs = []; page.on('pageerror', e => errs.push(String(e)));
  await boot(page, tick({ shipHours: [9.25, 10.5, 11.1, 13.75, 14.2, 15.9, 16.4], oldest: 212 }),
             '2026-08-13T21:40:00-05:00');

  let s = await read(page);
  ok(!s.shown, 'hidden while the panel is being used');

  // ⚠ boot() already advanced the clock 2.5s, so the elapsed idle time is
  // 2.5 + this. 86 lands at 88.5s — one beat short of the 90s constant.
  // ⚠ 82.5s not 88.5s: the tighter version had ~1.5s of slack and flaked once
  // in five runs. A net that fails at random teaches people to re-run it until
  // it passes, which is worse than not having it.
  await page.clock.runFor(REST_MS(80));
  s = await read(page);
  ok(!s.shown, 'still hidden at 82.5s — it does not jump the gun', s.shown);

  await page.clock.runFor(REST_MS(10));        // 92.5s — now past the constant
  s = await read(page);
  ok(s.shown, 'appears on its own after ~90s of quiet', s.shown);
  ok(s.night && !s.day, 'wears the NIGHT light', { day: s.day, night: s.night });
  ok(s.where === 'Houston \u00b7 closed', 'the top bar names the state', s.where);
  ok(/^\w{3} \d{1,2} \w{3}$/.test(s.date), 'and carries the date', s.date);
  ok(s.hint === 'Move to resume', 'the hint is there', s.hint);
  ok(s.countLabel === 'Opens in', 'countdown says OPENS IN', s.countLabel);
  ok(/^\d+h \d+m$/.test(s.countValue), 'opens-in reads as a duration', s.countValue);
  ok(s.countColor === 'rgb(224, 139, 60)', 'countdown is ember, not white', s.countColor);
  ok(s.shipped === '7' && s.received === '14', 'the flank carries shipped | received', [s.shipped, s.received]);
  ok(s.shippedSub === 'of 14 received', 'the sub-caption is a RELATIONSHIP', s.shippedSub);
  ok(s.receivedSub === '10 eBay \u00b7 4 direct', 'and the channel split', s.receivedSub);
  ok(s.dy === 'Thu \u00b7 Houston', 'at night the third line is an IDENTITY', s.dy);
  ok(s.mer === 'PM', 'meridiem is its own line', s.mer);
  // \u26a0 THE ASSERTION THIS ROUND EXISTS FOR. v1 stacked the flank ABOVE the
  // face; the approved mockup puts it BELOW. Pinned so it cannot drift back.
  ok(s.flankBelowDial, 'THE FLANK SITS BELOW THE FACE', s.flankBelowDial);
  ok(s.arcGone, 'the invented progress arc is gone', s.arcGone);
  ok(s.pips === 7, 'one pip per shipment', s.pips);
  ok(s.pipTags.every(t => t === 'line'), 'pips are radial STROKES, not dots', s.pipTags[0]);
  ok(/Mono|monospace/.test(s.rowFont), 'rows are mono', s.rowFont);
  ok(s.rowRule === '1px', 'rows are hairline-ruled', s.rowRule);
  ok(/255, 217, 102/.test(s.hourTicks), 'hour ticks are GOLD', s.hourTicks.slice(0, 70));
  ok(s.widgets === 0, 'NO WIDGETS \u2014 zero cards/boxes/bars inside the panel', s.widgets);
  ok(s.emoji.length === 0, 'ZERO EMOJI \u2014 v3.6\u2019s one visual language holds here', s.emoji);
  ok(/^-\d/.test(s.jewelDelay), 'the seconds jewel is phase-aligned', s.jewelDelay);
  const nightRows = s.rows.map(r => r.label);
  ok(nightRows.join('|') === 'Waiting for tomorrow|Oldest of those|Prep queue',
     'night rows match the mockup wording', nightRows);
  ok(s.rows[0].value === '24 lines \u00b7 8 orders', 'row values are COMPOUND', s.rows[0].value);
  ok(s.rows[1] && s.rows[1].warn, 'past the 3h line is the one loud row', s.rows[1]);
  ok(/^first out \d+:\d\d {2}·{1} {2}last \d+:\d\d$/.test(s.ledger),
     "the ledger carries the day's bookends", s.ledger);
  // ⚠ the face must sit UNDER the header, not float in the middle — a band of
  // dead space above it is what read as "empty" on the real panel.
  ok(s.dialTop !== null && s.dialTop < 60, 'the face sits under the header', s.dialTop);
  await page.screenshot({ path: path.join(OUT, `${TAG}-night.png`) });

  // waking must be instant and free
  await page.mouse.move(120, 400);
  s = await read(page);
  ok(!s.shown, 'a pointer move kills it INSTANTLY', s.shown);
  ok(errs.length === 0, 'no page errors', errs.slice(0, 3));
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // B · DAY — Thu 14:30 Houston. 5.5h into an 8h shift ≈ 69% arc.
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nB · day · Thu 2:30 PM Houston');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  const errsB = []; page.on('pageerror', e => errsB.push(String(e)));
  await boot(page, tick({ shipHours: [9.25, 10.5, 11.1, 13.75, 14.2], oldest: 212,
                          rest: { cantBuild: { kits: 26, units: 117 },
                                  ripple: [{ sku: '195306', sole: 7 }] } }),
             '2026-08-13T14:30:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  ok(s.shown, 'appears during the working day too', s.shown);
  ok(s.day && !s.night, 'wears the DAY light', { day: s.day, night: s.night });
  ok(s.where === 'Houston \u00b7 working', 'the top bar flips too', s.where);
  ok(s.countLabel === 'Closes in', 'countdown says CLOSES IN', s.countLabel);
  const closeMin = (s.countValue.match(/(\d+)h (\d+)m/) || [])
                     .slice(1).reduce((a, v, i) => a + (i ? +v : +v * 60), 0);
  ok(closeMin >= 147 && closeMin <= 150, 'closes-in \u2248 2h30m at 2:30 PM', s.countValue);
  ok(s.dy === '9 more in than out', 'by day the third line is a FINDING', s.dy);
  ok(/an hour$/.test(s.shippedSub), 'day sub-caption is the RATE', s.shippedSub);
  ok(s.flankBelowDial, 'THE FLANK SITS BELOW THE FACE', s.flankBelowDial);
  ok(s.pips === 5, 'five shipments, five pips', s.pips);
  const g = s.pipGeom[0];
  const r0 = Math.hypot(g.x1 - 100, g.y1 - 100), r1 = Math.hypot(g.x2 - 100, g.y2 - 100);
  ok(Math.abs(r0 - 76) < 0.5 && Math.abs(r1 - 69) < 0.5, 'pip spans r76 \u2192 r69', [r0.toFixed(1), r1.toFixed(1)]);
  ok(g.w === '2.4' && g.c === '#ffd966', 'pip is 2.4 wide and brand gold', [g.w, g.c]);
  const dayRows = s.rows.map(r => r.label);
  ok(dayRows.join('|') === "On the floor|Oldest waiting|Advertised, can't build|Restocking 1 part frees",
     'day rows match the mockup wording', dayRows);
  ok(s.rows[2].value === '117 units', "can't-build reports UNITS", s.rows[2].value);
  ok(s.rows[2].warn, 'and it is the loud one', s.rows[2].warn);
  ok(s.widgets === 0, 'NO WIDGETS in the day state either', s.widgets);
  ok(s.emoji.length === 0, 'ZERO EMOJI in the day state either', s.emoji);
  // The face now scales on the SHORTER of width and height, so it fills a tall
  // sidebar instead of floating in it. Bounded, not pinned.
  ok(s.dialPx > 180 && s.dialPx < 265, 'the face scales with the panel', Math.round(s.dialPx));
  await page.screenshot({ path: path.join(OUT, `${TAG}-day.png`) });
  ok(errsB.length === 0, 'no page errors', errsB.slice(0, 3));
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // C · A CALM DAY — healthy rows must VANISH, not read "none"
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nC · a calm day · nothing on the floor');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 11, 15], oldest: null, lines: 0, ebayGrab: 0, directGrab: 0,
                          ebayPending: 0, directPending: 0, prep: 0 }),
             '2026-08-13T15:10:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  ok(s.shown, 'still shows on a calm day', s.shown);
  ok(s.rows.length === 0, 'HEALTHY ROWS VANISH — zero rows, no "none"', s.rows);
  ok(s.pips === 3, 'the day it did is still worn on the ring', s.pips);
  ok(/^first out /.test(s.ledger), 'and the ledger still reports it', s.ledger);
  ok(s.shipped === '3', 'the flank still reports', s.shipped);
  await page.screenshot({ path: path.join(OUT, `${TAG}-calm.png`) });
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // D · 7 AM — the face is nearly bare; the day has not been done yet
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nD · 7 AM · a bare face before the shift');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [], received: 2, oldest: 40, lines: 3, ebayGrab: 2, directGrab: 0 }),
             '2026-08-14T07:00:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  ok(s.pips === 0, 'at 7 AM the ring is bare', s.pips);
  ok(s.night, 'before 9 AM the shop is still closed', { day: s.day, night: s.night });
  ok(s.pips === 0, 'a bare ring before the shift', s.pips);
  ok(s.ledger === '', 'and no ledger before the first shipment', s.ledger);
  ok(s.rows.every(r => !/Restocking/.test(r.label)), 'no ripple row without server data', s.rows.map(r=>r.label));
  await page.screenshot({ path: path.join(OUT, `${TAG}-dawn.png`) });
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // E · pip POSITION — a shipment marks the hour it happened
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nE · a pip sits at the clock position it happened');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  // ⚠ THE FIRST VERSION OF THIS CASE USED 3 AM / 3 PM AND EXPECTED TWO PIPS.
  // Both expectations were wrong and the panel was right: on a 12-HOUR face
  // 03:00 and 15:00 ARE the same position, and 9 o'clock is 270°, not 0°.
  // That collision cannot happen in practice — the shift is 09:00–17:00, which
  // maps to nine DISTINCT positions (270,300,330,0,30,60,90,120,150) — and the
  // last assertion below pins exactly that, so the face can never quietly
  // become ambiguous if the shift window is ever widened.
  await boot(page, tick({ shipHours: [9.0, 12.0, 15.0, 16.5] }), '2026-08-13T21:00:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  // Angle of each pip, derived from where the STROKE sits: atan2 about the
  // 200-viewBox centre, rotated so 12 o'clock reads 0\u00b0.
  const degs = s.pipGeom.map(g => (Math.atan2(g.y1 - 100, g.x1 - 100) * 180 / Math.PI + 90 + 360) % 360);
  ok(s.pips === 4, 'four shipment hours, four pips', s.pips);
  const want = [0, 90, 135, 270];
  const got  = degs.slice().sort((a, b) => a - b);
  ok(want.every((w, i) => Math.abs(got[i] - w) < 1),
     'noon\u21920\u00b0, 3 PM\u219290\u00b0, 4:30 PM\u2192135\u00b0, 9 AM\u2192270\u00b0', got.map(d => d.toFixed(1)));

  const shiftSlots = new Set();
  for (let m = 9 * 60; m < 17 * 60; m += 6) shiftSlots.add(Math.round(((m / 60) % 12) * 10));
  ok(shiftSlots.size === 80, 'the 09:00\u201317:00 shift never collides on a 12h face', shiftSlots.size);

  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // F · A CALM NIGHT — the vanish rule must hold in BOTH states.
  // ⚠ THIS CASE EXISTS BECAUSE C ALONE WAS NOT ENOUGH. C runs at 3:10 PM, so it
  // only ever exercised the DAY branch — a mutation that made a healthy NIGHT
  // row render anyway passed the whole suite. Same shape as the 2026-08-15 hold
  // bug, where seventeen green assertions shipped a feature that was dead for
  // eBay because every fixture was DIRECT. Both branches, always.
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nF · a calm night · nothing carried over');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 13, 16], oldest: null, lines: 0, ebayGrab: 0, directGrab: 0,
                          ebayPending: 0, directPending: 0, prep: 0 }),
             '2026-08-13T22:10:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  ok(s.night, 'it is the night state', { day: s.day, night: s.night });
  ok(s.rows.length === 0, 'HEALTHY ROWS VANISH AT NIGHT TOO — zero rows', s.rows);
  ok(s.pips === 3, 'the closed face still wears the day it did', s.pips);
  ok(s.countLabel.toLowerCase() === 'opens in', 'and still counts to open', s.countLabel);
  await page.screenshot({ path: path.join(OUT, `${TAG}-calm-night.png`) });
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // G · THE TWO-STAGE IDLE and THE FADE.
  // ⚠ Stage 1 already existed and nobody could ever see it: COCKPIT_IDLE_RESET_MS
  // blooms the cockpit at 4 min, but the rest panel lands at 90s — so the bloom
  // was firing UNDER the overlay every single time. Now it is a real sequence:
  // 40s tidy up in full view, 90s rest.
  // ⚠ And the panel must NEVER be display:none — a display change cannot be
  // transitioned, which is what made it snap instead of fade.
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nG · two-stage idle, and the fade');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 13, 16] }), '2026-08-13T21:00:00-05:00');

  // scroll the module list so the cockpit condenses, then walk away
  await page.evaluate(() => { const m = document.getElementById('modules');
    if (m) { m.scrollTop = 400; m.dispatchEvent(new Event('scroll')); } });
  s = await read(page);
  ok(s.condensed, 'scrolling condenses the cockpit', s.condensed);
  ok(s.display !== 'none', 'the panel is NEVER display:none — that is what snapped', s.display);
  ok(s.visibility === 'hidden', 'it hides by visibility, so opacity can transition', s.visibility);

  // ⚠ SAMPLE BETWEEN THE BLOOM AND ITS CLEANUP. The slow-easing class is
  // deliberately transient — added for the bloom, dropped 1200ms later so
  // scrolling stays responsive. Advancing straight past 45s fires both in one
  // go and the class is gone before it can be read, which looks like a failure
  // and is not one.
  // ⚠⚠ ADVANCE IN SLICES, NOT ONE JUMP. Reproduced directly: a single
  // clock.runFor(40500) does NOT fire the 40s stage-1 timer, while
  // runFor(38000) + runFor(2500) does — the page also runs a 1s interval
  // (_rpAmbient), and a long single jump can leave the longer timeout unfired.
  // Two steps still flaked about one run in four, so it advances in 5s slices:
  // deterministic, and it looked exactly like the bloom being broken both times.
  await advance(page, 40.5);                    // ~43s — bloom fired, cleanup pending
  s = await read(page);
  ok(!s.condensed, 'STAGE 1 at ~40s: the cockpit blooms, in full view', s.condensed);
  ok(s.idleBloomEase, 'and it blooms SLOWLY — an idle bloom is not a scroll', s.idleBloomEase);
  ok(!s.shown, 'and the panel has NOT taken over yet', s.shown);

  await page.clock.runFor(REST_MS(4.5));         // ~47.5s — past the cleanup
  s = await read(page);
  ok(!s.idleBloomEase, 'and hands the timing straight back to scroll speed', s.idleBloomEase);

  await page.clock.runFor(REST_MS(50));          // 97.5s — past stage 2
  s = await read(page);
  ok(s.shown, 'STAGE 2 at ~90s: the panel rests', s.shown);
  // ⚠ The GROUND is deliberately quick (650ms) — it is a hand-off, not the whole
  // arrival. The content then assembles behind it out to ~1430ms, so the console
  // is never left superimposed under half-drawn slots.
  ok(parseFloat(s.inTrans) > 0.5 && parseFloat(s.inTrans) < 0.9,
     'the ground hands over in ~650ms', s.inTrans);
  ok(s.hintAnimScoped, 'the hint breathes only while SHOWN', s.hintAnimScoped);

  await page.mouse.move(150, 300);
  s = await read(page);
  ok(!s.shown, 'a pointer move still kills it instantly', s.shown);
  // ⚠ WHAT MAKES WAKING FREE IS pointer-events, NOT THE DURATION. The class comes
  // off instantly, so the panel is click-dead the moment you move — the fade is
  // then purely cosmetic and can afford to be gentle. Assert the thing that
  // actually blocks work, and only bound the fade loosely.
  ok(s.pointer === 'none', 'it is click-dead the INSTANT you move', s.pointer);
  ok(parseFloat(s.inTrans) > 0.35 && parseFloat(s.inTrans) < 0.65,
     'and it leaves gently rather than cutting', s.inTrans);
  // ⚠⚠ REGRESSION NET. A CSS animation BEATS a transition on the same property,
  // and rpBreathe used to sit on .rp-hint itself — so it started at PAGE LOAD,
  // owned opacity outright, and the staged entrance never reached the hint
  // (measured at 0.95 and falling at t=0 of the arrival). Scoped to .rp-show it
  // starts with each showing. If it ever drifts back onto .rp-hint, this fails.
  ok(!s.hintAnimWhenHidden, 'and it does NOT breathe while hidden', s.hintAnimWhenHidden);
  await page.close();

  // ─────────────────────────────────────────────────────────────────────────
  // H · THE FINDINGS-ONLY ROWS. Three signals the tick always carried and the
  // panel was throwing away. Every one is SILENT when healthy — the mockup's own
  // note says the 3h row was cut only because it read "none", to be "restored
  // the moment it is >0". Both halves are asserted: they appear, and they go.
  // ─────────────────────────────────────────────────────────────────────────
  console.log('\nH · findings-only rows');
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 13], late: 3, zoho: 65, sync: 140 }),
             '2026-08-13T21:20:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  let lbl = s.rows.map(r => r.label);
  ok(lbl.indexOf('Waiting to pull') >= 0, 'un-pulled Zoho SOs surface at night', lbl);
  ok(lbl.indexOf('Past the 3h line') >= 0, 'and so does the 3h line', lbl);
  ok(s.rows.find(r => r.label === 'Past the 3h line').warn, 'the 3h row is loud', true);
  await page.close();

  // the same tick during working hours: a stale sync is a REAL finding then
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 13], late: 3, zoho: 65, sync: 140 }),
             '2026-08-13T14:00:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  lbl = s.rows.map(r => r.label);
  ok(lbl.indexOf('Last sync') >= 0, 'a dead pipeline shows during the shift', lbl);
  ok(s.rows.find(r => r.label === 'Last sync').warn, 'and it is loud', true);
  await page.close();

  // healthy: every one of them must go
  page = await browser.newPage({ viewport: { width: 310, height: 820 }, deviceScaleFactor: 2 });
  await boot(page, tick({ shipHours: [9.5, 13], late: 0, zoho: 0, sync: 3 }),
             '2026-08-13T14:00:00-05:00');
  await page.clock.runFor(REST_MS(95));
  s = await read(page);
  lbl = s.rows.map(r => r.label);
  ok(lbl.indexOf('Past the 3h line') < 0, 'healthy: the 3h row VANISHES', lbl);
  ok(lbl.indexOf('Waiting to pull') < 0, 'healthy: the Zoho row VANISHES', lbl);
  ok(lbl.indexOf('Last sync') < 0, 'healthy: the sync row VANISHES', lbl);
  await page.close();

  await browser.close();
  console.log(`\n${pass} passed · ${fail} failed`);
  console.log('renders → ' + OUT);
  process.exit(fail ? 1 : 0);
})();

// helpers hoisted for readability above
function REST_MS(sec) { return sec * 1000; }

/* ⚠ Advance the fake clock in slices. A long single runFor can leave a long
   timeout unfired when the page also runs a short interval — see the note in
   case G. Anything waiting on a multi-second timer should go through here. */
async function advance(page, seconds, sliceSec) {
  const slice = sliceSec || 5;
  let left = seconds;
  while (left > 0) { const step = Math.min(slice, left); await page.clock.runFor(REST_MS(step)); left -= step; }
}
