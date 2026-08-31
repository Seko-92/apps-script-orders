/**
 * diag-holdstop.js — THE HOLD THAT REACHES A BOX NOBODY IS PICKING ANY MORE.
 *
 * The 2026-08-21 incident: a buyer asked to change the shipping service, the
 * announcement went to WhatsApp, the pickers were busy and missed it, the label
 * was bought and the row went SHIPPED — at which point the board stopped
 * showing the order at all, because _dashOpenOrders filters to PENDING and
 * PREPARING. The box sat in the building for hours with no surface anywhere in
 * the system saying stop.
 *
 * ⚠ RUN IT AGAINST HEAD TO SEE IT BITE:
 *     BOARD_FILE=/tmp/head.html node diag-holdstop.js
 *   (git show HEAD:FloorBoard.html > /tmp/head.html)
 *
 * ⚠ EVERY CHECK FAILS SOFT. A before/after proof is only useful if every
 * section reports — the choosePicker lesson, re-learned in diag-pickgate.
 */
'use strict';
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');

const BOARD = process.env.BOARD_FILE || path.join(__dirname, '..', 'FloorBoard.html');
const html  = fs.readFileSync(BOARD, 'utf8');
const BASE  = JSON.parse(JSON.stringify(require('./mock-tick.js')));

// Thursday 2:30 PM Houston — inside working hours, so the siren is allowed to
// sound. Off-hours the board is deliberately visual-only.
const T0 = '2026-08-20T19:30:00Z';

let pass = 0, fail = 0;
const check = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const soft = async (label, fn) => {
  try { return await fn(); }
  catch (e) { fail++; console.log('  ✗ ' + label + '  → threw: ' + String(e.message || e).split('\n')[0]); }
};

// One held order, shaped exactly as Holds.holdScanRows emits it.
const held = (o, over) => Object.assign({
  orderId: o, channel: 'EBAY', note: 'HOLD — buyer wants expedited, change from Ground to 2-Day',
  acked: false, ackText: '', shipped: true, urgent: true, lines: 3,
  items: [ { sku: '165447', qty: 2, loc: 'B-30' },
           { sku: '172764', qty: 1, loc: 'A-14' } ]
}, over || {});

(async () => {
  console.log('BOARD: ' + BOARD + '\n');
  const browser = await chromium.launch();
  const ctx = await browser.newContext({
    viewport: { width: 1280, height: 800 }, hasTouch: true, timezoneId: 'America/Chicago'
  });
  // ⚠ SHORT ACTION TIMEOUT. Against HEAD the hold elements do not exist, and a
  // 30s default per click turns the before/after proof into a five-minute hang.
  ctx.setDefaultTimeout(4000);
  const page = await ctx.newPage();

  // Count real oscillator starts, so "did it make a noise" is measured at the
  // audio layer rather than inferred from a class name.
  await page.addInitScript(() => {
    window.__osc = 0;
    class FakeCtx {
      /* ⚠ STARTS SUSPENDED, like a real browser context on a page nobody has
         touched. The first version of this stub reported 'running' from birth,
         which is precisely why it could not see the async-resume bug: every
         sound scheduled fine against a clock that was already going. A stub that
         is healthier than the real thing tests nothing. */
      constructor() { this.state = 'suspended'; this.currentTime = 0; this.destination = {}; }
      resume() { const self = this; return Promise.resolve().then(() => { self.state = 'running'; }); }
      createGain()   { return { gain: { setValueAtTime(){}, linearRampToValueAtTime(){}, exponentialRampToValueAtTime(){} }, connect(){} }; }
      createBiquadFilter() { return { type: '', frequency: { value: 0 }, connect(){} }; }
      createOscillator() {
        const ctx = this;
        return { type: '', frequency: { value: 0 }, connect(){},
                 // ⚠ ONLY COUNTS WHILE THE CONTEXT IS RUNNING. A note scheduled
                 // against a suspended (frozen) clock lands in the past and is
                 // never heard — which is the bug, so the harness must not count it.
                 start() { if (ctx.state === 'running') window.__osc++; }, stop() {} };
      }
    }
    window.AudioContext = FakeCtx; window.webkitAudioContext = FakeCtx;
  });

  let TICK = Object.assign({}, BASE, { held: [] });
  const calls = [];
  await page.route('http://hqlab.test/**', r => {
    const u = r.request().url();
    if (u.includes('/api/board')) {
      const body = JSON.parse(r.request().postData() || '{}');
      calls.push(body);
      let res = { ok: false };
      if (body.action === 'boardTick')     res = Object.assign({ ok: true }, TICK);
      if (body.action === 'boardStatus')   res = { ok: true };
      if (body.action === 'boardRadio')    res = { ok: true, nowPlaying: '' };
      if (body.action === 'boardAckHold')  res = ackReply;
      return r.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
    }
    return r.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
  });
  await page.route(/aladhan|open-meteo|youtube|ytimg|somafm|walmradio|isekoi/, r => r.abort());

  let ackReply = { ok: true, rows: 3, tag: '✓ SEEN 2:31 PM by Yassin · 1' };

  await page.clock.install({ time: new Date(T0) });
  await page.goto('http://hqlab.test/', { waitUntil: 'load' });
  await page.clock.runFor(1500);
  await page.waitForTimeout(1200);

  const read = () => page.evaluate(() => {
    const strip = document.getElementById('holdStrip');
    const take  = document.getElementById('holdTake');
    return {
      stripShown: !!(strip && strip.classList.contains('show')),
      stripClass: strip ? strip.className : '(missing)',
      rows:       strip ? strip.querySelectorAll('.hold-row').length : -1,
      stripText:  strip ? strip.textContent.replace(/\s+/g, ' ').trim() : '',
      takeShown:  !!(take && !take.classList.contains('hidden')),
      takeOid:    (document.getElementById('htOid') || {}).textContent || '',
      takeState:  (document.getElementById('htState') || {}).textContent || '',
      takeTxt:    (document.getElementById('htTxt') || {}).textContent || '',
      takeMore:   (document.getElementById('htMore') || {}).textContent || '',
      liftShown:  (function(){ var e=document.getElementById('liftStrip');
                               return !!(e && e.classList.contains('show')); })(),
      liftText:   (document.getElementById('liftBody')||{}).textContent || '',
      items:      Array.prototype.map.call(
                    document.querySelectorAll('#htItems .ht-item'),
                    function (n) { return n.textContent.replace(/\s+/g, ' ').trim(); }),
      itemsShown: (function () { var e = document.getElementById('htItems');
                                 return !!(e && e.style.display !== 'none' && e.offsetParent !== null); })(),
      osc:        window.__osc || 0
    };
  });
  const push = async (h) => {
    TICK = Object.assign({}, BASE, { held: h });
    await page.evaluate(() => { if (typeof pollSoon === 'function') pollSoon(); });
    await page.clock.runFor(400);
    await page.waitForTimeout(900);
  };

  // ── A · THE REGRESSION NET ───────────────────────────────────────────────
  // A board with nothing held must look EXACTLY as it did before this feature.
  console.log('A · nothing held — the board is unchanged');
  await soft('A', async () => {
    const r = await read();
    check('  strip not drawn', r.stripShown, false);
    check('  no takeover', r.takeShown, false);
    check('  silent', r.osc, 0);
  });

  // ── B · THE REPORTED BUG ─────────────────────────────────────────────────
  console.log('\nB · HOLD on an order whose label is already bought  ← the incident');
  await soft('B', async () => {
    await push([held('24-14979-87359')]);
    const r = await read();
    check('  strip appears', r.stripShown, true);
    check('  strip is RED (unseen)', /unseen/.test(r.stripClass), true);
    check('  one row', r.rows, 1);
    check('  takeover fires', r.takeShown, true);
    check('  names the order', r.takeOid, '24-14979-87359');
    check('  says the label is bought', /LABEL ALREADY BOUGHT/.test(r.takeState), true);
    check('  shows what was typed', /change from Ground to 2-Day/.test(r.takeTxt), true);
    check('  the siren sounded', r.osc > 0, true);
  });

  // ── C · THE MUTE MUST NOT SILENCE IT ─────────────────────────────────────
  console.log('\nC · a muted board still sirens for a hold');
  await soft('C', async () => {
    await page.evaluate(() => { try { localStorage.setItem('floorSound', '0'); } catch (e) {} });
    await page.evaluate(() => { window.__osc = 0; if (typeof soundOn !== 'undefined') soundOn = false; });
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-15004-11290')]);
    const r = await read();
    check('  sound flag is off', await page.evaluate(() => (typeof soundOn === 'undefined') ? null : soundOn), false);
    check('  the siren sounded anyway', r.osc > 0, true);
    await page.evaluate(() => { try { localStorage.setItem('floorSound', '1'); } catch (e) {} if (typeof soundOn !== 'undefined') soundOn = true; });
  });

  // ── B2 · WHAT IS IN THE BOX ──────────────────────────────────────────────
  // The user's question: on a busy afternoon an order id does not identify a box
  // to a person standing in front of fifteen of them.
  console.log('\nB2 · the takeover names what is in the box');
  await soft('B2', async () => {
    const r = await read();
    check('  the contents are shown', r.itemsShown, true);
    // ⚠ textContent concatenates adjacent spans with no whitespace — the visual
    // spacing is flex `gap`, which does not appear in the text. Assert what the
    // DOM actually holds rather than what the picture looks like.
    check('  line one · qty, sku, shelf', r.items[0], '×2165447B-30');
    check('  line two', r.items[1], '×1172764A-14');
    check('  and says how many it did not list', r.items[2], '+1 more line');
    check('  the strip carries the line count', /3 lines/.test(r.stripText), true);
  });
  console.log('\nB3 · a server that sent no items draws no empty box');
  await soft('B3', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdCloseTakeover(); });
    await push([held('24-00000-00000', { items: [], lines: 0 })]);
    const r = await read();
    check('  the takeover still fires', r.takeShown, true);
    check('  but the contents block is not drawn', r.itemsShown, false);
  });

  // ── C2 · OFF-HOURS: ONE CALL, THEN QUIET ─────────────────────────────────
  // Corrected 2026-08-21 — the first build was SILENT off-hours and the user
  // reported no sound. Silent was wrong in both directions: off-hours is when
  // the box is most certainly still in the building, but a repeat that decays to
  // every 3 min and never expires must not run all night in an empty warehouse.
  console.log('\nC2 · off-hours it still SOUNDS — once — and then stops repeating');
  await soft('C2', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdAcked = {}; holdCloseTakeover();
                                window.__osc = 0; window.isOffHours = function () { return true; }; });
    await push([held('24-77777-77777')]);
    const r = await read();
    check('  the takeover still fires', r.takeShown, true);
    check('  it sounded once', r.osc > 0, true);
    const repeating = await page.evaluate(() => holdSirenT !== null);
    check('  but NO repeat is armed', repeating, false);
    await page.evaluate(() => { window.isOffHours = function () { return false; }; });
  });
  console.log('\nC3 · in hours the repeat IS armed');
  await soft('C3', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdCloseTakeover(); window.__osc = 0; });
    await push([held('24-66666-66666')]);
    check('  sounded', (await read()).osc > 0, true);
    check('  and keeps calling', await page.evaluate(() => holdSirenT !== null), true);
    await page.evaluate(() => holdSirenStop());
  });

  // ── D · A CALM HOLD MUST NOT SCREAM ──────────────────────────────────────
  console.log('\nD · HOLD on a still-PENDING order — strip yes, takeover no');
  await soft('D', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdCloseTakeover(); window.__osc = 0; });
    await push([held('24-99999-00001', { urgent: false, shipped: false })]);
    const r = await read();
    check('  strip appears', r.stripShown, true);
    check('  NO takeover', r.takeShown, false);
    check('  stayed silent', r.osc, 0);
  });

  // ── E · ACKNOWLEDGED ─────────────────────────────────────────────────────
  console.log('\nE · once acknowledged the strip goes calm — but does NOT vanish');
  await soft('E', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdCloseTakeover(); window.__osc = 0; });
    await push([held('24-14979-87359', { acked: true, ackText: '2:31 PM by Yassin · 1',
      note: 'HOLD — buyer wants expedited · ✓ SEEN 2:31 PM by Yassin · 1' })]);
    const r = await read();
    check('  strip STILL shown', r.stripShown, true);
    check('  no longer red', /unseen/.test(r.stripClass), false);
    check('  reads calm', /calm/.test(r.stripClass), true);
    check('  names who saw it', /Yassin/.test(r.stripText), true);
    check('  no takeover', r.takeShown, false);
    check('  silent', r.osc, 0);
  });

  // ── F · THE ACK WRITE ────────────────────────────────────────────────────
  console.log('\nF · ✓ Got it writes the acknowledgement, exactly once');
  await soft('F', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-14979-87359')]);
    calls.length = 0;
    await page.click('#htOk');
    await page.waitForTimeout(900);
    const acks = calls.filter(c => c.action === 'boardAckHold');
    check('  exactly one ack call', acks.length, 1);
    check('  for the right order', acks[0] && acks[0].orderId, '24-14979-87359');
    check('  takeover closed', (await read()).takeShown, false);
  });

  // ── F2 · THE FLOOR'S REPORT: THE STRIP MUST NOT WAIT FOR THE ROUND TRIP ──
  // 2026-08-21: the sheet updated instantly, the strip kept saying TAP TO
  // ACKNOWLEDGE for ~2 minutes. Nothing was broken — the board had no override.
  console.log('\nF2 · after ✓ Got it the strip goes calm AT ONCE  ← the floor report');
  await soft('F2', async () => {
    const r = await read();
    check('  calm immediately, before any new tick', /calm/.test(r.stripClass), true);
    check('  no longer red', /unseen/.test(r.stripClass), false);
    check('  names who saw it', /Yassin/.test(r.stripText), true);
    // ⚠ AND IT HOLDS while the server still reports it unacknowledged — that is
    // the whole 45-95s window the picker was staring at.
    await push([held('24-14979-87359')]);
    const r2 = await read();
    check('  survives a tick that still says unacked', /calm/.test(r2.stripClass), true);
    check('  and no takeover re-fires', r2.takeShown, false);
  });

  // ── F3 · THE OVERRIDE MUST RETIRE ON EVIDENCE ────────────────────────────
  console.log('\nF3 · once the server agrees, the override stands down');
  await soft('F3', async () => {
    await push([held('24-14979-87359', { acked: true, ackText: '2:31 PM by Yassin · 1',
      note: 'HOLD x · ✓ SEEN 2:31 PM by Yassin · 1' })]);
    const gone = await page.evaluate(() => !holdAcked['24-14979-87359']);
    check('  override dropped, server view takes over', gone, true);
    check('  still calm', /calm/.test((await read()).stripClass), true);
  });

  // ── G · "LATER" IS NOT AN ACKNOWLEDGEMENT ────────────────────────────────
  console.log('\nG · Later silences without putting anyone\'s name on it');
  await soft('G', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-15008-33107')]);
    calls.length = 0;
    await page.click('#htLater');
    await page.waitForTimeout(600);
    const r = await read();
    check('  no ack written', calls.filter(c => c.action === 'boardAckHold').length, 0);
    check('  takeover closed', r.takeShown, false);
    check('  strip STILL red', /unseen/.test(r.stripClass), true);
  });

  // ── H · THE STRIP IS THE WAY BACK IN ─────────────────────────────────────
  console.log('\nH · tapping the strip reopens the takeover');
  await soft('H', async () => {
    await page.click('#holdStrip');
    await page.waitForTimeout(500);
    check('  takeover reopened', (await read()).takeShown, true);
    await page.evaluate(() => holdCloseTakeover());
  });

  // ── I · A FAILED ACK RE-ARMS ─────────────────────────────────────────────
  console.log('\nI · a refused ack must NOT leave the board thinking it asked');
  await soft('I', async () => {
    ackReply = { ok: false, error: 'Sheet busy — try again' };
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-15008-99999')]);
    await page.click('#htOk');
    await page.waitForTimeout(900);
    const rearmed = await page.evaluate(() => !holdSeen['24-15008-99999']);
    check('  the order re-arms for the next tick', rearmed, true);
    await push([held('24-15008-99999')]);
    check('  and the takeover fires again', (await read()).takeShown, true);
    ackReply = { ok: true, rows: 1, tag: '✓ SEEN 2:31 PM by Yassin · 1' };
  });

  // ── J · MORE THAN ONE ────────────────────────────────────────────────────
  console.log('\nJ · two held orders — one takeover, both in the strip');
  await soft('J', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdCloseTakeover(); });
    await push([held('24-11111-11111'), held('24-22222-22222')]);
    const r = await read();
    check('  both rows in the strip', r.rows, 2);
    check('  one takeover', r.takeShown, true);
    check('  it says how many more', /\+1 more/.test(r.takeMore), true);
  });

  // ── J2 · ESCALATED ───────────────────────────────────────────────────────
  console.log('\nJ2 · an escalated hold says so — red for 40 min must not look like red for 30s');
  await soft('J2', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); holdAcked = {}; holdCloseTakeover(); });
    await push([held('24-33333-33333', { escalated: true, escText: '2:37 PM',
      note: 'HOLD — buyer wants 2-Day · ⚠ ESCALATED 2:37 PM' })]);
    const r = await read();
    check('  the strip says escalated', /ESCALATED 2:37 PM/.test(r.stripText), true);
    check('  and still asks for an ack', /TAP TO ACKNOWLEDGE/.test(r.stripText), true);
    check('  still red', /unseen/.test(r.stripClass), true);
    check('  the takeover says who was pulled in',
          /escalated to the shipping desk at 2:37 PM/.test(r.takeMore), true);
    check('  the machine tag is stripped from the sentence',
          /ESCALATED/.test(r.takeTxt), false);
    check('  … but what was typed survives', /buyer wants 2-Day/.test(r.takeTxt), true);
  });

  // ── K · LIFTED ───────────────────────────────────────────────────────────
  // ⭐ The loop closed the other way. Lifting is still just deleting the word
  // from the note; what was missing was anyone being TOLD.
  console.log('\nK · the word HOLD removed — the floor is told, once');
  await soft('K', async () => {
    await page.evaluate(() => { window.__osc = 0; });
    await push([]);
    const r = await read();
    check('  the red strip is gone', r.stripShown, false);
    check('  takeover gone', r.takeShown, false);
    check('  and the floor is TOLD it is cleared', r.liftShown, true);
    check('  naming the order', /24-33333-33333/.test(r.liftText), true);
    check('  with a sound', r.osc > 0, true);
  });
  console.log('\nK2 · it STAYS until somebody taps it — the picker may not be looking');
  await soft('K2', async () => {
    // Several ticks go by. A box cleared to ship is still sitting there.
    await push([]); await push([]);
    check('  still up after three ticks', (await read()).liftShown, true);
    check('  backstop is an HOUR, not two minutes',
          await page.evaluate(() => LIFT_LIFE_MS), 3600000);
    await page.click('#liftStrip');
    await page.waitForTimeout(300);
    check('  a tap retires it', (await read()).liftShown, false);
    await push([]);
    check('  and a later empty tick does NOT re-announce', (await read()).liftShown, false);
  });
  console.log('\nK2b · a second lift ACCUMULATES — it must not erase the first');
  await soft('K2b', async () => {
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-L1')]);
    await push([]);
    check('  first lift up', /24-L1/.test((await read()).liftText), true);
    await push([held('24-L2')]);
    await push([]);
    const r = await read();
    check('  second lift added', /24-L2/.test(r.liftText), true);
    check('  ⚠ and the FIRST is still named', /24-L1/.test(r.liftText), true);
    await page.evaluate(() => holdClearLift());
  });
  console.log('\nK4 · ⚠ a row SWEEP is not a lift');
  await soft('K4', async () => {
    // A person clears one hold. n8n drops every shipped row at ~1 AM in one go.
    // The shapes differ, so the shape is the guard — not the clock, which would
    // have silenced this for the two thirds of the day the board is watched
    // from Riyadh.
    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-s1'), held('24-s2'), held('24-s3')]);
    await page.evaluate(() => holdClearLift());
    await push([]);
    check('  three vanishing at once says nothing', (await read()).liftShown, false);

    await page.evaluate(() => { holdSeen = {}; holdKnown = {}; holdClearLift(); });
    await push([held('24-s4'), held('24-s5')]);
    await page.evaluate(() => holdClearLift());
    await push([]);
    check('  two IS a lift, and is announced', (await read()).liftShown, true);
    await page.evaluate(() => holdClearLift());
  });

  console.log('\nK3 · ⚠ a board that just booted must not announce history as news');
  await soft('K3', async () => {
    const p2 = await ctx.newPage();
    await p2.route('http://hqlab.test/**', r => {
      const u = r.request().url();
      if (u.includes('/api/board')) {
        const body = JSON.parse(r.request().postData() || '{}');
        let res = { ok: false };
        if (body.action === 'boardTick')   res = Object.assign({ ok: true }, Object.assign({}, BASE, { held: [] }));
        if (body.action === 'boardRadio')  res = { ok: true, nowPlaying: '' };
        return r.fulfill({ contentType: 'application/json', body: JSON.stringify(res) });
      }
      return r.fulfill({ contentType: 'text/html; charset=utf-8', body: html });
    });
    await p2.route(/aladhan|open-meteo|youtube|ytimg|somafm|walmradio|isekoi/, r => r.abort());
    await p2.goto('http://hqlab.test/', { waitUntil: 'load' });
    await p2.waitForTimeout(2200);
    const shown = await p2.evaluate(() => {
      const e = document.getElementById('liftStrip');
      return !!(e && e.classList.contains('show'));
    });
    check('  a fresh board with no holds says nothing', shown, false);
    await p2.close();
  });

  console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
  await browser.close();
  process.exit(fail ? 1 : 0);
})();
