/**
 * test-hold-escalation.js — the backstop's own logic, against the REAL Holds.js.
 *
 * ⚠ THE CASE THAT MATTERS MOST IS THE ONE THAT MUST **NOT** FIRE. A still-PENDING
 * hold produces no takeover and no siren, so nobody is ever prompted to
 * acknowledge it — which means escalating on it would fire on EVERY calm hold,
 * forever, on the one channel that has to stay believed.
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};

// A fresh sandbox per scenario so Script Property state cannot leak between them.
function makeEnv(live, opts) {
  opts = opts || {};
  const props = { HOLD_LIVE: JSON.stringify(live) };
  if (opts.state) props.HOLD_ESCALATION = JSON.stringify(opts.state);
  const sent = [];
  const stamped = [];
  const sandbox = {
    console: { log(){}, error(){} },
    Schema: { cols:{NOTE:5,SALES_ORDER:4,SKU:1,STATUS:6,QTY:2,LOCATION:3,HAND:7,LEFT:8,SHIPPING:9,SHIP_COST:10},
              idx: n => ({SKU:1,QTY:2,LOCATION:3,SALES_ORDER:4,NOTE:5,STATUS:6,HAND:7,LEFT:8,SHIPPING:9,SHIP_COST:10})[n]-1,
              status:{PENDING:'PENDING',PREPARING:'PREPARING',SHIPPED:'SHIPPED',CANCELED:'CANCELED'},
              boundaryMarker:'DIRECT', dataStartRow:4 },
    Utilities: { formatDate: () => '2:37 PM' },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; },
      deleteProperty: k => { delete props[k]; }
    })},
    SpreadsheetApp: null, LockService: null,
    logActivity(){}, getCurrentPicker: () => '', _resolveStatusTargetRows: () => [],
    _dashBustTickCache(){},
    _tgSend: (chat, text) => { if (opts.sendFails) return false; sent.push(text); return true; },
    TELEGRAM_ADMIN_CHAT_ID: 'admin',
    Date
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Holds.js'), 'utf8'), sandbox, { filename: 'Holds.js' });
  // holdStampEscalated needs a sheet; the stamp is best-effort by contract, so
  // record the attempt and return without one.
  sandbox.holdStampEscalated = oid => { stamped.push(oid); return 1; };
  return { run: () => sandbox.checkHoldEscalation(), sent, stamped,
           state: () => JSON.parse(props.HOLD_ESCALATION || '{}'),
           HOLDS: sandbox.HOLDS };
}

const H = (o, over) => Object.assign({ o, a: false, s: true, u: true, n: 'HOLD — change service' }, over || {});
const AGO = m => Date.now() - m * 60000;

console.log('THE TWO WINDOWS');
{
  const e = makeEnv([]);
  t('PREPARING waits 30 min', e.HOLDS.escalateAfterMin, 30);
  t('SHIPPED waits only 8',   e.HOLDS.escalateShippedAfterMin, 8);
  t('and shipped IS the shorter one',
    e.HOLDS.escalateShippedAfterMin < e.HOLDS.escalateAfterMin, true);
}

console.log('\n⭐ THE WINDOW FOLLOWS THE ORDER, NOT THE PASS');
{
  // 10 minutes old. PREPARING → still inside 30, must stay quiet.
  const e = makeEnv([H('24-a', { s: false })], { state: { '24-a': { first: AGO(10), alerted: 0 } } });
  e.run();
  t('10 min PREPARING stays quiet (limit 30)', e.sent.length, 0);
}
{
  // The SAME age, but the label got bought in the meantime → past the 8-min
  // limit, so it must fire on the very next pass rather than wait out the old
  // clock. This is the case a threshold stamped at first sight would miss.
  const e = makeEnv([H('24-b', { s: true })], { state: { '24-b': { first: AGO(10), alerted: 0 } } });
  e.run();
  t('the same 10 min, now SHIPPED, escalates at once', e.sent.length, 1);
  t('  … and the message names the limit it used', /limit 8/.test(e.sent[0]), true);
}
{
  const e = makeEnv([H('24-c', { s: true })], { state: { '24-c': { first: AGO(7), alerted: 0 } } });
  e.run();
  t('7 min SHIPPED is still inside its window', e.sent.length, 0);
}
{
  const e = makeEnv([H('24-d', { s: false })], { state: { '24-d': { first: AGO(31), alerted: 0 } } });
  e.run();
  t('31 min PREPARING finally escalates', e.sent.length, 1);
  t('  … naming its own limit', /limit 30/.test(e.sent[0]), true);
}

console.log('\n⚠ WHAT MUST NOT FIRE');
{
  // A PENDING hold gets no takeover and no siren, so nobody is ever asked to
  // acknowledge it. Escalating would fire on every calm hold, forever.
  const e = makeEnv([H('24-1', { u: false })], { state: { '24-1': { first: AGO(90), alerted: 0 } } });
  e.run();
  t('a still-PENDING hold NEVER escalates, however old', e.sent.length, 0);
  t('  … and is not even tracked', Object.keys(e.state()).indexOf('24-1'), -1);
}
{
  const e = makeEnv([H('24-2', { a: true })], { state: { '24-2': { first: AGO(90), alerted: 0 } } });
  e.run();
  t('an acknowledged hold never escalates', e.sent.length, 0);
}
{
  const e = makeEnv([H('24-3')]);
  e.run();
  t('a hold seen for the FIRST time does not fire immediately', e.sent.length, 0);
  t('  … but its clock starts now', Math.round((Date.now() - e.state()['24-3'].first) / 1000) < 3, true);
}
{
  const e = makeEnv([H('24-4', { s: false })], { state: { '24-4': { first: AGO(14), alerted: 0 } } });
  e.run();
  t('14 min on a PREPARING hold is still inside its window', e.sent.length, 0);
}
{
  /* ⚠ THE ONE THAT PROVES THE 15 → 30 CHANGE BIT. Twenty minutes fired under
     the old window and must now stay silent — a picker mid-walk on a big order
     is doing the job, not ignoring the board. Run this against the 15-minute
     build and it fails, which is the whole point of keeping it. */
  const e = makeEnv([H('24-4b', { s: false })], { state: { '24-4b': { first: AGO(20), alerted: 0 } } });
  e.run();
  t('20 min PREPARING stays quiet (fired under the old 15)', e.sent.length, 0);
}
{
  /* …and the same twenty minutes on a SHIPPED hold still fires, so the two
     windows genuinely diverged rather than both moving. */
  const e = makeEnv([H('24-4c', { s: true })], { state: { '24-4c': { first: AGO(20), alerted: 0 } } });
  e.run();
  t('  … while 20 min SHIPPED still escalates', e.sent.length, 1);
}

console.log('\nWHAT MUST FIRE');
{
  const e = makeEnv([H('24-5')], { state: { '24-5': { first: AGO(16), alerted: 0 } } });
  e.run();
  t('16 minutes on a SHIPPED hold escalates (well past its 8)', e.sent.length, 1);
  t('  … names the order', /24-5/.test(e.sent[0]), true);
  t('  … says the label is bought', /label already bought/.test(e.sent[0]), true);
  t('  … carries what was typed', /change service/.test(e.sent[0]), true);
  t('  … and stamps the sheet', e.stamped, ['24-5']);
  t('  … recording the key so it cannot repeat', e.state()['24-5'].alerted > 0, true);
}
{
  const e = makeEnv([H('24-6', { s: false })], { state: { '24-6': { first: AGO(35), alerted: 0 } } });
  e.run();
  t('a PREPARING hold escalates too', e.sent.length, 1);
  t('  … and says the label is NOT bought yet', /not bought yet/.test(e.sent[0]), true);
}

console.log('\n⭐⭐ THE SECOND GATE — PREPARING → SHIPPED is a NEW crossing, not a repeat');
{
  // Told at the prep gate, then somebody buys the label anyway.
  const e = makeEnv([H('24-g1', { s: false })],
                    { state: { '24-g1': { first: AGO(20), alerted: AGO(4), gate: 'prep' } } });
  e.run();
  t('nothing new while it is still PREPARING', e.sent.length, 0);

  const e2 = makeEnv([H('24-g1', { s: true })],
                     { state: { '24-g1': { first: AGO(20), alerted: AGO(4), gate: 'prep' } } });
  e2.run();
  t('the moment it SHIPS, a second message fires', e2.sent.length, 1);
  t('  … and it reads as an escalation, not a repeat',
    /JUST SHIPPED/.test(e2.sent[0]), true);
  t('  … saying they were already told', /nobody answered/.test(e2.sent[0]), true);
  t('  … and that this is the last free moment',
    /voiding it is still free/.test(e2.sent[0]), true);
  t('  … the gate advances', e2.state()['24-g1'].gate, 'ship');
}
{
  // ⚠ SHIP IS TERMINAL. There is nothing worse to report, so it must never nag.
  const e = makeEnv([H('24-g2', { s: true })],
                    { state: { '24-g2': { first: AGO(60), alerted: AGO(30), gate: 'ship' } } });
  e.run(); e.run(); e.run();
  t('an order already told at the SHIP gate never speaks again', e.sent.length, 0);
}
{
  // A hold written on an already-shipped order gets ONE message, at the ship gate.
  const e = makeEnv([H('24-g3', { s: true })], { state: { '24-g3': { first: AGO(10), alerted: 0, gate: '' } } });
  e.run();
  t('a born-shipped hold fires once', e.sent.length, 1);
  t('  … at the ship gate', e.state()['24-g3'].gate, 'ship');
  const before = e.sent.length; e.run(); e.run();
  t('  … and then stays quiet', e.sent.length, before);
}

console.log('\nALERT ONCE, NEVER NAG');
{
  const e = makeEnv([H('24-7')], { state: { '24-7': { first: AGO(60), alerted: AGO(40), gate: 'ship' } } });
  e.run();
  t('an already-alerted hold stays quiet', e.sent.length, 0);
}
{
  const e = makeEnv([H('24-8')], { state: { '24-8': { first: AGO(20), alerted: 0 } } });
  e.run(); const first = e.sent.length;
  e.run(); e.run();
  t('three passes, ONE message', [first, e.sent.length], [1, 1]);
}

console.log('\n⚠ A FAILED SEND MUST NOT BURN THE KEY');
{
  const e = makeEnv([H('24-9')], { state: { '24-9': { first: AGO(20), alerted: 0 } }, sendFails: true });
  e.run();
  t('nothing recorded when Telegram refused', e.state()['24-9'].alerted, 0);
  t('  … and the sheet is not stamped either', e.stamped.length, 0);
}

console.log('\nTHE STORE CLEANS ITSELF');
{
  const e = makeEnv([], { state: { 'gone': { first: AGO(20), alerted: AGO(5) } } });
  e.run();
  t('a lifted hold drops out at once', Object.keys(e.state()).length, 0);
}
{
  const e = makeEnv([H('24-10', { a: true })], { state: { '24-10': { first: AGO(20), alerted: AGO(5) } } });
  e.run();
  t('so does an acknowledged one — the next hold re-arms clean',
    Object.keys(e.state()).length, 0);
}
{
  const e = makeEnv([]);
  t('no snapshot yet → says so rather than throwing',
    /no live-hold snapshot|watching 0/.test(makeEnv([]).run()) || typeof e.run() === 'string', true);
}

console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
process.exit(fail ? 1 : 0);
