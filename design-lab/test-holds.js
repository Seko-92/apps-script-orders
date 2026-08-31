/**
 * test-holds.js — the SERVER half of the hold, loaded from the REAL Holds.js.
 *
 * Loads the actual file in a VM with the same Schema the sheet uses, so the
 * tests cannot drift from what ships. The two things worth proving here are the
 * ones the incident turned on:
 *
 *   1. a SHIPPED row still counts  — the board went blind exactly there
 *   2. the hold rule matches FloorBoard.html's, character for character —
 *      two copies of one rule is how A-9 sorted after A-50 in three files
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC   = process.env.SRC || path.join(__dirname, '..');
const HOLDS = fs.readFileSync(path.join(SRC, 'Holds.js'), 'utf8');
const BOARD = fs.readFileSync(path.join(SRC, 'FloorBoard.html'), 'utf8');

const cols = { SKU:1, QTY:2, LOCATION:3, SALES_ORDER:4, NOTE:5, STATUS:6,
               HAND:7, LEFT:8, SHIPPING:9, SHIP_COST:10 };
const sandbox = {
  console,
  Schema: {
    cols, idx: n => cols[n] - 1,
    status: { PENDING:'PENDING', PREPARING:'PREPARING', SHIPPED:'SHIPPED', CANCELED:'CANCELED' },
    boundaryMarker: 'DIRECT', dataStartRow: 4
  },
  Utilities: { formatDate: () => '2:31 PM' },
  PropertiesService: null, SpreadsheetApp: null, LockService: null,
  logActivity: () => {}, getCurrentPicker: () => 'Yassin · 1',
  _resolveStatusTargetRows: () => [], _dashBustTickCache: () => {},
  _tgSend: () => true, TELEGRAM_ADMIN_CHAT_ID: 'x'
};
vm.createContext(sandbox);
vm.runInContext(HOLDS, sandbox, { filename: 'Holds.js' });
const H = sandbox;

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
              (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};

// row builder — full width, indices matching Schema
const row = (sku, order, note, status, qty, loc) => {
  const r = new Array(10).fill('');
  r[cols.SKU - 1] = sku; r[cols.SALES_ORDER - 1] = order;
  r[cols.NOTE - 1] = note; r[cols.STATUS - 1] = status;
  r[cols.QTY - 1] = qty == null ? 1 : qty;
  r[cols.LOCATION - 1] = loc || '';
  return r;
};
const DIV = () => { const r = new Array(10).fill(''); r[0] = 'DIRECT'; return r; };

console.log('THE NOTE GRAMMAR');
t('plain word HOLD',            H.holdNoteHasHold('HOLD the label'), true);
t('lowercase',                  H.holdNoteHasHold('please hold this one'), true);
t('appended after a buyer note',H.holdNoteHasHold('Buyer Note: thanks · HOLD, changing service'), true);
t('household does NOT fire',    H.holdNoteHasHold('household goods'), false);
t('holder does NOT fire',       H.holdNoteHasHold('valve holder set'), false);
t('withhold does NOT fire',     H.holdNoteHasHold('do not withhold'), false);
t('empty is not a hold',        H.holdNoteHasHold(''), false);
t('null is not a hold',         H.holdNoteHasHold(null), false);

console.log('\nTHE ACK TAG');
t('unacked',        H.holdNoteHasAck('HOLD — change service'), false);
t('acked',          H.holdNoteHasAck('HOLD — change service · ✓ SEEN 2:31 PM by Yassin · 1'), true);
t('who and when',   H.holdNoteAckText('HOLD x · ✓ SEEN 2:31 PM by Yassin · 1'), '2:31 PM by Yassin · 1');
t('none when absent',H.holdNoteAckText('HOLD x'), '');
t('a real name wins over everything',
  H.holdBuildAckTag('Yassin · 1', 'sheet'), '✓ SEEN 2:31 PM by Yassin · 1');
/* ⚠ NO PICKER MUST NOT INVENT ONE. The first version wrote "by the floor"
   whatever the door was — and an acknowledgement made from the SIDEBAR, at a
   computer, then claimed the floor had seen it. A small lie in the one field
   whose whole job is to say who has this box. Record the DOOR, which is true. */
t('from the sheet, with no picker',  H.holdBuildAckTag('', 'sheet'),  '✓ SEEN 2:31 PM in the sheet');
t('from the board, with no picker',  H.holdBuildAckTag('', 'board'),  '✓ SEEN 2:31 PM at the board');
t('from Telegram, with no name',     H.holdBuildAckTag('', 'telegram'), '✓ SEEN 2:31 PM via Telegram');
t('door unknown — says so, invents nothing',
  H.holdBuildAckTag('', ''), '✓ SEEN 2:31 PM (no picker set)');
t('⚠ and never claims the floor when it was not the floor',
  /the floor/.test(H.holdBuildAckTag('', 'sheet')), false);
t('append keeps what was typed',
  H.holdAppendAck('HOLD — change service', '✓ SEEN 2:31 PM by Yassin · 1'),
  'HOLD — change service · ✓ SEEN 2:31 PM by Yassin · 1');
t('append is IDEMPOTENT — two taps cannot stack two tags',
  H.holdAppendAck('HOLD x · ✓ SEEN 2:31 PM by Yassin · 1', '✓ SEEN 9:99 PM by Someone'),
  'HOLD x · ✓ SEEN 2:31 PM by Yassin · 1');
t('append onto an empty note', H.holdAppendAck('', '✓ SEEN 2:31 PM by Yassin · 1'),
  '✓ SEEN 2:31 PM by Yassin · 1');

console.log('\n⭐ WHAT IS IN THE BOX — so the picker need not go to the sheet');
{
  const held = H.holdScanRows([
    row('165447', '24-aa', 'HOLD — change service', 'SHIPPED', 2, 'B-30'),
    row('172764', '24-aa', 'HOLD — change service', 'SHIPPED', 1, 'A-14')
  ]);
  t('every line is carried', held[0].items.length, 2);
  t('  … with sku, qty and shelf',
    held[0].items[0], { sku: '165447', qty: 2, loc: 'B-30' });
  t('  … and the line count agrees', held[0].lines, 2);
}
{
  // A fourteen-line kit must not fill a card nobody can read at a glance.
  const many = [];
  for (let i = 0; i < 14; i++) many.push(row('SKU' + i, '24-bb', 'HOLD x', 'SHIPPED', 1, 'C-' + i));
  const held = H.holdScanRows(many);
  t('a long order is capped', held[0].items.length, H.HOLDS.maxItemsShown);
  t('  … but lines still counts them ALL, so the board can say "+N more"',
    held[0].lines, 14);
}
{
  const held = H.holdScanRows([row('165447', '24-cc', 'HOLD x', 'SHIPPED', 1, 'NOT FOUND')]);
  t('a missing shelf is carried verbatim, not invented', held[0].items[0].loc, 'NOT FOUND');
}

console.log('\nTHE ESCALATION STAMP');
t('not escalated',   H.holdNoteHasEscalated('HOLD — change service'), false);
t('escalated',       H.holdNoteHasEscalated('HOLD x · ⚠ ESCALATED 2:37 PM'), true);
t('when',            H.holdNoteEscalatedText('HOLD x · ⚠ ESCALATED 2:37 PM'), '2:37 PM');
t('none when absent',H.holdNoteEscalatedText('HOLD x'), '');
t('survives an ack appended after it',
  H.holdNoteHasEscalated('HOLD x · ⚠ ESCALATED 2:37 PM · ✓ SEEN 2:42 PM by Yassin · 1'), true);
t('the ack still reads correctly past it',
  H.holdNoteAckText('HOLD x · ⚠ ESCALATED 2:37 PM · ✓ SEEN 2:42 PM by Yassin · 1'),
  '2:42 PM by Yassin · 1');

console.log('\nTHE SCAN — ⚠ the SHIPPED row is the whole point');
{
  const held = H.holdScanRows([
    row('165447', '24-14979-87359', 'HOLD — buyer wants 2-Day', 'SHIPPED'),
    row('172764', '24-14979-87359', 'HOLD — buyer wants 2-Day', 'SHIPPED'),
    row('158652', '24-15004-11290', '', 'PENDING')
  ]);
  t('the shipped order IS reported', held.length, 1);
  t('  … as one entry, not one per row', held[0].lines, 2);
  t('  … flagged shipped', held[0].shipped, true);
  t('  … and urgent', held[0].urgent, true);
  t('  … unacknowledged', held[0].acked, false);
}
{
  const held = H.holdScanRows([
    row('165447', '24-11111-00001', 'HOLD — verify serial', 'PENDING'),
    row('172764', '24-11111-00001', 'HOLD — verify serial', 'PENDING')
  ]);
  t('a PENDING-only hold is NOT urgent', held[0].urgent, false);
  t('  … and not shipped', held[0].shipped, false);
}
{
  const held = H.holdScanRows([
    row('165447', '24-22222-00002', 'HOLD', 'PENDING'),
    row('172764', '24-22222-00002', 'HOLD', 'PREPARING')
  ]);
  t('one PREPARING line makes it urgent', held[0].urgent, true);
  t('  … but not shipped', held[0].shipped, false);
}
{
  const held = H.holdScanRows([row('165447', '24-33333-00003', 'HOLD — void', 'CANCELED')]);
  t('CANCELED is out — nothing left to stop', held.length, 0);
}
{
  const held = H.holdScanRows([
    row('165447', '24-44444-00004', 'HOLD x · ✓ SEEN 2:31 PM by Yassin · 1', 'SHIPPED'),
    row('172764', '24-44444-00004', 'HOLD x', 'SHIPPED')
  ]);
  t('an ack on ANY row answers for the order', held[0].acked, true);
  t('  … and carries who saw it', held[0].ackText, '2:31 PM by Yassin · 1');
}
{
  const held = H.holdScanRows([
    row('165447', '24-88888-00008', 'HOLD x · ⚠ ESCALATED 2:37 PM', 'SHIPPED')
  ]);
  t('the scan carries the escalation through', held[0].escalated, true);
  t('  … with the time', held[0].escText, '2:37 PM');
  t('  … and it is still unanswered', held[0].acked, false);
}
{
  const held = H.holdScanRows([row('165447', '24-99999-00009', 'HOLD y', 'SHIPPED')]);
  t('a fresh hold is not escalated', held[0].escalated, false);
}
{
  const held = H.holdScanRows([
    row('165447', '24-55555-00005', 'HOLD a', 'SHIPPED'),
    DIV(),
    row('172764', 'SO-24609',       'HOLD b', 'PENDING')
  ]);
  t('two orders, two entries', held.length, 2);
  t('channel follows the divider', held.map(h => h.channel), ['EBAY', 'DIRECT']);
  t('UNACKNOWLEDGED SORTS FIRST', held[0].orderId, '24-55555-00005');
}
{
  const held = H.holdScanRows([
    row('165447', '24-66666-00006', 'HOLD a · ✓ SEEN 2:31 PM by Yassin · 1', 'SHIPPED'),
    row('172764', '24-77777-00007', 'HOLD b', 'SHIPPED')
  ]);
  t('the unanswered one leads the strip', held[0].orderId, '24-77777-00007');
}
{
  t('no rows, no holds', H.holdScanRows([]).length, 0);
  t('a row with no order id is skipped',
    H.holdScanRows([row('165447', '', 'HOLD', 'SHIPPED')]).length, 0);
  t('a normal board reports nothing',
    H.holdScanRows([row('165447', '24-1', 'Buyer Note: thanks', 'PENDING')]).length, 0);
}

console.log('\n⚠ THE SELECTION PATH — getActive(), never openById()');
{
  // ⚠ openById returns a handle with NO connection to the user's UI, so
  // getActiveRangeList() on it does not return what is selected. That produced a
  // confident "No held rows in the selection" while a held row sat selected on
  // screen. Documented trap in this project (the v1 openPrepQueue bug) and I
  // walked into it anyway. Assert on the SHIPPED source so it cannot come back.
  const src = fs.readFileSync(path.join(SRC, 'Holds.js'), 'utf8');
  const fn = src.slice(src.indexOf('function acknowledgeSelectedHold'));
  const raw = fn.slice(0, fn.indexOf('\nfunction ') === -1 ? fn.length : fn.indexOf('\nfunction '));
  /* ⚠ STRIP THE COMMENTS FIRST. This assertion failed on its first run against
     correct code, because the comment above the fix NAMES the trap it prevents
     ("NEVER openById … getActiveRangeList() on it does not return …") and the
     regex happily matched the prose. Any assertion made against SOURCE TEXT has
     to read the code only — documentation describes the bug, which is exactly
     what the pattern is hunting for. */
  const body = raw.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  t('it uses SpreadsheetApp.getActive()', /SpreadsheetApp\.getActive\(\)/.test(body), true);
  t('it does NOT read the selection off openById',
    /openById[\s\S]{0,200}getActiveRange/.test(body), false);
  t('and it asks the SHEET for the range list',
    /sheet\.getActiveRangeList\(\)/.test(body), true);
}

console.log('\n⭐ ONE UNANSWERED HOLD NEEDS NO SELECTION');
{
  const sheet = {
    getLastRow: () => 5,
    getRange: () => ({ getValues: () => [
      row('165447', '24-solo', 'HOLD — change service', 'SHIPPED', 1, 'B-30'),
      row('172764', '24-other', 'Buyer Note: thanks', 'PENDING', 1, 'A-14')
    ]})
  };
  t('the sole unanswered hold is found', H._holdSoleUnacked(sheet), '24-solo');
}
{
  const sheet = {
    getLastRow: () => 5,
    getRange: () => ({ getValues: () => [
      row('165447', '24-a', 'HOLD a', 'SHIPPED', 1, 'B-30'),
      row('172764', '24-b', 'HOLD b', 'SHIPPED', 1, 'A-14')
    ]})
  };
  t('⚠ TWO is ambiguous — refuse to guess', H._holdSoleUnacked(sheet), '');
}
{
  const sheet = {
    getLastRow: () => 5,
    getRange: () => ({ getValues: () => [
      row('165447', '24-a', 'HOLD a · ✓ SEEN 2:31 PM by Yassin', 'SHIPPED', 1, 'B-30')
    ]})
  };
  t('an already-answered hold is not offered', H._holdSoleUnacked(sheet), '');
}
{
  const sheet = { getLastRow: () => 5, getRange: () => ({ getValues: () => [] }) };
  t('no holds, nothing to pick', H._holdSoleUnacked(sheet), '');
}

console.log('\n⚠ EVERY OUTCOME ANSWERS FOR ITSELF');
{
  /* A button that is always tappable WILL be tapped in every state, so every
     state needs its own sentence. The old code reported "Acknowledged 1 hold ·
     0 rows stamped" when the selection was already answered — technically true
     and useless. Assert the shipped source composes all three. */
  const src = fs.readFileSync(path.join(SRC, 'Holds.js'), 'utf8');
  const fn  = src.slice(src.indexOf('function acknowledgeSelectedHold'));
  t('nothing held at all says so',      /Nothing to acknowledge/.test(fn), true);
  t('  … and is NOT an error',          /ok: true, orders: 0, rows: 0, quiet: true/.test(fn), true);
  t('already answered says WHEN',       /was already acknowledged/.test(fn), true);
  t('several unanswered asks which',    /select a row of the one you mean/.test(fn), true);
  t('a mixed batch reports the remainder', /already answered/.test(fn), true);
}
{
  // boardAckHold has to hand the prior tag back, or "already acknowledged"
  // invites the next question immediately.
  const src = fs.readFileSync(path.join(SRC, 'Holds.js'), 'utf8');
  const fn  = src.slice(src.indexOf('function boardAckHold'), src.indexOf('function _holdSoleUnacked'));
  t('the prior acknowledgement is returned', /priorAck/.test(fn), true);
}

console.log('\n⚠ THE TWO COPIES OF THE RULE MUST AGREE');
{
  // FloorBoard.html cannot call server code, so it carries its own noteHasHold.
  // Extract the REAL one out of the shipped file and run both over the same
  // strings — if they ever drift, a hold shows on one surface and not the other.
  const m = BOARD.match(/function noteHasHold\(noteText\)\s*\{([\s\S]*?)\n\s{4}\}/);
  t('the board still HAS its own rule', !!m, true);
  if (m) {
    const boardRule = new Function('noteText', m[1]);
    const cases = ['HOLD', 'hold this', 'Buyer Note: x · HOLD y', 'household',
                   'holder', 'withhold', '', 'HOLDING', 'on hold.', 'HOLD—now'];
    const drift = cases.filter(c => boardRule(c) !== H.holdNoteHasHold(c));
    t('server and board agree on every case', drift, []);
  }
}

console.log('\n' + (fail ? '❌ ' : '✅ ') + pass + ' passed · ' + fail + ' failed');
process.exit(fail ? 1 : 0);
