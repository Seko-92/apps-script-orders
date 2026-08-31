// Test auditBoardStockAdjustments() against the REAL StockAdjust.js, with the
// Apps Script globals stubbed. Loading the real file (rather than a re-typed
// copy) is the point — a test against a paraphrase proves nothing.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'StockAdjust.js'), 'utf8');

const ACTIVITY_LOG = {
  sheetName: 'Activity Log',
  cols: { TIMESTAMP:1, EVENT:2, ORDER_ID:3, SKU:4, QTY:5, SOURCE:6, DETAIL:7, NOTE:8, PICKER:9 },
  idx: function (n) { return ACTIVITY_LOG.cols[n] - 1; },
  dataWidth: 9, headerRow: 1, dataStartRow: 2
};

// [ts, event, orderId, sku, qty, source, detail, note, picker]
const row = (ts, sku, source, detail, picker, order) =>
  [ts, 'NOTE', order || '(stock)', sku, '', source, detail, '', picker || ''];

const D = s => new Date(s);

const ROWS = [
  // ── BEFORE the 2026-08-12 fix ──
  // classic bug signature: correct shelf, pushed counted+qty → inflated by 1
  row(D('2026-08-11T14:05:00'), '165447', 'board', 'Zoho stock 5 → 6 (+1) · adj 99001', ''),
  // a downward correction — the bug inflated, so this one is not its signature
  row(D('2026-08-11T15:20:00'), '155394', 'board', 'Zoho stock 9 → 4 (-5) · adj 99002', ''),
  // a no-op: logged on purpose, can never have caused harm
  row(D('2026-08-11T16:00:00'), '168138', 'board', 'Stock confirmed at 3 — no adjustment needed', ''),

  // ── AFTER the fix ──
  row(D('2026-08-13T09:30:00'), '194244', 'board', 'Zoho stock 8 → 9 (+1) · adj 99003', 'Hatem21332'),
  row(D('2026-08-13T11:00:00'), '176688', 'board', 'Zoho stock 12 → 10 (-2) · adj 99004', 'YAwiss 1'),

  // ── noise that must be ignored ──
  row(D('2026-08-13T12:00:00'), '190455', 'sidebar', 'Zoho stock 4 → 5 (+1) · adj 99005', 'x'),
  row(D('2026-08-13T12:05:00'), '190455', 'board',  'from PENDING', 'Hatem21332'),
  row(D('2026-08-13T12:10:00'), '',       'n8n',    'inserted', '')
];

function run(rows, days) {
  const logs = [];
  const sandbox = {
    SPREADSHEET_ID: 'x',
    ACTIVITY_LOG,
    // ⚠ Hand the sandbox OUR Date. A vm context gets its own constructor, so
    // `when instanceof Date` inside the script would be false for every Date
    // built out here — the rows would all be skipped and the audit would
    // report an empty log. Test-harness trap, not a bug in the function.
    Date,
    console: { log: m => logs.push(String(m)) },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const p = n => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
      }
    },
    SpreadsheetApp: {
      openById: () => ({
        getSpreadsheetTimeZone: () => 'America/Chicago',
        getSheetByName: name => name !== 'Activity Log' ? null : ({
          getLastRow: () => rows.length + 1,
          getRange: () => ({ getValues: () => rows })
        })
      })
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(SRC, sandbox);
  const text = sandbox.auditBoardStockAdjustments(days);
  return { text, logged: logs.join('\n') };
}

let fails = [];
const check = (name, cond, detail) => {
  console.log(`  ${cond ? '✓' : '✗'} ${name}${cond ? '' : '   ' + (detail || '')}`);
  if (!cond) fails.push(name);
};

console.log('\n── a log containing the bug window ──');
const a = run(ROWS);
check('output is also console.logged (Run button shows nothing otherwise)',
      a.logged.length > 0 && a.logged === a.text);
check('counts the 4 real moves', /4 adjustment\(s\) moved stock/.test(a.text));
check('names every board event kind it saw', /status change \(✓ Pick\)/.test(a.text));
check('counts the 1 no-op separately', /1 was a confirmation/.test(a.text));
check('ignores non-board sources', !/99005/.test(a.text));
check('ignores board rows that are not stock events', !/from PENDING/.test(a.text));
check('flags the pre-fix window', /2 adjustment\(s\) predate the fix/.test(a.text));
check('marks the positive delta as the bug signature', /⚠ INFLATED\?.*165447/.test(a.text));
check('does NOT mark the negative delta', /probably fine.*155394/.test(a.text));
check('surfaces the adjustment id for cross-reference', /adj 99001/.test(a.text));
check('says who, when known', /by Hatem21332/.test(a.text));
check('admits when the picker is unknown', /by \?\?\? \(see Zoho reason\)/.test(a.text));
check('warns that failures are not logged', /Failed pushes are NOT logged/.test(a.text));

console.log('\n── a clean log (nothing predates the fix) ──');
const clean = run(ROWS.filter(r => r[0] >= D('2026-08-12T00:00:00')));
check('reports CLEAR', /✅ CLEAR/.test(clean.text));
check('still lists the post-fix adjustments', /194244/.test(clean.text));

console.log('\n── the board has never adjusted anything ──');
const none = run(ROWS.filter(r => r[5] !== 'board'));
check('says so plainly', /never moved a stock number/.test(none.text));

// ⚠ THE REGRESSION THAT MATTERS. Live data on 2026-08-14 was 21 adjustments
// ALL logged as "Zoho stock undefined → 6 (undefined)" — the n8n proxy responds
// with an ARRAY (respondWith:allIncomingItems) and Apps Script read it as an
// object, so before/delta/adjustment_id were never recorded. The first audit
// dropped unreadable rows before the cutoff check and therefore announced
// "✅ CLEAR" — the most reassuring answer possible, produced by seeing nothing.
// An audit must never mistake "I could not read it" for "there is nothing".
console.log('\n── every adjustment unreadable, some inside the window ──');
const BROKEN = [
  row(D('2026-08-10T18:56:00'), '218204', 'board', 'Zoho stock undefined → 6 (undefined)', ''),
  row(D('2026-08-10T19:04:00'), '218204', 'board', 'Zoho stock undefined → 2 (undefined)', ''),
  row(D('2026-08-13T09:27:00'), '158715', 'board', 'Zoho stock undefined → 15 (undefined)', 'Hatem21332')
];
const broken = run(BROKEN);
check('does NOT claim the board never moved stock', !/never moved a stock number/.test(broken.text));
check('says the writes landed but the reporting was lost',
      /only the reporting of it was lost/.test(broken.text));
check('does NOT falsely report CLEAR', !/✅ CLEAR/.test(broken.text), 'reassurance from blindness');
check('still flags the 2 inside the window', /2 adjustment\(s\) predate the fix/.test(broken.text));
check('marks unreadable rows as UNREADABLE, not "probably fine"',
      /\? UNREADABLE/.test(broken.text) && !/probably fine/.test(broken.text));
check('points at Zoho for the missing figures', /check Zoho/.test(broken.text));
check('shows the raw text so the shape is visible', /\|Zoho stock undefined → 6 \(undefined\)\|/.test(broken.text));

console.log('\n── windowed: last 2 days only ──');
const win = run(ROWS, 2);
check('drops rows outside the window', !/99001/.test(win.text), 'pre-window row leaked in');
check('labels the window', /last 2 days/.test(win.text));

console.log('\n' + '='.repeat(58));
if (fails.length) {
  console.log('✗ ' + fails.join('\n✗ '));
  console.log('\n--- what it actually produced ---\n' + a.text);
  process.exit(1);
}
console.log('✓ audit reads the real log shape, flags the bug signature, and is');
console.log('  honest about what it cannot see.');
console.log('\n--- sample output ---\n' + a.text);
