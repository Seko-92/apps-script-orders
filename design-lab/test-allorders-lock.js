/**
 * test-allorders-lock.js — the All Orders lock, loaded from the REAL BrandTheme.js
 * so the tests cannot drift from what ships.
 *
 * What is worth proving here — this is the file where a wrong range is a locked-out
 * shift and a missing exception is a silently broken nightly sweep:
 *   1. ⚠⚠ IT REFUSES without the n8n account. That refusal IS the safety mechanism:
 *      n8n's `E5. Delete SHIPPED Row` writes to All Orders directly as a NAMED
 *      account, so locking without an exception stops the ~1 AM sweep and the
 *      symptom shows days later.
 *   2. The carve-outs are EXACTLY the five the floor needs — no more (a hole) and
 *      no fewer (a locked-out shift). Built from Schema, so a column move cannot
 *      silently open the wrong one.
 *   3. The merge is resolved, not hardcoded — Pick ID for Shipping is F2:G2 and
 *      that second column has already moved once (2026-05-19).
 *   4. A bad editor email ROLLS BACK rather than leaving a half-locked sheet.
 *
 * ⚠ EVERY SECTION FAILS SOFT — the choosePicker lesson.
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) drop the `if (!acct)` refusal            → A fails
 *   b) drop prot.removeEditors(...)             → E fails
 *   c) hardcode 'G2' instead of resolving merges → D fails
 *   d) drop the rollback in the addEditor catch  → H fails
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const CODE = fs.readFileSync(path.join(SRC, 'BrandTheme.js'), 'utf8');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const has = (label, hay, needle) => {
  const ok = String(hay).indexOf(needle) !== -1;
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label + (ok ? '' : '  → ' + JSON.stringify(String(hay).slice(0, 160))));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

/* A recording mock of just the surface the lock touches. Faithful where it matters:
   getEditors() returns a populated list so removeEditors has something to strip,
   and merged ranges are modelled because the merge is the subtle part. */
function build(opts) {
  opts = opts || {};
  const log = { protects: [], removed: [], editorsRemoved: 0, added: [], unprotected: null, domainEdit: null, written: [] };
  const a1 = (r, c, nr, nc) => {
    const col = String.fromCharCode(64 + c);
    return nr > 1 ? col + r + ':' + col : col + r;
  };
  const mkRange = (name, merges) => ({
    getA1Notation: () => name,
    getMergedRanges: () => merges || []
  });
  const mkProt = (type, desc, open) => {
    const p = {
      _type: type, _desc: desc, _open: (open || []),
      getDescription: () => p._desc,
      setDescription: d => { p._desc = d; return p; },
      getEditors: () => [{ getEmail: () => 'someone@else.com' }],
      removeEditors: list => { log.editorsRemoved += list.length; return p; },
      addEditor: e => {
        if (opts.badEditor) throw new Error('Invalid email: ' + e);
        log.added.push(e); return p;
      },
      canDomainEdit: () => true,
      setDomainEdit: v => { log.domainEdit = v; return p; },
      // ⚠ The protection must REMEMBER what was set on it. refreshAllOrdersLockCarveOuts
      //   reads the state back after writing — the whole bug class is the gap between
      //   what you asked for and what Sheets stored, so a stub that forgets proves nothing.
      setUnprotectedRanges: rs => {
        log.unprotected = rs.map(r => r.getA1Notation());
        log.writes = (log.writes || 0) + 1;
        p._open = log.unprotected.slice();
        return p;
      },
      getUnprotectedRanges: () => p._open.map(a => ({ getA1Notation: () => a })),
      remove: () => { log.removed.push(p._desc); }
    };
    return p;
  };
  const existing = (opts.existing || []).map((d, i) => mkProt('SHEET', d, i === 0 ? opts.existingOpen : null));
  const existingRange = (opts.existingRange || []).map(d => mkProt('RANGE', d));

  const MAXROWS = opts.maxRows || 1000;
  const sheet = {
    getMaxRows: () => MAXROWS,
    getProtections: type => (type === 'SHEET' ? existing : existingRange),
    protect: () => { const p = mkProt('SHEET', ''); log.protects.push(p); return p; },
    getRange: function () {
      if (arguments.length === 1) {
        let n = String(arguments[0]);
        // ⚠⚠ THE STUB USED TO BE HEALTHIER THAN THE REAL THING, AND THAT IS WHY THIS
        //    SUITE WENT GREEN THROUGH TWO SEPARATE STAFF LOCKOUTS. It returned the
        //    notation it was handed, so `getRange("E4:E")` came back as the string
        //    "E4:E" and the assertions cheerfully proved an UNBOUNDED carve-out that
        //    has never existed. Apps Script has no unbounded Range: open-ended notation
        //    is materialised against the grid the instant you ask, so the protection
        //    always stores a fixed box. Model that, or the harness cannot see the bug.
        const openEnded = n.match(/^([A-Z]+)(\d+):([A-Z]+)$/);
        if (openEnded) n = openEnded[1] + openEnded[2] + ':' + openEnded[3] + MAXROWS;
        // Pick ID for Shipping is a MERGE in the live layout.
        return mkRange(n, n === 'F2' ? [mkRange('F2:G2')] : []);
      }
      const [r, c, nr, nc] = arguments;
      return mkRange(a1(r, c, nr, nc));
    }
  };

  const sandbox = {
    console: { log: () => {} },
    SPREADSHEET_ID: 'x',
    MAIN_SHEET_NAME: 'All orders',
    Schema: {
      cols: { SKU: 1, QTY: 2, LOCATION: 3, SALES_ORDER: 4, NOTE: 5, STATUS: 6, HAND: 7, LEFT: 8 },
      dataStartRow: 4, headerRow: 3, dataWidth: 10,
      cellEmployeeId: 'F2', cellAdjustmentId: 'H2',
      cellEmployeeIdNext: 'I2', cellAdjustmentIdNext: 'J2',
      // ⚠⚠ WITHOUT THIS THE SUITE WAS 38/5 AND HAD BEEN FOR A WHILE. Five whole sections
      //    — including D, the one that checks the carve-outs are exactly what the floor
      //    needs — threw `Schema.pickIdA1 is not a function` and failed SOFT, so the file
      //    still printed a tidy tally while its most important assertions never ran. A
      //    permanently-red check is a check nobody reads; a red check that is red for a
      //    HARNESS reason is worse, because it hides the real ones behind it.
      // ⚠ This mirrors pickIdA1's BRANCH, not its plumbing — the resolver has its own
      //   suite (test-pickid-resolver.js). What this stub has to get right is that the
      //   lock ASKS for an address instead of hardcoding one, and gets a different answer
      //   when the mode flips. `pickIdMode: 'new'` in opts exercises the other arm.
      pickIdA1: function (which) {
        var isNew = opts.pickIdMode === 'new';
        return (which === 'adjustment') ? (isNew ? 'J2' : 'H2')
                                        : (isNew ? 'I2' : 'F2');
      }
    },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: n => (n === 'All orders' ? sheet : null) }),
      ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
      flush: () => {}
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => (opts.acct === undefined ? null : opts.acct),
        setProperty: (k, v) => { log.written.push({ k: k, v: v }); }
      })
    }
  };
  // ⚠ The production files call _obRequireOwner behind `typeof … === "function"`, so a
  //   sandbox WITHOUT it silently skips every owner gate. Injecting it is the only way
  //   these tests can see the gate at all — otherwise they pass vacuously.
  if (opts.notOwner) {
    sandbox._obRequireOwner = function (what) { return '\ud83d\udd12 ' + what + ' is owner-only.'; };
  } else if (opts.ownerGate) {
    sandbox._obRequireOwner = function () { return null; };
  }
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: 'BrandTheme.js' });
  return { B: sandbox, log };
}

// ===============================================================================
section('A · ⚠⚠ IT REFUSES WITHOUT THE n8n ACCOUNT', () => {
  const { B, log } = build({ acct: null });
  const r = B.protectAllOrdersSheet();
  has('A1 it refuses', r, '❌ REFUSED');
  has('A2 it names the property to set', r, 'N8N_SHEETS_ACCOUNT');
  has('A3 it explains the nightly sweep', r, 'E5. Delete SHIPPED Row');
  has('A4 ...and that the symptom is delayed', r, 'days later');
  has('A5 it offers the explicit opt-out', r, "'none'");
  t('A6 ⭐ NOTHING WAS LOCKED', log.protects.length, 0);

  const empty = build({ acct: '   ' });
  has('A7 whitespace is not an account either', empty.B.protectAllOrdersSheet(), '❌ REFUSED');
  t('A8 and it locked nothing', empty.log.protects.length, 0);
});

// ===============================================================================
section("B · 'none' is an explicit, deliberate opt-out", () => {
  const { B, log } = build({ acct: 'none' });
  const r = B.protectAllOrdersSheet();
  has('B1 it locks', r, '✅ All Orders LOCKED');
  t('B2 the sheet was protected', log.protects.length, 1);
  t('B3 no editor was granted', log.added, []);
  has('B4 the report says so', r, 'n8n exception: none');
  const { B: B2 } = build({ acct: 'NONE' });
  has('B5 case-insensitive', B2.protectAllOrdersSheet(), 'n8n exception: none');
});

// ===============================================================================
section('C · a real account gets its exception', () => {
  const { B, log } = build({ acct: 'n8n-bot@example.com' });
  const r = B.protectAllOrdersSheet();
  t('C1 the account was added as an editor', log.added, ['n8n-bot@example.com']);
  has('C2 the report names it', r, 'n8n exception: n8n-bot@example.com');
});

// ===============================================================================
section('D · ⭐ the carve-outs are EXACTLY the five the floor needs', () => {
  const { B, log } = build({ acct: 'none' });
  B.protectAllOrdersSheet();
  // NOTE(E) · STATUS(F) · LEFT(H) from the HEADER ROW to the last row, then both Pick IDs.
  t('D1 exactly five open ranges', log.unprotected.length, 5);
  // ⚠⚠ ANCHORED AT ROW 3, NOT ROW 4, AND THAT ONE ROW IS THE WHOLE 2026-09-09 FIX.
  //    doPost inserts arrivals with insertRowsBefore(dataStartRow) — immediately BEFORE
  //    row 4. Sheets expands a range only when rows land strictly INSIDE it; at the top
  //    boundary it SHIFTS the range down instead, so a row-4 anchor was pushed down by
  //    every batch of new orders and left the newest rows locked for staff.
  t('D2 and they are the right five, anchored ABOVE the insert point', log.unprotected,
    ['E3:E1000', 'F3:F1000', 'H3:H1000', 'F2:G2', 'H2']);

  const open = log.unprotected.join(' ');
  ['A', 'B', 'D', 'G'].forEach(function (c) {
    t('D3 col ' + c + ' (identity/derived) is NOT open', open.indexOf(c + '4:') !== -1, false);
  });
});

// ===============================================================================
section('E · the editor list is narrowed to the owner', () => {
  const { B, log } = build({ acct: 'none' });
  B.protectAllOrdersSheet();
  t('E1 pre-existing editors were stripped', log.editorsRemoved > 0, true);
  t('E2 domain edit was turned off', log.domainEdit, false);
});

// ===============================================================================
section('F · idempotent — a re-run refreshes rather than stacking', () => {
  const { B, log } = build({
    acct: 'none',
    existing: ['HQ-LOCK: All Orders — identity columns locked (2026-08-29)']
  });
  const r = B.protectAllOrdersSheet();
  t('F1 the prior lock was removed', log.removed.length, 1);
  has('F2 and it says it refreshed', r, 'refreshed');
});

// ===============================================================================
section('G · unprotect removes ONLY the lock', () => {
  const { B, log } = build({
    acct: 'none',
    existing: ['HQ-LOCK: All Orders — identity columns locked', 'SOMEONE-ELSE: do not touch'],
    existingRange: ['HQ-STRUCTURE: Banner rows 1-3 — accidental-edit guard']
  });
  const r = B.unprotectAllOrdersSheet();
  t('G1 exactly one protection removed', log.removed.length, 1);
  has('G2 and it was the lock', log.removed[0], 'HQ-LOCK');
  has('G3 it reports the count', r, 'Removed 1');
});

// ===============================================================================
section('H · a bad editor email ROLLS BACK — never a half-locked sheet', () => {
  const { B, log } = build({ acct: 'not-an-email', badEditor: true });
  const r = B.protectAllOrdersSheet();
  has('H1 it reports the failure', r, 'Could not add');
  has('H2 ...and says nothing was locked', r, 'Nothing was locked');
  t('H3 ⭐ the protection it had just created was removed', log.removed.length, 1);
  t('H4 no carve-outs were ever applied', log.unprotected, null);
});

// ===============================================================================
// ⚠⚠ THE "✅" HAS TO MEAN SOMETHING.
//
// 2026-08-30: setN8nSheetsAccountNow shipped in BOTH BrandTheme.js and OwnerBridge.js.
// Apps Script concatenates root files into ONE global scope in an unspecified order, so
// which body ran was undefined — and the two DIFFERED. The BrandTheme copy passed the
// literal "PUT-THE-ACCOUNT-EMAIL-OR-none-HERE" straight through, and the setter accepted
// it, because the only check was `if (!v)`. The user would have seen
//     ✅ N8N_SHEETS_ACCOUNT = PUT-THE-ACCOUNT-EMAIL-OR-none-HERE
// and believed the account was configured.
//
// The lock's own addEditor rollback (section H) catches the malformed value later, so this
// was never going to break the sweep silently — but it IS a green checkmark on a state that
// is wrong, which this codebase already rules is a bug. Reject it where the message can
// still name the fix.
section('I · the n8n account setter refuses anything that is not usable', () => {
  const { B, log } = build({ acct: null });
  const KEY = 'N8N_SHEETS_ACCOUNT';

  has('I1 an empty value is refused', B.setN8nSheetsAccount(''), '\u274c');

  const ph = B.setN8nSheetsAccount('PUT-THE-ACCOUNT-EMAIL-OR-none-HERE');
  has('I2 \u26a0\u26a0 the unedited placeholder is refused (the 2026-08-30 shape)', ph, '\u274c');
  has('I3 ...and the refusal explains the nightly-sweep cost', ph, 'shipped-row sweep');

  has('I4 a bare word is refused', B.setN8nSheetsAccount('yes'), '\u274c');
  has('I5 a half-typed address is refused', B.setN8nSheetsAccount('n8n@'), '\u274c');

  t('I6 \u2b50 NOTHING was written by any of those refusals', log.written.length, 0);

  const okMail = B.setN8nSheetsAccount('n8n-sheets@example.com');
  has('I7 a real address is accepted', okMail, '\u2705');
  t('I8 ...and it is what got stored',
    log.written[log.written.length - 1], { k: KEY, v: 'n8n-sheets@example.com' });

  has('I9 the literal sentinel is accepted', B.setN8nSheetsAccount('none'), '\u2705');
  has('I10 ...case-insensitively', B.setN8nSheetsAccount('NONE'), '\u2705');
  t('I11 surrounding whitespace is trimmed, not rejected',
    B.setN8nSheetsAccount('  none  ').indexOf('\u2705'), 0);

  t('I12 the sentinel is read from the constant, never re-typed',
    B.ALL_ORDERS_LOCK.noneSentinel, 'none');
});

// ===============================================================================
// ⭐ THE WHOLE LOCK SEQUENCE FROM THE SIDEBAR.
//
// Until 2026-08-30 the n8n account was the ONE step that lived in the Apps Script editor:
// you opened setN8nSheetsAccountNow(), edited `var VALUE` in code, and ran it.
// installAllOrdersLock() refuses without it, so that single value dragged the entire
// sequence out of the sidebar. It is a field now — which means the SETTER is reachable by
// anyone with the panel open, and therefore has to be gated like every other lock control.
//
// ⚠ The value becomes an EDITOR EXCEPTION on a protected sheet. Staff setting it would be
//   staff granting the edit rights the lock exists to withhold. Owner-only, and
//   deliberately NOT in OWNER_BRIDGE.actionNames.
section('J · the n8n account is settable from the sidebar, but only by the owner', () => {
  const staff = build({ acct: null, notOwner: true });
  const refused = staff.B.setN8nSheetsAccount('attacker@example.com');
  has('J1 ⚠⚠ a staff member cannot set the account', refused, 'owner-only');
  t('J2 ⭐ and NOTHING was written', staff.log.written.length, 0);

  const st = staff.B.getN8nSheetsAccountState();
  t('J3 the state reader tells the panel it is owner-only', st.ok, false);
  t('J4 ...and leaks no value', st.value, '');

  const fresh = build({ acct: null, ownerGate: true });
  const s0 = fresh.B.getN8nSheetsAccountState();
  t('J5 owner + unset → isSet false', [s0.ok, s0.isSet], [true, false]);
  has('J6 the panel is told which property it is', s0.key, 'N8N_SHEETS_ACCOUNT');

  const mail = build({ acct: 'n8n-sheets@example.com', ownerGate: true });
  const s1 = mail.B.getN8nSheetsAccountState();
  t('J7 owner + email → isSet true, isNone false', [s1.isSet, s1.isNone], [true, false]);
  t('J8 ...and the value is handed back to prefill the field', s1.value, 'n8n-sheets@example.com');
  has('J9 the owner CAN set it', mail.B.setN8nSheetsAccount('other@example.com'), '✅');

  const none = build({ acct: 'none', ownerGate: true });
  t('J10 ⭐ the sentinel is reported as such, not as an address',
    none.B.getN8nSheetsAccountState().isNone, true);

  t('J11 ⚠ state is a SHAPE, never prose the client must parse', typeof s1, 'object');
});

// The description the installer writes; sections K-O seed a lock that already exists.
const LOCKED = 'HQ-LOCK: All Orders \u2014 identity columns locked (2026-08-29)';

// ===============================================================================
// ⚠⚠⚠ THE 2026-09-09 INCIDENT: "I lock the sheet and after a while the picker tells me
//     the whole sheet is locked — even notes, even status." Hours sometimes, days
//     others. Unlock-then-relock fixed it every time, which is the tell: the LOCK was
//     fine, the CARVE-OUT had moved.
//
// THE MECHANISM. Apps Script has no unbounded Range — `getRange("E4:E")` is materialised
// against the grid on the spot, so the protection stored "E4:E<maxRows>". From then on
// Sheets adjusts it like any other range, and doPost inserts every arrival with
// `insertRowsBefore(Schema.dataStartRow, n)` — immediately BEFORE the carve-out's first
// row. A range EXPANDS only when rows land strictly INSIDE it; at the top boundary it
// SHIFTS DOWN. So each batch of orders pushed the carve-out down by N and left those N
// rows — the newest orders, the exact rows being picked — locked for staff and perfectly
// fine for the owner, because removeEditors() ignores the owner.
//
// THE FIX IS ONE ROW: anchor at headerRow, so the insert lands strictly inside and Sheets
// expands instead of shifting. Sections K–O are what makes that permanent.
section('K · ⭐ THE ANCHOR SITS ABOVE THE INSERT POINT — the fix, as a property', () => {
  const { B, log } = build({ acct: 'none' });
  B.protectAllOrdersSheet();

  const S = B.Schema;
  log.unprotected.forEach(function (a1) {
    const m = a1.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m || m[1] !== m[3]) return;                 // the Pick ID cells, not columns
    t('K1 ' + a1 + ' starts strictly ABOVE the row doPost inserts at',
      Number(m[2]) < S.dataStartRow, true);
    t('K2 ' + a1 + ' reaches the last row of the sheet', Number(m[4]), 1000);
  });

  t('K3 ⚠ and the anchor IS the header row — one row higher is what makes an ' +
    'insertRowsBefore(dataStartRow) land strictly inside',
    S.dataStartRow - S.headerRow, 1);

  // The bottom is materialised, so it MUST track the grid rather than a remembered number.
  const big = build({ acct: 'none', maxRows: 4200 });
  big.B.protectAllOrdersSheet();
  t('K4 on a bigger sheet the carve-out reaches ITS last row, not a hardcoded one',
    big.log.unprotected[0], 'E3:E4200');
});

// ===============================================================================
section('L · ⚠⚠ THE DETECTOR NAMES A TOP-SHIFTED CARVE-OUT — the reported incident', () => {
  // What the sheet looked like after ~5 rows of arrivals against a row-4 anchor.
  const { B } = build({
    acct: 'none',
    existing: [LOCKED],
    existingOpen: ['E9:E1000', 'F9:F1000', 'H9:H1000', 'F2:G2', 'H2']
  });
  const sheet = B.SpreadsheetApp.openById('x').getSheetByName('All orders');
  const st = B._lockCarveOutState(sheet);

  t('L1 it is seen as locked', st.locked, true);
  t('L2 ⭐ and as DRIFTED', st.drifted, true);
  const all = st.reasons.join(' | ');
  has('L3 it names the LOCKED ROWS, not just "stale"', all, 'LOCKED for staff on rows 4\u20138');
  has('L4 it names the mechanism so the next reader does not re-derive it', all,
      'insert at the first data row pushed the carve-out down');
  t('L5 all three floor columns are reported, not just the first',
    ['E', 'F', 'H'].every(c => all.indexOf('column ' + c + ' ') !== -1), true);

  // ⚠ THE HOLE IN THE OLD DETECTOR, PINNED: it only ever compared the END row, so a
  //   carve-out that had been pushed DOWN read as perfectly healthy. That is why a
  //   green `describeAllOrdersLock` sat next to a locked-out shift.
  t('L6 ⚠⚠ the END row is intact here — an end-only check would have called this healthy',
    st.reasons.some(r => r.indexOf('ends at row') !== -1), false);
});

// ===============================================================================
section('M · the 2026-08-31 shape too — a carve-out short of the last row', () => {
  const { B } = build({
    acct: 'none',
    existing: [LOCKED],
    existingOpen: ['E4:E51', 'F4:F51', 'H4:H51', 'F2:G2', 'H2']
  });
  const sheet = B.SpreadsheetApp.openById('x').getSheetByName('All orders');
  const st = B._lockCarveOutState(sheet);
  t('M1 drifted', st.drifted, true);
  has('M2 it says exactly where the floor stops being able to type', st.reasons.join(' | '),
      'LOCKED for staff on rows 52\u20131000');
  has('M2b ...and why the bottom ran out', st.reasons.join(' | '),
      'the sheet grew past where the carve-out ends');

  // A carve-out that is simply absent is the worst case and must not read as "fine".
  const gone = build({ acct: 'none', existing: [LOCKED], existingOpen: ['F2:G2', 'H2'] });
  const gs = gone.B._lockCarveOutState(
    gone.B.SpreadsheetApp.openById('x').getSheetByName('All orders'));
  t('M3 a MISSING column carve-out is drift, loudly', gs.drifted, true);
  has('M4 ...and it says so in words', gs.reasons.join(' | '), 'has NO carve-out at all');

  // A Pick ID carved as a lone cell against a live merge locks the whole merge.
  const merge = build({
    acct: 'none', existing: [LOCKED],
    existingOpen: ['E3:E1000', 'F3:F1000', 'H3:H1000', 'F2', 'H2']
  });
  const ms = merge.B._lockCarveOutState(
    merge.B.SpreadsheetApp.openById('x').getSheetByName('All orders'));
  t('M5 a partial carve-out over the F2:G2 merge is caught', ms.drifted, true);
  has('M6 ...and named as the merge problem it is', ms.reasons.join(' | '),
      'does not cover the whole merge');
});

// ===============================================================================
section('N · the self-heal repairs it, and reports what it repaired', () => {
  const { B, log } = build({
    acct: 'none',
    existing: [LOCKED],
    existingOpen: ['E9:E1000', 'F9:F1000', 'H9:H1000', 'F2:G2', 'H2']
  });
  const r = B.refreshAllOrdersLockCarveOuts();

  t('N1 it changed something', r.changed, true);
  t('N2 ⭐ and what it wrote is the canonical set', r.after,
    ['E3:E1000', 'F3:F1000', 'H3:H1000', 'F2:G2', 'H2']);
  has('N3 the message carries the BEFORE, so the log is the diagnosis', r.message, 'E9:E1000');
  has('N4 ...and the reason, not just the fact', r.message, 'LOCKED for staff on rows 4\u20138');
  t('N5 ⚠ it never creates a lock', log.protects.length, 0);
  t('N6 ⚠⚠ and never touches the editor list — it is a repair, not a lock control',
    [log.editorsRemoved, log.added], [0, []]);

  // ⚠ A re-anchor that re-fires forever would rewrite the protection every hour for
  //   nothing. Nothing wrong ⇒ nothing written.
  const clean = build({
    acct: 'none', existing: [LOCKED],
    existingOpen: ['E3:E1000', 'F3:F1000', 'H3:H1000', 'F2:G2', 'H2']
  });
  const c = clean.B.refreshAllOrdersLockCarveOuts();
  t('N7 ⭐ a healthy lock is a NO-OP', [c.changed, clean.log.writes || 0], [false, 0]);
  has('N8 ...and says so', c.message, 'already correct');

  // ⚠ An ABSENT lock is not a broken one. The self-heal must never install one.
  const unlocked = build({ acct: 'none' });
  const u = unlocked.B.refreshAllOrdersLockCarveOuts();
  t('N9 ⚠⚠ an unlocked sheet is left alone', [u.locked, u.changed, unlocked.log.protects.length],
    [false, false, 0]);
});

// ===============================================================================
section('O · the reporter and the detector share ONE rule', () => {
  const drift = {
    acct: 'none', existing: [LOCKED],
    existingOpen: ['E9:E1000', 'F9:F1000', 'H9:H1000', 'F2:G2', 'H2']
  };
  const { B } = build(drift);
  const sheet = B.SpreadsheetApp.openById('x').getSheetByName('All orders');

  // ⚠ setupMasthead's one-line warning and the full report must never disagree — a green
  //   status line beside a locked-out floor is how this survived two incidents.
  t('O1 _lockNeedsRefresh agrees with _lockCarveOutState', B._lockNeedsRefresh(sheet), true);

  const rep = B.describeAllOrdersLock();
  has('O2 the report says DRIFTED', rep, 'DRIFTED');
  has('O3 ...prints what it SHOULD be', rep, 'E3:E1000');
  has('O4 ...and names the one-tap fix', rep, 'Re-anchor carve-outs');

  const clean = build({
    acct: 'none', existing: [LOCKED],
    existingOpen: ['E3:E1000', 'F3:F1000', 'H3:H1000', 'F2:G2', 'H2']
  });
  const cs = clean.B.SpreadsheetApp.openById('x').getSheetByName('All orders');
  t('O5 and a healthy lock reads healthy in both', clean.B._lockNeedsRefresh(cs), false);
  has('O6 ...in words', clean.B.describeAllOrdersLock(), 'carve-outs: ✅ correct');
});

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-allorders-lock: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
