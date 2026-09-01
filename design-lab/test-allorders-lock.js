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
  const mkProt = (type, desc) => {
    const p = {
      _type: type, _desc: desc,
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
      setUnprotectedRanges: rs => { log.unprotected = rs.map(r => r.getA1Notation()); return p; },
      getUnprotectedRanges: () => [],
      remove: () => { log.removed.push(p._desc); }
    };
    return p;
  };
  const existing = (opts.existing || []).map(d => mkProt('SHEET', d));
  const existingRange = (opts.existingRange || []).map(d => mkProt('RANGE', d));

  const sheet = {
    getMaxRows: () => 1000,
    getProtections: type => (type === 'SHEET' ? existing : existingRange),
    protect: () => { const p = mkProt('SHEET', ''); log.protects.push(p); return p; },
    getRange: function () {
      if (arguments.length === 1) {
        const n = arguments[0];
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
  // NOTE(E) · STATUS(F) · LEFT(H) from row 4 down, then both Pick ID cells.
  t('D1 exactly five open ranges', log.unprotected.length, 5);
  t('D2 and they are the right five', log.unprotected,
    ['E4:E', 'F4:F', 'H4:H', 'F2:G2', 'H2']);

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

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-allorders-lock: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
