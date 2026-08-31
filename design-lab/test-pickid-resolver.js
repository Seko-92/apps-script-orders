/**
 * test-pickid-resolver.js — the grace-period address resolver.
 *
 * THE PROPERTY THAT MAKES THIS SHIPPABLE: with PICK_ID_ADDR unset, every reader must
 * resolve to the cells that hold the data TODAY. Shipping the resolver is then a strict
 * no-op against the live sheet, which is what lets the New Version be cut BEFORE the
 * cells move — and that ordering is the whole zero-downtime design.
 *
 * ⚠⚠ WHY A PROPERTY AND NOT A CONSTANT EDIT. getDashboardTick caches into
 *    CacheService.getScriptCache(), which is shared across the WHOLE script project. In
 *    any window where the pinned /exec reads F2 while the HEAD publish trigger reads I2,
 *    both write the SAME key — so the picker chip flickers on a 45-second coin flip.
 *    A property is read identically by both versions, so a flip moves everything at once.
 *
 * Run:  node test-pickid-resolver.js
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC  = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond, got) => {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (got !== undefined ? '  → got ' + JSON.stringify(got) : '')); }
};
const section = t => console.log('\n' + t);
const block = fn => { try { fn(); } catch (e) { fail++; console.log('  ✗ ⚠ SECTION ABORTED: ' + e.message); } };

// load Schema.js with a controllable PropertiesService
let PROP = null, READS = 0, THROW = false;
const sandbox = {
  console, JSON, String, Number, Boolean, Array, Object, RegExp, Math, Date,
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => { READS++; if (THROW) throw new Error('denied'); return PROP; }
    })
  }
};
vm.createContext(sandbox);
vm.runInContext(read('Schema.js'), sandbox, { filename: 'Schema.js' });
const S = sandbox.Schema;
const reset = () => { S._pickIdModeCache = null; READS = 0; THROW = false; };

section('A · ⭐ unset ⇒ TODAY\'S CELLS — this is what makes shipping it a no-op');
block(() => {
  reset(); PROP = null;
  ok('A1 employee resolves to the live Shipping cell',   S.pickIdA1() === 'F2', S.pickIdA1());
  reset();
  ok('A2 adjustment resolves to the live Adjustment cell', S.pickIdA1('adjustment') === 'H2', S.pickIdA1('adjustment'));
  reset();
  ok('A3 and they equal the constants the 8 call sites used before',
     S.pickIdA1() === S.cellEmployeeId && S.pickIdA1('adjustment') === S.cellAdjustmentId);
});

section('B · flipped ⇒ the hidden columns');
block(() => {
  reset(); PROP = 'new';
  ok('B1 employee → I2',   S.pickIdA1() === 'I2', S.pickIdA1());
  reset(); PROP = 'new';
  ok('B2 adjustment → J2', S.pickIdA1('adjustment') === 'J2', S.pickIdA1('adjustment'));
  reset(); PROP = 'NEW  ';
  ok('B3 case- and whitespace-tolerant', S.pickIdA1() === 'I2', S.pickIdA1());
});

section('C · ⚠ anything but "new" fails SAFE to the cells holding the data');
block(() => {
  [['old', 'the explicit old value'], ['', 'empty string'], ['yes', 'a plausible typo'],
   ['NEWER', 'a prefix match that must NOT count'], [undefined, 'undefined']].forEach(([v, why]) => {
    reset(); PROP = v;
    ok('C· ' + why + ' → F2', S.pickIdA1() === 'F2', S.pickIdA1());
  });
  reset(); THROW = true;
  ok('C· a throwing PropertiesService → F2 (never guesses "probably migrated")',
     S.pickIdA1() === 'F2', S.pickIdA1());
});

section('D · ⚠ memoised per execution — one read, and it cannot change mid-run');
block(() => {
  reset(); PROP = 'new';
  S.pickIdA1(); S.pickIdA1(); S.pickIdA1('adjustment'); S.pickIdA1();
  ok('D1 four calls, ONE property read', READS === 1, READS);

  // ⚠⚠ THE REGRESSION THAT MATTERS. A function that reads one address then writes
  //    another is the row-shift bug class wearing different clothes.
  PROP = 'old';
  ok('D2 ⚠⚠ flipping mid-execution does NOT move the address',
     S.pickIdA1() === 'I2', S.pickIdA1());
  reset();
  ok('D3 ...and the next execution picks the change up', S.pickIdA1() === 'F2', S.pickIdA1());
});

section('E · ⚠ no live reader still reads the raw constants');
block(() => {
  // ⚠ STRIP COMMENTS FIRST. Documentation describes the very pattern being hunted, and
  //   a source-text assertion that forgets this accuses its own explanation
  //   (CLAUDE.md, the openById round).
  const strip = t => t.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
  const FILES = ['ActivityLog.js', 'FulfillmentService.js', 'DashboardService.js', 'BrandTheme.js'];
  const offenders = [];
  FILES.forEach(f => {
    strip(read(f)).split('\n').forEach((ln, i) => {
      if (!/Schema\.(cellEmployeeId|cellAdjustmentId)\b/.test(ln)) return;
      if (/Next\b/.test(ln)) return;                       // the *Next constants are fine
      if (/cellEmployeeIdNext|cellAdjustmentIdNext/.test(ln)) return;
      offenders.push(f + ':' + (i + 1) + '  ' + ln.trim().slice(0, 72));
    });
  });

  // setupMasthead's sweep keep-list is the ONE legitimate reader: it must protect BOTH
  // pairs at once, precisely so a half-finished migration cannot lose a dropdown.
  const allowed = offenders.filter(o => /Schema\.cellEmployeeId,\s+Schema\.cellAdjustmentId/.test(o));
  const rogue   = offenders.filter(o => !allowed.includes(o));

  ok('E1 the scan can see the constants at all (not a vacuous pass)', offenders.length > 0, offenders.length);
  ok('E2 ⚠ exactly one legitimate reader — the sweep keep-list', allowed.length === 1, allowed);
  ok('E3 ⚠⚠ no other live code reads the raw constants', rogue.length === 0, rogue);
});

console.log('\n' + (fail ? '✗ ' : '✅ ') + 'test-pickid-resolver: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
