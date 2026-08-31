/**
 * test-pickid-migrate.js — the migration must never destroy an option list.
 *
 * THE STAKES: the dropdown option lists exist NOWHERE in this codebase. They were made
 * by hand in the Sheets UI, so the cell is the only copy of its own options and a botched
 * half-migration cannot be recovered from source. Every assertion here is about the same
 * promise — validate everything, then write; never write then discover.
 *
 * Run:  node test-pickid-migrate.js
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC  = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let pass = 0, fail = 0;
const ok = (n, c, g) => { if (c) { pass++; console.log('  ✓ ' + n); }
  else { fail++; console.log('  ✗ ' + n + (g !== undefined ? '  → ' + JSON.stringify(g) : '')); } };
const section = t => console.log('\n' + t);
const block = fn => { try { fn(); } catch (e) { fail++; console.log('  ✗ ⚠ SECTION ABORTED: ' + e.message); } };

const SHIP = ['Pick ID for Shipping','Shipping - AShamma 12343','Shipping - Hatem 21332',
              'Shipping - YAwiss 1','Shipping - Turkmani 43122'];
// ⚠ the REAL Adjustment list, trailing space and all — copied off the live sheet
const ADJ  = ['Pick ID for Adjustment ','Adjustments - AShamma 12343','Adjustments - AHafiz 21332',
              'Adjustments - YAwiss 1','Adjustments - AAlbri 43122'];

// ── a sheet that records every write ────────────────────────────────────────────────
function makeSheet(cfg) {
  const cells = {};
  const mk = a1 => cells[a1] || (cells[a1] = {
    a1, dv: null, val: '', merges: [], writes: [],
    getValue(){ return this.val; },
    getDataValidation(){ return this.dv; },
    getMergedRanges(){ return this.merges; },
    setDataValidation(d){ this.writes.push('dv'); this.dv = d; return this; },
    setValue(v){
      // ⚠⚠ SHEETS REJECTS THIS, SO THE STUB MUST TOO. A stub cheaper than the real
      //    thing tests nothing (the audio FakeCtx lesson) — and the entire failure
      //    being guarded against is setValue THROWING against a strict rule.
      if (this.dv && this.dv.getAllowInvalid() === false) {
        const cv = this.dv.getCriteriaValues() || [];
        if (Array.isArray(cv[0]) && cv[0].indexOf(v) === -1) {
          throw new Error('The data you entered in cell ' + this.a1 +
                          ' violates the data validation rules set on this cell.');
        }
      }
      this.writes.push('val'); this.val = v; return this;
    },
    clearContent(){ this.writes.push('clear'); this.val = ''; return this; },
    setBackground(){return this;}, setFontColor(){return this;}, setFontFamily(){return this;},
    setFontWeight(){return this;}, setFontSize(){return this;},
    setHorizontalAlignment(){return this;}, setVerticalAlignment(){return this;}, setWrap(){return this;}
  });
  Object.keys(cfg).forEach(a1 => {
    const c = mk(a1), s = cfg[a1];
    c.val = s.value || '';
    c.merges = s.merged ? [{ getA1Notation: () => s.merged }] : [];
    if (s.opts || s.range) {
      c.dv = {
        getCriteriaType: () => s.range ? 'VALUE_IN_RANGE' : 'VALUE_IN_LIST',
        getAllowInvalid: () => s.allowInvalid !== false,
        getHelpText: () => s.helpText || '',
        getCriteriaValues: () => [ s.range ? { getA1Notation: () => 'Z1:Z9' } : s.opts,
                                   s.showDropdown !== false ]
      };
    }
  });
  return { _cells: cells, getRange: a1 => mk(a1) };
}

function run(cfg, mode, propStart) {
  const props = { PICK_ID_ADDR: propStart || null };
  const order = [];
  const sandbox = {
    console: { log(){} }, JSON, String, Number, Boolean, Array, Object, RegExp, Math, Date,
    SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders',
    BRAND: { ink:'#1a1a1a', yellow:'#ffd400', fontDisplay:'Oswald' },
    _obRequireOwner: () => null,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => props[k],
      setProperty: (k,v) => { props[k]=v; order.push('PROP_SET'); },
      deleteProperty: k => { delete props[k]; order.push('PROP_DEL'); }
    })}
  };
  const sheet = makeSheet(cfg);
  sandbox.SpreadsheetApp = {
    openById: () => ({ getSheetByName: () => sheet }),
    flush: () => {},
    newDataValidation: () => {
      const st = { opts:null, show:null, allow:null, help:null };
      const api = {
        requireValueInList: (o,s) => { st.opts=o; st.show=s; return api; },
        setAllowInvalid: v => { st.allow=v; return api; },
        setHelpText: t => { st.help=t; return api; },
        build: () => ({
          _st: st,
          getCriteriaType: () => 'VALUE_IN_LIST',
          getAllowInvalid: () => st.allow,
          getHelpText: () => st.help || '',
          getCriteriaValues: () => [st.opts, st.show]
        })
      };
      return api;
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(read('Schema.js'), sandbox, { filename:'Schema.js' });
  vm.runInContext(read('ActivityLog.js'), sandbox, { filename:'ActivityLog.js' });
  vm.runInContext(read('BrandTheme.js'), sandbox, { filename:'BrandTheme.js' });
  sandbox.Schema._pickIdModeCache = null;
  const rep = sandbox.migratePickIdCells(mode);
  return { rep, sheet, props, order, S: sandbox };
}

const CLEAN = {
  F2: { opts: SHIP, allowInvalid: false, value: 'Pick ID for Shipping', merged: 'F2:G2' },
  H2: { opts: ADJ,  allowInvalid: false, value: 'Pick ID for Adjustment' },
  I2: {}, J2: {}
};
const clone = o => JSON.parse(JSON.stringify(o));

section('A · ⭐ the dry run writes NOTHING');
block(() => {
  const r = run(clone(CLEAN), undefined);
  const writes = Object.values(r.sheet._cells).reduce((n,c)=>n+c.writes.length,0);
  ok('A1 zero writes across every cell', writes === 0, writes);
  ok('A2 the property is untouched', r.props.PICK_ID_ADDR == null, r.props.PICK_ID_ADDR);
  ok('A3 and it says the preflight passed', /PREFLIGHT PASSES/.test(r.rep));
});

section('B · ⚠⚠ it refuses rather than half-migrating');
block(() => {
  const cases = [
    ['a VALUE_IN_RANGE rule (getCriteriaValues()[0] is a Range)',
      (c)=>{ c.F2 = { range:true, allowInvalid:false, value:'x' }; }, /VALUE_IN_RANGE|not VALUE_IN_LIST/],
    ['the destination already has a validation',
      (c)=>{ c.I2 = { opts:SHIP }; }, /ALREADY has a validation/],
    ['the destination is not empty',
      (c)=>{ c.I2 = { value:'leftover' }; }, /not empty/],
    ['the destination is inside a merge',
      (c)=>{ c.I2 = { merged:'I2:J2' }; }, /inside a merge/],
    ['no option could ever satisfy the gate',
      (c)=>{ c.F2 = { opts:['aaa','bbb'], allowInvalid:false, value:'aaa' }; }, /gate regex/],
    ['the source has no validation at all',
      (c)=>{ c.F2 = { value:'x' }; }, /NO validation/]
  ];
  cases.forEach(([why, mutate, re]) => {
    const c = clone(CLEAN); mutate(c);
    const r = run(c, 'APPLY');
    const writes = Object.values(r.sheet._cells).reduce((n,x)=>n+x.writes.length,0);
    ok('B· refuses: ' + why, re.test(r.rep) && /REFUSED/.test(r.rep), r.rep.slice(-160));
    ok('B·   ...with ZERO writes', writes === 0, writes);
  });
});

section('C · ⚠⚠ the 4am reset guard — the trap that was already live on H2');
block(() => {
  // placeholder absent from the list AND allowInvalid false ⇒ setValue would THROW at 4am
  const c = clone(CLEAN);
  c.H2 = { opts: ['Adjustments - A 1','Adjustments - B 2'], allowInvalid:false, value:'Adjustments - A 1' };
  const r = run(c, 'APPLY');
  ok('C1 refuses when the derived placeholder is not writable',
     /REFUSED/.test(r.rep) && /4am/.test(r.rep), r.rep.slice(-200));

  // same list but allowInvalid true ⇒ the write lands, so it is allowed through
  const c2 = clone(CLEAN);
  c2.H2 = { opts: ['Adjustments - A 1'], allowInvalid:true, value:'Adjustments - A 1' };
  const r2 = run(c2, undefined);
  ok('C2 ...but allows it when allowInvalid is true', /PREFLIGHT PASSES/.test(r2.rep));
});

section('D · ⭐ APPLY preserves the rule instead of re-authoring it');
block(() => {
  const c = clone(CLEAN);
  c.F2.helpText = 'pick your id'; c.F2.showDropdown = true;
  const r = run(c, 'APPLY');
  const dv = r.sheet._cells.I2.dv;
  ok('D1 migrated', /MIGRATED/.test(r.rep), r.rep.slice(-200));
  ok('D2 ⚠⚠ allowInvalid PRESERVED as false — never hardcoded',
     dv.getAllowInvalid() === false, dv.getAllowInvalid());
  ok('D3 helpText preserved', dv.getHelpText() === 'pick your id', dv.getHelpText());
  ok('D4 showDropdown preserved', dv.getCriteriaValues()[1] === true);
  ok('D5 ⚠ options copied VERBATIM, same order, nothing trimmed',
     JSON.stringify(dv.getCriteriaValues()[0]) === JSON.stringify(SHIP));
  ok('D6 ⚠ the Adjustment trailing space survives untouched',
     r.sheet._cells.J2.dv.getCriteriaValues()[0][0] === 'Pick ID for Adjustment ');
  ok('D7 the selected value moved too', r.sheet._cells.I2.val === 'Pick ID for Shipping');
  ok('D8 the sources are cleared', !r.sheet._cells.F2.dv && r.sheet._cells.F2.val === '');
});

section('E · ⚠⚠ the property is set LAST');
block(() => {
  const r = run(clone(CLEAN), 'APPLY');
  ok('E1 PICK_ID_ADDR = new', r.props.PICK_ID_ADDR === 'new', r.props.PICK_ID_ADDR);
  // the source clears happen inside the loop; the property write is the final act
  const lastClear = r.sheet._cells.H2.writes.lastIndexOf('clear');
  ok('E2 both sources were cleared before it', lastClear >= 0 && r.order[r.order.length-1] === 'PROP_SET',
     r.order);
  ok('E3 ⚠ and it is the ONLY property write', r.order.filter(x=>x==='PROP_SET').length === 1, r.order);
});

section('F · ⚠ already migrated ⇒ refuse, do not run twice');
block(() => {
  const c = clone(CLEAN);
  const r = run(c, 'APPLY', 'new');
  ok('F1 refuses when PICK_ID_ADDR is already "new"',
     /REFUSED/.test(r.rep) && /already/.test(r.rep), r.rep.slice(-200));
  const writes = Object.values(r.sheet._cells).reduce((n,x)=>n+x.writes.length,0);
  ok('F2 with zero writes', writes === 0, writes);
});

section('G · ⚠⚠ the LIVE case — a value that is not in its own option list');
block(() => {
  // exactly what the sheet holds: H2 = "Pick ID for Adjustment", list has it WITH a space
  const c = clone(CLEAN);
  ok('G0 the fixture is the real inconsistency, not an invention',
     c.H2.value === 'Pick ID for Adjustment' && ADJ.indexOf('Pick ID for Adjustment') === -1);

  const dry = run(clone(CLEAN), undefined);
  ok('G1 the DRY RUN warns before anyone runs APPLY',
     /is NOT in this list/.test(dry.rep), dry.rep.slice(-300));

  const r = run(clone(CLEAN), 'APPLY');
  ok('G2 ⚠⚠ APPLY does not throw — it completes', /MIGRATED/.test(r.rep), r.rep.slice(-260));
  ok('G3 J2 gets the list\'s OWN placeholder, not the invalid value',
     r.sheet._cells.J2.val === 'Pick ID for Adjustment ', r.sheet._cells.J2.val);
  ok('G4 ...and SHIPPING still migrated normally',
     r.sheet._cells.I2.val === 'Pick ID for Shipping', r.sheet._cells.I2.val);
  ok('G5 both sources cleared', !r.sheet._cells.F2.dv && !r.sheet._cells.H2.dv);
  ok('G6 the property still landed last', r.props.PICK_ID_ADDR === 'new');
});

section('H · ⚠ a REAL picker missing from its list is left BLANK, never guessed');
block(() => {
  const c = clone(CLEAN);
  c.F2.value = 'Shipping - Ghost 999';        // a real name, absent from the options
  const dry = run(clone(c), undefined);
  ok('H1 the dry run flags it as a real picker, not a placeholder',
     /REAL picker/.test(dry.rep), dry.rep.slice(-300));

  const r = run(clone(c), 'APPLY');
  ok('H2 it completes rather than throwing', /MIGRATED/.test(r.rep), r.rep.slice(-200));
  ok('H3 ⚠ I2 is left BLANK — substituting someone would misattribute the day',
     r.sheet._cells.I2.val === '', r.sheet._cells.I2.val);
  ok('H4 and the validation still moved, so it can be re-picked',
     !!r.sheet._cells.I2.dv && r.sheet._cells.I2.dv.getCriteriaValues()[0].length === 5);
});

console.log('\n' + (fail ? '✗ ' : '✅ ') + 'test-pickid-migrate: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
