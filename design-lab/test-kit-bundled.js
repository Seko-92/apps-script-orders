// =====================================================================================
// test-kit-bundled.js — the BUNDLED-COMPONENT rule, against the REAL KitRegistry.js
// and the REAL 223-kit registry (design-lab/fixture-kit-registry.json).
//
// The rule: a component named only AFTER a "+" on a Sales Description line ships
// INSIDE that parent, so it must not be priced, picked, or counted for buildability.
// Naming its own sales line (own MPN) means it is genuinely separate.
//
// ⚠ THE ASYMMETRY UNDER TEST: excluding a real part ships an incomplete box;
// including a bundled one costs one wasted walk. So every tie must resolve to KEEP.
// =====================================================================================
const fs = require('fs'), path = require('path'), vm = require('vm');

const src = fs.readFileSync(path.join(__dirname, '..', 'KitRegistry.js'), 'utf8');
const ctx = { console, Map, JSON, String, Number, parseInt, parseFloat, isNaN, Math, Array, Object, RegExp };
vm.createContext(ctx);
// Load ONLY the pure block — the rest of the file needs SpreadsheetApp.
const start = src.indexOf('var KIT_BUNDLE = {');
const end   = src.indexOf('// =======================================================================================\n// READ API');
if (start < 0 || end < 0) { console.error('FATAL: could not locate the pure block in KitRegistry.js'); process.exit(1); }
vm.runInContext(src.slice(start, end), ctx);
const { kitBundledComponents, _kbTokens, _kbSalesSegments, _kbSegNamesComponent, _kbAnnotateKit } = ctx;

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; }
  else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}
function section(t) { console.log('\n── ' + t); }

// ---------------------------------------------------------------- tokenising
section('tokens');
ok('drops MPNs and quantities', JSON.stringify(_kbTokens('-1 Full Gasket Set 19077-03310')) === JSON.stringify(['full','gasket']), JSON.stringify(_kbTokens('-1 Full Gasket Set 19077-03310')));
ok('drops parentheticals', JSON.stringify(_kbTokens('Head Gasket (Composite)')) === JSON.stringify(['head','gasket']));
ok('singularises', JSON.stringify(_kbTokens('Thrust Washers Set STD')) === JSON.stringify(['thrust','washer']));
ok('noise-only text yields nothing', _kbTokens('Set STD Kit').length === 0);
ok('empty is safe', _kbTokens('').length === 0 && _kbTokens(null).length === 0);

section('segment ↔ component naming');
ok('"Head Gasket" names "Head Gasket (Composite)"', _kbSegNamesComponent('Head Gasket', 'Head Gasket (Composite)'));
ok('"Thrust Washers" names "Thrust Washer set"', _kbSegNamesComponent('Thrust Washers', 'Thrust Washer set'));
ok('"Full Gasket Set" does NOT name "Head Gasket"', !_kbSegNamesComponent('-1 Full Gasket Set 19077-03310', 'Head Gasket'));
ok('"Metal Head Gasket" names "Head Gasket Metal"', _kbSegNamesComponent('Metal Head Gasket 1C020-03310', 'Head Gasket Metal'));
ok('a segment with no significant tokens never matches', !_kbSegNamesComponent('STD Set', 'Head Gasket'));

section('sales-description parsing');
const wrapped = _kbSalesSegments('-1 Main Bearings Set STD 15274-23480 +\nThrust Washers Set STD\n-3 Connecting Rod Bearings STD 15694-22310');
ok('a line ending in "+" joins the next line', wrapped.bundled.some(b => /Thrust Washers Set STD/.test(b.text)), JSON.stringify(wrapped.bundled));
ok('the continuation is NOT read as its own primary', !wrapped.primary.some(p => /^Thrust Washers/.test(p.trim())), JSON.stringify(wrapped.primary));
ok('normal lines still become primaries', wrapped.primary.some(p => /Connecting Rod Bearings/.test(p)));

// ---------------------------------------------------------------- the real cases
section('the real kits that started this');
function bundledFor(kit) { return kitBundledComponents(kit.components, kit.salesDescription); }

// 155547 — Lister Petter: PD has a separate thrust washer, SD bundles it.
const k155547 = { components: [
    { sku:'215225', qty:1, name:'Main Bearing Set STD' },
    { sku:'162288', qty:1, name:'Thrust Washer set' },
    { sku:'164259', qty:1, name:'Full Gasket Set' } ],
  salesDescription: '- 1 Full Gasket Set 657-34281\n- 1 Main Bearing Set STD 750-11250 + Thrust Washer\n- 4 Connecting Rod Bearing STD 750-11260' };
let b = bundledFor(k155547);
ok('155547 thrust washer 162288 is BUNDLED', !!b['162288'], JSON.stringify(b));
ok('155547 names the parent it hides in', b['162288'] && /Main Bearing/.test(b['162288'].into));
ok('155547 main bearing itself is NOT bundled', !b['215225']);
ok('155547 full gasket set is NOT bundled', !b['164259']);

// 155565 — Deutz: thrust washer has its OWN sales line with its own MPN.
const k155565 = { components: [
    { sku:'172332', qty:5, name:'Main Bearing STD' },
    { sku:'170163', qty:1, name:'Thrust Washer Set' } ],
  salesDescription: '-5 Main Bearing STD 03362379\n-1 Thrust Washer Set 04236912' };
ok('155565 thrust washer 170163 is KEPT (own MPN line)', !bundledFor(k155565)['170163']);

// 217565 — the clearest double-count: the purchase line itself says the set includes it.
const k217565 = { components: [
    { sku:'162927', qty:1, name:'Full Gasket Set + Metal Head Gasket' },
    { sku:'162801', qty:1, name:'Head Gasket Metal' },
    { sku:'168858', qty:1, name:'Main Bearing set STD' } ],
  salesDescription: '-1 Full Gasket Set + Metal Head Gasket 1C020-03310\n-1 Main Bearing set STD 1C020-23470' };
b = bundledFor(k217565);
ok('217565 the loose head gasket 162801 is BUNDLED', !!b['162801'], JSON.stringify(b));
ok('217565 the merged parent line is NEVER bundled into itself', !b['162927']);

// 217034 — Kubota, head gasket bundled by the "+ Head Gasket" annotation.
const k217034 = { components: [
    { sku:'156492', qty:1, name:'Full Gasket Set' },
    { sku:'168804', qty:1, name:'Head Gasket (Composite)' },
    { sku:'158328', qty:1, name:'Main Bearings Set STD' } ],
  salesDescription: '-1 Full Gasket Set 19077-03310 + Head Gasket\n-1 Main Bearings Set STD 17311-23470 + Thrust Washers' };
b = bundledFor(k217034);
ok('217034 head gasket 168804 is BUNDLED', !!b['168804']);
ok('217034 gasket set is KEPT', !b['156492']);
ok('217034 main bearings kept (it is the parent, not the washer)', !b['158328']);

// 157536 — Deutz: BOTH parts have their own MPN lines. Nothing may be dropped.
const k157536 = { components: [
    { sku:'166374', qty:1, name:'Full Gasket set' },
    { sku:'164988', qty:1, name:'Head Gasket' },
    { sku:'167697', qty:1, name:'Thrust Washer STD' } ],
  salesDescription: '-1 Full Gasket set 02929856\n-1 Head Gasket 04272391\n-1 Thrust Washer STD  02928961' };
b = bundledFor(k157536);
ok('157536 head gasket KEPT (own MPN)', !b['164988'], JSON.stringify(b));
ok('157536 thrust washer KEPT (own MPN)', !b['167697']);

section('the safe direction');
ok('no sales description at all → nothing is dropped',
   Object.keys(kitBundledComponents(k217034.components, '')).length === 0);
ok('sales description with no "+" → nothing is dropped',
   Object.keys(kitBundledComponents(k217034.components, '-1 Full Gasket Set 19077-03310')).length === 0);
ok('a one-word "+" segment is too weak to drop a part',
   Object.keys(kitBundledComponents(
     [{sku:'162288',qty:1,name:'Thrust Washer set'}],
     '-1 Main Bearing Set 750-11250 + Washers')).length === 0);
ok('empty component list is safe', Object.keys(kitBundledComponents([], 'x + Head Gasket')).length === 0);
ok('null inputs are safe', Object.keys(kitBundledComponents(null, null)).length === 0);

section('_kbAnnotateKit writes the flags onto the kit');
const kk = JSON.parse(JSON.stringify(k217034));
_kbAnnotateKit(kk);
ok('bundled component carries bundled=true', kk.components[1].bundled === true);
ok('bundled component names its parent', /Full Gasket Set/.test(kk.components[1].bundledInto || ''));
ok('kept components are explicitly bundled=false', kk.components[0].bundled === false);

// ---------------------------------------------------------------- whole catalog
section('the whole live registry (223 kits, 1509 components)');
const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture-kit-registry.json'), 'utf8'));
let nBundled = 0, nKept = 0, byKind = {};
const bundledList = [];
fixture.forEach(k => {
  const flags = kitBundledComponents(k.components, k.salesDescription);
  k.components.forEach(c => {
    if (flags[c.sku]) {
      nBundled++;
      bundledList.push({ kit: k.sku, sku: c.sku, name: c.name, into: flags[c.sku].into });
      const kind = /head gasket/i.test(c.name) ? 'head gasket'
                 : /thrust washer/i.test(c.name) ? 'thrust washer' : 'other';
      byKind[kind] = (byKind[kind] || 0) + 1;
    } else nKept++;
  });
});
console.log('   bundled: ' + nBundled + '   kept: ' + nKept + '   by kind: ' + JSON.stringify(byKind));
ok('the rule fires on a real, non-trivial share of the catalog', nBundled >= 40 && nBundled <= 70, 'got ' + nBundled);
ok('it never drops a majority of any kit', fixture.every(k => {
    const f = kitBundledComponents(k.components, k.salesDescription);
    const dropped = k.components.filter(c => f[c.sku]).length;
    return k.components.length === 0 || dropped <= k.components.length / 2;
  }));
ok('no kit is emptied by the rule', fixture.every(k => {
    const f = kitBundledComponents(k.components, k.salesDescription);
    return k.components.length === 0 || k.components.some(c => !f[c.sku]);
  }));
ok('every bundled component names a parent', bundledList.every(b => b.into && b.into.length > 0));
ok('only gasket/washer-class parts are being dropped today', !byKind.other, JSON.stringify(byKind));

// idempotency + purity
section('purity');
const before = JSON.stringify(k217034);
kitBundledComponents(k217034.components, k217034.salesDescription);
ok('kitBundledComponents does not mutate its input', JSON.stringify(k217034) === before);
const r1 = JSON.stringify(kitBundledComponents(k217034.components, k217034.salesDescription));
const r2 = JSON.stringify(kitBundledComponents(k217034.components, k217034.salesDescription));
ok('repeat calls agree', r1 === r2);

console.log('\n' + (fail === 0 ? '✅ ' : '❌ ') + pass + ' passed, ' + fail + ' failed');
fs.writeFileSync(path.join(__dirname, 'bundled-report.json'), JSON.stringify(bundledList, null, 1));
console.log('   wrote design-lab/bundled-report.json (' + bundledList.length + ' lines)');
process.exit(fail === 0 ? 0 : 1);
