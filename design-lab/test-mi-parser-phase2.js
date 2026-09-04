/**
 * test-mi-parser-phase2.js — proves the 17 new columns, on BOTH parsers.
 *
 * MAIN (hourly) and SUB (the weekend full pass) must emit an IDENTICAL key set.
 * SUB is what backfills these columns across all ~3,600 rows; if the two drift,
 * half the sheet fills and the other half quietly does not — which is exactly
 * the picture1 / pictureUrl1 divergence Phase 1 just spent a push repairing.
 * Section D is the guard against repeating it.
 */
const fs = require('fs');
const path = require('path');
const { XML_KIT, XML_FLAT, XML_METRIC } = require('./fixture-getitem');

const FRESH = path.join(__dirname, '..', '..', 'n8n WorkFlows', 'FInal', 'Feb 28-2026', 'Fresh');
const MAIN = fs.readFileSync(path.join(FRESH, 'MAIN-node10-parse-PHASE2-2026-09-04.js'), 'utf8');
const SUB  = fs.readFileSync(path.join(FRESH, 'SUB-node5-parse-PHASE2-2026-09-04.js'),  'utf8');

function run(code, xml) {
  const json = { itemId: '153889746781', sku: '157860', xmlResponse: xml };
  const $input = { first: () => ({ json }), all: () => [{ json }] };
  return new Function('$input', 'console', code)($input, { log(){}, error(){}, warn(){} })[0].json;
}

const m       = run(MAIN, XML_KIT);
const s       = run(SUB,  XML_KIT);
const mFlat   = run(MAIN, XML_FLAT);
const sFlat   = run(SUB,  XML_FLAT);
const mMetric = run(MAIN, XML_METRIC);
// oz arithmetic: 25 lbs 8 oz must land on 25.5, not 25 and not 33
const XML_OZ  = XML_KIT.replace(
  /(<ShippingPackageDetails>[\s\S]*?)<WeightMinor unit="oz"([^>]*)>0</,
  '$1<WeightMinor unit="oz"$2>8<');
const mOz     = run(MAIN, XML_OZ);

let pass = 0, fail = 0; const rows = [];
function ok(id, desc, got, want) {
  const good = JSON.stringify(got) === JSON.stringify(want);
  good ? pass++ : fail++;
  rows.push({ id, desc, got, want, good });
}

// ── A. PARCEL — the scoping trap ─────────────────────────────────────────────
ok('A1', 'weight comes from ShippingPackageDetails, NOT CalculatedShippingRate', m.packageWeightLbs, '25');
ok('A2', '…and SUB agrees',                                                     s.packageWeightLbs, '25');
ok('A3', 'length',                            m.packageLengthIn, '20.00');
ok('A4', 'width',                             m.packageWidthIn,  '18.00');
ok('A5', 'depth',                             m.packageDepthIn,  '10.00');
ok('A6', 'dims unit captured',                m.packageDimsUnit, 'inches');
ok('A7', 'shippingIrregular',                 m.shippingIrregular, 'false');
ok('A8', 'shippingPackage',                   m.shippingPackage,   'None');
ok('A9', 'lbs + oz/16 — 25 lbs 8 oz = 25.5',  mOz.packageWeightLbs, '25.5');
ok('A9b','…and the EXACT integer: 25 lbs 8 oz = 408 oz', mOz.packageWeightOz, '408');
ok('A9c','no parcel block → oz EMPTY, not 0',   mFlat.packageWeightOz, '');
ok('A10','metric listing is VISIBLE, not silently mixed', mMetric.packageDimsUnit, 'centimeters');

// ── B. NEGATIVE — flat-rate listing must yield EMPTY, never wrong ────────────
// The one that would ship silently: no parcel block at all, and the decoy
// CalculatedShippingRate gone too.
for (const [id, k] of [['B1','packageWeightLbs'], ['B2','packageLengthIn'],
                       ['B3','packageWidthIn'],  ['B4','packageDepthIn'],
                       ['B5','packageDimsUnit'], ['B6','shippingIrregular'],
                       ['B7','shippingPackage']]) {
  ok(id, `no parcel block → ${k} is EMPTY, not wrong`, mFlat[k], '');
}
ok('B8', '…and SUB behaves identically',            sFlat.packageWeightLbs, '');
ok('B9', 'flat-rate row still parses (sku intact)', mFlat.sku, '157860');

// ── C. SHIPPING · HEALTH · DEMAND ───────────────────────────────────────────
ok('C1', 'dispatchTimeMax',      m.dispatchTimeMax, '0');
ok('C2', 'shippingService',      m.shippingService, 'ShippingMethodStandard');
ok('C3', 'shippingCost',         m.shippingCost,    '0.0');
ok('C4', 'freeShipping',         m.freeShipping,    'true');
ok('C5', 'shipToLocations',      m.shipToLocations, 'US');
ok('C6', 'hideFromSearch — eBay\'s own verdict', m.hideFromSearch,       'true');
ok('C7', 'reasonHideFromSearch',                 m.reasonHideFromSearch, 'OutOfStock');
ok('C8', 'outOfStockControl',    m.outOfStockControl, 'true');
ok('C9', 'bestOfferEnabled',     m.bestOfferEnabled,  'false');
ok('C10','watchCount',           m.watchCount,        '47');
ok('C11','SUB emits the same health verdict',    s.reasonHideFromSearch, 'OutOfStock');

// ── D. THE DRIFT GUARD — MAIN and SUB must emit the SAME key set ────────────
const SEVENTEEN = ['packageWeightLbs','packageWeightOz','packageLengthIn','packageWidthIn','packageDepthIn',
  'packageDimsUnit','shippingIrregular','shippingPackage','dispatchTimeMax','shippingService',
  'shippingCost','freeShipping','shipToLocations','hideFromSearch','reasonHideFromSearch',
  'outOfStockControl','bestOfferEnabled','watchCount'];
ok('D1', 'all 17 present in MAIN', SEVENTEEN.filter(k => !(k in m)), []);
ok('D2', 'all 17 present in SUB',  SEVENTEEN.filter(k => !(k in s)), []);
const mk = Object.keys(m).filter(k => k.indexOf('C:') !== 0 && k !== 'lastUpdated').sort();
const sk = Object.keys(s).filter(k => k.indexOf('C:') !== 0 && k !== 'lastUpdated').sort();
ok('D3', 'MAIN has no key SUB lacks', mk.filter(k => sk.indexOf(k) < 0), []);
ok('D4', 'SUB has no key MAIN lacks', sk.filter(k => mk.indexOf(k) < 0), []);
ok('D5', 'no name collides with the existing 198 headers',
   SEVENTEEN.filter(k => ['sku','title','quantity','quantitySold','location','error',
     'shippingType','lastUpdated','itemId','postalCode','country'].indexOf(k) >= 0), []);

// ── E. PHASE 1 IS STILL INTACT (regression) ─────────────────────────────────
ok('E1', 'multi-value specifics still joined',
   m['C:Compatible Equipment Type'],
   'Ditch Witch, Crawler Tractor, JLG, Genie, Dynapac, Boom Lift, Crawler Dozer');
ok('E2', 'variation specifics still excluded', m['C:Bore Size'], undefined);
ok('E3', 'pictures still scoped',              m.pictureUrl3, '');
ok('E4', 'error still cleared on success',     m.error, '');
ok('E5', 'returnPeriod still present',         m.returnPeriod, '30 Days');
ok('E6', 'sku unchanged',                      m.sku, '157860');

// ── report ───────────────────────────────────────────────────────────────────
console.log('\n  MASTER INVENTORY — PHASE 2 PROOF  (MAIN + SUB)\n');
let sec = '';
const LBL = { A:'── PARCEL (scoped) ──', B:'── NEGATIVE: no parcel block ──',
              C:'── SHIPPING · HEALTH · DEMAND ──', D:'── DRIFT GUARD: MAIN ≡ SUB ──',
              E:'── PHASE 1 REGRESSION ──' };
for (const r of rows) {
  if (r.id[0] !== sec) { sec = r.id[0]; console.log('  ' + LBL[sec]); }
  console.log(`   ${r.good ? '✅' : '❌'}  ${r.id.padEnd(4)} ${r.desc}`);
  if (!r.good) console.log(`          want ${JSON.stringify(r.want)}  got ${JSON.stringify(r.got)}`);
}
console.log(`\n  ${pass} passed, ${fail} failed\n`);
process.exit(fail ? 1 : 0);
