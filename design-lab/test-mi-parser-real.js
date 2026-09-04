/**
 * test-mi-parser-real.js — both Phase 2 parsers against a REAL GetItem payload.
 *
 * fixture-real-153244362081.xml is the actual response for SKU 172827, captured
 * from node 8 on 2026-09-04 (Description stripped — 25 KB of escaped HTML that
 * no parser reads, and Gotcha #13 says it must never reach a cell).
 *
 * The hand-built fixture proves the traps we designed FOR. This one proves the
 * shapes eBay actually sends, which is a different question — the singular
 * <ShipToLocation> tag below is one nobody predicted.
 */
const fs = require('fs'), path = require('path');
const F = path.join(__dirname, '..', '..', 'n8n WorkFlows', 'FInal', 'Feb 28-2026', 'Fresh');
const xml = fs.readFileSync(path.join(__dirname, 'fixture-real-153244362081.xml'), 'utf8');

function run(file) {
  const json = { itemId: '153244362081', sku: '172827', xmlResponse: xml };
  const $input = { first: () => ({ json }), all: () => [{ json }] };
  return new Function('$input', 'console',
    fs.readFileSync(path.join(F, file), 'utf8'))($input, { log(){}, error(){}, warn(){} })[0].json;
}
const m = run('MAIN-node10-parse-PHASE2-2026-09-04.js');
const s = run('SUB-node5-parse-PHASE2-2026-09-04.js');

let pass = 0, fail = 0;
const ok = (d, got, want) => {
  const good = JSON.stringify(got) === JSON.stringify(want);
  good ? pass++ : fail++;
  console.log(`   ${good ? '✅' : '❌'} ${d}`);
  if (!good) console.log(`        want ${JSON.stringify(want)}  got ${JSON.stringify(got)}`);
};

console.log('\n  REAL PAYLOAD — item 153244362081 / SKU 172827\n');

// 0 lbs 4 oz. Both WeightMajor copies happen to agree on THIS item, so the
// scoping is not exercised here — the trap is unproven on real data, not absent.
ok('weight 0 lbs 4 oz → 0.25',              m.packageWeightLbs, '0.25');
ok('dims 7 x 5 x 1 inches',
   [m.packageLengthIn, m.packageWidthIn, m.packageDepthIn, m.packageDimsUnit],
   ['7.00','5.00','1.00','inches']);
ok('shippingPackage',                        m.shippingPackage, 'PackageThickEnvelope');

// ⚠ THE ONE NOBODY PREDICTED. The payload carries ONE <ShipToLocations> (plural,
// Item level) and FOUR <ShipToLocation> (singular, inside each
// InternationalShippingServiceOption). Sloppy matching yields
// "Worldwide, Worldwide, Worldwide, Worldwide, Worldwide".
ok('shipToLocations ignores the 4 singular <ShipToLocation> tags',
   m.shipToLocations, 'Worldwide');

// First domestic option only — ShippingService also appears in all four
// InternationalShippingServiceOption blocks.
ok('shippingService takes the DOMESTIC option', m.shippingService, 'USPSParcel');
ok('shippingCost / freeShipping',            [m.shippingCost, m.freeShipping], ['0.0','true']);
ok('dispatchTimeMax',                        m.dispatchTimeMax, '0');

// hideFromSearch present; ReasonHideFromSearch absent because it is NOT hidden.
ok('hideFromSearch false, reason blank',     [m.hideFromSearch, m.reasonHideFromSearch], ['false','']);
// ⚠ NOT a stock state — it is the seller's out-of-stock FEATURE flag. This item
// has 33 available and still reads true. Do not read this column as "is OOS".
ok('outOfStockControl is a SETTING, not a state', m.outOfStockControl, 'true');

// Absent from the response entirely; empty is the correct answer, not a bug.
ok('bestOfferEnabled absent → empty',        m.bestOfferEnabled, '');
ok('watchCount absent → empty',              m.watchCount, '');

// Phase 1 still holding on the real shape.
ok('multi-value specific joined',
   m['C:Compatible Equipment Type'], 'Marine, Ditch Witch, Tower light, Boom Lift');
ok('single value containing a comma is untouched', m['C:Model'], 'SR, LD');
ok('returnPeriod, not the sibling Refund/RefundOption', m.returnPeriod, '30 Days');
ok('error cleared',                          m.error, '');
ok('pictures scoped to PictureDetails',      m.totalPictures, '2');

// ⚠ SUB is the WEEKEND BACKFILL. If it drifts from MAIN, Saturday writes ~3,600
// rows without these columns and the sheet fills only where MAIN happened to go.
const KEYS = ['packageWeightLbs','packageLengthIn','packageWidthIn','packageDepthIn',
 'packageDimsUnit','shippingIrregular','shippingPackage','dispatchTimeMax','shippingService',
 'shippingCost','freeShipping','shipToLocations','hideFromSearch','reasonHideFromSearch',
 'outOfStockControl','bestOfferEnabled','watchCount'];
ok('MAIN ≡ SUB on all 17', KEYS.filter(k => JSON.stringify(m[k]) !== JSON.stringify(s[k])), []);

console.log(`\n  ${pass} passed, ${fail} failed\n`);
process.exit(fail ? 1 : 0);
