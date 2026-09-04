/**
 * test-mi-parser-phase1.js — proves the Phase 1 parser repairs.
 *
 * Loads the REAL code on both sides so the test cannot drift:
 *   BEFORE = "10. Parse & Flatten" lifted out of the live MAIN workflow export
 *   AFTER  = Fresh/MAIN-node10-parse-PHASE1-2026-09-04.js, the file you will paste
 *
 * The fixture is NOT a copy of a real eBay response — I don't have the golden
 * payload in this session. It is hand-built to carry the exact traps the plan
 * names, so it proves the parser LOGIC. The golden fixture (item 153889746781)
 * is still the acceptance test, run manually through the node in n8n.
 *
 * Every assertion is reported for BOTH parsers. Ones that FAIL on before and PASS
 * on after are the repairs; ones that pass on BOTH are the regression nets that
 * show the change is surgical.
 */

const fs = require('fs');
const path = require('path');

const FRESH = path.join(__dirname, '..', '..', 'n8n WorkFlows', 'FInal', 'Feb 28-2026', 'Fresh');

// ── load the two parsers ──────────────────────────────────────────────────────
function liveNodeCode(file, nodeName) {
  const wf = JSON.parse(fs.readFileSync(path.join(FRESH, file), 'utf8'));
  const n = wf.nodes.find(n => n.name === nodeName);
  if (!n) throw new Error('node not found: ' + nodeName);
  return n.parameters.jsCode;
}
const BEFORE = liveNodeCode('MAIN - Master Inventory Full Sync.json', '10. Parse & Flatten');
const AFTER  = fs.readFileSync(path.join(FRESH, 'MAIN-node10-parse-PHASE1-2026-09-04.js'), 'utf8');

// n8n Code-node shim: the body uses $input and console, and top-level `return`.
function runParser(code, inputJson) {
  const $input = { first: () => ({ json: inputJson }), all: () => [{ json: inputJson }] };
  const quietConsole = { log(){}, error(){}, warn(){} };
  const fn = new Function('$input', 'console', code);
  return fn($input, quietConsole)[0].json;
}

// ── fixture ───────────────────────────────────────────────────────────────────
// Element order follows eBay's schema order: PictureDetails, then ShippingDetails,
// then Variations near the end. That ordering is what makes the picture leak
// realistic rather than contrived.
const XML_KIT = `<?xml version="1.0" encoding="utf-8"?>
<GetItemResponse xmlns="urn:ebay:apis:eBLBaseComponents">
<Ack>Success</Ack>
<Item>
  <ItemID>153889746781</ItemID>
  <SKU>157860</SKU>
  <Title>Engine Overhaul Kit</Title>
  <ConditionID>1000</ConditionID>
  <ConditionDisplayName>New</ConditionDisplayName>
  <Country>US</Country>
  <Location>Houston, Texas</Location>
  <PostalCode>77passed</PostalCode>
  <ListingType>FixedPriceItem</ListingType>
  <ListingDuration>GTC</ListingDuration>
  <StartPrice currencyID="USD">980.00</StartPrice>
  <Quantity>7</Quantity>
  <HideFromSearch>true</HideFromSearch>
  <ReasonHideFromSearch>OutOfStock</ReasonHideFromSearch>
  <OutOfStockControl>true</OutOfStockControl>
  <BestOfferEnabled>false</BestOfferEnabled>
  <DispatchTimeMax>0</DispatchTimeMax>
  <ShipToLocations>US</ShipToLocations>
  <PictureDetails>
    <PictureURL>https://i.ebayimg.com/REAL-1.jpg</PictureURL>
    <PictureURL>https://i.ebayimg.com/REAL-2.jpg</PictureURL>
  </PictureDetails>
  <PrimaryCategory>
    <CategoryID>33615</CategoryID>
    <CategoryName>Heavy Equipment Parts</CategoryName>
  </PrimaryCategory>
  <Seller>
    <UserID>hqmotorservice</UserID>
    <FeedbackScore>4821</FeedbackScore>
    <PositiveFeedbackPercent>100.0</PositiveFeedbackPercent>
  </Seller>
  <SellingStatus>
    <CurrentPrice currencyID="USD">980.00</CurrentPrice>
    <QuantitySold>3</QuantitySold>
    <ListingStatus>Active</ListingStatus>
  </SellingStatus>
  <ShippingDetails>
    <ShippingType>Flat</ShippingType>
    <CalculatedShippingRate>
      <WeightMajor unit="lbs" measurementSystem="English">99</WeightMajor>
      <WeightMinor unit="oz" measurementSystem="English">0</WeightMinor>
    </CalculatedShippingRate>
    <ShippingServiceOptions>
      <ShippingService>ShippingMethodStandard</ShippingService>
      <ShippingServiceCost currencyID="USD">0.0</ShippingServiceCost>
      <FreeShipping>true</FreeShipping>
    </ShippingServiceOptions>
    <ShippingPackageDetails>
      <PackageDepth measurementSystem="English" unit="inches">10.00</PackageDepth>
      <PackageLength measurementSystem="English" unit="inches">20.00</PackageLength>
      <PackageWidth measurementSystem="English" unit="inches">18.00</PackageWidth>
      <ShippingIrregular>false</ShippingIrregular>
      <ShippingPackage>None</ShippingPackage>
      <WeightMajor unit="lbs" measurementSystem="English">25</WeightMajor>
      <WeightMinor unit="oz" measurementSystem="English">0</WeightMinor>
    </ShippingPackageDetails>
  </ShippingDetails>
  <ListingDetails>
    <StartTime>2026-01-04T18:00:00.000Z</StartTime>
    <EndTime>2026-12-04T18:00:00.000Z</EndTime>
    <ViewItemURL>https://www.ebay.com/itm/153889746781</ViewItemURL>
  </ListingDetails>
  <ReturnPolicy>
    <ReturnsAccepted>ReturnsAccepted</ReturnsAccepted>
    <ReturnsWithin>30 Days</ReturnsWithin>
  </ReturnPolicy>
  <ItemSpecifics>
    <NameValueList>
      <Name>Compatible Equipment Type</Name>
      <Value>Ditch Witch</Value>
      <Value>Crawler Tractor</Value>
      <Value>JLG</Value>
      <Value>Genie</Value>
      <Value>Dynapac</Value>
      <Value>Boom Lift</Value>
      <Value>Crawler Dozer</Value>
    </NameValueList>
    <NameValueList>
      <Name>Model Year</Name>
      <Value>K-55</Value>
    </NameValueList>
    <NameValueList>
      <Name>Brand</Name>
      <Value>HQ</Value>
    </NameValueList>
  </ItemSpecifics>
  <Variations>
    <Pictures>
      <VariationSpecificPictureSet>
        <PictureURL>https://i.ebayimg.com/VARIATION-A.jpg</PictureURL>
        <PictureURL>https://i.ebayimg.com/VARIATION-B.jpg</PictureURL>
        <PictureURL>https://i.ebayimg.com/VARIATION-C.jpg</PictureURL>
      </VariationSpecificPictureSet>
    </Pictures>
    <VariationSpecificsSet>
      <NameValueList>
        <Name>Bore Size</Name>
        <Value>STD</Value>
      </NameValueList>
    </VariationSpecificsSet>
  </Variations>
</Item>
</GetItemResponse>`;

// Flat-rate listing: NO CalculatedShippingRate, NO ShippingPackageDetails.
// This is the Phase 2 negative test — parcel fields must come back EMPTY, not wrong.
const XML_FLAT = XML_KIT
  .replace(/<CalculatedShippingRate>[\s\S]*?<\/CalculatedShippingRate>/, '')
  .replace(/<ShippingPackageDetails>[\s\S]*?<\/ShippingPackageDetails>/, '');

const XML_FAIL = `<?xml version="1.0"?><GetItemResponse><Ack>Failure</Ack>
<Errors><ShortMessage>Call usage limit has been reached.</ShortMessage></Errors></GetItemResponse>`;

// ── harness ───────────────────────────────────────────────────────────────────
let pass = 0, fail = 0;
const rows = [];
function check(id, desc, beforeFn, afterFn, wantBefore, wantAfter) {
  let b, a;
  try { b = beforeFn(); } catch (e) { b = 'THREW: ' + e.message; }
  try { a = afterFn();  } catch (e) { a = 'THREW: ' + e.message; }
  // bOk = the BEFORE parser behaved as predicted. If that misses, my model of the
  // CURRENT code is wrong — which is a finding in itself, not a pass.
  const bOk = JSON.stringify(b) === JSON.stringify(wantBefore);
  const aOk = JSON.stringify(a) === JSON.stringify(wantAfter);
  // A REPAIR is defined by the EXPECTATIONS differing, never by whether they matched.
  const isRepair = JSON.stringify(wantBefore) !== JSON.stringify(wantAfter);
  if (aOk && bOk) pass++; else fail++;
  rows.push({ id, desc, b, a, bOk, aOk, isRepair, wantBefore, wantAfter });
}

const before = runParser(BEFORE, { itemId: '153889746781', xmlResponse: XML_KIT });
const after  = runParser(AFTER,  { itemId: '153889746781', xmlResponse: XML_KIT });
const afterFlat = runParser(AFTER, { itemId: '1', xmlResponse: XML_FLAT });
const beforeFail = runParser(BEFORE, { itemId: 'X', xmlResponse: XML_FAIL });
const afterFail  = runParser(AFTER,  { itemId: 'X', xmlResponse: XML_FAIL });

const SEVEN = 'Ditch Witch, Crawler Tractor, JLG, Genie, Dynapac, Boom Lift, Crawler Dozer';

// ── A. THE REPAIRS (must fail before, pass after) ─────────────────────────────
check('A1', 'multi-value item specific keeps ALL 7 values',
  () => before['C:Compatible Equipment Type'], () => after['C:Compatible Equipment Type'],
  'Ditch Witch', SEVEN);

check('A2', 'variation specifics do NOT leak into C: columns',
  () => before['C:Bore Size'], () => after['C:Bore Size'],
  'STD', undefined);

check('A3', 'pictures land in pictureUrl1 (the key 4 readers use)',
  () => before.pictureUrl1, () => after.pictureUrl1,
  undefined, 'https://i.ebayimg.com/REAL-1.jpg');

check('A4', 'picture slots are NOT padded with variation images',
  () => before.picture3, () => after.pictureUrl3,
  'https://i.ebayimg.com/VARIATION-A.jpg', '');

check('A5', 'totalPictures counts only real PictureDetails images',
  () => before.totalPictures, () => after.totalPictures,
  undefined, '2');

check('A6', 'error is cleared on success',
  () => before.error, () => after.error,
  undefined, '');

check('A7', 'returns period uses SUB\'s key name',
  () => before.returnPeriod, () => after.returnPeriod,
  undefined, '30 Days');

check('A8', 'old picture1 key is gone',
  () => typeof before.picture1, () => typeof after.picture1,
  'string', 'undefined');

check('A9', 'extractBlock exists (Phase 2 prerequisite)',
  () => /function extractBlock/.test(BEFORE), () => /function extractBlock/.test(AFTER),
  false, true);

check('A10', 'no dead Parent>Child calls remain',
  () => (BEFORE.match(/extractValue\([^,]+,\s*'[^']*>[^']*'\)/g) || []).length,
  () => (AFTER.match(/extractValue\([^,]+,\s*'[^']*>[^']*'\)/g) || []).length,
  3, 0);

// ── B. REGRESSION NETS (must pass on BOTH — proves the change is surgical) ────
const nets = [
  ['B1', 'itemId',               'itemId',               '153889746781'],
  ['B2', 'sku',                  'sku',                  '157860'],
  ['B3', 'title',                'title',                'Engine Overhaul Kit'],
  ['B4', 'quantity',             'quantity',             '7'],
  ['B5', 'quantitySold',         'quantitySold',         '3'],
  ['B6', 'listingStatus',        'listingStatus',        'Active'],
  ['B7', 'currentPrice',         'currentPrice',         '980.00'],
  ['B8', 'viewItemURL',          'viewItemURL',          'https://www.ebay.com/itm/153889746781'],
  ['B9', 'primaryCategoryId (fallback path)', 'primaryCategoryId', '33615'],
  ['B10','sellerUserID (fallback path)',      'sellerUserID',      'hqmotorservice'],
  ['B11','single-value specific unchanged',   'C:Model Year',      'K-55'],
  ['B12','shippingType',         'shippingType',         'Flat'],
];
for (const [id, desc, key, want] of nets) {
  check(id, desc, () => before[key], () => after[key], want, want);
}
check('B13', 'Ack=Failure still returns the real error',
  () => beforeFail.error, () => afterFail.error,
  'Call usage limit has been reached.', 'Call usage limit has been reached.');
check('B14', 'Ack=Failure does NOT blank the error column',
  () => beforeFail.error !== '', () => afterFail.error !== '', true, true);

// ── C. PHASE 2 TRAPS (documented now, asserted against the fixture) ───────────
// Not parser behaviour — these prove the fixture carries the traps Phase 2 must survive.
const unscopedWeight = (XML_KIT.match(/<WeightMajor\b[^>]*>([^<]*)<\/WeightMajor>/) || [])[1];
const scopedBlock = (XML_KIT.match(/<ShippingPackageDetails\b[^>]*>[\s\S]*?<\/ShippingPackageDetails>/) || [])[0] || '';
const scopedWeight = (scopedBlock.match(/<WeightMajor\b[^>]*>([^<]*)<\/WeightMajor>/) || [])[1];
check('C1', 'PHASE 2 TRAP: unscoped WeightMajor grabs the WRONG copy',
  () => unscopedWeight, () => scopedWeight, '99', '25');
const flatBlock = (XML_FLAT.match(/<ShippingPackageDetails\b[^>]*>[\s\S]*?<\/ShippingPackageDetails>/) || [])[0] || '';
check('C2', 'PHASE 2 NEGATIVE: flat-rate listing yields NO parcel block',
  () => flatBlock === '', () => flatBlock === '', true, true);
check('C3', 'flat-rate listing still parses cleanly (no crash, sku intact)',
  () => afterFlat.sku, () => afterFlat.sku, '157860', '157860');

// ── report ────────────────────────────────────────────────────────────────────
const W = s => String(s === undefined ? 'undefined' : s).slice(0, 46);
console.log('\n  MASTER INVENTORY — PHASE 1 PARSER PROOF');
console.log('  BEFORE = live MAIN node 10   AFTER = MAIN-node10-parse-PHASE1-2026-09-04.js\n');
let section = '';
for (const r of rows) {
  const s = r.id[0];
  if (s !== section) {
    section = s;
    console.log('  ' + { A: '── REPAIRS (fail before → pass after) ──',
                         B: '── REGRESSION NETS (pass on both) ──',
                         C: '── PHASE 2 TRAPS (fixture carries them) ──' }[s]);
  }
  let verdict;
  if (!r.aOk)      verdict = ' ❌ AFTER ';
  else if (!r.bOk) verdict = ' ❌ BEFORE';   // my model of the live code was wrong
  else if (r.isRepair) verdict = ' ✅ FIXED ';
  else             verdict = '  ·  NET  ';
  console.log(`   ${verdict} ${r.id.padEnd(4)} ${r.desc}`);
  if (!r.aOk) console.log(`             AFTER  want ${JSON.stringify(r.wantAfter)}  got ${JSON.stringify(r.a)}`);
  if (!r.bOk) console.log(`             BEFORE want ${JSON.stringify(r.wantBefore)}  got ${JSON.stringify(r.b)}`);
  if (r.aOk && r.bOk && r.isRepair) console.log(`             was: ${W(r.b)}   now: ${W(r.a)}`);
}
const fixes = rows.filter(r => r.isRepair).length;
const nets2 = rows.filter(r => !r.isRepair).length;
console.log(`\n  ${pass} passed, ${fail} failed  —  ${fixes} behaviours REPAIRED, ${nets2} proven UNCHANGED\n`);
process.exit(fail ? 1 : 0);
