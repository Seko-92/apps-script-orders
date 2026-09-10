/**
 * test-product-health.js — the Product Data Health verdicts, loaded from the REAL
 * ProductHealth.js so the tests cannot drift from what ships.
 *
 * WHY THIS EXISTS. This sheet decides three different things at once — what eBay has wrong,
 * what data is missing, and what is Amazon-ready. Each has a different consequence, and the
 * one that matters most is the smallest:
 *
 *   ⚠⚠ eBay LIGHTER than the true shipped weight = UNDER-CHARGING SHIPPING TODAY.
 *      Measured 2026-09-09: 49 live listings. That population must never be diluted into
 *      "mismatch" — section A exists to pin exactly that, in both directions.
 *
 * ⚠ THE DIMENSION COMPARISON SORTS BEFORE COMPARING. "7x5x1" and "1x5x7" are the same box;
 *   comparing them positionally would manufacture hundreds of false mismatches (section C).
 *
 * ⚠ "Wright(Oz)" is a LIVE TYPO in the truth file. If the reader ever matched an exact
 *   header instead of "contains oz", every ounce value would silently vanish and half the
 *   catalogue would read as "needs weight" (section B).
 *
 * PROVE BY MUTATION (each should turn the run red):
 *   a) drop the `d < 0` test so light/heavy collapse into one band  → A2/A3 fail
 *   b) compare dims without sorting                                 → C2 fails
 *   c) treat missing weight as 0 instead of null                    → B fails
 *   d) sort INCOMPLETE by SKU instead of units sold                  → E2 fails
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log((ok ? '  ✓ ' : '  ✗ ') + label +
    (ok ? '' : '  → got ' + JSON.stringify(got) + ', want ' + JSON.stringify(want)));
};
const section = (name, fn) => {
  console.log('\n' + name);
  try { fn(); } catch (e) { fail++; console.log('  ✗ SECTION THREW (soft): ' + e.message); }
};

const sandbox = { console: { log: () => {} } };
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(SRC, 'ProductHealth.js'), 'utf8'), sandbox);
const { _phClassify, _phParseDims, _phTruthOz, _phNormSku, _phBuildRows, PRODUCT_HEALTH } = sandbox;
const B = PRODUCT_HEALTH.bands;

/* A complete, Amazon-ready MI record — each test perturbs one field. */
const MI = (o = {}) => Object.assign(
  { title: 'Head Gasket', oz: 8, L: 7, W: 5, D: 1, photos: 3, sold: 40, onHand: 20, active: true }, o);
const TR = (o = {}) => Object.assign({ pt: 'Cylinder Head Gasket', title: '', lb: '', oz: 8, dim: '7x5x1' }, o);

section('A · UNDER-CHARGING IS ITS OWN BAND  ⚠ the only one costing money', () => {
  t('A1 agreement → READY, no fix', (() => {
    const v = _phClassify(TR(), MI());
    return [v.band, v.ebayFix, v.amazon];
  })(), [B.READY, '', 'READY']);

  t('A2 eBay LIGHTER than truth → UNDER-CHARGING', (() => {
    const v = _phClassify(TR({ oz: 16 }), MI({ oz: 8 }));   // truth 16, eBay 8
    return [v.band, v.ebayFix];
  })(), [B.UNDER, 'eBay LIGHT by 8oz']);

  t('A3 eBay HEAVIER → plain mismatch, NOT the money band', (() => {
    const v = _phClassify(TR({ oz: 8 }), MI({ oz: 24 }));
    return [v.band, v.ebayFix];
  })(), [B.MISMATCH, 'eBay heavy by 16oz']);

  t('A4 a half-ounce difference is rounding, not a discrepancy',
    _phClassify(TR({ oz: 8 }), MI({ oz: 8.4 })).band, B.READY);
});

section('B · WEIGHT PARSING — lb OR oz, and the live "Wright(Oz)" typo', () => {
  t('B1 pounds convert to ounces', _phTruthOz(12, ''), 192);
  t('B2 ounces pass through', _phTruthOz('', 4), 4);
  t('B3 pounds win when both present', _phTruthOz(1, 4), 16);
  t('B4 ⚠ NO weight is null, never 0 — "we do not know" ≠ "weighs nothing"',
    _phTruthOz('', ''), null);
  t('B5 a row with no weight reports the gap and blocks Amazon', (() => {
    const v = _phClassify(TR({ oz: '' }), MI());
    return [v.band, v.gaps, v.amazon];
  })(), [B.INCOMPLETE, ['weight'], 'needs weight/dims']);
});

section('C · DIMENSIONS — sorted before comparing  ⚠ a box has no order', () => {
  t('C1 parses "7x5x1" sorted', _phParseDims('7x5x1'), [1, 5, 7]);
  t('C2 ⚠ same box, different order → NOT a mismatch',
    _phClassify(TR({ dim: '7x5x1' }), MI({ L: 1, W: 5, D: 7 })).ebayFix, '');
  t('C3 a genuinely different box IS a mismatch', (() => {
    const v = _phClassify(TR({ dim: '7x5x1' }), MI({ L: 12, W: 9, D: 4 }));
    return [v.band, v.ebayFix];
  })(), [B.MISMATCH, 'dims differ']);
  t('C4 tolerates separators and spacing', _phParseDims(' 7 X 5 x 1 '), [1, 5, 7]);
  t('C5 two numbers is not a box', _phParseDims('7x5'), null);
});

section('D · AMAZON READINESS — the gate order is the funnel order', () => {
  t('D1 inactive listing is never ready',
    _phClassify(TR(), MI({ active: false })).amazon, 'not Active on eBay');
  t('D2 missing dims blocks before photos',
    _phClassify(TR({ dim: '' }), MI({ photos: 0 })).amazon, 'needs weight/dims');
  t('D3 one photo (the logo) is not enough',
    _phClassify(TR(), MI({ photos: 1 })).amazon, 'needs photos');
  t('D4 never sold → not a launch candidate',
    _phClassify(TR(), MI({ sold: 0 })).amazon, 'no sales history');
  t('D5 thin stock cannot support a launch',
    _phClassify(TR(), MI({ onHand: 2 })).amazon, 'thin stock');
  t('D6 ⚠ a truth-only SKU is an OPPORTUNITY, not a fault', (() => {
    const v = _phClassify(TR(), null);
    return [v.band, v.amazon];
  })(), [B.NOT_LISTED, 'not listed on eBay']);
});

section('E · THE JOIN AND THE SORT', () => {
  const truth = {
    '100': TR({ oz: 16 }),                       // under-charging
    '200': TR(),                                 // ready
    '300': TR({ oz: '', dim: '' }),              // incomplete, sells a lot
    '400': TR({ oz: '', dim: '' }),              // incomplete, sells little
    '900': TR()                                  // not on eBay
  };
  const mi = {
    '100': MI({ oz: 8,  sold: 5 }),
    '200': MI({ sold: 50 }),
    // ⚠ SOLD IS DELIBERATELY INVERTED AGAINST SKU ORDER. With 300=900/400=3 both
    //   "sold desc" and "sku asc" yield [300,400], so E2 passed against a mutation that
    //   removed the sold ranking entirely. A fixture where the two rules agree tests
    //   nothing. (Found by the mutation run, 2026-09-09.)
    '300': MI({ oz: '', L: '', W: '', D: '', sold: 3 }),
    '400': MI({ oz: '', L: '', W: '', D: '', sold: 900 }),
    '500': MI({ active: false, sold: 0 })         // ended, no truth → skipped
  };
  const out = _phBuildRows(truth, mi, 'T');
  const col = n => out.rows.map(r => r[PRODUCT_HEALTH.cols[n] - 1]);

  t('E1 ⚠ UNDER-CHARGING sorts to the very top', col('SKU')[0], '100');
  t('E2 ⚠ INCOMPLETE is ranked by UNITS SOLD, so the worklist is actionable',
    col('SKU').filter(s => ['300', '400'].includes(s)), ['400', '300']);
  t('E3 the truth-only SKU is included, banded last',
    col('SKU')[col('SKU').length - 1], '900');
  t('E4 ⚠ an ended listing with no truth record is dropped, not reported as a gap',
    col('SKU').includes('500'), false);
  t('E5 counts add up', (() => {
    const c = out.counts;
    return c.under + c.mismatch + c.incomplete + c.ready + c.notListed === c.total;
  })(), true);
  t('E6 counts are right', [out.counts.under, out.counts.ready, out.counts.incomplete, out.counts.notListed],
    [1, 1, 2, 1]);
});

section('F · SKU normalisation — the float trap that silently breaks the join', () => {
  t('F1 "161361.0" → "161361"', _phNormSku('161361.0'), '161361');
  t('F2 a real number → integer string', _phNormSku(161361), '161361');
  t('F3 a genuine decimal is left alone', _phNormSku('12.5'), '12.5');
  t('F4 blank stays blank', _phNormSku(''), '');
});

section('G · the sheet never writes to its sources', () => {
  const code = fs.readFileSync(path.join(SRC, 'ProductHealth.js'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
  t('G1 no setValue(s) against the truth spreadsheet handle',
    /BY_PART_TYPE_ID[\s\S]{0,400}setValues?\(/.test(code), false);
  t('G2 it does read the truth file', /openById\(BY_PART_TYPE_ID\)/.test(code), true);
  t('G3 ⚠ MI is resolved through MiSchema, never positionally',
    /MiSchema\.readColumns/.test(code), true);
});

console.log('\n' + (fail === 0 ? '✅' : '❌') +
  ' test-product-health: ' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail === 0 ? 0 : 1);
