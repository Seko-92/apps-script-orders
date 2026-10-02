// test-weight.js — the shipping-weight calculator (Weight.js), loaded REAL with Apps Script
// stubbed: list parsing, weight sources and their order, kit handling, the message. (2026-10-02)
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) pass++; else { fail++; console.log('✗ ' + name + '\n    got  ' + g + '\n    want ' + w); }
}
// Product Data Health rows: SKU, BRAND, TYPE, TITLE, TRUTH oz, EBAY oz, Δ, TRUTH dims, EBAY dims
const PDH = [
  ['166500', 'Deutz', 'Piston', 'Piston With Rings', 20, 16, -4, '10x8x6', '9x8x6'],
  ['173817', 'Kubota', 'Gasket', 'Gasket set', '', 6, '', '', '12x8x2'],
  ['000000', '', '', 'Special package', '', '', '', '', ''],
];
const MI = [['sku', 'title', 'packageWeightOz', 'packageLengthIn', 'packageWidthIn', 'packageDepthIn'],
            ['171432', 'Air Filter Sleeve', 32, 14, 10, 8],
            ['999999', 'Mystery', '', '', '', '']];
const ROWS = [];   // All Orders rows for _findOrderRows
const sheet = rows => ({ getLastRow: () => rows.length + 1, getRange: (r, c, n, w) => ({ getValues: () => rows.slice(r - 2, r - 2 + n).map(x => x.slice(c - 1, c - 1 + w)) }) });
const ctx = {
  console, Math, JSON, String, Number, Array, Object, RegExp, isFinite, parseFloat,
  SPREADSHEET_ID: 'x', DB_SHEET_NAME: 'Master Inventory', DB_SKU_HEADER: 'sku', DB_TITLE_HEADER: 'title',
  PRODUCT_HEALTH: { sheetName: 'Product Data Health', dataStartRow: 2, cols: { SKU: 1, TITLE: 4, TRUTH_WT: 5, EBAY_WT: 6, TRUTH_DIMS: 8, EBAY_DIMS: 9 } },
  Schema: { status: { CANCELED: 'CANCELED' } },
  SpreadsheetApp: { openById: () => ({ getSheetByName: n => n === 'Product Data Health' ? sheet(PDH) : { mi: true } }) },
  MiSchema: { readColumns: () => { const h = MI[0], idx = {}; ['sku', 'title', 'packageWeightOz', 'packageLengthIn', 'packageWidthIn', 'packageDepthIn'].forEach(n => idx[n] = h.indexOf(n)); return { rows: MI.slice(1), idx }; } },
  _findOrderRows: () => ROWS.slice(),
  kitComponentTag: n => { const m = String(n || '').match(/^↳ (?:from|added to) KIT-(\S+)/); return m ? m[1] : ''; },
  computeZohoSoDiff: q => (q === 'SO-9' ? { ok: true, soNumber: 'SO-9', lines: [{ sku: '166500', zohoQty: 1, status: 'new', name: 'Piston' }, { sku: '173817', zohoQty: 2, status: 'removed' }] } : { ok: false })
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(SRC, 'Weight.js'), 'utf8'), ctx);

const P = t => ctx._wtParseList(t).items.map(x => x.sku + ':' + x.qty);
eq('L1 commas + x2', P('166500 x2, 173817, 171432 x3'), ['166500:2', '173817:1', '171432:3']);
eq('L2 glued, before, spaced', P('166500x2 3x173817 171432 x 4'), ['166500:2', '173817:3', '171432:4']);
eq('L3 "2 x SKU" and "SKU 4"', P('2 x 166500\n171432 4'), ['166500:2', '171432:4']);
eq('L4 repeats add up', P('166500, 166500 x2'), ['166500:3']);
eq('L5 SUP- and 000000 are SKUs', P('SUP-101 x5, 000000'), ['SUP-101:5', '000000:1']);
eq('L6 junk is reported, not guessed', ctx._wtParseList('166500 hello').bad, ['hello']);
eq('O1 order ids', ['SO-26018', 'INV-022496', '24-15008-33107', 'AMZ-111-2', '166500'].map(ctx._wtLooksLikeOrder), [true, true, true, true, false]);
eq('F1 formats', [ctx._wtFmt(151), ctx._wtFmt(6), ctx._wtFmt(32), ctx._wtFmt(15.96)], ['9 lb 7 oz', '6 oz', '2 lb', '1 lb']);

let w = ctx.weighItems([{ sku: '166500', qty: 2 }, { sku: '173817', qty: 1 }, { sku: '171432', qty: 1 }, { sku: '999999', qty: 1 }]);
const L = w.result.lines;
eq('W1 measured weight wins over eBay\'s', [L[0].ozEach, L[0].src], [20, 'measured']);
eq('W2 no measured → eBay\'s (Product Data Health)', [L[1].ozEach, L[1].src], [6, 'eBay']);
eq('W3 not in Product Data Health → Master Inventory', [L[2].ozEach, L[2].src, L[2].title], [32, 'eBay', 'Air Filter Sleeve']);
eq('W4 no weight anywhere → missing, NOT zero', [L[3].ozEach, w.result.missing], [null, ['999999']]);
eq('W5 total = 2×20 + 6 + 32', w.result.totalOz, 78);
eq('W6 largest by volume', w.result.largest.dims, [14, 10, 8]);
eq('W7 message: total, the unweighed line, the eBay note', [/TOTAL ≈ 4 lb 14 oz/.test(w.text), /NOT in the total — no weight on file: 999999/.test(w.text), /\(eBay listing weight\)/.test(w.text), /parts only/.test(w.text)], [true, true, true, true]);

ROWS.push({ salesOrder: 'SO-26018', sku: '160029', qty: 1, note: '', status: 'PENDING' },
          { salesOrder: 'SO-26018', sku: '166500', qty: 4, note: '↳ from KIT-160029', status: 'PENDING' },
          { salesOrder: 'SO-26018', sku: '173817', qty: 1, note: '', status: 'CANCELED' },
          { salesOrder: 'SO-260180', sku: '171432', qty: 1, note: '', status: 'PENDING' });
w = ctx.weighOrder('SO-26018');
eq('K1 an expanded kit counts its parts, not the box too; CANCELED out; exact order id only',
   w.result.lines.map(l => l.sku + ':' + l.qty), ['166500:4']);
w = ctx.weighOrder('SO-9');
eq('Z1 not on the sheet → the Zoho SO\'s own lines (removed ones out)', [w.result.lines.map(l => l.sku), /from Zoho/.test(w.text)], [['166500'], true]);
eq('Z2 unknown order → says so', /not on the sheet or in Pending/.test(ctx.weighQueryText('SO-1')), true);
eq('T1 /weight list', /TOTAL ≈ 2 lb 8 oz/.test(ctx.weighQueryText('166500 x2')), true);
eq('T2 /weight with no args → usage', /Usage: \/weight/.test(ctx.weighQueryText('')), true);
eq('T3 nothing weighed at all', /No weight on file for any/.test(ctx.weighQueryText('999999')), true);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
