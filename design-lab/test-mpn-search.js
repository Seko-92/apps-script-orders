// test-mpn-search.js — MPN Finder (MpnSearch.js), against the REAL file.
//
//   node test-mpn-search.js                     synthetic cases
//   MI_JSON=/path/mi.json node test-mpn-search.js   + a pass over a real MI export
//                                                  ({headers, rows} of sku/title/status + C: part-number columns)
// SRC=/other/dir overrides where MpnSearch.js is read from (before/after proofs).
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = process.env.SRC || path.join(__dirname, '..');
const ctx = { console, Map, Date };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(SRC, 'MpnSearch.js'), 'utf8'), ctx);

let pass = 0, fail = 0;
function eq(name, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) { pass++; } else { fail++; console.log('✗ ' + name + '\n    got  ' + g + '\n    want ' + w); }
}
const keys = q => ctx._mpnParseQuery(q).map(x => x.key);

// --- A · the match key ---------------------------------------------------------------
eq('A1 leading zero dropped', ctx._mpnKey('02102238'), '2102238');
eq('A2 stored-as-number reads the same', ctx._mpnKey(ctx._mpnCellText(2102238)), '2102238');
eq('A3 .0 float text', ctx._mpnCellText('4201560.0'), '4201560');
eq('A4 dashes/spaces ignored', ctx._mpnKey('1A033- 03043'), ctx._mpnKey('1A033-03043'));
eq('A5 case', ctx._mpnKey('1c010-74110'), '1C01074110');
eq('A6 all zeros keeps something', ctx._mpnKey('000000'), '000000');

// --- B · parsing what gets pasted ----------------------------------------------------
eq('B1 SERPIC spacing is one number', keys('0415 7075'), ['4157075']);
eq('B2 several SERPIC numbers', keys('0415 7075\n0416 1234'), ['4157075', '4161234']);
eq('B3 SERPIC row with position, description, qty',
   keys('1   0415 7075   Gasket   2\n2   0429 2547   Seal ring   1'), ['4157075', '4292547']);
eq('B4 comma list, mixed brands', keys('1C010-74110, 16851-22012; 02102238'),
   ['1C01074110', '1685122012', '2102238']);
eq('B5 dedupe incl. zero/no-zero', keys('02102238 2102238'), ['2102238']);
eq('B6 lone short number still searched', keys('821'), ['821']);
eq('B7 short noise dropped beside real numbers', keys('1 2 04157075 x2'), ['4157075']);
eq('B8 empty', keys('  \n '), []);

// --- C · cell splitting ---------------------------------------------------------------
eq('C1 comma cell', ctx._mpnSplitCell('027-03824, 366-08126'), ['027-03824', '366-08126']);
eq('C2 space-separated numbers + joined copy',
   ctx._mpnSplitCell('16851-22012 16851-22015'), ['16851-22012', '16851-22015', '16851-22012 16851-22015']);
eq('C3 Deutz pair joined', ctx._mpnSplitCell('0415 7075'), ['04157075']);
eq('C4 numeric cell', ctx._mpnSplitCell(4201560), ['4201560']);
eq('C5 blank', ctx._mpnSplitCell(''), []);

// --- D · index + answer ---------------------------------------------------------------
const rows = [
  ['111111', '02102238, 04157075'],            // main + extra
  ['222222', 4157075],                          // number, zero lost
  ['333333', '1C010-74110'],
  ['444444', '']
];
const cols = [{ name: 'C:MPN', off: 1 }, { name: 'C:Interchange Part Number', off: 2 }];
rows[3][2] = '04157075';                         // only in a one-off column
const index = ctx._mpnBuildIndex(rows, cols);
const describe = i => ({ sku: rows[i][0], active: i !== 1, available: 1 });
const ans = ctx._mpnAnswer(ctx._mpnParseQuery('0415 7075, 02102238, 9999999'), index, describe);
eq('D1 three answers', ans.length, 3);
eq('D2 0415 7075 hits all three listings', ans[0].matches.map(m => m.sku).sort(), ['111111', '222222', '444444']);
eq('D3 via labels', ans[0].matches.map(m => m.sku + ':' + m.via).sort(),
   ['111111:extra MPN', '222222:main', '444444:Interchange Part Number']);
eq('D4 inactive sorted last', ans[0].matches[ans[0].matches.length - 1].sku, '222222');
eq('D5 main match', ans[1].matches.map(m => m.sku + ':' + m.via), ['111111:main']);
eq('D6 miss', ans[2].matches.length, 0);

// --- F · the part before the kits that contain it ---------------------------------
const kidx = ctx._mpnBuildIndex([['K1', '04270701'], ['P1', '04270701'], ['K2', '04270701'], ['P0', '04270701']],
                                [{ name: 'C:MPN', off: 1 }]);
const kdesc = { 0: { sku: 'K1', isKit: true,  active: true, available: 9 },
                1: { sku: 'P1', isKit: false, active: true, available: 0 },
                2: { sku: 'K2', isKit: true,  active: true, available: 1 },
                3: { sku: 'P0', isKit: false, active: true, available: 203 } };
const kans = ctx._mpnAnswer(ctx._mpnParseQuery('0427 0701'), kidx, i => Object.assign({}, kdesc[i]));
eq('F1 parts first (stocked first), then kits', kans[0].matches.map(m => m.sku), ['P0', 'P1', 'K1', 'K2']);
eq('F2 SKU key agrees across number/text', ctx._mpnSkuKey(157554), ctx._mpnSkuKey('157554.0'));

// --- G · keyword search -------------------------------------------------------------
eq('G1 plural folds', ['gaskets', 'valves', 'glass', 'bus'].map(w => ctx._kwNorm(w)), ['gasket', 'valve', 'glass', 'bus']);
eq('G2 query words deduped, 1-char dropped', ctx._kwParseQuery('Head  gasket, head x'), ['head', 'gasket']);
const toks = ctx._kwTokens(['Cylinder Head Gasket for Deutz F3L912', 'BF6 M1013, 6 Cylinder']);
eq('G3 prefix', ctx._kwHit('gask', toks), true);
eq('G4 split model code glued', ctx._kwHit('bf6m1013', toks), true);
eq('G5 digits match the end of a model code', ctx._kwHit('912', toks), true);
eq('G6 digits do not match inside a plain number', ctx._kwHit('013', ctx._kwTokens(['part 4013'])), false);
eq('G7 absent word', ctx._kwHit('piston', toks), false);
const docs = [
  { title: 'Piston With Ring STD For Kubota V2203', other: ['V2203', 'Piston'] },
  { title: 'Engine Overhaul Kit', other: ['V2203, V2003', 'Overhaul Kit', 'piston rings included'] },
  { title: 'Piston For Deutz 912', other: ['912', 'Piston'] },
  { title: 'Head Gasket For Kubota V2203', other: ['V2203'] }
];
const km = ctx._kwMatch(ctx._kwParseQuery('v2203 pistons'), docs);
eq('G8 every word must appear (AND)', km.map(m => m.row).sort(), [0, 1]);
eq('G9 title match outranks spec-only', km.sort((a, b) => b.score - a.score).map(m => m.row), [0, 1]);
eq('G10 nothing matches', ctx._kwMatch(['crankshaft'], docs).length, 0);

// --- E · real MI export ---------------------------------------------------------------
if (process.env.MI_JSON) {
  const mi = JSON.parse(fs.readFileSync(process.env.MI_JSON, 'utf8'));
  const H = mi.headers, st = H.indexOf('listingStatus'), sk = H.indexOf('sku');
  const mcols = H.map((h, i) => ({ name: h, off: i })).filter(c => ctx.MPN_SEARCH.colPattern.test(c.name));
  const t0 = Date.now();
  const idx = ctx._mpnBuildIndex(mi.rows, mcols);
  console.log(`  real MI: ${mi.rows.length} rows · ${mcols.length} part-number columns · ` +
              `${idx.size} distinct keys · index built in ${Date.now() - t0}ms`);
  const multi = mi.rows.filter(r => r[st] === 'Active' && ctx._mpnSplitCell(r[H.indexOf('C:MPN')]).length > 1).length;
  console.log(`  active listings whose extra MPNs eBay search misses: ${multi}`);
  // every number in every cell must find its own row again
  let lost = 0;
  mi.rows.forEach((r, i) => mcols.forEach(c => ctx._mpnSplitCell(r[c.off]).forEach(n => {
    const k = ctx._mpnKey(n);
    if (k.length >= 3 && !(idx.get(k) || []).some(h => h.row === i)) lost++;
  })));
  eq('E1 every stored number finds its own listing', lost, 0);
  const desc = i => ({ sku: mi.rows[i][sk], active: mi.rows[i][st] === 'Active', available: 1 });
  const sample = ctx._mpnAnswer(ctx._mpnParseQuery('02102238\n0429 2547'), idx, desc);
  sample.forEach(x => console.log(`  sample ${x.query} → ${x.matches.map(m => m.sku + ' (' + m.via + ')').join(', ') || 'none'}`));
}

if (process.env.MI_ALL_JSON) {
  const mi = JSON.parse(fs.readFileSync(process.env.MI_ALL_JSON, 'utf8'));
  const H = mi.headers, ti = H.indexOf('title'), sk = H.indexOf('sku');
  const fi = ctx.KW_SEARCH.fields.map(f => H.indexOf(f)).filter(i => i >= 0);
  const docs2 = mi.rows.map(r => ({ title: r[ti], other: fi.map(i => r[i]) }));
  ['deutz 912 head gasket', 'v2203 piston', 'bf4m1011 crankshaft', 'perkins 404 gasket set',
   'yanmar 4tnv98 water pump', 'lister'].forEach(q => {
    const t0 = Date.now(), hits = ctx._kwMatch(ctx._kwParseQuery(q), docs2).sort((a, b) => b.score - a.score);
    console.log(`  "${q}" → ${hits.length} listings in ${Date.now() - t0}ms` +
      (hits[0] ? ` · top: ${mi.rows[hits[0].row][sk]} ${String(mi.rows[hits[0].row][ti]).slice(0, 55)}` : ''));
  });
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
