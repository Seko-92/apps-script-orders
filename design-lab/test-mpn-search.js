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

// --- H · ONE BOX: which search did it pick? ----------------------------------------
const mode = t => ctx._pfDetectMode(t);
eq('H1 a SKU', mode('163485'), 'sku');
eq('H2 a supplies SKU', mode('SUP-118'), 'sku');
eq('H3 a Deutz number (SERPIC spacing)', mode('0427 0701'), 'mpn');
eq('H4 the user\'s real paste', mode('0427 0701\n0417 9234\n0417 9921'), 'mpn');
eq('H5 SERPIC rows with pos/desc/qty', mode('1  0427 0701  Piston  1\n2  0417 3414  Piston pin  1'), 'mpn');
eq('H6 a Kubota dash number', mode('1G790-21050'), 'mpn');
eq('H7 several numbers on one line', mode('02102238, 1C010-74110'), 'mpn');
eq('H8 engine + part', mode('v2203 piston'), 'keywords');
eq('H9 engine code alone is a word', mode('4TNV98'), 'keywords');
eq('H10 perkins model with dash is a word', mode('404D-22 gasket'), 'keywords');
eq('H11 word + number on one line → keywords (AND)', mode('piston 04270701'), 'keywords');
eq('H12 a 7-digit legacy number', mode('4201560'), 'mpn');
eq('H13 leading-zero 6 digits is not a SKU', mode('041570'), 'mpn');
eq('I1 keyword digits ignore leading zeros', ctx._kwNorm('04157075'), ctx._kwNorm('4157075'));
eq('I2 keyword query joins SERPIC pairs', ctx._kwParseQuery('0415 7075 gasket'), ['4157075', 'gasket']);

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

// --- I · the part's identity (numbers + fit) -----------------------------------------
{
  const H = ['sku', 'C:Model Year', 'C:MPN', 'C:Model', 'C:Compatible Equipment Make',
             'C:Compatible Equipment Type', 'C:Engine Type', 'C:Additional Model',
             'C:Replace Part Number', 'C:Bobcat Models', 'C:Brand', 'C:More Part Number:'];
  const row = ['157554', 'M-16', '02102238, 0415 7075, 2102238', 'BF4M1011, 6 Cylinder', 'Deutz, Deutz-Fahr',
               'G940 Motor Grader, EC240CL Excavator', 'Diesel', 'S2600, bf4m1011',
               '1A033- 03043', '773, T190', 'HQ', '04292547'];
  const id = ctx._pfPartIdentity(H, row);
  eq('I1 main MPN first and flagged', id.numbers[0], { num: '02102238', main: true, via: '' });
  eq('I2 numbers: zero-variant deduped, SERPIC joined, trailing dash trimmed, labelled',
     id.numbers.map(n => n.num + '|' + n.via),
     ['02102238|', '04157075|', '1A033- 03043|Replace Part Number', '04292547|More Part Number']);
  eq('I3 engines: C:Model + extra model columns, cylinder/fuel dropped, case-deduped',
     id.engines, ['BF4M1011', 'S2600']);
  eq('I4 makes', id.brands, ['Deutz', 'Deutz-Fahr']);
  eq('I5 machines incl. Bobcat models', id.machines, ['G940 Motor Grader', 'EC240CL Excavator', '773', 'T190']);
  eq('I6 shelf (Model Year) and C:Brand never appear',
     JSON.stringify(id).includes('M-16') || JSON.stringify(id).includes('"HQ"'), false);
  eq('I7 fit kinds', ['C:Model Year', 'C:Model 2', 'C:Engine Info', 'C:Compatible Equipment Make', 'C:Bobcat Model', 'C:Part Type', 'title']
     .map(ctx._pfFitKind), [null, 'engines', 'engines', 'brands', 'machines', null, null]);
  eq('I8 empty row', ctx._pfPartIdentity(H, H.map(() => '')), { numbers: [], engines: [], brands: [], machines: [] });
  eq('I9 null-safe', ctx._pfPartIdentity(null, null).numbers, []);
}

// --- J · other sizes ------------------------------------------------------------------
{
  const cat = [
    ['166527', 'Piston With Rings STD For Kubota, 16423-21110, V2203 IDI, D1703, F2803, 87mm.'],
    ['173817', 'Piston With Rings 0.50 For Kubota, 16423-21910, V2203 IDI, D1703, F2803, 87mm.'],
    ['163332', 'Piston With Ring STD For Kubota, 1G796-21112, V2403-DI, V2403-MDI, V2203, 87mm.'],
    ['199095', 'Piston With Ring 0.50 For Kubota, 1G796-21110, V2403-DI, V2403-MDI, V2203, 87mm.'],
    ['216350', 'Piston With Rings STD For Kubota, 1J881-21110, V2403, V2203, 87mm.'],
    ['175617', 'Piston With Rings STD For Kubota 16641-21112, V2203 DI, D1803, 87mm.'],
    ['176724', 'Piston With Rings 0.50 For Kubota 16641-21912, V2203 DI, D1803, 87mm.'],
    ['155430', 'Piston rings STD For Kubota, 1G790-21050. V2203, V2403, V2203-M-DI For 1 Piston'],
    ['163413', 'Piston rings 0.50 For Kubota, 1G790-21053 V2203, V2203-M-DI, (For 1 Piston)'],
    ['163485', 'Piston With Ring STD For Deutz, 04179921, BF 1011.'],
    ['207069', 'Piston With Ring 0.50 For Deutz, 04270701, BF1011, 1011, 91.50mm.'],
    ['157572', 'Engine Overhaul, Rebuild Kit, Deutz STD, 04179914 F 4L1011F, 1011, 4 Cylinder.'],
    ['157590', 'Engine Overhaul, Rebuild Kit, For Deutz 0.50, 04179916, F 4L1011F, 1011.'],
    ['157599', 'Engine Overhaul, Rebuild Kit, Deutz 0.50, 04179916, F 3L1011F, 1011, 3 Cylinder'],
    ['195072', 'Main Bearing Set STD For Caterpillar, 156-6977, 3013C, C1.5, C1.7'],
    ['195090', 'Main Bearing Set 0.20 For Caterpillar, 161-2629, 3013C, C1.5, C1.7'],
    ['173214', 'Main Bearing Set STD For Caterpillar, 308-1854, C1.1'],
    ['175428', 'Main Bearing Set 0.20 For Caterpillar, 308-1854B, 294-4916B, C1.1'],
    ['172575', 'Main Bearing STD For Deutz 04231079, BF6L 913, F6L 914, TCD 914'],
    ['166122', 'Main Bearing 0.50 For Deutz 04231081, BF6L 913, F6L 914, TCD 914'],
    ['172508', 'Crankshaft Bushing STD For Kubota 1A091-23470, D1403, D1503, D1703'],
    ['158382', 'Main Bearing Set 0.50 For Kubota 1A091-23920, D1403, D1503, D1703, D1803, 60MM.'],
    ['171909', 'Fuel Injection Compensate Gasket for Deutz 04178523, 2011 Thickness 0.45mm'],
    ['172017', 'Fuel Injection Compensate Gasket for Deutz 04272924, 2011, Thickness 1.15mm'],
    ['199999', 'Water Pump For Kubota V2203, 1C010-73030'],
  ].map(([sku, title]) => ({ sku, title }));
  const lad = sku => ctx._szSiblings(cat.find(c => c.sku === sku).title, sku, cat).map(x => x.sku + ':' + x.size + (x.self ? '*' : ''));
  eq('J1 Kubota 16423 STD ↔ its own 0.50 only (not 1G796 / 16641 / 1J881)', lad('166527'), ['166527:STD*', '173817:0.50']);
  eq('J2 Kubota 1G796 family stays separate', lad('199095'), ['163332:STD', '199095:0.50*']);
  eq('J3 1J881 has no other size → no ladder', lad('216350'), []);
  eq('J4 rings ≠ pistons, rings pair with rings', lad('155430'), ['155430:STD*', '163413:0.50']);
  eq('J5 Deutz: different numbers, BF 1011 = BF1011, 91.50mm ignored', lad('163485'), ['163485:STD*', '207069:0.50']);
  eq('J6 kits: 4-cyl does not pair with 3-cyl', lad('157572'), ['157572:STD*', '157590:0.50']);
  eq('J7 Caterpillar C1.5/C1.7 ≠ C1.1', lad('195072'), ['195072:STD*', '195090:0.20']);
  eq('J8 Deutz main bearing STD ↔ 0.50', lad('172575'), ['172575:STD*', '166122:0.50']);
  eq('J9 bushing ≠ main bearing set', lad('172508'), []);
  eq('J10 mm thickness series, smallest first', lad('172017'), ['171909:0.45MM', '172017:1.15MM*']);
  eq('J11 no size in the title → nothing', lad('199999'), []);
  eq('J12 size parse', ['Piston STD For X', 'Rings Oversize 0.50 For', 'Gasket 1.25MM for', 'Bearing 0.100 For', 'Pump For V2203'].map(ctx._szSize),
     ['STD', '0.50', '1.25MM', '0.100', null]);
  const nums = sku => ctx._szSiblings(cat.find(c => c.sku === sku).title, sku, cat).map(x => x.num);
  eq('J14 card part number keeps the dash form (Kubota)', nums('166527'), ['16423-21110', '16423-21910']);
  eq('J15 card part number: plain 8-digit Deutz', nums('163485'), ['04179921', '04270701']);
  eq('J16 card part number: Caterpillar 3-digit dash', nums('195072'), ['156-6977', '161-2629']);
  eq('J13 ranking STD < 0.10 < 0.25 < 0.50', ['0.50', 'STD', '0.25', '0.10'].sort((a, b) => ctx._szRank(a) - ctx._szRank(b)), ['STD', '0.10', '0.25', '0.50']);
}

// --- K · findParts on the board (lite SKU mode) ----------------------------------------
{
  const seen = { dossier: 0, mpn: [] };
  ctx._mpnWho = () => '';
  ctx._buildPartDossier = () => { seen.dossier++; return { found: true }; };
  ctx.getPartBasics = q => ({ ok: true, basics: q === '157554' ? { found: true }
                                           : q === '300001' ? { found: false, zohoAvailable: 4 } : { found: false, zohoAvailable: null } });
  ctx.searchMpns = (t, o) => { seen.mpn.push(o.source); return { ok: true, results: [] }; };
  ctx.searchKeywords = (t, o) => { seen.mpn.push('kw:' + o.source); return { ok: true, matches: [], total: 0 }; };
  let r = ctx.findParts('157554', '', { source: 'board', lite: true });
  eq('K1 lite: a known SKU comes back as just the SKU, no dossier built', [r.mode, r.sku, seen.dossier], ['sku', '157554', 0]);
  r = ctx.findParts('300001', '', { source: 'board', lite: true });
  eq('K2 lite: a Zoho-only SKU still counts as ours', [r.mode, r.sku], ['sku', '300001']);
  r = ctx.findParts('412345', '', { source: 'board', lite: true });
  eq('K3 lite: not our SKU → searched as a part number, logged as board', [r.mode, /not one of our SKUs/.test(r.note), seen.mpn[0]], ['mpn', true, 'board']);
  r = ctx.findParts('v2203 piston', '', { source: 'board', lite: true });
  eq('K4 words → keyword search, logged as board', [r.mode, seen.mpn[1]], ['keywords', 'kw:board']);
  r = ctx.findParts('157554', '');
  eq('K5 sidebar path unchanged: full dossier, logged as console', [r.mode, !!r.dossier, seen.dossier], ['sku', true, 1]);
  ctx.findParts('04270701', '');
  eq('K6 sidebar default source is console', seen.mpn[2], 'console');
  eq('K7 empty text refused', ctx.findParts('  ', '', { lite: true }).ok, false);
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
