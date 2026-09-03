// ============================================================================
// test-kit-note.js — the kit component NOTE: composition, order, and the board
//
// WHY THIS FILE EXISTS (2026-09-03)
//   The picker used to expand a kit for a STOCK BUILD and get
//   "· deploy 3 total (1 for customer + 2 for us)" on every inserted row — on
//   an order with no customer. They deleted it by hand, row by row, after the
//   fact. The note is now EDITABLE in the modal before the rows exist.
//
//   That makes one segment of the note picker-controlled, and two things then
//   have to hold or the change is worse than the bug:
//
//   1. ⚠ THE TAG STAYS LOCKED. kitComponentTag() reads it, and this codebase
//      has TWICE shipped a live bug from that tag failing to match — a
//      custom-added part vanishing from the board's done/total, and
//      _kitParentFollowUp flipping a parent SHIPPED with a component open.
//
//   2. ⚠⚠ THE EDITED TEXT MUST COME AFTER THE MACHINE SEGMENTS. FloorBoard
//      walks " · " segments left to right, drops the machine ones, and STOPS
//      at the first it does not recognise. Editable text in the old leading
//      position would push "· swapped X → Y" and "· qty a→b" onto the tablet
//      as though a person wrote them.
//
// NOTHING HERE IS RE-TYPED. The composition is lifted out of the real
// KitExpansion.js and evaluated; the stripper is lifted out of the real
// FloorBoard.html and run. A re-typed copy would pass while the shipped code
// drifted — which is the whole failure mode this file is about.
//
// Run:  node test-kit-note.js
// HEAD: SRC=/tmp/head node test-kit-note.js
// ============================================================================
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

const SRC = process.env.SRC || path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');

let failed = 0, passed = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`  ${ok ? '✓' : '✗'} ${label}` +
    (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  ok ? passed++ : failed++;
};
const section = n => console.log('\n' + n);

// ── LIFT THE REAL COMPOSITION OUT OF KitExpansion.js ────────────────────────
const KE = read('KitExpansion.js');

function slice(startAnchor, endAnchor) {
  const i = KE.indexOf(startAnchor);
  if (i === -1) throw new Error('anchor not found in KitExpansion.js: ' + startAnchor);
  const j = KE.indexOf(endAnchor, i);
  if (j === -1) throw new Error('end anchor not found: ' + endAnchor);
  return KE.slice(i, j + endAnchor.length);
}

// ⚠ FAIL SOFT, NEVER THROW. Run against HEAD (SRC=/tmp/head) these anchors do not
//   exist yet, and a thrown error would abort the file before a single section
//   reported — which is exactly what makes a before/after proof useless. Missing
//   composition becomes a sentinel that fails every assertion loudly instead.
let partA = '', partB = '', ABSENT = '';
try {
  // (a) tag + default note text + the picker's override
  partA = slice('var tagFrom  = "↳ from KIT-"', 'KIT_NOTE_TEXT_MAX);\n  }');
  // (b) the per-row assembly
  partB = slice('var noteBase = comp.added ?', '(rowNote  ? " · " + rowNote  : "");');
} catch (e) {
  ABSENT = '<EDITABLE-NOTE COMPOSITION NOT PRESENT: ' + e.message + '>';
}

const NOTE_MAX = parseInt((KE.match(/var KIT_NOTE_TEXT_MAX\s*=\s*(\d+)/) || [])[1], 10) || 0;

/** Run the SHIPPED composition for one component. */
function compose(o) {
  if (ABSENT) return ABSENT;
  const sb = {
    rowSku: o.rowSku, extras: o.extras, totalKits: o.totalKits, rowQty: o.rowQty,
    alterations: o.alterations || null, comp: o.comp || {}, alter: o.alter || '',
    rowNote: o.rowNote || '', KIT_NOTE_TEXT_MAX: NOTE_MAX,
    String, parseInt, out: null
  };
  vm.createContext(sb);
  vm.runInContext(partA + '\n' + partB + '\nout = rowNoteFinal;', sb,
                  { filename: 'KitExpansion.js#note' });
  return sb.out;
}

// ── LIFT THE REAL STRIPPER OUT OF FloorBoard.html ───────────────────────────
const FB = read('FloorBoard.html');
const segStart = FB.indexOf('var KIT_MACHINE_SEG');
const segEnd   = FB.indexOf('}', FB.indexOf('function _isKitMachineSeg')) ;
const fnEnd    = FB.indexOf('\n    }', FB.indexOf('function _isKitMachineSeg')) + 6;
const stripSrc = FB.slice(segStart, fnEnd);
const sbFB = { console };
vm.createContext(sbFB);
vm.runInContext(stripSrc, sbFB, { filename: 'FloorBoard.html#seg' });

/** What the floor actually READS, per FloorBoard's own walk. */
function humanPart(note) {
  if (note.indexOf('↳') !== 0) return note;
  const segs = note.split(' · ');
  let s = 0;
  while (s < segs.length && sbFB._isKitMachineSeg(segs[s])) s++;
  return segs.slice(s).join(' · ').trim();
}

// ── the real parser, for the locked-tag assertions ──────────────────────────
const sbH = { SpreadsheetApp:{}, PropertiesService:{}, Utilities:{}, CacheService:{},
  SPREADSHEET_ID:'x', MAIN_SHEET_NAME:'All orders', Logger:{log(){}},
  Schema:{ idx:()=>0, cols:{}, dataStartRow:4, dataWidth:10 },
  console:{log(){},error(){}}, Date };
vm.createContext(sbH);
vm.runInContext(read('Helpers.js'), sbH);
const tag = sbH.kitComponentTag;

console.log('\n' + '='.repeat(74));
console.log('  kit component NOTE — composition, order, and what the floor reads');
console.log('='.repeat(74));

const BASE = { rowSku: '158679', rowQty: 1, extras: 0, totalKits: 1 };

// ============================================================================
section('A · THE DEFAULT IS UNCHANGED — leaving it alone must ship what shipped');
// ============================================================================
t('A1 no spares, no alterations, no row note',
  compose(BASE), '↳ from KIT-158679');

t('A2 ⭐ spares > 0 reads exactly as it always did',
  compose({ ...BASE, extras: 2, totalKits: 3 }),
  '↳ from KIT-158679 · deploy 3 total (1 for customer + 2 for us)');

t('A3 the parent row note still rides at the end',
  compose({ ...BASE, rowNote: 'Miguel' }),
  '↳ from KIT-158679 · Miguel');

t('A4 spares + a row note',
  compose({ ...BASE, extras: 2, totalKits: 3, rowNote: 'Miguel' }),
  '↳ from KIT-158679 · deploy 3 total (1 for customer + 2 for us) · Miguel');

// ============================================================================
section('B · THE PICKER EDIT — the whole point');
// ============================================================================
t('B1 ⭐ an edited note replaces the default sentence',
  compose({ ...BASE, extras: 2, totalKits: 3,
            alterations: { noteText: 'build 3 for stock' } }),
  '↳ from KIT-158679 · build 3 for stock');

t('B2 ⭐ clearing it leaves just the tag — the 2026-09-03 hand-deletion, automated',
  compose({ ...BASE, extras: 2, totalKits: 3, alterations: { noteText: '' } }),
  '↳ from KIT-158679');

t('B3 an edit survives alongside the parent row note',
  compose({ ...BASE, extras: 1, totalKits: 2, rowNote: 'call first',
            alterations: { noteText: 'build 2 for stock' } }),
  '↳ from KIT-158679 · build 2 for stock · call first');

t('B4 whitespace is trimmed',
  compose({ ...BASE, alterations: { noteText: '   padded   ' } }),
  '↳ from KIT-158679 · padded');

t('B5 ⚠ an over-long note is capped, not written whole',
  compose({ ...BASE, alterations: { noteText: 'x'.repeat(400) } }).length,
  ('↳ from KIT-158679 · ').length + NOTE_MAX);

t('B6 the Mini App path (no noteText) still gets the default',
  compose({ ...BASE, extras: 2, totalKits: 3, alterations: { overrides: {}, added: [] } }),
  '↳ from KIT-158679 · deploy 3 total (1 for customer + 2 for us)');

// ============================================================================
section('C · ⚠⚠ SEGMENT ORDER — machine first, human last');
// ============================================================================
const swapped = { ...BASE, extras: 2, totalKits: 3,
                  alter: ' · swapped 173763 → 173772', rowNote: 'call first' };

t('C1 ⭐ the alteration comes BEFORE the note text',
  compose(swapped),
  '↳ from KIT-158679 · swapped 173763 → 173772 ' +
  '· deploy 3 total (1 for customer + 2 for us) · call first');

t('C2 ⭐⭐ an EDITED note still leaves the swap in front of it',
  compose({ ...swapped, alterations: { noteText: 'build 3 for stock' } }),
  '↳ from KIT-158679 · swapped 173763 → 173772 ' +
  '· build 3 for stock · call first');

t('C3 ⭐ THE REASON: the floor reads only the human tail',
  humanPart(compose({ ...swapped, alterations: { noteText: 'build 3 for stock' } })),
  'build 3 for stock · call first');

t('C4 ⚠ with the OLD order the swap would have leaked onto the tablet',
  humanPart('↳ from KIT-158679 · build 3 for stock ' +
            '· swapped 173763 → 173772 · call first'),
  'build 3 for stock · swapped 173763 → 173772 · call first');

t('C5 unedited, the floor sees only the human note',
  humanPart(compose(swapped)), 'call first');

t('C6 unedited with no row note, the floor sees nothing at all',
  humanPart(compose({ ...BASE, extras: 2, totalKits: 3,
                      alter: ' · qty 2→4' })), '');

// ============================================================================
section('D · THE LOCKED TAG — every shape must still parse');
// ============================================================================
t('D1 default', tag(compose({ ...BASE, extras: 2, totalKits: 3 })), '158679');
t('D2 edited', tag(compose({ ...BASE, alterations: { noteText: 'build 3 for stock' } })), '158679');
t('D3 edited + swap + row note', tag(compose({ ...swapped,
    alterations: { noteText: 'anything at all' } })), '158679');
t('D4 note cleared', tag(compose({ ...BASE, alterations: { noteText: '' } })), '158679');

t('D5 ⭐ a custom add keeps ITS tag shape',
  compose({ ...BASE, comp: { added: true }, alter: ' · custom add' }),
  '↳ added to KIT-158679 · custom add');
t('D6 …and it parses too', tag(compose({ ...BASE, comp: { added: true },
    alter: ' · custom add' })), '158679');

t('D7 ⭐ a custom add NOW carries the note text too (it used to lose it)',
  compose({ ...BASE, comp: { added: true }, alter: ' · custom add',
            alterations: { noteText: 'build 3 for stock' } }),
  '↳ added to KIT-158679 · custom add · build 3 for stock');

// ============================================================================
section('E · THE SOURCE CONTRACT — the tag is never built from input');
// ============================================================================
if (ABSENT) console.log('  ⚠ ' + ABSENT);
const stripComments = s => s.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
const bodyA = stripComments(partA), bodyB = stripComments(partB);

t('E1 ⚠ the tag is a literal + the row SKU, never anything picker-supplied',
  /var tagFrom\s*=\s*"↳ from KIT-"\s*\+\s*rowSku;/.test(bodyA), true);
// ⚠ `alterations.` legitimately appears twice here (the null guard, then the read),
//   so counting occurrences accused correct code. Assert the INTENT instead.
const altRefs = bodyA.match(/alterations\.\w+/g) || [];
t('E2 ⚠ noteText is the ONLY field read out of alterations here',
  altRefs.length > 0 && altRefs.every(r => r === 'alterations.noteText'), true);
t('E3 ⚠ and it is capped',
  /slice\(0,\s*KIT_NOTE_TEXT_MAX\)/.test(bodyA), true);
t('E4 ⚠⚠ the assembly puts alter BEFORE noteText',
  bodyB.indexOf('alter') < bodyB.indexOf('noteText'), true);
// ⚠ `rowNoteFinal` CONTAINS the substring `rowNote`, so a bare indexOf found the
//   assignment target rather than the operand and accused correct code.
t('E5 …and noteText before the parent row note',
  bodyB.indexOf('noteText') < bodyB.indexOf('(rowNote'), true);

// ============================================================================
section('F · ⚠⚠ THE MODAL DEFAULT MUST EQUAL THE SERVER DEFAULT');
// ============================================================================
// The picker sees a pre-filled sentence and commits it verbatim. If the modal's
// wording and the server's wording ever diverge, the field shows one thing and the
// sheet gets another — a lie the picker has no way to catch. Both are lifted from
// the real files and compared across a matrix.
(function () {
  const MODAL = read('KitExpansionModal.html');
  const i = MODAL.indexOf('function noteDefaultFor');
  if (i === -1) {
    t('F0 the modal has a noteDefaultFor composer', false, true);
    return;
  }
  let depth = 0, started = false, end = -1;
  for (let j = i; j < MODAL.length; j++) {
    if (MODAL[j] === '{') { depth++; started = true; }
    else if (MODAL[j] === '}') { depth--; if (started && depth === 0) { end = j + 1; break; } }
  }
  const sbM = { parseInt, String };
  vm.createContext(sbM);
  vm.runInContext(MODAL.slice(i, end), sbM, { filename: 'KitExpansionModal.html#noteDefaultFor' });

  // Pull the server's default back out of a composed note (strip the locked tag).
  const serverDefault = (rowQty, extras) => {
    const whole = compose({ rowSku: 'X', rowQty, extras, totalKits: rowQty + extras });
    const tag = '↳ from KIT-X';
    return whole === tag ? '' : whole.slice((tag + ' · ').length);
  };

  let mismatches = [];
  for (let rowQty = 0; rowQty <= 4; rowQty++) {
    for (let extras = 0; extras <= 4; extras++) {
      const m = sbM.noteDefaultFor(rowQty, extras);
      const s = serverDefault(rowQty, extras);
      if (m !== s) mismatches.push(`rowQty=${rowQty} extras=${extras}: modal ${JSON.stringify(m)} vs server ${JSON.stringify(s)}`);
    }
  }
  t('F1 ⭐ modal and server agree on every rowQty × extras in 0..4', mismatches, []);

  t('F2 the modal caps at the same number the server slices to',
    (MODAL.match(/var NOTE_TEXT_MAX\s*=\s*(\d+)/) || [])[1], String(NOTE_MAX));

  t('F3 ⚠ the modal sends noteText inside alterations, not as an argument',
    /noteText:\s*currentNote/.test(MODAL) &&
    /commitKitFromModal\(SESSION_ID,\s*excluded,\s*currentExtras/.test(MODAL), true);

  t('F4 ⚠ the tag chip is a span, never an input',
    /class="note-tag"/.test(MODAL) &&
    !/<input[^>]*class="note-tag"/.test(MODAL), true);

  t('F5 the note resets per kit, like spares does',
    /noteEdited\s*=\s*false;[\s\S]{0,120}currentNote\s*=\s*noteDefaultFor/.test(MODAL), true);
})();

console.log('\n' + '='.repeat(74));
if (failed) { console.log(`❌ test-kit-note: ${passed} passed, ${failed} failed`); process.exit(1); }
console.log(`✅ test-kit-note: ${passed} passed, 0 failed`);
