// ============================================================================
// kitComponentTag — the ONE parser for "is this a kit component, and whose?"
//
// Loaded from the REAL Helpers.js. Every case below is a shape that has already
// caused a live bug, or is one tag-read away from causing one:
//
//   · `↳ added to KIT-` (custom adds) was invisible to two callers in two days.
//     In _dashOpenOrders it dropped the part from the board's done/total, so a
//     kit could read "5 of 5" with a sixth part still on its shelf. In
//     _kitParentFollowUp it was worse — the unmatched row fell through to the
//     candidate-PARENT branch, so the parent could flip SHIPPED while a
//     custom-added component sat unpicked.
//   · A Zoho flag PREPENDS its warning as its own first line and CASCADES onto
//     a removed kit's components, so the tag lands on line 2.
//   · The capture must be compared EXACTLY — `indexOf("↳ from KIT-" + sku)`
//     lets kit 1586 match kit 158652's components.
//
// Usage: node test-kit-tag.js
// ============================================================================
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const sandbox = {
  SpreadsheetApp: {}, PropertiesService: {}, Utilities: {}, CacheService: {},
  SPREADSHEET_ID: 'x', MAIN_SHEET_NAME: 'All orders', Logger: { log() {} },
  Schema: { idx: () => 0, cols: {}, dataStartRow: 4, dataWidth: 10 },
  console: { log() {}, error() {} }, Date
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Helpers.js'), 'utf8'), sandbox);
const tag = sandbox.kitComponentTag;

let failed = 0;
function check(label, got, want) {
  const ok = got === want;
  console.log(`  ${ok ? '✓' : '✗'} ${label}` + (ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`));
  if (!ok) failed++;
}

console.log('\n' + '='.repeat(70));
console.log('  kitComponentTag — both shapes, flags, and exact capture');
console.log('='.repeat(70));

// ── the ordinary case ─────────────────────────────────────────────────────
check('registered component', tag('↳ from KIT-159012 · Miguel'), '159012');
check('bare tag, no human note', tag('↳ from KIT-159012'), '159012');

// ── CUSTOM ADDS — the miss that caused both live bugs ─────────────────────
check('CUSTOM ADD is a component', tag('↳ added to KIT-159012 · custom add · Miguel'), '159012');
check('custom add, bare', tag('↳ added to KIT-159012'), '159012');

// ── machine annotations must not confuse the capture ──────────────────────
check('swap annotation', tag('↳ from KIT-159012 · swapped 173763 → 173772 · Miguel'), '159012');
check('qty annotation', tag('↳ from KIT-159012 · qty 4→6 · Miguel'), '159012');
check('deploy annotation',
      tag('↳ from KIT-159012 · deploy 3 total (1 for customer + 2 for us) · Miguel'), '159012');
check('legacy ⚠ FORCED', tag('↳ from KIT-159012 · ⚠ FORCED · Miguel'), '159012');

// ── a Zoho flag prepends its own LINE and cascades onto components ────────
check('flagged component still belongs to its kit',
      tag('⚠️ REMOVED IN ZOHO 8/14\n↳ from KIT-159012 · Miguel'), '159012');
check('qty-flagged component too',
      tag('⚠️ ZOHO QTY: 4 → 2 8/14\n↳ added to KIT-159012'), '159012');
check('flag with nothing under it', tag('⚠️ REMOVED IN ZOHO 8/14'), '');

// ── NOT components ────────────────────────────────────────────────────────
check('a plain human note', tag('Miguel'), '');
check('a buyer note', tag('Buyer Note: please pack carefully'), '');
check('empty', tag(''), '');
check('null', tag(null), '');
check('undefined', tag(undefined), '');
check('a floor note', tag('** HOLD — do not ship'), '');
// The tag must LEAD. Text that merely mentions it is not a component row.
check('tag must lead, not merely appear',
      tag('see the note ↳ from KIT-159012 below'), '');

// ── EXACT capture — the prefix trap ───────────────────────────────────────
check('captures the whole SKU, not a prefix', tag('↳ from KIT-158652 · Miguel'), '158652');
check('a shorter kit does not swallow a longer one',
      tag('↳ from KIT-158652 · Miguel') === '1586', false);
check('captures up to whitespace only', tag('↳ from KIT-159012 extra'), '159012');

console.log('\n' + '='.repeat(70));
if (failed) { console.log(`  ${failed} FAILURE(S)`); process.exit(1); }
console.log('  ALL CLEAR');
