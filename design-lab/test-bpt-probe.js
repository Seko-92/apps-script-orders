/**
 * test-bpt-probe.js — the By Part Type column probe's pure helpers.
 *
 * The load-bearing assertion is that `_bptIsConsumed` MIRRORS `_phReadTruth`'s real
 * matching, including its LOOSE `indexOf("oz")` rule — which exists because
 * "Wright(Oz)" is a live typo in the source file. If the probe used a stricter
 * rule than production, it would report columns as unread that production already
 * reads, and the whole "what have we never looked at" answer would be wrong.
 *
 *   SRC=/path/to/older/ProductHealth.js node test-bpt-probe.js   # before/after
 */
const fs = require('fs'), vm = require('vm'), path = require('path');
const SRC = process.env.SRC || path.join(__dirname, '..', 'ProductHealth.js');
const src = fs.readFileSync(SRC, 'utf8');

function grab(sig) {
  const i = src.indexOf(sig);
  if (i < 0) throw new Error('not found in source: ' + sig);
  let d = 0, j = src.indexOf('{', i);
  do { if (src[j] === '{') d++; else if (src[j] === '}') d--; j++; } while (d > 0 && j < src.length);
  return src.slice(i, j);
}
const ctx = {}; vm.createContext(ctx);
vm.runInContext(src.match(/var BPT_PROBE = \{[\s\S]*?\};/)[0], ctx);
['function _bptClassify', 'function _bptIsConsumed', 'function _bptTrim', 'function _bptPad']
  .forEach(f => vm.runInContext(grab(f), ctx));

let pass = 0, fail = 0;
const t = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${label}${ok ? '' : `  got ${JSON.stringify(got)} want ${JSON.stringify(want)}`}`);
};

console.log('A · _bptIsConsumed mirrors _phReadTruth');
t("'SKU'",                      ctx._bptIsConsumed('SKU'), true);
t("'Part Type'",                ctx._bptIsConsumed('Part Type'), true);
t("'Brand'",                    ctx._bptIsConsumed('Brand'), true);
t("'Weight (lb)'",              ctx._bptIsConsumed('Weight (lb)'), true);
t("'Wright(Oz)' (live typo)",   ctx._bptIsConsumed('Wright(Oz)'), true);
t("'Weight (oz)'",              ctx._bptIsConsumed('Weight (oz)'), true);
t("'Dimension LxWxD' (prefix)", ctx._bptIsConsumed('Dimension LxWxD'), true);
t("'Compatible Brands' is NOT", ctx._bptIsConsumed('Compatible Brands'), false);
t("'Replacement Part Numbers'", ctx._bptIsConsumed('Replacement Part Numbers'), false);
t("'Model' is NOT",             ctx._bptIsConsumed('Model'), false);
t("blank is NOT",               ctx._bptIsConsumed(''), false);

console.log('B · _bptClassify buckets');
t('UPC -> IDENTIFIER',          ctx._bptClassify('UPC'), 'IDENTIFIER');
t('GTIN -> IDENTIFIER',         ctx._bptClassify('GTIN'), 'IDENTIFIER');
t('Part Number -> PART NUMBER', ctx._bptClassify('Part Number'), 'PART NUMBER');
t('Replacement -> PART NUMBER', ctx._bptClassify('Replacement Part Numbers'), 'PART NUMBER');
t('Compatible Brands -> FITMENT', ctx._bptClassify('Compatible Brands'), 'FITMENT');
t('Model -> FITMENT',           ctx._bptClassify('Model'), 'FITMENT');
t('Weight (lb) -> PHYSICAL',    ctx._bptClassify('Weight (lb)'), 'PHYSICAL');
t('Title -> DESCRIPTIVE',       ctx._bptClassify('Title'), 'DESCRIPTIVE');
t('Photo -> MEDIA',             ctx._bptClassify('Photo'), 'MEDIA');
t('Zoho -> OTHER-SYSTEM',       ctx._bptClassify('Zoho'), 'OTHER-SYSTEM');
t('Widget -> UNCLASSIFIED',     ctx._bptClassify('Widget'), 'UNCLASSIFIED');
t('blank -> ""',                ctx._bptClassify(''), '');

console.log('C · formatting helpers');
t('trim passthrough',           ctx._bptTrim('abc'), 'abc');
t('trim null -> ""',            ctx._bptTrim(null), '');
t('trim caps at sampleLen',     ctx._bptTrim('x'.repeat(60)).length, ctx.BPT_PROBE.sampleLen);
t('pad to width',               ctx._bptPad('hi', 6), 'hi    ');
t('pad never exceeds width',    ctx._bptPad('abcdefgh', 5).length, 5);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
