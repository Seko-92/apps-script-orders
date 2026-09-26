/**
 * _layout-stub.js — injects the REAL three-table layout helpers from Helpers.js into a
 * test sandbox (2026-09-26), so harnesses that stub getBoundaryRow() keep working now
 * that range writers ask getTableLayout() instead.
 *
 * The pure pieces (_layoutRanges, isStructuralRowNum, tableOfRow, _tableSegment,
 * directTableEnd) are EXTRACTED FROM THE SHIPPED SOURCE, never retyped — a copy in a
 * test is how a test ends up proving a rule the product no longer follows.
 *
 * getTableLayout itself reads the sheet, so it is replaced by one that scans the
 * sandbox's own column-A model (or a {direct, amazon, maxRows} function when given).
 */
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');

function extract(src, name) {
  const start = src.indexOf('function ' + name + '(');
  if (start === -1) throw new Error('_layout-stub: ' + name + ' not found in Helpers.js');
  let i = src.indexOf('{', start), depth = 0;
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
  }
  throw new Error('_layout-stub: unbalanced ' + name);
}

/**
 * @param {object} sandbox  a vm context object (must already carry Schema, or pass schemaSrc)
 * @param {function} layoutFn  () => ({direct, amazon, maxRows, lastRow})
 */
function injectLayout(sandbox, layoutFn, root) {
  root = root || process.env.SRC || path.join(__dirname, '..');
  const src = fs.readFileSync(path.join(root, 'Helpers.js'), 'utf8');
  const names = ['_layoutRanges', 'isStructuralRowNum', 'tableOfRow', '_tableSegment', 'directTableEnd'];
  const code = names.filter(n => src.indexOf('function ' + n + '(') !== -1)
                    .map(n => extract(src, n)).join('\n');
  vm.runInContext(code, sandbox, { filename: 'Helpers.js (layout)' });
  sandbox.getTableLayout = function () {
    const b = layoutFn() || {};
    const out = { direct: b.direct == null ? -1 : b.direct,
                  amazon: b.amazon == null ? -1 : b.amazon,
                  lastRow: b.lastRow || 0, maxRows: b.maxRows || 0, structural: {} };
    return sandbox._layoutRanges(out);
  };
}

/** Layout from a column-A array (index 0 = row 1). */
function layoutFromColA(rows, schema) {
  const up = v => String(v == null ? '' : v).trim().toUpperCase();
  let direct = -1, amazon = -1, last = 0;
  rows.forEach((v, i) => {
    if (up(v) === 'DIRECT' && direct === -1) direct = i + 1;
    else if (up(v) === 'AMAZON' && amazon === -1) amazon = i + 1;
    if (up(v)) last = i + 1;
  });
  if (amazon !== -1 && direct !== -1 && amazon < direct) amazon = -1;
  return { direct, amazon, lastRow: last, maxRows: rows.length };
}

module.exports = { injectLayout, layoutFromColA, extract };
