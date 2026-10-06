// Engine-variant columns ("A:D1703-BB-EC-1  B:…") read from a scanned header.
"use strict";

/**
 * A SERIAL-NUMBER RANGE in the header ("A:<=15000  B:14000 to 15000") is not an engine variant —
 * the page offered buttons named "<=15000" (user test 2026-10-06, ~15 books). A real variant names
 * a model ("D1703-BB-EC-1") or is unreadable text ("XXXX"); a range is digits, symbols and "to".
 */
function isSerialRange(name) {
  const n = String(name || "");
  return /\d/.test(n) && !/[A-Z]{1,2}\d{3}/i.test(n) && !/[A-Z]/i.test(n.replace(/\bto\b/gi, ""));
}

/** Drop serial-range "variants" from a parsed book, and their quantity columns with them. */
function dropSerialRangeModels(parsed) {
  if (!Array.isArray(parsed.models) || !parsed.models.some(m => isSerialRange(m.name))) return false;
  const keep = parsed.models.map((m, i) => (isSerialRange(m.name) ? -1 : i)).filter(i => i >= 0);
  const cols = keep.length ? keep : [0];                 // nothing real left: one quantity column
  parsed.models = keep.map(i => parsed.models[i]);
  parsed.sections.forEach(s => s.parts.forEach(p => { if (Array.isArray(p.qty)) p.qty = cols.map(i => (p.qty[i] == null ? null : p.qty[i])); }));
  return true;
}

module.exports = { isSerialRange, dropSerialRangeModels };
