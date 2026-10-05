// Repair a misread LETTER in a Kubota prefix: tesseract reads the C of "1C010-5675-0" as 0, giving
// "10010-5675-0" — a valid-looking number, so no shape check can see it. Measured 2026-10-05: the
// clean text manuals hold ONE real "10…" number in 1,321 (10244-4232-0), the OCR books held 468.
//
// A suspect number ("10xyz", or "16xyz"/"14xyz"… — see CONFUSE) is corrected only on EVIDENCE, strongest first:
//   1. "known"  — exactly one letter makes it a number from a clean manual or one of our listings
//   2. "read"   — exactly one letter makes it a number another OCR book READ with that letter
//   3. "stem"   — the same book uses exactly one lettered prefix with these digits ("1C010")
// Still ambiguous → left alone (it stays unconfirmed, so the page shows it as "check").
// A "10…" number that is itself known is real and never touched.
"use strict";

// second character read as a digit → the letters it is mistaken for. Real prefixes "16…", "14…",
// "15…" are common, so for every digit except 0 only FULL-NUMBER evidence (tiers 1–2) counts —
// the stem guess (tier 3) is for "10…" alone, which is essentially never a real number.
// (V3600 p25: "1C010-5602-3" read as "16010-5602-3", "1C020-…" as "16020-…")
const CONFUSE = { "0": "CDGQU", "O": "CDGQU", "6": "GC", "4": "A", "3": "E", "1": "J", "8": "B" };
const SUSPECT = /^1([0O64318])(\d{3})-/;

/**
 * @param {Array<{parts:Array<{pn:string,ocr?:object}>}>} books  every OCR book (parts of all sections)
 * @param {(pn:string)=>boolean} isKnown  in a clean manual or on one of our listings
 * @return {{fixed:number, byTier:{known:number,read:number,stem:number}, left:number}}
 */
function fixMisreadPrefixes(books, isKnown) {
  // numbers OCR read WITH a letter in the prefix, in any book — the letter was really seen
  const letterRead = new Set();
  books.forEach(b => b.parts.forEach(p => { if (/^1[A-Z]\d{3}-/.test(p.pn || "") && !/^1O/.test(p.pn)) letterRead.add(p.pn); }));
  const out = { fixed: 0, byTier: { known: 0, read: 0, stem: 0 }, left: 0 };
  books.forEach(b => {
    const stems = {};   // "010" → Set of lettered prefixes seen in THIS book
    b.parts.forEach(p => { const m = (p.pn || "").match(/^1([A-Z])(\d{3})-/); if (m && m[1] !== "O") (stems[m[2]] = stems[m[2]] || new Set()).add(m[1]); });
    b.parts.forEach(p => {
      const m = (p.pn || "").match(SUSPECT);
      if (!m || isKnown(p.pn)) return;
      const zero = m[1] === "0" || m[1] === "O";
      const letters = CONFUSE[m[1]].split(""), rest = p.pn.slice(2);
      const as = L => "1" + L + rest;
      let pick = null, tier = "";
      const known = letters.filter(L => isKnown(as(L)));
      if (known.length === 1) { pick = known[0]; tier = "known"; }
      else if (!known.length) {
        const read = letters.filter(L => letterRead.has(as(L)));
        // The misread runs ONE way only (measured 2026-10-05): of 124 known-real "16…" numbers the
        // OCR saw, 0 were read with a letter; of 175 known-real "1C/1G…", 56 were read as "16…".
        // So a lettered read elsewhere is trusted over the digit — a guard demanding the letter be
        // read MORE often threw away 306 right corrections (V2607 p33: a whole column of 1C010).
        if (read.length === 1) { pick = read[0]; tier = "read"; }
        else if (!read.length && zero) {
          const s = [...(stems[m[2]] || [])].filter(L => letters.includes(L));
          if (s.length === 1) { pick = s[0]; tier = "stem"; }
        }
      }
      if (!pick) { if (zero) out.left++; return; }
      // the two reads that "agreed" agreed on the WRONG number — so an inferred fix is not confirmed
      // (the page shows "check"); a "known" fix is confirmed by import.js's own known-number test
      p.ocr = Object.assign({}, p.ocr, { prefixFixedFrom: p.pn, prefixFix: tier, agree: false });
      p.pn = as(pick);
      out.fixed++; out.byTier[tier]++;
    });
  });
  return out;
}

module.exports = { fixMisreadPrefixes };
