# Quantity-reader before/after check

Any change to how `lib/scan-book.js` reads quantities must be compared against the old reader on
real pages BEFORE a re-read is imported (a loosened rule once wiped two books' REFs silently).

1. `regress/jobs.tsv` lists `mode · book# · pdf · pages` (10 scanned books × 3 dense pages, 2026-10-07).
   "old" runs with `OCR_NO_SKEW=1` (the pre-2026-10-07 bands); "new" runs the current code.
   For a different change, swap that env var in `drive.js` for the switch your change adds.
2. `node catalogue/regress/drive.js <work dir> catalogue` — results land in `<work dir>/reg/` (create it);
   4 at a time, ~5 min per side, local only.
3. `node catalogue/regress/cmp.js <work dir>` — per book: same / filled / emptied / changed value,
   with every changed row. Then CHECK EACH CHANGED ROW AGAINST THE MANUAL PAGE — a change is not a fix.
4. `node catalogue/regress/one-page.js "$PWD/catalogue" <pdf> <pages>` reads single pages
   (`OCR_DEBUG_QTY=1` / `OCR_DEBUG_SKEW=1` explain each cell).

Result 2026-10-07 (qty bands follow the digits): 708 rows · 684 same · 24 changed, all verified
against the manuals as fixes (D1105 p12 ×21, D782 p36 ×3), 0 regressions.
