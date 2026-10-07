#!/bin/bash
# Re-read every scanned Kubota book with the current reader, safely.   (2026-10-07)
#
#   bash catalogue/reread.sh
#
# 1. backs up the current OCR cache (out/ocr → out/ocr-before-<date>) — nothing is lost
# 2. clears the cached reads (keeps the .skip files, so non-parts-books aren't re-scanned)
# 3. re-reads every scanned book, 4 at a time       — the long part (hours)
# 4. re-runs the glyph pass on the fresh reads       — ~10-20 min
# Writes a log to out/reread.log and ends with "REREAD FINISHED".
# Nothing is imported or published: that happens after the new reads are compared to the old.
# Local only — no server, no sheet. Keep the computer awake until it finishes.

set -u
cd "$(dirname "$0")/.."                      # repo root
PDFS="$HOME/Downloads/PDF files/New uploads/Kubota PDF"   # the same 46 books as the current data (the wider folder adds 7 unrelated ones)
OUT=catalogue/out
STAMP=$(date +%Y%m%d-%H%M)
LOG=$OUT/reread.log

{
  echo "=== reread started $(date)"
  cp -a "$OUT/ocr" "$OUT/ocr-before-$STAMP" || { echo "BACKUP FAILED — stopping"; exit 1; }
  echo "backup: $OUT/ocr-before-$STAMP ($(ls "$OUT/ocr-before-$STAMP" | wc -l) files)"
  find "$OUT/ocr" -name '*.pdf.json' -delete
  find "$OUT/ocr" -name '*.glyph.json' -delete

  node catalogue/ocr-batch.js "$PDFS" --jobs 4

  echo "--- glyph pass $(date)"
  ls "$OUT"/ocr/*.pdf.json | xargs -P 4 -I{} node catalogue/glyph-pass.js {}

  echo "books read: $(ls "$OUT"/ocr/*.pdf.json | wc -l) (before: $(ls "$OUT/ocr-before-$STAMP"/*.pdf.json | wc -l))"
  echo "=== REREAD FINISHED $(date)"
} 2>&1 | tee "$LOG"
