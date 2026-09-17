// =======================================================================================
// PrepDiag.js — TEMPORARY, READ-ONLY map of the Prep Queue sheet
// =======================================================================================
//
// Written 2026-09-17 for the sort-corruption incident. It WRITES NOTHING — no
// setValue, no setBackground, no insert/delete. Safe to run on a damaged sheet.
//
// Answers the questions the screenshots could not:
//   • how many structural markers exist, and at which rows (a SECOND stray
//     "NEEDS PHOTOS" or "INCOMING" anywhere in col A silently re-points every
//     boundary helper, because both scan top-down and return the FIRST hit)
//   • what each segment actually holds
//   • which rows still wear band / header FORMATTING while holding data —
//     those mark where a divider USED to sit (clearContent never clears format)
//   • how much of INCOMING is a duplicate of the NEEDS PHOTOS content
//   • what state is at risk (✔ DONE ticks, FIRST SEEN dates)
//
// DELETE THIS FILE once the incident is closed.
// =======================================================================================

function diagnosePrepQueue() {
  var out = [];
  function say(s) { out.push(s); console.log(s); }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) { say("✗ Prep Queue sheet not found."); return out.join("\n"); }

  var maxRows = sheet.getMaxRows();
  var lastRow = sheet.getLastRow();
  var W = PREP_QUEUE.dataWidth;

  say("=== PREP QUEUE MAP · " + new Date() + " ===");
  say("maxRows=" + maxRows + "  lastRow=" + lastRow + "  dataWidth=" + W);

  var n = Math.max(lastRow, 1);
  var vals = sheet.getRange(1, 1, n, W).getValues();
  var bgs  = sheet.getRange(1, 1, n, 1).getBackgrounds();

  var SKU = PREP_QUEUE.idx("SKU"), LOC = PREP_QUEUE.idx("LOCATION"),
      HAND = PREP_QUEUE.idx("HAND"), NOTE = PREP_QUEUE.idx("NOTE"),
      DATE = PREP_QUEUE.idx("DATE_ADDED"), DONE = PREP_QUEUE.idx("DONE");

  // ---- 1. every structural marker, not just the first -------------------------------
  var incRows = [], photoRows = [], hdrRows = [];
  for (var i = 0; i < n; i++) {
    var a = String(vals[i][SKU] || "").trim().toUpperCase();
    if (a === PREP_QUEUE.boundaryMarker) incRows.push(i + 1);
    if (a === PREP_PHOTO.marker)         photoRows.push(i + 1);
    if (a === String(PREP_QUEUE.headers[0]).trim().toUpperCase() ||
        a === String(PREP_PHOTO.headers[0]).trim().toUpperCase()) hdrRows.push(i + 1);
  }
  say("");
  say("MARKERS (helpers return the FIRST of each):");
  say("  'INCOMING'      rows: " + (incRows.join(", ")   || "(none)") +
      (incRows.length   > 1 ? "   ⚠⚠ MORE THAN ONE" : ""));
  say("  'NEEDS PHOTOS'  rows: " + (photoRows.join(", ") || "(none)") +
      (photoRows.length > 1 ? "   ⚠⚠ MORE THAN ONE" : ""));
  say("  header rows (col A = a header label): " + (hdrRows.join(", ") || "(none)"));
  say("  → _getPrepBoundaryRow  would return: " + (incRows.length   ? incRows[0]   : -1));
  say("  → _getPhotoBoundaryRow would return: " + (photoRows.length ? photoRows[0] : -1));
  say("  → _prepWalkEnd         would return: " +
      (photoRows.length ? Math.min(lastRow, photoRows[0] - 1) : lastRow));

  // ---- 2. orphan formatting: a data row still wearing band/header paint -------------
  // #ffd400 = the yellow band, #1d1d1b = the dark header. A row carrying one of
  // those while holding a normal SKU is where a divider/header USED to be.
  say("");
  say("ORPHAN BAND / HEADER FORMATTING (rows holding data but painted as structure):");
  var orphans = 0;
  for (var j = 0; j < n; j++) {
    var bg = String(bgs[j][0] || "").toLowerCase();
    var isBand = (bg === "#ffd400"), isHdr = (bg === "#1d1d1b");
    if (!isBand && !isHdr) continue;
    var a2 = String(vals[j][SKU] || "").trim();
    var au = a2.toUpperCase();
    var structural = (au === PREP_QUEUE.boundaryMarker || au === PREP_PHOTO.marker ||
                      hdrRows.indexOf(j + 1) !== -1 || j + 1 === PREP_QUEUE.titleRow ||
                      j + 1 === PREP_QUEUE.headerRow);
    if (structural || a2 === "") continue;
    orphans++;
    say("  row " + (j + 1) + "  bg=" + bg + "  SKU=" + a2 +
        "  LOC=" + vals[j][LOC] + "  NOTE=" + String(vals[j][NOTE]).slice(0, 48));
  }
  if (!orphans) say("  (none)");

  // ---- 3. segment contents ----------------------------------------------------------
  var incDiv   = incRows.length   ? incRows[0]   : -1;
  var photoDiv = photoRows.length ? photoRows[0] : -1;

  function segment(label, from, toInclusive) {
    if (toInclusive < from) { say("  " + label + ": (empty range " + from + ".." + toInclusive + ")"); return null; }
    var skus = [], ticked = 0, dates = {};
    for (var k = from; k <= toInclusive && k <= n; k++) {
      var s = String(vals[k - 1][SKU] || "").trim();
      if (!s) continue;
      skus.push({ row: k, sku: s, loc: vals[k - 1][LOC], date: vals[k - 1][DATE] });
      if (vals[k - 1][DONE] === true) ticked++;
      var d = vals[k - 1][DATE];
      var key = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), "M/d/yy")
        : String(d || "(blank)");
      dates[key] = (dates[key] || 0) + 1;
    }
    var top = Object.keys(dates).sort(function (a, b) { return dates[b] - dates[a]; }).slice(0, 5);
    say("  " + label + "  rows " + from + ".." + toInclusive +
        "   non-blank=" + skus.length + "   ✔DONE ticked=" + ticked);
    if (skus.length) {
      say("     first: row " + skus[0].row + " " + skus[0].sku + " @ " + skus[0].loc);
      say("     last:  row " + skus[skus.length - 1].row + " " + skus[skus.length - 1].sku +
          " @ " + skus[skus.length - 1].loc);
      say("     DATE ADDED spread: " + top.map(function (t) { return t + "×" + dates[t]; }).join("  "));
    }
    return skus;
  }

  say("");
  say("SEGMENTS:");
  var curEnd = (incDiv > 0) ? incDiv - 1 : lastRow;
  var current  = segment("CURRENT ", PREP_QUEUE.dataStartRow, curEnd);
  var incoming = (incDiv > 0)
    ? segment("INCOMING", incDiv + 2, (photoDiv > 0) ? photoDiv - 1 : lastRow)
    : null;
  var photos   = (photoDiv > 0) ? segment("PHOTOS  ", photoDiv + 2, lastRow) : null;

  // ---- 4. how much of INCOMING duplicates the photo section -------------------------
  if (incoming && photos) {
    var pset = {};
    photos.forEach(function (p) { pset[String(p.sku).toLowerCase()] = p.row; });
    var dup = incoming.filter(function (it) { return pset[String(it.sku).toLowerCase()]; });
    say("");
    say("OVERLAP  INCOMING ∩ PHOTOS: " + dup.length + " of " + incoming.length +
        " incoming rows share a SKU with the photo section" +
        (dup.length > incoming.length * 0.5 ? "   ⚠⚠ INCOMING IS MOSTLY PHOTO CONTENT" : ""));
    if (dup.length) {
      say("  e.g. " + dup.slice(0, 5).map(function (d) {
        return d.sku + " (inc row " + d.row + " / photo row " + pset[String(d.sku).toLowerCase()] + ")";
      }).join(",  "));
    }
    var uniq = incoming.filter(function (it) { return !pset[String(it.sku).toLowerCase()]; });
    say("  INCOMING rows NOT in the photo section (the likely REAL prep items): " + uniq.length);
    uniq.slice(0, 40).forEach(function (u) {
      say("     row " + u.row + "  " + u.sku + "  @ " + u.loc + "  added " +
          ((u.date instanceof Date)
            ? Utilities.formatDate(u.date, Session.getScriptTimeZone(), "M/d/yy h:mm a")
            : u.date));
    });
    if (uniq.length > 40) say("     … +" + (uniq.length - 40) + " more");
  }

  say("");
  say("=== END MAP ===");
  return out.join("\n");
}


// =======================================================================================
// REPAIR — strip the duplicated photo-queue content out of the INCOMING table
// =======================================================================================
//
// WHAT THE MAP ESTABLISHED (2026-09-17):
//   INCOMING (rows 18..508) holds 479 non-blank rows. 466 of them share a SKU with
//   the NEEDS PHOTOS section below, 459 of those are stamped 8/4/26 — the day the
//   photo queue shipped. Row 418 is an ORPHAN PHOTO HEADER (col A "◈ SKU", col C
//   "LOCATION", col F "FIRST SEEN") that got sorted in by its literal "LOCATION"
//   text. So the photo content has been sitting in the INCOMING region since the
//   2026-08-04 rollout, below the fold; today's Sort Incoming interleaved it with
//   the real items by location, which is what made it visible.
//
// THE KEEP RULE — two independent signals, so neither has to be trusted alone:
//   • DATE ADDED carries a TIME (h:mm AM/PM). Both prep write paths stamp
//     "M/d/yy h:mm a" (PrepQueue.js:700 quick-add, :1158 on-sheet edit);
//     refreshPhotoQueue stamps date-only "M/d/yy" (PhotoQueue.js:271). A row with
//     a time was typed by a person.
//   • SKU is absent from the photo section.
//   A row is DROPPED only when it is in the photo section AND has no time — so a
//   genuine prep item that ALSO needs photos survives on the time signal.
//   ⚠ Edge: an item added at exactly 12:00 AM reads as date-only. The housekeeping
//   window is 6am-6pm and nobody adds prep at midnight; the SKU signal covers it.
//
// Preview writes nothing. Commit is a separate call and re-reads from scratch.
// =======================================================================================

/**
 * True when DATE ADDED is DATE-ONLY (no time) — the photo-queue signature.
 *
 * ⚠⚠ THE TIMEZONE TRAP THAT BROKE THE FIRST CUT (2026-09-17). refreshPhotoQueue
 * writes a date-only STRING; Sheets coerces it to a Date at midnight in the
 * SPREADSHEET's timezone (America/Chicago). The SCRIPT runs in Asia/Amman
 * (UTC+3), so v.getHours() reads 8, not 0 — every photo row looked like it
 * carried a time and the rule kept 478 of 479. Compare against midnight in the
 * SHEET's zone, never the script's.
 */
function _prepDateOnly(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "HH:mm") === "00:00";
  return !/\d{1,2}:\d{2}/.test(String(v || ""));   // "M/d/yy" vs "M/d/yy h:mm a"
}

/**
 * Shared verdict for one INCOMING row. Returns null for a blank row.
 * DROPS only when BOTH signals agree — the SKU is in the photo section AND the
 * stamp is date-only. Either one alone leaves the row alone, so anything
 * ambiguous is kept.
 */
function _prepIncomingVerdict(rowVals, photoSkuSet, headerSet, tz) {
  var sku = String(rowVals[PREP_QUEUE.idx("SKU")] || "").trim();
  if (!sku) return null;
  if (headerSet[sku.toUpperCase()]) return { keep: false, why: "orphan photo HEADER row" };
  var inPhotos = !!photoSkuSet[sku.toLowerCase()];
  var dateOnly = _prepDateOnly(rowVals[PREP_QUEUE.idx("DATE_ADDED")], tz);
  if (inPhotos && dateOnly) return { keep: false, why: "duplicated photo-queue row" };
  return { keep: true, why: inPhotos ? "real prep item (also needs a photo)" : "real prep item" };
}

/**
 * Gather the INCOMING block and classify every row. Pure read — no writes.
 * Returns null (after logging why) when the sheet's structure isn't safe to act on.
 */
function _prepRepairPlan(sheet, say) {
  var tz = sheet.getParent().getSpreadsheetTimeZone();   // NOT the script's zone
  var lastRow = sheet.getLastRow();
  var colA = sheet.getRange(1, PREP_QUEUE.cols.SKU, lastRow, 1).getValues();

  var incRows = [], photoRows = [];
  for (var i = 0; i < colA.length; i++) {
    var a = String(colA[i][0] || "").trim().toUpperCase();
    if (a === PREP_QUEUE.boundaryMarker) incRows.push(i + 1);
    if (a === PREP_PHOTO.marker)         photoRows.push(i + 1);
  }
  // Both boundary helpers return the FIRST hit top-down, so a second stray marker
  // would silently re-point the block. Refuse rather than act on an ambiguous map.
  if (incRows.length !== 1) {
    say("✗ REFUSING — expected exactly one 'INCOMING' marker, found " + incRows.length +
        " at rows " + (incRows.join(", ") || "(none)")); return null;
  }
  if (photoRows.length !== 1) {
    say("✗ REFUSING — expected exactly one 'NEEDS PHOTOS' marker, found " + photoRows.length +
        " at rows " + (photoRows.join(", ") || "(none)")); return null;
  }

  var incDiv = incRows[0], photoDiv = photoRows[0];
  var segStart = incDiv + 2, segEnd = photoDiv - 1;
  if (segEnd < segStart) { say("✗ REFUSING — INCOMING block is empty."); return null; }

  // SKUs currently in the photo section, and the header labels to strip.
  var photoSkuSet = {};
  if (lastRow >= photoDiv + 2) {
    var pv = sheet.getRange(photoDiv + 2, PREP_QUEUE.cols.SKU, lastRow - (photoDiv + 2) + 1, 1)
                  .getValues();
    for (var p = 0; p < pv.length; p++) {
      var ps = String(pv[p][0] || "").trim();
      if (ps) photoSkuSet[ps.toLowerCase()] = true;
    }
  }
  var headerSet = {};
  PREP_QUEUE.headers.concat(PREP_PHOTO.headers).forEach(function (h) {
    headerSet[String(h).trim().toUpperCase()] = true;
  });

  var block = sheet.getRange(segStart, 1, segEnd - segStart + 1, PREP_QUEUE.dataWidth).getValues();
  var keep = [], drop = [], blanks = 0;
  for (var r = 0; r < block.length; r++) {
    var v = _prepIncomingVerdict(block[r], photoSkuSet, headerSet, tz);
    if (!v) { blanks++; continue; }
    (v.keep ? keep : drop).push({ row: segStart + r, vals: block[r], why: v.why });
  }
  return { tz: tz, incDiv: incDiv, photoDiv: photoDiv, segStart: segStart, segEnd: segEnd,
           keep: keep, drop: drop, blanks: blanks, photoCount: Object.keys(photoSkuSet).length };
}

/** READ-ONLY. Shows exactly what the repair would keep and what it would remove. */
function previewPrepIncomingRepair() {
  var out = [];
  function say(s) { out.push(s); console.log(s); }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) { say("✗ Prep Queue sheet not found."); return out.join("\n"); }

  var plan = _prepRepairPlan(sheet, say);
  if (!plan) return out.join("\n");

  var L = PREP_QUEUE.idx("LOCATION"), H = PREP_QUEUE.idx("HAND"),
      N = PREP_QUEUE.idx("NOTE"), D = PREP_QUEUE.idx("DATE_ADDED"),
      S = PREP_QUEUE.idx("SKU"), DN = PREP_QUEUE.idx("DONE");

  say("=== PREP INCOMING REPAIR · PREVIEW (nothing written) ===");
  say("INCOMING block rows " + plan.segStart + ".." + plan.segEnd +
      "   photo section starts row " + plan.photoDiv + " (" + plan.photoCount + " SKUs)");
  say("");
  // Render stamps in the SHEET's timezone — that is the zone the date-only test
  // uses, so the printed time is the actual evidence behind each verdict.
  function fmt(v) {
    return (v instanceof Date) ? Utilities.formatDate(v, plan.tz, "M/d/yy HH:mm") : String(v || "");
  }

  say("KEEP — " + plan.keep.length + " row(s)   [times shown in " + plan.tz + "]:");
  plan.keep.forEach(function (k) {
    say("  row " + k.row + "  " + String(k.vals[S]) +
        "  @ " + k.vals[L] + "  hand=" + k.vals[H] +
        "  added " + fmt(k.vals[D]) + (k.vals[DN] === true ? "  [✔ DONE]" : "") +
        "   — " + k.why);
    say("        note: " + String(k.vals[N] || "").slice(0, 70));
  });

  var byWhy = {};
  plan.drop.forEach(function (d) { byWhy[d.why] = (byWhy[d.why] || 0) + 1; });
  say("");
  say("REMOVE — " + plan.drop.length + " row(s):");
  Object.keys(byWhy).forEach(function (w) { say("  " + byWhy[w] + " × " + w); });
  say("  sample (note the 00:00 stamps — the photo-queue signature):");
  plan.drop.slice(0, 6).forEach(function (d) {
    say("     row " + d.row + "  " + String(d.vals[S]) + "  @ " + d.vals[L] +
        "  added " + fmt(d.vals[D]) + "   — " + d.why);
  });
  var ticked = plan.drop.filter(function (d) { return d.vals[DN] === true; });
  say("  of those, ✔ DONE ticked: " + ticked.length +
      (ticked.length ? "   ⚠ REVIEW BEFORE COMMITTING" : "   (none — no tick state at risk)"));
  say("  blank rows in block: " + plan.blanks);

  say("");
  say("AFTER: INCOMING holds " + plan.keep.length + " row(s) + " +
      PREP_PHOTO.incomingGap + " blank rows; 'NEEDS PHOTOS' moves row " + plan.photoDiv +
      " → " + (plan.segStart + plan.keep.length + PREP_PHOTO.incomingGap));
  say("The photo section itself is NOT touched.");
  say("");
  say("Happy with the KEEP list? Run  repairPrepIncomingNow()");
  return out.join("\n");
}

/**
 * COMMIT. Re-reads and re-classifies from scratch — never trusts the preview's row
 * numbers (rows can move between calls; the 2026-05-08 row-shift lesson). Takes the
 * script lock, captures every value into memory BEFORE any structural change, and
 * preserves DONE ticks + checkbox validation. Keeper ORDER is preserved — this
 * removes junk and nothing else; re-sorting is a separate button.
 */
function repairPrepIncomingNow() {
  var out = [];
  function say(s) { out.push(s); console.log(s); }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { say("✗ Could not take the script lock — try again in a moment."); return out.join("\n"); }

  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PREP_QUEUE.sheetName);
    if (!sheet) { say("✗ Prep Queue sheet not found."); return out.join("\n"); }

    var plan = _prepRepairPlan(sheet, say);
    if (!plan) return out.join("\n");

    // Sanity gates — a repair that removes everything, or nothing, is a bug not a fix.
    if (plan.keep.length === 0) {
      say("✗ REFUSING — the rule would keep 0 rows. That is not a repair."); return out.join("\n");
    }
    if (plan.keep.length > 150) {
      say("✗ REFUSING — would keep " + plan.keep.length +
          " rows; expected a couple of dozen. Re-run the preview and check."); return out.join("\n");
    }
    if (plan.drop.length === 0) {
      say("ℹ️ Nothing to remove — INCOMING is already clean."); return out.join("\n");
    }

    var keepers = plan.keep.map(function (k) { return k.vals; });   // captured before any write
    var segStart = plan.segStart;
    var blockCount = plan.segEnd - segStart + 1;
    var targetCount = keepers.length + PREP_PHOTO.incomingGap;
    var W = PREP_QUEUE.dataWidth;

    var diff = blockCount - targetCount;
    if (diff > 0)      sheet.deleteRows(segStart + targetCount, diff);
    else if (diff < 0) sheet.insertRowsAfter(segStart + blockCount - 1, -diff);

    sheet.getRange(segStart, 1, targetCount, W).clearContent();
    try { sheet.getRange(segStart, PREP_QUEUE.cols.DONE, targetCount, 1).removeCheckboxes(); } catch (e) {}

    sheet.getRange(segStart, 1, keepers.length, W).setValues(keepers);
    // requireCheckbox re-plants validation WITHOUT resetting the written TRUE/FALSE.
    sheet.getRange(segStart, PREP_QUEUE.cols.DONE, keepers.length, 1)
         .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    try {
      applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU,
                            segStart, segStart + keepers.length - 1, buildSkuEnrichmentMap());
    } catch (e) { say("⚠ SKU links not rebuilt: " + e); }

    try { _refreshPrepQueueDuplicates(sheet); } catch (e) {}
    SpreadsheetApp.flush();

    say("✅ INCOMING repaired — kept " + keepers.length + ", removed " + plan.drop.length + ".");
    say("   'NEEDS PHOTOS' divider is now row " + _getPhotoBoundaryRow(sheet) +
        "; sheet lastRow " + sheet.getLastRow() + ".");
    say("   Photo section untouched.");
    return out.join("\n");
  } finally {
    lock.releaseLock();
  }
}
