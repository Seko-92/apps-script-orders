// =======================================================================================
// PhotoQueue.js — the "▌ NEEDS PHOTOS" section of the Prep Queue sheet
// =======================================================================================
//
// The user's convention: no item is ever left with ZERO images — a logo placeholder
// is added instead. So an item that still needs a real photo has 0 or 1 images
// (0 = slipped through, 1 = the logo). This section auto-lists every ACTIVE listing
// (items AND kits — kits are MI rows too) with <= 1 image, so a photographer has a
// live worklist. It SELF-CLEANS: once a real photo is added on eBay, MAIN sync pulls
// the new images into MI, and the next refresh drops the item (image count > 1).
//
// WHY IT'S A THIRD TABLE IN PREP QUEUE (not its own sheet): the user is consolidating
// "work-needed" items into one hub to fight sheet sprawl. Design keeps it SAFE:
//   • It's the LAST section, below a `▌ NEEDS PHOTOS` divider (col-A value EXACTLY
//     "NEEDS PHOTOS" — same exact-marker contract as the INCOMING divider).
//   • It reuses the Prep columns A..G (B relabeled IMG, E relabeled TITLE, F = FIRST
//     SEEN) so the existing HAND (D) / LOCATION (C) refreshers write the RIGHT columns
//     if they ever reach a photo row — no corruption by construction.
//   • It's MACHINE-OWNED: refreshPhotoQueue() fully rebuilds the region (smart-merge
//     preserves the ✔ DONE check + FIRST SEEN by SKU, like the OOS sheet).
//   • The prep walkers get one guard — _prepWalkEnd() — so counting/clearing stops at
//     the divider and never touches the photo section.
//
// PINNED /exec NOTE: refreshPrepQueueHand runs on the pinned /exec every 2 min. Its
// bounding (via _prepWalkEnd) only takes effect there after a New Version — until then
// it BENIGNLY refreshes HAND (col D) for photo rows (col D IS the HAND column, so no
// corruption), just extra work. Cut a New Version to bound it for perf.
// =======================================================================================

var PREP_PHOTO = {
  marker:    "NEEDS PHOTOS",   // exact col-A value on the divider row
  maxImages: 1,                // items with <= this many images need a real photo
  incomingGap: 12,             // blank rows left for the INCOMING table to grow above the divider
  // Physical columns reuse the Prep schema (A..G); only the header LABELS differ.
  headers: ["◈ SKU", "IMG", "LOCATION", "◫ HAND", "TITLE", "FIRST SEEN", "✔ DONE"],
  rightLabel: "ONLY THE LOGO · SHOOT A PHOTO"
};


/**
 * Find the ▌ NEEDS PHOTOS divider row (exact col-A match). Returns the 1-based
 * row, or -1 if the section hasn't been created yet.
 */
function _getPhotoBoundaryRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var vals = sheet.getRange(1, PREP_QUEUE.cols.SKU, lastRow, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === PREP_PHOTO.marker) return i + 1;
  }
  return -1;
}

/**
 * The end row the CURRENT/INCOMING prep walkers should stop at — one BELOW the
 * NEEDS PHOTOS divider (so nothing walks into the machine-owned photo section).
 * Falls back to getLastRow() when there's no photo section yet.
 */
function _prepWalkEnd(sheet) {
  var pb = _getPhotoBoundaryRow(sheet);
  var last = sheet.getLastRow();
  return (pb > 0) ? Math.min(last, pb - 1) : last;
}


/**
 * Create the NEEDS PHOTOS divider + header below the INCOMING table if missing.
 * Idempotent — returns the existing divider row if already present.
 */
function _ensurePhotoDivider(sheet) {
  var existing = _getPhotoBoundaryRow(sheet);
  if (existing > 0) return existing;

  var incBoundary = _getPrepBoundaryRow(sheet);
  var afterIncoming = (incBoundary > 0) ? incBoundary + 2 : PREP_QUEUE.dataStartRow;
  var base = Math.max(sheet.getLastRow(), afterIncoming);
  var dividerRow = base + PREP_PHOTO.incomingGap + 1;   // leave INCOMING room to grow

  if (sheet.getMaxRows() < dividerRow + 3) {
    sheet.insertRowsAfter(sheet.getMaxRows(), (dividerRow + 3) - sheet.getMaxRows());
  }
  sheet.getRange(dividerRow, PREP_QUEUE.cols.SKU).setValue(PREP_PHOTO.marker);
  _stylePrepBand(sheet, dividerRow, PREP_PHOTO.marker, PREP_PHOTO.rightLabel);

  var hRow = dividerRow + 1;
  _stylePrepHeaderRow(sheet, hRow);   // dark band + writes the STANDARD prep labels…
  // …then overwrite with OUR labels (IMG / TITLE / FIRST SEEN). setValues keeps the styling.
  sheet.getRange(hRow, 1, 1, PREP_PHOTO.headers.length).setValues([PREP_PHOTO.headers]);
  return dividerRow;
}


/**
 * Scan Master Inventory for items that still need a real photo: ACTIVE listings
 * with <= PREP_PHOTO.maxImages non-empty pictureUrl1..5 (kits included — they're
 * MI rows). Returns [{sku, images, location, title, hand}].
 */
function _scanItemsNeedingPhotos() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mi = ss.getSheetByName(DB_SHEET_NAME);
    if (!mi) return out;
    var lastRow = mi.getLastRow(), lastCol = mi.getLastColumn();
    if (lastRow < 2) return out;

    var headers = mi.getRange(1, 1, 1, lastCol).getValues()[0];
    function col(name) {
      var t = String(name).toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i] || "").trim().toLowerCase() === t) return i;
      }
      return -1;
    }
    var skuC = col(DB_SKU_HEADER);
    if (skuC < 0) return out;
    var titleC = col(DB_TITLE_HEADER), locC = col(DB_LOCATION_HEADER),
        statusC = col(DB_LISTING_STATUS_HEADER), qtyC = col(DB_QUANTITY_HEADER),
        soldC = col(DB_QUANTITY_SOLD_HEADER);
    var picC = [];
    for (var p = 1; p <= 5; p++) picC.push(col('pictureUrl' + p));

    var data = mi.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var sku = String(row[skuC] || "").trim();
      if (!sku) continue;

      // ACTIVE only (fail-open: blank/unknown status is included so a header
      // rename can't silently empty the list; a NON-blank non-Active is skipped).
      if (statusC >= 0) {
        var st = String(row[statusC] || "").trim();
        if (st && st.toLowerCase() !== "active") continue;
      }

      var imgCount = 0;
      for (var pc = 0; pc < picC.length; pc++) {
        if (picC[pc] >= 0 && String(row[picC[pc]] || "").trim()) imgCount++;
      }
      if (imgCount > PREP_PHOTO.maxImages) continue;   // has real photos → not our problem

      var qty  = qtyC  >= 0 ? (parseFloat(row[qtyC])  || 0) : 0;
      var sold = soldC >= 0 ? (parseFloat(row[soldC]) || 0) : 0;
      out.push({
        sku:      sku,
        images:   imgCount,
        location: locC   >= 0 ? String(row[locC]   || "").trim() : "",
        title:    titleC >= 0 ? String(row[titleC] || "").trim() : "",
        hand:     qty - sold
      });
    }
  } catch (e) { try { console.log("_scanItemsNeedingPhotos: " + e); } catch (_) {} }
  return out;
}


/**
 * Rebuild the NEEDS PHOTOS section from the MI scan. SMART-MERGE: the ✔ DONE
 * check and FIRST SEEN date are preserved per SKU across refreshes (so an
 * in-progress "shot it, pending upload" mark survives, and chronic-needs items
 * keep their original date). Items that gained a real photo simply drop off.
 * Sorted by LOCATION (aisle walk; NOT FOUND last), then SKU.
 */
function refreshPhotoQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet not found.";

  var divider   = _ensurePhotoDivider(sheet);
  var headerRow = divider + 1;
  var dataStart = divider + 2;

  // Re-STYLE the band + header EVERY refresh. A prep "Re-style Sheet" repaints
  // banding over rows 2..1000, which FLATTENS the photo band + header — that was
  // the "Re-style breaks the Needs-Photos header" bug (delete+refresh looked
  // fine because that path re-creates the band). Re-apply the dark band + our
  // labels (IMG / TITLE / FIRST SEEN) so Re-style can never flatten it again.
  _stylePrepBand(sheet, divider, PREP_PHOTO.marker, PREP_PHOTO.rightLabel);
  sheet.setRowHeight(divider, 30);
  _stylePrepHeaderRow(sheet, headerRow);   // dark band + sets row height 36
  sheet.getRange(headerRow, 1, 1, PREP_PHOTO.headers.length).setValues([PREP_PHOTO.headers]);

  // Capture prior DONE + FIRST SEEN by SKU for the smart-merge.
  var prior = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= dataStart) {
    var n = lastRow - dataStart + 1;
    var ex = sheet.getRange(dataStart, 1, n, PREP_QUEUE.dataWidth).getValues();
    for (var i = 0; i < ex.length; i++) {
      var s = String(ex[i][PREP_QUEUE.idx("SKU")] || "").trim();
      if (!s) continue;
      prior[s.toLowerCase()] = {
        done:      ex[i][PREP_QUEUE.idx("DONE")] === true,
        firstSeen: ex[i][PREP_QUEUE.idx("DATE_ADDED")]   // col F = FIRST SEEN
      };
    }
  }

  var items = _scanItemsNeedingPhotos();
  items.sort(function(a, b) {
    var la = a.location || "", lb = b.location || "";
    var pa = (!la || la === "NOT FOUND") ? 1 : 0, pb = (!lb || lb === "NOT FOUND") ? 1 : 0;
    if (pa !== pb) return pa - pb;
    // NATURAL aisle order, not lexical — see compareLocations() in Helpers.js.
    var byLoc = compareLocations(la, lb);
    if (byLoc !== 0) return byLoc;
    return String(a.sku).localeCompare(String(b.sku));
  });

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yy");
  var rows = items.map(function(it) {
    var p = prior[it.sku.toLowerCase()] || {};
    var r = new Array(PREP_QUEUE.dataWidth);
    r[PREP_QUEUE.idx("SKU")]        = it.sku;
    r[PREP_QUEUE.idx("QTY")]        = it.images;                 // B = IMG count
    r[PREP_QUEUE.idx("LOCATION")]   = it.location || "NOT FOUND";
    r[PREP_QUEUE.idx("HAND")]       = (it.hand != null && !isNaN(it.hand)) ? it.hand : "";
    r[PREP_QUEUE.idx("NOTE")]       = it.title;                  // E = TITLE
    r[PREP_QUEUE.idx("DATE_ADDED")] = p.firstSeen || today;     // F = FIRST SEEN (preserved)
    r[PREP_QUEUE.idx("DONE")]       = (p.done === true);
    return r;
  });

  // Resize the photo data region (it's the LAST section, so we own everything
  // from dataStart to the bottom). Ensure enough rows, clear, then write.
  var maxRows = sheet.getMaxRows();
  var need = dataStart + Math.max(rows.length, 1) - 1;
  if (maxRows < need) sheet.insertRowsAfter(maxRows, need - maxRows);

  var clearN = sheet.getMaxRows() - dataStart + 1;
  if (clearN > 0) {
    var clearRange = sheet.getRange(dataStart, 1, clearN, PREP_QUEUE.dataWidth);
    // Also wipe any stale duplicate-highlight background/border the prep painter
    // may have left on photo rows before it was bounded out of this section.
    clearRange.clearContent().setBackground(null).setBorder(false, false, false, false, false, false);
    sheet.getRange(dataStart, PREP_QUEUE.cols.DONE, clearN, 1).removeCheckboxes();
  }

  if (rows.length > 0) {
    sheet.getRange(dataStart, 1, rows.length, PREP_QUEUE.dataWidth).setValues(rows);
    // ✔ DONE checkboxes — plant validation WITHOUT insertCheckboxes (which would
    // reset every box), then re-write the preserved TRUE/FALSE values.
    var doneRange = sheet.getRange(dataStart, PREP_QUEUE.cols.DONE, rows.length, 1);
    doneRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    doneRange.setValues(rows.map(function(r) { return [r[PREP_QUEUE.idx("DONE")]]; }));
    // SKU → eBay listing links (so the photographer can see the current logo-only listing).
    try {
      applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, dataStart, dataStart + rows.length - 1,
                            buildSkuEnrichmentMap());
    } catch (e) { try { console.log("photo SKU links: " + e); } catch (_) {} }
  }

  stampSheetPulse(sheet, SHEET_PULSE.prepQueue.stamp);
  return "✅ Needs-Photos refreshed — " + rows.length + " item(s) still need a photo.";
}


/** Count of items currently needing photos (reads the section, cheap) — for the sidebar badge. */
function getPhotoQueueCount() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
    if (!sheet) return 0;
    var divider = _getPhotoBoundaryRow(sheet);
    if (divider < 0) return 0;
    var dataStart = divider + 2;
    var lastRow = sheet.getLastRow();
    if (lastRow < dataStart) return 0;
    var vals = sheet.getRange(dataStart, PREP_QUEUE.cols.SKU, lastRow - dataStart + 1, 1).getValues();
    var c = 0;
    for (var i = 0; i < vals.length; i++) { if (String(vals[i][0] || "").trim()) c++; }
    return c;
  } catch (e) { return 0; }
}

/** Sidebar: open the Prep Queue sheet scrolled to the NEEDS PHOTOS section. */
function openPhotoQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet not found.";
  SpreadsheetApp.setActiveSheet(sheet);
  var divider = _getPhotoBoundaryRow(sheet);
  if (divider > 0) sheet.setActiveRange(sheet.getRange(divider, 1));
  return "✅ Opened Needs-Photos.";
}


// =======================================================================================
// ITS OWN TRIGGER (2026-08-19)
// =======================================================================================
//
// ⚠⚠ WHY THIS JOB LEFT THE HOUSEKEEPING PASS — MEASURED, not guessed.
//
//   _housekeepingPass WITH this job     352.748 s
//   _housekeepingPass WITHOUT it         18.219 s
//
// This one scan was ~334s — about 95% of the whole hourly pass — against Apps
// Script's hard 360s per-execution ceiling. Six unrelated jobs (HAND, Prep HAND,
// Out of Stock, Prep locations, the straggler watchdog, the tier-2 pulse) were
// sharing a budget that this job had already spent, and a timeout would have
// silently dropped whichever of them ran last. That is not a cost problem, it is
// a BLAST-RADIUS problem: an expensive job does not belong in a shared execution
// with cheap ones that have nothing to do with it.
//
// It is expensive for a structural reason, not a fixable one. _scanItemsNeedingPhotos
// reads the WHOLE of Master Inventory (3,635 x 198 = 719,730 cells) because it needs
// pictureUrl1..5, which the shared `maps` do not carry, and then rewrites ~464 rows
// with checkboxes, rich-text SKU links and a sort.
//
// ⚠ DO NOT "FIX" IT BY NARROWING THE READ. Tested on this exact sheet 2026-08-19:
//   198 columns vs a 39-column span measured 1,905 ms vs 1,872 ms — no faster.
//   In Apps Script the ROUND TRIP dominates, not the payload. The cost is the
//   464-row WRITE pipeline, not the read.
//
// ⚠ ONCE A DAY IS ENOUGH, AND PURE ON-DEMAND IS NOT.
//   Enough, because the backlog moves in single digits per day (460 on 08-18,
//   464 on 08-19) — it only changes when a photographer uploads to eBay and
//   MAIN's hourly sync carries the new pictureUrl into MI. Staleness also only
//   ever runs in the "already shot but still listed" direction, which the
//   photographer catches the moment they open the listing.
//
//   But NOT on-demand only, because getPhotoQueueCount() reads the SHEET rather
//   than rescanning, and it feeds the Alerts badge AND the weekly digest + PDF
//   report. With no scheduled run, a month of nobody pressing the button gives
//   Monday's report a month-old number while looking perfectly healthy. A number
//   that is quietly stale is worse than one that is openly a day old.
//
// ⚠ 5AM HOUSTON, PINNED EXPLICITLY. Before the 6am shift so it is fresh when
//   anyone reads it, and nothing else is running so its ~334s competes with
//   nobody. `.inTimezone()` is NOT optional here: atHour() otherwise uses the
//   SCRIPT's timezone, which is Asia/Amman on this project — 8 hours off Houston.

var PHOTO_QUEUE_TRIGGER = {
  handler:  "runPhotoQueueRefresh",
  hour:     5,                       // 5am …
  timezone: "America/Chicago"        // … Houston, NOT the script's Asia/Amman
};


/**
 * Trigger target. Deliberately NOT work-hours gated — 5am is the point.
 *
 * ⚠ Takes no meaningful arguments. A time trigger hands its target an EVENT
 * OBJECT as the first argument, so anything it accepted would be an event on
 * every scheduled run.
 */
function runPhotoQueueRefresh() {
  try {
    var msg = refreshPhotoQueue();
    console.log("Photo queue: " + msg);
    return msg;
  } catch (err) {
    // Never rethrow from a trigger target — a throw here only produces a failure
    // email, and the next daily run self-heals because the region is rebuilt
    // from scratch every time.
    console.log("Photo queue error: " + err);
    return "❌ Photo queue: " + err;
  }
}


/**
 * Install the daily trigger. Idempotent — safe to re-run.
 *
 * ⚠ MATCHES ON getHandlerFunction(), never on position. Picking trigger rows by
 * eye is what took nine handlers down in August.
 *
 * ⚠ It does NOT refresh the table itself. setupHousekeeping calls this, and a
 * ~334s scan on top of that function's own work would push it past the 360s
 * execution ceiling — the exact failure this whole change exists to remove. The
 * table is already populated; the first scheduled run keeps it that way, and the
 * sidebar's 📸 button forces one immediately if anyone cannot wait.
 */
function setupPhotoQueueTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === PHOTO_QUEUE_TRIGGER.handler) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  ScriptApp.newTrigger(PHOTO_QUEUE_TRIGGER.handler)
    .timeBased()
    .atHour(PHOTO_QUEUE_TRIGGER.hour)
    .inTimezone(PHOTO_QUEUE_TRIGGER.timezone)
    .everyDays(1)
    .create();

  var msg = "📸 Photo queue trigger installed — daily ~" + PHOTO_QUEUE_TRIGGER.hour +
            ":00 " + PHOTO_QUEUE_TRIGGER.timezone +
            " (" + removed + " old trigger(s) removed)";
  console.log(msg);
  return msg;
}


/** Remove it. Matches on handler name, same discipline as the installer. */
function removePhotoQueueTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === PHOTO_QUEUE_TRIGGER.handler) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  var msg = "📸 Photo queue trigger removed (" + removed + ").";
  console.log(msg);
  return msg;
}
