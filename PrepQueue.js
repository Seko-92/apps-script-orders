// =======================================================================================
// PREP_QUEUE.gs — "Prep Queue" sheet for the warehouse employee's personal todo list
// =======================================================================================
//
// PURPOSE — IMPORTANT BOUNDARY
//   This sheet is a PERSONAL TODO LIST for the warehouse employee. They add
//   SKUs they want to work on LATER — restock, repackage, recheck, organize.
//
//   This is NOT for deferred customer orders. The All Orders sheet is for
//   real customer orders that MUST be prepped and shipped immediately. The
//   sidebar deliberately does NOT have a "move from All Orders to Prep Queue"
//   button — the workflow rule is one-way: orders go OUT, never PARKED.
//
// ARCHITECTURE
//   Standalone sheet ("Prep Queue") with its own minimal schema.
//   Columns: SKU · QTY · LOCATION · HAND · NOTE · DATE ADDED
//
//   onEdit hook (prepQueueOnEdit, dispatched from Main.js onEdit):
//     - When SKU is entered/changed in column A, auto-fill LOCATION + HAND
//       from Master Inventory (same lookup as LiveSync uses for All Orders).
//     - When a SKU is first written to a previously-empty row, stamp DATE ADDED.
//
// PUBLIC API
//   setupPrepQueueSheet()  — one-time: create sheet, style, install validation
//   openPrepQueue()        — switch the user's active sheet to Prep Queue (sidebar button)
//   clearPrepQueue()       — wipe all data rows (sidebar danger button, two-step confirm)
//   prepQueueOnEdit(e)     — onEdit dispatcher (called from Main.js)
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
var PREP_QUEUE = {
  sheetName: "Prep Queue",

  // 1-based column positions
  cols: {
    SKU:        1,   // A
    QTY:        2,   // B
    LOCATION:   3,   // C
    HAND:       4,   // D
    NOTE:       5,   // E
    DATE_ADDED: 6    // F
  },

  // 0-based array indices for getValues() iteration
  idx: function(name) { return PREP_QUEUE.cols[name] - 1; },

  dataWidth:    6,
  headerRow:    1,
  dataStartRow: 2,

  // Headers — written by setupPrepQueueSheet, displayed in dark/yellow brand band
  headers: ["◈ SKU", "# QTY", "LOCATION", "◫ HAND", "NOTE", "DATE ADDED"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * One-time setup: creates "Prep Queue" sheet if missing, applies headers,
 * brand styling, banding, and HAND low-stock conditional formatting.
 * Idempotent — safe to re-run.
 */
function setupPrepQueueSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(PREP_QUEUE.sheetName);
  }

  // --- HEADERS ---
  sheet.getRange(PREP_QUEUE.headerRow, 1, 1, PREP_QUEUE.dataWidth)
    .setValues([PREP_QUEUE.headers])
    .setBackground('#1d1d1b')   // brand black
    .setFontColor('#ffd966')    // brand yellow
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Thick yellow underline below header
  sheet.getRange(PREP_QUEUE.headerRow, 1, 1, PREP_QUEUE.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(PREP_QUEUE.cols.SKU,        110);
  sheet.setColumnWidth(PREP_QUEUE.cols.QTY,         60);
  sheet.setColumnWidth(PREP_QUEUE.cols.LOCATION,    95);
  sheet.setColumnWidth(PREP_QUEUE.cols.HAND,        80);
  sheet.setColumnWidth(PREP_QUEUE.cols.NOTE,       260);
  sheet.setColumnWidth(PREP_QUEUE.cols.DATE_ADDED, 130);

  // --- DATA AREA: column-level format (so new rows inherit) ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - PREP_QUEUE.dataStartRow + 1;

  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.QTY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.NOTE, dataRows, 1)
    .setFontFamily('Roboto').setFontStyle('italic').setFontSize(10)
    .setFontColor('#434343').setHorizontalAlignment('left');
  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.DATE_ADDED, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');

  sheet.getRange(PREP_QUEUE.dataStartRow, 1, dataRows, PREP_QUEUE.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING (cream alternation) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, PREP_QUEUE.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- HAND LOW-STOCK CONDITIONAL FORMATTING ---
  var existingRules = sheet.getConditionalFormatRules();
  var keptRules = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) {
      return rg.getColumn() === PREP_QUEUE.cols.HAND;
    });
  });
  var handRange = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, dataRows, 1);
  var handRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      "=AND(ISNUMBER(D" + PREP_QUEUE.dataStartRow + "), D" + PREP_QUEUE.dataStartRow + "<=20)"
    )
    .setBackground('#ffd966').setFontColor('#1d1d1b').setBold(true)
    .setRanges([handRange])
    .build();
  keptRules.push(handRule);
  sheet.setConditionalFormatRules(keptRules);

  // --- FREEZE HEADER ROW ---
  sheet.setFrozenRows(1);

  // Paint any existing duplicate SKUs (idempotent — clears stale highlights too)
  _refreshPrepQueueDuplicates(sheet);

  // Backfill SKU → eBay listing links across existing rows (same as the All
  // Orders "Link SKUs" backfill). Best-effort.
  try { refreshPrepQueueSkuLinks(); }
  catch (e) { try { Logger.log("setupPrepQueueSheet: SKU link backfill error: " + e); } catch (_) {} }

  return "✅ Prep Queue sheet ready.";
}


/**
 * Shared LOCATION + HAND resolution for the three Prep Queue entry paths
 * (sidebar preview, quick-add, on-sheet edit). HAND prefers Zoho's
 * available_stock — Prep items are restock/personal, so Zoho is the unified
 * stock truth and covers items not listed on eBay — with MI as fallback. No
 * committed subtraction: HAND = available, matching the All Orders HAND
 * semantics (2026-05-09) so the Prep number equals the All Orders number.
 *
 * Pass pre-built maps for batch callers; omit them (null) for single lookups.
 * Returns { location, hand, found }. `found` is true when the SKU is known to
 * either MI (eBay shelf) or Zoho.
 */
function _prepResolveStock(skuLower, locationMap, inventoryMap, zohoMap) {
  var location = locationMap ? (locationMap.get(skuLower) || "NOT FOUND")
                             : getSingleLocation(skuLower);
  var inMi = (location !== "NOT FOUND");

  var zo = zohoMap ? zohoMap.get(skuLower) : getSingleZohoStock(skuLower);
  var hand = "";
  if (zo) {
    hand = zo.available;
  } else if (inMi) {
    var inv = inventoryMap ? inventoryMap.get(skuLower) : getSingleInventory(skuLower);
    hand = inv ? inv.available : "";
  }

  return { location: location, hand: hand, found: (inMi || zo != null) };
}


/**
 * Quick lookup for the sidebar's live preview as the user types a SKU.
 * Returns { sku, location, hand, found } — `found` is false when the SKU is
 * unknown to both MI and Zoho (sidebar uses this to flag NOT FOUND visually).
 *
 * HAND matches the All Orders sheet's HAND column (Zoho-first for direct/non-
 * eBay items, MI for eBay items; no committed subtraction).
 */
function lookupSkuForPrepQueue(sku) {
  var skuLower = String(sku || "").trim().toLowerCase();
  if (!skuLower) return { sku: "", location: "", hand: "", found: false };

  var r = _prepResolveStock(skuLower, null, null, null);
  return { sku: skuLower, location: r.location, hand: r.hand, found: r.found };
}


/**
 * Rewrite HAND for EVERY existing Prep Queue row from the current MI + Zoho
 * maps. This is the Prep-Queue counterpart to recomputeHand (which only walks
 * the All Orders sheet) — without it, Prep rows stayed frozen at their entry
 * value unless the picker re-typed the SKU. Zoho-first, MI fallback, no
 * committed subtraction. Only HAND is rewritten; LOCATION/QTY/NOTE/DATE are
 * left untouched. Returns a status string.
 */
function refreshPrepQueueHand() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet not found.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return "ℹ️ Prep Queue empty.";

  var nRows = lastRow - PREP_QUEUE.dataStartRow + 1;
  var skuVals  = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU,  nRows, 1).getValues();
  var handVals = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, nRows, 1).getValues();

  var maps = buildLocationAndInventoryMaps();
  var locationMap  = maps.locationMap;
  var inventoryMap = maps.inventoryMap;
  var zohoMap      = buildZohoStockMap();

  var out = [];
  var updated = 0;
  for (var i = 0; i < nRows; i++) {
    var sku = String(skuVals[i][0] || "").trim();
    if (!sku) { out.push([handVals[i][0]]); continue; }   // preserve blank rows as-is
    var r = _prepResolveStock(sku.toLowerCase(), locationMap, inventoryMap, zohoMap);
    out.push([r.hand]);
    updated++;
  }

  sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.HAND, nRows, 1).setValues(out);
  return "✅ Prep Queue HAND refreshed for " + updated + " row(s).";
}


/**
 * Quick-add path: append a new Prep Queue row with the given SKU, QTY, and
 * NOTE. LOCATION + HAND + DATE ADDED are auto-filled server-side (programmatic
 * setValues doesn't fire onEdit, so we inline the lookup here).
 *
 * Returns { success, row, sku, location, hand, found, error? } so the
 * sidebar can show what just landed.
 */
function addPrepQueueItem(sku, qty, note) {
  sku = String(sku || "").trim();
  if (!sku) return { success: false, error: "SKU required" };

  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) {
    setupPrepQueueSheet();
    sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  }

  qty  = parseInt(qty) || 1;
  note = String(note || "").trim();

  // Inline lookup (programmatic writes don't trigger prepQueueOnEdit).
  // HAND is Zoho-first (matches All Orders' HAND semantics, no committed sub).
  var skuLower = sku.toLowerCase();
  var r = _prepResolveStock(skuLower, null, null, null);
  var location = r.location;
  var hand = r.hand;
  var found = r.found;

  var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

  var lastRow = sheet.getLastRow();
  var insertAt = Math.max(lastRow + 1, PREP_QUEUE.dataStartRow);

  sheet.getRange(insertAt, 1, 1, PREP_QUEUE.dataWidth).setValues([[
    sku, qty, location, hand, note, nowStr
  ]]);

  // Programmatic writes don't fire prepQueueOnEdit, so refresh duplicate
  // highlighting here. If this SKU is already in the queue, both rows get
  // marked; if it's new, no-op.
  _refreshPrepQueueDuplicates(sheet);

  // Link the new SKU cell (programmatic setValues doesn't fire prepQueueOnEdit).
  try {
    applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, insertAt, insertAt, buildSkuEnrichmentMap());
  } catch (enrErr) { try { Logger.log("addPrepQueueItem: SKU link error: " + enrErr); } catch (_) {} }

  return {
    success:  true,
    row:      insertAt,
    sku:      sku,
    location: location,
    hand:     hand,
    found:    found
  };
}


/**
 * One-shot migration: realigns existing data to the 6-column schema.
 *
 * Symptom this fixes: when the sheet was originally built (or hand-edited)
 * with only 5 columns (SKU · LOCATION · HAND · NOTE · DATE), new entries
 * from `prepQueueOnEdit` start landing in the wrong columns because the
 * code uses the 6-column schema (QTY at B). Location appears under "HAND",
 * date drifts to a hidden col F, etc.
 *
 * Per-row detection:
 *   - If col B looks like a location code (e.g. "A-43", "L-233"), the row is
 *     in 5-col user layout → shift values right by one (insert blank QTY at B).
 *   - Otherwise the row is already in 6-col schema layout — leave it alone.
 *
 * Safe to re-run: rows already in schema layout are detected and not shifted.
 */
function repairPrepQueueLayout() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "❌ Prep Queue sheet not found.";

  var lastRow = sheet.getLastRow();

  if (lastRow < PREP_QUEUE.dataStartRow) {
    setupPrepQueueSheet();
    return "✅ Headers fixed — no data to migrate.";
  }

  // Read up to schema width (6) so we can detect both layouts in one shot.
  var nRows = lastRow - PREP_QUEUE.dataStartRow + 1;
  var data = sheet.getRange(PREP_QUEUE.dataStartRow, 1, nRows, PREP_QUEUE.dataWidth).getValues();

  // Location-code heuristic — Master Inventory uses formats like "A-43",
  // "L-233", "I-29". A column-B value matching that pattern is a strong
  // signal the row is in legacy 5-col user layout.
  var locPattern = /^[A-Z]+\s*-\s*\d+/i;

  var migrated = 0;
  var preserved = 0;
  var out = data.map(function(row) {
    var sku = row[0];
    if (!String(sku || "").trim()) {
      return row;  // empty row, leave as-is
    }
    var bStr = String(row[1] || "").trim();
    if (locPattern.test(bStr)) {
      // 5-col user layout: SKU | LOCATION | HAND | NOTE | DATE | (extra)
      // → 6-col schema:    SKU | QTY      | LOCATION | HAND | NOTE | DATE
      migrated++;
      return [sku, "", row[1], row[2], row[3], row[4]];
    }
    preserved++;
    return [sku, row[1], row[2], row[3], row[4], row[5]];
  });

  // Write the realigned data back
  sheet.getRange(PREP_QUEUE.dataStartRow, 1, out.length, PREP_QUEUE.dataWidth).setValues(out);

  // Clean up any orphan column-G data left over from old broken inserts.
  if (sheet.getLastColumn() > PREP_QUEUE.dataWidth) {
    var extraCols = sheet.getLastColumn() - PREP_QUEUE.dataWidth;
    sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.dataWidth + 1, nRows, extraCols).clearContent();
  }

  // Re-apply headers + styling so the visual layout matches the schema.
  setupPrepQueueSheet();

  return "✅ Prep Queue repaired — " + migrated + " row(s) migrated, " +
         preserved + " already aligned.";
}


/**
 * Sidebar: switch the user's active view to the Prep Queue sheet.
 *
 * IMPORTANT: must use SpreadsheetApp.getActive() (the BOUND spreadsheet),
 * not openById(). setActiveSheet() only changes the user's visible tab when
 * called on the active spreadsheet reference. Using openById would silently
 * succeed but the view wouldn't switch — that was the v1 bug.
 */
function openPrepQueue() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet (open this from the spreadsheet UI).";

  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) {
    setupPrepQueueSheet();
    sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  }

  ss.setActiveSheet(sheet);
  return "✅ Opened " + PREP_QUEUE.sheetName;
}


/**
 * Sidebar: clear all data rows in the Prep Queue (keeps the header).
 * The sidebar button asks for confirmation before calling this.
 */
function clearPrepQueue() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return "ℹ️ Prep Queue sheet doesn't exist yet.";

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return "ℹ️ Queue already empty.";

  sheet.getRange(PREP_QUEUE.dataStartRow, 1, lastRow - PREP_QUEUE.dataStartRow + 1, PREP_QUEUE.dataWidth)
    .clearContent();

  // Clear duplicate-highlight backgrounds/borders on the now-empty rows so
  // the sheet visually resets clean (not just empty cells with stale highlights).
  _refreshPrepQueueDuplicates(sheet);

  return "✅ Cleared " + (lastRow - PREP_QUEUE.dataStartRow + 1) + " row(s) from " + PREP_QUEUE.sheetName + ".";
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEdit(e)
// =======================================================================================

/**
 * SKU edit on Prep Queue → auto-fill LOCATION + HAND + DATE ADDED.
 * Mirrors LiveSync's location-lookup pattern but for this dedicated sheet.
 *
 * Called from Main.js onEdit(e). Defensive (try/catch swallowed) — never
 * blocks other edit handlers.
 */
function prepQueueOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== PREP_QUEUE.sheetName) return;
    if (e.range.getColumn() !== PREP_QUEUE.cols.SKU) return;
    if (e.range.getRow() < PREP_QUEUE.dataStartRow) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();

    // Pre-build lookup maps once for the whole batch (cheaper than per-row):
    // MI (location + eBay stock) + Zoho (direct/non-eBay stock). HAND is
    // Zoho-first, matching the All Orders HAND semantics.
    var useMap = edits.length > 3;
    var locationMap = null;
    var inventoryMap = null;
    var zohoMap = null;

    if (useMap) {
      var maps = buildLocationAndInventoryMaps();
      locationMap = maps.locationMap;
      inventoryMap = maps.inventoryMap;
      zohoMap = buildZohoStockMap();
    }

    var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe LOCATION/HAND/DATE so the row reads as empty.
        sheet.getRange(row, PREP_QUEUE.cols.LOCATION).setValue("");
        sheet.getRange(row, PREP_QUEUE.cols.HAND).setValue("");
        sheet.getRange(row, PREP_QUEUE.cols.DATE_ADDED).setValue("");
        continue;
      }

      // Resolve LOCATION + HAND (Zoho-first HAND, MI fallback; no committed sub).
      var r = _prepResolveStock(skuLower, locationMap, inventoryMap, zohoMap);
      sheet.getRange(row, PREP_QUEUE.cols.LOCATION).setValue(r.location);
      sheet.getRange(row, PREP_QUEUE.cols.HAND).setValue(r.hand);

      // Stamp DATE ADDED. Re-stamp on EVERY SKU edit (the row's "added" date
      // is when this SKU landed, not when the row was first created — if you
      // overwrite the SKU, that's a new todo item).
      sheet.getRange(row, PREP_QUEUE.cols.DATE_ADDED).setValue(nowStr);
    }

    // SKU → eBay listing link (same enrichment as All Orders). One MI read
    // covers the whole edited batch.
    try {
      applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, startRow,
                            startRow + edits.length - 1, buildSkuEnrichmentMap());
    } catch (enrErr) { try { Logger.log("prepQueueOnEdit: SKU link error: " + enrErr); } catch (_) {} }

    // Refresh duplicate-SKU highlight after every edit batch — surfaces
    // dupes the moment they're typed, clears highlights when one of a
    // dup-pair is deleted/changed.
    _refreshPrepQueueDuplicates(sheet);
  } catch (err) {
    // Swallow — don't break other onEdit handlers
    try { Logger.log("prepQueueOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (mirrors the OutOfStock + All Orders pattern)
// =======================================================================================

/**
 * Scans column A, identifies SKUs that appear two or more times (case-insensitive,
 * trimmed), and paints those cells with a soft amber background + thick yellow
 * border so duplicates are unmistakable.
 *
 * IMPORTANT — only touches background + border. The cell's font, alignment,
 * number format, banding etc. are untouched. When a duplicate is removed (or
 * the SKU is changed/cleared), the next call to this function explicitly
 * clears the background and border on the formerly-duped row, restoring its
 * original look without disturbing other formatting.
 *
 * Per-cell explicit clears (not a range-wide reset) are deliberate: a fast
 * follow-up edit on a previously-duplicate cell could otherwise see stale
 * formatting because Sheets batches background/border writes. The explicit
 * per-cell pass ensures every row's state is set in a single, ordered batch.
 *
 * Run after any change that could affect column A:
 *   - setupPrepQueueSheet (initial paint)
 *   - prepQueueOnEdit (live, on every SKU edit)
 *   - addPrepQueueItem (programmatic adds from sidebar Quick Add)
 *   - clearPrepQueue (removes any lingering highlights from cleared rows)
 *   - repairPrepQueueLayout (transitively, via setupPrepQueueSheet)
 */
function _refreshPrepQueueDuplicates(sheet) {
  if (!sheet) return;

  // CRITICAL: scan the entire data band (not just lastRow), otherwise rows
  // where the SKU was deleted but background formatting persists won't get
  // their highlight cleared. `getLastRow()` returns the last row with
  // CONTENT, so a previously-painted row that had its SKU cleared (and
  // dropped off lastRow) keeps its yellow forever. We walk the full
  // data band so every row's background is set explicitly to either
  // "dupe-yellow" or "default (null)".
  var maxScanRow = Math.min(sheet.getMaxRows(), 1000);
  if (maxScanRow < PREP_QUEUE.dataStartRow) return;

  var totalRows = maxScanRow - PREP_QUEUE.dataStartRow + 1;
  var skuRange  = sheet.getRange(PREP_QUEUE.dataStartRow, PREP_QUEUE.cols.SKU, totalRows, 1);
  var skus      = skuRange.getValues();

  // Count occurrences (case-insensitive, trimmed)
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }

  // Build the FULL backgrounds array for the entire scan range. Rows that
  // are dupes get yellow; everything else (empty rows, single SKUs, formerly
  // duped rows whose SKU was deleted) gets explicit null. One batched write.
  var bgs = [];
  var dupeIndexes = [];
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    if (k && counts[k] >= 2) {
      bgs.push(['#fff3b0']);
      dupeIndexes.push(i);
    } else {
      bgs.push([null]);
    }
  }
  skuRange.setBackgrounds(bgs);

  // Borders: clear the full range (single batched call), then add the
  // thick yellow border per dupe (small N, typically 0-4 calls). The
  // border-clear-then-paint sequence is safe within a single execution
  // because Apps Script applies queued writes in order; we flush at the
  // end so external follow-up edits see the final state, not an
  // intermediate one.
  skuRange.setBorder(false, false, false, false, false, false, null, null);
  for (var j = 0; j < dupeIndexes.length; j++) {
    var rowIdx = dupeIndexes[j];
    sheet.getRange(PREP_QUEUE.dataStartRow + rowIdx, PREP_QUEUE.cols.SKU)
      .setBorder(true, true, true, true, false, false,
                 '#ffb800', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // Force pending writes to land before the user's next interaction —
  // without this, a fast follow-up edit can see stale formatting because
  // Sheets batches background/border changes.
  SpreadsheetApp.flush();
}
