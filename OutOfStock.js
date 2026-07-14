// =======================================================================================
// OUT_OF_STOCK.gs — "Out of Stock" sheet + alert
// =======================================================================================
//
// PURPOSE
//   A weekly tracker of every SKU in Master Inventory whose available stock
//   has dropped to zero (or negative). Lives as its own sheet so the warehouse
//   can scan it once a week, decide what to restock, and check it off — without
//   having to scroll the much larger Master Inventory sheet.
//
//   The "out of stock count" feeds the sidebar's Alerts card. Click the alert
//   row → the user is dropped onto this sheet.
//
// DEFINITION
//   Out of stock = `quantity - quantitySold <= 0` in Master Inventory.
//   This is regardless of whether the SKU has open orders committing it
//   further into negative — if Master Inventory says we have nothing left,
//   it goes on the list.
//
// SMART-MERGE REFRESH (added 2026-04-30)
//   The weekly refresh is a MERGE, not a wipe-and-rewrite. This preserves the
//   FIRST SEEN date for SKUs that have been chronically out of stock — so you
//   can see at a glance whether a SKU just hit zero or has been out for weeks.
//
//   On each refresh:
//     - SKU still OOS → update QTY/SOLD/AVAILABLE/LAST CHECKED, keep FIRST SEEN
//     - SKU newly OOS → append, FIRST SEEN = today
//     - SKU restocked (Master Inventory now > 0) → drop the row
//     - SKU not in Master Inventory at all (manual lookup, NOT FOUND) → leave alone
//
// MANUAL ENTRY
//   Type a SKU into column A → outOfStockOnEdit (installable trigger) auto-fills
//   LOCATION, QTY, SOLD, AVAILABLE, LAST CHECKED, and stamps FIRST SEEN if empty.
//   Useful as a quick stock-check tool. Manually-added rows that are currently
//   IN STOCK get cleaned out on the next weekly refresh — by design, this sheet
//   is for tracking out-of-stock items, not a watch list.
//
// ARCHITECTURE
//   - Standalone sheet ("Out of Stock") with its own minimal schema
//     (separate from the main `Schema` — different columns, different sheet).
//   - `refreshOutOfStock(maps)` smart-merges from Master Inventory.
//     Runs HOURLY during work hours (6am–6pm Houston) via runHourlyHousekeeping
//     in Housekeeping.js (replaced the weekly Monday-6am trigger 2026-07-13 —
//     workers kept forgetting the manual button between weekly runs). The
//     optional `maps` param lets housekeeping share one MI read across jobs.
//   - The "⟳ last refreshed" pulse chip lives at I1 (stamp J1) — installed by
//     _installPulseChip (Housekeeping.js), stamped at the end of every
//     completed refresh.
//   - The alert in the sidebar reads the *count* of OOS rows in this sheet — it
//     does NOT re-scan Master Inventory on every poll. Cheap.
//
// PUBLIC API
//   setupOutOfStockSheet()      — one-time: create sheet, brand styling, headers
//   refreshOutOfStock()         — scan Master Inventory, smart-merge data rows
//   openOutOfStock()            — switch the user's active sheet to Out of Stock
//   getOutOfStockCount()        — count of rows where AVAILABLE ≤ 0 (for the alert)
//   setupOutOfStockTrigger()    — install weekly Mon 6am refresh trigger
//   removeOutOfStockTrigger()   — uninstall the weekly trigger (manual cleanup)
//   outOfStockOnEdit(e)         — onEdit dispatcher (called from Main.js)
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
var OUT_OF_STOCK = {
  sheetName: "Out of Stock",

  // 1-based column positions
  cols: {
    SKU:          1,   // A
    LOCATION:     2,   // B
    QTY:          3,   // C — Master Inventory `quantity`
    SOLD:         4,   // D — Master Inventory `quantitySold`
    AVAILABLE:    5,   // E — qty - sold (will be ≤ 0 for real OOS)
    FIRST_SEEN:   6,   // F — date this SKU first appeared as OOS (preserved across refreshes)
    LAST_CHECKED: 7,   // G — timestamp of latest refresh that confirmed state
    DAYS_OUT:     8    // H — derived: TODAY() - FIRST_SEEN. ARRAYFORMULA, never written by code.
  },

  idx: function(name) { return OUT_OF_STOCK.cols[name] - 1; },

  // dataWidth      = full visible width including the formula column
  // writableWidth  = how many cols the code actually writes/reads/clears.
  //                  DAYS_OUT lives in col 8 as an ARRAYFORMULA; if we wrote
  //                  there we'd erase the formula, so refresh/onEdit/clear
  //                  only touch cols 1..7.
  dataWidth:      8,
  writableWidth:  7,
  headerRow:      1,
  dataStartRow:   2,

  headers: ["◈ SKU", "LOCATION", "# QTY", "# SOLD", "AVAILABLE", "FIRST SEEN", "LAST CHECKED", "DAYS OUT"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * One-time setup: creates "Out of Stock" sheet if missing, applies brand
 * styling, banding, and conditional formatting on the AVAILABLE column.
 * Idempotent — safe to re-run.
 */
function setupOutOfStockSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(OUT_OF_STOCK.sheetName);
  }

  // --- HEADERS ---
  sheet.getRange(OUT_OF_STOCK.headerRow, 1, 1, OUT_OF_STOCK.dataWidth)
    .setValues([OUT_OF_STOCK.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(OUT_OF_STOCK.headerRow, 1, 1, OUT_OF_STOCK.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(OUT_OF_STOCK.cols.SKU,          120);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.LOCATION,     110);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.QTY,           70);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.SOLD,          70);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.AVAILABLE,     90);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.FIRST_SEEN,   120);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.LAST_CHECKED, 140);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.DAYS_OUT,      85);

  // --- DATA AREA: column-level formats so new rows inherit ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - OUT_OF_STOCK.dataStartRow + 1;

  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.QTY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.SOLD, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.AVAILABLE, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  // FIRST SEEN + LAST CHECKED are PLAIN TEXT ('@') on purpose: the code
  // writes "M/d/yy" strings, and without this format Sheets auto-coerces
  // them into real Date values — which the next refresh read turned into
  // "Thu Apr 30 2026 08:00:00 GMT+0300 (…)" dumps that broke the DAYS OUT
  // DATEVALUE parse (bug fixed 2026-07-13; see _normalizeOosFirstSeen).
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.FIRST_SEEN, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center')
    .setNumberFormat('@');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.LAST_CHECKED, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center')
    .setNumberFormat('@');
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.DAYS_OUT, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setFontColor('#1d1d1b').setHorizontalAlignment('center');

  sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, dataRows, OUT_OF_STOCK.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING (cream alternation) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, OUT_OF_STOCK.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- AVAILABLE: red when ≤ 0 (CF rule) ---
  // Every populated row gets it; the rule keeps the visual unambiguous even if
  // someone hand-types an in-stock SKU.
  // Strip any prior rules that target AVAILABLE (this rule) or SKU (legacy
  // CF-based dupe rule from a prior version — duplicates are now JS-painted).
  var existingRules = sheet.getConditionalFormatRules();
  var keptRules = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) {
      return rg.getColumn() === OUT_OF_STOCK.cols.AVAILABLE
          || rg.getColumn() === OUT_OF_STOCK.cols.SKU;
    });
  });
  var availRange = sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.AVAILABLE, dataRows, 1);
  var availRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      "=AND(ISNUMBER(E" + OUT_OF_STOCK.dataStartRow + "), E" + OUT_OF_STOCK.dataStartRow + "<=0)"
    )
    .setBackground('#ff6b6b').setFontColor('#ffffff').setBold(true)
    .setRanges([availRange])
    .build();
  keptRules.push(availRule);
  sheet.setConditionalFormatRules(keptRules);

  // --- Duplicate SKU highlight ---
  // Done JS-side (not via CF) — Apps Script CF formulas with COUNTIF on a
  // 1000-row range have been unreliable in this codebase. This mirrors the
  // setupDuplicateSalesOrderHighlighting pattern: scan column A, draw a thick
  // amber border on any SKU that appears more than once.
  _refreshOutOfStockDuplicates(sheet);

  // --- DAYS OUT formula (col H) ---
  // Single ARRAYFORMULA in H2 spills into the rest of the column. TODAY()
  // recalculates daily, so DAYS OUT is always accurate without rerunning the
  // refresh. IFERROR guards against unparseable FIRST SEEN strings.
  // ISNUMBER branch (added 2026-07-13): a FIRST SEEN cell that Sheets
  // already coerced into a real Date is a serial number — DATEVALUE errors
  // on those, so use the value directly; text dates go through DATEVALUE.
  // Clear the column first so the spill isn't blocked by any stale values.
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.DAYS_OUT, dataRows, 1).clearContent();
  sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.DAYS_OUT).setFormula(
    "=ARRAYFORMULA(IF(F" + OUT_OF_STOCK.dataStartRow + ":F=\"\", \"\", " +
    "IFERROR(TODAY()-IF(ISNUMBER(F" + OUT_OF_STOCK.dataStartRow + ":F), " +
    "F" + OUT_OF_STOCK.dataStartRow + ":F, " +
    "DATEVALUE(F" + OUT_OF_STOCK.dataStartRow + ":F)), \"\")))"
  );

  // --- FREEZE HEADER ROW ---
  sheet.setFrozenRows(1);

  // --- FRESHNESS PULSE CHIP (I1 chip / J1 stamp, right of the headers) ---
  try { _installPulseChip(sheet, SHEET_PULSE.outOfStock); }
  catch (e) { try { Logger.log("setupOutOfStockSheet: pulse chip error: " + e); } catch (_) {} }

  return "✅ Out of Stock sheet ready.";
}


/**
 * Normalize a FIRST SEEN cell value to the canonical "M/d/yy" string.
 *
 * WHY (bug found 2026-07-13, surfaced by the hourly refresh): Sheets
 * auto-coerces written "4/30/26" strings into real Date values. The next
 * refresh read then got a Date object, and String(date) produced the JS
 * dump "Thu Apr 30 2026 08:00:00 GMT+0300 (GMT+03:00)" — which got written
 * back, and which DATEVALUE can't parse, so the DAYS OUT column went blank
 * for every row. This helper repairs all three shapes on read:
 *   real Date            → format as M/d/yy
 *   legacy JS date-dump  → re-parse, format as M/d/yy
 *   clean "M/d/yy" text  → passthrough
 * Belt-and-suspenders with the plain-text ('@') format setupOutOfStockSheet
 * now applies to the FIRST SEEN / LAST CHECKED columns, which stops the
 * coercion at the source.
 */
function _normalizeOosFirstSeen(val, fallback) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "America/Chicago", "M/d/yy");
  }
  var s = String(val || "").trim();
  if (!s) return fallback;
  if (/GMT[+-]/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, "America/Chicago", "M/d/yy");
  }
  return s;
}


/**
 * Smart-merge refresh from Master Inventory. Preserves FIRST SEEN dates so
 * chronic out-of-stock items can be identified at a glance.
 *
 * Logic:
 *   - Build the OOS set from Master Inventory (qty - sold ≤ 0)
 *   - Index existing sheet rows by SKU
 *   - Walk existing rows and decide:
 *       a) SKU still OOS → KEEP, refresh QTY/SOLD/AVAILABLE/LAST CHECKED, preserve FIRST SEEN
 *       b) SKU restocked (Master Inventory now > 0) → DROP
 *       c) SKU not in Master Inventory at all (NOT FOUND / typo) → KEEP as-is
 *   - For each newly-OOS SKU not seen above → APPEND with FIRST SEEN = today
 *   - Re-sort by LOCATION asc → SKU asc
 *   - Write back
 */
function refreshOutOfStock(maps) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) {
    setupOutOfStockSheet();
    sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  }

  // Accept pre-built maps from runHourlyHousekeeping (shared MI read).
  // Detect by shape, not presence — the legacy weekly trigger (and any
  // future direct trigger) passes a time-event object as the first arg.
  if (!maps || !maps.inventoryMap) maps = buildLocationAndInventoryMaps();
  if (maps.inventoryMap.size === 0) {
    return "⚠️ Master Inventory empty or headers missing.";
  }

  var todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy");
  var nowStr   = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

  // ---- Read existing rows ----
  // Read only the writable columns (1..7). Column 8 (DAYS_OUT) is a formula
  // that we must not touch.
  var lastRow = sheet.getLastRow();
  var existingRows = [];
  if (lastRow >= OUT_OF_STOCK.dataStartRow) {
    existingRows = sheet.getRange(
      OUT_OF_STOCK.dataStartRow, 1,
      lastRow - OUT_OF_STOCK.dataStartRow + 1,
      OUT_OF_STOCK.writableWidth
    ).getValues();
  }

  // ---- Build merged row set ----
  var keptSkus = {};   // skus already accounted for from existing rows
  var merged = [];     // final rows array
  var dropped = 0;
  var refreshed = 0;
  var preserved = 0;   // manually-added NOT FOUND rows kept as-is

  for (var i = 0; i < existingRows.length; i++) {
    var row = existingRows[i];
    var rawSku = String(row[OUT_OF_STOCK.idx("SKU")] || "").trim();
    if (!rawSku) continue;  // skip empty rows
    var skuLower = rawSku.toLowerCase();

    var inv = maps.inventoryMap.get(skuLower);

    if (!inv) {
      // (c) SKU not in Master Inventory — leave row exactly as the user wrote it
      merged.push(row);
      keptSkus[skuLower] = true;
      preserved++;
      continue;
    }

    if (inv.available > 0) {
      // SKU is currently in stock per Master Inventory. Two sub-cases —
      // distinguished by what AVAILABLE was already in the sheet:
      var sheetAvail = row[OUT_OF_STOCK.idx("AVAILABLE")];
      if (typeof sheetAvail === 'number' && sheetAvail <= 0) {
        // Sheet treated this as OOS (auto-added by a prior refresh, or hand-
        // typed when it was zero). Master Inventory now says it's restocked.
        // Drop the row — the OOS list shouldn't show in-stock items.
        dropped++;
        continue;
      }
      // Sheet's own AVAILABLE was already > 0 — i.e., the user manually
      // typed a SKU that was in stock at the time, presumably as a watch
      // or lookup. Respect that: refresh values, preserve FIRST SEEN, keep.
      var locationKept  = maps.locationMap.get(skuLower) || row[OUT_OF_STOCK.idx("LOCATION")] || "";
      var firstSeenKept = _normalizeOosFirstSeen(row[OUT_OF_STOCK.idx("FIRST_SEEN")], todayStr);
      merged.push([
        rawSku,
        locationKept,
        inv.quantity,
        inv.sold,
        inv.available,
        firstSeenKept,
        nowStr
      ]);
      keptSkus[skuLower] = true;
      preserved++;
      continue;
    }

    // (a) Still OOS — refresh values, preserve FIRST SEEN
    var location  = maps.locationMap.get(skuLower) || row[OUT_OF_STOCK.idx("LOCATION")] || "";
    var firstSeen = _normalizeOosFirstSeen(row[OUT_OF_STOCK.idx("FIRST_SEEN")], todayStr);

    merged.push([
      rawSku,
      location,
      inv.quantity,
      inv.sold,
      inv.available,
      firstSeen,
      nowStr
    ]);
    keptSkus[skuLower] = true;
    refreshed++;
  }

  // ---- Append newly-OOS SKUs not already represented ----
  var added = 0;
  maps.inventoryMap.forEach(function(inv, skuLower) {
    if (inv.available > 0) return;
    if (keptSkus[skuLower]) return;

    var location = maps.locationMap.get(skuLower) || "";
    merged.push([
      skuLower,
      location,
      inv.quantity,
      inv.sold,
      inv.available,
      todayStr,   // FIRST SEEN
      nowStr      // LAST CHECKED
    ]);
    added++;
  });

  // ---- Sort: LOCATION asc (empty last), then SKU asc ----
  merged.sort(function(a, b) {
    var la = String(a[OUT_OF_STOCK.idx("LOCATION")] || "");
    var lb = String(b[OUT_OF_STOCK.idx("LOCATION")] || "");
    if (la === "" && lb !== "") return 1;
    if (lb === "" && la !== "") return -1;
    if (la !== lb) return la.localeCompare(lb);
    return String(a[OUT_OF_STOCK.idx("SKU")]).localeCompare(String(b[OUT_OF_STOCK.idx("SKU")]));
  });

  // ---- Wipe and write ----
  // Only touch writable cols (1..7). Leaving col 8 (DAYS_OUT formula) alone.
  if (lastRow >= OUT_OF_STOCK.dataStartRow) {
    sheet.getRange(
      OUT_OF_STOCK.dataStartRow, 1,
      lastRow - OUT_OF_STOCK.dataStartRow + 1,
      OUT_OF_STOCK.writableWidth
    ).clearContent();
  }

  if (merged.length > 0) {
    sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, merged.length, OUT_OF_STOCK.writableWidth)
      .setValues(merged);
  }

  // Refresh duplicate-SKU highlights (after data has been written)
  _refreshOutOfStockDuplicates(sheet);

  // Freshness chip — stamped only on a COMPLETED refresh, so the chip's
  // staleness tiers double as the "refresh trigger is dead" alarm.
  stampSheetPulse(sheet, SHEET_PULSE.outOfStock.stamp);

  return "✅ Out of Stock refreshed — " +
         added + " new, " +
         refreshed + " still out, " +
         dropped + " restocked" +
         (preserved > 0 ? ", " + preserved + " manual" : "") + ".";
}


/**
 * Sidebar: switch the user's active view to Out of Stock.
 * Same pattern as openPrepQueue — must use SpreadsheetApp.getActive() (the
 * BOUND spreadsheet) for setActiveSheet to actually switch the user's tab.
 */
function openOutOfStock() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet (open this from the spreadsheet UI).";

  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) {
    setupOutOfStockSheet();
    sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  }

  ss.setActiveSheet(sheet);
  return "✅ Opened " + OUT_OF_STOCK.sheetName;
}


/**
 * Fast count for the sidebar alert: counts only rows where AVAILABLE ≤ 0
 * (i.e. genuinely out of stock). Manual lookups for in-stock SKUs (which the
 * onEdit handler may have populated with AVAILABLE > 0) and NOT FOUND rows
 * are deliberately excluded.
 *
 * Reads the snapshot — does NOT re-scan Master Inventory.
 */
function getOutOfStockCount() {
  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < OUT_OF_STOCK.dataStartRow) return 0;

  // Read SKU + AVAILABLE in one range (cols A:E, dropping the rightmost two)
  var data = sheet.getRange(
    OUT_OF_STOCK.dataStartRow, 1,
    lastRow - OUT_OF_STOCK.dataStartRow + 1,
    OUT_OF_STOCK.cols.AVAILABLE
  ).getValues();

  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][OUT_OF_STOCK.idx("SKU")]).trim();
    if (!sku) continue;
    var avail = data[i][OUT_OF_STOCK.idx("AVAILABLE")];
    if (typeof avail === 'number' && avail <= 0) count++;
  }
  return count;
}


// =======================================================================================
// TRIGGER MANAGEMENT
// =======================================================================================

/**
 * ⚠️ SUPERSEDED 2026-07-13 — the weekly Monday-6am cadence was replaced by the
 * hourly work-hours pass in Housekeeping.js (runHourlyHousekeeping refreshes
 * this sheet AND Prep Queue locations from one shared MI read). Run
 * setupHousekeeping() instead — it also REMOVES any trigger this function
 * installed. Kept only as a manual fallback if the housekeeping layer is
 * ever ripped out. Do not run both: refreshOutOfStock is idempotent so a
 * double schedule wouldn't corrupt data, but it doubles the quota burn.
 *
 * (Original behavior: installs a weekly Monday 6am refreshOutOfStock trigger;
 * idempotent — removes any existing refreshOutOfStock trigger first.)
 */
function setupOutOfStockTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshOutOfStock') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('refreshOutOfStock')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  Logger.log("Weekly Out of Stock trigger installed: Monday 6am");

  try {
    SpreadsheetApp.getUi().alert(
      "Trigger Installed",
      "Out of Stock will refresh automatically every Monday at 6am.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Trigger installed. (No UI context for alert)");
  }
}

/** Removes the weekly refresh trigger. Manual cleanup helper. */
function removeOutOfStockTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshOutOfStock') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log("Removed " + removed + " refreshOutOfStock trigger(s).");
  return "Removed " + removed + " trigger(s).";
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEditInstallable(e)
// =======================================================================================

/**
 * SKU edit on Out of Stock → auto-fill LOCATION + QTY + SOLD + AVAILABLE +
 * stamp FIRST SEEN (only if currently empty) + LAST CHECKED.
 *
 * Mirrors prepQueueOnEdit. Lives on the INSTALLABLE trigger because the
 * Master Inventory lookup goes through openById which simple triggers can't
 * call reliably.
 *
 * Defensive try/catch — never blocks other edit handlers.
 */
function outOfStockOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== OUT_OF_STOCK.sheetName) return;
    if (e.range.getColumn() !== OUT_OF_STOCK.cols.SKU) return;
    if (e.range.getRow() < OUT_OF_STOCK.dataStartRow) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();

    // Pre-build inventory + location maps if multi-row paste, otherwise
    // single-row lookups are cheaper.
    var useMap = edits.length > 3;
    var locationMap = null;
    var inventoryMap = null;
    if (useMap) {
      var maps = buildLocationAndInventoryMaps();
      locationMap = maps.locationMap;
      inventoryMap = maps.inventoryMap;
    }

    var todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy");
    var nowStr   = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe the row's lookup fields so it reads as empty
        sheet.getRange(row, OUT_OF_STOCK.cols.LOCATION).setValue("");
        sheet.getRange(row, OUT_OF_STOCK.cols.QTY).setValue("");
        sheet.getRange(row, OUT_OF_STOCK.cols.SOLD).setValue("");
        sheet.getRange(row, OUT_OF_STOCK.cols.AVAILABLE).setValue("");
        sheet.getRange(row, OUT_OF_STOCK.cols.FIRST_SEEN).setValue("");
        sheet.getRange(row, OUT_OF_STOCK.cols.LAST_CHECKED).setValue("");
        continue;
      }

      // --- Resolve LOCATION + INVENTORY ---
      var location, inv;
      if (useMap) {
        location = locationMap.get(skuLower) || "NOT FOUND";
        inv = inventoryMap.get(skuLower);
      } else {
        location = getSingleLocation(skuLower);
        inv = getSingleInventory(skuLower);
      }

      var found = (location !== "NOT FOUND");
      var qty   = (found && inv) ? inv.quantity  : "";
      var sold  = (found && inv) ? inv.sold      : "";
      var avail = (found && inv) ? inv.available : "";

      // Preserve existing FIRST SEEN if non-empty (normalized — see
      // _normalizeOosFirstSeen); otherwise stamp today
      var firstSeen = _normalizeOosFirstSeen(
        sheet.getRange(row, OUT_OF_STOCK.cols.FIRST_SEEN).getValue(), todayStr);

      sheet.getRange(row, OUT_OF_STOCK.cols.LOCATION).setValue(location);
      sheet.getRange(row, OUT_OF_STOCK.cols.QTY).setValue(qty);
      sheet.getRange(row, OUT_OF_STOCK.cols.SOLD).setValue(sold);
      sheet.getRange(row, OUT_OF_STOCK.cols.AVAILABLE).setValue(avail);
      sheet.getRange(row, OUT_OF_STOCK.cols.FIRST_SEEN).setValue(firstSeen);
      sheet.getRange(row, OUT_OF_STOCK.cols.LAST_CHECKED).setValue(nowStr);
    }

    // Refresh duplicate highlights — every SKU edit could create or clear a dupe
    _refreshOutOfStockDuplicates(sheet);
  } catch (err) {
    try { Logger.log("outOfStockOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (JS-side, mirrors the All Orders pattern)
// =======================================================================================

/**
 * Scans column A, identifies SKUs that appear more than once, and paints the
 * SKU cell with a soft-amber background + thick yellow border so duplicates
 * stand out from both the cream banding and the red AVAILABLE highlight.
 *
 * Cleared cells (no SKU) get reset to default. Run after any change that
 * could affect column A: setupOutOfStockSheet, refreshOutOfStock, outOfStockOnEdit.
 */
function _refreshOutOfStockDuplicates(sheet) {
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < OUT_OF_STOCK.dataStartRow) return;

  var dataRows = lastRow - OUT_OF_STOCK.dataStartRow + 1;
  var skuRange = sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.SKU, dataRows, 1);
  var skus = skuRange.getValues();

  // Count occurrences (case-insensitive, trimmed)
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }

  // Walk every row and explicitly set its state — paint if dupe, else clear.
  // Per-cell explicit clears (rather than a range-wide reset followed by
  // selective paint) are more reliable: when a SKU is deleted and only its
  // counterpart remains, the counterpart's row gets an explicit clear in this
  // same pass instead of relying on the range reset to fully take effect
  // before the conditional paint.
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    var cell = sheet.getRange(OUT_OF_STOCK.dataStartRow + i, OUT_OF_STOCK.cols.SKU);
    if (k && counts[k] >= 2) {
      cell.setBackground('#fff3b0');
      cell.setBorder(true, true, true, true, false, false,
                     '#ffb800', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else {
      cell.setBackground(null);
      cell.setBorder(false, false, false, false, false, false, null, null);
    }
  }

  // Force pending writes to flush immediately. Without this, a fast follow-up
  // edit could see stale formatting because Sheets batches background/border
  // changes and the user sees the cell render before the batch lands.
  SpreadsheetApp.flush();
}
