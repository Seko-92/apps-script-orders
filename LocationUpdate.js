// =======================================================================================
// LOCATION_UPDATE.gs — "Location Update" sheet for tracking SKU→Location changes////
// =======================================================================================
//
// PURPOSE
//   Warehouse employees use this sheet to record SKUs whose physical location
//   changed. They type an SKU in column B; the system auto-fills:
//     - Col A: sequential counter (for the eye — "how many have I done today")
//     - Col E: timestamp in Houston time
//   The picker types the NEW location in column C manually (per 2026-05-14
//   user decision — KISS over auto-fill).
//
//   A second person later verifies the work: since 2026-08-11 they press
//   "⇊ Fetch eBay Locations" in the sidebar instead of opening eBay item by
//   item. Column D (◉ EBAY LOC) fills with each item's LIVE location as eBay
//   reports it right now, and the cell renders a VERDICT:
//     - quiet green  = eBay agrees with the typed LOCATION
//     - loud red     = eBay disagrees (or the listing has no location at all)
//     - dim italic   = SKU not found on eBay (usually a typo'd SKU) or the
//                      live check failed for that item
//   The checker scans for red, fixes what needs fixing, signs FINAL CHECK BY.
//
// LIVE FETCH ARCHITECTURE (2026-08-11)
//   Sidebar button → fetchEbayLocationsLive() → resolves each SKU's eBay
//   itemId from Master Inventory → POSTs the batch to the n8n
//   "eBay Location Check Proxy" (webhook: ebay-location-check) → n8n calls
//   Trading API GetItem per item and extracts the "Model Year" item specific
//   (the field this business uses for physical location — same field MAIN's
//   sync mirrors into MI's "C:Model Year" column) → responds synchronously →
//   we write values + verdicts in one batched pass.
//
//   TRULY LIVE, not the hourly mirror: the user chose the live call so a
//   check minutes after an eBay-side edit can never show a stale mismatch.
//   Cost: 1 GetItem per unique SKU per fetch (Trading API pool 5,000/day —
//   a 50-row sheet costs 1% of the pool).
//
// ARCHITECTURE — INSTALLABLE TRIGGER ONLY
//   The original implementation ran in the SIMPLE onEdit trigger
//   (locationUpdateTimestamp in Timestampfeature.js, now orphaned). Simple
//   triggers can fail silently when openById is called and permissions aren't
//   granted — that's why the location lookup and timestamp would sometimes
//   "fail to appear." This version runs in onEditInstallable, which has full
//   permissions. Same pattern as prepQueueOnEdit and outOfStockOnEdit.
//
// PUBLIC API
//   setupLocationUpdateSheet()       — one-time: brand-theme, headers, banding
//                                       (idempotent v3→v4 column migrator inside)
//   openLocationUpdate()             — sidebar: switch active sheet
//   refreshLocationUpdateSheet()     — sidebar: sweep all rows, fill any blanks
//                                       (manual escape hatch when auto-fill missed)
//   sortLocationUpdateByLocation()   — sidebar: sort data rows by LOCATION A→Z
//   fetchEbayLocationsLive()         — sidebar: live eBay location check, all rows
//   locationUpdateOnEdit(e)          — installable trigger dispatcher
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
//
// Schema v4 (2026-08-11): ◉ EBAY LOC inserted at column D, directly beside the
// typed LOCATION so the compare reads side-by-side. TIMESTAMP/EMPLOYEE/
// FINAL_CHECK_BY shifted one column right; setupLocationUpdateSheet migrates a
// v3 sheet in place (existing data survives the insert).
//
// Schema v3 (2026-05-14): user-requested simplification — LOCATION is
// MANUALLY edited by the picker (no auto-fill from Master Inventory). Only
// TIMESTAMP is auto-stamped (on SKU edit).
var LOCATION_UPDATE = {
  sheetName: "Location Update",

  // 1-based column positions
  cols: {
    COUNTER:        1,   // A — formula-driven row label (=ROW()-1), always visible
    SKU:            2,   // B — manually typed (rich-text link to the eBay listing)
    LOCATION:       3,   // C — manually typed (no auto-fill)
    EBAY_LOC:       4,   // D — machine-written by fetchEbayLocationsLive() ONLY
    TIMESTAMP:      5,   // E — auto-stamped at SKU edit time
    EMPLOYEE:       6,   // F — dropdown, user-selected
    FINAL_CHECK_BY: 7    // G — dropdown, user-selected
  },

  idx: function(name) { return LOCATION_UPDATE.cols[name] - 1; },

  dataWidth:    7,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["#", "◈ SKU", "LOCATION", "◉ EBAY LOC", "⏱ TIMESTAMP", "👤 EMPLOYEE", "✓ FINAL CHECK BY"],

  // Cap on how many UNIQUE items one fetch sends to n8n. The proxy loops a
  // GetItem per item; a synchronous round-trip must stay well inside
  // UrlFetchApp's timeout. Enforced HERE and in the proxy's validate node —
  // a guard that lives only in the caller is a guard a future caller forgets.
  fetchCap: 120,

  // Sentinel values the fetch writes into ◉ EBAY LOC when there is no real
  // location to show. _luVerdict() recognizes these by exact value — rename
  // them HERE only; everything else reads this table.
  sentinels: {
    notOnEbay:   "NOT ON EBAY",    // SKU has no MI row / no itemId → likely a typo'd SKU
    checkFailed: "CHECK FAILED",   // GetItem errored for this item (ended long ago, eBay 5xx…)
    blankOnEbay: "BLANK ON EBAY"   // listing exists but its location field is empty
  }
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * One-time setup: creates "Location Update" sheet if missing, applies brand
 * styling matching Prep Queue and Out of Stock. Idempotent — safe to re-run.
 *
 * v3 → v4 MIGRATION (2026-08-11): if the sheet still has the old layout
 * (⏱ TIMESTAMP at column D), a fresh ◉ EBAY LOC column is inserted before it.
 * insertColumnBefore shifts values, validations and formats together, so
 * existing rows + the EMPLOYEE / FINAL CHECK BY dropdowns survive intact.
 *
 * NOTE: previous versions of this sheet may have used a 2-row header (title
 * + column labels). This function writes the single-row brand header to row 1
 * and styles row 2+ as the data area. Existing data in row 3+ stays where it
 * is; if a user wants to compact it up, that's a manual fix.
 */
function setupLocationUpdateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(LOCATION_UPDATE.sheetName);
  }

  // --- v3 → v4 MIGRATOR: insert ◉ EBAY LOC before the old TIMESTAMP col D ---
  // Detection is by header VALUE, not position math, so re-running is a no-op
  // on an already-migrated sheet and harmless on a brand-new blank one.
  var oldD = String(sheet.getRange(LOCATION_UPDATE.headerRow, 4).getValue() || "");
  if (oldD.indexOf("TIMESTAMP") !== -1) {
    sheet.insertColumnBefore(4);
  }

  // --- HEADERS ---
  sheet.getRange(LOCATION_UPDATE.headerRow, 1, 1, LOCATION_UPDATE.dataWidth)
    .setValues([LOCATION_UPDATE.headers])
    .setBackground('#1d1d1b')   // brand black
    .setFontColor('#ffd966')    // brand yellow
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Thick yellow underline below header
  sheet.getRange(LOCATION_UPDATE.headerRow, 1, 1, LOCATION_UPDATE.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(LOCATION_UPDATE.cols.COUNTER,         55);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.SKU,            130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.LOCATION,       130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.EBAY_LOC,       130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.TIMESTAMP,      170);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.EMPLOYEE,       130);
  sheet.setColumnWidth(LOCATION_UPDATE.cols.FINAL_CHECK_BY, 150);

  // --- DATA AREA: column-level format (so new rows inherit) ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - LOCATION_UPDATE.dataStartRow + 1;

  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11)
    .setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  // ◉ EBAY LOC: machine-written value — mono like LOCATION but NOT bold, so
  // eBay's answer reads as an annotation beside the human entry, not a twin.
  // Per-cell verdict colors (green/red/gray) are painted by the fetch on top.
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EBAY_LOC, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('normal').setFontSize(10)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.TIMESTAMP, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
    .setFontFamily('Roboto').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
    .setFontFamily('Roboto').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  sheet.getRange(LOCATION_UPDATE.dataStartRow, 1, dataRows, LOCATION_UPDATE.dataWidth)
    .setVerticalAlignment('middle');

  // --- COUNTER FORMULA (col A) ---
  // =ROW()-1 in every data row. Always reflects current row position; survives
  // SKU clear (never disappears); auto-corrects on row insert/delete. Per-cell
  // formula (not ArrayFormula) so deleting one cell doesn't blank the column.
  // The refresh button re-paints these in case a cell got cleared accidentally.
  var counterFormulas = [];
  for (var cf = 0; cf < dataRows; cf++) counterFormulas.push(["=ROW()-1"]);
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, dataRows, 1)
    .setFormulas(counterFormulas);

  // --- DATA VALIDATION on Employee + Final Check By (dropdowns) ---
  // Placeholder list so the dropdown widget renders immediately. User edits
  // the list via Data → Data validation, OR calls setLocationUpdateDropdowns()
  // from the Apps Script editor with the real staff names.
  //
  // setAllowInvalid(true) means the cell accepts any typed value while the
  // list is still the placeholder — so warehouse staff aren't blocked before
  // the lists are configured.
  //
  // IMPORTANT: only install the placeholder if validation isn't already set on
  // these columns. Otherwise re-running setup would wipe out the user's
  // configured staff lists. We check the first data cell as a proxy for
  // "has this column been configured yet." (The v4 column insert shifts the
  // configured rules right along with their cells, so this stays correct.)
  var placeholderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["—"], true)
    .setAllowInvalid(true)
    .setHelpText("Edit this dropdown's list via Data → Data validation, or run setLocationUpdateDropdowns() from the script editor.")
    .build();
  var empProbe = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE).getDataValidation();
  if (!empProbe) {
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
      .setDataValidation(placeholderRule);
  }
  var fcProbe = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY).getDataValidation();
  if (!fcProbe) {
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
      .setDataValidation(placeholderRule);
  }

  // --- BANDING (cream alternation, brand-consistent) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, LOCATION_UPDATE.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- FREEZE HEADER ROW ---
  sheet.setFrozenRows(1);

  // Paint any existing duplicate SKUs (idempotent — clears stale highlights too)
  _refreshLocationUpdateDuplicates(sheet);

  return "✅ Location Update sheet ready.";
}


/**
 * v3 → v4 AUTO-MIGRATION GUARD (the Kit Health pattern: never write into an
 * old-shaped sheet; the one-cell probe is cheap). Every writer calls this
 * before touching column positions, so the column insert happens on the FIRST
 * touch after deploy — no manual Re-style step required, and the deploy-order
 * hazard (new code live via clasp push, sheet still v3, TIMESTAMP landing in
 * the old EMPLOYEE column) cannot occur.
 */
function _luEnsureV4(sheet) {
  if (!sheet) return;
  var d1 = String(sheet.getRange(LOCATION_UPDATE.headerRow, 4).getValue() || "");
  if (d1.indexOf("TIMESTAMP") !== -1) {
    setupLocationUpdateSheet();   // idempotent — performs the column insert + restyle
  }
}


/**
 * Sidebar: switch the user's active view to the Location Update sheet.
 *
 * Uses SpreadsheetApp.getActive() (BOUND spreadsheet), not openById().
 * setActiveSheet() only changes the visible tab when called on the active
 * spreadsheet reference. See PrepQueue.openPrepQueue for the same pattern
 * and the v1 bug history that established it.
 */
function openLocationUpdate() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet (open this from the spreadsheet UI).";

  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) {
    setupLocationUpdateSheet();
    sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  }

  ss.setActiveSheet(sheet);
  return "✅ Opened " + LOCATION_UPDATE.sheetName;
}


/**
 * Sidebar: sweep all data rows, fill any blank TIMESTAMPs where SKU exists,
 * and ensure the COUNTER formula is intact on every row. The manual escape
 * hatch — used when the installable trigger missed an edit, OR when a
 * counter cell got accidentally cleared.
 *
 * Behavior per row:
 *   - SKU empty → blank out LOCATION + EBAY LOC + TIMESTAMP. COUNTER stays.
 *   - SKU present + TIMESTAMP blank → stamp with "now" (best we can do — the
 *     original edit time isn't recoverable; an approximate stamp beats blank)
 *   - LOCATION → NEVER touched by refresh (manually edited per 2026-05-14 decision)
 *   - EBAY LOC → NEVER written for live rows (only fetchEbayLocationsLive owns it)
 *   - COUNTER formula always re-written (idempotent: if formula was deleted
 *     accidentally, it's restored)
 *   - EMPLOYEE + FINAL_CHECK_BY untouched (out of read range)
 *
 * Returns a status string the sidebar can show.
 */
function refreshLocationUpdateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";
  _luEnsureV4(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < LOCATION_UPDATE.dataStartRow) {
    return "ℹ️ No data rows to refresh.";
  }

  var nRows = lastRow - LOCATION_UPDATE.dataStartRow + 1;
  // Read columns B-E (SKU, LOCATION, EBAY_LOC, TIMESTAMP).
  // Col A is formula-driven (handled separately).
  // Cols F/G are user input (EMPLOYEE, FINAL_CHECK_BY) — never touched here.
  // LOCATION + EBAY_LOC are read so we can clear them when SKU is empty, but
  // never written when SKU is present (LOCATION is manual; EBAY_LOC belongs
  // to the live fetch).
  var workRange = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, nRows, 4);
  var workData = workRange.getValues();

  var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy h:mm a");

  var timestampsFilled = 0;
  var clearedEbayRows = [];

  // Local indices into workData (cols B/C/D/E → 0/1/2/3 in this slice)
  var SKU_I = 0, LOC_I = 1, EBAY_I = 2, TS_I = 3;

  for (var i = 0; i < workData.length; i++) {
    var sku = String(workData[i][SKU_I] || "").trim();
    var existingTimestamp = String(workData[i][TS_I] || "").trim();

    if (!sku) {
      // Empty SKU — clear LOCATION + EBAY_LOC + TIMESTAMP. The whole row is
      // one paired audit entry; losing the SKU loses the row's meaning.
      if (String(workData[i][EBAY_I] || "").trim()) {
        clearedEbayRows.push(LOCATION_UPDATE.dataStartRow + i);
      }
      workData[i][LOC_I] = "";
      workData[i][EBAY_I] = "";
      workData[i][TS_I] = "";
      continue;
    }

    // SKU present — only restore the timestamp if blank. LOCATION stays as
    // whatever the picker typed; EBAY_LOC stays as whatever the last live
    // fetch wrote (or blank if never fetched).
    if (!existingTimestamp) {
      workData[i][TS_I] = nowStr;
      timestampsFilled++;
    }
  }

  // One batched write for SKU + LOCATION + EBAY_LOC + TIMESTAMP columns
  workRange.setValues(workData);

  // ⚠ setValues just STRIPPED the rich-text listing links off col B (the
  // updateAllExistingRows lesson, 2026-06-03: any full-column setValues must
  // re-apply the SKU links). Best-effort — a missing MI never blocks refresh.
  try {
    var linkMap = buildSkuEnrichmentMap();
    if (linkMap.size) {
      applySkuLinksToColumn(sheet, LOCATION_UPDATE.cols.SKU,
                            LOCATION_UPDATE.dataStartRow, lastRow, linkMap);
    }
  } catch (linkErr) {
    console.log("refreshLocationUpdateSheet: link re-apply failed (non-fatal): " + linkErr);
  }

  // Rows whose EBAY_LOC we just blanked also need their verdict COLOR wiped —
  // a green/red background with no value in it reads as a rendering fault.
  for (var c = 0; c < clearedEbayRows.length; c++) {
    _luResetEbayLocCell(sheet.getRange(clearedEbayRows[c], LOCATION_UPDATE.cols.EBAY_LOC));
  }

  // Re-paint COUNTER formula on every data row. Idempotent — if formula was
  // already there, this is a no-op visually; if it was deleted accidentally,
  // it's restored. Use setFormulas so each cell gets =ROW()-1 individually.
  var counterFormulas = [];
  for (var cf = 0; cf < nRows; cf++) counterFormulas.push(["=ROW()-1"]);
  sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.COUNTER, nRows, 1)
    .setFormulas(counterFormulas);

  // Refresh duplicate highlighting too — defensive (no-op if nothing changed)
  _refreshLocationUpdateDuplicates(sheet);

  return "✅ Refreshed: " + timestampsFilled + " timestamp(s) restored. Counter formulas re-applied.";
}


/**
 * Sidebar: sort the data rows by LOCATION column A→Z.
 *
 * Sorts columns B–G (SKU … FINAL_CHECK_BY) only. Column A (COUNTER = ROW()-1
 * formula) is intentionally NOT included in the sort range — its per-row
 * formula always evaluates against its own row, so the numbering 1, 2, 3, ...
 * stays in order while the data shuffles beneath it.
 *
 * range.sort() is a REAL Sheets sort — values, per-cell verdict colors on
 * ◉ EBAY LOC, and the SKU rich-text links all travel with their rows (unlike
 * the getValues/setValues sorts elsewhere that need the paired-array carry).
 *
 * Empty-LOCATION rows naturally land at the bottom — useful cue that those
 * entries are still in-progress (picker typed SKU but hasn't filled LOCATION).
 *
 * After the sort, row positions have changed — re-paint the duplicate-SKU
 * highlight so amber/yellow lands on the right cells.
 */
function sortLocationUpdateByLocation() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";
  _luEnsureV4(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < LOCATION_UPDATE.dataStartRow) {
    return "ℹ️ No data rows to sort.";
  }

  var nRows = lastRow - LOCATION_UPDATE.dataStartRow + 1;

  var sortRange = sheet.getRange(
    LOCATION_UPDATE.dataStartRow,
    LOCATION_UPDATE.cols.SKU,             // B (col 2)
    nRows,
    LOCATION_UPDATE.dataWidth - 1         // 6 cols → B through G
  );

  sortRange.sort({ column: LOCATION_UPDATE.cols.LOCATION, ascending: true });

  _refreshLocationUpdateDuplicates(sheet);
  SpreadsheetApp.flush();

  return "✅ Sorted Location Update by LOCATION A→Z.";
}


/**
 * Sidebar/editor helper: install real dropdown values for Employee and Final
 * Check By columns. Called once after setup, or any time the staff list changes.
 *
 * Usage from the script editor:
 *   setLocationUpdateDropdowns(["Alice", "Bob"], ["Carol", "Dan"]);
 *
 * Passing an empty array for either argument leaves that column's dropdown
 * alone (so you can update just one list).
 *
 * Strict mode: once you call this with real names, the dropdown rejects invalid
 * values (setAllowInvalid(false)) — typos get flagged. Different from the
 * placeholder rule installed by setupLocationUpdateSheet, which allows any text.
 */
function setLocationUpdateDropdowns(employees, finalCheckers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";

  var maxDataRow = 1000;
  var dataRows = maxDataRow - LOCATION_UPDATE.dataStartRow + 1;
  var updates = [];

  if (Array.isArray(employees) && employees.length > 0) {
    var empRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(employees, true)
      .setAllowInvalid(false)
      .setHelpText("Re-run setLocationUpdateDropdowns() to update this list.")
      .build();
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EMPLOYEE, dataRows, 1)
      .setDataValidation(empRule);
    updates.push("Employee (" + employees.length + ")");
  }

  if (Array.isArray(finalCheckers) && finalCheckers.length > 0) {
    var fcRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(finalCheckers, true)
      .setAllowInvalid(false)
      .setHelpText("Re-run setLocationUpdateDropdowns() to update this list.")
      .build();
    sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.FINAL_CHECK_BY, dataRows, 1)
      .setDataValidation(fcRule);
    updates.push("Final Check By (" + finalCheckers.length + ")");
  }

  if (updates.length === 0) {
    return "ℹ️ Nothing to update — pass non-empty arrays for employees and/or finalCheckers.";
  }

  return "✅ Dropdowns updated: " + updates.join(", ");
}


// =======================================================================================
// LIVE EBAY LOCATION CHECK (2026-08-11)
// =======================================================================================

/**
 * Sidebar: fetch every row's LIVE location from eBay and render a verdict.
 *
 * One button, whole sheet — deliberately NOT per-row: the n8n proxy prices
 * per unique ITEM, and the checker's rhythm is "end of day, check everything."
 * A row-by-row gesture would save nothing and cost a tap per row.
 *
 * Flow:
 *   1. Read SKU + typed LOCATION for every data row (one read).
 *   2. Snapshot MI once (bounded read) → SKU → { itemId, title, url }.
 *   3. POST unique {sku, itemId} pairs to the n8n eBay Location Check Proxy.
 *      SKUs with no MI row / no itemId are NOT sent — they get the
 *      NOT ON EBAY sentinel locally (usually a typo'd SKU; free to catch).
 *   4. Write ◉ EBAY LOC values + verdict colors in batched passes.
 *   5. Re-apply SKU→listing links from the same MI snapshot (no extra read).
 *
 * The fetched value is a SNAPSHOT — the D1 header note records when it was
 * taken. Editing a row's LOCATION afterwards re-computes that row's verdict
 * locally (see locationUpdateOnEdit); changing the SKU clears the stale value.
 *
 * @returns {string} summary for the sidebar status bar
 */
function fetchEbayLocationsLive() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LOCATION_UPDATE.sheetName);
  if (!sheet) return "❌ Location Update sheet doesn't exist — run Re-style first.";
  _luEnsureV4(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < LOCATION_UPDATE.dataStartRow) {
    return "ℹ️ No data rows to check.";
  }

  var nRows = lastRow - LOCATION_UPDATE.dataStartRow + 1;
  var grid = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, nRows, 2).getValues();

  // --- MI snapshot: SKU → { itemId, title, url } ---
  var mi = _luMiSnapshot();
  if (!mi.size) return "❌ Master Inventory unavailable — cannot resolve eBay item ids.";

  // --- Build the unique fetch list ---
  var items = [];
  var seen = {};
  var overCap = 0;
  for (var i = 0; i < grid.length; i++) {
    var sku = String(grid[i][0] || "").trim();
    if (!sku) continue;
    var key = sku.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    var rec = mi.get(key);
    if (!rec || !rec.itemId) continue;             // → NOT ON EBAY locally, no API spend
    if (items.length >= LOCATION_UPDATE.fetchCap) { overCap++; continue; }
    items.push({ sku: sku, itemId: rec.itemId });
  }

  // --- Live call (only if there is anything to ask eBay about) ---
  var bySku = {};
  if (items.length > 0) {
    var resp = triggerEbayLocationCheck(items);
    if (!resp.ok) {
      return "❌ Live check failed — nothing was written. " + resp.message;
    }
    var results = (resp.data && resp.data.results) || [];
    for (var r = 0; r < results.length; r++) {
      bySku[String(results[r].sku || "").toLowerCase()] = results[r];
    }
  }

  // --- Compose column D: values + verdict styles, all batched ---
  var S = LOCATION_UPDATE.sentinels;
  var values = [], bgs = [], colors = [], styles = [];
  var counts = { match: 0, mismatch: 0, notOnEbay: 0, failed: 0, blank: 0 };

  for (var j = 0; j < grid.length; j++) {
    var rowSku = String(grid[j][0] || "").trim();
    var typedLoc = String(grid[j][1] || "").trim();
    var display = "";

    if (rowSku) {
      var mrec = mi.get(rowSku.toLowerCase());
      if (!mrec || !mrec.itemId) {
        display = S.notOnEbay;
        counts.notOnEbay++;
      } else {
        var res = bySku[rowSku.toLowerCase()];
        if (!res) {
          // In the fetch list but no result back (or dropped by the cap) —
          // never guess; say the check didn't happen for this row.
          display = S.checkFailed;
          counts.failed++;
        } else if (!res.ok) {
          display = S.checkFailed;
          counts.failed++;
        } else {
          var fetched = String(res.location || "").trim();
          if (!fetched) {
            display = S.blankOnEbay;
            counts.blank++;
          } else {
            display = fetched;
          }
        }
      }
    }

    var verdict = _luVerdict(typedLoc, display);
    // blank-on-eBay rows RENDER as mismatch (red) but are reported under
    // their own "blank on eBay" tally — counting them in both lines would
    // make the summary numbers disagree with the sheet.
    if (verdict === "match") counts.match++;
    else if (verdict === "mismatch" && display !== S.blankOnEbay) counts.mismatch++;

    var st = _luVerdictStyle(verdict);
    values.push([display]);
    bgs.push([st.bg]);
    colors.push([st.color]);
    styles.push([st.italic ? "italic" : "normal"]);
  }

  var dRange = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.EBAY_LOC, nRows, 1);
  dRange.setValues(values);
  dRange.setBackgrounds(bgs);
  dRange.setFontColors(colors);
  dRange.setFontStyles(styles);

  // Stamp WHEN this snapshot was taken on the column header (hover to read).
  var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy h:mm a");
  sheet.getRange(LOCATION_UPDATE.headerRow, LOCATION_UPDATE.cols.EBAY_LOC)
    .setNote("Live eBay check: " + nowStr + " · " + items.length + " item(s) queried");

  // --- SKU → listing links, from the SAME MI snapshot (zero extra reads) ---
  try {
    applySkuLinksToColumn(sheet, LOCATION_UPDATE.cols.SKU,
                          LOCATION_UPDATE.dataStartRow, lastRow, mi);
  } catch (linkErr) {
    console.log("fetchEbayLocationsLive: link pass failed (non-fatal): " + linkErr);
  }

  SpreadsheetApp.flush();

  var parts = [];
  parts.push(items.length + " checked");
  parts.push(counts.match + " match");
  if (counts.mismatch) parts.push(counts.mismatch + " MISMATCH");
  if (counts.blank) parts.push(counts.blank + " blank on eBay");
  if (counts.notOnEbay) parts.push(counts.notOnEbay + " not on eBay");
  if (counts.failed) parts.push(counts.failed + " failed");
  if (overCap) parts.push(overCap + " skipped (over " + LOCATION_UPDATE.fetchCap + " cap)");

  return "✅ " + parts.join(" · ");
}


/**
 * Normalize a location string for comparison.
 * "g-35 * 2" → "G-35" (the trailing "* N" is the warehouse PACK-SIZE note —
 * "one sellable unit is pre-packed as N pieces" — never part of the address).
 * All whitespace is dropped so "A - 57" and "A-57" compare equal; dual
 * locations like "L-208 / C-51" become "L-208/C-51" on both sides.
 * PURE — Node-tested.
 */
function _luNormalizeLoc(s) {
  return String(s == null ? "" : s)
    .trim()
    .replace(/\s*\*\s*\d+\s*$/, "")   // strip pack-size suffix
    .replace(/\s+/g, "")              // drop ALL whitespace
    .toUpperCase();
}


/**
 * Classify one row's typed LOCATION vs the value about to sit in ◉ EBAY LOC.
 * PURE — Node-tested.
 *
 * @param {string} typedLoc  what the worker typed in col C (may be "")
 * @param {string} display   the value for col D: a real location, one of
 *                           LOCATION_UPDATE.sentinels, or "" (empty row)
 * @returns {"match"|"mismatch"|"info"|"none"}
 *   match    — eBay agrees with the typed location
 *   mismatch — eBay disagrees, OR the listing's location is blank while the
 *              sheet says it should have one (same operational meaning: the
 *              eBay side was not updated correctly)
 *   info     — nothing to verify: no typed location yet, SKU not on eBay,
 *              or the live check failed → dim, never loud
 *   none     — empty cell (row has no SKU)
 */
function _luVerdict(typedLoc, display) {
  var S = LOCATION_UPDATE.sentinels;
  var d = String(display == null ? "" : display).trim();
  var t = String(typedLoc == null ? "" : typedLoc).trim();

  if (!d) return "none";
  if (d === S.notOnEbay || d === S.checkFailed) return "info";
  if (d === S.blankOnEbay) return t ? "mismatch" : "info";
  if (!t) return "info";
  return _luNormalizeLoc(t) === _luNormalizeLoc(d) ? "match" : "mismatch";
}


/**
 * Visual treatment per verdict — colors only ever touch column D, so the
 * duplicate-SKU amber on col B and the banding elsewhere are never disturbed.
 * bg null = banding shows through.
 */
function _luVerdictStyle(verdict) {
  switch (verdict) {
    case "match":    return { bg: "#e8f5e9", color: "#1b5e20", italic: false };
    case "mismatch": return { bg: "#ffcdd2", color: "#b71c1c", italic: false };
    case "info":     return { bg: null,      color: "#9e9e9e", italic: true  };
    default:         return { bg: null,      color: "#434343", italic: false };
  }
}


/**
 * Reset one ◉ EBAY LOC cell to its base (unfetched) look. Used when the row's
 * SKU changes/clears — the old fetched value belonged to the old SKU and a
 * leftover verdict color with no meaning reads as a rendering fault.
 */
function _luResetEbayLocCell(range) {
  range.setValue("")
       .setBackground(null)
       .setFontColor("#434343")
       .setFontStyle("normal");
}


/**
 * Bounded Master Inventory snapshot for the live check:
 *   Map skuLower → { itemId, title, url }
 *
 * MI is ~174 cols × ~3,500 rows — a getDataRange() here would be the same
 * 600,000-cell read that slowed the Part Console (see the 2026-08-06 perf
 * hunt). Instead: read the header row, locate the four columns we need, and
 * read ONE contiguous block spanning them (span-guarded with a per-column
 * fallback if the columns ever drift far apart) — the _pcMasterSnapshot
 * pattern.
 *
 * `url` doubles as the SKU-link target, so applySkuLinksToColumn can reuse
 * this map directly (it only reads `.url`).
 */
function _luMiSnapshot() {
  var map = new Map();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var db = ss.getSheetByName(DB_SHEET_NAME);
  if (!db) return map;

  var lastCol = db.getLastColumn();
  var lastRowMi = db.getLastRow();
  if (lastRowMi < 2) return map;

  var headers = db.getRange(1, 1, 1, lastCol).getValues()[0];
  var cSku    = headers.indexOf(DB_SKU_HEADER);
  var cItemId = headers.indexOf("itemId");        // same literal updateMiRows resolves
  var cTitle  = headers.indexOf(DB_TITLE_HEADER);
  var cUrl    = headers.indexOf(DB_VIEWURL_HEADER);
  if (cSku === -1 || cItemId === -1) return map;

  var wanted = [cSku, cItemId, cTitle, cUrl].filter(function(c) { return c >= 0; });
  var minC = Math.min.apply(null, wanted);
  var maxC = Math.max.apply(null, wanted);
  var nMiRows = lastRowMi - 1;

  var cols = {};
  if (maxC - minC <= 40) {
    // One contiguous block read covering all wanted columns.
    var block = db.getRange(2, minC + 1, nMiRows, maxC - minC + 1).getValues();
    cols.sku    = function(r) { return block[r][cSku - minC]; };
    cols.itemId = function(r) { return block[r][cItemId - minC]; };
    cols.title  = cTitle >= 0 ? function(r) { return block[r][cTitle - minC]; } : function() { return ""; };
    cols.url    = cUrl   >= 0 ? function(r) { return block[r][cUrl   - minC]; } : function() { return ""; };
  } else {
    // Columns drifted apart — fall back to one read per column.
    var vSku    = db.getRange(2, cSku + 1,    nMiRows, 1).getValues();
    var vItemId = db.getRange(2, cItemId + 1, nMiRows, 1).getValues();
    var vTitle  = cTitle >= 0 ? db.getRange(2, cTitle + 1, nMiRows, 1).getValues() : null;
    var vUrl    = cUrl   >= 0 ? db.getRange(2, cUrl   + 1, nMiRows, 1).getValues() : null;
    cols.sku    = function(r) { return vSku[r][0]; };
    cols.itemId = function(r) { return vItemId[r][0]; };
    cols.title  = vTitle ? function(r) { return vTitle[r][0]; } : function() { return ""; };
    cols.url    = vUrl   ? function(r) { return vUrl[r][0]; }   : function() { return ""; };
  }

  for (var r = 0; r < nMiRows; r++) {
    var sku = String(cols.sku(r) || "").trim().toLowerCase();
    if (!sku) continue;
    map.set(sku, {
      itemId: String(cols.itemId(r) || "").trim(),
      title:  String(cols.title(r)  || "").trim(),
      url:    String(cols.url(r)    || "").trim()
    });
  }
  return map;
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEditInstallable(e)
// =======================================================================================

/**
 * SKU edit on Location Update → auto-stamp TIMESTAMP only (LOCATION is manual
 * per user request 2026-05-14). Since v4 it ALSO:
 *   - clears the row's ◉ EBAY LOC (value + verdict color) — a fetched value
 *     belongs to the SKU it was fetched for; a new/changed SKU makes it stale
 *   - links the SKU cell to its eBay listing (same treatment as All Orders /
 *     Prep Queue — hover shows Google's title+photo preview card)
 *
 * A LOCATION (col C) edit re-computes that row's verdict LOCALLY against the
 * already-fetched ◉ EBAY LOC value — no API call — so a checker who fixes the
 * typed location sees the red clear immediately instead of wondering whether
 * they need to re-fetch.
 *
 * Runs in the INSTALLABLE trigger because the duplicate-highlight refresh
 * uses openById via SpreadsheetApp.flush() context — simple triggers can
 * fail silently.
 *
 * Defensive: any error is logged and swallowed so this never blocks other
 * onEditInstallable handlers (Prep Queue, Out of Stock, manual receive, etc.).
 */
function locationUpdateOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== LOCATION_UPDATE.sheetName) return;
    if (e.range.getRow() < LOCATION_UPDATE.dataStartRow) return;

    // One-time v3→v4 auto-migration before any column-positioned write.
    // The edited range (col B) sits LEFT of the inserted column, so e.range
    // stays valid across the insert.
    _luEnsureV4(sheet);

    var editedCol = e.range.getColumn();

    // --- LOCATION edited → re-verdict against the existing fetched value ---
    if (editedCol === LOCATION_UPDATE.cols.LOCATION) {
      _luReVerdictRows(sheet, e.range.getRow(), e.range.getNumRows());
      return;
    }

    if (editedCol !== LOCATION_UPDATE.cols.SKU) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();

    var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy h:mm a");
    var linkRows = [];

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      var skuRaw = String(edits[i][0]).trim();
      var skuLower = skuRaw.toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe LOCATION + EBAY_LOC + TIMESTAMP so the row reads
        // as empty. SKU and LOCATION are a paired audit entry; clearing the
        // SKU clears the location record with it.
        //
        // COUNTER is a formula (=ROW()-1) that's intentionally left alone — the
        // user wanted # to PERSIST across SKU clears (fixed 2026-05-13).
        // EMPLOYEE + FINAL_CHECK_BY are user-input only; never touched here.
        sheet.getRange(row, LOCATION_UPDATE.cols.LOCATION).setValue("");
        _luResetEbayLocCell(sheet.getRange(row, LOCATION_UPDATE.cols.EBAY_LOC));
        sheet.getRange(row, LOCATION_UPDATE.cols.TIMESTAMP).setValue("");
        continue;
      }

      // Non-empty SKU — stamp TIMESTAMP, drop any stale fetched location
      // (it was fetched for whatever SKU used to be here), queue the link.
      sheet.getRange(row, LOCATION_UPDATE.cols.TIMESTAMP).setValue(nowStr);
      _luResetEbayLocCell(sheet.getRange(row, LOCATION_UPDATE.cols.EBAY_LOC));
      linkRows.push({ row: row, sku: skuRaw });
    }

    // --- SKU → listing links for the edited rows ---
    // Single-row edits (the typing case) use the cheap single lookup; a
    // multi-row paste builds the MI map once instead of N full scans.
    try {
      if (linkRows.length === 1) {
        var rec = getSingleSkuEnrichment(linkRows[0].sku.toLowerCase());
        sheet.getRange(linkRows[0].row, LOCATION_UPDATE.cols.SKU)
          .setRichTextValue(_skuRichText(linkRows[0].sku, rec));
      } else if (linkRows.length > 1) {
        var map = buildSkuEnrichmentMap();
        if (map.size) {
          applySkuLinksToColumn(sheet, LOCATION_UPDATE.cols.SKU,
                                startRow, startRow + edits.length - 1, map);
        }
      }
    } catch (linkErr) {
      try { Logger.log("locationUpdateOnEdit link error: " + linkErr); } catch (_) {}
    }

    // Refresh duplicate-SKU highlight after every edit batch — surfaces dupes
    // the moment they're typed, clears highlights when one of a pair is changed.
    _refreshLocationUpdateDuplicates(sheet);
  } catch (err) {
    try { Logger.log("locationUpdateOnEdit error: " + err); } catch (_) {}
  }
}


/**
 * Re-compute the verdict color for rows whose typed LOCATION just changed,
 * using the fetched value already sitting in ◉ EBAY LOC. Pure repaint — the
 * fetched VALUE is never altered here, and rows that were never fetched
 * (empty D) are left untouched.
 */
function _luReVerdictRows(sheet, startRow, numRows) {
  var band = sheet.getRange(startRow, LOCATION_UPDATE.cols.LOCATION, numRows, 2);
  var vals = band.getValues();   // [ [LOCATION, EBAY_LOC], ... ]

  for (var i = 0; i < numRows; i++) {
    var display = String(vals[i][1] || "").trim();
    if (!display) continue;
    var verdict = _luVerdict(String(vals[i][0] || "").trim(), display);
    var st = _luVerdictStyle(verdict);
    sheet.getRange(startRow + i, LOCATION_UPDATE.cols.EBAY_LOC)
      .setBackground(st.bg)
      .setFontColor(st.color)
      .setFontStyle(st.italic ? "italic" : "normal");
  }
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (mirrors the PrepQueue + OutOfStock pattern)
// =======================================================================================

/**
 * Scans column B (SKU), identifies SKUs that appear two or more times
 * (case-insensitive, trimmed), and paints those cells with soft amber
 * background + thick yellow border. Removing a duplicate clears the highlight
 * on its formerly-duped counterpart in the same pass.
 *
 * Only touches background + border — never font, alignment, number format,
 * banding, etc. When a duplicate is removed, the highlight clears cleanly
 * and the row's original look is preserved.
 *
 * Scans the full data band (Math.min(maxRows, 1000)) — not just to lastRow —
 * so previously-highlighted cells whose SKU has been deleted still get their
 * background explicitly cleared. See _refreshPrepQueueDuplicates docstring
 * for the full rationale.
 *
 * Run after any change that could affect col B:
 *   - setupLocationUpdateSheet (initial paint)
 *   - locationUpdateOnEdit (live, on every SKU edit)
 *   - refreshLocationUpdateSheet (post-sweep)
 */
function _refreshLocationUpdateDuplicates(sheet) {
  if (!sheet) return;

  var maxScanRow = Math.min(sheet.getMaxRows(), 1000);
  if (maxScanRow < LOCATION_UPDATE.dataStartRow) return;

  var totalRows = maxScanRow - LOCATION_UPDATE.dataStartRow + 1;
  var skuRange  = sheet.getRange(LOCATION_UPDATE.dataStartRow, LOCATION_UPDATE.cols.SKU, totalRows, 1);
  var skus      = skuRange.getValues();

  // Count occurrences (case-insensitive, trimmed)
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    var k = String(skus[i][0]).trim().toLowerCase();
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }

  // Build the FULL backgrounds array — dupes get amber, everything else null.
  var bgs = [];
  var dupeIndexes = [];
  for (var j = 0; j < skus.length; j++) {
    var key = String(skus[j][0]).trim().toLowerCase();
    if (key && counts[key] >= 2) {
      bgs.push(['#fff3b0']);
      dupeIndexes.push(j);
    } else {
      bgs.push([null]);
    }
  }
  skuRange.setBackgrounds(bgs);

  // Borders: clear the full range (single batched call), then add thick yellow
  // borders per dupe (small N, typically 0-4 calls).
  skuRange.setBorder(false, false, false, false, false, false, null, null);
  for (var k = 0; k < dupeIndexes.length; k++) {
    var rowIdx = dupeIndexes[k];
    sheet.getRange(LOCATION_UPDATE.dataStartRow + rowIdx, LOCATION_UPDATE.cols.SKU)
      .setBorder(true, true, true, true, false, false,
                 '#ffb800', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // Force pending writes to land before the user's next interaction.
  SpreadsheetApp.flush();
}
