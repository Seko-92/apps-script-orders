// =======================================================================================
// ROW_MANAGEMENT.gs - v2.5 SIMPLE (Copies from eBay which works!)
// =======================================================================================

/**
 * Deletes empty rows while preserving buffer rows
 * @param {number} t - Table number (1 or 2)
 * @returns {string} - Status message
 */
function deleteEmptyRows(t) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var b = getBoundaryRow();
  var start = (t === 1) ? Schema.dataStartRow : b + 2;
  var end   = (t === 1) ? b - 1               : sheet.getMaxRows();
  var last  = findLastDataRowInSegment(start, end);

  var delStart = (t === 1) ? last + 4 : last + MAX_EMPTY_ROWS_TO_KEEP + 1;

  if (t === 1 && delStart >= b) return "ℹ️ 3-row buffer already exists.";

  if (delStart < end) {
    sheet.deleteRows(delStart, end - delStart + 1);
    return "✅ Cleanup complete (3-row buffer preserved).";
  }
  return "ℹ️ Already clean.";
}

function runDeleteEmptyRowsTableOne() {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('runDeleteEmptyRowsTableOne', []);
 return deleteEmptyRows(1); }
function runDeleteEmptyRowsTableTwo() {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('runDeleteEmptyRowsTableTwo', []);
 return deleteEmptyRows(2); }

/**
 * Ensures the DIRECT table always has at least 3 empty buffer rows
 * with proper data formatting (not header formatting).
 * Called automatically via onChange when rows are deleted.
 */
function ensureDirectTableBuffer() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var boundary = getBoundaryRow();
  if (boundary === -1) return;

  var BUFFER_SIZE = 3;
  var directDataStart = boundary + 2; // First data row after DIRECT header
  var lastRow = sheet.getLastRow();

  // Find last data row in DIRECT table
  var lastDataRow = findLastDataRowInSegment(directDataStart, lastRow);

  // Count empty rows after last data (or after header if no data)
  var emptyStart = (lastDataRow >= directDataStart) ? lastDataRow + 1 : directDataStart;
  var emptyCount = lastRow - emptyStart + 1;
  if (emptyStart > lastRow) emptyCount = 0;

  if (emptyCount >= BUFFER_SIZE) return; // Buffer already exists

  var rowsToAdd = BUFFER_SIZE - emptyCount;

  // Add rows at the end of the sheet
  sheet.insertRowsAfter(lastRow, rowsToAdd);

  // Copy formatting from eBay data row (which always has correct format)
  var sourceRange = sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth);
  var targetRange = sheet.getRange(lastRow + 1, 1, rowsToAdd, Schema.dataWidth);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.setRowHeights(lastRow + 1, rowsToAdd, 30);
}

/** How many blank rows each table carries below its last row of data. */
var TABLE_BUFFER_ROWS = 3;

/**
 * balanceTableBuffers(n) — give BOTH tables exactly the same number of trailing
 * blank rows. Owner-run, idempotent, reports what it did.
 *
 * ⚠⚠ WHY THE TWO DRIFTED APART. `ensureDirectTableBuffer` only ever ADDS — it tops the
 *    DIRECT table back up to 3 and has no path that removes anything. So every row
 *    deleted from the middle, every n8n shipped-row sweep, every manual tidy leaves the
 *    tail one row longer than it found it, and the gap grows in one direction forever.
 *    Observed 2026-08-31: eBay 3, DIRECT 6. **Exactly the Prep Queue's buffer bug of
 *    2026-07-20** — and the resolution is the same one: the automatic path keeps only
 *    GROWING (it must never delete a row someone is mid-way through typing into), and the
 *    trim lives behind a deliberate call.
 *
 * ⚠ NEVER DELETES A ROW THAT HOLDS ANYTHING. Both ends are measured from
 *   findLastDataRowInSegment, which reads column A, so only the genuinely blank tail
 *   past the buffer is ever removed. A table with no data at all is handled: the scanner
 *   returns start-1, so "last data + buffer" still resolves to the right target.
 *
 * ⚠ eBay's tail sits ABOVE the divider and DIRECT's sits at the BOTTOM OF THE SHEET, so
 *   they are two different edits — the first moves the boundary, the second does not.
 *   eBay is done FIRST and the boundary is re-read afterwards, because inserting or
 *   deleting above it invalidates every row number below.
 *
 * ⚠ Deliberately NOT bridged through OwnerBridge. The allowlist is EXECUTED BY doPost on
 *   the pinned /exec, so adding a name to it costs a New Version — and this is owner-side
 *   housekeeping, not something the floor ever calls.
 */
function balanceTableBuffersNow() { return balanceTableBuffers(); }

function balanceTableBuffers(n) {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Balancing the table buffers");
    if (denied) return denied;
  }
  var want = Math.max(1, parseInt(n, 10) || TABLE_BUFFER_ROWS);
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { var m = "❌ Main sheet not found."; console.log(m); return m; }

  var out = ["── BALANCE TABLE BUFFERS · target " + want + " blank row(s) each ──"];

  // ---- eBay: the tail between its last data row and the DIRECT divider ----------------
  var b = getBoundaryRow();
  if (b === -1) {
    var msg = "❌ DIRECT divider not found — refusing to touch row structure.\n" +
              "   getBoundaryRow() is strict equality on \"DIRECT\" in column A.";
    console.log(msg); return msg;
  }
  var eLast = findLastDataRowInSegment(Schema.dataStartRow, b - 1);
  var eHave = (b - 1) - eLast;
  out.push("eBay:   last data row " + (eLast < Schema.dataStartRow ? "(none)" : eLast) +
           " · " + eHave + " blank → " + want);

  if (eHave > want) {
    var cut = eHave - want;
    sheet.deleteRows(eLast + want + 1, cut);
    out.push("        ✓ removed " + cut + " blank row(s)");
  } else if (eHave < want) {
    var add = want - eHave;
    // ⚠ insertRowsBefore on this sheet can corrupt the header text — a documented
    //   Sheets bug (gotcha #4). Restore immediately, before anything reads the columns.
    sheet.insertRowsBefore(b, add);
    try { verifyAndRestoreHeaders(); } catch (e) { out.push("        ⚠ header restore: " + e); }
    // Format from a real data row, never from whatever sat above the insert point.
    sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth)
         .copyTo(sheet.getRange(b, 1, add, Schema.dataWidth),
                 SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.setRowHeights(b, add, 30);
    out.push("        ✓ added " + add + " blank row(s)");
  } else {
    out.push("        – already right");
  }

  // ---- DIRECT: the tail between its last data row and the end of the sheet ------------
  // ⚠ RE-READ THE BOUNDARY. The edit above moved it.
  SpreadsheetApp.flush();
  b = getBoundaryRow();
  var dStart = b + 2;
  var maxRow = sheet.getMaxRows();
  var dLast  = findLastDataRowInSegment(dStart, maxRow);
  var dHave  = maxRow - dLast;
  out.push("DIRECT: last data row " + (dLast < dStart ? "(none)" : dLast) +
           " · " + dHave + " blank → " + want);

  if (dHave > want) {
    var dcut = dHave - want;
    sheet.deleteRows(dLast + want + 1, dcut);
    out.push("        ✓ removed " + dcut + " blank row(s)");
  } else if (dHave < want) {
    var dadd = want - dHave;
    sheet.insertRowsAfter(maxRow, dadd);
    sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth)
         .copyTo(sheet.getRange(maxRow + 1, 1, dadd, Schema.dataWidth),
                 SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.setRowHeights(maxRow + 1, dadd, 30);
    out.push("        ✓ added " + dadd + " blank row(s)");
  } else {
    out.push("        – already right");
  }

  // The divider's per-order boxes are drawn by ROW POSITION, so anything that shifts
  // rows has to leave them repainted rather than stranded one row off.
  try { setupDuplicateSalesOrderHighlighting(); } catch (e) {
    out.push("⚠ duplicate/divider repaint failed: " + e);
  }

  var rep = out.join("\n");
  console.log(rep);
  return rep;
}

/**
 * Adds rows to Table 1 (eBay) - PUSHES DIRECT TABLE DOWN
 * @param {number} n - Number of rows to add
 * @returns {string} - Status message
 */
function addRowsTableOne(n) {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('addRowsTableOne', [n]);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var lastUsedRow = findLastDataRowInSegment(Schema.dataStartRow, boundary - 1);
  var insertionPoint = (lastUsedRow < Schema.dataStartRow) ? Schema.dataStartRow : lastUsedRow + 1;
  var rowsToInsert = parseInt(n);

  sheet.insertRowsAfter(insertionPoint, rowsToInsert);

  return "✅ Inserted " + rowsToInsert + " rows. DIRECT moved to Row " + (boundary + rowsToInsert) + ".";
}

/**
 * SIMPLEST FIX: Copy format from eBay table (which works perfectly!)
 * Since both tables should have the same format anyway
 * @param {number} n - Number of rows to add
 * @returns {string} - Status message
 */
function addRowsTableTwo(n) {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('addRowsTableTwo', [n]);

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var numRows = parseInt(n);
  var lastRow = sheet.getLastRow();

  // Insert the rows at the end
  sheet.insertRowsAfter(lastRow, numRows);

  // Copy format from eBay table's first data row (which always has correct format)
  var sourceRange = sheet.getRange(Schema.dataStartRow, 1, 1, Schema.dataWidth);
  var targetRange = sheet.getRange(lastRow + 1, 1, numRows, Schema.dataWidth);

  // Copy ONLY the format (not content)
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Set row heights to match
  sheet.setRowHeights(lastRow + 1, numRows, 30);

  return "✅ Added " + numRows + " rows (format copied from eBay table).";
}

// =======================================================================================
// BOUNDARY PROTECTION FUNCTIONS
// =======================================================================================

function protectBoundaryRow() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  removeExistingBoundaryProtection(sheet);

  var boundaryRange = sheet.getRange(boundary, 1, 1, Schema.dataWidth);
  var protection = boundaryRange.protect();
  protection.setDescription('DIRECT_BOUNDARY_PROTECTED');
  protection.setWarningOnly(true);

  var headerRange = sheet.getRange(boundary + 1, 1, 1, Schema.dataWidth);
  var headerProtection = headerRange.protect();
  headerProtection.setDescription('DIRECT_HEADER_PROTECTED');
  headerProtection.setWarningOnly(true);
  
  return "✅ Protected DIRECT boundary (Row " + boundary + ") and header (Row " + (boundary + 1) + ").";
}

function removeExistingBoundaryProtection(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var desc = protections[i].getDescription();
    if (desc === 'DIRECT_BOUNDARY_PROTECTED' || desc === 'DIRECT_HEADER_PROTECTED') {
      protections[i].remove();
    }
  }
}

function unprotectBoundaryRow() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  removeExistingBoundaryProtection(sheet);
  return "✅ Boundary protection removed.";
}

function validateBoundaryIntegrity() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  if (boundary === -1) {
    Logger.log("⚠️ CRITICAL: DIRECT boundary row not found!");
    return false;
  }
  
  var cellValue = sheet.getRange(boundary, Schema.cols.SKU).getValue();
  if (String(cellValue).toUpperCase().indexOf(Schema.boundaryMarker) === -1) {
    Logger.log("⚠️ WARNING: Boundary row " + boundary + " doesn't contain '" + Schema.boundaryMarker + "'. Value: " + cellValue);
    return false;
  }
  
  Logger.log("✅ Boundary integrity OK. DIRECT is at row " + boundary);
  return true;
}

// =======================================================================================
// HIGHLIGHT DUPLICATES - Shared Infrastructure
// =======================================================================================

/**
 * Bright color palette for duplicate SKU groups.
 * Each entry: [background, fontColor] — visually matched pairs.
 * 20 pairs, cycles if more groups exist.
 */
var SKU_DUPE_COLORS = [
  ["#ff6d6d", "#7a0000"],  // Bright Red / Dark Red
  ["#4fc3f7", "#01579b"],  // Sky Blue / Navy
  ["#81c784", "#1b5e20"],  // Bright Green / Forest
  ["#ffb74d", "#e65100"],  // Bright Orange / Burnt Orange
  ["#ba68c8", "#4a148c"],  // Bright Purple / Deep Purple
  ["#4dd0e1", "#006064"],  // Bright Cyan / Dark Cyan
  ["#e57373", "#b71c1c"],  // Vivid Coral / Crimson
  ["#fff176", "#f57f17"],  // Bright Yellow / Amber
  ["#aed581", "#33691e"],  // Lime Green / Olive
  ["#ff8a65", "#bf360c"],  // Tangerine / Mahogany
  ["#7986cb", "#1a237e"],  // Bright Indigo / Deep Indigo
  ["#4db6ac", "#004d40"],  // Bright Teal / Dark Teal
  ["#f06292", "#880e4f"],  // Hot Pink / Wine
  ["#dce775", "#827717"],  // Chartreuse / Olive Gold
  ["#64b5f6", "#0d47a1"],  // Dodger Blue / Royal Blue
  ["#ffab91", "#bf360c"],  // Salmon / Rust
  ["#a1887f", "#3e2723"],  // Mocha / Espresso
  ["#90caf9", "#0d47a1"],  // Cornflower / Dark Blue
  ["#ce93d8", "#6a1b9a"],  // Orchid / Plum
  ["#80cbc4", "#00695c"],  // Aquamarine / Emerald
];

// (ORDER_BORDER_COLORS removed 2026-07-14 — the colored left-border tabs
// were replaced by the SO badge glyphs, which survive B&W printing and
// scattered rows. Git history has the palette if ever wanted back.)

// =======================================================================================
// HIGHLIGHT DUPLICATE SKUs (Per-Group, Bright, Auto-Refresh)
// =======================================================================================

/**
 * Sets up per-group duplicate SKU highlighting with matched font colors.
 * Each duplicate SKU group gets its own bright color + dark complementary font.
 * Skips DIRECT boundary row.
 * Called from onOpen() and auto-refreshed on edits.
 */
function setupDuplicateHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return null;

  removeDuplicateHighlightRules(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return null;

  var allData = sheet.getRange(Schema.dataStartRow, Schema.cols.SKU, lastRow - Schema.dataStartRow + 1, 1).getValues();
  var boundary = getBoundaryRow();

  var skuCount = {};
  for (var i = 0; i < allData.length; i++) {
    var currentRow = Schema.dataStartRow + i;
    if (boundary > 0 && (currentRow === boundary || currentRow === boundary + 1)) continue;
    var sku = String(allData[i][0]).trim().toUpperCase();
    if (sku && sku !== Schema.boundaryMarker) {
      skuCount[sku] = (skuCount[sku] || 0) + 1;
    }
  }

  var duplicateSkus = [];
  for (var sku in skuCount) {
    if (skuCount[sku] > 1) duplicateSkus.push(sku);
  }

  if (duplicateSkus.length === 0) return null;

  var rules = sheet.getConditionalFormatRules();
  var skuRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SKU, 1000, 1);
  var ref = "A" + Schema.dataStartRow;

  for (var i = 0; i < duplicateSkus.length; i++) {
    var escapedSku = duplicateSkus[i].replace(/"/g, '""');
    var pair = SKU_DUPE_COLORS[i % SKU_DUPE_COLORS.length];

    var formula = '=UPPER(TRIM(' + ref + '))="' + escapedSku + '"';
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(pair[0])
      .setFontColor(pair[1])
      .setRanges([skuRange])
      .build();

    rules.push(rule);
  }

  sheet.setConditionalFormatRules(rules);
}

function highlightAllDuplicates() {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('highlightAllDuplicates', []);

  setupDuplicateHighlighting();
  return "✅ Duplicate SKU highlighting enabled.";
}

function removeDuplicateHighlightRules(sheet) {
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        if (formula === '=1=1' || formula === '=2=2') continue;
        var ranges = rules[i].getRanges();
        // ⚠⚠ EVERY range must be column A — see the identical note in
        //   removeLegacySalesOrderCFRules. The identity rules (2026-08-30) span cols A AND
        //   D and their status guard contains UPPER(TRIM(, so an "ANY range" test deleted
        //   them here too. A duplicate-SKU rule targets column A and nothing else.
        var isSkuColumn = ranges.length > 0;
        for (var j = 0; j < ranges.length; j++) {
          if (!(ranges[j].getColumn() === Schema.cols.SKU &&
                ranges[j].getNumColumns() === 1)) {
            isSkuColumn = false;
            break;
          }
        }
        if (isSkuColumn && (formula.indexOf('COUNTIF') !== -1 || formula.indexOf('UPPER(TRIM(') !== -1)) {
          continue;
        }
      }
    }
    filtered.push(rules[i]);
  }
  sheet.setConditionalFormatRules(filtered);
}

function clearAllDuplicateHighlights() {
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('clearAllDuplicateHighlights', []);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  removeDuplicateHighlightRules(sheet);
  PropertiesService.getScriptProperties().deleteProperty('DUPE_ORIGINAL_BGS');
  return "✅ Duplicate SKU highlights cleared.";
}

// =======================================================================================
// DUPLICATE SALES ORDER BORDERS (Per-Group, Colored Left Border Tabs)
// =======================================================================================

// Keycap digits 1-10 — the SO BADGE glyph set (2026-07-14, final form).
// Rendered via NUMBER FORMAT prefix ONLY (display layer): the cell VALUE
// stays the clean order id, so every machine reader is untouched — n8n's
// four All-Orders reader nodes (pinned to UNFORMATTED_VALUE), the
// updateOrderStatus col-D matcher, dedupe, and the rich-text order links.
// Same device as the ▣ kit marker and the ▌ DIRECT divider.
//
// WHY KEYCAP EMOJI (the glyph saga, condensed): the filled circled digits
// (❶❷❸) are illegible at the table's 10px, and making ONLY the glyph render
// larger is IMPOSSIBLE — Sheets normalizes cell-level font writes and
// rich-text run styles into each other on every write (setFontSizes stomps
// run sizes; setRichTextValue resets the cell default from the runs), so a
// format prefix can never durably out-size the value text beside it. Two
// shipped attempts proved this. Keycap emoji solve it a different way:
// they render as colored squares visually LARGER than letterforms at the
// SAME font size — instantly distinguishable, table stays uniform 10px.
// The PRINT pick list maps the badge to a drawn ink circle-digit instead
// (emoji print as gray mush on B&W; see _badgeFromFormat + .so-badge).
// Numbering restarts per TABLE (eBay / DIRECT) and cycles past 10 — a badge
// only needs to be unique among the currently-visible multi-item orders of
// its own table.
//
// STICKY NUMBERING (2026-07-16): a group KEEPS the digit it already wears
// across every repaint — numbers are no longer positional (top-to-bottom
// renumbering made every n8n arrival shift ALL badges, so paper printed an
// hour earlier disagreed with the sheet). The persistent store is the sheet
// itself (the col-D number formats — cloud-side, identical for every device
// and for the print template); a digit frees only when its group leaves the
// sheet or shrinks to one row, and new groups draw the lowest free digit.
var SO_BADGE_GLYPHS = ["1️⃣","2️⃣","3️⃣","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"];

/**
 * Multi-item Sales Order marking on Column D — the SO BADGE:
 * a keycap digit (1️⃣2️⃣3️⃣…) prefixed via number format on every row of the
 * group. Survives the aisle sort's scattering (repeated on every row of the
 * group); the print pick list reads the SAME format and renders it as a
 * drawn ink circle-digit (B&W-crisp). The table stays uniform 10px — the
 * keycap's salience comes from the glyph itself, not font size. (The
 * colored left-border tabs this replaced were dropped 2026-07-14; the
 * band-wide border clear remains so legacy bars self-wipe.)
 * Badge digits are STICKY per Sales Order (2026-07-16): each repaint first
 * reads the digit a group already wears from the current formats and keeps
 * it, so printed pick lists stay consistent with the sheet when new orders
 * land on top. Only new groups draw a fresh (lowest free) digit.
 * Clears stale borders AND stale badge formats first, then re-applies for
 * current duplicates. Skips DIRECT boundary row and its header.
 * Called from onOpen(), auto-refreshed on edits/inserts, and re-run by
 * sortTableByStatusAndLocation (formats travel with the sort, but the
 * repaint keeps assignment canonical — ▣ lesson).
 */
function setupDuplicateSalesOrderHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return null;

  var boundary = getBoundaryRow();

  // Data read is bounded by lastRow (only cells that COULD have a SO value).
  var dataRows = lastRow - Schema.dataStartRow + 1;
  var allData = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, dataRows, 1).getValues();

  // 1. Remove any legacy CF rules on column D (from previous highlight approach)
  removeLegacySalesOrderCFRules(sheet);

  // 2. Clear all left borders on column D — using a WIDER range than lastRow.
  //
  //    Why: getLastRow() returns the position of the last row with CONTENT in
  //    ANY column. When a row's SO is cleared (or n8n removes a shipped order
  //    and empties the whole row), lastRow shrinks. The previously-bordered
  //    row falls outside the lastRow-bounded clear range, so its stale left-
  //    border survives forever. The user sees: "I removed the duplicate but
  //    the highlight stays on the empty cell."
  //
  //    Same class of bug we fixed in _refreshPrepQueueDuplicates 2026-05-06.
  //    Cheap fix: extend the clear band a generous margin past lastRow
  //    (capped at sheet.getMaxRows() to stay in bounds).
  var clearLastRow = Math.min(sheet.getMaxRows(), lastRow + 200);
  var clearRowCount = clearLastRow - Schema.dataStartRow + 1;
  var fullRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, clearRowCount, 1);
  fullRange.setBorder(null, false, null, null, null, null);

  // 2. Count occurrences, skipping boundary rows
  var orderCount = {};
  var orderRows = {};  // Map order → [row numbers]
  for (var i = 0; i < allData.length; i++) {
    var currentRow = Schema.dataStartRow + i;
    if (boundary > 0 && (currentRow === boundary || currentRow === boundary + 1)) continue;
    var order = String(allData[i][0]).trim();
    if (order) {
      orderCount[order] = (orderCount[order] || 0) + 1;
      if (!orderRows[order]) orderRows[order] = [];
      orderRows[order].push(currentRow);
    }
  }

  // 3. Identify duplicates and assign border colors
  var duplicateOrders = [];
  for (var order in orderCount) {
    if (orderCount[order] > 1) duplicateOrders.push(order);
  }

  // 3b. SO BADGES — rebuild column D's number formats over the whole clear
  // band in ONE batched write. Default '@' (plain text: clears stale badges
  // when a group dissolves AND guards order ids against date coercion);
  // multi-item groups get '"❶ "@' etc. Boundary + DIRECT-header rows keep
  // whatever format they already carry. Runs even when there are NO
  // duplicates — that's what erases the last badge when a group shrinks to
  // one row.
  var bandFormats = fullRange.getNumberFormats();

  // STICKY BADGES — capture the digit each group CURRENTLY wears before the
  // band is reset. The formats themselves are the persistent store (no
  // Script Properties, no helper sheet — cloud-side, identical for every
  // device and the print). First badged row of a group wins; a format that
  // doesn't parse to a known glyph counts as unbadged (that group re-draws).
  var BADGE_FMT_RE = /^"(.+) "@$/;
  var existingBadge = {};   // order → index into SO_BADGE_GLYPHS
  for (var e = 0; e < duplicateOrders.length; e++) {
    var eOrd = duplicateOrders[e];
    var eRows = orderRows[eOrd];
    for (var r = 0; r < eRows.length; r++) {
      var em = BADGE_FMT_RE.exec(bandFormats[eRows[r] - Schema.dataStartRow][0]);
      if (!em) continue;
      var eIdx = SO_BADGE_GLYPHS.indexOf(em[1]);
      if (eIdx >= 0) { existingBadge[eOrd] = eIdx; break; }
    }
  }

  for (var f = 0; f < bandFormats.length; f++) {
    var fRow = Schema.dataStartRow + f;
    if (boundary > 0 && (fRow === boundary || fRow === boundary + 1)) continue;
    bandFormats[f][0] = '@';
  }

  // Per-table sequences (numbering is independent per table, eBay / DIRECT).
  var ebaySeq = [];
  var directSeq = [];
  for (var d = 0; d < duplicateOrders.length; d++) {
    var dupFirstRow = orderRows[duplicateOrders[d]][0];   // rows collected top-down
    if (boundary > 0 && dupFirstRow > boundary) directSeq.push(duplicateOrders[d]);
    else ebaySeq.push(duplicateOrders[d]);
  }
  var byFirstRow = function(a, b) { return orderRows[a][0] - orderRows[b][0]; };
  ebaySeq.sort(byFirstRow);
  directSeq.sort(byFirstRow);
  [ebaySeq, directSeq].forEach(function(seq) {
    // Pass 1 — keepers: a group that already wears a digit keeps it. On a
    // collision (two groups claiming the same digit — only possible via >10
    // concurrent groups cycling, or a hand-pasted format) the topmost group
    // keeps it and the other re-draws as new.
    var used = {};       // glyph index → true
    var assigned = {};   // order → glyph index
    var newcomers = [];
    for (var s = 0; s < seq.length; s++) {
      var kIdx = existingBadge.hasOwnProperty(seq[s]) ? existingBadge[seq[s]] : -1;
      if (kIdx >= 0 && !used[kIdx]) { used[kIdx] = true; assigned[seq[s]] = kIdx; }
      else newcomers.push(seq[s]);
    }
    // Pass 2 — newcomers (top-to-bottom) draw the lowest FREE digit; when
    // all 10 are worn, cycle positionally (same overflow rule as before —
    // a badge only needs to be unique among currently-visible groups).
    var overflow = 0;
    for (var n = 0; n < newcomers.length; n++) {
      var free = -1;
      for (var fi = 0; fi < SO_BADGE_GLYPHS.length; fi++) {
        if (!used[fi]) { free = fi; break; }
      }
      if (free >= 0) { used[free] = true; assigned[newcomers[n]] = free; }
      else assigned[newcomers[n]] = (overflow++) % SO_BADGE_GLYPHS.length;
    }
    // Write every row of every group with its assigned digit.
    for (var w = 0; w < seq.length; w++) {
      var glyph = SO_BADGE_GLYPHS[assigned[seq[w]]];
      var gRows = orderRows[seq[w]];
      for (var g = 0; g < gRows.length; g++) {
        bandFormats[gRows[g] - Schema.dataStartRow][0] = '"' + glyph + ' "@';
      }
    }
  });
  fullRange.setNumberFormats(bandFormats);

  // 3c. Column font: uniform 10px across the band. Heals the 12px/14px
  // residue from the abandoned size-split attempts (see the SO_BADGE_GLYPHS
  // comment for why per-glyph sizing is impossible — the keycap glyphs carry
  // their own visual weight at 10px instead).
  var bandSizes = fullRange.getFontSizes();
  for (var fs = 0; fs < bandSizes.length; fs++) {
    var fsRow = Schema.dataStartRow + fs;
    if (boundary > 0 && (fsRow === boundary || fsRow === boundary + 1)) continue;
    bandSizes[fs][0] = 10;
  }
  fullRange.setFontSizes(bandSizes);

  // COLOR BARS DROPPED (2026-07-14, user's call after living with both):
  // the badge answers "which order" — the bar only ever answered "some
  // group", spent scarce color budget, and died on B&W prints. The band-wide
  // border CLEAR above (step 2) stays so legacy bars wipe themselves on the
  // first repaint. Border-apply loop deleted; git history has it.

  // 3d. DIRECT per-order DIVIDER — a heavy top rule at the start of each
  // sales-order group in the DIRECT table (drawn wherever the SO changes from
  // the row above), so multi-item direct orders separate visually WITHOUT
  // blank rows. Pure formatting keyed to col D → survives sort / shipped-delete
  // / kit expansion / manual edit because this painter re-runs on all those
  // paths. DIRECT ONLY (the eBay table is location-sorted single-item orders —
  // a rule per row would be noise).
  if (boundary > 0) {
    try { _paintDirectOrderDividers(sheet, boundary, allData, lastRow); }
    catch (dividerErr) { console.log("SO divider paint error: " + dividerErr); }
  }

  // Force the clear-then-apply sequence to land before any subsequent reads.
  SpreadsheetApp.flush();
}

/**
 * Draws a heavy top border at the start of each sales-order group in the DIRECT
 * table — the "divider between orders" that retires the manual blank-row habit.
 * A border is drawn on any DIRECT data row whose SALES ORDER differs from the
 * row directly above it. PURE FORMATTING (no rows added) — so it is immune to
 * the sort, the shipped-delete workflow, kit expansion, and manual edits: this
 * painter simply re-runs and re-derives the group starts from col D.
 *
 * @param {Sheet}  sheet     the main sheet
 * @param {number} boundary  the DIRECT divider row
 * @param {Array}  colDData  col-D values from Schema.dataStartRow (reused from the caller)
 * @param {number} lastRow   sheet.getLastRow()
 */
function _paintDirectOrderDividers(sheet, boundary, colDData, lastRow) {
  var firstData = boundary + 2;      // first DIRECT data row
  var clearLast = Math.min(sheet.getMaxRows(), lastRow + 200);
  if (clearLast < firstData) return;
  var W = Schema.dataWidth;
  var bandRows = clearLast - firstData + 1;
  var dataN = lastRow - firstData + 1;

  // ---- Restore the row BANDING under the DIRECT data (clears any prior
  // per-order tint). The block-shade experiment was reverted 2026-08-03: with
  // many small orders it just collapses back into per-row banding, and it
  // overrode the banding (stickier to undo than a border). We keep the clean
  // per-order DIVIDER below — it scales to any number of orders and is trivially
  // reversible. SAFETY: Zoho ⚠-flagged rows keep their soft-red background;
  // every other DIRECT row is set to null → the sheet's banding shows through.
  var noteVals = (dataN > 0) ? sheet.getRange(firstData, Schema.cols.NOTE, dataN, 1).getValues() : [];
  var curBg    = (dataN > 0) ? sheet.getRange(firstData, 1, dataN, W).getBackgrounds() : [];
  var bg = [];
  for (var r = 0; r < bandRows; r++) { var arr = new Array(W); for (var c = 0; c < W; c++) arr[c] = null; bg.push(arr); }
  for (var row = firstData; row <= lastRow; row++) {
    var di = row - firstData;
    var note = String((noteVals[di] && noteVals[di][0]) || '');
    if (note.indexOf('⚠') !== -1) {                  // Zoho flag → preserve its bg
      for (var cf = 0; cf < W; cf++) bg[di][cf] = curBg[di][cf];
    }
    // else: leave bg[di] null → the sheet's banding shows (no tint)
  }
  sheet.getRange(firstData, 1, bandRows, W).setBackgrounds(bg);

  // ---- GOLD BOX per order — box each sales-order group (top + left + right +
  // bottom) in brand gold, so every order is its own contained block that ties
  // into the yellow DIRECT band. Pure formatting; the shared edge between two
  // adjacent orders reads as the divider between them.
  //
  // Clear the box borders across the band FIRST — but leave the range's OUTER
  // TOP (= the header / first-data edge) untouched (top=null) so we never
  // recolor the header's underline. (2026-08-03: on a multi-row range `top`
  // clears only the outer edge; the per-row lines are inner horizontals, so we
  // clear left/right/bottom/vertical/horizontal here.)
  sheet.getRange(firstData, 1, bandRows, W).setBorder(null, false, false, false, false, false);

  var boxColor  = "#c9a227";                             // brand gold — order is whole
  var splitColor = "#b71c1c";                            // alarm red — order is SPLIT
  var boxStyle  = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  var boxW = Schema.cols.LEFT;                           // box the visible DIRECT columns (A..LEFT)

  // ── PASS 1: collect the contiguous blocks, don't draw yet ──────────────────
  // Drawing inside the walk (the old shape) forced a decision about each block
  // before we knew whether its order appears again further down. See the
  // SPLIT-ORDER note below for why that matters.
  var blocks = [];
  var gStart = -1, gSO = null;
  for (var row2 = firstData; row2 <= lastRow + 1; row2++) {   // +1 flushes the final group
    var bi = row2 - Schema.dataStartRow;
    var bso = (row2 <= lastRow && bi >= 0 && bi < colDData.length)
              ? String(colDData[bi][0]).trim() : "";
    if (bso !== gSO) {
      if (gSO && gStart > 0) blocks.push({ so: gSO, start: gStart, end: row2 - 1 });
      gSO = bso;
      gStart = bso ? row2 : -1;
    }
  }

  // ── PASS 2: how many separate blocks does each order occupy? ───────────────
  var blockCount = {};
  for (var b = 0; b < blocks.length; b++) {
    blockCount[blocks[b].so] = (blockCount[blocks[b].so] || 0) + 1;
  }

  // ── PASS 3: draw ──────────────────────────────────────────────────────────
  //
  // ⚠ THE SPLIT-ORDER CASE — this is a SAFETY NET, and it is the reason the
  // painter is two-pass.
  //
  // This painter used to assume every sales order occupies ONE contiguous run of
  // rows, and closed a box on every change of col D. When that assumption broke,
  // it drew TWO CLOSED GOLD BOXES for a single order — and a closed box is the
  // floor's signal for "this is the whole order." A picker working the first box
  // sees it close and reads the order as complete. Observed live 2026-08-07 after
  // a Zoho line was added to an in-progress order: SO-24696 rendered as a 2-row
  // box at the top and a 13-row box further down. That ships 2 of 15 lines.
  //
  // Contiguity is now also protected at the source (inserts land next to their
  // order's existing rows — see _insertAddedItemsToDirect). But that protects
  // only the paths we know about today; a future insert site, a manual row move,
  // a paste or a delete can all break it again. So the painter must never be the
  // thing that renders a broken layout as a trustworthy one.
  //
  // A split order is therefore drawn RED and OPEN-ENDED: no bottom edge except on
  // its final block, no top edge except on its first. It reads as one damaged,
  // continuing thing rather than two tidy complete ones. Same rule the rest of
  // this system runs on — never show a number or a boundary you can't stand
  // behind (OOS BUILDABLE, Kit Health computed price, the ripple's "frees alone").
  var seenBlocks = {};
  for (var k = 0; k < blocks.length; k++) {
    var blk    = blocks[k];
    var isSplit = blockCount[blk.so] > 1;
    var nth     = (seenBlocks[blk.so] = (seenBlocks[blk.so] || 0) + 1);

    // First order on the sheet skips its top edge so we never recolor the
    // header's underline.
    var drawTop = (blk.start !== firstData);
    var top     = drawTop ? (!isSplit || nth === 1) : null;
    var bottom  = (!isSplit || nth === blockCount[blk.so]);

    sheet.getRange(blk.start, 1, blk.end - blk.start + 1, boxW)
         .setBorder(top, true, bottom, true, false, false,
                    isSplit ? splitColor : boxColor, boxStyle);
  }
}

function highlightAllDuplicateSalesOrders() {
  setupDuplicateSalesOrderHighlighting();
  return "✅ Duplicate Sales Order border tabs applied.";
}

/**
 * Removes leftover CF rules from the old background-highlight approach.
 * Safe to call repeatedly — only strips rules targeting column D with TRIM/COUNTIF formulas.
 */
function removeLegacySalesOrderCFRules(sheet) {
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var bc = rules[i].getBooleanCondition();
    if (bc && bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      var values = bc.getCriteriaValues();
      if (values.length > 0) {
        var formula = values[0];
        var ranges = rules[i].getRanges();
        // ⚠⚠ EVERY range must be column D — not merely ONE of them. 2026-08-30: this
        //   read "does ANY range touch col D", and the identity rules span cols A AND D
        //   while their status guard contains TRIM( — so all three were deleted on the
        //   FIRST edit after they were installed. The feature simply never marked
        //   anything, and because conditional formatting is a display layer there was no
        //   residue to notice: an uninstalled rule and a clean sheet look identical.
        //
        //   A LEGACY rule from the old background-highlight approach targeted column D
        //   and nothing else, so requiring that is both tighter and semantically right.
        var isOrderColumn = ranges.length > 0;
        for (var j = 0; j < ranges.length; j++) {
          if (!(ranges[j].getColumn() === Schema.cols.SALES_ORDER &&
                ranges[j].getNumColumns() === 1)) {
            isOrderColumn = false;
            break;
          }
        }
        if (isOrderColumn && (formula.indexOf('COUNTIF') !== -1 || formula.indexOf('TRIM(') !== -1)) {
          continue;  // Skip (remove) this legacy rule
        }
      }
    }
    filtered.push(rules[i]);
  }
  if (filtered.length !== rules.length) {
    sheet.setConditionalFormatRules(filtered);
  }
}

function clearAllDuplicateSalesOrderHighlights() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return "✅ Nothing to clear.";

  var fullRange = sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, lastRow - Schema.dataStartRow + 1, 1);
  fullRange.setBorder(null, false, null, null, null, null);
  return "✅ Duplicate Sales Order border tabs cleared.";
}

// =======================================================================================
// AUTO-REFRESH: Unified handler for both SKU and Sales Order duplicate highlights
// =======================================================================================

/**
 * Refreshes both duplicate highlight systems on any edit to the main sheet.
 * Called from onEditInstallable(e) — triggers on ANY data-area edit,
 * not just column A or D, so highlights update when you clear/edit any cell.
 */
function refreshDuplicateHighlightsOnEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== MAIN_SHEET_NAME) return;
    if (range.getRow() < Schema.dataStartRow) return;
    // Only auto-refresh Sales Order highlights (SKU is manual-only)
    setupDuplicateSalesOrderHighlighting();
  } catch (err) { /* silent */ }
}

// consolidateTable, runMergeEbayDuplicates, runMergeDirectDuplicates —
// REMOVED 2026-04-29.
//
// These were the old "merge duplicate SKU rows" feature. Verified zero callers
// in Apps Script, HTML sidebars, and n8n workflows. The function itself
// carried a self-warning ("May affect n8n duplicate detection") because
// merging rows breaks the SKU+SalesOrder dedup contract that doPost relies on.
//
// Modern duplicate handling lives in:
//   - setupDuplicateSalesOrderHighlighting() above (visual color tabs)
//   - setupDuplicateHighlighting() above (SKU group colors)
// These VISUALIZE duplicates rather than destroying data.