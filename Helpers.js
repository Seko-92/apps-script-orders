// =======================================================================================
// HELPERS.gs - Core Utility Functions//
// =======================================================================================

/**
 * Gets the column index (0-based) for a given header name
 * @param {Sheet} sheet - The sheet to search
 * @param {string} headerName - The header name to find
 * @returns {number} - 0-based column index, or -1 if not found
 */
function getColumnIndexByHeader(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === String(headerName).trim().toLowerCase()) {
      return i;
    }
  }
  return -1;
}

/**
 * Finds the boundary row (the divider whose column A value is Schema.boundaryMarker)
 * @returns {number} - Row number of the boundary, or -1 if not found
 */
function getBoundaryRow() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return -1;
  var values = sheet.getRange(1, Schema.cols.SKU, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === Schema.boundaryMarker) return i + 1;
  }
  return -1;
}

/**
 * Finds the last row with data in a segment
 * @param {number} startRow - Start row of the segment
 * @param {number} endRow - End row of the segment
 * @returns {number} - Last row with data
 */
function findLastDataRowInSegment(startRow, endRow) {
  if (endRow < startRow) return startRow - 1;
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var values = sheet.getRange(startRow, Schema.cols.SKU, endRow - startRow + 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim().length > 0) return startRow + i;
  }
  return startRow - 1;
}

/**
 * Builds a map of SKU → total committed quantity (PENDING + PREPARING orders).
 * Used to subtract from Master Inventory available stock for accurate HAND values.
 * @returns {Map} - Map of SKU (lowercase) → total committed qty
 */
function getCommittedQuantities() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < Schema.dataStartRow) return new Map();

  // Read SKU, QTY, ..., STATUS — first 6 columns cover everything we need
  var data = sheet.getRange(
    Schema.dataStartRow, 1,
    lastRow - Schema.dataStartRow + 1,
    Schema.cols.STATUS
  ).getValues();
  var committed = new Map();

  for (var i = 0; i < data.length; i++) {
    var sku    = String(data[i][Schema.idx("SKU")]).trim().toLowerCase();
    var qty    = parseInt(data[i][Schema.idx("QTY")]) || 0;
    var status = String(data[i][Schema.idx("STATUS")]).trim().toUpperCase();

    if (!sku || sku === Schema.boundaryMarker.toLowerCase()) continue;
    if (status !== Schema.status.PENDING && status !== Schema.status.PREPARING) continue;

    committed.set(sku, (committed.get(sku) || 0) + qty);
  }

  return committed;
}

/**
 * Recomputes the HAND column for every active (non-terminal) order row.
 *
 * HAND = MI.available for the SKU. Same value for every non-terminal row
 * of the same SKU. Tells the picker "what eBay's listing shows as available
 * for new buyers right now."
 *
 * Why no per-row decrement: with the per-order GetItem refresh wired into
 * n8n's eBay-orders workflow, MI is fresh at the moment each order arrives.
 * eBay's QuantitySold already counts every PENDING / PREPARING qty (the sale
 * was registered when the buyer clicked Buy). Decrementing per-row would
 * re-subtract the same units, producing values lower than reality —
 * exactly the symptom that surfaced 2026-05-09 for SKU 165447 (HAND=158
 * when MI.available was 162; the 4-unit gap matched the SKU's pre-existing
 * PENDING qty).
 *
 * Idempotent. Run from sidebar, or schedule via setupHandRecomputeTrigger()
 * for automatic refresh after each MI sync.
 */
/**
 * @param {{inventoryMap:Map}=} sharedMaps  a Master Inventory snapshot already built
 * @param {Map=} sharedZoho                 a Zoho Stock map already built
 *
 * ⚠ BOTH ARE SHAPE-DETECTED, NOT TRUSTED BY POSITION. recomputeHand is a TIME
 *   TRIGGER target (setupHandRecomputeTrigger), and a trigger hands its function
 *   an EVENT OBJECT as the first argument — so a positional `maps` would be an
 *   event on every scheduled run. Same guard refreshOutOfStock uses.
 *
 * ⚠ WHY SHARING MATTERS (measured 2026-08-19). buildLocationAndInventoryMaps is
 *   1,872 ms — getDataRange on 3,635 × 198 — and buildZohoStockMap is 415 ms.
 *   The every-2-minutes Zoho job calls recomputeHand AND refreshPrepQueueHand
 *   back to back, and each built its own copy of both: 2,287 ms of the 5,789 ms
 *   cycle spent reading the same two sheets twice. Over a 9–5 day that is
 *   ~9 minutes of the ~90-minute budget, thrown away.
 *
 * ⚠ AND NARROWING THE READ DOES NOT HELP — TESTED, DO NOT RETRY. Reading only
 *   the 39-column span the headers occupy instead of all 198 measured 1,905 ms
 *   against 1,872 ms: no faster, fractionally slower, with identical maps. In
 *   Apps Script the ROUND TRIP dominates, not the payload — the same finding the
 *   PartConsole work hit in July when its "lean" four-column read barely helped.
 *   The only way to make this cheaper is to do it LESS OFTEN, not more narrowly.
 */
function recomputeHand(sharedMaps, sharedZoho) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return "ℹ️ No data to recompute.";

  // Live snapshot of Master Inventory — reused when the caller already has one.
  var maps = (sharedMaps && sharedMaps.inventoryMap) ? sharedMaps
                                                     : buildLocationAndInventoryMaps();
  var inventoryMap = maps.inventoryMap;
  if (inventoryMap.size === 0) {
    return "⚠️ Master Inventory empty or headers missing.";
  }

  // Read SKU/QTY/STATUS for every data row
  var nRows = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, nRows, Schema.cols.HAND).getValues();

  var boundary = getBoundaryRow();
  // Zoho stock mirror (SKU → {available}). Empty map if the sheet doesn't exist
  // yet → every row falls back to MI, i.e. identical to pre-Zoho behavior.
  var zohoMap = (sharedZoho && typeof sharedZoho.get === 'function') ? sharedZoho
                                                                     : buildZohoStockMap();
  var newHandValues = [];      // 2D array (one cell per row) for batched setValues
  var updatedCount = 0;
  var zohoSourced = 0;         // how many rows took their HAND from Zoho (debug signal)

  for (var i = 0; i < nRows; i++) {
    var rowNum = Schema.dataStartRow + i;
    var rawSku = String(data[i][Schema.idx("SKU")] || "").trim();

    // Skip boundary divider + DIRECT header row + empty rows
    var skipRow = false;
    if (boundary > 0 && (rowNum === boundary || rowNum === boundary + 1)) skipRow = true;
    if (!rawSku) skipRow = true;
    if (rawSku.toUpperCase() === Schema.boundaryMarker) skipRow = true;
    if (rawSku.toUpperCase().indexOf('SKU') === 0) skipRow = true;  // header leak guard

    if (skipRow) {
      newHandValues.push([data[i][Schema.idx("HAND")]]);   // preserve as-is
      continue;
    }

    var skuLower = rawSku.toLowerCase();
    var status = String(data[i][Schema.idx("STATUS")] || "").trim().toUpperCase();

    // Terminal-state rows keep their historical HAND — Master Inventory's
    // quantitySold (and Zoho's committed netting) already account for them.
    if (Schema.isTerminal(status)) {
      newHandValues.push([data[i][Schema.idx("HAND")]]);
      continue;
    }

    // Source routing (see resolveHandValue in ZohoStock.js):
    //   DIRECT-table row (below the boundary header) → Zoho first.
    //   Manually-typed eBay row (SALES ORDER isn't a clean eBay order id, e.g.
    //     a "Replacement #: …" row)                  → Zoho first.
    //   Automated eBay-order row (clean order id)    → MI first (eBay truth).
    // No decrement either way: MI.available and Zoho.available_stock both
    // already net committed qty (per the 2026-05-09 HAND semantics).
    var isDirect = (boundary > 0 && rowNum > boundary + 1);
    var preferZoho = isDirect || _isManualSalesOrder(data[i][Schema.idx("SALES_ORDER")]);
    var miInv  = inventoryMap.get(skuLower);
    var miAvail = miInv ? miInv.available : null;
    var zo     = zohoMap.get(skuLower);
    var zoAvail = zo ? zo.available : null;

    var hand = resolveHandValue(miAvail, zoAvail, preferZoho);
    if (preferZoho ? (zoAvail != null) : (miAvail == null && zoAvail != null)) zohoSourced++;

    newHandValues.push([hand]);
    updatedCount++;
  }

  // Single batched write to col G (HAND)
  sheet.getRange(Schema.dataStartRow, Schema.cols.HAND, nRows, 1).setValues(newHandValues);

  return "✅ HAND recomputed for " + updatedCount + " active row(s)" +
         (zohoSourced ? " (" + zohoSourced + " from Zoho)" : "") + ".";
}


/**
 * Run ONCE from the Apps Script Editor (or sidebar) to install a 15-minute
 * trigger that auto-runs recomputeHand(). Pairs naturally with the n8n
 * Master Inventory full-sync workflow that runs every 15 min — this picks
 * up the fresh quantitySold data and rewrites HAND across the sheet without
 * manual clicks.
 *
 * Idempotent — removes any existing recomputeHand trigger before creating
 * the new one, so re-running won't pile up duplicates.
 *
 * Quota cost: ~96 invocations/day × ~2s avg = 3 min/day execution time.
 * Well within Apps Script's 90 min/day allowance.
 */
function setupHandRecomputeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'recomputeHand') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('recomputeHand')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("HAND recompute trigger installed: every 15 minutes");
  try {
    SpreadsheetApp.getUi().alert(
      "Trigger Installed",
      "HAND will auto-recompute every 15 minutes — pairs with the 15-min " +
      "Master Inventory sync.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { /* no UI context */ }
}

/** Removes the auto-recompute trigger. Manual cleanup helper. */
function removeHandRecomputeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'recomputeHand') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log("Removed " + removed + " recomputeHand trigger(s).");
  return "Removed " + removed + " trigger(s).";
}


/**
 * Sets up conditional formatting on the HAND column (Column G)
 * so low-stock highlighting is ALWAYS accurate and never stale.
 * Replaces all manual setBackground calls.
 */
function setupHandConditionalFormatting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  // Remove any existing HAND highlight rules to avoid duplicates
  var rules = sheet.getConditionalFormatRules();
  var filtered = [];
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var isHandRule = false;
    for (var j = 0; j < ranges.length; j++) {
      if (ranges[j].getColumn() === Schema.cols.HAND && ranges[j].getNumColumns() === 1) {
        isHandRule = true;
        break;
      }
    }
    if (!isHandRule) filtered.push(rules[i]);
  }

  // Clear stale manual backgrounds ONLY on data rows (skip boundary + header).
  // Important after the v6 cutover: the OLD rule painted a red bg, so any cell
  // that ever tripped low-stock under the old rule kept that bg even after the
  // CF rule was swapped to font-only. Strip those legacy bgs here so the new
  // font-only treatment isn't visually drowned out by leftover red fills.
  var boundary = getBoundaryRow();
  var lastRow = Math.max(sheet.getLastRow(), Schema.dataStartRow);

  // eBay data rows
  if (boundary > Schema.dataStartRow) {
    var ebayCount = boundary - 1 - Schema.dataStartRow + 1;
    if (ebayCount > 0) {
      sheet.getRange(Schema.dataStartRow, Schema.cols.HAND, ebayCount, 1).setBackground(null);
    }
  }

  // DIRECT data rows (skip boundary row and DIRECT header row)
  if (boundary > 0 && boundary + 2 <= lastRow) {
    var directStart = boundary + 2;
    var directCount = lastRow - directStart + 1;
    if (directCount > 0) {
      sheet.getRange(directStart, Schema.cols.HAND, directCount, 1).setBackground(null);
    }
  }

  // Build conditional formatting rule — Service Bay v6 font-only treatment.
  // Mirrors _buildHandLowStockRule in BrandTheme.js: dark red `#b71c1c` + bold,
  // NO background. Cell backgrounds are reserved for the highest-priority
  // alerts (STATUS + paid SHIP COST). HAND stays a "noted but not screaming"
  // secondary signal — bold font compensates for the lost bg.
  var handRange = sheet.getRange(Schema.dataStartRow, Schema.cols.HAND, 1000, 1);
  var formula = "=AND(ISNUMBER(G" + Schema.dataStartRow + "), G" + Schema.dataStartRow + "<=20)";
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setFontColor("#b71c1c")
    .setBold(true)
    .setRanges([handRange])
    .build();

  filtered.push(rule);
  sheet.setConditionalFormatRules(filtered);
}

/**
 * Gets the live update state from Settings sheet
 * @returns {string} - "ON" or "OFF"
 */
function getLiveUpdateState() {
  var s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LIVE_UPDATE_SHEET);
  return s ? s.getRange(LIVE_UPDATE_TOGGLE_CELL).getValue() : "OFF";
  // (Note: LIVE_UPDATE_TOGGLE_CELL is "B1" in the Settings sheet —
  //  not in Schema because Schema is for the All Orders sheet's structure)
}

/**
 * Toggles the live update state
 * @param {string} st - "ON" or "OFF"
 * @returns {string} - The new state
 */
function toggleLiveUpdate(st) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var s = ss.getSheetByName(LIVE_UPDATE_SHEET) || ss.insertSheet(LIVE_UPDATE_SHEET).hideSheet();
  s.getRange(LIVE_UPDATE_TOGGLE_CELL).setValue(st);
  return st;
}

/**
 * Updates Master Inventory rows by itemId with fresh qty/quantitySold/qtyLastSync.
 * Called by doPost when n8n sends action=updateMiRows. Skips itemIds not found
 * in MI (defensive — never inserts new rows; appendOrUpdate could pollute the
 * 174-column structure with mostly-empty rows).
 *
 * Input: rows = [{ itemId, sku?, quantity, quantitySold }]
 * Returns: { updated, notFound }
 *
 * Use case: n8n's eBay-orders workflow calls this per batch BEFORE doPost
 * inserts the order. Net effect: by the time doPost reads MI for HAND, the
 * affected SKUs are already current. No more "ordered N units, sheet still
 * shows pre-sale qty" gap.
 */
function updateMiRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return { updated: 0, notFound: 0 };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var miSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!miSheet) {
    throw new Error(DB_SHEET_NAME + " sheet not found");
  }

  var lastRow = miSheet.getLastRow();
  var lastCol = miSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { updated: 0, notFound: rows.length };
  }

  var headers = miSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Find columns by header name (case-insensitive trim).
  function findCol(name) {
    var target = String(name).trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toLowerCase() === target) return i + 1;
    }
    return -1;
  }

  var itemIdCol   = findCol('itemId');
  var qtyCol      = findCol(DB_QUANTITY_HEADER);       // "quantity"
  var qtySoldCol  = findCol(DB_QUANTITY_SOLD_HEADER);  // "quantitySold"
  var lastSyncCol = findCol('qtyLastSync');             // optional — only stamped if present

  if (itemIdCol < 0 || qtyCol < 0 || qtySoldCol < 0) {
    throw new Error("Required MI columns not found (need: itemId, " +
      DB_QUANTITY_HEADER + ", " + DB_QUANTITY_SOLD_HEADER + ")");
  }

  // Build itemId → row map (one read of the itemId column).
  var itemIdValues = miSheet.getRange(2, itemIdCol, lastRow - 1, 1).getValues();
  var rowByItemId = {};
  for (var r = 0; r < itemIdValues.length; r++) {
    var id = String(itemIdValues[r][0] || '').trim();
    if (id) rowByItemId[id] = r + 2;  // sheet row number, accounting for header
  }

  var nowIso = new Date().toISOString();
  var updated = 0, notFound = 0;

  for (var k = 0; k < rows.length; k++) {
    var input = rows[k] || {};
    var targetId = String(input.itemId || '').trim();
    if (!targetId) { notFound++; continue; }

    var rowNum = rowByItemId[targetId];
    if (!rowNum) { notFound++; continue; }

    miSheet.getRange(rowNum, qtyCol).setValue(parseInt(input.quantity, 10) || 0);
    miSheet.getRange(rowNum, qtySoldCol).setValue(parseInt(input.quantitySold, 10) || 0);
    if (lastSyncCol > 0) {
      miSheet.getRange(rowNum, lastSyncCol).setValue(nowIso);
    }
    updated++;
  }

  // Force the writes to land before subsequent reads (the doPost insert that
  // triggered this call is going to read MI in its next step).
  SpreadsheetApp.flush();

  return { updated: updated, notFound: notFound };
}


// =======================================================================================
// SHELF-LOCATION ORDERING  (natural, not lexical)
// =======================================================================================
//
// FOUND 2026-08-05 by reading a /oos reply top-to-bottom in Telegram:
//
//     A-30 · A-36 · A-47 · A-50 · A-9        ← 9 lands AFTER 50
//     B-22 … B-52 · B-8                      ← same
//
// Plain string comparison puts "A-9" after "A-50" because "9" > "5" character
// by character. Every surface that ordered by LOCATION did this, so the
// restock walk has been wrong at every single-digit shelf — the list's entire
// job is "walk the aisle in order," and it wasn't.
//
// It only became visible in Telegram because the sheet is usually scanned, not
// read straight down. Worth remembering: a new surface onto old data is a
// cheap way to find bugs the original surface was hiding.
//
// This compares the AISLE LETTER first, then the shelf NUMBER numerically, then
// any trailing note. It deliberately does NOT handle "blank / NOT FOUND last" —
// every caller already does that before reaching here, and folding two rules
// into one comparator makes both harder to reason about.
//
// Handles the warehouse's real notations, incl. the pack-size suffix that
// appears in ~59 MI locations ("G-35 * 2" — one sellable unit is pre-packed as
// N pieces; an assurance note, never a qty multiplier).

/**
 * Split "G-35 * 2" into {prefix:"G", num:35, rest:"* 2"}.
 * @returns {?{prefix:string,num:number,rest:string}} null when it isn't
 *          letter-then-number shaped (e.g. "NOT FOUND", "", "K55B").
 */
function _parseShelfLocation(v) {
  var s = String(v == null ? "" : v).trim();
  if (!s) return null;
  var m = s.match(/^([A-Za-z]+)\s*[-–—]?\s*(\d+)(.*)$/);
  if (!m) return null;
  return {
    prefix: m[1].toUpperCase(),
    num:    parseInt(m[2], 10),
    rest:   m[3].trim().toUpperCase()
  };
}

/**
 * Compare two shelf locations in the order a picker physically walks them.
 * Unparseable values sort after parseable ones, then fall back to plain text
 * so the ordering is still stable and deterministic.
 *
 * @returns {number} negative / 0 / positive, suitable for Array.prototype.sort
 */
/**
 * THE ONE PARSER for "is this row a kit component, and whose?"
 *
 * Returns the PARENT KIT SKU, or "" if the row is not an expansion component.
 *
 * ⚠ THERE ARE TWO TAG SHAPES AND BOTH ARE REAL. KitExpansion writes
 * `↳ from KIT-<sku>` for registered components and `↳ added to KIT-<sku>` for
 * CUSTOM ADDS (the modal's "+ add component"). Matching only the first is a
 * live bug that has now been found twice in two days:
 *   · _dashOpenOrders — a custom-added part was missing from the board's
 *     `done/total`, so a kit could read "5 of 5" with a sixth part on its shelf
 *   · _kitParentFollowUp — worse: the unmatched row fell through to the
 *     candidate-PARENT branch, so the parent could flip to SHIPPED while a
 *     custom-added component was still open
 *
 * ⚠ SURVIVES A ZOHO FLAG. _flagDirectRow PREPENDS its warning as its own first
 * LINE and CASCADES onto a removed kit's components, so a flagged component's
 * note starts with "⚠️" and the tag sits on line 2. Anchoring at ^ without
 * skipping that line drops the row out of its kit entirely.
 *
 * ⚠ RETURNS THE CAPTURE — callers must compare it EXACTLY, never with a prefix
 * test. `note.indexOf("↳ from KIT-" + sku) === 0` lets kit `1586` match kit
 * `158652`'s components. (Latent rather than live while every SKU is 6 digits,
 * since no 6-digit SKU can prefix a different one — but do not rely on that.)
 *
 * ⚠ ONE IMPLEMENTATION ON PURPOSE. The KitRegistry parser was once duplicated
 * across the CSV importer and the webhook handler, and "every past fix had to
 * land twice". This is the same tag read in at least seven places.
 *
 * @param {string} note the row's NOTE cell
 * @returns {string} parent kit SKU, or ""
 */
function kitComponentTag(note) {
  var s = String(note || "").trim();
  if (s.charAt(0) === "⚠") {              // ⚠ — a Zoho flag line, skip it
    var nl = s.indexOf("\n");
    s = (nl === -1) ? "" : s.slice(nl + 1).trim();
  }
  var m = s.match(/^↳ (?:from|added to) KIT-(\S+)/);
  return m ? m[1] : "";
}


function compareLocations(a, b) {
  var pa = _parseShelfLocation(a);
  var pb = _parseShelfLocation(b);

  if (pa && pb) {
    if (pa.prefix !== pb.prefix) return pa.prefix < pb.prefix ? -1 : 1;
    if (pa.num    !== pb.num)    return pa.num - pb.num;
    if (pa.rest   !== pb.rest)   return pa.rest < pb.rest ? -1 : 1;
    return 0;
  }
  if (pa && !pb) return -1;
  if (!pa && pb) return 1;

  var sa = String(a == null ? "" : a).trim().toUpperCase();
  var sb = String(b == null ? "" : b).trim().toUpperCase();
  return sa < sb ? -1 : (sa > sb ? 1 : 0);
}


// =======================================================================================
// TIMING — where does the every-2-minutes job actually spend itself?
// Added 2026-08-19. Editor-run, prints to the EXECUTION LOG.
// =======================================================================================
//
// WHY THIS EXISTS. The Zoho stock sync fires writeZohoStock every 2 minutes, 9–5,
// and that one call does three things: write ~28,000 cells to the Zoho Stock
// sheet, then recomputeHand, then refreshPrepQueueHand. BOTH recomputes call
// buildLocationAndInventoryMaps() independently, and that does getDataRange() on
// Master Inventory — 3,655 rows × 198 columns ≈ 724,000 cells — to use FIVE of
// those columns. Twice. Every two minutes.
//
// So the plan to move the Zoho WRITE to n8n would move ~2% of the job and leave
// the rest. That is worth knowing before spending half a day on it.
//
// ⚠ BUT CELL COUNTS ARE NOT SECONDS. Apps Script does not scale linearly — the
//   round trip dominates a small read, the payload dominates a big one. Every
//   claim above is an inference from SIZE. This function replaces it with a
//   measurement, because this project has twice been wrong optimising from a
//   guess (the "lean" four-column read that barely helped, and the two wrong
//   calls chased while inferring publish causes).
//
// SAFE TO RUN ANY TIME: the two candidate reads are READ-ONLY, and they are
// compared for exact equality so a faster read that is also a WRONG read cannot
// pass unnoticed. recomputeHand / refreshPrepQueueHand DO write, but they are
// idempotent and already run every 2 minutes — running one more costs nothing.
// =======================================================================================

/**
 * EDITOR-RUN. Pick timeHandPath in the function dropdown, press Run, then read
 * the EXECUTION LOG — the Run button does not display return values.
 *
 * ⚠⚠ IT TAKES SEVERAL SAMPLES ON PURPOSE. Two runs 30 minutes apart on an idle
 *    script measured the same Master Inventory read at 1,872 ms and 4,943 ms —
 *    2.2x apart, with nothing on the floor and nothing changed. Apps Script
 *    timing is that noisy, so ONE sample cannot support a decision. This reports
 *    min / median / max and projects from the MEDIAN.
 *
 * ⚠ IT ALSO ALTERNATES the shared and unshared order between passes. Reading the
 *   same ranges repeatedly inside one execution warms them, so always measuring
 *   "shared" last would systematically flatter it — which is exactly what the
 *   first version of this function did.
 *
 * Safe any time: recomputeHand / refreshPrepQueueHand are idempotent and already
 * run every 2 minutes. Run it during WORK HOURS too — contention is the case
 * that decides anything.
 */
function timeHandPath() {
  var PASSES = 3;
  var L = [];
  function say(x) { L.push(x); console.log(x); }
  function ms(fn) { var t = Date.now(); fn(); return Date.now() - t; }
  function stats(a) {
    var s = a.slice().sort(function (x, y) { return x - y; });
    return { min: s[0], med: s[(s.length - 1) >> 1], max: s[s.length - 1] };
  }
  function line(label, st) {
    say("  " + label.padEnd(42) + String(st.med).padStart(6) + " ms" +
        "   (min " + st.min + " / max " + st.max + ")");
  }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DB_SHEET_NAME);
  var rows = sheet ? sheet.getLastRow() - 1 : 0, cols = sheet ? sheet.getLastColumn() : 0;
  say("=== HAND PATH TIMING · " + PASSES + " passes · " + new Date().toISOString() + " ===");
  say("Master Inventory: " + rows + " rows x " + cols + " cols = " + (rows * cols) + " cells");
  say("");

  var mi = [], zo = [], unshared = [], shared = [];
  for (var p = 0; p < PASSES; p++) {
    mi.push(ms(function () { buildLocationAndInventoryMaps(); }));
    zo.push(ms(function () { buildZohoStockMap(); }));

    // alternate, so neither arrangement always gets the warm cache
    if (p % 2 === 0) {
      unshared.push(ms(function () { recomputeHand(); refreshPrepQueueHand(); }));
      shared.push(ms(function () {
        var m = buildLocationAndInventoryMaps(), z = buildZohoStockMap();
        recomputeHand(m, z); refreshPrepQueueHand(m, z);
      }));
    } else {
      shared.push(ms(function () {
        var m = buildLocationAndInventoryMaps(), z = buildZohoStockMap();
        recomputeHand(m, z); refreshPrepQueueHand(m, z);
      }));
      unshared.push(ms(function () { recomputeHand(); refreshPrepQueueHand(); }));
    }
  }

  line("buildLocationAndInventoryMaps (" + cols + " cols)", stats(mi));
  line("buildZohoStockMap", stats(zo));
  say("");
  line("per cycle, UNSHARED (each builds its own)", stats(unshared));
  line("per cycle, SHARED (one build, handed to both)", stats(shared));
  var u = stats(unshared).med, sh = stats(shared).med;
  say("  → saved per cycle (median):              " + (u - sh) + " ms  (" +
      Math.round(100 * (u - sh) / Math.max(1, u)) + "%)");
  say("");

  say("  PROJECTION over a 9-5 day, from the MEDIAN:");
  say("     interval    cycles     today      shared     saved");
  [2, 3, 4, 5, 6, 8, 10].forEach(function (m) {
    var c = Math.round(8 * 60 / m);
    say("      " + (m + " min").padEnd(11) + String(c).padEnd(10) +
        (u * c / 60000).toFixed(1).padEnd(11) +
        (sh * c / 60000).toFixed(1).padEnd(11) +
        ((u - sh) * c / 60000).toFixed(1) + "  min/day");
  });
  say("");
  say("  ⚠ Spread matters more than any single figure — if min and max are far");
  say("    apart the script is contended and the median is the honest number.");
  return L.join("\n");
}

