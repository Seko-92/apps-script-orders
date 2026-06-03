// =======================================================================================
// ZohoStock.js — Zoho stock mirror for live-ish HAND on DIRECT / Prep / non-eBay items
// Shipped 2026-05-28 (Slice 1 of the "live HAND" plan).///////////////
// =======================================================================================
//
// WHY
// ---
// HAND in the eBay table is kept fresh by the eBay-orders workflow (per-order GetItem
// refresh writes MI). But MI is an EBAY snapshot, so two surfaces read stale-or-zero:
//   - DIRECT table rows  — many DIRECT items aren't even listed on eBay → no MI row → 0.
//   - Prep Queue rows    — snapshot frozen at entry time.
//   - Manual entries of non-eBay items — no MI row.
//
// Zoho is the inventory MASTER (it pushes stock → eBay), so Zoho's `available_stock`
// is the authoritative number for direct sales and for items that never hit eBay.
//
// WHAT THIS IS
// -----------
// A SKU-keyed mirror of Zoho's whole active catalog, refreshed wholesale from the
// existing "Zoho Items Bulk Fetch Proxy" (~4,000 items in ~20 Zoho API calls). We
// join by SKU because Zoho's `reference_id` (which would bridge to eBay's itemId)
// comes back EMPTY — SKU is the only reliable key.
//
// MI IS NEVER TOUCHED. This is a separate sheet. `recomputeHand` routes each row to
// the right source (see Helpers.js):
//   - eBay-table row : MI.available if SKU in MI, else Zoho.
//   - DIRECT row     : Zoho if SKU in Zoho Stock, else MI.
//   - Prep Queue     : Zoho, else MI.
//
// SEMANTICS — no double-count. Zoho's `available_stock` already nets committed open
// SOs, and our DIRECT rows ARE those SOs. So HAND = available_stock DIRECTLY, no
// decrement — the same rule the eBay side adopted 2026-05-09 (HAND = MI.available).
//
// DATA SOURCE FIELD: the bulk-fetch node must keep `available_stock` (and we keep
// `stock_on_hand` so committed = on_hand − available is eyeball-able on the sheet).
// =======================================================================================


var ZOHO_STOCK = {
  sheetName: "Zoho Stock",

  cols: {
    SKU:           1,   // A
    ITEM_NAME:     2,   // B — Zoho item name; lets Price Audit read names from this sheet
    ITEM_ID:       3,   // C — Zoho internal id (not eBay's)
    AVAILABLE:     4,   // D — available_stock (net of committed SOs) — the HAND source
    ON_HAND:       5,   // E — stock_on_hand (gross physical) — for committed visibility
    SELLING_PRICE: 6,   // F — Zoho rate; lets Price Audit read from this sheet (no own fetch)
    SYNCED:        7    // G — wholesale refresh timestamp
  },

  dataWidth:    7,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["SKU", "ITEM NAME", "ITEM ID", "AVAILABLE", "ON HAND", "SELLING PRICE", "SYNCED"]
};


// =======================================================================================
// SETUP — idempotent
// =======================================================================================

function setupZohoStockSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
  if (!sheet) sheet = ss.insertSheet(ZOHO_STOCK.sheetName);

  // --- HEADERS ---
  sheet.getRange(ZOHO_STOCK.headerRow, 1, 1, ZOHO_STOCK.dataWidth)
    .setValues([ZOHO_STOCK.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(ZOHO_STOCK.headerRow, 1, 1, ZOHO_STOCK.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(ZOHO_STOCK.cols.SKU,           110);
  sheet.setColumnWidth(ZOHO_STOCK.cols.ITEM_NAME,     320);
  sheet.setColumnWidth(ZOHO_STOCK.cols.ITEM_ID,       190);
  sheet.setColumnWidth(ZOHO_STOCK.cols.AVAILABLE,     100);
  sheet.setColumnWidth(ZOHO_STOCK.cols.ON_HAND,       100);
  sheet.setColumnWidth(ZOHO_STOCK.cols.SELLING_PRICE, 110);
  sheet.setColumnWidth(ZOHO_STOCK.cols.SYNCED,        150);

  // --- DATA AREA FORMATS ---
  var maxDataRow = 12000;   // generous — catalog ~4K, headroom for growth
  var dataRows = maxDataRow - ZOHO_STOCK.dataStartRow + 1;

  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.ITEM_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.ITEM_ID, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.AVAILABLE, dataRows, 1)
    .setNumberFormat('0').setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.ON_HAND, dataRows, 1)
    .setNumberFormat('0').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.SELLING_PRICE, dataRows, 1)
    .setNumberFormat('$#,##0.00').setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.SYNCED, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');

  sheet.getRange(ZOHO_STOCK.dataStartRow, 1, dataRows, ZOHO_STOCK.dataWidth)
    .setVerticalAlignment('middle');

  // --- BANDING ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, ZOHO_STOCK.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- AVAILABLE low/zero highlight (font-only, matches HAND alert family) ---
  var existingRules = sheet.getConditionalFormatRules() || [];
  var keep = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    return !ranges.some(function(rg) {
      return rg.getSheet().getName() === ZOHO_STOCK.sheetName &&
             rg.getColumn() === ZOHO_STOCK.cols.AVAILABLE;
    });
  });
  var availRange = sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.AVAILABLE, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setFontColor('#b71c1c').setBold(true)
    .setRanges([availRange]).build());
  sheet.setConditionalFormatRules(keep);

  sheet.setFrozenRows(1);

  return "✅ Zoho Stock sheet ready.";
}


/** Sidebar: switch view to the Zoho Stock sheet. */
function openZohoStock() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
  if (!sheet) {
    setupZohoStockSheet();
    sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened Zoho Stock";
}


/**
 * Wholesale-rewrite the Zoho Stock sheet from an items array (each item:
 * { sku, item_name, item_id, available_stock, stock_on_hand, selling_price }).
 * Called by the writeZohoStock doPost action (the scheduled n8n "Zoho Stock
 * Sync" push). Clears prior data, writes fresh, stamps SYNCED.
 * Returns {count, skipped}.
 */
function _writeZohoStockRows(items) {
  if (!Array.isArray(items)) items = [];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
  if (!sheet) {
    setupZohoStockSheet();
    sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
  }

  var now = new Date();
  var rows = [];
  var skipped = 0;
  for (var i = 0; i < items.length; i++) {
    var z = items[i] || {};
    var sku = String(z.sku || "").trim();
    if (!sku) { skipped++; continue; }     // can't key without a SKU
    rows.push([
      sku,                                        // A: SKU
      String(z.item_name || ""),                  // B: ITEM NAME
      String(z.item_id || ""),                    // C: ITEM ID
      parseFloat(z.available_stock) || 0,         // D: AVAILABLE (the HAND source)
      parseFloat(z.stock_on_hand)   || 0,         // E: ON HAND
      parseFloat(z.selling_price)   || 0,         // F: SELLING PRICE
      now                                         // G: SYNCED
    ]);
  }

  // Wipe prior data (preserve headers + formats), then write fresh.
  var lastRow = sheet.getLastRow();
  if (lastRow >= ZOHO_STOCK.dataStartRow) {
    sheet.getRange(ZOHO_STOCK.dataStartRow, 1,
                   lastRow - ZOHO_STOCK.dataStartRow + 1, ZOHO_STOCK.dataWidth)
         .clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(ZOHO_STOCK.dataStartRow, 1, rows.length, ZOHO_STOCK.dataWidth)
         .setValues(rows);
  }
  SpreadsheetApp.flush();

  return { count: rows.length, skipped: skipped };
}


/**
 * Recompute HAND on BOTH All Orders (recomputeHand) and Prep Queue
 * (refreshPrepQueueHand) from the Zoho Stock sheet AS-IS — does NOT pull from
 * Zoho. This is the sidebar button (repurposed 2026-05-28).
 *
 * WHY no pull: the scheduled n8n workflow ("Zoho Stock Sync Scheduled") owns
 * sheet freshness — it pushes the whole catalog every 2 min during work hours
 * and recomputes server-side. The old button pulled via the bulk-fetch proxy,
 * but that proxy only returns price/name fields, NOT available_stock /
 * stock_on_hand — clicking it would have written 0 to every AVAILABLE/ON_HAND
 * cell. So the button no longer pulls; it just re-derives HAND from whatever
 * the scheduled push last wrote. For a true on-demand live pull, fire the
 * scheduled n8n workflow manually (▶ Execute Workflow).
 *
 * Returns { ok, message, handMessage, prepMessage } for the sidebar.
 */
function recomputeHandFromZohoStock() {
  var handMsg = "", prepMsg = "";
  try { handMsg = recomputeHand(); }        catch (e) { handMsg = "HAND recompute error: " + e; }
  try { prepMsg = refreshPrepQueueHand(); } catch (e) { prepMsg = "Prep refresh error: " + e; }
  return {
    ok:          true,
    message:     "HAND recomputed · " + handMsg + " · " + prepMsg,
    handMessage: handMsg,
    prepMessage: prepMsg
  };
}


// =======================================================================================
// READ HELPERS — consumed by recomputeHand / LiveSync / Prep Queue
// =======================================================================================

/**
 * Build a SKU(lowercased) → { available, onHand, itemId, itemName, sellingPrice,
 * skuOriginal } map from the Zoho Stock sheet. Returns an EMPTY map if the
 * sheet is missing/empty — HAND callers then fall back to MI, so behavior
 * matches pre-Zoho.
 *
 * The audit consumer needs itemName/sellingPrice/skuOriginal (the original
 * case-preserved SKU for sheet display). HAND consumers only read .available
 * and are unaffected by the added fields.
 */
function buildZohoStockMap() {
  var map = new Map();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < ZOHO_STOCK.dataStartRow) return map;

    var n = lastRow - ZOHO_STOCK.dataStartRow + 1;
    var data = sheet.getRange(ZOHO_STOCK.dataStartRow, 1, n, ZOHO_STOCK.dataWidth).getValues();
    for (var i = 0; i < data.length; i++) {
      var skuRaw = String(data[i][ZOHO_STOCK.cols.SKU - 1] || "").trim();
      if (!skuRaw) continue;
      map.set(skuRaw.toLowerCase(), {
        skuOriginal:  skuRaw,
        itemName:     String(data[i][ZOHO_STOCK.cols.ITEM_NAME - 1] || ""),
        itemId:       String(data[i][ZOHO_STOCK.cols.ITEM_ID - 1] || ""),
        available:    parseFloat(data[i][ZOHO_STOCK.cols.AVAILABLE - 1])     || 0,
        onHand:       parseFloat(data[i][ZOHO_STOCK.cols.ON_HAND - 1])       || 0,
        sellingPrice: parseFloat(data[i][ZOHO_STOCK.cols.SELLING_PRICE - 1]) || 0
      });
    }
  } catch (e) {
    try { console.log("buildZohoStockMap error: " + e); } catch (_) {}
  }
  return map;
}


/**
 * Returns the SYNCED timestamp of the Zoho Stock sheet (any data row — all
 * rows share the same wholesale-rewrite timestamp). Returns null if the sheet
 * is missing/empty or the cell isn't a Date. Used by Price Audit to surface
 * data freshness to the picker.
 */
function getZohoStockSyncedAt() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
    if (!sheet) return null;
    if (sheet.getLastRow() < ZOHO_STOCK.dataStartRow) return null;
    var v = sheet.getRange(ZOHO_STOCK.dataStartRow, ZOHO_STOCK.cols.SYNCED).getValue();
    return (v instanceof Date) ? v : null;
  } catch (e) {
    try { console.log("getZohoStockSyncedAt error: " + e); } catch (_) {}
    return null;
  }
}


/**
 * Single-SKU Zoho stock lookup (for per-keystroke Prep Queue preview + manual
 * single-cell edits). Returns { available, onHand } or null if not in Zoho.
 * The Zoho Stock sheet is narrow (5 cols), so this read is cheap.
 */
function getSingleZohoStock(skuLower) {
  skuLower = String(skuLower || "").trim().toLowerCase();
  if (!skuLower) return null;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ZOHO_STOCK.sheetName);
    if (!sheet) return null;
    var lastRow = sheet.getLastRow();
    if (lastRow < ZOHO_STOCK.dataStartRow) return null;

    var n = lastRow - ZOHO_STOCK.dataStartRow + 1;
    var data = sheet.getRange(ZOHO_STOCK.dataStartRow, 1, n, ZOHO_STOCK.dataWidth).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === skuLower) {
        return {
          available: parseFloat(data[i][ZOHO_STOCK.cols.AVAILABLE - 1]) || 0,
          onHand:    parseFloat(data[i][ZOHO_STOCK.cols.ON_HAND - 1])   || 0
        };
      }
    }
  } catch (e) {
    try { console.log("getSingleZohoStock error: " + e); } catch (_) {}
  }
  return null;
}


/**
 * THE single source-routing decision for HAND. Shared by recomputeHand,
 * LiveSync (manual entry), Prep Queue, and the DIRECT pull insert.
 *
 * @param {number|null} miAvail    MI.available, or null if SKU not in MI
 * @param {number|null} zoAvail    Zoho available_stock, or null if SKU not in Zoho
 * @param {boolean}     preferZoho true → use Zoho first (DIRECT rows, Prep rows,
 *                                  AND manually-typed eBay rows). false → MI first
 *                                  (automated eBay-order rows = eBay truth).
 * @returns {number} HAND value (no committed subtraction — see file header)
 */
function resolveHandValue(miAvail, zoAvail, preferZoho) {
  if (preferZoho) {
    if (zoAvail != null) return zoAvail;
    if (miAvail != null) return miAvail;
    return 0;
  }
  if (miAvail != null) return miAvail;
  if (zoAvail != null) return zoAvail;
  return 0;
}


/**
 * Decide whether an eBay-table row was MANUALLY typed (→ prefer Zoho) vs.
 * inserted by the n8n eBay-orders workflow (→ keep MI / eBay truth).
 *
 * Signal: a real eBay order id in the SALES ORDER cell is digits-and-dashes
 * only (e.g. "02-14623-46718"). Manual rows never look like that — they carry
 * "Replacement for order # …", "Replacement #: …", or other free text, or are
 * blank. This is the SAME rule the S4 shipped-check uses to exclude replacement
 * rows, so the two stay consistent by construction.
 *
 * Returns true for manual/replacement/blank rows, false for clean order ids.
 * (DIRECT rows don't need this — they're always Zoho-first regardless.)
 */
function _isManualSalesOrder(so) {
  var s = String(so || "").trim();
  if (!s) return true;                                  // blank → treat as manual
  return !(/^[\d-]+$/.test(s) && s.indexOf('-') !== -1); // not a clean eBay order id → manual
}

// Back-fill idx() helper if the schema object didn't define one (parity with
// PRICE_AUDIT.idx / Schema.idx style). Kept defensive for buildZohoStockMap.
ZOHO_STOCK.idx = function(name) { return ZOHO_STOCK.cols[name] - 1; };
