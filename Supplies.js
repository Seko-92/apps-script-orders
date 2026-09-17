// =======================================================================================
// SUPPLIES — internal packing consumables (boxes, bags, tape, labels, void fill)
// =======================================================================================
//
// WHAT THIS IS
//   The ~60 materials used to PACK orders. They are bought outside Zoho (card /
//   supplier direct, no PO), they never sell, and until now nothing in the system
//   knew they existed — so the first sign of running out was an empty shelf with an
//   order waiting on it.
//
// THE SPLIT — ZOHO HOLDS THE NUMBER, THIS SHEET IS THE REGISTRY AND THE LENS
//   Same shape as Kit Registry (a local registry over Zoho-sourced data) and Price
//   Audit (a local lens over the Zoho mirror). The reason the NUMBER lives in Zoho
//   rather than here is that the picker wants three doors onto it — the tablet, this
//   sheet, and Zoho's own UI — and the 2-minute Zoho Stock sync already fetches every
//   non-inactive Zoho item with NO filter. So the moment a supply exists in Zoho it
//   flows to the mirror for free. A sheet-owned number could never offer the Zoho
//   door without building a sync back.
//
//   The write half already existed and is proven: pushSingleStockAdjustBySku →
//   the Zoho Inventory Adjustment Proxy, which computes the delta INSIDE Zoho's own
//   execution from live stock_on_hand. Nothing new was built for the write.
//
// ⚠⚠ THE SEPARATION CONTRACT — why ~60 extra Zoho items cannot touch real inventory
//   Only ONE place in the codebase iterates the whole Zoho map: PriceAudit's
//   zohoMap.forEach. Every other consumer (recomputeHand, LiveSync, Prep Queue, the
//   kit paths, Part Console, Low Stock, StockAdjust) does a keyed .get(skuLower), so
//   extra rows are never visited and are inert BY CONSTRUCTION, not by luck.
//
//   The MI-driven surfaces are closed by a different fact: a supply has no Master
//   Inventory row, so Out of Stock, Product Data Health and Price Audit's eBay side
//   cannot see it. Low Stock is demand-scoped and a thing that never sells never
//   enters it. Kit Registry's webhook fails its name regex and skips silently.
//
//   That leaves PriceAudit's one full iteration, and it is closed explicitly there by
//   a SUPPLIES.isSupplySku(sku) skip. ⚠ Do NOT rely on the incidental zero-price
//   escape (`if (zohoPrice <= 0) return;`) alone — nothing documents it as a contract,
//   and one person typing a price on a box would flood the hygiene worklist with 60
//   permanent rows.
//
// ⚠⚠ THE GUARD THAT MATTERS — pushSingleStockAdjustBySku will adjust ANY SKU present
//   in the mirror, real parts included. adjustSupplyCount is the only supplies write
//   path and it refuses twice before reaching that call: the SKU must carry the
//   SUP- prefix (a 6-digit part SKU can NEVER pass this — that is the hard boundary),
//   AND it must be declared on this sheet. This is the one place the two worlds could
//   genuinely touch.
//
// ⚠ ZOHO SETUP REQUIREMENTS (they are not optional)
//   · Create each supply as a TRACKED INVENTORY item. A "Sales/Purchase" item returns
//     no stock_on_hand and the proxy's Compute Delta node throws "refusing to adjust
//     blind" — the whole write path fails for it.
//   · Leave the SELLING PRICE EMPTY. Purchase-only is the honest shape anyway, and it
//     is a second line of defence behind the isSupplySku skip in Price Audit.
//   · Prefix every SKU SUP- and keep it consistent.
//
// THE SKU SCHEME — SUP-<nnn>, SEQUENTIAL, BLOCKED BY FAMILY
//   ⚠⚠ THE DECIDING CONSTRAINT IS THAT THIS GETS PRINTED. The SKU goes on a label
//   stuck to the container holding the stock, so it is short, fixed-width and
//   all-digits after the prefix — nothing a human can misread off a bin at arm's
//   length (no 0/O, no 1/I). A descriptive SKU (SUP-BOX-3007A-HQ) is unreadable at
//   label size and was rejected for that reason, not on taste.
//
//   THREE digits, deliberately: the picker's own box codes are FOUR (1003, 7014),
//   so SUP-118 can never be mistaken for one of those on a shelf where both appear.
//
//   The descriptive code lives in the ITEM TITLE, where it has room —
//   "Box (Hq) 3007A". That is what names and SKUs are each for, and it is also why
//   his two duplicated codes (1003 and 1006 exist as both Hq and Brown) are a
//   non-issue: sequential SKUs are unique by construction.
//
//   Sequential also survives the scope growing. The user intends to extend past
//   packing to other operational consumables, and a descriptive scheme has to be
//   redesigned the moment that happens.
//
//   The hundred-block is a NUMBERING HABIT, not a rule — nothing in this file enforces
//   or reads it. It exists because ZOHO HAS NO KIND COLUMN: there the SKU and the item
//   name are all you get, so the leading digit is what groups boxes away from bags in
//   a 3,600-item list. Our own sheet uses the KIND column for the same job.
//     1xx boxes · 2xx bags & mailers · 3xx tape & adhesives · 4xx labels & documents
//     5xx void fill, wrap, padding  · 7xx cleaning · 8xx safety/PPE · 9xx shop/office
//
//   ⚠ WHERE THIS SCHEME STOPS: durable EQUIPMENT (a scale, a label printer) is not
//   consumed, has no reorder point, and wants service dates rather than counts. That
//   is a different sheet, not a different prefix — "supplies" covering things we use
//   up is the right boundary, and it tells you when you have left it.
//
// PUBLIC API
//   setupSuppliesSheet()          — idempotent layout/migrator (sidebar "Re-style")
//   refreshSupplies(maps)         — hourly: pull ON HAND from the mirror, derive STATUS
//   openSupplies()                — activate the sheet
//   adjustSupplyCount(sku, n, o)  — THE guarded write (opening count / pack / recount)
//   getSuppliesLowCount()         — cheap snapshot read for badges
//   getLowSupplies()              — the low list, for the watchdog / Telegram / board
//   suppliesOnEdit(e)             — SKU entry → fill ITEM + ON HAND (from Main.js)
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
var SUPPLIES = {
  sheetName: "Supplies",

  cols: {
    SKU:          1,   // A — SUP-… (user-owned)
    ITEM:         2,   // B — Zoho item name (mirrored)
    KIND:         3,   // C — box / bag / tape / label (user-owned, groups 60 rows)
    LOCATION:     4,   // D — where it is stored (user-owned)
    ON_HAND:      5,   // E — from the Zoho mirror
    REORDER_AT:   6,   // F — user-owned; the level that fires the alert
    STATUS:       7,   // G — derived: OUT / LOW / OK
    LAST_COUNTED: 8,   // H — stamped by adjustSupplyCount
    COUNTED_BY:   9,   // I — picker at the time of the count
    NOTE:        10    // J — supplier, pack size, what it fits (user-owned)
  },

  idx: function(name) { return SUPPLIES.cols[name] - 1; },

  dataWidth: 10,

  titleRow:     1,   // ▌ SUPPLIES band (chip in-band at J1, stamp L1 hidden)
  headerRow:    2,
  dataStartRow: 3,

  // ⚠ THE HARD BOUNDARY. Real part SKUs in this business are 6 DIGITS, so a
  // letter prefix can never collide with one. This is what stops a supplies
  // surface from ever reaching real stock — see the guard note in the header.
  skuPrefix: "SUP-",

  isSupplySku: function(sku) {
    return String(sku || "").trim().toUpperCase().indexOf(SUPPLIES.skuPrefix) === 0;
  },

  // ⚠ Supplies legitimately move in PACKS — a bundle of 25 boxes, a bag of 500
  // poly bags, and an opening count that goes 0 → 500. STOCK_ADJUST.maxDelta (50)
  // is sized for correcting a sellable part's shelf count and would refuse all of
  // those. This is the supplies-specific ceiling: generous enough for a real pack,
  // tight enough that a fat-fingered extra zero still gets refused.
  maxDelta: 2000,

  // ⚠ ONE CONSTANT STRING, NEVER per-item or per-picker text. Zoho's adjustment
  // `reason` is a MANAGED DROPDOWN — every distinct string we send becomes a
  // permanent entry in the list humans pick from, growing fastest exactly when the
  // feature succeeds. The picker's name rides in the adjustment DESCRIPTION, which
  // is free text (the same ruling as the Floor Board's stock fix).
  reason: "Packing supplies count (warehouse floor)",

  maxDataRow: 500,

  headers: ["◈ SKU", "ITEM", "KIND", "LOCATION", "ON HAND", "REORDER AT",
            "STATUS", "LAST COUNTED", "COUNTED BY", "NOTE"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * Idempotent setup + layout migrator. Safe to re-run any time (sidebar
 * "Re-style Sheet"). Creates the sheet on first run.
 */
function setupSuppliesSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(SUPPLIES.sheetName);
  }

  // --- TITLE-BAND MIGRATION ---
  // A layout with the column headers on row 1 predates the title band. Shift
  // everything down one so row 1 becomes the ▌ SUPPLIES band (Prep Queue pattern).
  var a1 = String(sheet.getRange(1, 1).getValue()).trim();
  if (a1.charAt(0) === '◈') {
    sheet.insertRowsBefore(1, 1);
  }

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(SUPPLIES.cols.SKU,          130);
  sheet.setColumnWidth(SUPPLIES.cols.ITEM,         240);
  sheet.setColumnWidth(SUPPLIES.cols.KIND,          90);
  sheet.setColumnWidth(SUPPLIES.cols.LOCATION,     110);
  sheet.setColumnWidth(SUPPLIES.cols.ON_HAND,       90);
  sheet.setColumnWidth(SUPPLIES.cols.REORDER_AT,   100);
  sheet.setColumnWidth(SUPPLIES.cols.STATUS,        90);
  sheet.setColumnWidth(SUPPLIES.cols.LAST_COUNTED, 120);
  sheet.setColumnWidth(SUPPLIES.cols.COUNTED_BY,   120);
  sheet.setColumnWidth(SUPPLIES.cols.NOTE,         260);

  // --- DATA AREA: column-level formats so new rows inherit ---
  var dataRows = SUPPLIES.maxDataRow - SUPPLIES.dataStartRow + 1;
  _applySuppliesDataRowFormats(sheet, SUPPLIES.dataStartRow, dataRows);

  // --- BANDING (cream alternation) ---
  // Range starts at the HEADER row (2), NOT row 1 — the banding header slot paints
  // over manual fills and would black out the row-1 title band (Prep Queue rollout
  // lesson 2026-07-16). Row 1 stays OUTSIDE the banding, and the band/header manual
  // styling below runs AFTER this step.
  // Wrapped: if applyRowBanding ever throws (an overlapping banding on a grown
  // sheet) it must NOT abort setup before the band-LABEL styling below.
  try {
    sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
    var bandRange = sheet.getRange(SUPPLIES.headerRow, 1,
                                   SUPPLIES.maxDataRow - SUPPLIES.headerRow + 1,
                                   SUPPLIES.dataWidth);
    var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    band.setHeaderRowColor('#1d1d1b')
        .setFirstRowColor('#ffffff')
        .setSecondRowColor('#fff8e7');
  } catch (bErr) {
    try { Logger.log("setupSuppliesSheet: banding: " + bErr); } catch (_) {}
  }

  // --- TITLE BAND + HEADERS (after banding, so manual styling wins) ---
  _styleSuppliesBand(sheet, SUPPLIES.titleRow, "SUPPLIES", "PACKING MATERIALS · INTERNAL");
  _styleSuppliesHeaderRow(sheet, SUPPLIES.headerRow, SUPPLIES.headers);

  // --- CONDITIONAL FORMATTING (wipe ours in cols A..J, rebuild) ---
  // Rule ORDER is load-bearing — Sheets applies the FIRST matching rule per cell.
  // OUT before LOW before OK. Chip rules (col J row 1) are excluded by the
  // headerRow floor below.
  var keptRules = sheet.getConditionalFormatRules().filter(function(r) {
    return !r.getRanges().some(function(rg) {
      return rg.getRow() >= SUPPLIES.dataStartRow &&
             rg.getColumn() <= SUPPLIES.dataWidth;
    });
  });

  var dsr = SUPPLIES.dataStartRow;
  var statusRange = sheet.getRange(dsr, SUPPLIES.cols.STATUS, dataRows, 1);

  var outRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G' + dsr + '="OUT"')
    .setBackground('#ff6b6b').setFontColor('#ffffff').setBold(true)
    .setRanges([statusRange]).build();
  var lowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G' + dsr + '="LOW"')
    .setBackground('#ffd400').setFontColor('#1d1d1b').setBold(true)
    .setRanges([statusRange]).build();
  var okRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G' + dsr + '="OK"')
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([statusRange]).build();

  // ⚠ A SKU that does not carry the SUP- prefix should never be on this sheet —
  // it is how a real part SKU would get into the registry and become adjustable.
  // adjustSupplyCount refuses it regardless, but the sheet says so loudly too.
  var skuRange = sheet.getRange(dsr, SUPPLIES.cols.SKU, dataRows, 1);
  var badSkuRule = SpreadsheetApp.newConditionalFormatRule()
    // ⚠ BOTH the length and the literal are DERIVED from SUPPLIES.skuPrefix. A CF
    // formula cannot call isSupplySku, so this rule is necessarily a second copy of
    // the same decision — and two copies of one rule is how "A-9 sorted after A-50"
    // survived in three files here. Deriving them means changing the prefix moves
    // the sheet and the guard together, in one edit.
    .whenFormulaSatisfied('=AND($A' + dsr + '<>"",LEFT(UPPER($A' + dsr + '),' +
                          SUPPLIES.skuPrefix.length + ')<>"' + SUPPLIES.skuPrefix + '")')
    .setBackground('#ff6b6b').setFontColor('#ffffff').setBold(true)
    .setRanges([skuRange]).build();

  // A missing reorder point is not an error, but it IS why a row can never alert —
  // surface it quietly so 60 rows don't hide one that was never configured.
  var reorderRange = sheet.getRange(dsr, SUPPLIES.cols.REORDER_AT, dataRows, 1);
  var noReorderRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + dsr + '<>"",$F' + dsr + '="")')
    .setBackground('#fff3b0').setFontColor('#7a5c00')
    .setRanges([reorderRange]).build();

  sheet.setConditionalFormatRules(
    [outRule, lowRule, okRule, badSkuRule, noReorderRule].concat(keptRules));

  // --- FREEZE TITLE BAND + HEADER ---
  sheet.setFrozenRows(2);

  // --- FRESHNESS PULSE CHIP (J1 in-band / L1 hidden stamp) ---
  try { _installPulseChip(sheet, SHEET_PULSE.supplies); }
  catch (e) { try { Logger.log("setupSuppliesSheet: pulse chip error: " + e); } catch (_) {} }

  return "✅ Supplies sheet ready (" + SUPPLIES.skuPrefix + " items, Zoho-backed).";
}


/**
 * Pull ON HAND from the Zoho mirror for every declared supply, derive STATUS,
 * and auto-discover any SUP- item that exists in Zoho but is not yet on the sheet.
 *
 * ⚠ THIS SHEET IS USER-OWNED, UNLIKE Out of Stock. The refresh never DELETES a row
 *   and never touches SKU / KIND / LOCATION / REORDER AT / NOTE. A supply that
 *   vanishes from the mirror (transient sync gap, or someone marked it inactive in
 *   Zoho) keeps its row with a blank ON HAND — "I could not read it" must never be
 *   rendered as "there is none", which is the reassuring-label-on-a-dangerous-state
 *   bug this system has ruled against repeatedly.
 *
 * @param {Object} [maps] pre-built maps from runHourlyHousekeeping (unused here but
 *                        accepted so this matches every other housekeeping job).
 */
function refreshSupplies(maps) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);
  if (!sheet) {
    setupSuppliesSheet();
    sheet = ss.getSheetByName(SUPPLIES.sheetName);
  }

  // Self-heal a legacy header-on-row-1 layout.
  if (String(sheet.getRange(1, 1).getValue()).trim().charAt(0) === '◈') {
    setupSuppliesSheet();
  }

  var zohoMap;
  try {
    zohoMap = buildZohoStockMap();
  } catch (ze) {
    try { console.log("refreshSupplies: zoho map unavailable: " + ze); } catch (_) {}
    return "⚠️ Supplies: Zoho Stock unreadable — left untouched.";
  }
  if (!zohoMap || zohoMap.size === 0) {
    // Same guard the Photo Queue learned the hard way: an unreadable source must
    // bail, never rewrite the sheet with an empty answer.
    return "⚠️ Supplies: Zoho Stock sheet is empty — left untouched.";
  }

  var rows = _supReadRows(sheet);
  var seen = {};
  var low = 0, out = 0, unknown = 0;

  // ---- Refresh the declared rows ----
  var writes = [];   // [rowNumber, [item, onHand], status]
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r.sku) continue;
    seen[r.sku.toLowerCase()] = true;

    var z = zohoMap.get(r.sku.toLowerCase());
    var onHand = (z && z.onHand !== "" && z.onHand !== null && !isNaN(parseFloat(z.onHand)))
      ? parseFloat(z.onHand) : "";
    var itemName = (z && z.itemName) ? z.itemName : r.item;

    var status = _supStatus(onHand, r.reorderAt);
    if (status === "OUT") out++;
    else if (status === "LOW") low++;
    else if (status === "") unknown++;

    writes.push({ row: r.row, item: itemName, onHand: onHand, status: status });
  }

  for (var w = 0; w < writes.length; w++) {
    sheet.getRange(writes[w].row, SUPPLIES.cols.ITEM).setValue(writes[w].item);
    sheet.getRange(writes[w].row, SUPPLIES.cols.ON_HAND).setValue(writes[w].onHand);
    sheet.getRange(writes[w].row, SUPPLIES.cols.STATUS).setValue(writes[w].status);
  }

  // ---- Auto-discover SUP- items that exist in Zoho but not on the sheet ----
  // This is what makes the opening setup painless: create the ~60 items in Zoho and
  // they populate themselves here, leaving only KIND / LOCATION / REORDER AT to fill.
  var discovered = [];
  zohoMap.forEach(function(rec, skuLower) {
    if (seen[skuLower]) return;
    if (!SUPPLIES.isSupplySku(rec.skuOriginal)) return;
    discovered.push(rec);
  });

  if (discovered.length > 0) {
    var appendAt = Math.max(sheet.getLastRow() + 1, SUPPLIES.dataStartRow);
    var newRows = discovered.map(function(rec) {
      var onHand = (rec.onHand !== "" && rec.onHand !== null && !isNaN(parseFloat(rec.onHand)))
        ? parseFloat(rec.onHand) : "";
      var blank = new Array(SUPPLIES.dataWidth).fill("");
      blank[SUPPLIES.idx("SKU")]     = rec.skuOriginal;
      blank[SUPPLIES.idx("ITEM")]    = rec.itemName || "";
      blank[SUPPLIES.idx("ON_HAND")] = onHand;
      blank[SUPPLIES.idx("STATUS")]  = _supStatus(onHand, "");
      return blank;
    });
    sheet.getRange(appendAt, 1, newRows.length, SUPPLIES.dataWidth).setValues(newRows);
    // Rows appended past the previously formatted band (or inheriting the header's
    // dark format on an empty sheet) need their formats re-asserted — the same
    // degenerate-insert case _applyOosDataRowFormats takes (startRow, numRows) for.
    _applySuppliesDataRowFormats(sheet, appendAt, newRows.length);
  }

  // ⚠ Stamp ONLY on a completed pass, so the chip's staleness tiers double as the
  // failure alarm — a run that threw never reaches this line.
  try { stampSheetPulse(sheet, SHEET_PULSE.supplies.stamp); } catch (_) {}

  var msg = "📦 Supplies: " + writes.length + " tracked";
  if (out > 0)         msg += " · " + out + " OUT";
  if (low > 0)         msg += " · " + low + " LOW";
  if (unknown > 0)     msg += " · " + unknown + " no reading";
  if (discovered.length > 0) msg += " · " + discovered.length + " newly discovered";
  return msg;
}


/** Activate the Supplies sheet in the user's own window. */
function openSupplies() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);
  if (!sheet) {
    setupSuppliesSheet();
    sheet = ss.getSheetByName(SUPPLIES.sheetName);
  }
  if (!sheet) return "⚠️ Supplies sheet could not be created.";
  ss.setActiveSheet(sheet);
  return "📦 Supplies opened.";
}


/**
 * THE ONLY SUPPLIES WRITE PATH. Sets a supply's Zoho on-hand to an explicit count.
 *
 * ⚠⚠ THE GUARD IS THE POINT. pushSingleStockAdjustBySku will happily adjust any SKU
 *   present in the Zoho mirror — including a real sellable part. This refuses twice
 *   before it gets there:
 *     1. the SUP- prefix (a 6-digit part SKU can never pass — the hard boundary), and
 *     2. membership of the Supplies registry (scoped to what was actually declared).
 *   Neither check alone is sufficient: the prefix stops a part reaching Zoho's write,
 *   the registry stops a typo'd SUP- SKU adjusting an item nobody is tracking.
 *
 * @param {string} sku    the SUP- SKU
 * @param {number} target what the count SHOULD be
 * @param {Object} [opts] { picker }
 * @returns {Object} { ok, message, ... } — mirrors pushSingleStockAdjustBySku
 */
function adjustSupplyCount(sku, target, opts) {
  opts = opts || {};
  sku = String(sku || '').trim();
  if (!sku) return { ok: false, error: 'No SKU given.' };

  var n = parseFloat(target);
  if (isNaN(n) || n < 0) {
    return { ok: false, error: 'Count must be zero or more.' };
  }

  // --- GUARD 1: the prefix. A real part SKU cannot pass this. ---
  if (!SUPPLIES.isSupplySku(sku)) {
    return { ok: false,
             error: sku + ' is not a packing supply — supplies start with ' +
                    SUPPLIES.skuPrefix + '. Refusing, so this can never reach real stock.' };
  }

  // --- GUARD 2: it must be declared on the Supplies sheet. ---
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);
  if (!sheet) return { ok: false, error: 'Supplies sheet does not exist yet.' };

  var rows = _supReadRows(sheet);
  var hit = null;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].sku && rows[i].sku.toLowerCase() === sku.toLowerCase()) { hit = rows[i]; break; }
  }
  if (!hit) {
    return { ok: false, error: sku + ' is not on the Supplies sheet — add it there first.' };
  }

  var picker = String(opts.picker || '');
  if (!picker) { try { picker = getCurrentPicker() || ''; } catch (_) { picker = ''; } }

  var res = pushSingleStockAdjustBySku(sku, n, {
    maxDelta: SUPPLIES.maxDelta,
    reason:   SUPPLIES.reason,
    picker:   picker
  });

  if (!res || !res.ok) {
    return { ok: false, error: (res && res.message) ? res.message : 'Adjustment failed.', detail: res };
  }

  // Write the confirmed number straight back rather than waiting up to 2 minutes for
  // the mirror — the picker is standing there and should see what they just set.
  // The next refresh re-reads it from Zoho and would correct any disagreement.
  try {
    var nowStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");
    sheet.getRange(hit.row, SUPPLIES.cols.ON_HAND).setValue(n);
    sheet.getRange(hit.row, SUPPLIES.cols.STATUS).setValue(_supStatus(n, hit.reorderAt));
    sheet.getRange(hit.row, SUPPLIES.cols.LAST_COUNTED).setValue(nowStr);
    sheet.getRange(hit.row, SUPPLIES.cols.COUNTED_BY).setValue(picker || 'in the sheet');
  } catch (wErr) {
    try { console.log("adjustSupplyCount: sheet write-back failed: " + wErr); } catch (_) {}
  }

  try {
    logActivity('NOTE', '', sku, n, 'supplies',
                'Supply count ' + (res.noop ? 'confirmed at ' + n
                                            : res.before + ' → ' + n +
                                              ' (' + (res.delta > 0 ? '+' : '') + res.delta + ')') +
                (res.adjustmentId ? ' · adj ' + res.adjustmentId : ''));
  } catch (_) {}

  return { ok: true, sku: sku, before: res.before, target: n, delta: res.delta,
           noop: !!res.noop, message: res.message };
}


/**
 * Cheap snapshot count of supplies at or below their reorder point (OUT + LOW).
 * Reads the sheet's own derived STATUS — it does NOT re-derive or hit Zoho, so it
 * is safe on a poll. Returns 0 when the sheet is absent.
 */
function getSuppliesLowCount() {
  try {
    var list = getLowSupplies();
    return list.length;
  } catch (e) {
    return 0;
  }
}


/**
 * The low list — every declared supply whose STATUS is OUT or LOW, worst first.
 * Feeds the watchdog bucket, the Telegram route and the board strip.
 *
 * ⚠ Rows with no reorder point are NOT included. They cannot be judged, and a row
 *   nobody configured must not masquerade as a row that is fine — the STATUS cell
 *   stays blank and the sheet's amber tint on REORDER AT is where that shows.
 */
function getLowSupplies() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);
  if (!sheet) return [];

  var rows = _supReadRows(sheet);
  var hits = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r.sku) continue;
    if (r.status !== 'OUT' && r.status !== 'LOW') continue;
    hits.push({
      sku:       r.sku,
      item:      r.item,
      kind:      r.kind,
      location:  r.location,
      onHand:    r.onHand,
      reorderAt: r.reorderAt,
      status:    r.status,
      shortBy:   (typeof r.onHand === 'number' && typeof r.reorderAt === 'number')
                   ? Math.max(0, r.reorderAt - r.onHand) : ''
    });
  }

  // OUT before LOW, then by how far under the reorder point (biggest gap first).
  hits.sort(function(a, b) {
    if (a.status !== b.status) return a.status === 'OUT' ? -1 : 1;
    var ag = (typeof a.shortBy === 'number') ? a.shortBy : -1;
    var bg = (typeof b.shortBy === 'number') ? b.shortBy : -1;
    if (ag !== bg) return bg - ag;
    return String(a.sku).localeCompare(String(b.sku));
  });
  return hits;
}


// =======================================================================================
// onEdit DISPATCHER — called from Main.js's onEditInstallable(e)
// =======================================================================================

/**
 * SKU edit on Supplies → fill ITEM name + ON HAND from the Zoho mirror and derive
 * STATUS. Clearing the SKU clears only the MACHINE-owned cells; KIND, LOCATION,
 * REORDER AT and NOTE are hand-typed and are never wiped by an edit to column A.
 *
 * Lives on the INSTALLABLE trigger because the mirror read goes through openById,
 * which simple triggers cannot call reliably. Defensive try/catch — never blocks
 * the other edit handlers.
 */
function suppliesOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SUPPLIES.sheetName) return;
    if (e.range.getColumn() !== SUPPLIES.cols.SKU) return;
    if (e.range.getRow() < SUPPLIES.dataStartRow) return;

    var edits = e.range.getValues();
    var startRow = e.range.getRow();
    var zohoMap = null;

    for (var i = 0; i < edits.length; i++) {
      var row = startRow + i;
      var sku = String(edits[i][0] || "").trim();

      if (!sku) {
        // Machine-owned cells only. B, E, G, H, I — never C/D/F/J.
        sheet.getRange(row, SUPPLIES.cols.ITEM).clearContent();
        sheet.getRange(row, SUPPLIES.cols.ON_HAND).clearContent();
        sheet.getRange(row, SUPPLIES.cols.STATUS).clearContent();
        sheet.getRange(row, SUPPLIES.cols.LAST_COUNTED).clearContent();
        sheet.getRange(row, SUPPLIES.cols.COUNTED_BY).clearContent();
        continue;
      }

      if (!SUPPLIES.isSupplySku(sku)) {
        // The CF rule already paints this red; say why in the one cell that has room.
        sheet.getRange(row, SUPPLIES.cols.ITEM)
             .setValue("⚠ not a supply SKU — must start with " + SUPPLIES.skuPrefix);
        sheet.getRange(row, SUPPLIES.cols.ON_HAND).clearContent();
        sheet.getRange(row, SUPPLIES.cols.STATUS).clearContent();
        continue;
      }

      if (zohoMap === null) {
        try { zohoMap = buildZohoStockMap(); } catch (ze) { zohoMap = new Map(); }
      }
      var z = zohoMap.get(sku.toLowerCase());
      if (!z) {
        sheet.getRange(row, SUPPLIES.cols.ITEM)
             .setValue("⚠ not in Zoho yet — create it as a tracked Inventory item");
        sheet.getRange(row, SUPPLIES.cols.ON_HAND).clearContent();
        sheet.getRange(row, SUPPLIES.cols.STATUS).clearContent();
        continue;
      }

      var onHand = (z.onHand !== "" && z.onHand !== null && !isNaN(parseFloat(z.onHand)))
        ? parseFloat(z.onHand) : "";
      var reorderAt = sheet.getRange(row, SUPPLIES.cols.REORDER_AT).getValue();

      sheet.getRange(row, SUPPLIES.cols.ITEM).setValue(z.itemName || "");
      sheet.getRange(row, SUPPLIES.cols.ON_HAND).setValue(onHand);
      sheet.getRange(row, SUPPLIES.cols.STATUS).setValue(_supStatus(onHand, reorderAt));
    }
  } catch (err) {
    try { Logger.log("suppliesOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: pure decision helper (Node-testable)
// =======================================================================================

/**
 * Derive a supply's STATUS. PURE — no Sheets calls.
 *
 * ⚠ BLANK IS A REAL ANSWER AND IT IS NOT "OK". Two distinct cases return "":
 *   no reading from Zoho, and no reorder point configured. Neither can be judged,
 *   and rendering either as OK would be a reassuring label on a state nobody has
 *   checked — the exact class of bug this project has been bitten by repeatedly.
 *
 * @param {number|string} onHand    from the Zoho mirror ("" when unreadable)
 * @param {number|string} reorderAt the hand-set threshold ("" when unset)
 * @returns {string} "OUT" | "LOW" | "OK" | ""
 */
function _supStatus(onHand, reorderAt) {
  var h = parseFloat(onHand);
  if (onHand === "" || onHand === null || onHand === undefined || isNaN(h)) return "";
  if (h <= 0) return "OUT";

  var r = parseFloat(reorderAt);
  if (reorderAt === "" || reorderAt === null || reorderAt === undefined || isNaN(r)) return "";
  return (h <= r) ? "LOW" : "OK";
}


// =======================================================================================
// PRIVATE: sheet helpers
// =======================================================================================

/**
 * Read every declared row as objects. One getValues call; blank-SKU rows are kept
 * (with row numbers) so callers can distinguish "no rows" from "unreadable".
 */
function _supReadRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < SUPPLIES.dataStartRow) return [];

  var n = lastRow - SUPPLIES.dataStartRow + 1;
  var vals = sheet.getRange(SUPPLIES.dataStartRow, 1, n, SUPPLIES.dataWidth).getValues();

  var out = [];
  for (var i = 0; i < n; i++) {
    var v = vals[i];
    var sku = String(v[SUPPLIES.idx("SKU")] || "").trim();
    if (!sku) continue;
    out.push({
      row:       SUPPLIES.dataStartRow + i,
      sku:       sku,
      item:      String(v[SUPPLIES.idx("ITEM")] || ""),
      kind:      String(v[SUPPLIES.idx("KIND")] || ""),
      location:  String(v[SUPPLIES.idx("LOCATION")] || ""),
      onHand:    v[SUPPLIES.idx("ON_HAND")],
      reorderAt: v[SUPPLIES.idx("REORDER_AT")],
      status:    String(v[SUPPLIES.idx("STATUS")] || "").trim().toUpperCase(),
      note:      String(v[SUPPLIES.idx("NOTE")] || "")
    });
  }
  return out;
}


/**
 * Style the brand-yellow ▌ title band. markerText lands as the cell VALUE (the ▌ is
 * number-format dressing) — the same Gotcha #1 discipline the other sheets use, so
 * anything matching on this cell sees a clean string.
 */
function _styleSuppliesBand(sheet, row, markerText, rightLabel) {
  sheet.getRange(row, 1, 1, SUPPLIES.dataWidth)
    .setBackground('#ffd400')   // brand action yellow
    .setFontColor('#1d1d1b')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontLine('none')
    .setFontStyle('normal')
    .setVerticalAlignment('middle')
    .setBorder(true, null, true, null, null, null,
               '#1d1d1b', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(row, 1)
    .setValue(markerText)
    .setNumberFormat('"▌  "@')
    .setHorizontalAlignment('left');
  // Right label sits mid-band, right-aligned so it overflows LEFT over the empty
  // band cells. J1 stays clear — it is the in-band chip's home, and the chip
  // installer restyles that one cell after this runs.
  sheet.getRange(row, SUPPLIES.cols.REORDER_AT)
    .setValue(rightLabel)
    .setHorizontalAlignment('right')
    .setFontSize(9);
  sheet.setRowHeight(row, 36);
}


/** Dark brand header band, yellow Oswald text, thick yellow underline. */
function _styleSuppliesHeaderRow(sheet, row, headers) {
  sheet.getRange(row, 1, 1, SUPPLIES.dataWidth)
    .setValues([headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontLine('none')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(row, 1, 1, SUPPLIES.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
}


/**
 * Column-level data formats for a run of rows. Takes (startRow, numRows) so it
 * serves BOTH the setup whole-band pass and refresh's append case, where rows added
 * past the formatted band (or onto an empty sheet, inheriting the header's dark
 * format) need their formats re-asserted.
 *
 * ⚠ GOTCHA #16 — every code-written NUMBER column gets an explicit number format.
 *   clearContent() preserves formats, so a column whose meaning ever changes can keep
 *   a stale DATE format and render correct integers as 1900-era dates — and then
 *   getValues() returns Date objects, so parseFloat is NaN and a reader silently
 *   reports zero. Same reason LAST COUNTED is pinned to plain text: it is written as
 *   an "M/d/yy h:mm a" STRING, and without '@' Sheets coerces it into a real Date.
 */
function _applySuppliesDataRowFormats(sheet, startRow, numRows) {
  sheet.getRange(startRow, SUPPLIES.cols.SKU, numRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setNumberFormat('@');
  sheet.getRange(startRow, SUPPLIES.cols.ITEM, numRows, 1)
    .setFontFamily('Roboto').setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.getRange(startRow, SUPPLIES.cols.KIND, numRows, 1)
    .setFontFamily('Oswald').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, SUPPLIES.cols.LOCATION, numRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, SUPPLIES.cols.ON_HAND, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setNumberFormat('0');
  sheet.getRange(startRow, SUPPLIES.cols.REORDER_AT, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setFontColor('#434343').setHorizontalAlignment('center').setNumberFormat('0');
  sheet.getRange(startRow, SUPPLIES.cols.STATUS, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, SUPPLIES.cols.LAST_COUNTED, numRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center')
    .setNumberFormat('@');
  sheet.getRange(startRow, SUPPLIES.cols.COUNTED_BY, numRows, 1)
    .setFontFamily('Roboto').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(startRow, SUPPLIES.cols.NOTE, numRows, 1)
    .setFontFamily('Roboto').setFontSize(9)
    .setHorizontalAlignment('left');

  sheet.getRange(startRow, 1, numRows, SUPPLIES.dataWidth)
    .setVerticalAlignment('middle');
}

// =======================================================================================
// ONE-OFF SEED — the picker's 2026-09-17 shelf survey
// =======================================================================================
//
// ⚠ THE SKU IS A LABEL ID, NOT A DESCRIPTION. These get PRINTED and stuck on the
//   containers that hold the stock, so they are short, fixed-width and all-digits
//   after the prefix: no 0/O or 1/I to misread off a bin at arm's length. THREE
//   digits on purpose — the picker's own box codes are FOUR (1003, 7014), so
//   SUP-118 can never be mistaken for one of those.
//
//   His code lives in the TITLE, where it has room: "Box (Hq) 3007A". That is what
//   names and SKUs are each for, and it is why the Hq-vs-Brown duplicates at 1003
//   and 1006 are a non-issue here — sequential SKUs are unique by construction.
//
//   Blocked so a shelf of labelled containers groups by eye, and so Zoho's item
//   list (which has no KIND column) sorts sensibly:
//     101-145 boxes · 151-154 ring boxes · 201-209 shipping bags · 211-213 poly bags
//   Gaps are deliberate — growth inside a family must not force a renumber.
//
// ⚠ ON HAND IS NOT SEEDED. It comes from Zoho and only from Zoho. The picker's
//   count rides in the NOTE instead, as the number to type when creating the item —
//   seeding it would put a figure in the column that nothing verified, and the next
//   refresh would silently overwrite it anyway.
//
// Run ONCE from the editor. Refuses if the sheet already has rows.

/** The picker's 2026-09-17 shelf survey. SKU = label id; his own code is in the TITLE. */
var SUPPLIES_SEED = [
  ["SUP-101", "Box (Hq) 1000", "box", "opening count 12 SET"],
  ["SUP-102", "Box (Hq) 1002A", "box", "opening count 04 SET"],
  ["SUP-103", "Box (Hq) 1003", "box", "opening count 12 SET"],
  ["SUP-104", "Box (Brown) 1003", "box", "opening count 12 SET"],
  ["SUP-105", "Box (Hq) 1004", "box", "opening count 10 SET"],
  ["SUP-106", "Box (Hq) 1005", "box", "opening count 04 SET"],
  ["SUP-107", "Box (Hq) 1006", "box", "opening count 05 SET"],
  ["SUP-108", "Box (Brown) 1006", "box", "opening count 05 SET"],
  ["SUP-109", "Box (Hq) 1031", "box", "opening count 08 SET"],
  ["SUP-110", "Box (Hq) 1051", "box", "opening count 06 SET"],
  ["SUP-111", "Box (Hq) 1052", "box", "opening count 03 SET"],
  ["SUP-112", "Box (Hq) 1053", "box", "opening count 08 SET"],
  ["SUP-113", "Box (Hq) 1053A", "box", "opening count 14 SET"],
  ["SUP-114", "Box (Hq) 2021", "box", "opening count 16 SET"],
  ["SUP-115", "Box (Hq) 3002", "box", "opening count 08 SET"],
  ["SUP-116", "Box (Hq) 3003", "box", "opening count 06 SET"],
  ["SUP-117", "Box (Hq) 3005", "box", "opening count 14 SET"],
  ["SUP-118", "Box (Hq) 3006", "box", "opening count 08 SET"],
  ["SUP-119", "Box (Hq) 3007A", "box", "opening count 25 SET"],
  ["SUP-120", "Box (Hq) 3008", "box", "opening count 40 SET"],
  ["SUP-121", "Box (Hq) 3009", "box", "opening count 04 SET"],
  ["SUP-122", "Box (Hq) 3011", "box", "opening count 50 SET"],
  ["SUP-123", "Box (Hq) 3013", "box", "opening count 04 SET"],
  ["SUP-124", "Box (Hq) 3015", "box", "opening count 12 SET"],
  ["SUP-125", "Box (Hq) 4000", "box", "opening count 08 SET"],
  ["SUP-126", "Box (Hq) 4001", "box", "opening count 40 SET"],
  ["SUP-127", "Box (Hq) 4003", "box", "opening count 06 SET"],
  ["SUP-128", "Box (Hq) 4009", "box", "opening count 50 SET"],
  ["SUP-129", "Box (Hq) 4010A", "box", "opening count 20 SET"],
  ["SUP-130", "Box (Hq) 5002", "box", "opening count 16 SET"],
  ["SUP-131", "Box (Hq) 5004", "box", "opening count 03 SET"],
  ["SUP-132", "Box (Hq) 5019", "box", "opening count 10 SET"],
  ["SUP-133", "Box (Hq) 5020", "box", "opening count 08 SET"],
  ["SUP-134", "Box (Hq) 5022", "box", "opening count 60 SET"],
  ["SUP-135", "Box (Hq) 5025", "box", "opening count 18 SET"],
  ["SUP-136", "Box (Hq) 7001", "box", "opening count 01 SET"],
  ["SUP-137", "Box (Hq) 7002", "box", "opening count 03 SET"],
  ["SUP-138", "Box (Hq) 7003", "box", "opening count 14 SET"],
  ["SUP-139", "Box (Hq) 7003A", "box", "opening count 45 SET"],
  ["SUP-140", "Box (Hq) 7004", "box", "opening count 04 SET"],
  ["SUP-141", "Box (Hq) 7005", "box", "opening count 15 SET"],
  ["SUP-142", "Box (Hq) 7007", "box", "opening count 03 SET"],
  ["SUP-143", "Box (Hq) 7009", "box", "opening count 15 SET"],
  ["SUP-144", "Box (Hq) 7011", "box", "opening count 06 SET"],
  ["SUP-145", "Box (Hq) 7014", "box", "opening count 20 SET"],
  ["SUP-151", "Box (Hq) 2024 (For Ring)", "ring box", "For Ring · opening count 06 BOX"],
  ["SUP-152", "Box (Brown) 2025 (For Ring)", "ring box", "For Ring · opening count 01 BOX"],
  ["SUP-153", "Box (Hq) 2026 (For Ring)", "ring box", "For Ring · opening count 01 BOX"],
  ["SUP-154", "Box (Hq) 2027 (For Ring)", "ring box", "For Ring · opening count 01 BOX"],
  ["SUP-201", "Shipping Bag 1 (16*23*5)", "ship bag", "opening count 19 SET"],
  ["SUP-202", "Shipping Bag 2 (20*28*5)", "ship bag", "opening count 12 SET"],
  ["SUP-203", "Shipping Bag 2.5 (10*13)", "ship bag", "opening count 60 SET"],
  ["SUP-204", "Shipping Bag 3 (22*5*50)", "ship bag", "opening count 40 SET"],
  ["SUP-205", "Shipping Bag 4 (28*40)", "ship bag", "opening count 50 SET"],
  ["SUP-206", "Shipping Bag 5 (24*50)", "ship bag", "opening count 60 SET"],
  ["SUP-207", "Shipping Bag 5.5 (14.5*19)", "ship bag", "opening count 60 SET"],
  ["SUP-208", "Shipping Bag 6 (40*50)", "ship bag", "opening count 40 SET"],
  ["SUP-209", "Shipping Bag 8 (24*24)", "ship bag", "opening count 300 SET"],
  ["SUP-211", "Bag 1 YELLOW", "poly bag", "opening count 10 SET"],
  ["SUP-212", "Bag 2 BLUE", "poly bag", "opening count 00 SET"],
  ["SUP-213", "Bag 3 BLUE", "poly bag", "opening count 10 SET"],
];

/**
 * Write the 61 surveyed supplies into the Supplies sheet as a worklist.
 * REFUSES on a non-empty sheet — this is a seed, never a sync.
 */
function seedSuppliesRegistry() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SUPPLIES.sheetName);
  if (!sheet) { setupSuppliesSheet(); sheet = ss.getSheetByName(SUPPLIES.sheetName); }

  var existing = _supReadRows(sheet);
  if (existing.length > 0) {
    return "⚠ Supplies already has " + existing.length + " row(s) — refusing to seed over them. " +
           "Clear the sheet first if you really mean to re-seed.";
  }

  var rows = SUPPLIES_SEED.map(function (s) {
    var r = new Array(SUPPLIES.dataWidth).fill("");
    r[SUPPLIES.idx("SKU")]  = s[0];
    r[SUPPLIES.idx("ITEM")] = s[1];   // proposed Zoho name; refresh replaces it with Zoho's own
    r[SUPPLIES.idx("KIND")] = s[2];
    r[SUPPLIES.idx("NOTE")] = s[3];
    return r;
  });

  sheet.getRange(SUPPLIES.dataStartRow, 1, rows.length, SUPPLIES.dataWidth).setValues(rows);
  _applySuppliesDataRowFormats(sheet, SUPPLIES.dataStartRow, rows.length);

  return "✅ Seeded " + rows.length + " supplies. Next: create each in Zoho as a TRACKED " +
         "INVENTORY item with NO selling price, using the SKU and the count in its NOTE. " +
         "ON HAND fills itself on the next hourly refresh.";
}
