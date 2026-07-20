// =======================================================================================
// OUT_OF_STOCK.gs — "Out of Stock" sheet + alert  ·  TWO-TABLE LAYOUT (2026-07-18)
// =======================================================================================
//
// PURPOSE
//   The restock decision surface. One sheet, two stacked tables (the Prep
//   Queue pattern):
//
//     Row 1        ▌ OUT OF STOCK yellow title band (sync chip in-band at H1)
//     Row 2        main-table headers (frozen rows = 2)
//     Row 3+       MAIN TABLE — plain SKUs to REORDER from the supplier
//     ▌ KITS       divider (col A value EXACTLY "KITS" — same exact-marker
//                  contract as All Orders' DIRECT / Prep Queue's INCOMING)
//     next row     kit-table headers
//     below        KIT TABLE — out-of-stock KITS with a BUILDABLE count, so
//                  the decision "open this kit from components?" is one glance
//
//   The two tables answer the two different restock actions: the main table
//   is "order more from the supplier"; the kit table is "we may already OWN
//   this — as components." A kit SKU never appears in both.
//
// DEFINITION (both tables)
//   Out of stock = `quantity - quantitySold <= 0` in Master Inventory
//   AND `listingStatus == "Active"` (added 2026-07-18: deleted/ended eBay
//   listings get flipped to "Completed"/"Ended" by the n8n sync but their MI
//   rows are never removed — without the filter they lingered here forever as
//   phantom restock candidates; 15 of 112 rows at the time of the fix).
//
// KIT TABLE MATH
//   For each OOS kit (SKU present in the Kit Registry):
//     BUILDABLE  = min over components of floor(available ÷ qty-per-kit)
//                  — the number of complete kits the shelf can assemble.
//     LIMITED BY = the bottleneck component ("167517 · has 12 / needs 6").
//     COMPONENTS = health ("5 ok", or a loud "⚠ …" when the registry has
//                  unparsed PD lines / unknown components — an untrustable
//                  number is never shown, same honesty rule as the Kit
//                  Expansion modal).
//   Component availability is ZOHO-FIRST (Zoho Stock sheet, ≤2 min stale via
//   the n8n push) with MI fallback — the same routing every other kit surface
//   uses. Zero new API calls: both sources are sheets already in the file.
//   READY kits are included too (their K-* LOCATION identifies them) — for a
//   READY kit the number reads "how many more boxes could be assembled".
//
// DAYS OUT IS WRITTEN, NOT A FORMULA (changed 2026-07-18)
//   The old col-H ARRAYFORMULA would be #REF!-blocked by the kit table's
//   second header row (an array spill cannot expand over a value). The hourly
//   refresh now writes DAYS OUT as plain numbers — day-granular data refreshed
//   hourly loses nothing, and the whole DATEVALUE/date-coercion fragility
//   class (see _normalizeOosFirstSeen) stops mattering for display.
//
// SMART-MERGE REFRESH
//   Still a MERGE, not a wipe: FIRST SEEN survives across refreshes — and
//   across SEGMENTS (a kit that sat in the main table before the two-table
//   layout keeps its date when it moves to the kit table). On each refresh:
//     - still OOS + Active → refresh values, keep FIRST SEEN
//     - restocked          → drop
//     - listing not Active → drop (phantom)
//     - kit SKU            → kit table (never the main table)
//     - not in MI at all   → preserved as-is in the main table (manual entry)
//   The divider slides via insertRowsAfter/deleteRows so band formatting
//   travels with it; ~3 blank buffer rows above it stay typable for manual
//   SKU checks.
//
// MANUAL ENTRY (main table + buffer rows only)
//   Type a SKU into col A above the divider → outOfStockOnEdit auto-fills
//   LOCATION/QTY/SOLD/AVAILABLE/DAYS OUT + stamps FIRST SEEN if empty. The
//   kit table is machine-owned: edits at/below the divider are ignored.
//
// PUBLIC API
//   setupOutOfStockSheet()      — idempotent setup + LAYOUT MIGRATOR (old
//                                 single-table sheets get the title band, the
//                                 KITS divider, and the in-band chip)
//   refreshOutOfStock(maps)     — smart-merge both tables from MI + Kit
//                                 Registry + Zoho Stock (hourly via
//                                 runHourlyHousekeeping, or sidebar button)
//   openOutOfStock()            — activate the sheet
//   getOutOfStockCount()        — main OOS rows + kit rows (for the alert)
//   outOfStockOnEdit(e)         — onEdit dispatcher (called from Main.js)
//   setupOutOfStockTrigger()    — ⚠ superseded weekly trigger (see below)
//   removeOutOfStockTrigger()   — uninstall the weekly trigger
// =======================================================================================

// ---------- LOCAL SCHEMA (kept here, not in Schema.js — different sheet) ----------
var OUT_OF_STOCK = {
  sheetName: "Out of Stock",

  // 1-based column positions — MAIN table (plain SKUs to reorder)
  cols: {
    SKU:          1,   // A
    LOCATION:     2,   // B
    QTY:          3,   // C — Master Inventory `quantity`
    SOLD:         4,   // D — Master Inventory `quantitySold`
    AVAILABLE:    5,   // E — qty - sold (will be ≤ 0 for real OOS)
    FIRST_SEEN:   6,   // F — date this SKU first appeared as OOS (preserved)
    LAST_CHECKED: 7,   // G — timestamp of latest refresh that confirmed state
    DAYS_OUT:     8    // H — TODAY - FIRST SEEN, WRITTEN by refresh (no formula)
  },

  // The KIT table reuses the same 8-col grid with different meanings in A..E:
  kitCols: {
    KIT_SKU:      1,   // A
    LOCATION:     2,   // B — kit's own shelf (K-* = READY pre-assembled box)
    BUILDABLE:    3,   // C — min over components of floor(avail / qty-per-kit)
    LIMITED_BY:   4,   // D — bottleneck component + its numbers
    COMPONENTS:   5,   // E — "N ok" or "⚠ …" health
    FIRST_SEEN:   6,   // F
    LAST_CHECKED: 7,   // G
    DAYS_OUT:     8    // H
  },

  idx: function(name) { return OUT_OF_STOCK.cols[name] - 1; },

  dataWidth: 8,

  titleRow:     1,   // ▌ OUT OF STOCK band (chip in-band at H1, stamp J1 hidden)
  headerRow:    2,
  dataStartRow: 3,

  // Divider contract — col A value must be EXACTLY this (▌ dressing is a
  // number-format display prefix; Gotcha #1 discipline).
  boundaryMarker: "KITS",

  // Blank typable rows kept above the divider for manual SKU checks.
  bufferRows: 3,

  headers:    ["◈ SKU", "LOCATION", "# QTY", "# SOLD", "AVAILABLE", "FIRST SEEN", "LAST CHECKED", "DAYS OUT"],
  kitHeaders: ["📦 KIT SKU", "LOCATION", "BUILDABLE", "LIMITED BY", "COMPONENTS", "FIRST SEEN", "LAST CHECKED", "DAYS OUT"]
};


// =======================================================================================
// PUBLIC API
// =======================================================================================

/**
 * Idempotent setup + layout migrator. Safe to re-run any time (sidebar
 * "Re-style Sheet" button). Migrates pre-2026-07-18 single-table sheets:
 * inserts the row-1 title band, moves the sync chip in-band (H1, stamp stays
 * J1), removes the legacy DAYS OUT ARRAYFORMULA, and creates the KITS divider.
 */
function setupOutOfStockSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(OUT_OF_STOCK.sheetName);
  }

  // --- TITLE-BAND MIGRATION ---
  // Old layouts have the column headers on row 1. Shift everything down one
  // so row 1 becomes the ▌ OUT OF STOCK title band (Prep Queue pattern).
  var a1 = String(sheet.getRange(1, 1).getValue()).trim();
  if (a1.charAt(0) === '◈') {
    sheet.insertRowsBefore(1, 1);
  }

  // --- CHIP MIGRATION (dark I1 chip / J1 stamp era → H1 in-band / J1) ---
  // After the title insert the old chip sits at I2 and the old stamp at J2.
  // Carry the stamp Date back to J1, wipe the dark-chip leftovers, and strip
  // stale chip CF rules (they shifted to row 2 with the insert).
  try {
    var stampCell = sheet.getRange(SHEET_PULSE.outOfStock.stamp);   // J1
    if (!(stampCell.getValue() instanceof Date)) {
      var carried = sheet.getRange("J2").getValue();
      if (carried instanceof Date) stampCell.setValue(carried);
    }
    sheet.getRange("J2").clearContent();
    ["I1", "I2"].forEach(function(oldHome) {
      sheet.getRange(oldHome).clearContent()
           .setBackground(null).setFontColor(null)
           .setBorder(false, false, false, false, false, false);
    });
    sheet.getRange(1, 8, 2, 3).clearNote();   // chip notes, rows 1-2 × H..J
    sheet.setColumnWidth(9, 40);              // col I was 180 for the dark chip
    var cfMigrated = sheet.getConditionalFormatRules().filter(function(r) {
      return !r.getRanges().some(function(rg) {
        return rg.getRow() <= 2 && rg.getNumRows() === 1 &&
               rg.getColumn() >= 9 && rg.getColumn() <= 10;
      });
    });
    sheet.setConditionalFormatRules(cfMigrated);
  } catch (mErr) { try { Logger.log("setupOutOfStockSheet: chip migration: " + mErr); } catch (_) {} }

  // --- LEGACY DAYS OUT ARRAYFORMULA REMOVAL ---
  // DAYS OUT is written by refresh now — the kit table's second header row
  // would block an array spill in col H with #REF!.
  try {
    var h3 = sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.DAYS_OUT);
    if (String(h3.getFormula()).indexOf("=ARRAYFORMULA") === 0) h3.clearContent();
  } catch (fErr) { try { Logger.log("setupOutOfStockSheet: formula cleanup: " + fErr); } catch (_) {} }

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(OUT_OF_STOCK.cols.SKU,          120);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.LOCATION,     110);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.QTY,           70);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.SOLD,         140);  // doubles as LIMITED BY in the kit table
  sheet.setColumnWidth(OUT_OF_STOCK.cols.AVAILABLE,     90);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.FIRST_SEEN,   120);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.LAST_CHECKED, 140);
  sheet.setColumnWidth(OUT_OF_STOCK.cols.DAYS_OUT,      85);

  // --- DATA AREA: column-level formats so new rows inherit ---
  var maxDataRow = 1000;
  var dataRows = maxDataRow - OUT_OF_STOCK.dataStartRow + 1;
  _applyOosDataRowFormats(sheet, OUT_OF_STOCK.dataStartRow, dataRows);

  // --- BANDING (cream alternation) ---
  // Range starts at the HEADER row (2), NOT row 1 — the banding header slot
  // paints over manual fills, and it would black out the row-1 title band
  // (Prep Queue rollout lesson 2026-07-16). Row 1 stays OUTSIDE the banding,
  // and the band/header manual styling below runs AFTER this step.
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(OUT_OF_STOCK.headerRow, 1,
                                 maxDataRow - OUT_OF_STOCK.headerRow + 1, OUT_OF_STOCK.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  // --- TITLE BAND + MAIN HEADERS (after banding, so manual styling wins) ---
  _styleOosBand(sheet, OUT_OF_STOCK.titleRow, "OUT OF STOCK", "RESTOCK LIST · ACTIVE LISTINGS");
  _styleOosHeaderRow(sheet, OUT_OF_STOCK.headerRow, OUT_OF_STOCK.headers);

  // --- KITS DIVIDER + KIT HEADERS ---
  // Create below existing data (with typing buffer) on first run; re-style in
  // place on every re-run.
  var boundary = _getOosBoundaryRow(sheet);
  if (boundary < 0) {
    var lastContent = Math.max(sheet.getLastRow(), OUT_OF_STOCK.headerRow);
    boundary = lastContent + OUT_OF_STOCK.bufferRows + 1;
    sheet.getRange(boundary, 1).setValue(OUT_OF_STOCK.boundaryMarker);
  }
  _styleOosBand(sheet, boundary, OUT_OF_STOCK.boundaryMarker, "OOS KITS · BUILDABLE FROM COMPONENTS");
  _styleOosHeaderRow(sheet, boundary + 1, OUT_OF_STOCK.kitHeaders);

  // --- CONDITIONAL FORMATTING (wipe ours in cols A..E, rebuild) ---
  // Rule ORDER is load-bearing — Sheets applies the FIRST matching rule per
  // cell (Prep Queue DONE-strike lesson). Chip rules (col H) are untouched.
  var keptRules = sheet.getConditionalFormatRules().filter(function(r) {
    return !r.getRanges().some(function(rg) {
      return rg.getColumn() <= OUT_OF_STOCK.cols.AVAILABLE &&
             rg.getRow() >= OUT_OF_STOCK.headerRow;
    });
  });

  // Kit-table guard: a cell is "in the kit table" when its row is below the
  // KITS divider. MATCH is exact — the divider's underlying value is bare
  // "KITS" (the ▌ is number-format dressing).
  var kitsGuard = 'ROW()>IFERROR(MATCH("KITS",$A$1:$A,0),100000)';
  var dsr = OUT_OF_STOCK.dataStartRow;

  // Zoho-resolved mute (2026-07-20): MI still says OOS but Zoho's OWN entry
  // for this SKU already shows stock — a restock (or an "open this kit"
  // decision) that hasn't reached MI yet via the hourly eBay sync. refresh
  // stamps "⟳ Zoho: N" into LAST CHECKED (col G) for exactly these rows —
  // main-table restocks AND a kit's own SKU — so a picker reads "already
  // handled, MI is catching up" instead of "still broken." Applies to BOTH
  // tables from one full-width rule. ORDERED FIRST — Sheets applies the
  // first matching rule per cell, so a resolved row reads muted gray instead
  // of still-red/still-green (same rule-order discipline as Prep Queue's
  // DONE-strike, 2026-07-16). MI stays the decider on whether the row is
  // still listed at all — this rule only changes how a resolved row LOOKS
  // while it waits its turn; annotate, never override (same philosophy as
  // the queued eBay-row Zoho divergence flag).
  var muteRange = sheet.getRange(dsr, 1, dataRows, OUT_OF_STOCK.dataWidth);
  var muteRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($G' + dsr + ',"⟳ Zoho")')
    .setBackground('#e0e0e0').setFontColor('#757575').setStrikethrough(true)
    .setRanges([muteRange]).build();

  var buildableRange = sheet.getRange(dsr, OUT_OF_STOCK.kitCols.BUILDABLE, dataRows, 1);
  var buildableGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(ISNUMBER($C" + dsr + "),$C" + dsr + ">0," + kitsGuard + ")")
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([buildableRange]).build();
  var buildableRed = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(ISNUMBER($C" + dsr + "),$C" + dsr + "<=0," + kitsGuard + ")")
    .setBackground('#ff6b6b').setFontColor('#ffffff').setBold(true)
    .setRanges([buildableRange]).build();

  // ⚠ warning cells (kit table C..E: unparsed PD / missing components)
  var warnRange = sheet.getRange(dsr, OUT_OF_STOCK.kitCols.BUILDABLE, dataRows, 3);
  var warnRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT(C' + dsr + ',1)="⚠"')
    .setBackground('#fff3b0').setFontColor('#7a5c00').setBold(true)
    .setRanges([warnRange]).build();

  // Main table: AVAILABLE red when ≤ 0. ISNUMBER keeps it off the kit table
  // (kit col E holds text) and off hand-typed in-stock lookups.
  var availRange = sheet.getRange(dsr, OUT_OF_STOCK.cols.AVAILABLE, dataRows, 1);
  var availRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(ISNUMBER(E" + dsr + "), E" + dsr + "<=0)")
    .setBackground('#ff6b6b').setFontColor('#ffffff').setBold(true)
    .setRanges([availRange]).build();

  sheet.setConditionalFormatRules(
    [muteRule, buildableGreen, buildableRed, warnRule, availRule].concat(keptRules));

  // --- Duplicate SKU highlight (JS-side, structure-aware) ---
  _refreshOutOfStockDuplicates(sheet);

  // --- FREEZE TITLE BAND + HEADER ---
  sheet.setFrozenRows(2);

  // --- FRESHNESS PULSE CHIP (H1 in-band / J1 hidden stamp) ---
  try { _installPulseChip(sheet, SHEET_PULSE.outOfStock); }
  catch (e) { try { Logger.log("setupOutOfStockSheet: pulse chip error: " + e); } catch (_) {} }

  return "✅ Out of Stock sheet ready (two tables: RESTOCK + KITS).";
}


/**
 * Normalize a FIRST SEEN cell value to the canonical "M/d/yy" string.
 *
 * WHY (bug found 2026-07-13, surfaced by the hourly refresh): Sheets
 * auto-coerces written "4/30/26" strings into real Date values. The next
 * refresh read then got a Date object, and String(date) produced the JS
 * dump "Thu Apr 30 2026 08:00:00 GMT+0300 (GMT+03:00)" — which got written
 * back and broke downstream parsing. This helper repairs all three shapes on
 * read; the '@' plain-text format on the FIRST SEEN / LAST CHECKED columns
 * stops the coercion at the source.
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
 * Smart-merge refresh of BOTH tables.
 *
 *   MAIN table: plain Active-listing SKUs with available ≤ 0 (reorder list).
 *   KIT table:  Active OOS SKUs found in the Kit Registry, enriched with
 *               BUILDABLE / LIMITED BY / COMPONENTS from Kit Registry ×
 *               Zoho Stock (MI fallback), sorted most-buildable first.
 *
 * FIRST SEEN is preserved per SKU across refreshes AND across segments.
 * The KITS divider slides (insert/delete rows) so band formatting travels;
 * OUT_OF_STOCK.bufferRows blank rows above it stay typable.
 */
function refreshOutOfStock(maps) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) {
    setupOutOfStockSheet();
    sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  }

  // Accept pre-built maps from runHourlyHousekeeping (shared MI read).
  // Detect by shape, not presence — a time trigger passes an event object.
  if (!maps || !maps.inventoryMap) maps = buildLocationAndInventoryMaps();
  if (maps.inventoryMap.size === 0) {
    return "⚠️ Master Inventory empty or headers missing.";
  }

  // Self-heal the two-table layout: legacy header-on-row-1 sheets and lost
  // dividers (paste-over) both route through the setup migrator.
  var boundary = _getOosBoundaryRow(sheet);
  if (boundary < 0 || String(sheet.getRange(1, 1).getValue()).trim().charAt(0) === '◈') {
    setupOutOfStockSheet();
    boundary = _getOosBoundaryRow(sheet);
    if (boundary < 0) return "⚠️ Out of Stock: KITS divider missing after setup.";
  }

  // Kit + Zoho lookups (best-effort — an unavailable registry degrades to
  // "no kit table rows", never blocks the main refresh).
  var kitLookup = new Map();   // sku lower → kit object
  try {
    buildKitMap().forEach(function(kit, kitSku) {
      kitLookup.set(String(kitSku).trim().toLowerCase(), kit);
    });
  } catch (ke) { try { console.log("refreshOutOfStock: kit map unavailable: " + ke); } catch (_) {} }
  var zohoMap = new Map();
  try { zohoMap = buildZohoStockMap(); }
  catch (ze) { try { console.log("refreshOutOfStock: zoho map unavailable: " + ze); } catch (_) {} }
  var resolveAvail = _oosResolveAvailFactory(zohoMap, maps.inventoryMap);

  var todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy");
  var nowStr   = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

  // ---- Read existing rows (both segments) ----
  var FS = OUT_OF_STOCK.idx("FIRST_SEEN");
  var mainExisting = [];
  var mainSlots = boundary - OUT_OF_STOCK.dataStartRow;
  if (mainSlots > 0) {
    mainExisting = sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, mainSlots, 7).getValues();
  }
  var kitExisting = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= boundary + 2) {
    kitExisting = sheet.getRange(boundary + 2, 1, lastRow - boundary - 1, 7).getValues();
  }

  // FIRST SEEN index across BOTH segments — a SKU keeps its date when it
  // moves between tables (e.g. first two-table refresh moves kits down).
  var firstSeenBySku = {};
  [mainExisting, kitExisting].forEach(function(rows) {
    for (var i = 0; i < rows.length; i++) {
      var s = String(rows[i][0] || "").trim().toLowerCase();
      if (!s || s === OUT_OF_STOCK.boundaryMarker.toLowerCase()) continue;
      var fs = _normalizeOosFirstSeen(rows[i][FS], "");
      if (fs && !firstSeenBySku[s]) firstSeenBySku[s] = fs;
    }
  });

  // ---- Build MAIN rows ----
  var keptSkus = {};
  var mainRows = [];
  var dropped = 0, refreshed = 0, preserved = 0, inactiveDropped = 0, zohoResolvedCount = 0;

  for (var i = 0; i < mainExisting.length; i++) {
    var row = mainExisting[i];
    var rawSku = String(row[OUT_OF_STOCK.idx("SKU")] || "").trim();
    if (!rawSku) continue;   // buffer / empty rows
    var skuLower = rawSku.toLowerCase();

    var inv = maps.inventoryMap.get(skuLower);

    if (!inv) {
      // Not in MI — leave the row exactly as the user wrote it (manual entry)
      mainRows.push(row.slice(0, 7).concat([_oosDaysOut(_normalizeOosFirstSeen(row[FS], ""))]));
      keptSkus[skuLower] = true;
      preserved++;
      continue;
    }

    if (!_oosIsActive(inv)) {
      // Listing deleted/ended on eBay — phantom, drop it.
      inactiveDropped++;
      continue;
    }

    if (inv.available <= 0 && kitLookup.has(skuLower)) {
      // OOS kit — the kit table owns it now.
      continue;
    }

    if (inv.available > 0) {
      var sheetAvail = row[OUT_OF_STOCK.idx("AVAILABLE")];
      if (typeof sheetAvail === 'number' && sheetAvail <= 0) {
        // Was OOS, MI says restocked → drop.
        dropped++;
        continue;
      }
      // User deliberately typed an in-stock SKU (watch/lookup row) — keep,
      // refresh values, preserve FIRST SEEN.
      var locationKept  = maps.locationMap.get(skuLower) || row[OUT_OF_STOCK.idx("LOCATION")] || "";
      var firstSeenKept = _normalizeOosFirstSeen(row[FS], todayStr);
      mainRows.push([rawSku, locationKept, inv.quantity, inv.sold, inv.available,
                     firstSeenKept, nowStr, _oosDaysOut(firstSeenKept)]);
      keptSkus[skuLower] = true;
      preserved++;
      continue;
    }

    // Still OOS — refresh values, preserve FIRST SEEN
    var location  = maps.locationMap.get(skuLower) || row[OUT_OF_STOCK.idx("LOCATION")] || "";
    var firstSeen = _normalizeOosFirstSeen(row[FS], todayStr);
    var zohoR1 = _oosZohoResolved(skuLower, zohoMap);
    if (zohoR1 !== null) zohoResolvedCount++;
    mainRows.push([rawSku, location, inv.quantity, inv.sold, inv.available,
                   firstSeen, _oosLastCheckedText(nowStr, zohoR1), _oosDaysOut(firstSeen)]);
    keptSkus[skuLower] = true;
    refreshed++;
  }

  // Append newly-OOS plain SKUs (Active only, kits excluded)
  var added = 0;
  maps.inventoryMap.forEach(function(inv, skuLower) {
    if (inv.available > 0) return;
    if (!_oosIsActive(inv)) return;
    if (keptSkus[skuLower]) return;
    if (kitLookup.has(skuLower)) return;

    var location = maps.locationMap.get(skuLower) || "";
    var zohoR2 = _oosZohoResolved(skuLower, zohoMap);
    if (zohoR2 !== null) zohoResolvedCount++;
    mainRows.push([skuLower, location, inv.quantity, inv.sold, inv.available,
                   todayStr, _oosLastCheckedText(nowStr, zohoR2), 0]);
    added++;
  });

  mainRows.sort(_oosMainRowCompare);

  // ---- Build KIT rows ----
  var kitRows = [];
  var kitBuildableNow = 0;
  maps.inventoryMap.forEach(function(inv, skuLower) {
    if (inv.available > 0) return;
    if (!_oosIsActive(inv)) return;
    var kit = kitLookup.get(skuLower);
    if (!kit) return;

    var build = _oosComputeKitBuild(kit, resolveAvail);
    if (typeof build.buildable === 'number' && build.buildable > 0) kitBuildableNow++;

    // System-wide convention (user ruling 2026-07-18): an item with no
    // location shows the literal "NOT FOUND" — same as every other surface.
    // locationMap already stores "NOT FOUND" for MI rows with a blank
    // location (typical for MANUAL kits — no pre-assembled box on a shelf);
    // READY kits carry their real K-* aisle. Registry KIT LOC is the
    // fallback for the theoretical MI-absent case.
    var kitLoc = maps.locationMap.get(skuLower) || kit.location || "NOT FOUND";

    var kitFirstSeen = firstSeenBySku[skuLower] || todayStr;
    // Mute check on the KIT'S OWN sku (not its components) — same "MI still
    // says OOS but Zoho already shows stock" signal as the main table. Fires
    // e.g. right after the floor opens a kit and adjusts its qty in Zoho.
    var zohoR3 = _oosZohoResolved(skuLower, zohoMap);
    if (zohoR3 !== null) zohoResolvedCount++;
    kitRows.push([kit.sku, kitLoc, build.buildable, build.limitedBy, build.components,
                  kitFirstSeen, _oosLastCheckedText(nowStr, zohoR3), _oosDaysOut(kitFirstSeen)]);
  });
  kitRows.sort(_oosKitRowCompare);

  // ---- Structural write ----
  // Slide the divider so the main segment is exactly mainRows + buffer.
  var targetSlots = mainRows.length + OUT_OF_STOCK.bufferRows;
  var currentSlots = boundary - OUT_OF_STOCK.dataStartRow;
  if (targetSlots > currentSlots) {
    var addN = targetSlots - currentSlots;
    sheet.insertRowsAfter(boundary - 1, addN);
    if (currentSlots === 0) {
      // Degenerate case: the inheritance source row was the HEADER — reset
      // the inserted rows to a data-row look (Prep Queue buffer lesson).
      var fresh = sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, addN, OUT_OF_STOCK.dataWidth);
      fresh.setBackground(null).setFontColor(null).setFontLine('none')
           .setBorder(false, false, false, false, false, false);
      _applyOosDataRowFormats(sheet, OUT_OF_STOCK.dataStartRow, addN);
    }
  } else if (targetSlots < currentSlots) {
    sheet.deleteRows(OUT_OF_STOCK.dataStartRow + targetSlots, currentSlots - targetSlots);
  }
  boundary = OUT_OF_STOCK.dataStartRow + targetSlots;

  // Main segment: clear + write
  sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, targetSlots, OUT_OF_STOCK.dataWidth).clearContent();
  if (mainRows.length > 0) {
    sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, mainRows.length, OUT_OF_STOCK.dataWidth)
      .setValues(mainRows);
  }

  // Kit segment: clear + write (content-only — nothing lives below it)
  var kitStart = boundary + 2;
  var clearTo = Math.max(sheet.getLastRow(), kitStart + kitRows.length - 1);
  if (clearTo >= kitStart) {
    sheet.getRange(kitStart, 1, clearTo - kitStart + 1, OUT_OF_STOCK.dataWidth).clearContent();
  }
  if (kitRows.length > 0) {
    sheet.getRange(kitStart, 1, kitRows.length, OUT_OF_STOCK.dataWidth).setValues(kitRows);
    // LIMITED BY + COMPONENTS are prose-ish — compact mono instead of the
    // main table's big Oswald digits (re-asserted every refresh; the whole
    // segment is rewritten anyway).
    sheet.getRange(kitStart, OUT_OF_STOCK.kitCols.LIMITED_BY, kitRows.length, 2)
      .setFontFamily('Roboto Mono').setFontWeight('normal').setFontSize(9)
      .setFontColor('#434343');
  }

  // Divider + kit header self-heal (cheap, keeps the KITS band styled even
  // after a stray paste-over; mirrors the Prep Queue re-style discipline)
  _styleOosBand(sheet, boundary, OUT_OF_STOCK.boundaryMarker, "OOS KITS · BUILDABLE FROM COMPONENTS");
  _styleOosHeaderRow(sheet, boundary + 1, OUT_OF_STOCK.kitHeaders);

  // Duplicate-SKU highlights (structure-aware)
  _refreshOutOfStockDuplicates(sheet);

  // Freshness chip — stamped only on a COMPLETED refresh, so the chip's
  // staleness tiers double as the "refresh trigger is dead" alarm.
  stampSheetPulse(sheet, SHEET_PULSE.outOfStock.stamp);

  return "✅ Out of Stock refreshed — " +
         added + " new, " +
         refreshed + " still out, " +
         dropped + " restocked, " +
         kitRows.length + " OOS kit(s)" +
         (kitBuildableNow > 0 ? " (" + kitBuildableNow + " buildable now)" : "") +
         (inactiveDropped > 0 ? ", " + inactiveDropped + " inactive dropped" : "") +
         (zohoResolvedCount > 0 ? ", " + zohoResolvedCount + " zoho-resolved (MI catching up)" : "") +
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
 * Manual re-sort of the MAIN table — LOCATION asc (empty last), then SKU
 * asc: the exact rule the hourly refresh already applies. Useful right
 * after typing several SKUs into the buffer rows without waiting for the
 * next refresh. Pure in-place reorder (getValues → sort → setValues, no
 * MI/Kit Registry/Zoho reads) — instant. Column-level fonts and the '@'
 * text format are uniform per column, so they travel automatically with
 * the values; only the position-dependent duplicate-SKU highlight needs a
 * repaint after (same reasoning as All Orders' sort — see Gotcha #3 — minus
 * the per-cell NUMBER FORMAT complication, since nothing here varies it
 * row-by-row).
 */
function sortOutOfStockMain() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) return "⚠️ Out of Stock sheet not found.";

  var boundary = _getOosBoundaryRow(sheet);
  if (boundary < 0) return "⚠️ KITS divider missing — run Re-style Sheet first.";

  var n = boundary - OUT_OF_STOCK.dataStartRow;
  if (n <= 0) return "Nothing to sort — the restock list is empty.";

  var range = sheet.getRange(OUT_OF_STOCK.dataStartRow, 1, n, OUT_OF_STOCK.dataWidth);
  var rows = range.getValues();
  rows.sort(_oosMainRowCompare);
  range.setValues(rows);

  _refreshOutOfStockDuplicates(sheet);
  return "✅ Restock list sorted by location.";
}

/**
 * Manual re-sort of the KIT table — BUILDABLE desc ("⚠" rows last), then kit
 * SKU asc. The kit table is fully rewritten in this order on every refresh
 * already; this button is for re-asserting the order on demand (e.g. right
 * after eyeballing the sheet) without kicking off a full MI/Kit Registry/
 * Zoho refresh cycle.
 */
function sortOutOfStockKits() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) return "⚠️ Out of Stock sheet not found.";

  var boundary = _getOosBoundaryRow(sheet);
  if (boundary < 0) return "⚠️ KITS divider missing — run Re-style Sheet first.";

  var kitStart = boundary + 2;
  var lastRow = sheet.getLastRow();
  var n = lastRow - kitStart + 1;
  if (n <= 0) return "Nothing to sort — no OOS kits right now.";

  var range = sheet.getRange(kitStart, 1, n, OUT_OF_STOCK.dataWidth);
  var rows = range.getValues();
  rows.sort(_oosKitRowCompare);
  range.setValues(rows);

  _refreshOutOfStockDuplicates(sheet);
  return "✅ Kit table sorted by buildable count.";
}


/**
 * Fast count for the sidebar alert. Main table: rows where AVAILABLE ≤ 0
 * (manual in-stock lookups and NOT FOUND rows excluded, as before). Kit
 * table: every row counts — kit rows are OOS by construction.
 *
 * Reads the snapshot — does NOT re-scan Master Inventory.
 */
function getOutOfStockCount() {
  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OUT_OF_STOCK.sheetName);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < OUT_OF_STOCK.dataStartRow) return 0;

  var boundary = _getOosBoundaryRow(sheet);
  var data = sheet.getRange(
    OUT_OF_STOCK.dataStartRow, 1,
    lastRow - OUT_OF_STOCK.dataStartRow + 1,
    OUT_OF_STOCK.cols.AVAILABLE
  ).getValues();

  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var rowNum = OUT_OF_STOCK.dataStartRow + i;
    if (_isOosStructureRow(rowNum, boundary)) continue;
    var sku = String(data[i][OUT_OF_STOCK.idx("SKU")]).trim();
    if (!sku) continue;
    if (boundary > 0 && rowNum > boundary) { count++; continue; }   // kit table
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
 * SKU edit on Out of Stock (MAIN table + buffer rows only) → auto-fill
 * LOCATION + QTY + SOLD + AVAILABLE + DAYS OUT + stamp FIRST SEEN (only if
 * currently empty) + LAST CHECKED.
 *
 * Edits at/below the KITS divider are ignored — the kit table is machine-
 * owned and rewritten wholesale by every refresh.
 *
 * Mirrors prepQueueOnEdit. Lives on the INSTALLABLE trigger because the
 * Master Inventory lookup goes through openById which simple triggers can't
 * call reliably. Defensive try/catch — never blocks other edit handlers.
 */
function outOfStockOnEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== OUT_OF_STOCK.sheetName) return;
    if (e.range.getColumn() !== OUT_OF_STOCK.cols.SKU) return;
    if (e.range.getRow() < OUT_OF_STOCK.dataStartRow) return;

    var boundary = _getOosBoundaryRow(sheet);
    var startRow = e.range.getRow();
    if (boundary > 0 && startRow >= boundary) return;   // kit table is machine-owned

    var edits = e.range.getValues();
    // Clamp a multi-row paste so it never spills into the divider/kit table
    var rowCount = (boundary > 0) ? Math.min(edits.length, boundary - startRow) : edits.length;

    // Pre-build inventory + location maps if multi-row paste, otherwise
    // single-row lookups are cheaper.
    var useMap = rowCount > 3;
    var locationMap = null;
    var inventoryMap = null;
    if (useMap) {
      var maps = buildLocationAndInventoryMaps();
      locationMap = maps.locationMap;
      inventoryMap = maps.inventoryMap;
    }

    var todayStr = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy");
    var nowStr   = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");

    for (var i = 0; i < rowCount; i++) {
      var row = startRow + i;
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      if (skuLower === "") {
        // SKU cleared — wipe the row's lookup fields (B..H) so it reads empty
        sheet.getRange(row, OUT_OF_STOCK.cols.LOCATION, 1, 7).clearContent();
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

      // Preserve existing FIRST SEEN if non-empty (normalized); else today
      var firstSeen = _normalizeOosFirstSeen(
        sheet.getRange(row, OUT_OF_STOCK.cols.FIRST_SEEN).getValue(), todayStr);

      sheet.getRange(row, OUT_OF_STOCK.cols.LOCATION, 1, 7).setValues([[
        location, qty, sold, avail, firstSeen, nowStr, _oosDaysOut(firstSeen)
      ]]);
    }

    // Keep typable blank rows above the divider (light version of the Prep
    // Queue buffer machinery — the hourly refresh re-normalizes to exactly
    // OUT_OF_STOCK.bufferRows anyway).
    if (boundary > 0) {
      var lastEdited = startRow + rowCount - 1;
      if (boundary - lastEdited <= 2) {
        sheet.insertRowsAfter(lastEdited, OUT_OF_STOCK.bufferRows);
      }
    }

    // Refresh duplicate highlights — every SKU edit could create or clear a dupe
    _refreshOutOfStockDuplicates(sheet);
  } catch (err) {
    try { Logger.log("outOfStockOnEdit error: " + err); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE: two-table structure helpers
// =======================================================================================

/**
 * Find the KITS divider row — exact-match contract on col A from the data
 * band down (same discipline as All Orders' getBoundaryRow / Prep Queue's
 * INCOMING). Returns the 1-based row number, or -1 when the sheet has no
 * divider yet (legacy single-table layout).
 */
function _getOosBoundaryRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < OUT_OF_STOCK.dataStartRow) return -1;
  var vals = sheet.getRange(OUT_OF_STOCK.dataStartRow, 1,
                            lastRow - OUT_OF_STOCK.dataStartRow + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === OUT_OF_STOCK.boundaryMarker) {
      return OUT_OF_STOCK.dataStartRow + i;
    }
  }
  return -1;
}

/** True for the two structural rows every walker must skip: divider + kit header. */
function _isOosStructureRow(row, boundary) {
  return boundary > 0 && (row === boundary || row === boundary + 1);
}

/**
 * Style a brand-yellow ▌ band (row-1 title + KITS divider — both wear the
 * same identity, mirroring Prep Queue). markerText lands as the cell VALUE
 * (the ▌ is number-format dressing), so for the divider it MUST be exactly
 * OUT_OF_STOCK.boundaryMarker to keep _getOosBoundaryRow's strict match.
 */
function _styleOosBand(sheet, row, markerText, rightLabel) {
  sheet.getRange(row, 1, 1, OUT_OF_STOCK.dataWidth)
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
  // Right label sits in col E, right-aligned — it overflows LEFT over the
  // empty band cells. F..H stay empty (H1 is the in-band chip's home on the
  // title row; the chip installer restyles that one cell after this runs).
  sheet.getRange(row, OUT_OF_STOCK.cols.AVAILABLE)
    .setValue(rightLabel)
    .setHorizontalAlignment('right')
    .setFontSize(9);
  sheet.setRowHeight(row, 36);
}

/**
 * Style one header row — dark brand band, yellow Oswald text, thick yellow
 * underline. Used for the main headers (row 2) and the kit-table header
 * (divider + 1), which carry DIFFERENT header sets on the same 8-col grid.
 */
function _styleOosHeaderRow(sheet, row, headers) {
  sheet.getRange(row, 1, 1, OUT_OF_STOCK.dataWidth)
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
  sheet.getRange(row, 1, 1, OUT_OF_STOCK.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/**
 * Column-level data formats for a run of rows. Used by setup for the whole
 * band and by refresh's degenerate insert case (rows inserted when the main
 * segment was empty inherit the HEADER's format and need a reset).
 */
function _applyOosDataRowFormats(sheet, startRow, numRows) {
  sheet.getRange(startRow, OUT_OF_STOCK.cols.SKU, numRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, OUT_OF_STOCK.cols.LOCATION, numRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, OUT_OF_STOCK.cols.QTY, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, OUT_OF_STOCK.cols.SOLD, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(startRow, OUT_OF_STOCK.cols.AVAILABLE, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  // FIRST SEEN + LAST CHECKED are PLAIN TEXT ('@') on purpose: the code
  // writes "M/d/yy" strings, and without this format Sheets auto-coerces
  // them into real Date values (date-corruption bug fixed 2026-07-13; see
  // _normalizeOosFirstSeen).
  sheet.getRange(startRow, OUT_OF_STOCK.cols.FIRST_SEEN, numRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center')
    .setNumberFormat('@');
  sheet.getRange(startRow, OUT_OF_STOCK.cols.LAST_CHECKED, numRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9)
    .setFontColor('#434343').setHorizontalAlignment('center')
    .setNumberFormat('@');
  // Explicit '0' number format is LOAD-BEARING: the retired TODAY()-minus
  // ARRAYFORMULA left a stale DATE format on this column, which rendered the
  // newly WRITTEN day counts as "1/13/1900"-style dates (serial N = N days —
  // values were right, display was garbage; caught in the 2026-07-18 live
  // audit). Converting a formula column to written values must always reset
  // the number format (sibling of the clearContent-keeps-formats gotcha).
  sheet.getRange(startRow, OUT_OF_STOCK.cols.DAYS_OUT, numRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(13)
    .setFontColor('#1d1d1b').setHorizontalAlignment('center')
    .setNumberFormat('0');

  sheet.getRange(startRow, 1, numRows, OUT_OF_STOCK.dataWidth)
    .setVerticalAlignment('middle');
}


// =======================================================================================
// PRIVATE: pure decision/compute helpers (Node-validated 2026-07-18)
// =======================================================================================

/**
 * Active-listing test. Blank status fails OPEN (treated as Active) so a
 * renamed/missing listingStatus column in MI can never make the whole
 * catalog read as inactive and empty the sheet.
 */
function _oosIsActive(inv) {
  var s = String((inv && inv.status) || "").trim();
  return s === "" || s === "Active";
}

/**
 * Component-availability resolver: Zoho Stock first (fresh every 2 min,
 * covers DIRECT-side items not listed on eBay), Master Inventory fallback,
 * null when neither knows the SKU. Same routing rule as the Kit Expansion
 * modal and recomputeHand.
 */
function _oosResolveAvailFactory(zohoMap, inventoryMap) {
  return function(skuLower) {
    var z = zohoMap && zohoMap.get(skuLower);
    if (z) return z.available;
    var mi = inventoryMap && inventoryMap.get(skuLower);
    if (mi) return mi.available;
    return null;
  };
}

/**
 * Zoho-resolved check for a SKU MI already judges OOS. DIRECT zohoMap lookup
 * — deliberately NO MI fallback (the whole point is comparing the two
 * sources against each other, not resolving one from the other). Returns
 * the Zoho `available` number (rounded to 2dp) when Zoho's OWN entry for
 * this exact SKU already shows stock — a restock (or an "open this kit"
 * qty adjustment) that hasn't reached MI yet via the hourly eBay sync.
 * Returns null when there's nothing to report (Zoho doesn't carry the SKU,
 * or Zoho agrees it's still out).
 */
function _oosZohoResolved(skuLower, zohoMap) {
  var z = zohoMap && zohoMap.get(skuLower);
  if (!z || !(z.available > 0)) return null;
  return Math.round(z.available * 100) / 100;
}

/**
 * LAST CHECKED cell text: the plain timestamp, or the timestamp plus a
 * "⟳ Zoho: N" tag when _oosZohoResolved found a pending restock. The mute
 * CF rule (setupOutOfStockSheet) keys off this EXACT tag via
 * REGEXMATCH($G,"⟳ Zoho") — the two must stay in sync if either changes.
 */
function _oosLastCheckedText(nowStr, zohoResolvedValue) {
  return zohoResolvedValue === null ? nowStr : (nowStr + "  ⟳ Zoho: " + zohoResolvedValue);
}

/**
 * The kit table's headline math.
 *
 *   BUILDABLE = min over components of floor(available ÷ qty-per-kit)
 *   — the number of COMPLETE kits assemblable right now; ≥1 only when every
 *   single component covers at least one full kit.
 *
 * Untrustable states never show a number (buildable = "⚠", painted amber by
 * CF): unparsed PD lines in the registry, zero registered components, or a
 * component unknown to both Zoho Stock and MI.
 *
 * @param {Object} kit          buildKitMap() entry ({components, unparsedLines, …})
 * @param {Function} resolveAvail  skuLower → available (number) or null
 * @return {{buildable: (number|string), limitedBy: string, components: string}}
 */
function _oosComputeKitBuild(kit, resolveAvail) {
  var unparsed = (kit.unparsedLines || []).length;
  if (unparsed > 0) {
    return {
      buildable: "⚠",
      limitedBy: "⚠ PD unreadable — fix in Zoho",
      components: "⚠ " + unparsed + " unparsed"
    };
  }

  var comps = kit.components || [];
  if (comps.length === 0) {
    return {
      buildable: "⚠",
      limitedBy: "⚠ no components registered",
      components: "⚠ none"
    };
  }

  var missing = [];
  var minBuild = Infinity;
  var limiter = null;

  for (var i = 0; i < comps.length; i++) {
    var comp = comps[i];
    var qtyPer = (comp.qty > 0) ? comp.qty : 1;
    var avail = resolveAvail(String(comp.sku).trim().toLowerCase());
    if (avail === null) {
      missing.push(String(comp.sku));
      continue;
    }
    var can = Math.floor(avail / qtyPer);
    if (can < minBuild) {
      minBuild = can;
      limiter = { sku: String(comp.sku), avail: avail, qtyPer: qtyPer };
    }
  }

  if (missing.length > 0) {
    return {
      buildable: "⚠",
      limitedBy: "⚠ not found: " + missing.join(", "),
      components: "⚠ " + missing.length + " of " + comps.length + " missing"
    };
  }

  var availDisp = Math.round(limiter.avail * 100) / 100;
  return {
    buildable: minBuild,
    limitedBy: limiter.sku + " · has " + availDisp + " / needs " + limiter.qtyPer,
    components: comps.length + " ok"
  };
}

/**
 * DAYS OUT from a "M/d/yy" FIRST SEEN string (today in America/Chicago).
 * Returns "" for unparseable input or future dates. Written as a plain value
 * by refresh/onEdit — see the file header for why the ARRAYFORMULA died.
 */
function _oosDaysOut(firstSeenStr) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(String(firstSeenStr || "").trim());
  if (!m) return "";
  var y = parseInt(m[3], 10);
  if (y < 100) y += 2000;
  var first = new Date(y, parseInt(m[1], 10) - 1, parseInt(m[2], 10));
  var t = Utilities.formatDate(new Date(), "America/Chicago", "M/d/yyyy").split("/");
  var today = new Date(parseInt(t[2], 10), parseInt(t[0], 10) - 1, parseInt(t[1], 10));
  var days = Math.round((today.getTime() - first.getTime()) / 86400000);
  return days >= 0 ? days : "";
}

/** Main-table sort: LOCATION asc (empty last), then SKU asc. */
function _oosMainRowCompare(a, b) {
  var la = String(a[1] || "");
  var lb = String(b[1] || "");
  if (la === "" && lb !== "") return 1;
  if (lb === "" && la !== "") return -1;
  if (la !== lb) return la.localeCompare(lb);
  return String(a[0]).localeCompare(String(b[0]));
}

/**
 * Kit-table sort: BUILDABLE desc (what can be opened NOW on top), "⚠" rows
 * after all numbers, kit SKU asc as the tiebreak.
 */
function _oosKitRowCompare(a, b) {
  var ba = a[2], bb = b[2];
  var na = (typeof ba === 'number');
  var nb = (typeof bb === 'number');
  if (na && nb && ba !== bb) return bb - ba;
  if (na !== nb) return na ? -1 : 1;
  return String(a[0]).localeCompare(String(b[0]));
}


// =======================================================================================
// PRIVATE: duplicate-SKU highlight (JS-side, mirrors the All Orders pattern)
// =======================================================================================

/**
 * Scans column A, identifies SKUs that appear more than once, and paints the
 * SKU cell with a soft-amber background + thick yellow border so duplicates
 * stand out from both the cream banding and the red AVAILABLE highlight.
 *
 * STRUCTURE-AWARE (2026-07-18): the KITS divider and the kit header row are
 * never counted OR painted/cleared — clearing them would strip the yellow
 * band / dark header styling. Both tables' data rows participate.
 *
 * Cleared cells (no SKU) get reset to default. Run after any change that
 * could affect column A: setupOutOfStockSheet, refreshOutOfStock, outOfStockOnEdit.
 */
function _refreshOutOfStockDuplicates(sheet) {
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < OUT_OF_STOCK.dataStartRow) return;

  var boundary = _getOosBoundaryRow(sheet);
  var dataRows = lastRow - OUT_OF_STOCK.dataStartRow + 1;
  var skuRange = sheet.getRange(OUT_OF_STOCK.dataStartRow, OUT_OF_STOCK.cols.SKU, dataRows, 1);
  var skus = skuRange.getValues();

  // Count occurrences (case-insensitive, trimmed), skipping structure rows
  var counts = {};
  for (var i = 0; i < skus.length; i++) {
    if (_isOosStructureRow(OUT_OF_STOCK.dataStartRow + i, boundary)) continue;
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
    var rowNum = OUT_OF_STOCK.dataStartRow + i;
    if (_isOosStructureRow(rowNum, boundary)) continue;   // never touch band styling
    var k = String(skus[i][0]).trim().toLowerCase();
    var cell = sheet.getRange(rowNum, OUT_OF_STOCK.cols.SKU);
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
