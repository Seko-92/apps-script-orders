// =======================================================================================
// KIT_REGISTRY.gs — Repair-kit composition registry + expansion engine
// =======================================================================================
//
// PURPOSE
//   Repair/overhaul kits sold as single SKUs on eBay physically represent 5-12
//   components at different warehouse aisles. The picker today either knows the
//   composition by heart or manually types each component row. This registry +
//   the expansion button shipped alongside it turn that into one click.
//
// SOURCE OF TRUTH
//   The kit composition data lives in Zoho's per-item "Purchase Description"
//   field — already maintained by the team, structured enough to parse with a
//   single regex (validated 97.9% per-component-line success across 209 kits).
//   This sheet is a CACHED PROJECTION of that data — refreshed from a Zoho CSV
//   export. The human source of truth stays in Zoho; no double maintenance.
//
// DESIGN — MANUAL BUTTON, NOT AUTO-EXPAND
//   Expansion fires from a sidebar button on selected rows, NOT automatically
//   on eBay arrival. Picker-in-the-loop matters for kits because:
//     1. Registry can go stale (until Zoho sync is solved) — manual button gives
//        the picker a chance to spot-check before committing
//     2. Kits are higher-stakes (more parts = more chances for error)
//     3. Same UX for eBay-arrived kits AND DIRECT-typed kits — no surprises
//     4. Reversible — wrong expansion can be undone before pollution spreads
//
// KIT TYPES
//   READY  — pre-assembled box, lives at K-* aisles, ships as one unit.
//            Expansion is REFUSED for READY kits in the UI.
//   MANUAL — components live at separate aisles; picker walks the list.
//            Expansion produces N component rows under the kit's SALES_ORDER.
//
// SCHEMA
//   KIT_SKU · KIT_NAME · KIT_TYPE · KIT_LOCATION · KIT_ENGINE · COMPONENT_SKU
//   · COMPONENT_QTY · COMPONENT_NAME · SALES_DESC · LAST_UPDATED
//   One row per (kit, component) pair. Denormalized for read simplicity —
//   kit-level fields repeat across the kit's component rows. Sheet is
//   read-only after import; no manual maintenance hazard from duplication.
//
//   SALES_DESC (added 2026-05-19) — Zoho's Sales Description for the kit, kept
//   as a raw text blob (NOT parsed). Used by the Kit Expansion modal as a
//   reference panel so the picker can see which BOM components are physically
//   packed together (Sales Description shows the as-shipped packaging view;
//   Purchase Description shows the BOM view). Repeated across the kit's
//   component rows — duplication is mild on ~7 component rows per kit, lookup
//   is trivial.
//
// PUBLIC API
//   setupKitRegistrySheet()              — one-time: create sheet, brand styling
//   importKitsFromZohoCsv(driveFileId)   — re-import from a Zoho CSV in Drive
//   buildKitMap()                        — returns Map<kitSku, kitObject>
//                                          (lazy-built per call; cheap on ~1500 rows)
//   getKitInfo(kitSku)                   — Map lookup, returns kit obj or null
//   openKitRegistry()                    — sidebar: switch to Kit Registry sheet
// =======================================================================================


// ---------- LOCAL SCHEMA ----------
var KIT_REGISTRY = {
  sheetName: "Kit Registry",

  // 1-based column positions
  cols: {
    KIT_SKU:        1,   // A
    KIT_NAME:       2,   // B
    KIT_TYPE:       3,   // C — "READY" or "MANUAL"
    KIT_LOCATION:   4,   // D — kit's own aisle in MI (for KIT_TYPE derivation/audit)
    KIT_ENGINE:     5,   // E — first line of Purchase Description (engine model)
    COMPONENT_SKU:  6,   // F
    COMPONENT_QTY:  7,   // G
    COMPONENT_NAME: 8,   // H
    SALES_DESC:     9,   // I — Zoho Sales Description raw text (kit-level, repeated across rows)
    LAST_UPDATED:  10    // J
  },

  idx: function(name) { return KIT_REGISTRY.cols[name] - 1; },

  dataWidth:   10,
  headerRow:    1,
  dataStartRow: 2,

  headers: [
    "📦 KIT SKU", "KIT NAME", "TYPE", "KIT LOC", "ENGINE",
    "◈ COMP SKU", "# QTY", "COMPONENT NAME", "SALES DESCRIPTION", "⏱ UPDATED"
  ],

  types: { READY: "READY", MANUAL: "MANUAL" }
};


// =======================================================================================
// SETUP
// =======================================================================================

/**
 * Idempotent. Creates the Kit Registry sheet if missing, applies brand styling.
 * Safe to re-run — preserves existing data, just refreshes formatting.
 */
function setupKitRegistrySheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);

  if (!sheet) sheet = ss.insertSheet(KIT_REGISTRY.sheetName);

  // --- HEADERS ---
  sheet.getRange(KIT_REGISTRY.headerRow, 1, 1, KIT_REGISTRY.dataWidth)
    .setValues([KIT_REGISTRY.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(KIT_REGISTRY.headerRow, 1, 1, KIT_REGISTRY.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(KIT_REGISTRY.cols.KIT_SKU,         95);
  sheet.setColumnWidth(KIT_REGISTRY.cols.KIT_NAME,       260);
  sheet.setColumnWidth(KIT_REGISTRY.cols.KIT_TYPE,        80);
  sheet.setColumnWidth(KIT_REGISTRY.cols.KIT_LOCATION,    80);
  sheet.setColumnWidth(KIT_REGISTRY.cols.KIT_ENGINE,     130);
  sheet.setColumnWidth(KIT_REGISTRY.cols.COMPONENT_SKU,  105);
  sheet.setColumnWidth(KIT_REGISTRY.cols.COMPONENT_QTY,   55);
  sheet.setColumnWidth(KIT_REGISTRY.cols.COMPONENT_NAME, 260);
  sheet.setColumnWidth(KIT_REGISTRY.cols.SALES_DESC,     320);
  sheet.setColumnWidth(KIT_REGISTRY.cols.LAST_UPDATED,   140);

  // --- DATA AREA: column-level format so new imports inherit ---
  var maxDataRow = 2500;  // ~209 kits × ~7 components + headroom
  var dataRows = maxDataRow - KIT_REGISTRY.dataStartRow + 1;

  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_TYPE, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_LOCATION, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(10).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_ENGINE, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#434343').setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.COMPONENT_SKU, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.COMPONENT_QTY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.COMPONENT_NAME, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left');
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.SALES_DESC, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(9).setFontColor('#434343').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.LAST_UPDATED, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');

  sheet.getRange(KIT_REGISTRY.dataStartRow, 1, dataRows, KIT_REGISTRY.dataWidth)
    .setVerticalAlignment('middle');

  // --- KIT_TYPE CONDITIONAL FORMATTING (READY = green, MANUAL = yellow) ---
  // Refresh CF rules — strip prior KitRegistry rules, re-add.
  var existingRules = sheet.getConditionalFormatRules() || [];
  var keep = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    // Drop any rule scoped to KIT_TYPE column on this sheet
    return !ranges.some(function(rng) {
      return rng.getSheet().getName() === KIT_REGISTRY.sheetName
          && rng.getColumn() === KIT_REGISTRY.cols.KIT_TYPE;
    });
  });
  var typeRange = sheet.getRange(KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_TYPE, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_REGISTRY.types.READY)
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([typeRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(KIT_REGISTRY.types.MANUAL)
    .setBackground('#fff4b0').setFontColor('#1d1d1b').setBold(true)
    .setRanges([typeRange]).build());
  sheet.setConditionalFormatRules(keep);

  // --- BANDING (cream alternation) ---
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  var bandRange = sheet.getRange(1, 1, maxDataRow, KIT_REGISTRY.dataWidth);
  var band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b')
      .setFirstRowColor('#ffffff')
      .setSecondRowColor('#fff8e7');

  sheet.setFrozenRows(1);

  return "✅ Kit Registry sheet ready.";
}


/**
 * Sidebar: switch view to Kit Registry sheet.
 */
function openKitRegistry() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  if (!sheet) {
    setupKitRegistrySheet();
    sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened " + KIT_REGISTRY.sheetName;
}


// =======================================================================================
// READ API — used by the expansion engine + sidebar preview
// =======================================================================================

/**
 * Builds a Map<kitSku, kitObject> from the Kit Registry sheet.
 *
 * kitObject = {
 *   sku, name, type ("READY"|"MANUAL"), location, engine, salesDescription,
 *   components: [{sku, qty, name}]
 * }
 *
 * One read of the full data range, grouped client-side. ~1500 rows is well
 * within a single getValues call — cost is negligible.
 *
 * Returns an empty Map if the sheet doesn't exist (caller can decide whether
 * that's an error or a "kits feature is unconfigured" state).
 */
function buildKitMap() {
  var map = new Map();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  if (!sheet) return map;

  var lastRow = sheet.getLastRow();
  if (lastRow < KIT_REGISTRY.dataStartRow) return map;

  var data = sheet.getRange(
    KIT_REGISTRY.dataStartRow, 1,
    lastRow - KIT_REGISTRY.dataStartRow + 1,
    KIT_REGISTRY.dataWidth
  ).getValues();

  var KSKU_I  = KIT_REGISTRY.idx("KIT_SKU");
  var KNAME_I = KIT_REGISTRY.idx("KIT_NAME");
  var KTYPE_I = KIT_REGISTRY.idx("KIT_TYPE");
  var KLOC_I  = KIT_REGISTRY.idx("KIT_LOCATION");
  var KENG_I  = KIT_REGISTRY.idx("KIT_ENGINE");
  var CSKU_I  = KIT_REGISTRY.idx("COMPONENT_SKU");
  var CQTY_I  = KIT_REGISTRY.idx("COMPONENT_QTY");
  var CNAME_I = KIT_REGISTRY.idx("COMPONENT_NAME");
  var SD_I    = KIT_REGISTRY.idx("SALES_DESC");

  for (var i = 0; i < data.length; i++) {
    var kitSku = String(data[i][KSKU_I] || "").trim();
    if (!kitSku) continue;

    var existing = map.get(kitSku);
    if (!existing) {
      existing = {
        sku:              kitSku,
        name:             String(data[i][KNAME_I] || ""),
        type:             String(data[i][KTYPE_I] || "MANUAL").toUpperCase(),
        location:         String(data[i][KLOC_I] || ""),
        engine:           String(data[i][KENG_I] || ""),
        salesDescription: String(data[i][SD_I]   || ""),
        components:       []
      };
      map.set(kitSku, existing);
    }

    var compSku = String(data[i][CSKU_I] || "").trim();
    if (compSku) {
      existing.components.push({
        sku:  compSku,
        qty:  parseInt(data[i][CQTY_I]) || 1,
        name: String(data[i][CNAME_I] || "")
      });
    }
  }

  return map;
}


/**
 * Convenience: single-kit lookup. Builds map per call (cheap on ~1500 rows
 * but obviously O(n) on registry size; if this is called in a hot path, the
 * caller should buildKitMap() once and reuse).
 */
function getKitInfo(kitSku) {
  if (!kitSku) return null;
  var map = buildKitMap();
  return map.get(String(kitSku).trim()) || null;
}


// =======================================================================================
// IMPORT FROM ZOHO CSV
// =======================================================================================
//
// Refresh path for v1: user manually exports CSV from Zoho, uploads to Drive,
// runs this function from the Apps Script editor with the Drive file ID:
//
//   importKitsFromZohoCsv("1aB2cD3eF4gH5iJ6kL7mN8oP")
//
// A sidebar "Re-import" button is parked for v2 (file upload via HtmlService
// is a real build, not just a button). For v1 the editor path is acceptable
// because the registry refresh cadence is low (kits don't change daily).

/**
 * Convenience wrapper — runs `importKitsFromZohoCsv` with the current Zoho
 * export's Drive file ID. Bumps the file ID here whenever you re-upload.
 *
 * Why a wrapper: the Apps Script editor's Run dropdown can only invoke
 * parameterless functions. Calling importKitsFromZohoCsv('<id>') directly
 * requires editing the function-call line and remembering quotes. This
 * wrapper lets you just hit Run.
 */
function importKitsNow() {
  var result = importKitsFromZohoCsv('13TyKalS1LOj0CDa_OdpjXKr6L_Jy1R2f');
  // Surface the return value in the Execution Log so the user can SEE the
  // outcome. Without this, Apps Script's "Execution completed" message hides
  // whether the import succeeded (with summary) or failed (with reason).
  console.log(result);
  return result;
}


/**
 * Reads the given Drive CSV file, parses Zoho item data, populates the
 * Kit Registry. Idempotent — clears existing data rows before writing.
 *
 * Returns a status string with the import summary.
 *
 * @param {string} driveFileId — Google Drive file ID of the Zoho CSV export
 */
function importKitsFromZohoCsv(driveFileId) {
  if (!driveFileId) {
    return "❌ Pass a Drive file ID: importKitsFromZohoCsv('1aB2cD...')";
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  if (!sheet) {
    setupKitRegistrySheet();
    sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  }

  // --- Step 1: read the CSV ---
  var csvText;
  try {
    var file = DriveApp.getFileById(driveFileId);
    csvText = file.getBlob().getDataAsString("UTF-8");
  } catch (err) {
    return "❌ Failed to read Drive file " + driveFileId + ": " + err;
  }

  // --- Step 2: parse the CSV ---
  // Use Utilities.parseCsv for robust quoted-field handling.
  // Strip BOM if present (Zoho exports include UTF-8 BOM).
  if (csvText.charCodeAt(0) === 0xFEFF) csvText = csvText.substring(1);
  var rows;
  try {
    rows = Utilities.parseCsv(csvText);
  } catch (err) {
    return "❌ CSV parse failed: " + err;
  }
  if (!rows || rows.length < 2) return "❌ CSV has no data rows.";

  var headers = rows[0];
  var col = {};
  for (var h = 0; h < headers.length; h++) {
    col[String(headers[h]).trim()] = h;
  }

  // Required column lookups — fail loud if Zoho's schema drifts
  var skuCol  = col["SKU"];
  var nameCol = col["Item Name"];
  var pdCol   = col["Purchase Description"];
  if (skuCol == null || nameCol == null || pdCol == null) {
    return "❌ CSV missing required columns. Need: SKU, Item Name, Purchase Description";
  }
  // Sales Description is OPTIONAL — older CSV exports may not have it. If
  // missing, SALES_DESC col is left blank and the modal's reference panel
  // will show "(no Sales Description available)".
  var sdCol = col["Sales Description"];

  // --- Step 3: build MI map (SKU → location/qty) for cross-resolution + KIT_TYPE ---
  var miMap = _buildMasterInventoryMap();

  // --- Step 4: iterate rows, identify kits, parse components ---
  //
  // KIT NAME PATTERN
  //   Catches "Engine Overhaul" / "Major Overhaul" / "Overhaul Kit" /
  //   "Rebuild Kit" / "Repair Kit". "Repair Kit" added 2026-05-15 after
  //   SKU 215153 ("Engine repair kit 0.50") didn't match the original regex.
  //   Side-effect: ~20 single-package "Oil Cooler Repair Kit" / "Fuel Pump
  //   Repair Kit" items now also match, but they have empty/non-parseable PD
  //   so they get correctly skipped at the component-parsing step.
  var kitNamePattern = /(engine\s+overhaul|major\s+overhaul|overhaul\s+kit|rebuild\s+kit|repair\s+kit)/i;

  // COMPONENT LINE PATTERN
  //   ^[\s'\-]*           leading dash/apostrophe/whitespace (Zoho exports vary)
  //   (\d+)               quantity
  //   [\s\-]+             separator
  //   (.+?)               component name (lazy)
  //   \s+                 space before SKU
  //   (\d{6})             6-digit component SKU
  //   (?:\s*\([^)]*\))?   optional parenthetical suffix like "(Deleted)"
  //   \s*$                trailing whitespace
  //
  // The "(Deleted)" tolerance was added 2026-05-15. Without it, 8 component
  // lines across 6 kits silently dropped. The deleted marker itself is
  // discarded — the component IS registered so the preview surfaces it as
  // "NOT FOUND" in MI, giving the picker the explicit "this kit has a
  // deleted component that needs a substitute" signal.
  var componentRegex = /^[\s'\-]*(\d+)[\s\-]+(.+?)\s+(\d{6})(?:\s*\([^)]*\))?\s*$/;

  var registryRows = [];
  var stats = {
    kitsScanned:       0,
    kitsWithData:      0,
    kitsSkippedEmpty:  0,
    kitsSkippedUnreal: 0,
    componentLines:    0,
    componentMissing:  0,
    readyKits:         0,
    manualKits:        0
  };

  var now = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd HH:mm");

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var name = String(row[nameCol] || "");
    if (!kitNamePattern.test(name)) continue;
    stats.kitsScanned++;

    var kitSku = String(row[skuCol] || "").trim();
    if (!kitSku) { stats.kitsSkippedEmpty++; continue; }
    if (/^TEMP[-\s]?\d+$/i.test(kitSku)) { stats.kitsSkippedUnreal++; continue; }

    var pd = String(row[pdCol] || "").trim();
    if (!pd) { stats.kitsSkippedEmpty++; continue; }
    var sd = (sdCol != null) ? String(row[sdCol] || "").trim() : "";

    // First non-empty line of PD = engine model
    var pdLines = pd.split(/\r?\n/);
    var engineModel = "";
    var componentLines = [];
    for (var j = 0; j < pdLines.length; j++) {
      var line = pdLines[j].trim();
      if (!line) continue;
      // Heuristic: lines starting with -, ', or a digit followed by a dash are components
      // Anything else on the first non-empty position is the engine model
      if (!engineModel && !/^[\s'\-]*\d+[\s\-]/.test(line)) {
        engineModel = line;
        continue;
      }
      componentLines.push(line);
    }

    // Parse components
    var components = [];
    for (var k = 0; k < componentLines.length; k++) {
      var m = componentRegex.exec(componentLines[k]);
      if (!m) continue;
      var compQty  = parseInt(m[1]) || 1;
      var compName = m[2].trim();
      var compSku  = m[3];
      components.push({ sku: compSku, qty: compQty, name: compName });
      stats.componentLines++;
      if (!miMap[compSku]) stats.componentMissing++;
    }
    if (components.length === 0) { stats.kitsSkippedEmpty++; continue; }
    stats.kitsWithData++;

    // Derive KIT_TYPE from kit's own MI location
    var kitMi = miMap[kitSku];
    var kitLocation = kitMi ? String(kitMi.location || "") : "";
    var kitType = (/^K[-\s]/i.test(kitLocation)) ? KIT_REGISTRY.types.READY
                                                  : KIT_REGISTRY.types.MANUAL;
    if (kitType === KIT_REGISTRY.types.READY) stats.readyKits++;
    else stats.manualKits++;

    // One row per (kit, component) pair — denormalized
    for (var c = 0; c < components.length; c++) {
      registryRows.push([
        kitSku,
        name.replace(/ /g, " "),  // Zoho uses non-breaking spaces, normalize
        kitType,
        kitLocation,
        engineModel,
        components[c].sku,
        components[c].qty,
        components[c].name,
        sd,
        now
      ]);
    }
  }

  // --- Step 5: write rows (clear existing first) ---
  var lastRow = sheet.getLastRow();
  if (lastRow >= KIT_REGISTRY.dataStartRow) {
    sheet.getRange(KIT_REGISTRY.dataStartRow, 1,
                   lastRow - KIT_REGISTRY.dataStartRow + 1,
                   KIT_REGISTRY.dataWidth).clearContent();
  }

  if (registryRows.length > 0) {
    sheet.getRange(KIT_REGISTRY.dataStartRow, 1, registryRows.length, KIT_REGISTRY.dataWidth)
         .setValues(registryRows);
  }

  return "✅ Imported " + stats.kitsWithData + " kits ("
       + stats.manualKits + " MANUAL, " + stats.readyKits + " READY) · "
       + registryRows.length + " component rows · "
       + stats.componentMissing + " components not in MI · "
       + (stats.kitsSkippedUnreal + stats.kitsSkippedEmpty) + " kits skipped (TEMP/empty)";
}


// =======================================================================================
// LIVE WEBHOOK: refresh ONE kit from a Zoho Item payload
// =======================================================================================
//
// Companion to importKitsFromZohoCsv — same parser logic, applied to a single
// item. Called by doPost when Zoho's Workflow Rule (Item.Purchase Description
// updated) fires with action=zohoKitUpdated.
//
// UPSERT SEMANTICS
//   This kit's existing rows in Kit Registry are DELETED, then new rows are
//   written based on the current PD content. If parsing yields no components
//   (PD cleared, name no longer kit-like, TEMP-* SKU), rows are removed but
//   NOT replaced — keeping the registry in sync with Zoho's reality.
//
// @param {object} zohoItem  — Zoho payload's `item` field. Expected keys:
//                              item_id, sku, name, purchase_description,
//                              description (= sales description per items.yml
//                              line 409 — Zoho's API names the sales-side
//                              field plain `description`, not `sales_description`).
//                              Other fields ignored. Custom fields not used —
//                              KIT_TYPE is still derived from MI location to
//                              match the CSV importer's logic exactly.
//
// @returns {{ kitSku, actionTaken, componentsWritten, reason }}
//   actionTaken: "added" | "updated" | "removed" | "skipped"

function refreshOneKitFromZohoPayload(zohoItem) {
  if (!zohoItem || typeof zohoItem !== 'object') {
    return { kitSku: "", actionTaken: "skipped", componentsWritten: 0, reason: "Empty payload" };
  }

  var kitSku = String(zohoItem.sku || "").trim();
  var name   = String(zohoItem.name || "");
  var pd     = String(zohoItem.purchase_description || "").trim();
  var sd     = String(zohoItem.description || "").trim();   // Sales Description

  if (!kitSku) {
    return { kitSku: "", actionTaken: "skipped", componentsWritten: 0,
             reason: "No SKU in payload" };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(KIT_REGISTRY.sheetName);
  if (!sheet) {
    return { kitSku: kitSku, actionTaken: "skipped", componentsWritten: 0,
             reason: "Kit Registry sheet doesn't exist — run setupKitRegistrySheet first" };
  }

  // Locate any existing rows for this kit BEFORE deciding what to do
  var existingRows = _findKitRegistryRows(sheet, kitSku);
  var hadExisting  = existingRows.length > 0;

  // --- Validate kit-name + SKU + PD (same checks as the bulk CSV importer) ---
  var kitNamePattern = /(engine\s+overhaul|major\s+overhaul|overhaul\s+kit|rebuild\s+kit|repair\s+kit)/i;
  var tempPattern    = /^TEMP[-\s]?\d+$/i;

  if (tempPattern.test(kitSku)) {
    if (hadExisting) _deleteKitRegistryRows(sheet, existingRows);
    return { kitSku: kitSku, actionTaken: hadExisting ? "removed" : "skipped",
             componentsWritten: 0, reason: "TEMP-* placeholder SKU" };
  }
  if (!kitNamePattern.test(name)) {
    if (hadExisting) _deleteKitRegistryRows(sheet, existingRows);
    return { kitSku: kitSku, actionTaken: hadExisting ? "removed" : "skipped",
             componentsWritten: 0, reason: "Item name no longer matches kit pattern" };
  }
  if (!pd) {
    if (hadExisting) _deleteKitRegistryRows(sheet, existingRows);
    return { kitSku: kitSku, actionTaken: hadExisting ? "removed" : "skipped",
             componentsWritten: 0, reason: "Purchase Description is empty" };
  }

  // --- Parse PD: first non-component line = engine model, rest = components ---
  // Identical regex to importKitsFromZohoCsv (including 2026-05-15 fixes for
  // "repair kit" pattern + "(Deleted)" suffix tolerance).
  var componentRegex = /^[\s'\-]*(\d+)[\s\-]+(.+?)\s+(\d{6})(?:\s*\([^)]*\))?\s*$/;
  var pdLines        = pd.split(/\r?\n/);
  var engineModel    = "";
  var components     = [];

  for (var i = 0; i < pdLines.length; i++) {
    var line = pdLines[i].trim();
    if (!line) continue;
    if (!engineModel && !/^[\s'\-]*\d+[\s\-]/.test(line)) {
      engineModel = line;
      continue;
    }
    var m = componentRegex.exec(line);
    if (!m) continue;
    components.push({
      qty:  parseInt(m[1]) || 1,
      name: m[2].trim(),
      sku:  m[3]
    });
  }

  if (components.length === 0) {
    if (hadExisting) _deleteKitRegistryRows(sheet, existingRows);
    return { kitSku: kitSku, actionTaken: hadExisting ? "removed" : "skipped",
             componentsWritten: 0, reason: "No parseable component lines in Purchase Description" };
  }

  // --- KIT_TYPE from MI location (same as CSV importer) ---
  var miMap       = _buildMasterInventoryMap();
  var kitMi       = miMap[kitSku];
  var kitLocation = (kitMi && kitMi.location != null) ? String(kitMi.location) : "";
  var kitType     = (/^K[-\s]/i.test(kitLocation)) ? KIT_REGISTRY.types.READY
                                                    : KIT_REGISTRY.types.MANUAL;

  // Zoho exports embed non-breaking spaces in some item names; normalize
  var nameNorm = name.replace(/ /g, " ");
  var now      = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd HH:mm");

  var newRows = components.map(function(c) {
    return [
      kitSku, nameNorm, kitType, kitLocation, engineModel,
      c.sku, c.qty, c.name, sd, now
    ];
  });

  // --- Upsert: delete existing rows, then append new ones ---
  if (hadExisting) _deleteKitRegistryRows(sheet, existingRows);

  var appendStart = sheet.getLastRow() + 1;
  sheet.getRange(appendStart, 1, newRows.length, KIT_REGISTRY.dataWidth).setValues(newRows);
  SpreadsheetApp.flush();

  return {
    kitSku:            kitSku,
    actionTaken:       hadExisting ? "updated" : "added",
    componentsWritten: newRows.length,
    reason:            ""
  };
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/**
 * Returns 1-based sheet row numbers in Kit Registry whose KIT_SKU matches.
 * O(n) scan of the SKU column; cheap on ~1500 rows.
 */
function _findKitRegistryRows(sheet, kitSku) {
  var lastRow = sheet.getLastRow();
  if (lastRow < KIT_REGISTRY.dataStartRow) return [];

  var skus = sheet.getRange(
    KIT_REGISTRY.dataStartRow, KIT_REGISTRY.cols.KIT_SKU,
    lastRow - KIT_REGISTRY.dataStartRow + 1, 1
  ).getValues();

  var target = String(kitSku).trim();
  var matches = [];
  for (var i = 0; i < skus.length; i++) {
    if (String(skus[i][0]).trim() === target) {
      matches.push(KIT_REGISTRY.dataStartRow + i);
    }
  }
  return matches;
}

/**
 * Deletes rows by 1-based number. Handles contiguous + non-contiguous cases
 * via descending iteration (so each deletion doesn't invalidate the next
 * row number).
 */
function _deleteKitRegistryRows(sheet, rowNumbers) {
  rowNumbers
    .slice()
    .sort(function(a, b) { return b - a; })   // descending
    .forEach(function(rn) { sheet.deleteRow(rn); });
}

/**
 * Builds {sku → {location, qty, title}} map from Master Inventory.
 *
 * IMPORTANT: MI stores SKUs as floats in some rows (e.g., 161361.0 vs '161361').
 * Normalize to string-of-integer for reliable matching against the string SKUs
 * Zoho exports.
 *
 * Reads only the 4 columns we need (skip the 195-col-wide MI scan) — keeps
 * the read fast and memory tight.
 */
function _buildMasterInventoryMap() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var miSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!miSheet) return {};

  var lastRow = miSheet.getLastRow();
  if (lastRow < 2) return {};

  // We need: col 2 (sku), col 6 (quantity), col 3 (title), col 40 (C:Model Year = aisle).
  // Read col 2-40 in one batch; cheaper than 4 separate reads on large MI.
  var data = miSheet.getRange(2, 2, lastRow - 1, 39).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var raw = data[i][0];
    if (raw === "" || raw == null) continue;
    var skuStr;
    if (typeof raw === "number") {
      skuStr = String(Math.trunc(raw));
    } else {
      skuStr = String(raw).trim();
      // Sometimes Sheets stores numeric SKUs with a trailing .0
      if (/^\d+\.0+$/.test(skuStr)) skuStr = skuStr.replace(/\.0+$/, "");
    }
    if (!skuStr) continue;
    map[skuStr] = {
      title:    data[i][1],     // col 3
      qty:      data[i][4],     // col 6 (offset 4 within col-2-to-40 slice)
      location: data[i][38]     // col 40 (offset 38 within slice) - C:Model Year = aisle
    };
  }
  return map;
}
