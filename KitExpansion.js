// =======================================================================================
// KIT_EXPANSION.gs — Sidebar-driven expansion of repair-kit rows
// =======================================================================================
//
// COMPANION TO KitRegistry.js. The registry is the data layer; this file is
// the action layer. Picker selects kit rows on the sheet, clicks Preview in
// the sidebar, reviews in a popup modal, commits the expansion.
//
// TWO COMMIT PATHS coexist (both production):
//
//   1. MODAL PATH (primary, shipped 2026-05-19) — Preview opens a popup
//      modal showing the components checklist (left) + Sales Description
//      reference panel (right). Picker unchecks components that are
//      physically packed together (Sales Description shows packaging hints
//      like "Full Gasket Set ABC + Head Gasket" — Head Gasket is bundled
//      inside ABC's box, doesn't need its own row). Expand-and-continue
//      commits the kit minus excluded SKUs, advances to next kit in queue.
//
//   2. BULK PATH (fallback, predates the modal) — sidebar Expand button
//      commits ALL selected kits with ALL components, no exclusion controls.
//      Kept during the modal's bake-in period; will be removed once the
//      modal is proven in production.
//
// FLOW (modal path)
//   1. Picker selects N rows in All Orders, enters optional Deploy×
//   2. Picker clicks Preview → openKitExpansionModal reads selection,
//      builds queue, caches it, opens the modal with first kit's data
//   3. Modal renders kit metadata + per-component checkboxes + Sales
//      Description reference text
//   4. Picker reviews, unchecks bundled components, clicks Expand
//   5. commitKitFromModal commits this one kit, returns next queue entry
//   6. Modal swaps body to next kit, repeats until queue exhausted
//   7. Done screen summarizes committed / skipped / failed
//
// PUBLIC API
//   expandKit(kitSku, deployQty)                — pure: kit info + scaled components
//   previewSelectedKits(deployQty)              — read-only: enriched preview payload
//   expandSelectedKits(deployQty, exclusionMap) — BULK COMMIT (fallback path)
//   openKitExpansionModal(deployQty)            — MODAL: entry point from sidebar
//   commitKitFromModal(sessionId, excludedSkus) — MODAL: commit current + advance
//   skipKitFromModal(sessionId)                 — MODAL: skip current + advance
//   closeKitExpansionSession(sessionId)         — MODAL: cleanup on close
// =======================================================================================


// =======================================================================================
// PUBLIC: expandKit(kitSku, deployQty)
// =======================================================================================

/**
 * Pure function over the Kit Registry. Returns the expansion plan for one
 * kit at a given deploy multiplier — does NOT touch the sheet.
 *
 * @param {string} kitSku       — kit's SKU (e.g., "160029")
 * @param {number} deployQty    — multiplier (kit_component_qty * deployQty);
 *                                defaults to 1 if missing/invalid
 * @returns {{
 *   found: boolean,
 *   reason: string,            — only set when found=false
 *   kitSku, kitName, kitType, kitLocation, kitEngine, salesDescription,
 *   deployQty: number,
 *   components: Array<{sku, qty, name}>
 * }}
 *
 * Returns found=false in two cases:
 *   - Kit not in registry (caller decides whether to surface "unknown kit" or
 *     fall through to treating it as a non-kit row)
 *   - Kit is registered but has no components (shouldn't happen; the importer
 *     skips kits with zero parsed components — but defensive)
 *
 * For READY kits, this still returns found=true with the component list. The
 * caller (preview UI / commit function) decides what to do with READY kits;
 * typically: show the components as informational but refuse to commit.
 */
function expandKit(kitSku, deployQty) {
  var sku = String(kitSku == null ? "" : kitSku).trim();
  if (!sku) {
    return { found: false, reason: "Empty SKU", kitSku: sku };
  }

  var kit = getKitInfo(sku);
  if (!kit) {
    return { found: false, reason: "Kit SKU not in registry", kitSku: sku };
  }
  if (!kit.components || kit.components.length === 0) {
    return { found: false, reason: "Kit has no components in registry", kitSku: sku };
  }

  var dq = parseInt(deployQty);
  if (!dq || dq < 1) dq = 1;

  // Scale each component qty by the deploy multiplier
  var components = kit.components.map(function(c) {
    return {
      sku:  c.sku,
      name: c.name,
      qty:  (parseInt(c.qty) || 1) * dq
    };
  });

  return {
    found:            true,
    kitSku:           kit.sku,
    kitName:          kit.name,
    kitType:          kit.type,
    kitLocation:      kit.location,
    kitEngine:        kit.engine,
    salesDescription: kit.salesDescription || "",
    deployQty:        dq,
    components:       components
  };
}


// =======================================================================================
// PUBLIC: previewSelectedKits(deployQty)
// =======================================================================================

/**
 * Sidebar entry point. Reads the picker's current selection on the All Orders
 * sheet, classifies each selected row as kit-or-not, builds a preview payload
 * the sidebar can render. READ-ONLY — no sheet writes.
 *
 * Selection-based (not row-number-based) because picker selects on the sheet
 * naturally; asking them to type row numbers in the sidebar is friction.
 *
 * @param {number} deployQty — multiplier applied to all selected kit rows.
 *                              Defaults to 1 (deploy exactly the row qty).
 * @returns {{
 *   ok: boolean,
 *   message: string,
 *   selectedCount: number,
 *   kitsFound: number,
 *   nonKitRows: Array<{row, sku, reason}>,
 *   kitRows: Array<{
 *     row, table, sourceSku, sourceQty, sourceSalesOrder, sourceNote,
 *     plan: {kitName, kitType, kitLocation, kitEngine, deployQty,
 *            components: [{sku, qty, name, location, available, missing}]}
 *   }>
 * }}
 */
function previewSelectedKits(deployQty) {
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    return { ok: false, message: "No active spreadsheet.", selectedCount: 0,
             kitsFound: 0, nonKitRows: [], kitRows: [] };
  }

  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    return { ok: false, message: "All Orders sheet not found.", selectedCount: 0,
             kitsFound: 0, nonKitRows: [], kitRows: [] };
  }

  // --- Collect distinct row numbers from the selection ---
  var selectedRows = _collectSelectedRows(sheet);
  if (selectedRows.length === 0) {
    return { ok: false, message: "No rows selected on the sheet.", selectedCount: 0,
             kitsFound: 0, nonKitRows: [], kitRows: [], debug: { selectedRows: [] } };
  }

  var boundaryRow = getBoundaryRow();   // for eBay vs DIRECT labeling

  // --- Batch-read the data band so we can pull each selected row's fields ---
  var lastRow = sheet.getLastRow();
  var allData = sheet.getRange(1, 1, lastRow, Schema.dataWidth).getValues();

  // --- Build the location + inventory maps ONCE so component enrichment is
  // O(1) per lookup. These are the same maps used by doPost / LiveSync. ---
  var locInvMaps = buildLocationAndInventoryMaps();   // { locationMap, inventoryMap }
  // Kit components are DIRECT-side items — many aren't listed on eBay, so MI
  // reads null/stale. Zoho is the authoritative stock source for them (same
  // routing as recomputeHand's DIRECT rows). Enrich availability Zoho-first.
  var zohoMap = buildZohoStockMap();

  var kitRows = [];
  var nonKitRows = [];

  var SKU_I = Schema.idx("SKU");
  var QTY_I = Schema.idx("QTY");
  var SO_I  = Schema.idx("SALES_ORDER");
  var NOTE_I = Schema.idx("NOTE");

  for (var i = 0; i < selectedRows.length; i++) {
    var sheetRow = selectedRows[i];

    // EXPLICIT rejection for non-data rows — surface in nonKitRows with a
    // reason rather than silently dropping. This is essential for debugging
    // "why didn't my row show up?"
    if (sheetRow < Schema.dataStartRow) {
      nonKitRows.push({ row: sheetRow, sku: "", reason: "Banner/header row (rows 1-3)" });
      continue;
    }
    if (boundaryRow > 0 && sheetRow === boundaryRow) {
      nonKitRows.push({ row: sheetRow, sku: "", reason: "DIRECT divider row" });
      continue;
    }
    if (boundaryRow > 0 && sheetRow === boundaryRow + 1) {
      nonKitRows.push({ row: sheetRow, sku: "", reason: "DIRECT table header row" });
      continue;
    }

    var rowVals = allData[sheetRow - 1];   // 0-based index into the read buffer
    if (!rowVals) {
      nonKitRows.push({ row: sheetRow, sku: "", reason: "Row index out of read buffer (past last data row)" });
      continue;
    }

    var rowSku = String(rowVals[SKU_I] || "").trim();
    var rowQty = parseInt(rowVals[QTY_I]) || 1;
    var rowSo  = String(rowVals[SO_I]  || "");
    var rowNote = String(rowVals[NOTE_I] || "");

    if (!rowSku) {
      nonKitRows.push({ row: sheetRow, sku: "", reason: "Empty SKU on this row" });
      continue;
    }

    // Effective deploy = user multiplier × this row's order qty
    var effectiveDeploy = (parseInt(deployQty) || 1) * rowQty;
    var plan = expandKit(rowSku, effectiveDeploy);

    if (!plan.found) {
      nonKitRows.push({ row: sheetRow, sku: rowSku, reason: plan.reason });
      continue;
    }

    // Enrich each component with current LOCATION (MI) + AVAILABLE (Zoho-first,
    // MI fallback). null if neither source has the SKU → UI shows "unknown".
    var enriched = plan.components.map(function(c) {
      var skuLower = c.sku.toLowerCase();
      var loc = locInvMaps.locationMap.get(skuLower) || "NOT FOUND";
      var inv = locInvMaps.inventoryMap.get(skuLower);
      var zo  = zohoMap.get(skuLower);
      var miAvail = (inv && inv.available != null) ? inv.available : null;
      var zoAvail = zo ? zo.available : null;
      var available = (zoAvail != null) ? zoAvail : miAvail;
      return {
        sku:       c.sku,
        qty:       c.qty,
        name:      c.name,
        location:  loc,
        available: available,
        missing:   (loc === "NOT FOUND")
      };
    });

    var table = (boundaryRow > 0 && sheetRow > boundaryRow) ? "DIRECT" : "eBay";

    kitRows.push({
      row:              sheetRow,
      table:            table,
      sourceSku:        rowSku,
      sourceQty:        rowQty,
      sourceSalesOrder: rowSo,
      sourceNote:       rowNote,
      plan: {
        kitName:          plan.kitName,
        kitType:          plan.kitType,
        kitLocation:      plan.kitLocation,
        kitEngine:        plan.kitEngine,
        salesDescription: plan.salesDescription,
        deployQty:        plan.deployQty,
        components:       enriched
      }
    });
  }

  return {
    ok:            true,
    message:       "Preview ready",
    selectedCount: selectedRows.length,
    kitsFound:     kitRows.length,
    nonKitRows:    nonKitRows,
    kitRows:       kitRows,
    debug: {
      selectedRows:    selectedRows,
      boundaryRow:     boundaryRow,
      dataStartRow:    Schema.dataStartRow,
      lastDataRow:     lastRow,
      activeSheetName: ss.getActiveSheet().getName()
    }
  };
}


// =======================================================================================
// PUBLIC: expandSelectedKits(deployQty)
// =======================================================================================
//
// COMMIT path. Reads sheet selection, inserts component rows below each MANUAL
// kit row in the selection. Sibling of previewSelectedKits — same selection
// logic, but actually writes.
//
// CONTRACT
//   - LockService serializes commits (30s wait)
//   - Original kit row STAYS. Only its position is the anchor — components are
//     inserted directly below it via sheet.insertRowsAfter(kitRow, N).
//   - Component rows inherit: SALES_ORDER, STATUS, SHIPPING (the picker-facing
//     identifiers), buyer NOTE merged into the kit-expansion tag, and current
//     MI location + hand.
//   - SHIP_COST stays on the kit row only; not duplicated onto components
//     (preserves single-line-item shipping accounting).
//   - LEFT column stays blank — picker fills it after counting at the shelf.
//   - Activity Log: one RECEIVED event per component row, source=sidebar.
//   - READY kits are REFUSED (return refused[] entry) — they ship as one box.
//   - Already-expanded kits are REFUSED — defense: check the row immediately
//     below; if its NOTE starts with "↳ from KIT-<sku>" for this same kit,
//     skip to prevent double-expansion when the picker clicks twice.
//   - Filter-corruption defense (gotcha #4): save headers before each insert,
//     verify-and-restore after. Multi-kit batches use bottom-up processing
//     order so kit row numbers stay valid across inserts.
//
// @param {number} extrasQty   — ADDITIVE extras-for-us count, applied to
//                                EACH selected kit row. 0 (default) = ship
//                                exactly the customer's rowQty (no spares).
//                                N = ship rowQty + N total kits per row.
//                                Same semantic as the modal's per-kit input
//                                (unified 2026-05-20). Old "deployQty=N means
//                                multiply rowQty by N" semantic was retired
//                                because it broke at rowQty>1.
// @param {object} [exclusionMap] — optional { "<kitSku>": ["<compSku1>", ...] }
//                                map of components to SKIP per kit. Provided by
//                                the Kit Expansion modal when the picker
//                                unchecks components that are physically packed
//                                together (Sales Description bundling hint).
//                                Missing kits default to no exclusions.
// @returns {{
//   ok: boolean, message: string,
//   expanded: number,
//   refused: Array<{row, sku, reason}>,
//   skipped: Array<{row, reason}>,
//   details: Array<{row, sku, componentsAdded, excludedSkus, extras, totalKits, rowQty}>
// }}
function expandSelectedKits(extrasQty, exclusionMap) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: "Another operation is in progress. Try again.",
             expanded: 0, refused: [], skipped: [], details: [] };
  }

  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) {
      return { ok: false, message: "No active spreadsheet.",
               expanded: 0, refused: [], skipped: [], details: [] };
    }
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) {
      return { ok: false, message: "All Orders sheet not found.",
               expanded: 0, refused: [], skipped: [], details: [] };
    }

    var selectedRows = _collectSelectedRows(sheet);
    if (selectedRows.length === 0) {
      return { ok: false, message: "No rows selected on the sheet.",
               expanded: 0, refused: [], skipped: [], details: [] };
    }

    // Additive extras-for-us count (replaces old row-qty multiplier)
    var extras = parseInt(extrasQty);
    if (isNaN(extras) || extras < 0) extras = 0;
    if (extras > 99) extras = 99;

    var boundaryRow = getBoundaryRow();
    var locInvMaps = buildLocationAndInventoryMaps();
    var zohoMap = buildZohoStockMap();   // DIRECT-side HAND source (Zoho-first)

    // Save headers ONCE upfront — verifyAndRestoreHeaders defends against the
    // Sheets bug where inserting rows inside a filtered area replaces headers
    // with "Column 1", "Column 2", ... (CLAUDE.md gotcha #4).
    var savedHeaders = sheet.getRange(Schema.headerRow, 1, 1, Schema.dataWidth).getValues()[0];

    // Process BOTTOM-UP — insertions shift rows BELOW the insertion point. By
    // working from the highest row number downward, the row numbers we haven't
    // processed yet stay valid.
    selectedRows.sort(function(a, b) { return b - a; });

    var SKU_I        = Schema.idx("SKU");
    var QTY_I        = Schema.idx("QTY");
    var SO_I         = Schema.idx("SALES_ORDER");
    var NOTE_I       = Schema.idx("NOTE");
    var STATUS_I     = Schema.idx("STATUS");
    var SHIPPING_I   = Schema.idx("SHIPPING");

    var expanded = 0;
    var refused = [];
    var skipped = [];
    var details = [];
    var activityLog = [];

    for (var i = 0; i < selectedRows.length; i++) {
      var kitRow = selectedRows[i];

      // Skip non-data rows
      if (kitRow < Schema.dataStartRow) {
        skipped.push({ row: kitRow, reason: "Banner/header row" });
        continue;
      }
      if (boundaryRow > 0 && (kitRow === boundaryRow || kitRow === boundaryRow + 1)) {
        skipped.push({ row: kitRow, reason: "DIRECT divider/header row" });
        continue;
      }

      // Re-read the kit row freshly to catch concurrent edits
      var rowVals = sheet.getRange(kitRow, 1, 1, Schema.dataWidth).getValues()[0];
      var rowSku    = String(rowVals[SKU_I]    || "").trim();
      var rowQty    = parseInt(rowVals[QTY_I]) || 1;
      var rowSo     = String(rowVals[SO_I]     || "");
      var rowNote   = String(rowVals[NOTE_I]   || "");
      var rowStatus = String(rowVals[STATUS_I] || Schema.status.PENDING);
      var rowShip   = String(rowVals[SHIPPING_I] || "");

      if (!rowSku) {
        skipped.push({ row: kitRow, reason: "Empty SKU on this row" });
        continue;
      }

      // Already-expanded check: if the row directly below has a NOTE that
      // starts with "↳ from KIT-<rowSku>", this kit was already expanded.
      // Defensive against double-clicks or repeated commits on stale selection.
      if (kitRow + 1 <= sheet.getLastRow()) {
        var belowNote = String(sheet.getRange(kitRow + 1, NOTE_I + 1).getValue() || "");
        if (belowNote.indexOf("↳ from KIT-" + rowSku) === 0) {
          refused.push({
            row: kitRow, sku: rowSku,
            reason: "Already expanded (row below has matching kit tag)"
          });
          continue;
        }
      }

      // Additive math: total kits = rowQty + extras
      var totalKits       = rowQty + extras;
      if (totalKits < 1) {
        skipped.push({ row: kitRow, reason: "rowQty=" + rowQty + " + extras=" + extras + " yields 0 kits" });
        continue;
      }
      var effectiveDeploy = totalKits;
      var plan = expandKit(rowSku, effectiveDeploy);

      if (!plan.found) {
        refused.push({ row: kitRow, sku: rowSku, reason: plan.reason });
        continue;
      }
      if (plan.kitType === KIT_REGISTRY.types.READY) {
        refused.push({
          row: kitRow, sku: rowSku,
          reason: "READY kit (lives at " + (plan.kitLocation || "K-*") + ") — ships pre-assembled, no expansion"
        });
        continue;
      }

      // --- Apply per-kit exclusions from the modal (if any) ---
      // exclusionMap shape: { "<kitSku>": ["<compSku1>", "<compSku2>"] }
      // Components matching an excluded SKU are filtered out BEFORE insert.
      // Excluded SKUs are recorded for the Activity Log DETAIL field so the
      // audit trail captures "what didn't get inserted and why".
      var excludedSkus = [];
      if (exclusionMap && exclusionMap[rowSku] && exclusionMap[rowSku].length) {
        var excludeSet = {};
        for (var ex = 0; ex < exclusionMap[rowSku].length; ex++) {
          excludeSet[String(exclusionMap[rowSku][ex]).trim()] = true;
        }
        var keptComponents = [];
        for (var fc = 0; fc < plan.components.length; fc++) {
          var compSku = String(plan.components[fc].sku).trim();
          if (excludeSet[compSku]) {
            excludedSkus.push(compSku);
          } else {
            keptComponents.push(plan.components[fc]);
          }
        }
        plan.components = keptComponents;
      }

      var N = plan.components.length;
      if (N === 0) {
        var emptyReason = excludedSkus.length > 0
          ? "All components excluded by picker (" + excludedSkus.join(", ") + ")"
          : "Kit has no components in registry";
        refused.push({ row: kitRow, sku: rowSku, reason: emptyReason });
        continue;
      }

      // Build the kit-expansion NOTE prefix
      // - Default (extras=0): "↳ from KIT-160029"
      // - extras>0: "↳ from KIT-160029 · deploy 3 total (1 for customer + 2 for us)"
      // - Then merge in the kit row's original NOTE (which may contain a
      //   "Buyer Note: ..." prefix from n8n, or supervisor remarks) — picker
      //   sees both signals on each component row.
      var notePrefix = "↳ from KIT-" + rowSku;
      if (extras > 0) {
        notePrefix += " · deploy " + totalKits + " total ("
                    + rowQty + " for customer + " + extras + " for us)";
      }
      var componentNote = rowNote ? (notePrefix + " · " + rowNote) : notePrefix;

      // Insert N blank rows below the kit row
      sheet.insertRowsAfter(kitRow, N);

      // Build component row values
      var newRows = [];
      for (var c = 0; c < N; c++) {
        var comp = plan.components[c];
        var skuLower = comp.sku.toLowerCase();
        var loc = locInvMaps.locationMap.get(skuLower) || "NOT FOUND";
        var inv = locInvMaps.inventoryMap.get(skuLower);
        var zo  = zohoMap.get(skuLower);
        // DIRECT-side rows take HAND Zoho-first, MI fallback (matches recomputeHand).
        // "" when neither source has it — the next recompute resolves it.
        var hand = zo ? zo.available
                 : (inv && inv.available != null) ? inv.available : "";

        var row = new Array(Schema.dataWidth);
        row[Schema.idx("SKU")]         = comp.sku;
        row[Schema.idx("QTY")]         = comp.qty;
        row[Schema.idx("LOCATION")]    = loc;
        row[Schema.idx("SALES_ORDER")] = rowSo;
        row[Schema.idx("NOTE")]        = componentNote;
        row[Schema.idx("STATUS")]      = rowStatus;
        row[Schema.idx("HAND")]        = hand;
        row[Schema.idx("LEFT")]        = "";
        row[Schema.idx("SHIPPING")]    = rowShip;
        row[Schema.idx("SHIP_COST")]   = "";   // stays on parent kit row only
        newRows.push(row);
      }

      sheet.getRange(kitRow + 1, 1, N, Schema.dataWidth).setValues(newRows);
      verifyAndRestoreHeaders(sheet, savedHeaders);

      // One Activity Log RECEIVED entry per component row.
      // DETAIL field carries kit-expansion context + (if any) the list of
      // SKUs excluded by the picker — so reviewing any component row in the
      // log surfaces "this was inserted as part of KIT-X, and these other
      // SKUs were excluded as bundled-with-another-SKU calls".
      var baseDetail = "kit expansion from " + rowSku;
      if (extras > 0) baseDetail += " (deploy " + totalKits + " total: " + rowQty + "+" + extras + ")";
      if (excludedSkus.length > 0) {
        baseDetail += " · excluded: " + excludedSkus.join(", ");
      }
      for (var c2 = 0; c2 < N; c2++) {
        var comp2 = plan.components[c2];
        activityLog.push([
          "RECEIVED",
          rowSo,                         // orderId
          comp2.sku,                     // sku
          comp2.qty,                     // qty
          "sidebar",                     // source — picker auto-captured from G2
          baseDetail,                    // detail
          undefined,                     // picker: let logActivityBatch resolve via G2
          componentNote                  // note
        ]);
      }

      expanded++;
      details.push({
        row:             kitRow,
        sku:             rowSku,
        componentsAdded: N,
        excludedSkus:    excludedSkus,
        extras:          extras,
        totalKits:       totalKits,
        rowQty:          rowQty
      });

      // Boundary row shifted if we just inserted in the eBay table — recompute
      // for the next iteration. (Bottom-up processing means later iterations
      // are at LOWER row numbers, so they're unaffected — but for safety.)
      boundaryRow = getBoundaryRow();
    }

    // Batch-log activity (best-effort; logging failure does NOT roll back inserts)
    if (activityLog.length > 0) {
      try { logActivityBatch(activityLog); } catch (logErr) {
        try { console.log("expandSelectedKits: activity log failed: " + logErr); } catch (_) {}
      }
    }

    // Refresh Kit SKU markers — programmatic insertRowsAfter + setValues
    // doesn't fire kitSkuOnEdit, so without this explicit call:
    //   (a) the parent kit row keeps its ▣ correctly (untouched by expansion)
    //   (b) but newly-inserted component rows would inherit whatever number
    //       format their template row had (could be ▣ if pasted near a kit
    //       row, or plain if pasted near a non-kit row)
    //   (c) and the suppression rule for "↳ from KIT-" NOTE prefix wouldn't
    //       run, so component rows whose SKU is itself a kit (sub-assemblies)
    //       would show ▣ alongside the parent — exactly the noise the user
    //       flagged 2026-05-19.
    // refreshKitSkuMarkers reads each row's NOTE column and skips components
    // (via the ↳ from KIT- prefix check) regardless of whether their SKU is
    // a standalone kit in the registry.
    if (expanded > 0) {
      try { refreshKitSkuMarkers(); }
      catch (kitErr) { console.log("expandSelectedKits: kit marker refresh failed: " + kitErr); }
      // New component rows share the kit row's SO# — re-paint duplicate-SO
      // borders so the group is surfaced immediately instead of waiting for
      // the next onEdit (which only fires on user typing, not programmatic
      // insertRowsAfter / setValues).
      try { setupDuplicateSalesOrderHighlighting(); }
      catch (dupErr) { console.log("expandSelectedKits: dup-SO refresh failed: " + dupErr); }

      // Enrich the inserted component SKUs (title note + listing link) from MI.
      try { refreshAllOrdersEnrichment(); }
      catch (enrErr) { console.log("expandSelectedKits: SKU enrichment failed: " + enrErr); }
    }

    return {
      ok:       true,
      message:  expanded + " kit(s) expanded · " + refused.length + " refused · " + skipped.length + " skipped",
      expanded: expanded,
      refused:  refused,
      skipped:  skipped,
      details:  details
    };

  } catch (err) {
    try { console.log("expandSelectedKits error: " + err + "\n" + err.stack); } catch (_) {}
    return {
      ok: false,
      message: "Expand failed: " + (err.message || err),
      expanded: 0, refused: [], skipped: [], details: []
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// =======================================================================================
// EDITOR TEST WRAPPER — for verifying preview output before sidebar UI exists
// =======================================================================================
//
// To use:
//   1. Open All Orders, select one or more rows containing kit SKUs
//   2. Open Apps Script editor → pick `previewSelectedKitsNow` from dropdown
//   3. Run → check Execution Log for the JSON output
//
// Change `deploy` below if you want to test the multiplier (e.g., deploy=3
// to see what "deploy 3 kits per row" produces).
function previewSelectedKitsNow() {
  var deploy = 1;
  var result = previewSelectedKits(deploy);
  console.log(JSON.stringify(result, null, 2));
  return result;
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/**
 * Collects distinct 1-based row numbers from the sheet's current selection.
 * Handles single ranges, multi-selections, and full-column selections (we cap
 * at the last data row to avoid iterating thousands of empty rows).
 */
function _collectSelectedRows(sheet) {
  var rangeList = sheet.getActiveRangeList();
  if (!rangeList) return [];

  var lastRow = sheet.getLastRow();
  var seen = {};
  var out = [];

  var ranges = rangeList.getRanges();
  for (var i = 0; i < ranges.length; i++) {
    var rng = ranges[i];
    var start = rng.getRow();
    var nRows = rng.getNumRows();
    var end = Math.min(start + nRows - 1, lastRow);
    for (var r = start; r <= end; r++) {
      if (!seen[r]) {
        seen[r] = true;
        out.push(r);
      }
    }
  }

  out.sort(function(a, b) { return a - b; });
  return out;
}


// =======================================================================================
// =======================================================================================
// MODAL ORCHESTRATION (shipped 2026-05-19) — popup-based kit expansion with
// per-component exclusion controls and Sales Description reference panel.
// =======================================================================================
// =======================================================================================
//
// WHY A MODAL — the sidebar is 310px wide and the Sales Description reference
// panel needs side-by-side layout with the components checklist. Modal dialogs
// give us ~900px to work with.
//
// FLOW
//   1. Picker selects N rows on All Orders, clicks Preview in sidebar
//   2. Sidebar → openKitExpansionModal(deployQty) — server reads selection,
//      builds the MANUAL-kit queue, stashes it in CacheService with a session
//      ID, opens the modal with the first kit's data
//   3. Modal renders kit checklist (components) + Sales Description (reference)
//   4. Picker unchecks components packed-together-with-another-SKU (Sales
//      Description shows packaging hints like "1G790-03612 + Head Gasket")
//   5. Picker clicks Expand → commitKitFromModal(sessionId, excludedSkus) —
//      server commits this one kit's expansion (with exclusions applied),
//      returns the next kit's data or a done signal
//   6. Modal swaps body to next kit response, repeats until queue exhausted
//   7. Picker can Skip any kit (skipKitFromModal advances without commit)
//   8. Closing the modal (✕) abandons remaining kits but keeps anything
//      already committed
//
// ROW DRIFT — between modal-open and each commit, other operations (n8n
// inserts, manual edits, sort, even our own prior commits in this session)
// can shift kit row numbers. _findKitRowBySkuAndSo re-locates the kit row at
// each commit by scanning for {SKU, SALES_ORDER} match — robust to ALL shift
// causes. The cached row number is only a hint for tie-breaking when multiple
// rows of the same kit exist on the same SO.
//
// SESSION STATE (CacheService.getUserCache)
//   key: KitExpansionModal:<sessionId>
//   value: JSON.stringify({
//     queue: [{kitSku, sourceSalesOrder, originalRow, userMultiplier, sourceQty,
//             kitName, kitType, kitEngine, salesDescription, components: [...]}],
//     currentIndex: number,
//     results: { committed: [], skipped: [], failed: [] }
//   })
//   TTL: 1800s (30 min) — modal abandonment cleans itself up on its own.
//
// PUBLIC API
//   openKitExpansionModal(deployQty)              — sidebar entry point
//   commitKitFromModal(sessionId, excludedSkus)   — modal: commit & advance
//   skipKitFromModal(sessionId)                   — modal: skip & advance
//   closeKitExpansionSession(sessionId)           — modal: ✕ cleanup
// =======================================================================================


var KIT_MODAL_CACHE_PREFIX = "KitExpansionModal:";
var KIT_MODAL_CACHE_TTL    = 1800;   // 30 min


/**
 * Sidebar entry point. Builds the modal queue from the current sheet selection,
 * stashes it in CacheService, opens the modal with the first MANUAL kit's data.
 *
 * READY kits ARE included in the queue (since 2026-05-22) so the picker can
 * Force Expand them per-kit (DIRECT-order use case — keep the pre-assembled
 * K-* box reserved for eBay, pick components individually for this order).
 * The modal shows a warning banner on READY pages with Skip-as-default and
 * Force Expand as a secondary action. Non-kit rows and kits with no
 * registered components are still filtered out — those have no actionable
 * choice. skippedReady stays in the return shape for back-compat but is
 * always 0 going forward.
 *
 * @param {number} deployQty — user-entered multiplier (default 1).
 * @returns {{
 *   ok: boolean,
 *   message: string,
 *   sessionId: string,
 *   queueLength: number,
 *   skippedReady: number,     — always 0 going forward (kept for back-compat)
 *   skippedNonKit: number,    — non-kit rows in selection (with reason)
 *   modalOpened: boolean
 * }}
 */
function openKitExpansionModal(deployQty) {
  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) {
      return { ok: false, message: "No active spreadsheet.", sessionId: "",
               queueLength: 0, skippedReady: 0, skippedNonKit: 0, modalOpened: false };
    }
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) {
      return { ok: false, message: "All Orders sheet not found.", sessionId: "",
               queueLength: 0, skippedReady: 0, skippedNonKit: 0, modalOpened: false };
    }

    // The modal flow IGNORES the sidebar Deploy× input — each kit's modal
    // page starts at multiplier=1, picker explicitly types higher per-kit
    // if needed. So we always preview at ×1 (base qtys).
    // (The sidebar input still controls the bulk Expand-button path which
    // doesn't go through the modal — that fallback is unchanged.)
    var preview = previewSelectedKits(1);
    if (!preview.ok) {
      return { ok: false, message: preview.message, sessionId: "",
               queueLength: 0, skippedReady: 0, skippedNonKit: 0, modalOpened: false };
    }

    // READY kits ARE included now — modal renders a warning page for them,
    // picker decides per-kit (Skip ships pre-assembled box; Force Expand
    // breaks down to components). Only kits with no registered components
    // are dropped (no actionable choice).
    var queue = [];
    var skippedReady = 0;   // always 0 going forward; kept for back-compat
    var skippedEmpty = 0;
    for (var i = 0; i < preview.kitRows.length; i++) {
      var k = preview.kitRows[i];
      if (!k.plan.components || k.plan.components.length === 0) { skippedEmpty++; continue; }

      // Each kit's modal page starts at multiplier=1 (safest default — no
      // accidental over-shipping). Components are ALREADY at base × rowQty
      // because we called previewSelectedKits(1) above. Modal's per-kit
      // multiplier input multiplies these on display.
      queue.push({
        kitSku:           k.sourceSku,
        sourceSalesOrder: k.sourceSalesOrder,
        sourceQty:        k.sourceQty,
        sourceNote:       k.sourceNote,
        originalRow:      k.row,
        table:            k.table,
        userMultiplier:   1,                   // each kit's modal page starts fresh
        kitName:          k.plan.kitName,
        kitType:          k.plan.kitType,
        kitLocation:      k.plan.kitLocation,
        kitEngine:        k.plan.kitEngine,
        salesDescription: k.plan.salesDescription || "",
        components:       k.plan.components    // base qty (×1) — modal re-scales
      });
    }

    var skippedNonKit = preview.nonKitRows.length;

    if (queue.length === 0) {
      var reasonBits = [];
      if (skippedEmpty > 0)  reasonBits.push(skippedEmpty + " empty kit(s)");
      if (skippedNonKit > 0) reasonBits.push(skippedNonKit + " non-kit row(s)");
      var summary = reasonBits.length > 0 ? " (" + reasonBits.join(", ") + ")" : "";
      return {
        ok: false,
        message: "No expandable kits in selection" + summary + ".",
        sessionId: "", queueLength: 0,
        skippedReady: skippedReady, skippedNonKit: skippedNonKit,
        modalOpened: false
      };
    }

    // --- Stash session in cache ---
    var sessionId = Utilities.getUuid();
    var state = {
      queue:        queue,
      currentIndex: 0,
      results:      { committed: [], skipped: [], failed: [] }
    };
    var cache = CacheService.getUserCache();
    cache.put(KIT_MODAL_CACHE_PREFIX + sessionId, JSON.stringify(state), KIT_MODAL_CACHE_TTL);

    // --- Open the modal (HtmlService template, body passes session + first kit) ---
    // CRITICAL: kitJson is injected into a <script> block via <?!= ?> (force
    // unescaped) so the JSON literal parses as a JS object. Apps Script's
    // standard <?= ?> HTML-escapes quotes/angle-brackets which would corrupt
    // the JSON. The `</` → `<\/` substitution defends against the one
    // remaining risk — a literal `</script>` substring inside any value (e.g.,
    // from a Sales Description or kit name with that text) would otherwise
    // close the script tag prematurely. `<\/script>` in JS source is valid
    // (the `\/` is a no-op forward-slash escape) and parses to `</script>`
    // as a string value, but the HTML parser doesn't see it as a tag close.
    var firstKit = queue[0];
    var template = HtmlService.createTemplateFromFile("KitExpansionModal");
    template.sessionId   = JSON.stringify(sessionId);             // becomes a valid JS string literal
    template.queueLength = queue.length;                          // number — safe inline
    template.kitJson     = JSON.stringify(firstKit).replace(/<\//g, "<\\/");
    template.kitIndex    = 0;

    var html = template.evaluate()
      .setWidth(920)
      .setHeight(620);
    SpreadsheetApp.getUi().showModalDialog(html, "Kit Expansion · " + queue.length + " kit" + (queue.length > 1 ? "s" : ""));

    return {
      ok: true,
      message: "Modal opened — " + queue.length + " kit(s) queued",
      sessionId: sessionId,
      queueLength: queue.length,
      skippedReady: skippedReady,
      skippedNonKit: skippedNonKit,
      modalOpened: true
    };
  } catch (err) {
    try { console.log("openKitExpansionModal error: " + err + "\n" + err.stack); } catch (_) {}
    return {
      ok: false,
      message: "Failed to open modal: " + (err.message || err),
      sessionId: "", queueLength: 0, skippedReady: 0, skippedNonKit: 0, modalOpened: false
    };
  }
}


/**
 * Modal commit handler. Commits the current kit in the session with the
 * picker's exclusion choices + per-kit multiplier applied, advances the
 * queue, returns the next kit's data (or a done signal).
 *
 * Re-locates the kit row at commit time by SKU + SALES_ORDER scan — robust
 * to row shifts from prior commits in this same session, concurrent n8n
 * inserts, manual sheet edits, etc.
 *
 * @param {string} sessionId
 * @param {string[]} excludedSkus — component SKUs to skip for this kit
 * @param {number}   [extras]     — additive "spares for us" count. 0 = ship
 *                                  exactly the customer's rowQty (no spares).
 *                                  N = ship rowQty + N total kits. Min 0,
 *                                  defaults to 0 if omitted/invalid.
 *                                  Replaces the old "multiplier" semantic
 *                                  (which scaled rowQty and broke at rowQty>1).
 * @param {boolean}  [force]      — when true, expand even if kit is READY type.
 *                                  Use case: DIRECT order where keeping the
 *                                  pre-assembled K-* box reserved for eBay
 *                                  matters more than picker time. Stamps
 *                                  "READY · forced" in Activity Log DETAIL
 *                                  and "⚠ FORCED" in the component NOTE
 *                                  prefix for audit visibility.
 * @returns {{
 *   ok: boolean,
 *   committed: { kitSku, componentsAdded, excludedSkus, extras, totalKits, rowQty, forced } | null,
 *   reason: string,            — only if !ok or skipped
 *   done: boolean,             — true when no more kits in queue
 *   next: object | null,       — next kit's data (same shape as queue items)
 *   nextIndex: number,         — 0-based position of next kit
 *   queueLength: number
 * }}
 */
function commitKitFromModal(sessionId, excludedSkus, extras, force) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, reason: "Another operation in progress. Try again.",
             committed: null, done: false, next: null, nextIndex: -1, queueLength: 0 };
  }

  try {
    var state = _loadKitModalSession(sessionId);
    if (!state) {
      return { ok: false, reason: "Modal session expired (>30 min). Close and re-preview.",
               committed: null, done: true, next: null, nextIndex: -1, queueLength: 0 };
    }

    if (state.currentIndex >= state.queue.length) {
      return { ok: true, reason: "", committed: null, done: true, next: null,
               nextIndex: state.currentIndex, queueLength: state.queue.length };
    }

    var current = state.queue[state.currentIndex];

    // Extras (spares for us) — additive on top of rowQty. Min 0.
    var extrasNum = parseInt(extras);
    if (isNaN(extrasNum) || extrasNum < 0) extrasNum = 0;

    var commitResult = _commitOneKitForModal(current, excludedSkus || [], extrasNum, !!force);

    if (commitResult.ok) {
      state.results.committed.push({
        kitSku:          current.kitSku,
        kitType:         current.kitType,
        componentsAdded: commitResult.componentsAdded,
        excludedSkus:    commitResult.excludedSkus,
        extras:          commitResult.extras,
        totalKits:       commitResult.totalKits,
        rowQty:          commitResult.rowQty,
        forced:          commitResult.forced
      });
    } else {
      state.results.failed.push({
        kitSku:  current.kitSku,
        kitType: current.kitType,
        reason:  commitResult.reason
      });
    }

    state.currentIndex++;
    _saveKitModalSession(sessionId, state);

    var done = state.currentIndex >= state.queue.length;
    return {
      ok:          commitResult.ok,
      reason:      commitResult.reason || "",
      committed:   commitResult.ok ? {
        kitSku:          current.kitSku,
        kitType:         current.kitType,
        componentsAdded: commitResult.componentsAdded,
        excludedSkus:    commitResult.excludedSkus,
        extras:          commitResult.extras,
        totalKits:       commitResult.totalKits,
        rowQty:          commitResult.rowQty,
        forced:          commitResult.forced
      } : null,
      done:        done,
      next:        done ? null : state.queue[state.currentIndex],
      nextIndex:   state.currentIndex,
      queueLength: state.queue.length,
      sessionResults: done ? state.results : null
    };
  } catch (err) {
    try { console.log("commitKitFromModal error: " + err + "\n" + err.stack); } catch (_) {}
    return { ok: false, reason: "Commit failed: " + (err.message || err),
             committed: null, done: false, next: null, nextIndex: -1, queueLength: 0 };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


/**
 * Modal skip handler. Advances the queue without committing the current kit.
 * Used when the picker decides to handle a kit manually instead.
 *
 * @param {string} sessionId
 * @returns same shape as commitKitFromModal (committed is always null)
 */
function skipKitFromModal(sessionId) {
  var state = _loadKitModalSession(sessionId);
  if (!state) {
    return { ok: false, reason: "Modal session expired. Close and re-preview.",
             committed: null, done: true, next: null, nextIndex: -1, queueLength: 0 };
  }

  if (state.currentIndex < state.queue.length) {
    var skipped = state.queue[state.currentIndex];
    // READY-kit skips read as "ship as box" in the audit (the picker
    // deliberately chose to keep the kit as one unit); MANUAL skips are
    // generic "Picker skipped".
    var skipReason = (skipped.kitType === KIT_REGISTRY.types.READY)
      ? "READY · ship as box"
      : "Picker skipped";
    state.results.skipped.push({
      kitSku:  skipped.kitSku,
      kitType: skipped.kitType,
      reason:  skipReason
    });
    state.currentIndex++;
    _saveKitModalSession(sessionId, state);
  }

  var done = state.currentIndex >= state.queue.length;
  return {
    ok:          true,
    reason:      "",
    committed:   null,
    done:        done,
    next:        done ? null : state.queue[state.currentIndex],
    nextIndex:   state.currentIndex,
    queueLength: state.queue.length,
    sessionResults: done ? state.results : null
  };
}


/**
 * Modal cleanup on close (✕). Best-effort — cache will TTL out anyway.
 */
function closeKitExpansionSession(sessionId) {
  try {
    CacheService.getUserCache().remove(KIT_MODAL_CACHE_PREFIX + sessionId);
  } catch (_) {}
  return { ok: true };
}


// =======================================================================================
// MODAL-PRIVATE HELPERS
// =======================================================================================

function _loadKitModalSession(sessionId) {
  if (!sessionId) return null;
  try {
    var raw = CacheService.getUserCache().get(KIT_MODAL_CACHE_PREFIX + sessionId);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_) { return null; }
}

function _saveKitModalSession(sessionId, state) {
  try {
    CacheService.getUserCache().put(
      KIT_MODAL_CACHE_PREFIX + sessionId,
      JSON.stringify(state),
      KIT_MODAL_CACHE_TTL
    );
  } catch (_) {}
}


/**
 * Per-kit commit for the modal path. Mirrors the inner-loop body of
 * expandSelectedKits but works on a single pre-resolved kit object (from the
 * cached queue), re-locates the row at the moment of commit to handle drift.
 *
 * @param {object}   queueItem    — one entry from the session queue
 * @param {string[]} excludedSkus — component SKUs to skip
 * @param {number}   multiplier   — per-kit deploy multiplier (from the modal's
 *                                   per-kit input, NOT the cached userMultiplier).
 *                                   Min 1.
 * @param {boolean}  [force]      — when true, expand even if kit is READY.
 *                                   Annotates the component NOTE prefix
 *                                   ("⚠ FORCED") and Activity Log DETAIL
 *                                   ("READY · forced") so the override is
 *                                   visible on the sheet and in the audit.
 * @returns {{ok, reason, componentsAdded, excludedSkus, extras, totalKits, rowQty, forced}}
 */
function _commitOneKitForModal(queueItem, excludedSkus, multiplier, force) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return { ok: false, reason: "All Orders sheet not found" };

  // Re-locate the kit row by SKU + SALES_ORDER — handles row drift from any
  // cause (prior commits this session, n8n inserts, manual edits, sort).
  var kitRow = _findKitRowBySkuAndSo(sheet,
                                      queueItem.kitSku,
                                      queueItem.sourceSalesOrder,
                                      queueItem.originalRow);
  if (kitRow < 0) {
    return { ok: false,
             reason: "Kit row not found (may have been deleted, expanded by another picker, or its SO# changed). Cancel and re-preview." };
  }

  var savedHeaders = sheet.getRange(Schema.headerRow, 1, 1, Schema.dataWidth).getValues()[0];
  var locInvMaps   = buildLocationAndInventoryMaps();
  var zohoMap      = buildZohoStockMap();   // DIRECT-side HAND source (Zoho-first)

  var SKU_I      = Schema.idx("SKU");
  var QTY_I      = Schema.idx("QTY");
  var SO_I       = Schema.idx("SALES_ORDER");
  var NOTE_I     = Schema.idx("NOTE");
  var STATUS_I   = Schema.idx("STATUS");
  var SHIPPING_I = Schema.idx("SHIPPING");

  // Fresh re-read of the kit row's current values (NOTE / STATUS / SHIPPING
  // may have changed since modal opened — picker-edited buyer notes, etc.).
  var rowVals    = sheet.getRange(kitRow, 1, 1, Schema.dataWidth).getValues()[0];
  var rowSku     = String(rowVals[SKU_I]    || "").trim();
  var rowQty     = parseInt(rowVals[QTY_I]) || 1;
  var rowSo      = String(rowVals[SO_I]     || "");
  var rowNote    = String(rowVals[NOTE_I]   || "");
  var rowStatus  = String(rowVals[STATUS_I] || Schema.status.PENDING);
  var rowShip    = String(rowVals[SHIPPING_I] || "");

  // Sanity: the located row should still match — _findKitRowBySkuAndSo
  // already filtered on these, but defensive against extreme race conditions
  if (rowSku !== queueItem.kitSku) {
    return { ok: false, reason: "Row drift: located row " + kitRow + " has SKU " + rowSku + " (expected " + queueItem.kitSku + ")" };
  }

  // ADDITIVE semantic: multiplier param is "extras for us" (0 = just the
  // customer's order, no spares; N = N extra kits beyond rowQty).
  // Total kits to ship = rowQty + extras.
  // This semantic is correct regardless of rowQty (the row-multiplier semantic
  // we had earlier only worked at rowQty=1; broke at rowQty=2+).
  var extras = parseInt(multiplier);
  if (isNaN(extras) || extras < 0) extras = 0;
  if (extras > 99) extras = 99;

  // Defensive: rowQty=0 + extras=0 means nothing to ship. parseInt above
  // floored rowQty to 1 if 0/negative, so this is mostly belt-and-suspenders.
  var totalKits       = rowQty + extras;
  if (totalKits < 1) {
    return { ok: false, reason: "Total kits to build is 0 (rowQty=" + rowQty + " + extras=" + extras + ")" };
  }

  var effectiveDeploy = totalKits;  // expandKit scales components by this
  var plan            = expandKit(rowSku, effectiveDeploy);

  if (!plan.found) {
    return { ok: false, reason: plan.reason };
  }
  // READY kits require force=true. Picker uses Skip-default + Force Expand
  // secondary on the modal page; if Force was clicked, we proceed and stamp
  // the override in the audit trail (NOTE prefix + Activity Log DETAIL).
  var isReadyForced = (plan.kitType === KIT_REGISTRY.types.READY) && !!force;
  if (plan.kitType === KIT_REGISTRY.types.READY && !force) {
    return { ok: false, reason: "READY kit — use Force Expand to override" };
  }

  // Apply exclusions
  var excludedSet = {};
  var capturedExclusions = [];
  for (var ex = 0; ex < (excludedSkus || []).length; ex++) {
    excludedSet[String(excludedSkus[ex]).trim()] = true;
  }
  var keptComponents = [];
  for (var fc = 0; fc < plan.components.length; fc++) {
    var compSku = String(plan.components[fc].sku).trim();
    if (excludedSet[compSku]) {
      capturedExclusions.push(compSku);
    } else {
      keptComponents.push(plan.components[fc]);
    }
  }
  plan.components = keptComponents;

  var N = plan.components.length;
  if (N === 0) {
    return { ok: false,
             reason: "All components excluded — nothing to insert. "
                   + "(Excluded: " + capturedExclusions.join(", ") + ")",
             componentsAdded: 0,
             excludedSkus: capturedExclusions };
  }

  // Build kit-expansion NOTE prefix + merge buyer note.
  // When extras > 0, append the breakdown so picker (and audit reader) can
  // see why the QTY column is higher than what the customer ordered.
  // When forced (READY kit overridden), append "⚠ FORCED" so the sheet
  // itself surfaces that this kit was expanded instead of shipped as a box.
  var notePrefix = "↳ from KIT-" + rowSku;
  if (isReadyForced) notePrefix += " · ⚠ FORCED";
  if (extras > 0) {
    notePrefix += " · deploy " + totalKits + " total ("
                + rowQty + " for customer + " + extras + " for us)";
  }
  var componentNote = rowNote ? (notePrefix + " · " + rowNote) : notePrefix;

  // Insert + populate
  sheet.insertRowsAfter(kitRow, N);
  var newRows = [];
  for (var c = 0; c < N; c++) {
    var comp = plan.components[c];
    var skuLower = comp.sku.toLowerCase();
    var loc = locInvMaps.locationMap.get(skuLower) || "NOT FOUND";
    var inv = locInvMaps.inventoryMap.get(skuLower);
    var zo  = zohoMap.get(skuLower);
    // DIRECT-side rows take HAND Zoho-first, MI fallback (matches recomputeHand).
    // "" when neither source has it — the next recompute resolves it.
    var hand = zo ? zo.available
             : (inv && inv.available != null) ? inv.available : "";

    var row = new Array(Schema.dataWidth);
    row[Schema.idx("SKU")]         = comp.sku;
    row[Schema.idx("QTY")]         = comp.qty;
    row[Schema.idx("LOCATION")]    = loc;
    row[Schema.idx("SALES_ORDER")] = rowSo;
    row[Schema.idx("NOTE")]        = componentNote;
    row[Schema.idx("STATUS")]      = rowStatus;
    row[Schema.idx("HAND")]        = hand;
    row[Schema.idx("LEFT")]        = "";
    row[Schema.idx("SHIPPING")]    = rowShip;
    row[Schema.idx("SHIP_COST")]   = "";
    newRows.push(row);
  }
  sheet.getRange(kitRow + 1, 1, N, Schema.dataWidth).setValues(newRows);
  verifyAndRestoreHeaders(sheet, savedHeaders);

  // Activity Log — one RECEIVED per inserted component, DETAIL carries
  // exclusion + extras + force-override context when applicable
  var baseDetail = "kit expansion from " + rowSku;
  if (isReadyForced) baseDetail += " (READY · forced)";
  if (extras > 0) baseDetail += " (deploy " + totalKits + " total: " + rowQty + "+" + extras + ")";
  if (capturedExclusions.length > 0) {
    baseDetail += " · excluded: " + capturedExclusions.join(", ");
  }
  var activityLog = [];
  for (var c2 = 0; c2 < N; c2++) {
    var comp2 = plan.components[c2];
    activityLog.push([
      "RECEIVED",
      rowSo,                         // orderId
      comp2.sku,                     // sku
      comp2.qty,                     // qty
      "sidebar",                     // source
      baseDetail,                    // detail
      undefined,                     // picker: resolved by logActivityBatch via G2
      componentNote                  // note
    ]);
  }
  try { logActivityBatch(activityLog); }
  catch (logErr) { try { console.log("modal commit: activity log failed: " + logErr); } catch (_) {} }

  // Refresh kit SKU markers (same reasoning as expandSelectedKits — setValues
  // doesn't fire onEdit, plus the "↳ from KIT-" NOTE suppression rule needs
  // to re-run for any sub-assembly component rows)
  try { refreshKitSkuMarkers(); }
  catch (kitErr) { try { console.log("modal commit: kit marker refresh failed: " + kitErr); } catch (_) {} }

  // Enrich the inserted component SKUs (title note + listing link) from MI.
  try { refreshAllOrdersEnrichment(); }
  catch (enrErr) { try { console.log("modal commit: SKU enrichment failed: " + enrErr); } catch (_) {} }

  // Repaint duplicate-SO borders — newly-inserted component rows share the
  // kit row's SO# so the group needs to be surfaced immediately.
  try { setupDuplicateSalesOrderHighlighting(); }
  catch (dupErr) { try { console.log("modal commit: dup-SO refresh failed: " + dupErr); } catch (_) {} }

  return {
    ok: true,
    reason: "",
    componentsAdded: N,
    excludedSkus: capturedExclusions,
    extras:        extras,
    totalKits:     totalKits,
    rowQty:        rowQty,
    forced:        isReadyForced
  };
}


/**
 * Re-locates a kit row by {SKU, SALES_ORDER} match. Resilient to row drift
 * from any source (prior commits, n8n inserts, manual edits, sort).
 *
 * If multiple unexpanded rows match (e.g., customer ordered 2 of the same
 * kit on the same SO), returns the one closest to hintRow (the row we saw
 * at modal-open time). Skips rows that are already expanded — i.e. the row
 * directly below carries the "↳ from KIT-<sku>" tag for this kit.
 *
 * @returns 1-based row number, or -1 if no match found.
 */
function _findKitRowBySkuAndSo(sheet, kitSku, soNumber, hintRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return -1;

  var nRows = lastRow - Schema.dataStartRow + 1;
  var data = sheet.getRange(Schema.dataStartRow, 1, nRows, Schema.dataWidth).getValues();

  var SKU_I  = Schema.idx("SKU");
  var SO_I   = Schema.idx("SALES_ORDER");
  var NOTE_I = Schema.idx("NOTE");

  var targetSku = String(kitSku).trim();
  var targetSo  = String(soNumber).trim();
  var matches = [];

  for (var i = 0; i < data.length; i++) {
    var rowSku = String(data[i][SKU_I] || "").trim();
    if (rowSku !== targetSku) continue;

    var rowSo = String(data[i][SO_I] || "").trim();
    if (rowSo !== targetSo) continue;

    var actualRow = Schema.dataStartRow + i;

    // Skip already-expanded matches (row directly below has the kit tag)
    if (i + 1 < data.length) {
      var belowNote = String(data[i + 1][NOTE_I] || "");
      if (belowNote.indexOf("↳ from KIT-" + targetSku) === 0) continue;
    }

    matches.push(actualRow);
  }

  if (matches.length === 0) return -1;
  if (matches.length === 1) return matches[0];

  // Multiple matches — pick closest to hint (last-seen position)
  var best = matches[0];
  var bestDist = Math.abs(matches[0] - hintRow);
  for (var j = 1; j < matches.length; j++) {
    var dist = Math.abs(matches[j] - hintRow);
    if (dist < bestDist) { best = matches[j]; bestDist = dist; }
  }
  return best;
}
