// =======================================================================================
// ZohoPull.js — Pull modal architecture (Option C)
// =======================================================================================
//
// FOUNDATION FILE for the new picker-driven Pull-from-Zoho flow. Replaces
// the background auto-propagation in ZohoSalesOrders.js with an explicit
// modal-based review-and-apply UX, mirroring the Kit Expansion pattern.
//
// SHIPPING IN STAGES:
//   Step 2 (this file, initial)  — computeZohoSoDiff: pure diff computation
//                                  (no UI, no sheet writes). Editor-Run
//                                  testable via previewZohoPullDiffNow().
//   Step 3 — sanity-check output against real SOs
//   Step 4 — applyZohoPullSelection: server-side commit handler
//   Step 5 — Pull modal HTML/JS
//   Step 6 — sidebar wiring
//   Step 7 — shadow-test
//   Step 8 — delete old propagation
//
// DEPENDS ON (from ZohoSalesOrders.js):
//   - PENDING_SO schema
//   - _resolvePendingRow / _findPendingRow / _findPendingRowByInvoice
//   - _readDirectStateForSo
//   - _normalizeSku
// Apps Script has no module system — these are globals in the same project.
// =======================================================================================


// =======================================================================================
// PUBLIC: computeZohoSoDiff(query)
// =======================================================================================

/**
 * Compute the diff between Zoho's current state (cached in Pending PAYLOAD)
 * and our DIRECT table state for a given SO.
 *
 * READ-ONLY. No sheet writes. No mutations.
 *
 * Resolves query through the same path as Pull (accepts SO# or INV#), reads
 * the cached PAYLOAD column on Pending for Zoho's last-known line items,
 * reads DIRECT state via _readDirectStateForSo (which already excludes
 * kit-expanded component rows). Returns a per-line classification the
 * upcoming Pull modal will render.
 *
 * @param {string} query — SO# (e.g. "SO-22750") or INV# (e.g. "INV-021184")
 *
 * @returns {{
 *   ok: boolean,
 *   reason: string,                  — populated only when ok=false
 *   soNumber: string,                — canonical SO# resolved from query
 *   matchedVia: string,              — "so_number" | "invoice"
 *   pendingRowFound: boolean,
 *   isFirstPull: boolean,            — true if zero existing DIRECT rows for this SO
 *   pulled: boolean,                 — true if PULLED column = "PULLED" on Pending row
 *   pulledAt: string,                — formatted date if pulled, else ""
 *   customerName: string,
 *   totalFormatted: string,
 *   zohoStatus: string,              — confirmed / draft / closed / void
 *   zohoShippedStatus: string,       — pending / partially_shipped / shipped / fulfilled
 *   invoiceNumber: string,
 *   pendingLastUpdated: string,      — when Pending row was last refreshed
 *   lines: Array<{
 *     sku: string,
 *     name: string,
 *     status: "unchanged" | "new" | "qty_changed" | "removed",
 *     zohoQty: number,               — 0 if removed (SKU no longer in Zoho)
 *     directQty: number,             — sum across non-CANCELED DIRECT rows
 *     delta: number,                 — zohoQty - directQty (signed)
 *     location: string,              — from MI (NOT FOUND if absent)
 *     available: number | null,      — Zoho-first (MI fallback); null if in neither
 *     missing: boolean,              — true if SKU not in MI (location-wise)
 *     directRows: Array<{row, qty, status}>  — existing DIRECT rows for this SKU+SO
 *   }>,
 *   summary: {
 *     totalLines: number,
 *     unchanged: number,
 *     new: number,
 *     qtyChanged: number,
 *     removed: number,
 *     anyChanges: boolean            — true if any non-unchanged line present
 *   }
 * }}
 */
function computeZohoSoDiff(query) {
  var out = {
    ok: false,
    reason: "",
    soNumber: "",
    matchedVia: "",
    pendingRowFound: false,
    isFirstPull: false,
    pulled: false,
    pulledAt: "",
    customerName: "",
    totalFormatted: "",
    zohoStatus: "",
    zohoShippedStatus: "",
    invoiceNumber: "",
    pendingLastUpdated: "",
    lines: [],
    summary: { totalLines: 0, unchanged: 0, new: 0, qtyChanged: 0, removed: 0, anyChanges: false }
  };

  var q = String(query || "").trim();
  if (!q) {
    out.reason = "Empty query — pass an SO# or INV# string.";
    return out;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var pendingSheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!pendingSheet) {
    out.reason = "Pending Sales Orders sheet not found.";
    return out;
  }

  // --- 1. Resolve the query to a Pending row ---
  var pendingRow = _resolvePendingRow(pendingSheet, q);
  if (pendingRow < 0) {
    out.reason = "No Pending row matches '" + q + "'. Either Pull failed earlier or use Fetch from Zoho to backfill.";
    return out;
  }
  out.pendingRowFound = true;
  out.matchedVia = /^INV[-_]/i.test(q) ? "invoice" : "so_number";

  // --- 2. Read the Pending row's cached payload + summary fields ---
  var pendingRowValues = pendingSheet.getRange(
    pendingRow, 1, 1, PENDING_SO.dataWidth
  ).getValues()[0];

  var soNumber = String(pendingRowValues[PENDING_SO.idx("SO_NUMBER")] || "").trim();
  if (!soNumber) {
    out.reason = "Pending row matched but SO_NUMBER cell is empty.";
    return out;
  }
  out.soNumber = soNumber;

  out.customerName       = String(pendingRowValues[PENDING_SO.idx("CUSTOMER")]    || "");
  out.totalFormatted     = String(pendingRowValues[PENDING_SO.idx("TOTAL")]       || "");
  out.zohoStatus         = String(pendingRowValues[PENDING_SO.idx("ORDER_STATUS")] || "").toLowerCase();
  out.zohoShippedStatus  = String(pendingRowValues[PENDING_SO.idx("SHIPMENT")]    || "").toLowerCase();
  out.invoiceNumber      = String(pendingRowValues[PENDING_SO.idx("INVOICE")]     || "");
  var pulledCell         = String(pendingRowValues[PENDING_SO.idx("PULLED")]      || "").trim().toUpperCase();
  out.pulled             = (pulledCell === "PULLED");
  out.pulledAt           = _formatDateCell(pendingRowValues[PENDING_SO.idx("PULLED_AT")]);
  out.pendingLastUpdated = _formatDateCell(pendingRowValues[PENDING_SO.idx("LAST_UPDATED")]);

  // --- 3. Parse the cached PAYLOAD ---
  var rawPayload = pendingRowValues[PENDING_SO.idx("PAYLOAD")];
  if (!rawPayload || String(rawPayload).trim() === "") {
    out.reason = "Cached payload missing for " + soNumber + ". Use Refresh from Zoho to repopulate.";
    return out;
  }
  var payload;
  try {
    payload = JSON.parse(String(rawPayload));
  } catch (e) {
    out.reason = "Cached payload is not valid JSON for " + soNumber + ". Use Refresh from Zoho to repopulate.";
    return out;
  }
  if (payload && payload._truncated === true) {
    out.reason = "Cached payload was truncated (line items > 49K chars). Use Refresh from Zoho.";
    return out;
  }

  // --- 4. Build Zoho line totals per SKU ---
  // Same shape as the propagation's algorithm — sum qty per SKU, skip lines
  // without a SKU (Zoho includes shipping / discount / service lines that
  // shouldn't translate to picker rows).
  var zohoBySku = {};
  var zohoNames = {};
  var zohoSkuOrder = [];
  var lineItems = Array.isArray(payload.line_items) ? payload.line_items : [];
  for (var i = 0; i < lineItems.length; i++) {
    var li = lineItems[i] || {};
    var sku = _normalizeSku(li.sku);
    if (!sku) continue;
    var qty = parseInt(li.quantity, 10);
    if (isNaN(qty) || qty < 1) continue;
    if (!(sku in zohoBySku)) {
      zohoBySku[sku] = 0;
      zohoNames[sku] = String(li.name || "");
      zohoSkuOrder.push(sku);
    }
    zohoBySku[sku] += qty;
  }

  // --- 5. Read DIRECT state for this SO ---
  // _readDirectStateForSo already filters out kit-expanded component rows
  // (NOTE prefix "↳ from KIT-") and tracks per-SKU active/CANCELED split.
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var directState = mainSheet ? _readDirectStateForSo(mainSheet, soNumber) : { skus: {} };

  // Total active rows on DIRECT for this SO — drives isFirstPull
  var totalDirectActiveRows = 0;
  Object.keys(directState.skus).forEach(function(sku) {
    if (directState.skus[sku].rows) totalDirectActiveRows += directState.skus[sku].rows.length;
  });
  out.isFirstPull = (totalDirectActiveRows === 0);

  // --- 6. Build MI lookup maps for enrichment ---
  // Location + available — picker sees this in the modal next to each line
  // to make pick-or-defer decisions.
  var locInvMaps;
  try {
    locInvMaps = buildLocationAndInventoryMaps();
  } catch (e) {
    // If MI is unreachable for any reason, fall back to empty maps. The
    // diff itself doesn't need MI; the enrichment is best-effort UI data.
    locInvMaps = { locationMap: new Map(), inventoryMap: new Map() };
    console.log("computeZohoSoDiff: MI map build failed: " + e);
  }
  // These are DIRECT-side lines — availability is Zoho-first (MI fallback),
  // matching recomputeHand's DIRECT routing. Best-effort: empty map on failure.
  var zohoMap;
  try { zohoMap = buildZohoStockMap(); }
  catch (e) { zohoMap = new Map(); console.log("computeZohoSoDiff: Zoho map build failed: " + e); }

  // --- 7. Union SKU set and classify each ---
  var allSkus = {};
  zohoSkuOrder.forEach(function(s) { allSkus[s] = true; });
  Object.keys(directState.skus).forEach(function(s) { allSkus[s] = true; });

  // Preserve Zoho's order for the SKUs Zoho has; appended any DIRECT-only
  // SKUs (removed lines) at the end. Predictable UI ordering.
  var orderedSkus = zohoSkuOrder.slice();
  Object.keys(directState.skus).forEach(function(s) {
    if (orderedSkus.indexOf(s) === -1) orderedSkus.push(s);
  });

  orderedSkus.forEach(function(sku) {
    var zohoQty   = zohoBySku[sku] || 0;
    var directRec = directState.skus[sku];
    var directQty = directRec ? directRec.totalActiveQty : 0;

    var status;
    if (zohoQty > 0 && directQty === 0) status = "new";
    else if (zohoQty === 0 && directQty > 0) status = "removed";
    else if (zohoQty === directQty) status = "unchanged";
    else status = "qty_changed";   // both > 0 but different

    var skuLower = sku.toLowerCase();
    var loc = locInvMaps.locationMap.get(skuLower) || "NOT FOUND";
    var inv = locInvMaps.inventoryMap.get(skuLower);
    var zo  = zohoMap.get(skuLower);
    // Zoho-first availability (MI fallback), null if neither — DIRECT routing.
    var miAvail = (inv && inv.available != null) ? inv.available : null;
    var available = (zo && zo.available != null) ? zo.available : miAvail;

    var lineOut = {
      sku:        sku,
      name:       zohoNames[sku] || "",
      status:     status,
      zohoQty:    zohoQty,
      directQty:  directQty,
      delta:      zohoQty - directQty,
      location:   loc,
      available:  available,
      missing:    (loc === "NOT FOUND"),
      directRows: directRec && directRec.rows ? directRec.rows.slice() : []
    };
    out.lines.push(lineOut);

    out.summary.totalLines++;
    if (status === "unchanged")    out.summary.unchanged++;
    else if (status === "new")     out.summary.new++;
    else if (status === "qty_changed") out.summary.qtyChanged++;
    else if (status === "removed") out.summary.removed++;
  });

  out.summary.anyChanges = (out.summary.new + out.summary.qtyChanged + out.summary.removed) > 0;

  out.ok = true;
  return out;
}


// =======================================================================================
// EDITOR TEST WRAPPER
// =======================================================================================

/**
 * Editor-Run testing entry point. Edit the SO# below to point at whatever
 * you want to inspect. Logs the JSON output of computeZohoSoDiff to the
 * Apps Script execution log.
 *
 * Recommended test cases for step 3 sanity check:
 *   - "SO-22750"  — lifecycle-complete case (the 2026-05-23 bug example).
 *                   Expected: all lines status="removed" because DIRECT rows
 *                   were cleaned up. If shown as "removed", we confirm the
 *                   classifier handles the post-shipping cleanup case.
 *   - A clean active SO that's been Pulled and is mid-pick. Expected:
 *                   all lines status="unchanged" if no Zoho-side edits since.
 *   - An SO with a known qty change on one line. Expected: that line
 *                   status="qty_changed", delta = new - old.
 *   - An SO with a known added line. Expected: that line status="new".
 *   - An SO that's never been Pulled (just sitting in Pending). Expected:
 *                   isFirstPull=true, all lines status="new".
 */
function previewZohoPullDiffNow() {
  var query = "SO-22750";  // change me to test other SOs
  var result = computeZohoSoDiff(query);
  console.log(JSON.stringify(result, null, 2));
  return result;
}


// =======================================================================================
// PRIVATE HELPERS
// =======================================================================================

/**
 * Format a date cell value for display. Handles Date objects (typical of
 * sheet datetime cells) and string fallbacks.
 *
 * Returns "" for empty cells so the modal can safely concat without
 * showing "null" or "undefined".
 */
function _formatDateCell(val) {
  if (val == null || val === "") return "";
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, "America/Chicago", "M/d/yy h:mm a");
    } catch (e) {
      return val.toString();
    }
  }
  return String(val);
}


// =======================================================================================
// PUBLIC: applyZohoPullSelection(query, selections)
// =======================================================================================

/**
 * Commit the picker's per-line decisions from the Pull modal.
 *
 * CONTRACT
 *   - LockService serializes (15s wait) against other Pull/propagation work.
 *   - Recomputes the diff fresh against current state, then validates EACH
 *     selection's action matches the line's current status. Any mismatch
 *     aborts the WHOLE operation (all-or-nothing) with a clear "state
 *     changed since modal opened — re-Preview" message.
 *   - On success: batched inserts via _insertAddedItemsToDirect (preserves
 *     existing duplicate-highlight + kit-marker + activity-log behavior),
 *     per-row flags via _flagDirectRow, cancellations via direct STATUS
 *     write + manual Activity Log entry.
 *   - Pending row's PULLED + PULLED_AT cells are stamped on first apply
 *     (or refreshed timestamp on subsequent applies).
 *
 * ACCEPTED ACTIONS (per selection entry):
 *   "insert"          — line.status must be "new". Inserts one DIRECT row
 *                       at qty = zohoQty.
 *   "insert_delta"    — line.status must be "qty_changed". Inserts ONE row
 *                       with qty = delta. NOTE explains "↳ delta from Zoho".
 *   "flag_existing"   — line.status must be "qty_changed". Flags each
 *                       existing non-CANCELED DIRECT row with the qty-change
 *                       annotation. No insert.
 *   "flag_removed"    — line.status must be "removed". Flags existing rows
 *                       with strikethrough + ⚠ REMOVED IN ZOHO note.
 *   "cancel_removed"  — line.status must be "removed". Flips existing non-
 *                       CANCELED DIRECT rows to CANCELED.
 *
 * @param {string} query — SO# or INV#
 * @param {Array<{sku: string, action: string}>} selections
 *
 * @returns {{
 *   ok: boolean,
 *   reason: string,
 *   soNumber: string,
 *   applied: {
 *     inserted: number,        — count of new DIRECT rows written
 *     flaggedQty: number,      — count of rows flagged with qty-change note
 *     flaggedRemoved: number,  — count of rows flagged as REMOVED
 *     canceled: number         — count of rows flipped to CANCELED
 *   },
 *   skipped: Array<{sku, reason}>,
 *   summary: string            — human-readable for status bar
 * }}
 */
function applyZohoPullSelection(query, selections) {
  var out = {
    ok: false,
    reason: "",
    soNumber: "",
    applied: { inserted: 0, flaggedQty: 0, flaggedRemoved: 0, canceled: 0 },
    skipped: [],
    summary: ""
  };

  var q = String(query || "").trim();
  if (!q) {
    out.reason = "Empty query.";
    return out;
  }
  if (!Array.isArray(selections)) {
    out.reason = "Selections must be an array.";
    return out;
  }
  if (selections.length === 0) {
    // Nothing selected — picker hit Apply with empty checks. Not an error,
    // just a no-op. UI should normally prevent this but be defensive.
    out.ok = true;
    out.summary = "Nothing selected — no changes applied.";
    return out;
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) {
    out.reason = "Server busy, try again.";
    return out;
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var pendingSheet = ss.getSheetByName(PENDING_SO.sheetName);
    var mainSheet    = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!pendingSheet || !mainSheet) {
      out.reason = "Required sheets not found (Pending Sales Orders / All Orders).";
      return out;
    }

    // --- Recompute diff fresh ---
    var diff = computeZohoSoDiff(q);
    if (!diff.ok) {
      out.reason = "Diff failed: " + diff.reason;
      return out;
    }
    out.soNumber = diff.soNumber;

    // Index diff lines by SKU for validation
    var diffBySku = {};
    diff.lines.forEach(function(line) { diffBySku[line.sku] = line; });

    // --- Validate every selection BEFORE applying anything (all-or-nothing) ---
    // The all-or-nothing rule prevents partial application when state has
    // drifted since the modal opened. Picker sees a coherent error and can
    // re-Preview rather than reconcile partially-applied changes.
    var validActions = {
      "insert":         ["new"],
      "insert_delta":   ["qty_changed"],
      "flag_existing":  ["qty_changed"],
      "flag_removed":   ["removed"],
      "cancel_removed": ["removed"]
    };

    for (var v = 0; v < selections.length; v++) {
      var sel = selections[v] || {};
      var selSku = String(sel.sku || "").trim();
      var selAction = String(sel.action || "").trim();

      if (!selSku || !selAction) {
        out.reason = "Selection " + v + " missing sku or action.";
        return out;
      }
      if (!(selAction in validActions)) {
        out.reason = "Unknown action '" + selAction + "' for SKU " + selSku + ".";
        return out;
      }
      var line = diffBySku[selSku];
      if (!line) {
        out.reason = "SKU " + selSku + " no longer present in diff. State changed since modal opened — re-Preview.";
        return out;
      }
      if (validActions[selAction].indexOf(line.status) === -1) {
        out.reason = "Cannot " + selAction + " SKU " + selSku
                   + " — line is currently '" + line.status + "'. "
                   + "State changed since modal opened — re-Preview.";
        return out;
      }
    }

    // --- Bucket selections by action class for efficient processing ---
    var inserts       = [];   // for _insertAddedItemsToDirect batch
    var flagsQty      = [];   // [{line, sku, ...}]
    var flagsRemoved  = [];
    var cancels       = [];

    selections.forEach(function(sel) {
      var line = diffBySku[sel.sku];
      switch (sel.action) {
        case "insert":
          // Fresh line on this SO from Zoho — full qty
          inserts.push({
            sku:             line.sku,
            quantity:        line.zohoQty,
            name:            line.name,
            _noteOverride:   "",   // per user policy: leave NOTE empty on Pull
            _detailOverride: "Pulled from Zoho · " + (diff.customerName || "no customer")
          });
          break;
        case "insert_delta":
          // Qty increase — insert delta row with explanation in NOTE
          inserts.push({
            sku:             line.sku,
            quantity:        line.delta,
            name:            line.name,
            _noteOverride:   "↳ delta from Zoho · was " + line.directQty
                             + " total, now " + line.zohoQty,
            _detailOverride: "Pull delta on existing SKU · was " + line.directQty
                             + ", now " + line.zohoQty
          });
          break;
        case "flag_existing":
          flagsQty.push(line);
          break;
        case "flag_removed":
          flagsRemoved.push(line);
          break;
        case "cancel_removed":
          cancels.push(line);
          break;
      }
    });

    var todayLabel = Utilities.formatDate(new Date(), "America/Chicago", "M/d");

    // OPERATION ORDER MATTERS — row-targeting ops (flags, cancels) reference
    // specific row numbers captured in the diff. `_insertAddedItemsToDirect`
    // calls `insertRowsBefore` which shifts EVERY row below the insert point
    // down by N, invalidating the captured row numbers. So flags + cancels
    // MUST run BEFORE inserts. Sheet preserves cell formatting (strikethrough,
    // background) when rows shift during insertRowsBefore, so any visual
    // changes we make to specific rows here travel correctly to their new
    // post-insert positions.
    //
    // Bug regression test (2026-05-23): inserts-first order caused the qty
    // flag for SKU 171279 to land on a freshly-inserted 163890 row instead
    // of its target. See CLAUDE.md history.

    // --- 1. Flags (qty changes — annotate, no strike) ---
    flagsQty.forEach(function(line) {
      var prefix = "⚠️ ZOHO QTY: " + line.directQty + " → " + line.zohoQty + " " + todayLabel;
      line.directRows.forEach(function(rowInfo) {
        if (rowInfo.status === Schema.status.CANCELED) return;
        _flagDirectRow(mainSheet, rowInfo.row, prefix, false);
        out.applied.flaggedQty++;
      });
    });

    // --- 2. Flags (removed — strikethrough + REMOVED annotation) ---
    flagsRemoved.forEach(function(line) {
      var prefix = "⚠️ REMOVED IN ZOHO " + todayLabel;
      line.directRows.forEach(function(rowInfo) {
        if (rowInfo.status === Schema.status.CANCELED) return;
        _flagDirectRow(mainSheet, rowInfo.row, prefix, true);
        out.applied.flaggedRemoved++;
      });
    });

    // --- 3. Cancels (removed lines the picker chose to flip to CANCELED) ---
    // Direct cell write + manual Activity Log entry. We don't use
    // updateOrderStatus here because it targets ALL rows for an SO, not
    // specific rows. Per-row precision matters when only some SKUs of an
    // SO are being canceled.
    var cancelLogBatch = [];
    cancels.forEach(function(line) {
      line.directRows.forEach(function(rowInfo) {
        if (rowInfo.status === Schema.status.CANCELED) return;
        try {
          mainSheet.getRange(rowInfo.row, Schema.cols.STATUS).setValue(Schema.status.CANCELED);
          // Also strikethrough + soft-red tint to match flag_removed look,
          // since the user-facing meaning is similar (line gone from Zoho,
          // picker confirmed cancellation).
          mainSheet.getRange(rowInfo.row, 1, 1, Schema.dataWidth).setFontLine('line-through');
          mainSheet.getRange(rowInfo.row, 1, 1, Schema.dataWidth).setBackground('#ffe5e5');
          out.applied.canceled++;
          cancelLogBatch.push([
            "CANCELED", diff.soNumber, line.sku, rowInfo.qty,
            "sidebar",
            "Canceled via Pull modal · line removed in Zoho",
            undefined, ""
          ]);
        } catch (e) {
          console.log("applyZohoPullSelection: cancel write failed row " + rowInfo.row + ": " + e);
        }
      });
    });
    if (cancelLogBatch.length > 0) {
      try { logActivityBatch(cancelLogBatch); }
      catch (e) { console.log("applyZohoPullSelection: cancel log error: " + e); }
    }

    // --- 4. Batch inserts LAST (reuses existing helper for consistency) ---
    // Inserts shift rows below, but the flag/cancel work above has already
    // landed on the right rows, and Sheets carries cell formatting along
    // during the shift.
    if (inserts.length > 0) {
      out.applied.inserted = _insertAddedItemsToDirect(
        mainSheet, diff.soNumber, inserts
      );
    }

    // --- 5. Stamp PULLED on Pending row (idempotent on re-pull) ---
    // Any action implies the SO has been touched; refresh the timestamp.
    var pendingRow = _resolvePendingRow(pendingSheet, q);
    if (pendingRow > 0) {
      pendingSheet.getRange(pendingRow, PENDING_SO.cols.PULLED).setValue(PENDING_SO.pulledFlag);
      pendingSheet.getRange(pendingRow, PENDING_SO.cols.PULLED_AT).setValue(new Date());
    }

    // --- Done ---
    out.ok = true;
    var bits = [];
    if (out.applied.inserted > 0)       bits.push(out.applied.inserted + " inserted");
    if (out.applied.flaggedQty > 0)     bits.push(out.applied.flaggedQty + " qty-flagged");
    if (out.applied.flaggedRemoved > 0) bits.push(out.applied.flaggedRemoved + " removed-flagged");
    if (out.applied.canceled > 0)       bits.push(out.applied.canceled + " canceled");
    out.summary = diff.soNumber + " · " + (bits.length > 0 ? bits.join(" · ") : "no changes");
    return out;

  } catch (err) {
    out.reason = "Apply failed: " + (err.message || err);
    try { console.log("applyZohoPullSelection error: " + err + "\n" + err.stack); } catch (_) {}
    return out;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


/**
 * Editor-Run test wrapper for applyZohoPullSelection.
 *
 * Edit the `query` and `selections` constants below to match a real test
 * scenario. CAUTION: this DOES write to the sheet — use on a test SO or
 * accept the changes will be applied.
 *
 * Example: re-pull SO-XXXXX where Zoho has a new line for SKU 165447 qty=2.
 *   query      = "SO-XXXXX"
 *   selections = [{sku: "165447", action: "insert"}]
 */
function applyZohoPullSelectionNow() {
  var query = "SO-22750";   // edit me
  var selections = [];      // edit me — e.g., [{sku: "171279", action: "insert_delta"}]
  var result = applyZohoPullSelection(query, selections);
  console.log(JSON.stringify(result, null, 2));
  return result;
}


// =======================================================================================
// PUBLIC: openZohoPullModal(query)
// =======================================================================================

/**
 * Sidebar entry point. Computes the diff and either opens the Pull modal
 * (when there are changes to review) OR returns a status-only result when
 * the diff is empty (per lock: empty-diff case shows status message, no
 * modal opens).
 *
 * @param {string} query — SO# or INV# typed by the picker
 *
 * @returns {{
 *   ok: boolean,
 *   modalOpened: boolean,
 *   soNumber: string,
 *   reason: string,           — populated when !ok OR modal not opened (status msg)
 *   summary: object | null    — copy of diff.summary when ok
 * }}
 */
function openZohoPullModal(query) {
  try {
    var diff = computeZohoSoDiff(query);
    if (!diff.ok) {
      return {
        ok: false,
        modalOpened: false,
        soNumber: diff.soNumber || "",
        reason: diff.reason,
        summary: null
      };
    }

    // FIRST-PULL CASE — diff treats every Zoho line as "new" (because zero
    // active DIRECT rows). That's a legitimate case to open the modal for —
    // the picker reviews and applies the full list.
    //
    // RE-PULL CASE WITH ZERO CHANGES — diff has lines but all are
    // "unchanged" (anyChanges === false). Per lock: status message only,
    // don't open the modal. Picker sees "nothing to do."
    if (!diff.summary.anyChanges && !diff.isFirstPull) {
      return {
        ok: true,
        modalOpened: false,
        soNumber: diff.soNumber,
        reason: "No changes since " + (diff.pulledAt || "last pull")
              + " · " + diff.summary.unchanged + " line(s) already in DIRECT.",
        summary: diff.summary
      };
    }

    // --- Open the modal ---
    // diffJson is force-unescaped via <?!= ?> in the template (the diff is
    // a JSON string, HTML-escaping would corrupt the quotes). Pre-substitute
    // `</` → `<\/` so any value containing the literal `</script>` substring
    // can't close the script tag prematurely.
    var diffJson = JSON.stringify(diff).replace(/<\//g, "<\\/");
    var template = HtmlService.createTemplateFromFile("ZohoPullModal");
    template.diffJson = diffJson;

    var html = template.evaluate()
      .setWidth(920)
      .setHeight(620);

    var titleBits = [];
    titleBits.push("Pull " + diff.soNumber);
    if (diff.isFirstPull) titleBits.push("First pull · " + diff.summary.totalLines + " lines");
    else                  titleBits.push(diff.summary.new + " new · "
                                       + diff.summary.qtyChanged + " qty · "
                                       + diff.summary.removed + " removed");

    SpreadsheetApp.getUi().showModalDialog(html, titleBits.join(" · "));

    return {
      ok: true,
      modalOpened: true,
      soNumber: diff.soNumber,
      reason: "",
      summary: diff.summary
    };
  } catch (err) {
    try { console.log("openZohoPullModal error: " + err + "\n" + err.stack); } catch (_) {}
    return {
      ok: false,
      modalOpened: false,
      soNumber: "",
      reason: "Failed to open Pull modal: " + (err.message || err),
      summary: null
    };
  }
}


/**
 * Editor-Run wrapper for the Pull modal. Edit the `query` constant to test
 * against any SO# or INV#. Opens the modal for visual review of the diff
 * + UI rendering before wiring the sidebar (step 6).
 *
 * Safe to run from the editor — opens the modal without applying anything
 * until the picker clicks Apply.
 */
function previewZohoPullModalNow() {
  var query = "SO-22750";   // edit me
  var result = openZohoPullModal(query);
  console.log(JSON.stringify(result, null, 2));
  return result;
}
