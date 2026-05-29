// =======================================================================================
// STATUS_SERVICE.gs — Canonical status-transition function
// =======================================================================================
//
// PURPOSE
//   Single entry point for ALL order status transitions. Consolidates the
//   logic that was previously scattered across:
//       findAndUpdateOrder        (OrderService.js)
//       updateStatus              (OrderService.js)
//       handleManualStatusChange  (OrderService.js)
//       markSelectedPreparing     (FulfillmentService.js)
//
//   Every status change — whether triggered by Telegram callback, n8n webhook,
//   manual cell edit, or Sidebar bulk button — now flows through this function
//   in the SAME canonical sequence:
//
//       1. validate input (Schema.isValidStatus)
//       2. acquire lock
//       3. resolve target → list of row numbers
//       4. read current state for those rows (one batch)
//       5. partition into rowsToWrite vs blockedRows (terminal-state guard)
//       6. batch-write status for non-blocked rows (consecutive runs grouped)
//       7. refresh sheet stats
//       8. (optional) sort eBay table by status + location
//       9. (optional) sync each unique order to Telegram
//      10. release lock
//      11. return structured result
//
//   Result of consolidation:
//   - The SHIPPED-revert hole in markSelectedPreparing is closed
//   - Identical behavior across every entry point
//   - Adding a new status (e.g., ON_HOLD) means editing one function
//   - Logging consistently includes the `source` so issues are traceable
//
// USAGE
//   updateOrderStatus(orderId,  Schema.status.PREPARING, { source: "telegram" });
//   updateOrderStatus(rowNum,   Schema.status.SHIPPED,   { source: "n8n-direct" });
//   updateOrderStatus([5,6,7],  Schema.status.PREPARING, { source: "sidebar-bulk" });
//   updateOrderStatus({ startRow: 5, numRows: 10 }, Schema.status.PREPARING,
//                     { source: "sidebar-bulk", sortAfter: false });
//
// TARGET TYPES (the function dispatches based on type):
//   - string   → Sales Order ID. Resolves to all rows whose SALES_ORDER matches.
//   - number   → Specific row number.
//   - array    → Array of row numbers.
//   - object   → { startRow, numRows } range spec.
//
// OPTIONS
//   source        — "telegram" | "n8n" | "manual-edit" | "sidebar-bulk" | "n8n-direct"
//                   Used in logs to trace which path triggered a change.
//   syncTelegram  — boolean. Default true. Skip Telegram sync (e.g. for n8n-direct
//                   where the webhook handler already manages Telegram).
//   sortAfter     — boolean. Default true. Set false for manual-edit (don't disturb
//                   the user mid-edit) and for paths that don't want a sort cascade.
//   force         — boolean. Default false. When true, BYPASSES the terminal-state
//                   guard so SHIPPED/CANCELED rows can transition. Use only for
//                   manual-edit paths where the user has already typed the value
//                   into the cell — refusing to sync would diverge sheet from Telegram.
//                   Programmatic callers (Telegram callback, Sidebar bulk, n8n)
//                   should leave this false so accidental terminal-state reverts
//                   are blocked.
//
// RETURN SHAPE
//   {
//     success:       boolean,
//     count:         number of rows actually written,
//     blockedCount:  number of rows blocked due to terminal state,
//     blockedRows:   [{row, orderId, blockedStatus}, ...],
//     ordersSynced:  number of unique orders synced to Telegram,
//     source:        the source string passed in,
//     currentStatus: (legacy) first blocked status — preserved for callers that
//                    expect findAndUpdateOrder's old return shape,
//     error:         (only on failure) error message string
//   }
// =======================================================================================

/**
 * Canonical status-transition function. See file header for full contract.
 */
function updateOrderStatus(target, newStatus, options) {
  options = options || {};
  var source       = options.source || "unknown";
  var syncTelegram = options.syncTelegram !== false;   // default true
  var sortAfter    = options.sortAfter    !== false;   // default true
  var force        = options.force === true;            // default false

  // 1. Validate + normalize
  // Schema.normalize maps spelling/format aliases to the canonical form
  // (e.g. eBay's "Cancelled" / "CANCELLED" → "CANCELED"). Doing this BEFORE
  // validation fixes the silent-reject bug where n8n status updates with
  // eBay's British spelling were dropped on the floor. Doing it BEFORE the
  // setValue ensures the cell ends up with a value that matches the F-column
  // dropdown validation list.
  newStatus = Schema.normalize(newStatus);
  if (!Schema.isValidStatus(newStatus)) {
    return _statusResult(false, 0, [], 0, source, "Invalid status: " + newStatus);
  }

  // 2. Lock
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    return _statusResult(false, 0, [], 0, source, "Could not acquire lock");
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) {
      return _statusResult(false, 0, [], 0, source, "No data rows");
    }

    // 3. Resolve target → list of row numbers
    var targetRows = _resolveStatusTargetRows(sheet, target, lastRow);
    if (targetRows.length === 0) {
      return _statusResult(false, 0, [], 0, source, "No matching rows for target");
    }

    // 4. Read current state for all target rows in one batch
    targetRows = targetRows.slice().sort(function(a, b) { return a - b; });
    var minRow = targetRows[0];
    var maxRow = targetRows[targetRows.length - 1];
    var spanData = sheet.getRange(
      minRow, 1,
      maxRow - minRow + 1,
      Schema.cols.STATUS
    ).getValues();

    // 5. Partition: rowsToWrite vs blockedRows (terminal state guard)
    var rowsToWrite = [];
    var blockedRows = [];
    var ordersToSync = {};
    var logEntries = [];   // staged for Activity Log; written after successful batch write

    targetRows.forEach(function(row) {
      var arrIdx    = row - minRow;
      var oldStatus = String(spanData[arrIdx][Schema.idx("STATUS")]).trim().toUpperCase();
      var orderId   = String(spanData[arrIdx][Schema.idx("SALES_ORDER")]).trim();
      var sku       = String(spanData[arrIdx][Schema.idx("SKU")]).trim();
      var qty       = parseInt(spanData[arrIdx][Schema.idx("QTY")]) || 0;
      var note      = String(spanData[arrIdx][Schema.idx("NOTE")] || "").trim();

      // Terminal-state guard: SHIPPED and CANCELED rows can't transition
      // (unless we're "writing" the same terminal status — then it's a no-op).
      // BYPASS via force: true (manual-edit path — cell already changed, blocking
      // would diverge sheet from Telegram).
      if (!force && Schema.isTerminal(oldStatus) && oldStatus !== newStatus) {
        blockedRows.push({ row: row, orderId: orderId, blockedStatus: oldStatus });
      } else {
        rowsToWrite.push(row);
        if (orderId) ordersToSync[orderId] = true;
        // No-op writes (oldStatus === newStatus) skip the log so we don't
        // pollute it with redundant entries.
        if (oldStatus !== newStatus) {
          // Slots: [event, orderId, sku, qty, source, detail, picker?, note]
          logEntries.push([
            newStatus,
            orderId,
            sku,
            qty,
            source,
            oldStatus ? ("from " + oldStatus) : "",
            undefined,   // picker — let logActivityBatch resolve from G2 if warehouse-side
            note         // note from the order row
          ]);
        }
      }
    });

    // If everything was blocked, return early with the blocked details
    if (rowsToWrite.length === 0) {
      console.log("updateOrderStatus[" + source + "] all " + blockedRows.length +
                  " row(s) blocked (terminal state)");
      return _statusResult(true, 0, blockedRows, 0, source, null);
    }

    // 6. Batch-write status for non-blocked rows (group consecutive runs)
    _writeStatusBatched(sheet, rowsToWrite, newStatus);

    // 6b. Activity Log — one event per row written. Best-effort (silent on
    // failure) so a logging error never rolls back a successful status change.
    try { logActivityBatch(logEntries); } catch (e) {
      console.log("updateOrderStatus: activity log error: " + e);
    }

    // 7. Refresh sheet stats
    try { updateOrderStatsInSheet(); } catch (e) {
      console.log("updateOrderStatus: stats refresh error: " + e);
    }

    // 8. Optional sort — BOTH tables. A status change can land in the eBay
    // table, the DIRECT table, or both (an SO with rows in each). Previously
    // this only sorted table 1 (eBay), so DIRECT-table rows flipped to SHIPPED
    // (Telegram callback, n8n verify, sidebar bulk) wrote correctly but the
    // DIRECT segment never re-sorted — statuses stayed interleaved while eBay
    // stayed clean. Sort is idempotent and cheap, so sort both unconditionally.
    if (sortAfter) {
      try { sortTableByStatusAndLocation(1); } catch (e) {
        console.log("updateOrderStatus: eBay sort error: " + e);
      }
      try { sortTableByStatusAndLocation(2); } catch (e) {
        console.log("updateOrderStatus: DIRECT sort error: " + e);
      }
    }

    // 9. Optional Telegram sync — one editMessageText per unique order
    var orderIds = Object.keys(ordersToSync);
    var syncedCount = 0;
    if (syncTelegram) {
      orderIds.forEach(function(orderId) {
        try {
          syncStatusToTelegram(orderId, newStatus);
          syncedCount++;
        } catch (e) {
          console.log("updateOrderStatus: telegram sync error for " + orderId + ": " + e);
        }
      });
    }

    // 10. Log
    console.log("updateOrderStatus[" + source + "] " + rowsToWrite.length + " row(s) → " +
                newStatus + " · " + blockedRows.length + " blocked · " +
                syncedCount + "/" + orderIds.length + " telegram synced");

    // 11. Return
    return _statusResult(true, rowsToWrite.length, blockedRows, syncedCount, source, null);

  } finally {
    lock.releaseLock();
  }
}


// =======================================================================================
// PRIVATE HELPERS
// =======================================================================================

/**
 * Builds a structured result object. Preserves legacy `currentStatus` field for
 * callers that expected findAndUpdateOrder's old return shape.
 */
function _statusResult(success, count, blockedRows, syncedCount, source, error) {
  var r = {
    success:      success,
    count:        count,
    blockedCount: blockedRows.length,
    blockedRows:  blockedRows,
    ordersSynced: syncedCount,
    source:       source
  };
  if (error) r.error = error;
  // Legacy: expose first blocked status as `currentStatus`.
  // Old callers (findAndUpdateOrder wrapper, handleTelegramCallback) check
  // `result.currentStatus === "SHIPPED"` to detect terminal-state blocks.
  if (blockedRows.length > 0) {
    r.currentStatus = blockedRows[0].blockedStatus;
  } else {
    r.currentStatus = "";
  }
  return r;
}


/**
 * Resolves a target spec to an array of row numbers in the data area.
 *   - string   → sales order ID, find all matching rows
 *   - number   → single row number
 *   - array    → array of row numbers (filtered for validity)
 *   - object   → { startRow, numRows } range spec
 */
function _resolveStatusTargetRows(sheet, target, lastRow) {
  // Single row number
  if (typeof target === 'number') {
    return (target >= Schema.dataStartRow && target <= lastRow) ? [target] : [];
  }

  // Sales Order ID — find all matching rows
  if (typeof target === 'string') {
    var clean = String(target).trim().toLowerCase();
    if (!clean) return [];
    var data = sheet.getRange(
      Schema.dataStartRow,
      Schema.cols.SALES_ORDER,
      lastRow - Schema.dataStartRow + 1,
      1
    ).getValues();
    var rows = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === clean) {
        rows.push(Schema.dataStartRow + i);
      }
    }
    return rows;
  }

  // Array of row numbers
  if (Array.isArray(target)) {
    return target.filter(function(r) {
      return typeof r === 'number' && r >= Schema.dataStartRow && r <= lastRow;
    });
  }

  // Range spec { startRow, numRows }
  if (target && typeof target === 'object' &&
      typeof target.startRow === 'number' &&
      typeof target.numRows === 'number') {
    var rows2 = [];
    for (var r = target.startRow; r < target.startRow + target.numRows; r++) {
      if (r >= Schema.dataStartRow && r <= lastRow) rows2.push(r);
    }
    return rows2;
  }

  return [];
}


/**
 * Writes the same status to many rows efficiently.
 * Groups consecutive row numbers into single setValues() calls for speed.
 *
 *   rows = [5, 6, 7, 12, 15, 16] →
 *     one setValues(5..7), one setValue(12), one setValues(15..16)
 *
 * For the common Sidebar-bulk case (employee selects rows 4-30 contiguously)
 * this is a single setValues call instead of 27 separate setValue calls —
 * roughly 10x faster.
 */
function _writeStatusBatched(sheet, rows, newStatus) {
  if (rows.length === 0) return;
  var sorted = rows.slice().sort(function(a, b) { return a - b; });

  var i = 0;
  while (i < sorted.length) {
    var runStart = sorted[i];
    var runEnd   = sorted[i];
    while (i + 1 < sorted.length && sorted[i + 1] === runEnd + 1) {
      i++;
      runEnd = sorted[i];
    }
    var runLen = runEnd - runStart + 1;
    if (runLen === 1) {
      sheet.getRange(runStart, Schema.cols.STATUS).setValue(newStatus);
    } else {
      var values = [];
      for (var k = 0; k < runLen; k++) values.push([newStatus]);
      sheet.getRange(runStart, Schema.cols.STATUS, runLen, 1).setValues(values);
    }
    i++;
  }
}
