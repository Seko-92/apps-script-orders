// =======================================================================================
// SCHEMA.gs — Single source of truth for All Orders sheet data structure///
// =======================================================================================
//
// PURPOSE
//   Owns column geometry, status enum, and the boundary-marker contract that
//   the rest of the codebase depends on. Eliminates magic numbers and magic
//   strings spread across files. To reorder, rename, or add a column you
//   change THIS file and nothing else.
//
// CONTRACT
//   - Schema.cols.* are 1-BASED column numbers (matches Apps Script getRange convention)
//   - Schema.idx(name) returns the 0-BASED index for getValues()[i][j] arrays
//   - Schema.boundaryMarker MUST stay exactly "DIRECT" — getBoundaryRow() does
//     a strict equality check. Decorate the RIGHT merge (G:J) of the divider row,
//     NEVER column A.
//   - Schema.status.* are the canonical strings; cell values are normalized
//     via .trim().toUpperCase() before comparison.
//
// USAGE
//   sheet.getRange(row, Schema.cols.STATUS).setValue(Schema.status.PREPARING);
//   var status = row[Schema.idx("STATUS")];
//   if (status === Schema.status.SHIPPED) { ... }
//   if (Schema.isValidStatus(status)) { ... }
//   if (Schema.isTerminal(currentStatus)) { return; }  // SHIPPED or CANCELED
//
// BACKWARDS COMPATIBILITY
//   The old Config.js constants (SKU_COLUMN, STATUS_COLUMN, DATA_START_ROW,
//   DATA_WIDTH, TABLE_TWO_IDENTIFIER, etc.) are kept as duplicate literals
//   so any not-yet-migrated reference still resolves correctly. Schema and
//   Config use the SAME numeric values — they don't drift.
// =======================================================================================

var Schema = {

  // =====================================================================================
  // COLUMNS — 1-based numbers (Apps Script getRange convention)
  // =====================================================================================
  cols: {
    SKU:         1,   // A
    QTY:         2,   // B
    LOCATION:    3,   // C
    SALES_ORDER: 4,   // D
    NOTE:        5,   // E
    STATUS:      6,   // F
    HAND:        7,   // G
    LEFT:        8,   // H
    SHIPPING:    9,   // I
    SHIP_COST:  10    // J
  },

  /**
   * 0-based array index for the named column.
   * Use when reading from getValues()[rowIdx][colIdx] arrays.
   * Throws if the name is unknown — fail loud over silent wrong index.
   */
  idx: function(name) {
    var c = Schema.cols[name];
    if (typeof c !== 'number') {
      throw new Error("Schema.idx: unknown column name '" + name + "'");
    }
    return c - 1;
  },

  /** Total number of data columns (full row width) */
  dataWidth: 10,


  // =====================================================================================
  // ROW GEOMETRY
  // =====================================================================================
  bannerRows:    3,    // rows 1-3 carry banner content (logo, stats, banner cells)
  headerRow:     3,    // column headers live in this row
  dataStartRow:  4,    // first data row in eBay table


  // =====================================================================================
  // BOUNDARY CONTRACT
  // =====================================================================================

  /**
   * The exact string that MUST live in column A of the divider row separating
   * the eBay table (above) from the DIRECT table (below).
   *
   * getBoundaryRow() does a strict equality check (after trim + uppercase).
   * If anything writes a decorated value here ("HQ DIRECT", "▌ DIRECT", etc.)
   * the entire downstream system silently breaks: row inserts go to wrong
   * positions, sorts can't find tables, live-sync writes "NOT FOUND" into
   * existing LOCATION cells.
   *
   * Visual richness goes in the RIGHT MERGE (G:J) of the divider row.
   */
  boundaryMarker: "DIRECT",

  /** Width of the boundary row's left merge (A:F) */
  boundaryLeftWidth:  6,

  /** Width of the boundary row's right merge (G:J) */
  // ⚠ REVERTED 2026-08-31, same day. I widened this to 14 (G:T) for a "bands bleed,
  //   tables don't" rule that was in the plan and approved on paper — and the moment it
  //   was on screen the owner's verdict was that it just extended the band into empty
  //   columns and added nothing. They were right: it was paint, not engineering, and it
  //   was folded in beside real defect fixes where it did not belong.
  //   Design gets judged in the eye, not in a document. Back to G:J.
  boundaryRightWidth: 4,


  // =====================================================================================
  // STATUS ENUM
  // =====================================================================================
  status: {
    PENDING:   "PENDING",
    PREPARING: "PREPARING",
    SHIPPED:   "SHIPPED",
    CANCELED:  "CANCELED"
  },

  /** All valid status strings (also the data-validation list on column F) */
  validStatuses: ["PENDING", "PREPARING", "SHIPPED", "CANCELED"],

  /**
   * Statuses from which an order cannot transition further.
   * Used as guards in findAndUpdateOrder, handleTelegramCallback, etc.
   * to prevent reverting a SHIPPED order back to PREPARING.
   */
  terminalStatuses: ["SHIPPED", "CANCELED"],

  /**
   * Spelling/format aliases for inbound status strings. Maps the alias (UPPER,
   * trimmed) to the canonical form. eBay's API returns "Cancelled" (British,
   * double-L); the rest of the codebase + the F-column dropdown use the
   * American single-L "CANCELED". Without this map, n8n's status updates with
   * eBay's spelling were silently rejected by isValidStatus and the order
   * stayed stuck in PREPARING.
   *
   * Add new aliases here whenever an external system sends a variant.
   */
  aliases: {
    "CANCELLED": "CANCELED",  // eBay / British spelling → canonical
    "CANCEL":    "CANCELED",
    "SHIP":      "SHIPPED",
    "SHIPED":    "SHIPPED"    // common typo
  },

  /**
   * Normalizes any inbound status string to its canonical form: trim,
   * uppercase, then alias-map. Use this on any value that comes from outside
   * the codebase (n8n payload, manual cell edits, future Telegram fields)
   * BEFORE validating or writing it to the sheet.
   */
  normalize: function(s) {
    var u = String(s == null ? "" : s).trim().toUpperCase();
    return Schema.aliases[u] || u;
  },

  /**
   * Returns true if the given string is one of the four valid statuses.
   * Normalizes via trim + uppercase + alias lookup before checking, so
   * "Cancelled" and "CANCELLED" both resolve to "CANCELED" → valid.
   */
  isValidStatus: function(s) {
    return Schema.validStatuses.indexOf(Schema.normalize(s)) !== -1;
  },

  /**
   * Returns true if the given string is a terminal status (SHIPPED or CANCELED).
   * Used to short-circuit further state transitions. Uses normalize() so
   * British/typo variants are still recognized.
   */
  isTerminal: function(s) {
    return Schema.terminalStatuses.indexOf(Schema.normalize(s)) !== -1;
  },


  // =====================================================================================
  // BANNER CELLS — single source of truth for code that writes to them.
  // (Avoids hardcoding "G1", "D1", etc. in helper functions.)
  // =====================================================================================
  cellSyncTime:     "E1",   // last successful n8n sync timestamp — visible, user-facing
  // Banner cells re-anchored 2026-05-19 after the operator manually compacted
  // row 1 + row 2 to absorb the now-hidden cols I + J. The full picture:
  //
  //   Row 1: A1=HQ chip, B1:E1=date+pulse, F1:H1=stats banner (was G1:J1)
  //   Row 2: A2:E2=eBay logo, F2:G2=Shipping (was G2:H2), H2=Adjustment (was I2:J2)
  //   Cols I + J hidden — part of the SHIPPING + SHIP COST soft-delete (see CLAUDE.md)
  //
  // The merges were resized manually via the Sheets UI rather than my
  // programmatic E2:F2 migration plan — the operator's layout is tighter and
  // is what's now in production. Schema is aligned to that.
  //
  // To revert to default layout: change these three constants back to G1/G2/I2,
  // restore the original merges (G1:J1, G2:H2, I2:J2, A2:F2), and unhide cols I+J.
  // ⚠⚠ cellSyncTime ("E1") MUST NOT MOVE. ActivityLog.js regex-parses "h:mm AM/PM"
  // out of that cell into cockpit.lastSyncMinutes, which drives the Floor Board's
  // heartbeat dot, the sidebar System Pulse, /status and the published tick.
  // getDashboardSnapshot IS reachable from doPost, so relocating E1 needs a New
  // Version AND breaks the heartbeat in the window either side of the cut.
  //
  // MASTHEAD LAYOUT (2026-08-30). Row 1 is four zones, bounded by the data table's
  // own column edges — the widths are shared with the table below and cannot move:
  //
  //   A1:C1 (260px)  the state face — A 103 · B 70 · C 87
  //   D1    (232px)  the live headline — cellStats, moved here from F1
  //   E1    (307px)  the pulse — UNMOVED, still carries "h:mm AM/PM"
  //   F1:H1 (337px)  SPARKLINE of today — F 130 · G 100 · H 107
  //
  // ⚠⚠ THESE FIVE NUMBERS WERE WRONG UNTIL 2026-09-02 and are now MEASURED, not claimed.
  //    This block used to read 287 / 233 / 306 / 344; diagnoseBanner()'s new row-1 section
  //    read the live sheet and every one of them was off. banner-mock.js's copy was the
  //    correct one all along, which is the whole problem with widths living in comments:
  //    two files disagreed and neither was checkable without opening the spreadsheet.
  //    Same class as MASTHEAD.dialW claiming 280 against a 260px merge — a constant that
  //    stayed right in code while the sheet moved underneath it.
  // ⭐ Re-run diagnoseBanner() rather than trusting this block. It measures.
  cellMasthead:     "A1",   // state face (A1:C1 merge anchor)
  cellStats:        "D1",   // live headline (was F1 until 2026-08-30)
  cellDayCurve:     "F1",   // SPARKLINE of today (F1:H1 merge anchor)
  // ── PICK ID ADDRESSES · GRACE PERIOD (2026-08-31) ──────────────────────────────
  //
  // The pickers are moving off the banner into the hidden columns so row 2 can carry
  // the eBay table's nameplate. Both addresses are live at once during the cut-over,
  // and WHICH ONE IS CURRENT IS RUNTIME STATE, NOT A CONSTANT.
  //
  // ⚠⚠ WHY A SCRIPT PROPERTY AND NOT JUST EDITING THESE CONSTANTS. A dual-read placed
  //    in HEAD cannot help, because THE PINNED /exec DOES NOT HAVE THE HEAD CODE.
  //    Worse: getDashboardTick caches into CacheService.getScriptCache(), which is
  //    shared across the WHOLE script project — so in any window where the pinned
  //    /exec reads F2 while the HEAD publish trigger reads I2, both write the SAME
  //    KEY. The picker chip would then flicker on a 45-second coin flip: ✓ Pick works,
  //    then asks who you are, then works. That reads as flaky wifi, not as a deploy.
  //
  //    A Script Property is read identically by BOTH versions, so flipping it moves
  //    every surface in the same instant. It is also the one-click rollback: delete
  //    the property and everything resolves 'old' on its next execution, with no
  //    deploy and no New Version.
  //
  // ⚠ UNSET, UNREADABLE OR ANYTHING BUT "new" ⇒ 'old'. Failing safe means failing to
  //   the cells that hold the data today. Guessing "probably migrated" would point
  //   every reader at empty cells.
  cellEmployeeId:       "F2",   // Pick ID for Shipping (F2:G2 merge anchor)
  cellAdjustmentId:     "H2",   // Pick ID for Adjustment (single cell, no merge)
  cellEmployeeIdNext:   "I2",   // where migratePickIdCells puts it
  cellAdjustmentIdNext: "J2",

  /**
   * Which cell currently holds a Pick ID. Every reader goes through this.
   * @param {string} which - 'employee' (default) or 'adjustment'
   * @returns {string} an A1 address
   */
  pickIdA1: function (which) {
    var isNew = Schema._pickIdMode() === 'new';
    return (which === 'adjustment')
      ? (isNew ? Schema.cellAdjustmentIdNext : Schema.cellAdjustmentId)
      : (isNew ? Schema.cellEmployeeIdNext   : Schema.cellEmployeeId);
  },

  // ⚠ MEMOISED PER EXECUTION, and that is deliberate in both directions. _currentPicker
  //   runs inside getDashboardSnapshot, which is on the board tick — one property read
  //   per execution is fine, one per call is not. And the address must NOT be able to
  //   change halfway through a single run: a function that reads F2 and then writes I2
  //   is the row-shift bug class wearing different clothes.
  //   Apps Script resets globals per execution, so a flip takes effect on the next one.
  _pickIdModeCache: null,
  _pickIdMode: function () {
    if (Schema._pickIdModeCache !== null) return Schema._pickIdModeCache;
    var m = 'old';
    try {
      var v = PropertiesService.getScriptProperties().getProperty('PICK_ID_ADDR');
      if (String(v || '').trim().toLowerCase() === 'new') m = 'new';
    } catch (e) {
      m = 'old';   // no PropertiesService, no permission, anything at all → today's cells
    }
    Schema._pickIdModeCache = m;
    return m;
  }
};
