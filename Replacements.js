// =======================================================================================
// REPLACEMENTS.JS — the door for MISSING and REPLACEMENT lines
// =======================================================================================
//
// WHY THIS EXISTS
//   2026-08-28: a row's SALES_ORDER cell was overwritten by hand. `09-15094-35132`
//   became `Missing #: 05-15052-93025` at 12:21:40 — on a row that had already been
//   picked (12:15:30) and shelf-counted (LEFT = 0 at 12:15:48).
//
//   doPost dedupes on SALES_ORDER + "|" + SKU, rebuilt live from the sheet on every
//   sync. The signature `09-15094-35132|212498` stopped existing, so n8n's next sync
//   at 12:25:30 correctly re-inserted the "missing" line. Nothing was corrupted and
//   nothing alerted — what was destroyed was the floor's work, and the only trace was
//   two rows that looked like twins.
//
//   It was a SLIP, not an intrusion: the person had the right intent (add a missing
//   line) and hit the wrong cell. So the fix is not authentication. It is removing the
//   reason anyone hand-types into column D at all.
//
// ⭐ THE ONE PROPERTY THAT MAKES THIS SAFE
//   `addReplacementLine` INSERTS ONLY. It has no code path that writes to an existing
//   row. The 08-28 incident is impossible here by construction, not by care.
//
//   Hand-typing into All Orders was the last major operation in this system with no
//   proper door. Kit expansion got a modal, Zoho pull got a modal, price push got a
//   modal, stock adjust got a numpad. This is the same move, one surface later.
//
// TWO KINDS, BECAUSE THE CASE DECIDES
//   /missing     — an item omitted from a shipment that already went out
//   /replacement — a replacement for a wrong or damaged item
//   Same insert engine, different prefix. The type is for the humans reading the
//   sheet; n8n treats both identically (see the anchored-filter note below).
//
// ⚠ THE ANCHORED-FILTER CONTRACT — load-bearing, do not weaken
//   n8n's shipped-check (node S4) filters order ids with an ANCHORED /^[\d-]+$/ test
//   and only sends clean ids to eBay. That is what keeps these rows out of the batch:
//   a malformed id in a comma-joined `orderIds=` URL poisons the WHOLE call (the
//   2026-05-26 incident, where one `Replacement #: …` row silently stopped every
//   order flipping to SHIPPED for the entire batch).
//
//   Any prefix achieves that — but a composed SO that somehow still matched
//   /^[\d-]+$/ would rejoin the shipped-check and get auto-flipped to SHIPPED, which
//   is the exact failure the prefix exists to prevent. `_rlValidate` ASSERTS the
//   composed value fails that test rather than assuming it. See `_rlSelfCheck`.
//
// HAND / LOCATION RESOLUTION
//   A composed SO is not a clean eBay order id, so `_isManualSalesOrder` returns true
//   → the house rule for manual rows applies: ZOHO-FIRST, MI fallback. Same routing
//   `recomputeHand`, LiveSync, Prep Queue and the DIRECT pull insert already use, so
//   this row's HAND cannot disagree with the next recompute.
//
// SURFACES
//   - Telegram /missing + /replacement  → doPost → OWNER context (writes past any
//                                          column protection Phase 2 installs)
//   - Sidebar card                      → runs as the user; owner-only in practice
//
// ⚠ DEPLOY: the sidebar half is editor-bound (`clasp push` is the whole deploy).
//   The Telegram routes are read by handleTelegramCommand, which doPost serves on
//   /exec — those need a NEW VERSION. Manage Deployments → Edit → New version.
// =======================================================================================


var REPLACEMENT = {

  // The two kinds. `prefix` is what lands in column D ahead of the original order id.
  // ⚠ Changing a prefix changes what the floor reads on the sheet AND what the
  //   Activity Log DETAIL says — it does not change n8n behaviour (any non-numeric
  //   prefix is excluded by the anchored filter). Keep them short; column D is narrow.
  kinds: {
    missing: {
      prefix: "Missing #: ",
      label:  "MISSING",
      blurb:  "an item omitted from a shipment"
    },
    replacement: {
      prefix: "Replacement #: ",
      label:  "REPLACEMENT",
      blurb:  "a replacement for a wrong or damaged item"
    }
  },

  maxQty:       99,    // a replacement line is a handful of units, never a pallet
  maxNoteChars: 200,   // same cap /note uses — column E is read at a glance
  lockWaitMs:   15000, // matches the other insert paths

  // How far back the Activity Log is searched when the original order is no longer
  // on the sheet (shipped rows get swept by n8n at ~1 AM Houston). The log keeps 90
  // days; searching all of it is one read either way, so this is a sanity bound only.
  logLookbackDays: 90
};


// =======================================================================================
// PURE HELPERS — no Sheets access, Node-testable
// =======================================================================================

/**
 * Normalize an order id for MATCHING: strip every non-alphanumeric, uppercase.
 * Same approach as OrderLookup's `_normalizeOrderId`, so a query pasted with
 * surrounding noise ("Order # 05-15052-93025") still resolves.
 *
 * ⭐ Substring matching on the normalized form is what lets an EXISTING replacement
 * row prove the original exists: "Missing #: 05-15052-93025" normalizes to
 * "MISSING051505293025", which contains "051505293025".
 */
function _rlNormalizeOrderId(s) {
  return String(s == null ? "" : s).replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
}


/**
 * THE SELF-CHECK. A composed SALES_ORDER must NOT look like a clean eBay order id.
 *
 * Mirrors `_isManualSalesOrder` (ZohoStock.js) and n8n node S4's filter, which are
 * the same rule stated in two places. If this ever returns false, the row would
 * rejoin the shipped-check and get auto-flipped to SHIPPED behind everyone's back.
 *
 * @returns {boolean} true when the value is SAFE (i.e. will be excluded by S4)
 */
function _rlSelfCheck(composedSo) {
  var s = String(composedSo || "").trim();
  if (!s) return false;
  return !(/^[\d-]+$/.test(s) && s.indexOf("-") !== -1);
}


/**
 * Validate every argument and compose the SALES_ORDER value. PURE — this is the part
 * the Node tests drive.
 *
 * @returns {{ok: boolean, error: string, clean: object}}
 *   clean = { kind, label, blurb, originalOrder, salesOrder, sku, qty, note }
 */
function _rlValidate(kind, originalOrder, sku, qty, note) {
  var kindKey = String(kind || "").trim().toLowerCase();
  var spec = REPLACEMENT.kinds[kindKey];
  if (!spec) {
    return {
      ok: false,
      error: "Unknown kind '" + kind + "' — expected " +
             Object.keys(REPLACEMENT.kinds).join(" or ") + "."
    };
  }

  var original = String(originalOrder || "").trim();
  if (!original) {
    return { ok: false, error: "Original order id is required — which order was this missing from?" };
  }
  // Guard against someone pasting an already-composed value back in ("Missing #: …").
  // Silently re-prefixing would produce "Missing #: Missing #: 05-…" on the sheet.
  var alreadyPrefixed = null;
  Object.keys(REPLACEMENT.kinds).forEach(function (k) {
    var p = String(REPLACEMENT.kinds[k].prefix || "").trim().toLowerCase();
    // ⚠ An EMPTY prefix must never enter this test: "x".indexOf("") === 0 is
    //   ALWAYS true, so a blanked prefix would make every input look already-prefixed
    //   and would mask the self-check refusal below behind a wrong, misleading
    //   message -- sending whoever is debugging at the wrong problem. And a blanked
    //   prefix is precisely the "dropped or mangled" case the self-check exists for.
    //   Caught by test A3b.
    if (!p) return;
    if (original.toLowerCase().indexOf(p) === 0) alreadyPrefixed = k;
  });
  if (alreadyPrefixed) {
    return {
      ok: false,
      error: "That already carries the '" + REPLACEMENT.kinds[alreadyPrefixed].prefix.trim() +
             "' prefix — pass the ORIGINAL order id on its own."
    };
  }
  if (_rlNormalizeOrderId(original).length < 6) {
    return { ok: false, error: "'" + original + "' is too short to be an order id." };
  }

  var skuClean = String(sku || "").trim();
  if (!skuClean) {
    return { ok: false, error: "SKU is required." };
  }

  // Qty: accept "1", 1, " 2 " — reject 0, negatives, fractions, and anything absurd.
  var qtyNum = (qty === "" || qty === null || typeof qty === "undefined")
    ? 1
    : Number(String(qty).trim());
  if (!isFinite(qtyNum) || Math.floor(qtyNum) !== qtyNum) {
    return { ok: false, error: "Qty must be a whole number." };
  }
  if (qtyNum < 1) {
    return { ok: false, error: "Qty must be at least 1." };
  }
  if (qtyNum > REPLACEMENT.maxQty) {
    return { ok: false, error: "Qty " + qtyNum + " exceeds the " + REPLACEMENT.maxQty +
                               " cap — split it, or add the line by hand if it is genuinely that large." };
  }

  var noteClean = String(note || "").trim();
  if (noteClean.length > REPLACEMENT.maxNoteChars) {
    noteClean = noteClean.slice(0, REPLACEMENT.maxNoteChars - 1) + "…";
  }

  var composed = spec.prefix + original;

  // ⚠ THE ASSERTION, not an assumption. See the header note.
  if (!_rlSelfCheck(composed)) {
    return {
      ok: false,
      error: "Refusing: the composed order id '" + composed + "' would still match n8n's " +
             "shipped-check filter, so the row could be auto-flipped to SHIPPED. " +
             "This means a prefix was dropped or mangled — fix REPLACEMENT.kinds."
    };
  }

  return {
    ok: true,
    error: "",
    clean: {
      kind:          kindKey,
      label:         spec.label,
      blurb:         spec.blurb,
      originalOrder: original,
      salesOrder:    composed,
      sku:           skuClean,
      qty:           qtyNum,
      note:          noteClean
    }
  };
}


/**
 * Compose the Activity Log DETAIL string. Pure, so the tests can pin the wording.
 */
function _rlDetail(clean, where) {
  var d = clean.label + " line for " + clean.originalOrder;
  if (where) d += " (original " + where + ")";
  return d;
}


/**
 * Parse a Telegram command tail into arguments. PURE.
 *
 *   /missing 05-15052-93025 212498 2 ship separately
 *            └ order ────┘ └ sku ┘ └q┘ └── note ───┘
 *
 * ⭐ QTY IS OPTIONAL AND POSITIONALLY AMBIGUOUS, so it is detected by SHAPE:
 * token 3 counts as a qty only when it is a bare integer. That makes
 *   /missing 05-15052-93025 212498 ship urgently
 * read as qty 1 with the note "ship urgently", instead of silently swallowing
 * "ship" as a quantity and losing the first word of the note. Guessing wrong in
 * the other direction would be invisible on the sheet.
 *
 * @returns {{originalOrder: string, sku: string, qty: (number|string), note: string}}
 */
function _rlParseCommandArgs(argStr) {
  var parts = String(argStr || "").trim().split(/\s+/).filter(function (x) { return x !== ""; });
  var out = { originalOrder: "", sku: "", qty: "", note: "" };
  if (!parts.length) return out;

  out.originalOrder = parts.shift();
  if (!parts.length) return out;

  out.sku = parts.shift();
  if (!parts.length) return out;

  if (/^\d+$/.test(parts[0])) out.qty = Number(parts.shift());
  out.note = parts.join(" ");
  return out;
}


// =======================================================================================
// LOOKUP — does the original order actually exist?
// =======================================================================================

/**
 * Validate that the original order is one of ours.
 *
 * ⭐ WHY THIS EARNS ITS KEEP: today `Missing #: 05-15052-93025` is unvalidated free
 * text. One wrong digit references an order that was never yours, and nothing ever
 * says so — the row just sits on the pick list forever pointing at nothing.
 *
 * Two stages, cheapest first:
 *   1. All Orders column D (the order may still be open, or be a sibling line)
 *   2. Activity Log column C (catches orders already shipped and swept off the sheet)
 *
 * @returns {{found: boolean, where: string, sample: string}}
 */
function _rlFindOriginal(originalOrder, composedSo, sku) {
  var needle = _rlNormalizeOrderId(originalOrder);
  if (!needle) return { found: false, where: "", sample: "", duplicateRow: 0 };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hit = null;
  var duplicateRow = 0;

  // Normalized signature for the duplicate check — the SAME SALES_ORDER + "|" + SKU
  // pair doPost dedupes on. Only computed when the caller asks for it.
  var dupSo  = composedSo ? _rlNormalizeOrderId(composedSo) : "";
  var dupSku = sku ? String(sku).trim().toLowerCase() : "";

  // ---- Stage 1: All Orders — ONE read answers both questions ----
  try {
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= Schema.dataStartRow) {
        var n = lastRow - Schema.dataStartRow + 1;
        // Read A..D once rather than column D twice — the round trip dominates.
        var block = sheet.getRange(Schema.dataStartRow, 1, n, Schema.cols.SALES_ORDER).getValues();
        for (var i = 0; i < block.length; i++) {
          var raw = String(block[i][Schema.cols.SALES_ORDER - 1] || "");
          if (!raw) continue;
          var norm = _rlNormalizeOrderId(raw);

          if (!hit && norm.indexOf(needle) !== -1) {
            hit = { found: true, where: "on the sheet", sample: raw.trim() };
          }
          if (!duplicateRow && dupSo && norm === dupSo &&
              String(block[i][Schema.cols.SKU - 1] || "").trim().toLowerCase() === dupSku) {
            duplicateRow = Schema.dataStartRow + i;
          }
          if (hit && (duplicateRow || !dupSo)) break;
        }
      }
    }
  } catch (e) {
    console.log("_rlFindOriginal: sheet scan failed: " + e);
  }

  if (hit) { hit.duplicateRow = duplicateRow; return hit; }

  // ---- Stage 2: Activity Log, column C ----
  // Shipped rows get swept by n8n at ~1 AM Houston, so a legitimate "missing item"
  // report on yesterday's order will NOT be on the sheet. The log is the only place
  // that still remembers it.
  try {
    var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (log) {
      var logLast = log.getLastRow();
      if (logLast >= ACTIVITY_LOG.dataStartRow) {
        var m = logLast - ACTIVITY_LOG.dataStartRow + 1;
        var ids = log.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.ORDER_ID, m, 1).getValues();
        var cutoff = new Date().getTime() - (REPLACEMENT.logLookbackDays * 86400000);
        var stamps = log.getRange(ACTIVITY_LOG.dataStartRow, ACTIVITY_LOG.cols.TIMESTAMP, m, 1).getValues();
        // Walk newest-first — the log is append-only and chronological.
        for (var j = ids.length - 1; j >= 0; j--) {
          var ts = stamps[j][0];
          if (ts instanceof Date && ts.getTime() < cutoff) break;
          var oid = String(ids[j][0] || "");
          if (!oid) continue;
          if (_rlNormalizeOrderId(oid).indexOf(needle) !== -1) {
            return { found: true, where: "in the Activity Log", sample: oid.trim(), duplicateRow: duplicateRow };
          }
        }
      }
    }
  } catch (e2) {
    console.log("_rlFindOriginal: log scan failed: " + e2);
  }

  return { found: false, where: "", sample: "", duplicateRow: duplicateRow };
}


/**
 * Resolve LOCATION and HAND for the new row, on the manual-row house rule.
 *
 * ⚠ ZOHO-FIRST. A composed SO is not a clean eBay order id, so `_isManualSalesOrder`
 * is true and the row must read Zoho's `available_stock` first (MI fallback) — the
 * same routing `recomputeHand` will apply on its next pass. Using MI-first here
 * would make the row flash a different number until the next recompute overwrote it.
 */
function _rlResolveStock(sku) {
  var skuLower = String(sku || "").trim().toLowerCase();
  var location = "NOT FOUND";
  var miAvail = null;
  var zoAvail = null;

  try {
    var loc = getSingleLocation(sku);
    if (loc) location = String(loc);
  } catch (e) { console.log("_rlResolveStock: location failed: " + e); }

  try {
    var inv = getSingleInventory(skuLower);
    if (inv && inv.available != null) miAvail = parseFloat(inv.available);
  } catch (e2) { console.log("_rlResolveStock: MI failed: " + e2); }

  try {
    var zo = getSingleZohoStock(skuLower);
    if (zo && zo.available != null) zoAvail = parseFloat(zo.available);
  } catch (e3) { console.log("_rlResolveStock: Zoho failed: " + e3); }

  // ⚠ Guarded for the same reason as the three lookups above: every other step here
  //   degrades rather than blocking, and a HAND we could not resolve must not be the
  //   thing that stops a pick line reaching the floor. In Apps Script all root files
  //   share one global scope so this is belt-and-braces — but an inconsistent
  //   defensive posture is how a surprise gets in later.
  var hand = 0;
  try {
    hand = resolveHandValue(miAvail, zoAvail, true);   // true = prefer Zoho (manual row)
  } catch (e4) {
    console.log("_rlResolveStock: resolveHandValue failed, defaulting HAND to 0: " + e4);
  }

  return {
    location: location,
    hand:     hand,
    knownSku: (miAvail != null || zoAvail != null)
  };
}


// =======================================================================================
// PREVIEW — everything the writer does, minus the write
// =======================================================================================

/**
 * Dry run. Same validation and the same lookups, no sheet mutation.
 * Used by the sidebar to render a confirm card, and by /missing + /replacement to
 * build the card that carries the confirm button.
 *
 * @returns {{ok, error, clean, original, stock, warnings: string[]}}
 */
function previewReplacementLine(kind, originalOrder, sku, qty, note) {
  var v = _rlValidate(kind, originalOrder, sku, qty, note);
  if (!v.ok) return { ok: false, error: v.error };

  var original = _rlFindOriginal(v.clean.originalOrder, v.clean.salesOrder, v.clean.sku);

  // ⚠ REFUSE AN EXACT DUPLICATE. This is what makes a double-tap safe by
  //   construction rather than by the button being removed on success — the same
  //   property /pull gets from re-running its diff at apply time.
  //   The asymmetry picks the direction: wrongly refusing costs friction and the
  //   person bumps the qty on the line that already exists; wrongly allowing puts
  //   TWO identical pick lines on the floor, and someone walks the aisle twice or
  //   ships the part twice. Refusing is the safe direction.
  if (original.duplicateRow) {
    return {
      ok: false,
      error: "That exact line already exists — row " + original.duplicateRow + " is " +
             v.clean.sku + " on '" + v.clean.salesOrder + "'. If you need more units, " +
             "raise the qty on that row instead of adding a second one."
    };
  }

  if (!original.found) {
    return {
      ok: false,
      error: "Order " + v.clean.originalOrder + " is not on the sheet and not in the " +
             "last " + REPLACEMENT.logLookbackDays + " days of the Activity Log. " +
             "Check the digits — a wrong id would sit on the pick list pointing at nothing."
    };
  }

  var stock = _rlResolveStock(v.clean.sku);

  // Warnings NEVER block. They are facts the person should weigh, in the board's own
  // grammar: state the problem, let the human decide.
  var warnings = [];
  if (!stock.knownSku) {
    warnings.push("SKU " + v.clean.sku + " is in neither Zoho nor Master Inventory — " +
                  "HAND will read 0 and LOCATION will say NOT FOUND.");
  } else if (stock.location === "NOT FOUND" || !stock.location) {
    warnings.push("No shelf location for " + v.clean.sku + " — the picker will have to find it.");
  }
  if (stock.knownSku && stock.hand < v.clean.qty) {
    warnings.push("Only " + stock.hand + " on hand against a qty of " + v.clean.qty + ".");
  }

  return {
    ok: true,
    error: "",
    clean: v.clean,
    original: original,
    stock: stock,
    warnings: warnings
  };
}


// =======================================================================================
// THE WRITER — INSERT ONLY
// =======================================================================================

/**
 * Insert ONE replacement/missing line at the top of the eBay table.
 *
 * ⭐ There is deliberately no row-targeting parameter and no update path. This
 * function cannot modify an existing row. That is the whole point.
 *
 * @param {string} kind          "missing" | "replacement"
 * @param {string} originalOrder the ORIGINAL eBay order id, un-prefixed
 * @param {string} sku
 * @param {number|string} qty    defaults to 1
 * @param {string} note          optional, appears in column E
 * @param {string} source        Activity Log source — "telegram" | "sidebar"
 * @returns {{ok, message, salesOrder, row, location, hand, warnings}}
 */
function addReplacementLine(kind, originalOrder, sku, qty, note, source) {
  var pre = previewReplacementLine(kind, originalOrder, sku, qty, note);
  if (!pre.ok) return { ok: false, message: pre.error };

  var clean = pre.clean;
  var stock = pre.stock;
  // Activity Log source. All three are in ACTIVITY_LOG.warehouseSources, so the
  // PICKER column auto-captures from G2 for every one of them.
  var srcIn = String(source || "").trim().toLowerCase();
  var src = (srcIn === "telegram" || srcIn === "board") ? srcIn : "sidebar";

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(REPLACEMENT.lockWaitMs);
  } catch (lockErr) {
    return { ok: false, message: "The sheet is busy right now — try again in a few seconds." };
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, message: "Main sheet not found." };

    var W = Schema.dataWidth;

    // Build the row. LEFT / SHIPPING / SHIP_COST stay blank — LEFT is the picker's
    // own post-pick count, and there is no carrier data for a line we are adding.
    var row = new Array(W).fill("");
    row[Schema.idx("SKU")]         = clean.sku;
    row[Schema.idx("QTY")]         = clean.qty;
    row[Schema.idx("LOCATION")]    = stock.location;
    row[Schema.idx("SALES_ORDER")] = clean.salesOrder;
    row[Schema.idx("NOTE")]        = clean.note;
    row[Schema.idx("STATUS")]      = Schema.status.PENDING;
    row[Schema.idx("HAND")]        = stock.hand;

    // ---- A. Save headers (Google's filter-corruption bug on insertRowsBefore) ----
    var savedHeaders = sheet.getRange(Schema.headerRow, 1, 1, W).getValues()[0];

    // ---- B. Insert at the top of the eBay table ----
    sheet.insertRowsBefore(Schema.dataStartRow, 1);

    // ---- C. Write ----
    sheet.getRange(Schema.dataStartRow, 1, 1, W).setValues([row]);

    // ---- D. Restore headers if Sheets corrupted them ----
    verifyAndRestoreHeaders(sheet, savedHeaders);

    // ---- E. Copy formatting from the row below (borders, fonts, banding) ----
    var templateRow = Schema.dataStartRow + 1;
    sheet.getRange(templateRow, 1, 1, W).copyFormatToRange(
      sheet, 1, W, Schema.dataStartRow, Schema.dataStartRow
    );

    // ---- F. ⚠ NEUTRALIZE THE INHERITED SO BADGE (2026-07-14 / 2026-07-17) ----
    // Step E copies EVERY format from the template row — including the col-D badge
    // number format ('"1️⃣ "@') when that row belongs to a multi-item group. Without
    // this reset the new row wears a false badge, and the sticky painter then
    // faithfully PRESERVES the scramble. New rows must start badge-free no matter
    // what; the painter below re-derives the real badges afterwards.
    sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, 1, 1).setNumberFormat("@");

    // ---- G. Also clear any row-level decoration the template may have carried ----
    // Same anti-bleed as the Zoho pull insert: a flagged template row would
    // otherwise donate its strikethrough + red tint to a brand-new line.
    sheet.getRange(Schema.dataStartRow, 1, 1, W)
      .setBackground(null)
      .setFontLine("none");

  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }

  // ---- Post-insert refreshes. Each isolated: a repaint failure must never make the
  //      insert look like it failed, because the row is already on the sheet. ----

  try {
    logActivityBatch([[
      "RECEIVED",
      clean.salesOrder,
      clean.sku,
      clean.qty,
      src,
      _rlDetail(clean, pre.original.where),
      undefined,          // picker — warehouse source, logActivity reads G2 itself
      clean.note
    ]]);
  } catch (e) { console.log("addReplacementLine: activity log failed: " + e); }

  try { _dashBustTickCache(); } catch (e) { console.log("addReplacementLine: cache bust failed: " + e); }
  try { refreshKitSkuMarkers(); } catch (e) { console.log("addReplacementLine: kit markers failed: " + e); }
  try { refreshAllOrdersEnrichment(); } catch (e) { console.log("addReplacementLine: enrichment failed: " + e); }
  try { setupDuplicateSalesOrderHighlighting(); } catch (e) { console.log("addReplacementLine: badge repaint failed: " + e); }

  // Publish inline — a human is standing there having just added a pick line, and the
  // board is where they will look for it. Same reasoning as the doPost arrival path.
  try {
    if (typeof publishBoardTickInline === "function") publishBoardTickInline("replacement");
  } catch (e) { console.log("addReplacementLine: inline publish failed: " + e); }

  var msg = "✅ Added " + clean.label + " line · " + clean.sku + " ×" + clean.qty +
            " · " + stock.location + " · for " + clean.originalOrder;

  return {
    ok: true,
    message: msg,
    salesOrder: clean.salesOrder,
    row: Schema.dataStartRow,
    location: stock.location,
    hand: stock.hand,
    warnings: pre.warnings
  };
}


// =======================================================================================
// SIDEBAR ENTRY POINTS
// =======================================================================================

/**
 * Sidebar preview. Returns a plain object the panel renders into a confirm card.
 */
function previewReplacementFromSidebar(kind, originalOrder, sku, qty, note) {
  try {
    var p = previewReplacementLine(kind, originalOrder, sku, qty, note);
    if (!p.ok) return { ok: false, error: p.error };
    return {
      ok: true,
      kind:          p.clean.kind,
      label:         p.clean.label,
      salesOrder:    p.clean.salesOrder,
      originalOrder: p.clean.originalOrder,
      originalWhere: p.original.where,
      sku:           p.clean.sku,
      qty:           p.clean.qty,
      note:          p.clean.note,
      location:      p.stock.location,
      hand:          p.stock.hand,
      warnings:      p.warnings
    };
  } catch (e) {
    return { ok: false, error: "Preview failed: " + e };
  }
}


/**
 * Sidebar commit.
 */
function addReplacementFromSidebar(kind, originalOrder, sku, qty, note) {
  try {
    return addReplacementLine(kind, originalOrder, sku, qty, note, "sidebar");
  } catch (e) {
    return { ok: false, message: "Failed: " + e };
  }
}


// =======================================================================================
// FLOOR BOARD ENTRY POINT
// =======================================================================================

/**
 * The floor's door. Reached from the board's ⋯ menu via doPost, which runs as the
 * OWNER — so it writes past the column protection, which is the whole reason the
 * board can be the floor's entry point once the sheet is locked.
 *
 * ⚠ GATED ON THE PICK ID, unlike the hold acknowledgement. The standing ruling there
 * is that "an urgent alert is the worst possible moment to force someone through a
 * dropdown" — but this is deliberate data entry, not an emergency, and a line that
 * lands on the pick list should carry the name of whoever put it there. Returning
 * `needsPicker` is the shape the board already knows: it opens the picker drawer and
 * replays the action afterwards.
 *
 * ⚠ THE BOARD IS PUBLIC AND LOGIN-FREE. Its safety story has always been that
 * `boardSetStatus` is narrowed server-side to PENDING/PREPARING and CANNOT insert.
 * This adds an insert, so state the bound plainly: the engine can only ever INSERT,
 * the original order must validate against real sheet or Activity Log data, and an
 * exact duplicate is refused. The worst case is a junk pick line that is visible and
 * deletable — never a shipment, a cancellation or a deletion.
 */
function boardAddReplacementLine(kind, originalOrder, sku, qty, note) {
  var gate = _boardRequirePicker();
  if (!gate.ok) return gate;                 // { ok:false, needsPicker:true, error }
  return addReplacementLine(kind, originalOrder, sku, qty, note, "board");
}


// =======================================================================================
// EDITOR TEST WRAPPERS — the Run button cannot pass arguments
// =======================================================================================

/**
 * Safe: previews only, writes nothing. Edit the values and Run.
 * ⚠ Output goes to the EXECUTION LOG — the Run button does not display return values
 *   (a lesson that cost an evening on getPublishedTick).
 */
function previewReplacementNow() {
  var out = previewReplacementLine("missing", "05-15052-93025", "212498", 1, "");
  console.log(JSON.stringify(out, null, 2));
  return out;
}
