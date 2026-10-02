// =======================================================================================
// AMAZON.gs — THE THIRD TABLE + ITS DOOR (2026-09-26)
// =======================================================================================
//
// All Orders carries three standalone tables, top to bottom:
//
//     eBay rows  ·  ▌DIRECT band + header  ·  Direct rows  ·  ▌AMAZON band + header  ·  Amazon rows
//
// Amazon is LAST on purpose: new eBay orders insert at the top and Zoho pulls insert at
// DIRECT+2, so neither insert path had to change. Where the three tables are is decided
// in ONE place — getTableLayout() in Helpers.js — and every range writer asks it.
//
// THIS FILE
//   setupAmazonTable()        owner-only · adds the ▌AMAZON band + header + a blank row
//   removeAmazonTable(force)  owner-only · THE ROLLBACK — refuses while Amazon rows exist
//   previewAmazonOrder(…)     dry run: validation + shelf + HAND, no write
//   addAmazonOrder(…)         THE DOOR — INSERT-ONLY, like Replacements.js
//   purgeAmazonTestEvents()   removes Activity Log rows left by a test order (AMZ-000-…)
//
// ⚠⚠ THE DOOR CAN ONLY INSERT. No parameter can address an existing row, so it cannot
//    overwrite a live order — the property that made the /missing door safe on a locked
//    sheet (the 2026-08-28 incident).
//
// ⚠⚠ THE AMZ- PREFIX IS REQUIRED, AND FOR ONE REASON: a raw Amazon id
//    (114-3941689-8772232) is digits-and-dashes, so it would PASS n8n S4's shipped-check
//    filter and be sent to eBay's API. _rlSelfCheck (Replacements.js) asserts it.
//
// ⚠ NOTHING AUTOMATIC TOUCHES THESE ROWS YET (by design — manual first). Marking them
//   SHIPPED is manual, and so is confirming the shipment in Seller Central: if the box
//   leaves and nobody confirms, Amazon records a LATE SHIPMENT. The board says so on
//   every Amazon order band, and the print says so on the section head.
//
// ⚠ ACTIVITY LOG SOURCE IS "amazon", deliberately NOT a warehouse source: the owner
//   enters these from Riyadh, and auto-capturing the shift's Pick ID from G2 would
//   attribute the entry to a picker who never touched it.
// =======================================================================================

var AMAZON = {
  lockWaitMs:   15000,
  maxLines:     20,
  maxQty:       50,
  maxNoteChars: 200,
  // Amazon order ids are 3-7-7 digits: 114-3941689-8772232
  idRe:         /^\d{3}-\d{7}-\d{7}$/,
  // Test orders use this id range so their Activity Log trail can be purged afterwards.
  testIdPrefix: "000-",
  // Ids entered during testing that are NOT in the 000- range, purged the same way.
  // 114-3941689-8772232 was the example printed in the sidebar field and in /amazon's
  // usage text, and the first live test (2026-09-26) used it verbatim.
  testOrderIds: ["114-3941689-8772232"],
  source:       "amazon",
  headerHeight: 36
};


// =======================================================================================
// PURE HELPERS — Node-testable, no Sheets calls
// =======================================================================================

/**
 * "114-3941689-8772232" / "AMZ-114-…" / " #114 3941689 8772232 " → the canonical pair.
 * @returns {{ok, error, id, salesOrder}}
 */
function _amzNormalizeOrderId(raw) {
  var s = String(raw == null ? "" : raw).trim().toUpperCase();
  s = s.replace(/^#/, "");
  if (s.indexOf(Schema.amazonOrderPrefix) === 0) s = s.slice(Schema.amazonOrderPrefix.length);
  s = s.replace(/\s+/g, "-").replace(/-+/g, "-");
  if (!AMAZON.idRe.test(s)) {
    return { ok: false, error: "'" + String(raw || "").trim() + "' is not an Amazon order number " +
                               "(expected 3-7-7 digits, like 114-3941689-8772232)." };
  }
  return { ok: true, error: "", id: s, salesOrder: Schema.amazonOrderPrefix + s };
}

/**
 * [{sku, qty}] → validated lines. Same SKU twice is MERGED (quantities added) rather
 * than refused — typing a part twice is a slip, not a second line.
 * @returns {{ok, error, lines:[{sku, qty}]}}
 */
function _amzCleanLines(lines) {
  if (!lines || !lines.length) return { ok: false, error: "At least one SKU is required." };
  if (lines.length > AMAZON.maxLines) {
    return { ok: false, error: "More than " + AMAZON.maxLines + " lines — split the order." };
  }
  var bySku = {}, order = [];
  for (var i = 0; i < lines.length; i++) {
    var sku = String((lines[i] && lines[i].sku) || "").trim();
    if (!sku) return { ok: false, error: "Line " + (i + 1) + " has no SKU." };
    if (Schema.isStructuralMarker(sku)) return { ok: false, error: "'" + sku + "' is not a SKU." };
    var q = lines[i].qty;
    var qty = (q === "" || q === null || typeof q === "undefined") ? 1 : Number(String(q).trim());
    if (!isFinite(qty) || Math.floor(qty) !== qty || qty < 1) {
      return { ok: false, error: "Qty for " + sku + " must be a whole number of at least 1." };
    }
    var key = sku.toUpperCase();
    if (!bySku[key]) { bySku[key] = { sku: sku, qty: 0 }; order.push(key); }
    bySku[key].qty += qty;
    if (bySku[key].qty > AMAZON.maxQty) {
      return { ok: false, error: "Qty for " + sku + " exceeds the " + AMAZON.maxQty + " cap." };
    }
  }
  return { ok: true, error: "", lines: order.map(function (k) { return bySku[k]; }) };
}

/**
 * Ship-by date → "M/d" text for the NOTE, or "" when not given.
 * Accepts 9/29 · 9/29/2026 · 2026-09-29 · today · tomorrow.
 * @param {string} raw
 * @param {Date} [now]  injectable for tests
 * @returns {{ok, error, text}}
 */
function _amzShipBy(raw, now) {
  var s = String(raw == null ? "" : raw).trim().toLowerCase();
  if (!s) return { ok: true, error: "", text: "" };
  now = now || new Date();
  var d = null;
  if (s === "today") d = new Date(now.getTime());
  else if (s === "tomorrow") d = new Date(now.getTime() + 86400000);
  else {
    var m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/.exec(s);
    var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (m) {
      var y = m[3] ? Number(m[3]) : now.getFullYear();
      if (y < 100) y += 2000;
      d = new Date(y, Number(m[1]) - 1, Number(m[2]), 12);
    } else if (iso) {
      d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 12);
    }
  }
  if (!d || isNaN(d.getTime()) || (m && (d.getMonth() !== Number(m[1]) - 1))) {
    return { ok: false, error: "Ship-by '" + raw + "' is not a date (try 9/29 or 2026-09-29)." };
  }
  return { ok: true, error: "", text: (d.getMonth() + 1) + "/" + d.getDate() };
}

/** The NOTE written on every line of the order. Pure so a test can pin the wording. */
function _amzNote(shipByText, note) {
  // ⭐ 2026-09-26 (user's call): no bare "AMAZON" word — the table and the AMZ- id already
  //   say which channel it is, and on the board every NOTE renders as a 📌, which is for
  //   real exceptions. Only a deadline or a real note earns the cell; otherwise it is blank.
  var parts = [];
  if (shipByText) parts.push("ship by " + shipByText);
  var n = String(note || "").trim();
  if (n.length > AMAZON.maxNoteChars) n = n.slice(0, AMAZON.maxNoteChars - 1) + "…";
  if (n) parts.push(n);
  return parts.join(" · ");
}

/**
 * Parse "/amazon 114-3941689-8772232 167517 2 171378 by 9/29 note gift wrap".
 * SKU/qty pairs follow the id; a qty may be omitted (→1) when the next token is not a
 * number. "by <date>" and "note <text…>" are optional and may come in either order.
 * @returns {{orderId, lines:[{sku,qty}], shipBy, note}}
 */
function _amzParseCommand(argStr) {
  var toks = String(argStr || "").trim().split(/\s+/).filter(function (t) { return t; });
  var out = { orderId: toks.shift() || "", lines: [], shipBy: "", note: "" };
  var i = 0;
  while (i < toks.length) {
    var t = toks[i], tl = t.toLowerCase();
    if (tl === "note") { out.note = toks.slice(i + 1).join(" "); break; }
    if (tl === "by" && i + 1 < toks.length) { out.shipBy = toks[i + 1]; i += 2; continue; }
    var line = { sku: t, qty: 1 };
    if (i + 1 < toks.length && /^\d+$/.test(toks[i + 1])) { line.qty = Number(toks[i + 1]); i += 2; }
    else i += 1;
    out.lines.push(line);
  }
  return out;
}

/**
 * Where a new Amazon order's rows go: into the order's OWN block when it already has
 * rows (so a second entry for the same order stays contiguous — the split-box lesson of
 * 2026-08-07), otherwise the top of the Amazon table.
 * @param {Array[]} colD   column-D values from row 1 (index 0 = row 1)
 * @param {Object} layout  getTableLayout()
 * @param {string} salesOrder
 * @returns {number} sheet row to insertRowsBefore
 */
function _amzInsertRow(colD, layout, salesOrder) {
  var start = layout.amazon + 2;
  var target = String(salesOrder).trim().toUpperCase();
  for (var r = start; r <= colD.length; r++) {
    if (String(colD[r - 1][0] || "").trim().toUpperCase() === target) return r;
  }
  return start;
}


// =======================================================================================
// THE TABLE — setup and rollback
// =======================================================================================

/**
 * Adds the ▌AMAZON band, its header row and one blank row at the bottom of the sheet.
 * Owner-only, idempotent (refuses if the table already exists).
 */
function setupAmazonTable() {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Adding the Amazon table");
    if (denied) return denied;
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var L = getTableLayout(sheet);
  if (L.direct <= 0) return "❌ DIRECT divider not found — refusing to touch row structure.";
  if (L.amazon > 0)  return "ℹ️ The Amazon table already exists (divider at row " + L.amazon + ").";

  var W = Schema.dataWidth;
  var buffer = Math.max(1, (typeof TABLE_BUFFER_ROWS === "number") ? TABLE_BUFFER_ROWS : 1);

  // A previous attempt that died half-way leaves an orphan header at the bottom with no
  // AMAZON band above it (2026-09-26: the band was dropped, the header half-written).
  // Clear it first, or this run would stack a second table under the debris.
  var cleaned = _amzClearFailedSetup(sheet, L);
  if (cleaned.error) return "❌ " + cleaned.error;
  if (cleaned.removed) L = getTableLayout(sheet);

  var at = sheet.getMaxRows();                 // append below the Direct table's tail
  sheet.insertRowsAfter(at, 2 + buffer);
  var band = at + 1, header = at + 2, firstData = at + 3;

  try {
    _amzBuildStructure(sheet, L, band, header, firstData, buffer);
    SpreadsheetApp.flush();      // surface any refused write HERE, not inside a later step
    var bad = _amzVerifyStructure(sheet, L, band, header);
    if (bad) throw new Error(bad);
  } catch (e) {
    // All or nothing: never leave a half-built table on the live sheet.
    try { sheet.deleteRows(band, 2 + buffer); } catch (e2) {}
    try { SpreadsheetApp.flush(); } catch (e3) {}
    var fail = "❌ Amazon table NOT added — the rows were removed again, the sheet is as it was.\n" +
               "   Reason: " + (e && e.message || e);
    console.log(fail);
    return fail;
  }

  var notes = [];
  var step = function (label, fn) {
    try { fn(); notes.push("✓ " + label); } catch (e) { notes.push("⚠ " + label + ": " + e); }
  };
  step("row banding", refreshDynamicBandings);
  step("order boxes", setupDuplicateSalesOrderHighlighting);
  step("row-1 counts", function () { _ensureSparkData(ss); });
  step("staff lock carve-outs", function () {
    if (typeof refreshAllOrdersLockCarveOuts === "function") refreshAllOrdersLockCarveOuts();
  });
  step("board cache", _dashBustTickCache);

  var msg = "✅ Amazon table added — divider at row " + band + ", header " + header +
            ", first row " + firstData + "." +
            (cleaned.removed ? "\n✓ removed " + cleaned.removed + " leftover row(s) from the failed attempt" : "") +
            "\n" + notes.join("\n");
  console.log(msg);
  return msg;
}

/**
 * Band + header + blank rows. Separate from setupAmazonTable so the caller can wrap it
 * in one all-or-nothing try.
 */
function _amzBuildStructure(sheet, L, band, header, firstData, buffer) {
  var W = Schema.dataWidth;
  // Data-row format first (fonts, borders) from a real data row, for all the new rows —
  // then the band and header get their own styling on top.
  sheet.getRange(Schema.dataStartRow, 1, 1, W)
       .copyTo(sheet.getRange(band, 1, 2 + buffer, W),
               SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.getRange(band, 1, 2 + buffer, W).clearContent().setBackground(null).setFontLine("none");
  // ⚠⚠ PASTE_FORMAT CARRIES DATA VALIDATION. The data row's STATUS dropdown lands on the
  //    band and the header too, and then writing "AMAZON" / "STATUS" there is refused —
  //    which is exactly how the first live run lost its band (2026-09-26). The blank data
  //    rows KEEP the dropdown (they need it); the two structural rows must not have it.
  sheet.getRange(band, 1, 2, W).clearDataValidations();
  sheet.getRange(firstData, Schema.cols.SALES_ORDER, buffer, 1).setNumberFormat("@");

  // Header row = the DIRECT header's own labels (the same columns, same words).
  var headerVals = sheet.getRange(L.direct + 1, 1, 1, W).getValues();
  sheet.getRange(header, 1, 1, W).setValues(headerVals);

  // Band: the one shared styler builds the merges, the mark and the nameplate.
  _styleAmazonDivider(sheet, band);
  _styleHeaderRow(sheet, header);
  sheet.setRowHeight(header, AMAZON.headerHeight);
  sheet.setRowHeights(firstData, buffer, 30);
}

/**
 * Reads the structure back. Returns "" when it is right, else what is wrong.
 * Checks what the rest of the system depends on: the exact marker, the header labels,
 * and that getTableLayout() now finds the table where it was built.
 */
function _amzVerifyStructure(sheet, L, band, header) {
  var W = Schema.dataWidth;
  var marker = String(sheet.getRange(band, 1).getValue()).trim().toUpperCase();
  if (marker !== Schema.amazonMarker) {
    return "the band at row " + band + " reads '" + marker + "', not '" + Schema.amazonMarker + "'.";
  }
  var want = sheet.getRange(L.direct + 1, 1, 1, W).getValues()[0];
  var got  = sheet.getRange(header, 1, 1, W).getValues()[0];
  for (var c = 0; c < W; c++) {
    if (String(got[c]) !== String(want[c])) {
      return "header cell " + String.fromCharCode(65 + c) + header + " reads '" + got[c] +
             "', expected '" + want[c] + "'.";
    }
  }
  var L2 = getTableLayout(sheet);
  if (L2.amazon !== band) return "the layout finds the Amazon divider at row " + L2.amazon + ", not " + band + ".";
  if (L2.direct !== L.direct) return "the DIRECT divider moved (" + L.direct + " → " + L2.direct + ").";
  return "";
}

/**
 * Removes the debris of a setup that died half-way: a copy of the table header sitting
 * below the Direct table with no AMAZON band. Only acts when every row from the row
 * above that header to the bottom of the sheet is empty apart from the header itself —
 * anything else and it refuses, because then it is not debris.
 * @returns {{removed:number, error:string}}
 */
function _amzClearFailedSetup(sheet, L) {
  if (!(L.direct > 0) || L.amazon > 0) return { removed: 0, error: "" };
  var W = Schema.dataWidth;
  var top = L.direct + 2, max = sheet.getMaxRows();
  if (max < top) return { removed: 0, error: "" };
  var headA = String(sheet.getRange(L.direct + 1, 1).getValue()).trim();
  var vals = sheet.getRange(top, 1, max - top + 1, W).getValues();
  var h = -1;
  for (var i = 0; i < vals.length; i++) {
    if (headA && String(vals[i][0]).trim() === headA) { h = top + i; break; }
  }
  if (h === -1) return { removed: 0, error: "" };
  var from = h - 1;
  if (from < top) {
    return { removed: 0, error: "Found a stray copy of the table header at row " + h +
             " directly under the Direct header. Not touching it — check the sheet by hand." };
  }
  for (var r = from; r <= max; r++) {
    if (r === h) continue;
    var row = vals[r - top];
    for (var c = 0; c < W; c++) {
      if (String(row[c]).trim() !== "") {
        return { removed: 0, error: "Found a leftover table header at row " + h + " from an earlier " +
                 "attempt, but row " + r + " below it holds data ('" + row[c] + "'). Not deleting " +
                 "anything — clear those rows by hand, then run setupAmazonTable() again." };
      }
    }
  }
  try { sheet.getRange(from, 1, max - from + 1, W).breakApart(); } catch (e) {}
  sheet.deleteRows(from, max - from + 1);
  console.log("setupAmazonTable: removed " + (max - from + 1) + " leftover row(s) from a failed attempt (" + from + "–" + max + ").");
  return { removed: max - from + 1, error: "" };
}

/**
 * THE ROLLBACK. Deletes the ▌AMAZON band, its header and every row below it.
 * Refuses while any Amazon row holds data unless force === true — a rollback must never
 * silently throw away a real order.
 */
function removeAmazonTable(force) {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Removing the Amazon table");
    if (denied) return denied;
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";
  var L = getTableLayout(sheet);
  if (!(L.amazon > 0)) return "ℹ️ There is no Amazon table to remove.";

  var seg = _tableSegment(3, L);
  var last = findLastDataRowInSegment(seg.start, seg.end);
  var rowsHeld = last >= seg.start ? (last - seg.start + 1) : 0;
  if (rowsHeld > 0 && force !== true) {
    return "❌ The Amazon table still has " + rowsHeld + " row(s) with data (rows " + seg.start +
           "–" + last + "). Mark them shipped / move them first, or run removeAmazonTableForce().";
  }

  // Drop the warning-only protections that point at rows about to disappear.
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      var d = String(p.getDescription() || "");
      if (d.indexOf("AMAZON") !== -1) { try { p.remove(); } catch (e) {} }
    });
  } catch (e) {}

  var count = L.maxRows - L.amazon + 1;
  if (count >= sheet.getMaxRows()) return "❌ Refusing: that would delete every row on the sheet.";
  sheet.deleteRows(L.amazon, count);

  var notes = [];
  var step = function (label, fn) {
    try { fn(); notes.push("✓ " + label); } catch (e) { notes.push("⚠ " + label + ": " + e); }
  };
  step("row banding", refreshDynamicBandings);
  step("order boxes", setupDuplicateSalesOrderHighlighting);
  step("row-1 counts", function () { _ensureSparkData(ss); });
  step("board cache", _dashBustTickCache);

  var msg = "✅ Amazon table removed (" + count + " rows" +
            (rowsHeld ? ", including " + rowsHeld + " with data" : "") + ").\n" + notes.join("\n");
  console.log(msg);
  return msg;
}

/** Editor wrapper — the Run button cannot pass arguments. Deletes Amazon rows too. */
function removeAmazonTableForce() { return removeAmazonTable(true); }


// =======================================================================================
// THE DOOR
// =======================================================================================

/**
 * Dry run: same validation and lookups as addAmazonOrder, no sheet mutation.
 * @param {string} orderId
 * @param {Array<{sku, qty}>} lines
 * @param {string} [shipBy]
 * @param {string} [note]
 * @returns {{ok, error, clean, lines, warnings}}
 */
function previewAmazonOrder(orderId, lines, shipBy, note) {
  var id = _amzNormalizeOrderId(orderId);
  if (!id.ok) return { ok: false, error: id.error };
  if (typeof _rlSelfCheck === "function" && !_rlSelfCheck(id.salesOrder)) {
    return { ok: false, error: "Refusing: '" + id.salesOrder + "' would match n8n's eBay shipped-check filter." };
  }
  var cl = _amzCleanLines(lines);
  if (!cl.ok) return { ok: false, error: cl.error };
  var sb = _amzShipBy(shipBy);
  if (!sb.ok) return { ok: false, error: sb.error };

  var out = [], warnings = [], unknown = [];
  for (var i = 0; i < cl.lines.length; i++) {
    var st = _rlResolveStock(cl.lines[i].sku);
    if (!st.knownSku) unknown.push(cl.lines[i].sku);
    if (st.location === "NOT FOUND") warnings.push(cl.lines[i].sku + " has no shelf location.");
    if (st.knownSku && st.hand < cl.lines[i].qty) {
      warnings.push(cl.lines[i].sku + ": only " + st.hand + " on hand for " + cl.lines[i].qty + ".");
    }
    out.push({ sku: cl.lines[i].sku, qty: cl.lines[i].qty,
               location: st.location, hand: st.hand, knownSku: st.knownSku });
  }
  // ⚠ An unknown SKU is REFUSED, not warned: a typo would put a line on the floor that
  //   no shelf holds, on the one channel where a late shipment is an account-health event.
  if (unknown.length) {
    return { ok: false, error: "Unknown SKU" + (unknown.length > 1 ? "s" : "") + ": " +
                               unknown.join(", ") + " — not in Master Inventory or Zoho. Check the number." };
  }
  if (!sb.text) warnings.push("No ship-by date given — Amazon shows it on the order.");
  return {
    ok: true, error: "",
    clean: { orderId: id.id, salesOrder: id.salesOrder, shipBy: sb.text,
             note: _amzNote(sb.text, note) },
    lines: out,
    warnings: warnings
  };
}

/**
 * THE DOOR. Inserts one row per line into the Amazon table, PENDING, as one order.
 * @param {string} source  "telegram" | "sidebar" | "editor"  (recorded in DETAIL)
 * @returns {{ok, message, salesOrder, rows, warnings}}
 */
function addAmazonOrder(orderId, lines, shipBy, note, source) {
  var pre = previewAmazonOrder(orderId, lines, shipBy, note);
  if (!pre.ok) return { ok: false, message: pre.error };
  var clean = pre.clean;
  var via = String(source || "editor").trim().toLowerCase();

  var lock = LockService.getScriptLock();
  try { lock.waitLock(AMAZON.lockWaitMs); }
  catch (e) { return { ok: false, message: "The sheet is busy right now — try again in a few seconds." }; }

  var insertedAt = -1, n = pre.lines.length;
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, message: "Main sheet not found." };
    var L = getTableLayout(sheet);
    if (!(L.amazon > 0)) {
      return { ok: false, message: "The Amazon table is not on the sheet yet — run setupAmazonTable() first." };
    }
    var W = Schema.dataWidth;

    // ---- Duplicate guard: the same SALES_ORDER|SKU signature doPost dedupes on ----
    var last = sheet.getLastRow();
    var colsAD = last >= 1 ? sheet.getRange(1, 1, last, Schema.cols.SALES_ORDER).getValues() : [];
    var have = {};
    for (var r = Schema.dataStartRow - 1; r < colsAD.length; r++) {
      have[String(colsAD[r][Schema.idx("SALES_ORDER")] || "").trim().toUpperCase() + "|" +
           String(colsAD[r][Schema.idx("SKU")] || "").trim().toUpperCase()] = true;
    }
    var dups = pre.lines.filter(function (ln) {
      return have[clean.salesOrder.toUpperCase() + "|" + ln.sku.toUpperCase()];
    });
    if (dups.length) {
      return { ok: false, message: clean.salesOrder + " already has " +
               dups.map(function (d) { return d.sku; }).join(", ") +
               " on the sheet — nothing was added. Change the qty on the existing row instead." };
    }

    // ---- Where: the order's own block, else the top of the Amazon table ----
    var colD = colsAD.map(function (row) { return [row[Schema.idx("SALES_ORDER")]]; });
    var at = _amzInsertRow(colD, L, clean.salesOrder);

    var rows = pre.lines.map(function (ln) {
      var row = new Array(W).fill("");
      row[Schema.idx("SKU")]         = ln.sku;
      row[Schema.idx("QTY")]         = ln.qty;
      row[Schema.idx("LOCATION")]    = ln.location;
      row[Schema.idx("SALES_ORDER")] = clean.salesOrder;
      row[Schema.idx("NOTE")]        = clean.note;
      row[Schema.idx("STATUS")]      = Schema.status.PENDING;
      row[Schema.idx("HAND")]        = ln.hand;
      return row;
    });

    var savedHeaders = sheet.getRange(Schema.headerRow, 1, 1, W).getValues()[0];
    if (at > sheet.getMaxRows()) {
      // The table has no blank row left to insert above — grow the sheet at the end.
      sheet.insertRowsAfter(sheet.getMaxRows(), n);
      at = sheet.getMaxRows() - n + 1;
    } else {
      sheet.insertRowsBefore(at, n);
    }
    sheet.getRange(at, 1, n, W).setValues(rows);
    verifyAndRestoreHeaders(sheet, savedHeaders);

    // Format from a real data row (never the header above), then strip anything that
    // row may have carried: a SO badge, a Zoho flag tint, a strikethrough.
    sheet.getRange(Schema.dataStartRow, 1, 1, W).copyFormatToRange(sheet, 1, W, at, at + n - 1);
    sheet.getRange(at, Schema.cols.SALES_ORDER, n, 1).setNumberFormat("@");
    sheet.getRange(at, 1, n, W).setBackground(null).setFontLine("none");
    insertedAt = at;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  // ---- Post-insert refreshes, each isolated: the rows are already on the sheet ----
  try {
    logActivityBatch(pre.lines.map(function (ln) {
      return ["RECEIVED", clean.salesOrder, ln.sku, ln.qty, AMAZON.source,
              "Amazon order entered via " + via + (clean.shipBy ? " · ship by " + clean.shipBy : ""),
              "", clean.note];
    }));
  } catch (e) { console.log("addAmazonOrder: activity log failed: " + e); }
  try { _dashBustTickCache(); } catch (e) { console.log("addAmazonOrder: cache bust failed: " + e); }
  try { refreshKitSkuMarkers(); } catch (e) { console.log("addAmazonOrder: kit markers failed: " + e); }
  try { refreshAllOrdersEnrichment(); } catch (e) { console.log("addAmazonOrder: enrichment failed: " + e); }
  try { setupDuplicateSalesOrderHighlighting(); } catch (e) { console.log("addAmazonOrder: boxes failed: " + e); }
  try {
    if (typeof refreshAllOrdersLockCarveOuts === "function") refreshAllOrdersLockCarveOuts();
  } catch (e) { console.log("addAmazonOrder: carve-out refresh failed: " + e); }
  try {
    if (typeof publishBoardTickInline === "function") publishBoardTickInline(undefined, "amazon");
  } catch (e) { console.log("addAmazonOrder: inline publish failed: " + e); }

  var msg = "✅ " + clean.salesOrder + " added to the Amazon table — " +
            pre.lines.map(function (l) { return l.qty + "× " + l.sku + " @ " + l.location; }).join(", ") +
            (clean.shipBy ? " · ship by " + clean.shipBy : "");
  return { ok: true, message: msg, salesOrder: clean.salesOrder, row: insertedAt,
           rows: n, warnings: pre.warnings };
}


// =======================================================================================
// SIDEBAR ENTRY POINTS
// =======================================================================================

/** Sidebar preview — lines arrive as [{sku, qty}]. */
function previewAmazonOrderFromSidebar(orderId, lines, shipBy, note) {
  try { return previewAmazonOrder(orderId, lines, shipBy, note); }
  catch (e) { return { ok: false, error: "Preview failed: " + e }; }
}

/** Sidebar commit — hops to the owner under the All Orders lock (see OwnerBridge.js). */
function addAmazonOrderFromSidebar(orderId, lines, shipBy, note) {
  if (!_obIsOwner()) return _asOwner('addAmazonOrderFromSidebar', [orderId, lines, shipBy, note]);
  try { return addAmazonOrder(orderId, lines, shipBy, note, "sidebar"); }
  catch (e) { return { ok: false, message: "Failed: " + e }; }
}

/** Updates LOCATION for the Amazon table (the Direct button's twin). */
function runUpdateLocationsTableThree() {
  if (!_obIsOwner()) return _asOwner('runUpdateLocationsTableThree', []);
  return updateAllExistingRows(3);
}


// =======================================================================================
// TEST CLEAN-UP
// =======================================================================================

/**
 * A test order's RECEIVED rows SURVIVE deleting the order and inflate "Received Today"
 * for good (the KitBuild lesson). Test orders use the AMZ-000-… id range, and this
 * deletes exactly those Activity Log rows — nothing else can match that prefix.
 */
function purgeAmazonTestEvents() {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Purging Amazon test events");
    if (denied) return denied;
  }
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ACTIVITY_LOG.sheetName);
  if (!sh) return "ℹ️ No Activity Log sheet.";
  var last = sh.getLastRow();
  if (last < 2) return "ℹ️ Activity Log is empty.";
  var col = ACTIVITY_LOG.cols.ORDER_ID;
  var ids = sh.getRange(2, col, last - 1, 1).getValues();
  var prefix = Schema.amazonOrderPrefix + AMAZON.testIdPrefix;
  var exact = {};
  AMAZON.testOrderIds.forEach(function (id) { exact[Schema.amazonOrderPrefix + id] = true; });
  var removed = 0;
  for (var i = ids.length - 1; i >= 0; i--) {          // bottom-up so row numbers hold
    var v = String(ids[i][0] || "").trim().toUpperCase();
    if (v.indexOf(prefix) === 0 || exact[v]) {
      sh.deleteRow(i + 2); removed++;
    }
  }
  try { _dashBustTickCache(); } catch (e) {}
  return "✅ Removed " + removed + " Activity Log row(s) for test orders (" + prefix + "… and " +
         Object.keys(exact).join(", ") + ").";
}


// =======================================================================================
// EDITOR TEST WRAPPERS — output goes to the EXECUTION LOG
// =======================================================================================

/** Safe: previews a test order, writes nothing. */
function previewAmazonOrderNow() {
  var out = previewAmazonOrder("000-0000000-0000001", [{ sku: "167517", qty: 1 }], "tomorrow", "test");
  console.log(JSON.stringify(out, null, 2));
  return out;
}
