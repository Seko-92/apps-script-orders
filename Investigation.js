// =======================================================================================
// Investigation.js — Order Case module: one-screen dossier + permanent documentation
// Shipped 2026-07-28.
// =======================================================================================
//
// WHY
//   When an order goes sideways (missing part, wrong item, buyer says "not
//   delivered"), staff had to Ctrl+F two sheets AND jump across eBay + Zoho +
//   the printed label to reconstruct the story — and it left NO record, so the
//   next person re-does the whole dig. This module gives that investigation
//   ONE roomy home (a modal, like Kit Expansion) and, critically, a place to
//   WRITE DOWN what was found so the order carries its own memory.
//
// TWO HALVES
//   THE STORY (read)  — the consolidated dossier: All Orders rows + the full
//                       Activity Log timeline (both via lookupOrder, unchanged)
//                       + one-click jump-outs to the eBay order, the Zoho SO,
//                       and the print record (URLs reused from OrderLinks.js).
//   THE CASE FILE (write) — human documentation logged to a NEW permanent
//                       "Investigations" sheet: Category · Findings · Resolution
//                       · Status (Open→Resolved) · Investigator. NOT purged
//                       (unlike the 90-day Activity Log) — documentation must
//                       outlive the timeline.
//
// SCOPE (v1, deliberately lean — Tier 1, rides on what we already built):
//   - Category is an IN-HOUSE human classification. It is NOT pulled from eBay.
//     Auto-ingesting eBay buyer cases (INR / returns via the Post-Order API) is
//     a parked Tier-3 idea — a separate external dependency, not this build.
//   - No live eBay/Zoho status reconciliation inside the modal (the jump-outs
//     get you there). That "side-by-side, mismatch lights up red" view is the
//     parked v2.
//   - No label IMAGE — we link to the eBay order (where the label lives) and
//     surface the PRINTED event; we never captured the PDF.
//
// DEPLOY: editor-bound (sidebar google.script.run + a showModalDialog modal).
//   `clasp push` is the whole deploy — nothing here touches /exec, no New Version.
// =======================================================================================


// ---------- LOCAL SCHEMA ----------
var INVESTIGATIONS = {
  sheetName: "Investigations",

  cols: {
    TIMESTAMP:    1,   // A — real Date
    ORDER_ID:     2,   // B
    CATEGORY:     3,   // C
    FINDINGS:     4,   // D — free text: what was wrong / discovered
    RESOLUTION:   5,   // E — free text: how it was resolved (may be blank while Open)
    STATUS:       6,   // F — Open / Resolved
    INVESTIGATOR: 7    // G — Pick ID for Shipping at time of note (blank = supervisor/unknown)
  },

  idx: function (name) { return INVESTIGATIONS.cols[name] - 1; },

  dataWidth:    7,
  headerRow:    1,
  dataStartRow: 2,

  headers: ["⏱ TIMESTAMP", "ORDER ID", "CATEGORY", "FINDINGS", "RESOLUTION", "STATUS", "👤 INVESTIGATOR"],

  // Human classification of WHY we're investigating. Covers both internally-found
  // ("Wrong item", "Missing") and buyer-triggered ("Not received", "Return/refund")
  // reasons — all logged BY HAND (nothing is auto-pulled from eBay in v1).
  categories: ["Not received", "Wrong item", "Missing item/part", "Damaged",
               "Shipping issue", "Return/refund", "Other"],

  statuses: ["Open", "Resolved"]
};


// =======================================================================================
// SETUP — idempotent (mirrors the other module sheets)
// =======================================================================================

function setupInvestigationsSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
  if (!sheet) sheet = ss.insertSheet(INVESTIGATIONS.sheetName);

  // --- HEADERS ---
  sheet.getRange(INVESTIGATIONS.headerRow, 1, 1, INVESTIGATIONS.dataWidth)
    .setValues([INVESTIGATIONS.headers])
    .setBackground('#1d1d1b')
    .setFontColor('#ffd966')
    .setFontFamily('Oswald')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(INVESTIGATIONS.headerRow, 1, 1, INVESTIGATIONS.dataWidth)
    .setBorder(null, null, true, null, null, null,
               '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(INVESTIGATIONS.cols.TIMESTAMP,    150);
  sheet.setColumnWidth(INVESTIGATIONS.cols.ORDER_ID,     150);
  sheet.setColumnWidth(INVESTIGATIONS.cols.CATEGORY,     140);
  sheet.setColumnWidth(INVESTIGATIONS.cols.FINDINGS,     360);
  sheet.setColumnWidth(INVESTIGATIONS.cols.RESOLUTION,   360);
  sheet.setColumnWidth(INVESTIGATIONS.cols.STATUS,       100);
  sheet.setColumnWidth(INVESTIGATIONS.cols.INVESTIGATOR, 130);

  // --- DATA AREA FORMATS ---
  var maxDataRow = 4000;
  var dsr = INVESTIGATIONS.dataStartRow;
  var dataRows = maxDataRow - dsr + 1;

  sheet.getRange(dsr, INVESTIGATIONS.cols.TIMESTAMP, dataRows, 1)
    .setNumberFormat('M/d/yy h:mm am/pm')
    .setFontFamily('Roboto Mono').setFontSize(9).setFontColor('#5f5f5f').setHorizontalAlignment('center');
  sheet.getRange(dsr, INVESTIGATIONS.cols.ORDER_ID, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, INVESTIGATIONS.cols.CATEGORY, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  sheet.getRange(dsr, INVESTIGATIONS.cols.FINDINGS, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(dsr, INVESTIGATIONS.cols.RESOLUTION, dataRows, 1)
    .setFontFamily('Roboto').setFontSize(10).setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(dsr, INVESTIGATIONS.cols.STATUS, dataRows, 1)
    .setFontFamily('Oswald').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(dsr, INVESTIGATIONS.cols.INVESTIGATOR, dataRows, 1)
    .setFontFamily('Roboto Mono').setFontSize(9).setHorizontalAlignment('center');

  sheet.getRange(dsr, 1, dataRows, INVESTIGATIONS.dataWidth).setVerticalAlignment('middle');

  // --- DATA VALIDATION (so hand-edits on the sheet stay consistent with the modal) ---
  sheet.getRange(dsr, INVESTIGATIONS.cols.CATEGORY, dataRows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(INVESTIGATIONS.categories, true).setAllowInvalid(true).build());
  sheet.getRange(dsr, INVESTIGATIONS.cols.STATUS, dataRows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(INVESTIGATIONS.statuses, true).setAllowInvalid(true).build());

  // --- BANDING ---
  sheet.getBandings().forEach(function (b) { try { b.remove(); } catch (e) {} });
  var band = sheet.getRange(1, 1, maxDataRow, INVESTIGATIONS.dataWidth)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor('#1d1d1b').setFirstRowColor('#ffffff').setSecondRowColor('#fff8e7');

  // --- CONDITIONAL FORMATTING on STATUS ---
  var existing = sheet.getConditionalFormatRules() || [];
  var keep = existing.filter(function (r) {
    var ranges = r.getRanges();
    if (!ranges || ranges.length === 0) return true;
    return !ranges.some(function (rg) {
      return rg.getSheet().getName() === INVESTIGATIONS.sheetName
          && rg.getColumn() === INVESTIGATIONS.cols.STATUS;
    });
  });
  var statusRange = sheet.getRange(dsr, INVESTIGATIONS.cols.STATUS, dataRows, 1);
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Open")
    .setBackground('#fff4b0').setFontColor('#7d5d00').setBold(true)
    .setRanges([statusRange]).build());
  keep.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Resolved")
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setBold(true)
    .setRanges([statusRange]).build());
  sheet.setConditionalFormatRules(keep);

  sheet.setFrozenRows(1);

  return "✅ Investigations sheet ready.";
}


/** Sidebar: switch view to the Investigations sheet (create if missing). */
function openInvestigations() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) return "❌ No active spreadsheet";
  var sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
  if (!sheet) {
    setupInvestigationsSheet();
    sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
  }
  ss.setActiveSheet(sheet);
  return "✅ Opened Investigations";
}


// =======================================================================================
// CASE LINKS — reuse OrderLinks.js builders to derive the jump-out URLs
// =======================================================================================
//
// From the query + the matched All Orders rows, resolve the eBay order URL and
// the Zoho SO URL (deep-link if we have the salesorder_id in the Pending cache,
// else a Zoho search). Reuses _ebayOrderUrl / _zohoSoUrl / _zohoSoSearchUrl /
// buildZohoSoIdMap / _isEbayOrderId / _isZohoSo from OrderLinks.js.
// =======================================================================================

function _caseLinks(query, rows) {
  var out = { ebay: null, zoho: null, zohoLabel: "" };

  var candidates = [String(query || "")];
  (rows || []).forEach(function (r) { if (r && r.salesOrder) candidates.push(String(r.salesOrder)); });

  var zohoIdMap = null;   // lazily built only if a SO shows up

  for (var i = 0; i < candidates.length; i++) {
    var v = String(candidates[i] || "").trim();
    if (!v) continue;

    // eBay order id (exact, or embedded in "Replacement #: …" text)
    if (!out.ebay) {
      if (typeof _isEbayOrderId === 'function' && _isEbayOrderId(v)) {
        out.ebay = _ebayOrderUrl(v);
      } else {
        var em = v.match(/\d{2,3}-\d{4,6}-\d{4,6}/);
        if (em && typeof _ebayOrderUrl === 'function') out.ebay = _ebayOrderUrl(em[0]);
      }
    }

    // Zoho SO (exact SO-… or embedded)
    if (!out.zoho) {
      var som = v.match(/SO-\d+/i);
      var so = (typeof _isZohoSo === 'function' && _isZohoSo(v)) ? v : (som ? som[0] : null);
      if (so) {
        if (zohoIdMap === null) {
          try { zohoIdMap = buildZohoSoIdMap(); } catch (e) { zohoIdMap = new Map(); }
        }
        var id = zohoIdMap.get(so.toUpperCase());
        out.zoho = id ? _zohoSoUrl(id) : _zohoSoSearchUrl(so);
        out.zohoLabel = so;
      }
    }
  }
  return out;
}


// =======================================================================================
// ZOHO STATUS — the authoritative direct-order status, FREE (Tier 1)
// =======================================================================================
//
// For direct/Zoho orders we already mirror Zoho's order/payment/shipment status
// into the Pending Sales Orders sheet on every webhook (ZohoSalesOrders.js). So
// the console can show "what Zoho says" with a plain SHEET READ — no API call,
// no new dependency. Returns null for eBay orders (not in Pending) or when the
// sheet is missing. Freshness = the last Zoho webhook (surfaced via updatedMs).
//
// (The eBay live-status pull — the full sheet-vs-eBay reconciliation — is a
// deliberately PARKED Tier-3 build; it needs a live Fulfillment API call.)
// =======================================================================================

function getZohoStatusForOrder(query) {
  try {
    var q = String(query || "").trim();
    if (!q) return null;
    if (typeof PENDING_SO === 'undefined') return null;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PENDING_SO.sheetName);
    if (!sheet) return null;

    var row = _resolvePendingRow(sheet, q);   // handles SO# and INV# (ZohoSalesOrders.js)
    if (row < 1) return null;

    var C = PENDING_SO.cols;
    var vals = sheet.getRange(row, 1, 1, C.INVOICE).getValues()[0];   // A..M covers all we read
    var updated = vals[C.LAST_UPDATED - 1];

    return {
      soNumber:    String(vals[C.SO_NUMBER - 1]    || ""),
      customer:    String(vals[C.CUSTOMER - 1]     || ""),
      orderStatus: String(vals[C.ORDER_STATUS - 1] || ""),
      payment:     String(vals[C.PAYMENT - 1]      || ""),
      shipment:    String(vals[C.SHIPMENT - 1]     || ""),
      total:       String(vals[C.TOTAL - 1]        || ""),
      invoice:     String(vals[C.INVOICE - 1]      || ""),
      pulled:      String(vals[C.PULLED - 1] || "").trim().toUpperCase() === PENDING_SO.pulledFlag,
      updatedMs:   (updated instanceof Date) ? updated.getTime() : null
    };
  } catch (e) {
    try { console.log("getZohoStatusForOrder error: " + e); } catch (_) {}
    return null;
  }
}


// =======================================================================================
// READ — investigation notes for an order
// =======================================================================================

/**
 * All case notes for an order (normalized substring match on ORDER_ID),
 * newest-first. Returns [] when the sheet is missing.
 */
function getInvestigationNotes(orderIdOrNormalized) {
  var normalized = _normalizeOrderId(orderIdOrNormalized);
  if (!normalized) return [];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < INVESTIGATIONS.dataStartRow) return [];

  var n = lastRow - INVESTIGATIONS.dataStartRow + 1;
  var data = sheet.getRange(INVESTIGATIONS.dataStartRow, 1, n, INVESTIGATIONS.dataWidth).getValues();

  var TS = INVESTIGATIONS.idx("TIMESTAMP"), OID = INVESTIGATIONS.idx("ORDER_ID"),
      CAT = INVESTIGATIONS.idx("CATEGORY"), FIN = INVESTIGATIONS.idx("FINDINGS"),
      RES = INVESTIGATIONS.idx("RESOLUTION"), ST = INVESTIGATIONS.idx("STATUS"),
      INV = INVESTIGATIONS.idx("INVESTIGATOR");

  var notes = [];
  for (var i = 0; i < data.length; i++) {
    var oid = _normalizeOrderId(data[i][OID]);
    if (!oid || oid.indexOf(normalized) === -1) continue;
    var ts = data[i][TS];
    notes.push({
      timestamp:    (ts instanceof Date) ? ts.getTime() : null,
      orderId:      String(data[i][OID]  || ""),
      category:     String(data[i][CAT]  || ""),
      findings:     String(data[i][FIN]  || ""),
      resolution:   String(data[i][RES]  || ""),
      status:       String(data[i][ST]   || ""),
      investigator: String(data[i][INV]  || "")
    });
  }
  notes.sort(function (a, b) { return (b.timestamp || 0) - (a.timestamp || 0); });   // newest first
  return notes;
}

/**
 * Count of OPEN cases for one order = 1 if the order's MOST RECENT note is
 * "Open", else 0 (an order has one live case state). Cheap; used by the sidebar
 * lookup's "⚠ open case" flag. Returns 0 defensively.
 */
function getOpenCasesForOrder(orderIdOrNormalized) {
  try {
    var notes = getInvestigationNotes(orderIdOrNormalized);   // newest first
    if (!notes.length) return 0;
    return (String(notes[0].status).trim().toLowerCase() === "open") ? 1 : 0;
  } catch (e) {
    return 0;
  }
}

/**
 * Catalog-wide open-case count = number of DISTINCT orders whose latest note is
 * "Open". Powers the Alerts badge — a cheap snapshot read, safe on the 30s poll.
 */
function getOpenCaseCount() {
  try {
    var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
    if (!sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < INVESTIGATIONS.dataStartRow) return 0;

    var n = lastRow - INVESTIGATIONS.dataStartRow + 1;
    var data = sheet.getRange(INVESTIGATIONS.dataStartRow, 1, n, INVESTIGATIONS.dataWidth).getValues();
    var TS = INVESTIGATIONS.idx("TIMESTAMP"), OID = INVESTIGATIONS.idx("ORDER_ID"), ST = INVESTIGATIONS.idx("STATUS");

    // Track the latest note (by timestamp) per normalized order id.
    var latest = {};   // normOrderId → {ts, status}
    for (var i = 0; i < data.length; i++) {
      var oid = _normalizeOrderId(data[i][OID]);
      if (!oid) continue;
      var ts = (data[i][TS] instanceof Date) ? data[i][TS].getTime() : 0;
      var status = String(data[i][ST] || "").trim().toLowerCase();
      if (!latest[oid] || ts >= latest[oid].ts) latest[oid] = { ts: ts, status: status };
    }
    var count = 0;
    for (var k in latest) { if (latest[k].status === "open") count++; }
    return count;
  } catch (e) {
    console.log("getOpenCaseCount error: " + e);
    return 0;
  }
}


// =======================================================================================
// WRITE — log a case note (the documentation action)
// =======================================================================================

/**
 * Append one investigation note. Called by the modal's Save button.
 *
 * payload = { orderId, category, findings, resolution, status }
 *   - orderId + findings are required (a note with no findings is meaningless).
 *   - category defaults to "Other" if unrecognized; status defaults to "Open".
 *   - investigator + timestamp are captured server-side (never trusted from the client).
 *
 * @returns {{ ok, note }|{ ok:false, reason }}  note = the saved row echo for optimistic render.
 */
function logInvestigationNote(payload) {
  try {
    payload = payload || {};
    var orderId  = String(payload.orderId  || "").trim();
    var findings = String(payload.findings || "").trim();
    if (!orderId)  return { ok: false, reason: "Missing order ID." };
    if (!findings) return { ok: false, reason: "Add at least a finding before saving." };

    var category = String(payload.category || "").trim();
    if (INVESTIGATIONS.categories.indexOf(category) === -1) category = "Other";

    var status = String(payload.status || "").trim();
    if (INVESTIGATIONS.statuses.indexOf(status) === -1) status = "Open";

    var resolution = String(payload.resolution || "").trim();
    var investigator = "";
    try { investigator = (typeof getCurrentPicker === 'function') ? String(getCurrentPicker() || "") : ""; } catch (e) {}

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(INVESTIGATIONS.sheetName);
    if (!sheet) { setupInvestigationsSheet(); sheet = ss.getSheetByName(INVESTIGATIONS.sheetName); }

    var now = new Date();
    var row = [now, orderId, category, findings, resolution, status, investigator];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, INVESTIGATIONS.dataWidth).setValues([row]);
    SpreadsheetApp.flush();

    return {
      ok: true,
      note: {
        timestamp:    now.getTime(),
        orderId:      orderId,
        category:     category,
        findings:     findings,
        resolution:   resolution,
        status:       status,
        investigator: investigator
      }
    };
  } catch (err) {
    try { console.log("logInvestigationNote error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


// =======================================================================================
// MODAL — open the Order Case dossier
// =======================================================================================

/**
 * Build the dossier object for an order — rows + timeline + jump-out links +
 * prior case notes + investigator + the category/status option lists. Returns a
 * populated object even when the order isn't found (found:false, empty arrays)
 * so the modal can render an empty state AND let the user search again in-window.
 */
function _buildOrderCaseDossier(raw) {
  var lookup = lookupOrder(raw);                         // rows + events + summary (OrderLookup.js)
  var notes  = getInvestigationNotes(raw);
  var links  = _caseLinks(raw, lookup.rows);

  var picker = "";
  try { picker = (typeof getCurrentPicker === 'function') ? String(getCurrentPicker() || "") : ""; } catch (e) {}

  return {
    orderId:      raw,
    found:        lookup.found,
    channel:      (lookup.rows && lookup.rows.length) ? lookup.rows[0].table : "",
    statuses:     (lookup.summary && lookup.summary.statuses) || [],
    rows:         lookup.rows   || [],
    events:       lookup.events || [],
    notes:        notes,
    links:        links,
    zoho:         getZohoStatusForOrder(raw),   // null for eBay orders — the FREE authoritative direct status
    investigator: picker,
    categories:   INVESTIGATIONS.categories,
    statusOptions: INVESTIGATIONS.statuses
  };
}

/**
 * Modal → server: fetch a dossier for a NEW order without closing the window.
 * Powers the in-modal search bar (the "look up another order in the same window"
 * flow). Returns { ok, dossier } or { ok:false, reason }.
 */
function getOrderCaseData(query) {
  try {
    var raw = String(query == null ? "" : query).trim();
    if (!raw) return { ok: false, reason: "Type an order ID." };
    return { ok: true, dossier: _buildOrderCaseDossier(raw) };
  } catch (err) {
    try { console.log("getOrderCaseData error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}

/**
 * Open the Order Case console (the modal). ALWAYS opens on a non-empty query —
 * even a not-found id opens the console (empty state + a working search bar) so
 * the user can retry in-window rather than bouncing back to the sidebar. The
 * dossier is pre-injected so the first render is instant; subsequent searches
 * re-fetch via getOrderCaseData and re-render in place.
 */
function openOrderCase(query) {
  try {
    var raw = String(query == null ? "" : query).trim();
    if (!raw) return { ok: false, reason: "Type an order ID first." };

    var dossier = _buildOrderCaseDossier(raw);

    var template = HtmlService.createTemplateFromFile("OrderCaseModal");
    // <?!= ?> force-unescaped JSON injection (Gotcha #2) + </-guard.
    template.dossierJson = JSON.stringify(dossier).replace(/<\//g, "<\\/");

    var html = template.evaluate().setWidth(1080).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, "🔍 Order Case");
    return { ok: true, found: dossier.found, rows: dossier.rows.length,
             events: dossier.events.length, notes: dossier.notes.length };
  } catch (err) {
    try { console.log("openOrderCase error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


/** EDITOR-RUN test: dump a case dossier's data (no modal). */
function testOrderCase(orderId) {
  var raw = orderId || "";
  var lookup = lookupOrder(raw);
  var out = { found: lookup.found, rows: lookup.rows.length, events: lookup.events.length,
              notes: getInvestigationNotes(raw).length, links: _caseLinks(raw, lookup.rows) };
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}
