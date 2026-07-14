/**
 * OrderLinks.js — turns the SALES ORDER cell (col D) into a clickable link that
 * opens the actual order, killing the "copy order # → go find it" round-trip.
 *
 *   eBay order  (XX-XXXXX-XXXXX)  → eBay Seller Hub order details, built DIRECTLY
 *                                   from the id (no lookup needed).
 *   Zoho SO     (SO-…)            → the Zoho Inventory sales order, via the
 *                                   salesorder_id joined from the Pending Sales
 *                                   Orders payload cache. Fallback (manually-typed
 *                                   SOs not in Pending): a Zoho SO search for that
 *                                   number — still one click.
 *   Anything else (Replacement #:, INV-, free text) → left plain (no link).
 *
 * Same mechanics + wiring as the SKU link (SkuEnrichment.js): live read, no
 * cache, brand-ink styling, applied on edit + at insert sites, and carried
 * through the sort. URL patterns verified live 2026-06-04.
 */

// Fixed Zoho Inventory org for this account (appears in every Zoho URL — not a
// secret). Same value recorded in CLAUDE.md.
var ZOHO_ORG_ID = '803368514';

/* ── URL builders ──────────────────────────────────────────────────────── */
function _ebayOrderUrl(orderId) {
  return 'https://www.ebay.com/mesh/ord/details?mode=SH&orderid='
       + encodeURIComponent(orderId) + '&source=Orders';
}
function _zohoSoUrl(salesorderId) {
  return 'https://inventory.zoho.com/app/' + ZOHO_ORG_ID
       + '#/salesorders/' + encodeURIComponent(salesorderId);
}
function _zohoSoSearchUrl(soNumber) {
  return 'https://inventory.zoho.com/app/' + ZOHO_ORG_ID
       + '#/salesorders?per_page=200&search_criteria='
       + encodeURIComponent(JSON.stringify({ search_text: soNumber }));
}

/* ── Classifiers ───────────────────────────────────────────────────────── */
// eBay order id: digits + dashes only, contains a dash (same convention as the
// S4 shipped-check, so "Replacement #: …" free text is correctly excluded).
function _isEbayOrderId(v) { return /^[\d-]+$/.test(v) && v.indexOf('-') !== -1; }
function _isZohoSo(v)      { return /^SO-/i.test(v); }

// On-brand link styling: brand ink (not link-blue) + a thin underline as the
// subtle "this is clickable" cue. Matches the SKU link.
// setFontSize(10) is LOAD-BEARING (2026-07-14): on SO-badge rows the CELL
// font is raised to 14 so the badge glyph (a number-format prefix, which
// renders at the cell's default font) reads large — the ID text must stay
// pinned at the table's 10px via its run style or it would inherit the 14.
var _ORDER_LINK_STYLE = SpreadsheetApp.newTextStyle()
  .setForegroundColor('#1d1d1b')
  .setUnderline(true)
  .setFontSize(10)
  .build();

/**
 * Build SO#(upper) → salesorder_id from the Pending Sales Orders payload cache.
 * Returns an empty Map if the sheet is missing. Only built when a Zoho SO is
 * actually present in the range being linked (eBay rows don't need it).
 */
function buildZohoSoIdMap() {
  var map = new Map();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PENDING_SO.sheetName);
  if (!sheet) return map;
  var lastRow = sheet.getLastRow();
  if (lastRow < PENDING_SO.dataStartRow) return map;

  var n = lastRow - PENDING_SO.dataStartRow + 1;
  var soVals = sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.SO_NUMBER, n, 1).getValues();
  var plVals = sheet.getRange(PENDING_SO.dataStartRow, PENDING_SO.cols.PAYLOAD,   n, 1).getValues();
  for (var i = 0; i < n; i++) {
    var so = String(soVals[i][0] || '').trim().toUpperCase();
    if (!so) continue;
    var raw = String(plVals[i][0] || '');
    if (!raw) continue;
    try {
      var id = String(JSON.parse(raw).salesorder_id || '').trim();
      if (id) map.set(so, id);
    } catch (e) { /* corrupted payload — skip */ }
  }
  return map;
}

/**
 * Build the linked RichTextValue for one SALES ORDER value, or null if the value
 * isn't a linkable order (leave the cell as-is).
 * @param {string} rawVal  the cell's SALES ORDER text
 * @param {Map} zohoIdMap  SO#(upper) → salesorder_id (may be null/empty)
 */
function _orderRichText(rawVal, zohoIdMap) {
  var v = String(rawVal || '').trim();
  if (!v) return null;

  var url = null;
  if (_isEbayOrderId(v)) {
    url = _ebayOrderUrl(v);
  } else if (_isZohoSo(v)) {
    var id = zohoIdMap ? zohoIdMap.get(v.toUpperCase()) : null;
    url = id ? _zohoSoUrl(id) : _zohoSoSearchUrl(v);   // deep-link if we have the id, else search
  } else {
    // Free-text row (e.g. "Replacement #: 19-14597-26309") — extract an embedded
    // eBay order id or SO# and link the whole cell to it, keeping the text as-is.
    var em = v.match(/\d{2,3}-\d{4,6}-\d{4,6}/);   // embedded eBay order id
    if (em) {
      url = _ebayOrderUrl(em[0]);
    } else {
      var sm = v.match(/SO-\d+/i);                 // embedded Zoho SO#
      if (sm) {
        var sid = zohoIdMap ? zohoIdMap.get(sm[0].toUpperCase()) : null;
        url = sid ? _zohoSoUrl(sid) : _zohoSoSearchUrl(sm[0]);
      }
    }
    if (!url) return null;   // no id found anywhere — leave plain
  }

  return SpreadsheetApp.newRichTextValue()
    .setText(v)
    .setLinkUrl(url)
    .setTextStyle(_ORDER_LINK_STYLE)
    .build();
}

/**
 * GENERIC core — apply order links to a SALES ORDER column on any sheet.
 * Skips the boundary marker + header-glyph rows. Cells that aren't a linkable
 * order are left untouched (their existing rich text is preserved).
 * @return {number} rows linked
 */
function applyOrderLinksToColumn(sheet, soCol, startRow, endRow, zohoIdMap) {
  if (!sheet || endRow < startRow) return 0;
  var n = endRow - startRow + 1;
  var range = sheet.getRange(startRow, soCol, n, 1);
  var values = range.getValues();
  var rich = range.getRichTextValues();   // preserve non-linkable cells as-is
  var linked = 0;

  for (var i = 0; i < n; i++) {
    var raw = String(values[i][0] || '').trim();
    if (!raw) continue;
    if (raw.toUpperCase() === Schema.boundaryMarker) continue;
    if (raw.charAt(0) === '◈') continue;

    var rtv = _orderRichText(raw, zohoIdMap);
    if (rtv) { rich[i][0] = rtv; linked++; }
  }

  range.setRichTextValues(rich);
  return linked;
}

/** Bulk backfill / re-apply order links across the All Orders SALES ORDER column. */
function refreshOrderLinks() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ Main sheet not found';

  var startRow = Schema.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '✅ No data rows for order links.';

  var zohoIdMap = buildZohoSoIdMap();   // cheap; covers the SO-… rows
  var linked = applyOrderLinksToColumn(sheet, Schema.cols.SALES_ORDER, startRow, lastRow, zohoIdMap);
  SpreadsheetApp.flush();
  return '✅ Order links refreshed — ' + linked + ' order(s) linked.';
}

/**
 * Combined backfill — SKU links + order links in one call. Used by the sidebar
 * button and the programmatic insert sites so one refresh covers both columns.
 */
function refreshAllOrdersEnrichment() {
  var sku = refreshSkuEnrichment();
  var ord = refreshOrderLinks();
  return sku + '  ·  ' + ord;
}

/**
 * Per-edit handler — SALES ORDER (col D) edit → apply the order link. Dispatched
 * from Main.onEditInstallable. Builds the Zoho id map only if a SO-… value is in
 * the edit (eBay rows need no lookup). Best-effort; wrapped in try/catch upstream.
 */
function orderLinkOnEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== MAIN_SHEET_NAME) return;

  var firstCol = e.range.getColumn();
  var lastCol  = firstCol + e.range.getNumColumns() - 1;
  if (firstCol !== Schema.cols.SALES_ORDER || lastCol !== Schema.cols.SALES_ORDER) return;
  if (e.range.getRow() < Schema.dataStartRow) return;

  var values = e.range.getValues();
  var rich = e.range.getRichTextValues();

  // Only pay for the Pending-sheet read if a Zoho SO (exact or embedded in
  // replacement text) is actually in this edit. eBay rows need no lookup.
  var needsZoho = false;
  for (var k = 0; k < values.length; k++) {
    if (/SO-\d+/i.test(String(values[k][0] || ''))) { needsZoho = true; break; }
  }
  var zohoIdMap = needsZoho ? buildZohoSoIdMap() : null;

  for (var i = 0; i < values.length; i++) {
    var rtv = _orderRichText(values[i][0], zohoIdMap);
    if (rtv) rich[i][0] = rtv;
    // non-linkable / cleared cells: leave existing rich text (empty when cleared)
  }

  e.range.setRichTextValues(rich);
}
