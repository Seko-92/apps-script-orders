/**
 * SkuEnrichment.js — turns each SKU cell into a clickable LISTING LINK (a
 * rich-text link to the item's eBay viewItemURL), looked up LIVE from Master
 * Inventory by SKU. Hovering/clicking the link shows Google's native preview
 * card — which renders the listing's TITLE + PHOTO + URL, so a separate title
 * note is unnecessary (it was dropped 2026-06-03 — it duplicated the card's
 * title and overlapped it). No new columns, no cache: the lookup reads MI on
 * the spot. viewItemURL is effectively immutable per listing, so there is
 * nothing to go stale (qty/stock live elsewhere and stay MI-synced via n8n).
 *
 * Same join key + same wiring as the location auto-fill and the ▣ kit marker:
 *   - typing a SKU in col A → skuEnrichmentOnEdit (Main.onEditInstallable)
 *   - programmatic inserts (n8n doPost, Zoho pull, kit expansion) call
 *     refreshSkuEnrichment() at their insert sites (setValues doesn't fire onEdit)
 *   - the sort carries the links + notes with their rows (see sortTable...)
 *
 * Fallback: a SKU not found in MI (e.g. a never-listed part) gets no title note,
 * and its link points at an eBay SEARCH for that SKU — i.e. the picker's manual
 * "copy SKU → paste into eBay" workflow, pre-built into a single click. Works
 * identically for eBay-table and DIRECT-table rows since the SKU is the same.
 */

/** eBay search URL for a bare SKU — the manual workflow, one click. */
function _ebaySearchUrl(sku) {
  return 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(String(sku).trim());
}

// On-brand link styling: brand ink (black, NOT default link-blue) + a thin
// underline as the subtle "this is clickable" cue. Adds no foreign color to the
// table's monochrome design. Hover/click still pops Google's preview card.
var _SKU_LINK_STYLE = SpreadsheetApp.newTextStyle()
  .setForegroundColor('#1d1d1b')
  .setUnderline(true)
  .build();

/** Build the RichTextValue (linked SKU) for one cell. */
function _skuRichText(rawSku, rec) {
  var url = (rec && rec.url) ? rec.url : _ebaySearchUrl(rawSku);
  return SpreadsheetApp.newRichTextValue()
    .setText(rawSku)
    .setLinkUrl(url)
    .setTextStyle(_SKU_LINK_STYLE)
    .build();
}

/**
 * Build SKU(lower) → { title, url } from Master Inventory in one read.
 * `url` is '' when viewItemURL is blank (caller substitutes the eBay-search URL).
 * Returns an empty Map if MI / its headers are missing (graceful degradation).
 */
function buildSkuEnrichmentMap() {
  var map = new Map();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var db = ss.getSheetByName(DB_SHEET_NAME);
  if (!db) return map;

  var data = db.getDataRange().getValues();
  if (!data.length) return map;
  var headers = data[0];
  var skuCol   = headers.indexOf(DB_SKU_HEADER);
  var titleCol = headers.indexOf(DB_TITLE_HEADER);
  var urlCol   = headers.indexOf(DB_VIEWURL_HEADER);
  if (skuCol === -1) return map;

  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][skuCol] || '').trim().toLowerCase();
    if (!sku) continue;
    map.set(sku, {
      title: titleCol !== -1 ? String(data[i][titleCol] || '').trim() : '',
      url:   urlCol   !== -1 ? String(data[i][urlCol]   || '').trim() : ''
    });
  }
  return map;
}

/** Single-SKU lookup (live MI scan). Returns { title, url } or null. */
function getSingleSkuEnrichment(skuLower) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var db = ss.getSheetByName(DB_SHEET_NAME);
  if (!db) return null;
  var data = db.getDataRange().getValues();
  if (!data.length) return null;
  var headers = data[0];
  var skuCol   = headers.indexOf(DB_SKU_HEADER);
  var titleCol = headers.indexOf(DB_TITLE_HEADER);
  var urlCol   = headers.indexOf(DB_VIEWURL_HEADER);
  if (skuCol === -1) return null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][skuCol] || '').trim().toLowerCase() === skuLower) {
      return {
        title: titleCol !== -1 ? String(data[i][titleCol] || '').trim() : '',
        url:   urlCol   !== -1 ? String(data[i][urlCol]   || '').trim() : ''
      };
    }
  }
  return null;
}

/**
 * Bulk (re)apply clickable links to every data row's SKU cell (and clear any
 * old title notes left from earlier testing).
 * Use as the one-time backfill (editor Run) AND as the post-insert refresh at
 * programmatic insert sites. Boundary / header / empty cells are left untouched.
 * Batched: 2 writes total (setRichTextValues + setNotes) regardless of row count.
 */
/**
 * GENERIC core — apply SKU→listing links to a SKU column in ANY sheet
 * (All Orders, Prep Queue, Out of Stock, ...). Skips empty cells, the DIRECT
 * boundary marker, and header-glyph (◈) rows. Also clears each cell's note.
 * Batched: 2 writes total regardless of row count.
 * @return {number} rows linked
 */
function applySkuLinksToColumn(sheet, skuCol, startRow, endRow, map) {
  if (!sheet || endRow < startRow) return 0;
  var n = endRow - startRow + 1;
  var range = sheet.getRange(startRow, skuCol, n, 1);
  var values = range.getValues();
  var rich = range.getRichTextValues();   // preserve non-owned cells as-is
  var blankNotes = [];                     // title note was dropped — clear notes
  var linked = 0;

  for (var i = 0; i < n; i++) {
    blankNotes.push(['']);
    var raw = String(values[i][0] || '').trim();
    if (!raw) continue;                                       // empty — no link
    if (raw.toUpperCase() === Schema.boundaryMarker) continue; // DIRECT divider — leave
    if (raw.charAt(0) === '◈') continue;                      // header glyph — leave

    var rec = map.get(raw.toLowerCase()) || null;
    rich[i][0] = _skuRichText(raw, rec);
    linked++;
  }

  range.setRichTextValues(rich);
  range.setNotes(blankNotes);
  return linked;
}

/** Bulk backfill / re-apply links across the All Orders sheet (both tables). */
function refreshSkuEnrichment() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ Main sheet not found';

  var startRow = Schema.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return '✅ No data rows to enrich.';

  var map = buildSkuEnrichmentMap();
  if (!map.size) return '❌ Master Inventory unavailable (no title/URL data).';

  var linked = applySkuLinksToColumn(sheet, Schema.cols.SKU, startRow, lastRow, map);
  SpreadsheetApp.flush();
  return '✅ SKU links refreshed — ' + linked + ' row(s) linked.';
}

/** Bulk backfill / re-apply links across the Prep Queue sheet. */
function refreshPrepQueueSkuLinks() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PREP_QUEUE.sheetName);
  if (!sheet) return 'ℹ️ Prep Queue sheet not found.';

  var lastRow = sheet.getLastRow();
  if (lastRow < PREP_QUEUE.dataStartRow) return 'ℹ️ Prep Queue empty.';

  var map = buildSkuEnrichmentMap();
  if (!map.size) return '❌ Master Inventory unavailable (no title/URL data).';

  var linked = applySkuLinksToColumn(sheet, PREP_QUEUE.cols.SKU, PREP_QUEUE.dataStartRow, lastRow, map);
  SpreadsheetApp.flush();
  return '✅ Prep Queue SKU links refreshed — ' + linked + ' row(s) linked.';
}

/**
 * Per-edit handler — col-A SKU edit → live MI lookup → clickable link (and
 * clears any old col-A note, incl. when a SKU is removed/changed).
 * Dispatched from Main.onEditInstallable next to kitSkuOnEdit. Single-cell and
 * multi-row paste supported (one MI read covers the whole edited range).
 * Best-effort; wrapped in try/catch upstream so any error stays contained.
 */
function skuEnrichmentOnEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== MAIN_SHEET_NAME) return;

  // Only react to col-A-only edits (single cell or multi-row paste within col A)
  var firstCol = e.range.getColumn();
  var lastCol  = firstCol + e.range.getNumColumns() - 1;
  if (firstCol !== Schema.cols.SKU || lastCol !== Schema.cols.SKU) return;
  if (e.range.getRow() < Schema.dataStartRow) return;

  var values = e.range.getValues();
  var rich = e.range.getRichTextValues();
  // Always clear the col-A note on edited cells: no title note anymore, and this
  // also wipes the stale note when a SKU is removed/changed.
  var blankNotes = [];
  var map = buildSkuEnrichmentMap();

  for (var i = 0; i < values.length; i++) {
    blankNotes.push(['']);
    var raw = String(values[i][0] || '').trim();
    var upper = raw.toUpperCase();
    if (!raw) continue;                                       // cleared SKU — link cleared with the value
    if (upper === Schema.boundaryMarker) continue;
    if (raw.charAt(0) === '◈') continue;

    var rec = map.get(raw.toLowerCase()) || null;
    rich[i][0] = _skuRichText(raw, rec);
  }

  e.range.setRichTextValues(rich);
  e.range.setNotes(blankNotes);
}
