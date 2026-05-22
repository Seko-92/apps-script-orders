// =======================================================================================
// VisualLab.js — Design experiments on a duplicate sheet.
//
// Service Bay GRADUATED to production on 2026-05-17. Its implementation now
// lives in BrandTheme.js as the production applyBrandTheme() function, which
// is parameterized by sheet name (defaults to MAIN_SHEET_NAME). The lab's
// testServiceBay() runner just delegates to that — guaranteeing the lab and
// production sheets always render identical Service Bay output.
//
// Telemetry Console stays as a lab-only experiment (dark cockpit aesthetic) —
// distinctive but rejected for production (eBay logo conflict on dark bg,
// print-template misalignment, long-session eye fatigue).
//
// Test runners — call from the Apps Script editor function dropdown:
//   testServiceBay()        → applies production theme to "Copy of All orders"
//   testTelemetry()         → applies Telemetry Console (lab-only) to "Copy of All orders"
//   resetVisualTestSheet()  → strips CF + bandings on test sheet (between A/B compare)
//
// Both runners preserve =IMAGE() in A2 and =TEXT(NOW()…) in B1 — formulas
// are never touched. Only formatting + non-formula cell values are written.
// =======================================================================================

// Visual-Lab color + font tokens. Kept local to this file so production BRAND tokens
// stay untouched.
var VL = {
  // Brand foundation
  yellow:      '#ffd966',   // brand yellow (data-on-dark)
  yellowDeep:  '#ffd400',   // brand action yellow (heavy fills)
  yellowSoft:  '#fff4b0',   // soft yellow surface
  black:       '#1d1d1b',   // brand black (warmer than pure)
  paper:       '#ffffff',
  paperWarm:   '#fff8e7',   // workshop paper cream
  paperWarm2:  '#fdf0d8',   // slightly darker cream — banding pair
  inkSoft:     '#5d5a4a',   // secondary text on cream

  // Service Bay state colors (font-only, no fills)
  goldDark:    '#b8860b',   // PREPARING text on cream
  greenDark:   '#2e7d32',   // SHIPPED text on cream
  graySoft:    '#9b958a',   // CANCELED text on cream (strikethrough)

  // Telemetry palette (dark theme)
  bgCharcoal:  '#1a1a1a',
  bgCharcoal2: '#181818',   // banding pair (very subtle)
  bgBlack:     '#000000',
  rule:        '#2a2a2a',   // scanline divider between rows
  textDim:     '#b8b8b8',   // gray secondary text on dark
  textMuted:   '#666666',
  yellowBright:'#ffeb3b',   // PREPARING on dark
  greenBright: '#66bb6a',   // SHIPPED on dark
  redAlert:    '#ff6b6b',   // low-stock signal on dark

  // Fonts
  display:     'Oswald',
  data:        'Roboto',
  mono:        'Roboto Mono'
};

// =======================================================================================
// TEST RUNNERS — call these from the Apps Script editor function dropdown.
// =======================================================================================

function testServiceBay() {
  return applyServiceBayTheme("Copy of All orders");
}

function testTelemetry() {
  return applyTelemetryTheme("Copy of All orders");
}

function resetVisualTestSheet() {
  return resetVisualLabSheet("Copy of All orders");
}

// =======================================================================================
// THEME A — "THE SERVICE BAY" (graduated to production 2026-05-17)
// Delegates to BrandTheme.applyBrandTheme(sheetName). The production function
// is parameterized to target any sheet, so lab + production share the exact
// same implementation. To customize the test, edit BrandTheme.js, not here.
// =======================================================================================

function applyServiceBayTheme(sheetName) {
  return applyBrandTheme(sheetName);
}

// (Service Bay helpers graduated to BrandTheme.js on 2026-05-17.
// _vlEnsureSparkData / _vlSetServiceBayBannerFormulas / _vlApplyServiceBayDataFormat /
// _vlApplyServiceBayCF were removed from this file — their logic now lives in
// BrandTheme.js as _ensureSparkData / _setSystemPulseBannerFormulas /
// _applyColumnLevelDataFormats / _applyAllConditionalFormatting. To customize
// Service Bay, edit BrandTheme.js; testServiceBay() picks up the changes.)


// =======================================================================================
// THEME B — "TELEMETRY CONSOLE"
// F1 pit-wall live timing aesthetic. Charcoal backgrounds, Roboto Mono everywhere,
// yellow primary + dim gray secondary. DIRECT divider INVERTED to brand-yellow on
// black text for maximum contrast against the dark data rows.
// =======================================================================================

function applyTelemetryTheme(sheetName) {
  var sheet = _vlGetSheet(sheetName);
  if (!sheet) return "❌ Sheet '" + sheetName + "' not found.";

  var lock = LockService.getDocumentLock();
  try { lock.waitLock(15000); } catch (e) { return "❌ Sheet busy — try again."; }

  try {
    _vlClearExistingStyles(sheet);

    var boundary = _vlFindBoundary(sheet);
    if (boundary < 0) return "❌ DIRECT boundary not found in '" + sheetName + "'.";

    // ---- ROW HEIGHTS ----
    sheet.setRowHeight(1, 38);
    sheet.setRowHeight(2, 65);
    sheet.setRowHeight(3, 32);
    sheet.setRowHeight(boundary, 36);
    sheet.setRowHeight(boundary + 1, 32);
    _vlSetDataRowHeights(sheet, 4, boundary - 1, 26);
    _vlSetDataRowHeights(sheet, boundary + 2, sheet.getMaxRows(), 26);

    // ---- COLUMN WIDTHS — same tightened values as Service Bay ----
    sheet.setColumnWidth(1,  110);
    sheet.setColumnWidth(2,   70);
    sheet.setColumnWidth(3,   95);
    sheet.setColumnWidth(4,  130);
    sheet.setColumnWidth(5,  300);
    sheet.setColumnWidth(6,  130);
    sheet.setColumnWidth(7,  100);
    sheet.setColumnWidth(8,  100);
    sheet.setColumnWidth(9,  180);
    sheet.setColumnWidth(10,  90);

    // ---- BANNER ROW 1 — pure black with mono uppercase tracked tight ----
    var bannerRow1 = sheet.getRange(1, 1, 1, 10);
    bannerRow1.setBackground(VL.bgBlack).setFontColor(VL.yellow)
      .setFontFamily(VL.mono).setFontWeight('bold')
      .setFontSize(10).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true);
    // A1 inverted: yellow chip with black HQ monogram
    sheet.getRange('A1').setBackground(VL.yellow).setFontColor(VL.bgBlack)
      .setFontFamily(VL.display).setFontWeight('bold').setFontSize(15);

    // ---- BANNER ROW 2 ----
    sheet.getRange('A2:F2').setBackground(VL.bgCharcoal)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange('G2:J2').setBackground(VL.bgBlack).setFontColor(VL.yellow)
      .setFontFamily(VL.mono).setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

    // ---- HEADER ROW 3 ----
    var headerRow = sheet.getRange(3, 1, 1, 10);
    headerRow.setBackground(VL.bgBlack).setFontColor(VL.yellow)
      .setFontFamily(VL.mono).setFontWeight('bold')
      .setFontSize(10).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true);
    // 2px yellow rule under header — like a screen guide line
    headerRow.setBorder(null, null, true, null, null, null,
                        VL.yellow, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // ---- DATA AREA ----
    _vlApplyTelemetryDataFormat(sheet, 4, boundary - 1);
    _vlApplyTelemetryDataFormat(sheet, boundary + 2, sheet.getMaxRows());

    // ---- DIRECT DIVIDER — INVERTED: yellow on black, the loudest visual cue ----
    var divider = sheet.getRange(boundary, 1, 1, 10);
    divider.setBackground(VL.yellow).setFontColor(VL.bgBlack)
      .setFontFamily(VL.display).setFontWeight('bold').setFontSize(15)
      .setVerticalAlignment('middle').setWrap(false);
    sheet.getRange(boundary, 1).setHorizontalAlignment('left')
      .setValue('▌  DIRECT')
      .setFontSize(17);
    sheet.getRange(boundary, 7).setHorizontalAlignment('right')
      .setValue('HQ MS · DIRECT FEED')
      .setFontSize(9).setFontFamily(VL.mono).setFontWeight('bold');
    // Black borders top + bottom — frames the yellow stripe
    divider.setBorder(true, null, true, null, null, null,
                      VL.bgBlack, SpreadsheetApp.BorderStyle.SOLID_THICK);

    // ---- DIRECT HEADER ROW ----
    var directHeader = sheet.getRange(boundary + 1, 1, 1, 10);
    directHeader.setBackground(VL.bgBlack).setFontColor(VL.yellow)
      .setFontFamily(VL.mono).setFontWeight('bold')
      .setFontSize(10).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true);
    directHeader.setBorder(null, null, true, null, null, null,
                           VL.yellow, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // ---- BANDINGS — very subtle charcoal alternation (scanline-like) ----
    _vlApplyBanding(sheet, 4,             boundary - 1, VL.bgCharcoal, VL.bgCharcoal2);
    _vlApplyBanding(sheet, boundary + 2,  sheet.getMaxRows(), VL.bgCharcoal, VL.bgCharcoal2);

    // ---- CF ----
    _vlApplyTelemetryCF(sheet, boundary);

    SpreadsheetApp.flush();
    return "✅ Telemetry Console theme applied to '" + sheetName + "'.";
  } finally {
    lock.releaseLock();
  }
}

function _vlApplyTelemetryDataFormat(sheet, startRow, endRow) {
  if (endRow < startRow) return;
  var rows = endRow - startRow + 1;

  // A: SKU — bright yellow, mono, center
  sheet.getRange(startRow, 1, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('bold')
    .setFontColor(VL.yellow).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // B: QTY — yellow, center
  sheet.getRange(startRow, 2, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('bold')
    .setFontColor(VL.yellow).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // C: LOC — dim gray secondary
  sheet.getRange(startRow, 3, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // D: SALES ORDER — dim gray secondary
  sheet.getRange(startRow, 4, rows, 1)
    .setFontFamily(VL.mono).setFontSize(9).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // E: NOTE — dim gray, wrap
  sheet.getRange(startRow, 5, rows, 1)
    .setFontFamily(VL.mono).setFontSize(9).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('left')
    .setVerticalAlignment('middle').setWrap(true);

  // F: STATUS — CF handles color, base is bold yellow
  sheet.getRange(startRow, 6, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('bold')
    .setFontColor(VL.yellow).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // G: HAND — bright yellow, right
  sheet.getRange(startRow, 7, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('bold')
    .setFontColor(VL.yellow).setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  // H: LEFT — dim gray, right
  sheet.getRange(startRow, 8, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  // I: SHIPPING — dim, left
  sheet.getRange(startRow, 9, rows, 1)
    .setFontFamily(VL.mono).setFontSize(9).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('left')
    .setVerticalAlignment('middle').setWrap(true);

  // J: SHIP COST — dim, right
  sheet.getRange(startRow, 10, rows, 1)
    .setFontFamily(VL.mono).setFontSize(10).setFontWeight('normal')
    .setFontColor(VL.textDim).setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
}

function _vlApplyTelemetryCF(sheet, boundary) {
  var maxRow = sheet.getMaxRows();
  var rules = [];

  var statusRange = sheet.getRange(4, 6, maxRow - 3, 1);
  function statusRule(value, fg, strike) {
    var formula =
      '=AND(UPPER(TRIM($F4))="' + value + '",$A4<>"",' +
      'UPPER(TRIM($A4))<>"DIRECT",LEFT(TRIM($A4),1)<>"◈")';
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setFontColor(fg).setBold(true);
    if (strike) rule = rule.setStrikethrough(true);
    return rule.setRanges([statusRange]).build();
  }
  rules.push(statusRule('PREPARING', VL.yellowBright, false));
  rules.push(statusRule('SHIPPED',   VL.greenBright, false));
  rules.push(statusRule('CANCELED',  VL.textMuted, true));

  // HAND low-stock — red on dark, no bg flip (just font color)
  var handRange = sheet.getRange(4, 7, maxRow - 3, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($G4), $G4<=20)')
    .setFontColor(VL.redAlert).setBold(true)
    .setRanges([handRange]).build());

  // Buyer Note — italic + yellow (instead of muted gold, since gold is hard to
  // see on charcoal). Stands out as "the buyer's words" against the dim gray notes.
  var noteRange = sheet.getRange(4, 5, maxRow - 3, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E4<>"", REGEXMATCH(TO_TEXT($E4), "(?i)^buyer note:"))')
    .setItalic(true).setFontColor(VL.yellow)
    .setRanges([noteRange]).build());

  sheet.setConditionalFormatRules(rules);
}

// =======================================================================================
// SHARED HELPERS
// =======================================================================================

function _vlGetSheet(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(sheetName);
}

function _vlFindBoundary(sheet) {
  // Mirror of getBoundaryRow but takes a sheet param so it works on the test sheet.
  // Looser match than production: the themes write "▌  DIRECT" into the divider's
  // col A (the glyph IS the visual stripe). Strict-equals would fail on re-runs.
  // First try strict equals, then fall back to a CONTAINS check capped at 40 chars
  // so we don't mistakenly match an SKU description that contains "DIRECT".
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  // Pass 1: strict equals (fresh duplicate from production)
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === 'DIRECT') return i + 1;
  }
  // Pass 2: contains (theme has been applied, glyph injected)
  for (var j = 0; j < values.length; j++) {
    var s = String(values[j][0]).trim().toUpperCase();
    if (s.length > 0 && s.length <= 40 && s.indexOf('DIRECT') !== -1) return j + 1;
  }
  return -1;
}

function _vlClearExistingStyles(sheet) {
  // Wipe CF rules + bandings so the new theme writes onto a clean slate.
  // Cell-level formatting (bg, fg, fonts) gets overwritten explicitly by the
  // theme so we don't strip it here — saves a sheet round-trip.
  sheet.setConditionalFormatRules([]);
  sheet.getBandings().forEach(function(b) {
    try { b.remove(); } catch (e) {}
  });
  // Also wipe any borders left by a previous theme run on the divider/header
  // rows so the new theme can paint its own without ghost lines.
  // Use a moderately wide range to catch everything realistic.
  var maxRow = Math.min(sheet.getMaxRows(), 100);
  sheet.getRange(1, 1, maxRow, 10).setBorder(false, false, false, false, false, false);
}

function _vlApplyBanding(sheet, startRow, endRow, firstColor, secondColor) {
  if (endRow < startRow) return;
  var range = sheet.getRange(startRow, 1, endRow - startRow + 1, 10);
  var banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null)
         .setFirstRowColor(firstColor)
         .setSecondRowColor(secondColor);
}

function _vlSetDataRowHeights(sheet, startRow, endRow, height) {
  if (endRow < startRow) return;
  // setRowHeights(startRow, numRows, height) — batch is much faster than per-row
  var capped = Math.min(endRow, sheet.getMaxRows());
  sheet.setRowHeights(startRow, capped - startRow + 1, height);
}

/**
 * Strips CF + bandings + borders. Doesn't restore cell formatting (that's
 * what the theme functions do). Use between A/B comparisons if you want a
 * clean baseline before applying the other theme.
 */
function resetVisualLabSheet(sheetName) {
  var sheet = _vlGetSheet(sheetName);
  if (!sheet) return "❌ Sheet '" + sheetName + "' not found.";
  _vlClearExistingStyles(sheet);
  return "✅ Cleared CF + bandings + borders on '" + sheetName + "'. Re-apply a theme to restyle.";
}
