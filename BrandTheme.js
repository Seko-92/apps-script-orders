// =======================================================================================
// BRAND_THEME.gs — High Quality Motor Service brand system
// Single source of truth for All Orders sheet styling.
//
// USAGE:
//   1. Run applyBrandTheme()       — once, to apply the full theme
//   2. Run setupBrandLogo(<id>)    — once, after uploading new-logo.png to your Drive
//   3. Run refreshDynamicBandings() MANUALLY if banding ranges drift over time
//      (NOT auto-fired by row-add paths — that was overwriting user formatting).
//
// DYNAMIC TABLE CONTRACT:
//   Cell formats use COLUMN-LEVEL ranges (A4:A1000 etc) so new rows inherit.
//   Status colors use CONDITIONAL FORMATTING with wide ranges (F4:F1000) so
//   they recompute the moment a value changes. Sheets natively extends
//   bandings when rows are inserted INSIDE the banded range, so most row-add
//   operations don't need any banding refresh at all.
// =======================================================================================

var BRAND = {
  // Colors
  ink:         '#1a1a1a',  // structure / headers
  inkSoft:     '#434343',  // secondary/auxiliary text
  paper:       '#ffffff',  // row base
  paperWarm:   '#fff8e7',  // row banding (warm cream)
  yellow:      '#ffd400',  // brand action color (matches logo)
  yellowSoft:  '#fff4b0',  // soft yellow surface
  redAlert:    '#ff6b6b',  // low-stock alert (existing)
  greenSubtle: '#e8f5e9',  // SHIPPED bg
  greenInk:    '#1b5e20',  // SHIPPED fg
  graySubtle:  '#f0f0f0',  // CANCELED bg
  grayInk:     '#5f5f5f',  // CANCELED fg

  // Fonts (all available in Google Sheets font picker)
  fontDisplay: 'Oswald',         // labels, headers, "DIRECT", status
  fontMono:    'Roboto Mono',    // codes (SKU, ORDER, time, doc#)
  fontData:    'Roboto',         // body data, notes

  // Upper bound for column-level format + CF ranges.
  // Set well above realistic row counts so growth never escapes the format.
  dataLast: 1000
};

// =======================================================================================
// MAIN ENTRY POINTS
// =======================================================================================

/**
 * Applies the full brand theme — "Service Bay v6" design system (2026-05-17).
 *
 * Service Bay design language:
 *   - Cream paper data area (works with row banding)
 *   - Black/yellow banner rows (HQ + date + System Pulse + live stats)
 *   - DIRECT divider as heavy brand-yellow band with black Oswald text
 *   - Status: BG + bold text (PENDING red, PREPARING yellow, SHIPPED green, CANCELED gray)
 *   - HAND low-stock: font-only red (no bg) — disciplined secondary signal
 *   - Paid SHIP COST: yellow bg + bold (the "money on the line" cue)
 *   - Buyer Note CF: italic muted gold (subtle audit overlay)
 *   - Banner E1: live System Pulse from Activity Log MAX(A:A) + minutes-since
 *   - Banner G1: live COUNTIF-based status counts + today total from __SparkData
 *
 * Parameterized: when called without arguments, targets MAIN_SHEET_NAME (production).
 * Pass a sheet name to target a different sheet — used by VisualLab.testServiceBay()
 * to apply the exact same code to "Copy of All orders" for design experiments.
 *
 * Idempotent — safe to re-run. All CF rules and bandings get stripped before reapplied.
 *
 * @param {string} [sheetName] - Optional target sheet name. Defaults to MAIN_SHEET_NAME.
 */
function applyBrandTheme(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var targetName = sheetName || MAIN_SHEET_NAME;
  var sheet = ss.getSheetByName(targetName);
  if (!sheet) return "❌ Sheet '" + targetName + "' not found";

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return "❌ Server busy — try again"; }

  try {
    // ── GEOMETRY ──
    // Frozen rows = 0 deliberately. The two-table architecture would create
    // header confusion if eBay's header were frozen while scrolling into DIRECT
    // (same column labels, different data table). User's decision after testing.
    sheet.setFrozenRows(0);

    // ── COLUMN WIDTHS (v6 tightened) ──
    // Honest sizes for the data each column holds. Math verified against the
    // banner merges above — B1:D1 stays 285px (date fits at 200px); G1:J1 stays
    // 470px (stats fit at ~300px); G2:H2 Pick ID merge stays 200px (dropdown
    // text ~140px fits with margin).
    sheet.setColumnWidth(Schema.cols.SKU,         110);
    sheet.setColumnWidth(Schema.cols.QTY,          70);
    sheet.setColumnWidth(Schema.cols.LOCATION,     95);
    sheet.setColumnWidth(Schema.cols.SALES_ORDER, 130);
    sheet.setColumnWidth(Schema.cols.NOTE,        300);
    sheet.setColumnWidth(Schema.cols.STATUS,      130);  // tightened from 250 → 130
    sheet.setColumnWidth(Schema.cols.HAND,        100);  // tightened from 145 → 100
    sheet.setColumnWidth(Schema.cols.LEFT,        100);  // tightened from 145 → 100
    sheet.setColumnWidth(Schema.cols.SHIPPING,    180);
    sheet.setColumnWidth(Schema.cols.SHIP_COST,    90);

    // ── ROW HEIGHTS ──
    // Lab values that produced the right visual rhythm. Set BEFORE _style*
    // calls so the styled rows have the right height when their content lands.
    sheet.setRowHeight(1, 42);   // banner row 1 — date+pulse+stats strip
    sheet.setRowHeight(2, 65);   // logo + Pick-ID badges (taller for eBay logo + dropdown breathing room)
    sheet.setRowHeight(3, 36);   // eBay header row
    // Data rows: uniform 30px breathable read. setRowHeights is a batch op —
    // much faster than per-row. Boundary + DIRECT header heights get overridden
    // below (after _styleDirectDivider runs) so they don't stay at 30px.
    var dataLast = Math.min(sheet.getMaxRows(), BRAND.dataLast);
    if (dataLast >= 4) {
      sheet.setRowHeights(4, dataLast - 3, 30);
    }

    // ── BANNER TYPOGRAPHY ──
    _styleBannerRow1(sheet);
    _styleBannerRow2(sheet);
    _styleHeaderRow(sheet, Schema.headerRow);

    _ensureDateFormula(sheet);

    // ── DATA AREA TYPOGRAPHY ──
    _applyColumnLevelDataFormats(sheet);

    // ── DIRECT DIVIDER + DIRECT HEADER ──
    var boundary = _findBoundaryInSheet(sheet);
    if (boundary > 0) {
      _styleDirectDivider(sheet, boundary);     // sets row height 40 internally
      _styleHeaderRow(sheet, boundary + 1);
      sheet.setRowHeight(boundary + 1, 36);     // DIRECT header row — same as eBay header
    }

    // ── CONDITIONAL FORMATTING (v6 — all in one wipe-and-rebuild pass) ──
    // We consolidate all CF here so re-running the theme produces a clean,
    // deterministic rule set. Order matters: status rules paint backgrounds
    // (most prominent), HAND/SHIP COST/Buyer Note paint specific cells, then
    // bandings sit underneath everything.
    _applyAllConditionalFormatting(sheet);

    // ── LIVE BANNER FORMULAS ──
    _ensureSparkData(ss);            // hidden helper sheet for hourly counts + sync pulse
    _setSystemPulseBannerFormulas(sheet);

    // ── BANDINGS ──
    if (sheetName) {
      // For test-sheet runs, apply a simple banding directly (the production
      // refreshDynamicBandings() targets MAIN_SHEET_NAME specifically).
      _applyTestSheetBanding(sheet, boundary);
    } else {
      refreshDynamicBandings();
    }

    return "✅ Service Bay theme applied to '" + targetName + "'.";
  } finally {
    lock.releaseLock();
  }
}

/**
 * Installs a self-updating date+time formula in B1 (the merge anchor of the
 * banner's date area). Reads as e.g. "Tuesday, May 1, 2026 · 8:32 PM".
 *
 * Implementation: `=TEXT(NOW(), "...")`. NOW() recalculates whenever the
 * sheet recalculates — for a busy warehouse sheet that's effectively every
 * few minutes (every n8n insert, every status change, every cell edit). No
 * trigger overhead, no quota cost. Slight staleness during long idle periods
 * (visible only if the sheet sits untouched for hours).
 *
 * Idempotent — safe to re-run. Brand styling on B1 (font, color, alignment)
 * is preserved because setFormula only changes the cell's value, not its
 * formatting. The spreadsheet timezone (set to America/Chicago by
 * setupActivityLogSheet) governs the displayed time.
 */
function setupBannerDateTime() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  // B1:D1 is merged. Writing to the anchor (B1) populates the merged display.
  sheet.getRange("B1").setFormula(
    '=TEXT(NOW(), "dddd, mmmm d, yyyy · h:mm AM/PM")'
  );
  return "✅ Banner date+time formula installed. Updates on every recalc.";
}


/**
 * Restores the eBay logo banner above the eBay table (cell A2 — anchor of
 * the A2:F2 merge that BrandTheme reserves as "eBay logo zone").
 *
 * Uses an `=IMAGE()` formula pointing at Wikimedia Commons' public eBay logo
 * — a stable, retina-quality, license-clean source. Mode 4 sets explicit
 * pixel dimensions (height 32, width 120) so the logo fits proportionally
 * inside the merged banner without stretching.
 *
 * Idempotent — overwrites whatever's currently in A2.
 */
function setupEbayLogo() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  // Wikimedia mirror of eBay's official logo. PNG renders crisp inside Sheets'
  // IMAGE formula on any zoom level. The "1200px" variant balances quality
  // with load time — the cell renders ~120px wide, so anything larger is
  // overkill.
  var logoUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/EBay_logo.svg/1200px-EBay_logo.svg.png";

  sheet.getRange("A2").setFormula('=IMAGE("' + logoUrl + '", 4, 32, 120)');
  return "✅ eBay logo restored to A2.";
}


/**
 * Adds WARNING-ONLY protections to the structural rows that should not be
 * edited casually — banner (rows 1-3), the DIRECT divider, and the DIRECT
 * header row. Warning-only means: anyone can still edit (no hard lock), but
 * Sheets pops a "you're editing a protected range — are you sure?" dialog
 * first. This catches accidental edits without blocking intentional ones.
 *
 * SELF-HEALS the DIRECT marker. `getBoundaryRow()` does a strict equality
 * check on column A === "DIRECT" — if someone (or some past code path) wrote
 * "HQ DIRECT" / "Direct Sales" / "▌ DIRECT" there, the lookup returns -1 and
 * a bunch of downstream things silently break (sort, row inserts, this
 * protection, etc.). Before protecting, we fall back to a case-insensitive
 * contains-search; if found, we write back the canonical value so the rest
 * of the system starts working again.
 *
 * Idempotent — re-running removes any prior HQ-STRUCTURE protections before
 * adding fresh ones (so it stays in sync if the boundary row moves).
 */
function protectSheetStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  // ---- 1. Strip any prior HQ-STRUCTURE protections (idempotent re-run) ----
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var removed = 0;
  existing.forEach(function(p) {
    var d = String(p.getDescription() || "");
    if (d.indexOf("HQ-STRUCTURE") === 0) {
      try { p.remove(); removed++; } catch (e) { /* ignore */ }
    }
  });

  // ---- 2. Banner rows 1-3 (logo, stats, column headers) ----
  sheet.getRange(1, 1, 3, Schema.dataWidth)
    .protect()
    .setDescription("HQ-STRUCTURE: Banner rows 1-3 — accidental-edit guard")
    .setWarningOnly(true);

  // ---- 3. Locate (or self-heal) the DIRECT divider row ----
  var boundary = getBoundaryRow();
  var healed = false;

  if (boundary <= 0) {
    // Strict match failed. Scan column A for ANY row whose value contains
    // "DIRECT" (case-insensitive). Most likely culprit: a decorative prefix
    // got typed in (e.g. "HQ DIRECT") which broke getBoundaryRow.
    var lastRow = sheet.getLastRow();
    if (lastRow >= Schema.dataStartRow) {
      var colA = sheet.getRange(1, Schema.cols.SKU, lastRow, 1).getValues();
      for (var i = 0; i < colA.length; i++) {
        var s = String(colA[i][0] || "").trim().toUpperCase();
        // Match "HQ DIRECT", "DIRECT SALES", "▌ DIRECT", etc.
        // Skip rows that just have "DIRECT" inside a longer word (defensive).
        if (s.indexOf("DIRECT") !== -1 && s.length < 32) {
          boundary = i + 1;
          // Write back the canonical value so getBoundaryRow works henceforth
          sheet.getRange(boundary, Schema.cols.SKU).setValue(Schema.boundaryMarker);
          healed = true;
          break;
        }
      }
    }
  }

  // ---- 4. Apply DIRECT-divider protection (rows boundary + boundary+1) ----
  var boundaryNote;
  if (boundary > 0) {
    sheet.getRange(boundary, 1, 2, Schema.dataWidth)
      .protect()
      .setDescription("HQ-STRUCTURE: DIRECT divider + header (rows " +
                      boundary + "-" + (boundary + 1) + ") — accidental-edit guard")
      .setWarningOnly(true);
    boundaryNote = " · DIRECT divider at row " + boundary +
                   (healed ? " (self-healed col A → '" + Schema.boundaryMarker + "')" : "");
  } else {
    boundaryNote = " · ⚠️ no DIRECT divider found anywhere in column A — " +
                   "manually verify the divider row exists and re-run.";
  }

  return "✅ Sheet structure protected (warning-only)" + boundaryNote +
         (removed > 0 ? " · refreshed (" + removed + " prior protection(s))" : "");
}


/**
 * Removes ALL HQ-STRUCTURE protections. Use if you want to disable the
 * accidental-edit guards entirely.
 */
function unprotectSheetStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var removed = 0;
  protections.forEach(function(p) {
    var d = String(p.getDescription() || "");
    if (d.indexOf("HQ-STRUCTURE") === 0) {
      try { p.remove(); removed++; } catch (e) {}
    }
  });
  return "✅ Removed " + removed + " HQ-STRUCTURE protection(s).";
}


/**
 * Inserts the HQ logo over cell A1 from a Drive file.
 * @param {string} driveFileIdOrUrl - Drive file ID or share URL
 */
function setupBrandLogo(driveFileIdOrUrl) {
  if (!driveFileIdOrUrl) {
    return "❌ Provide a Drive file ID or share URL.\n" +
           "Example: setupBrandLogo('1abc...XYZ')\n" +
           "Or:      setupBrandLogo('https://drive.google.com/file/d/1abc...XYZ/view')";
  }

  // Extract file ID from URL if needed
  var fileId = String(driveFileIdOrUrl);
  var match = fileId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) fileId = match[1];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);

  // Remove any image already anchored at A1
  sheet.getImages().forEach(function(img) {
    try {
      if (img.getAnchorCell().getA1Notation() === 'A1') img.remove();
    } catch (e) {}
  });

  // Pull the file
  var file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    return "❌ Could not access Drive file. Make sure: (1) the ID is correct, " +
           "(2) the script has Drive access (it should — accept the OAuth prompt).\n   " + e.toString();
  }

  var blob = file.getBlob();
  var image = sheet.insertImage(blob, 1, 1);  // anchor at A1

  // Fit image to A1 cell, preserve aspect ratio, leave 4px padding
  var rowH = sheet.getRowHeight(1);
  var colW = sheet.getColumnWidth(1);
  var aspect = image.getWidth() / image.getHeight();
  var targetH = rowH - 4;
  var targetW = targetH * aspect;
  if (targetW > colW - 4) {
    targetW = colW - 4;
    targetH = targetW / aspect;
  }
  image.setWidth(Math.round(targetW)).setHeight(Math.round(targetH));

  // Hide the placeholder "HQ" text now that the image sits over A1
  sheet.getRange('A1').setValue('');

  return "✅ HQ logo installed over A1.";
}

/**
 * Rebuilds the eBay and DIRECT bandings to span the current dynamic ranges.
 * Call after ANY row insert/delete that could move the DIRECT boundary or
 * extend the data area past existing banding edges.
 *
 * Banding theme: white / paperWarm cream alternation, on top of which CF
 * rules paint status and low-stock colors.
 */
function refreshDynamicBandings() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;

  var boundary = getBoundaryRow();
  if (boundary <= Schema.dataStartRow) return;  // Sheet not in expected shape

  // Remove existing bandings (clean slate)
  sheet.getBandings().forEach(function(b) {
    try { b.remove(); } catch (e) {}
  });

  var maxRow = Math.max(sheet.getMaxRows(), boundary + 5);

  // eBay banding: header row + data rows up to boundary - 1
  // (Header row gets its own dark format on top of the banding header color.)
  var ebayHeight = boundary - Schema.headerRow;
  if (ebayHeight > 0) {
    var ebayRange = sheet.getRange(Schema.headerRow, 1, ebayHeight, Schema.dataWidth);
    var ebayBand  = ebayRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    ebayBand.setHeaderRowColor(BRAND.ink)
            .setFirstRowColor(BRAND.paper)
            .setSecondRowColor(BRAND.paperWarm);
  }

  // DIRECT banding: DIRECT header (boundary + 1) + data rows to maxRow
  var directHeaderRow = boundary + 1;
  if (directHeaderRow <= maxRow) {
    var directHeight = maxRow - directHeaderRow + 1;
    var directRange  = sheet.getRange(directHeaderRow, 1, directHeight, Schema.dataWidth);
    var directBand   = directRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    directBand.setHeaderRowColor(BRAND.ink)
              .setFirstRowColor(BRAND.paper)
              .setSecondRowColor(BRAND.paperWarm);
  }
}

/**
 * One-shot repair for sheets where the divider value drifted away from "DIRECT"
 * (e.g., previous theme version wrote "▌ HQ · DIRECT" and broke getBoundaryRow).
 * Searches column A by substring, restores the canonical "DIRECT" value, then
 * re-applies the brand theme. Safe to run any time.
 */
function repairBrandTheme() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  // Find the divider row by substring match on column A
  var lastRow = sheet.getLastRow();
  var values  = sheet.getRange(1, Schema.cols.SKU, lastRow, 1).getValues();
  var brokenRow = -1;
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0]).trim().toUpperCase();
    if (v.indexOf(Schema.boundaryMarker) !== -1 && v.length < 50) {
      brokenRow = i + 1;
      break;
    }
  }
  if (brokenRow === -1) {
    return "❌ Could not locate the " + Schema.boundaryMarker + " divider. Manually set its column A cell to exactly '" + Schema.boundaryMarker + "' and re-run applyBrandTheme().";
  }

  // Restore the canonical value
  sheet.getRange(brokenRow, Schema.cols.SKU).setValue(Schema.boundaryMarker);

  // Verify getBoundaryRow can now find it
  var boundary = getBoundaryRow();
  if (boundary !== brokenRow) {
    return "⚠️ Restored row " + brokenRow + " to '" + Schema.boundaryMarker + "' but getBoundaryRow returned " + boundary + ". Inspect manually.";
  }

  // Re-apply theme; divider will now be styled correctly
  var result = applyBrandTheme();
  return "✅ Repaired divider on row " + brokenRow + ". " + result;
}

/**
 * Relocates the Pick ID for Adjustment dropdown between its two supported
 * layouts. Shipped 2026-05-19 to support hiding cols I + J as part of the
 * SHIPPING + SHIP COST soft-delete (Schema.cellAdjustmentId moved from I2 → E2).
 *
 *   target === "E2" — hidden-cols layout (cols I + J hidden on sheet)
 *                     Merges: A2:D2 (logo) + E2:F2 (Adjustment) + G2:H2 (Shipping)
 *
 *   target === "I2" — default layout (cols I + J visible on sheet)
 *                     Merges: A2:F2 (logo) + G2:H2 (Shipping) + I2:J2 (Adjustment)
 *
 * Operation is fully symmetric — call with the OTHER target to revert.
 * Preserves the validation rule + currently-selected value through the move.
 * Idempotent — re-running with the same target returns "already in place" and
 * makes no changes.
 *
 * IMPORTANT after running: update Schema.cellAdjustmentId to match the new
 * location ("E2" or "I2"), then clasp push. The function itself only moves
 * the sheet-side artifacts; the code's source of truth is Schema.
 *
 * @param {string} target - "E2" or "I2"
 * @returns {string} Status message
 */
function relocateAdjustmentBadge(target) {
  if (target !== "E2" && target !== "I2") {
    return "❌ target must be 'E2' (hidden-cols layout) or 'I2' (default layout)";
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  // Detect current source by checking which of I2/E2 has the validation.
  // (Validations are cell-level; survive merge/unmerge operations.)
  var source = null;
  if (sheet.getRange('I2').getDataValidation()) source = 'I2';
  else if (sheet.getRange('E2').getDataValidation()) source = 'E2';

  if (!source) {
    return "❌ No Pick ID for Adjustment validation found at I2 OR E2. " +
           "Set up the dropdown first via Sheets UI (Data → Data validation), " +
           "then re-run this function.";
  }

  if (source === target) {
    return "✅ Pick ID for Adjustment already at " + target + " — no changes needed.";
  }

  // Capture source state before disturbing anything
  var sourceRange = sheet.getRange(source);
  var validation = sourceRange.getDataValidation();
  var value = sourceRange.getValue();

  // Break only the row-2 merges we're going to recreate. G2:H2 (Shipping) is
  // never touched — its merge stays intact across both layouts.
  if (target === "E2") {
    try { sheet.getRange('A2:F2').breakApart(); } catch (e) { /* not merged — ignore */ }
    try { sheet.getRange('I2:J2').breakApart(); } catch (e) { /* not merged — ignore */ }
  } else {
    try { sheet.getRange('A2:D2').breakApart(); } catch (e) { /* not merged — ignore */ }
    try { sheet.getRange('E2:F2').breakApart(); } catch (e) { /* not merged — ignore */ }
  }

  // Clear the source cell's validation + value (before recreating merges)
  sourceRange.setDataValidation(null).setValue('');

  // Recreate the layout's merges
  if (target === "E2") {
    sheet.getRange('A2:D2').merge();
    sheet.getRange('E2:F2').merge();
  } else {
    sheet.getRange('A2:F2').merge();
    sheet.getRange('I2:J2').merge();
  }

  // Apply validation + value + brand styling to target
  sheet.getRange(target)
    .setDataValidation(validation)
    .setValue(value)
    .setBackground(BRAND.ink)
    .setFontColor(BRAND.yellow)
    .setFontFamily(BRAND.fontDisplay)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Re-paint the logo merge zone with the warm cream background
  var logoMerge = (target === "E2") ? 'A2:D2' : 'A2:F2';
  sheet.getRange(logoMerge).setBackground(BRAND.paperWarm);

  return "✅ Pick ID for Adjustment relocated " + source + " → " + target +
         ". Logo merge: " + logoMerge + ". Now update Schema.cellAdjustmentId " +
         "to \"" + target + "\" if not already, then clasp push.";
}

/**
 * One-shot repair for the live banner formulas in E1 (System Pulse) and
 * G1 (status counts + TODAY total). Use when those cells show stale static
 * text — typically the OLD format "🔴 Pending: N   🟡 Preparing: N …" left
 * behind from before updateOrderStatsInSheet was converted to a no-op.
 *
 * Touches ONLY:
 *   - __SparkData helper sheet (ensured/refreshed)
 *   - E1 formula
 *   - G1 formula
 * Does NOT re-apply theme, banding, CF, column widths, or anything else.
 */
function repairLiveBannerFormulas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";
  _ensureSparkData(ss);
  _setSystemPulseBannerFormulas(sheet);
  return "✅ Live banner formulas re-installed in E1 and G1.";
}

/**
 * Reverts the brand theme. Use if the team wants the old look back.
 * Note: this restores defaults but cannot recover any pre-existing
 * manual cell colors that the theme overwrote — those are gone.
 */
function revertBrandTheme() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  sheet.setFrozenRows(0);
  sheet.setColumnWidth(Schema.cols.STATUS, 250);

  // Strip status CF rules; preserve HAND CF and any other rules
  var rules = sheet.getConditionalFormatRules();
  var keep = [];
  rules.forEach(function(rule) {
    var bc = rule.getBooleanCondition();
    if (!bc) { keep.push(rule); return; }
    var formula = (bc.getCriteriaValues() || [''])[0] || '';
    if (formula.indexOf('PREPARING') === -1 &&
        formula.indexOf('SHIPPED')   === -1 &&
        formula.indexOf('CANCELED')  === -1) {
      keep.push(rule);
    }
  });
  sheet.setConditionalFormatRules(keep);

  return "✅ Brand theme reverted (frozen rows, F width, status CF). Cell formats and bandings retained — re-run applyBrandTheme to restore.";
}

// =======================================================================================
// PRIVATE HELPERS
// =======================================================================================

function _styleBannerRow1(sheet) {
  // Service Bay v6 — UNIFORM banner row 1. Entire row gets the same base
  // styling (black bg + brand yellow + Roboto bold 11pt + center + wrap),
  // then A1 only gets bumped to Oswald 16pt as the brand monogram.
  //
  // Why uniform: previously each cell got individual treatment (Oswald for A1/
  // B1/G1, Roboto Mono for E1), which left gaps when legacy writers (updateLast
  // SyncTimestamp / updateOrderStatsInSheet) were converted to no-ops on
  // 2026-05-17. The legacy functions used to set E1/G1's font color; the
  // converted no-ops don't, so prior cell colors (white from old stats code)
  // persisted. Uniform whole-row styling here guarantees E1 + G1 inherit
  // brand yellow on black even when no legacy writer exists. Emoji bullets
  // (🔴🟡🟢⚫🟢) in the G1/E1 formulas render their own colors regardless.
  var fullRow = sheet.getRange(1, 1, 1, Schema.dataWidth);
  fullRow.setBackground(BRAND.ink)
         .setFontColor(BRAND.yellow)
         .setFontFamily(BRAND.fontData)
         .setFontWeight('bold')
         .setFontSize(11)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle')
         .setWrap(true);

  // A1 — HQ brand monogram. Oswald display 16pt for the chip identity.
  // Image overlay (via setupBrandLogo) will sit on top once installed.
  sheet.getRange('A1')
    .setFontFamily(BRAND.fontDisplay)
    .setFontSize(16);
  if (!sheet.getRange('A1').getValue()) sheet.getRange('A1').setValue('HQ');
}

function _styleBannerRow2(sheet) {
  // Row 2: logo zone + Pick ID badges. Three valid layouts seen in production:
  //
  //   Default layout (cols I + J visible):
  //     A2:F2 = eBay logo zone, G2:H2 = Shipping, I2:J2 = Adjustment
  //     → Schema.cellAdjustmentId === "I2"
  //
  //   Hidden-cols (programmatic E2 migration, never actually used in prod):
  //     A2:D2 = eBay logo zone, E2:F2 = Adjustment, G2:H2 = Shipping
  //     → Schema.cellAdjustmentId === "E2"
  //
  //   Hidden-cols (manual compaction, current production state 2026-05-19):
  //     A2:E2 = eBay logo zone, F2:G2 = Shipping, H2 = Adjustment
  //     → Schema.cellAdjustmentId === "H2", Schema.cellEmployeeId === "F2"
  //     Row 1 also compacted: F1:H1 = stats banner (Schema.cellStats === "F1")
  //
  // This function only PAINTS — it does not create or break merges.
  var logoZone;
  if (Schema.cellAdjustmentId === 'E2') {
    logoZone = 'A2:D2';
  } else if (Schema.cellAdjustmentId === 'H2') {
    logoZone = 'A2:E2';
  } else {
    logoZone = 'A2:F2';
  }
  sheet.getRange(logoZone).setBackground(BRAND.paperWarm);

  // Both Pick ID cells have data validation (dropdowns of allowed values).
  // We MUST NOT write VALUES to these cells during the theme apply, or the
  // validation will reject anything not in its list and abort the whole theme.
  // Style only — values stay untouched, dropdown stays functional.
  [Schema.cellEmployeeId, Schema.cellAdjustmentId].forEach(function(a1) {
    sheet.getRange(a1)
      .setBackground(BRAND.ink)
      .setFontColor(BRAND.yellow)
      .setFontFamily(BRAND.fontDisplay)
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
  });
}

/**
 * OPT-IN: rewrites the G2 (Shipping) and I2 (Adjustment) dropdowns as two-line
 * badge values. After running this:
 *   - The dropdown options become "SHIPPING\nYAwiss · 1" instead of "Shipping - YAwiss 1"
 *   - Each cell's currently-selected value is migrated to its new two-line equivalent
 *   - Validation still works — the new options are the only valid choices
 *
 * Run this AFTER applyBrandTheme() succeeds. Safe to re-run; converts only what's
 * still in the old single-line format.
 */
function setupPickIdBadges() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  var report = [];
  report.push(_rewritePickIdValidation(sheet.getRange(Schema.cellEmployeeId),   'SHIPPING',   /^shipping\s*[-:·]\s*(.+)$/i));
  report.push(_rewritePickIdValidation(sheet.getRange(Schema.cellAdjustmentId), 'ADJUSTMENT', /^adjustment(?:s)?\s*[-:·]\s*(.+)$/i));
  return "Pick ID badge migration:\n" + report.join("\n");
}

function _rewritePickIdValidation(range, label, parsePattern) {
  var a1 = range.getA1Notation();
  var validation = range.getDataValidation();
  if (!validation) return "  • " + a1 + ": no validation found — skipped";

  var criteriaType = validation.getCriteriaType();
  if (criteriaType !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return "  • " + a1 + ": validation type is " + criteriaType + " (not a list) — skipped";
  }

  var raw = validation.getCriteriaValues()[0];   // [0] is the list of options
  var oldOptions = (raw || []).map(function(o) { return String(o); });

  var newOptions = [];
  var migrationMap = {};

  oldOptions.forEach(function(opt) {
    var s = opt.trim();
    if (s.indexOf('\n') !== -1) {
      // Already two-line — leave it
      newOptions.push(s);
      migrationMap[opt] = s;
      return;
    }
    var match = s.match(parsePattern);
    if (match) {
      var data = match[1].trim().replace(/\s+(\d+)$/, ' · $1');
      var newOpt = label + '\n' + data;
      newOptions.push(newOpt);
      migrationMap[opt] = newOpt;
    } else {
      // Anything that doesn't match (e.g., the default "Pick ID for Shipping"
      // placeholder) → keep as-is so the dropdown still has a "no selection" option
      newOptions.push(s);
      migrationMap[opt] = s;
    }
  });

  // Build and apply the new validation
  var newValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(newOptions, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(newValidation);

  // Migrate the currently-selected value
  var current = String(range.getValue()).trim();
  if (migrationMap.hasOwnProperty(current)) {
    range.setValue(migrationMap[current]);
  }

  return "  • " + a1 + ": " + oldOptions.length + " option(s) migrated to two-line badges";
}

function _styleHeaderRow(sheet, row) {
  // Black bg, brand yellow uppercase Oswald, thick yellow underline
  var range = sheet.getRange(row, 1, 1, Schema.dataWidth);
  range.setBackground(BRAND.ink)
       .setFontColor(BRAND.yellow)
       .setFontFamily(BRAND.fontDisplay)
       .setFontWeight('bold')
       .setFontSize(10)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  range.setBorder(null, null, true, null, null, null,
                  BRAND.yellow, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function _styleDirectDivider(sheet, boundary) {
  // Service Bay v6 divider — full-row brand-yellow band, the loudest section
  // break in the sheet. Reads from across the warehouse.
  //
  // Architecture:
  //   A:F merge (Schema.boundaryLeftWidth) → "DIRECT" left-aligned, big Oswald
  //   G:J merge → "HQ MS · DIRECT TABLE" right-aligned, smaller Oswald
  //   Whole row: brand-yellow #ffd400 bg + brand-black text
  //   Top + bottom thick black borders frame the band visually
  //
  // CRITICAL: The left-merge value MUST stay exactly Schema.boundaryMarker
  // ("DIRECT"). getBoundaryRow() does strict equality on this constant —
  // prepending glyphs ("▌ DIRECT") or branding ("HQ DIRECT") silently breaks
  // every function downstream (sort, row inserts, live sync, fulfillment,
  // protection self-heal). The yellow band itself provides the visual cue;
  // we don't need decorative prefixes in the canonical marker cell.
  var leftMerge  = sheet.getRange(boundary, 1, 1, Schema.boundaryLeftWidth);                             // A:F
  var rightMerge = sheet.getRange(boundary, Schema.boundaryLeftWidth + 1, 1, Schema.boundaryRightWidth); // G:J

  leftMerge.setValue(Schema.boundaryMarker)      // ← Underlying value MUST be exactly this
           .setNumberFormat('"▌  "@')             // ← DISPLAY prepends the bar glyph; underlying value untouched
           .setBackground(BRAND.yellow)
           .setFontColor(BRAND.ink)
           .setFontFamily(BRAND.fontDisplay)
           .setFontWeight('bold')
           .setFontSize(16)
           .setHorizontalAlignment('left')
           .setVerticalAlignment('middle');
  // The ▌ glyph is a number-format prefix, NOT a value. getValue() returns
  // "DIRECT" (the underlying value), so getBoundaryRow()'s strict-equality
  // contract stays intact. The visual stripe lives purely in the cell's
  // displayed render. Sheets persists number formats per-cell, so the prefix
  // survives re-runs of this function.

  rightMerge.setValue('HQ MS · DIRECT TABLE')
            .setBackground(BRAND.yellow)
            .setFontColor(BRAND.ink)
            .setFontFamily(BRAND.fontDisplay)
            .setFontWeight('bold')
            .setFontSize(10)
            .setHorizontalAlignment('right')
            .setVerticalAlignment('middle');

  // Black borders top + bottom across the entire row — frames the yellow band
  // so it reads as a defined section break (not just a colored row).
  sheet.getRange(boundary, 1, 1, Schema.dataWidth)
       .setBorder(true, null, true, null, null, null,
                  BRAND.ink, SpreadsheetApp.BorderStyle.SOLID_THICK);

  sheet.setRowHeight(boundary, 40);
}

function _applyColumnLevelDataFormats(sheet) {
  // Service Bay v6 column typography. Applied at the entire data band so any
  // inserted row inherits format automatically.
  //
  // Design rules:
  //   - Roboto Mono for codes (SKU, QTY, LOC, ORDER, HAND, LEFT, SHIP COST) —
  //     feels like part-number readouts on a service spec sheet.
  //   - Roboto regular for prose-like text (NOTE only).
  //   - All numerics CENTER (HAND/LEFT/QTY/SHIP COST) — warehouse typography
  //     favors center over right-align for 1-3 digit values; matches column rhythm.
  //   - inkSoft secondary color for LEFT, SHIPPING, SHIP COST (auxiliary data
  //     the picker reads but doesn't act on directly).
  //   - NO italic on NOTE — italic is reserved for buyer-note CF only.
  //
  // STATUS column (F) is intentionally NOT touched — data validation dropdown
  // would conflict. All status visuals come from CF.
  var rows = BRAND.dataLast - Schema.bannerRows;
  var startRow = Schema.dataStartRow;

  // A: SKU — Roboto Mono, bold, center, ink (primary anchor)
  sheet.getRange(startRow, Schema.cols.SKU, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.ink)
    .setFontWeight('bold').setFontSize(11)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // B: QTY — Roboto Mono, bold, center
  sheet.getRange(startRow, Schema.cols.QTY, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.ink)
    .setFontWeight('bold').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // C: LOCATION — Roboto Mono, regular weight, center (codes like E-30)
  sheet.getRange(startRow, Schema.cols.LOCATION, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.ink)
    .setFontWeight('normal').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // D: SALES ORDER — Roboto Mono, regular, center (matches eBay convention,
  //    centered since 2026-05-16 to fix DIRECT-table column-alignment drift)
  sheet.getRange(startRow, Schema.cols.SALES_ORDER, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.ink)
    .setFontWeight('normal').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // E: NOTE — Roboto regular, ink, left, WRAP (prose-style buyer/supervisor notes)
  //    NO italic at column level. Italic + muted gold is added per-cell via the
  //    buyer-note CF rule (when the cell starts with "Buyer Note:").
  sheet.getRange(startRow, Schema.cols.NOTE, rows, 1)
    .setFontFamily(BRAND.fontData).setFontColor(BRAND.ink)
    .setFontWeight('normal').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setWrap(true);

  // F: STATUS — DELIBERATELY UNTOUCHED. Validation dropdown + CF own this column.

  // G: HAND — Roboto Mono, bold, center, ink (CF paints red font when ≤20)
  sheet.getRange(startRow, Schema.cols.HAND, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.ink)
    .setFontWeight('bold').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // H: LEFT — Roboto Mono, regular, center, inkSoft (auxiliary, picker fills post-pick)
  sheet.getRange(startRow, Schema.cols.LEFT, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.inkSoft)
    .setFontWeight('normal').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // I: SHIPPING — Roboto regular, center, inkSoft (auxiliary; v6 changed from
  //    left to center to match the surrounding columns' rhythm)
  sheet.getRange(startRow, Schema.cols.SHIPPING, rows, 1)
    .setFontFamily(BRAND.fontData).setFontColor(BRAND.inkSoft)
    .setFontWeight('normal').setFontSize(9)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  // J: SHIP COST — Roboto Mono, regular, center, inkSoft (CF paints yellow bg on paid)
  sheet.getRange(startRow, Schema.cols.SHIP_COST, rows, 1)
    .setFontFamily(BRAND.fontMono).setFontColor(BRAND.inkSoft)
    .setFontWeight('normal').setFontSize(10)
    .setFontStyle('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Vertical alignment middle on the whole data band (belt-and-suspenders;
  // individual cols already set it but this guarantees consistency)
  sheet.getRange(startRow, 1, rows, Schema.dataWidth).setVerticalAlignment('middle');
}

/**
 * Buyer Note highlighting (2026-05-16 — designer pass).
 * ─────────────────────────────────────────────────────────────────────────
 * Adds ONE conditional-formatting rule to the NOTE column (E) on All Orders:
 *   Cells starting with "Buyer Note:" (case-insensitive)
 *     → italic + muted gold-brown font color (#8a7434)
 *     → no background change (preserves banding, status CF, low-stock highlights)
 *
 * Supervisor notes (anything else non-empty in the NOTE cell) are deliberately
 * left UNSTYLED — they're the common case, and dressing them up would add color
 * noise to an already-busy sheet. The buyer note is the exception; that's what
 * gets the visual cue.
 *
 * Edit workflow consequence: when a supervisor rewrites a buyer note and
 * removes the "Buyer Note:" prefix as part of the edit, the CF rule no longer
 * matches → italic/gold disappear → cell snaps back to default. The act of
 * editing IS the act of taking ownership; the sheet shows it back to you.
 *
 * Idempotent — strips any prior buyer-note rule (identified by NOTE-column
 * range + formula containing "buyer note") before re-adding.
 *
 * Standalone for v1 — NOT wired into applyBrandTheme() yet (per user, pending
 * a sheet-design audit via SheetInspector.inspectSheetDesign()).
 */
/**
 * Buyer Note highlighting — public entry point. Delegates to the private
 * helper so it can target any sheet (production "All orders" by default,
 * "Copy of All orders" or other test sheets when called from VisualLab).
 * Standalone idempotent — re-run safely. Now ALSO wired into applyBrandTheme()
 * via _applyAllConditionalFormatting() so the theme owns its full CF set.
 */
function setupBuyerNoteHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  // Re-apply the buyer-note rule in place, preserving every other CF rule.
  // (Different from the theme path which rebuilds the FULL CF set; this
  // one-rule update is safer for standalone "fix the highlight" runs.)
  var rules = sheet.getConditionalFormatRules();
  rules = _stripBuyerNoteRule(rules);
  rules.push(_buildBuyerNoteRule(sheet));
  sheet.setConditionalFormatRules(rules);

  return "✅ Buyer Note highlighting applied — italic + muted gold for cells starting with 'Buyer Note:'.";
}

/**
 * Kit row highlighting — public entry point. Prepends a "▣ " glyph to the
 * SKU display on any All Orders row whose SKU is a member of the Kit Registry.
 *
 * Design intent (2026-05-18, glyph-prefix iteration):
 *   - Multi-kit DIRECT orders are the motivating case — the picker has to
 *     mentally tag several consecutive rows as needing kit handling, and
 *     attention fatigues across that kind of cluster. A subtle SKU marker
 *     prevents the "missed the 4th kit in the stack" failure mode.
 *
 *   - Treatment is a NUMBER-FORMAT GLYPH PREFIX (not a CF rule). The cell's
 *     underlying value stays the SKU exactly; the display renders "▣ <SKU>".
 *     Same trick used by the DIRECT divider's "▌  DIRECT" rendering. No
 *     chromatic cost, no font weight change (SKU column is already bold), and
 *     a glyph prefix is unambiguously visible at a glance — addresses the
 *     "italic and font-color were both barely visible" feedback from the
 *     pure-typography iteration earlier today.
 *
 *   - Cell value unchanged. `getValue()` returns the SKU, not "▣ <SKU>" —
 *     downstream code (lookups, exports, formulas) is unaffected.
 *
 *   - The marker is a typographic FACT ("this SKU is a kit"), not a workflow
 *     STATE ("act on this"). Stays for the row's lifetime. Kit Expansion
 *     sidebar card is the workflow action; this is just an at-a-glance label.
 *
 *   - Trade-off vs the CF approach: number format is per-cell, not CF-
 *     conditional, so it doesn't auto-sync with Kit Registry changes. Re-run
 *     this function (or wire it into insert paths in a v2 pass) to refresh.
 *     Sidebar button gives a one-click refresh entry point.
 *
 * On first run, also strips the legacy italic CF rule from earlier today's
 * ship-cycle so the two approaches don't double up.
 *
 * Idempotent — re-run safely.
 */
function setupKitRowHighlighting() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  // Migration: strip the deprecated italic CF rule if it's still on the sheet
  // from earlier today's CF-based iteration. Safe no-op if already gone.
  var rules = sheet.getConditionalFormatRules();
  var beforeCount = rules.length;
  rules = _stripKitSkuRule(rules);
  if (rules.length !== beforeCount) {
    sheet.setConditionalFormatRules(rules);
  }

  return refreshKitSkuMarkers();
}

/**
 * Walks every data row in All Orders, applies the "▣ " number-format prefix to
 * SKU cells whose value is in the Kit Registry, and clears the prefix from
 * non-kit cells (so a previously-marked SKU that gets re-typed cleans up).
 *
 * Skips: empty cells, the DIRECT boundary divider, header rows (col-A SKUs
 * that start with the "◈" SKU header glyph). One batched setNumberFormats
 * call writes all formats in a single API trip, no matter how many rows.
 */
function refreshKitSkuMarkers() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";

  // Build a Set of kit SKUs (normalized uppercase+trim) for fast lookup.
  // buildKitMap() returns an empty Map if Kit Registry is missing — safe degradation.
  var kitSkus = new Set();
  try {
    buildKitMap().forEach(function(_value, sku) {
      kitSkus.add(String(sku).toUpperCase().trim());
    });
  } catch (e) {
    return "❌ Kit Registry unavailable: " + e.message;
  }

  var startRow = Schema.dataStartRow;
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return "✅ No data rows to scan.";

  var range = sheet.getRange(startRow, Schema.cols.SKU, lastRow - startRow + 1, 1);
  var values = range.getValues();
  // Read existing formats so we can PRESERVE them on cells we don't own.
  // Critical: the DIRECT boundary row carries '"▌  "@' (the divider's bar-glyph
  // number-format prefix from Service Bay v6). If we blindly write '@' on every
  // non-kit row, we clobber the DIRECT divider's ▌. Preserve any cell whose
  // role is "not a regular SKU" (boundary / header / empty).
  var existingFormats = range.getNumberFormats();
  // Also read the NOTE column so we can suppress the marker on rows that are
  // already EXPANSION COMPONENTS of another kit (NOTE starts with "↳ from KIT-").
  // Without this, a sub-component that happens to also be a standalone kit in
  // the registry (common for sub-assemblies) would get the same ▣ as its parent,
  // producing visual noise where every row in the kit block looks like a kit.
  var notes = sheet.getRange(startRow, Schema.cols.NOTE, lastRow - startRow + 1, 1).getValues();
  var formats = [];
  var kitCount = 0;

  for (var i = 0; i < values.length; i++) {
    var raw = String(values[i][0] || "").trim();
    var upper = raw.toUpperCase();
    var isEmpty = !raw;
    var isBoundary = upper === Schema.boundaryMarker;
    var isHeader = raw.charAt(0) === "◈";
    var noteRaw = String(notes[i][0] || "").trim();
    var isExpansionComponent = noteRaw.indexOf("↳ from KIT-") === 0;

    if (isEmpty || isBoundary || isHeader) {
      // Preserve whatever was there (e.g. DIRECT divider's '"▌  "@' glyph).
      formats.push([existingFormats[i][0]]);
      continue;
    }
    if (isExpansionComponent) {
      // Component row inserted by Kit Expansion — never marked as a kit, even
      // if its SKU happens to be a registered kit on its own (sub-assemblies).
      formats.push(['@']);
      continue;
    }
    if (kitSkus.has(upper)) {
      formats.push(['"▣ "@']);
      kitCount++;
    } else {
      formats.push(['@']);
    }
  }

  range.setNumberFormats(formats);
  SpreadsheetApp.flush();
  return "✅ Kit markers refreshed — " + kitCount + " row(s) marked with ▣ prefix.";
}

/**
 * Per-row Kit SKU marker handler — applies/clears the "▣ " number-format prefix
 * on a single col-A cell when its value changes. Dispatched from Main.js's
 * onEditInstallable trigger.
 *
 * v2 hook (2026-05-19) for the user-edit case: picker types a SKU into col A,
 * and if it's a Kit Registry SKU, the ▣ glyph appears immediately. Same dispatch
 * pattern as locationUpdateOnEdit, prepQueueOnEdit — single-cell write per edit.
 *
 * Programmatic inserts (n8n doPost, Zoho Pull, Zoho propagation) do NOT fire
 * onEdit, so those paths get a batched refreshKitSkuMarkers() call at their
 * respective insert sites instead. This handler covers user edits only.
 *
 * Skips: edits off the All Orders sheet, edits outside col A, edits inside the
 * banner zone (rows 1-3), edits to boundary marker / header glyph cells.
 * Multi-cell edits (paste / autofill) supported — formats array matches range
 * dimensions; each cell evaluated independently.
 *
 * Best-effort — wrapped in try/catch upstream so any error stays contained.
 */
function kitSkuOnEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== MAIN_SHEET_NAME) return;

  // Only react to col-A-only edits (single col or multi-row paste within col A)
  var firstCol = e.range.getColumn();
  var lastCol  = firstCol + e.range.getNumColumns() - 1;
  if (firstCol !== Schema.cols.SKU || lastCol !== Schema.cols.SKU) return;

  // Skip banner rows entirely
  if (e.range.getRow() < Schema.dataStartRow) return;

  // Build kit-SKU set (cheap — Kit Registry is typically a few hundred rows)
  var kitSkus = new Set();
  try {
    buildKitMap().forEach(function(_v, sku) {
      kitSkus.add(String(sku).toUpperCase().trim());
    });
  } catch (err) {
    return;   // Kit Registry unavailable — silent skip, no marker applied
  }

  var values  = e.range.getValues();
  // Read the NOTE column for the same row range to detect expansion components
  // (rows whose NOTE starts with "↳ from KIT-" — written by KitExpansion). Those
  // rows must never get the ▣ marker even if their SKU happens to be a
  // registered kit standalone (sub-assemblies are common).
  var noteRange = sheet.getRange(e.range.getRow(), Schema.cols.NOTE, values.length, 1);
  var notes = noteRange.getValues();
  var formats = [];
  for (var i = 0; i < values.length; i++) {
    var raw        = String(values[i][0] || "").trim();
    var upper      = raw.toUpperCase();
    var isEmpty    = !raw;
    var isBoundary = upper === Schema.boundaryMarker;
    var isHeader   = raw.charAt(0) === "◈";
    var noteRaw    = String(notes[i][0] || "").trim();
    var isExpansionComponent = noteRaw.indexOf("↳ from KIT-") === 0;

    if (isEmpty || isBoundary || isHeader) {
      formats.push(['@']);              // plain text
    } else if (isExpansionComponent) {
      formats.push(['@']);              // expansion component — never marked
    } else if (kitSkus.has(upper)) {
      formats.push(['"▣ "@']);          // kit marker
    } else {
      formats.push(['@']);              // plain text — clears stale marker
    }
  }
  e.range.setNumberFormats(formats);
}

/**
 * Surgical repair: restores the "▌  " glyph prefix on the DIRECT boundary row's
 * column-A cell. Run if the divider's ▌ ever disappears (e.g. earlier today's
 * refreshKitSkuMarkers bug clobbered it). Only touches the one cell's number
 * format — does not re-apply theme, banding, CF, or anything else.
 *
 * Safe to run any time; idempotent.
 */
function repairDirectDividerGlyph() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found";
  var boundary = _findBoundaryInSheet(sheet);
  if (boundary <= 0) return "❌ DIRECT boundary row not found";
  sheet.getRange(boundary, 1).setNumberFormat('"▌  "@');
  return "✅ DIRECT divider ▌ glyph restored on row " + boundary + ".";
}

/**
 * Consolidated CF rebuilder — wipes ALL theme-owned CF rules and rebuilds them
 * in a deterministic order. Called from applyBrandTheme() to ensure the full
 * Service Bay v6 CF rule set is present, idempotently.
 *
 * Rule set:
 *   1. STATUS — PENDING red, PREPARING yellow, SHIPPED green, CANCELED gray
 *   2. HAND low-stock — font-only #b71c1c bold (no bg, disciplined secondary signal)
 *   3. SHIP COST paid — yellow bg #fff4b0 + bold black ("money on the line" cue)
 *   4. Buyer Note — italic muted gold #8a7434 (subtle audit overlay)
 *
 * Kit SKU markers (▣ glyph prefix) are NOT a CF rule — they live in the cells'
 * number-format property and are applied/refreshed by refreshKitSkuMarkers().
 * The italic-CF approach used earlier 2026-05-18 was scrapped after the
 * pure-typography signal proved too subtle in real use. The strip map below
 * still includes the SKU column so any stale italic rule gets cleared on
 * theme re-apply (migration cleanup).
 *
 * Theme-owned rules are identified by their cell-range signature so non-theme
 * rules (anything the user might have added manually) are preserved.
 */
function _applyAllConditionalFormatting(sheet) {
  // Strip every theme-owned rule before rebuilding. Identify by range signature:
  // each theme rule lives on exactly one column (SKU A migration-strip-only,
  // NOTE E, STATUS F, HAND G, SHIP COST J). Anything ranging over multiple
  // columns is non-theme → keep.
  var existing = sheet.getConditionalFormatRules();
  var keep = [];
  var themeColumns = {};
  themeColumns[Schema.cols.SKU]       = true;   // legacy italic Kit-SKU rule → stripped, not rebuilt
  themeColumns[Schema.cols.NOTE]      = true;
  themeColumns[Schema.cols.STATUS]    = true;
  themeColumns[Schema.cols.HAND]      = true;
  themeColumns[Schema.cols.SHIP_COST] = true;
  existing.forEach(function(rule) {
    var ranges = rule.getRanges();
    var isThemeRule = ranges.some(function(r) {
      return r.getNumColumns() === 1 && themeColumns[r.getColumn()];
    });
    if (!isThemeRule) keep.push(rule);
  });

  keep.push.apply(keep, _buildStatusRules(sheet));
  keep.push(_buildHandLowStockRule(sheet));
  keep.push(_buildShipCostPaidRule(sheet));
  keep.push(_buildBuyerNoteRule(sheet));

  sheet.setConditionalFormatRules(keep);
}

function _buildStatusRules(sheet) {
  // STATUS column CF — saturated palette matching the banner emoji intensity
  // (🔴 PEND, 🟡 PREP, 🟢 SHIP, ⚫ CXL). All 4 states get cell bg + bold text;
  // CANCELED uses gray instead of strikethrough (user preferred "muted ignored"
  // semantics over the dramatic crossing).
  //
  // Smart formula: paints only on REAL data rows. Excludes:
  //   - Empty col-A rows (no SKU = empty data row)
  //   - The DIRECT boundary divider (col A === "DIRECT")
  //   - Header rows (col A starts with the "◈" SKU header glyph)
  // This prevents the DIRECT header's literal "PREPARING" header text from
  // being painted as if it were a live status cell.
  var statusRange = sheet.getRange(
    Schema.dataStartRow, Schema.cols.STATUS,
    BRAND.dataLast - Schema.bannerRows, 1
  );

  function buildFormula(statusValue) {
    return '=AND(' +
              'UPPER(TRIM($F' + Schema.dataStartRow + '))="' + statusValue + '",' +
              '$A' + Schema.dataStartRow + '<>"",' +
              'UPPER(TRIM($A' + Schema.dataStartRow + '))<>"' + Schema.boundaryMarker + '",' +
              'LEFT(TRIM($A' + Schema.dataStartRow + '),1)<>"◈"' +
            ')';
  }
  function rule(value, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(buildFormula(value))
      .setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([statusRange]).build();
  }
  return [
    rule(Schema.status.PENDING,   '#ffcdd2', '#b71c1c'),   // medium-light red + dark red
    rule(Schema.status.PREPARING, '#ffd400', BRAND.ink),   // full brand action yellow + black
    rule(Schema.status.SHIPPED,   '#c8e6c9', '#1b5e20'),   // medium green + dark green
    rule(Schema.status.CANCELED,  '#e0e0e0', '#424242')    // medium gray + near-black, NO strikethrough
  ];
}

function _buildHandLowStockRule(sheet) {
  // HAND low-stock — font-only red (no bg). Cell backgrounds are reserved for
  // the highest-priority alerts (status + paid shipping). HAND becomes a
  // "noted but not screaming" secondary signal. Darker red `#b71c1c` + bold
  // compensates for the lost bg by giving the font more visual weight.
  var handRange = sheet.getRange(
    Schema.dataStartRow, Schema.cols.HAND,
    BRAND.dataLast - Schema.bannerRows, 1
  );
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($G' + Schema.dataStartRow + '), $G' + Schema.dataStartRow + '<=20)')
    .setFontColor('#b71c1c').setBold(true)
    .setRanges([handRange]).build();
}

function _buildShipCostPaidRule(sheet) {
  // SHIP COST paid — soft brand-yellow bg + bold black text. Yellow because
  // "money on the line, has to ship before refund window closes." Picker scans
  // the SHIP COST column purely by color; any yellow cell = paid order.
  // Match condition: cell non-empty, not "FREE", contains a digit (rules out
  // header text like "SHIP COST" or any non-numeric label).
  var range = sheet.getRange(
    Schema.dataStartRow, Schema.cols.SHIP_COST,
    BRAND.dataLast - Schema.bannerRows, 1
  );
  var anchor = '$J' + Schema.dataStartRow;
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(' + anchor + '<>"", UPPER(TRIM(' + anchor + '))<>"FREE", ' +
      'REGEXMATCH(TO_TEXT(' + anchor + '), "[0-9]"))'
    )
    .setBackground('#fff4b0').setFontColor(BRAND.ink).setBold(true)
    .setRanges([range]).build();
}

function _buildBuyerNoteRule(sheet) {
  // Buyer Note — italic + muted gold-brown #8a7434, NO bg. Asymmetric design:
  // only buyer notes get styled; supervisor notes stay default. The buyer
  // note IS the exception (raw input from outside the system).
  // Edit workflow: when a supervisor rewrites a buyer note and removes the
  // "Buyer Note:" prefix as part of the edit, the CF rule stops matching →
  // italic/gold disappear → cell snaps back to default. Acts as live visual
  // feedback for "taking ownership" of the note.
  var noteRange = sheet.getRange(
    Schema.dataStartRow, Schema.cols.NOTE,
    BRAND.dataLast - Schema.bannerRows, 1
  );
  var anchor = '$E' + Schema.dataStartRow;
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(' + anchor + '<>"", REGEXMATCH(TO_TEXT(' + anchor + '), "(?i)^buyer note:"))'
    )
    .setItalic(true).setFontColor('#8a7434')
    .setRanges([noteRange]).build();
}

function _stripBuyerNoteRule(rules) {
  // Strip just the buyer-note rule from a rules array (preserves all others).
  // Identifies by range = NOTE column + formula contains "buyer note".
  return rules.filter(function(rule) {
    var bc = rule.getBooleanCondition();
    if (!bc) return true;
    var formula = (bc.getCriteriaValues() || [''])[0] || '';
    var ranges = rule.getRanges();
    var isNoteRange = ranges.some(function(r) {
      return r.getColumn() === Schema.cols.NOTE && r.getNumColumns() === 1;
    });
    return !(isNoteRange && formula.toLowerCase().indexOf('buyer note') !== -1);
  });
}

function _buildKitSkuRule(sheet) {
  // Kit Registry membership cue — italicizes the SKU cell when the SKU appears
  // in the Kit Registry. Pure typography, no chromatic signal. The picker
  // reads italic SKU as "this is a kit, check Kit Expansion if unsure."
  //
  // Multi-kit clusters (DIRECT orders with several kits) appear as a stack of
  // italic SKUs in col A — visually distinct against upright neighbors at a
  // glance, which is the failure-mode this rule prevents (missing one kit in
  // a cluster of consecutive rows).
  //
  // Guard conditions match the status rules' pattern (skip empty, skip DIRECT
  // divider, skip header rows starting with the "◈" SKU header glyph). The
  // MATCH is IFERROR-wrapped so a missing Kit Registry sheet degrades safely
  // to "no rows highlighted" instead of breaking the CF chain.
  //
  // INDIRECT wrap on the Kit Registry reference is REQUIRED — Sheets CF
  // formulas cannot reference another sheet by direct name (`'Kit Registry'!A:A`
  // throws "Conditional format rule cannot reference a different sheet"). The
  // INDIRECT runtime-resolves the reference, sidestepping the static check.
  var range = sheet.getRange(
    Schema.dataStartRow, Schema.cols.SKU,
    BRAND.dataLast - Schema.bannerRows, 1
  );
  var anchor = '$A' + Schema.dataStartRow;
  var formula =
    '=AND(' + anchor + '<>"", ' +
    'UPPER(TRIM(' + anchor + '))<>"' + Schema.boundaryMarker + '", ' +
    'LEFT(TRIM(' + anchor + '),1)<>"◈", ' +
    'IFERROR(ISNUMBER(MATCH(' + anchor + ', INDIRECT("\'Kit Registry\'!A:A"), 0)), FALSE))';
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setItalic(true)
    .setRanges([range]).build();
}

function _stripKitSkuRule(rules) {
  // Strip just the kit-SKU rule from a rules array (preserves all others).
  // Identifies by range = SKU column + formula contains "Kit Registry".
  return rules.filter(function(rule) {
    var bc = rule.getBooleanCondition();
    if (!bc) return true;
    var formula = (bc.getCriteriaValues() || [''])[0] || '';
    var ranges = rule.getRanges();
    var isSkuRange = ranges.some(function(r) {
      return r.getColumn() === Schema.cols.SKU && r.getNumColumns() === 1;
    });
    return !(isSkuRange && formula.indexOf('Kit Registry') !== -1);
  });
}

function _ensureDateFormula(sheet) {
  // Service Bay v6: B1 holds live date+time. The NOW() formula re-evaluates on
  // every spreadsheet recalc — which happens on every edit, every n8n insert,
  // every status change. Banner feels alive without any trigger overhead.
  // Force-write to ensure the canonical v6 format (older versions used just
  // TODAY() without the time, or static text — both should be upgraded).
  sheet.getRange('B1').setFormula('=TEXT(NOW(), "dddd, mmmm d, yyyy · h:mm AM/PM")');
}


// ═══════════════════════════════════════════════════════════════════════════════
// SERVICE BAY v6 HELPERS (2026-05-17 — added during VisualLab → production port)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Parameterized boundary lookup. Mirror of getBoundaryRow() but takes a sheet
 * argument so applyBrandTheme can target any sheet (production or test).
 * Strict equality on Schema.boundaryMarker ("DIRECT") — same contract as the
 * production getBoundaryRow().
 */
function _findBoundaryInSheet(sheet) {
  if (!sheet) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var values = sheet.getRange(1, Schema.cols.SKU, lastRow, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === Schema.boundaryMarker) return i + 1;
  }
  return -1;
}

/**
 * Creates (or refreshes) the hidden helper sheet `__SparkData` that drives the
 * banner's live System Pulse + TODAY total.
 *
 * Layout (all in the same hidden sheet):
 *   Row 1, A:X (24 cells) — hourly EVENT COUNTS for today. Each cell:
 *     =IFERROR(COUNTIFS('Activity Log'!A:A,">="&TODAY()+H/24,
 *                        'Activity Log'!A:A,"<"&TODAY()+(H+1)/24),0)
 *     A1 = 00:00-00:59, X1 = 23:00-23:59. IFERROR returns 0 when Activity Log
 *     is missing or empty — banner formulas degrade gracefully.
 *
 *   A3 — latest timestamp anywhere in Activity Log.
 *        =IFERROR(MAX('Activity Log'!A:A),0)
 *
 *   A4 — minutes since A3.
 *        =IF(A3>0,(NOW()-A3)*1440,-1)
 *     Returns -1 when there's no activity data so the banner can show
 *     "🔴 OFFLINE" explicitly instead of nonsense like "67000m ago".
 *
 * Sheet is hidden by default (try/catch — already-hidden throws).
 * Idempotent — safe to re-run; formulas are rewritten on every call.
 */
function _ensureSparkData(ss) {
  var name = '__SparkData';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  try { sheet.hideSheet(); } catch (e) { /* already hidden — fine */ }

  // Row 1: hourly counts (today, 00:00 → 23:59 by hour)
  var countFormulas = [];
  for (var h = 0; h < 24; h++) {
    countFormulas.push(
      "=IFERROR(COUNTIFS('Activity Log'!A:A,\">=\"&TODAY()+" + h + "/24," +
      "'Activity Log'!A:A,\"<\"&TODAY()+" + (h + 1) + "/24),0)"
    );
  }
  sheet.getRange(1, 1, 1, 24).setFormulas([countFormulas]);

  // System Pulse helpers (A3 timestamp, A4 minutes-since)
  sheet.getRange('A3').setFormula("=IFERROR(MAX('Activity Log'!A:A),0)");
  sheet.getRange('A4').setFormula("=IF(A3>0,(NOW()-A3)*1440,-1)");

  return sheet;
}

/**
 * Writes the live-cockpit formulas into banner row 1:
 *   B1:D1 → live date+time (NOW formula, already set by _ensureDateFormula)
 *   E1:F1 → SYSTEM PULSE: sync time + 🟢/🟡/🔴 ALIVE/IDLE/STALE + freshness
 *   G1:J1 → live status counts + TODAY total
 *
 * IMPORTANT: This OVERWRITES whatever was in E1 and G1. That's intentional —
 * the old updateLastSyncTimestamp() and updateOrderStatsInSheet() functions
 * were converted to no-ops on this date (2026-05-17) precisely so they don't
 * clobber these formulas after every n8n sync.
 */
function _setSystemPulseBannerFormulas(sheet) {
  // E1 — System Pulse. Color-codes by minutes-since-last-Activity-Log-event.
  //   <15min  → 🟢 ALIVE (system humming)
  //   15-60m  → 🟡 IDLE  (slowing but OK)
  //   >60m    → 🔴 STALE (needs attention)
  //   no data → 🔴 OFFLINE (Activity Log missing or empty)
  sheet.getRange('E1').setFormula(
    "=IF('__SparkData'!A4<0,\"⏱ Last sync · — · 🔴 OFFLINE\"," +
    "\"⏱ Last sync · \"&TEXT('__SparkData'!A3,\"h:mm AM/PM\")&\"   \"&" +
    "IF('__SparkData'!A4<15,\"🟢 ALIVE\"," +
    "IF('__SparkData'!A4<60,\"🟡 IDLE\",\"🔴 STALE\"))&" +
    "\"  \"&ROUND('__SparkData'!A4)&\"m ago\")"
  );

  // Stats banner — live status counts + TODAY total. Emoji bullets render their
  // own colors regardless of cell font color. ◢ glyph as a "drill-into" hint.
  // Anchor cell is Schema.cellStats (default G1; F1 since the 2026-05-19
  // layout compaction that hid cols I + J).
  //
  // CRITICAL: COUNTIF range MUST start at Schema.dataStartRow (typically F4),
  // not F:F. When Schema.cellStats was at G1 the old F:F range was fine, but
  // since the layout compaction moved cellStats into F1, a F:F range now
  // INCLUDES the formula cell itself — Sheets refuses with #REF! "Circular
  // dependency detected." Starting at F4 skips the entire banner zone (rows
  // 1-3) and only counts real data rows.
  var statusRangeStart = 'F' + Schema.dataStartRow;        // e.g. "F4"
  sheet.getRange(Schema.cellStats).setFormula(
    '="🔴 "&COUNTIF(' + statusRangeStart + ':F,"PENDING")&' +
    '"   🟡 "&COUNTIF(' + statusRangeStart + ':F,"PREPARING")&' +
    '"   🟢 "&COUNTIF(' + statusRangeStart + ':F,"SHIPPED")&' +
    '"   ⚫ "&COUNTIF(' + statusRangeStart + ':F,"CANCELED")&' +
    '"      ◢ "&IFERROR(SUM(\'__SparkData\'!A1:X1),0)&" TODAY"'
  );
}

/**
 * Test-sheet banding (used when applyBrandTheme is called with a sheet name).
 * Production runs go through refreshDynamicBandings() which targets MAIN_SHEET_NAME.
 */
function _applyTestSheetBanding(sheet, boundary) {
  if (!sheet) return;
  sheet.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });

  if (boundary > Schema.dataStartRow) {
    var ebayHeight = boundary - Schema.headerRow;
    var ebayBand = sheet.getRange(Schema.headerRow, 1, ebayHeight, Schema.dataWidth)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    ebayBand.setHeaderRowColor(BRAND.ink)
            .setFirstRowColor(BRAND.paper)
            .setSecondRowColor(BRAND.paperWarm);
  }

  var directHeaderRow = boundary + 1;
  var maxRow = Math.max(sheet.getMaxRows(), directHeaderRow + 5);
  if (directHeaderRow > 0 && directHeaderRow <= maxRow) {
    var directHeight = maxRow - directHeaderRow + 1;
    var directBand = sheet.getRange(directHeaderRow, 1, directHeight, Schema.dataWidth)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    directBand.setHeaderRowColor(BRAND.ink)
              .setFirstRowColor(BRAND.paper)
              .setSecondRowColor(BRAND.paperWarm);
  }
}
