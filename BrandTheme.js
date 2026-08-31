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

/**
 * MASTHEAD — row 1's state face (2026-08-30).
 *
 * The face is a pre-rendered GIF per state, picked live by formula. Pre-rendering
 * rather than rendering on demand is what removes the only unsolved piece: there is
 * no PNG renderer on the VPS, and =IMAGE() will not take SVG. Caddy already serves
 * /opt/hq-app via try_files, so this needs no /api route and NO Caddyfile edit —
 * the blast radius that killed the YouTube search box (2026-08-21).
 *
 * Faces are drawn by design-lab/shoot-masthead.js and scp'd to /opt/hq-app/mast/.
 *
 * ⚠⚠ THE ANIMATION QUESTION, SETTLED 2026-08-30 BY TWO PROBES. Read this before
 *    re-proposing an animated banner — it took two tests and one wrong conclusion.
 *
 *    1 · =IMAGE() DOES NOT ANIMATE. It renders an animated GIF's first frame and never
 *        advances it. Proven with a known-good GIF.
 *
 *    2 · insertImage() DOES ANIMATE. A floating OverGridImage plays the GIF properly.
 *        These are two different code paths and the first proves NOTHING about the
 *        second — I concluded "a Sheet cannot animate" from (1) alone, and the user was
 *        right to push back. @ Note the asymmetry it implies: a Blob throws "blob format
 *        is unsupported", but a public URL or data URL works.
 *
 *    3 · ⚠⚠ BUT A FLOATING IMAGE SCROLLS WITH THE GRID. It is anchored to a CELL, not to
 *        the viewport, so it does NOT stay on the frozen banner — scroll one page and
 *        the masthead slides away. That is what rules it out HERE, not the animation.
 *
 *    NET: motion is possible in a Sheet, but not in a PINNED position. Row 1 is frozen,
 *    so the banner can only use =IMAGE() — pinned, formula-driven, and still. The faces
 *    are PNG for that reason. The ruling that puts smooth motion on the Floor Board
 *    survives, but for a sharper reason than "cells cannot animate".
 *
 *    ⭐ The animated GIFs are still on the VPS beside the PNGs, so any future surface
 *      where scrolling does not matter can use them with no rework.
 *
 * ⚠ VERSION THE FILENAME. Sheets caches images per-URL, so re-drawing the art at the
 *   same path can leave the old frame showing indefinitely. Bump `version` instead.
 *
 * ⚠⚠ A MISSING FACE ANSWERS 200, NOT 404 — Caddy's try_files serves the Floor Board's
 *    HTML (303KB, text/html) for any unmatched path under /mast/. IMAGE() cannot decode
 *    that, so the formula errors and IFERROR falls back to the text "HQ" chip: safe, and
 *    visible. But it is the same "success status, failure content" shape as Zoho's
 *    200-with-an-error-body and the /exec error page n8n banked as success for weeks.
 *    The practical consequence: bumping `version` here WITHOUT scp-ing the new files
 *    turns every face into the text chip at once. Ship the art first — and
 *    design-lab/ship-masthead.sh byte-verifies against what the server actually serves.
 */
var MASTHEAD = {
  baseUrl:      "https://hq.yassinqurabi.com/mast/",
  // ⚠⚠ BUMP THIS ON EVERY RE-RENDER. Sheets caches =IMAGE() per URL, so re-drawing the
  //    art at the SAME path leaves the OLD frame on screen indefinitely — which is exactly
  //    what happened on 2026-08-30: the sky kept showing a broken-logo build long after
  //    the file was replaced, and it read as "the images are broken" when they were fine.
  //    This rule was written in the docblock above and then violated by its author.
  version:      "v5",  // 2026-08-31 — the horizon/mark fix, the seam vignette, 280x56
                       // ⚠ BUMPED ONLY AFTER ship-masthead.sh byte-verified all 144
                       //   v4 faces against the served copy. Reversed, every face
                       //   becomes the text chip at once.
  ext:          "png",  // ⚠ NOT gif — Sheets shows only a GIF's first frame (see above)
  imgH:         56,    // px — mode-4 explicit sizing, so nothing letterboxes
  imgW:         280,   // px — EQUALS A1:C1 (107+70+103), asserted by setupMasthead
                       // ⚠ MEASURED, NOT READ OFF A COMMENT. Schema's layout note says
                       //   "A1:C1 (287px)" and the live sheet is 280 — the comment was
                       //   aspirational and I sized the v4 art to it. The assertion
                       //   caught the 7px gap on its first real run, which is exactly
                       //   what it exists for. The art matches the SHEET; the sheet's
                       //   widths belong to the table below and are the operator's.
  restAccent:   "#7e8894",  // the cool tone the rest face wears — shared with its curve
  // ⚠ The neutral row-1 ink. D1 and E1 wear it so the face's state word is the ONLY
  //   coloured thing in the row. Three yellows meant nothing led.
  quietInk:     "#e8e8e8",
  // BANDS BLEED. TABLES DON'T. The masthead and the gold DIRECT divider run off the
  // right edge; the two table blocks sit between them as defined shapes ending at
  // column J. The white strip to the right stops reading as "the sheet ran out" and
  // starts reading as a margin between two full-bleed rules.
  // Nothing reads past dataWidth -- formats, CF and banding all stop there -- so K:T is
  // free canvas. Hidden columns take zero width, so with I:J hidden the band is
  // continuous from A to the viewport edge.
  // ⚠⚠ ROW 2'S LOOK IS AN EXPLICIT CHOICE, NOT SOMETHING DERIVED. It used to be
  //    inferred from WHERE THE PICKERS LIVE, because when that was written the two were
  //    the same question — moving the dropdowns off the banner was what freed the space.
  //    revertRow2() split them apart and the coupling broke immediately: the cream
  //    layout came back carrying the DARK-grounded logo, a black rectangle floating in
  //    a cream row. Two independent decisions must not share one input.
  //
  //      'cream'  the long-standing layout — A2:E2 merged, 65px, transparent logo
  //      'plate'  the nameplate — A2:F2 + G2:H2, 44px, charcoal, logo with its bar
  //
  //    'plate' additionally REQUIRES the pickers to have moved (G2 is half of the
  //    F2:G2 Shipping merge otherwise), so setupMasthead checks both.
  row2Style:    'cream',
  // The warm charcoal plate the 'plate' style wears. One dark family with row 1; gold
  // stays exclusive to the DIRECT divider so the sheet has exactly one gold object.
  row2Plate:    "#2a2724",
  nameplateInk: "#8a8f98",
  // ONE PREFIX, TWO NAMEPLATES. Retyping the house mark in two places is how the live
  // sheet ended up saying "HQMS . DIRECT ORDERS" while the code wrote
  // "HQ MS . DIRECT TABLE" -- a hand-edit the next applyBrandTheme would have
  // normalised away without anyone noticing which one was intended.
  namePrefix:   "HQMS",
  nameEbay:     "EBAY ORDERS",
  nameDirect:   "DIRECT ORDERS",
  // ⭐ THE FACE KEEPS THE HOUR. It cannot MOVE (a floating image scrolls off the frozen
  //    banner; =IMAGE() shows a GIF's first frame only) — but it can be LIT. Light is the
  //    one thing that reads correctly at ~1 frame per minute, because a day changes
  //    slowly anyway. Cold before dawn, amber as the floor opens, clean at noon, an ember
  //    horizon at 5pm, quiet by midnight. The Floor Board's night dial, on the sheet.
  rowHeight:    56,   // px — imgH matches, so the art fills the row edge to edge
  row2Height:   44,   // px — only applied once the pickers have moved off the banner
  // ---- ROW 2: THE DAY -----------------------------------------------------------------
  // ⭐ 826x65px of cream sat under the lit face doing nothing but holding a small logo —
  //    the biggest dead space in the banner, and the reason rows 1-2 read as two objects
  //    instead of one. It is now a SKY: the sun walks 6a->6p, the moon walks 6p->6a, stars
  //    at night, and a quiet 9-5 marker along the horizon.
  // ⚠ The eBay logo is COMPOSITED INTO the art, not layered over it — A2 can hold exactly
  //   one =IMAGE(), and a separate logo cell would mean splitting the A2:E2 merge and
  //   moving Schema.cellEmployeeId. Re-render with LOGO=0 to retire it; that is a brand
  //   decision, not a design one.
  // ⚠ OFF. The sky was beautiful and it cost the eBay table its NAME. The sheet is two
  //   stacked tables and row 2 is the eBay one's label, the counterpart to the gold
  //   "▌ DIRECT" divider — that cell was never empty space. Row 1 carries the creative
  //   load; row 2 carries an identity, and identity wins.
  sky:          false,
  skyCell:      "A2",
  skyH:         65,
  skyW:         826,    // 42 is too cramped for two lines — proven in the render
  lateMinutes:  180,   // matches the Floor Board's own 3h redline
  staleMinutes: 60     // matches the System Pulse's STALE tier
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
 *   - Banner F1 (Schema.cellStats): live COUNTIF status counts + an event-driven
 *     "work-shape" block bar (█ to-grab / ▒ in-prep / ░ headroom) + today total
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
    // ⚠⚠ MEASURED OFF THE LIVE SHEET 2026-08-31, NOT DESIGNED. These had drifted badly
    //    from the operator's layout, and a theme re-apply would have silently undone it:
    //    SALES_ORDER alone was 130 here against 232 live, which would CLIP the masthead's
    //    headline; SKU/LOCATION would have moved A1:C1 out from under the face image,
    //    and =IMAGE() mode 4 takes explicit pixel dimensions, so the art would stretch.
    //    setupMasthead now asserts A1:C1 == MASTHEAD.imgW and says so loudly.
    //    Re-measure with diagnoseMasthead() before changing any of these.
    sheet.setColumnWidth(Schema.cols.SKU,         107);   // A ┐
    sheet.setColumnWidth(Schema.cols.QTY,          70);   // B ├ = 280 = MASTHEAD.imgW
    sheet.setColumnWidth(Schema.cols.LOCATION,    103);   // C ┘
    sheet.setColumnWidth(Schema.cols.SALES_ORDER, 232);   // D — the masthead's headline
    sheet.setColumnWidth(Schema.cols.NOTE,        307);   // E — the pulse (MUST NOT MOVE)
    sheet.setColumnWidth(Schema.cols.STATUS,      130);  // tightened from 250 → 130
    sheet.setColumnWidth(Schema.cols.HAND,        100);
    sheet.setColumnWidth(Schema.cols.LEFT,        107);
    sheet.setColumnWidth(Schema.cols.SHIPPING,    180);
    sheet.setColumnWidth(Schema.cols.SHIP_COST,    90);

    // ── ROW HEIGHTS ──
    // Lab values that produced the right visual rhythm. Set BEFORE _style*
    // calls so the styled rows have the right height when their content lands.
    // ⚠ setupMasthead OWNS row 1's height (MASTHEAD.rowHeight). This used to set 42
    //   while setupMasthead set 52, so the banner's height depended on which ran last.
    sheet.setRowHeight(1, MASTHEAD.rowHeight);
    sheet.setRowHeight(2, 65);   // logo + Pick-ID badges; setupMasthead drops it to 44
                                 // once the pickers move off the banner
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

  // ⚠ RETIRED 2026-08-31. Its comment below describes a layout that no longer exists:
  //   B1:D1 has not been the merge since the masthead shipped — B1 now sits INSIDE
  //   A1:C1, so this formula wrote a date into a cell nobody could see.
  //   The date's job belongs to the face (which carries the hour as LIGHT) and to D1.
  return "⚠ RETIRED — B1 is inside the A1:C1 masthead merge, so this wrote an " +
         "invisible formula. The date is carried by the face and D1 now. " +
         "Run setupMasthead() instead.";
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
function setupEbayLogo(plate) {
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Restoring the eBay logo");
    if (_denied) return _denied;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  // ⚠⚠ SELF-HOSTED SINCE 2026-08-30, AND THAT IS THE POINT. This pointed at
  //    Wikimedia's "1200px-EBay_logo.svg.png" thumbnail for months. Wikimedia now
  //    ALLOWLISTS thumbnail sizes and that one 400s with "Use thumbnail sizes listed
  //    on w.wiki/GHai" — 500 and 1280 still work, 320/640/800/1024/1200 do not.
  //
  //    ⚠ It failed SILENTLY for an unknown length of time because Sheets had the image
  //      CACHED: the cell kept showing a logo whose URL had already died, and only went
  //      blank once something rewrote A2 and busted the cache. A dependency can rot
  //      long before you find out.
  //
  //    The logo is ours now, served from the same VPS directory as the masthead faces
  //    (Caddy already serves /opt/hq-app via try_files — no route, no Caddyfile edit).
  //    An upstream policy change can no longer take the eBay table's LABEL off the sheet.
  // ⭐ TWO ASSETS, CHOSEN BY THE LAYOUT — because a logo's GROUND has to match the cell
  //    it sits in, and row 2's ground changes with the nameplate.
  //
  //    ebay-v1  transparent, sized for today's CREAM row 2 (65px tall)
  //    ebay-v2  the plate colour baked in, with the ▌ bar drawn at its left edge in
  //             #8a8f98 — so row 2 and the gold DIRECT divider rhyme on BOTH sides:
  //             bar + name on the left, count on the right. Two nameplates, one grammar.
  //
  // ⚠ The bar has to live INSIDE the PNG. A cell cannot hold both an =IMAGE() and a
  //   number-format prefix, which is how the divider draws its ▌ — so the only way to
  //   give row 2 the same mark is to composite it into the art.
  // ⚠⚠ THE CALLER DECIDES, because the ASSET has to match the CELL it lands on and
  //    nothing else. This used to read Schema.pickIdA1() — i.e. it inferred the look
  //    from where the dropdowns were — and the moment those two decisions came apart it
  //    painted the dark-grounded logo onto the cream row. A logo's ground and a
  //    dropdown's address are unrelated facts.
  var plateMode = (arguments.length > 0)
    ? !!plate
    : (MASTHEAD.row2Style === 'plate');
  var logoUrl   = MASTHEAD.baseUrl + (plateMode ? "ebay-v2.png" : "ebay-v1.png");

  // ⚠ 500x201 source = 2.49:1. An older call passed (32, 120) = 3.75:1, so the logo had
  //   been horizontally STRETCHED all along. Both sizes below hold their true aspect.
  // ⚠ 40px was sized for row 2 at 65px. The nameplate row is 44px, so the plate asset
  //   ships at 30px — 40 would leave 4px of margin and read as jammed.
  var h = plateMode ? 30 : 40;
  var w = plateMode ? 80 : 100;   // v2 carries the bar + a gap, so it is wider per unit height

  sheet.getRange("A2").setFormula(
    '=IFERROR(IMAGE("' + logoUrl + '", 4, ' + h + ', ' + w + '),"eBay")'
  );
  return "✅ eBay logo restored to A2 — self-hosted, true aspect, " +
         (plateMode ? "plate ground + ▌ bar (nameplate layout)." : "transparent (cream layout).");
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
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Protecting the sheet structure");
    if (_denied) return _denied;
  }

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
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Removing structure protection");
    if (_denied) return _denied;
  }

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


// =======================================================================================
// THE ALL-ORDERS LOCK — hard protection, scoped to one sheet
// =======================================================================================
//
// WHY (2026-08-28): a SALES_ORDER cell was overwritten by hand on a row that had
// already been picked and shelf-counted. A slip, not an intrusion — the right intent,
// the wrong cell. `protectSheetStructure` above is WARNING-ONLY: a dialog you click
// through. This is a real wall.
//
// ⭐ WHY A WALL IS NOW POSSIBLE. It was ruled out on the grounds that hard protection
// needs an editor list and would break the sidebar for employees. Measured in incognito
// 2026-08-29: an anonymous link-editor gets NO sidebar and NO ⚙️ menu — no
// google.script.run at all. They can only type in cells. So there is no asymmetry to
// create, and "block every anonymous editor on the identity columns" is exactly the
// policy wanted. Every legitimate writer runs as the OWNER:
//
//   n8n doPost inserts · Telegram · /kits          → owner  ✅
//   onEditInstallable handlers (LiveSync, guards)  → owner  ✅ (installable = owner)
//   your sidebar / the editor                      → owner is never restricted
//   n8n's DIRECT Sheets-API writes                 → ❌ a NAMED account — see below
//   an anonymous employee                          → ❌ blocked, which is the point
//
// ⚠⚠ SHEET-LEVEL WITH CARVE-OUTS, NOT THREE COLUMN PROTECTIONS. This flips the default
// to DENY: a column added later is protected automatically instead of silently open —
// the same posture as DOPOST_LOCK_FREE, where anything unnamed keeps the lock. It also
// stops accidental row/column DELETION, which lives right next to the 08-28 slip. And
// it is scoped: Prep Queue, Out of Stock, Location Update and the employees' own temp
// tabs are untouched, so the manual-SKU workflows that live there keep working.

var ALL_ORDERS_LOCK = {
  tag: "HQ-LOCK",

  // Script Property holding the Google account n8n's Sheets credential runs as.
  // ⚠⚠ THE INSTALLER REFUSES WITHOUT IT, and that refusal is the whole point:
  //    `E5. Delete SHIPPED Row` in the eBay orders workflow writes to All Orders
  //    DIRECTLY through the Sheets node — not through Apps Script — so it is a
  //    NAMED account, not the owner. Protection blocks non-editors from structural
  //    changes, so forgetting this stops the ~1 AM shipped-row sweep. The symptom
  //    would be "the sheet is filling up with shipped rows", noticed days later.
  //    Fail loudly up front instead.
  n8nAccountKey: "N8N_SHEETS_ACCOUNT",

  // Set the property to this literal to assert, deliberately, that no external
  // named account writes to All Orders and none needs an exception.
  noneSentinel: "none"
};


/**
 * Report the current lock state WITHOUT changing anything. Run this first, and
 * again after installing — the Run button shows no return value, so it logs.
 */
/**
 * True when a lock is installed AND its carve-out no longer covers a Pick ID cell's
 * full merged range. Cheap enough to run at the end of every setupMasthead().
 *
 * ⚠ Returns FALSE when nothing is locked — an absent lock is not a stale one, and a
 *   warning on an unlocked sheet would train the reader to skip the line.
 */
function _lockNeedsRefresh(sheet) {
  try {
    var mine = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .filter(function (p) {
        return String(p.getDescription() || "").indexOf(ALL_ORDERS_LOCK.tag) === 0;
      });
    if (!mine.length) return false;
    var open = mine[0].getUnprotectedRanges().map(function (r) { return r.getA1Notation(); });
    return [Schema.pickIdA1(), Schema.pickIdA1('adjustment')].some(function (a1) {
      var mg = sheet.getRange(a1).getMergedRanges();
      var want = (mg && mg.length) ? mg[0].getA1Notation() : a1;
      return open.indexOf(want) === -1;
    });
  } catch (e) { return false; }
}

function describeAllOrdersLock() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { console.log("❌ Main sheet not found."); return "❌ Main sheet not found."; }

  var out = ["── ALL ORDERS LOCK ──"];
  var sheetProts = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var mine = sheetProts.filter(function (p) {
    return String(p.getDescription() || "").indexOf(ALL_ORDERS_LOCK.tag) === 0;
  });

  if (!mine.length) {
    out.push("state:      NOT LOCKED (no " + ALL_ORDERS_LOCK.tag + " sheet protection)");
  } else {
    var p = mine[0];
    out.push("state:      LOCKED");
    out.push("editors:    " + (p.getEditors().map(function (u) { return u.getEmail(); }).join(", ") || "(owner only)"));
    out.push("open cells: " + p.getUnprotectedRanges().map(function (r) { return r.getA1Notation(); }).join(" · "));
  }

  // ⚠⚠ A CARVE-OUT THAT NAMES ONLY PART OF A MERGE LOCKS THE WHOLE MERGE.
  //    Sheets requires write access to the ENTIRE merged range, so an unprotected "F2"
  //    against a merged F2:G2 leaves the Shipping picker unusable for staff — and
  //    perfectly usable for the owner, because removeEditors() ignores the owner. That
  //    is this file's oldest trap wearing a new face.
  //
  // ⭐ IT IS A SEQUENCING BUG, WHICH IS WHY A DETECTOR BEATS AN INSTRUCTION.
  //    protectAllOrdersSheet() resolves the merge correctly — but only the merge that
  //    exists WHEN IT RUNS. Run it before setupMasthead() rebuilds row 2 and it carves a
  //    lone cell, then the merge appears underneath it and nothing complains. Happened
  //    2026-08-31: lock at 12:59:48, merge restored at 1:01:03.
  if (mine.length) {
    var openA1 = mine[0].getUnprotectedRanges().map(function (r) { return r.getA1Notation(); });
    [Schema.pickIdA1(), Schema.pickIdA1('adjustment')].forEach(function (a1) {
      try {
        var mg = sheet.getRange(a1).getMergedRanges();
        var want = (mg && mg.length) ? mg[0].getA1Notation() : a1;
        if (openA1.indexOf(want) === -1) {
          out.push("⚠⚠ PICKER " + a1 + " IS MERGED AS " + want + " BUT THE CARVE-OUT SAYS '" +
                   (openA1.filter(function (x) { return x.indexOf(a1) === 0; })[0] || "nothing") +
                   "' — staff CANNOT edit it. Re-run protectAllOrdersSheet().");
        }
      } catch (e) { /* never let the reporter throw */ }
    });
  }

  var acct = "";
  try { acct = PropertiesService.getScriptProperties().getProperty(ALL_ORDERS_LOCK.n8nAccountKey) || ""; } catch (e) {}
  out.push("n8n acct:   " + (acct || "⚠ NOT SET — the installer will refuse"));

  var warn = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(function (p) {
    return String(p.getDescription() || "").indexOf("HQ-STRUCTURE") === 0;
  }).length;
  out.push("warn-only:  " + warn + " HQ-STRUCTURE range protection(s) (kept — they still warn YOU)");

  var msg = out.join("\n");
  console.log(msg);
  return msg;
}


/**
 * Install the lock. Idempotent — strips any prior HQ-LOCK protection first, so
 * re-running refreshes the carve-outs rather than stacking.
 *
 * WHAT STAYS EDITABLE (everything else on this sheet is locked):
 *   NOTE   — notes, holds, and now location corrections
 *   STATUS — PENDING → PREPARING → SHIPPED, the floor's main action
 *   LEFT   — the picker's shelf count
 *   Pick ID for Shipping + Pick ID for Adjustment — printing hard-refuses without them
 *
 * ⚠ LOCATION IS DELIBERATELY LOCKED (user's call 2026-08-29). It is auto-filled, and a
 *   correction goes in the NOTE. A real shelf change belongs on the Location Update sheet.
 *
 * ⚠ KNOWN AND ACCEPTED: the carve-outs are whole-column-from-dataStartRow, so the DIRECT
 *   header row's NOTE/STATUS/LEFT label cells fall inside them and stay editable.
 *   Excluding them would need ranges recomputed every time the boundary moves, which it
 *   does all day. The cells are cosmetic, getBoundaryRow() reads only column A (locked),
 *   and protectSheetStructure()'s warning-only protection still covers those rows.
 */
function protectAllOrdersSheet() {
  // ⚠ Owner-only, and deliberately NOT reachable through the owner bridge — see
  //   _obRequireOwner. A control that decides who may edit must not be runnable by the
  //   people it constrains.
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Locking All Orders");
    if (denied) return denied;
  }

  var props = PropertiesService.getScriptProperties();
  var acct = String(props.getProperty(ALL_ORDERS_LOCK.n8nAccountKey) || "").trim();

  // ⚠⚠ THE REFUSAL THAT PROTECTS THE NIGHTLY SWEEP. See n8nAccountKey above.
  if (!acct) {
    return "❌ REFUSED — set the Script Property '" + ALL_ORDERS_LOCK.n8nAccountKey + "' first.\n\n" +
           "n8n's `E5. Delete SHIPPED Row` writes to All Orders DIRECTLY via the Sheets\n" +
           "API, as a NAMED account rather than as the owner. Locking this sheet without\n" +
           "granting that account an exception STOPS the ~1 AM shipped-row sweep, and the\n" +
           "symptom (a sheet filling up with shipped rows) shows days later.\n\n" +
           "Set it to the account email from the n8n Google Sheets credential, or to the\n" +
           "literal 'none' if you have CONFIRMED in the live n8n UI that nothing writes to\n" +
           "All Orders outside Apps Script.";
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  // ---- 1. Strip any prior HQ-LOCK protection (idempotent re-run) ----
  var removed = 0;
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (String(p.getDescription() || "").indexOf(ALL_ORDERS_LOCK.tag) === 0) {
      try { p.remove(); removed++; } catch (e) { /* ignore */ }
    }
  });

  // ---- 2. Protect the whole sheet, then narrow the editor list to the owner ----
  var prot = sheet.protect().setDescription(
    ALL_ORDERS_LOCK.tag + ": All Orders — identity columns locked (2026-08-29)");
  prot.removeEditors(prot.getEditors());
  try { if (prot.canDomainEdit()) prot.setDomainEdit(false); } catch (e) { /* not a Workspace domain */ }

  // ---- 3. Grant the ONE named external writer its exception ----
  var granted = "none";
  if (acct.toLowerCase() !== ALL_ORDERS_LOCK.noneSentinel) {
    try { prot.addEditor(acct); granted = acct; }
    catch (e) {
      try { prot.remove(); } catch (e2) {}
      return "❌ Could not add '" + acct + "' as an editor: " + e + "\n" +
             "Nothing was locked — fix the account and re-run.";
    }
  }

  // ---- 4. Carve out what the floor actually works in ----
  // ⚠ Built from Schema, never from A1 literals, so a column move cannot silently
  //   open the wrong one — this is the file where a wrong range is a locked-out shift.
  var maxRow = sheet.getMaxRows();
  var span = maxRow - Schema.dataStartRow + 1;
  var open = [
    sheet.getRange(Schema.dataStartRow, Schema.cols.NOTE,   span, 1),
    sheet.getRange(Schema.dataStartRow, Schema.cols.STATUS, span, 1),
    sheet.getRange(Schema.dataStartRow, Schema.cols.LEFT,   span, 1)
  ];

  // The two Pick ID cells. ⚠ Pick ID for Shipping is a MERGE (F2:G2 today), and a
  // merge is only editable through its whole range — so resolve the merge rather
  // than hardcoding the second column, which has already moved once (2026-05-19).
  // ⚠ The carve-out must follow the pickers. A STALE one fails SILENTLY for staff
  //   while working perfectly for you, because removeEditors() ignores the owner —
  //   the 2026-08-29 "measured the wrong population" trap. Re-run
  //   protectAllOrdersSheet() after any flip, and verify as a STAFF account.
  [Schema.pickIdA1(), Schema.pickIdA1('adjustment')].forEach(function (a1) {
    var cell = sheet.getRange(a1);
    var merges = cell.getMergedRanges();
    open.push(merges && merges.length ? merges[0] : cell);
  });

  prot.setUnprotectedRanges(open);
  SpreadsheetApp.flush();

  return "✅ All Orders LOCKED" +
         (removed ? " (refreshed — " + removed + " prior)" : "") +
         "\n   editable: " + open.map(function (r) { return r.getA1Notation(); }).join(" · ") +
         "\n   n8n exception: " + granted +
         "\n\n⚠ Verify in an incognito window before trusting it: col D must refuse," +
         "\n   NOTE / STATUS / LEFT must accept, and both Pick ID dropdowns must work.";
}


// =======================================================================================
// THE WARNING LAYER — the identity columns get a speed bump, not a wall
// =======================================================================================
//
// ⚠⚠ WHY THE HARD LOCK WAS ABANDONED (2026-08-29, and this is the important note).
//   protectAllOrdersSheet works, and it was installed and verified — then removed, because
//   it broke the sidebar for everyone who is not the owner.
//
//   The measurement that misled me: an INCOGNITO window has no sidebar and no ⚙️ menu, so
//   I concluded hard protection breaks nothing. But the staff are not anonymous — they are
//   signed in with a COMPANY GOOGLE ACCOUNT, so they do have the sidebar and they use all
//   of it. I measured the wrong population and generalised from it.
//
//   google.script.run executes as the INVOKING USER. So under the lock, every sidebar
//   action that writes All Orders fails for them: kit expansion, Zoho pull, sort, update
//   locations, cleanup, add rows, bulk status.
//
//   ⭐ AND THE CONSTRAINT IS ABSOLUTE, not a gap in our knowledge. Confirmed against
//   Google's own docs and the canonical workaround (tanaikech, 2020): Sheets protection
//   checks WHO, never WHERE FROM. A cell edit and a setValues from the sidebar are the
//   same user doing the same operation. The ONLY documented way to separate them is to
//   run the write as the owner via a Web App round trip — rejected here for latency and
//   for the blast radius across seven working workflows.
//
//   ⚠ Also learned: removeEditors() SILENTLY IGNORES the owner's email. The owner can
//   never be locked out of their own protected ranges. That is why identityEditGuard
//   exists and is not optional — no protection scheme can ever restrict you.
//
// SO: a dialog before (here) and an undo plus an alert after (IdentityGuard.js).
// A speed bump and an alarm, not a wall — and the 08-28 event was a slip, which is
// exactly what a speed bump catches.
//
// ⭐ WHY THIS DOES NOT BREAK THE SIDEBAR. setWarningOnly produces a UI DIALOG ONLY.
//   There is no human to warn on a script write, so setValues from the sidebar passes
//   straight through. protectSheetStructure has covered rows 1-3 this way for months
//   without one complaint.
//
// ⚠ AND WHY IT ONLY WORKS NOW. The standing ruling is that a warning firing on the
//   NORMAL case gets clicked through and tuned out inside a week — the same reasoning
//   that killed the "not counted" marker. Until Phase 1 shipped, typing into column D
//   WAS normal: it was the only way to add a missing line. /missing, /replacement and
//   the Floor Board button changed that, so a hand-edit there is now genuinely abnormal.
//   Do not install this on a system where the door does not exist.

var IDENTITY_WARN = {
  tag: "HQ-IDENTITY",
  // ⚠ WHOLE COLUMNS, in A1 notation, NOT a row range. This table inserts rows all day —
  //   n8n at the top, kit expansion mid-table, Zoho pull into DIRECT, the 1 AM sweep
  //   deleting them again — and a fixed range would need re-applying every time the
  //   boundary moved. "A:A" is unbounded, so it covers every row that exists now and
  //   every row created later, for free and forever.
  columns: ["SKU", "QTY", "SALES_ORDER"]
};


/** 1 → "A", 27 → "AA". Small, but a wrong letter here protects the wrong column. */
function _iwColumnLetter(n) {
  var s = "";
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - r - 1) / 26;
  }
  return s;
}


/**
 * Install the warning dialog on the identity columns. Idempotent — strips any prior
 * HQ-IDENTITY protection first, so re-running refreshes rather than stacking.
 */
function warnOnIdentityEdits() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var removed = 0;
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (String(p.getDescription() || "").indexOf(IDENTITY_WARN.tag) === 0) {
      try { p.remove(); removed++; } catch (e) { /* ignore */ }
    }
  });

  var done = [];
  IDENTITY_WARN.columns.forEach(function (name) {
    var col = Schema.cols[name];
    if (!col) return;
    var a1 = _iwColumnLetter(col) + ":" + _iwColumnLetter(col);
    sheet.getRange(a1)
      .protect()
      .setDescription(IDENTITY_WARN.tag + ": " + name + " (" + a1 + ") — identity column, confirm before editing")
      .setWarningOnly(true);
    done.push(name + " " + a1);
  });

  return "✅ Warning installed on " + done.join(" · ") +
         (removed ? "  (refreshed — " + removed + " prior)" : "") +
         "\n\nA dialog now appears before a hand edit. Script writes — the sidebar, n8n," +
         "\nTelegram — pass through untouched, because there is no human to warn." +
         "\n\nWhole columns, so rows added or deleted later are covered automatically.";
}


/** Remove the warning layer. */
function unwarnIdentityEdits() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";
  var removed = 0;
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (String(p.getDescription() || "").indexOf(IDENTITY_WARN.tag) === 0) {
      try { p.remove(); removed++; } catch (e) {}
    }
  });
  return "✅ Removed " + removed + " " + IDENTITY_WARN.tag + " protection(s).";
}


/**
 * Remove the lock. One call, full reversal — the escape hatch that makes the
 * whole thing safe to try. Leaves HQ-STRUCTURE warning-only protections alone.
 */
function unprotectAllOrdersSheet() {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Unlocking All Orders");
    if (denied) return denied;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";

  var removed = 0;
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (String(p.getDescription() || "").indexOf(ALL_ORDERS_LOCK.tag) === 0) {
      try { p.remove(); removed++; } catch (e) {}
    }
  });
  return "✅ Removed " + removed + " " + ALL_ORDERS_LOCK.tag + " protection(s). All Orders is open again.";
}


/**
 * One-shot setter so the account can be configured without touching code.
 *
 * ⚠ THE ZERO-ARG WRAPPER LIVES IN OwnerBridge.js (setN8nSheetsAccountNow), beside the
 *   rollout it feeds. It was DUPLICATED here until 2026-08-30 — two top-level functions
 *   of the same name in Apps Script's single global scope, where file concatenation order
 *   is unspecified, so which body ran was undefined and editing one could leave the other
 *   executing. This stays the single WRITER; the wrapper delegates to it.
 */
function setN8nSheetsAccount(email) {
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, like every other lock control. This value becomes
  //   an EDITOR EXCEPTION on a protected sheet, so letting staff set it would hand them a
  //   way to grant edit rights the lock exists to withhold. Deliberately NOT bridged.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Setting the n8n Sheets account");
    if (_denied) return _denied;
  }

  var v = String(email || "").trim();
  if (!v) return "❌ Pass the account email, or the literal '" +
                 ALL_ORDERS_LOCK.noneSentinel + "'.";

  // ⚠⚠ THE "✅" HAS TO MEAN SOMETHING. The only values this property may ever hold are an
  //   email address (handed straight to prot.addEditor) or the literal sentinel. Anything
  //   else is GUARANTEED to fail later, inside protectAllOrdersSheet, where the error reads
  //   as a lock fault rather than a typo — and an unedited placeholder would be stored
  //   under a green checkmark, which is the "reassuring label on a dangerous state" this
  //   codebase already rules is a bug. Reject it here, where the message can name the fix.
  var isNone  = v.toLowerCase() === ALL_ORDERS_LOCK.noneSentinel;
  var isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
  if (!isNone && !isEmail) {
    return "❌ '" + v + "' is neither an email address nor the literal '" +
           ALL_ORDERS_LOCK.noneSentinel + "'.\n\n" +
           "   Use the account email from n8n's Google Sheets credential, or '" +
           ALL_ORDERS_LOCK.noneSentinel + "' if you have\n" +
           "   CONFIRMED in the live n8n UI that nothing outside Apps Script writes to\n" +
           "   All Orders. Storing anything else stops the ~1 AM shipped-row sweep.";
  }

  PropertiesService.getScriptProperties().setProperty(ALL_ORDERS_LOCK.n8nAccountKey, v);
  return "✅ " + ALL_ORDERS_LOCK.n8nAccountKey + " = " + v;
}


/**
 * The n8n-account state, shaped for the sidebar's Sheet Protection card.
 *
 * ⭐ WHY THIS EXISTS: until 2026-08-30 this one value was the ONLY step of the lock
 *   sequence that could not be done from the sidebar — it meant opening the editor,
 *   editing `var VALUE` in code, and running it. installAllOrdersLock() refuses without
 *   it, so that single gap forced the whole sequence into the editor.
 *
 * ⚠ Returns a SHAPE, never a formatted sentence. getN8nSheetsAccount below returns prose
 *   for the execution log; a client that had to parse that string would break the first
 *   time the wording changed.
 */
function getN8nSheetsAccountState() {
  if (typeof _obRequireOwner === "function" && _obRequireOwner("x")) {
    return { ok: false, owner: false, isSet: false, value: "" };
  }
  var v = "";
  try {
    v = String(PropertiesService.getScriptProperties()
          .getProperty(ALL_ORDERS_LOCK.n8nAccountKey) || "").trim();
  } catch (e) {}
  return {
    ok: true,
    owner: true,
    isSet: !!v,
    isNone: v.toLowerCase() === ALL_ORDERS_LOCK.noneSentinel,
    value: v,
    key: ALL_ORDERS_LOCK.n8nAccountKey
  };
}


/** Read it back, for confirming what is actually stored. */
function getN8nSheetsAccount() {
  var v = PropertiesService.getScriptProperties().getProperty(ALL_ORDERS_LOCK.n8nAccountKey);
  var out = ALL_ORDERS_LOCK.n8nAccountKey + " = " + (v || "(not set)");
  console.log(out);
  return out;
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
 * ⚰ relocateAdjustmentBadge — DELETED 2026-08-31. Do not resurrect.
 *
 * Shipped 2026-05-19 to move the Pick ID for Adjustment dropdown between an
 * "E2" and an "I2" layout. It had ZERO callers for its entire life (verified by
 * grep across every .js and .html before deleting) and was always editor-run.
 *
 * ⚠⚠ WHY IT HAD TO GO RATHER THAN JUST SIT THERE. It identified which cell held
 *    the Adjustment picker by asking "does I2 have a validation?" — a question
 *    that was unambiguous only while I2 was empty. The 2026-08-31 migration puts
 *    the SHIPPING picker at I2. So the first person to run this afterwards would
 *    have had it confidently identify the Shipping dropdown as the Adjustment
 *    one and drag it to E2, taking its validation and value with it. Silent, and
 *    it reads like success ("✅ relocated I2 → E2").
 *
 *    That is the same shape as every other trap in this file: a heuristic that
 *    was correct about the layout it was written for, left standing after the
 *    layout moved underneath it.
 *
 * Git history has the body if a future layout ever needs a mover — but write it
 * against Schema.pickIdA1(), never against "which cell happens to be validated".
 */

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
  // ⚠ It writes A1/D1/E1/F1. The pre-masthead layout put the stats in G1 and this
  //   string was never updated — a report that names the wrong cells sends the next
  //   reader to the wrong place.
  return "✅ Live banner formulas re-installed (A1 face · D1 headline · E1 pulse · F1 curve).";
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
  // ⚠⚠ ROW 1 HAS ONE OWNER NOW, AND IT IS setupMasthead. Until 2026-08-31 this
  //    function and setupMasthead BOTH styled row 1, and they disagreed about
  //    everything: 42px vs 56px, Roboto vs Oswald, brand yellow vs white, centred vs
  //    left. Whichever ran last won, so the banner's appearance depended on the order
  //    of two unrelated calls.
  //
  // ⚠⚠ AND IT COULD CLOBBER THE FACE. The old body ended with
  //    `if (!A1.getValue()) A1.setValue('HQ')` — a guard written when A1 held a text
  //    chip. A1 now holds the =IMAGE() formula, and getValue() on a formula cell
  //    returns its RESULT: for an image that is an empty string. So the guard read
  //    "empty", wrote the literal "HQ", and destroyed the masthead. Silently, and it
  //    looks exactly like the documented IFERROR fallback.
  //
  // This is left as the ground fill only, so a theme re-apply cannot repaint the
  // banner into the pre-masthead layout. setupMasthead does the rest.
  sheet.getRange(1, 1, 1, Schema.dataWidth).setBackground(BRAND.ink);
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
  //   Nameplate layout (2026-08-31 — pickers moved into the hidden columns):
  //     A2:F2 = eBay logo zone, G2:H2 = the EBAY nameplate, I2 = Shipping, J2 = Adjustment
  //     → Schema.cellAdjustmentId === "J2", Schema.cellEmployeeId === "I2"
  //
  // ⚠ The 'J2' arm below is NOT decoration. Without it this chain falls through
  //   to its A2:F2 default — which happens to be the right answer for the new
  //   layout, and being right by luck is not the same as being right. The next
  //   person to add a layout would inherit a branch that silently guesses.
  //
  // This function only PAINTS — it does not create or break merges.
  // ⚠ Branch on the RESOLVER, not the constant. Schema.cellAdjustmentId stays "H2"
  //   for the whole grace period even after the pickers move, so branching on it
  //   would paint the OLD layout over the new one.
  var adj = Schema.pickIdA1('adjustment');
  var logoZone;
  if (adj === 'E2') {
    logoZone = 'A2:D2';
  } else if (adj === 'H2') {
    logoZone = 'A2:E2';
  } else if (adj === 'J2') {
    logoZone = 'A2:F2';       // pickers are in the hidden columns; row 2 is the nameplate
  } else {
    logoZone = 'A2:F2';
  }
  sheet.getRange(logoZone).setBackground(BRAND.paperWarm);

  // Both Pick ID cells have data validation (dropdowns of allowed values).
  // We MUST NOT write VALUES to these cells during the theme apply, or the
  // validation will reject anything not in its list and abort the whole theme.
  // Style only — values stay untouched, dropdown stays functional.
  [Schema.pickIdA1(), adj].forEach(function(a1) {
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
 * ⚠⚠ DISABLED 2026-08-31 — IT REFUSES. Read this before re-enabling it.
 *
 * OPT-IN: rewrites the Shipping and Adjustment dropdowns as two-line badge values
 * ("SHIPPING\nYAwiss · 1" instead of "Shipping - YAwiss 1").
 *
 * ⚠⚠ THE TWO-LINE FORM BREAKS THREE REGEXES AT ONCE, PERMANENTLY, AND SILENTLY:
 *
 *   _currentPicker        ^Shipping\s*-\s*                 case-insensitive (ActivityLog.js)
 *   getBoardPickers       ^Shipping\s*-\s*                 case-insensitive (DashboardService.js)
 *   _extractPickIdData    ^(?:shipping|adjustments?)\s*[-:·]\s*
 *
 * (written without their delimiters on purpose — a literal \s*<slash>i inside a
 *  block comment contains the sequence that ENDS the comment, and node --check
 *  caught exactly that when this docblock was first written.)
 *
 * Every one of them wants a SEPARATOR after the label. In "SHIPPING\nYAwiss · 1"
 * the `\s*` happily eats the newline and then the pattern demands `-` and finds
 * `Y`. So all three fail together. The consequences are quiet and total:
 * getCurrentPicker() returns "" forever, so every Activity Log row goes
 * unattributed, the print gate refuses with a message about a cell nobody can
 * see, and the Floor Board's picker list offers nobody while reporting ok:true.
 *
 * Nothing has ever called this function — it has been opt-in and caller-free for
 * its whole life. It stays refused rather than deleted because its body is the
 * only worked example of reading and rewriting a validation rule in place, which
 * migratePickIdCells() was modelled on.
 *
 * If badges are ever genuinely wanted, the fix is NOT to re-enable this: it is to
 * change the three regexes FIRST, in one commit, with a test pinning them
 * together — the A-9/A-50 drift class this project has been bitten by three times.
 */
function setupPickIdBadges() {
  return "❌ REFUSED — setupPickIdBadges is disabled (2026-08-31).\n\n" +
         "It rewrites the Pick ID options into a two-line form that breaks the\n" +
         "gate regex in _currentPicker, getBoardPickers AND _extractPickIdData at\n" +
         "once. getCurrentPicker() would return \"\" permanently: unattributed\n" +
         "Activity Log rows, a print gate that refuses, and a Floor Board picker\n" +
         "list that offers nobody while reporting ok:true.\n\n" +
         "Read the docblock above this function before re-enabling it.";
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

/**
 * describePickIdCells — READ-ONLY. Zero args. Writes to the execution log.
 *
 * The Run button in the Apps Script editor does not display a return value, so this
 * console.log()s everything AND returns it. Same shape as auditBoardStockAdjustments
 * and diagnoseIdentityFlags, for the same reason.
 *
 * ⚠⚠ SAVE THIS OUTPUT BEFORE MIGRATING ANYTHING. The dropdown option lists exist
 *    NOWHERE in this codebase — they were authored by hand in the Sheets UI, and
 *    _rewritePickIdValidation only ever REWRITES a rule that already exists. The cell
 *    is the only copy of its own options, so this log is the only backup that will
 *    exist short of a Drive version-history restore.
 *
 * It answers three questions the 2026-08-31 migration depends on:
 *
 *   1 · Are the options SINGLE-line? A two-line form ("SHIPPING\nYAwiss · 1") defeats
 *       the gate regex in _currentPicker, getBoardPickers AND _extractPickIdData at
 *       once, so getCurrentPicker() returns "" permanently. That is a PRE-EXISTING
 *       outage, not the migration's fault — but it must be fixed first, separately,
 *       or the move gets blamed for it.
 *
 *   2 · Is the rule VALUE_IN_LIST and not VALUE_IN_RANGE? For a range-backed rule
 *       getCriteriaValues()[0] returns a *Range*, so getBoardPickers' loop reads
 *       undefined.length, never runs, and returns {ok:true, pickers:[]} — healthy
 *       looking, and it offers nobody.
 *
 *   3 · Are the destinations genuinely empty? I2/J2 carrying any validation or value
 *       means a half-finished migration, and writing over it would destroy evidence.
 *
 * Run describeAllOrdersLock() in the same sitting — its `editable:` line is the only
 * thing that says whether the lock is installed and which cells it currently opens.
 *
 * @returns {string} the same report that was logged
 */
function describePickIdCells() {
  var out = [];
  var say = function (s) { out.push(s); };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    var miss = "❌ Main sheet not found: " + MAIN_SHEET_NAME;
    console.log(miss);
    return miss;
  }

  var RULE = '════════════════════════════════════════════════════════════════════';
  var THIN = '────────────────────────────────────────────────────────────────────';

  say(RULE);
  say(' PICK ID CELLS · read-only · ' +
      Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd HH:mm') + ' Houston');
  say(RULE);
  say(' ⚠ SAVE THIS OUTPUT. The option lists exist nowhere in code.');
  say('');

  // Tallied across BOTH cells — these are what the gate reads.
  var twoLineTotal = 0;
  var rangeBacked  = 0;
  var unresolvable = 0;   // a list offering nothing the gate regex would ever accept
  var resetUnsafe  = [];  // cells where resetDailyPickIds' literal is not a valid option

  var describe = function (label, a1, gateRe, resetLiteral) {
    say(THIN);
    say(' ' + label + ' · ' + a1);
    say(THIN);
    if (!a1) { say('  ✗ address is undefined in Schema'); return; }

    var cell = sheet.getRange(a1);
    var col  = cell.getColumn();

    // JSON.stringify is deliberate — it is what makes an embedded \n VISIBLE.
    var raw = cell.getValue();
    say('  value          : ' + JSON.stringify(String(raw)));
    say('  gate regex     : ' + (gateRe.test(String(raw)) ? '✓ PASSES' : '✗ FAILS') +
        '   (' + String(gateRe) + ')');

    var merges = cell.getMergedRanges();
    say('  merge          : ' + (merges.length ? merges[0].getA1Notation() : 'none'));
    say('  column hidden  : ' + (sheet.isColumnHiddenByUser(col) ? 'YES' : 'no') +
        '   (col ' + col + ')');

    var dv = cell.getDataValidation();
    if (!dv) { say('  validation     : ✗ NONE'); return; }

    say('  validation     : ' + dv.getCriteriaType());
    say('    allowInvalid : ' + dv.getAllowInvalid());
    say('    helpText     : ' + JSON.stringify(dv.getHelpText() || ''));

    var cv = dv.getCriteriaValues() || [];
    say('    showDropdown : ' + cv[1]);

    // ⚠ VALUE_IN_RANGE returns a Range here, not an Array. Detect rather than assume.
    if (!Array.isArray(cv[0])) {
      rangeBacked++;
      var where = 'unknown';
      try { where = cv[0].getA1Notation(); } catch (e) {}
      say('    options      : ✗ RANGE-BACKED (' + where + ') — NOT a literal list.');
      say('                   getBoardPickers reads undefined.length on this and');
      say('                   returns {ok:true, pickers:[]} — offers nobody.');
      return;
    }

    var opts = cv[0];
    say('    options (' + opts.length + ')  :');
    for (var i = 0; i < opts.length; i++) {
      var o = String(opts[i]);
      var multi = o.indexOf('\n') !== -1;
      if (multi) twoLineTotal++;
      say('      [' + i + '] ' + JSON.stringify(o) + (multi ? '   ⚠ TWO-LINE' : ''));
    }

    // ---- can a picker be RESOLVED from this list at all? ------------------------
    // ⚠⚠ NOT the same question as "is one selected right now". The placeholder is
    //    DELIBERATELY rejected — _currentPicker says so in its own comment — so an
    //    empty getCurrentPicker() between the 4am reset and the first pick of the
    //    day is CORRECT, not an outage. What matters is whether the list holds
    //    anything the gate would accept once somebody picks.
    var matching = 0;
    for (var m = 0; m < opts.length; m++) if (gateRe.test(String(opts[m]))) matching++;
    if (!matching) unresolvable++;
    say('    gate-valid   : ' + matching + ' of ' + opts.length +
        (matching ? '' : '   ✗ NOTHING here would ever satisfy the gate'));

    // ---- what would the 4am reset ACTUALLY write, and would it land? ------------
    // ⚠⚠ IT CALLS THE REAL _pickIdPlaceholder RATHER THAN RE-DERIVING THE RULE.
    //    Apps Script puts every root file in one global scope, so the diagnostic can
    //    exercise the production function directly — which means it tests the real
    //    path and cannot drift from it. A third copy of a rule is how A-9 sorted
    //    after A-50 in three files until August.
    //
    // ⚠ The first version of this check asked "is resetDailyPickIds' LITERAL in the
    //   list?" That was right while the reset hardcoded one. It no longer does — it
    //   derives the placeholder off the cell — so the literal is now only a fallback
    //   for an unreadable rule, and gating on it flagged a cell that was already
    //   fixed. Test the mechanism, not a proxy for it.
    if (resetLiteral !== undefined) {
      var derived = null, derr = null;
      try { derived = _pickIdPlaceholder(cell, gateRe, resetLiteral); }
      catch (e) { derr = e; }

      if (derr) {
        resetUnsafe.push(a1 + ' (derivation threw)');
        say('    reset writes : ✗ could not derive a placeholder — ' + derr);
      } else {
        var lands = (opts.indexOf(derived) !== -1) || dv.getAllowInvalid();
        if (!lands) resetUnsafe.push(a1);
        say('    reset writes : ' + JSON.stringify(derived) +
            (lands ? '   ✓ lands' : '   ❌ REJECTED → setValue THROWS at 4am'));

        // Cosmetic, not gating: the hardcoded fallback is reached ONLY when the rule
        // cannot be read, so a mismatch here is drift worth tidying rather than a fault.
        if (opts.indexOf(resetLiteral) === -1) {
          say('    ⚠ cosmetic   : the fallback literal ' + JSON.stringify(resetLiteral));
          say('                   is not a member of this list. Harmless while the');
          say('                   derivation works — it is only used when the rule');
          say('                   is unreadable — but worth tidying in the UI.');
        }
      }
    }
  };

  // ⚠ These literals are a SECOND COPY of what resetDailyPickIds writes — they are
  //   inline strings there, not exported. That is the drift class this project keeps
  //   getting bitten by, so it is flagged rather than hidden: if that function ever
  //   changes what it writes, change these too. The durable fix is to have the reset
  //   write option [0] read off the cell and delete both copies.
  // ⚠ Report whichever cells are LIVE, via the resolver — not the constants. After the
  //   flip, Schema.cellEmployeeId still says "F2" while the data is at I2, so a
  //   constant-driven diagnostic would confidently describe two empty cells.
  var liveShip  = Schema.pickIdA1();
  var liveAdj   = Schema.pickIdA1('adjustment');
  var isNew     = (liveShip === Schema.cellEmployeeIdNext);
  var otherShip = isNew ? Schema.cellEmployeeId   : Schema.cellEmployeeIdNext;
  var otherAdj  = isNew ? Schema.cellAdjustmentId : Schema.cellAdjustmentIdNext;

  say(' PICK_ID_ADDR mode: ' + (isNew ? 'new (migrated)' : 'old (pre-migration)'));
  say('');

  describe('SHIPPING   · live at', liveShip,
           /^Shipping\s*-\s*/i,                 'Pick ID for Shipping');
  say('');
  describe('ADJUSTMENT · live at', liveAdj,
           /^adjustment(?:s)?\s*[-:·]\s*/i,     'Pick ID for Adjustment');

  // ---- destinations -------------------------------------------------------------
  say('');
  say(THIN);
  say(isNew ? ' THE VACATED CELLS (should be empty — the data has moved)'
            : ' DESTINATIONS (must be empty and unvalidated before migrating)');
  say(THIN);
  var destDirty = 0;
  [otherShip, otherAdj].forEach(function (a1) {
    var c = sheet.getRange(a1);
    var hasDv = !!c.getDataValidation();
    var val   = String(c.getValue());
    var mg    = c.getMergedRanges();
    if (hasDv || val) destDirty++;
    say('  ' + a1 + '  validation: ' + (hasDv ? '⚠ PRESENT' : 'none') +
        '   value: ' + JSON.stringify(val) +
        '   merge: ' + (mg.length ? mg[0].getA1Notation() : 'none') +
        '   col hidden: ' + (sheet.isColumnHiddenByUser(c.getColumn()) ? 'YES' : 'no'));
  });

  // ---- the live verdict ---------------------------------------------------------
  say('');
  say(THIN);
  say(' LIVE VERDICT');
  say(THIN);
  var picker = '';
  try { picker = String(getCurrentPicker() || ''); }
  catch (e) { picker = ''; say('  getCurrentPicker() THREW: ' + e); }
  say('  getCurrentPicker() : ' + JSON.stringify(picker) +
      (picker ? '' : '   ✗ EMPTY — every Activity Log row is unattributed'));

  var boardCount = 0;
  try {
    var bp = getBoardPickers() || {};
    var list = bp.pickers || [];
    boardCount = list.length;
    say('  getBoardPickers()  : ok=' + bp.ok +
        '  current=' + JSON.stringify(String(bp.current || '')) +
        '  pickers=' + list.length);
    for (var k = 0; k < list.length; k++) say('      [' + k + '] ' + JSON.stringify(String(list[k])));
    if (!list.length) say('      ✗ NO PICKERS OFFERED — the board drawer would be empty.');
  } catch (e2) {
    say('  getBoardPickers() THREW: ' + e2);
  }

  // ---- the gate -----------------------------------------------------------------
  // ⚠⚠ THE CRITERION IS "CAN A PICKER BE RESOLVED", NOT "IS ONE SELECTED".
  //    The first cut of this diagnostic gated on getCurrentPicker() being non-empty
  //    and duly FAILED on a Saturday night with the floor closed — because the 4am
  //    reset had correctly restored the placeholder and _currentPicker correctly
  //    rejects it. That would have sent someone hunting a pre-existing outage that
  //    does not exist, which is the precise mistake the plan warns against in the
  //    other direction. Plumbing health is what getBoardPickers reports; whether
  //    anyone has picked today is a fact about the shift, so it is informational.
  var blocking = (twoLineTotal > 0) || (rangeBacked > 0) || (unresolvable > 0) ||
                 (destDirty > 0)    || (resetUnsafe.length > 0);
  say('');
  say(RULE);
  say(' GATE');
  say(RULE);
  say('  two-line options     : ' + twoLineTotal + (twoLineTotal ? '   ❌' : '   ✓'));
  say('  range-backed rules   : ' + rangeBacked  + (rangeBacked  ? '   ❌' : '   ✓'));
  say('  lists with no valid  : ' + unresolvable + (unresolvable ? '   ❌' : '   ✓'));
  say('  destinations dirty   : ' + destDirty    + (destDirty    ? '   ❌' : '   ✓'));
  say('  reset-unsafe cells   : ' + (resetUnsafe.length
                                      ? resetUnsafe.join(', ') + '   ❌'
                                      : '0   ✓'));
  say('');
  say('  ── informational, NOT gating ──');
  say('  pickers offered      : ' + boardCount);
  say('  picker selected NOW  : ' + (picker ? JSON.stringify(picker)
                                            : 'none (expected off-shift / before the first pick)'));
  say('');
  if (!blocking) {
    // ⚠ The verdict has to know which side of the migration it is on. Printing
    //   "safe to proceed with the migration" AFTER it has already run is a stale label
    //   on a finished state — small, but this codebase's own ruling is that a reassuring
    //   message about the wrong state is a bug.
    say(isNew ? '  ✅ ALL HEALTHY — the pickers are migrated and every check passes.'
              : '  ✅ GATE PASSES — safe to proceed with the migration.');
    if (isNew) {
      say('');
      say('  ⏭ If protectAllOrdersSheet() has not been re-run since the move, its');
      say('     carve-out still opens the OLD cells. Nothing breaks today (the board');
      say('     writes through doPost, which runs as the owner) — but any sidebar');
      say('     control writing these cells would be refused for STAFF.');
    }
  } else {
    say('  ❌ GATE FAILS — STOP. Fix in its OWN commit before migrating, or the');
    say('     move will be blamed for a fault that was already there.');
  }
  say('');
  say('  Next: run describeAllOrdersLock() and save its `editable:` line too.');
  say(RULE);

  var report = out.join('\n');
  console.log(report);
  return report;
}

/**
 * migratePickIdCells — move the two Pick ID dropdowns into the hidden columns.
 *
 *   migratePickIdCells()        DRY RUN. Reads, checks, reports. Writes NOTHING.
 *   migratePickIdCells("APPLY") does it.
 *
 * ⚠⚠ VALIDATE EVERYTHING, THEN WRITE. NEVER WRITE THEN DISCOVER. The option lists exist
 *    nowhere in this codebase — they were authored by hand in the Sheets UI, so the cell
 *    is the only copy of its own options and a botched half-migration is not recoverable
 *    from source. Every preflight below is a hard refuse with ZERO writes behind it.
 *
 * ⚠⚠ IT DOES NOT HARDCODE setAllowInvalid(false). Both live rules are allowInvalid:false,
 *    and resetDailyPickIds writes a value into these cells at 4am — against a strict rule
 *    that does not list that value, setValue THROWS, into that function's own try/catch,
 *    where nobody reads it. Yesterday's picker then rolls forward onto today's work. So
 *    allowInvalid, showDropdown and helpText are all PRESERVED from the source rule, and
 *    the preflight refuses unless the placeholder the reset would derive is actually
 *    writable. (That failure was already live on H2 before this migration existed.)
 *
 * ⚠ COPY THE OPTIONS VERBATIM. Never author, normalize, dedupe, sort or trim them.
 *   _rewritePickIdValidation is the cautionary tale: it rewrote the options into a
 *   two-line form that broke three gate regexes at once, permanently and silently.
 *
 * ⚠ THE PROPERTY IS SET LAST, after every cell has been written AND read back. Until
 *   that moment every reader still resolves to the old cells, so an abort at any earlier
 *   point leaves a sheet that still works.
 *
 * ⚠ It deliberately does NOT call applyBrandTheme, setupMasthead, setupPickIdBadges or
 *   relocateAdjustmentBadge; does not unhide I or J; does not touch the lock; and leaves
 *   the F2:G2 merge alone. An empty unvalidated cell inside a surviving merge is inert,
 *   and breaking that merge would make rollback non-trivial.
 *
 * ⏭ AFTER A SUCCESSFUL APPLY: re-run protectAllOrdersSheet() — the lock's carve-out
 *   follows Schema.pickIdA1() and a stale one fails SILENTLY for staff while working
 *   perfectly for the owner. Then setupMasthead() to build row 2's nameplate.
 *
 * @param {string} mode - "APPLY" to write. Anything else is a dry run.
 * @returns {string} the report, also console.log'd
 */
function migratePickIdCells(mode) {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Migrating the Pick ID cells");
    if (denied) return denied;
  }

  var APPLY = (String(mode || '').toUpperCase() === 'APPLY');
  var out = [], say = function (s) { out.push(s); };
  var RULE = '════════════════════════════════════════════════════════════════════';

  say(RULE);
  say(' MIGRATE PICK ID CELLS · ' + (APPLY ? '⚠ APPLY (writing)' : 'DRY RUN (no writes)'));
  say(RULE);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { var m = '❌ Main sheet not found.'; console.log(m); return m; }

  var PLAN = [
    { label: 'SHIPPING',   from: Schema.cellEmployeeId,   to: Schema.cellEmployeeIdNext,
      gate: /^Shipping\s*-\s*/i,               fallback: 'Pick ID for Shipping' },
    { label: 'ADJUSTMENT', from: Schema.cellAdjustmentId, to: Schema.cellAdjustmentIdNext,
      gate: /^adjustment(?:s)?\s*[-:·]\s*/i,   fallback: 'Pick ID for Adjustment' }
  ];

  // ── PREFLIGHT. Every check runs before ANY write. ────────────────────────────────
  var refuse = [], captured = [];

  if (Schema._pickIdMode() === 'new') {
    refuse.push('PICK_ID_ADDR is already "new" — the migration has already run. ' +
                'Run rollbackPickIdCells() first if you need to redo it.');
  }

  PLAN.forEach(function (p) {
    var src = sheet.getRange(p.from), dst = sheet.getRange(p.to);
    var dv  = src.getDataValidation();

    if (!dv) { refuse.push(p.label + ': ' + p.from + ' has NO validation to move.'); return; }

    var type = String(dv.getCriteriaType());
    if (type !== 'VALUE_IN_LIST') {
      // ⚠ A VALUE_IN_RANGE rule returns a *Range* from getCriteriaValues()[0]. Copying it
      //   would leave getBoardPickers reading undefined.length — it never loops, returns
      //   {ok:true, pickers:[]}, and offers nobody while looking perfectly healthy.
      refuse.push(p.label + ': ' + p.from + ' is ' + type + ', not VALUE_IN_LIST. ' +
                  'A range-backed rule cannot be copied safely.');
      return;
    }

    var cv = dv.getCriteriaValues() || [];
    if (!Array.isArray(cv[0]) || !cv[0].length) {
      refuse.push(p.label + ': ' + p.from + ' has no readable option list.');
      return;
    }

    // destination must be genuinely untouched — this is the anti-half-migration gate
    if (dst.getDataValidation()) refuse.push(p.label + ': ' + p.to + ' ALREADY has a validation.');
    if (String(dst.getValue()))  refuse.push(p.label + ': ' + p.to + ' is not empty.');
    if (dst.getMergedRanges().length) {
      refuse.push(p.label + ': ' + p.to + ' is inside a merge (' +
                  dst.getMergedRanges()[0].getA1Notation() + ').');
    }

    var opts = cv[0].slice();          // verbatim copy, never re-authored
    var gateValid = 0;
    for (var i = 0; i < opts.length; i++) if (p.gate.test(String(opts[i]))) gateValid++;
    if (!gateValid) {
      refuse.push(p.label + ': no option satisfies the gate regex — a picker could ' +
                  'never be resolved from this list.');
    }

    // ⚠ Would the 4am reset still land after the move? It derives the placeholder off
    //   the cell, so ask the REAL function rather than restating its rule here.
    var placeholder = (typeof _pickIdPlaceholder === 'function')
      ? _pickIdPlaceholder(src, p.gate, p.fallback) : p.fallback;
    var allowInvalid = dv.getAllowInvalid();
    if (opts.indexOf(placeholder) === -1 && !allowInvalid) {
      refuse.push(p.label + ': the placeholder ' + JSON.stringify(placeholder) +
                  ' is not in the list and allowInvalid is false — resetDailyPickIds ' +
                  'would THROW at 4am. Fix the option list first.');
    }

    // ⚠⚠ CAN THE CURRENT VALUE EVEN BE COPIED? Found on the live sheet 2026-08-31:
    //    H2 holds "Pick ID for Adjustment" while its own option [0] is
    //    "Pick ID for Adjustment " (trailing space). The cell's value is NOT a member of
    //    its own list. Writing it into a fresh allowInvalid:false rule would THROW —
    //    after the FIRST cell had already migrated and cleared its source, leaving the
    //    Shipping dropdown gone from the banner and not yet live at I2. A half-migration.
    //
    //    So the value is classified here, not discovered mid-write:
    //      valid       copy it verbatim
    //      placeholder it fails the gate regex, so it means "unset" — write the list's
    //                  OWN placeholder instead. Same meaning, and it lands.
    //      real picker a name that is not in its own list. Do NOT substitute; skip it
    //                  and say so. Guessing at someone's identity is worse than blank.
    var val = String(src.getValue());
    var valueMode = 'valid';
    if (val && opts.indexOf(val) === -1 && !allowInvalid) {
      valueMode = p.gate.test(val) ? 'skip' : 'placeholder';
    }

    captured.push({
      label: p.label, from: p.from, to: p.to,
      opts: opts, showDropdown: cv[1], allowInvalid: allowInvalid,
      helpText: dv.getHelpText() || '', value: val,
      gateValid: gateValid, placeholder: placeholder, valueMode: valueMode
    });
  });

  captured.forEach(function (c) {
    say('');
    say(' ' + c.label + '   ' + c.from + '  →  ' + c.to);
    say('   value        : ' + JSON.stringify(c.value));
    say('   options (' + c.opts.length + ')  : ' + c.gateValid + ' satisfy the gate');
    for (var i = 0; i < c.opts.length; i++) say('       [' + i + '] ' + JSON.stringify(String(c.opts[i])));
    say('   allowInvalid : ' + c.allowInvalid + '   showDropdown: ' + c.showDropdown);
    say('   helpText     : ' + JSON.stringify(c.helpText));
    say('   4am reset    : would write ' + JSON.stringify(c.placeholder) + ' → ' +
        ((c.opts.indexOf(c.placeholder) !== -1 || c.allowInvalid) ? '✓ lands' : '❌ THROWS'));
    if (c.valueMode === 'placeholder') {
      say('   ⚠ the value    : ' + JSON.stringify(c.value) + ' is NOT in this list.');
      say('                    It fails the gate regex, so it means "unset" — the list\'s');
      say('                    own placeholder ' + JSON.stringify(c.placeholder) +
          ' will be written instead.');
      say('                    (Copying it verbatim would THROW against allowInvalid:false.)');
    } else if (c.valueMode === 'skip') {
      say('   ⚠⚠ the value   : ' + JSON.stringify(c.value) + ' is a REAL picker that is');
      say('                    not in its own list. It will be LEFT BLANK rather than');
      say('                    guessed at — set it again from the dropdown afterwards.');
    }
  });

  say('');
  say(RULE);
  if (refuse.length) {
    say(' ❌ REFUSED — ' + refuse.length + ' blocking problem(s). NOTHING was written.');
    refuse.forEach(function (r) { say('   • ' + r); });
    say(RULE);
    var rep = out.join('\n'); console.log(rep); return rep;
  }

  if (!APPLY) {
    say(' ✓ PREFLIGHT PASSES. No writes were made.');
    say('');
    say('   Re-run as migratePickIdCells("APPLY") to move them.');
    say('   Then: protectAllOrdersSheet()  ·  setupMasthead()');
    say(RULE);
    var rep2 = out.join('\n'); console.log(rep2); return rep2;
  }

  // ── APPLY. Per cell: write → flush → READ BACK AND ASSERT → only then clear source. ──
  var done = [];
  for (var k = 0; k < captured.length; k++) {
    var c = captured[k];
    var src = sheet.getRange(c.from), dst = sheet.getRange(c.to);

    var b = SpreadsheetApp.newDataValidation()
      .requireValueInList(c.opts, c.showDropdown !== false)
      .setAllowInvalid(c.allowInvalid);           // PRESERVED, never hardcoded
    if (c.helpText) b.setHelpText(c.helpText);

    dst.setDataValidation(b.build());
    // ⚠ Write only what the rule will accept — see valueMode in the preflight.
    var toWrite = (c.valueMode === 'placeholder') ? c.placeholder
                : (c.valueMode === 'skip')        ? ''
                : c.value;
    if (toWrite) dst.setValue(toWrite);
    c.wrote = toWrite;
    dst.setBackground(BRAND.ink).setFontColor(BRAND.yellow)
       .setFontFamily(BRAND.fontDisplay).setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

    SpreadsheetApp.flush();

    // read back — a write that reports success and did not land is the whole risk here
    var back = dst.getDataValidation();
    var bad  = [];
    if (!back) bad.push('no validation landed');
    else {
      var bv = back.getCriteriaValues() || [];
      if (!Array.isArray(bv[0]) || bv[0].length !== c.opts.length) bad.push('option count differs');
      else for (var q = 0; q < c.opts.length; q++) {
        if (String(bv[0][q]) !== String(c.opts[q])) { bad.push('option ' + q + ' differs'); break; }
      }
      if (back.getAllowInvalid() !== c.allowInvalid) bad.push('allowInvalid differs');
    }
    if (String(dst.getValue()) !== String(c.wrote || '')) bad.push('value differs');

    if (bad.length) {
      say('');
      say(' ❌ ' + c.label + ' FAILED THE READ-BACK: ' + bad.join(', '));
      say('   ' + c.from + ' was NOT cleared and PICK_ID_ADDR was NOT set, so every');
      say('   reader still points at the old cell and the sheet still works.');
      say('   Clear ' + c.to + ' by hand, then investigate before retrying.');
      say(RULE);
      var rep3 = out.join('\n'); console.log(rep3); return rep3;
    }

    // only now is it safe to let go of the original
    src.setDataValidation(null);
    src.clearContent();
    SpreadsheetApp.flush();
    done.push(c.label + ' ' + c.from + ' → ' + c.to);
  }

  // ⚠ LAST. Until this line every reader still resolves to the old cells.
  PropertiesService.getScriptProperties().setProperty('PICK_ID_ADDR', 'new');

  say(' ✅ MIGRATED — ' + done.join('  ·  '));
  say('   PICK_ID_ADDR = new');
  say('');
  say(' ⏭ NOW, IN THIS ORDER:');
  say('   1. protectAllOrdersSheet()   — the carve-out must follow the pickers.');
  say('      ⚠ Verify as a STAFF account. removeEditors() ignores the owner, so a');
  say('        stale carve-out fails silently for staff and works fine for you.');
  say('   2. setupMasthead()           — builds row 2 as the nameplate.');
  say('   3. describePickIdCells()     — confirm, and check I/J are still hidden.');
  say('');
  say(' ↩ Rollback: delete the PICK_ID_ADDR property (instant, no deploy), THEN');
  say('   rollbackPickIdCells().');
  say(RULE);
  var rep4 = out.join('\n'); console.log(rep4); return rep4;
}

/**
 * rollbackPickIdCells — the mirror. Moves the dropdowns back to the banner cells.
 *
 * ⚠⚠ THE PROPERTY GOES FIRST, AND THE ORDER IS LOAD-BEARING. Deleting it makes every
 *    reader — including the PINNED /exec, which cannot be redeployed instantly — resolve
 *    to the old cells on its next execution. Doing the cells first would leave a window
 *    where every surface points at cells that are already empty.
 *
 * The instant, no-deploy rollback is deleting the property ALONE. This function is only
 * needed to put the data back afterwards.
 */
/**
 * migratePickIdCellsAPPLY — the zero-arg door for the Run button.
 *
 * ⚠⚠ THE APPS SCRIPT RUN BUTTON CANNOT PASS ARGUMENTS. It calls the selected function
 *    with none, so migratePickIdCells("APPLY") is simply not runnable from the editor
 *    dropdown — it would execute as a dry run forever while looking like it applied.
 *    This project has walked into that trap before (checkPublishedTickNow,
 *    setMyPricePushPassphraseNow, importKitsNow, runSingleStockAdjustTest all exist for
 *    the same reason), which is why every one-shot here gets a zero-arg wrapper.
 *
 * ⚠ NAMED IN CAPITALS ON PURPOSE. The Run dropdown lists functions alphabetically, so
 *   this sits directly beneath migratePickIdCells — and the difference between the safe
 *   one and the writing one has to be readable at a glance in that list, not inferred.
 *
 * @returns {string} the migration report
 */
function migratePickIdCellsAPPLY() {
  return migratePickIdCells("APPLY");
}

/**
 * rollbackPickIdCellsAPPLY — same, for the mirror.
 *
 * ⚠ The INSTANT rollback is not this function. It is deleting the PICK_ID_ADDR Script
 *   Property: that alone sends every reader — including the pinned /exec, which cannot
 *   be redeployed quickly — back to the banner cells on its next execution, with no
 *   deploy at all. This only moves the data back afterwards.
 *
 * @returns {string} the rollback report
 */
function rollbackPickIdCellsAPPLY() {
  return rollbackPickIdCells("APPLY");
}

function rollbackPickIdCells(mode) {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Rolling back the Pick ID cells");
    if (denied) return denied;
  }
  var APPLY = (String(mode || '').toUpperCase() === 'APPLY');
  var out = [], say = function (s) { out.push(s); };

  say('=== ROLLBACK PICK ID CELLS · ' + (APPLY ? '⚠ APPLY' : 'DRY RUN') + ' ===');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) { var m = '❌ Main sheet not found.'; console.log(m); return m; }

  if (APPLY) {
    // FIRST — every reader goes back to the banner cells on its next execution.
    PropertiesService.getScriptProperties().deleteProperty('PICK_ID_ADDR');
    Schema._pickIdModeCache = null;
    say('  ✓ PICK_ID_ADDR deleted — all readers resolve to the old cells again.');
  } else {
    say('  would delete PICK_ID_ADDR first (this alone is the instant rollback)');
  }

  var PAIRS = [
    [Schema.cellEmployeeIdNext,   Schema.cellEmployeeId,   'SHIPPING'],
    [Schema.cellAdjustmentIdNext, Schema.cellAdjustmentId, 'ADJUSTMENT']
  ];

  PAIRS.forEach(function (p) {
    var from = sheet.getRange(p[0]), to = sheet.getRange(p[1]);
    var dv = from.getDataValidation();
    if (!dv) { say('  – ' + p[2] + ': nothing at ' + p[0] + ' to move back.'); return; }
    var cv = dv.getCriteriaValues() || [];
    if (!Array.isArray(cv[0])) { say('  ✗ ' + p[2] + ': ' + p[0] + ' is not a literal list.'); return; }

    say('  ' + p[2] + ': ' + p[0] + ' → ' + p[1] + '  (' + cv[0].length + ' options)');
    if (!APPLY) return;

    var b = SpreadsheetApp.newDataValidation()
      .requireValueInList(cv[0], cv[1] !== false)
      .setAllowInvalid(dv.getAllowInvalid());
    if (dv.getHelpText()) b.setHelpText(dv.getHelpText());
    to.setDataValidation(b.build());
    var v = String(from.getValue());
    if (v) to.setValue(v);
    SpreadsheetApp.flush();
    from.setDataValidation(null);
    from.clearContent();
    SpreadsheetApp.flush();
  });

  // ⚠⚠ ORDER IS LOAD-BEARING: setupMasthead() FIRST. It rebuilds row 2's merges, and
  //    the lock can only carve out the merge that exists when it runs — reversed, it
  //    names a lone cell, the merge appears underneath, and the picker silently refuses
  //    STAFF. That is exactly what happened on 2026-08-31.
  say(APPLY ? '  ⏭ NEXT, IN THIS ORDER:  1) setupMasthead()   2) protectAllOrdersSheet()'
            : '  Re-run as rollbackPickIdCellsAPPLY() to act.');
  var rep = out.join('\n'); console.log(rep); return rep;
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

/**
 * One nameplate, two callers. Row 2 names the eBay table; the gold divider names the
 * DIRECT one. Same grammar, same prefix, same count-suffix rule -- built here so the two
 * cannot drift into saying the house mark differently.
 *
 * A BLANK COUNT DROPS THE SUFFIX ENTIRELY. __SparkData A17/A18 use pubNumBlank, which
 * yields "" when the published key cannot be read -- so an unreadable payload renders
 * "HQMS . EBAY ORDERS" rather than "HQMS . EBAY ORDERS . 0 waiting". A reassuring label
 * on a wrong state is a bug; saying less is the honest degradation. A genuine 0 still
 * prints "0 waiting", because an empty table saying so is true.
 *
 * A FORMULA IN THE BOUNDARY ROW'S RIGHT MERGE IS SAFE. It already held a static string,
 * getBoundaryRow() matches column A only, and n8n's All-Orders readers take
 * UNFORMATTED_VALUE -- which returns the computed string exactly as today.
 */
function _nameplateFormula(label, sparkCell) {
  var SD = "'__SparkData'!";
  return '="' + MASTHEAD.namePrefix + ' \u00b7 ' + label + '"&' +
         'IF(' + SD + sparkCell + '="","",' +
         '" \u00b7 "&' + SD + sparkCell + '&" waiting")';
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

  rightMerge.setFormula(_nameplateFormula(MASTHEAD.nameDirect, 'A18'))
            .setBackground(BRAND.yellow)
            .setFontColor(BRAND.ink)
            .setFontFamily(BRAND.fontDisplay)
            .setFontWeight('bold')
            .setFontSize(10)
            .setHorizontalAlignment('right')
            .setVerticalAlignment('middle');

  // Black borders top + bottom across the entire row — frames the yellow band
  // so it reads as a defined section break (not just a colored row).
  // ⚠ Stops at dataWidth. A "bleed past J" version was tried and reverted the same day.
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
  //    centered since 2026-05-16 to fix DIRECT-table column-alignment drift).
  //    Size stays 10 = the table's uniform rhythm. The SO badge glyph gets
  //    its LARGER size from the painter (setupDuplicateSalesOrderHighlighting
  //    raises the CELL font to 14 on badge rows while the id text is pinned
  //    at 10 via its rich-text run) — do NOT bump this column-level size,
  //    that was tried 2026-07-14 and read as inconsistent with the table.
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
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Applying the buyer-note highlight");
    if (_denied) return _denied;
  }

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
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Applying the kit-SKU highlight");
    if (_denied) return _denied;
  }

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
  // ⚠ WRITES A PROTECTED SHEET. google.script.run runs as the INVOKING USER, so under
  //   the All Orders lock a staff call would be refused. Come back in through /exec,
  //   where doPost executes as the OWNER — see OwnerBridge.js.
  if (!_obIsOwner()) return _asOwner('refreshKitSkuMarkers', []);

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

  // Build kit-SKU set. Cache for 60s in CacheService — a paste of N rows
  // fires this handler once with a multi-row range, but rapid successive
  // edits (typing several SKUs in a row) would otherwise re-read the whole
  // Kit Registry sheet each time. Trade-off: for up to 60s after a Zoho
  // webhook adds a new kit, that brand-new SKU may briefly miss the ▣ —
  // acceptable since registry changes are rare relative to All Orders edits.
  var kitSkus = new Set();
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("kit_skus_v1");
    if (cached) {
      JSON.parse(cached).forEach(function(s) { kitSkus.add(s); });
    } else {
      buildKitMap().forEach(function(_v, sku) {
        kitSkus.add(String(sku).toUpperCase().trim());
      });
      cache.put("kit_skus_v1", JSON.stringify(Array.from(kitSkus)), 60);
    }
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
 *   5. Identity — deep red #b71c1c on cols A+D when a row's identity is broken
 *      (missing half, never received, or duplicated). See _buildIdentityRules.
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
  // ⚠ THE HOLD RULES GO BEFORE THE BUYER-NOTE RULE, AND ORDER IS LOAD-BEARING.
  // Sheets applies the FIRST matching rule per cell. A buyer note that happens
  // to contain the word "hold" ("please hold for pickup") must light up as a
  // hold, not sit quietly in italic gold — FAIL TOWARD SHOWING is the standing
  // rule for every note surface here, and it is what this whole feature exists
  // to enforce.
  keep.push.apply(keep, _buildHoldRules(sheet));
  keep.push(_buildBuyerNoteRule(sheet));
  // ⚠⚠ THE IDENTITY RULES MUST BE REBUILT HERE OR applyBrandTheme() DELETES THEM.
  // They span cols A + D, and the strip above discards ANY single-column rule on col A
  // (it is in themeColumns to clear the retired italic Kit-SKU rule). Without this line
  // a theme re-apply would silently remove the whole feature — and, because the mark is
  // a display layer, leave no trace that it had ever been there.
  keep.push.apply(keep, _buildIdentityRules(sheet));

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

/**
 * ⏸ HOLD — the table's half of the 2026-08-21 fix.
 *
 * The board has carried an amber HOLD chip since 2026-08-15; the SHEET had
 * nothing, so a hold sitting in the NOTE column looked like ordinary text on
 * the surface the whole team actually reads all day. These two rules make the
 * note cell say its own state, with no extra column and no hidden helper sheet:
 *
 *   contains HOLD, no ✓ SEEN  → 🔴 nobody has answered for this
 *   contains HOLD and ✓ SEEN  → 🟡 held, and it names who has it
 *
 * ⭐ SO THE COLOUR OF THE CELL IS THE RECEIPT. Someone reading the table sees
 * red and knows to phone; sees amber and reads the name and time right there in
 * the cell. No Telegram, nothing to hover over, nothing to open.
 *
 * ⚠ ONE CELL DRIVES EVERYTHING. The board reads the same note text through its
 * own noteHasHold(), so the sheet and the tablet cannot disagree — and
 * acknowledging from ANY door repaints both within one poll.
 *
 * ⚠ FAILS TOWARD SHOWING, in both directions. Delete the ack text and the cell
 * goes back to red; delete the word HOLD and both rules stop matching, which is
 * exactly how a hold is lifted. The clearing path is the same act as editing
 * the note you were already reading, so there is no separate mechanism to rot.
 *
 * @param {Sheet} sheet
 * @returns {Array} two rules, unseen first
 */
function _buildHoldRules(sheet) {
  var noteRange = sheet.getRange(
    Schema.dataStartRow, Schema.cols.NOTE,
    BRAND.dataLast - Schema.bannerRows, 1
  );
  var a = '$E' + Schema.dataStartRow;
  // Whole word, anywhere, any case — the SAME rule the board and Holds.js use.
  var hasHold = 'REGEXMATCH(TO_TEXT(' + a + '), "(?i)\\bhold\\b")';
  var hasAck  = 'REGEXMATCH(TO_TEXT(' + a + '), "✓\\s*SEEN")';
  var hasEsc  = 'REGEXMATCH(TO_TEXT(' + a + '), "ESCALATED")';

  /* ⚠ ESCALATED AND STILL UNANSWERED — the loudest state on the sheet, and it
     must come FIRST because Sheets applies the first matching rule per cell.
     Without it, a hold red for forty minutes looks exactly like one red for
     thirty seconds, and "the shipping desk has already been pulled in" is
     precisely the thing the next person needs to know. */
  var escalated = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + a + '<>"", ' + hasHold + ', ' + hasEsc + ', NOT(' + hasAck + '))')
    .setBackground('#b71c1c').setFontColor('#ffffff').setBold(true)
    .setRanges([noteRange]).build();

  // UNSEEN — red, and loud. This is the state that cost a label.
  var unseen = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + a + '<>"", ' + hasHold + ', NOT(' + hasAck + '))')
    .setBackground('#ffcdd2').setFontColor('#b71c1c').setBold(true)
    .setRanges([noteRange]).build();

  // ACKNOWLEDGED — calm, but still marked. Seen is not the same as handled:
  // the box is still sitting there needing its label voided, so the cell must
  // not go back to looking like an ordinary note.
  var seen = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + a + '<>"", ' + hasHold + ', ' + hasAck + ')')
    .setBackground('#fff3c4').setFontColor('#7a5c00')
    .setRanges([noteRange]).build();

  return [escalated, unseen, seen];
}


/**
 * Install the two hold rules on their own, without re-running the whole theme.
 * Mirrors setupBuyerNoteHighlighting — same idempotent strip-then-add shape.
 */
function setupHoldHighlighting() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ No " + MAIN_SHEET_NAME + " sheet.";

  var rules = _stripHoldRules(sheet.getConditionalFormatRules());
  // Ahead of everything else on the NOTE column, for the reason in
  // _applyAllConditionalFormatting.
  rules = _buildHoldRules(sheet).concat(rules);
  sheet.setConditionalFormatRules(rules);
  return "✅ Hold highlighting applied — a NOTE containing HOLD now reads red " +
         "until someone acknowledges it, then amber.";
}

/** Strip just the two hold rules, preserving every other rule on the sheet. */
function _stripHoldRules(rules) {
  return rules.filter(function (rule) {
    var bc = rule.getBooleanCondition();
    if (!bc) return true;
    var formula = (bc.getCriteriaValues() || [''])[0] || '';
    var ranges = rule.getRanges();
    var onNote = ranges.some(function (r) {
      return r.getNumColumns() === 1 && r.getColumn() === Schema.cols.NOTE;
    });
    // All three hold rules carry the same \bhold\b token in their formula, so
    // this catches the escalated one too — and the buyer-note rule, which does
    // not, survives untouched. There is a test for exactly that.
    return !(onNote && /\\bhold\\b/i.test(String(formula)));
  });
}



/**
 * ⚠ IDENTITY — the sheet's half of the 2026-08-28 fix, and the third attempt at it.
 *
 * ⭐⭐ WHY THESE ARE CF RULES AND NOT A SCRIPT PAINTING CELLS. Five rounds on 2026-08-29
 * painted a static background on the offending row, and every single failure came from
 * that one write: the paint could not be told from the #fff8e7 banding; clearing it
 * needed a SECOND write, so it depended on a trigger that Ctrl+Z does not reliably queue
 * (undo pops liveUpdateTrigger's LOCATION/HAND write, not the SKU); and a stale mark had
 * no way to retire itself. The feature ended read-only — a Telegram message and nothing
 * on the sheet.
 *
 * Conditional formatting is a DISPLAY LAYER. Sheets recomputes it live and stores NOTHING
 * on the row, so fixing the cell removes the red in the same keystroke — no trigger, no
 * property, no clearing pass, nothing that can go stale. That is the entire point.
 *
 * THREE STATES, all on the two identity cells (A + D) and nowhere else, so the row keeps
 * its STATUS colour:
 *
 *   GONE        an established row missing its SKU or its SALES ORDER
 *   UNKNOWN     the pair is on the reconcile's published mismatch list
 *   DUPLICATED  the same pair sits on more than one row
 *
 * ⭐ TWO OF THE THREE ARE PURELY LOCAL — GONE and DUPLICATED read nothing but the sheet,
 * so they are instant in BOTH directions and cannot be affected by any timing. Only
 * UNKNOWN needs a lookup, and it deliberately reads a list of ALREADY-JUDGED pairs rather
 * than the Activity Log: doPost inserts rows before writing their RECEIVED events, so a
 * rule that searched the log would flash red on every legitimate arrival.
 *
 * ⚠ EVERY LITERAL IS BUILT FROM A CONSTANT — Schema.validStatuses, Schema.cols,
 * IDENTITY_GUARD.sheetName / .listMax / .deltaNoteToken. Nothing here is typed twice, so
 * there is nothing to drift. A test asserts that.
 *
 * @param {Sheet} sheet
 * @returns {Array} three rules
 */
/**
 * The three formulas, anchored at a row you choose.
 *
 * ⭐ SPLIT OUT 2026-08-30 SO A DIAGNOSTIC CAN EVALUATE THE REAL THING. The rules were
 *   reported as doing nothing twice running, and every round of reasoning about WHY was
 *   wrong. `diagnoseIdentityCF()` now writes these same strings into a scratch cell and
 *   reads back what Sheets actually returns — which is the only way to know. If this were
 *   a second copy of the formula the answer would be worthless, so there is exactly one.
 *
 * ⚠ Absolute COLUMNS, relative ROW ($A4, not A4 or $A$4). A rule with two ranges anchors
 *   relative references at the first range's top-left, then offsets per range; with the
 *   column pinned, only the row moves, which is what makes A and D behave identically.
 *
 * @param {number} anchorRow
 * @returns {{gone: string, unknown: string, duplicated: string, established: string}}
 */
function _identityFormulas(anchorRow) {
  var a = '$A' + anchorRow;   // SKU
  var b = '$B' + anchorRow;   // QTY
  var d = '$D' + anchorRow;   // SALES ORDER
  var f = '$F' + anchorRow;   // STATUS

  // ⚠ THE STATUS LIST IS THE HEADER GUARD, and it earns its keep twice. The DIRECT header
  //   row carries the literal "STATUS" in col F and the boundary divider's F is blank, so
  //   both fall out for free — the same trick _buildStatusRules uses. It also means a row
  //   being typed right now (no status yet) is never judged.
  var established = 'OR(' + Schema.validStatuses.map(function (s) {
    return 'UPPER(TRIM(' + f + '))="' + s + '"';
  }).join(',') + ')';

  var bothPresent = a + '<>"", ' + d + '<>""';

  function listRefFor(colLetter) {
    return 'INDIRECT("\'' + IDENTITY_GUARD.sheetName + '\'!$' + colLetter + '$1:$' +
           colLetter + '$' + IDENTITY_GUARD.listMax + '")';
  }
  var listRef    = listRefFor('A');   // order|sku      — identity
  var qtyListRef = listRefFor('B');   // order|sku|qty  — quantity

  var pairCount =
    'COUNTIFS($A$' + Schema.dataStartRow + ':$A,' + a +
    ',$D$' + Schema.dataStartRow + ':$D,' + d + ')';
  var deltaCount =
    'COUNTIFS($A$' + Schema.dataStartRow + ':$A,' + a +
    ',$D$' + Schema.dataStartRow + ':$D,' + d +
    ',$E$' + Schema.dataStartRow + ':$E,"*' + IDENTITY_GUARD.deltaNoteToken + '*")';

  return {
    established: '=' + established,

    /* GONE — half an identity on a live row. The v2 guard returned "skip" here and was
       therefore blind to the likeliest slip of all: hitting Delete on the wrong cell. */
    gone: '=AND(' + established + ', OR(' + a + '="", ' + d + '=""))',

    /* UNKNOWN — a dumb lookup in the published list, no verdict logic of its own.
       ⚠ IFERROR-wrapped so a missing __Identity sheet FAILS SILENT rather than screaming:
         the MATCH errors, the rule reads FALSE, nothing flags. Same ruling as "an
         unreadable fingerprint means SKIP, not PUBLISH". */
    unknown: '=AND(' + established + ', ' + bothPresent + ', ' +
             'IFERROR(ISNUMBER(MATCH(LOWER(TRIM(' + d + '))&"|"&LOWER(TRIM(' + a + ')), ' +
             listRef + ', 0)), FALSE))',

    /* DUPLICATED — the copied-row case, which GONE and UNKNOWN are both blind to because
       the pair is complete AND was genuinely received, on the source row.
       ⭐ The one legitimate twin identifies itself: Zoho Pull's insert_delta writes a note
         that exists precisely to tell a delta row from a duplicate, so subtracting the
         delta-noted rows means a delta pair reads 2-1=1 and stays quiet.
       ⚠ Open-ended $A$4:$A, never a bounded absolute range — n8n inserts at the top all
         day and a fixed end drifts. */
    duplicated: '=AND(' + established + ', ' + bothPresent + ', ' +
                pairCount + '-' + deltaCount + '>1)',

    /* QTY — the identity is right and the quantity is not.
       ⚠ NOTHING in this codebase writes column B after insert (verified by grep), and Zoho
         Pull's insert_delta creates a NEW row with its own RECEIVED event rather than
         editing an existing qty — so a qty that does not match what was received is always
         a hand-edit. That is what makes this safe to flag at all.
       ⚠ It paints column B alone: the identity is fine, so saying so on A and D would
         point at the wrong cells. */
    qty: '=AND(' + established + ', ' + bothPresent + ', ' + b + '<>"", ' +
         'IFERROR(ISNUMBER(MATCH(LOWER(TRIM(' + d + '))&"|"&LOWER(TRIM(' + a + '))&"|"&TRIM(' + b + '), ' +
         qtyListRef + ', 0)), FALSE))'
  };
}


/**
 * ⚠ IDENTITY — the sheet's half of the 2026-08-28 fix, and the third attempt at it.
 *
 * ⭐⭐ WHY THESE ARE CF RULES AND NOT A SCRIPT PAINTING CELLS. Five rounds on 2026-08-29
 * painted a static background on the offending row, and every single failure came from
 * that one write: the paint could not be told from the #fff8e7 banding; clearing it needed
 * a SECOND write, so it depended on a trigger that Ctrl+Z does not reliably queue; and a
 * stale mark had no way to retire itself.
 *
 * Conditional formatting is a DISPLAY LAYER. Sheets recomputes it live and stores NOTHING
 * on the row, so fixing the cell removes the red in the same keystroke.
 *
 * ⚠⚠ TWO OTHER FUNCTIONS STRIP CF RULES AND BOTH DELETED THESE ON EVERY EDIT until
 * 2026-08-30 — removeLegacySalesOrderCFRules and removeDuplicateHighlightRules, each asking
 * "does ANY range touch MY column?" while these span A AND D. Before changing anything
 * here, grep setConditionalFormatRules across the whole project. Sixteen files write CF.
 *
 * @param {Sheet} sheet
 * @returns {Array} three rules
 */
function _buildIdentityRules(sheet) {
  var rows = BRAND.dataLast - Schema.bannerRows;
  var ranges = [
    sheet.getRange(Schema.dataStartRow, Schema.cols.SKU, rows, 1),
    sheet.getRange(Schema.dataStartRow, Schema.cols.SALES_ORDER, rows, 1)
  ];
  var qtyRange = [sheet.getRange(Schema.dataStartRow, Schema.cols.QTY, rows, 1)];
  var F = _identityFormulas(Schema.dataStartRow);

  function rule(formula, on) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      // The deep red the escalated-hold rule already owns — "act now", and unmistakable
      // against the cream banding that made the 08-29 amber invisible.
      .setBackground('#b71c1c').setFontColor('#ffffff').setBold(true)
      .setRanges(on || ranges).build();
  }

  // ⚠ The qty rule paints column B ALONE. The identity is fine on such a row, so marking
  //   A and D would point the reader at the wrong two cells.
  return [rule(F.gone), rule(F.unknown), rule(F.duplicated), rule(F.qty, qtyRange)];
}


/**
 * Install the three identity rules on their own, without re-running the whole theme.
 * Mirrors setupHoldHighlighting — same idempotent strip-then-add shape.
 */
function setupIdentityHighlighting() {
  // ⚠ ADMIN, NOT WORKFLOW — owner-only, and deliberately NOT bridged. The owner
  //   bridge exists so staff WRITES run as the owner; routing a SETUP function
  //   through it would let anyone re-theme, re-protect or rewrite rules on a locked
  //   sheet. Refusing in a sentence beats an unexplained permission error.
  if (typeof _obRequireOwner === "function") {
    var _denied = _obRequireOwner("Installing the identity rules");
    if (_denied) return _denied;
  }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ No " + MAIN_SHEET_NAME + " sheet.";

  var rules = _stripIdentityRules(sheet.getConditionalFormatRules());
  rules = _buildIdentityRules(sheet).concat(rules);
  sheet.setConditionalFormatRules(rules);

  return "✅ Identity highlighting applied — the SKU and SALES ORDER cells turn red when " +
         "a row is missing half its identity, carries a pair that was never received, or " +
         "duplicates another row. Fix the cell and the red clears itself instantly.";
}

/**
 * Strip just the three identity rules, preserving every other rule on the sheet.
 * ⚠ Identified by RANGE SIGNATURE — exactly two single-column ranges, on SKU and on
 *   SALES ORDER. Nothing else in this codebase spans that pair, and unlike a formula
 *   substring it cannot be broken by rewording the rule.
 */
function _stripIdentityRules(rules) {
  return rules.filter(function (rule) {
    var ranges = rule.getRanges();
    if (!ranges) return true;

    // the three identity rules: exactly two single-column ranges, on SKU and SALES ORDER
    if (ranges.length === 2) {
      var cols = ranges.map(function (r) {
        return (r.getNumColumns() === 1) ? r.getColumn() : -1;
      }).sort(function (x, y) { return x - y; });
      if (cols[0] === Schema.cols.SKU && cols[1] === Schema.cols.SALES_ORDER) return false;
    }

    // the qty rule: column B alone, and it names our helper sheet. ⚠ The sheet-name test
    // matters — column B is otherwise unclaimed, and a future rule there should survive.
    if (ranges.length === 1 && ranges[0].getNumColumns() === 1 &&
        ranges[0].getColumn() === Schema.cols.QTY) {
      var bc = rule.getBooleanCondition();
      var f = bc ? String((bc.getCriteriaValues() || [''])[0] || '') : '';
      if (f.indexOf(IDENTITY_GUARD.sheetName) !== -1) return false;
    }

    return true;
  });
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
  // ⚠ DELETED 2026-08-31. B1 is INSIDE the A1:C1 masthead merge, so this formula has
  //   been invisible since the masthead shipped — it wrote a date nobody could see, and
  //   writing into a merged range's non-anchor cell is a good way to be surprised later.
  //   The date's job moved to the face and to D1. setupBannerDateTime went with it.
  //   (Kept as a no-op rather than deleted outright so applyBrandTheme's call site and
  //   any stale trigger stay harmless.)
  return;
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
/**
 * Minutes -> "8m" / "2h 14m", as a formula FRAGMENT. Shared by the pulse and the
 * headline so the two can never render the same age differently.
 */
function _fmtMinsExpr(ref) {
  // ⚠ ROUND THE TOTAL FIRST, THEN DECOMPOSE. Rounding the remainder independently
  //    produces "5h 60m" (INT(359.6/60)=5, ROUND(MOD(359.6,60))=60) — seen live on the
  //    banner within minutes of install. Rounding first also promotes 59.6 to "1h 0m"
  //    rather than a bare "60m".
  var r = 'ROUND(' + ref + ')';
  return 'IF(' + r + '<60,' + r + '&"m",' +
         'INT(' + r + '/60)&"h "&MOD(' + r + ',60)&"m")';
}

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

  // Row 2: YESTERDAY, same shape, read only by the RESTING face. At 4am "today" is four
  // hours old and empty, so a live curve renders as a blank third of the banner. The
  // resting masthead wears the day it just did instead — the Floor Board's own resting
  // panel idea ("the clock ends the day wearing the day it just did"), finally on the
  // sheet. Same COUNTIFS shape, one day back; costs nothing until off-hours.
  var yesterdayFormulas = [];
  for (var y = 0; y < 24; y++) {
    yesterdayFormulas.push(
      "=IFERROR(COUNTIFS('Activity Log'!A:A,\">=\"&TODAY()-1+" + y + "/24," +
      "'Activity Log'!A:A,\"<\"&TODAY()-1+" + (y + 1) + "/24),0)"
    );
  }
  sheet.getRange(2, 1, 1, 24).setFormulas([yesterdayFormulas]);

  // System Pulse helpers (A3 timestamp, A4 minutes-since)
  sheet.getRange('A3').setFormula("=IFERROR(MAX('Activity Log'!A:A),0)");
  sheet.getRange('A4').setFormula("=IF(A3>0,(NOW()-A3)*1440,-1)");

  // ---- MASTHEAD verdict + its inputs (2026-08-30) ------------------------------------
  // ⭐ ONE verdict, computed once in A6, read by BOTH the face image and the headline.
  //    Nothing downstream may re-derive it — a verdict and a formula that disagree is
  //    the whole drift class the identity guard was bitten by (_igVerdict, 2026-08-30).
  //
  // ⚠ Every read of __Published is IFERROR-wrapped and falls back to 0. That cell can
  //    be empty, stale or trimmed, and a masthead that errors is worse than one saying
  //    "clear". Trim order in Published.js is sidebar -> timeline -> openOrders, so the
  //    cockpit SCALARS read here are effectively never trimmed.
  var PUB = "'__Published'!A1";
  var pubNum = function (key) {
    return '=IFERROR(VALUE(REGEXEXTRACT(' + PUB + ',' + '"""' + key + '"":(\\d+)")),0)';
  };
  // BLANK, NEVER ZERO. pubNum returns 0 through IFERROR when a key is absent, and
  // "0 waiting" on a busy table is a reassuring label on a wrong state -- which this
  // codebase already rules is a bug. The nameplates use this variant so a key that
  // cannot be read drops the suffix entirely rather than lying about the count.
  // A genuine 0 still renders "0 waiting", because an empty table saying so is TRUE.
  var pubNumBlank = function (key) {
    return '=IFERROR(VALUE(REGEXEXTRACT(' + PUB + ',' + '\x22\x22\x22' + key + '\x22\x22:(\\d+)")),"")';
  };
  // Both counts come from getDashboardSnapshot, which the sidebar's queue strip also
  // reads -- so the nameplates and the strip cannot disagree. And the 2026-06-02 fix
  // already un-merged the Zoho mirror out of directPending, so it counts DIRECT-table
  // rows only. No boundary derivation anywhere.
  sheet.getRange('A17').setFormula(pubNumBlank('ebayPending'));
  sheet.getRange('A18').setFormula(pubNumBlank('directPending'));

  sheet.getRange('A7').setFormula(pubNum('oldestPendingMinutes'));
  sheet.getRange('A9').setFormula(pubNum('receivedToday'));
  sheet.getRange('A10').setFormula(pubNum('shippedToday'));

  // The queue is read straight off the sheet — local, instant, and it keeps working
  // when the publish cycle does not.
  var qCol = "'" + MAIN_SHEET_NAME + "'!F" + Schema.dataStartRow + ":F";
  sheet.getRange('A8').setFormula(
    '=IFERROR(COUNTIF(' + qCol + ',"PENDING")+COUNTIF(' + qCol + ',"PREPARING"),0)'
  );

  sheet.getRange('A11').setFormula('=' + _fmtMinsExpr('A7'));
  sheet.getRange('A12').setFormula('=' + _fmtMinsExpr('A4'));

  // ---- off-hours ---------------------------------------------------------------------
  // ⭐⚠ THE SHEET NEVER KNEW ABOUT OFF-HOURS, AND THAT MADE THE MASTHEAD LIE. The Floor
  //     Board has isOffHours() and the sidebar has _isOffHoursHouston(); row 1 had
  //     neither, so a quiet Saturday night read as a dead pipeline in the loudest place
  //     on the screen. A false alarm on the loudest surface is how alarms stop being
  //     believed — this project's own ruling, applied to itself.
  //
  // Matches the BOARD's definition, which is the fuller of the two: weekends off all
  // day, otherwise 9-17. (⚠ The sidebar's copy omits weekends — a real drift between
  //  those two that predates this and is not fixed here.)
  // The spreadsheet timezone is America/Chicago, so NOW() is already Houston time.
  var base = 'TODAY()+IF(HOUR(NOW())<9,0,1)';
  sheet.getRange('A13').setFormula(
    '=OR(WEEKDAY(NOW())=1,WEEKDAY(NOW())=7,HOUR(NOW())<9,HOUR(NOW())>=17)'
  );
  // Minutes until the next working 9am. ⚠ No LET() — it threw "Formula parse error" on
  // this sheet (2026-06-05), so the base expression is repeated rather than bound.
  sheet.getRange('A14').setFormula(
    '=MAX(0,((' + base +
      '+IF(WEEKDAY(' + base + ')=7,2,IF(WEEKDAY(' + base + ')=1,1,0))' +
      '+9/24)-NOW())*1440)'
  );
  sheet.getRange('A15').setFormula('=' + _fmtMinsExpr('A14'));

  // ⚠ A4 is -1 when the Activity Log is unreadable. That is a DEAD pipeline, not a
  //    healthy one — it must land on "stale", never fall through to "clear".
  //
  // ⚠ REST OUTRANKS EVERYTHING, including "late". The asymmetry decides it: a false
  //    alarm at 3am costs the credibility of the one surface that has to stay believed,
  //    while a missed one costs nothing because nobody can act until 9. The facts are
  //    not lost — the headline still carries what is waiting.
  sheet.getRange('A6').setFormula(
    '=IF(A13,"rest",' +
     'IF(A7>' + MASTHEAD.lateMinutes + ',"late",' +
     'IF(OR(A4<0,A4>' + MASTHEAD.staleMinutes + '),"stale",' +
     'IF(A8>0,"busy","clear"))))'
  );

  return sheet;
}

/**
 * Writes the four masthead formulas into banner row 1 (2026-08-30).
 *
 *   A1:C1  the state face  — one =IMAGE(), chosen by the verdict in __SparkData!A6
 *   D1     the headline    — branches on the SAME verdict, so the two cannot disagree
 *   E1     the pulse       — UNMOVED, still carrying "h:mm AM/PM"
 *   F1:H1  the day         — SPARKLINE over __SparkData!A1:X1
 *
 * ⚠⚠ E1 MUST KEEP a "h:mm AM/PM" substring. ActivityLog.js regex-parses it into
 *    cockpit.lastSyncMinutes, which drives the Floor Board heartbeat, the sidebar
 *    System Pulse, /status and the published tick. There is a Node assertion on this.
 *
 * ⚠ E1 KEEPS ITS EMOJI LAMP, deliberately. A cell has ONE font colour, so 🟢/🟡/🔴 is
 *   the only way it can carry the tier's own colour — the same reason the stats bullets
 *   survived the 2026-08-30 emoji audit. A CF rule would also be at the mercy of
 *   _applyAllConditionalFormatting's strip-and-rebuild, which is exactly what silently
 *   deleted the identity rules for an hour.
 *
 * IMPORTANT: this OVERWRITES A1, D1, E1 and F1. Intentional — updateLastSyncTimestamp()
 * and updateOrderStatsInSheet() became no-ops on 2026-05-17 precisely so they cannot
 * clobber these formulas after every n8n sync.
 */
function _setSystemPulseBannerFormulas(sheet) {
  var SD = "'__SparkData'!";

  // ---- A1:C1 — THE FACE --------------------------------------------------------------
  // ⚠ IFERROR is not optional. If the endpoint dies at 3am the header degrades to the
  //   text "HQ" chip — never a broken-image icon, never a dead header.
  sheet.getRange(Schema.cellMasthead).setFormula(
    '=IFERROR(IMAGE("' + MASTHEAD.baseUrl + '"&' + SD + 'A6&"-h"&TEXT(HOUR(NOW()),"00")&"-' +
    MASTHEAD.version + '.' + MASTHEAD.ext + '",4,' + MASTHEAD.imgH + ',' + MASTHEAD.imgW + '),"HQ")'
  );

  // ---- D1 — THE HEADLINE -------------------------------------------------------------
  // Branches on the same A6 verdict the face uses, so the picture and the words always
  // agree by construction. Two lines via CHAR(10); the cell is set to WRAP.
  // ⭐ D1 CARRIES NUMBERS. THE FACE CARRIES THE STATE. Until 2026-08-31 both narrated
  //    the same thing in two typefaces 287px apart — the face said "ALL CAUGHT UP /
  //    nothing waiting" while D1 said "N in · N out / nothing waiting", the SAME WORDS
  //    twice on the clear verdict. That is the "two objects" feeling: a redundancy
  //    problem, not a lighting one. Splitting the voices is what makes them one object.
  //
  // ⚠ Lowercase, and neutral ink. The face does the shouting — uppercase, tracked, and
  //   the only coloured thing in row 1. If D1 shouts too, nothing leads.
  //
  // ⚠ NO LET(). It throws "Formula parse error" on this sheet (2026-06-05), so the
  //   __SparkData refs are repeated rather than bound.
  sheet.getRange(Schema.cellStats).setFormula(
    '=IF(' + SD + 'A6="rest","opens in "&' + SD + 'A15&' +
      'IF(' + SD + 'A8>0," · "&' + SD + 'A8&" waiting",""),' +
    'IF(' + SD + 'A6="late","oldest "&' + SD + 'A11&CHAR(10)&' + SD + 'A8&" still waiting",' +
    'IF(' + SD + 'A6="stale","last seen "&' + SD + 'A12,' +
    'IF(' + SD + 'A6="busy",' + SD + 'A8&" to grab"&CHAR(10)&' + SD + 'A9&" in · "&' +
    SD + 'A10&" out",' +
    SD + 'A9&" in · "&' + SD + 'A10&" out"))))'
  );

  // ---- E1 — THE PULSE (cell deliberately UNMOVED) ------------------------------------
  // ⚠ THE PULSE MUST AGREE WITH THE FACE. Shipped 2026-08-30 saying "🔴 STALE" while the
  //    face beside it said CLOSED — the banner contradicting itself in the same row.
  //    Both now read the SAME A13, so off-hours is answered once and rendered twice.
  //    ⚪ joins the 🟢🟡🔴 traffic light as the "not applicable" tier; it still carries
  //    its own colour, which is the whole reason these are emoji and not glyphs.
  sheet.getRange(Schema.cellSyncTime).setFormula(
    '=IF(' + SD + 'A4<0,"⏱ OFFLINE · no activity logged",' +
    '"⏱ "&IF(' + SD + 'A13,"⚪ RESTING",' +
    'IF(' + SD + 'A4<15,"🟢 ALIVE",IF(' + SD + 'A4<60,"🟡 IDLE","🔴 STALE")))&' +
    '" · "&TEXT(' + SD + 'A3,"h:mm AM/PM")&" · "&' + SD + 'A12&" ago")'
  );

  // ---- A2:E2 — THE SKY ---------------------------------------------------------------
  // Same hour, same light curve as the face, so the two halves of the banner can never
  // disagree about what time it is.
  if (MASTHEAD.sky) {
    sheet.getRange(MASTHEAD.skyCell).setFormula(
      '=IFERROR(IMAGE("' + MASTHEAD.baseUrl + 'sky-h"&TEXT(HOUR(NOW()),"00")&"-' +
      MASTHEAD.version + '.' + MASTHEAD.ext + '",4,' + MASTHEAD.skyH + ',' +
      MASTHEAD.skyW + '),"")'
    );
  }

  // ---- F1:H1 — THE DAY CURVE ---------------------------------------------------------
  // ⚠ The curve is NOT lit by the hour, deliberately. It is the one element in the banner
  //   carrying DATA, and brand yellow means "action" everywhere else in this system.
  //   Colouring it by time of day would spend a meaning-bearing colour on decoration. It
  //   has exactly two states — working (yellow) and resting (cool) — and that is enough.
  // ⭐ __SparkData!A1:X1 has held 24 live hourly counts since May, built for a heatmap
  //    that was pivoted away from — nothing has read them since. Future hours are
  //    genuinely 0, so the curve GROWS RIGHTWARD through the day on its own; no dimming
  //    trick is needed.
  // ⭐ SPARKLINE also sidesteps the block-character approach's divide-by-MAX, so a quiet
  //    morning cannot produce a #DIV/0.
  // ⚠ Off-hours it shows YESTERDAY (row 2) in the rest tone, not today's empty row —
  //    and the whole banner goes cool together rather than the face resting beside a
  //    live-yellow chart.
  var spark = function (range, colour) {
    return 'SPARKLINE(' + SD + range + ',{"charttype","column";"color","' + colour +
           '";"empty","zero"})';
  };
  sheet.getRange(Schema.cellDayCurve).setFormula(
    '=IFERROR(IF(' + SD + 'A13,' + spark('A2:X2', MASTHEAD.restAccent) + ',' +
    spark('A1:X1', BRAND.yellow) + '),"")'
  );
}

/**
 * setupMasthead — the one-shot installer for row 1. Editor-run, zero args.
 *
 * ⚠ Row 1's merges were made BY HAND in the Sheets UI during the 2026-05-19 compaction,
 *   not by code — applyBrandTheme has only ever merged row 2. So this breaks whatever
 *   shape is there before building ours, one try/catch per range so a range that is not
 *   currently merged cannot abort the install.
 *
 * Deliberately NARROW. A full applyBrandTheme() repaints the whole sheet and would need
 * its own verification pass; this touches row 1, __SparkData, and step 6's validation
 * sweep.
 *
 * ⚠⚠ STEP 6 IS THE DANGEROUS PART — read its comment before editing it. It used to
 *    hardcode a wipe of I2:J2, which the 2026-08-31 migration turns into the LIVE pick-ID
 *    cells. It now protects the live addresses and sweeps the rest.
 */
/**
 * _buildRow2 — row 2 is the eBay table's NAME, and this owns it end to end.
 *
 * ⚠⚠ IT REPORTS EVERY STEP, AND THAT IS THE POINT. On 2026-08-31 a run of setupMasthead
 *    left row 2 HALF-APPLIED: the row height moved to the nameplate's 44px while the
 *    merges and the logo stayed on the old layout. The result looked worse than either
 *    design, and working out which half had landed took reading the sheet back through a
 *    separate diagnostic. A build step that cannot say what it did turns every future
 *    failure into archaeology — so each step here is isolated, and the caller prints the
 *    outcome of all of them.
 *
 * ⚠ IDEMPOTENT BY CONSTRUCTION. Every merge is broken before it is made, every colour is
 *   set explicitly rather than inherited, and the VACATED picker cells are actively
 *   repainted — they keep their dark badge styling otherwise and read as two empty black
 *   boxes in the middle of the row, which is exactly what they did after the migration.
 *
 * @param {Sheet}   sheet
 * @param {boolean} plate  true once the pickers have moved off the banner
 * @returns {string[]} one line per step
 */
function _buildRow2(sheet, plate) {
  var log = [];
  var step = function (name, fn) {
    try { fn(); log.push('✓ ' + name); }
    catch (e) { log.push('✗ ' + name + ' — ' + e); }
  };

  // Break EVERY shape row 2 has ever worn, whichever direction we are going.
  step('clear merges', function () {
    ['A2:D2', 'A2:E2', 'A2:F2', 'E2:F2', 'F2:G2', 'G2:H2', 'I2:J2'].forEach(function (r) {
      try { sheet.getRange(r).breakApart(); } catch (e) { /* not merged — fine */ }
    });
  });

  if (plate) {
    step('merge A2:F2 + G2:H2', function () {
      sheet.getRange('A2:F2').merge();
      sheet.getRange('G2:H2').merge();
    });

    // ⚠ A:H only. I2/J2 hold the pickers and are hidden — leave their badge styling be.
    step('plate A2:H2', function () {
      sheet.getRange(2, 1, 1, 8)
        .setBackground(MASTHEAD.row2Plate)
        .setFontColor(MASTHEAD.nameplateInk)
        .setFontFamily(BRAND.fontDisplay)
        .setVerticalAlignment('middle')
        .setBorder(false, false, false, false, false, false);
    });

    step('nameplate G2:H2', function () {
      sheet.getRange('G2:H2')
        .setFormula(_nameplateFormula(MASTHEAD.nameEbay, 'A17'))
        .setFontColor(MASTHEAD.nameplateInk)
        .setFontFamily(BRAND.fontDisplay).setFontWeight('bold').setFontSize(10)
        .setHorizontalAlignment('right').setVerticalAlignment('middle').setWrap(false);
    });

    step('eBay logo (plate asset)', function () { setupEbayLogo(true); });
    step('row height ' + MASTHEAD.row2Height, function () {
      sheet.setRowHeight(2, MASTHEAD.row2Height);
    });

  } else {
    // The pre-2026-08-31 layout, restored exactly: cream logo zone, pickers on the banner.
    //
    // ⚠⚠ THE PLATE'S GROUND HAS TO BE SCRUBBED FIRST. The first cut of this branch only
    //    painted A2:E2 cream and left F2:H2 wearing the charcoal — so a revert produced a
    //    row that was half cream and half plate, which reads as a design choice rather
    //    than the residue it is. A revert has to undo the WHOLE of what it reverses, not
    //    the part that is easy to name.
    step('scrub the plate off A2:H2', function () {
      sheet.getRange(2, 1, 1, 8)
        .setBackground(BRAND.paperWarm)
        .setFontColor(BRAND.ink)
        .setBorder(false, false, false, false, false, false);
      sheet.getRange('G2:H2').clearContent();
    });

    // ⚠ BOTH merges, not just the logo's. The Shipping picker lives in a MERGED F2:G2 —
    //   the plate broke it to build A2:F2 + G2:H2, and re-merging A2:E2 alone would send
    //   the dropdown home to a cell half the width it expects.
    step('merge A2:E2 + F2:G2', function () {
      sheet.getRange('A2:E2').merge();
      sheet.getRange('F2:G2').merge();
    });

    // Reuse the styler rather than re-deriving the badge look here — it branches on the
    // resolver, so after a rollback it paints the pickers exactly as they were.
    step('logo zone + Pick ID badges', function () { _styleBannerRow2(sheet); });
    step('eBay logo (transparent asset)', function () { setupEbayLogo(false); });
    step('row height 65', function () { sheet.setRowHeight(2, 65); });
  }

  // ⚠⚠ READ IT BACK. Every assertion above is about what we ASKED for; this is the only
  //    part that says what the sheet actually holds. It is what would have caught the
  //    half-applied run immediately instead of hours later.
  SpreadsheetApp.flush();
  try {
    var mg = [];
    sheet.getRange(2, 1, 1, Schema.dataWidth).getMergedRanges()
         .forEach(function (r) { mg.push(r.getA1Notation()); });
    var wantMerge = plate ? 'A2:F2' : 'A2:E2';
    var a2 = String(sheet.getRange('A2').getFormula() || '');
    var h  = sheet.getRowHeight(2);

    log.push('— read-back —');
    var wantAlso = plate ? 'G2:H2' : 'F2:G2';
    log.push((mg.indexOf(wantMerge) !== -1 ? '✓' : '✗') + ' merges: ' +
             (mg.length ? mg.join(' ') : 'none') + '   (want ' + wantMerge + ')');
    log.push((mg.indexOf(wantAlso) !== -1 ? '✓' : '✗') + ' second merge ' + wantAlso +
             (plate ? '   (the nameplate)' : '   (the Shipping picker)'));
    log.push((a2.indexOf(plate ? 'ebay-v2' : 'ebay-v1') !== -1 ? '✓' : '✗') +
             ' logo: ' + (a2.match(/ebay-v\d/) || ['none'])[0] +
             '   (want ' + (plate ? 'ebay-v2' : 'ebay-v1') + ')');
    log.push((h === (plate ? MASTHEAD.row2Height : 65) ? '✓' : '✗') + ' height: ' + h + 'px');
  } catch (e) {
    log.push('✗ read-back threw — ' + e);
  }
  return log;
}

/**
 * _applyDividerNameplate — the DIRECT band's right-hand count.
 *
 * ⚠⚠ THE MISSING CALL SITE. _styleDirectDivider is only ever reached through
 *    applyBrandTheme, which repaints the ENTIRE sheet and is not something anyone runs
 *    casually. So the divider's nameplate formula — written on 2026-08-31 — had never
 *    once been applied, and the band was still showing a hand-edited static string that
 *    happened to match. Two nameplates were designed to rhyme and only one existed.
 *    setupMasthead owns the banner's grammar, so it owns this too.
 *
 * ⚠ The formula goes on the ANCHOR CELL, not the span. setFormula on a multi-cell range
 *   writes it into every cell, and the right merge is not created by code — it was made
 *   by hand — so the span may or may not actually be merged.
 */
function _applyDividerNameplate(sheet) {
  var boundary = _findBoundaryInSheet(sheet);
  if (boundary <= 0) return '✗ divider: boundary row not found';
  try {
    var col = Schema.boundaryLeftWidth + 1;                       // G
    sheet.getRange(boundary, col).setFormula(_nameplateFormula(MASTHEAD.nameDirect, 'A18'));
    sheet.getRange(boundary, col, 1, Schema.boundaryRightWidth)
      .setFontFamily(BRAND.fontDisplay).setFontWeight('bold').setFontSize(10)
      .setFontColor(BRAND.ink)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    SpreadsheetApp.flush();
    var got = String(sheet.getRange(boundary, col).getDisplayValue() || '');
    return (got.indexOf(MASTHEAD.namePrefix) !== -1 ? '✓' : '✗') +
           ' divider nameplate row ' + boundary + ': "' + got + '"';
  } catch (e) {
    return '✗ divider nameplate — ' + e;
  }
}

/**
 * revertRow2 — put row 2 back exactly as it was before 2026-08-31.
 *
 * Cream A2:E2 logo zone, the transparent logo asset at its old size, 65px.
 * Editor-run, zero args, owner-gated. This is the promised one-call escape hatch: the
 * nameplate is a taste decision, and a taste decision that cannot be undone in one
 * step is not a decision, it is a commitment.
 *
 * ⚠ It does NOT move the pickers back — that is rollbackPickIdCells, and the two are
 *   deliberately separate. Row 2's appearance and where the dropdowns live are
 *   different questions, and conflating them would make a cosmetic revert into a
 *   data migration.
 */
function revertRow2() {
  if (typeof _obRequireOwner === "function") {
    var denied = _obRequireOwner("Reverting row 2");
    if (denied) return denied;
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return "❌ Main sheet not found.";
  var out = _buildRow2(sheet, false);
  var rep = "↩ Row 2 reverted to the pre-nameplate layout\n  " + out.join('\n  ') +
            (MASTHEAD.row2Style === 'plate'
              ? "\n\n  ⚠⚠ MASTHEAD.row2Style is still 'plate', so the NEXT setupMasthead()" +
                "\n     will build the nameplate again and undo this. Set it to 'cream'" +
                "\n     in BrandTheme.js and push — a revert the installer overwrites is" +
                "\n     not a revert."
              : "") +
            "\n\n  The pickers were NOT moved — run rollbackPickIdCells() for that," +
            "\n  and delete the PICK_ID_ADDR property first if you do.";
  console.log(rep);
  return rep;
}

function setupMasthead(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName || MAIN_SHEET_NAME);
  if (!sheet) return "❌ Sheet not found: " + (sheetName || MAIN_SHEET_NAME);

  // 1 — clear every row-1 merge shape this layout has ever worn, then build ours.
  ['A1:C1', 'A1:D1', 'A1:H1', 'B1:D1', 'B1:E1', 'D1:E1', 'F1:H1', 'G1:J1']
    .forEach(function (r) {
      try { sheet.getRange(r).breakApart(); } catch (e) { /* not merged — fine */ }
    });
  sheet.getRange('A1:C1').merge();
  sheet.getRange('F1:H1').merge();

  // 2 — the ground. Row 1 is ink end to end; each face's own art carries the colour.
  //     ⚠ It stops at dataWidth. A version of this ran to column T so the band would
  //       "bleed" off the right edge; on screen that read as nothing but colour poured
  //       into empty columns, and it was reverted the same day.
  sheet.getRange(1, 1, 1, Schema.dataWidth)
    .setBackground(BRAND.ink)
    .setFontColor('#ffffff')
    .setFontFamily(BRAND.fontDisplay)
    .setVerticalAlignment('middle');

  // 3 — the zones. WRAP on D1 so CHAR(10) renders as two lines.
  //     ⚠ setRowHeight comes AFTER, so the explicit height wins over auto-fit — the
  //       2026-06-05 "an embedded newline expands the row" lesson.
  sheet.getRange(Schema.cellMasthead)
    .setFontColor(BRAND.yellow).setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('left');                      // styles the "HQ" fallback only
  // ⚠ ONE ACCENT AT A TIME. Row 1 used to paint three separate things brand yellow —
  //   the face, the headline and the pulse — so nothing led. The face's state word is
  //   the accent now; D1 and E1 go neutral, same size, same weight, same alignment, so
  //   they read as one continuous line of information rather than two competing labels.
  // ⚠ D1 WRAPS (it uses CHAR(10)); E1 must NOT (its string is parsed by ActivityLog).
  [Schema.cellStats, Schema.cellSyncTime].forEach(function (a1) {
    sheet.getRange(a1)
      .setFontColor(MASTHEAD.quietInk).setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  sheet.getRange(Schema.cellStats).setWrap(true);
  sheet.getRange(Schema.cellSyncTime).setWrap(false);

  // 4 — ROW 2, BUILT FOR WHERE THE PICKERS ACTUALLY ARE.
  //
  // ⚠⚠ THIS BRANCH IS WHAT MAKES THE WHOLE STEP DORMANT. The nameplate's home is G2:H2
  //    — and until migratePickIdCells runs, G2 is the second half of the F2:G2 Shipping
  //    merge. Building the nameplate unconditionally would tear a live dropdown in half.
  //    So the layout follows Schema.pickIdA1(): with PICK_ID_ADDR unset this rebuilds
  //    row 2 EXACTLY as it is today, and the nameplate appears only when the property
  //    flips. Shipping this changes nothing on the sheet, which is the point.
  // ⚠ TWO CONDITIONS, AND THEY MEAN DIFFERENT THINGS. row2Style is the look you want;
  //   pickersMoved is whether the space is physically free (G2 is half of the F2:G2
  //   Shipping merge until the dropdowns move). Wanting the plate is not enough.
  var pickersMoved = (Schema.pickIdA1() === Schema.cellEmployeeIdNext);
  var wantPlate    = (MASTHEAD.row2Style === 'plate') && pickersMoved && !MASTHEAD.sky;
  var row2Log = _buildRow2(sheet, wantPlate);
  if (MASTHEAD.row2Style === 'plate' && !pickersMoved) {
    row2Log.push('⚠ row2Style is "plate" but the pickers are still on the banner — ' +
                 'built the cream layout instead. Run migratePickIdCells first.');
  }

  // 5 — the inputs, then the formulas that read them.
  _ensureSparkData(ss);
  _setSystemPulseBannerFormulas(sheet);

  sheet.setRowHeight(1, MASTHEAD.rowHeight);

  // ⚠⚠ ASSERT THE FACE'S CANVAS. =IMAGE() mode 4 takes explicit pixel dimensions, so if
  //    A1:C1 does not sum to imgW the art STRETCHES OR CLIPS — silently, and it still
  //    looks like a masthead. applyBrandTheme sets explicit widths on A-J, so a theme
  //    re-apply is exactly what would move this out from under the image.
  var span = 0;
  for (var wc = 1; wc <= 3; wc++) span += sheet.getColumnWidth(wc);
  var widthNote = (span === MASTHEAD.imgW) ? '' :
    ' · ⚠ A1:C1 is ' + span + 'px but MASTHEAD.imgW is ' + MASTHEAD.imgW +
    ' — the face will stretch. Reconcile applyBrandTheme\'s column widths.';

  // 6 — ⚠⚠ SWEEP STALE ROW-2 VALIDATIONS — AND NEVER THE LIVE PICK IDs.
  //
  //     The 2026-05-19 compaction left phantom dropdowns at I2:J2 when it moved the
  //     Adjustment picker to H2. They are invisible while I+J are hidden and become
  //     two extra pickers the moment anyone unhides those columns. Worth sweeping.
  //
  // ⚠⚠ THIS BLOCK USED TO HARDCODE `I2:J2`, AND THAT WAS A LOADED GUN. The 2026-08-31
  //    migration moves the LIVE pickers to exactly those two cells, so the next run of
  //    this installer would have destroyed both dropdowns — and the loss is
  //    UNRECOVERABLE, because NOTHING IN THIS CODEBASE AUTHORS THAT OPTION LIST.
  //    _rewritePickIdValidation only rewrites a rule that already exists; the lists were
  //    made by hand in the Sheets UI, so the cell is the only copy of its own options.
  //    The whole block sits in a try/catch and returned a success-shaped string, so it
  //    would have read as a clean install for weeks.
  //
  //    The rule is now INVERTED: protect the live addresses, sweep whatever else in
  //    row 2 carries a validation. Protecting rather than targeting means a future
  //    layout change moves the pickers without ever re-arming this.
  var phantoms = [];
  try {
    // ⚠ Protect the WHOLE MERGE each address sits in, not just the anchor. A rule
    //   applied to F2:G2 lives on BOTH cells, so sweeping the non-anchor half would be
    //   an invisible partial wipe. Same "merges[0] or the cell" shape the lock carve-out
    //   uses at :619. The *Next constants may not exist yet — the guard tolerates that,
    //   which is what keeps this correct DURING the migration, when both the old and the
    //   new addresses are briefly live at once.
    var keepCol = {};
    [Schema.cellEmployeeId,     Schema.cellAdjustmentId,
     Schema.cellEmployeeIdNext, Schema.cellAdjustmentIdNext].forEach(function (a1) {
      if (!a1) return;
      var r      = sheet.getRange(a1);
      var merges = r.getMergedRanges();
      var span   = merges.length ? merges[0] : r;
      for (var i = 0; i < span.getNumColumns(); i++) keepCol[span.getColumn() + i] = true;
    });

    // One batched read of the row rather than a getDataValidation() per cell.
    var rules = sheet.getRange(2, 1, 1, Schema.dataWidth).getDataValidations()[0];
    for (var c = 0; c < rules.length; c++) {
      if (!rules[c] || keepCol[c + 1]) continue;
      var stale = sheet.getRange(2, c + 1);
      stale.setDataValidation(null);
      phantoms.push(stale.getA1Notation());
    }
  } catch (e) {
    // Best effort — never block the masthead on a cleanup. But SAY SO: swallowing
    // this silently is half of what made the old version dangerous.
    console.log('setupMasthead.sweepRow2Validations: ' + e);
  }

  SpreadsheetApp.flush();
  // ⚠ The divider is part of the banner's grammar, so the installer owns it. It used to
  //   be reachable only through applyBrandTheme, which is why its nameplate had never
  //   actually been applied.
  var dividerLine = _applyDividerNameplate(sheet);

  var rep = "✅ Masthead installed · row 1 = " + MASTHEAD.rowHeight + "px · faces " +
            MASTHEAD.version + widthNote +
            (phantoms.length ? "\n   swept stale row-2 validation(s): " + phantoms.join(', ') : "") +
            "\n\n  ROW 2 (style=" + MASTHEAD.row2Style + " · pickers at " +
              Schema.pickIdA1() + "/" + Schema.pickIdA1('adjustment') + ")\n    " +
            row2Log.join("\n    ") +
            "\n\n  DIVIDER\n    " + dividerLine +
            (_lockNeedsRefresh(sheet)
              ? "\n\n  ⚠⚠ RE-RUN protectAllOrdersSheet() — row 2's merges just changed and the" +
                "\n     lock's carve-out still names the OLD shape. A partial carve-out over a" +
                "\n     merge locks the whole merge, so the Pick ID would refuse STAFF while" +
                "\n     working fine for you (removeEditors ignores the owner)."
              : "") +
            "\n\n  ↩ revertRow2() puts row 2 back the way it was, in one call.";
  console.log(rep);
  return rep;
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

/**
 * ⏭ PROBE — does an animated GIF actually MOVE when floated OVER cells?
 *
 * ⚠⚠ THIS IS A DIFFERENT MECHANISM FROM =IMAGE(), AND THAT DISTINCTION IS THE WHOLE
 *    POINT. =IMAGE() renders INSIDE a cell and is proven static (tested 2026-08-30 with
 *    a known-good animated GIF — it sat motionless). insertImage() creates a floating
 *    OverGridImage that ignores the grid entirely. They are separate code paths and the
 *    first proves nothing about the second, which is what I wrongly assumed.
 *
 * Deliberately the SAME GIF used for the =IMAGE() test, so this is a clean A/B on the
 * mechanism rather than on the file.
 *
 * ⚠ Safe by construction: an OverGridImage is a floating object. It writes to NO cell,
 *   changes no value, no format, no validation. removeProbeAnimatedGif() deletes it.
 *   Anchored at column 12 (L) — past dataWidth, so it floats over empty space.
 */
var GIF_PROBE = {
  url: "https://upload.wikimedia.org/wikipedia/commons/2/2c/Rotating_earth_%28large%29.gif"
};

// ⚠ The anchor is RESOLVED, not hardcoded. The first cut anchored at column 12 and threw
//   "Those columns are out of bounds" — All Orders is exactly Schema.dataWidth (10)
//   columns wide, so there is no column L to float over. Park it below the data instead,
//   clamped to whatever the sheet actually is.
function _gifProbeAnchor(sheet) {
  var row = Math.min(sheet.getMaxRows(), Math.max(1, sheet.getLastRow() + 3));
  return { row: row, column: 2 };               // B, a few rows under the last data row
}

function probeAnimatedGif() {
  var sheet = SpreadsheetApp.getActiveSheet();
  removeProbeAnimatedGif();                     // idempotent — never stack probes
  var at = _gifProbeAnchor(sheet);
  sheet.insertImage(GIF_PROBE.url, at.column, at.row);
  SpreadsheetApp.flush();
  return "✅ Floated an animated GIF over " +
         sheet.getRange(at.row, at.column).getA1Notation() +
         "  (sheet is " + sheet.getMaxRows() + " rows x " + sheet.getMaxColumns() + " cols)\n" +
         "   Scroll DOWN — it is parked just below the last data row.\n\n" +
         "WATCH IT for ~5 seconds:\n" +
         "  · the globe SPINS  → insertImage animates, and the masthead can move\n" +
         "  · it sits still    → the platform genuinely cannot, on either path\n\n" +
         "Then run removeProbeAnimatedGif() to clear it.";
}

function removeProbeAnimatedGif() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var gone = 0;
  // ⚠ Match on the ANCHOR, never remove all images — setupBrandLogo() also uses
  //   insertImage, and a blanket getImages().remove() would take the HQ mark with it.
  sheet.getImages().forEach(function (img) {
    try {
      var a = img.getAnchorCell();
      if (a.getColumn() === 2 && a.getRow() > 3) {   // below the banner, in col B
        img.remove(); gone++;
      }
    } catch (e) { /* an image with no readable anchor is not ours */ }
  });
  SpreadsheetApp.flush();
  return gone ? "✅ Removed " + gone + " probe image(s)." : "Nothing to remove.";
}

/**
 * ⏭ PROBE 2 — float the REAL animated face over the masthead, and find out whether it
 *             stays PINNED when you scroll.
 *
 * ⭐ Probe 1 settled the big question: insertImage() DOES animate a GIF (the globe spun),
 *    even though =IMAGE() renders a dead first frame. Two different code paths, and the
 *    first proved nothing about the second — which is exactly what I got wrong.
 *
 * ⚠⚠ THIS IS THE REMAINING UNKNOWN, AND IT DECIDES THE DESIGN. An OverGridImage floats on
 *    the GRID, not on the viewport. Row 1 is inside the FROZEN pane (frozen rows = 3). If
 *    a floating image scrolls away with the content, an animated masthead is useless the
 *    moment anyone scrolls — which is always. Nothing in the docs settles it; scrolling
 *    does.
 *
 * ⚠ Non-destructive: an OverGridImage writes to no cell. A1's =IMAGE() formula stays
 *   underneath as the fallback and is untouched.
 */
var MAST_ANIM = { state: 'clear', width: 280, height: 44 };

function probeMastheadAnimated(state) {
  var sheet = SpreadsheetApp.getActiveSheet();
  removeMastheadAnimated();
  var url = MASTHEAD.baseUrl + (state || MAST_ANIM.state) + '-' + MASTHEAD.version + '.gif';
  var img = sheet.insertImage(url, 1, 1);       // anchored at A1 — inside the frozen pane
  img.setWidth(MAST_ANIM.width).setHeight(MAST_ANIM.height);
  SpreadsheetApp.flush();
  return "✅ Floated the ANIMATED face over A1.\n" +
         "   " + url + "\n\n" +
         "TWO THINGS TO CHECK:\n" +
         "  1 · does it move?      (the shine should sweep across every ~3s)\n" +
         "  2 · SCROLL DOWN a page — does it STAY on the frozen banner, or scroll off?\n\n" +
         "      stays  → an animated masthead is real, and this becomes the mechanism\n" +
         "      scrolls off → floating images cannot carry the banner; motion has to\n" +
         "                    live somewhere that is not row 1\n\n" +
         "Then run removeMastheadAnimated().";
}

function removeMastheadAnimated() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var gone = 0;
  // ⚠ Match the ANCHOR, never sweep getImages() — setupBrandLogo() also floats an image
  //   and a blanket remove would take the HQ mark with it.
  sheet.getImages().forEach(function (img) {
    try {
      var a = img.getAnchorCell();
      if (a.getRow() === 1 && a.getColumn() === 1) { img.remove(); gone++; }
    } catch (e) { /* no readable anchor — not ours */ }
  });
  SpreadsheetApp.flush();
  return gone ? "✅ Removed " + gone + " floating masthead image(s)." : "Nothing to remove.";
}

/**
 * diagnoseMasthead — why is the banner showing the "HQ" text chip instead of the face?
 *
 * ⚠ A MISSING FACE ANSWERS 200 WITH THE BOARD'S HTML, not 404 (Caddy try_files). IMAGE()
 *   cannot decode that, so it errors and IFERROR falls back — silently, behind a success
 *   status. Same "success status, failure content" shape as Zoho's 200-with-an-error-body.
 *   This builds the URL exactly as the formula does and FETCHES it, so the answer is
 *   measured rather than reasoned about.
 *
 * ⚠ Zero-arg on purpose: the editor Run button cannot pass arguments (walked into three
 *   times in this project). Output goes to the EXECUTION LOG.
 */
function diagnoseMasthead() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var sd    = ss.getSheetByName('__SparkData');
  var L = [];
  var say = function (s) { L.push(s); console.log(s); };

  say('=== MASTHEAD DIAGNOSIS ===');
  if (!sd) { say('✗ __SparkData is MISSING — run setupMasthead() first.'); return L.join('\n'); }

  var state = String(sd.getRange('A6').getValue() || '');
  var off   = sd.getRange('A13').getValue();
  var hour  = Number(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'H'));
  var hh    = (hour < 10 ? '0' : '') + hour;
  say('state (__SparkData!A6) : "' + state + '"' + (state ? '' : '   ✗ EMPTY — the URL will 404'));
  say('off-hours (A13)        : ' + off);
  say('sheet timezone         : ' + ss.getSpreadsheetTimeZone() + '  → hour ' + hh);

  say('\n--- the formulas actually IN the cells ---');
  say('A1: ' + sheet.getRange(Schema.cellMasthead).getFormula());
  // ⚠ A2 holds the eBay LOGO while MASTHEAD.sky is off. Reading it unconditionally and
  //   labelling it "sky" described a feature that is switched OFF and mislabelled the
  //   one that is on — row 2 is the eBay table's NAME, not a canvas.
  say('A2 (' + (MASTHEAD.sky ? 'sky' : 'eBay logo') + '): ' +
      sheet.getRange(MASTHEAD.skyCell).getFormula());
  say('D1: ' + sheet.getRange(Schema.cellStats).getFormula());

  // ---- GEOMETRY. The masthead is sized in PIXELS against columns it does not own. ----
  // ⚠⚠ applyBrandTheme sets explicit widths on A-J, so a theme re-apply can silently
  //    resize the zones the face and the headline live in. These are the numbers to pin.
  say('\n--- geometry (pin these into applyBrandTheme) ---');
  var w = [], tot = 0;
  for (var c = 1; c <= Schema.dataWidth; c++) {
    var cw = sheet.getColumnWidth(c);
    w.push(String.fromCharCode(64 + c) + '=' + cw);
    if (c <= 3) tot += cw;
  }
  say('widths : ' + w.join('  '));
  say('A1:C1  : ' + tot + 'px   vs MASTHEAD.imgW ' + MASTHEAD.imgW +
      (tot === MASTHEAD.imgW ? '   ✓' : '   ⚠ MISMATCH — mode-4 sizing will stretch or clip'));
  say('rows   : 1=' + sheet.getRowHeight(1) + 'px (MASTHEAD.rowHeight ' + MASTHEAD.rowHeight +
      ', imgH ' + MASTHEAD.imgH + ')   2=' + sheet.getRowHeight(2) + 'px');
  say('hidden : I=' + sheet.isColumnHiddenByUser(9) + '  J=' + sheet.isColumnHiddenByUser(10));
  var mg = [];
  sheet.getRange(1, 1, 2, Schema.dataWidth).getMergedRanges()
       .forEach(function (r) { mg.push(r.getA1Notation()); });
  say('merges : ' + (mg.length ? mg.join('  ') : 'none'));
  say('pickers: shipping=' + Schema.pickIdA1() + '  adjustment=' + Schema.pickIdA1('adjustment'));

  say('\n--- what the cells RENDER ---');
  say('A1 shows: "' + sheet.getRange(Schema.cellMasthead).getDisplayValue() + '"' +
      '   (the "HQ" chip means IMAGE() errored and IFERROR caught it)');

  say('\n--- can the URLs actually be fetched? ---');
  // ⚠ Only probe the sky when it is switched ON. It is off, no sky art of any version
  //   has ever been on the server, and probing it printed a guaranteed ✗ every run —
  //   a diagnostic that always shows one failure teaches you to skim past its failures.
  var probes = [['face', MASTHEAD.baseUrl + state + '-h' + hh + '-' + MASTHEAD.version + '.' + MASTHEAD.ext]];
  if (MASTHEAD.sky) {
    probes.push(['sky', MASTHEAD.baseUrl + 'sky-h' + hh + '-' + MASTHEAD.version + '.' + MASTHEAD.ext]);
  }
  probes.forEach(function (pair) {
    var url = pair[1];
    try {
      var r  = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      var ct = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '?');
      var n  = r.getContent().length;
      var ok = r.getResponseCode() === 200 && ct.indexOf('image/') === 0;
      say((ok ? '✓ ' : '✗ ') + pair[0] + '  ' + r.getResponseCode() + '  ' + ct + '  ' + n + ' bytes');
      say('   ' + url);
      if (!ok && ct.indexOf('html') > -1) {
        say('   ⚠ HTML, not an image — this URL does not exist and Caddy served the board.');
      }
    } catch (e) {
      say('✗ ' + pair[0] + ' fetch threw: ' + e);
      say('   ' + url);
    }
  });
  say('\n⚠ If BOTH fetch fine here but the cells still show the fallback, the URL is right and');
  say('  the problem is Sheets refusing to render it — paste this into any empty cell to');
  say('  isolate that (a literal URL, no formula):');
  say('  =IMAGE("' + MASTHEAD.baseUrl + state + '-h' + hh + '-' + MASTHEAD.version +
      '.' + MASTHEAD.ext + '",4,' + MASTHEAD.imgH + ',' + MASTHEAD.imgW + ')');
  return L.join('\n');
}
