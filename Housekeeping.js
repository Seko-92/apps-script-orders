// =======================================================================================
// HOUSEKEEPING.gs — hourly freshness pass for the satellite sheets + "pulse chip" UI
// =======================================================================================
//
// PURPOSE
//   The satellite sheets used to depend on a weekly trigger (Out of Stock,
//   Mon 6am) and manual sidebar buttons that workers forget to press. This
//   file owns the FRESHNESS LAYER that replaces that:
//
//   1) runHourlyHousekeeping() — ONE hourly time trigger (work-hours gated,
//      America/Chicago) that runs:
//        · refreshOutOfStock(maps)         — smart-merge OOS from Master Inventory
//        · refreshPrepQueueLocations(maps) — re-mirror Prep Queue LOCATION from MI
//      Both share a single buildLocationAndInventoryMaps() read — one MI scan
//      per hour instead of two. Hourly matches the upstream cadence: MI's own
//      qty/location data only changes hourly (MAIN Smart Sync) plus instantly
//      per order (per-order GetItem refresh), so running more often would
//      re-read the same data.
//
//   2) PULSE CHIP — a live "⟳ 9:04 AM · 12m ago" cell in each sheet's frozen
//      header row, so anyone looking at the sheet knows how fresh it is
//      without opening the sidebar:
//        · Each full refresh stamps a real Date into a hidden "stamp" cell
//          (invisible ink: dark font on the dark header band, one column
//          right of the chip).
//        · The chip cell is a NOW()-based formula rendering the stamp as
//          "⟳ h:mm AM/PM · Xm ago" — re-renders on every recalc.
//        · CF font-color tiers on the chip: GREEN fresh (< 2h) · AMBER aging
//          (2–26h — expected overnight/weekend) · RED (> 26h, or never) =
//          the hourly trigger has missed a full day, something is broken.
//        · SELF-TICKING requires the spreadsheet setting File → Settings →
//          Calculation → Recalculation = "On change and every minute".
//          Without it the chip still re-renders on any edit anywhere (fine
//          during the workday, frozen overnight). Zero Apps Script quota
//          either way — recalc is Sheets-side.
//
//   The chip deliberately means "last FULL refresh of this sheet". Per-row
//   lookups (onEdit auto-fill, sidebar Quick Add) do NOT stamp it.
//
// WHY APPS SCRIPT, NOT AN n8n WORKFLOW
//   The refresh logic already lives here and reads MI directly. An n8n
//   workflow would just be an external clock calling the same functions via
//   /exec — extra moving parts AND it would drag this onto the pinned-version
//   deployment surface (Gotcha #12). A time trigger always runs latest pushed
//   code: `clasp push` is the whole deploy.
//
// SETUP (one-time, from the Apps Script editor)
//   Run setupHousekeeping(). It (a) installs the pulse chips on both sheets,
//   (b) removes the old weekly refreshOutOfStock trigger + any prior hourly
//   housekeeping trigger, (c) installs the hourly trigger, (d) runs one pass
//   immediately so the chips show real data. Idempotent — safe to re-run.
//
// PUBLIC API
//   runHourlyHousekeeping()      — trigger handler (also runnable manually)
//   setupHousekeeping()          — one-time: chips + trigger swap + first pass
//   removeHousekeepingTrigger()  — uninstall the hourly trigger
//   stampSheetPulse(sheet, a1)   — write "now" into a sheet's stamp cell
//   _installPulseChip(sheet, cfg)— (private) chip formula + styling + CF tiers
// =======================================================================================

// ---------- PULSE CHIP GEOMETRY ----------
// Chip + stamp live in the frozen header row, in the first free columns to
// the RIGHT of each sheet's headers — no schema shift, no dataStartRow change.
// Sheet names are literals (not OUT_OF_STOCK.sheetName etc.) so this global
// has no cross-file load-order dependency at init time.
var SHEET_PULSE = {
  outOfStock: { sheetName: "Out of Stock", chip: "I1", stamp: "J1" },
  prepQueue:  { sheetName: "Prep Queue",   chip: "G1", stamp: "H1" }
};

// Work-hours gate (America/Chicago). 6am start so the sheets are fresh
// before the Houston shift; 6pm stop — nobody reads them overnight and MI
// barely moves.
var HOUSEKEEPING_START_HOUR = 6;   // inclusive
var HOUSEKEEPING_END_HOUR   = 18;  // exclusive


// =======================================================================================
// TRIGGER HANDLER
// =======================================================================================

/**
 * Hourly trigger handler. Gates to Houston work hours, then runs the shared
 * freshness pass. Time triggers pass an event object as the first argument —
 * deliberately ignored (this function takes no meaningful params).
 */
function runHourlyHousekeeping() {
  var hour = parseInt(Utilities.formatDate(new Date(), "America/Chicago", "H"), 10);
  if (hour < HOUSEKEEPING_START_HOUR || hour >= HOUSEKEEPING_END_HOUR) {
    console.log("Housekeeping: off-hours skip (Houston hour " + hour + ")");
    return "⏸ Off-hours skip (Houston hour " + hour + ")";
  }
  return _housekeepingPass();
}


/**
 * The actual freshness pass — no time gate, so setupHousekeeping can force an
 * immediate first run at any hour. Each job is isolated in its own try/catch:
 * one failing refresh never blocks the other (same defense-in-depth rule as
 * Main.js's onEditInstallable handler chain).
 */
function _housekeepingPass() {
  var maps = buildLocationAndInventoryMaps();   // ONE MI read shared by both jobs

  var parts = [];
  try { parts.push(refreshOutOfStock(maps)); }
  catch (e) { parts.push("❌ OOS refresh: " + e); console.log("Housekeeping OOS error: " + e); }

  try { parts.push(refreshPrepQueueLocations(maps)); }
  catch (e) { parts.push("❌ Prep locations: " + e); console.log("Housekeeping Prep error: " + e); }

  var summary = parts.join("  |  ");
  console.log("Housekeeping: " + summary);
  return summary;
}


// =======================================================================================
// SETUP / TEARDOWN (run from the Apps Script editor)
// =======================================================================================

/**
 * One-time setup — idempotent, safe to re-run any time:
 *   1) Re-runs both sheet setups (creates the sheets if missing, re-applies
 *      styling + the pulse chips + the plain-text date format / robust
 *      DAYS OUT formula on Out of Stock).
 *   2) Removes the old weekly refreshOutOfStock trigger (superseded) and any
 *      existing hourly housekeeping trigger, then installs a fresh hourly one.
 *   3) Runs one pass immediately (ignores the work-hours gate) so the chips
 *      show real data right away instead of "NEVER SYNCED". Order matters:
 *      the pass runs AFTER setup so the Out of Stock rewrite lands on the
 *      plain-text-formatted columns (no date re-coercion).
 */
function setupHousekeeping() {
  var msgs = [];

  // --- 1) Sheet setups (each installs its own pulse chip) ---
  try { setupOutOfStockSheet(); msgs.push("Out of Stock re-styled + chip"); }
  catch (e) { msgs.push("⚠ Out of Stock setup: " + e); console.log("setupHousekeeping OOS setup error: " + e); }

  try { setupPrepQueueSheet(); msgs.push("Prep Queue re-styled + chip"); }
  catch (e) { msgs.push("⚠ Prep Queue setup: " + e); console.log("setupHousekeeping Prep setup error: " + e); }

  // --- 2) Trigger swap ---
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'refreshOutOfStock' || h === 'runHourlyHousekeeping') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger('runHourlyHousekeeping').timeBased().everyHours(1).create();
  msgs.push("hourly trigger installed (" + removed + " old trigger(s) removed)");

  // --- 3) Immediate first pass so the chips populate now ---
  msgs.push(_housekeepingPass());

  var summary = "✅ Housekeeping ready — " + msgs.join(" · ");
  console.log(summary);
  return summary;
}


/** Uninstall the hourly housekeeping trigger. Manual cleanup helper. */
function removeHousekeepingTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runHourlyHousekeeping') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  console.log("Removed " + removed + " housekeeping trigger(s).");
  return "Removed " + removed + " housekeeping trigger(s).";
}


// =======================================================================================
// PULSE CHIP — stamp + install
// =======================================================================================

/**
 * Write "now" into a sheet's stamp cell. Called at the END of every full
 * refresh so the chip means "last completed refresh" — a refresh that throws
 * never stamps, and the chip's staleness tiers become the failure alarm.
 * Best-effort: a stamp failure must never break the refresh that called it.
 */
function stampSheetPulse(sheet, stampA1) {
  try {
    sheet.getRange(stampA1).setValue(new Date());
  } catch (e) {
    try { console.log("stampSheetPulse error: " + e); } catch (_) {}
  }
}


/**
 * Install (or re-install) the pulse chip on a sheet: formula + header-band
 * styling + CF color tiers. Idempotent — strips any prior CF rules that
 * target the chip cell before re-adding, and leaves every other rule alone
 * (single-cell row-1 match, so the sheets' own column-level CF filters never
 * collide with it).
 */
function _installPulseChip(sheet, cfg) {
  var chip  = sheet.getRange(cfg.chip);
  var stamp = sheet.getRange(cfg.stamp);

  // "J1" → "$J$1" for formula references that survive any future edits around them
  var stampAbs = cfg.stamp.replace(/([A-Z]+)(\d+)/, "$$$1$$$2");

  // --- Chip formula: "⟳ 9:04 AM · 12m ago" (m → h → d as the gap grows) ---
  chip.setFormula(
    '=IF(' + stampAbs + '="","⟳ NEVER SYNCED",' +
    '"⟳ "&TEXT(' + stampAbs + ',"h:mm AM/PM")&" · "&' +
    'IF(NOW()-' + stampAbs + '<1/24,MAX(0,ROUND((NOW()-' + stampAbs + ')*1440))&"m ago",' +
    'IF(NOW()-' + stampAbs + '<1,ROUND((NOW()-' + stampAbs + ')*24,1)&"h ago",' +
    'ROUND(NOW()-' + stampAbs + ',1)&"d ago")))'
  );

  // --- Styling: extend the dark header band through the chip ---
  // The thick yellow BOTTOM border matters: every real header cell carries it
  // (setup functions apply it A1:<last>1), so without it the band visibly
  // "breaks" at the chip — that was the first-screenshot feedback (2026-07-13).
  chip.setBackground('#1d1d1b')
      .setFontColor('#81c784')          // fresh-green base; CF overrides below
      .setFontFamily('Oswald')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(null, null, true, null, null, null,
                 '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
  chip.setNote(
    "Last full refresh of this sheet.\n" +
    "Green < 2h · amber 2–26h (normal overnight) · red > 26h or never (refresh trigger is dead)."
  );
  sheet.setColumnWidth(chip.getColumn(), 180);

  // Stamp cell: real Date (not a formatted string) so the chip formula and
  // the CF tiers can do date math on it directly. The whole COLUMN is hidden
  // below — a visible empty dark cell after the chip read as a glitch.
  stamp.setBackground('#1d1d1b')
       .setFontColor('#1d1d1b')
       .setFontSize(8)
       .setHorizontalAlignment('center')
       .setNumberFormat('M/d/yy h:mm:ss');
  try { sheet.hideColumns(stamp.getColumn()); } catch (e) {}

  // --- CF tiers on the chip (font color only; bg stays the dark band) ---
  var chipCol = chip.getColumn();
  var rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) {
      return rg.getRow() === 1 && rg.getNumRows() === 1 && rg.getColumn() === chipCol;
    });
  });

  function chipRule(formula, fontColor) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setFontColor(fontColor)
      .setRanges([chip])
      .build();
  }

  // Order matters — first match wins. Red (never / dead trigger), then amber
  // (aging — expected overnight), else the base green set above shows through.
  rules.push(chipRule('=' + stampAbs + '=""', '#ff6b6b'));
  rules.push(chipRule('=AND(' + stampAbs + '<>"",NOW()-' + stampAbs + '>=26/24)', '#ff6b6b'));
  rules.push(chipRule('=AND(' + stampAbs + '<>"",NOW()-' + stampAbs + '>=2/24)', '#ffd966'));
  sheet.setConditionalFormatRules(rules);
}

