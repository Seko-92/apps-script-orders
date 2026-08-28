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
  // OOS chip moved IN-BAND 2026-07-18 (two-table overhaul): H1, the right edge
  // of the row-1 ▌ OUT OF STOCK band (was a dark chip at I1 outside the
  // table). Stamp stays J1 — same cell as the old layout, hidden column.
  // setupOutOfStockSheet's chip migration wipes the old I1/I2 dark-chip homes.
  // Nothing on the pinned /exec writes this stamp (hourly trigger + sidebar
  // only), so the move needs no New Version.
  outOfStock: { sheetName: "Out of Stock", chip: "H1", stamp: "J1", inBand: true },
  // Prep chip FINAL home (2026-07-16, after the two-table + title-band
  // passes): F1, INSIDE the ▌ CURRENT band (inBand style — quiet ink on
  // yellow; a dark chip cell beside the band read as a floating island).
  // Stamp stays I1 (hidden col) — the pinned /exec 2-min writer targets it,
  // so the DISPLAY cell can move freely without a New Version.
  // setupPrepQueueSheet's chip migration cleans the older homes
  // (G1/H1 → H1/I1 → H2/I2 → H1 dark → F1 in-band).
  prepQueue:  { sheetName: "Prep Queue",   chip: "F1", stamp: "I1", inBand: true }
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
  var maps = buildLocationAndInventoryMaps();   // ONE MI read shared by every job below

  var parts = [];

  // HAND RECOMPUTE (2026-08-19) — moved here from a standalone every-15-minutes
  // trigger, and it belongs here for three reasons.
  //
  //   1. THE MAPS ARE ALREADY BUILT. That standalone trigger spent 10.7s a run
  //      almost entirely on re-reading Master Inventory and Zoho Stock; here the
  //      MI read above is already paid for, so this costs ~3s instead. 96 runs a
  //      day at 10.7s = 17.1 min of a 90-minute quota, against ~0.6 min here.
  //
  //   2. IT RUNS WHEN PEOPLE ARE THERE. The old trigger fired 24/7 — two thirds
  //      of its runs were overnight, recomputing for nobody, and mostly re-reading
  //      a Zoho Stock sheet that is not being refreshed then either (same gate).
  //      This pass is 6am–6pm Houston, which also covers the 6–9am window the
  //      Zoho push does not reach: that push is gated 9–17, so an early shift
  //      would otherwise be reading HAND last refreshed at 5pm yesterday.
  //
  //   3. HOURLY MATCHES THE SOURCE. MAIN's Smart Sync refreshes Master Inventory
  //      HOURLY, so recomputing every 15 minutes was 4x oversampling a number
  //      that only moves once an hour. Nothing is lost by matching its cadence.
  //
  // ⚠ THIS IS THE FALLBACK, NOT THE PRIMARY. During work hours the n8n Zoho push
  //   recomputes both of these every 2 minutes; this pass exists for the 6–9am
  //   and 5–6pm edges and for any hour the push does not arrive. Removing the
  //   push would leave HAND hourly, not stale.
  try {
    var hkZoho = buildZohoStockMap();
    parts.push(recomputeHand(maps, hkZoho));
    parts.push(refreshPrepQueueHand(maps, hkZoho));
  } catch (e) {
    parts.push("❌ HAND recompute: " + e);
    console.log("Housekeeping HAND error: " + e);
  }
  try { parts.push(refreshOutOfStock(maps)); }
  catch (e) { parts.push("❌ OOS refresh: " + e); console.log("Housekeeping OOS error: " + e); }

  try { parts.push(refreshPrepQueueLocations(maps)); }
  catch (e) { parts.push("❌ Prep locations: " + e); console.log("Housekeeping Prep error: " + e); }

  // ⚠ NEEDS PHOTOS IS NOT HERE ANY MORE (2026-08-19). It ran ~334s of this
  // pass's measured 352.7s — 95% of a 360s ceiling — for a backlog that moves in
  // single digits per day. It now has its OWN daily trigger at 5am Houston with
  // its own full execution budget: see runPhotoQueueRefresh in PhotoQueue.js.
  // Removing it took this pass from 352.7s to 18.2s.

  // STRAGGLER WATCHDOG (Telegram Tier C) — orders past the 3h line, SOs
  // part-shipped >24h, kits awaiting a decision. Alerts once per item, so this
  // is silent on most runs. Reads All Orders / Activity Log / Kit Registry
  // itself (none of that is in `maps`). runStragglerWatchdog never throws, but
  // it stays inside the same try/catch discipline as every other job here.
  try { parts.push(runStragglerWatchdog(maps)); }   // reuses the MI read above
  catch (e) { parts.push("❌ Watchdog: " + e); console.log("Housekeeping Watchdog error: " + e); }

  // RESTING-PANEL SNAPSHOT (2026-08-20) — the sidebar's two dark rows,
  // "Advertised, can't build" and "Restocking N parts frees N kits".
  //
  // ⚠⚠ THIS IS THE REASON IT LIVES HERE AND NOWHERE ELSE. The ripple half costs
  // ~6s because it builds the kit map plus the MI and Zoho availability maps.
  // The obvious homes are both wrong: the sidebar heartbeat would pay it per
  // poll, and runPublishTick fires EVERY MINUTE — 1,440 × 6s is ~2.4 hours of
  // runtime a day against a ~90-minute trigger budget, for a row nobody reads at
  // 3am. Once an hour is finer than the underlying numbers move.
  //
  // ⚠ AND IT RIDES THE MAPS THIS PASS ALREADY BUILT — `maps` from the top of the
  // function and the Zoho map built for the HAND recompute — so most of that ~6s
  // is already paid for. Same sharing rule as recomputeHand and the watchdog.
  // ⚠ hkZoho is `var`, so it is function-scoped and readable here even though it
  // is assigned inside the HAND try block above — and it is `undefined` if that
  // block threw before the assignment. Passing that through is safe: the shape
  // check inside analyzeRestockRipple falls back to building the map itself.
  try { parts.push(refreshRestSnapshot(maps, hkZoho || null)); }
  catch (e) { parts.push("❌ Rest snapshot: " + e); console.log("Housekeeping rest error: " + e); }

  // PUBLISHED-CELL PULSE (2026-08-18) — is the board's FREE path still free?
  // One outbound probe of the public board URL, compared against the last known
  // state; silent unless it CHANGED. Costs one HTTP call an hour and catches
  // the failure that silently ran for a week in Aug 2026 and makes every screen
  // start charging Apps Script for itself. See the long note in Watchdog.js.
  try { parts.push(checkPublishedPulse()); }
  catch (e) { parts.push("❌ Pulse: " + e); console.log("Housekeeping Pulse error: " + e); }

  // ORDER ARCHIVE (2026-08-28) — roll each COMPLETED order into one durable row
  // before purgeOldActivityLog destroys the events it was built from.
  //
  // ⚠ THIS IS THE CHEAPEST JOB IN THE PASS ON 11 OF 12 RUNS. It archives whole
  //   COMPLETED DAYS only, so once the watermark reaches yesterday it returns after a
  //   single Script Property read — no sheet access at all. The one expensive run per
  //   day reads the Activity Log once, in full, which is why it lives here rather than
  //   on a per-transition hook: n8n's shipped sweep flips orders in BATCHES, so
  //   hooking updateOrderStatus would mean one log read per flipped order AND would
  //   put work on the floor's hot write path, which this project's own ruling forbids.
  //
  // ⚠ It reads the Activity Log itself — none of that is in `maps` — but it stays
  //   inside the same try/catch discipline as every other job here.
  try { parts.push(runOrderArchiveSweep()); }
  catch (e) { parts.push("❌ Order archive: " + e); console.log("Housekeeping archive error: " + e); }

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

  // ⚠ No pulse chip on this one, deliberately — see the note on _oaPaintStatus.
  //   The shared chip's tiers go RED past 26h, and the archive is swept ONCE A DAY
  //   by design, so a perfectly healthy sheet would sit amber and drift red before
  //   every sweep. It carries a plain factual status line instead.
  try { setupOrderArchiveSheet(); msgs.push("Order Archive ready"); }
  catch (e) { msgs.push("⚠ Order Archive setup: " + e); console.log("setupHousekeeping Archive setup error: " + e); }

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

  // The photo scan lives on its own daily trigger now — installed here so this
  // stays the single setup entry point. Best-effort: a trigger-install failure
  // must not abort the rest of setup.
  try { msgs.push(setupPhotoQueueTrigger()); }
  catch (e) { msgs.push("❌ photo trigger: " + e); }

  // --- 3) Immediate first pass so the chips populate now ---
  // ⚠ This deliberately does NOT refresh the photo table. That scan is ~334s and
  // would push this function past the 360s execution ceiling. Its own trigger
  // covers it; the sidebar's 📸 button forces one on demand.
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

  // --- Styling: two variants ---
  if (cfg.inBand) {
    // IN-BAND (Prep Queue, 2026-07-16): the chip lives INSIDE the sheet's
    // yellow ▌ title band — quiet band-ink text at the band's right edge.
    // A separate dark chip cell next to the yellow band read as a floating
    // island (user feedback); in-band, the sync time is part of the design
    // and only turns loud (dark red/amber CF tiers below) when stale.
    // No own border or column width: the band's row styling rules the look.
    chip.setBackground('#ffd400')
        .setFontColor('#1d1d1b')
        .setFontFamily('Oswald')
        .setFontWeight('bold')
        .setFontSize(9)
        .setHorizontalAlignment('right')
        .setVerticalAlignment('middle');
  } else {
    // DARK: extends the dark header band through the chip. The thick yellow
    // BOTTOM border matters: every real header cell carries it (setup
    // functions apply it A1:<last>1), so without it the band visibly
    // "breaks" at the chip — first-screenshot feedback (2026-07-13).
    chip.setBackground('#1d1d1b')
        .setFontColor('#81c784')          // fresh-green base; CF overrides below
        .setFontFamily('Oswald')
        .setFontWeight('bold')
        .setFontSize(10)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(null, null, true, null, null, null,
                   '#ffd966', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.setColumnWidth(chip.getColumn(), 180);
  }
  chip.setNote(
    "Last full refresh of this sheet.\n" +
    "Quiet color < 2h · amber 2–26h (normal overnight) · red > 26h or never (refresh trigger is dead)."
  );

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
  // Strip prior rules on the chip's OWN cell (row from cfg, not hardcoded 1 —
  // the Prep chip lives on row 2 since the 2026-07-16 title-band layout).
  var chipCol = chip.getColumn();
  var chipRow = chip.getRow();
  var rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) {
      return rg.getRow() === chipRow && rg.getNumRows() === 1 && rg.getColumn() === chipCol;
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
  // (aging — expected overnight), else the base color set above shows
  // through. In-band chips sit on brand yellow, so their alert tiers use
  // DARK red/amber (the light palette is invisible on yellow).
  var tierRed   = cfg.inBand ? '#b71c1c' : '#ff6b6b';
  var tierAmber = cfg.inBand ? '#7a5c00' : '#ffd966';
  rules.push(chipRule('=' + stampAbs + '=""', tierRed));
  rules.push(chipRule('=AND(' + stampAbs + '<>"",NOW()-' + stampAbs + '>=26/24)', tierRed));
  rules.push(chipRule('=AND(' + stampAbs + '<>"",NOW()-' + stampAbs + '>=2/24)', tierAmber));
  sheet.setConditionalFormatRules(rules);
}

