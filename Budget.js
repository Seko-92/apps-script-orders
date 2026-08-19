/**
 * Budget.js — WHERE DOES THE DAY ACTUALLY GO?
 *
 * A SELF-LIMITING PROBE, not a permanent instrument. It arms a one-minute
 * trigger, records what the two biggest suspects are doing, and then DELETES
 * ITS OWN TRIGGER when the window closes.
 *
 * ⚠⚠ WHY A WINDOW AND NOT AN ALL-DAY LOG — the probe is itself load.
 *   Every trigger run bills ~1.5-2s of container start whatever it does, so a
 *   permanent one-minute observer would cost ~40 min/day of the very budget it
 *   exists to measure. Absurd. One hour costs ~2 minutes and answers the
 *   question. Run a window, read it, stop. (Same ruling as design-lab/probe-load.js.)
 *
 * ⚠⚠ IT TOUCHES NO PRODUCTION CODE. Nothing in Published.js, UIService.js or
 *   Housekeeping.js is modified, wrapped or called. This file only READS two
 *   things that already exist:
 *
 *     1. __Published!A1:A4 — the board's published cell. A2 is a real Date
 *        stamped on every publish, and the tick JSON in A1 carries
 *        _publishReason. So "did runPublishTick actually publish this minute,
 *        and why" is answerable from outside.
 *
 *     2. The sidebar's SCRIPT CACHE entries. getSidebarTick stamps _builtAt
 *        into the cached payload, so a moving _builtAt IS a rebuild — visible
 *        to any function in this project without touching getSidebarTick.
 *
 * ⚠ THE BLIND SPOT, STATED SO NOBODY OVER-READS THE OUTPUT.
 *   This counts sidebar REBUILDS, not sidebar CALLS. A cache HIT leaves no
 *   trace anywhere observable, and a hit still bills ~1.5-2s of container
 *   start. So the sidebar's real cost is HIGHER than this probe can show, never
 *   lower. Counting calls needs two lines inside getSidebarTick — deliberately
 *   not done here, because that would mean editing a production path.
 *
 * ⚠ AND THE DURATIONS ARE NOT HERE ON PURPOSE. In-function timers report a
 *   FLOOR — measured 20-50x below the billed cost on 2026-08-19. The Executions
 *   panel is the only honest source for how long something took. This probe
 *   answers HOW OFTEN and WHY, which the panel makes you scroll thousands of
 *   rows to count. Pair the two: panel for duration, this for frequency.
 *
 * HOW TO USE IT
 *   budgetProbeStartNow()   arm for the default window (editor Run button)
 *   budgetProbeStatus()     how much longer, how many samples so far
 *   budgetProbeStop()       stop early — safe to call any time, idempotent
 *   budgetReport()          read the ledger and print the breakdown
 *
 * ⚠ Output goes to console.log, NOT a return value — the editor's Run button
 *   does not display returns (a lesson that cost an evening on getPublishedTick).
 */

/* ⚠ NAMED `PROBE`, NOT `BUDGET`, ON PURPOSE. Helpers.js:754 already declares a
   top-level `var BUDGET` (the config for timeBudget()). Apps Script puts every
   root .js file in ONE global scope, so a second `var BUDGET` here would
   silently clobber it — no error, no warning, timeBudget() just starts reading
   the wrong object. Caught by a duplicate-global scan before this shipped;
   run that scan before adding any new top-level var to this project. */
var PROBE = {
  sheetName: "__Budget",

  // One property holds the whole probe state. No state == no probe: an
  // unreadable or missing property makes budgetObserve stop rather than run
  // forever. FAIL TOWARD STOPPING is the whole safety story for a self-armed
  // trigger — the opposite default would be a job nobody remembers installing.
  propKey: "BUDGET_PROBE_STATE",

  // ⚠ Trigger rows are matched on this HANDLER NAME, never picked by eye.
  // Deleting a trigger by position is what caused the August nine-handler
  // outage; getHandlerFunction() is the only safe discriminator.
  handler: "budgetObserve",

  defaultMinutes: 60,
  maxMinutes: 240,          // hard ceiling — a probe that outlives its purpose is a bug

  headers: ["WHEN", "PUBLISHED?", "REASON", "CHARS", "SIDEBAR REBUILT?", "SLOW REBUILT?", "NOTE"]
};


// =======================================================================================
// ARM / DISARM
// =======================================================================================

/** Zero-arg wrapper — the editor Run button cannot pass arguments. */
function budgetProbeStartNow() {
  var msg = budgetProbeStart(PROBE.defaultMinutes);
  console.log(msg);
  return msg;
}


/**
 * Arm the probe for a window of N minutes.
 * @param {number} [minutes] defaults to PROBE.defaultMinutes, capped at maxMinutes
 */
function budgetProbeStart(minutes) {
  minutes = Math.min(PROBE.maxMinutes,
                     Math.max(1, parseInt(minutes, 10) || PROBE.defaultMinutes));

  var sheet = _budgetSheet(true);
  if (!sheet) return "❌ Could not create the " + PROBE.sheetName + " sheet.";

  _budgetClearTriggers();                       // never stack two observers

  var state = {
    endMs:    Date.now() + minutes * 60 * 1000,
    lastPub:  _budgetReadPublished().stampMs,   // seed, so the first sample is a delta
    lastTick: _budgetCacheBuiltAt(typeof SIDEBAR_TICK_CACHE_KEY === "string" ? SIDEBAR_TICK_CACHE_KEY : ""),
    lastSlow: _budgetCacheBuiltAt(typeof SIDEBAR_SLOW_CACHE_KEY === "string" ? SIDEBAR_SLOW_CACHE_KEY : ""),
    samples:  0,
    startedAt: Date.now()
  };
  _budgetSaveState(state);

  ScriptApp.newTrigger(PROBE.handler).timeBased().everyMinutes(1).create();

  sheet.appendRow([new Date(), "", "", "", "", "",
                   "▶ probe armed for " + minutes + " min (~" +
                   Math.round(minutes * 2 / 60 * 10) / 10 + " min of budget)"]);

  return "▶ Probe armed for " + minutes + " minutes. It disarms itself. " +
         "Read it with budgetReport().";
}


/**
 * Stop early, or clean up after the window closed. Idempotent — safe to run
 * twice, safe to run when nothing is armed.
 */
function budgetProbeStop() {
  var removed = _budgetClearTriggers();
  var state   = _budgetLoadState();

  try { PropertiesService.getScriptProperties().deleteProperty(PROBE.propKey); }
  catch (e) { /* nothing more we can do */ }

  var n = state ? state.samples : 0;
  var sheet = _budgetSheet(false);
  if (sheet) {
    try { sheet.appendRow([new Date(), "", "", "", "", "",
                           "■ probe stopped — " + n + " sample(s)"]); }
    catch (e) { /* the ledger is the point, not this line */ }
  }

  var msg = "■ Probe stopped. " + removed + " trigger(s) removed, " + n + " sample(s) recorded.";
  console.log(msg);
  return msg;
}


/** How much longer, and how many samples so far. */
function budgetProbeStatus() {
  var state = _budgetLoadState();
  if (!state) { console.log("No probe armed."); return "No probe armed."; }

  var leftMin = Math.max(0, Math.round((state.endMs - Date.now()) / 60000));
  var msg = "Probe running — " + state.samples + " sample(s), ~" + leftMin + " min left.";
  console.log(msg);
  return msg;
}


// =======================================================================================
// THE OBSERVER (trigger target)
// =======================================================================================

/**
 * One sample. Reads only things that already exist; writes only to __Budget.
 *
 * ⚠ THE WINDOW CHECK IS THE FIRST THING THAT HAPPENS, before any read, any
 *   sheet touch, any parse. Whatever else goes wrong below, a probe past its
 *   end time disarms — that is what makes a self-armed one-minute trigger safe
 *   to hand someone an hour before a shift.
 */
function budgetObserve() {
  var state;
  try { state = _budgetLoadState(); } catch (e) { state = null; }

  // No state means no probe. Stop rather than guess — an observer with no
  // memory cannot produce a delta anyway, and a runaway trigger is far worse
  // than a missing sample.
  if (!state || !state.endMs) { try { budgetProbeStop(); } catch (e) {} return; }
  if (Date.now() >= state.endMs) { try { budgetProbeStop(); } catch (e) {} return; }

  try {
    var pub  = _budgetReadPublished();
    var tick = _budgetCacheBuiltAt(typeof SIDEBAR_TICK_CACHE_KEY === "string" ? SIDEBAR_TICK_CACHE_KEY : "");
    var slow = _budgetCacheBuiltAt(typeof SIDEBAR_SLOW_CACHE_KEY === "string" ? SIDEBAR_SLOW_CACHE_KEY : "");

    var pubMoved  = (pub.stampMs > 0 && pub.stampMs !== state.lastPub);
    var tickMoved = (tick > 0 && tick !== state.lastTick);
    var slowMoved = (slow > 0 && slow !== state.lastSlow);

    var sheet = _budgetSheet(true);
    if (sheet) {
      sheet.appendRow([
        new Date(),
        pubMoved  ? "✓ published" : "—",
        pubMoved  ? (pub.reason || "(no reason stamped)") : "",
        pubMoved  ? pub.chars : "",
        tickMoved ? "✓ rebuilt" : "—",
        slowMoved ? "✓ rebuilt" : "—",
        pub.note
      ]);
    }

    state.lastPub  = pub.stampMs  || state.lastPub;
    state.lastTick = tick         || state.lastTick;
    state.lastSlow = slow         || state.lastSlow;
    state.samples  = (state.samples || 0) + 1;
    _budgetSaveState(state);

  } catch (err) {
    // A failed sample must never leave the trigger running blind. Record the
    // gap so the report reads it as a gap and not as a quiet zero — the stock
    // audit's rule: never mistake "I could not read it" for "nothing happened".
    try {
      var s = _budgetSheet(false);
      if (s) s.appendRow([new Date(), "?", "", "", "?", "?", "sample failed: " + err]);
    } catch (e) { /* give up on this sample, keep the probe alive */ }
    try {
      state.samples = (state.samples || 0) + 1;
      _budgetSaveState(state);
    } catch (e) { /* state is best-effort too */ }
  }
}


// =======================================================================================
// THE REPORT
// =======================================================================================

/**
 * Read the ledger and print what it means. Console output, never a return —
 * the Run button does not display return values.
 */
function budgetReport() {
  var sheet = _budgetSheet(false);
  if (!sheet) { console.log("No " + PROBE.sheetName + " sheet yet — run budgetProbeStartNow() first."); return; }

  var last = sheet.getLastRow();
  if (last < 2) { console.log("Ledger is empty."); return; }

  var all = sheet.getRange(2, 1, last - 1, PROBE.headers.length).getValues();

  /* ⚠ ONLY THE MOST RECENT WINDOW. Each arm writes a "▶ probe armed" banner, so
     everything after the LAST banner is one continuous run. Without this, a
     5-minute window at 5pm and a 60-minute window the next morning would be
     averaged across the 16-hour gap between them and every per-day figure would
     be nonsense — the measuring instrument quietly lying, which is worse than
     no instrument. Rows before the last banner stay in the sheet as history. */
  var startIdx = 0;
  for (var b0 = all.length - 1; b0 >= 0; b0--) {
    if (String(all[b0][6] || "").indexOf("\u25b6") === 0) { startIdx = b0 + 1; break; }
  }
  var rows = all.slice(startIdx);
  if (!rows.length) { console.log("No samples since the last probe was armed."); return; }

  var samples = 0, publishes = 0, tickBuilds = 0, slowBuilds = 0, failed = 0;
  var reasons = {}, firstWhen = null, lastWhen = null, chars = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var note = String(r[6] || "");
    if (note.indexOf("▶") === 0 || note.indexOf("■") === 0) continue;   // banner rows

    samples++;
    if (r[0] instanceof Date) { if (!firstWhen) firstWhen = r[0]; lastWhen = r[0]; }

    if (String(r[1]).indexOf("?") === 0) { failed++; continue; }

    if (String(r[1]).indexOf("✓") === 0) {
      publishes++;
      var why = String(r[2] || "(none)");
      reasons[why] = (reasons[why] || 0) + 1;
      chars = Math.max(chars, parseInt(r[3], 10) || 0);
    }
    if (String(r[4]).indexOf("✓") === 0) tickBuilds++;
    if (String(r[5]).indexOf("✓") === 0) slowBuilds++;
  }

  var spanMin = (firstWhen && lastWhen)
    ? Math.max(1, Math.round((lastWhen.getTime() - firstWhen.getTime()) / 60000) + 1)
    : samples;

  var perDay = function (n) { return Math.round(n / spanMin * 1440); };

  var out = [];
  out.push("═══ BUDGET PROBE ═══");
  out.push("window            " + spanMin + " min   (" + samples + " samples" +
           (failed ? ", " + failed + " failed" : "") + ")");
  out.push("");
  out.push("PUBLISH TICK  (runPublishTick fires 1440x/day)");
  out.push("  published       " + publishes + " in " + spanMin + " min  →  ~" + perDay(publishes) + "/day");
  out.push("  skipped         " + (samples - publishes - failed) + " in " + spanMin + " min  →  ~" +
           (1440 - perDay(publishes)) + "/day");
  out.push("  ⚠ EVERY skip still READS the All Orders range (the fingerprint net).");
  out.push("    There is no cheap path here — a run either fingerprints or rebuilds.");
  for (var k in reasons) { if (reasons.hasOwnProperty(k)) out.push("    · " + k + " x" + reasons[k]); }
  if (chars) out.push("  largest payload " + chars + " chars (cap " + PUBLISHED.maxChars + ")");
  out.push("");
  out.push("SIDEBAR");
  out.push("  tick rebuilds   " + tickBuilds + " in " + spanMin + " min  →  ~" + perDay(tickBuilds) + "/day");
  out.push("  slow rebuilds   " + slowBuilds + " in " + spanMin + " min  →  ~" + perDay(slowBuilds) + "/day");
  out.push("  ⚠ REBUILDS ONLY — cache hits are invisible from here and still bill");
  out.push("    ~1.5-2s of container start each. True sidebar cost is HIGHER than this.");
  out.push("");
  out.push("⚠ Durations are NOT here by design. Take them from the Executions panel,");
  out.push("  the only honest source: in-function timers report a floor 20-50x low.");
  out.push("  billed/day  ≈  (calls/day)  x  (mean panel duration)");
  out.push("═════════════════════");

  var text = out.join("\n");
  console.log(text);
  return text;
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/**
 * Read the published cell.
 *
 * ⚠ A1:A4 IN ONE getValues, DELIBERATELY. A1 is the whole ~15KB tick and we
 * only want one field out of it — but in Apps Script the ROUND TRIP dominates,
 * not the payload (proven twice: PartConsole in July, the narrow-read test on
 * 2026-08-19). Reading A2 alone would cost the same as reading all four, and
 * would throw away _publishReason for nothing.
 */
function _budgetReadPublished() {
  var out = { stampMs: 0, reason: "", chars: 0, note: "" };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(PUBLISHED.sheetName);
    if (!sh) { out.note = "no " + PUBLISHED.sheetName + " sheet"; return out; }

    var vals = sh.getRange("A1:A4").getValues();
    var json  = String(vals[0][0] || "");
    var when  = vals[1][0];
    var err   = String(vals[2][0] || "");

    if (when instanceof Date) out.stampMs = when.getTime();
    out.chars = json.length;
    if (err) out.note = "publish error cell: " + err;

    if (json) {
      try { out.reason = String(JSON.parse(json)._publishReason || ""); }
      catch (e) { out.reason = "(unparseable tick)"; }
    }
  } catch (e) {
    out.note = "read failed: " + e;
  }
  return out;
}


/**
 * When did a cached payload last get BUILT? The sidebar stamps _builtAt into
 * its own cache entry, so this reads a rebuild without touching getSidebarTick.
 * Returns 0 when there is nothing to read — never throws.
 */
function _budgetCacheBuiltAt(key) {
  if (!key) return 0;
  try {
    var hit = CacheService.getScriptCache().get(key);
    if (!hit) return 0;
    return parseInt(JSON.parse(hit)._builtAt, 10) || 0;
  } catch (e) { return 0; }
}


/** The ledger sheet. Hidden, like __Published and __SparkData. */
function _budgetSheet(createIfMissing) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(PROBE.sheetName);
    if (sh) return sh;
    if (!createIfMissing) return null;

    sh = ss.insertSheet(PROBE.sheetName);
    sh.getRange(1, 1, 1, PROBE.headers.length)
      .setValues([PROBE.headers])
      .setFontWeight("bold")
      .setBackground("#1d1d1b")
      .setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidth(7, 320);
    try { sh.hideSheet(); } catch (e) { /* visible is fine too */ }
    return sh;
  } catch (e) {
    try { console.log("_budgetSheet: " + e); } catch (_) {}
    return null;
  }
}


/**
 * Delete every trigger pointing at our handler.
 * ⚠ MATCHES ON getHandlerFunction(), never on position. Picking trigger rows by
 * eye is what took nine handlers down in August.
 */
function _budgetClearTriggers() {
  var n = 0;
  try {
    var all = ScriptApp.getProjectTriggers();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getHandlerFunction() === PROBE.handler) {
        ScriptApp.deleteTrigger(all[i]);
        n++;
      }
    }
  } catch (e) { try { console.log("_budgetClearTriggers: " + e); } catch (_) {} }
  return n;
}


function _budgetLoadState() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PROBE.propKey);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}


function _budgetSaveState(state) {
  try {
    PropertiesService.getScriptProperties().setProperty(PROBE.propKey, JSON.stringify(state));
  } catch (e) { try { console.log("_budgetSaveState: " + e); } catch (_) {} }
}
