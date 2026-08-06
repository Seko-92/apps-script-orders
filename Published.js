// =======================================================================================
// PUBLISHED.gs — compute once, write it down, let everyone read it
// =======================================================================================
//
// THE PROBLEM IT SOLVES
//   The board recomputes the tick on every poll, on every device, and Apps
//   Script runs one execution at a time. Cost scales with VIEWERS, not with
//   ACTIVITY — two devices already queue behind each other. Caching helped, but
//   a cache is still a recomputation waiting to happen on every miss.
//
// THE MOVE
//   Apps Script writes the finished tick to a cell. n8n reads that cell. The
//   board reads n8n. Nobody recomputes anything to look at a screen.
//
//   This is NOT a new pattern here — Kit Health, Out of Stock and Price Audit
//   are already computed views that everything else merely reads. The tick is
//   the one hot read that never got the treatment, which is why it is the one
//   that broke.
//
// WHY THE RESULT AND NOT THE LOGIC
//   The alternative was reimplementing the tick in n8n so it could read Sheets
//   directly. That is 200-300 lines of real logic (boundary detection, natural
//   aisle sort, kit detection, aging, paid-shipping, Activity Log tallies)
//   living in a browser-edited Code node on the VPS: not version controlled,
//   not diffable, not editable by Claude. Publishing the RESULT gets the same
//   latency with ZERO duplicated logic.
//
// COST MODEL — the whole point
//                       before                    after
//   scales with         viewers x poll rate       order activity
//   10 boards open      10x the executions        unchanged
//   quiet evening       still polling all night   a flag check
//
// DIRTY FLAG, NOT INLINE PUBLISH
//   Publishing inline at each chokepoint would add seconds to every status
//   flip, and a Telegram PREP tap must stay instant. So chokepoints only set a
//   flag (one property write, ~50ms) and a 2-minute trigger does the work —
//   and only when something actually changed.
//
// NOTHING HERE IS LOAD-BEARING YET. The board still runs on the live path; this
// only writes a copy. n8n is pointed at it as a separate, reversible step.
// =======================================================================================

var PUBLISHED = {
  sheetName: "__Published",

  cells: {
    TICK:      "A1",   // the JSON
    UPDATED:   "A2",   // real Date — when it was written
    ERROR:     "A3",   // last publish failure, blank when healthy
    SIZE:      "A4"    // char count, so growth toward the cap is visible
  },

  // A cell holds 50,000 characters (Gotcha #13 — the limit that broke the Zoho
  // payload cache). Trim BEFORE the write, never let it throw: a failed write
  // would leave readers on a stale copy with no explanation.
  maxChars: 45000,

  dirtyPropKey: "PUBLISHED_TICK_DIRTY",

  // Republish once the copy reaches this age even if NOTHING changed.
  //
  // ⚠ FOUND 2026-08-07, reading a live payload. Publishing only on "dirty" has
  // two failure modes on a quiet night, and they compound:
  //   1. The tick carries ELAPSED values (lastSyncMinutes, oldestPendingMinutes)
  //      computed at publish time. Frozen, `lastSyncMinutes: 5` stays 5 forever
  //      and the board keeps showing LIVE while the sync is long dead.
  //   2. n8n only trusts a payload for 10 minutes, so a never-republished copy
  //      means EVERY poll falls through to Apps Script — reinstating exactly the
  //      per-viewer cost this whole design removes.
  //
  // Must stay comfortably UNDER the reader's trust window (10 min in the n8n
  // proxy) or there is a gap where nothing is servable.
  maxAgeMinutes: 8,

  // ⚠ Apps Script's everyMinutes() accepts ONLY 1, 5, 10, 15 or 30. Anything
  // else throws at trigger-creation time. (Cost 2026-08-07: I picked 2.)
  //
  // WHY 5 AND NOT 1. A trigger run is not free even when it does nothing —
  // project load plus the flag read is roughly a second of runtime. At 1 minute
  // that is ~1,440 executions a day, call it 35 min of the ~90 min daily quota,
  // spent mostly on runs that find nothing to do. At 5 minutes it is ~7 min/day.
  //
  // The cost of 5 is FRESHNESS: a new order could sit up to 5 minutes before
  // the published copy shows it. That is only acceptable because (a) the board
  // is an ambient display, (b) the Telegram ping is the real-time alert for
  // arrivals, and (c) ✓ Pick already flips optimistically on the device.
  //
  // If 5 minutes proves too slow once the board actually reads this, the fix is
  // NOT a faster timer — it is publishing INLINE on the doPost insert path,
  // which is machine-facing (n8n waits, no human does) and costs only as much
  // as real order volume. Deliberately not built yet: measure first.
  triggerMinutes: 5
};


// =======================================================================================
// PUBLISH
// =======================================================================================

/**
 * Compute the tick and write it to the published cell.
 *
 * Always publishes, dirty or not — callable by hand to force a refresh.
 * runPublishTick() is the trigger target that respects the flag.
 *
 * @returns {{ok:boolean, bytes:number, trimmed:boolean, ms:number, message:string}}
 */
function publishBoardTick() {
  var t0 = Date.now();
  try {
    var sheet = _pubSheet();
    if (!sheet) return { ok: false, bytes: 0, trimmed: false, ms: 0, message: "No __Published sheet." };

    // Build FRESH. Deliberately not getDashboardTick(): that would happily hand
    // back a cached copy, and publishing a cache of a cache is how staleness
    // compounds invisibly.
    var tick = _buildDashboardTick();
    tick._publishedAt = new Date().toISOString();

    var json = JSON.stringify(tick);
    var trimmed = false;

    // Shed the least valuable payload first: the timeline is a nice-to-have,
    // the pick list is the reason the board exists.
    if (json.length > PUBLISHED.maxChars && tick.cockpit && tick.cockpit.timeline) {
      tick.cockpit.timeline = [];
      tick._trimmed = trimmed = true;
      json = JSON.stringify(tick);
    }
    if (json.length > PUBLISHED.maxChars && tick.openOrders) {
      tick.openOrders = tick.openOrders.slice(0, 25);
      tick._trimmed = trimmed = true;
      json = JSON.stringify(tick);
    }
    if (json.length > PUBLISHED.maxChars) {
      // Should be unreachable. Refuse rather than throw — readers keep the last
      // good copy and the error cell says why.
      sheet.getRange(PUBLISHED.cells.ERROR).setValue(
        "payload " + json.length + " chars — over the " + PUBLISHED.maxChars + " cap, not written");
      return { ok: false, bytes: json.length, trimmed: trimmed, ms: Date.now() - t0,
               message: "Payload too large — not written." };
    }

    sheet.getRange(PUBLISHED.cells.TICK).setValue(json);
    sheet.getRange(PUBLISHED.cells.UPDATED).setValue(new Date());
    sheet.getRange(PUBLISHED.cells.ERROR).setValue("");
    sheet.getRange(PUBLISHED.cells.SIZE).setValue(json.length);

    _pubClearDirty();
    return { ok: true, bytes: json.length, trimmed: trimmed, ms: Date.now() - t0, message: "" };

  } catch (err) {
    try {
      var sh = _pubSheet();
      if (sh) sh.getRange(PUBLISHED.cells.ERROR).setValue(String(err.message || err));
    } catch (e) { /* nothing more we can do */ }
    try { console.log("publishBoardTick: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, bytes: 0, trimmed: false, ms: Date.now() - t0, message: String(err.message || err) };
  }
}


/**
 * Trigger target. Publishes ONLY when a chokepoint marked the tick dirty, so a
 * quiet night costs a property read rather than a rebuild every two minutes.
 */
function runPublishTick() {
  try {
    var dirty = _pubIsDirty();
    var stale = _pubIsStale();
    if (!dirty && !stale) return "clean and fresh — skipped";
    var res = publishBoardTick();
    return res.ok
      ? ("published " + res.bytes + " chars in " + res.ms + "ms  (" +
         (dirty ? "changed" : "keep-fresh") + ")")
      : ("publish failed: " + res.message);
  } catch (err) {
    try { console.log("runPublishTick: " + err); } catch (_) {}
    return "error: " + String(err.message || err);
  }
}


/**
 * Read the published payload back — what n8n will do, and a way to verify by
 * hand before anything is pointed at it.
 *
 * @returns {{ok:boolean, ageSec:number, bytes:number, tick:Object|null, message:string}}
 */
function getPublishedTick() {
  try {
    var sheet = _pubSheet(true);
    if (!sheet) return { ok: false, ageSec: -1, bytes: 0, tick: null, message: "No __Published sheet." };

    var json = String(sheet.getRange(PUBLISHED.cells.TICK).getValue() || "");
    if (!json) return { ok: false, ageSec: -1, bytes: 0, tick: null, message: "Nothing published yet." };

    var when = sheet.getRange(PUBLISHED.cells.UPDATED).getValue();
    var ageSec = (when instanceof Date) ? Math.round((Date.now() - when.getTime()) / 1000) : -1;

    return { ok: true, ageSec: ageSec, bytes: json.length, tick: JSON.parse(json), message: "" };
  } catch (err) {
    return { ok: false, ageSec: -1, bytes: 0, tick: null, message: String(err.message || err) };
  }
}


// =======================================================================================
// DIRTY FLAG
// =======================================================================================
//
// Set by _dashBustTickCache(), which every write chokepoint already calls —
// updateOrderStatus, the doPost insert, and boardSetStatus. Hooking there means
// one edit covers every path, present and future, instead of three.

function _pubMarkDirty() {
  try { PropertiesService.getScriptProperties().setProperty(PUBLISHED.dirtyPropKey, "1"); }
  catch (e) { /* a missed flag costs one late publish, not correctness */ }
}

function _pubIsDirty() {
  try { return PropertiesService.getScriptProperties().getProperty(PUBLISHED.dirtyPropKey) === "1"; }
  catch (e) { return true; }   // unreadable -> assume dirty; publishing spuriously is the safe error
}

/**
 * Is the published copy old enough to need a refresh regardless of changes?
 * Unreadable or never-published counts as stale — publishing spuriously is the
 * safe direction, serving a frozen heartbeat is not.
 */
function _pubIsStale() {
  try {
    var sheet = _pubSheet(true);
    if (!sheet) return true;
    var when = sheet.getRange(PUBLISHED.cells.UPDATED).getValue();
    if (!(when instanceof Date)) return true;
    return (Date.now() - when.getTime()) > (PUBLISHED.maxAgeMinutes * 60 * 1000);
  } catch (e) {
    return true;
  }
}

function _pubClearDirty() {
  try { PropertiesService.getScriptProperties().deleteProperty(PUBLISHED.dirtyPropKey); }
  catch (e) { /* worst case: one extra publish */ }
}


// =======================================================================================
// SETUP
// =======================================================================================

/** The sheet, created on demand. @param {boolean} [readOnly] don't create it */
function _pubSheet(readOnly) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PUBLISHED.sheetName);
  if (sheet || readOnly) return sheet;

  sheet = ss.insertSheet(PUBLISHED.sheetName);
  sheet.getRange("B1").setValue("tick JSON  ·  written by publishBoardTick()  ·  DO NOT EDIT");
  sheet.getRange("B2").setValue("published at");
  sheet.getRange("B3").setValue("last error");
  sheet.getRange("B4").setValue("size (chars)");
  sheet.setColumnWidth(1, 120);          // keep the JSON cell narrow — it is not for reading
  sheet.hideSheet();
  return sheet;
}

/** One-time: create the sheet, publish once, arm the trigger. */
function setupPublishedTick() {
  var out = [];
  try { _pubSheet(); out.push("sheet ready"); }
  catch (e) { out.push("sheet failed: " + e); }

  var res = publishBoardTick();
  out.push(res.ok ? ("first publish: " + res.bytes + " chars in " + res.ms + "ms")
                  : ("first publish FAILED: " + res.message));

  out.push(setupPublishTrigger());
  var msg = out.join("\n");
  try { console.log(msg); } catch (_) {}
  return msg;
}

/** Arm the publish trigger. Idempotent — removes any previous one first. */
function setupPublishTrigger() {
  // Fail loudly on a bad constant rather than at ScriptApp, which reports it as
  // a runtime exception halfway through setup with the sheet already created.
  var ALLOWED = [1, 5, 10, 15, 30];
  if (ALLOWED.indexOf(PUBLISHED.triggerMinutes) === -1) {
    return "trigger NOT armed — triggerMinutes must be one of " + ALLOWED.join(", ") +
           " (got " + PUBLISHED.triggerMinutes + ")";
  }
  removePublishTrigger();
  ScriptApp.newTrigger("runPublishTick").timeBased()
    .everyMinutes(PUBLISHED.triggerMinutes).create();
  return "trigger armed — every " + PUBLISHED.triggerMinutes + " min";
}

function removePublishTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "runPublishTick") { ScriptApp.deleteTrigger(t); n++; }
  });
  return "removed " + n + " trigger(s)";
}
