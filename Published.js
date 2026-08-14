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

  // ⚠ INLINE PUBLISH — the debounce that makes it safe (2026-08-14).
  //
  // The dirty flag alone means a floor-visible change waits for the next
  // trigger run: up to 60s, plus n8n's 15s cache and the board's 20s poll —
  // ~95s worst case. Fine for an ambient change, far too slow for one a HUMAN
  // just made and is standing there waiting on. Expanding a kit rewrites the
  // pick list under the picker's feet: five new rows across five aisles.
  //
  // So the human-initiated paths publish INLINE instead of only marking dirty.
  // They are already multi-second operations (insert + enrichment + dup-repaint
  // + activity log), they are rare, and the person is watching a modal — the
  // exact opposite profile from a Telegram PREP tap, which is why the original
  // design chose a flag and why that choice still stands everywhere else.
  //
  // ⚠ THE DEBOUNCE IS WHAT KEEPS A BULK BATCH HONEST. The kit modal commits one
  // kit at a time through a queue, so a 9-kit session would otherwise pay for
  // nine full rebuilds back to back. Below this gap we skip and leave the dirty
  // flag set, so the trigger picks it up — never SLOWER than flag-only, just
  // not faster for the ones in the middle of a burst.
  inlineMinGapSec: 15,

  dirtyPropKey: "PUBLISHED_TICK_DIRTY",

  // ⚠ THE NET UNDER THE DIRTY FLAG (2026-08-14).
  //
  // The flag depends on every write path REMEMBERING to call
  // _dashBustTickCache(). That is an enumeration a human maintains, and it has
  // now been missed FIVE TIMES IN TWO DAYS: human sheet edits (08-14), then kit
  // expansion, the Zoho Pull insert and the Pull cancel (08-14). Apps Script
  // triggers do not fire for changes made BY a script, so every programmatic
  // path has to announce itself by hand — and nothing forces anyone to extend a
  // list. It will be missed again.
  //
  // So the publisher no longer trusts the enumeration: on the run where it would
  // otherwise SKIP, it fingerprints what the board actually renders from and
  // republishes if that moved without anyone saying so. A forgotten bust then
  // costs ~1 minute of staleness instead of the full keep-fresh window — the
  // same as a door that did report in.
  //
  // ⚠ THE FLAG IS STILL WORTH KEEPING. It is FASTER (it also clears the 45s
  // CacheService copy that the in-sheet board and the tier-3 live fallback
  // read), and the inline publish is faster still. This is a net, not a
  // replacement.
  fpPropKey: "PUBLISHED_TICK_FP",

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
  // ⚠ WAS 5, NOW 1 (2026-08-14) — and the note above about "measure first" is
  // exactly what happened, so here is the measurement.
  //
  // Until 2026-08-13 n8n could not read this cell at ALL (its Google Sheets
  // credential had an allowed-domains restriction that matched nothing), so the
  // publish was written every 5 min and thrown away every time while Apps
  // Script rebuilt the tick on essentially every poll — roughly 960 rebuilds a
  // day against a ~90 min quota. Fixing the credential took the board from
  // 7.5-35s to 0.65s AND handed back nearly all of that runtime budget.
  //
  // That budget is what pays for this. At 5 min the served tick was measured
  // between 31s and 272s old; cross-device status (tablet A picks a row, tablet
  // B still lists it) could lag ~5 min, which is a double-walk risk once two
  // pickers are on the floor. A 1-minute timer bounds it at ~1 min, and a run
  // that finds nothing dirty is a property read, not a rebuild.
  //
  // ⚠ STILL NOT inline-publish-on-write. publishBoardTick() does a FULL rebuild
  // (~7-14s) and Apps Script serialises executions, so publishing inline on
  // ✓ Pick would let a picker tapping five rows queue them into a jam. Arrivals
  // publish inline (OrderService doPost) because that path is machine-facing
  // and nobody waits on it. The dirty flag plus this timer is the right lever.
  //
  // ⚠ CHANGING THIS NUMBER DOES NOTHING ON ITS OWN — the interval is baked into
  // the trigger when it is created. Re-run setupPublishedTick() to re-arm.
  triggerMinutes: 1
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
    // ⚠ RECORD THE FINGERPRINT AFTER the write, so the next run compares against
    // the state this copy represents. Taken fresh rather than reused from the
    // build: if the sheet moved DURING the build, the two differ and the next
    // run republishes — which is the safe direction to be wrong in.
    _pubStoreFingerprint(_pubFingerprint());
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
/**
 * Publish RIGHT NOW for a change a human just made — unless we published a
 * moment ago, in which case ride the dirty flag instead.
 *
 * Call AFTER _dashBustTickCache(): that marks dirty, and this either clears the
 * flag by publishing or deliberately leaves it set for the next trigger run.
 * Either way the change is never lost — the only question is whether the floor
 * sees it in ~20s or ~60s.
 *
 * ⚠ BEST-EFFORT, ALWAYS. A publish failure must never fail the operation that
 * called it: the rows are already on the sheet, and the keep-fresh republish is
 * the backstop. Callers wrap this in try/catch and ignore the result.
 *
 * @param {number} [minGapSec] override the debounce (PUBLISHED.inlineMinGapSec)
 * @returns {{ok:boolean, skipped:boolean=, ms:number=, message:string=}}
 */
function publishBoardTickInline(minGapSec) {
  minGapSec = (typeof minGapSec === "number") ? minGapSec : PUBLISHED.inlineMinGapSec;
  try {
    // One cell read, deliberately NOT getPublishedTick() — that parses the whole
    // payload back out of JSON to answer a question about its age.
    var sheet = _pubSheet(true);
    if (sheet) {
      var when = sheet.getRange(PUBLISHED.cells.UPDATED).getValue();
      if (when instanceof Date) {
        var ageSec = (Date.now() - when.getTime()) / 1000;
        if (ageSec < minGapSec) {
          return { ok: true, skipped: true,
                   message: "published " + Math.round(ageSec) + "s ago — left dirty for the trigger" };
        }
      }
    }
  } catch (e) {
    // Could not read the age — publish rather than skip. Publishing spuriously
    // is the safe direction; serving a stale pick list is not.
    try { console.log("publishBoardTickInline age check: " + e); } catch (_) {}
  }
  return publishBoardTick();
}


function runPublishTick() {
  try {
    var dirty = _pubIsDirty();
    var stale = _pubIsStale();
    var why   = dirty ? "changed" : "keep-fresh";

    if (!dirty && !stale) {
      // ⚠ THE NET. Nobody said anything changed — check for ourselves rather
      // than take the enumeration's word for it. See PUBLISHED.fpPropKey.
      var fp = _pubFingerprint();

      // ⚠ AN UNREADABLE FINGERPRINT MEANS SKIP, NOT PUBLISH — the opposite of
      // _pubIsDirty's "assume dirty" rule, and deliberately so. That flag is one
      // property read that essentially cannot fail; this reads the sheet. If it
      // ever failed persistently, "assume changed" would rebuild every minute:
      // 1,440 full rebuilds a day, which is the entire runtime quota. Degrading
      // to today's behaviour (the 8-minute keep-fresh) is the safe direction.
      if (!fp) return "clean and fresh — skipped (fingerprint unavailable)";

      if (fp === _pubLastFingerprint()) return "clean and fresh — skipped";
      dirty = true;
      why   = "sheet moved, unannounced";
    }

    var res = publishBoardTick();
    return res.ok
      ? ("published " + res.bytes + " chars in " + res.ms + "ms  (" + why + ")")
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

/**
 * Fold a 2-D block of cell values into a short change-detection fingerprint.
 *
 * ⚠ PURE ON PURPOSE — no Sheets calls — so the part that actually decides
 * "did anything move?" can be tested in Node against real row shapes rather
 * than trusted. `_pubFingerprint()` is the thin sheet-reading wrapper.
 *
 * ⚠ NOT A CRYPTOGRAPHIC HASH, and deliberately not Utilities.computeDigest:
 * this needs to run every minute inside a trigger, and a plain-JS fold keeps it
 * both cheap and testable off-platform. Two independent 32-bit accumulators
 * (FNV-1a and djb2) plus the cell count give ~2^-64 odds of two different
 * sheets colliding — for "has this changed since a minute ago", that is far
 * past sufficient.
 *
 * Dates stringify stably, so a DATE cell does not churn the fingerprint.
 *
 * @param {Array<Array>} values the data range, row-major
 * @returns {string} short, stable, safe to keep in a Script Property
 */
function _pubFingerprintOf(values, skipIdx) {
  var rows = values || [];
  var skip = skipIdx || [];
  var h = 0x811c9dc5;          // FNV-1a offset basis
  var g = 5381;                // djb2
  var cells = 0;

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [];
    for (var c = 0; c < row.length; c++) {
      if (skip.indexOf(c) !== -1) continue;
      var v = row[c];
      var s = (v === null || v === undefined) ? "" : String(v);
      cells++;
      //  between cells so ["ab",""] and ["a","b"] cannot fold alike.
      s += "";
      for (var i = 0; i < s.length; i++) {
        var ch = s.charCodeAt(i);
        h = Math.imul(h ^ ch, 16777619);
        g = ((g << 5) + g + ch) | 0;
      }
    }
    h = Math.imul(h ^ 0x0a, 16777619);   // row terminator
    g = ((g << 5) + g + 10) | 0;
  }

  var hex = function (n) { return (n >>> 0).toString(16); };
  return rows.length + "." + cells + "." + hex(h) + hex(g);
}


/**
 * Fingerprint the All Orders DATA RANGE — exactly what the pick list is built
 * from. Banner rows are excluded: the clock in B1 changes every recalc and
 * would make the sheet look permanently modified.
 *
 * @returns {string} "" when it cannot be read — see the caller for why that
 *                   deliberately means "skip", not "publish".
 */
function _pubFingerprint() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "";
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return "empty";
    var n = lastRow - Schema.dataStartRow + 1;
    // ⚠ HAND IS DELIBERATELY EXCLUDED — this is the difference between a net and
    // a quota fire. recomputeHand rewrites column G on EVERY Zoho stock push
    // (every 2 minutes, 9–5) and does NOT bust the cache today, so folding it in
    // would make the sheet look permanently modified and turn a heartbeat into a
    // near-constant rebuilder — 3.5s each, several hundred a day.
    //
    // Excluding it changes nothing about HAND's freshness: it already rides the
    // 8-minute keep-fresh and always has. The net exists to catch a PICK-LIST
    // change nobody announced — a row appearing, a status flipping, a note being
    // written — not to shorten the staleness budget of a scheduled numeric sync.
    // If fresher HAND on the board is ever wanted, that is its own deliberate
    // decision with its own cost, not a side effect of this.
    return _pubFingerprintOf(
      sheet.getRange(Schema.dataStartRow, 1, n, Schema.dataWidth).getValues(),
      [Schema.idx("HAND")]
    );
  } catch (err) {
    try { console.log("_pubFingerprint: " + err); } catch (_) {}
    return "";
  }
}

function _pubStoreFingerprint(fp) {
  if (!fp) return;
  try { PropertiesService.getScriptProperties().setProperty(PUBLISHED.fpPropKey, fp); }
  catch (e) { /* a lost fingerprint costs one extra publish, not correctness */ }
}

function _pubLastFingerprint() {
  try { return PropertiesService.getScriptProperties().getProperty(PUBLISHED.fpPropKey) || ""; }
  catch (e) { return ""; }
}


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
