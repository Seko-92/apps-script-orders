// =======================================================================================
// WATCHDOG.gs — the straggler watchdog (Telegram Layer, Tier C: "TELL ME")
// =======================================================================================
//
// PURPOSE
//   Every other Telegram command is something you ASK. This is the half that
//   speaks first. The user runs Houston from Riyadh, so `/status` only helps
//   on the days they remember to type it — a watchdog means they stop checking.
//
//   The motivating case, observed live 2026-08-05: order 24-14979-87359 sat
//   PENDING from 9:01 AM and was only noticed at 17:06, because someone
//   happened to run `/order` on it. Eight hours. This pings at hour three.
//
// WHAT IT WATCHES (user-chosen 2026-08-05)
//   1) PAST THE 3H LINE   — a PENDING order older than WATCHDOG.redlineMinutes.
//                           Same 180-min threshold the Floor Board already
//                           paints red, so the two surfaces can never disagree.
//   2) PART-SHIPPED >24h  — a SALES ORDER with at least one SHIPPED row AND at
//                           least one still-open row, where the last shipment
//                           was over a day ago. That's the forgotten-remainder
//                           case. It is ALSO the general form of the 2026-07-14
//                           kit-parent gap: if kit-parent auto-follow ever fails
//                           again, this catches it without knowing about kits.
//   3) KIT NEEDS A DECISION — a kit row sitting in DIRECT that nobody has
//                           expanded. Covers both outcomes: MANUAL kits waiting
//                           to be expanded, and READY kits waiting for the
//                           ship-as-box call. Exactly the 155583 / SO-24609 case.
//
// THE ANTI-NOISE RULE (the whole design, really)
//   ALERT ONCE PER ITEM, ON FIRST CROSSING. An alert you mute is worse than no
//   alert, so nothing here ever repeats itself. Every item gets a stable key
//   (`rl:<order>` / `ps:<so>` / `kx:<so>:<sku>`); keys already in the store are
//   silently skipped. One batched message per run, listing only what is NEW.
//
//   Keys are pruned after WATCHDOG.retentionDays. Deliberate consequence: an
//   item still unresolved after two weeks re-pings exactly once. That is a
//   feature — something stuck that long has earned a second look.
//
// COLD START
//   The first ever run would otherwise dump every currently-late order in one
//   wall of text. Instead it SEEDS the store silently and sends a one-line
//   "armed" notice with the counts — honest about what it found, without the
//   wall, and never silently swallowing the fact that things are late.
//
// ORDER OF OPERATIONS (matters)
//   Keys are recorded ONLY AFTER a successful send. If Telegram fails, nothing
//   is recorded and the next run retries. The failure mode is therefore a
//   possible duplicate, never a silent miss — the right way round for an alert.
//
// WORK-HOURS GATE
//   Runs from _housekeepingPass(), so it inherits the 6am–6pm America/Chicago
//   gate. An order crossing the line at 7pm is caught at 6am — correct, since
//   nothing can be picked at 7pm anyway, and it means no 3am pings.
//
// PUBLIC API
//   runStragglerWatchdog()      — gather → diff → send. Trigger + manual entry.
//   previewStragglerWatchdog()  — returns the text, SENDS NOTHING, MUTATES NOTHING.
//   resetStragglerWatchdog()    — clear the alerted store (re-arms cold start).
//   getStragglerCounts()        — {redline, partShipped, kits} for a future badge.
// =======================================================================================


var WATCHDOG = {
  redlineMinutes:   180,          // PENDING older than this = past the line
  partShippedHours: 24,           // last shipment older than this on a split SO
  alertedPropKey:   "WATCHDOG_ALERTED",
  retentionDays:    14,           // key lifetime; see "ALERT ONCE" above
  maxListed:        10            // items listed per section before summarising
};


// =======================================================================================
// ENTRY POINTS
// =======================================================================================

/**
 * The watchdog run: gather, diff against what we've already said, send the new.
 * Never throws — it is called from the housekeeping chain and must never be
 * able to take the hourly refresh down with it.
 *
 * @returns {string} short summary for the housekeeping log line
 */
function runStragglerWatchdog() {
  try {
    var found = _gatherStragglers();
    var store = _wdLoadAlerted();
    var isColdStart = (store === null);
    var known = isColdStart ? {} : store;

    // Split findings into already-told vs new.
    var fresh = { redline: [], partShipped: [], kits: [] };
    var allKeys = [];
    ["redline", "partShipped", "kits"].forEach(function (bucket) {
      found[bucket].forEach(function (item) {
        allKeys.push(item.key);
        if (!known[item.key]) fresh[bucket].push(item);
      });
    });

    var totalFound = allKeys.length;
    var totalFresh = fresh.redline.length + fresh.partShipped.length + fresh.kits.length;

    // ---- COLD START: seed silently, announce the counts only -------------------
    if (isColdStart) {
      var armed = "🛡 WATCHDOG ARMED\n\n" +
                  "Now watching: orders past the 3h line, SOs part-shipped over " +
                  WATCHDOG.partShippedHours + "h, and kits sitting unexpanded in DIRECT.\n\n" +
                  "Found right now (seeded, not alerting on these):\n" +
                  "  " + found.redline.length + " past the 3h line\n" +
                  "  " + found.partShipped.length + " part-shipped >" + WATCHDOG.partShippedHours + "h\n" +
                  "  " + found.kits.length + " kit(s) awaiting a decision\n\n" +
                  "From here you'll only hear about NEW ones, once each.";
      var armedSent = _tgSend(TELEGRAM_ADMIN_CHAT_ID, armed);
      if (armedSent) _wdRecord(allKeys, {});
      return "🛡 Watchdog armed (seeded " + totalFound + ")";
    }

    // ---- NORMAL RUN ------------------------------------------------------------
    if (totalFresh === 0) {
      // Still prune, so the store can't grow without bound on quiet days.
      _wdRecord([], known);
      return "🛡 Watchdog: nothing new (" + totalFound + " open)";
    }

    var text = _wdBuildText(fresh, totalFresh);
    var sent = _tgSend(TELEGRAM_ADMIN_CHAT_ID, text);

    // Record ONLY on success — a failed send must retry, never vanish.
    if (sent) {
      var newKeys = [];
      ["redline", "partShipped", "kits"].forEach(function (b) {
        fresh[b].forEach(function (i) { newKeys.push(i.key); });
      });
      _wdRecord(newKeys, known);
      return "🛡 Watchdog: alerted " + totalFresh;
    }
    return "⚠ Watchdog: send failed, will retry (" + totalFresh + " pending)";

  } catch (err) {
    try { console.log("runStragglerWatchdog failed: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return "❌ Watchdog: " + err;
  }
}


/**
 * Eyeball path — build the message from EVERYTHING currently found, ignoring
 * the already-told store. Sends nothing, writes nothing. Safe to run from the
 * editor or a sidebar button at any time.
 *
 * @returns {string} the text the watchdog would produce for the current state
 */
function previewStragglerWatchdog() {
  var found = _gatherStragglers();
  var total = found.redline.length + found.partShipped.length + found.kits.length;
  if (total === 0) return "✅ Nothing late — no stragglers, no part-shipped SOs, no kits waiting.";
  return _wdBuildText(found, total) +
         "\n\n(preview — shows everything currently late, not just new)";
}


/** Counts only, for a possible sidebar badge later. Never throws. */
function getStragglerCounts() {
  try {
    var f = _gatherStragglers();
    return { redline: f.redline.length, partShipped: f.partShipped.length, kits: f.kits.length };
  } catch (e) {
    return { redline: 0, partShipped: 0, kits: 0 };
  }
}


/** Clear the alerted store — the next run behaves like a cold start again. */
function resetStragglerWatchdog() {
  PropertiesService.getScriptProperties().deleteProperty(WATCHDOG.alertedPropKey);
  return "🛡 Watchdog store cleared — next run will re-arm and re-seed.";
}


// =======================================================================================
// DETECTION
// =======================================================================================

/**
 * Find every current straggler across the three categories.
 *
 * Reads: Activity Log once (RECEIVED + SHIPPED timestamps), All Orders once,
 * Kit Registry once. Each detector is isolated so one bad read degrades that
 * category to empty instead of killing the whole run.
 *
 * @returns {{redline:Array, partShipped:Array, kits:Array}}
 */
function _gatherStragglers() {
  var out = { redline: [], partShipped: [], kits: [] };
  var now = Date.now();

  // ---- Activity Log: earliest RECEIVED and latest SHIPPED, per order id -------
  var receivedMap = {};
  var lastShipMap = {};
  try {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var log = ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (log) {
      var logLast = log.getLastRow();
      if (logLast >= ACTIVITY_LOG.dataStartRow) {
        var logData = log.getRange(
          ACTIVITY_LOG.dataStartRow, 1,
          logLast - ACTIVITY_LOG.dataStartRow + 1,
          ACTIVITY_LOG.dataWidth
        ).getValues();

        var TS_I = ACTIVITY_LOG.idx("TIMESTAMP");
        var EV_I = ACTIVITY_LOG.idx("EVENT");
        var ID_I = ACTIVITY_LOG.idx("ORDER_ID");

        for (var i = 0; i < logData.length; i++) {
          var ts = logData[i][TS_I];
          if (!(ts instanceof Date)) continue;
          var ev  = String(logData[i][EV_I] || "").trim().toUpperCase();
          var oid = String(logData[i][ID_I] || "").trim();
          if (!oid) continue;
          var ms = ts.getTime();

          if (ev === "RECEIVED") {
            if (!receivedMap[oid] || ms < receivedMap[oid]) receivedMap[oid] = ms;
          } else if (ev === "SHIPPED") {
            if (!lastShipMap[oid] || ms > lastShipMap[oid]) lastShipMap[oid] = ms;
          }
        }
      }
    }
  } catch (e) {
    console.log("Watchdog: Activity Log read failed — " + e);
  }

  // ---- Kit registry (for the unexpanded-kit detector) ------------------------
  var kitMap = null;
  try { kitMap = buildKitMap(); }
  catch (e) { console.log("Watchdog: kit map read failed — " + e); }

  // ---- One pass over All Orders ----------------------------------------------
  var rows = [];              // {sku, loc, so, note, status, inDirect, rowNum}
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return out;
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return out;

    var data = sheet.getRange(
      Schema.dataStartRow, 1,
      lastRow - Schema.dataStartRow + 1,
      Schema.dataWidth
    ).getValues();

    var SKU_I = Schema.idx("SKU");
    var LOC_I = Schema.idx("LOCATION");
    var SO_I  = Schema.idx("SALES_ORDER");
    var NOTE_I = Schema.idx("NOTE");
    var ST_I  = Schema.idx("STATUS");

    var boundaryIdx = -1;
    for (var j = 0; j < data.length; j++) {
      var skuCell = String(data[j][SKU_I] || "").trim();

      if (skuCell.toUpperCase() === Schema.boundaryMarker) { boundaryIdx = j; continue; }
      if (boundaryIdx !== -1 && j === boundaryIdx + 1) continue;   // DIRECT header
      if (!skuCell) continue;

      rows.push({
        sku:      skuCell,
        loc:      String(data[j][LOC_I]  || "").trim(),
        so:       String(data[j][SO_I]   || "").trim(),
        note:     String(data[j][NOTE_I] || ""),
        status:   String(data[j][ST_I]   || "").trim().toUpperCase(),
        inDirect: (boundaryIdx !== -1 && j > boundaryIdx),
        arrIdx:   j
      });
    }
  } catch (e) {
    console.log("Watchdog: All Orders read failed — " + e);
    return out;
  }

  // ---- (1) PAST THE 3H LINE ---------------------------------------------------
  // One entry per ORDER, not per row — a 9-line order is one late order.
  try {
    var seenLate = {};
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (row.status !== Schema.status.PENDING) continue;
      if (!row.so || seenLate[row.so]) continue;

      var recv = receivedMap[row.so];
      if (!recv) continue;                       // no RECEIVED event → can't age it
      var ageMin = Math.floor((now - recv) / 60000);
      if (ageMin < WATCHDOG.redlineMinutes) continue;

      seenLate[row.so] = true;
      out.redline.push({
        key:     "rl:" + row.so,
        orderId: row.so,
        channel: row.inDirect ? "Direct" : "eBay",
        ageMin:  ageMin,
        sku:     row.sku,
        loc:     row.loc
      });
    }
    out.redline.sort(function (a, b) { return b.ageMin - a.ageMin; });
  } catch (e) {
    console.log("Watchdog: redline detector failed — " + e);
  }

  // ---- (2) PART-SHIPPED > 24h -------------------------------------------------
  try {
    var bySo = {};
    for (var k = 0; k < rows.length; k++) {
      var rw = rows[k];
      if (!rw.so) continue;
      if (!bySo[rw.so]) bySo[rw.so] = { shipped: 0, open: 0 };
      if (rw.status === Schema.status.SHIPPED) bySo[rw.so].shipped++;
      else if (rw.status === Schema.status.PENDING || rw.status === Schema.status.PREPARING) {
        bySo[rw.so].open++;
      }
    }
    Object.keys(bySo).forEach(function (so) {
      var g = bySo[so];
      if (g.shipped < 1 || g.open < 1) return;           // not a split order
      var lastShip = lastShipMap[so];
      if (!lastShip) return;                             // no timestamp → don't guess
      var ageMin = Math.floor((now - lastShip) / 60000);
      if (ageMin < WATCHDOG.partShippedHours * 60) return;
      out.partShipped.push({
        key:     "ps:" + so,
        so:      so,
        shipped: g.shipped,
        open:    g.open,
        ageMin:  ageMin
      });
    });
    out.partShipped.sort(function (a, b) { return b.ageMin - a.ageMin; });
  } catch (e) {
    console.log("Watchdog: part-shipped detector failed — " + e);
  }

  // ---- (3) KIT AWAITING A DECISION -------------------------------------------
  // A DIRECT row whose SKU is a registered kit, still open, that is not itself
  // an expansion component, and whose next row does NOT carry this kit's
  // "↳ from KIT-<sku>" tag. Same already-expanded test KitExpansion uses.
  try {
    if (kitMap && kitMap.size) {
      for (var m = 0; m < rows.length; m++) {
        var kr = rows[m];
        if (!kr.inDirect) continue;
        if (kr.status !== Schema.status.PENDING && kr.status !== Schema.status.PREPARING) continue;
        if (kr.note.indexOf("↳ from KIT-") === 0) continue;      // it IS a component
        var kit = kitMap.get(kr.sku);
        if (!kit) continue;

        var next = rows[m + 1];
        if (next && next.note.indexOf("↳ from KIT-" + kr.sku) === 0) continue;  // expanded

        out.kits.push({
          key:  "kx:" + kr.so + ":" + kr.sku,
          so:   kr.so,
          sku:  kr.sku,
          loc:  kr.loc || kit.location || "",
          name: kit.name || "",
          type: kit.type || "MANUAL"
        });
      }
    }
  } catch (e) {
    console.log("Watchdog: kit detector failed — " + e);
  }

  return out;
}


// =======================================================================================
// MESSAGE
// =======================================================================================

/**
 * Compose the alert. PLAIN TEXT, no parse_mode — same robustness choice as the
 * weekly digest and _sendKitParseAlert: SKUs, names and emoji can never break
 * a Markdown/HTML parse and silently drop the whole message.
 */
function _wdBuildText(items, total) {
  var L = ["⏳ WATCHDOG · " + total + " new", ""];

  if (items.redline.length) {
    L.push("PAST THE 3H LINE (" + items.redline.length + ")");
    _wdEach(items.redline, L, function (o) {
      return "  " + o.orderId + " · " + o.channel + " · " + _wdAge(o.ageMin) +
             "\n      " + o.sku + (o.loc ? "  " + o.loc : "");
    });
    L.push("");
  }

  if (items.partShipped.length) {
    L.push("PART-SHIPPED >" + WATCHDOG.partShippedHours + "H (" + items.partShipped.length + ")");
    _wdEach(items.partShipped, L, function (o) {
      return "  " + o.so + " · " + o.shipped + " shipped, " + o.open +
             " still open · last ship " + _wdAge(o.ageMin) + " ago";
    });
    L.push("");
  }

  if (items.kits.length) {
    L.push("KIT NEEDS A DECISION (" + items.kits.length + ")");
    _wdEach(items.kits, L, function (o) {
      var head = "  " + o.so + " · " + o.sku + (o.loc ? "  " + o.loc : "") + " · " + o.type;
      return o.name ? head + "\n      " + _wdClip(o.name, 48) : head;
    });
    L.push("");
  }

  return L.join("\n").replace(/\n+$/, "");
}

/** Push up to maxListed formatted entries, then summarise the remainder. */
function _wdEach(list, lines, fmt) {
  var shown = Math.min(list.length, WATCHDOG.maxListed);
  for (var i = 0; i < shown; i++) lines.push(fmt(list[i]));
  if (list.length > shown) lines.push("  … +" + (list.length - shown) + " more");
}

/** 195 → "3h 15m"; 2950 → "2d 1h". */
function _wdAge(min) {
  if (min < 60) return min + "m";
  var h = Math.floor(min / 60);
  if (h < 24) {
    var rm = min % 60;
    return rm ? h + "h " + rm + "m" : h + "h";
  }
  var d = Math.floor(h / 24);
  var rh = h % 24;
  return rh ? d + "d " + rh + "h" : d + "d";
}

function _wdClip(s, n) {
  s = String(s || "").trim();
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}


// =======================================================================================
// "ALREADY TOLD" STORE
// =======================================================================================

/**
 * @returns {Object|null} key → ms-recorded map, or NULL when the property has
 *          never been written (the cold-start signal — deliberately distinct
 *          from an empty object, which means "armed, nothing outstanding").
 */
function _wdLoadAlerted() {
  var raw = PropertiesService.getScriptProperties().getProperty(WATCHDOG.alertedPropKey);
  if (raw === null || raw === undefined || raw === "") return null;
  try {
    var parsed = JSON.parse(raw);
    return (parsed && typeof parsed === "object") ? parsed : {};
  } catch (e) {
    console.log("Watchdog: alerted store unparseable, treating as empty — " + e);
    return {};
  }
}

/** Merge new keys into the store, prune expired ones, write back. */
function _wdRecord(newKeys, known) {
  var now = Date.now();
  var cutoff = now - (WATCHDOG.retentionDays * 24 * 60 * 60 * 1000);
  var next = {};

  Object.keys(known || {}).forEach(function (k) {
    var t = known[k];
    if (typeof t === "number" && t >= cutoff) next[k] = t;
  });
  (newKeys || []).forEach(function (k) { next[k] = now; });

  try {
    PropertiesService.getScriptProperties()
      .setProperty(WATCHDOG.alertedPropKey, JSON.stringify(next));
  } catch (e) {
    // Script Properties cap a value at ~9KB. If we ever get there, drop the
    // oldest half rather than losing the store entirely (which would re-alert
    // everything on the next run).
    console.log("Watchdog: store write failed (" + e + ") — retrying trimmed");
    try {
      var keys = Object.keys(next).sort(function (a, b) { return next[b] - next[a]; });
      var trimmed = {};
      keys.slice(0, Math.floor(keys.length / 2)).forEach(function (k) { trimmed[k] = next[k]; });
      PropertiesService.getScriptProperties()
        .setProperty(WATCHDOG.alertedPropKey, JSON.stringify(trimmed));
    } catch (e2) {
      console.log("Watchdog: trimmed store write also failed — " + e2);
    }
  }
}
