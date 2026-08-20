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
function runStragglerWatchdog(maps) {
  try {
    var found = _gatherStragglers(maps);
    var store = _wdLoadAlerted();
    var isColdStart = (store === null);
    var known = isColdStart ? {} : store;

    // Split findings into already-told vs new.
    var fresh = { redline: [], partShipped: [], kits: [], overListed: [], critical: [] };
    var allKeys = [];
    ["redline", "partShipped", "kits", "overListed", "critical"].forEach(function (bucket) {
      found[bucket].forEach(function (item) {
        allKeys.push(item.key);
        if (!known[item.key]) fresh[bucket].push(item);
      });
    });

    var totalFound = allKeys.length;
    var totalFresh = fresh.redline.length + fresh.partShipped.length + fresh.kits.length
                   + fresh.overListed.length + fresh.critical.length;

    // ---- COLD START: seed silently, announce the counts only -------------------
    if (isColdStart) {
      var armed = "🛡 WATCHDOG ARMED\n\n" +
                  "Now watching: orders past the 3h line, SOs part-shipped over " +
                  WATCHDOG.partShippedHours + "h, kits sitting unexpanded in DIRECT, and " +
                  "MANUAL kits listed for more than their parts can build, and parts that too many kits depend on running thin.\n\n" +
                  "Found right now (seeded, not alerting on these):\n" +
                  "  " + found.redline.length + " past the 3h line\n" +
                  "  " + found.partShipped.length + " part-shipped >" + WATCHDOG.partShippedHours + "h\n" +
                  "  " + found.kits.length + " kit(s) awaiting a decision\n" +
                  "  " + found.overListed.length + " kit(s) listed beyond what we can build\n" +
                  "  " + found.critical.length + " load-bearing part(s) running thin\n\n" +
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
      ["redline", "partShipped", "kits", "overListed", "critical"].forEach(function (b) {
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
  var total = found.redline.length + found.partShipped.length + found.kits.length
            + found.overListed.length;
  if (total === 0) return "✅ Nothing late — no stragglers, no part-shipped SOs, no kits waiting, nothing over-listed.";
  return _wdBuildText(found, total) +
         "\n\n(preview — shows everything currently late, not just new)";
}


/** Counts only, for a possible sidebar badge later. Never throws. */
function getStragglerCounts() {
  try {
    var f = _gatherStragglers();
    return { critical: f.critical.length, redline: f.redline.length, partShipped: f.partShipped.length,
             kits: f.kits.length, overListed: f.overListed.length };
  } catch (e) {
    return { redline: 0, partShipped: 0, kits: 0, overListed: 0, critical: 0 };
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
function _gatherStragglers(maps) {
  var out = { redline: [], partShipped: [], kits: [], overListed: [], critical: [] };
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

  // ---- 4 · OVER-LISTED KITS — advertising more than the components can build ---
  //
  // The picker's own words (2026-08-19): "when every piston is sold or repair kit,
  // I go and check if I need to adjust the stock for this kit." That check is
  // manual and easy to forget, and forgetting it means eBay keeps offering a kit
  // nobody can assemble. This does the check every hour instead.
  //
  // ⚠ MANUAL KITS ONLY. A READY kit's quantity is a pre-assembled box on a K-*
  //   shelf — real inventory. "Can't build more" there is a resupply fact, not a
  //   promise we can't keep. Alarming on it would bury the real cases.
  // ⚠ ADVERTISED = max(eBay, Zoho) — either channel can take the order.
  // ⚠ Bundled components are already excluded by _oosComputeKitBuild, so a head
  //   gasket that ships inside the gasket set can't raise a false alarm here.
  // ⚠ Reuses the MI maps the hourly pass already built; falls back to its own read
  //   when called standalone from the sidebar preview.
  try {
    var olMaps = (maps && maps.inventoryMap) ? maps : buildLocationAndInventoryMaps();
    var olZoho = buildZohoStockMap();
    var olKits = buildKitMap();
    var olAvail = _oosResolveAvailFactory(olZoho, olMaps.inventoryMap);

    olKits.forEach(function (kit, skuLower) {
      if (String(kit.type || "MANUAL").toUpperCase() === "READY") return;

      var build = _oosComputeKitBuild(kit, olAvail);
      if (typeof build.buildable !== 'number') return;      // ⚠ untrustable — say nothing

      var zo = olZoho.get(skuLower);
      var mi = olMaps.inventoryMap.get(skuLower);
      var advertised = Math.max(
        (zo && zo.available != null) ? zo.available : 0,
        (mi && mi.available != null) ? mi.available : 0
      );
      if (advertised <= 0 || advertised <= build.buildable) return;

      out.overListed.push({
        key:   "ol:" + kit.sku,
        sku:   kit.sku,
        name:  kit.name || "",
        adv:   advertised,
        build: build.buildable,
        gap:   advertised - build.buildable,
        limitedBy: build.limitedBy || "",
        limiterName: (build.limiter && build.limiter.name) || ""
      });
    });

    // Worst first: nothing buildable at all before partly-covered, then by size.
    out.overListed.sort(function (a, b) {
      if ((a.build === 0) !== (b.build === 0)) return a.build === 0 ? -1 : 1;
      return b.gap - a.gap;
    });
  } catch (e) {
    console.log("Watchdog: over-listed detector failed — " + e);
  }

  /* ---- LOAD-BEARING PARTS RUNNING THIN --------------------------------------
     ⚠⚠ THE ONLY PREVENTIVE ITEM THE WATCHDOG CARRIES. Every other bucket here
     reports something that has ALREADY gone wrong — an order past the line, an SO
     part-shipped, a kit sitting unexpanded. This one fires BEFORE the damage: a
     component that N kits depend on is down to its last couple of builds.

     The ripple cannot raise it, by construction — it returns early on any kit that
     is not already blocked, so a part holding up twelve HEALTHY kits is invisible
     to it right until the morning all twelve stop at once.

     ⚠ RIDES THE MAPS THE PASS ALREADY HOLDS. analyzeKitCriticality costs a kit-map
     walk; handing it the caller's maps keeps this to little more than that. Best
     effort — a failure here must never cost the four buckets above it. */
  try {
    if (typeof analyzeKitCriticality === 'function') {
      var crit = analyzeKitCriticality(maps);
      (crit.atRisk || []).forEach(function (p) {
        out.critical.push({
          /* ⚠ KEYED ON THE SKU ALONE, deliberately. Stock oscillates, so keying on
             the level would re-alert on every wobble across the threshold. One
             alert per part, then silence until the 14-day prune — the same
             once-per-crossing contract every other bucket uses. The trade-off is
             stated rather than hidden: a part that dips, is restocked, and dips
             again inside a fortnight will not alert twice. Flapping is worse. */
          key:   "kc:" + String(p.sku).toLowerCase(),
          sku:   String(p.sku),
          name:  String(p.name || ""),
          loc:   String(p.location || ""),
          avail: p.avail,
          kits:  p.kits,
          cover: p.coverKits
        });
      });
    }
  } catch (e) {
    try { console.log("watchdog.critical: " + e); } catch (_) {}
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

  /* ⚠ THIS BUCKET LEADS. Everything below it reports damage already done;
     this is the only one you can still act on before it costs anything. */
  if (items.critical.length) {
    L.push("🏗 LOAD-BEARING AND RUNNING THIN (" + items.critical.length + ")");
    _wdEach(items.critical, L, function (o) {
      return "  " + o.sku + (o.loc ? "  " + o.loc : "") + " · have " + o.avail +
             "\n      " + o.kits + " kits depend on it · covers " + o.cover + " more" +
             (o.name ? "\n      " + _wdClip(o.name, 44) : "");
    });
    L.push("");
  }

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

  if (items.overListed.length) {
    var cant = 0;
    for (var q = 0; q < items.overListed.length; q++) if (items.overListed[q].build === 0) cant++;
    L.push("LISTED BUT NOT BUILDABLE (" + items.overListed.length +
           (cant ? " · " + cant + " can't build at all" : "") + ")");
    _wdEach(items.overListed, L, function (o) {
      var head = "  " + o.sku + " · listed " + o.adv + ", can build " + o.build +
                 (o.build === 0 ? "  ← next sale fails" : "");
      var why  = o.limiterName ? (_wdClip(o.limiterName, 30) + " · " + o.limitedBy) : o.limitedBy;
      return head + (o.name ? "\n      " + _wdClip(o.name, 46) : "") +
             (why ? "\n      short: " + _wdClip(why, 56) : "");
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


// =======================================================================================
// ⭐ THE PUBLISHED-CELL PULSE — is the board's FREE path still free? (2026-08-18)
// =======================================================================================
/**
 * WHY THIS EXISTS, AND WHY IT IS THE ONE ALARM WORTH HAVING.
 *
 * A board tick is normally answered by n8n out of the published cell — tier 2 —
 * and Apps Script is never involved. THAT is why monitors and wall screens cost
 * nothing and why the tenth viewer costs what the first did.
 *
 * When that path fails, every poll falls through to Apps Script rebuilding the
 * tick (~3.5s). Do the arithmetic on a four-screen floor:
 *
 *     4 devices x 3 polls/min x 3.5s = 42 SECONDS of script time per minute
 *
 * — about 70% of the script consumed by people merely LOOKING at screens,
 * permanently, and it gets worse with every screen added. It would starve the
 * floor's writes exactly the way 2026-08-17 did, but with nothing on the floor
 * having changed to explain it.
 *
 * ⚠⚠ AND IT FAILS SILENTLY. Nothing errors. n8n's tier-2 block is wrapped so a
 * failed read simply falls through. It ran BROKEN FOR A WEEK in August 2026 and
 * was found only because pickers said the board "felt slow". That combination —
 * invisible, cheap to detect, expensive to miss — is what earns an alarm.
 *
 * ⚠ TWO DIFFERENT FAULTS, AND THEY NEED DIFFERENT FIXES:
 *   READER broken — n8n cannot read the cell (the August case was a credential
 *                   domain allowlist). Tick arrives with _liveFallback set.
 *   WRITER broken — n8n reads fine but the cell is ancient, i.e. Apps Script
 *                   stopped publishing (a dead trigger). Tick says _published
 *                   but _publishedAgeMs is far past keep-fresh.
 * Reporting them as one "tier 2 is down" would send someone to the wrong system.
 *
 * ⚠ IT PROBES THE PUBLIC URL so it exercises Caddy + n8n + the tier logic, the
 * same path a tablet takes. It never re-enters Apps Script — n8n reads the cell
 * through the Sheets API — so there is no recursion and no self-measurement.
 *
 * ⚠ ALERTS ON THE CROSSING, IN BOTH DIRECTIONS. One message when it breaks, one
 * when it recovers, and silence in between. A monitor that repeats itself every
 * hour is one people mute — the same discipline as the straggler watchdog.
 */
var PULSE = {
  propKey:       "PULSE_LAST_STATE",
  maxTrayAgeMin: 15,      // keep-fresh republishes by 8 min; 15 is unambiguously wrong
  timeoutMs:     20000
};


/**
 * Probe the board's public endpoint and classify what answered.
 * Pure-ish: no sends, no state writes. Safe to call any time.
 *
 * @returns {{state: string, detail: string}} state is one of
 *          ok | reader | writer | unreachable
 */
function _pulseProbe() {
  var res;
  try {
    res = UrlFetchApp.fetch(HQ_BOARD_API_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ action: "boardTick" }),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
  } catch (e) {
    return { state: "unreachable", detail: "the request failed: " + e };
  }

  var code = res.getResponseCode();
  if (code !== 200) return { state: "unreachable", detail: "HTTP " + code };

  var tick;
  try {
    tick = JSON.parse(res.getContentText());
  } catch (e) {
    return { state: "unreachable", detail: "the answer was not JSON" };
  }

  // ⚠ A tick with no cockpit is not a tick. The proxy answers HTTP 200 even
  // when the hop failed, so shape is the only honest test — the same lesson
  // the board itself learned on 2026-08-05 and again on 2026-08-17.
  if (!tick || typeof tick !== "object" || !tick.cockpit) {
    return { state: "unreachable", detail: "a 200 with no tick in it" };
  }

  if (!tick._published) {
    return {
      state: "reader",
      detail: tick._liveFallback
        ? "Apps Script rebuilt this tick (tier 3)"
        : "the published cell did not answer" +
          (tick._tier2 ? " — n8n says: " + tick._tier2 : "")
    };
  }

  var ageMin = Math.round((Number(tick._publishedAgeMs) || 0) / 60000);
  if (ageMin > PULSE.maxTrayAgeMin) {
    return {
      state: "writer",
      detail: "the cell is being READ fine but was last written " + ageMin +
              " min ago (keep-fresh should republish within 8)"
    };
  }

  return { state: "ok", detail: "served from the published cell, " + ageMin + " min old" };
}


/**
 * The hourly check. Alerts only when the state CHANGES.
 * Never throws — it runs inside the housekeeping chain.
 *
 * @returns {string} one-line summary for the housekeeping log
 */
function checkPublishedPulse() {
  try {
    var probe = _pulseProbe();
    var props = PropertiesService.getScriptProperties();
    var last  = props.getProperty(PULSE.propKey);   // null on a cold start

    // Nothing changed — say nothing. This is the case almost every hour.
    if (last === probe.state) return "💓 Pulse: " + probe.state + " (unchanged)";

    var msg = null;
    if (last === null) {
      // Cold start: record silently unless we are ALREADY broken, which is
      // worth hearing about immediately.
      if (probe.state === "ok") {
        props.setProperty(PULSE.propKey, probe.state);
        return "💓 Pulse armed (healthy)";
      }
      msg = _pulseText(probe, true);
    } else {
      msg = _pulseText(probe, false);
    }

    var sent = _tgSend(TELEGRAM_ADMIN_CHAT_ID, msg);
    // Record ONLY on a successful send, so a failed send retries next hour
    // rather than silently swallowing the transition.
    if (sent) props.setProperty(PULSE.propKey, probe.state);
    return "💓 Pulse: " + (last || "cold") + " → " + probe.state +
           (sent ? " (alerted)" : " (SEND FAILED, will retry)");

  } catch (err) {
    try { console.log("checkPublishedPulse failed: " + err); } catch (_) {}
    return "❌ Pulse: " + err;
  }
}


/**
 * The message. Written so the reader knows WHAT broke, WHY it matters and WHERE
 * to look — an alert that only says "something is wrong" costs a person twenty
 * minutes finding out what.
 */
function _pulseText(probe, coldStart) {
  if (probe.state === "ok") {
    return "🟢 BOARD BACK ON THE PUBLISHED CELL\n\n" +
           probe.detail + ".\n\n" +
           "Screens are free again — Apps Script is no longer rebuilding the " +
           "tick for every poll.";
  }

  var head, why, look;
  if (probe.state === "reader") {
    head = "🔴 THE BOARD'S FREE PATH IS DOWN";
    why  = "Every board poll is now making Apps Script rebuild the tick (~3.5s " +
           "each). With 4 screens that is roughly 70% of the script spent on " +
           "people just LOOKING — it will starve the pickers' taps.";
    look = "LOOK AT: the n8n hq-board workflow, node 2 (the Sheets read) and " +
           "its Google credential. In Aug 2026 this was the credential's " +
           "Allowed Domains — it wants a BARE hostname (sheets.googleapis.com), " +
           "not a URL.";
  } else if (probe.state === "writer") {
    head = "🟠 THE PUBLISHED CELL HAS GONE STALE";
    why  = "n8n can still READ it, so the board is fast — but it is serving an " +
           "old picture, so the floor may be looking at orders that have moved.";
    look = "LOOK AT: Apps Script → Triggers → runPublishTick. It should run " +
           "every minute. Also check __Published!A3 for the last publish error.";
  } else {
    head = "🔴 THE BOARD ENDPOINT DID NOT ANSWER";
    why  = "A tablet asking for the pick list right now would get nothing.";
    look = "LOOK AT: Caddy and the n8n container on the VPS, then the hq-board " +
           "workflow is Active.";
  }

  return head + "\n\n" +
         (coldStart ? "(found already broken when the pulse check first ran)\n\n" : "") +
         probe.detail + ".\n\n" +
         why + "\n\n" + look + "\n\n" +
         "You'll get one message when this clears. Nothing more until then.";
}


/** Eyeball path — probe and show, send nothing, write nothing. */
function previewPublishedPulse() {
  var p = _pulseProbe();
  var out = "state: " + p.state + "\n" + p.detail + "\n\n--- the message it would send ---\n" +
            (p.state === "ok" ? "(nothing — healthy states are silent unless recovering)"
                              : _pulseText(p, false));
  // ⚠ BOTH loggers on purpose. The editor's inline Execution-log panel often
  // shows only start/complete for a fast run, and the Run button does not
  // display return values at all (the getPublishedTick lesson). Logger.log is
  // the one that reliably surfaces there.
  try { console.log(out); } catch (_) {}
  try { Logger.log(out); }  catch (_) {}
  return out;
}

/** Forget the last state, so the next run re-announces whatever it finds. */
function resetPublishedPulse() {
  PropertiesService.getScriptProperties().deleteProperty(PULSE.propKey);
  return "Pulse state cleared — the next check will re-announce.";
}

/**
 * TEST PATH — prove the alert text and the crossing logic WITHOUT breaking
 * production. Deliberately does not touch tier 2: taking the board's fast path
 * down to test an alarm would put the floor on the slow path to prove a point.
 *
 * @param {string} state - ok | reader | writer | unreachable
 * @returns {string} the exact message that state would send
 */
function simulatePulseAlert(state) {
  var fake = {
    ok:          { state: "ok",          detail: "served from the published cell, 1 min old" },
    reader:      { state: "reader",      detail: "Apps Script rebuilt this tick (tier 3)" },
    writer:      { state: "writer",      detail: "the cell is being READ fine but was last written 41 min ago (keep-fresh should republish within 8)" },
    unreachable: { state: "unreachable", detail: "HTTP 502" }
  }[state || "reader"];
  var out = _pulseText(fake, false);
  try { console.log(out); } catch (_) {}
  return out;
}


/**
 * Send ONE test alert to the admin chat so the whole chain can be proven —
 * message composition → _tgSend → Telegram → your phone — without breaking
 * anything. Marked as a drill in the text so nobody mistakes it for real.
 *
 * ⚠ Does NOT touch the stored state, so it cannot suppress or trigger a real
 * transition afterwards.
 *
 * Run from the editor: sendTestPulseAlert('reader' | 'writer' | 'unreachable' | 'ok')
 */
function sendTestPulseAlert(state) {
  var body = simulatePulseAlert(state || "reader");
  var text = "🧪 DRILL — this is a TEST of the board pulse alarm.\n" +
             "Nothing is wrong. This is what you would receive:\n\n" +
             "─────────────\n" + body;
  var ok = _tgSend(TELEGRAM_ADMIN_CHAT_ID, text);
  var out = ok ? "Test alert sent to the admin chat."
               : "SEND FAILED — check TELEGRAM_ADMIN_CHAT_ID and the bot token.";
  try { console.log(out); } catch (_) {}
  return out;
}
