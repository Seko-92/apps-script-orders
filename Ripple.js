// =======================================================================================
// RIPPLE.gs — "which parts should I restock first?"  (shared-component ripple)
// =======================================================================================
//
// WHY THIS EXISTS
//   The 2026-08-05 KPI snapshot showed in-stock-BLOCKED kits jumping 37 → 46 in
//   five days while buildable fell 169 → 162. Components running out are taking
//   kits down with them. Kit Health already reports the 46; it does not tell you
//   which purchase order clears the most of them.
//
//   The parked roadmap idea was REACTIVE — "you restocked 195306, that unblocked
//   kits X and Y." Useful, but it arrives after the buying decision. This is the
//   FORWARD version: rank every blocking component by how many kits it frees, so
//   one order clears the maximum backlog.
//
// THE HONESTY RULE (the reason this is not a one-liner)
//   `_oosComputeKitBuild` returns ONE limiter — the single worst component. If a
//   kit is short on three parts, restocking the worst one alone does NOT make it
//   buildable. Counting "kits where I am the limiter" would therefore promise
//   unblocks that never happen.
//
//   So we compute EVERY short component per blocked kit and report two numbers:
//     · SOLE   — kits where this part is the ONLY thing short. Restocking it
//                alone genuinely makes those kits buildable. This is the number
//                the ranking is built on, because it is the one that's a promise.
//     · SHORT IN — kits where it's among the missing. Bigger, softer, useful
//                context for a bulk order.
//   And "restock the top N → M kits" is computed as a SET operation (a kit
//   counts only when ALL of its short parts are inside the chosen set), never as
//   a sum of per-part counts, which would double-count shared kits.
//
//   Kits whose composition is untrustworthy (unparsed PD lines, components not
//   in MI/Zoho) are EXCLUDED and counted separately — same rule as OOS,
//   Kit Health and the kit-expansion modal: never show a number we can't stand
//   behind.
//
// PURE TIER 1 — no new data, no new dependency. It is a lens over three things
//   already owned: buildKitMap() · buildZohoStockMap() + MI (Zoho-first
//   availability, the established kit routing) · _oosComputeKitBuild() for the
//   authoritative buildable verdict, so this can never disagree with the OOS
//   sheet or the Kit Health cockpit.
//
// PUBLIC API
//   analyzeRestockRipple()      — structured data (parts, blocked kits, totals)
//   previewRestockRipple()      — the plain-text report (sidebar / editor / TG)
//   getRippleTop(n)             — top-n parts only, for the weekly report later
// =======================================================================================


var RIPPLE = {
  topParts:        8,   // parts listed in the report
  kitsPerPart:     3,   // example kit SKUs shown per part before "+N"
  planSize:        5,   // "restock these N →" headline set size
  unreadableLimit: 10   // unassessable kits NAMED before summarising
};


// =======================================================================================
// ENGINE
// =======================================================================================

/**
 * Work out which components are blocking which kits, and what restocking each
 * would actually free.
 *
 * @returns {{
 *   parts: Array<{sku,name,location,avail,sole:number,shortIn:number,soleKits:Array,allKits:Array}>,
 *   blockedCount: number,      // kits with buildable 0 (trustworthy ones only)
 *   multiBlocked: number,      // blocked kits that need >1 part restocked
 *   skipped: number,           // kits excluded as untrustworthy
 *   totalKits: number,
 *   planParts: Array<string>,  // the top-planSize part SKUs
 *   planFrees: number          // kits made buildable if ALL planParts are restocked
 * }}
 */
function analyzeRestockRipple(sharedMaps, sharedZoho) {
  /* ⚠⚠ SHAPE-DETECTED, NEVER TRUSTED BY POSITION — the same guard recomputeHand
     and refreshOutOfStock use, and for the same reason: a TIME TRIGGER hands its
     target an EVENT OBJECT as the first argument, so a positional `maps` becomes
     that event the day anyone schedules this. Nothing schedules it today. The
     guard costs two comparisons and removes the whole class.

     ⚠ WHY SHARING MATTERS HERE AND NOT ELSEWHERE: this function costs ~6s and
     ~4.2s of that is map-building (measured 2026-08-19) — buildLocationAndInventoryMaps
     alone is the ~600,000-cell Master Inventory read. A caller that already holds
     them (the hourly housekeeping pass) hands them over and pays for the kit map
     only. Callers with nothing in hand pass nothing and behave exactly as before. */
  var maps = (sharedMaps && sharedMaps.inventoryMap && sharedMaps.locationMap)
               ? sharedMaps : buildLocationAndInventoryMaps();
  var zohoMap = (sharedZoho && typeof sharedZoho.get === 'function')
               ? sharedZoho : buildZohoStockMap();
  var kitMap = buildKitMap();
  var resolveAvail = _oosResolveAvailFactory(zohoMap, maps.inventoryMap);
  var locationMap = maps.locationMap;

  var blocked = [];    // {kitSku, kitName, short: [compSkuLower]}
  var partInfo = {};   // compSkuLower → {sku, name, avail}
  var unreadable = []; // {sku, name, reason, raw} — kits nothing can assess
  var skipped = 0;
  var totalKits = 0;

  kitMap.forEach(function (kit) {
    totalKits++;

    // Authoritative verdict — same call the OOS sheet and Kit Health make, so
    // the three surfaces can never disagree about whether a kit is blocked.
    var build = _oosComputeKitBuild(kit, resolveAvail);
    if (build.buildable === "⚠") {
      // NAME them. A bare count isn't actionable — and these kits are a blind
      // spot in EVERY surface: Kit Health can't price them, OOS can't compute
      // buildable, this can't rank them. `limitedBy` already carries which of
      // the three causes it was (unreadable PD / component not found / no
      // components), and each needs a different fix in Zoho.
      skipped++;
      unreadable.push({
        sku:    String(kit.sku),
        name:   String(kit.name || ""),
        reason: String(build.limitedBy || "unknown").replace(/^⚠\s*/, ""),
        raw:    (kit.unparsedLines && kit.unparsedLines.length)
                  ? String(kit.unparsedLines[0]) : ""
      });
      return;
    }
    if (build.buildable > 0) return;              // not blocked — nothing to free

    // buildable === 0 → find EVERY component that can't cover one full kit.
    var comps = kit.components || [];
    var short = [];
    for (var i = 0; i < comps.length; i++) {
      var qtyPer = (comps[i].qty > 0) ? comps[i].qty : 1;
      var key = String(comps[i].sku).trim().toLowerCase();
      var avail = resolveAvail(key);
      if (avail === null) continue;               // can't happen post-⚠ gate; guard anyway
      if (Math.floor(avail / qtyPer) >= 1) continue;

      short.push(key);
      if (!partInfo[key]) {
        partInfo[key] = {
          sku:      String(comps[i].sku).trim(),
          name:     String(comps[i].name || ""),
          avail:    Math.round(avail * 100) / 100,
          location: locationMap.get(key) || ""
        };
      }
    }
    if (!short.length) return;                    // defensive: 0-buildable with nothing short

    blocked.push({
      kitSku:  String(kit.sku),
      kitName: String(kit.name || ""),
      short:   short
    });
  });

  // ---- Per-part tallies -------------------------------------------------------
  var tally = {};   // key → {sole, shortIn, soleKits[], allKits[]}
  blocked.forEach(function (b) {
    var isSolo = (b.short.length === 1);
    b.short.forEach(function (key) {
      if (!tally[key]) tally[key] = { sole: 0, shortIn: 0, soleKits: [], allKits: [] };
      tally[key].shortIn++;
      tally[key].allKits.push(b.kitSku);
      if (isSolo) {
        tally[key].sole++;
        tally[key].soleKits.push(b.kitSku);
      }
    });
  });

  var parts = Object.keys(tally).map(function (key) {
    var t = tally[key];
    var info = partInfo[key];
    return {
      key:      key,
      sku:      info.sku,
      name:     info.name,
      location: info.location,
      avail:    info.avail,
      sole:     t.sole,
      shortIn:  t.shortIn,
      soleKits: t.soleKits,
      allKits:  t.allKits
    };
  });

  // Rank by the number we can PROMISE (sole), then by breadth, then by
  // emptiest shelf — a part at 0 is more urgent than one at 2 when tied.
  parts.sort(function (a, b) {
    if (b.sole !== a.sole)       return b.sole - a.sole;
    if (b.shortIn !== a.shortIn) return b.shortIn - a.shortIn;
    return a.avail - b.avail;
  });

  // ---- "Restock these N →" as a SET operation, never a sum -------------------
  var planParts = parts.slice(0, RIPPLE.planSize).map(function (p) { return p.key; });
  var planSet = {};
  planParts.forEach(function (k) { planSet[k] = true; });
  var planFrees = 0;
  blocked.forEach(function (b) {
    for (var i = 0; i < b.short.length; i++) {
      if (!planSet[b.short[i]]) return;           // one part outside the set → still blocked
    }
    planFrees++;
  });

  var multiBlocked = blocked.filter(function (b) { return b.short.length > 1; }).length;

  unreadable.sort(function (a, b) { return String(a.sku).localeCompare(String(b.sku)); });

  return {
    parts:        parts,
    blockedCount: blocked.length,
    multiBlocked: multiBlocked,
    skipped:      skipped,
    unreadable:   unreadable,
    totalKits:    totalKits,
    planParts:    parts.slice(0, RIPPLE.planSize).map(function (p) { return p.sku; }),
    planFrees:    planFrees
  };
}


/** Top-n parts only — for the weekly report / digest to reuse without re-formatting. */
function getRippleTop(n, sharedMaps, sharedZoho) {
  try {
    // Shared maps are passed straight through; see the note on analyzeRestockRipple.
    var d = analyzeRestockRipple(sharedMaps, sharedZoho);
    return d.parts.slice(0, n || RIPPLE.topParts);
  } catch (e) {
    console.log("getRippleTop failed: " + e);
    return [];
  }
}


// =======================================================================================
// REPORT
// =======================================================================================

/**
 * The plain-text ripple report. PLAIN TEXT, no parse_mode — same robustness
 * choice as the digest and the watchdog: part names carry punctuation that
 * would break a Markdown parse and silently drop the whole message.
 */
function previewRestockRipple() {
  var d;
  try { d = analyzeRestockRipple(); }
  catch (e) { return "⚠ Ripple failed: " + (e.message || e); }

  if (d.blockedCount === 0) {
    var clear = "✅ Nothing blocked — every kit with a readable composition can be built.";
    var tail = _rpUnreadableBlock(d);
    return tail.length ? clear + "\n\n" + tail.join("\n") : clear;
  }

  var L = [];
  L.push("🔗 RESTOCK RIPPLE");
  L.push(d.blockedCount + " kits blocked · " + d.parts.length + " parts are the cause");
  L.push("");
  L.push("BIGGEST UNBLOCKERS");

  var shown = Math.min(d.parts.length, RIPPLE.topParts);
  for (var i = 0; i < shown; i++) {
    var p = d.parts[i];
    L.push("  " + p.sku + (p.location ? "  " + p.location : "") + "   have " + p.avail);
    if (p.name) L.push("      " + _rpClip(p.name, 46));

    var line = "      ";
    if (p.sole > 0) {
      line += "frees " + p.sole + " kit" + (p.sole === 1 ? "" : "s") + " on its own";
      if (p.shortIn > p.sole) line += " · short in " + p.shortIn;
    } else {
      // Honest about the weaker case: restocking this alone frees nothing.
      line += "short in " + p.shortIn + " kit" + (p.shortIn === 1 ? "" : "s") +
              " · none freed by this part alone";
    }
    L.push(line);

    if (p.soleKits.length) {
      var ex = p.soleKits.slice(0, RIPPLE.kitsPerPart).join(", ");
      if (p.soleKits.length > RIPPLE.kitsPerPart) {
        ex += ", +" + (p.soleKits.length - RIPPLE.kitsPerPart);
      }
      L.push("      → " + ex);
    }
  }
  if (d.parts.length > shown) L.push("  … +" + (d.parts.length - shown) + " more parts");

  L.push("");
  if (d.planFrees > 0) {
    L.push("Restock these " + d.planParts.length + " → " + d.planFrees +
           " of " + d.blockedCount + " kits become buildable:");
    L.push("  " + d.planParts.join(" · "));
  } else {
    L.push("No single group of " + RIPPLE.planSize + " parts frees a kit outright —");
    L.push("every blocked kit is short on something outside the top " + RIPPLE.planSize + ".");
  }

  if (d.multiBlocked > 0) {
    L.push("");
    L.push(d.multiBlocked + " kit" + (d.multiBlocked === 1 ? " needs" : "s need") +
           " more than one part restocked.");
  }
  var un = _rpUnreadableBlock(d);
  if (un.length) {
    L.push("");
    L = L.concat(un);
  }

  return L.join("\n");
}


/**
 * The "can't be assessed" section — NAMED, with the specific reason per kit.
 *
 * These kits are a blind spot in every surface at once (Kit Health can't price
 * them, OOS can't compute buildable, the ripple can't rank them), so a bare
 * count leaves the user with nothing to act on. Each of the three causes needs
 * a different fix, which is why the reason is printed per kit rather than as a
 * single blanket sentence.
 *
 * @returns {Array<string>} lines, or [] when everything parsed cleanly
 */
function _rpUnreadableBlock(d) {
  var list = d.unreadable || [];
  if (!list.length) return [];

  var L = ["⚠ CAN'T BE ASSESSED (" + list.length + ")"];
  var shown = Math.min(list.length, RIPPLE.unreadableLimit);
  for (var i = 0; i < shown; i++) {
    var u = list[i];
    L.push("  " + u.sku + (u.name ? "  " + _rpClip(u.name, 38) : ""));
    L.push("      " + _rpClip(u.reason, 54));
    if (u.raw) L.push("      \"" + _rpClip(u.raw, 50) + "\"");
  }
  if (list.length > shown) L.push("  … +" + (list.length - shown) + " more");
  L.push("  Fix the Purchase Description in Zoho — the registry re-parses on save.");
  return L;
}


function _rpClip(s, n) {
  s = String(s || "").trim();
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}
