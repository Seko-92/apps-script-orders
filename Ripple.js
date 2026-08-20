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
function analyzeRestockRipple(sharedMaps, sharedZoho, sharedKitMap) {
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
  /* ⚠ THIRD SHARED MAP (2026-08-21). analyzeKitCriticality walks the same kit map,
     and the hourly pass now runs BOTH — so without this the pass builds it twice.
     Shape-detected like the other two, and a caller passing nothing is unchanged. */
  var kitMap = (sharedKitMap && typeof sharedKitMap.forEach === 'function'
                             && typeof sharedKitMap.get === 'function')
               ? sharedKitMap : buildKitMap();
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
function getRippleTop(n, sharedMaps, sharedZoho, sharedKitMap) {
  try {
    // Shared maps are passed straight through; see the note on analyzeRestockRipple.
    var d = analyzeRestockRipple(sharedMaps, sharedZoho, sharedKitMap);
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


// =======================================================================================
// KIT CRITICALITY — which components the kit catalogue LEANS ON, blocked or not
// =======================================================================================
//
// ⚠⚠ WHY THIS IS NOT THE RIPPLE. analyzeRestockRipple deliberately returns early on
// any kit that is not ALREADY blocked:
//
//     if (build.buildable > 0) return;      // not blocked — nothing to free
//
// which is exactly right for its question ("what do I restock to unblock the most").
// But it means a component that TWELVE kits depend on, and that happens to be in
// stock today, is invisible — right up until the morning it hits zero and twelve
// kits go down at once. That is a single point of failure with no surface.
//
// This is the other lens: rank every component by HOW MANY KITS DEPEND ON IT,
// regardless of whether anything is short right now. The ripple is reactive
// (what is broken); this is structural (what would break).
//
// ⚠ ONLY THIS SYSTEM CAN ANSWER IT. eBay and Zoho both report sales volume well —
// that is the channel/financial view. Neither holds the kit composition map, so
// neither can say "if this part dies, N kits die with it". That is why the raw
// top-sellers-by-volume report was NOT built alongside this one: it would have
// duplicated two systems that already do it, and /lowstock already carries soldQty.
//
// ⚠ UNTRUSTWORTHY KITS ARE EXCLUDED AND COUNTED, never silently folded in — the
// same honesty rule the OOS sheet, Kit Health and the ripple all follow. A kit whose
// PD will not parse cannot tell us what it depends on, so counting it would understate
// every part it contains.

var KIT_CRITICAL = {
  topN:          12,   // how many parts the text surface lists
  maxKitsListed:  6,   // sample kit SKUs shown per part before "+N"
  // Fan-out at which a part stops being "a part" and becomes infrastructure.
  // Below this, losing it is a normal restock; at or above it, one empty shelf
  // takes a whole shelf of products off the market at once.
  minKits:        5,
  // Alert when it can no longer cover this many complete kits of its hungriest
  // recipe. 2 rather than 1 on purpose: firing at 1 fires when it is ALREADY the
  // last one, which is a report, not a warning.
  warnCoverKits:  2
};

/**
 * Rank components by how much of the kit catalogue rests on them.
 *
 * @returns {{
 *   parts: Array<{sku,name,location,avail,kits,maxQtyPer,cannotCover,coverKits,kitSkus}>,
 *   atRisk: Array<Object>,   // fan-out >= minKits AND running thin — the alert set
 *   totalKits: number,
 *   totalParts: number,
 *   skipped: number          // kits excluded as untrustworthy
 * }}
 */
function analyzeKitCriticality(sharedMaps, sharedZoho, sharedKitMap) {
  var maps = (sharedMaps && sharedMaps.inventoryMap && sharedMaps.locationMap)
               ? sharedMaps : buildLocationAndInventoryMaps();
  var zohoMap = (sharedZoho && typeof sharedZoho.get === 'function')
               ? sharedZoho : buildZohoStockMap();
  var kitMap = (sharedKitMap && typeof sharedKitMap.forEach === 'function'
                             && typeof sharedKitMap.get === 'function')
               ? sharedKitMap : buildKitMap();
  var resolveAvail = _oosResolveAvailFactory(zohoMap, maps.inventoryMap);
  var locationMap = maps.locationMap;

  var tally = {};
  var totalKits = 0, skipped = 0;

  kitMap.forEach(function (kit) {
    totalKits++;

    // Same authoritative verdict the OOS sheet and Kit Health make, so the
    // surfaces can never disagree about which kits are assessable at all.
    var build = _oosComputeKitBuild(kit, resolveAvail);
    if (build.buildable === "⚠") { skipped++; return; }

    var comps = kit.components || [];
    for (var i = 0; i < comps.length; i++) {
      var qtyPer = (comps[i].qty > 0) ? comps[i].qty : 1;
      var key = String(comps[i].sku).trim().toLowerCase();
      var avail = resolveAvail(key);
      if (avail === null) continue;              // unknown to both sources

      if (!tally[key]) {
        tally[key] = {
          key: key, sku: String(comps[i].sku).trim(),
          name: String(comps[i].name || ""),
          location: locationMap.get(key) || "",
          avail: Math.round(avail * 100) / 100,
          kits: 0, maxQtyPer: 0, cannotCover: 0, kitSkus: []
        };
      }
      var t = tally[key];
      t.kits++;
      t.kitSkus.push(String(kit.sku));
      if (qtyPer > t.maxQtyPer) t.maxQtyPer = qtyPer;
      // Already short for THIS recipe — it is not a future risk, it is current.
      if (Math.floor(avail / qtyPer) < 1) t.cannotCover++;
    }
  });

  var parts = Object.keys(tally).map(function (k) {
    var t = tally[k];
    /* ⚠ DEPTH IS MEASURED AGAINST THE HUNGRIEST RECIPE, not an average. A part
       used 1-per-kit in ten kits and 6-per-kit in one is limited by the 6 — and
       the conservative number is the one worth alarming on. */
    t.coverKits = (t.maxQtyPer > 0) ? Math.floor(t.avail / t.maxQtyPer) : 0;
    return t;
  });

  /* ⚠ RANKED BY FAN-OUT, because that is the question being asked — "what does the
     catalogue lean on most". Deliberately NOT a blended risk score: this project
     already learned (Kit Health's calibrated discount) that a composite number
     hides the signal that produced it. Fan-out ranks, cover is shown beside it,
     and the human weighs them. Ties break toward the one already hurting, then
     toward the emptier shelf. */
  parts.sort(function (a, b) {
    if (b.kits !== a.kits)               return b.kits - a.kits;
    if (b.cannotCover !== a.cannotCover) return b.cannotCover - a.cannotCover;
    return a.avail - b.avail;
  });

  var atRisk = parts.filter(function (p) {
    return p.kits >= KIT_CRITICAL.minKits && p.coverKits < KIT_CRITICAL.warnCoverKits;
  });

  return {
    parts: parts, atRisk: atRisk,
    totalKits: totalKits, totalParts: parts.length, skipped: skipped
  };
}


/** Top N critical parts — for the weekly report, without re-formatting text. */
function getKitCriticalTop(n, sharedMaps, sharedZoho, sharedKitMap) {
  var d = analyzeKitCriticality(sharedMaps, sharedZoho, sharedKitMap);
  return {
    parts:     d.parts.slice(0, n || KIT_CRITICAL.topN),
    atRisk:    d.atRisk,
    totalKits: d.totalKits,
    skipped:   d.skipped
  };
}


/**
 * The plain-text criticality report. PLAIN TEXT, no parse_mode — same robustness
 * choice as the ripple and the digest: part names carry punctuation that would
 * break a Markdown parse and silently drop the whole message.
 */
function previewKitCriticality() {
  var d;
  try { d = analyzeKitCriticality(); }
  catch (e) { return "⚠ Criticality failed: " + (e.message || e); }

  if (!d.parts.length) {
    return "No component data — the Kit Registry may be empty or unreadable.";
  }

  var L = [];
  L.push("🏗 WHAT THE CATALOGUE LEANS ON");
  L.push(d.totalParts + " parts across " + d.totalKits + " kits"
         + (d.skipped ? "  ·  " + d.skipped + " kits unreadable, excluded" : ""));

  /* ⚠ THE ALERT SET LEADS. A ranked list is a reference; the parts that are BOTH
     load-bearing and running thin are the reason to read it today. Silent when
     empty — a heading over nothing trains people to skip the section. */
  if (d.atRisk.length) {
    L.push("");
    L.push("🚨 LOAD-BEARING AND RUNNING THIN");
    for (var a = 0; a < d.atRisk.length; a++) {
      var r = d.atRisk[a];
      L.push("  " + r.sku + (r.location ? "  " + r.location : "") + "   have " + r.avail);
      if (r.name) L.push("      " + _rpClip(r.name, 46));
      L.push("      " + r.kits + " kits depend on it · covers " + r.coverKits + " more");
    }
  }

  L.push("");
  L.push("MOST DEPENDED ON");

  var shown = Math.min(d.parts.length, KIT_CRITICAL.topN);
  for (var i = 0; i < shown; i++) {
    var p = d.parts[i];
    L.push("  " + p.sku + (p.location ? "  " + p.location : "") + "   have " + p.avail);
    if (p.name) L.push("      " + _rpClip(p.name, 46));

    // The headline sentence: what happens when this shelf is empty.
    var line = "      " + p.kits + " kit" + (p.kits === 1 ? "" : "s") + " go down if it hits 0";
    if (p.maxQtyPer > 1) line += " · needs up to " + p.maxQtyPer + " each";
    L.push(line);

    if (p.cannotCover > 0) {
      // Already true, not a forecast — say so in the present tense.
      L.push("      ⚠ already short for " + p.cannotCover + " of them");
    } else {
      L.push("      covers " + p.coverKits + " more");
    }

    var ex = p.kitSkus.slice(0, KIT_CRITICAL.maxKitsListed).join(", ");
    if (p.kitSkus.length > KIT_CRITICAL.maxKitsListed) {
      ex += ", +" + (p.kitSkus.length - KIT_CRITICAL.maxKitsListed);
    }
    L.push("      → " + ex);
  }
  if (d.parts.length > shown) L.push("  … +" + (d.parts.length - shown) + " more parts");

  L.push("");
  L.push("Ranked by how many kits depend on each part — not by what is");
  L.push("short today. /ripple answers that one.");
  return L.join("\n");
}
