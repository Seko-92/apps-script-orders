// =======================================================================================
// KitBuild.js — build a kit NOBODY bought: expand it, gather it, print it
// =======================================================================================
//
// WHY THIS EXISTS (2026-09-03)
//   Pickers sometimes assemble a repair kit that has no customer behind it —
//   restocking an empty K-* shelf, or building ahead of demand. There was no door
//   for that. The workaround was to hand-type a kit SKU into a blank DIRECT row,
//   expand it, print, and delete the rows again.
//
//   That workaround cost more than it looked. A blank SALES ORDER on a live row is
//   exactly what the identity guard is built to catch, so every inserted component
//   turned red the instant it landed — correctly, because those rows genuinely had
//   no identity. And the note read "1 for customer + 2 for us" on an order with no
//   customer, which the picker deleted by hand afterwards, row by row.
//
// ⭐⭐ THE ONE PROPERTY THAT MAKES THIS SAFE: IT WRITES NOTHING.
//   No sheet rows. No Activity Log. No Script Properties. No cache. It reads the
//   registry, reads stock, and renders a page. There is nothing to clean up, nothing
//   to go stale, and nothing that can be half-done if it fails midway.
//
//   That is not laziness — it is the whole design, and it was the user's call. Routing
//   throwaway rows through All Orders charges an ordinary order's full price for a
//   piece of paper:
//     · every insert fires the dedupe painter, kit markers, enrichment, board publish
//     · the straggler watchdog Telegrams "past the 3h line" on a build left PENDING
//     · each component's RECEIVED event SURVIVES the delete, so "Received Today" on
//       the cockpit silently inflates by N every time you build an N-part kit
//     · a deleted row is the one shape the identity guard explicitly cannot see
//
// ⭐ WHAT THE WINDOW BUYS THAT THE SHEET STRUCTURALLY CANNOT: ROLL-UP.
//   On the sheet each kit expands under its own parent row, so two kits sharing a
//   piston send the picker to that shelf twice. Here they merge into one line with
//   one total. That is the reason this is a window and not a better hand-typed row.
//
// WHAT IT REUSES, UNCHANGED — nothing about kit rules is re-implemented here:
//   expandKit()                     — already PURE and sheet-free (KitExpansion.js)
//   buildKitMap()                   — the registry, for autocomplete (KitRegistry.js)
//   buildLocationAndInventoryMaps() — MI location + availability (LiveSync.js)
//   buildZohoStockMap()             — Zoho-first availability (ZohoStock.js)
//   compareLocations()              — the aisle walk, natural not lexical (Helpers.js)
//
// ⚠ THE PICKER GATE IS DELIBERATELY NOT APPLIED. preparePrintSheet refuses without a
//   Pick ID because every pick, count and adjustment for the shift is stamped with it.
//   This surface stamps nothing, so there is nothing to attribute and a refusal would
//   be friction with no safety behind it. The name is PRINTED when one is set, and
//   the sheet prints fine without it.
//
// ⚠ DEPLOY: editor-bound. The sidebar opens the modal via google.script.run and the
//   modal talks back the same way — nothing here is on the /exec surface, so
//   `clasp push` is the whole deploy for this file.
// =======================================================================================


var KIT_BUILD = {
  // Bounds, so a typo cannot produce a thousand-line pick sheet. Both are generous
  // against real use — the observed case is 1-3 kits of 1-5 each.
  maxKits:      12,
  maxQtyPerKit: 99,

  // Modal geometry. Matches the Kit Expansion modal so the two feel like one system.
  dialogWidth:  1180,
  dialogHeight: 820
};


// =======================================================================================
// SIDEBAR ENTRY
// =======================================================================================

/**
 * Open the Kit Build window.
 *
 * Passes a lightweight {sku, name} list of every registry kit for input autocomplete —
 * the same move openKitPriceCalculator makes, and for the same reason: the picker knows
 * the kit by name, not by six digits.
 */
function openKitBuildModal() {
  try {
    var kits = [];
    try {
      buildKitMap().forEach(function (k) {
        kits.push({ sku: k.sku, name: k.name, type: k.type, location: k.location });
      });
      kits.sort(function (a, b) { return String(a.sku).localeCompare(String(b.sku)); });
    } catch (e) {
      console.log("openKitBuildModal kit list error: " + e);
    }

    var template = HtmlService.createTemplateFromFile("KitBuildModal");

    // ⚠ EVERY value goes in as JSON through the FORCE-UNESCAPED printing scriptlet.
    //   The escaping variant always html-escapes, so quotes become &quot; and the
    //   JSON silently fails to parse — the modal then renders "undefined" everywhere
    //   with no error. The </-guard stops a value containing a closing script tag
    //   from ending the block early; "<\/" is a no-op escape in JS source but the
    //   HTML parser does not see it as a tag close.
    var guard = function (v) { return JSON.stringify(v).replace(/<\//g, "<\\/"); };
    template.kitListJson    = guard(kits);
    template.pickerNameJson = guard(_kbPickerName());
    template.maxKitsJson    = guard(KIT_BUILD.maxKits);

    var html = template.evaluate()
      .setWidth(KIT_BUILD.dialogWidth)
      .setHeight(KIT_BUILD.dialogHeight);
    SpreadsheetApp.getUi().showModalDialog(html, "Build Kits for Stock");
    return { ok: true, kits: kits.length };
  } catch (err) {
    try { console.log("openKitBuildModal error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


/**
 * The shift's picker, for the printed sheet's header. "—" when unset.
 *
 * ⚠ Schema.pickIdA1(), never the raw cellEmployeeId constant — the address resolves
 *   F2 vs H2 at runtime from a Script Property, and reading the constant directly is
 *   how a surface ends up pointing at the wrong cell after a layout migration.
 */
function _kbPickerName() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "—";
    var raw = String(sheet.getRange(Schema.pickIdA1()).getValue() || "").trim();
    if (!raw || !/^Shipping\s*-\s*/i.test(raw)) return "—";   // the dropdown placeholder
    return _extractPickIdData(raw);
  } catch (e) {
    return "—";
  }
}


// =======================================================================================
// THE ENGINE
// =======================================================================================

/**
 * Expand one or more kits, enrich, and roll up into a single gather list.
 *
 * @param {Object} payload
 *        { kits: [{sku, qty}], excluded: { "<kitSku>": ["<componentSku>", ...] } }
 * @returns {Object}
 *        { ok, kits: [...], gather: [...], warnings: [...], picker, builtAt, totals }
 *
 * PURE-ISH: reads the registry and the stock maps, writes nothing.
 */
function getKitBuildPlan(payload) {
  try {
    payload = payload || {};
    var wanted   = payload.kits || [];
    var excluded = payload.excluded || {};

    if (!wanted.length) {
      return { ok: false, reason: "No kits selected." };
    }
    if (wanted.length > KIT_BUILD.maxKits) {
      return { ok: false, reason: "Too many kits at once (max " + KIT_BUILD.maxKits + ")." };
    }

    // ⚠ ONE read of each map for the whole plan, never one per kit. Master Inventory
    //   is ~3,600 rows × ~198 columns; building these per kit is the O(n²) shape that
    //   made computeKitPriceBySku unusable in a loop.
    var locInvMaps = buildLocationAndInventoryMaps();
    var zohoMap    = buildZohoStockMap();

    var kits = [], warnings = [], gatherMap = {};

    for (var i = 0; i < wanted.length; i++) {
      var reqSku = String(wanted[i] && wanted[i].sku || "").trim();
      var reqQty = parseInt(wanted[i] && wanted[i].qty);
      if (!reqSku) continue;
      if (isNaN(reqQty) || reqQty < 1) reqQty = 1;
      if (reqQty > KIT_BUILD.maxQtyPerKit) reqQty = KIT_BUILD.maxQtyPerKit;

      // expandKit is already pure and already scales component qty by the multiplier.
      var plan = expandKit(reqSku, reqQty);

      if (!plan.found) {
        warnings.push({ kind: "unknown", kitSku: reqSku, message: plan.reason });
        kits.push({ kitSku: reqSku, kitName: "", qty: reqQty, found: false,
                    reason: plan.reason, components: [], unparsedLines: [] });
        continue;
      }

      // ⚠ SURFACE UNREADABLE PD LINES LOUDLY. A kit whose Purchase Description has
      //   lines the registry parser could not read has an INCOMPLETE checklist, and a
      //   gather list that looks whole when it is not is worse than no list at all.
      //   Same honesty rule as OOS BUILDABLE, Kit Health's blank computed price, and
      //   the expansion modal's red banner.
      if (plan.unparsedLines && plan.unparsedLines.length) {
        warnings.push({ kind: "unparsed", kitSku: plan.kitSku,
                        message: plan.unparsedLines.length + " line(s) in this kit's Zoho " +
                                 "Purchase Description could not be read — the list below " +
                                 "is INCOMPLETE",
                        lines: plan.unparsedLines.slice() });
      }

      // ⚠⚠ WHO OWNS THE DEFAULT — "no opinion" is NOT the same as "include everything".
      //   A kit the picker has not touched yet has NO KEY in `excluded`, and for that
      //   kit the bundled parts start UNCHECKED. A kit they HAVE touched has a key —
      //   possibly an empty array, meaning "I looked, and I want all of it" — and then
      //   their list wins outright.
      //
      //   Getting this wrong is not cosmetic: seed the default only in the client and
      //   the first toggle silently drops it, so unchecking one part quietly re-adds a
      //   gasket that ships inside another box. Ignore the client and a re-checked
      //   bundled part can never be gathered. The distinction is the key's PRESENCE.
      var kitKey = Object.prototype.hasOwnProperty.call(excluded, reqSku) ? reqSku
                 : (Object.prototype.hasOwnProperty.call(excluded, plan.kitSku) ? plan.kitSku
                                                                                : null);
      var kitExcluded = {};
      if (kitKey !== null) {
        (excluded[kitKey] || []).forEach(function (x) {
          kitExcluded[String(x).toLowerCase()] = true;
        });
      }

      var comps = [];
      for (var c = 0; c < plan.components.length; c++) {
        var comp = plan.components[c];
        var skuLower = String(comp.sku).toLowerCase();

        var enriched = _kbEnrich(comp, skuLower, locInvMaps, zohoMap);
        // Untouched kit → the bundled rule proposes. Touched kit → the picker decides.
        enriched.excluded = (kitKey === null) ? (comp.bundled === true)
                                              : !!kitExcluded[skuLower];
        comps.push(enriched);

        if (enriched.excluded) continue;   // not gathered → not in the roll-up

        // ── THE ROLL-UP. Keyed on the SKU, so two kits sharing a part become one
        //    line and one walk. This is the capability the sheet path cannot have,
        //    because there each kit expands under its own parent row.
        var g = gatherMap[skuLower];
        if (!g) {
          g = gatherMap[skuLower] = {
            sku: comp.sku, name: comp.name, location: enriched.location,
            available: enriched.available, missing: enriched.missing,
            totalQty: 0, usedBy: []
          };
        }
        g.totalQty += enriched.qty;
        g.usedBy.push({ kitSku: plan.kitSku, kits: reqQty,
                        perKit: reqQty > 0 ? (enriched.qty / reqQty) : enriched.qty,
                        subtotal: enriched.qty });
      }

      kits.push({
        found:         true,
        kitSku:        plan.kitSku,
        kitName:       plan.kitName,
        kitType:       plan.kitType,
        kitLocation:   plan.kitLocation || "NOT FOUND",
        kitEngine:     plan.kitEngine || "",
        qty:           reqQty,
        components:    comps,
        unparsedLines: (plan.unparsedLines || []).slice()
      });
    }

    // ⚠ compareLocations, NEVER localeCompare — lexically "A-9" sorts AFTER "A-50",
    //   and this list's entire job is to be walked in aisle order. That bug shipped
    //   in three files and was only caught when a new surface printed it in a column.
    var gather = Object.keys(gatherMap).map(function (k) { return gatherMap[k]; });
    gather.sort(function (a, b) {
      var c = compareLocations(a.location, b.location);
      return c !== 0 ? c : String(a.sku).localeCompare(String(b.sku));
    });

    var totals = {
      kits:       kits.filter(function (k) { return k.found; }).length,
      kitUnits:   kits.reduce(function (n, k) { return n + (k.found ? k.qty : 0); }, 0),
      lines:      gather.length,
      pieces:     gather.reduce(function (n, g) { return n + g.totalQty; }, 0),
      shelves:    _kbDistinctShelves(gather),
      notFound:   gather.filter(function (g) { return g.missing; }).length,
      shared:     gather.filter(function (g) { return g.usedBy.length > 1; }).length
    };

    return {
      ok: true, kits: kits, gather: gather, warnings: warnings, totals: totals,
      picker: _kbPickerName(), builtAt: _kbStamp()
    };
  } catch (err) {
    try { console.log("getKitBuildPlan error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, reason: String(err.message || err) };
  }
}


/**
 * One component, enriched with where it is and how many there are.
 *
 * ⚠ ZOHO-FIRST, MI FALLBACK — lifted from previewSelectedKits rather than re-derived.
 *   Kit components are DIRECT-side items and many are not listed on eBay, so MI reads
 *   null or stale for them. This is the same routing recomputeHand, LiveSync, the Prep
 *   Queue and the DIRECT pull insert all use; a surface that resolves stock differently
 *   is a surface whose numbers disagree with HAND on the next recompute.
 */
function _kbEnrich(comp, skuLower, locInvMaps, zohoMap) {
  var loc = locInvMaps.locationMap.get(skuLower) || "NOT FOUND";
  var inv = locInvMaps.inventoryMap.get(skuLower);
  var zo  = zohoMap.get(skuLower);
  var miAvail = (inv && inv.available != null) ? inv.available : null;
  var zoAvail = zo ? zo.available : null;
  var available = (zoAvail != null) ? zoAvail : miAvail;

  return {
    sku:         comp.sku,
    name:        comp.name,
    qty:         comp.qty,
    location:    loc,
    available:   available,
    missing:     (loc === "NOT FOUND"),
    // ⭐ Ships INSIDE another component's box. expandKit already worked this out from
    //   the Sales Description. Starting these UNCHECKED is the whole point: a picker
    //   sent after a gasket already sitting in the gasket set walks for nothing.
    bundled:     comp.bundled === true,
    bundledInto: comp.bundledInto || "",
    // Short by our best current reading. Informational only — the shelf decides.
    short:       (available != null && available < comp.qty)
  };
}


/** How many distinct shelves the walk touches. NOT FOUND is not a shelf. */
function _kbDistinctShelves(gather) {
  var seen = {};
  for (var i = 0; i < gather.length; i++) {
    var l = String(gather[i].location || "").trim();
    if (!l || l === "NOT FOUND") continue;
    seen[l.toUpperCase()] = true;
  }
  return Object.keys(seen).length;
}


/** Houston wall-clock stamp for the printed header. */
function _kbStamp() {
  try {
    return Utilities.formatDate(new Date(), "America/Chicago", "M/d/yy h:mm a");
  } catch (e) {
    return "";
  }
}


// =======================================================================================
// EDITOR WRAPPER — the Run button shows no return value, so it logs.
// =======================================================================================

/** Editor-run smoke test: expand the first two registry kits and print the roll-up. */
function previewKitBuildNow() {
  var kits = [];
  try {
    buildKitMap().forEach(function (k) { if (kits.length < 2) kits.push({ sku: k.sku, qty: 2 }); });
  } catch (e) {}
  if (!kits.length) { console.log("No kits in the registry."); return; }

  var out = getKitBuildPlan({ kits: kits });
  console.log(JSON.stringify(out, null, 2));
  return out;
}
