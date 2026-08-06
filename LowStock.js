// =======================================================================================
// LOWSTOCK.gs — "what sold this week and is now running out"
// =======================================================================================
//
// THE SCOPE, AND WHY IT IS NARROWER THAN "everything under 5"
//   This is a REORDER report, so it is driven by what MOVED, not by the whole
//   catalogue. A part with 3 on the shelf that has not sold in a year is dead
//   stock — putting it in a purchasing report is noise that hides the real
//   lines. So:
//
//       Activity Log (what SHIPPED in the window)  ->  the candidate list
//       Master Inventory (what is LEFT)            ->  the verdict
//
//   Catalogue-wide "what is low regardless of demand" is a different question
//   and already has a home: the Out of Stock sheet.
//
// THE RANKING — DAYS OF COVER, not quantity
//   A flat list sorted by quantity is a count, not a decision. Both of these
//   are "3 left":
//       3 left · sold 12 this week  ->  under 2 days of cover   ORDER NOW
//       3 left · sold 1  this week  ->  three weeks of cover    fine
//   Only one is urgent. days of cover = (available / sold) * windowDays, so the
//   report sorts itself into the order a buyer would actually work down.
//
// ⚠ ZERO IS INCLUDED, DELIBERATELY. An item that sold this week and is now at
//   zero is the most urgent line on the page, not an exclusion. It is tagged
//   OUT so it reads differently from "getting low".
//
// HONESTY NOTES
//   * Demand comes from the ACTIVITY LOG (real SHIPPED events in the window),
//     never from MI's `quantitySold` — that field is LIFETIME for the listing
//     and would badly overstate a week. MI's number is reported separately and
//     labelled as lifetime.
//   * The Activity Log keeps 90 days (ACTIVITY_LOG.retentionDays), so any
//     window beyond that silently under-reports. We clamp and say so.
//   * Both channels are covered: eBay and DIRECT rows both flow through
//     updateOrderStatus, so both write SHIPPED events.
//
// Pure Tier 1 — no new data, no new dependency. One MI read, one log read.
// =======================================================================================

var LOW_STOCK = {
  days:      7,     // lookback window, in days
  threshold: 5,     // "low" = available at or below this
  tgLimit:   25,    // rows listed in a Telegram reply before summarising
  pdfLimit:  60,    // rows listed in the PDF
  nameClip:  46,    // item-title clip length in text output

  // Embed an item thumbnail per row in the PDF. Flip to false if Google's
  // HTML->PDF converter turns out not to fetch remote images on your account —
  // the table degrades cleanly to text-only.
  images:    true,

  // Flag when Zoho and Master Inventory disagree by at least this much. Small
  // gaps are normal timing jitter between a 2-min push and an hourly sync.
  divergeBy: 2,

  // Days of cover below which a line is genuinely urgent. Above it the item is
  // low in COUNT but not in TIME, so it drops below a divider instead of
  // competing for attention at the top of the page.
  urgentDays: 14
};


// =======================================================================================
// ENGINE
// =======================================================================================

/**
 * What sold in the window and is now at/below the threshold.
 *
 * @param {Object} [opts] { days, threshold }
 * @returns {Object} see the return literal at the bottom
 */
function analyzeLowStock(opts) {
  opts = opts || {};

  var days = parseInt(opts.days, 10);
  if (isNaN(days) || days < 1) days = LOW_STOCK.days;
  var clamped = false;
  var maxDays = (typeof ACTIVITY_LOG !== 'undefined' && ACTIVITY_LOG.retentionDays) || 90;
  if (days > maxDays) { days = maxDays; clamped = true; }   // never claim data we purged

  var threshold = parseInt(opts.threshold, 10);
  if (isNaN(threshold) || threshold < 0) threshold = LOW_STOCK.threshold;

  try {
    // ---- 1. what SHIPPED in the window -------------------------------------
    var sold = _lsSoldInWindow(days);
    if (!sold.ok) {
      return _lsEmpty(days, threshold, sold.message);
    }
    // Zoho sales that never reached the sheet — merged into the SAME demand map
    // so every downstream number (cover, totals, ranking) counts both channels.
    var zsales = _lsZohoSalesInWindow(days);
    zsales.map.forEach(function (z, key) {
      if (!sold.map.has(key)) {
        sold.map.set(key, { sku: z.sku, qty: 0, orders: new Set(), last: z.last });
      }
      var rec = sold.map.get(key);
      rec.qty += z.qty;
      z.orders.forEach(function (o) { rec.orders.add(o); });
      if (z.last > rec.last) rec.last = z.last;
    });
    sold.totalQty += zsales.totalQty;

    if (!sold.map.size) {
      return _lsEmpty(days, threshold,
        "Nothing shipped in the last " + days + " day" + (days === 1 ? "" : "s") + ".");
    }

    // ---- 2. what is LEFT ----------------------------------------------------
    // TWO sources, on purpose:
    //   Zoho  = the inventory MASTER (it pushes stock TO eBay) and the fresher
    //           of the two (2-min push during work hours vs MI's hourly). It
    //           also covers DIRECT-only parts that were never listed on eBay.
    //   MI    = eBay's view. Needed for shelf LOCATION and listing status, and
    //           as the fallback for anything Zoho does not know.
    // Stock verdict is ZOHO-FIRST -> MI fallback, matching the routing rule the
    // rest of the system already uses for DIRECT-side items.
    var inv = _lsInventorySnapshot();
    if (!inv.ok) return _lsEmpty(days, threshold, inv.message);

    var zoho = new Map();
    try {
      if (typeof buildZohoStockMap === 'function') zoho = buildZohoStockMap() || new Map();
    } catch (e) {
      try { console.log("analyzeLowStock: Zoho stock unavailable — " + e); } catch (_) {}
    }

    // ---- 3. which kits a part would block (best effort) ----------------------
    var kitIndex = null;
    try {
      if (typeof buildKitMap === 'function' && typeof _invertKitMap === 'function') {
        kitIndex = _invertKitMap(buildKitMap());
      }
    } catch (e) {
      try { console.log("analyzeLowStock: kit index unavailable — " + e); } catch (_) {}
    }

    // ---- 4. join -------------------------------------------------------------
    var items = [], healthy = 0, unknown = 0, outCount = 0;

    sold.map.forEach(function (s, skuLower) {
      var rec = inv.map.get(skuLower);
      var zr  = zoho.get(skuLower);

      // Not in EITHER source. Report it rather than drop it — silently losing a
      // line from a reorder report is the worst failure here.
      if (!rec && !zr) {
        unknown++;
        items.push({
          sku: s.sku, name: "", image: "", listing: "",
          location: "NOT FOUND", available: null,
          miAvail: null, zohoAvail: null, diverges: false, zohoOnly: false,
          soldQty: s.qty, orders: s.orders.size, lastSold: s.last,
          daysCover: 0, lifetimeSold: null, status: "", unknown: true,
          blocksKits: []
        });
        return;
      }

      // A Zoho-only part (never listed on eBay) has no MI row at all — that is
      // normal for DIRECT stock, not an error. Synthesise the shape.
      if (!rec) {
        rec = { sku: zr.skuOriginal || s.sku, name: zr.itemName || "", image: "",
                listing: "", location: "NOT FOUND", available: zr.available,
                sold: null, status: "" };
      }

      // Active-only, FAIL-OPEN on blank (a renamed header must never empty the
      // report) — the same rule the OOS sheet uses. Only applies to items that
      // ARE on eBay; a Zoho-only part has no listing status to judge.
      var st = String(rec.status || "").trim();
      if (st && st.toLowerCase() !== "active") return;

      // ZOHO-FIRST verdict, MI fallback.
      var miAvail   = rec.available;
      var zoAvail   = zr ? zr.available : null;
      var available = (zoAvail != null) ? zoAvail : miAvail;

      // Divergence is signal in its own right: Zoho lower than MI usually means
      // a DIRECT sale that eBay's hourly sync has not caught up with yet.
      var diverges = (zoAvail != null && miAvail != null &&
                      Math.abs(zoAvail - miAvail) >= LOW_STOCK.divergeBy);

      if (available > threshold) { healthy++; return; }
      if (available <= 0) outCount++;

      var cover = (s.qty > 0) ? (Math.max(0, available) / s.qty) * days : null;

      var blocks = [];
      if (kitIndex) {
        var norm = (typeof _normPartSku === 'function') ? _normPartSku(s.sku) : skuLower;
        var uses = kitIndex.get(norm) || [];
        for (var u = 0; u < uses.length; u++) blocks.push(uses[u].kitSku);
      }

      items.push({
        sku:          rec.sku || s.sku,
        name:         rec.name || "",
        image:        rec.image || "",
        listing:      rec.listing || "",
        location:     rec.location || "NOT FOUND",
        available:    available,
        miAvail:      miAvail,
        zohoAvail:    zoAvail,
        diverges:     diverges,
        zohoOnly:     !inv.map.get(skuLower),
        soldQty:      s.qty,
        orders:       s.orders.size,
        lastSold:     s.last,
        daysCover:    cover,
        lifetimeSold: rec.sold,
        status:       st,
        unknown:      false,
        blocksKits:   blocks
      });
    });

    items.sort(_lsCompare);

    return {
      ok: true,
      days: days, threshold: threshold, clamped: clamped,
      items: items,
      count: items.length,
      outCount: outCount,
      healthyCount: healthy,
      unknownCount: unknown,
      soldSkuCount: sold.map.size,
      totalUnitsSold: sold.totalQty,
      zohoUnits: zsales.totalQty,
      zohoOrders: zsales.orders,
      generatedAt: new Date(),
      message: ""
    };

  } catch (err) {
    try { console.log("analyzeLowStock: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return _lsEmpty(days, threshold, String(err.message || err));
  }
}


/**
 * Sort: most urgent first.
 *   1. anything we could not resolve in MI (needs a human)
 *   2. fewest days of cover
 *   3. biggest seller breaks the tie
 */
function _lsCompare(a, b) {
  if (a.unknown !== b.unknown) return a.unknown ? -1 : 1;
  var ac = (a.daysCover == null) ? Number.MAX_VALUE : a.daysCover;
  var bc = (b.daysCover == null) ? Number.MAX_VALUE : b.daysCover;
  if (ac !== bc) return ac - bc;
  if (b.soldQty !== a.soldQty) return b.soldQty - a.soldQty;
  return String(a.sku).localeCompare(String(b.sku));
}


// =======================================================================================
// PRIVATE READS
// =======================================================================================

/**
 * SHIPPED events in the last N days, summed per SKU.
 * @returns {{ok:boolean, map:Map, totalQty:number, message:string}}
 *          map: skuLower -> { sku, qty, orders:Set, last:Date }
 */
function _lsSoldInWindow(days) {
  var out = new Map(), totalQty = 0;
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss && ss.getSheetByName(ACTIVITY_LOG.sheetName);
    if (!sheet) return { ok: false, map: out, totalQty: 0, message: "Activity Log sheet not found." };

    var lastRow = sheet.getLastRow();
    if (lastRow < ACTIVITY_LOG.dataStartRow) {
      return { ok: true, map: out, totalQty: 0, message: "" };
    }

    var n = lastRow - ACTIVITY_LOG.dataStartRow + 1;
    var data = sheet.getRange(ACTIVITY_LOG.dataStartRow, 1, n, ACTIVITY_LOG.dataWidth).getValues();

    var TS = ACTIVITY_LOG.idx("TIMESTAMP"),
        EV = ACTIVITY_LOG.idx("EVENT"),
        SK = ACTIVITY_LOG.idx("SKU"),
        QT = ACTIVITY_LOG.idx("QTY"),
        OR = ACTIVITY_LOG.idx("ORDER_ID");

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    for (var i = 0; i < data.length; i++) {
      var ts = data[i][TS];
      if (!(ts instanceof Date) || ts < cutoff) continue;
      if (String(data[i][EV] || "").trim().toUpperCase() !== "SHIPPED") continue;

      var sku = String(data[i][SK] || "").trim();
      if (!sku) continue;
      var key = sku.toLowerCase();

      var qty = parseInt(data[i][QT], 10);
      if (isNaN(qty) || qty < 0) qty = 0;

      if (!out.has(key)) out.set(key, { sku: sku, qty: 0, orders: new Set(), last: ts });
      var rec = out.get(key);
      rec.qty += qty;
      totalQty += qty;
      var oid = String(data[i][OR] || "").trim();
      if (oid) rec.orders.add(oid);
      if (ts > rec.last) rec.last = ts;
    }

    return { ok: true, map: out, totalQty: totalQty, message: "" };

  } catch (err) {
    try { console.log("_lsSoldInWindow: " + err); } catch (_) {}
    return { ok: false, map: out, totalQty: 0, message: String(err.message || err) };
  }
}


/**
 * Zoho sales the Activity Log CANNOT know about.
 *
 * THE GAP THIS CLOSES (found 2026-08-06): demand was read only from the
 * Activity Log, which records SHIPPED events for rows that exist on the All
 * Orders sheet. A Zoho sales order is only ever written there if somebody
 * PULLED it. Everything sold in Zoho and never pulled — and there is a standing
 * backlog of those — was completely invisible to this report, so parts were
 * being under-counted and dropping off the reorder list.
 *
 * ⚠ DEDUPE RULE, and it is the whole correctness story: count a Pending row
 * ONLY when PULLED is blank. A pulled SO already has DIRECT rows on the sheet,
 * whose SHIPPED events the Activity Log has counted — counting it here too
 * would double every pulled Zoho sale.
 *
 * Line items come from the cached slim PAYLOAD (col L), which carries
 * { sku, quantity, name } per line — see _slimSalesOrder in ZohoSalesOrders.js.
 *
 * @returns {{ok:boolean, map:Map, totalQty:number, orders:number}}
 *          map: skuLower -> { sku, qty, orders:Set, last:Date }
 */
function _lsZohoSalesInWindow(days) {
  var out = new Map(), totalQty = 0, orders = 0;
  try {
    if (typeof PENDING_SO === 'undefined') return { ok: true, map: out, totalQty: 0, orders: 0 };
    var ss = SpreadsheetApp.getActive();
    var sheet = ss && ss.getSheetByName(PENDING_SO.sheetName);
    if (!sheet) return { ok: true, map: out, totalQty: 0, orders: 0 };

    var lastRow = sheet.getLastRow();
    if (lastRow < PENDING_SO.dataStartRow) return { ok: true, map: out, totalQty: 0, orders: 0 };

    var n = lastRow - PENDING_SO.dataStartRow + 1;
    var data = sheet.getRange(PENDING_SO.dataStartRow, 1, n, PENDING_SO.dataWidth).getValues();

    var C_DATE = PENDING_SO.idx("DATE"),
        C_PULL = PENDING_SO.idx("PULLED"),
        C_PAY  = PENDING_SO.idx("PAYLOAD"),
        C_SO   = PENDING_SO.idx("SO_NUMBER"),
        C_STAT = PENDING_SO.idx("ORDER_STATUS");

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    for (var i = 0; i < data.length; i++) {
      // already on the sheet -> the Activity Log owns it
      if (String(data[i][C_PULL] || "").trim()) continue;

      var when = data[i][C_DATE];
      if (!(when instanceof Date)) {
        var parsed = new Date(when);
        if (isNaN(parsed.getTime())) continue;
        when = parsed;
      }
      if (when < cutoff) continue;

      // a voided order is not a sale
      if (String(data[i][C_STAT] || "").trim().toUpperCase() === "VOID") continue;

      var raw = String(data[i][C_PAY] || "").trim();
      if (!raw) continue;
      var payload;
      try { payload = JSON.parse(raw); } catch (e) { continue; }
      var lines = (payload && payload.line_items) || [];
      if (!lines.length) continue;

      var soNum = String(data[i][C_SO] || "").trim() || ("row" + i);
      orders++;

      for (var li = 0; li < lines.length; li++) {
        var sku = String(lines[li].sku || "").trim();
        if (!sku) continue;
        var key = sku.toLowerCase();
        var q = parseInt(lines[li].quantity, 10);
        if (isNaN(q) || q < 0) q = 0;

        if (!out.has(key)) out.set(key, { sku: sku, qty: 0, orders: new Set(), last: when });
        var rec = out.get(key);
        rec.qty += q;
        totalQty += q;
        rec.orders.add(soNum);
        if (when > rec.last) rec.last = when;
      }
    }
    return { ok: true, map: out, totalQty: totalQty, orders: orders };

  } catch (err) {
    try { console.log("_lsZohoSalesInWindow: " + err); } catch (_) {}
    return { ok: false, map: out, totalQty: 0, orders: 0 };
  }
}


/**
 * ONE Master Inventory pass: sku -> { sku, name, location, available, sold, status }.
 * Deliberately its own read rather than buildLocationAndInventoryMaps() — that
 * helper carries no TITLE, and a reorder report without item names is unusable
 * for whoever has to place the order.
 */
function _lsInventorySnapshot() {
  var map = new Map();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(DB_SHEET_NAME);
    if (!sheet) return { ok: false, map: map, message: "Master Inventory sheet not found." };

    var data = sheet.getDataRange().getValues();
    if (!data.length) return { ok: false, map: map, message: "Master Inventory is empty." };

    var h = data[0];
    var cSku    = h.indexOf(DB_SKU_HEADER);
    var cTitle  = h.indexOf(DB_TITLE_HEADER);
    var cLoc    = h.indexOf(DB_LOCATION_HEADER);
    var cQty    = h.indexOf(DB_QUANTITY_HEADER);
    var cSold   = h.indexOf(DB_QUANTITY_SOLD_HEADER);
    var cStatus = h.indexOf(DB_LISTING_STATUS_HEADER);
    var cView   = h.indexOf(DB_VIEWURL_HEADER);
    // First non-empty pictureUrl1..5 — same columns PhotoQueue reads.
    var picCols = [];
    for (var p = 1; p <= 5; p++) {
      var ci = h.indexOf('pictureUrl' + p);
      if (ci !== -1) picCols.push(ci);
    }

    if (cSku === -1 || cQty === -1 || cSold === -1) {
      return { ok: false, map: map, message: "Master Inventory is missing SKU/quantity columns." };
    }

    for (var i = 1; i < data.length; i++) {
      var sku = String(data[i][cSku] || "").trim();
      if (!sku) continue;
      var qty  = parseInt(data[i][cQty], 10)  || 0;
      var sold = parseInt(data[i][cSold], 10) || 0;
      var pic = "";
      for (var pc = 0; pc < picCols.length; pc++) {
        var cand = String(data[i][picCols[pc]] || "").trim();
        if (cand) { pic = cand; break; }
      }

      map.set(sku.toLowerCase(), {
        sku:       sku,
        name:      cTitle  !== -1 ? String(data[i][cTitle] || "").trim() : "",
        image:     pic,
        listing:   cView   !== -1 ? String(data[i][cView] || "").trim() : "",
        location:  cLoc    !== -1 ? (String(data[i][cLoc] || "").trim() || "NOT FOUND") : "NOT FOUND",
        available: qty - sold,
        sold:      sold,
        status:    cStatus !== -1 ? String(data[i][cStatus] || "").trim() : ""
      });
    }
    return { ok: true, map: map, message: "" };

  } catch (err) {
    try { console.log("_lsInventorySnapshot: " + err); } catch (_) {}
    return { ok: false, map: map, message: String(err.message || err) };
  }
}


function _lsEmpty(days, threshold, message) {
  return {
    ok: true, days: days, threshold: threshold, clamped: false,
    items: [], count: 0, outCount: 0, healthyCount: 0, unknownCount: 0,
    soldSkuCount: 0, totalUnitsSold: 0,
    generatedAt: new Date(), message: message || ""
  };
}


// =======================================================================================
// FORMATTERS
// =======================================================================================

/** Round days-of-cover to something a human reads at a glance. */
function _lsCoverText(d) {
  if (d == null) return "—";
  if (d <= 0)  return "OUT";
  if (d < 1)   return "<1d";
  if (d < 10)  return (Math.round(d * 10) / 10) + "d";
  return Math.round(d) + "d";
}

function _lsClip(s, n) {
  s = String(s == null ? "" : s);
  n = n || LOW_STOCK.nameClip;
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}


/**
 * Plain-text report — used by Telegram and by the sidebar preview.
 * PLAIN TEXT, no parse_mode: the same robustness choice as the weekly digest
 * (item titles contain *, _, [ and would break Markdown at random).
 */
function buildLowStockText(data) {
  var d = data || analyzeLowStock();
  if (!d.ok) return "⚠ Low-stock report failed: " + (d.message || "unknown error");

  var L = [];
  L.push("📉 LOW STOCK · sold in the last " + d.days + " day" + (d.days === 1 ? "" : "s"));
  L.push("");

  if (!d.count) {
    L.push(d.message || ("Nothing sold in the window has fallen to " + d.threshold + " or below."));
    if (d.soldSkuCount) {
      L.push("");
      L.push(d.soldSkuCount + " SKUs sold · all still above " + d.threshold + ".");
    }
    return L.join("\n");
  }

  L.push(d.count + " item" + (d.count === 1 ? "" : "s") + " at or below " + d.threshold +
         (d.outCount ? "  (" + d.outCount + " already OUT)" : ""));
  L.push("of " + d.soldSkuCount + " SKUs sold · " + d.totalUnitsSold + " units" +
         (d.zohoUnits ? "  (incl. " + d.zohoUnits + " from " + d.zohoOrders + " un-pulled Zoho SOs)" : ""));
  L.push("");

  var show = d.items.slice(0, LOW_STOCK.tgLimit);
  for (var i = 0; i < show.length; i++) {
    var it = show[i];
    var head = (it.available == null ? "?" : it.available) + " left";
    if (it.available != null && it.available <= 0) head = "OUT";

    L.push("• " + it.sku + "  " + head +
           "  · sold " + it.soldQty +
           "  · cover " + _lsCoverText(it.daysCover));
    if (it.name)     L.push("    " + _lsClip(it.name));
    var meta = "    " + (it.location || "NOT FOUND");
    if (it.unknown)  meta += "  · ⚠ in neither Zoho nor MI";
    if (it.zohoOnly) meta += "  · Zoho only";
    if (it.diverges) meta += "  · Zoho " + it.zohoAvail + " vs eBay " + it.miAvail;
    if (it.blocksKits && it.blocksKits.length) {
      meta += "  · blocks " + it.blocksKits.length + " kit" + (it.blocksKits.length === 1 ? "" : "s");
    }
    L.push(meta);
  }

  if (d.count > show.length) L.push("", "… +" + (d.count - show.length) + " more");
  if (d.clamped) L.push("", "(window clamped to the Activity Log's " + d.days + "-day retention)");

  return L.join("\n");
}


/** Editor/sidebar preview — never sends anything. */
function previewLowStock(days, threshold) {
  return buildLowStockText(analyzeLowStock({ days: days, threshold: threshold }));
}


// =======================================================================================
// PDF  —  the standalone "what to reorder" report
// =======================================================================================
//
// ⚠ HTML→PDF CAVEAT (the same one Report.js is built around): Google's converter
// is BASIC — no web fonts, no flexbox, no grid. So this is deliberately TABLE
// based with inline styles. Clean and reliable beats clever and broken.
//
// ⚠ AND THE ONE THAT BIT US 2026-08-06 — TWO WRONG GUESSES BEFORE THE RIGHT ONE.
// The first live PDF rendered the header band's yellow text on WHITE and the red
// hook band invisible. Text `color:`, borders, `align=` and `width=` all survived;
// only fills died.
//   guess 1: "css background is unsupported"      -> WRONG
//   guess 2: "use the legacy bgcolor attribute"   -> ALSO WRONG, changed nothing
//   actual : the converter honours `background` on BLOCK elements (<div>) but
//            drops it on TABLE CELLS (<td>/<tr>).
// The evidence was already in the repo: Report.js builds its bands as
// <div style="background:#1a1a1a"> and that PDF renders correctly.
// SO: every coloured band here is a <div>. Inside tables, emphasis comes from
// BORDERS and TYPE, never a fill. Do not "tidy" a band back into a table cell.

/**
 * Build the report HTML.
 * @param {Object} d  an analyzeLowStock() result
 */
function _lsBuildHtml(d, dateStr, genTs) {
  var F = "font-family:Helvetica,Arial,sans-serif;";
  var h = [];

  h.push('<div style="' + F + 'color:#1d1d1b;">');

  // header band
  // ⚠ MUST be a <div>, not a table cell. See the converter note above.
  h.push('<div style="background:#1d1d1b;padding:18px 22px;">' +
           '<span style="' + F + 'color:#ffd400;font-size:22px;font-weight:bold;letter-spacing:1px;">HQ MOTOR SERVICE</span>' +
           '<span style="' + F + 'color:#ffffff;font-size:13px;"> &nbsp;·&nbsp; LOW STOCK &amp; REORDER</span>' +
           '<div style="' + F + 'color:#bdbdbd;font-size:11px;padding-top:5px;">' + _rEscLs(dateStr) +
           ' &nbsp;·&nbsp; sold in the last ' + d.days + ' day' + (d.days === 1 ? '' : 's') + '</div>' +
         '</div>');

  // hook — also a div, same reason
  var hookBg = d.outCount ? '#b71c1c' : (d.count ? '#ffd400' : '#e8f5e9');
  var hookFg = d.outCount ? '#ffffff' : '#1d1d1b';
  var hookTxt = d.count
    ? (d.count + ' item' + (d.count === 1 ? '' : 's') + ' need reordering' +
       (d.outCount ? '  ·  ' + d.outCount + ' already OUT OF STOCK' : ''))
    : 'Nothing sold this period has fallen to ' + d.threshold + ' or below.';
  h.push('<div style="background:' + hookBg + ';padding:13px 22px;' + F +
         'color:' + hookFg + ';font-size:16px;font-weight:bold;">' + _rEscLs(hookTxt) + '</div>');

  // at a glance
  h.push('<table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;margin-top:14px;' + F + '">');
  h.push('<tr>' +
    _lsKpi(d.soldSkuCount,   'SKUs SOLD') +
    _lsKpi(d.totalUnitsSold, d.zohoUnits ? 'UNITS OUT (eBay+Zoho)' : 'UNITS OUT') +
    _lsKpi(d.count,          'NEED REORDER', d.count ? '#8a6d00' : null) +
    _lsKpi(d.outCount,       'AT ZERO', d.outCount ? '#b71c1c' : null) +
  '</tr></table>');

  if (d.count) {
    h.push('<div style="' + F + 'font-size:12px;color:#555;margin:16px 0 6px;">' +
           'Ordered by <b>days of cover</b> — how long the shelf lasts at this period\'s rate. ' +
           'Lowest first. "Sold" is real shipments in the window, not lifetime.</div>');

    h.push('<table width="100%" cellpadding="6" cellspacing="0" ' +
           'style="border-collapse:collapse;' + F + 'font-size:11px;">');
    // <thead> so the column labels REPEAT on every page. Page 2 of the first
    // live run had an unlabelled table — the same lesson already learned in
    // PrintFulfillment.html: a per-page repeating element belongs in <thead>,
    // never in a plain <tr>.
    h.push('<thead>');
    // Table cells cannot carry a fill in this converter, so the header row
    // earns its weight from BORDERS + type, which do render.
    h.push('<tr style="color:#1d1d1b;border-bottom:2px solid #1d1d1b;">' +
      (LOW_STOCK.images ? _lsTh('') : '') +
      _lsTh('SKU') + _lsTh('ITEM') + _lsTh('SHELF') +
      _lsTh('LEFT', 'right') + _lsTh('SOLD', 'right') + _lsTh('COVER', 'right') + _lsTh('NOTE') +
    '</tr>');
    h.push('</thead><tbody>');

    var rows = d.items.slice(0, LOW_STOCK.pdfLimit);
    var tailOpened = false;
    for (var i = 0; i < rows.length; i++) {
      var it = rows[i];

      // Everything here is at/below the quantity threshold, but "5 left selling
      // 1 a week" is 35 days of cover — not a purchase decision. Since the list
      // is already sorted by cover, one divider separates what must be ordered
      // now from what is merely worth watching, without hiding anything.
      if (!tailOpened && it.daysCover != null && it.daysCover >= LOW_STOCK.urgentDays) {
        tailOpened = true;
        h.push('<tr><td colspan="' + (LOW_STOCK.images ? 8 : 7) + '" ' +
               'style="padding:14px 0 6px;' + F + 'font-size:10px;letter-spacing:1.5px;' +
               'color:#8a8a80;border-bottom:2px solid #d8d2c0;">' +
               'WATCH &nbsp;·&nbsp; over ' + LOW_STOCK.urgentDays +
               ' days of cover at this rate &nbsp;·&nbsp; not urgent</td></tr>');
      }
      var out = (it.available != null && it.available <= 0);
      var note = [];
      if (it.unknown)  note.push('in neither Zoho nor Master Inventory');
      if (it.zohoOnly) note.push('Zoho only · not listed on eBay');
      if (it.diverges) note.push('Zoho ' + it.zohoAvail + ' vs eBay ' + it.miAvail);
      if (it.blocksKits && it.blocksKits.length) {
        note.push('blocks ' + it.blocksKits.length + ' kit' + (it.blocksKits.length === 1 ? '' : 's'));
      }
      var thumb = "";
      if (LOW_STOCK.images) {
        thumb = it.image
          ? _lsTd('<img src="' + _rEscLs(_lsThumb(it.image, 90)) + '" width="46" ' +
                  'style="border:1px solid #ddd;">')
          : _lsTd('<span style="color:#bbb;font-size:9px;">no photo</span>');
      }

      h.push('<tr style="border-bottom:1px solid ' + (out ? '#e0b4b4' : '#eeeeee') + ';">' + thumb +
        _lsTd('<b>' + _rEscLs(it.sku) + '</b>') +
        _lsTd(_rEscLs(_lsClip(it.name, 62))) +
        _lsTd(_rEscLs(it.location)) +
        _lsTd(out ? '<b style="color:#b71c1c;">OUT</b>'
                  : (it.available == null ? '?' : String(it.available)), 'right') +
        _lsTd(String(it.soldQty), 'right') +
        _lsTd(_lsCoverText(it.daysCover), 'right') +
        _lsTd('<span style="color:#8a6d00;">' + _rEscLs(note.join(' · ')) + '</span>') +
      '</tr>');
    }
    h.push('</tbody></table>');

    if (d.count > rows.length) {
      h.push('<div style="' + F + 'font-size:11px;color:#777;padding-top:6px;">… +' +
             (d.count - rows.length) + ' more below the cut.</div>');
    }
  }

  var foot = d.healthyCount + ' other SKUs sold and are still comfortably stocked.' +
             ' Stock is Zoho-first (the inventory master), falling back to eBay\'s Master Inventory.' +
             ' Demand counts eBay shipments plus Zoho sales orders that were never pulled to the sheet' +
             (d.zohoOrders ? ' (' + d.zohoOrders + ' this period).' : '.');
  if (d.clamped) foot += ' · window clamped to the Activity Log\'s ' + d.days + '-day retention';
  h.push('<div style="' + F + 'font-size:10px;color:#888;border-top:1px solid #ddd;margin-top:18px;padding-top:8px;">' +
         _rEscLs(foot) + ' &nbsp;·&nbsp; generated ' + _rEscLs(genTs) + ' &nbsp;·&nbsp; HQ Motor Service</div>');

  h.push('</div>');
  return h.join('');
}

/**
 * eBay serves every size of a picture from the same path — `s-lNNN` is the size
 * token — so a thumbnail costs a string swap rather than a fetch. Keeping these
 * SMALL matters: the PDF embeds one per row.
 *
 * ⚠ Whether Google's HTML->PDF converter actually fetches remote images is not
 * guaranteed (it is a basic renderer — no web fonts, no flexbox). If the first
 * PDF comes back with blank boxes, set LOW_STOCK.images = false and the layout
 * degrades cleanly to the text-only table it was before.
 */
function _lsThumb(url, px) {
  if (!url) return "";
  px = px || 90;
  return String(url).replace(/\/s-l\d+\.(jpg|jpeg|png|webp)/i, '/s-l' + px + '.$1');
}

function _lsKpi(v, label, tone) {
  return '<td width="25%" align="center" style="border:1px solid #e0e0e0;">' +
    '<div style="font-size:23px;font-weight:bold;color:' + (tone || '#1d1d1b') + ';">' + v + '</div>' +
    '<div style="font-size:9px;color:#777;letter-spacing:1px;">' + label + '</div></td>';
}
function _lsTh(t, align) {
  return '<td align="' + (align || 'left') + '" style="font-size:9px;letter-spacing:1px;">' + t + '</td>';
}
function _lsTd(t, align) {
  return '<td align="' + (align || 'left') + '" style="border-bottom:1px solid #eee;">' + t + '</td>';
}
/** Local escape — Report.js's _rEsc may not be loaded in every context. */
function _rEscLs(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}


/** Render the current low-stock report as a PDF blob. */
function _lowStockPdfBlob(opts) {
  var d  = analyzeLowStock(opts);
  var tz = (typeof WEEKLY_DIGEST !== 'undefined' && WEEKLY_DIGEST.timezone) || Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(new Date(), tz, "EEEE, MMMM d, yyyy");
  var genTs   = Utilities.formatDate(new Date(), tz, "MMM d, h:mm a");
  var stamp   = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var html    = _lsBuildHtml(d, dateStr, genTs);
  return Utilities.newBlob(html, "text/html", "lowstock.html")
    .getAs("application/pdf")
    .setName("HQ_Low_Stock_" + stamp + ".pdf");
}


/**
 * Save the low-stock PDF to Drive and hand back a link the sidebar can open.
 * Same delivery pattern as generateReportPdf() — a Drive link always works,
 * an in-iframe blob download does not.
 */
function generateLowStockPdf(days, threshold) {
  try {
    var blob = _lowStockPdfBlob({ days: days, threshold: threshold });
    var name = (typeof REPORT !== 'undefined' && REPORT.driveFolderName) || "HQ Weekly Reports";
    var it = DriveApp.getFoldersByName(name);
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
    var file = folder.createFile(blob);
    return { ok: true, name: file.getName(), url: file.getUrl() };
  } catch (e) {
    try { console.log("generateLowStockPdf error: " + e); } catch (_) {}
    return { ok: false, error: String(e) };
  }
}


/** Post the low-stock PDF to the admin Telegram chat. */
function sendLowStockToTelegram(days, threshold) {
  try {
    if (typeof TELEGRAM_ADMIN_CHAT_ID === 'undefined' || !TELEGRAM_ADMIN_CHAT_ID) {
      return { sent: false, error: "TELEGRAM_ADMIN_CHAT_ID not set" };
    }
    var d = analyzeLowStock({ days: days, threshold: threshold });
    var blob = _lowStockPdfBlob({ days: days, threshold: threshold });
    var payload = {
      chat_id: TELEGRAM_ADMIN_CHAT_ID,
      document: blob,
      caption: "📉 Low stock · " + d.count + " to reorder (last " + d.days + "d)"
    };
    var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendDocument", {
      method: "post", payload: payload, muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var ok = (code >= 200 && code < 300);
    return { sent: ok, count: d.count, error: ok ? "" : ("Telegram HTTP " + code) };
  } catch (e) {
    try { console.log("sendLowStockToTelegram error: " + e); } catch (_) {}
    return { sent: false, error: String(e) };
  }
}
