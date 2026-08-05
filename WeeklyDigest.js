// =======================================================================================
// WeeklyDigest.js — Monday-morning ops digest to the admin Telegram chat
// =======================================================================================
//
// Reaches the user in Riyadh where a sidebar badge only helps when the sheet is
// open (roadmap #9). PURE TIER 1 — reads snapshots the system already maintains,
// sends via the Telegram bot we already own. No new data, no new dependency.
//
// WHAT IT READS (all best-effort — a missing sheet/getter degrades to 0, never
// throws, so one broken source can't kill the digest):
//   • Kit Health sheet   → underpriced kits (+ $ on the table), buildable-now,
//                           in-stock-but-blocked  (snapshot read, NOT a re-audit)
//   • Out of Stock       → getOutOfStockCount()
//   • Investigations     → getOpenCaseCount()
//   • Price Audit        → getPriceDriftCount()   (eBay↔Zoho drift)
//
// FRESHNESS: this is a READER, not an auditor — it reflects the last Kit Health /
// Price Audit run. Keep the Monday ~4am Kit Health + ~5am Price Audit weekly
// triggers installed (setupKitHealthTrigger / setupPriceAuditTrigger) so the
// ~8am digest reads fresh snapshots. Re-running the heavy audits here would risk
// the off-hours "anti-yank" rule (both rewrite + re-sort their whole sheet), so
// the digest deliberately does NOT re-audit.
//
// SEND SAFETY: previewWeeklyDigest() returns the exact text WITHOUT sending, so
// the message can be eyeballed before arming the weekly send. sendWeeklyDigest()
// is the only path that actually posts to Telegram.
//
// Message is PLAIN TEXT (no parse_mode) — same robustness choice as
// _sendKitParseAlert: numbers + emoji can never break Markdown/HTML parsing.
// =======================================================================================

var WEEKLY_DIGEST = {
  timezone: "America/Chicago",   // Houston shop time (matches the sheet clock)
  hour:     8                    // ~8am Monday (script-timezone)
};


/**
 * Gather every digest metric. Each source is isolated in its own try/catch so a
 * single failure degrades that line to 0 (or null for "sheet not run yet") and
 * the rest of the digest still sends.
 * @returns {{kitHealthRan, totalKits, underpriced, underpricedDollars,
 *            buildableNow, inStockBlocked, outOfStock, openCases, priceDrift}}
 */
function _gatherWeeklyDigest() {
  var d = {
    kitHealthRan: false,
    totalKits: 0,
    underpriced: 0,
    underpricedDollars: 0,
    buildableNow: 0,
    inStockBlocked: 0,
    outOfStock: 0,
    openCases: 0,
    priceDrift: 0,
    needPhotos: 0
  };

  // --- Kit Health (snapshot read of the sheet) ---
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(KIT_HEALTH.sheetName);
    if (sh) {
      var last = sh.getLastRow();
      var dsr = KIT_HEALTH.dataStartRow;
      if (last >= dsr) {
        var rows = sh.getRange(dsr, 1, last - dsr + 1, KIT_HEALTH.dataWidth).getValues();
        var iStatus = KIT_HEALTH.idx("PRICE_STATUS");
        var iDelta  = KIT_HEALTH.idx("DELTA");
        var iBuild  = KIT_HEALTH.idx("BUILDABLE");
        var iStock  = KIT_HEALTH.idx("STOCK_STATUS");
        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          if (!r[KIT_HEALTH.idx("KIT_SKU")]) continue;   // skip any blank row
          d.totalKits++;
          if (r[iStatus] === KIT_HEALTH.status.UNDER) {
            d.underpriced++;
            var delta = r[iDelta];
            if (typeof delta === 'number' && !isNaN(delta)) d.underpricedDollars += Math.abs(delta);
          }
          if (typeof r[iBuild] === 'number' && r[iBuild] > 0) d.buildableNow++;
          if (r[iStock] === KIT_HEALTH.stock.IN_STOCK) d.inStockBlocked++;
        }
        d.kitHealthRan = true;
        d.underpricedDollars = Math.round(d.underpricedDollars);
      }
    }
  } catch (e) { try { console.log("digest.kitHealth: " + e); } catch (_) {} }

  try { d.outOfStock = getOutOfStockCount() || 0; } catch (e) { try { console.log("digest.oos: " + e); } catch (_) {} }
  try { d.openCases  = getOpenCaseCount()   || 0; } catch (e) { try { console.log("digest.cases: " + e); } catch (_) {} }
  try { d.priceDrift = getPriceDriftCount() || 0; } catch (e) { try { console.log("digest.drift: " + e); } catch (_) {} }
  try { d.needPhotos = getPhotoQueueCount() || 0; } catch (e) { try { console.log("digest.photos: " + e); } catch (_) {} }

  return d;
}


/**
 * Compose the digest text from a gathered-data object. PURE (no I/O) so it's
 * Node-testable and previewable. Plain text, no parse_mode.
 */
function _buildWeeklyDigestText(d, dateStr) {
  var L = [];
  L.push("📊 HQ WEEKLY OPS DIGEST · " + dateStr);
  L.push("");

  // KITS
  L.push("KITS");
  if (d.kitHealthRan) {
    var dollars = d.underpricedDollars > 0 ? "  (~$" + d.underpricedDollars + " on the table)" : "";
    L.push("💸 Underpriced: " + d.underpriced + dollars);
    L.push("🔧 Buildable now: " + d.buildableNow);
    L.push("⛔ In stock but blocked: " + d.inStockBlocked);
  } else {
    L.push("· Kit Health hasn't run yet — no snapshot to report");
  }
  L.push("");

  // INVENTORY
  L.push("INVENTORY");
  L.push("📦 Out of stock: " + d.outOfStock + (d.outOfStock === 1 ? " item" : " items"));
  L.push("📸 Needs photos: " + d.needPhotos + (d.needPhotos === 1 ? " item" : " items"));
  L.push("");

  // ORDERS / QUALITY
  L.push("ORDERS");
  L.push("⚠ Open investigations: " + d.openCases);
  L.push("📉 eBay↔Zoho price drift: " + d.priceDrift);
  L.push("");

  // one-line "all clear" celebration when nothing needs attention
  var anythingHot = d.underpriced || d.inStockBlocked || d.outOfStock || d.openCases || d.priceDrift || d.needPhotos;
  L.push(anythingHot
    ? "— open the sheets for the full breakdown —"
    : "✅ All clear — nothing flagged this week.");

  return L.join("\n");
}


/**
 * Build the digest text for RIGHT NOW without sending. Safe to run from the
 * editor to eyeball the message before arming the weekly send. Also logged.
 * @returns {string} the message text
 */
function previewWeeklyDigest() {
  var d = _gatherWeeklyDigest();
  var dateStr = Utilities.formatDate(new Date(), WEEKLY_DIGEST.timezone, "EEE MMM d");
  var text = _buildWeeklyDigestText(d, dateStr);
  try { console.log(text); } catch (_) {}
  return text;
}


/**
 * Gather → build → POST to the admin Telegram chat. The ONLY path that actually
 * sends. Returns {sent, text, error}. Best-effort HTTP (muteHttpExceptions).
 */
function sendWeeklyDigest() {
  var text = previewWeeklyDigest();   // reuse the gather+build (also logs)
  try {
    if (typeof TELEGRAM_ADMIN_CHAT_ID === 'undefined' || !TELEGRAM_ADMIN_CHAT_ID) {
      return { sent: false, text: text, error: "TELEGRAM_ADMIN_CHAT_ID not set" };
    }
    var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: TELEGRAM_ADMIN_CHAT_ID, text: text }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var ok = (code >= 200 && code < 300);
    return { sent: ok, text: text, error: ok ? "" : ("Telegram HTTP " + code + ": " + res.getContentText()) };
  } catch (e) {
    try { console.log("sendWeeklyDigest error: " + e); } catch (_) {}
    return { sent: false, text: text, error: String(e) };
  }
}


/**
 * Weekly-trigger target. Pure send (the Mon ~4am Kit Health + ~5am Price Audit
 * triggers keep the snapshots fresh; the digest does NOT re-audit — see header).
 */
function runWeeklyDigest() {
  var r = sendWeeklyDigest();
  if (!r.sent) { try { console.log("runWeeklyDigest not sent: " + r.error); } catch (_) {} }
  // Attach the full PDF report right after the text glance (best-effort — a PDF
  // failure must never block the text digest that already sent).
  try {
    if (typeof sendReportToTelegram === 'function') {
      var pr = sendReportToTelegram("📄 HQ Weekly Operations Report");
      if (!pr.sent) { try { console.log("runWeeklyDigest report not sent: " + pr.error); } catch (_) {} }
    }
  } catch (e) { try { console.log("runWeeklyDigest report error: " + e); } catch (_) {} }
  return r;
}


/**
 * Install the Monday ~8am weekly trigger (script timezone). Idempotent — removes
 * any prior WeeklyDigest trigger first. Run ONCE from the editor.
 */
function setupWeeklyDigestTrigger() {
  removeWeeklyDigestTrigger();
  ScriptApp.newTrigger("runWeeklyDigest")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(WEEKLY_DIGEST.hour)
    .create();
  return "Weekly digest armed — Monday ~" + WEEKLY_DIGEST.hour + ":00 (script timezone).";
}

/** Remove the weekly digest trigger. */
function removeWeeklyDigestTrigger() {
  var removed = 0;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runWeeklyDigest") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return "Removed " + removed + " weekly digest trigger(s).";
}
