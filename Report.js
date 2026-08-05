// =======================================================================================
// Report.js — the "HQ Weekly Operations Report" PDF (the digest's deep-dive sibling)
// =======================================================================================
//
// The weekly Telegram DIGEST (WeeklyDigest.js) is the 10-second glance; THIS is the
// full report: an executive KPI band + the detail tables that make the numbers
// actionable — the underpriced kits leaking margin ("money on the table") and the
// kits that are in stock but can't be replenished (blocked). Rendered to a branded
// PDF via Apps Script's native HTML→PDF converter, delivered two ways:
//   • DOWNLOAD  — generateReportPdf() saves it to a Drive folder + returns the URL;
//                 the sidebar opens it in a new tab (reliable — HtmlService iframes
//                 can't always trigger a raw blob download, a Drive link always can).
//   • TELEGRAM  — sendReportToTelegram() posts it to the admin chat as a real file
//                 (sendDocument). runWeeklyDigest() attaches it to the Monday send.
//
// PURE TIER 1 — no new dependency. Reads the same snapshots the digest does; the
// PDF converter + Telegram bot are already ours.
//
// HTML→PDF CAVEAT: Google's converter is BASIC — no web fonts, no flexbox/grid. So
// the report is deliberately TABLE-based with inline styles + brand colors. Clean
// and reliable, iterated from a screenshot (not pixel-art).
// =======================================================================================

var REPORT = {
  topUnderpricedLimit: 12,   // how many underpriced kits to detail
  blockedLimit:        15,   // how many in-stock-but-blocked kits to detail
  driveFolderName:     "HQ Weekly Reports"
};


/**
 * Gather everything the report needs: the digest KPIs + two detail arrays read
 * from the Kit Health sheet (top underpriced kits by $ lost, and the in-stock-
 * but-blocked kits with their named bottleneck). Best-effort — a missing sheet
 * degrades the detail to [] and the KPIs to the digest's own safe defaults.
 */
function _gatherReportData() {
  var kpis = _gatherWeeklyDigest();   // reuse the digest's reader (counts + $)
  var topUnderpriced = [];
  var blocked = [];

  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(KIT_HEALTH.sheetName);
    if (sh) {
      var last = sh.getLastRow();
      var dsr = KIT_HEALTH.dataStartRow;
      if (last >= dsr) {
        var rows = sh.getRange(dsr, 1, last - dsr + 1, KIT_HEALTH.dataWidth).getValues();
        var iSku = KIT_HEALTH.idx("KIT_SKU"), iName = KIT_HEALTH.idx("KIT_NAME");
        var iListed = KIT_HEALTH.idx("LISTED"), iComp = KIT_HEALTH.idx("COMPUTED");
        var iDelta = KIT_HEALTH.idx("DELTA"), iDisc = KIT_HEALTH.idx("DISC_PCT");
        var iStatus = KIT_HEALTH.idx("PRICE_STATUS"), iStock = KIT_HEALTH.idx("STOCK_STATUS");
        var iLimit = KIT_HEALTH.idx("LIMITED_BY");

        for (var i = 0; i < rows.length; i++) {
          var r = rows[i];
          if (!r[iSku]) continue;
          if (r[iStatus] === KIT_HEALTH.status.UNDER) {
            var delta = (typeof r[iDelta] === 'number') ? r[iDelta] : 0;
            topUnderpriced.push({
              sku: r[iSku], name: r[iName] || "",
              listed: (typeof r[iListed] === 'number') ? r[iListed] : null,
              computed: (typeof r[iComp] === 'number') ? r[iComp] : null,
              underBy: Math.abs(delta),
              discPct: (typeof r[iDisc] === 'number') ? r[iDisc] : null
            });
          }
          if (r[iStock] === KIT_HEALTH.stock.IN_STOCK) {
            blocked.push({ sku: r[iSku], name: r[iName] || "", limitedBy: r[iLimit] || "" });
          }
        }
        topUnderpriced.sort(function (a, b) { return b.underBy - a.underBy; });
        topUnderpriced = topUnderpriced.slice(0, REPORT.topUnderpricedLimit);
        blocked = blocked.slice(0, REPORT.blockedLimit);
      }
    }
  } catch (e) { try { console.log("report.gather: " + e); } catch (_) {} }

  return { kpis: kpis, topUnderpriced: topUnderpriced, blocked: blocked };
}


/* ---- small HTML helpers (kept tiny; the converter is table+inline-style only) ---- */
function _rEsc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function _rMoney(n) {
  if (n == null || isNaN(n)) return "—";
  var v = Math.round(n);
  return "$" + String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
function _rPct(f) { return (f == null || isNaN(f)) ? "—" : (Math.round(f * 100) + "%"); }


/** One KPI cell for the executive band. */
function _kpiCell(value, label, tone) {
  var numColor = tone === 'bad' ? '#b71c1c' : (tone === 'good' ? '#1b5e20' : '#1a1a1a');
  return '<td width="33%" valign="top" style="border:1px solid #e0dccb;background:#fffdf5;padding:10px 12px;">' +
           '<div style="font-family:Arial,sans-serif;font-size:24px;font-weight:bold;color:' + numColor + ';">' + value + '</div>' +
           '<div style="font-family:Arial,sans-serif;font-size:10px;letter-spacing:1px;color:#6b6b6b;text-transform:uppercase;">' + label + '</div>' +
         '</td>';
}


/**
 * Build the branded report HTML (table-based, inline styles — survives the
 * HTML→PDF converter). PURE — Node-inspectable.
 */
function _buildReportHtml(data, dateStr, genTs) {
  var k = data.kpis;
  var ran = k.kitHealthRan;

  var html = '<div style="font-family:Arial,sans-serif;color:#1a1a1a;max-width:760px;">';

  // ---- Header band ----
  html +=
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:4px solid #ffd400;">' +
      '<tr>' +
        '<td valign="bottom" style="padding:0 0 8px 0;">' +
          '<span style="font-size:22px;font-weight:bold;letter-spacing:1px;">HQ MOTOR SERVICE</span><br>' +
          '<span style="font-size:12px;letter-spacing:3px;color:#6b6b6b;">WEEKLY OPERATIONS REPORT</span>' +
        '</td>' +
        '<td align="right" valign="bottom" style="padding:0 0 8px 0;">' +
          '<span style="font-size:13px;font-weight:bold;">' + _rEsc(dateStr) + '</span><br>' +
          '<span style="font-size:10px;color:#9a9a9a;">generated ' + _rEsc(genTs) + '</span>' +
        '</td>' +
      '</tr>' +
    '</table>';

  // ---- Headline hook ----
  if (ran && k.underpriced > 0) {
    html +=
      '<div style="background:#1a1a1a;color:#ffd400;padding:12px 14px;margin:14px 0;">' +
        '<span style="font-size:15px;font-weight:bold;">💸 ' + _rMoney(k.underpricedDollars) +
        ' of margin on the table</span> ' +
        '<span style="font-size:12px;color:#f0e6a8;">across ' + k.underpriced + ' underpriced kits</span>' +
      '</div>';
  }

  // ---- Executive KPI band (2 rows × 3) ----
  html += '<div style="font-size:11px;font-weight:bold;letter-spacing:2px;color:#6b6b6b;margin:6px 0;">AT A GLANCE</div>';
  html += '<table width="100%" cellpadding="0" cellspacing="4"><tr>';
  html += _kpiCell(ran ? String(k.underpriced) : "—", "Underpriced kits", ran && k.underpriced ? 'bad' : '');
  html += _kpiCell(ran ? String(k.buildableNow) : "—", "Buildable now", 'good');
  html += _kpiCell(ran ? String(k.inStockBlocked) : "—", "In stock, blocked", ran && k.inStockBlocked ? 'bad' : '');
  html += '</tr><tr>';
  html += _kpiCell(String(k.outOfStock), "Out of stock", k.outOfStock ? 'bad' : '');
  html += _kpiCell(String(k.openCases), "Open investigations", k.openCases ? 'bad' : '');
  html += _kpiCell(String(k.priceDrift), "eBay↔Zoho drift", k.priceDrift ? 'bad' : '');
  html += '</tr></table>';

  // ---- Section: money on the table ----
  html += '<div style="font-size:13px;font-weight:bold;letter-spacing:1px;margin:20px 0 6px;border-bottom:2px solid #ffd400;padding-bottom:3px;">💸 MONEY ON THE TABLE — TOP UNDERPRICED KITS</div>';
  if (!ran) {
    html += '<div style="font-size:12px;color:#8a8a8a;">Kit Health hasn\'t run yet — no snapshot to detail.</div>';
  } else if (data.topUnderpriced.length === 0) {
    html += '<div style="font-size:12px;color:#1b5e20;">✅ No underpriced kits vs your catalog norm. Nice.</div>';
  } else {
    html += '<table width="100%" cellpadding="6" cellspacing="0" style="font-size:11px;border-collapse:collapse;">';
    html += '<tr style="background:#1a1a1a;color:#ffffff;">' +
              '<th align="left">KIT</th><th align="left">NAME</th>' +
              '<th align="right">LISTED</th><th align="right">SHOULD BE</th>' +
              '<th align="right">UNDER BY</th><th align="right">DISC%</th></tr>';
    for (var i = 0; i < data.topUnderpriced.length; i++) {
      var u = data.topUnderpriced[i];
      var bg = (i % 2 === 0) ? '#ffffff' : '#fff8e7';
      html += '<tr style="background:' + bg + ';border-bottom:1px solid #eee;">' +
                '<td style="font-family:monospace;">' + _rEsc(u.sku) + '</td>' +
                '<td>' + _rEsc(String(u.name).substring(0, 42)) + '</td>' +
                '<td align="right">' + _rMoney(u.listed) + '</td>' +
                '<td align="right">' + _rMoney(u.computed) + '</td>' +
                '<td align="right" style="color:#b71c1c;font-weight:bold;">' + _rMoney(u.underBy) + '</td>' +
                '<td align="right">' + _rPct(u.discPct) + '</td>' +
              '</tr>';
    }
    html += '</table>';
    html += '<div style="font-size:10px;color:#9a9a9a;margin-top:4px;">SHOULD BE = computed at your catalog\'s median discount. DISC% = implied discount vs parts value — a high DISC% may be an intentional deep discount, not a leak.</div>';
  }

  // ---- Section: in stock but blocked ----
  html += '<div style="font-size:13px;font-weight:bold;letter-spacing:1px;margin:20px 0 6px;border-bottom:2px solid #ffd400;padding-bottom:3px;">⛔ IN STOCK BUT BLOCKED — CAN\'T REPLENISH</div>';
  if (!ran) {
    html += '<div style="font-size:12px;color:#8a8a8a;">Kit Health hasn\'t run yet.</div>';
  } else if (data.blocked.length === 0) {
    html += '<div style="font-size:12px;color:#1b5e20;">✅ Nothing blocked — every stocked kit can be rebuilt.</div>';
  } else {
    html += '<table width="100%" cellpadding="6" cellspacing="0" style="font-size:11px;border-collapse:collapse;">';
    html += '<tr style="background:#1a1a1a;color:#ffffff;"><th align="left">KIT</th><th align="left">NAME</th><th align="left">BLOCKED BY</th></tr>';
    for (var b = 0; b < data.blocked.length; b++) {
      var row = data.blocked[b];
      var bg2 = (b % 2 === 0) ? '#ffffff' : '#fff8e7';
      html += '<tr style="background:' + bg2 + ';border-bottom:1px solid #eee;">' +
                '<td style="font-family:monospace;">' + _rEsc(row.sku) + '</td>' +
                '<td>' + _rEsc(String(row.name).substring(0, 40)) + '</td>' +
                '<td>' + _rEsc(String(row.limitedBy).substring(0, 46)) + '</td>' +
              '</tr>';
    }
    html += '</table>';
  }

  // ---- Footer ----
  html += '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:22px;border-top:1px solid #ddd;">' +
            '<tr><td style="padding-top:6px;font-size:10px;color:#9a9a9a;">' +
              'HQ Motor Service · Order Management System · decision-support snapshot' +
            '</td><td align="right" style="padding-top:6px;font-size:10px;color:#9a9a9a;">' +
              'Out of stock ' + k.outOfStock + ' · Needs photos ' + k.needPhotos + ' · Open cases ' + k.openCases + ' · Price drift ' + k.priceDrift +
            '</td></tr></table>';

  html += '</div>';
  return html;
}


/** Build the report PDF blob (named with today's date). */
function _reportPdfBlob() {
  var data = _gatherReportData();
  var dateStr = Utilities.formatDate(new Date(), WEEKLY_DIGEST.timezone, "EEEE, MMMM d, yyyy");
  var genTs   = Utilities.formatDate(new Date(), WEEKLY_DIGEST.timezone, "MMM d, h:mm a");
  var html = _buildReportHtml(data, dateStr, genTs);
  var stamp = Utilities.formatDate(new Date(), WEEKLY_DIGEST.timezone, "yyyy-MM-dd");
  return Utilities.newBlob(html, "text/html", "report.html")
    .getAs("application/pdf")
    .setName("HQ_Weekly_Report_" + stamp + ".pdf");
}


/**
 * Generate the report PDF, save it to the Drive folder, and return the file URL
 * for the sidebar to open (reliable download path — a Drive link always works,
 * an in-iframe blob download does not). Returns {ok, name, url} or {ok:false}.
 */
function generateReportPdf() {
  try {
    var blob = _reportPdfBlob();
    var folder;
    var it = DriveApp.getFoldersByName(REPORT.driveFolderName);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(REPORT.driveFolderName);
    var file = folder.createFile(blob);
    return { ok: true, name: file.getName(), url: file.getUrl() };
  } catch (e) {
    try { console.log("generateReportPdf error: " + e); } catch (_) {}
    return { ok: false, error: String(e) };
  }
}


/**
 * Post the report PDF to the admin Telegram chat as a document. Optional caption.
 * Blob-in-payload → Apps Script sends multipart/form-data automatically.
 */
function sendReportToTelegram(caption) {
  try {
    if (typeof TELEGRAM_ADMIN_CHAT_ID === 'undefined' || !TELEGRAM_ADMIN_CHAT_ID) {
      return { sent: false, error: "TELEGRAM_ADMIN_CHAT_ID not set" };
    }
    var blob = _reportPdfBlob();
    var payload = { chat_id: TELEGRAM_ADMIN_CHAT_ID, document: blob };
    if (caption) payload.caption = caption;
    var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendDocument", {
      method: "post",
      payload: payload,             // NO contentType — let Apps Script build the multipart boundary
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var ok = (code >= 200 && code < 300);
    return { sent: ok, error: ok ? "" : ("Telegram HTTP " + code + ": " + res.getContentText()) };
  } catch (e) {
    try { console.log("sendReportToTelegram error: " + e); } catch (_) {}
    return { sent: false, error: String(e) };
  }
}
