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

  // Restock ripple — which parts to buy first. Best-effort: a failure here
  // degrades the report to what it was before, never blocks it.
  var ripple = null;
  try { ripple = analyzeRestockRipple(); }
  catch (e) { try { console.log("report.ripple: " + e); } catch (_) {} }

  // Week-over-week movement. null until KPI History has a row from a prior day.
  var trend = null;
  try { trend = getKpiTrend(kpis); }
  catch (e) { try { console.log("report.trend: " + e); } catch (_) {} }

  return {
    kpis: kpis, topUnderpriced: topUnderpriced, blocked: blocked,
    ripple: ripple, trend: trend
  };
}


// =======================================================================================
// "THIS WEEK'S 3 MOVES" — the decision box
// =======================================================================================
//
// The whole point of this report, per the strategic direction set 2026-07-31:
// eBay's and Zoho's reports are CHANNEL and FINANCIAL views. This one is the
// IN-FIELD view, and a manager should be able to read the top of page 1 and
// know what to DO — not just what is true.
//
// Each move carries a $ or a count so it can be ranked against the others.
// PURE — no I/O, so the ranking logic stays Node-testable.

/**
 * Pick the highest-impact actions from the gathered data, best first.
 * @returns {Array<{title:string, detail:string, impact:string}>} 0–3 moves
 */
function _reportMoves(data) {
  var k = data.kpis || {};
  var moves = [];

  // 1. Restock — the only move that unblocks physical throughput.
  var rip = data.ripple;
  if (rip && rip.planFrees > 0) {
    moves.push({
      weight: 1000 + rip.planFrees,
      title:  "Restock " + rip.planParts.length + " parts",
      detail: rip.planParts.join(" · "),
      impact: "unblocks " + rip.planFrees + " of " + rip.blockedCount + " blocked kits"
    });
  }

  // 2. Reprice — money already on the table, no purchase required.
  if (k.kitHealthRan && k.underpricedDollars > 0) {
    moves.push({
      weight: 900,
      title:  "Reprice " + k.underpriced + " underpriced kits",
      detail: "listed below what their own parts cost, at the catalog's median discount",
      impact: _rMoney(k.underpricedDollars) + " of margin recoverable"
    });
  }

  // 3. Whatever is deteriorating fastest — only when history proves it moved.
  var t = data.trend;
  if (t && t.deltas) {
    if (t.deltas.inStockBlocked > 0) {
      moves.push({
        weight: 800 + t.deltas.inStockBlocked,
        title:  "Blocked kits are climbing",
        detail: "+" + t.deltas.inStockBlocked + " in the last " + t.daysAgo + " days" +
                " (now " + k.inStockBlocked + " on the shelf but not replenishable)",
        impact: "shared components running dry"
      });
    } else if (t.deltas.outOfStock > 0) {
      moves.push({
        weight: 700 + t.deltas.outOfStock,
        title:  "Out-of-stock list is growing",
        detail: "+" + t.deltas.outOfStock + " in the last " + t.daysAgo + " days",
        impact: k.outOfStock + " items to reorder"
      });
    }
  }

  // 4. Fallbacks so the box is never empty on a healthy week.
  if (k.openCases > 0) {
    moves.push({
      weight: 600, title: "Close " + k.openCases + " open investigation" + (k.openCases === 1 ? "" : "s"),
      detail: "orders with an unresolved finding", impact: "customer-facing"
    });
  }
  if (k.needPhotos > 0) {
    moves.push({
      weight: 500, title: "Shoot " + k.needPhotos + " listings still on the logo",
      detail: "active listings with one image or none", impact: "conversion drag"
    });
  }
  if (k.priceDrift > 0) {
    moves.push({
      weight: 400, title: "Resolve " + k.priceDrift + " eBay↔Zoho price difference" + (k.priceDrift === 1 ? "" : "s"),
      detail: "quotes are made off Zoho's number", impact: "mis-quote risk"
    });
  }

  moves.sort(function (a, b) { return b.weight - a.weight; });
  return moves.slice(0, 3);
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


/**
 * Week-over-week movement for one KPI, as a coloured HTML fragment.
 * Direction is coloured by MEANING, not by sign — `goodWhenUp` comes from
 * KPI_HISTORY.fieldMap, so "buildable ▲" reads green while "blocked ▲" reads
 * red. Returns "" when there's no history yet or nothing moved, so a first-ever
 * report is byte-identical to the pre-trend one.
 */
function _rTrend(trend, field) {
  if (!trend || !trend.deltas) return "";
  var delta = trend.deltas[field];
  if (typeof delta !== 'number' || isNaN(delta) || delta === 0) return "";

  var good = true;
  for (var i = 0; i < KPI_HISTORY.fieldMap.length; i++) {
    if (KPI_HISTORY.fieldMap[i].field === field) { good = KPI_HISTORY.fieldMap[i].goodWhenUp; break; }
  }
  var up = delta > 0;
  var color = (up === good) ? '#1b5e20' : '#b71c1c';
  return '<span style="font-size:11px;font-weight:bold;color:' + color + ';">&nbsp;' +
         (up ? '▲' : '▼') + ' ' + Math.abs(delta) + '</span>';
}


/** One KPI cell for the executive band. `trendHtml` is optional. */
function _kpiCell(value, label, tone, trendHtml) {
  var numColor = tone === 'bad' ? '#b71c1c' : (tone === 'good' ? '#1b5e20' : '#1a1a1a');
  return '<td width="33%" valign="top" style="border:1px solid #e0dccb;background:#fffdf5;padding:10px 12px;">' +
           '<div style="font-family:Arial,sans-serif;font-size:24px;font-weight:bold;color:' + numColor + ';">' +
             value + (trendHtml || "") +
           '</div>' +
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
  // ---- THIS WEEK'S 3 MOVES — the decision box ----
  // Deliberately ABOVE the KPI band: a manager reading only the top of page 1
  // should get instructions, not statistics.
  var moves = _reportMoves(data);
  if (moves.length) {
    html += '<div style="border:2px solid #1a1a1a;margin:16px 0;">' +
              '<div style="background:#1a1a1a;color:#ffd400;padding:7px 12px;font-size:12px;' +
              'font-weight:bold;letter-spacing:2px;">▌ THIS WEEK\'S ' + moves.length + ' MOVES</div>' +
              '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fffdf5;">';
    for (var mi = 0; mi < moves.length; mi++) {
      var mv = moves[mi];
      html += '<tr>' +
                '<td width="34" valign="top" align="center" style="padding:10px 0 10px 10px;' +
                  'font-size:20px;font-weight:bold;color:#c9a227;">' + (mi + 1) + '</td>' +
                '<td valign="top" style="padding:10px 12px 10px 4px;' +
                  (mi < moves.length - 1 ? 'border-bottom:1px solid #ece7d5;' : '') + '">' +
                  '<div style="font-size:13px;font-weight:bold;color:#1a1a1a;">' + _rEsc(mv.title) + '</div>' +
                  '<div style="font-size:11px;color:#6b6b6b;padding-top:2px;">' + _rEsc(mv.detail) + '</div>' +
                  '<div style="font-size:11px;font-weight:bold;color:#1b5e20;padding-top:3px;">→ ' +
                    _rEsc(mv.impact) + '</div>' +
                '</td>' +
              '</tr>';
    }
    html += '</table></div>';
  }

  var tr = data.trend;
  html += '<div style="font-size:11px;font-weight:bold;letter-spacing:2px;color:#6b6b6b;margin:6px 0;">AT A GLANCE' +
          (tr ? '<span style="font-weight:normal;letter-spacing:0;text-transform:none;">' +
                '  ·  ▲▼ vs ' + tr.daysAgo + ' days ago</span>' : '') + '</div>';
  html += '<table width="100%" cellpadding="0" cellspacing="4"><tr>';
  html += _kpiCell(ran ? String(k.underpriced) : "—", "Underpriced kits", ran && k.underpriced ? 'bad' : '', _rTrend(tr, 'underpriced'));
  html += _kpiCell(ran ? String(k.buildableNow) : "—", "Buildable now", 'good', _rTrend(tr, 'buildableNow'));
  html += _kpiCell(ran ? String(k.inStockBlocked) : "—", "In stock, blocked", ran && k.inStockBlocked ? 'bad' : '', _rTrend(tr, 'inStockBlocked'));
  html += '</tr><tr>';
  html += _kpiCell(String(k.outOfStock), "Out of stock", k.outOfStock ? 'bad' : '', _rTrend(tr, 'outOfStock'));
  html += _kpiCell(String(k.openCases), "Open investigations", k.openCases ? 'bad' : '', _rTrend(tr, 'openCases'));
  html += _kpiCell(String(k.priceDrift), "eBay↔Zoho drift", k.priceDrift ? 'bad' : '', _rTrend(tr, 'priceDrift'));
  html += '</tr></table>';

  // ---- Section: restock ripple ----
  var rip = data.ripple;
  if (rip && rip.parts && rip.parts.length) {
    html += '<div style="font-size:13px;font-weight:bold;letter-spacing:1px;margin:20px 0 6px;border-bottom:2px solid #ffd400;padding-bottom:3px;">🔗 RESTOCK RIPPLE — WHAT ONE ORDER UNBLOCKS</div>';
    html += '<div style="font-size:11px;color:#6b6b6b;margin-bottom:6px;">' +
            '<strong>Frees alone</strong> = kits where this is the ONLY part missing, so restocking it is enough. ' +
            '<strong>Short in</strong> = kits where it is one of several missing.</div>';
    html += '<table width="100%" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-size:11px;">' +
              '<tr style="background:#1a1a1a;color:#ffffff;text-align:left;">' +
                '<th>PART</th><th>SHELF</th><th align="right">ON HAND</th>' +
                '<th align="right">FREES ALONE</th><th align="right">SHORT IN</th><th>NAME</th>' +
              '</tr>';
    var rn = Math.min(rip.parts.length, 10);
    for (var ri = 0; ri < rn; ri++) {
      var p = rip.parts[ri];
      html += '<tr style="background:' + (ri % 2 ? '#fffdf5' : '#ffffff') + ';border-bottom:1px solid #ece7d5;">' +
                '<td><strong>' + _rEsc(p.sku) + '</strong></td>' +
                '<td>' + _rEsc(p.location || "—") + '</td>' +
                '<td align="right">' + _rEsc(String(p.avail)) + '</td>' +
                '<td align="right" style="font-weight:bold;color:' + (p.sole ? '#1b5e20' : '#9a9a9a') + ';">' + p.sole + '</td>' +
                '<td align="right">' + p.shortIn + '</td>' +
                '<td>' + _rEsc(String(p.name).substring(0, 34)) + '</td>' +
              '</tr>';
    }
    html += '</table>';
    if (rip.multiBlocked > 0) {
      html += '<div style="font-size:11px;color:#6b6b6b;margin-top:5px;">' +
              rip.multiBlocked + ' of ' + rip.blockedCount +
              ' blocked kits need more than one part restocked — no single order frees them.</div>';
    }
  }

  // ---- Section: kits nothing can assess ----
  // A blind spot in every surface at once — Kit Health can't price these, OOS
  // can't compute buildable, the ripple can't rank them. Named, with the
  // specific cause, because each of the three needs a different fix in Zoho.
  if (rip && rip.unreadable && rip.unreadable.length) {
    html += '<div style="font-size:13px;font-weight:bold;letter-spacing:1px;margin:20px 0 6px;border-bottom:2px solid #ffd400;padding-bottom:3px;">⚠ KITS NOTHING CAN ASSESS (' + rip.unreadable.length + ')</div>';
    html += '<div style="font-size:11px;color:#6b6b6b;margin-bottom:6px;">' +
            'Invisible to pricing, buildability and the ripple until their Purchase Description is fixed in Zoho. ' +
            'The registry re-parses on save.</div>';
    html += '<table width="100%" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-size:11px;">' +
              '<tr style="background:#1a1a1a;color:#ffffff;text-align:left;">' +
                '<th>KIT</th><th>NAME</th><th>WHY</th>' +
              '</tr>';
    var un = rip.unreadable.slice(0, 12);
    for (var ui = 0; ui < un.length; ui++) {
      var u = un[ui];
      html += '<tr style="background:' + (ui % 2 ? '#fffdf5' : '#ffffff') + ';border-bottom:1px solid #ece7d5;">' +
                '<td><strong>' + _rEsc(u.sku) + '</strong></td>' +
                '<td>' + _rEsc(String(u.name).substring(0, 32)) + '</td>' +
                '<td>' + _rEsc(String(u.reason).substring(0, 46)) +
                  (u.raw ? '<br><span style="color:#8a8a8a;">&ldquo;' +
                           _rEsc(String(u.raw).substring(0, 44)) + '&rdquo;</span>' : '') +
                '</td>' +
              '</tr>';
    }
    html += '</table>';
    if (rip.unreadable.length > un.length) {
      html += '<div style="font-size:11px;color:#6b6b6b;margin-top:4px;">… +' +
              (rip.unreadable.length - un.length) + ' more</div>';
    }
  }

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
