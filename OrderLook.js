/**
 * OrderLook.js — the quieter DIRECT table (2026-09-30, step 1 of the order-box round).
 *
 * Agreed with the user on the mockup "Direct Order Boxes" after rejecting the busier
 * rounds (three box weights, aisle lines): the table should get CALMER, not louder.
 * Four changes, all display only — no cell value, row, width or n8n contract changes:
 *
 *   1. A FINISHED order (every line SHIPPED or CANCELED) — its box drops from gold to a
 *      thin pale line (RowManagement._paintDirectOrderDividers), and its text goes quiet
 *      (the "done" colour rule below). Live orders keep today's gold box.
 *      ⭐ 2026-10-03: the quiet text now reaches the eBay table too (no boxes there).
 *   2. After an order's first line, the SALES ORDER text goes grey. The keycap badge is
 *      an emoji and keeps its colour, so every line still says which order it is.
 *   3. A note that only repeats the order's first note goes grey. Kit notes (↳), Zoho
 *      flags (⚠) and anything containing the whole word HOLD never do — the same
 *      whole-word rule the HOLD alarms use.
 *   4. The DIRECT band's nameplate ends "· N ready" — orders with every line picked and
 *      none waiting (__SparkData!A31). Hidden at 0.
 *
 *   PART 2 (2026-10-01) — kits:
 *   5. K TAGS. Within one order, every EXPANDED kit is numbered K1, K2… in sheet order.
 *      The parent's SKU cell shows "▣ K1 157644", each of its parts "K1 164979" — display
 *      only, the ▣ number-format trick; the cell still holds the plain SKU. An unexpanded
 *      kit keeps today's plain "▣". _kitTagPlan() is the ONE numbering rule; the printed
 *      pick list (part 3) must call it too, so a part is K3 on screen and K3 on paper.
 *   6. KIT READY. The parent's SKU cell turns green once no part of that kit is still
 *      PENDING (and the parent itself is live). A live colour rule — flips on the pick.
 *
 * ⭐ ONE SWITCH. setupOrderLook() turns it on; removeOrderLook() is the whole rollback.
 *   The switch is a Script Property because the box painter runs on every edit and must
 *   know, cheaply, whether to draw the pale state.
 *
 * ⚠⚠ THE COLOUR RULES GO LAST, ON PURPOSE. Sheets applies the FIRST matching rule per
 *   cell, so appending means every existing signal wins where they overlap: the identity
 *   red, the HOLD rules, the buyer-note italic, the HAND low-stock red. These rules can
 *   only ever QUIET a cell nothing else has claimed.
 *
 * ⚠⚠ SAFE FROM THE STRIPPERS — checked against all of them (2026-09-30):
 *   · _applyAllConditionalFormatting strips single-column rules on A/E/F/G/J. The "done"
 *     rule spans A:E + G:J (never one column) and survives; the SALES ORDER rule is on D
 *     (not a theme column) and survives; the NOTE rule is single-column E and WOULD be
 *     stripped — so that function rebuilds these rules itself when the switch is on.
 *   · removeLegacySalesOrderCFRules strips col-D-only rules containing COUNTIF or TRIM( —
 *     the SALES ORDER rule contains neither.
 *   · removeDuplicateHighlightRules strips col-A-only rules; the HAND stripper col-G-only
 *     rules. Nothing here is single-column A or G.
 *   · Every rule carries N("hq-look") so this file can find and remove exactly its own.
 *
 * ⚠ STATUS (F) is never in any range: the first-match rule would otherwise hide the
 *   SHIPPED green behind a grey font rule.
 */

var ORDER_LOOK = {
  prop:      'ORDER_LOOK_ON',
  marker:    'hq-look',
  quiet:     '#a39b86',   // the mockup's quiet ink
  paleBox:   '#ddd3b0',   // a finished order's box
  readyCell: 'A31'        // __SparkData — orders fully picked, none waiting
};

/** Cheap: one Script Property read. The box painter calls this on every repaint. */
function _orderLookOn() {
  try { return PropertiesService.getScriptProperties().getProperty(ORDER_LOOK.prop) === 'on'; }
  catch (e) { return false; }
}

/** The three formulas, pure (Node-testable). Row r is the first data row of the ranges. */
function _orderLookFormulas(r) {
  var tag = 'N("' + ORDER_LOOK.marker + '")=0';
  // Below the DIRECT divider's header: the look is for the DIRECT (and AMAZON) tables only.
  var below = 'ROW()>MATCH("' + Schema.boundaryMarker + '",$A$1:$A,0)+1';
  var prev = r - 1;
  return {
    // A line of a finished order: its own status is terminal AND no line of the order is open.
    // ⭐ 2026-10-03: ALL THREE TABLES (the owner's call — "so all feel the same"). Unlike the
    //   SO-grey and repeated-note rules below, this one never assumes an order's lines are
    //   contiguous: the COUNTIFS look across the whole column, so it is just as right on the
    //   aisle-sorted eBay table. eBay gets the quiet text only — it has no order boxes.
    done: '=AND(' + tag + ',$D' + r + '<>"",' +
          'OR($F' + r + '="SHIPPED",$F' + r + '="CANCELED"),' +
          'COUNTIFS($D$' + Schema.dataStartRow + ':$D,$D' + r + ',$F$' + Schema.dataStartRow + ':$F,"PENDING")+' +
          'COUNTIFS($D$' + Schema.dataStartRow + ':$D,$D' + r + ',$F$' + Schema.dataStartRow + ':$F,"PREPARING")=0)',
    // Every line after an order's first — orders are contiguous under the DIRECT sort.
    so:   '=AND(' + tag + ',' + below + ',$D' + r + '<>"",$D' + r + '=$D' + prev + ')',
    // A note that only repeats the order's FIRST line's note. MATCH finds that first line.
    note: '=AND(' + tag + ',' + below + ',$E' + r + '<>"",$D' + r + '<>"",$D' + r + '=$D' + prev + ',' +
          '$E' + r + '=INDEX($E:$E,MATCH($D' + r + ',$D:$D,0)),' +
          'ISERROR(FIND("↳",$E' + r + '&"")),LEFT($E' + r + ',1)<>"⚠",' +
          'NOT(REGEXMATCH($E' + r + '&"","(?i)\\bhold\\b")))',
    // A kit PARENT whose parts are all picked. A part's note reads "↳ from KIT-<sku>…"
    // (or "added to", or after a ⚠ Zoho flag line) — matched as "*↳ * KIT-<sku>" alone or
    // followed by a space, so KIT-1586 can never match KIT-158652's parts.
    kit:  (function () {
      var D = '$D$' + Schema.dataStartRow + ':$D', E = '$E$' + Schema.dataStartRow + ':$E',
          Fs = '$F$' + Schema.dataStartRow + ':$F';
      var p1 = '"*↳ * KIT-"&$A' + r, p2 = '"*↳ * KIT-"&$A' + r + '&" *"';
      return '=AND(' + tag + ',$A' + r + '<>"",$D' + r + '<>"",' +
             '$F' + r + '<>"SHIPPED",$F' + r + '<>"CANCELED",ISERROR(FIND("↳",$E' + r + '&"")),' +
             'COUNTIFS(' + D + ',$D' + r + ',' + E + ',' + p1 + ')+COUNTIFS(' + D + ',$D' + r + ',' + E + ',' + p2 + ')>0,' +
             'COUNTIFS(' + D + ',$D' + r + ',' + E + ',' + p1 + ',' + Fs + ',"PENDING")+' +
             'COUNTIFS(' + D + ',$D' + r + ',' + E + ',' + p2 + ',' + Fs + ',"PENDING")=0)';
    })()
  };
}

/**
 * ⭐ THE ONE K-NUMBERING RULE (sheet now, print in part 3). Pure — Node-testable.
 *
 * @param {Array<{sku,so,note}>} rows  sheet order
 * @param {Set} kitSkus  registered kit SKUs, UPPERCASE
 * @returns {Array<{k:number, parent:boolean}|null>}  per row: its kit number, whether it is
 *   the parent line, or null (not part of an expanded kit)
 *
 * A kit is numbered only when it is EXPANDED in that order (a part's note names it).
 * ⚠ NUMBERED BY THE KIT'S OWN SKU, LOWEST FIRST — never by row position. The DIRECT sort
 *   reorders an order's lines when statuses change, so position-numbering would turn K2
 *   into K3 after a sort and disagree with a pick list printed ten minutes earlier. The
 *   SKU order never moves while the same kits are on the order.
 * A part whose parent line is not on the sheet gets no number — a tag pointing at nothing
 * would send someone looking.
 */
function _kitTagPlan(rows, kitSkus, prior) {
  // ⚠⚠ 2026-10-09 — K NUMBERS ARE UNIQUE ACROSS THE WHOLE SHEET, AND STICKY.
  //   This used to number kits PER ORDER, so every order's first kit was "K1". Two kit
  //   orders on one table (the floor's report: 157563 on 12-15269-54269 and 159093 on
  //   06-15279-95092) both read K1, and on the aisle-sorted eBay table their parts
  //   interleave — a picker could not tell which part belongs to which box.
  //   Now: one number per KIT (order + kit SKU), never shared with any other kit on the
  //   sheet. A kit keeps its number for as long as it is on the sheet (`prior`, the map
  //   the last run assigned — see _kitTagPlanSticky), so paper printed an hour ago still
  //   matches; a new kit takes the LOWEST FREE number; a number frees when its kit leaves.
  //   With no prior map, numbering is deterministic: order id, then kit SKU lowest first.
  prior = prior || {};
  var partsOf = {};                                   // "SO|PARENTSKU" → true
  rows.forEach(function (r) {
    var p = kitComponentTag(r.note);
    if (p && r.so) partsOf[String(r.so).trim() + '|' + String(p).toUpperCase()] = true;
  });
  function isParent(r) {
    var so = String(r.so || '').trim(), sku = String(r.sku || '').trim().toUpperCase();
    return so && sku && !kitComponentTag(r.note) && kitSkus.has(sku) && partsOf[so + '|' + sku];
  }
  var keys = [], seen = {};
  rows.forEach(function (r) {
    if (!isParent(r)) return;
    var key = String(r.so).trim() + '|' + String(r.sku).trim().toUpperCase();
    if (!seen[key]) { seen[key] = true; keys.push(key); }
  });
  var numOf = {}, used = {};
  keys.forEach(function (k) {                         // keep every number already given
    var n = Math.floor(Number(prior[k]));
    if (n >= 1 && !used[n]) { numOf[k] = n; used[n] = true; }
  });
  keys.filter(function (k) { return !(k in numOf); }).sort(function (a, b) {
    var pa = a.split('|'), pb = b.split('|');
    if (pa[0] !== pb[0]) return pa[0] < pb[0] ? -1 : 1;
    var na = Number(pa[1]), nb = Number(pb[1]);
    return (isFinite(na) && isFinite(nb)) ? na - nb : (pa[1] < pb[1] ? -1 : pa[1] > pb[1] ? 1 : 0);
  }).forEach(function (k) {                            // new kits: lowest free number
    var n = 1; while (used[n]) n++;
    numOf[k] = n; used[n] = true;
  });
  var out = rows.map(function (r) {
    if (!isParent(r)) return null;
    return { k: numOf[String(r.so).trim() + '|' + String(r.sku).trim().toUpperCase()], parent: true };
  });
  rows.forEach(function (r, i) {
    var p = kitComponentTag(r.note);
    if (!p) return;
    var key = String(r.so || '').trim() + '|' + String(p).toUpperCase();
    if (key in numOf) out[i] = { k: numOf[key], parent: false };
  });
  out.numbers = numOf;
  return out;
}

/** The K-number memory. Document property, not cell formats: a new row INHERITS its
 *  neighbour's number format on insert, so a format-based memory would let a brand-new
 *  kit steal an old kit's number. The sheet's marker refresh and the print both go
 *  through here, so paper and sheet always agree. */
var KIT_TAG_PROP = 'KIT_TAG_NUMBERS';
function _kitTagPlanSticky(rows, kitSkus) {
  var props = null, prior = {};
  try {
    props = PropertiesService.getDocumentProperties();
    prior = JSON.parse(props.getProperty(KIT_TAG_PROP) || '{}') || {};
  } catch (e) { prior = {}; }
  var plan = _kitTagPlan(rows, kitSkus, prior);
  try {
    var next = JSON.stringify(plan.numbers);
    if (props && next !== JSON.stringify(prior)) props.setProperty(KIT_TAG_PROP, next);
  } catch (e) { console.log('kit tag memory not saved: ' + e); }
  return plan;
}

/** The number format a SKU cell wears for a plan entry (null = decide as before). */
function _kitTagFormat(entry) {
  if (!entry) return null;
  return entry.parent ? '"▣ K' + entry.k + ' "@' : '"K' + entry.k + ' "@';
}

/** The rules, in the order they are appended (all LAST — see the header). */
function _buildOrderLookRules(sheet) {
  var r0 = Schema.dataStartRow;
  var n = BRAND.dataLast - Schema.bannerRows;
  var F = _orderLookFormulas(r0);
  function rule(formula, ranges) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setFontColor(ORDER_LOOK.quiet)
      .setRanges(ranges).build();
  }
  var W = Schema.dataWidth;
  var doneRanges = [
    sheet.getRange(r0, 1, n, Schema.cols.STATUS - 1),                         // A:E
    sheet.getRange(r0, Schema.cols.STATUS + 1, n, W - Schema.cols.STATUS)      // G:J
  ];
  var kitReady = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(F.kit)
    .setBackground('#c8e6c9').setFontColor('#1b5e20').setBold(true)
    .setRanges([sheet.getRange(r0, Schema.cols.SKU, n, 1)]).build();
  return [
    kitReady,
    rule(F.done, doneRanges),
    rule(F.so,   [sheet.getRange(r0, Schema.cols.SALES_ORDER, n, 1)]),
    rule(F.note, [sheet.getRange(r0, Schema.cols.NOTE, n, 1)])
  ];
}

/** Drop this file's own rules (found by the N("hq-look") marker), keep everything else. */
function _stripOrderLookRules(rules) {
  return rules.filter(function (rl) {
    var bc = rl.getBooleanCondition();
    if (!bc || bc.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    var v = bc.getCriteriaValues();
    return !(v.length && String(v[0]).indexOf(ORDER_LOOK.marker) !== -1);
  });
}

/**
 * Turn the quieter DIRECT table ON. Idempotent. Owner-only (it rewrites colour rules and
 * the band on a locked sheet). Returns a readable report, including what the band reads.
 */
function setupOrderLook() {
  if (typeof _obRequireOwner === 'function') {
    var denied = _obRequireOwner('Turning on the quieter DIRECT table');
    if (denied) return denied;
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ No ' + MAIN_SHEET_NAME + ' sheet.';

  PropertiesService.getScriptProperties().setProperty(ORDER_LOOK.prop, 'on');
  var out = [];

  var rules = _stripOrderLookRules(sheet.getConditionalFormatRules());
  sheet.setConditionalFormatRules(rules.concat(_buildOrderLookRules(sheet)));
  out.push('✓ colour rules installed (finished orders · grey SO# · repeated notes · kit ready)');
  try { out.push(refreshKitSkuMarkers()); } catch (e) { out.push('✗ kit tags: ' + e); }

  try { _ensureSparkData(ss); out.push('✓ ready count formula (__SparkData!' + ORDER_LOOK.readyCell + ')'); }
  catch (e) { out.push('✗ ready count: ' + e); }
  try { out.push(_applyDividerNameplate(sheet)); }
  catch (e) { out.push('✗ band nameplate: ' + e); }
  try { setupDuplicateSalesOrderHighlighting(); out.push('✓ order boxes repainted'); }
  catch (e) { out.push('✗ box repaint: ' + e); }

  SpreadsheetApp.flush();
  try {
    var sd = ss.getSheetByName('__SparkData');
    out.push('  ready now: ' + JSON.stringify(sd.getRange(ORDER_LOOK.readyCell).getDisplayValue()));
    var L = getTableLayout(sheet);
    if (L.direct > 0) out.push('  DIRECT band reads: ' + JSON.stringify(sheet.getRange(L.direct, Schema.bandPlateCol).getDisplayValue()));
  } catch (e) { out.push('  (could not read back: ' + e + ')'); }

  var msg = '✅ Quieter DIRECT table ON\n' + out.join('\n') + '\nRollback: removeOrderLook()';
  console.log(msg);
  return msg;
}

/** The whole rollback: switch off, remove the rules, restore the band and the boxes. */
function removeOrderLook() {
  if (typeof _obRequireOwner === 'function') {
    var denied = _obRequireOwner('Turning off the quieter DIRECT table');
    if (denied) return denied;
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ No ' + MAIN_SHEET_NAME + ' sheet.';
  PropertiesService.getScriptProperties().deleteProperty(ORDER_LOOK.prop);
  sheet.setConditionalFormatRules(_stripOrderLookRules(sheet.getConditionalFormatRules()));
  try { _applyDividerNameplate(sheet); } catch (e) {}
  try { setupDuplicateSalesOrderHighlighting(); } catch (e) {}
  try { refreshKitSkuMarkers(); } catch (e) {}          // K tags back to plain ▣
  return '✅ Quieter DIRECT table OFF — rules removed, band and boxes back to today.';
}
