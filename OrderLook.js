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
 *   2. After an order's first line, the SALES ORDER text goes grey. The keycap badge is
 *      an emoji and keeps its colour, so every line still says which order it is.
 *   3. A note that only repeats the order's first note goes grey. Kit notes (↳), Zoho
 *      flags (⚠) and anything containing the whole word HOLD never do — the same
 *      whole-word rule the HOLD alarms use.
 *   4. The DIRECT band's nameplate ends "· N ready" — orders with every line picked and
 *      none waiting (__SparkData!A31). Hidden at 0.
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
    done: '=AND(' + tag + ',' + below + ',$D' + r + '<>"",' +
          'OR($F' + r + '="SHIPPED",$F' + r + '="CANCELED"),' +
          'COUNTIFS($D$' + Schema.dataStartRow + ':$D,$D' + r + ',$F$' + Schema.dataStartRow + ':$F,"PENDING")+' +
          'COUNTIFS($D$' + Schema.dataStartRow + ':$D,$D' + r + ',$F$' + Schema.dataStartRow + ':$F,"PREPARING")=0)',
    // Every line after an order's first — orders are contiguous under the DIRECT sort.
    so:   '=AND(' + tag + ',' + below + ',$D' + r + '<>"",$D' + r + '=$D' + prev + ')',
    // A note that only repeats the order's FIRST line's note. MATCH finds that first line.
    note: '=AND(' + tag + ',' + below + ',$E' + r + '<>"",$D' + r + '<>"",$D' + r + '=$D' + prev + ',' +
          '$E' + r + '=INDEX($E:$E,MATCH($D' + r + ',$D:$D,0)),' +
          'LEFT($E' + r + ',1)<>"↳",LEFT($E' + r + ',1)<>"⚠",' +
          'NOT(REGEXMATCH($E' + r + '&"","(?i)\\bhold\\b")))'
  };
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
  return [
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
  out.push('✓ colour rules installed (finished orders · grey SO# · repeated notes)');

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
  return '✅ Quieter DIRECT table OFF — rules removed, band and boxes back to today.';
}
