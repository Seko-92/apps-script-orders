// =======================================================================================
// SHEET_INSPECTOR.gs - Read-only design audit for the All Orders sheet
// Run inspectSheetDesign() once. Output lands in:
//   1. Apps Script Logger (View > Logs)
//   2. A hidden sheet "__Inspection" (full report, easy to copy)
// Safe to run repeatedly. Writes nothing to the live data area.
// Kept as a permanent diagnostic — useful before any sheet-visual change to
// audit current bg/fg/fonts/CF/bandings/widths/heights/merges/protections.
// Run via the editor's function dropdown; not wired to any trigger.
// =======================================================================================

function inspectSheetDesign() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    Logger.log("ERROR: Sheet '" + MAIN_SHEET_NAME + "' not found.");
    return;
  }

  var lines = [];
  var pad = function(s, n) { s = String(s); return s.length >= n ? s : s + new Array(n - s.length + 1).join(' '); };

  lines.push("═══════════════════════════════════════════════════════════════");
  lines.push("  ALL ORDERS — SHEET DESIGN INSPECTION");
  lines.push("  " + new Date().toLocaleString());
  lines.push("═══════════════════════════════════════════════════════════════");
  lines.push("");

  // 1. GEOMETRY
  lines.push("── GEOMETRY ──");
  lines.push("Frozen rows:    " + sheet.getFrozenRows());
  lines.push("Frozen columns: " + sheet.getFrozenColumns());
  lines.push("Last row:       " + sheet.getLastRow());
  lines.push("Last column:    " + sheet.getLastColumn());
  lines.push("Max rows:       " + sheet.getMaxRows());
  lines.push("Max columns:    " + sheet.getMaxColumns());
  lines.push("Schema.dataStartRow: " + Schema.dataStartRow);
  var boundary = getBoundaryRow();
  lines.push("DIRECT row:     " + (boundary > 0 ? boundary : "(not found)"));
  lines.push("");

  // 2. COLUMN WIDTHS
  lines.push("── COLUMN WIDTHS (A–J) ──");
  var colNames = ["A: SKU", "B: QTY", "C: LOC", "D: ORDER", "E: NOTE", "F: STATUS", "G: HAND", "H: LEFT", "I: SHIPPING", "J: SHIP COST"];
  for (var c = 1; c <= 10; c++) {
    lines.push("  " + pad(colNames[c-1], 18) + sheet.getColumnWidth(c) + "px");
  }
  lines.push("");

  // 3. ROW HEIGHTS (banner + header + first data row)
  lines.push("── ROW HEIGHTS (1 to " + (Schema.dataStartRow + 1) + ") ──");
  for (var r = 1; r <= Schema.dataStartRow + 1; r++) {
    lines.push("  Row " + pad(r, 4) + sheet.getRowHeight(r) + "px");
  }
  lines.push("");

  // 4. MERGED RANGES (banner + boundary)
  lines.push("── MERGED RANGES — banner area (rows 1-" + (Schema.dataStartRow - 1) + ") ──");
  var bannerScan = sheet.getRange(1, 1, Schema.dataStartRow - 1, Schema.dataWidth).getMergedRanges();
  if (bannerScan.length === 0) lines.push("  (none)");
  else bannerScan.forEach(function(m) { lines.push("  " + m.getA1Notation()); });
  lines.push("");

  if (boundary > 0) {
    lines.push("── MERGED RANGES — DIRECT divider area (rows " + boundary + "-" + (boundary + 1) + ") ──");
    var bScan = sheet.getRange(boundary, 1, 2, Schema.dataWidth).getMergedRanges();
    if (bScan.length === 0) lines.push("  (none)");
    else bScan.forEach(function(m) { lines.push("  " + m.getA1Notation()); });
    lines.push("");
  }

  // 5. BANNER CELL FORMATS — every individual cell in rows 1 to Schema.dataStartRow
  lines.push("── BANNER FORMATS (rows 1-" + Schema.dataStartRow + ", cols A-J) ──");
  lines.push("  cell  | bg       | fontColor | fontFamily          | size | weight | h-align | v-align | wrap   | value");
  lines.push("  " + new Array(140).join('-'));
  for (var rr = 1; rr <= Schema.dataStartRow; rr++) {
    for (var cc = 1; cc <= 10; cc++) {
      var rng = sheet.getRange(rr, cc);
      var a1 = rng.getA1Notation();
      var bg = rng.getBackground();
      var fc = rng.getFontColor();
      var ff = rng.getFontFamily();
      var fs = rng.getFontSize();
      var fw = rng.getFontWeight();
      var ha = rng.getHorizontalAlignment();
      var va = rng.getVerticalAlignment();
      var wr = rng.getWrap();
      var val = String(rng.getValue() || "").replace(/\n/g, " ").substring(0, 30);
      lines.push("  " + pad(a1, 5) + " | " + pad(bg, 8) + " | " + pad(fc, 9) + " | " + pad(ff, 19) + " | " + pad(fs, 4) + " | " + pad(fw, 6) + " | " + pad(ha, 7) + " | " + pad(va, 7) + " | " + pad(wr, 6) + " | " + val);
    }
  }
  lines.push("");

  // 6. FIRST DATA ROW SAMPLE
  lines.push("── DATA ROW SAMPLE (row " + Schema.dataStartRow + ", cols A-J) ──");
  lines.push("  col | bg       | fontColor | fontFamily          | size | weight | h-align | border");
  lines.push("  " + new Array(95).join('-'));
  for (var c2 = 1; c2 <= 10; c2++) {
    var r2 = sheet.getRange(Schema.dataStartRow, c2);
    lines.push("  " + pad(c2, 3) + " | " + pad(r2.getBackground(), 8) + " | " + pad(r2.getFontColor(), 9) + " | " + pad(r2.getFontFamily(), 19) + " | " + pad(r2.getFontSize(), 4) + " | " + pad(r2.getFontWeight(), 6) + " | " + pad(r2.getHorizontalAlignment(), 7));
  }
  lines.push("");

  // 7. DIRECT BOUNDARY + DIRECT HEADER FORMATS
  if (boundary > 0) {
    lines.push("── DIRECT BOUNDARY ROW (row " + boundary + ") ──");
    for (var cb = 1; cb <= 10; cb++) {
      var rb = sheet.getRange(boundary, cb);
      var val = String(rb.getValue() || "").substring(0, 20);
      lines.push("  Col " + pad(cb, 3) + "bg:" + pad(rb.getBackground(), 8) + " fc:" + pad(rb.getFontColor(), 8) + " w:" + pad(rb.getFontWeight(), 6) + " val: " + val);
    }
    lines.push("");

    if (boundary + 1 <= sheet.getLastRow()) {
      lines.push("── DIRECT HEADER ROW (row " + (boundary + 1) + ") ──");
      for (var ch = 1; ch <= 10; ch++) {
        var rh = sheet.getRange(boundary + 1, ch);
        var val = String(rh.getValue() || "").substring(0, 20);
        lines.push("  Col " + pad(ch, 3) + "bg:" + pad(rh.getBackground(), 8) + " fc:" + pad(rh.getFontColor(), 8) + " w:" + pad(rh.getFontWeight(), 6) + " val: " + val);
      }
      lines.push("");
    }
  }

  // 8. BANDINGS
  lines.push("── BANDINGS ──");
  var bandings = sheet.getBandings();
  if (bandings.length === 0) {
    lines.push("  (none)");
  } else {
    bandings.forEach(function(b, i) {
      lines.push("  Banding " + (i + 1) + ": " + b.getRange().getA1Notation());
      try { lines.push("    Header: " + (b.getHeaderColor()  || "(none)")); } catch (e) {}
      try { lines.push("    First:  " + (b.getFirstRowColor()  || "(none)")); } catch (e) {}
      try { lines.push("    Second: " + (b.getSecondRowColor() || "(none)")); } catch (e) {}
      try { lines.push("    Footer: " + (b.getFooterColor()  || "(none)")); } catch (e) {}
    });
  }
  lines.push("");

  // 9. CONDITIONAL FORMATTING
  lines.push("── CONDITIONAL FORMATTING ──");
  var rules = sheet.getConditionalFormatRules();
  lines.push("  Total rules: " + rules.length);
  rules.forEach(function(r, i) {
    var ranges = r.getRanges().map(function(rg) { return rg.getA1Notation(); }).join(", ");
    var bc = r.getBooleanCondition();
    var ct = bc ? bc.getCriteriaType() : "GRADIENT";
    var cv = bc ? (bc.getCriteriaValues() || []).join(" | ") : "(gradient)";
    var bgC = bc ? (bc.getBackgroundObject ? "..." : "") : "";
    lines.push("  " + (i + 1) + ". " + ranges + "  →  " + ct + "  →  " + cv);
  });
  lines.push("");

  // 10. RANGE PROTECTIONS
  lines.push("── RANGE PROTECTIONS ──");
  var prots = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (prots.length === 0) {
    lines.push("  (none)");
  } else {
    prots.forEach(function(p) {
      lines.push("  " + p.getRange().getA1Notation() + "  —  " + (p.getDescription() || "(no description)"));
    });
  }
  lines.push("");

  // 11. EMBEDDED IMAGES
  lines.push("── EMBEDDED IMAGES (over-cell, not =IMAGE() formulas) ──");
  var imgs = sheet.getImages();
  if (imgs.length === 0) {
    lines.push("  (none — any visible logos are likely =IMAGE() formulas inside cells)");
  } else {
    imgs.forEach(function(img, i) {
      try {
        var anchor = img.getAnchorCell().getA1Notation();
        lines.push("  " + (i + 1) + ". anchor: " + anchor + "  size: " + img.getWidth() + "×" + img.getHeight() + "px  url: " + (img.getUrl() ? img.getUrl().substring(0, 60) + "..." : "(inline)"));
      } catch (e) {
        lines.push("  " + (i + 1) + ". (could not read: " + e + ")");
      }
    });
  }
  lines.push("");

  // 12. CELLS WITH FORMULAS in banner (catches =IMAGE(), =NOW(), etc.)
  lines.push("── FORMULAS in banner area (rows 1-" + Schema.dataStartRow + ") ──");
  var bannerFormulas = sheet.getRange(1, 1, Schema.dataStartRow, Schema.dataWidth).getFormulas();
  var foundAny = false;
  for (var fr = 0; fr < bannerFormulas.length; fr++) {
    for (var fc2 = 0; fc2 < bannerFormulas[fr].length; fc2++) {
      var f = bannerFormulas[fr][fc2];
      if (f) {
        foundAny = true;
        var a1 = sheet.getRange(fr + 1, fc2 + 1).getA1Notation();
        lines.push("  " + a1 + ": " + f.substring(0, 100));
      }
    }
  }
  if (!foundAny) lines.push("  (none)");
  lines.push("");

  lines.push("═══════════════════════════════════════════════════════════════");
  lines.push("  END OF INSPECTION");
  lines.push("═══════════════════════════════════════════════════════════════");

  var output = lines.join("\n");
  Logger.log(output);

  // Also drop into a hidden inspection sheet so you can copy the whole thing easily
  try {
    var inspName = "__Inspection";
    var inspSheet = ss.getSheetByName(inspName);
    if (!inspSheet) {
      inspSheet = ss.insertSheet(inspName);
      inspSheet.hideSheet();
    }
    inspSheet.clear();
    inspSheet.getRange(1, 1).setValue(output)
      .setFontFamily("Roboto Mono")
      .setFontSize(9)
      .setVerticalAlignment("top")
      .setWrap(false);
    inspSheet.setColumnWidth(1, 1100);
    Logger.log("\nFull report also saved to hidden sheet '" + inspName + "' — unhide it to copy.");
  } catch (e) {
    Logger.log("(could not write to inspection sheet: " + e + ")");
  }

  return output;
}
