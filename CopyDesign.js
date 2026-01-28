// =======================================================================================
// SHEET_DESIGN_COPIER.gs - Copy Sheet Design/Template
// =======================================================================================

/**
 * Copies the entire sheet design to a new sheet or another spreadsheet
 * Includes: formatting, column widths, row heights, merged cells, 
 * conditional formatting, data validation, and protected ranges
 */

/**
 * Creates a complete template copy of the current sheet (without data)
 * @returns {string} - URL of the new spreadsheet
 */
function createTemplateFromCurrentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  // Create a new spreadsheet
  var newSS = SpreadsheetApp.create("HQ Motor Service - Template " + new Date().toLocaleDateString());
  var newSheet = newSS.getActiveSheet();
  newSheet.setName(MAIN_SHEET_NAME);
  
  // Copy the design
  copySheetDesign(sourceSheet, newSheet, true); // true = include structure rows
  
  var url = newSS.getUrl();
  Logger.log("✅ Template created: " + url);
  
  return "✅ Template created! URL: " + url;
}

/**
 * Copies sheet design from source to target
 * @param {Sheet} sourceSheet - The sheet to copy from
 * @param {Sheet} targetSheet - The sheet to copy to
 * @param {boolean} includeHeaderData - Whether to include header row data
 */
function copySheetDesign(sourceSheet, targetSheet, includeHeaderData) {
  
  var sourceRows = sourceSheet.getMaxRows();
  var sourceCols = sourceSheet.getMaxColumns();
  
  // ═══════════════════════════════════════════════════════════════
  // 1. Set up target sheet dimensions
  // ═══════════════════════════════════════════════════════════════
  var targetRows = targetSheet.getMaxRows();
  var targetCols = targetSheet.getMaxColumns();
  
  // Add rows/columns if needed
  if (targetRows < sourceRows) {
    targetSheet.insertRowsAfter(targetRows, sourceRows - targetRows);
  }
  if (targetCols < sourceCols) {
    targetSheet.insertColumnsAfter(targetCols, sourceCols - targetCols);
  }
  
  // ═══════════════════════════════════════════════════════════════
  // 2. Copy Column Widths
  // ═══════════════════════════════════════════════════════════════
  for (var col = 1; col <= sourceCols; col++) {
    var width = sourceSheet.getColumnWidth(col);
    targetSheet.setColumnWidth(col, width);
  }
  Logger.log("✅ Column widths copied");
  
  // ═══════════════════════════════════════════════════════════════
  // 3. Copy Row Heights
  // ═══════════════════════════════════════════════════════════════
  for (var row = 1; row <= sourceRows; row++) {
    var height = sourceSheet.getRowHeight(row);
    targetSheet.setRowHeight(row, height);
  }
  Logger.log("✅ Row heights copied");
  
  // ═══════════════════════════════════════════════════════════════
  // 4. Copy All Formatting (fonts, colors, borders, alignment)
  // ═══════════════════════════════════════════════════════════════
  var sourceRange = sourceSheet.getRange(1, 1, sourceRows, sourceCols);
  var targetRange = targetSheet.getRange(1, 1, sourceRows, sourceCols);
  
  // Copy format only
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  Logger.log("✅ Formatting copied");
  
  // ═══════════════════════════════════════════════════════════════
  // 5. Copy Header Data (Row 1, 2, 3 and DIRECT row)
  // ═══════════════════════════════════════════════════════════════
  if (includeHeaderData) {
    // Copy header rows (1-3)
    var headerSource = sourceSheet.getRange(1, 1, 3, sourceCols);
    var headerTarget = targetSheet.getRange(1, 1, 3, sourceCols);
    headerSource.copyTo(headerTarget, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    // Find and copy DIRECT boundary row
    var boundary = findBoundaryRowInSheet(sourceSheet);
    if (boundary > 0) {
      var directSource = sourceSheet.getRange(boundary, 1, 2, sourceCols); // DIRECT + header
      var directTarget = targetSheet.getRange(boundary, 1, 2, sourceCols);
      directSource.copyTo(directTarget, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    }
    Logger.log("✅ Header data copied");
  }
  
  // ═══════════════════════════════════════════════════════════════
  // 6. Copy Merged Cells
  // ═══════════════════════════════════════════════════════════════
  var mergedRanges = sourceSheet.getRange(1, 1, sourceRows, sourceCols).getMergedRanges();
  mergedRanges.forEach(function(mergedRange) {
    var a1 = mergedRange.getA1Notation();
    try {
      targetSheet.getRange(a1).merge();
    } catch(e) {
      Logger.log("Could not merge: " + a1);
    }
  });
  Logger.log("✅ Merged cells copied: " + mergedRanges.length);
  
  // ═══════════════════════════════════════════════════════════════
  // 7. Copy Conditional Formatting Rules
  // ═══════════════════════════════════════════════════════════════
  var rules = sourceSheet.getConditionalFormatRules();
  if (rules.length > 0) {
    targetSheet.setConditionalFormatRules(rules);
    Logger.log("✅ Conditional formatting rules copied: " + rules.length);
  }
  
  // ═══════════════════════════════════════════════════════════════
  // 8. Copy Data Validation (dropdowns like STATUS)
  // ═══════════════════════════════════════════════════════════════
  copyDataValidations(sourceSheet, targetSheet, sourceRows, sourceCols);
  Logger.log("✅ Data validations copied");
  
  // ═══════════════════════════════════════════════════════════════
  // 9. Copy Filter Views (if any)
  // ═══════════════════════════════════════════════════════════════
  var filter = sourceSheet.getFilter();
  if (filter) {
    var filterRange = filter.getRange();
    targetSheet.getRange(filterRange.getA1Notation()).createFilter();
    Logger.log("✅ Filter copied");
  }
  
  // ═══════════════════════════════════════════════════════════════
  // 10. Copy Frozen Rows/Columns
  // ═══════════════════════════════════════════════════════════════
  var frozenRows = sourceSheet.getFrozenRows();
  var frozenCols = sourceSheet.getFrozenColumns();
  if (frozenRows > 0) targetSheet.setFrozenRows(frozenRows);
  if (frozenCols > 0) targetSheet.setFrozenColumns(frozenCols);
  Logger.log("✅ Frozen rows/columns copied");
  
  Logger.log("═══════════════════════════════════════════════════════");
  Logger.log("✅ SHEET DESIGN COPY COMPLETE!");
  Logger.log("═══════════════════════════════════════════════════════");
}

/**
 * Copies data validations from source to target
 */
function copyDataValidations(sourceSheet, targetSheet, rows, cols) {
  // We'll check each cell for data validation
  // This is slower but comprehensive
  
  for (var row = 1; row <= Math.min(rows, 100); row++) { // Limit to first 100 rows for performance
    for (var col = 1; col <= cols; col++) {
      var sourceCell = sourceSheet.getRange(row, col);
      var validation = sourceCell.getDataValidation();
      
      if (validation) {
        var targetCell = targetSheet.getRange(row, col);
        targetCell.setDataValidation(validation);
      }
    }
  }
}

/**
 * Helper: Find boundary row in any sheet
 */
function findBoundaryRowInSheet(sheet) {
  var data = sheet.getRange("A:A").getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase().indexOf("DIRECT") !== -1) {
      return i + 1;
    }
  }
  return -1;
}

// =======================================================================================
// EXPORT DESIGN AS JSON (For sharing or backup)
// =======================================================================================

/**
 * Exports the sheet design as a JSON object
 * Can be saved and imported later
 * @returns {Object} - Design configuration
 */
function exportSheetDesignAsJSON() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  var rows = sheet.getMaxRows();
  var cols = sheet.getMaxColumns();
  
  var design = {
    name: sheet.getName(),
    exportDate: new Date().toISOString(),
    dimensions: {
      rows: rows,
      cols: cols,
      frozenRows: sheet.getFrozenRows(),
      frozenCols: sheet.getFrozenColumns()
    },
    columnWidths: [],
    rowHeights: [],
    headerRows: [],
    boundaryRow: findBoundaryRowInSheet(sheet),
    conditionalFormatRulesCount: sheet.getConditionalFormatRules().length
  };
  
  // Column widths
  for (var col = 1; col <= cols; col++) {
    design.columnWidths.push({
      column: col,
      width: sheet.getColumnWidth(col)
    });
  }
  
  // Row heights (first 50 rows)
  for (var row = 1; row <= Math.min(rows, 50); row++) {
    design.rowHeights.push({
      row: row,
      height: sheet.getRowHeight(row)
    });
  }
  
  // Header row data
  var headerData = sheet.getRange(1, 1, 3, cols).getValues();
  design.headerRows = headerData;
  
  // Log the JSON
  var jsonString = JSON.stringify(design, null, 2);
  Logger.log(jsonString);
  
  return design;
}

/**
 * Creates a Google Doc with the design specifications
 * Useful for documentation
 */
function exportDesignToDoc() {
  var design = exportSheetDesignAsJSON();
  
  var doc = DocumentApp.create("HQ Sheet Design Spec - " + new Date().toLocaleDateString());
  var body = doc.getBody();
  
  body.appendParagraph("HQ MOTOR SERVICE - SHEET DESIGN SPECIFICATION")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Export Date: " + design.exportDate);
  body.appendParagraph("");
  
  body.appendParagraph("DIMENSIONS")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("Rows: " + design.dimensions.rows);
  body.appendParagraph("Columns: " + design.dimensions.cols);
  body.appendParagraph("Frozen Rows: " + design.dimensions.frozenRows);
  body.appendParagraph("Frozen Columns: " + design.dimensions.frozenCols);
  body.appendParagraph("DIRECT Boundary Row: " + design.boundaryRow);
  body.appendParagraph("");
  
  body.appendParagraph("COLUMN WIDTHS")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  design.columnWidths.forEach(function(cw) {
    body.appendParagraph("Column " + cw.column + ": " + cw.width + "px");
  });
  
  body.appendParagraph("");
  body.appendParagraph("HEADER ROW VALUES")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("Row 1: " + design.headerRows[0].join(" | "));
  body.appendParagraph("Row 2: " + design.headerRows[1].join(" | "));
  body.appendParagraph("Row 3: " + design.headerRows[2].join(" | "));
  
  var url = doc.getUrl();
  Logger.log("✅ Design doc created: " + url);
  
  return "✅ Design documentation created: " + url;
}

// =======================================================================================
// QUICK COPY FUNCTIONS (For Sidebar)
// =======================================================================================

/**
 * Duplicates the current sheet within the same spreadsheet
 * Creates an exact copy with all design elements
 */
function duplicateSheetWithDesign() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  var newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(MAIN_SHEET_NAME + " - Copy " + new Date().toLocaleTimeString());
  
  return "✅ Sheet duplicated: " + newSheet.getName();
}

/**
 * Copies sheet design to another spreadsheet by URL
 * @param {string} targetUrl - URL of the target spreadsheet
 */
function copyDesignToSpreadsheet(targetUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  try {
    var targetSS = SpreadsheetApp.openByUrl(targetUrl);
    var targetSheet = targetSS.getSheets()[0]; // Use first sheet
    
    copySheetDesign(sourceSheet, targetSheet, true);
    
    return "✅ Design copied to: " + targetSS.getName();
  } catch (e) {
    return "❌ Error: " + e.message + ". Make sure you have edit access to the target spreadsheet.";
  }
}

// =======================================================================================
// SIDEBAR COMMAND (Optional)
// =======================================================================================

/**
 * Shows a dialog to copy design to another spreadsheet
 */

function devTestCopy() {
  // Use this to test without needing the UI Popup
  var targetUrl = "https://docs.google.com/spreadsheets/d/1gpaXtZpDSFvQlbvsRodUZW6JxsKqkm5waGg7JQXpdmQ/edit?gid=0#gid=0"; 
  copyDesignToSpreadsheet(targetUrl);
  Logger.log("Copy complete!");
}
function showCopyDesignDialog() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 20px; }
      input { width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px; }
      button { background: #D4AF37; color: #000; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-weight: bold; }
      button:hover { background: #F7CB4D; }
      .info { background: #f5f5f5; padding: 10px; border-radius: 4px; margin: 10px 0; font-size: 12px; }
    </style>
    <h3>📋 Copy Sheet Design</h3>
    <div class="info">
      This will copy all formatting, column widths, row heights, and header data to another spreadsheet.
    </div>
    <input type="text" id="targetUrl" placeholder="Paste target spreadsheet URL here...">
    <button onclick="copyDesign()">Copy Design</button>
    <div id="result" style="margin-top: 15px;"></div>
    <script>
      function copyDesign() {
        var url = document.getElementById('targetUrl').value;
        if (!url) {
          document.getElementById('result').innerHTML = '⚠️ Please enter a URL';
          return;
        }
        document.getElementById('result').innerHTML = '⏳ Copying...';
        google.script.run
          .withSuccessHandler(function(msg) {
            document.getElementById('result').innerHTML = msg;
          })
          .withFailureHandler(function(err) {
            document.getElementById('result').innerHTML = '❌ Error: ' + err.message;
          })
          .copyDesignToSpreadsheet(url);
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Copy Sheet Design');
}