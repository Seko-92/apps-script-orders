// =======================================================================================
// FULFILLMENT_SERVICE.gs - Fulfillment and Printing Functions//
// =======================================================================================

/**
 * Gathers 'PREPARING' rows, separating them into eBay vs DIRECT lists
 * Now includes HAND (Col G) and LEFT (Col H) columns
 */
function preparePrintSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  // Get Employee ID from E1
  var employeeId = sheet.getRange("E1").getValue();

  var STATUS_COL_INDEX = 5; // Column F is index 5
  var HAND_COL_INDEX = 6;   // Column G is index 6
  var LEFT_COL_INDEX = 7;   // Column H is index 7
  
  var ebayItems = [];
  var directItems = [];
  var isDirectSection = false; // Flag to track which section we are in

  // Iterate through data rows
  for (var i = DATA_START_ROW - 1; i < data.length; i++) {
    var row = data[i];
    
    // --- 1. DETECT SECTION SPLIT ---
    var rowString = row.join("||").toUpperCase();
    if (rowString.indexOf("DIRECT") > -1 && rowString.length < 200) { 
      isDirectSection = true;
      continue; // Skip the header row itself
    }

    // --- 2. CHECK STATUS ---
    var status = String(row[STATUS_COL_INDEX] || "").trim().toUpperCase();
    
    if (status === "PREPARING") {
      var itemData = [
        row[0],                // 0: SKU (Col A)
        row[1],                // 1: Qty (Col B)
        row[2],                // 2: Location (Col C)
        row[3],                // 3: Sales Order (Col D)
        row[4] || "",          // 4: Note (Col E)
        row[STATUS_COL_INDEX], // 5: Status (Col F)
        row[HAND_COL_INDEX] || "",  // 6: HAND (Col G)
        row[LEFT_COL_INDEX] || ""   // 7: LEFT (Col H)
      ];

      if (isDirectSection) {
        directItems.push(itemData);
      } else {
        ebayItems.push(itemData);
      }
    }
  }

  // Check if both are empty
  if (ebayItems.length === 0 && directItems.length === 0) {
    throw new Error("No items found with status 'PREPARING' in Column F.");
  }

  var htmlTemplate = HtmlService.createTemplateFromFile('PrintFulfillment');
  htmlTemplate.ebayItems = ebayItems;
  htmlTemplate.directItems = directItems;
  htmlTemplate.employeeId = employeeId; 
  
  var ui = htmlTemplate.evaluate()
      .setTitle('Print Picking List')
      .setWidth(1000)
      .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}

/**
 * Changes the status in Column F to "PREPARING" for all selected rows
 */
/**
 * Changes the status in Column F to "PREPARING" and syncs to Telegram.
 */
function markSelectedPreparing() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return "Server busy, please try again.";
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    var range = ss.getActiveRange();
    if (!range) return;

    var startRow = range.getRow();
    var numRows = range.getNumRows();

    if (startRow < DATA_START_ROW) {
      var adj = DATA_START_ROW - startRow;
      startRow = DATA_START_ROW;
      numRows -= adj;
    }
    if (numRows <= 0) return;

    // 1. Batch Update the Sheet
    var statusRange = sheet.getRange(startRow, 6, numRows, 1);
    var values = Array(numRows).fill(["PREPARING"]);
    statusRange.setValues(values);

    // 2. Sync each unique Order ID in the selection to Telegram
    var orderIdData = sheet.getRange(startRow, 4, numRows, 1).getValues();
    var processedIds = new Set();

    for (var i = 0; i < orderIdData.length; i++) {
      var id = orderIdData[i][0];
      if (id && !processedIds.has(id)) {
        syncStatusToTelegram(id, "PREPARING");
        processedIds.add(id);
      }
    }

    updateOrderStatsInSheet();
  } finally {
    lock.releaseLock();
  }
}