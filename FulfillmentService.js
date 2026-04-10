// =======================================================================================
// FULFILLMENT_SERVICE.gs - Fulfillment and Printing Functions/
// =======================================================================================

/**
 * Gathers 'PREPARING' rows, separating them into eBay vs DIRECT lists
 * Now includes HAND (Col G) and LEFT (Col H) columns
 */
function preparePrintSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  // Get Employee ID from F2
  var employeeId = sheet.getRange("F2").getValue();

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

  // Build duplicate Sales Order border color map for print highlighting
  var allOrders = {};
  var allItems = ebayItems.concat(directItems);
  for (var j = 0; j < allItems.length; j++) {
    var so = String(allItems[j][3]).trim();
    if (so) allOrders[so] = (allOrders[so] || 0) + 1;
  }
  // Assign border colors only to duplicates (matches ORDER_BORDER_COLORS in RowManagement)
  var ORDER_PRINT_BORDER_COLORS = [
    "#1a73e8", "#e53935", "#43a047", "#fb8c00", "#8e24aa",
    "#00acc1", "#d81b60", "#6d4c41", "#3949ab", "#00897b",
    "#c0ca33", "#f4511e", "#5e35b1", "#039be5", "#7cb342",
    "#ffb300", "#1e88e5", "#e91e63", "#26a69a", "#546e7a"
  ];
  var orderColorMap = {};
  var colorIdx = 0;
  for (var so in allOrders) {
    if (allOrders[so] > 1) {
      orderColorMap[so] = ORDER_PRINT_BORDER_COLORS[colorIdx % ORDER_PRINT_BORDER_COLORS.length];
      colorIdx++;
    }
  }

  // Format print timestamp in Houston timezone
  var now = new Date();
  var houstonTZ = "America/Chicago";
  var printDate = Utilities.formatDate(now, houstonTZ, "M/d/yyyy");
  var printTime = Utilities.formatDate(now, houstonTZ, "h:mm a");

  var htmlTemplate = HtmlService.createTemplateFromFile('PrintFulfillment');
  htmlTemplate.ebayItems = ebayItems;
  htmlTemplate.directItems = directItems;
  htmlTemplate.employeeId = employeeId;
  htmlTemplate.orderColorMap = orderColorMap;
  htmlTemplate.printDate = printDate;
  htmlTemplate.printTime = printTime; 
  
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
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    var range = SpreadsheetApp.getActiveRange();
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