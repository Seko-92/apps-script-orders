// =======================================================================================
// LOCATION_SERVICE.gs - Location Lookup Functions (FIXED)
// =======================================================================================

/**
 * Builds a SKU → Location map from Master Inventory using HEADER NAMES
 * This is the FIXED version that uses headers instead of hardcoded column numbers
 * @returns {Map} - Map of SKU (lowercase) to Location
 */
function buildLocationMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet) {
    Logger.log("ERROR: Master Inventory sheet not found!");
    return new Map();
  }
  
  // Get all headers from row 1
  var lastCol = dbSheet.getLastColumn();
  var lastRow = dbSheet.getLastRow();
  var headers = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Find column indexes by header name
  var skuColIndex = -1;
  var locationColIndex = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var headerLower = String(headers[i]).trim().toLowerCase();
    if (headerLower === DB_SKU_HEADER.toLowerCase()) {
      skuColIndex = i;
    }
    if (headerLower === DB_LOCATION_HEADER.toLowerCase()) {
      locationColIndex = i;
    }
  }
  
  // Debug logging
  Logger.log("SKU column '" + DB_SKU_HEADER + "' found at index: " + skuColIndex);
  Logger.log("Location column '" + DB_LOCATION_HEADER + "' found at index: " + locationColIndex);
  
  if (skuColIndex === -1) {
    Logger.log("ERROR: SKU header '" + DB_SKU_HEADER + "' not found in Master Inventory!");
    return new Map();
  }
  
  if (locationColIndex === -1) {
    Logger.log("ERROR: Location header '" + DB_LOCATION_HEADER + "' not found in Master Inventory!");
    return new Map();
  }
  
  // Get all data
  var allData = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues(); // Start from row 2 (skip headers)
  
  // Build the map
  var skuMap = new Map();
  
  allData.forEach(function(row) {
    var sku = String(row[skuColIndex] || "").trim().toLowerCase();
    var location = String(row[locationColIndex] || "").trim();
    
    if (sku && location) {
      skuMap.set(sku, location);
    }
  });
  
  Logger.log("Built location map with " + skuMap.size + " entries");
  return skuMap;
}

/**
 * Gets location for a single SKU
 * @param {string} sku - The SKU to look up (will be lowercased)
 * @returns {string} - Location or "NOT FOUND"
 */
function getSingleLocation(sku) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet) return "NOT FOUND";
  
  var lastCol = dbSheet.getLastColumn();
  var lastRow = dbSheet.getLastRow();
  var headers = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Find column indexes
  var skuColIndex = -1;
  var locationColIndex = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var headerLower = String(headers[i]).trim().toLowerCase();
    if (headerLower === DB_SKU_HEADER.toLowerCase()) {
      skuColIndex = i;
    }
    if (headerLower === DB_LOCATION_HEADER.toLowerCase()) {
      locationColIndex = i;
    }
  }
  
  if (skuColIndex === -1 || locationColIndex === -1) {
    return "NOT FOUND";
  }
  
  // Search for the SKU
  var skuLower = String(sku).trim().toLowerCase();
  var data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][skuColIndex]).trim().toLowerCase() === skuLower) {
      var location = String(data[i][locationColIndex]).trim();
      return location || "NOT FOUND";
    }
  }
  
  return "NOT FOUND";
}

/**
 * Updates locations for all rows in a table
 * @param {number} tableNumber - 1 for eBay table, 2 for Direct table
 * @returns {string} - Status message
 */
function updateAllExistingRows(tableNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  
  // startRow: Table 1 starts at row 4. Table 2 starts 2 rows after the "Direct" title.
  var startRow = (tableNumber === 2) ? boundary + 2 : DATA_START_ROW;
  var endRow = (tableNumber === 1) ? boundary - 2 : sheet.getLastRow();
  
  var lastDataRow = findLastDataRowInSegment(startRow, endRow);
  
  // Validation: If lastDataRow points to the header or above, the table is empty
  if (lastDataRow < startRow) return "Table is empty.";

  var skuMap = buildLocationMap();
  if (skuMap.size === 0) {
    return "⚠️ Could not build location map. Check Master Inventory headers.";
  }

  var range = sheet.getRange(startRow, 1, lastDataRow - startRow + 1, DATA_WIDTH);
  var data = range.getValues();
  var updates = 0;

  for (var i = 0; i < data.length; i++) {
    var rawSku = String(data[i][SKU_COLUMN-1]).trim();
    var sku = rawSku.toLowerCase();
    var oldLoc = String(data[i][LOCATION_COLUMN-1]).trim();
    
    // --- THE HEADER SHIELD ---
    // Skip row if: Empty, is the Table Title, OR contains the word "SKU"
    if (sku === "" || 
        sku === TABLE_TWO_IDENTIFIER.toLowerCase() || 
        rawSku.includes("SKU")) { 
      continue; 
    }
    
    var found = skuMap.get(sku) || "NOT FOUND";
    
    if (oldLoc !== String(found)) { 
      data[i][LOCATION_COLUMN-1] = found; 
      updates++; 
    }
  }

  if (updates > 0) range.setValues(data);
  return "✅ Updated " + updates + " rows.";
}

// Wrapper functions for sidebar
function runUpdateLocationsTableOne() { return updateAllExistingRows(1); }
function runUpdateLocationsTableTwo() { return updateAllExistingRows(2); }
