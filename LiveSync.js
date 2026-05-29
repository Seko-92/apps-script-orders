// =======================================================================================
// LIVE_SYNC.gs - Live Update Trigger Functions (Format-Safe Version)//
// =======================================================================================

/**
 * Turbo Live Trigger - Optimized for 4000+ items
 * Now syncs BOTH Location AND Available Quantity with smart color coding
 * Called automatically when cells are edited (requires onEdit trigger)
 * @param {Event} e - The edit event
 */
function liveUpdateTrigger(e) {
  if (getLiveUpdateState() !== 'ON') return;
  
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== MAIN_SHEET_NAME) return;

  // Only trigger when SKU column is edited
  if (range.getColumn() === Schema.cols.SKU) {
    var edits = range.getValues();
    var startRow = range.getRow();
    var locationResults = [];
    var quantityResults = [];

    // Build maps once (MI for location + eBay stock, Zoho for direct/non-eBay
    // stock). This is cheaper than the old per-row single lookups, which each
    // did a full MI read anyway. The Zoho map is empty if the Zoho Stock sheet
    // doesn't exist → HAND falls back to MI exactly as before this feature.
    var maps = buildLocationAndInventoryMaps();
    var locationMap = maps.locationMap;
    var inventoryMap = maps.inventoryMap;
    var zohoMap = buildZohoStockMap();

    // Get boundary row to (a) protect divider/header and (b) route HAND source.
    var boundary = getBoundaryRow();

    // SALES ORDER values for the edited rows — used to tell a manually-typed
    // eBay row (Zoho-first) from an automated eBay-order row (MI-first).
    var soVals = sheet.getRange(startRow, Schema.cols.SALES_ORDER, edits.length, 1).getValues();

    for (var i = 0; i < edits.length; i++) {
      var currentRow = startRow + i;
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      // Protect boundary row and DIRECT header row - preserve their content
      if (boundary > 0 && (currentRow === boundary || currentRow === boundary + 1)) {
        locationResults.push([sheet.getRange(currentRow, Schema.cols.LOCATION).getValue()]);
        quantityResults.push([sheet.getRange(currentRow, Schema.cols.HAND).getValue()]);
        continue;
      }

      if (skuLower === "" || skuLower === Schema.boundaryMarker.toLowerCase()) {
        // Empty row or table separator
        locationResults.push([""]);
        quantityResults.push([""]);
        continue;
      }

      // LOCATION from MI (the eBay shelf map).
      locationResults.push([locationMap.get(skuLower) || "NOT FOUND"]);

      // HAND source routing — DIRECT rows and manually-typed eBay rows prefer
      // Zoho; automated eBay-order rows (clean order id) prefer MI (eBay truth).
      // No committed subtraction: HAND = available, matching recomputeHand /
      // the 2026-05-09 HAND semantics, so entry-time == the recompute.
      var isDirect = (boundary > 0 && currentRow > boundary + 1);
      var preferZoho = isDirect || _isManualSalesOrder(soVals[i][0]);
      var miInv  = inventoryMap.get(skuLower);
      var miAvail = miInv ? miInv.available : null;
      var zo     = zohoMap.get(skuLower);
      var zoAvail = zo ? zo.available : null;
      quantityResults.push([resolveHandValue(miAvail, zoAvail, preferZoho)]);
    }

    // Update LOCATION column
    sheet.getRange(startRow, Schema.cols.LOCATION, locationResults.length, 1).setValues(locationResults);

    // Update HAND column — conditional formatting handles highlighting
    var handRange = sheet.getRange(startRow, Schema.cols.HAND, quantityResults.length, 1);
    handRange.setValues(quantityResults);
  }

  // Only update stats when a data cell is edited (SKU or status column)
  if (range.getColumn() === Schema.cols.SKU || range.getColumn() === Schema.cols.STATUS) {
    updateOrderStatsInSheet();
  }
}

/**
 * Build both location AND inventory maps at once
 * More efficient than building separately
 */
function buildLocationAndInventoryMaps() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet) {
    return {
      locationMap: new Map(),
      inventoryMap: new Map()
    };
  }
  
  var data = dbSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indices
  var skuCol = headers.indexOf(DB_SKU_HEADER);
  var locCol = headers.indexOf(DB_LOCATION_HEADER);
  var qtyCol = headers.indexOf(DB_QUANTITY_HEADER);
  var soldCol = headers.indexOf(DB_QUANTITY_SOLD_HEADER);
  
  if (skuCol === -1) {
    return {
      locationMap: new Map(),
      inventoryMap: new Map()
    };
  }
  
  var locationMap = new Map();
  var inventoryMap = new Map();
  
  // Build both maps in one pass (efficient!)
  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][skuCol] || "").trim().toLowerCase();
    
    if (sku) {
      // Location map
      if (locCol !== -1) {
        var location = String(data[i][locCol] || "").trim();
        locationMap.set(sku, location || "NOT FOUND");
      }
      
      // Inventory map
      if (qtyCol !== -1 && soldCol !== -1) {
        var qty = parseInt(data[i][qtyCol]) || 0;
        var sold = parseInt(data[i][soldCol]) || 0;
        var available = qty - sold;
        
        inventoryMap.set(sku, {
          quantity: qty,
          sold: sold,
          available: available
        });
      }
    }
  }
  
  return {
    locationMap: locationMap,
    inventoryMap: inventoryMap
  };
}

/**
 * Get inventory for a single SKU
 * Used for single-cell edits (faster than building full map)
 */
function getSingleInventory(skuLower) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet) {
    return { quantity: 0, sold: 0, available: 0 };
  }
  
  var data = dbSheet.getDataRange().getValues();
  var headers = data[0];
  
  var skuCol = headers.indexOf(DB_SKU_HEADER);
  var qtyCol = headers.indexOf(DB_QUANTITY_HEADER);
  var soldCol = headers.indexOf(DB_QUANTITY_SOLD_HEADER);
  
  if (skuCol === -1 || qtyCol === -1 || soldCol === -1) {
    return { quantity: 0, sold: 0, available: 0 };
  }
  
  // Search for SKU
  for (var i = 1; i < data.length; i++) {
    var dbSku = String(data[i][skuCol] || "").trim().toLowerCase();
    
    if (dbSku === skuLower) {
      var qty = parseInt(data[i][qtyCol]) || 0;
      var sold = parseInt(data[i][soldCol]) || 0;
      
      return {
        quantity: qty,
        sold: sold,
        available: qty - sold
      };
    }
  }
  
  return { quantity: 0, sold: 0, available: 0 };
}