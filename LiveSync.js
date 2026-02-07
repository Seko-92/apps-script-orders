// =======================================================================================
// LIVE_SYNC.gs - Live Update Trigger Functions (Format-Safe Version)
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
  
  // Only trigger when SKU column (Column A) is edited
  if (range.getColumn() === SKU_COLUMN) {
    var edits = range.getValues();
    var startRow = range.getRow();
    var locationResults = [];
    var quantityResults = [];

    var useMap = edits.length > 5;
    var locationMap = null;
    var inventoryMap = null;

    if (useMap) {
      // Build both maps for bulk operations
      var maps = buildLocationAndInventoryMaps();
      locationMap = maps.locationMap;
      inventoryMap = maps.inventoryMap;
    }

    // Build committed orders map to subtract from available stock
    var committedMap = getCommittedQuantities();

    for (var i = 0; i < edits.length; i++) {
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();

      if (skuLower === "" || skuLower === TABLE_TWO_IDENTIFIER.toLowerCase()) {
        // Empty row or table separator
        locationResults.push([""]);
        quantityResults.push([""]);
      } else {
        var committedQty = committedMap.get(skuLower) || 0;

        if (useMap) {
          // Use pre-built maps (fast for bulk)
          locationResults.push([locationMap.get(skuLower) || "NOT FOUND"]);

          var inventory = inventoryMap.get(skuLower);
          if (inventory) {
            quantityResults.push([inventory.available - committedQty]);
          } else {
            quantityResults.push([0 - committedQty]);
          }
        } else {
          // Direct lookup (fast for single edits)
          locationResults.push([getSingleLocation(skuLower)]);

          var stockInfo = getSingleInventory(skuLower);
          quantityResults.push([stockInfo.available - committedQty]);
        }
      }
    }
    
    // Update LOCATION column (Column C)
    sheet.getRange(startRow, LOCATION_COLUMN, locationResults.length, 1).setValues(locationResults);
    
    // ✅ UPDATED: Update HAND column (Column G) with SMART color coding
    var handRange = sheet.getRange(startRow, HAND_COLUMN, quantityResults.length, 1);
    handRange.setValues(quantityResults);

    // ✅ NEW: Apply colors row-by-row to preserve formatting
    for (var j = 0; j < quantityResults.length; j++) {
      var qty = quantityResults[j][0];
      var cellRow = startRow + j;
      var cell = sheet.getRange(cellRow, HAND_COLUMN);
      
      if (qty === "" || qty === null) {
        // Clear any leftover red highlight when row is emptied
        cell.setBackground(null);
        continue;
      } else if (typeof qty === 'number') {
        if (qty <= 20) {
          // Low stock: Red background, keep everything else
          cell.setBackground("#FF6B6B");
        } else {
          // Good stock: Clear any custom background (reverts to sheet default)
          cell.setBackground(null);
        }
      }
    }
  }
  
  // Only update stats when a data cell is edited (SKU or status column)
  if (range.getColumn() === SKU_COLUMN || range.getColumn() === STATUS_COLUMN) {
    updateOrderStatsInSheet();
  }
}

/**
 * Build both location AND inventory maps at once
 * More efficient than building separately
 */
function buildLocationAndInventoryMaps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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