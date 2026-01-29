// =======================================================================================
// STATS_HELPER.gs - v3.1 - Added CANCELED Support
// =======================================================================================

/**
 * Gets current order statistics for sidebar
 * Used by notification system to detect new orders
 * 
 * @returns {Object} - Order counts by status
 */
function getOrderStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    return { pending: 0, preparing: 0, shipped: 0, canceled: 0 };
  }
  
  var boundary = getBoundaryRow();
  var startRow = DATA_START_ROW;
  var endRow = (boundary > 0) ? boundary - 1 : sheet.getLastRow();
  
  var stats = {
    pending: 0,
    preparing: 0,
    shipped: 0,
    canceled: 0  // ✅ Added
  };
  
  if (endRow < startRow) {
    return stats;
  }
  
  // Get all status values (column F = column 6)
  var statusData = sheet.getRange(startRow, 6, endRow - startRow + 1, 1).getValues();
  
  statusData.forEach(function(row) {
    var status = String(row[0]).trim().toUpperCase();
    
    if (status === 'PENDING') {
      stats.pending++;
    } else if (status === 'PREPARING') {
      stats.preparing++;
    } else if (status === 'SHIPPED') {
      stats.shipped++;
    } else if (status === 'CANCELED') {  // ✅ Added
      stats.canceled++;
    }
  });
  
  return stats;
}

/**
 * Alternative function that returns formatted stats string
 * 
 * @returns {string} - Formatted stats message
 */
function getStatsMessage() {
  var stats = getOrderStats();
  return "📊 PENDING: " + stats.pending + 
         " | PREPARING: " + stats.preparing + 
         " | SHIPPED: " + stats.shipped +
         " | CANCELED: " + stats.canceled;  // ✅ Added
}