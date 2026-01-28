// =======================================================================================
// LIVE_SYNC.gs - Live Update Trigger Functions
// =======================================================================================

/**
 * Turbo Live Trigger - Optimized for 4000+ items
 * Called automatically when cells are edited (requires onEdit trigger)
 * @param {Event} e - The edit event
 */
function liveUpdateTrigger(e) {
  if (getLiveUpdateState() !== 'ON') return;
  
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== MAIN_SHEET_NAME) return;
  
  if (range.getColumn() === SKU_COLUMN) {
    var edits = range.getValues();
    var startRow = range.getRow();
    var results = [];

    var useMap = edits.length > 5; 
    var skuMap = null;

    if (useMap) {
      skuMap = buildLocationMap();
    }

    for (var i = 0; i < edits.length; i++) {
      var rawSku = String(edits[i][0]).trim();
      var skuLower = rawSku.toLowerCase();
      
      if (skuLower === "" || skuLower === TABLE_TWO_IDENTIFIER.toLowerCase()) {
        results.push([""]);
      } else {
        if (useMap) {
          results.push([skuMap.get(skuLower) || "NOT FOUND"]);
        } else {
          results.push([getSingleLocation(skuLower)]);
        }
      }
    }
    
    sheet.getRange(startRow, LOCATION_COLUMN, results.length, 1).setValues(results);
  }
  
  // THIS MUST BE OUTSIDE THE IF BLOCK - updates on ANY edit
  updateOrderStatsInSheet();
}




