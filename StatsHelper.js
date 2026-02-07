// =======================================================================================
// STATS_HELPER.gs - v3.2 - getOrderStats() is now in OrderService.js (single source)
// =======================================================================================

/**
 * Alternative function that returns formatted stats string
 * Note: getOrderStats() is defined in OrderService.js to avoid duplicate definitions
 *
 * @returns {string} - Formatted stats message
 */
function getStatsMessage() {
  var stats = getOrderStats();
  return "PENDING: " + stats.pending +
         " | PREPARING: " + stats.preparing +
         " | SHIPPED: " + stats.shipped +
         " | CANCELED: " + stats.canceled;
}
