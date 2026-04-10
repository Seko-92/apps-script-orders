// =======================================================================================
// CONFIG.gs - All Configuration Variables
// =======================================================================================

// Spreadsheet ID
var SPREADSHEET_ID = '1yCsQsRL5WPOwWPCFcUekZpVgsN-aSfh6Efx3GzQv8Pg';

// Sheet Names
var MAIN_SHEET_NAME = "All orders";
var DB_SHEET_NAME = "Master Inventory";
var LIVE_UPDATE_SHEET = "Settings";
var LOCATION_UPDATE_SHEET = "Location Update";

// Main Sheet Columns (1-based)
var SKU_COLUMN = 1;                   // Column A
var LOCATION_COLUMN = 3;              // Column C
var HAND_COLUMN = 7;                  // ✅ NEW: Column G (◫ HAND)
var SALES_ORDER_COLUMN = 4;           // Column D
var DATA_START_ROW = 4;
var CLOCK_CELL = "D1";
var STATUS_COLUMN = 6;              
var MAX_EMPTY_ROWS_TO_KEEP = 5;       
var TABLE_TWO_IDENTIFIER = "DIRECT";  
var DATA_WIDTH = 7;  // A, B, C, D, E, F, G (includes HAND column)
var LIVE_UPDATE_TOGGLE_CELL = "B1";

// Master Inventory Column HEADERS (not numbers!)
var DB_SKU_HEADER = "sku";
var DB_LOCATION_HEADER = "C:Model Year";
var DB_QUANTITY_HEADER = "quantity";        // ✅ NEW: Total quantity column
var DB_QUANTITY_SOLD_HEADER = "quantitySold"; // ✅ NEW: Sold quantity column