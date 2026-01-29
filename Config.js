// =======================================================================================
// CONFIG.gs - All Configuration Variables{}
// =======================================================================================

// Sheet Names
var MAIN_SHEET_NAME = "All orders"; 
var DB_SHEET_NAME = "Master Inventory"; 
var LIVE_UPDATE_SHEET = "Settings";

// Main Sheet Columns (1-based)
var SKU_COLUMN = 1;                   // Column A
var LOCATION_COLUMN = 3;              // Column C
var DATA_START_ROW = 4;               
var CLOCK_CELL = "C1";  
var STATUS_COLUMN = 6;              
var MAX_EMPTY_ROWS_TO_KEEP = 5;       
var TABLE_TWO_IDENTIFIER = "DIRECT";  
var DATA_WIDTH = 7;                   
var LIVE_UPDATE_TOGGLE_CELL = "B1";   

// Master Inventory Column HEADERS (not numbers!)
// This is the key fix - we use header names instead of column numbers
var DB_SKU_HEADER = "sku";
var DB_LOCATION_HEADER = "C:Model Year";
