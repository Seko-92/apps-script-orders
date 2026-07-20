// =======================================================================================
// CONFIG.gs - Environment & installation configuration
// =======================================================================================
//
// MIGRATION NOTE (2026-04-28):
//   Column geometry, status enum, and the boundary marker have moved to
//   Schema.js (single source of truth for sheet data structure).
//
//   The old constants below (SKU_COLUMN, STATUS_COLUMN, DATA_START_ROW,
//   DATA_WIDTH, TABLE_TWO_IDENTIFIER, etc.) are KEPT as identical literals
//   so any not-yet-migrated reference still resolves correctly.
//   New code should reference Schema.cols.*, Schema.dataStartRow, etc.
//
//   Both files use the SAME numeric values — they don't drift.
//
// WHAT BELONGS HERE
//   - Spreadsheet ID, sheet names (environment / installation)
//   - Master Inventory header names (data-source headers)
//
// WHAT BELONGS IN Schema.js
//   - Column positions, row geometry
//   - Status enum, boundary marker
//   - Banner cell references (G1, D1, etc.)
// =======================================================================================

// ---------- ENVIRONMENT / INSTALLATION ----------
var SPREADSHEET_ID = '1yCsQsRL5WPOwWPCFcUekZpVgsN-aSfh6Efx3GzQv8Pg';

// Sheet Names
var MAIN_SHEET_NAME       = "All orders";
var DB_SHEET_NAME         = "Master Inventory";
var LIVE_UPDATE_SHEET     = "Settings";
var LOCATION_UPDATE_SHEET = "Location Update";

// Master Inventory column HEADERS (not numbers — they're looked up by name in
// LocationService and OrderService because the Master Inventory sheet's column
// order may shift independently from the All Orders sheet)
var DB_SKU_HEADER           = "sku";
var DB_LOCATION_HEADER      = "C:Model Year";
var DB_QUANTITY_HEADER      = "quantity";
var DB_QUANTITY_SOLD_HEADER = "quantitySold";
// listingStatus: "Active" for live listings; the n8n sync flips ended/deleted
// listings to "Completed"/"Ended" (MI rows are never deleted — Gotcha, see
// Price Audit INACTIVE CANDIDATE work 2026-05-25). Out of Stock filters on it
// (2026-07-18) so deleted listings can't linger as phantom restock candidates.
var DB_LISTING_STATUS_HEADER = "listingStatus";
// SKU enrichment (title-on-hover + clickable listing link). Looked up by name,
// live from MI — no cache (title/URL are effectively immutable per listing).
var DB_TITLE_HEADER         = "title";          // human-readable item title
var DB_VIEWURL_HEADER       = "viewItemURL";    // eBay listing URL


// ---------- BACKWARDS-COMPAT MIRRORS OF Schema.js ----------
// These exist so legacy code that hasn't been migrated to Schema.* still works.
// DO NOT add new references to these in new code — use Schema.* instead.
// (Both sources use the same literal values, so they don't drift.)

// Column numbers (1-based) — see Schema.cols
var SKU_COLUMN          = 1;    // Column A — Schema.cols.SKU
var QTY_COLUMN          = 2;    // Column B — Schema.cols.QTY
var LOCATION_COLUMN     = 3;    // Column C — Schema.cols.LOCATION
var SALES_ORDER_COLUMN  = 4;    // Column D — Schema.cols.SALES_ORDER
var NOTE_COLUMN         = 5;    // Column E — Schema.cols.NOTE
var STATUS_COLUMN       = 6;    // Column F — Schema.cols.STATUS
var HAND_COLUMN         = 7;    // Column G — Schema.cols.HAND
var LEFT_COLUMN         = 8;    // Column H — Schema.cols.LEFT
var SHIPPING_COLUMN     = 9;    // Column I — Schema.cols.SHIPPING
var SHIP_COST_COLUMN    = 10;   // Column J — Schema.cols.SHIP_COST

// Row geometry — see Schema.dataStartRow / Schema.bannerRows / Schema.headerRow
var DATA_START_ROW = 4;

// Full row width — see Schema.dataWidth
var DATA_WIDTH = 10;

// Boundary marker — see Schema.boundaryMarker. MUST stay exactly "DIRECT".
var TABLE_TWO_IDENTIFIER = "DIRECT";

// Banner cells — see Schema.cellSyncTime / Schema.cellStats / etc.
// CLOCK_CELL is kept as a back-compat alias only. Live code references
// Schema.cellSyncTime ("E1" — last n8n sync timestamp).
var CLOCK_CELL              = "E1";
var LIVE_UPDATE_TOGGLE_CELL = "B1";

// Misc tunables (no Schema equivalent — internal behavior, not data structure)
var MAX_EMPTY_ROWS_TO_KEEP = 5;
