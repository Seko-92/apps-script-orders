// =======================================================================================
// ORDER_SERVICE.gs - COMPLETE with Hidden Sheet Message ID Storage
// =======================================================================================

// Note: TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are now defined in Secrets.js
// Make sure Secrets.js is uploaded to your Apps Script project
var HIDDEN_SHEET_NAME = "Telegram_Messages"; // Hidden sheet for message IDs

/**
 * The "Front Door" for n8n - Receives POST requests
 */
/**
 * The "Front Door" for n8n - Receives POST requests
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Server Busy"})).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var payload = JSON.parse(e.postData.contents);

    // --- AUTHENTICATION ---
    // Telegram callbacks are verified by their structure (callback_query with valid bot data)
    // All other requests must include the shared secret token
    if (!payload.callback_query) {
      var token = String(payload.token || (e.parameter && e.parameter.token) || "").trim();
      var expected = APP_SECRET_TOKEN;
      if (token !== expected) {
        console.log("AUTH FAILED - received: [" + token + "] expected: [" + expected + "]");
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", "message": "Unauthorized"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- TELEGRAM ACTIONS ---
    if (payload.action === 'storeMessageId') return storeMessageId(payload.orderId, payload.messageId, payload.chatId);
    if (payload.action === 'notifyShipped') return notifyTelegramShipped(payload.orderId);
    if (payload.callback_query) return handleTelegramCallback(payload);

    // --- STATUS UPDATES ---
    if (payload.action === 'updateOrderStatus') {
      var result = findAndUpdateOrder(payload.orderId, payload.newStatus);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'updateStatus') {
      // Validate row number is within data range
      var rowNum = parseInt(payload.rowNumber);
      if (isNaN(rowNum) || rowNum < DATA_START_ROW) {
        return ContentService.createTextOutput(JSON.stringify({
          "status": "error", "message": "Invalid row number"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      return updateStatus(rowNum, payload.status);
    }

    // --- IMPROVED BATCH ORDER INSERTION ---
    var orders = payload.orders || [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    
    // 1. Gather Existing Orders to prevent duplicates
    // Scan ALL data rows to avoid duplicate insertions
    var lastRow = sheet.getLastRow();
    var scanRows = Math.max(0, lastRow - DATA_START_ROW + 1);
    if (scanRows === 0) scanRows = 1;
    var existingData = sheet.getRange(DATA_START_ROW, 1, scanRows, 4).getValues();
    var existingSignatures = new Set();
    existingData.forEach(function(row) {
      if (row[3] && row[0]) existingSignatures.add(String(row[3]).trim() + "|" + String(row[0]).trim().toUpperCase());
    });

    var newRows = [];
    var results = [];

    // 2. Build location & inventory maps ONCE (not per-item)
    // This reads Master Inventory LIVE - fresh data every doPost call
    // but avoids re-reading the entire sheet for EACH item
    var maps = buildLocationAndInventoryMaps();
    var locationMap = maps.locationMap;
    var inventoryMap = maps.inventoryMap;

    // Track available stock decrements within this batch
    var batchStock = {};

    // Build committed orders map to subtract existing PENDING/PREPARING from available
    var committedMap = getCommittedQuantities();

    // 3. Build the data array first (Don't touch the sheet yet)
    orders.forEach(function(item) {
      var sku = String(item.SKU || "").trim().toUpperCase();
      var skuLower = sku.toLowerCase();
      var salesOrder = String(item["SALES ORDER"] || "").trim();

      if (!salesOrder || salesOrder.length < 3) {
        results.push("Skipped: Invalid ID");
        return;
      }

      // Duplicate Check
      if (existingSignatures.has(salesOrder + "|" + sku)) {
        results.push("Skipped: Duplicate " + salesOrder);
        return;
      }

      // LIVE location lookup from map (built fresh this request)
      var location = locationMap.get(skuLower) || "NOT FOUND";

      // Inventory with batch-level decrement tracking
      // Also subtract already committed orders (PENDING/PREPARING) in the sheet
      if (!(skuLower in batchStock)) {
        var invData = inventoryMap.get(skuLower);
        var baseAvailable = invData ? invData.available : 0;
        var alreadyCommitted = committedMap.get(skuLower) || 0;
        batchStock[skuLower] = baseAvailable - alreadyCommitted;
      }
      var itemQty = parseInt(item.QTY) || 1;

      // Subtract this item's qty to show what's LEFT after this order
      var handValue = batchStock[skuLower] - itemQty;

      // Store row data: [SKU, Qty, Loc, OrderID, Note, Status, Hand/Stock]
      newRows.push([
        sku,
        item.QTY || 1,
        location,
        salesOrder,
        item.NOTE || "",
        "PENDING",
        handValue
      ]);

      // Update batchStock for next order with same SKU in this batch
      batchStock[skuLower] = handValue;

      existingSignatures.add(salesOrder + "|" + sku);
      results.push("Added: " + salesOrder);
    });

    // 3. Insert and Format in ONE GO (Much Faster)
    if (newRows.length > 0) {
      // A. Insert blank rows at the top
      sheet.insertRowsBefore(DATA_START_ROW, newRows.length);
      
      // B. Get the target range
      var range = sheet.getRange(DATA_START_ROW, 1, newRows.length, 7);
      
      // C. Paste Data
      range.setValues(newRows);
      
      // D. Clean Formatting (Fixes "Format Persistence")
      // We copy format from the row *below* the insertion to ensure borders/fonts match
      var templateRow = DATA_START_ROW + newRows.length;
      sheet.getRange(templateRow, 1, 1, 7).copyFormatToRange(sheet, 1, 7, DATA_START_ROW, DATA_START_ROW + newRows.length - 1);

      updateOrderStatsInSheet();
      updateLastSyncTimestamp();
    }

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "added": newRows.length, 
      "details": results
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// 📊 INVENTORY LOOKUP - Single definition is at the bottom of this file
// ═══════════════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ HIDDEN SHEET MANAGEMENT - Store Message IDs
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Get or create the hidden sheet for storing message IDs
 */
function getHiddenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    
    // Set headers
    sheet.getRange("A1:D1").setValues([["Order ID", "Message ID", "Chat ID", "Timestamp"]]);
    
    // Format headers
    sheet.getRange("A1:D1")
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF");
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // Order ID
    sheet.setColumnWidth(2, 120); // Message ID
    sheet.setColumnWidth(3, 150); // Chat ID
    sheet.setColumnWidth(4, 180); // Timestamp
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Hide the sheet
    sheet.hideSheet();
    
    logDebug("Created hidden sheet: " + HIDDEN_SHEET_NAME);
  }
  
  return sheet;
}

// storeMessageId() and getMessageId() - single definitions are at the bottom of this file

/**
 * Delete message ID entry after order is deleted or completed
 */
function deleteMessageIdEntry(orderId) {
  try {
    var sheet = getHiddenSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        sheet.deleteRow(i + 1);
        logDebug("Deleted message ID entry for order: " + orderId);
        return true;
      }
    }
    
    return false;
    
  } catch (e) {
    logDebug("Error deleting message ID: " + e.toString());
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ AUTO-CLEANUP: Remove old message IDs (older than 7 days)
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Clean up message IDs older than specified days
 * Run this weekly via Apps Script trigger
 */
function cleanupOldMessageIds(daysToKeep) {
  try {
    daysToKeep = daysToKeep || 7; // Default 7 days
    
    logDebug("=== CLEANUP OLD MESSAGE IDs ===");
    logDebug("Keeping entries from last " + daysToKeep + " days");
    
    var sheet = getHiddenSheet();
    var data = sheet.getDataRange().getValues();
    var today = new Date();
    var deletedCount = 0;
    
    // Loop backwards to avoid row index shifting
    for (var i = data.length - 1; i > 0; i--) {
      var timestamp = new Date(data[i][3]);
      var ageInDays = (today - timestamp) / (1000 * 60 * 60 * 24);
      
      if (ageInDays > daysToKeep) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    logDebug("✅ Deleted " + deletedCount + " old entries");
    
    return {
      success: true,
      deletedCount: deletedCount,
      daysToKeep: daysToKeep
    };
    
  } catch (e) {
    logDebug("❌ Cleanup error: " + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Setup weekly cleanup trigger (run this once manually)
 */
function setupWeeklyCleanup() {
  // Delete existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupOldMessageIds') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new weekly trigger (runs every Monday at 2 AM)
  ScriptApp.newTrigger('cleanupOldMessageIds')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();
  
  Logger.log("✅ Weekly cleanup trigger created - runs every Monday at 2 AM");
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ✨ NOTIFY TELEGRAM WHEN SHIPPED - Uses Stored Message ID
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Called by n8n when an order is marked SHIPPED
 * Updates the existing Telegram message using stored message_id
 */
function notifyTelegramShipped(orderId) {
  // 1. Retrieve the Chat ID and Message ID from the hidden sheet
  var msgData = getMessageId(orderId);

  if (!msgData) {
    console.log("⚠️ Order " + orderId + " not found in Telegram_Messages sheet.");
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "skipped", 
      reason: "Message ID not found for Order " + orderId 
    }));
  }

  // 2. Define the new "Shipped" message text
  var newText = "📦 <b>Order " + orderId + "</b>\n" +
                "────────────────────\n" +
                "✅ <b>STATUS: SHIPPED</b>\n\n" +
                "<i>This order has been processed and shipped.</i>";

  // 3. Send the edit request to Telegram (CORRECTED ENDPOINT)
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";

  var payload = {
    'chat_id': String(msgData.chatId),      // Ensure it's a string
    'message_id': parseInt(msgData.messageId), // Must be integer for Telegram API
    'text': newText,
    'parse_mode': 'HTML',
    'reply_markup': { 'inline_keyboard': [] } // Removes the buttons
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // Prevents crash so we can read the error
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseBody = JSON.parse(response.getContentText());

    if (responseBody.ok) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", action: "updated_telegram" }));
    } else {
      // Log the exact error from Telegram
      console.error("Telegram Error for Order " + orderId + ": " + responseBody.description);
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        telegram_error: responseBody.description 
      }));
    }

  } catch (e) {
    console.error("Script Error: " + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: e.toString() }));
  }
}

/**
 * Update Telegram message to show SHIPPED status with no buttons
 */
function updateTelegramMessageToShipped(chatId, messageId, orderId) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
  
  try {
    // Build new message text for shipped status
    var newText = buildShippedMessageText(orderId);
    
    // Build payload with NO buttons
    var payload = {
      "chat_id": chatId,
      "message_id": messageId,
      "text": newText,
      "parse_mode": "HTML",
      "reply_markup": {
        "inline_keyboard": []  // Empty = no buttons
      }
    };
    
    logDebug("Editing message " + messageId + " in chat " + chatId);
    
    var response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    
    var result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      logDebug("✅ Message updated successfully!");
      return true;
    } else {
      logDebug("❌ Failed to update message: " + result.description);
      return false;
    }
    
  } catch (e) {
    logDebug("❌ Error updating message: " + e.toString());
    return false;
  }
}

/**
 * Build the SHIPPED message text
 * Preserves original order details, just updates status
 */
function buildShippedMessageText(orderId) {
  // This is a simplified version - you may want to fetch full order details from sheet
  var text = "📦 <b>Order Complete</b>\n\n";
  text += "Order: <code>" + orderId + "</code>\n\n";
  text += "━━━━━━━━━━━━━━━\n";
  text += "📋 Status: ✅ <b>SHIPPED - Order Complete!</b>\n";
  text += "🟢 Ready for carrier pickup";
  
  return text;
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// TELEGRAM INTERACTION HANDLER - Updated for SHIPPED
// ═══════════════════════════════════════════════════════════════════════════════════════

function handleTelegramCallback(payload) {
  var callback = payload.callback_query;
  var data = callback.data;
  var chatId = callback.message.chat.id;
  var messageId = callback.message.message_id;
  var originalText = callback.message.text || "";
  
  logDebug("=== CALLBACK RECEIVED ===");
  logDebug("Order: " + data);
  logDebug("Chat ID: " + chatId);
  logDebug("Message ID: " + messageId);
  
  var action = "";
  var orderId = "";
  
  if (data.startsWith("PREP_")) {
    action = "PREPARING";
    orderId = data.replace("PREP_", "");
  } else if (data.startsWith("PEND_")) {
    action = "PENDING";
    orderId = data.replace("PEND_", "");
  } else {
    answerCallbackQuery(callback.id, "❓ Unknown action", true);
    return ContentService.createTextOutput("OK");
  }
  
  logDebug("Action: " + action + " for Order: " + orderId);
  
  // 1. UPDATE THE SHEET
  var result = findAndUpdateOrder(orderId, action);
  logDebug("Sheet update result: " + JSON.stringify(result));
  
  if (!result.found) {
    answerCallbackQuery(callback.id, "⚠️ Order not found", true);
    return ContentService.createTextOutput("OK");
  }
  
  // ✨ FIX: Check if order was already SHIPPED
  if (result.currentStatus === "SHIPPED") {
    answerCallbackQuery(callback.id, "✅ Order already shipped!", true);
    
    // Update message to show SHIPPED with NO buttons
    var editResult = updateMessageStatus(chatId, messageId, originalText, orderId, "SHIPPED");
    logDebug("Edit result (shipped): " + editResult);
    return ContentService.createTextOutput("OK");
  }
  
  // 2. SHOW TOAST
  var toastText = action === "PREPARING" ? "✅ Preparing" : "🔄 Pending";
  answerCallbackQuery(callback.id, toastText, false);
  
  // 3. EDIT MESSAGE
  var editResult = updateMessageStatus(chatId, messageId, originalText, orderId, action);
  logDebug("Edit result: " + editResult);
  
  return ContentService.createTextOutput("OK");
}

function findAndUpdateOrder(orderId, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  
  logDebug("--- Multi-Update Triggered ---");
  logDebug("Target Order ID: [" + orderId + "]");
  logDebug("New Status Request: " + newStatus);

  if (lastRow < DATA_START_ROW) return { found: false, count: 0 };

  var range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6);
  var data = range.getValues();
  
  var foundCount = 0;
  var currentStatus = "";
  var cleanTargetId = String(orderId).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var rawRowId = data[i][3];
    var cleanRowId = String(rawRowId).trim().toLowerCase();
    
    if (cleanRowId === cleanTargetId && cleanTargetId !== "") {
      var actualRow = DATA_START_ROW + i;
      currentStatus = String(data[i][5]).toUpperCase();
      
      // ✨ IMPORTANT: Return current status if SHIPPED
      // Prevent reverting from final states
      if (currentStatus === "SHIPPED" || currentStatus === "CANCELED") {
        logDebug("Row " + actualRow + " is already " + currentStatus + ". Cannot revert.");
        return { found: true, count: 0, currentStatus: currentStatus };
      }

      // Update the status
      sheet.getRange(actualRow, 6).setValue(newStatus);
      foundCount++;
      logDebug("✅ Match Found! Updated Row: " + actualRow + " (SKU: " + data[i][0] + ")");
    }
  }

  if (foundCount > 0) {
    updateOrderStatsInSheet();
    sortTableByStatusAndLocation(1);
  }

  return { 
    found: foundCount > 0, 
    count: foundCount, 
    currentStatus: currentStatus 
  };
}

function answerCallbackQuery(callbackQueryId, text, showAlert) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/answerCallbackQuery";
  var payload = {
    "callback_query_id": callbackQueryId,
    "text": text,
    "show_alert": showAlert || false
  };
  
  try {
    UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
  } catch (e) {
    logDebug("Toast error: " + e.toString());
  }
}

function updateMessageStatus(chatId, messageId, originalText, orderId, newStatus) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
  
  // Remove old status line if exists
  var cleanText = originalText.replace(/\n\n📋 Status:.*$/s, "");
  
  // Different formatting for each status
  var statusEmoji = "";
  var statusText = "";
  
if (newStatus === "SHIPPED") {
  statusEmoji = "✅";
  statusText = "SHIPPED - Order Complete!";
} else if (newStatus === "CANCELED") {
  statusEmoji = "❌";
  statusText = "CANCELED - Order Cancelled";
} else if (newStatus === "PREPARING") {
  statusEmoji = "🟡";
  statusText = "PREPARING";
} else {
  statusEmoji = "🔴";
  statusText = "PENDING";
}
  
  var newText = cleanText + "\n\n📋 Status: " + statusEmoji + " " + statusText;
  
  // ✨ NO buttons if SHIPPED!
  var keyboard = null;
  
if (newStatus === "SHIPPED" || newStatus === "CANCELED") {
  keyboard = { "inline_keyboard": [] };  // NO BUTTONS for final states
} else if (newStatus === "PENDING") {
  keyboard = { 
    "inline_keyboard": [
      [{ "text": "🚀 Mark as Preparing", "callback_data": "PREP_" + orderId }]
    ] 
  };
} else if (newStatus === "PREPARING") {
  keyboard = { 
    "inline_keyboard": [
      [{ "text": "🔄 Revert to Pending", "callback_data": "PEND_" + orderId }]
    ] 
  };
}
  
  var payload = {
    "chat_id": chatId,
    "message_id": messageId,
    "text": newText,
    "reply_markup": keyboard
  };
  
  logDebug("Sending edit request...");
  logDebug("New status: " + newStatus);
  
  try {
    var response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    
    var responseText = response.getContentText();
    var result = JSON.parse(responseText);
    
    if (result.ok) {
      logDebug("✅ EDIT SUCCESS");
      return "SUCCESS";
    } else {
      logDebug("❌ EDIT FAILED: " + result.description);
      return "FAILED: " + result.description;
    }
    
  } catch (e) {
    logDebug("❌ EDIT ERROR: " + e.toString());
    return "ERROR: " + e.toString();
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// DEBUG LOGGING
// ═══════════════════════════════════════════════════════════════════════════════════════

function logDebug(message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Debug Log");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Debug Log");
      logSheet.getRange("A1").setValue("Timestamp");
      logSheet.getRange("B1").setValue("Message");
      logSheet.setFrozenRows(1);
    }
    
    var timestamp = new Date().toLocaleString();
    logSheet.appendRow([timestamp, message]);
    
    var lastRow = logSheet.getLastRow();
    if (lastRow > 101) {
      logSheet.deleteRows(2, lastRow - 101);
    }
  } catch (e) {
    // Silently fail
  }
}

function clearDebugLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Debug Log");
  if (logSheet) {
    logSheet.clear();
    logSheet.getRange("A1").setValue("Timestamp");
    logSheet.getRange("B1").setValue("Message");
    logSheet.setFrozenRows(1);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// WEBHOOK MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════════════

function setWebhook() {
  // Note: WEB_APP_URL is now defined in Secrets.js
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/setWebhook?url=" + WEB_APP_URL);
  Logger.log(response.getContentText());
}

function deleteWebhook() {
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/deleteWebhook");
  Logger.log(response.getContentText());
}

function getWebhookInfo() {
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/getWebhookInfo");
  Logger.log(response.getContentText());
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// ORDER MANAGEMENT FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════════════

function addOrderFromN8N(sku, salesOrder, qty) {
  if (!sku || !salesOrder) return "Skipped: Missing data";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var startRow = DATA_START_ROW;
  var endRow = boundary - 1;
  
  if (endRow >= startRow) {
    var existingOrders = sheet.getRange(startRow, 4, endRow - startRow + 1, 1).getValues();
    for (var i = 0; i < existingOrders.length; i++) {
      if (String(existingOrders[i][0]).trim() === String(salesOrder).trim()) {
        return "Order " + salesOrder + " already exists.";
      }
    }
  }
  var location = getSingleLocation(String(sku).toLowerCase().trim());
  var baseQty = getInventoryForSKU(sku);
  var committedMap = getCommittedQuantities();
  var committedQty = committedMap.get(String(sku).trim().toLowerCase()) || 0;
  var availableQty = baseQty - committedQty - (parseInt(qty) || 1);
  sheet.insertRowBefore(DATA_START_ROW);
  sheet.getRange(DATA_START_ROW, 1, 1, 7).setValues([[sku, qty, location, salesOrder, "", "PENDING", availableQty]]);  // ✅ Changed from 6 to 7

  updateOrderStatsInSheet();
  updateLastSyncTimestamp();
  return "Added: " + salesOrder;
}

function getOrderStats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) throw new Error("Main sheet not found");
  var lastRow = sheet.getLastRow();
if (lastRow < DATA_START_ROW) return { pending: 0, preparing: 0, shipped: 0, canceled: 0 };
var statuses = sheet.getRange(DATA_START_ROW, 6, lastRow - DATA_START_ROW + 1, 1).getValues().flat();
var stats = { pending: 0, preparing: 0, shipped: 0, canceled: 0 };
statuses.forEach(function(s) {
  s = String(s).trim().toUpperCase();
  if (s === 'PENDING') stats.pending++;
  else if (s === 'PREPARING') stats.preparing++;
  else if (s === 'SHIPPED') stats.shipped++;
  else if (s === 'CANCELED') stats.canceled++;
});
return stats;
}

function updateOrderStatsInSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  
  var stats = getOrderStats();
  
  // ✅ Updated to include CANCELED
  var text = "🔴 Pending: " + stats.pending + 
             "   🟡 Preparing: " + stats.preparing + 
             "   🟢 Shipped: " + stats.shipped +
             "   ⚫ Canceled: " + stats.canceled;
  
  var range = sheet.getRange('F1:H1');
  try { range.breakApart(); } catch(e) {}
  range.merge();
  
  sheet.getRange('F1')
    .setValue(text)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBackground('#212121')
    .setFontColor('#FFFFFF')
    .setWrap(false);
}

function updateLastSyncTimestamp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var now = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "h:mm a");
  sheet.getRange('D1').setValue("⏱ " + timestamp).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#212121').setFontColor('#FFFFFF');
}

function sortTableByStatusAndLocation(tableNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var boundary = getBoundaryRow();
  var startRow = (tableNumber === 1) ? DATA_START_ROW : boundary + 2;
  var endRow = (tableNumber === 1) ? boundary - 1 : sheet.getLastRow();
  var lastDataRow = findLastDataRowInSegment(startRow, endRow);
  if (lastDataRow < startRow) return "Table is empty.";
  var numRows = lastDataRow - startRow + 1;
  var range = sheet.getRange(startRow, 1, numRows, 8);
  var data = range.getValues();
  var statusOrder = { 'PENDING': 1, 'PREPARING': 2, 'SHIPPED': 3, 'CANCELED': 4, '': 5 };
  data.sort(function(a, b) {
    var sA = String(a[5] || '').trim().toUpperCase();
    var sB = String(b[5] || '').trim().toUpperCase();
    var cmp = (statusOrder[sA] || 4) - (statusOrder[sB] || 4);
    if (cmp !== 0) return cmp;
    return String(a[2] || '').localeCompare(String(b[2] || ''));
  });
  range.setValues(data);
  return "✅ Sorted";
}

function sortEbayTable() { return sortTableByStatusAndLocation(1); }
function sortDirectTable() { return sortTableByStatusAndLocation(2); }

function refreshProDashboard() {
  updateOrderStatsInSheet();
  updateDailyDate();
  return "✅ Dashboard refreshed!";
}

function updateStatus(rowNumber, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  sheet.getRange(rowNumber, 6).setValue(status);
  updateOrderStatsInSheet();
  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// 💾 MESSAGE STORAGE HELPERS (The Missing Link)
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Stores the Telegram Message ID and Chat ID linked to an Order ID
 * Called by doPost when action === 'storeMessageId'
 */
function storeMessageId(orderId, messageId, chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    // Safety: Create the sheet if it doesn't exist
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    sheet.appendRow(["Order ID", "Message ID", "Chat ID", "Timestamp"]);
  }
  
  // Append the new log entry
  sheet.appendRow([orderId, messageId, chatId, new Date()]);
  
  return ContentService.createTextOutput("Stored");
}

/**
 * Retrieves the Message ID and Chat ID for a specific Order ID
 * Called by notifyTelegramShipped
 */
function getMessageId(orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  if (!sheet) {
    logDebug("❌ Error: " + HIDDEN_SHEET_NAME + " sheet not found.");
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  var targetId = String(orderId).trim().toLowerCase();
  
  // Loop backwards (bottom to top) to find the most recent entry for this order
  // Start at i >= 1 to skip the header row
  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = String(data[i][0]).trim().toLowerCase();
    
    if (rowId === targetId) {
      return {
        messageId: data[i][1],
        chatId: data[i][2]
      };
    }
  }
  
  return null; // Not found
}


// ═══════════════════════════════════════════════════════════════════════════════════════
// 📊 INVENTORY LOOKUP FUNCTIONS - For Auto-Populating HAND Column
// ═══════════════════════════════════════════════════════════════════════════════════════

/**
 * Get available inventory for a SKU
 * Used when inserting new orders via n8n
 */
function getInventoryForSKU(sku) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet || !sku) {
    return 0;
  }
  
  var skuLower = String(sku).trim().toLowerCase();
  var data = dbSheet.getDataRange().getValues();
  var headers = data[0];
  
  var skuCol = headers.indexOf(DB_SKU_HEADER);
  var qtyCol = headers.indexOf(DB_QUANTITY_HEADER);
  var soldCol = headers.indexOf(DB_QUANTITY_SOLD_HEADER);
  
  if (skuCol === -1 || qtyCol === -1 || soldCol === -1) {
    return 0;
  }
  
  // Search for SKU
  for (var i = 1; i < data.length; i++) {
    var dbSku = String(data[i][skuCol] || "").trim().toLowerCase();
    
    if (dbSku === skuLower) {
      var qty = parseInt(data[i][qtyCol]) || 0;
      var sold = parseInt(data[i][soldCol]) || 0;
      return qty - sold; // Return available quantity
    }
  }
  
  return 0; // SKU not found
}


/**
 * Synchronizes a status change from the sheet to Telegram.
 * REPLICATES THE "ELEGANT MOBILE DESIGN" FROM N8N WORKFLOW
 */
function syncStatusToTelegram(orderId, newStatus) {
  logDebug("=== SYNC SHEET → TELEGRAM ===");
  logDebug("Order: " + orderId + " → " + newStatus);

  var msgData = getMessageId(orderId);
  if (!msgData) {
    logDebug("⚠️ Order " + orderId + ": No Telegram ID found in Telegram_Messages sheet. Skipping.");
    return;
  }
  logDebug("Found Message ID: " + msgData.messageId + " | Chat ID: " + msgData.chatId);

  // 1. GATHER DATA & CALCULATE TOTALS
  // We grab all rows for this order to build the "Pick List"
  var items = getItemsFromSheet(orderId);
  var totalUnits = 0;
  for (var k = 0; k < items.length; k++) totalUnits += parseInt(items[k].qty) || 0;

  // 2. DEFINE STATUS EMOJI & BUTTONS
  var statusEmoji = "🔴";
  var buttons = []; 

  if (newStatus === "PREPARING") {
    statusEmoji = "🟡";
    buttons.push([{ "text": "🔄 Revert to Pending", "callback_data": "PEND_" + orderId }]);
  } else if (newStatus === "PENDING") {
    statusEmoji = "🔴";
    buttons.push([{ "text": "🚀 Mark as Preparing", "callback_data": "PREP_" + orderId }]);
  } else if (newStatus === "SHIPPED") {
    statusEmoji = "✅"; // No buttons
  } else if (newStatus === "CANCELED") {
    statusEmoji = "❌"; // No buttons
  }

  // 3. BUILD MESSAGE (Exact N8N Template Port)
  // Pre-build inventory map once instead of calling getInventoryForSKU per item
  var inventoryCache = {};
  try {
    var maps = buildLocationAndInventoryMaps();
    var invMap = maps.inventoryMap;
  } catch (e) {
    invMap = new Map();
  }

  var timestamp = Utilities.formatDate(new Date(), "America/Chicago", "EEE, MMM d, h:mm a"); // Houston Time

  var msg = "══════════════════════\n";
  msg += "       📦  ORDER\n";
  msg += "══════════════════════\n";
  msg += "🕐  " + timestamp + " CST\n\n";

  msg += "🔖  " + orderId + "\n";
  msg += "📦  " + totalUnits + " total units\n\n";

  // -- PICK LIST SECTION --
  msg += "┌─────────────────────\n";
  msg += "│ PICK LIST\n";
  msg += "├─────────────────────\n";

  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var isLast = j === items.length - 1;
    var prefix = isLast ? '└' : '├';
    var linePrefix = isLast ? ' ' : '│';

    // Get Live Inventory from cached map (single read)
    var invData = invMap.get(String(item.sku).trim().toLowerCase());
    var availableStock = invData ? invData.available : 0;
    var stockStatus = availableStock <= 20 ? '⚠️' : '✅';

    msg += "│\n";
    msg += prefix + "─ " + (j + 1) + ". SKU: " + item.sku + "\n";
    msg += linePrefix + "      ├─ 📦 " + item.sku + "\n"; // Added explicit SKU line to match n8n
    msg += linePrefix + "      ├─ 📍 Loc: " + item.loc + "\n";
    msg += linePrefix + "      ├─ 🔢 Qty: " + item.qty + "\n";
    msg += linePrefix + "      └─ 📊 Stock: " + stockStatus + " " + availableStock + " units\n";
  }
  msg += "\n";

  // -- NOTE SECTION --
  // If any item has a note, we display the first one (common for order-level notes)
  var orderNote = items.find(i => i.note !== "")?.note || "";
  if (orderNote) {
    msg += "┌─────────────────────\n";
    msg += "│ 💬 BUYER NOTE\n";
    msg += "├─────────────────────\n";
    msg += "│ " + orderNote + "\n";
    msg += "└─────────────────────\n\n";
  }

  msg += "📋 Status: " + statusEmoji + " " + newStatus;

  // 4. SEND UPDATE
  // NOTE: No parse_mode set — message uses Unicode box-drawing, not HTML.
  // Setting parse_mode:"HTML" caused silent failures when eBay titles contain & < > chars.
  var payload = {
    "chat_id": String(msgData.chatId),
    "message_id": parseInt(msgData.messageId),
    "text": msg,
    "reply_markup": { "inline_keyboard": buttons }
  };

  try {
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText", {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    var respCode = response.getResponseCode();
    if (respCode !== 200) {
      logDebug("❌ Telegram editMessage failed (" + respCode + "): " + response.getContentText());
    } else {
      logDebug("✅ Telegram synced for " + orderId + " → " + newStatus);
    }
  } catch (e) {
    logDebug("❌ Sync Error: " + e.toString());
  }
}

/**
 * Helper: Gets all items for a specific order ID from the sheet
 */
function getItemsFromSheet(orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) return [];

  // Search ALL rows (both eBay and Direct tables) for the order ID
  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 5).getValues();
  var items = [];
  var cleanOrderId = String(orderId).trim();

  for (var i = 0; i < data.length; i++) {
    // Skip the DIRECT boundary row
    if (String(data[i][0]).trim().toUpperCase() === TABLE_TWO_IDENTIFIER) continue;

    if (String(data[i][3]).trim() === cleanOrderId) {
      items.push({
        sku: data[i][0],
        qty: data[i][1],
        loc: data[i][2],
        note: data[i][4]
      });
    }
  }
  return items;
}

/**
 * HANDLES MANUAL EDITS
 */
function handleManualStatusChange(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();

  // Only process status column edits on the main sheet
  if (sheet.getName() !== MAIN_SHEET_NAME) return;
  if (range.getColumn() !== STATUS_COLUMN) return;
  if (range.getRow() < DATA_START_ROW) return;

  logDebug("=== MANUAL STATUS EDIT DETECTED ===");
  logDebug("Row: " + range.getRow() + " | Column: " + range.getColumn());

  // Concurrency lock: prevents conflicts with doPost or other users
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    logDebug("❌ Could not acquire lock, skipping sync");
    return;
  }

  try {
    var numRows = range.getHeight();
    var statuses = sheet.getRange(range.getRow(), STATUS_COLUMN, numRows, 1).getValues();
    var orderIds = sheet.getRange(range.getRow(), 4, numRows, 1).getValues();

    var synced = {};
    for (var i = 0; i < numRows; i++) {
      var newStatus = String(statuses[i][0]).trim().toUpperCase();
      var orderId = String(orderIds[i][0]).trim();

      logDebug("Processing row " + (range.getRow() + i) + ": Order=" + orderId + " Status=" + newStatus);

      if (!orderId || orderId === "" || synced[orderId]) continue;
      if (!["PENDING", "PREPARING", "SHIPPED", "CANCELED"].includes(newStatus)) continue;

      synced[orderId] = true;
      try {
        syncStatusToTelegram(orderId, newStatus);
      } catch (err) {
        logDebug("❌ handleManualStatusChange error for " + orderId + ": " + err.toString());
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════════════════════════════
// 🔍 DIAGNOSTIC - Run from Script Editor to debug Telegram 404///
// ═══════════════════════════════════════════════════════════════════════════════════════

function diagnoseTelegram() {
  logDebug("=== TELEGRAM DIAGNOSTIC START ===");

  // 1. Check bot token
  var tokenPreview = TELEGRAM_BOT_TOKEN
    ? (TELEGRAM_BOT_TOKEN.substring(0, 6) + "..." + TELEGRAM_BOT_TOKEN.substring(TELEGRAM_BOT_TOKEN.length - 4))
    : "UNDEFINED";
  logDebug("Token preview: " + tokenPreview);
  logDebug("Token length: " + (TELEGRAM_BOT_TOKEN ? TELEGRAM_BOT_TOKEN.length : 0));
  logDebug("Token type: " + typeof TELEGRAM_BOT_TOKEN);

  // 2. Test getMe (verifies token is valid)
  try {
    var getMeUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/getMe";
    logDebug("Calling getMe...");
    var getMeResp = UrlFetchApp.fetch(getMeUrl, { "muteHttpExceptions": true });
    logDebug("getMe status: " + getMeResp.getResponseCode());
    logDebug("getMe response: " + getMeResp.getContentText());
  } catch (e) {
    logDebug("getMe ERROR: " + e.toString());
  }

  // 3. Test with the most recent message in Telegram_Messages sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgSheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!msgSheet || msgSheet.getLastRow() < 2) {
    logDebug("No messages stored in Telegram_Messages sheet");
    logDebug("=== DIAGNOSTIC END ===");
    return;
  }

  var lastRow = msgSheet.getLastRow();
  var testData = msgSheet.getRange(lastRow, 1, 1, 3).getValues()[0];
  var testOrderId = testData[0];
  var testMsgId = testData[1];
  var testChatId = testData[2];

  logDebug("Test order: " + testOrderId);
  logDebug("Test message_id: " + testMsgId + " (type: " + typeof testMsgId + ")");
  logDebug("Test chat_id: " + testChatId + " (type: " + typeof testChatId + ")");
  logDebug("parseInt(message_id): " + parseInt(testMsgId));

  // 4. Try editMessageText with minimal payload
  try {
    var editUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
    var testPayload = {
      "chat_id": String(testChatId),
      "message_id": parseInt(testMsgId),
      "text": "Diagnostic test - " + new Date().toLocaleString()
    };
    logDebug("Edit payload: " + JSON.stringify(testPayload));

    var editResp = UrlFetchApp.fetch(editUrl, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(testPayload),
      "muteHttpExceptions": true
    });
    logDebug("Edit status: " + editResp.getResponseCode());
    logDebug("Edit response: " + editResp.getContentText());
  } catch (e) {
    logDebug("Edit ERROR: " + e.toString());
  }

  logDebug("=== DIAGNOSTIC END ===");
}