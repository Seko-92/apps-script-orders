// =======================================================================================
// ORDER_SERVICE.gs - COMPLETE with Hidden Sheet Message ID Storage
// =======================================================================================

// Note: TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are now defined in Secrets.js
// Make sure Secrets.js is uploaded to your Apps Script project
var HIDDEN_SHEET_NAME = "Telegram_Messages"; // Hidden sheet for message IDs

/**
 * The "Front Door" for n8n - Receives POST requests
 */
function doPost(e) {
  // 🔒 LOCK: Prevent race conditions (Crucial for n8n stability)
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Server Busy"})).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var payload = JSON.parse(e.postData.contents);

    // --- PRESERVED TELEGRAM & STATUS ACTIONS ---
    if (payload.action === 'storeMessageId') return storeMessageId(payload.orderId, payload.messageId, payload.chatId);
    if (payload.action === 'notifyShipped') return notifyTelegramShipped(payload.orderId);
    if (payload.callback_query) return handleTelegramCallback(payload);
    
    if (payload.action === 'updateOrderStatus') {
      var result = findAndUpdateOrder(payload.orderId, payload.newStatus);
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success",
        "found": result.found,
        "orderId": payload.orderId,
        "newStatus": payload.newStatus
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'updateStatus') {
      return updateStatus(payload.rowNumber, payload.status);
    }

    // --- IMPROVED ORDER INSERTION LOGIC ---
    var orders = payload.orders || [];
    var results = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    
    // 1. Find the eBay boundary (Don't look past the DIRECT table)
    var fullData = sheet.getDataRange().getValues();
    var ebayBoundaryRow = fullData.length; 

    for (var i = 0; i < fullData.length; i++) {
      var rowStr = fullData[i].join("||").toUpperCase();
      if (rowStr.indexOf("DIRECT") > -1 && rowStr.length < 200) {
        ebayBoundaryRow = i; 
        break;
      }
    }

    // 2. Identify existing orders in the eBay section ONLY
    var existingEbaySignatures = new Set();
    for (var j = DATA_START_ROW - 1; j < ebayBoundaryRow; j++) {
      var existingSku = String(fullData[j][0] || "").trim().toUpperCase();
      var existingOrder = String(fullData[j][3] || "").trim();
      if (existingSku && existingOrder) {
        existingEbaySignatures.add(existingOrder + "|" + existingSku);
      }
    }

    var addedCount = 0;

    orders.forEach(function(item) {
      var sku = String(item.SKU || "").trim().toUpperCase();
      var salesOrder = String(item["SALES ORDER"] || "").trim();
      
      // Validation: Fixes the "wrong/broken sales order" issue
      if (!salesOrder || salesOrder.toLowerCase() === "undefined" || salesOrder.length < 3) {
        results.push("Skipped: Invalid ID (" + salesOrder + ")");
        return;
      }

      var currentSignature = salesOrder + "|" + sku;

      // 3. Duplicate Check (Targeted to eBay section)
      if (existingEbaySignatures.has(currentSignature)) {
        results.push("Skipped: Duplicate " + salesOrder);
        return;
      }

      // 4. Insert Row
      sheet.insertRowBefore(DATA_START_ROW);
      var location = getSingleLocation(sku.toLowerCase());
      var rowData = [[sku, item.QTY || 1, location, salesOrder, item.NOTE || "", "PENDING"]];
      sheet.getRange(DATA_START_ROW, 1, 1, 6).setValues(rowData);
      
      // Update local set so same-batch duplicates are also caught
      existingEbaySignatures.add(currentSignature);
      addedCount++;
      results.push("Added: " + salesOrder);
    });

    if (addedCount > 0) {
      updateOrderStatsInSheet();
      updateLastSyncTimestamp();
      sortTableByStatusAndLocation(1);
    }

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "added": addedCount, 
      "details": results
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock(); // ALWAYS release the lock
  }
}

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

/**
 * Store message ID after Telegram sends message
 * Called by n8n after sending Telegram message
 */
function storeMessageId(orderId, messageId, chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  // 1. Check if sheet exists
  if (!sheet) {
    return ContentService.createTextOutput("Error: Sheet [" + HIDDEN_SHEET_NAME + "] not found!");
  }
  
  // 2. Log what we are trying to save
  logDebug("Attempting to store: Order=" + orderId + ", Msg=" + messageId);
  
  // 3. Perform the save
  try {
    sheet.appendRow([
      String(orderId), 
      String(messageId), 
      String(chatId), 
      new Date()
    ]);
    
    // 4. Return confirmation to n8n
    return ContentService.createTextOutput("✅ Success: Added Order " + orderId + " to sheet.");
  } catch (e) {
    return ContentService.createTextOutput("❌ Script Error: " + e.toString());
  }
}

/**
 * Get message ID for an order
 */
function getMessageId(orderId) {
  try {
    var sheet = getHiddenSheet();
    var data = sheet.getDataRange().getValues();
    
    logDebug("Searching for message ID for order: " + orderId);
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        logDebug("✅ Found message ID: " + data[i][1]);
        return {
          messageId: data[i][1],
          chatId: data[i][2],
          timestamp: data[i][3]
        };
      }
    }
    
    logDebug("❌ No message ID found for order: " + orderId);
    return null;
    
  } catch (e) {
    logDebug("Error getting message ID: " + e.toString());
    return null;
  }
}

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
// ═══════════════════════════════════════════════════════════════════════════════════════
// 🚀 NOTIFICATION WORKERS
// ═══════════════════════════════════════════════════════════════════════════════════════

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
    'message_id': String(msgData.messageId), // Ensure it's a string
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
    // Get current message text to preserve order details
    var getMessage = UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + 
      "/getUpdates?offset=-1",
      { muteHttpExceptions: true }
    );
    
    // Build new message text (keep original content, just update status)
    // We'll rebuild from the original message structure
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
      if (currentStatus === "SHIPPED") {
        logDebug("Row " + actualRow + " is already SHIPPED. Cannot revert.");
        return { found: true, count: 0, currentStatus: "SHIPPED" };
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
  
  if (newStatus === "SHIPPED") {
    keyboard = { "inline_keyboard": [] };  // NO BUTTONS
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
  sheet.insertRowBefore(DATA_START_ROW);
  sheet.getRange(DATA_START_ROW, 1, 1, 6).setValues([[sku, qty, location, salesOrder, "", "PENDING"]]);
  updateOrderStatsInSheet();
  updateLastSyncTimestamp();
  return "Added: " + salesOrder;
}

function getOrderStats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) throw new Error("Main sheet not found");
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { pending: 0, preparing: 0, shipped: 0 };
  var statuses = sheet.getRange(DATA_START_ROW, 6, lastRow - DATA_START_ROW + 1, 1).getValues().flat();
  var stats = { pending: 0, preparing: 0, shipped: 0 };
  statuses.forEach(function(s) {
    s = String(s).trim().toUpperCase();
    if (s === 'PENDING') stats.pending++;
    else if (s === 'PREPARING') stats.preparing++;
    else if (s === 'SHIPPED') stats.shipped++;
  });
  return stats;
}

function updateOrderStatsInSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  var stats = getOrderStats();
  var text = "🔴 Pending: " + stats.pending + "   🟡 Preparing: " + stats.preparing + "   🟢 Shipped Today: " + stats.shipped;
  var range = sheet.getRange('F1:H1');
  try { range.breakApart(); } catch(e) {}
  range.merge();
  sheet.getRange('F1').setValue(text).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#212121').setFontColor('#FFFFFF').setWrap(false);
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
  var statusOrder = { 'PENDING': 1, 'PREPARING': 2, 'SHIPPED': 3, '': 4 };
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
  for (var i = data.length - 1; i >= 0; i--) {
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
// 💾 DATABASE WORKERS - These actually talk to the Telegram_Messages sheet
// ═══════════════════════════════════════════════════════════════════════════════════════

function storeMessageId(orderId, messageId, chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  
  // If the sheet doesn't exist, create it
  if (!sheet) {
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    sheet.appendRow(["Order ID", "Message ID", "Chat ID", "Timestamp"]);
  }
  
  // Add the info to the sheet
  sheet.appendRow([orderId, messageId, chatId, new Date()]);
  
  return ContentService.createTextOutput("✅ Stored successfully");
}

function getMessageId(orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  var targetId = String(orderId).trim().toLowerCase();
  
  // Look from bottom to top (get most recent message for this order)
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim().toLowerCase() === targetId) {
      return {
        messageId: data[i][1],
        chatId: data[i][2]
      };
    }
  }
  return null;
}

/**
 * Main function to sync a sheet status change to Telegram.
 * Rebuilds the message to show the new status and removes all buttons.
 */
/**
/**
 * Synchronizes a status change from the sheet to Telegram.
 * Rebuilds the "Elegant Mobile-Friendly Design" from the n8n template.
 */
function syncStatusToTelegram(orderId, newStatus) {
  var msgData = getMessageId(orderId);
  if (!msgData) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  // 1. GATHER ALL ITEMS FOR THIS ORDER ID
  var items = [];
  var buyerNote = "";
  for (var i = DATA_START_ROW - 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === String(orderId).trim()) {
      items.push({
        sku: data[i][0],
        qty: data[i][1],
        loc: data[i][2],
        note: data[i][4]
      });
      if (data[i][4]) buyerNote = data[i][4]; // Take note from row
    }
  }

  if (items.length === 0) return;

  // 2. REPLICATE N8N ELEGANT DESIGN
  var msg = "══════════════════════\n";
  msg += "         📦  ORDER UPDATE\n";
  msg += "══════════════════════\n\n";
  
  msg += "🔖  " + orderId + "\n";
  // Note: Buyer Name/City isn't in your main sheet columns, so we skip or use Order ID
  msg += "\n";
  
  msg += "┌─────────────────────\n";
  msg += "│  📋  PICK LIST\n";
  msg += "├─────────────────────\n";

  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var isLast = j === items.length - 1;
    var prefix = isLast ? '└' : '├';
    var linePrefix = isLast ? ' ' : '│';
    
    msg += "│\n";
    msg += prefix + "─ " + (j + 1) + ". " + item.sku + "\n";
    msg += linePrefix + "      ├─ 📦 SKU: " + item.sku + "\n";
    msg += linePrefix + "      ├─ 📍 Loc: " + item.loc + "\n";
    msg += linePrefix + "      └─ 🔢 Qty: " + item.qty + "\n";
  }
  msg += "\n";

  if (buyerNote) {
    msg += "┌─────────────────────\n";
    msg += "│ 💬 BUYER NOTE\n";
    msg += "├─────────────────────\n";
    msg += "│ " + buyerNote + "\n";
    msg += "└─────────────────────\n\n";
  }

  var statusEmoji = newStatus === "PREPARING" ? "🟡" : (newStatus === "SHIPPED" ? "✅" : "🔴");
  msg += "📋 Status: " + statusEmoji + " " + newStatus;

  // 3. BUILD BUTTONS (Keep "Revert" for Preparing, "Mark Prep" for Pending)
  var buttons = [];
  if (newStatus === "PREPARING") {
    buttons.push([{ "text": "↩️ Revert to Pending", "callback_data": "PEND_" + orderId }]);
  } else if (newStatus === "PENDING") {
    buttons.push([{ "text": "🚀 Mark as Preparing", "callback_data": "PREP_" + orderId }]);
  }

  var payload = {
    "chat_id": String(msgData.chatId),
    "message_id": String(msgData.messageId),
    "text": msg,
    "parse_mode": "HTML",
    "reply_markup": { "inline_keyboard": buttons }
  };

  UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText", {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  });
}

/**
 * Automatically triggers when a cell in the sheet is changed manually.
 */
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  
  // Only trigger if we are on the Main Sheet and Column F (6)
  if (sheet.getName() === MAIN_SHEET_NAME && range.getColumn() === 6) {
    var row = range.getRow();
    if (row < DATA_START_ROW) return;

    var newStatus = range.getValue();
    var orderId = sheet.getRange(row, 4).getValue(); // Get Order ID from Column D

    if (orderId && (newStatus === "PREPARING" || newStatus === "PENDING")) {
      syncStatusToTelegram(orderId, newStatus);
    }
  }
}