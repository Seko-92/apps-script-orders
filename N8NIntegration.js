// =======================================================================================
// N8N_INTEGRATION.gs - v3.0 with Timestamp Support
// =======================================================================================

// Note: N8N_WEBHOOK_URL is now defined in Secrets.js
// Make sure Secrets.js is uploaded to your Apps Script project 

/**
 * Triggers the n8n Awaiting Shipments workflow via webhook
 * Called from the Sidebar when user clicks "Sync Orders"
 * NOW UPDATES TIMESTAMP IN CELL F2!
 */
function triggerN8NWebhook() {
  if (!N8N_WEBHOOK_URL) {
    return "⚠️ Webhook URL not configured.";
  }
  
  try {
    var options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'followRedirects': true,
      'timeout': 30000,
      'headers': {
        'ngrok-skip-browser-warning': 'true',
        'User-Agent': 'GoogleAppsScript'
      }
    };
    
    var response = UrlFetchApp.fetch(N8N_WEBHOOK_URL, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode === 200) {
      // 1. Update stats and timestamp
      try { updateOrderStatsInSheet(); } catch(e) {} 
      try { updateLastSyncTimestamp(); } catch(e) {}
      
      // 2. ✨ NEW: Update timestamp in cell F2
      try { 
        updateLastOrderTimestamp("F2"); 
      } catch(e) {
        Logger.log("Timestamp update failed: " + e.toString());
      }
      
      // 3. Parse response
      try {
        var data = JSON.parse(responseText);
        if (data.message) return "✅ " + data.message;
        if (data.added) return "✅ Synced! " + data.added + " orders added.";
      } catch(e) {}
      
      return "✅ Sync complete!";
      
    } else if (responseCode === 404) {
      return "❌ 404: Webhook not found. Is n8n workflow ACTIVE?";
    } else if (responseCode === 502) {
      return "❌ 502: ngrok cannot reach n8n. Is n8n running?";
    } else {
      return "⚠️ Error: n8n responded with code " + responseCode;
    }
    
  } catch (error) {
    Logger.log("n8n webhook error: " + error.toString());
    return "❌ Connection error: " + error.message;
  }
}

/**
 * Tests the n8n webhook connection
 */
function testN8NConnection() {
  try {
    var response = UrlFetchApp.fetch(N8N_WEBHOOK_URL, {
      'method': 'get',
      'muteHttpExceptions': true,
      'headers': {
        'ngrok-skip-browser-warning': 'true'
      }
    });
    
    var code = response.getResponseCode();
    if (code === 200) return "✅ Connection successful!";
    return "⚠️ Connection failed. Code: " + code;
  } catch (e) {
    return "❌ Connection failed: " + e.message;
  }
}

function getN8NStatus() {
  return {
    configured: true,
    url: N8N_WEBHOOK_URL
  };
}
