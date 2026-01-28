// =======================================================================================
// SECRETS.TEMPLATE.js - Template for Secret Configuration
// =======================================================================================
// 📋 SETUP INSTRUCTIONS:
// 1. Copy this file and rename it to "Secrets.js"
// 2. Fill in your actual values below
// 3. NEVER commit Secrets.js to git (it's already in .gitignore)
// =======================================================================================

/**
 * Telegram Bot Configuration
 * Get your bot token from @BotFather on Telegram
 * Tutorial: https://core.telegram.org/bots#6-botfather
 */
var TELEGRAM_BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN_HERE";
var TELEGRAM_CHAT_ID = "YOUR_TELEGRAM_CHAT_ID_HERE"; // Can be negative number for groups

/**
 * N8N Webhook Configuration
 * This is the webhook URL from your n8n workflow
 * Example: https://your-domain.com/webhook/your-webhook-id
 */
var N8N_WEBHOOK_URL = "YOUR_N8N_WEBHOOK_URL_HERE";

/**
 * Google Apps Script Web App URL
 * How to get this:
 * 1. In Apps Script Editor, click "Deploy" > "New Deployment"
 * 2. Select type "Web App"
 * 3. Set "Execute as" to "Me"
 * 4. Set "Who has access" to "Anyone"
 * 5. Click "Deploy" and copy the Web App URL
 */
var WEB_APP_URL = "YOUR_WEB_APP_URL_HERE";

/**
 * Google Apps Script Project ID
 * Found in: Apps Script Editor > Project Settings > IDs > Script ID
 */
var SCRIPT_ID = "YOUR_SCRIPT_ID_HERE";
