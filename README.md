# Google Apps Script Order Management System

A complete order fulfillment system built with Google Apps Script, integrating Google Sheets, Telegram Bot, and n8n workflows for automated order processing and warehouse management.

## Features

- **Order Syncing**: Automatically sync orders from n8n workflows to Google Sheets
- **Telegram Integration**: Real-time notifications and interactive order management via Telegram
- **Status Management**: Track orders through PENDING → PREPARING → SHIPPED workflow
- **Location Lookup**: Automatic SKU-to-location mapping from Master Inventory
- **Print Fulfillment**: Generate picking lists for warehouse staff
- **Live Updates**: Real-time sheet updates with n8n integration
- **Control Panel**: Interactive sidebar for all operations

## Project Structure

```
Excel Code/
├── Config.js                 # Configuration variables (sheet names, columns)
├── Main.js                   # Entry points and menu setup
├── OrderService.js           # Order management and Telegram integration
├── N8NIntegration.js         # n8n webhook integration
├── LocationService.js        # SKU location lookup
├── FulfillmentService.js     # Print picking lists
├── RowManagement.js          # Row operations
├── LiveSync.js               # Real-time sync features
├── Helpers.js                # Utility functions
├── UIService.js              # UI components
├── Sidebar.html              # Control panel interface
├── PrintFulfillment.html     # Print picking list UI
├── Snake.html                # Fun arcade game
├── Secrets.js                # Secret credentials (NOT in git)
├── Secrets.template.js       # Template for secrets setup
├── .gitignore                # Git ignore rules
└── README.md                 # This file
```

## Setup Instructions

### 1. Clone the Repository

```bash
git clone <your-repo-url>
cd "Excel Code"
```

### 2. Configure Secrets

**IMPORTANT:** Never commit your actual secrets to Git!

1. Copy the template file:
   ```bash
   cp Secrets.template.js Secrets.js
   ```

2. Edit `Secrets.js` and fill in your actual values:
   - **TELEGRAM_BOT_TOKEN**: Get from [@BotFather](https://t.me/BotFather) on Telegram
   - **TELEGRAM_CHAT_ID**: Your Telegram group/channel ID
   - **N8N_WEBHOOK_URL**: Your n8n workflow webhook URL
   - **WEB_APP_URL**: Your deployed Google Apps Script web app URL
   - **SCRIPT_ID**: Your Google Apps Script project ID

### 3. Setup Google Apps Script Project

#### Option A: Using CLASP (Recommended)

1. Install CLASP:
   ```bash
   npm install -g @google/clasp
   ```

2. Login to Google:
   ```bash
   clasp login
   ```

3. Create a new project or clone existing:
   ```bash
   # Create new
   clasp create --title "Order Management System" --type sheets

   # Or clone existing (update scriptId in .clasp.json first)
   clasp clone <YOUR_SCRIPT_ID>
   ```

4. Push your code:
   ```bash
   clasp push
   ```

#### Option B: Manual Upload

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project linked to your Google Sheet
3. Copy each `.js` file's content into separate script files
4. Copy each `.html` file's content into separate HTML files
5. Save the project

### 4. Deploy as Web App

1. In Apps Script Editor, click **Deploy** → **New Deployment**
2. Select type: **Web App**
3. Configure:
   - **Execute as**: Me
   - **Who has access**: Anyone
4. Click **Deploy**
5. Copy the **Web App URL** and add it to your `Secrets.js` as `WEB_APP_URL`

### 5. Setup Telegram Webhook

Run the `setWebhook()` function once to register your web app with Telegram:

1. In Apps Script Editor, select `setWebhook` function
2. Click **Run**
3. Check logs to confirm success

### 6. Configure Your Google Sheet

Your Google Sheet should have these sheets:

#### **All orders** Sheet
- Column A: SKU
- Column B: Qty
- Column C: Location
- Column D: Sales Order ID
- Column E: Note
- Column F: Status (PENDING/PREPARING/SHIPPED)
- Column G: HAND
- Column H: LEFT

#### **Master Inventory** Sheet
- Must have a column header "sku"
- Must have a column header "C:Model Year" (or configure in Config.js)

#### **Settings** Sheet (optional)
- For live sync toggle

#### **Telegram_Messages** Sheet (auto-created)
- Hidden sheet for message ID storage

### 7. Setup n8n Workflows

Configure your n8n workflows to:

1. **POST orders to your Web App URL** with this format:
   ```json
   {
     "orders": [
       {
         "SKU": "ABC123",
         "QTY": 2,
         "SALES ORDER": "ORDER-001",
         "NOTE": "Handle with care"
       }
     ]
   }
   ```

2. **Store Telegram message IDs** after sending:
   ```json
   {
     "action": "storeMessageId",
     "orderId": "ORDER-001",
     "messageId": "12345",
     "chatId": "YOUR_CHAT_ID"
   }
   ```

3. **Notify when shipped**:
   ```json
   {
     "action": "notifyShipped",
     "orderId": "ORDER-001"
   }
   ```

## Usage

### Control Panel

Open the Control Panel from the Google Sheets menu: **⚙️ Control Panel** → **Open Control Panel**

Features:
- Sync orders from n8n
- Update SKU locations
- Mark orders as PREPARING
- Print picking lists
- Sort tables
- Toggle focus mode

### Telegram Commands

The Telegram bot supports interactive buttons:
- **🚀 Mark as Preparing**: Change order status to PREPARING
- **🔄 Revert to Pending**: Change order status back to PENDING

Orders automatically update when marked as SHIPPED in the sheet.

### Manual Operations

You can also:
- Manually edit status in Column F
- Change will sync to Telegram automatically
- Delete rows (message IDs are cleaned up weekly)

## Configuration

Edit `Config.js` to customize:
- Sheet names
- Column positions
- Cell references
- Table identifiers

## Security Notes

- **Never commit `Secrets.js`** to version control
- **Never share your Web App URL publicly** without authentication
- **Rotate your Telegram bot token** if compromised
- **Use environment-specific URLs** for n8n (dev/staging/prod)

## Troubleshooting

### Orders not syncing
1. Check n8n webhook is active
2. Verify Web App URL is correct
3. Check Apps Script execution logs

### Telegram not updating
1. Verify webhook is set: Run `getWebhookInfo()`
2. Check bot token is correct
3. Check message IDs are stored in `Telegram_Messages` sheet

### Location not found
1. Verify "sku" column exists in Master Inventory
2. Check SKU is in lowercase in the lookup
3. Verify header name matches `DB_SKU_HEADER` in Config.js

## Development

### Running Tests

Open Apps Script Editor and run test functions:
- `testN8NConnection()` - Test n8n webhook
- `getWebhookInfo()` - Check Telegram webhook status

### Debugging

Enable debug logging by checking the **Debug Log** sheet (auto-created).

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is private and proprietary.

## Support

For issues or questions, contact the development team.

---

**Built with ❤️ for efficient warehouse operations**
