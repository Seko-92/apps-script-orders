# Order Management System

A production warehouse order-management system built on **Google Apps Script**, integrating **Google Sheets**, a **Telegram bot**, and **n8n** automation, with deep two-way integrations into **eBay** (Trading + Fulfillment APIs) and **Zoho Inventory**.

Orders flow in from two channels — eBay and direct sales — get tracked through a `PENDING → PREPARING → SHIPPED` (or `CANCELED`) lifecycle, surface as interactive cards in a Telegram group for warehouse staff, and exit through a printed picking list. Inventory, pricing, and kit composition stay synchronized across eBay, Zoho, and the sheet automatically.

> This repository is a code backup and a record of the work. Secrets, real endpoint URLs, spreadsheet IDs, and business-identifying details are intentionally excluded.

---

## Architecture & data flow

```
        ┌──────────┐        ┌──────────────┐
        │   eBay   │        │ Zoho (direct │
        │  orders  │        │  sales / SOs)│
        └────┬─────┘        └──────┬───────┘
             │ Trading API         │ Webhooks
             ▼                     ▼
        ┌─────────────────────────────────┐
        │              n8n                 │  polling, enrichment,
        │   (workflows + proxies)          │  scheduling, error sink
        └────────────────┬────────────────┘
                         │ POST { action, ... }
                         ▼
        ┌─────────────────────────────────┐
        │   doPost()  —  Apps Script web   │ ──► Google Sheet (orders,
        │   app  (single dispatch surface) │     inventory mirrors, logs)
        └────────────────┬────────────────┘
                         │ status changes
                         ▼
        ┌─────────────────────────────────┐
        │  Telegram bot (interactive       │ ◄── button callbacks
        │  order cards, inline updates)    │
        └─────────────────────────────────┘
```

**The web app exposes one POST surface (`doPost`)** that dispatches on an `action` field. Current actions include: `storeMessageId`, `notifyShipped`, `updateStatus`, `updateOrderStatus`, `updateMiRows`, `recomputeHand`, `runPriceAudit`, `writeZohoStock`, `zohoSalesOrder`, `zohoBackfillSalesOrder`, and `zohoKitUpdated`. n8n and Zoho call these; the sheet's own UI calls server functions directly.

---

## Key capabilities

### Order pipeline
- **Two stacked tables on one sheet** — an eBay table and a `DIRECT` table separated by a movable divider row, with a strict boundary contract that all row math depends on.
- **Canonical status transitions** — every status change (Telegram callback, webhook, manual cell edit, bulk action) flows through one function with validation, locking, terminal-state guards, and status-alias normalization (e.g. eBay's `Cancelled` → canonical `CANCELED`).
- **Telegram interactivity** — order cards with inline buttons; status edits in the sheet sync back to the original message in place (no message spam).

### Inventory & stock
- **HAND (available stock) routing** — resolves per row from the freshest source: eBay-channel rows from a Master Inventory mirror, direct/manual rows from a live Zoho stock mirror, with graceful fallback.
- **Per-order live qty refresh** — fetches current quantity for an order's SKUs at arrival time so stock math is correct the moment an order lands.
- **Zoho stock mirror** — a scheduled push keeps a SKU-keyed sheet in sync with Zoho's live `available_stock`, refreshed every couple of minutes during work hours.
- **Out-of-Stock tracker** — weekly smart-merge sheet that preserves "first seen" dates so chronic out-of-stock items are visible.
- **Prep Queue** — a warehouse employee's personal restock/repackage to-do list (separate from customer orders).

### Zoho integration suite
- **Direct sales-order pipeline** — Zoho SOs auto-mirror into a Pending sheet via webhook; a picker-driven Pull modal is the single, explicit bridge into the working table (with a per-line diff: new / changed / removed).
- **Invoice lookup** — pull by sales-order number *or* invoice number.
- **Backfill** — fetch a missed/historical SO on demand through an n8n proxy.
- **Kit registry sync** — kit composition is parsed from Zoho item descriptions and live-updated via webhook; manual CSV import remains as a safety net.
- **Kit expansion** — a modal expands multi-component kits into individual pickable rows, with per-component exclusion and "spares for us" quantities.

### Pricing
- **Price Audit** — compares Zoho selling prices against live eBay prices across the catalog, bucketing results into price drift / inactive-listing candidates / out-of-stock, sorted for bulk review.
- **Price write-back** — push eBay's price into Zoho (single-item safe test + selection-driven bulk push), passphrase-gated, with an append-only push log for full price-change history.

### Operational visibility
- **Ops Cockpit** — live "shipped today / received today / oldest pending" readouts, a day-tape timeline, a system-pulse heartbeat, and a per-channel work queue, all in the sidebar.
- **Activity Log** — append-only event log (received / status transitions / printed) with 90-day rolling retention; powers the cockpit and order lookup.
- **Alerts** — actionable counts (paid shipping awaiting label, international, low stock, location-not-found, out of stock, prep queue, new Zoho SOs) with click-to-jump.
- **Order Lookup** — a customer-service consolidated view: every sheet row plus the full activity timeline for one order ID or invoice.
- **Print fulfillment** — a spec-sheet-styled picking list with batch metrics, sign-off, and a closing audit page.

---

## Project structure

```
.
├── Config.js              # Back-compat constants (aliases over Schema)
├── Schema.js              # Single source of truth: columns, status enum, boundary contract
├── Main.js                # onOpen / onEdit / onChange + installable-trigger dispatch
├── Helpers.js             # Boundary detection, HAND recompute, MI updates
├── Secrets.js             # Credentials & endpoint URLs (gitignored — NOT in repo)
│
│ ── Order pipeline ──
├── OrderService.js        # doPost dispatch, Telegram helpers, sort, manual sync
├── StatusService.js       # Canonical updateOrderStatus()
├── N8NIntegration.js      # Outbound webhook triggers
├── LiveSync.js            # onEdit-driven SKU → location + inventory lookup
├── RowManagement.js       # Row add/delete, boundary protection, duplicate highlighting
│
│ ── Inventory & stock ──
├── LocationService.js     # SKU → location maps
├── ZohoStock.js           # Zoho stock mirror + HAND routing
├── OutOfStock.js          # Out-of-Stock tracker sheet
├── PrepQueue.js           # Personal prep/restock to-do sheet
├── LocationUpdate.js      # Location-update entry sheet
│
│ ── Zoho integration ──
├── ZohoSalesOrders.js     # Pending SO mirror + void/shipped propagation
├── ZohoPull.js            # Pull modal — diff + all-or-nothing apply
├── ZohoPullModal.html
├── KitRegistry.js         # Kit composition registry (Zoho-synced)
├── KitExpansion.js        # Kit expansion engine
├── KitExpansionModal.html
│
│ ── Pricing ──
├── PriceAudit.js          # eBay-vs-Zoho price drift / hygiene audit
├── PriceWriteback.js      # eBay → Zoho price push (single + bulk)
├── PricePushModal.html
│
│ ── Fulfillment & print ──
├── FulfillmentService.js  # Picking-list prep + batch metrics
├── PrintFulfillment.html  # Print template
│
│ ── Dashboards & UI ──
├── Sidebar.html           # Control panel (zoned, with Ops Cockpit)
├── UIService.js           # Sidebar service helpers (consolidated tick)
├── Dashboard.html
├── DashboardService.js
├── Alerts.js              # Actionable alert counts + jump-to-rows
├── ActivityLog.js         # Append-only event log + cockpit data
├── OrderLookup.js         # Consolidated order + timeline lookup
├── StatsHelper.js         # Stat formatters
├── ApiMonitor.js          # API-usage sidebar readout
│
│ ── Theming & diagnostics ──
├── BrandTheme.js          # Visual design system (theme, banding, conditional formats)
├── SheetInspector.js      # Diagnostic dump of merges / formats / CF rules
├── VisualLab.js           # Visual experiments (lab sheet)
├── Snake.html             # Arcade easter egg
│
│ ── Legacy (audit before relying on) ──
├── Timestampfeature.js    # Deprecated
├── CopyDesign.js          # Legacy/speculative
└── updateStatus.js        # Small status helper
```

---

## Tech stack

| Layer | Technology |
|---|---|
| App logic | Google Apps Script (V8) |
| Data store / UI | Google Sheets + HtmlService (sidebar & modals) |
| Automation | n8n (self-hosted, Docker) |
| Messaging | Telegram Bot API |
| Marketplace | eBay Trading API + Fulfillment API |
| ERP / inventory | Zoho Inventory API (OAuth 2.0) |
| Deploy tooling | clasp |

---

## Setup

### 1. Get the code into an Apps Script project

```bash
npm install -g @google/clasp
clasp login
# point .clasp.json at your container-bound script, then:
clasp push
```

The script is **container-bound** to a Google Sheet — most code runs in that context (menu, sidebar, `onEdit`).

### 2. Configure secrets

Secrets live in `Secrets.js`, which is **gitignored and not in this repo**. Create it locally with these constants (values are yours to fill in):

```js
// Secrets.js  — DO NOT COMMIT
const TELEGRAM_BOT_TOKEN  = '...';   // from @BotFather
const TELEGRAM_CHAT_ID    = '...';   // target group/channel id
const WEB_APP_URL         = '...';   // your deployed /exec URL
const SPREADSHEET_ID      = '...';   // the bound spreadsheet id
const APP_SECRET_TOKEN    = '...';   // shared secret for authed webhooks

// n8n webhook endpoints (one per integration)
const N8N_WEBHOOK_URL                     = '...';
const N8N_API_USAGE_WEBHOOK_URL           = '...';
const N8N_STATUS_CHECK_WEBHOOK_URL        = '...';
const N8N_VERIFY_SHIPPED_WEBHOOK_URL      = '...';
const N8N_INVENTORY_LITE_WEBHOOK_URL      = '...';
const N8N_ZOHO_SO_WEBHOOK_URL             = '...';
const N8N_ZOHO_FETCH_WEBHOOK_URL          = '...';
const N8N_ZOHO_PRICE_WRITE_WEBHOOK_URL    = '...';
const N8N_ZOHO_PRICE_BULK_WRITE_WEBHOOK_URL = '...';
```

Some secrets are **not** stored in source at all — e.g. the bulk price-push passphrase lives in a Script Property (`PRICE_PUSH_PASSPHRASE`), set once from the editor, never in the HTML or in git.

### 3. Deploy as a web app

> ⚠️ **Deployment rule:** to update the production URL, always use **Manage Deployments → Edit → New version**. **Never** use **Deploy → New Deployment** — that mints a fresh URL and orphans every external integration that hardcodes the old one.

1. **Deploy → Manage Deployments → Edit → New version**
2. Execute as **Me**, access as appropriate for your integrations.
3. The `/exec` URL stays stable across versions — store it as `WEB_APP_URL`.

### 4. Wire up integrations

- **Telegram:** run `setWebhook()` once to register the web app with the bot.
- **n8n:** point each workflow's HTTP node at `WEB_APP_URL` and authenticate with `APP_SECRET_TOKEN`.
- **Zoho:** create workflow rules (Sales Order / Invoice / Item) that POST to the corresponding n8n proxy. External callers must route through n8n — Apps Script's `/exec` 302-redirects and can't read request headers, so n8n handles redirect-following and header auth.

### 5. Expected sheets

| Sheet | Purpose |
|---|---|
| **All orders** | The two-table order sheet (eBay + DIRECT) |
| **Master Inventory** | eBay-sourced inventory mirror (`sku`, `listingStatus`, qty, price, …) |
| **Zoho Stock** | Live Zoho `available_stock` mirror |
| **Pending Sales Orders** | Zoho SO mirror (Pull source) |
| **Kit Registry** | Kit → component composition |
| **Activity Log** | Append-only event log (90-day retention) |
| **Prep Queue / Out of Stock / Price Audit / Price Push Log / …** | Operational sheets |

The **All orders** data row is 10 columns wide:

```
A SKU · B QTY · C LOCATION · D SALES ORDER · E NOTE ·
F STATUS · G HAND · H LEFT · I SHIPPING · J SHIP COST
```

---

## Conventions & invariants

A few load-bearing rules that keep the system from silently breaking:

- **The `DIRECT` boundary cell must stay exactly `"DIRECT"`.** Boundary detection is a strict equality match; decorative text in that cell breaks sorting, row inserts, and live sync. Visual styling goes in adjacent merged cells or via number-format display tricks, never the underlying value.
- **Cells with data validation (status dropdown, pick-ID dropdowns) are read-only to code** except through their validation list.
- **Range read/write must use the full data width** so trailing columns travel with their owning row during sorts.
- **External API responses are never cached whole into a cell** (50k-char-per-cell limit) — always slimmed to a known field set.
- **eBay XML parsing uses word-boundary regex** for attributed tags, and regex metacharacters are double-escaped inside template literals.
- **Programmatic writes don't fire `onEdit`** — any code path that inserts/edits rows explicitly re-triggers the relevant handlers (duplicate highlight, kit markers, etc.).

---

## Security notes

- `Secrets.js`, deployment configs, and any data exports are **gitignored**; never commit them.
- The web app is reachable by URL — gate sensitive actions with `APP_SECRET_TOKEN` (header auth via n8n) and, for destructive operations, an additional server-checked passphrase in Script Properties.
- Rotate the Telegram bot token and any OAuth refresh tokens if exposed.
- Treat every external-facing doc as a potential leak surface — keep real IDs, URLs, and tokens out of source.

---

## Status

Private project. Built and iterated in production over several months. Not licensed for redistribution.
