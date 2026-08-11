# eBay Location Check Proxy — Location Update live verify

Serves the **Location Update** sheet's new `⇊ Fetch eBay Locations` sidebar button
(2026-08-11). Apps Script sends `{ items: [{sku, itemId}] }`; this workflow calls
Trading API **GetItem** per item, extracts the **"Model Year"** item specific (the
field this business uses for the physical location on eBay — the same field MAIN's
sync mirrors into Master Inventory's `C:Model Year` column), and responds
**synchronously** so the sheet paints match/mismatch verdicts instantly.

**READ-ONLY.** The only eBay call is GetItem. This proxy cannot write anything.

```
Sidebar button → fetchEbayLocationsLive() (LocationUpdate.js)
  → resolves SKU → itemId from Master Inventory (bounded read)
  → POST https://n8n.yassinqurabi.com/webhook/ebay-location-check   (X-API-Token)
  → n8n: validate → eBay OAuth → GetItem ×N (chunks of 5, 200ms between chunks)
  → responds { ok, count, results: [{sku, itemId, ok, location, error}] }
  → Apps Script writes ◉ EBAY LOC + verdict colors in one batched pass
```

## Import & activation (one time, ~3 minutes)

1. **n8n UI → Workflows → Import from File** → `eBay Location Check Proxy.json`.
   ⚠ Import creates a NEW workflow — if you re-import later, delete the old copy
   first (two active workflows on one webhook path is how the board went dark once).
2. On **🔔 Webhook (from Apps Script)**: confirm the existing **Header Auth account**
   credential is bound (X-API-Token = APP_SECRET_TOKEN). If the import didn't
   auto-bind it, pick it from the credential dropdown.
3. On **🔑 Get eBay Access Token**: replace the two `<<PASTE…>>` placeholders with
   the SAME two values used in **MAIN - Master Inventory Full Sync → "1. Get Access
   Token"**:
   - the `Authorization` header value (`Basic …`)
   - the `refresh_token=…` part of the body
   (Or simply copy that whole node from MAIN, paste it here, and rewire it in
   place of the placeholder node.) Placeholders are deliberate — this file must
   never carry real credentials because `deploy/` is not gitignored.
4. **Activate** the workflow. The webhook path `ebay-location-check` must stay
   exactly as-is — it matches `N8N_EBAY_LOCATION_WEBHOOK_URL` in `Secrets.js`.

## Test recipe

1. Location Update sheet has a few rows with SKUs (mix in one junk SKU on purpose).
2. Sidebar → **Location Update** → `⇊ Fetch eBay Locations`.
3. Expect within ~5–15s: `◉ EBAY LOC` fills — green where eBay agrees with the
   typed LOCATION, red where it disagrees (or the listing's location is blank),
   dim gray `NOT ON EBAY` on the junk SKU. The D1 header note records the fetch
   time; the sidebar status bar shows the tally.
4. n8n Executions should show one run with a green **📤 Respond to Caller**.

## Notes

- **Quota:** shared Trading API pool (5,000/day). One GetItem per unique SKU per
  press — a 50-row sheet ≈ 1% of the pool. Sidebar enforces a 15s cooldown.
- **Caps:** Apps Script sends at most 120 unique items (`LOCATION_UPDATE.fetchCap`);
  the validate node independently hard-caps at 150. Both on purpose — a guard that
  lives only in the caller is a guard a future caller forgets.
- **404 from the webhook = workflow not Active** (routing, not auth). 401/403 =
  Header Auth mismatch.
- **No Apps Script New Version needed** — the Apps Script half is editor-bound
  (sidebar → `google.script.run`); `clasp push` is the whole Apps Script deploy.
