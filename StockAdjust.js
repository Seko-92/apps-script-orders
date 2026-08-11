// =======================================================================================
// StockAdjust.js — correct a SKU's on-hand quantity in Zoho from a physical count
// =======================================================================================
//
// WHY THIS IS ITS OWN MODULE and not part of PriceWriteback.js:
//
//   · Different endpoint entirely. Price is PUT /items/{id}; stock is
//     POST /inventoryadjustments. Different payload, different failure modes.
//   · Different OAuth scope — ZohoInventory.inventoryadjustments.CREATE, which
//     the price token does NOT carry. Kept as a SEPARATE refresh token on
//     purpose: a bug or leak in the stock path then cannot rewrite prices, and
//     the price path cannot move stock. Each capability revocable on its own.
//   · Different log. The Price Push Log records before→after prices; a stock
//     log wants counted / expected / delta / picker / Zoho's adjustment id.
//   · Deletable. If the floor workflow does not stick, one file and one n8n
//     workflow come out cleanly — the standing rule here is delete code, don't
//     hide it behind a toggle.
//
// THE WORKFLOW IT SERVES (from the floor, 2026-08-11): HAND is near-live, but
// the physical shelf occasionally disagrees. On LOW-QUANTITY items — the only
// ones anyone can count in one go — the picker records what is really there in
// ◩ LEFT. Today that number goes on paper and someone walks to a PC to correct
// Zoho. The board already captures the count (boardSetLeft); this is the other
// half of the loop.
//
// ⚠ ONE SKU AT A TIME, AND AN ABSOLUTE TARGET.
// pushSingleStockAdjustBySku is the primitive: "make Zoho's on-hand equal N".
// Unambiguous, and proven end to end 2026-08-11 (Apps Script → n8n → Zoho, a
// junk SKU taken 0 → 5) before anything was allowed to call it.
//
// It takes a TARGET rather than a picker's shelf count on purpose. Deriving one
// from the other depends on when Zoho deducts a pulled unit, and the board
// cannot know that with certainty — so the board SUGGESTS and the person at the
// shelf CONFIRMS. See boardAdjustStock for the full reasoning.
//
// A batch/review surface for many counts at once is deliberately NOT built yet:
// the capture has to prove itself on the floor first, and that stage is where a
// dedicated Stock Adjustments log sheet belongs. Today the audit trail is
// Zoho's own adjustment record plus the Activity Log.
//
// THE DELTA IS COMPUTED IN THE PROXY, NOT HERE — Zoho's inventory adjustment
// API takes `quantity_adjusted`, a CHANGE, not a new total. So the proxy reads
// the item's CURRENT stock and subtracts, inside one execution. Computing the
// delta here from the Zoho Stock mirror would use a number up to two minutes
// old, and any sale in that window would land a wrong correction.
// =======================================================================================


var STOCK_ADJUST = {
  // Sanity bounds. A physical count outside these is a typo, not a shelf.
  minQty: 0,
  maxQty: 10000,
  // Refuse a correction larger than this unless explicitly forced. A count that
  // moves stock by more than this is far more likely a miscount or a wrong SKU
  // than a real discrepancy — and on a live listing it either oversells or
  // hides sellable stock within minutes of the push.
  maxDelta: 50,
  reason: 'Physical count (warehouse floor)'
};


/**
 * Set a SKU's Zoho on-hand quantity to an explicit target.
 *
 * @param {string} sku          the SKU to correct
 * @param {number} targetOnHand what Zoho's on-hand SHOULD be
 * @param {Object} [opts]       { force: true } to allow a delta beyond maxDelta
 * @returns {Object} { ok, message, sku, itemId, before, target, delta, adjustmentId }
 */
function pushSingleStockAdjustBySku(sku, targetOnHand, opts) {
  opts = opts || {};
  sku = String(sku || '').trim();
  if (!sku) return { ok: false, message: 'No SKU given.' };

  var target = parseFloat(targetOnHand);
  if (isNaN(target)) {
    return { ok: false, message: 'Target quantity must be a number.' };
  }
  if (target < STOCK_ADJUST.minQty || target > STOCK_ADJUST.maxQty) {
    return { ok: false,
             message: 'Target ' + target + ' is outside sanity bounds [' +
                      STOCK_ADJUST.minQty + ', ' + STOCK_ADJUST.maxQty + '] — refusing.' };
  }

  // --- Resolve the Zoho item_id from the mirror ----------------------------
  // The mirror is used ONLY to find the item_id, which never changes. Its
  // quantity is NOT trusted for the arithmetic — see the header note.
  var zMap = buildZohoStockMap();
  if (!zMap || zMap.size === 0) {
    return { ok: false, message: 'Zoho Stock sheet is empty — nothing to resolve against.' };
  }
  var z = zMap.get(sku.toLowerCase());
  if (!z)          return { ok: false, message: 'SKU not found in Zoho Stock sheet: ' + sku };
  if (!z.itemId)   return { ok: false, message: 'No Zoho item_id for SKU ' + sku + " — can't target the write." };

  try {
    console.log('STOCK ADJUST (test) → SKU=' + sku + ' item_id=' + z.itemId +
                ' | target on-hand ' + target + ' (mirror last saw ' + z.onHand + ')');
  } catch (_) {}

  var res = triggerZohoStockAdjust({
    item_id:   z.itemId,
    sku:       sku,
    target:    target,
    max_delta: opts.force ? STOCK_ADJUST.maxQty : STOCK_ADJUST.maxDelta,
    reason:    opts.reason || STOCK_ADJUST.reason
  });

  if (!res || !res.ok) {
    return { ok: false,
             message: 'Adjustment failed: ' + (res && res.message ? res.message : 'unknown'),
             sku: sku, itemId: z.itemId, target: target, detail: res };
  }

  var d = res.data || {};

  // A no-op is a SUCCESS, and saying so plainly matters: "already correct" is
  // the answer the floor wants most of the time, and dressing it as a failure
  // would teach people to re-run it.
  if (d.noop) {
    return { ok: true, noop: true, sku: sku, itemId: z.itemId,
             before: d.before, target: target, delta: 0,
             message: '✓ ' + sku + ' already at ' + d.before + ' — nothing to adjust.' };
  }

  return {
    ok:           true,
    sku:          sku,
    itemId:       z.itemId,
    before:       d.before,
    target:       target,
    delta:        d.delta,
    adjustmentId: d.adjustment_id || '',
    message:      '✅ ' + sku + ': ' + d.before + ' → ' + target +
                  ' (' + (d.delta > 0 ? '+' : '') + d.delta + ')',
    detail:       d
  };
}


/**
 * EDITOR-RUN TEST WRAPPER.
 *
 * Put a JUNK SKU and the quantity you want it set to below, then Run this from
 * the Apps Script editor. Prove it on something disposable first — this is the
 * same staging the price write went through, and that write was verified on a
 * throwaway item before it was ever pointed at real stock.
 *
 * SAFE BY DEFAULT: the placeholder SKU will not resolve, so running it
 * unedited does nothing.
 */
function runSingleStockAdjustTest() {
  var TEST_SKU    = 'PUT-A-JUNK-SKU-HERE';
  var TEST_TARGET = 5;

  var out = pushSingleStockAdjustBySku(TEST_SKU, TEST_TARGET);
  try { console.log(JSON.stringify(out, null, 2)); } catch (_) { console.log(out); }
  return out;
}


// =======================================================================================
// PUBLIC: board console — correct stock from a shelf count
// =======================================================================================

/**
 * Push one SKU's corrected on-hand quantity to Zoho, from the Floor Board.
 *
 * ⚠ WHY THIS TAKES A TARGET THE PICKER CONFIRMED, RATHER THAN DERIVING ONE.
 * The board can suggest a number but it cannot know one for certain. Zoho
 * reduces `available_stock` when a sales order is created but only reduces
 * `stock_on_hand` when the order SHIPS — so at pick time the units already in
 * the picker's hand are still counted in Zoho's on-hand. The board's best
 * guess is therefore (shelf count + qty pulled for this line), and that is
 * wrong the moment another open order's units have also been pulled and not
 * yet shipped.
 *
 * The board shows its arithmetic and lets the person standing at the shelf
 * confirm or correct the final number. They can see what has been pulled;
 * the server cannot. Same decision-support line the rest of this system
 * draws — suggest, show the working, never silently assert.
 *
 * ⚠ THE GATE IS THE PICK ID, NOT A PASSPHRASE. A passphrase on a shared floor
 * tablet is theatre: every picker would know it within a day. What actually
 * protects this is that the write is narrow (one SKU), bounded (the proxy
 * refuses a swing beyond maxDelta), reversible (another adjustment), and
 * ATTRIBUTED — the same accountability chokepoint that already gates printing.
 *
 * @param {string} sku
 * @param {number} target  the on-hand quantity Zoho should end up with
 * @param {string} [orderId] the line this count came from, for the audit trail
 * @returns {Object}
 */
function boardAdjustStock(sku, target, orderId) {
  sku = String(sku || '').trim();
  orderId = String(orderId || '').trim();
  if (!sku) return { ok: false, error: 'No SKU' };

  var picker = '';
  try { picker = getCurrentPicker(); } catch (e) { picker = ''; }
  if (!picker) {
    return { ok: false,
             error: 'Set the Pick ID before adjusting stock — a correction has to be ' +
                    'attributable to someone.' };
  }

  var res = pushSingleStockAdjustBySku(sku, target,
                                       { reason: STOCK_ADJUST.reason + ' · ' + picker });
  if (!res || !res.ok) {
    return { ok: false, error: (res && res.message) || 'Adjustment failed' };
  }

  // Best-effort audit. Zoho keeps its own adjustment record; this is the half
  // that lives where the rest of the order's story lives.
  try {
    logActivity('NOTE', orderId || '(stock)', sku, '', 'board',
                res.noop
                  ? 'Stock confirmed at ' + res.before + ' — no adjustment needed'
                  : 'Zoho stock ' + res.before + ' → ' + res.target +
                    ' (' + (res.delta > 0 ? '+' : '') + res.delta + ')' +
                    (res.adjustmentId ? ' · adj ' + res.adjustmentId : ''));
  } catch (e) { console.log('boardAdjustStock log: ' + e); }

  // The mirror is now a step behind Zoho; drop the cached tick so the board
  // stops showing the old on-hand next to the number it just corrected.
  try { _dashBustTickCache(); } catch (e) {}

  return { ok: true, noop: !!res.noop, sku: sku,
           before: res.before, after: res.target, delta: res.delta || 0,
           picker: picker, message: res.message };
}
