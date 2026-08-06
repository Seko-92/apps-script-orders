// =======================================================================================
// WEBKITS.gs — kit expansion for the hosted page (web-app slice 2)
// =======================================================================================
//
// WHY THIS EXISTS
//   Kit expansion is the ONE workflow with a hard platform wall: the Google
//   Sheets mobile app cannot render an Apps Script modal at all, so on the
//   warehouse tablet the answer today is simply "you can't." Everything else on
//   the web-app wish-list already has a home that works (the sidebar, the
//   consoles, the Telegram commands). This is the piece that doesn't.
//
// TWO ENDPOINTS, DELIBERATELY
//   getKitQueueForWeb()  — everything the page needs to render, in ONE read.
//   commitKitFromWeb()   — the write.
//   A per-kit "detail" call was considered and dropped: the queue is a handful
//   of kits and already carries components + Sales Description, so a second
//   round-trip would buy nothing on a tablet over warehouse wifi.
//
// THE TRUST RULE (mirrors applyZohoPullSelection)
//   The client sends IDENTIFIERS and CHOICES — never data. commitKitFromWeb
//   re-derives the queue item from a fresh scan at commit time and hands that
//   to the existing engine. So a stale page, a tampered payload, or a kit that
//   somebody expanded from the sidebar thirty seconds ago all fail loudly
//   instead of writing something wrong.
//
// WHAT IT REUSES, UNCHANGED
//   previewSelectedKits(1, rows)  — the whole enrichment pipeline (components,
//                                   live location, Zoho-first availability,
//                                   unparsed-line detection). It gained an
//                                   optional `rowsOverride` argument so this
//                                   path can hand it scanned rows instead of a
//                                   sheet selection.
//   _commitOneKitForModal(...)    — the actual insert, including the row-drift
//                                   re-location that makes a selection-less
//                                   caller safe in the first place.
//   Nothing about the expansion rules is re-implemented here.
//
// AUTH lives OUTSIDE this file. n8n verifies the Telegram initData HMAC before
// forwarding; by the time doPost dispatches here the caller is already known.
// =======================================================================================


/**
 * Every kit row currently awaiting a decision, with everything the page needs
 * to render it.
 *
 * "Awaiting a decision" = a kit-registry SKU sitting on a still-open row that
 * is not itself an expansion component and has no components under it yet.
 * Covers BOTH outcomes: a MANUAL kit waiting to be expanded, and a READY kit
 * waiting for the ship-as-box call.
 *
 * @returns {{ok:boolean, kits:Array, count:number, message:string}}
 */
function getKitQueueForWeb() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss && ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, kits: [], count: 0, message: "All Orders sheet not found." };

    var rows = _webFindKitRows(sheet);
    if (!rows.length) {
      return { ok: true, kits: [], count: 0, message: "No kits waiting — nothing to expand." };
    }

    // ×1 on purpose: base quantities. Spares are a per-kit choice made on the
    // page, exactly as in the modal, so nothing is pre-scaled here.
    var preview = previewSelectedKits(1, rows);
    if (!preview.ok) return { ok: false, kits: [], count: 0, message: preview.message };

    var kits = [];
    for (var i = 0; i < preview.kitRows.length; i++) {
      var k = preview.kitRows[i];
      if (!k.plan.components || !k.plan.components.length) continue;   // no actionable choice
      kits.push({
        kitSku:           k.sourceSku,
        salesOrder:       k.sourceSalesOrder,
        rowQty:           k.sourceQty,
        note:             k.sourceNote,
        row:              k.row,
        table:            k.table,
        kitName:          k.plan.kitName,
        kitType:          k.plan.kitType,
        kitLocation:      k.plan.kitLocation,
        kitEngine:        k.plan.kitEngine,
        salesDescription: k.plan.salesDescription || "",
        components:       k.plan.components,
        unparsedLines:    k.plan.unparsedLines || []
      });
    }

    return { ok: true, kits: kits, count: kits.length, message: "" };

  } catch (err) {
    try { console.log("getKitQueueForWeb: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, kits: [], count: 0, message: String(err.message || err) };
  }
}


/**
 * Commit one kit's expansion.
 *
 * Identified by SKU + SALES ORDER rather than a row number, because rows drift
 * (n8n inserts, sorts, other commits). The queue item is re-derived server-side
 * from a fresh scan — the client's copy is never trusted.
 *
 * @param {string}  kitSku       the kit row's SKU
 * @param {string}  salesOrder   its SALES ORDER
 * @param {Array}   excludedSkus component SKUs the user unchecked
 * @param {number}  extras       spare kits for us, on top of the row qty
 * @param {boolean} force        expand a READY kit anyway
 * @param {Object}  alterations  { overrides:{...}, added:[...] } — same shape
 *                               the modal sends
 * @returns {{ok:boolean, message:string, inserted:number, forced:boolean}}
 */
function commitKitFromWeb(kitSku, salesOrder, excludedSkus, extras, force, alterations) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, message: "Sheet is busy — try again in a moment.", inserted: 0 }; }

  try {
    kitSku     = String(kitSku || "").trim();
    salesOrder = String(salesOrder || "").trim();

    // Only the SKU is required. A BLANK sales order is legitimate — a
    // hand-typed kit row that has not been given one yet — and the engine
    // handles it: _findKitRowBySkuAndSo compares trimmed strings, so blank
    // matches blank, and when several rows tie it disambiguates by proximity
    // to the hint row we pass below. The desktop modal has never required an
    // SO either, and a rule that exists on only one surface is a defect.
    // (An earlier version of this guard demanded both and blocked expansion
    // of any SO-less row — 2026-08-06.)
    if (!kitSku) {
      return { ok: false, message: "Missing kit SKU.", inserted: 0 };
    }

    // Re-derive from a FRESH scan. If it isn't in the queue any more — expanded
    // by someone else, row deleted, status gone terminal — refuse loudly rather
    // than acting on what the page happened to be showing.
    var queue = getKitQueueForWeb();
    if (!queue.ok) return { ok: false, message: queue.message, inserted: 0 };

    var item = null;
    for (var i = 0; i < queue.kits.length; i++) {
      if (String(queue.kits[i].kitSku) === kitSku &&
          String(queue.kits[i].salesOrder) === salesOrder) { item = queue.kits[i]; break; }
    }
    if (!item) {
      return { ok: false, inserted: 0,
               message: "That kit is no longer waiting — it may have been expanded already, " +
                        "or its row changed. Refresh and check." };
    }

    // ⚠ The engine's 3rd argument is EXTRAS (spares for us), NOT total kits.
    // It computes total = rowQty + extras itself — the additive semantic settled
    // on 2026-05-20. Passing a total here would double-ship: a qty-6 row would
    // build 12. Pass the spare count and nothing else.
    var spares = parseInt(extras, 10);
    if (isNaN(spares) || spares < 0) spares = 0;
    if (spares > 99) spares = 99;

    var queueItem = {
      kitSku:           item.kitSku,
      sourceSalesOrder: item.salesOrder,
      sourceQty:        item.rowQty,
      sourceNote:       item.note,
      originalRow:      item.row,
      table:            item.table,
      userMultiplier:   spares,
      kitName:          item.kitName,
      kitType:          item.kitType,
      kitLocation:      item.kitLocation,
      kitEngine:        item.kitEngine,
      salesDescription: item.salesDescription,
      components:       item.components,
      unparsedLines:    item.unparsedLines
    };

    var res = _commitOneKitForModal(
      queueItem,
      excludedSkus || [],
      spares,
      force === true,
      alterations || {}
    );

    // The engine reports the count as `componentsAdded` — NOT `inserted`.
    // Reading the wrong key here would silently report 0 rows on every
    // successful commit.
    return {
      ok:          !!(res && res.ok),
      message:     (res && res.reason) || "",
      inserted:    (res && res.componentsAdded) || 0,
      forced:      !!(res && res.forced),
      excluded:    (res && res.excludedSkus) || [],
      extras:      (res && res.extras) || 0,
      totalKits:   (res && res.totalKits) || 0,
      swapped:     (res && res.swapped) || 0,
      qtyChanged:  (res && res.qtyChanged) || 0,
      addedCustom: (res && res.addedCustom) || 0,
      kitSku:      kitSku,
      salesOrder:  salesOrder
    };

  } catch (err) {
    try { console.log("commitKitFromWeb: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, message: String(err.message || err), inserted: 0 };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/**
 * Row numbers of every kit still awaiting a decision.
 *
 * Same test the straggler watchdog uses, so the two surfaces can never disagree
 * about what's outstanding: a registered kit SKU, on a still-open row, that is
 * not itself an expansion component, with no matching "↳ from KIT-<sku>" tag on
 * the row below.
 *
 * @returns {Array<number>} 1-based sheet rows
 */
function _webFindKitRows(sheet) {
  var out = [];
  var lastRow = sheet.getLastRow();
  if (lastRow < Schema.dataStartRow) return out;

  var kitMap;
  try { kitMap = buildKitMap(); }
  catch (e) { console.log("_webFindKitRows: kit map failed — " + e); return out; }
  if (!kitMap || !kitMap.size) return out;

  var data = sheet.getRange(Schema.dataStartRow, 1,
                            lastRow - Schema.dataStartRow + 1, Schema.dataWidth).getValues();

  var SKU_I  = Schema.idx("SKU");
  var NOTE_I = Schema.idx("NOTE");
  var ST_I   = Schema.idx("STATUS");

  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][SKU_I] || "").trim();
    if (!sku || sku.toUpperCase() === Schema.boundaryMarker) continue;
    if (!kitMap.get(sku)) continue;

    var status = String(data[i][ST_I] || "").trim().toUpperCase();
    if (status !== Schema.status.PENDING && status !== Schema.status.PREPARING) continue;

    var note = String(data[i][NOTE_I] || "");
    if (note.indexOf("↳ from KIT-") === 0) continue;          // it IS a component

    var next = data[i + 1];
    if (next && String(next[NOTE_I] || "").indexOf("↳ from KIT-" + sku) === 0) continue;  // expanded

    out.push(Schema.dataStartRow + i);
  }
  return out;
}
