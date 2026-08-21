/**
 * Holds.js — THE HOLD, END TO END.
 * ---------------------------------------------------------------------------
 * A hold is "picked is fine, do NOT hand this to the carrier." It is written by
 * whoever talks to the buyer, and it has to reach whoever is standing next to
 * the box — who is in another city, on another surface, possibly mid-pick.
 *
 * ⚠ WHY THIS FILE EXISTS (the 2026-08-21 incident). A buyer asked to change the
 * shipping service. The announcement went to a WhatsApp group. The pickers were
 * busy, missed it, bought the label, and the row went SHIPPED — at which point
 * `_dashOpenOrders` stops emitting it (it filters to PENDING/PREPARING), so the
 * board had nowhere to show a hold even if one had been written. The box sat in
 * the building for hours with no surface anywhere in the system saying "stop".
 *
 * ⚠⚠ AND THE TIMING IS NOT A COINCIDENCE. Buying the label is what sends the
 * buyer their shipping notification, which is exactly what prompts "wait, can
 * you upgrade this / hold it". So these requests arrive PREFERENTIALLY inside
 * the one window the board was blind to. No amount of better messaging closes
 * that; the window itself had to be opened.
 *
 * ── THE GRAMMAR ────────────────────────────────────────────────────────────
 * Everything reads ONE cell — the row's NOTE. There is no parallel store, no
 * hidden register, no second source that can drift:
 *
 *   "hold this and change from x to z"                      → held, unseen
 *   "hold this and change from x to z · ✓ SEEN 2:32 PM by …" → held, acknowledged
 *   (the word HOLD removed)                                 → lifted
 *
 * That means acknowledging from ANY door — the tablet, the sidebar, Telegram —
 * is the same write, and the siren stops everywhere within one poll. It also
 * means LIFTING needs no mechanism at all: you delete the word while you are
 * already editing the note you were reading. The clearing path is the one that
 * usually rots, and here there isn't one to rot.
 *
 * ⚠ HOLD DETECTION MUST STAY IDENTICAL TO THE BOARD'S. FloorBoard.html has its
 * own `noteHasHold()` — whole word, anywhere, any case — because the board
 * cannot call server code. Two copies of a rule is how `A-9` sorted after
 * `A-50` in three files until 2026-08-05. If you change one, change both, and
 * there is a test asserting they agree.
 *
 * ⚠ THE MATCH IS DELIBERATELY LOOSE — whole word, ANYWHERE in the note, ANY
 * case. A hold is usually APPENDED to a row that already carries a buyer note,
 * so requiring it to lead would miss the common case. A buyer note containing
 * "hold" lighting up is loud and harmless; a real hold going quiet is the
 * failure that has now happened twice. FAIL TOWARD SHOWING.
 *
 * ⚠ AND THE SYSTEM CANNOT ENFORCE ANY OF IT. The label is bought in eBay's
 * Seller Hub, outside this system entirely. Everything here makes a hold
 * impossible to MISS; nothing here makes it impossible to IGNORE. That is why
 * the physical half — a red tag on the box, a taped-off hold area — is not a
 * nice-to-have. It is the only layer that actually stops a box, and it is the
 * only layer that still works when the tablet is dead.
 *
 * ── WHAT IS PUBLIC ─────────────────────────────────────────────────────────
 *   holdNoteHasHold(note)      → bool     (pure — mirrors the board's rule)
 *   holdNoteHasAck(note)       → bool     (pure)
 *   holdNoteAckText(note)      → string   (pure — "2:32 PM by Yassin · 1")
 *   holdBuildAckTag(picker)    → string   (pure-ish — reads the clock)
 *   holdAppendAck(note, tag)   → string   (pure)
 *   boardAckHold(orderId)      → {ok,…}   THE WRITE. Every door lands here.
 *   getHeldOrderCount()        → int      cheap-ish, for the sidebar badge
 *   acknowledgeSelectedHold()  → {ok,…}   sidebar entry (sheet selection)
 *   holdRecordLive(held)       → void     called at publish time
 *   checkHoldEscalation()      → string   called once a minute
 *   resetHoldEscalation()      → string   manual re-seed
 */

var HOLDS = {
  // How long an unacknowledged hold may sit before the system stops shouting at
  // the floor and goes to find the person who wrote it.
  //
  // ⚠ THE ASYMMETRY PICKS THE NUMBER, not a guess. Escalate too early and it
  // costs one Telegram message. Escalate too late and it costs a voided label,
  // a re-ship and a customer. The real deadline is not a clock at all — it is
  // carrier pickup, which we cannot see — so this is a floor, not an estimate.
  //
  // ⚠ RAISED 5 → 15 ON 2026-08-21, and the reason matters more than the number.
  // Five minutes fires while a picker is still WALKING. They may be six aisles
  // away with the siren going, on their way back to the board — and escalating
  // on someone who is actively responding is how an alert becomes noise, which
  // is the precise failure that made the WhatsApp group stop working. Fifteen
  // means "nobody is coming", not "nobody has arrived yet".
  //
  // This one is for a PREPARING hold: pulled from the shelf, label NOT bought.
  escalateAfterMin: 15,

  // ⭐⭐ SHIPPED GETS ITS OWN, SHORTER WINDOW — the user's call, and the cost of
  // being late is genuinely different, not just felt differently:
  //
  //   PREPARING  · nothing irreversible has happened. The packing bench and the
  //                label purchase are both still ahead, so a hold caught late
  //                here costs approximately nothing.
  //   SHIPPED    · the label is bought. Money is already committed, and the box
  //                is waiting on a carrier who can arrive at any minute — an
  //                EXTERNAL deadline this system cannot see. Late here costs the
  //                label, a re-ship and a customer.
  //
  // ⚠ The picker's walk does not get shorter because the order shipped, so this
  // is deliberately not tiny: eight minutes is still ~24 sirens and a real
  // chance for someone to reach the board. It trades a little of the walking
  // allowance for the case where being late actually costs something.
  escalateShippedAfterMin: 8,

  // Script Property holding the last published view of every live hold, written
  // by holdRecordLive() at publish time. The per-minute check reads THIS, never
  // the sheet — see checkHoldEscalation for why that matters.
  liveKey: "HOLD_LIVE",

  // Script Property holding {orderId: {first, alerted}} — the alert-once-per-
  // crossing memory, same shape and same reasoning as the straggler watchdog.
  stateKey: "HOLD_ESCALATION",

  // A key whose hold has vanished is dropped immediately; this only bounds a
  // store that somehow stops being cleaned.
  pruneDays: 14,

  // A held order is 1-3 lines in practice; the cap is a guard against a
  // fourteen-line kit filling a takeover card nobody can read at a glance.
  maxItemsShown: 6,

  ackMark: "✓ SEEN"
};


/* ═══════════════════════ THE NOTE GRAMMAR (pure) ═══════════════════════════ */

/**
 * ⚠ MUST STAY IDENTICAL TO FloorBoard.html's noteHasHold(). See the file header.
 * @param {string} note
 * @returns {boolean}
 */
function holdNoteHasHold(note) {
  return /\bHOLD\b/i.test(String(note == null ? "" : note));
}

/**
 * Has someone already said they have seen it?
 * @param {string} note
 * @returns {boolean}
 */
function holdNoteHasAck(note) {
  return /✓\s*SEEN\b/i.test(String(note == null ? "" : note));
}

/**
 * The human-readable half of the ack, for the board strip: "2:32 PM by Yassin · 1".
 * Empty string when there is no ack.
 * @param {string} note
 * @returns {string}
 */
function holdNoteAckText(note) {
  var m = String(note == null ? "" : note).match(/✓\s*SEEN\s+(.+)$/i);
  return m ? String(m[1]).trim() : "";
}

/**
 * Has this hold already been escalated to the person who wrote it?
 * @param {string} note
 * @returns {boolean}
 */
function holdNoteHasEscalated(note) {
  return /⚠\s*ESCALATED\b/i.test(String(note == null ? "" : note));
}

/**
 * When it was escalated — "2:37 PM". Empty string when it has not been.
 * @param {string} note
 * @returns {string}
 */
function holdNoteEscalatedText(note) {
  var m = String(note == null ? "" : note).match(/⚠\s*ESCALATED\s+([0-9]{1,2}:[0-9]{2}\s*[AP]M)/i);
  return m ? String(m[1]).trim() : "";
}

/**
 * The tag appended to the note when someone acknowledges.
 *
 * ⚠ THE CLOCK IS PINNED TO HOUSTON. The script timezone is Asia/Amman (see the
 * 2026-08-19 trigger-hour trap) — eight hours out in summer — and this string is
 * read by a person standing on the Houston floor. Formatting it in the script's
 * own zone would stamp a time that never happened for them.
 *
 * ⚠ AN UNSET PICKER IS RECORDED, NOT REFUSED. This is the one write on the board
 * that must never argue: an urgent alert is the worst possible moment to force
 * someone through a dropdown, and a hold seen by "somebody" beats a hold nobody
 * could acknowledge. Safety beats bookkeeping in that thirty seconds.
 *
 * @param {string} picker  cleaned Pick ID, or "" when unset
 * @returns {string}
 */
function holdBuildAckTag(picker) {
  var t = Utilities.formatDate(new Date(), "America/Chicago", "h:mm a");
  var who = String(picker || "").trim();
  return HOLDS.ackMark + " " + t + " by " + (who || "the floor");
}

/**
 * Append the tag to a note, using the same " · " segment separator the rest of
 * the system writes. Idempotent: a note that already carries an ack is returned
 * untouched, so two people tapping at once cannot stack two tags.
 *
 * @param {string} note
 * @param {string} tag
 * @returns {string}
 */
function holdAppendAck(note, tag) {
  var raw = String(note == null ? "" : note).trim();
  if (holdNoteHasAck(raw)) return raw;
  return raw ? (raw + " · " + tag) : tag;
}


/* ═══════════════════════ SCANNING (pure) ═══════════════════════════════════ */

/**
 * Build the board's held-order list from raw All Orders row values.
 *
 * PURE on purpose — it takes the values array the tick builder has already read
 * and returns plain objects, so the whole "which orders are held, and is the
 * shipped one included" decision is testable in Node against real row shapes
 * without a spreadsheet.
 *
 * ⚠ ONE ENTRY PER ORDER, NOT PER ROW. A hold is an order-level fact — you
 * cannot part-ship an order, so holding any line holds all of it. Repeating the
 * same sentence on nine lines of the alert strip would push everything else off
 * the screen to say one thing. The pick list below still marks every line
 * individually, because THAT surface's job is to meet the picker at each shelf.
 *
 * ⚠⚠ SHIPPED ROWS COUNT — THIS IS THE WHOLE POINT OF THE FILE. n8n deletes
 * shipped rows at about 1 AM Houston, so a SHIPPED row still present on the
 * sheet is, by construction, a box from today that has not yet been cleaned up
 * — which is exactly the window where it is probably still in the building.
 * That means presence on the sheet IS the window, and no Activity Log join is
 * needed to work out which shipped orders still matter.
 *
 * @param {Array<Array>} data   rows from Schema.dataStartRow, full dataWidth
 * @returns {Array<Object>}
 */
function holdScanRows(data) {
  data = data || [];
  var byOrder = {};
  var order = [];
  var inDirect = false;

  for (var i = 0; i < data.length; i++) {
    var sku = String(data[i][Schema.idx("SKU")] || "").trim();
    if (sku.toUpperCase() === Schema.boundaryMarker) { inDirect = true; continue; }
    if (!sku) continue;

    var status = String(data[i][Schema.idx("STATUS")] || "").trim().toUpperCase();
    // CANCELED is out: that line is no longer going anywhere, so a hold on it
    // has nothing left to stop. PENDING / PREPARING / SHIPPED are all "still in
    // the building" as far as this file is concerned.
    if (status === Schema.status.CANCELED) continue;

    var note = String(data[i][Schema.idx("NOTE")] || "").trim();
    if (!holdNoteHasHold(note)) continue;

    var oid = String(data[i][Schema.idx("SALES_ORDER")] || "").trim();
    if (!oid) continue;

    if (!byOrder[oid]) {
      byOrder[oid] = {
        orderId: oid,
        channel: inDirect ? "DIRECT" : "EBAY",
        note:    note,
        acked:   holdNoteHasAck(note),
        ackText: holdNoteAckText(note),
        // TRUE once EVERY live line of the order has shipped. A part-shipped
        // order is still pickable work, so it keeps the calmer treatment.
        shipped: true,
        // ⚠ URGENT = SOMEBODY HAS ALREADY ACTED ON THIS BOX. Any line PREPARING
        // (pulled from the shelf) or SHIPPED (label bought) means the hold has
        // arrived late and there is something to undo. THAT is what earns the
        // siren and the full-screen takeover.
        //
        // A hold on a still-PENDING order deliberately does NOT: nothing has
        // happened yet, the existing amber chip already meets the picker at
        // every shelf, and a board that screams at the calm case is a board
        // that gets muted before it ever sees the dangerous one.
        urgent:  false,
        // ⚠ THE ESCALATION HAS TO LEAVE A MARK ON THE SHEET, not only in a chat.
        // Escalating means the system gave up on the floor and went to find the
        // person who wrote the hold — and if the only trace of that is a Telegram
        // message, it scrolls away, which is EXACTLY the failure this whole
        // feature exists to fix. A cell that has been red for 40 minutes and one
        // that has been red for 30 seconds otherwise look identical.
        escalated: false,
        escText:   "",
        /* ⭐ WHAT IS ACTUALLY IN THE BOX (2026-08-21, the user's question).
           An order id identifies a box to the SYSTEM. It does not identify one
           to a person standing in an outbound area on a busy afternoon with
           fifteen boxes in front of them — and a hold that makes you walk to a
           computer and look up what you are holding has spent most of the time
           it just saved.
           ⚠ FREE: these are columns this scan has already read. Deliberately NOT
           the item title — that lives in Master Inventory, and the cheapest read
           of it is ~1.9s (the round trip dominates, so narrowing does not help —
           2026-05-22). SKU + qty + shelf is what identifies a part everywhere
           else on this board, so it is what identifies one here. */
        items:   [],
        lines:   0
      };
      order.push(oid);
    }
    var h = byOrder[oid];
    h.lines++;
    if (h.items.length < HOLDS.maxItemsShown) {
      h.items.push({
        sku: sku,
        qty: data[i][Schema.idx("QTY")],
        loc: String(data[i][Schema.idx("LOCATION")] || "").trim()
      });
    }
    if (status !== Schema.status.SHIPPED) h.shipped = false;
    if (status === Schema.status.PREPARING || status === Schema.status.SHIPPED) h.urgent = true;
    // ⚠ AN ACK ANYWHERE ON THE ORDER COUNTS. The ack is written to every held
    // row, but a row inserted afterwards (a Zoho line-item add) would not carry
    // it — and re-arming the siren because a NEW row appeared on an order
    // somebody already answered for is exactly the nagging that gets alarms
    // muted. The longest note wins for display so the strip shows the fullest
    // version of what was written.
    if (holdNoteHasAck(note)) {
      h.acked = true;
      if (!h.ackText) h.ackText = holdNoteAckText(note);
    }
    if (holdNoteHasEscalated(note)) {
      h.escalated = true;
      if (!h.escText) h.escText = holdNoteEscalatedText(note);
    }
    if (note.length > String(h.note).length) h.note = note;
  }

  var out = [];
  for (var k = 0; k < order.length; k++) out.push(byOrder[order[k]]);
  // UNACKNOWLEDGED FIRST — the strip is read top-down and the unanswered ones
  // are the only actionable rows in it.
  out.sort(function (a, b) {
    if (a.acked !== b.acked) return a.acked ? 1 : -1;
    return String(a.orderId).localeCompare(String(b.orderId));
  });
  return out;
}


/* ═══════════════════════ THE WRITE ═════════════════════════════════════════ */

/**
 * Acknowledge a hold. Every door — the tablet's takeover, the sidebar button,
 * Telegram's /ack — lands here, so there is exactly one implementation of what
 * "seen" means.
 *
 * ⚠ WRITES TO EVERY HELD ROW OF THE ORDER, not just the first. The hold may
 * have been typed on any line, and the sheet's colour is per cell — leaving the
 * others red would make the table disagree with itself.
 *
 * ⚠ ACKNOWLEDGING IS NOT RESOLVING. This records that a human has seen it. The
 * strip stays up and the note keeps the word HOLD until somebody deletes it. If
 * the tap made everything vanish we would have replaced a missed message with a
 * dismissed one, and the box would still be sitting there needing its label
 * voided.
 *
 * @param {string} orderId
 * @returns {{ok:boolean, rows:number, tag:string, error:string=}}
 */
function boardAckHold(orderId) {
  orderId = String(orderId || "").trim();
  if (!orderId) return { ok: false, error: "Need an order id" };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) return { ok: false, error: "Sheet busy — try again" };

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { ok: false, error: "No sheet" };
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return { ok: false, error: "No rows" };

    var rows = _resolveStatusTargetRows(sheet, orderId, lastRow);
    if (!rows.length) return { ok: false, error: "Order not found — it may have been cleaned up" };

    // Deliberately NOT gated on the Pick ID — see holdBuildAckTag.
    var picker = "";
    try { picker = getCurrentPicker() || ""; } catch (e) {}
    var tag = holdBuildAckTag(picker);

    var touched = 0;
    for (var i = 0; i < rows.length; i++) {
      var cell = sheet.getRange(rows[i], Schema.cols.NOTE);
      var note = String(cell.getValue() || "").trim();
      if (!holdNoteHasHold(note) || holdNoteHasAck(note)) continue;
      cell.setValue(holdAppendAck(note, tag));
      touched++;
    }
    SpreadsheetApp.flush();

    if (!touched) {
      return { ok: true, rows: 0, tag: tag, already: true };
    }

    try {
      // ⚠ POSITIONAL: (event, orderId, sku, qty, source, detail, picker, note)
      // picker left undefined — logActivity reads G2 itself for warehouse-side
      // sources, which is the one place that value is authoritative.
      logActivity("NOTE", orderId, "", "", "board", "Hold acknowledged · " + tag);
    } catch (e) { console.log("boardAckHold log: " + e); }

    _dashBustTickCache();
    return { ok: true, rows: touched, tag: tag };
  } catch (e) {
    console.error("boardAckHold: " + e);
    return { ok: false, error: String(e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Stamp "⚠ ESCALATED h:mm AM" onto every held row of an order.
 *
 * ⚠ BEST-EFFORT, AND CALLED ONLY AFTER THE MESSAGE HAS ACTUALLY GONE OUT. The
 * Telegram send is the thing that must not be lost; this is the record of it. If
 * the lock is busy or a write fails, the escalation still happened and still
 * alerted — the sheet just does not say so, which is a smaller loss than a
 * duplicate alarm or a blocked trigger.
 *
 * ⚠ IT TAKES THE LOCK BRIEFLY AND GIVES UP FAST. This runs inside
 * runPublishTick, a per-minute trigger; waiting on a picker's ✓ Pick to finish
 * would be the wrong trade for a decoration.
 *
 * @param {string} orderId
 * @returns {number} rows stamped
 */
function holdStampEscalated(orderId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) return 0;
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return 0;

    var rows = _resolveStatusTargetRows(sheet, String(orderId || "").trim(), lastRow);
    var tag = "⚠ ESCALATED " + Utilities.formatDate(new Date(), "America/Chicago", "h:mm a");
    var n = 0;
    for (var i = 0; i < rows.length; i++) {
      var cell = sheet.getRange(rows[i], Schema.cols.NOTE);
      var note = String(cell.getValue() || "").trim();
      // Only a live, unanswered hold gets stamped — and never twice.
      if (!holdNoteHasHold(note) || holdNoteHasAck(note) || holdNoteHasEscalated(note)) continue;
      cell.setValue(note + " · " + tag);
      n++;
    }
    if (n) {
      SpreadsheetApp.flush();
      try { _dashBustTickCache(); } catch (e) {}
    }
    return n;
  } catch (e) {
    console.log("holdStampEscalated: " + e);
    return 0;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/* ═══════════════════════ ESCALATION ════════════════════════════════════════ */

/**
 * Record the live hold picture, called from publishBoardTick with the tick's own
 * `held` array.
 *
 * ⚠ THIS EXISTS SO THE PER-MINUTE CHECK NEVER READS THE SHEET. runPublishTick
 * fires 1,440 times a day and skips most of them; hanging a sheet read on it
 * would be the same mistake as dropping maxAgeMinutes to 1 — a full read every
 * minute to answer a question that changes three to five times a WEEK. A hold
 * can only appear or clear via a NOTE edit, every NOTE edit busts the cache and
 * marks the tick dirty, and a dirty tick republishes within a minute. So this
 * property is refreshed exactly when the answer can have changed, and the check
 * below costs one property read.
 *
 * @param {Array<Object>} held  from the tick
 */
function holdRecordLive(held) {
  try {
    var lean = (held || []).map(function (h) {
      return { o: h.orderId, a: !!h.acked, s: !!h.shipped, u: !!h.urgent,
               n: String(h.note || "").slice(0, 120) };
    });
    PropertiesService.getScriptProperties()
      .setProperty(HOLDS.liveKey, JSON.stringify(lean));
  } catch (e) { console.log("holdRecordLive: " + e); }
}

/**
 * The empty-room backstop. Runs once a minute from runPublishTick.
 *
 * ⚠⚠ IT HAS TO BE SERVER-SIDE, AND THAT IS THE WHOLE DESIGN. If the board were
 * counting the minutes, then the exact case escalation exists for — the room is
 * empty, the tablet is asleep, nobody is there to hear the siren — is the case
 * where nothing would ever fire. A clock that only runs when someone is
 * watching is not a backstop.
 *
 * ⚠ ALERT ONCE PER HOLD, on first crossing. Stable keys in a Script Property,
 * same shape and same reasoning as the straggler watchdog: a repeating alarm
 * about something already reported is how alarms get muted. It re-arms only
 * when the hold clears and a new one appears.
 *
 * ⚠ THE KEY IS RECORDED ONLY AFTER A SUCCESSFUL SEND, so the failure mode is a
 * possible duplicate rather than a silent miss — the right way round for the
 * message that exists to stop a wrong shipment.
 *
 * ⚠ IT FIRES OUTSIDE WORKING HOURS TOO, deliberately. Off-hours is precisely
 * when nobody can acknowledge, so it is when the escalation is most needed —
 * and the person it reaches is remote anyway. If a 1 AM Riyadh ping ever proves
 * worse than a wrong shipment, mute that Telegram chat overnight rather than
 * gate this, because the gate would silence the case it was built for.
 *
 * @returns {string} a one-line summary, for the editor log
 */
function checkHoldEscalation() {
  try {
    var props = PropertiesService.getScriptProperties();
    var liveRaw = props.getProperty(HOLDS.liveKey);
    if (!liveRaw) return "no live-hold snapshot yet";

    var live;
    try { live = JSON.parse(liveRaw) || []; } catch (e) { live = []; }

    var state = {};
    try { state = JSON.parse(props.getProperty(HOLDS.stateKey) || "{}") || {}; }
    catch (e) { state = {}; }

    var now     = Date.now();
    var changed = false;
    var present = {};
    var fired   = 0;

    for (var i = 0; i < live.length; i++) {
      var h = live[i];
      var oid = String(h.o || "").trim();
      if (!oid || h.a) continue;          // acknowledged → nothing to chase

      /* ⚠⚠ URGENT ONLY — PREPARING or SHIPPED. Found 2026-08-21 by the user
         asking whether that was already true. It was NOT, and the gap was worse
         than an inconsistency: the takeover and the siren fire only for urgent
         holds, so a still-PENDING hold produces no prompt to acknowledge at all
         — which means it could never be answered, which means it would
         ALWAYS escalate at the threshold. Every calm hold, guaranteed noise, on
         the channel that exists to be believed.

         A PENDING hold has nothing to undo: no label bought, no box packed, and
         the picker meets it at every shelf they walk for that order. It is
         carried by the strip and the row chips, which is the right weight for it.

         ⭐ AND THE CLOCK NOW STARTS AT THE RIGHT MOMENT. Because no state is
         recorded while it is calm, `first` is stamped the moment the order
         becomes PREPARING — so the fifteen minutes are counted from when the
         hold acquired a deadline, not from when it was written. */
      if (!h.u) continue;

      present[oid] = true;

      if (!state[oid]) { state[oid] = { first: now, alerted: 0, gate: "" }; changed = true; continue; }

      /* ⭐⭐ TWO GATES, NOT ONE (2026-08-21, the user's call).
         The rule everywhere else in this system is ALERT ONCE PER CROSSING — and
         the point is the word CROSSING, not the word once. PREPARING → SHIPPED
         on a hold nobody has answered for IS a new crossing: the thing you were
         warned about got materially worse, because a label now exists and money
         is committed. A second message there is not a repeat, it is a different
         fact.
         ⚠ AND IT CANNOT BECOME NOISE, because it only ever fires when somebody
         SHIPPED an order that had an unacknowledged hold on it — which is the
         exact disaster this whole feature was built to prevent. If that message
         is arriving often, the message is not the problem. */
      var gate = h.s ? "ship" : "prep";
      if (state[oid].gate === gate) continue;               // already told, same state
      if (state[oid].gate === "ship") continue;             // ship is terminal — nothing worse to report

      // ⚠ THE WINDOW IS CHOSEN PER ORDER, NOT PER PASS. Reading `h.s` here
      // rather than stamping a threshold at first sight is what lets a hold that
      // has been waiting 10 minutes and then ships escalate on the NEXT pass.
      var thresh = (h.s ? HOLDS.escalateShippedAfterMin : HOLDS.escalateAfterMin) * 60000;
      if (now - state[oid].first < thresh) continue;

      var mins  = Math.round((now - state[oid].first) / 60000);
      var again = (gate === "ship" && state[oid].gate === "prep");
      var text = again
        ? ("🚨 THE HELD ORDER JUST SHIPPED\n\n" +
           oid + "  (label bought — still not acknowledged)\n" +
           String(h.n || "").trim() + "\n\n" +
           "You were told about this " + mins + " min ago and nobody answered.\n" +
           "A label now exists, so the box is waiting on the carrier — this is the " +
           "last point where voiding it is still free.")
        : ("⏸ HOLD NOT ACKNOWLEDGED\n\n" +
           oid + (h.s ? "  (label already bought)" : "  (picked, label not bought yet)") + "\n" +
           String(h.n || "").trim() + "\n\n" +
           "Nobody on the floor has acknowledged this for " + mins + " min" +
           " (limit " + (h.s ? HOLDS.escalateShippedAfterMin : HOLDS.escalateAfterMin) + ").\n" +
           (h.s ? "The box may still be in the building — worth a call before the carrier takes it."
                : "Worth a call before the label is bought."));

      var sent = false;
      try { sent = _tgSend(TELEGRAM_ADMIN_CHAT_ID, text); }
      catch (e) { console.log("checkHoldEscalation send: " + e); }

      if (sent) {
        state[oid].alerted = now; state[oid].gate = gate; changed = true; fired++;
        // The sheet says it too — see holdStampEscalated. Best-effort by
        // contract: the message is already out and the key is already set, so a
        // failure here can never produce a second alarm.
        // ⚠ holdStampEscalated refuses to stamp twice, so the second gate adds
        // no second mark — the cell already says ESCALATED, and repeating it
        // would lengthen the note without telling anyone anything new.
        try { holdStampEscalated(oid); } catch (e) { console.log("stamp: " + e); }
      }
    }

    // A hold that has been acknowledged or lifted drops out immediately, so the
    // next one on the same order arms cleanly.
    var cutoff = now - HOLDS.pruneDays * 86400000;
    for (var k in state) {
      if (!Object.prototype.hasOwnProperty.call(state, k)) continue;
      if (!present[k] || state[k].first < cutoff) { delete state[k]; changed = true; }
    }

    if (changed) props.setProperty(HOLDS.stateKey, JSON.stringify(state));
    return fired ? ("escalated " + fired) : ("watching " + Object.keys(present).length);
  } catch (err) {
    console.log("checkHoldEscalation: " + err);
    return "error: " + String(err.message || err);
  }
}

/** Forget every escalation key — the next pass re-arms from now. */
function resetHoldEscalation() {
  PropertiesService.getScriptProperties().deleteProperty(HOLDS.stateKey);
  return "✅ Hold escalation memory cleared — re-arms on the next pass.";
}


/* ═══════════════════════ SIDEBAR / SHEET DOOR ═════════════════════════════ */

/**
 * How many orders are held, and how many of those nobody has answered for.
 * Reads six columns, not the full width — the boundary marker, the order id,
 * the status and the note are all this needs.
 *
 * @returns {{total:number, unacked:number}}
 */
function getHeldOrderCount() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return { total: 0, unacked: 0 };
    var lastRow = sheet.getLastRow();
    if (lastRow < Schema.dataStartRow) return { total: 0, unacked: 0 };

    var n = lastRow - Schema.dataStartRow + 1;
    var data = sheet.getRange(Schema.dataStartRow, 1, n, Schema.cols.STATUS).getValues();
    var held = holdScanRows(data);
    var un = 0;
    for (var i = 0; i < held.length; i++) if (!held[i].acked) un++;
    return { total: held.length, unacked: un };
  } catch (e) {
    console.log("getHeldOrderCount: " + e);
    return { total: 0, unacked: 0 };
  }
}

/**
 * Sidebar entry: acknowledge the hold on whatever rows are selected.
 *
 * ⚠ THIS IS THE DOOR FOR THE DAY THE TABLET IS NOT USED. It exists because a
 * capability reachable from only one surface is a capability the floor loses
 * the moment that surface is flat, forgotten or being charged — the same
 * reasoning that turned the board's picker chip from a label into a button.
 *
 * ⚠ IT DELIBERATELY DOES NOT LET ANYONE TYPE THE ACK BY HAND. Hand-typed text
 * would carry no real timestamp and no name, and "seen at 2:32 by Yassin" IS
 * the value — a red cell that merely turned amber tells you nothing about who
 * has it.
 *
 * @returns {{ok:boolean, orders:number, rows:number, message:string}}
 */
function acknowledgeSelectedHold() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getActiveSheet();
    if (!sheet || sheet.getName() !== MAIN_SHEET_NAME) {
      return { ok: false, orders: 0, rows: 0,
               message: "Open the All Orders sheet and select a held row first." };
    }

    var ranges = ss.getActiveRangeList();
    if (!ranges) return { ok: false, orders: 0, rows: 0, message: "Select a row first." };

    var list = ranges.getRanges();
    var wanted = {}, order = [];
    for (var r = 0; r < list.length; r++) {
      var start = list[r].getRow(), count = list[r].getNumRows();
      for (var i = 0; i < count; i++) {
        var row = start + i;
        if (row < Schema.dataStartRow) continue;
        var oid  = String(sheet.getRange(row, Schema.cols.SALES_ORDER).getValue() || "").trim();
        var note = String(sheet.getRange(row, Schema.cols.NOTE).getValue() || "").trim();
        if (!oid || !holdNoteHasHold(note) || wanted[oid]) continue;
        wanted[oid] = true; order.push(oid);
      }
    }

    if (!order.length) {
      return { ok: false, orders: 0, rows: 0,
               message: "No held rows in the selection — a hold is a NOTE containing the word HOLD." };
    }

    var rows = 0;
    for (var k = 0; k < order.length; k++) {
      var res = boardAckHold(order[k]);
      if (res && res.ok) rows += (res.rows || 0);
    }
    return { ok: true, orders: order.length, rows: rows,
             message: "✅ Acknowledged " + order.length + " hold" +
                      (order.length === 1 ? "" : "s") + " · " + rows + " row" +
                      (rows === 1 ? "" : "s") + " stamped." };
  } catch (e) {
    console.error("acknowledgeSelectedHold: " + e);
    return { ok: false, orders: 0, rows: 0, message: "Failed: " + String(e.message || e) };
  }
}
