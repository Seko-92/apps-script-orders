/**
 * TelegramCommands.js — the Telegram command router (shipped 2026-08-05)
 * ============================================================================
 *
 * PHASE 1 of "THE TELEGRAM LAYER" (see the roadmap section in CLAUDE.md).
 *
 * WHAT THIS IS
 * ------------
 * Until now the bot only listened for BUTTON TAPS (`callback_query` → the
 * PREP/PEND status flips). It never listened for TYPED MESSAGES. This file adds
 * that half: a text command arrives, we parse it, run it, and reply in the chat.
 *
 * It is deliberately built as a ROUTER rather than a feature, because the whole
 * point is that every command after the first is a small job:
 *
 *     TG_ROUTES["/part"] = { help: "...", run: function (args) { ... } }
 *
 * Add a route, done. No n8n change, no new deployment surface.
 *
 * THE FLOW
 * --------
 *   Telegram  →  n8n webhook (telegram-callback, ALREADY EXISTS)
 *             →  IF callback_query?  → the EXISTING chain, UNTOUCHED
 *             →  IF message + text?  → thin forward to Apps Script /exec
 *             →  doPost action 'telegramCommand'
 *             →  handleTelegramCommand(update)
 *             →  replies to Telegram directly via UrlFetchApp
 *
 * WHY THE LOGIC LIVES HERE AND NOT IN n8n
 * ---------------------------------------
 * Apps Script is version-controlled and diffable; the n8n instance lives on the
 * VPS behind a VPN and is edited by hand in a browser. Keeping n8n as a THIN
 * PROXY (the same role it plays for the Zoho webhooks — see Gotcha #12 and the
 * Zoho Kit Refresh Proxy) means the logic stays reviewable and the wiring stays
 * boring.
 *
 * SECURITY — chat allowlist, no passphrase
 * ----------------------------------------
 * The webhook is public, so the gate is a SERVER-SIDE ALLOWLIST of chat IDs.
 * Telegram tells us who sent every update, which is exactly why this layer was
 * chosen over a phone web page (that would have needed a passphrase on a public
 * /exec URL). Allowed = TELEGRAM_ADMIN_CHAT_ID, plus any extra IDs in the
 * optional Script Property TELEGRAM_COMMAND_CHATS (comma-separated) so the
 * warehouse group can be added WITHOUT a code change or a redeploy.
 *
 * An update from a chat that is not on the list is IGNORED IN SILENCE — never
 * answered with "unauthorized", which would only confirm the bot is live to
 * anyone probing it.
 *
 * REPLIES ARE PLAIN TEXT — no parse_mode. Same robustness call as the weekly
 * digest and _sendKitParseAlert: item names and free text routinely contain
 * characters that break Markdown/HTML parsing, and a digest that fails to send
 * is worse than one without bold.
 *
 * READ-ONLY BY DESIGN IN v1. Every route here only reads. Write commands
 * (/pull, /note, photo upload) are deliberately a later phase — see CLAUDE.md.
 */

var TG_COMMANDS = {
  // Optional Script Property holding extra allowed chat IDs, comma-separated.
  // Lets the warehouse group be added without touching code.
  allowlistPropKey: "TELEGRAM_COMMAND_CHATS",

  maxReplyChars:  3800,   // Telegram's hard cap is 4096 — leave headroom
  rippleLimit:    8,      // kits listed in /part before we summarise the rest
  componentLimit: 10,     // components listed for a kit
  oosLimit:       25,     // OOS items listed before we summarise the rest
  eventLimit:     8,      // timeline events shown by /order

  // Namespace for OUR inline-button callbacks. Must NOT collide with the
  // PREP_/PEND_ prefixes owned by the original button handler — that parser
  // returns skip:true for anything it doesn't recognise, so our taps flow
  // harmlessly past it and down the command branch instead.
  // Telegram caps callback_data at 64 BYTES, so keep payloads short.
  callbackPrefix: "HQ:"
};


// =======================================================================================
// ENTRY POINT
// =======================================================================================

/**
 * Handle one Telegram update — either a typed `/command` or a tap on one of
 * OUR inline buttons (callback data namespaced `HQ:`).
 *
 * Never throws — doPost must always be able to answer n8n. Returns a small
 * status object describing what happened (useful in the n8n execution log).
 *
 * @param {Object} update  the raw Telegram update object
 * @returns {{ok:boolean, handled:boolean, reason:string, command:string, chatId:(string|number)}}
 */
function handleTelegramCommand(update) {
  try {
    // --- BUTTON TAPS on our own inline keyboards -------------------------
    // Namespaced `HQ:` so we can never collide with the PREP_/PEND_ callbacks
    // that the original handler owns. Its parser already returns skip:true for
    // unrecognised data, so those keep flowing harmlessly down the old branch.
    var cbq = update && update.callback_query;
    if (cbq && String(cbq.data || "").indexOf(TG_COMMANDS.callbackPrefix) === 0) {
      return _tgHandleCallback(cbq);
    }

    var msg = update && (update.message || update.edited_message);
    if (!msg) return { ok: true, handled: false, reason: "not a message update", command: "", chatId: "" };

    var chatId = msg.chat && msg.chat.id;
    var text   = String(msg.text == null ? "" : msg.text).trim();
    if (!chatId) return { ok: true, handled: false, reason: "no chat id", command: "", chatId: "" };

    // Ignore ordinary conversation. The warehouse group is a real chat — the
    // bot must only speak when explicitly addressed with a slash command.
    if (text.charAt(0) !== "/") {
      return { ok: true, handled: false, reason: "not a command", command: "", chatId: chatId };
    }

    // Silent refusal for chats that are not on the allowlist.
    if (!_tgIsAllowed(chatId)) {
      try { console.log("telegramCommand: ignored update from non-allowlisted chat " + chatId); } catch (_) {}
      return { ok: true, handled: false, reason: "chat not allowlisted", command: "", chatId: chatId };
    }

    var parsed = _tgParse(text);
    var route  = TG_ROUTES[parsed.cmd];

    var reply;
    if (!route) {
      reply = "Unknown command: " + parsed.cmd + "\n\nTry /help";
    } else {
      try {
        reply = route.run(parsed.argStr, parsed.args, msg);
      } catch (runErr) {
        try { console.log("telegramCommand route " + parsed.cmd + " failed: " + runErr + "\n" + (runErr.stack || "")); } catch (_) {}
        reply = "⚠ " + parsed.cmd + " failed: " + String(runErr.message || runErr);
      }
    }

    // A route may return a plain string, or {text, buttons} when it wants an
    // inline keyboard. Sending lives HERE so routes never need a chat id.
    if (reply && typeof reply === 'object') _tgSend(chatId, reply.text, reply.buttons);
    else _tgSend(chatId, reply);

    return { ok: true, handled: true, reason: "", command: parsed.cmd, chatId: chatId };

  } catch (err) {
    try { console.log("handleTelegramCommand error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, handled: false, reason: String(err.message || err), command: "", chatId: "" };
  }
}


/**
 * Handle a tap on one of our own inline buttons.
 *
 * Callback data is `HQ:<action>:<arg>` — kept short because Telegram caps
 * callback_data at 64 BYTES.
 *
 * The reply EDITS the original message rather than sending a new one, so the
 * card the button lived on becomes its own result. That also removes the
 * button, which is the first line of defence against a double-tap.
 */
function _tgHandleCallback(cbq) {
  var chatId    = cbq.message && cbq.message.chat && cbq.message.chat.id;
  var messageId = cbq.message && cbq.message.message_id;
  var raw       = String(cbq.data || "").slice(TG_COMMANDS.callbackPrefix.length);

  if (!_tgIsAllowed(chatId)) {
    try { console.log("telegramCallback: ignored tap from non-allowlisted chat " + chatId); } catch (_) {}
    return { ok: true, handled: false, reason: "chat not allowlisted", command: "", chatId: chatId || "" };
  }

  var sep    = raw.indexOf(":");
  var action = sep < 0 ? raw : raw.slice(0, sep);
  var arg    = sep < 0 ? ""  : raw.slice(sep + 1);
  var handler = TG_ACTIONS[action];

  if (!handler) {
    _tgAnswerCallback(cbq.id, "Unknown action");
    return { ok: true, handled: false, reason: "unknown action " + action, command: action, chatId: chatId };
  }

  // Answer FIRST — the button stays visibly spinning until this lands, and the
  // work below can take several seconds.
  _tgAnswerCallback(cbq.id, handler.toast || "Working…");

  var text;
  try {
    text = handler.run(arg, cbq);
  } catch (err) {
    try { console.log("telegramCallback " + action + " failed: " + err + "\n" + (err.stack || "")); } catch (_) {}
    text = "⚠ " + action + " failed: " + String(err.message || err);
  }

  if (chatId && messageId) _tgEdit(chatId, messageId, text);   // no buttons => removed
  else _tgSend(chatId, text);

  return { ok: true, handled: true, reason: "", command: action, chatId: chatId };
}


// =======================================================================================
// PARSING + AUTH + REPLY
// =======================================================================================

/**
 * Split "/part@HQBot 167517" into { cmd:"/part", args:["167517"], argStr:"167517" }.
 * The "@botname" suffix is how Telegram disambiguates commands in GROUP chats —
 * it must be stripped or every group command would miss its route.
 */
function _tgParse(text) {
  var parts = String(text).trim().split(/\s+/);
  var cmd = (parts.shift() || "").toLowerCase();
  var at = cmd.indexOf("@");
  if (at > 0) cmd = cmd.slice(0, at);
  return { cmd: cmd, args: parts, argStr: parts.join(" ").trim() };
}

/** Allowed chat IDs: the admin chat plus any in the Script Property. */
function _tgAllowedChats() {
  var out = [];
  try {
    if (typeof TELEGRAM_ADMIN_CHAT_ID !== 'undefined' && TELEGRAM_ADMIN_CHAT_ID) {
      out.push(String(TELEGRAM_ADMIN_CHAT_ID).trim());
    }
  } catch (_) {}
  try {
    var extra = PropertiesService.getScriptProperties().getProperty(TG_COMMANDS.allowlistPropKey);
    if (extra) {
      String(extra).split(",").forEach(function (s) {
        var v = String(s).trim();
        if (v) out.push(v);
      });
    }
  } catch (_) {}
  return out;
}

/** True when this chat may run commands. Compared as strings — chat IDs are
 *  negative for groups and JSON may deliver them as number OR string. */
function _tgIsAllowed(chatId) {
  var id = String(chatId).trim();
  var allowed = _tgAllowedChats();
  for (var i = 0; i < allowed.length; i++) if (allowed[i] === id) return true;
  return false;
}

/**
 * EDITOR-RUN: add a chat to the command allowlist (e.g. the warehouse group).
 * Send any message in the target chat, then read its id from the n8n execution
 * log, then run this once with that id.
 */
function addTelegramCommandChat(chatId) {
  var id = String(chatId == null ? "" : chatId).trim();
  if (!id) return "❌ Pass a chat id, e.g. addTelegramCommandChat('-1001234567890')";
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty(TG_COMMANDS.allowlistPropKey) || "";
  var list = cur.split(",").map(function (s) { return String(s).trim(); }).filter(function (s) { return !!s; });
  if (list.indexOf(id) >= 0) return "Already allowed: " + id;
  list.push(id);
  props.setProperty(TG_COMMANDS.allowlistPropKey, list.join(","));
  return "✅ Allowed chat " + id + ". Full list: " + _tgAllowedChats().join(", ");
}

/** Show the current allowlist (admin chat + Script Property extras). */
function listTelegramCommandChats() {
  var l = _tgAllowedChats();
  return l.length ? l.join(", ") : "(none — TELEGRAM_ADMIN_CHAT_ID not set and no extras)";
}

/** Low-level Telegram API POST. Best-effort; failures are logged, never thrown. */
function _tgApi(method, payload) {
  try {
    var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/" + method, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      try { console.log("_tgApi " + method + " HTTP " + code + ": " + res.getContentText()); } catch (_) {}
      return false;
    }
    return true;
  } catch (e) {
    try { console.log("_tgApi " + method + " error: " + e); } catch (_) {}
    return false;
  }
}

/**
 * Post a plain-text reply, optionally with an inline keyboard.
 * @param {Array<Array<{text:string,data:string}>>} [buttons] rows of buttons;
 *        `data` is appended to TG_COMMANDS.callbackPrefix.
 */
function _tgSend(chatId, text, buttons) {
  var payload = {
    chat_id: chatId,
    text: _tgTruncate(String(text == null ? "" : text)),
    disable_web_page_preview: true
  };
  var kb = _tgKeyboard(buttons);
  if (kb) payload.reply_markup = kb;
  return _tgApi("sendMessage", payload);
}

/** Rewrite an existing message in place — used so a tapped button's card
 *  becomes its own result instead of pushing a second message into the chat. */
function _tgEdit(chatId, messageId, text, buttons) {
  var payload = {
    chat_id: chatId,
    message_id: messageId,
    text: _tgTruncate(String(text == null ? "" : text)),
    disable_web_page_preview: true
  };
  var kb = _tgKeyboard(buttons);
  if (kb) payload.reply_markup = kb;   // omit entirely => buttons are REMOVED
  return _tgApi("editMessageText", payload);
}

/**
 * Answer a callback query — this is what stops the spinner on the tapped
 * button. Telegram leaves the button visibly "loading" until it arrives, so
 * it must be sent BEFORE any slow work, not after.
 */
function _tgAnswerCallback(callbackQueryId, text) {
  return _tgApi("answerCallbackQuery", {
    callback_query_id: callbackQueryId,
    text: String(text || "").slice(0, 200)
  });
}

/** Build an inline_keyboard, or null when there are no buttons. */
function _tgKeyboard(buttons) {
  if (!buttons || !buttons.length) return null;
  var rows = buttons.map(function (row) {
    return row.map(function (b) {
      return { text: b.text, callback_data: TG_COMMANDS.callbackPrefix + b.data };
    });
  });
  return { inline_keyboard: rows };
}

/** Keep replies under Telegram's 4096-char cap, and say so when we cut. */
function _tgTruncate(s) {
  var max = TG_COMMANDS.maxReplyChars;
  if (s.length <= max) return s;
  return s.slice(0, max - 24).replace(/\n[^\n]*$/, "") + "\n\n… (truncated)";
}


// =======================================================================================
// FORMAT HELPERS  (pure — Node-testable)
// =======================================================================================

/** "—" for null/NaN, otherwise the number as-is. */
function _tgNum(v) {
  return (typeof v === 'number' && !isNaN(v)) ? String(v) : "—";
}

/** Money with no cents when whole, 2dp otherwise. "—" when unknown. */
function _tgMoney(v) {
  if (typeof v !== 'number' || isNaN(v)) return "—";
  return "$" + (Math.abs(v % 1) < 0.005 ? String(Math.round(v)) : v.toFixed(2));
}

/** Trim a name so a line stays readable in a phone-width chat bubble. */
function _tgClip(s, n) {
  var t = String(s == null ? "" : s).trim();
  if (!t) return "";
  var lim = n || 42;
  return t.length <= lim ? t : t.slice(0, lim - 1) + "…";
}


// =======================================================================================
// ROUTES
// =======================================================================================

/**
 * The command table. Adding a command = adding one entry here.
 * Each `run(argStr, args, msg)` returns the reply STRING.
 */
var TG_ROUTES = {

  "/help": {
    help: "this list",
    run: function () {
      var L = ["🤖 HQ BOT — commands", ""];
      Object.keys(TG_ROUTES).forEach(function (k) {
        L.push(k + (TG_ROUTES[k].usage ? " " + TG_ROUTES[k].usage : "") + " — " + TG_ROUTES[k].help);
      });
      L.push("");
      L.push("Read-only for now. Buttons on order cards still do PREP / PEND.");
      return L.join("\n");
    }
  },

  "/part": {
    help:  "stock, shelf, price, and which kits use a SKU",
    usage: "<sku>",
    run: function (argStr) {
      if (!argStr) return "Usage: /part <sku>\nExample: /part 167517";
      return _tgFormatPart(argStr);
    }
  },

  "/status": {
    help: "today at a glance — to grab, oldest pending, shipped, last sync",
    run: function () { return _tgFormatStatus(); }
  },

  "/order": {
    help:  "status, timeline and open cases for an order",
    usage: "<order id or SO>",
    run: function (argStr) {
      if (!argStr) return "Usage: /order <id>\nExample: /order SO-23219";
      return _tgFormatOrder(argStr);
    }
  },

  "/oos": {
    help: "what's out of stock, by aisle",
    run: function () { return _tgFormatOos(); }
  },

  "/pull": {
    help:  "pull a Zoho sales order into the DIRECT table",
    usage: "<SO or INV>",
    run: function (argStr) {
      if (!argStr) return "Usage: /pull <SO or INV>\nExample: /pull SO-23219";
      return _tgPullPreview(argStr);
    }
  }
};


/**
 * Callback-button actions. Same shape as TG_ROUTES: `run(arg)` returns the
 * text that REPLACES the message the button was attached to.
 */
var TG_ACTIONS = {
  "pull": {
    toast: "Pulling…",
    run: function (soNumber) { return _tgPullApply(soNumber); }
  },
  "cancel": {
    toast: "Cancelled",
    run: function (soNumber) { return "✖ Cancelled — " + soNumber + " was not pulled."; }
  }
};


// =======================================================================================
// /pull  —  DELIBERATELY LIMITED TO SIMPLE FIRST PULLS
// =======================================================================================
//
// A first pull (no DIRECT rows yet, every Zoho line new) is one decision:
// "take all of it." That is a button.
//
// A RE-pull is not. It carries per-line qty changes, removals, and a choice
// between insert-delta and flag-existing for each one. Squeezing that review
// onto a phone screen is exactly the cramped, flow-state-hostile surface that
// caused the 2026-05-23 incident where two kits shipped with missing
// components. So anything that isn't a clean first pull is REFUSED here and
// sent to the Pull modal, which was built for precisely that review.
//
// Double-tap is safe by construction: applyZohoPullSelection recomputes the
// diff at apply time and validates every selection against the line's CURRENT
// status, aborting all-or-nothing on any mismatch. A second tap therefore
// finds nothing still "new" and refuses. Removing the button on success is
// just the friendlier first line of defence.

/** Build the /pull confirmation card, or explain why this one needs the sheet. */
function _tgPullPreview(query) {
  var d = computeZohoSoDiff(query);
  if (!d || !d.ok) return "⚠ " + ((d && d.reason) || "Could not read that sales order.");

  var s = d.summary || {};
  var head = "⬇ " + d.soNumber +
             (d.customerName ? " · " + _tgClip(d.customerName, 34) : "") + "\n" +
             (d.totalFormatted ? d.totalFormatted + " · " : "") +
             s.totalLines + " line" + (s.totalLines === 1 ? "" : "s");

  // --- the simple-case gate ---
  if (!d.isFirstPull || s.new !== s.totalLines) {
    var why = [];
    if (!d.isFirstPull)   why.push("already partly on the sheet");
    if (s.qtyChanged)     why.push(s.qtyChanged + " qty change" + (s.qtyChanged === 1 ? "" : "s"));
    if (s.removed)        why.push(s.removed + " removed in Zoho");
    if (s.unchanged && d.isFirstPull) why.push(s.unchanged + " unchanged");
    return head + "\n\n🔒 Needs the Pull modal — " + (why.join(" · ") || "not a clean first pull") +
           ".\n\nPer-line decisions belong on a real screen. Open the sheet → Pull from Zoho.";
  }

  var L = [head, ""];
  d.lines.forEach(function (ln) {
    L.push("  " + ln.zohoQty + "× " + ln.sku +
           (ln.location && ln.location !== "NOT FOUND" ? "  " + ln.location : "  ⚠ no shelf") +
           (ln.name ? "\n      " + _tgClip(ln.name, 40) : ""));
  });
  L.push("");
  L.push("All " + s.new + " lines are new.");

  // Returning {text, buttons} keeps the sending in ONE place (the entry point)
  // instead of handing routes a chat id to send with themselves.
  return {
    text: L.join("\n"),
    buttons: [[
      { text: "✅ Pull all", data: "pull:" + d.soNumber },
      { text: "✖ Cancel",   data: "cancel:" + d.soNumber }
    ]]
  };
}

/** Apply the pull for every line, then report. Called from the button tap. */
function _tgPullApply(soNumber) {
  var d = computeZohoSoDiff(soNumber);
  if (!d || !d.ok) return "⚠ " + ((d && d.reason) || "Could not re-read that sales order.");

  // Re-assert the gate at apply time — state may have moved since the card
  // was drawn (someone could have pulled it on the sheet in the meantime).
  var s = d.summary || {};
  if (!d.isFirstPull || s.new !== s.totalLines) {
    return "⬇ " + d.soNumber + "\n\n🔒 State changed since this card was sent — it's no longer a clean first pull.\nOpen the sheet → Pull from Zoho.";
  }

  var selections = d.lines.map(function (ln) { return { sku: ln.sku, action: "insert" }; });
  var r = applyZohoPullSelection(soNumber, selections, "");

  if (!r || !r.ok) return "⬇ " + soNumber + "\n\n⚠ Pull failed: " + ((r && r.reason) || "unknown");

  return "✅ PULLED · " + r.soNumber +
         "\n\n" + r.applied.inserted + " row" + (r.applied.inserted === 1 ? "" : "s") +
         " added to DIRECT." +
         (r.skipped && r.skipped.length ? "\n⚠ " + r.skipped.length + " skipped." : "");
}


/** Render a Part dossier as a chat message. Reuses getPartData() unchanged. */
function _tgFormatPart(query) {
  var res = getPartData(query);
  if (!res || !res.ok) return "⚠ " + ((res && res.reason) || "Lookup failed.");

  var d = res.dossier;
  if (!d.found) {
    return "🔍 " + d.sku + "\n\nNot found in Master Inventory, Zoho, or the Kit Registry.";
  }

  var p = d.part;
  var L = [];

  L.push("🧩 " + d.sku + (p.title ? " · " + _tgClip(p.title, 46) : ""));
  var head = "📍 " + (p.location || "NOT FOUND");
  if (p.listingStatus) head += "   ·   " + p.listingStatus;
  L.push(head);
  L.push("");

  // --- stock ---
  L.push("STOCK");
  L.push("  On hand: " + _tgNum(p.available) + (p.oos ? "   ⛔ OUT OF STOCK" : ""));
  if (p.zohoAvailable != null || p.miAvailable != null) {
    L.push("  (Zoho " + _tgNum(p.zohoAvailable) + " · MI " + _tgNum(p.miAvailable) + ")");
  }
  if (p.committed) L.push("  Committed to open orders: " + p.committed);
  L.push("  eBay " + _tgMoney(p.ebayPrice) + " · Zoho " + _tgMoney(p.zohoPrice));

  // --- if this SKU is itself a kit ---
  if (d.isKit && d.kit) {
    var k = d.kit;
    L.push("");
    L.push("📦 THIS IS A KIT · " + (k.type || "MANUAL"));
    if (k.name) L.push("  " + _tgClip(k.name, 46));
    L.push("  Buildable now: " + _tgNum(k.buildable));
    if (k.limitedBy) L.push("  Blocked by: " + _tgClip(k.limitedBy, 52));
    if (k.complete && k.computed != null) {
      L.push("  Parts " + _tgMoney(k.partsValue) + " · computed " + _tgMoney(k.computed) +
             " · listed " + _tgMoney(k.listed));
    } else {
      L.push("  Price: ⚠ incomplete — not all components priced");
    }
    if (k.priceStatus) L.push("  " + k.priceStatus);
    if (k.unparsed) L.push("  ⚠ " + k.unparsed + " unreadable line(s) in the Zoho description");

    var comps = k.components || [];
    if (comps.length) {
      L.push("");
      L.push("  COMPONENTS (" + comps.length + ")");
      comps.slice(0, TG_COMMANDS.componentLimit).forEach(function (c) {
        L.push("   " + c.qty + "× " + c.sku + "  has " + _tgNum(c.available) +
               (c.name ? "  " + _tgClip(c.name, 26) : ""));
      });
      if (comps.length > TG_COMMANDS.componentLimit) {
        L.push("   … +" + (comps.length - TG_COMMANDS.componentLimit) + " more");
      }
    }
  }

  // --- the ripple: kits that USE this part ---
  if (d.usedIn && d.usedIn.length) {
    L.push("");
    L.push("USED IN " + d.usedIn.length + " KIT" + (d.usedIn.length === 1 ? "" : "S"));
    d.usedIn.slice(0, TG_COMMANDS.rippleLimit).forEach(function (u) {
      L.push("  " + (u.blocked ? "⛔ " : "   ") + u.kitSku +
             "  ×" + u.qtyPer +
             "  buildable " + _tgNum(u.buildable) +
             (u.blocked ? "   ← blocked by this part" : ""));
    });
    if (d.usedIn.length > TG_COMMANDS.rippleLimit) {
      L.push("  … +" + (d.usedIn.length - TG_COMMANDS.rippleLimit) + " more");
    }
  }

  // --- the payoff line ---
  if (d.unblock && d.unblock.length) {
    L.push("");
    L.push("💡 Restocking this unblocks: " + d.unblock.join(", "));
  }

  if (d.ebayUrl) { L.push(""); L.push(d.ebayUrl); }

  return L.join("\n");
}


/** Today at a glance, from the same snapshot the sidebar cockpit uses. */
function _tgFormatStatus() {
  var s;
  try { s = getDashboardSnapshot(); } catch (e) { s = null; }
  if (!s) return "⚠ Couldn't read the dashboard snapshot.";

  var L = [];
  L.push("📊 HQ · RIGHT NOW");
  L.push("");
  L.push("🛒 To grab: " + _tgNum(s.ebayGrab != null ? s.ebayGrab + (s.directGrab || 0) : s.pendingCount) +
         "   (eBay " + _tgNum(s.ebayGrab) + " · Direct " + _tgNum(s.directGrab) + ")");

  if (s.oldestPendingMinutes != null && s.oldestPendingMinutes >= 0) {
    var m = s.oldestPendingMinutes;
    var age = m < 60 ? (m + "m") : (Math.floor(m / 60) + "h " + (m % 60) + "m");
    L.push("⏱ Oldest pending: " + age + (m >= 180 ? "   ⚠ past the 3h line" : ""));
  }
  L.push("📦 Shipped today: " + _tgNum(s.shippedToday));
  L.push("📥 Received today: " + _tgNum(s.receivedToday));

  if (s.lastSyncMinutes != null && s.lastSyncMinutes >= 0) {
    L.push("");
    L.push("🔄 Last sync: " + s.lastSyncMinutes + "m ago");
  }
  return L.join("\n");
}


/** Order dossier as a chat message. Reuses getOrderCaseData() unchanged. */
function _tgFormatOrder(query) {
  var res = getOrderCaseData(query);
  if (!res || !res.ok) return "⚠ " + ((res && res.reason) || "Lookup failed.");

  var d = res.dossier;
  if (!d.found) return "🔍 " + d.orderId + "\n\nNo rows and no activity found for that order.";

  var L = [];
  L.push("📂 " + d.orderId + (d.channel ? "  ·  " + d.channel : ""));
  if (d.statuses && d.statuses.length) L.push("   " + d.statuses.join(" · "));

  // Zoho's own view, when we have it (direct orders only)
  if (d.zoho) {
    L.push("🔗 Zoho: " + (d.zoho.orderStatus || "?") +
           " · pay " + (d.zoho.payment || "?") +
           " · ship " + (d.zoho.shipment || "?") +
           (d.zoho.pulled ? " · PULLED" : ""));
  }

  var rows = d.rows || [];
  if (rows.length) {
    L.push("");
    L.push("ROWS (" + rows.length + ")");
    rows.forEach(function (r) {
      L.push("  " + r.qty + "× " + r.sku +
             (r.location ? "  " + r.location : "") +
             (r.status ? "  [" + r.status + "]" : ""));
      if (r.note) L.push("      " + _tgClip(r.note, 44));
    });
  } else {
    L.push("");
    L.push("No rows on the sheet — shipped and cleaned up, or never landed.");
  }

  var ev = d.events || [];
  if (ev.length) {
    L.push("");
    L.push("TIMELINE (last " + Math.min(ev.length, TG_COMMANDS.eventLimit) + " of " + ev.length + ")");
    ev.slice(-TG_COMMANDS.eventLimit).forEach(function (e) {
      var when = e.timestamp
        ? Utilities.formatDate(new Date(e.timestamp), WEEKLY_DIGEST.timezone, "M/d h:mm a")
        : "—";
      L.push("  " + when + "  " + e.event +
             (e.source ? " · " + e.source : "") +
             (e.picker ? " · " + e.picker : ""));
    });
  }

  var notes = d.notes || [];
  if (notes.length) {
    L.push("");
    L.push("⚠ " + notes.length + " investigation note" + (notes.length === 1 ? "" : "s") + " on file");
  }

  // NOTE: _caseLinks returns .ebay / .zoho (NOT .ebayUrl / .zohoUrl).
  if (d.links) {
    L.push("");
    if (d.links.ebay) L.push(d.links.ebay);
    if (d.links.zoho) L.push(d.links.zoho);
  }
  return L.join("\n");
}


/** The restock list, aisle-ordered — a snapshot read of the Out of Stock sheet. */
function _tgFormatOos() {
  var sheet;
  try { sheet = SpreadsheetApp.getActive().getSheetByName(OUT_OF_STOCK.sheetName); } catch (e) { sheet = null; }
  if (!sheet) return "⚠ Out of Stock sheet not found.";

  // Read ONLY the main (reorder) table — stop at the KITS divider, which is
  // a different schema and answers a different question.
  var last = sheet.getLastRow();
  var start = OUT_OF_STOCK.dataStartRow;
  if (last < start) return "📦 Out of stock: nothing listed.";

  var vals = sheet.getRange(start, 1, last - start + 1, OUT_OF_STOCK.dataWidth).getValues();
  var iSku = OUT_OF_STOCK.idx("SKU"), iLoc = OUT_OF_STOCK.idx("LOCATION"),
      iAvail = OUT_OF_STOCK.idx("AVAILABLE");

  var items = [];
  for (var i = 0; i < vals.length; i++) {
    var sku = String(vals[i][iSku] || "").trim();
    if (!sku) continue;
    if (sku.toUpperCase() === "KITS") break;           // divider — kit table below
    var av = vals[i][iAvail];
    if (typeof av === 'number' && av > 0) continue;    // in-stock watch rows
    items.push({ sku: sku, loc: String(vals[i][iLoc] || "").trim() || "NOT FOUND" });
  }

  if (!items.length) return "📦 Out of stock: nothing to reorder. ✅";

  var L = ["📦 OUT OF STOCK · " + items.length + " to reorder", ""];
  items.slice(0, TG_COMMANDS.oosLimit).forEach(function (it) {
    L.push("  " + it.loc + "   " + it.sku);
  });
  if (items.length > TG_COMMANDS.oosLimit) {
    L.push("  … +" + (items.length - TG_COMMANDS.oosLimit) + " more — open the sheet for the full list");
  }
  return L.join("\n");
}


// =======================================================================================
// WEBHOOK SUBSCRIPTION
// =======================================================================================

/**
 * EDITOR-RUN, ONCE: re-register the bot webhook with `message` updates ENABLED.
 *
 * WHY THIS EXISTS. `setWebhook()` in OrderService.js does not send
 * `allowed_updates`, and Telegram's rule for that field is RETAIN-PREVIOUS, not
 * reset-to-default. So if the subscription was ever narrowed to
 * ["callback_query"] — which is exactly what you'd do for a buttons-only bot —
 * typed commands would never reach n8n at all, with no error anywhere to show
 * for it. This sets both types explicitly so the router can actually be fed.
 *
 * Points at the n8n webhook, NOT at WEB_APP_URL: Apps Script /exec always
 * answers 302 and Telegram refuses redirects. That mistake killed every button
 * after the 2026-05-31 VPS migration (fixed 2026-06-10) — do not "simplify" it.
 *
 * Verify afterwards with getWebhookInfo(): allowed_updates should list both,
 * and last_error_message should be empty.
 */
function setWebhookWithCommands() {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/setWebhook";
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      url: N8N_TELEGRAM_CALLBACK_WEBHOOK_URL,
      allowed_updates: ["message", "callback_query"]
    }),
    muteHttpExceptions: true
  });
  var body = res.getContentText();
  Logger.log(body);
  return body;
}

/**
 * EDITOR-RUN, optional: register the command list with Telegram so typing "/"
 * in the chat shows an autocomplete menu. Cosmetic — the router works without
 * it — but it makes the commands discoverable to anyone in the chat.
 */
function registerTelegramCommandMenu() {
  var cmds = [];
  Object.keys(TG_ROUTES).forEach(function (k) {
    cmds.push({ command: k.replace(/^\//, ""), description: TG_ROUTES[k].help });
  });
  var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/setMyCommands", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ commands: cmds }),
    muteHttpExceptions: true
  });
  var body = res.getContentText();
  Logger.log(body);
  return body;
}


// =======================================================================================
// EDITOR-RUN TEST WRAPPERS  (no Telegram send — safe to run from the editor)
// =======================================================================================

/** Build the /part reply for a SKU and log it, WITHOUT sending to Telegram. */
function testTelegramPart(sku) {
  var t = _tgFormatPart(sku || "167517");
  Logger.log(t);
  return t;
}

/** Build the /status reply and log it, WITHOUT sending. */
function testTelegramStatus() {
  var t = _tgFormatStatus();
  Logger.log(t);
  return t;
}

/** Build the /help reply and log it, WITHOUT sending. */
function testTelegramHelp() {
  var t = TG_ROUTES["/help"].run("");
  Logger.log(t);
  return t;
}
