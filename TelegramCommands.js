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
  callbackPrefix: "HQ:",

  // Floor-note marker. MUST stay in sync with FLOOR_MARK in FloorBoard.html —
  // the board scans each open order's NOTE for this and, when found, shows the
  // text as a 📌 floor note. This is what makes /note reach the warehouse wall
  // instead of just sitting in a cell.
  floorMark: "**",

  // NOTE cells hold several appended segments; the board splits them on a
  // 3-space run, so that is the separator we must append with.
  noteSep: "   ",

  // --- /kits — the hosted kit-expansion page --------------------------------
  // The page identifies its user from Telegram's SIGNED initData, so it only
  // works when Telegram launches it as a Mini App. A plain link opened in a
  // browser tab arrives with empty initData and the server refuses — correct,
  // but it means the BUTTON TYPE matters:
  //
  //   web_app button  — launches the Mini App directly, but Telegram only
  //                     honours it in PRIVATE chats.
  //   t.me/<bot>/<app> URL button — works EVERYWHERE, groups included. Requires
  //                     creating the app once: BotFather -> /newapp.
  //
  // So: set miniAppLink once and /kits works for the whole team in the group.
  // Leave it blank and /kits still works, but only in a private chat with the
  // bot — it detects that and says so rather than sending a dead button.
  webAppUrl:   "https://hq.yassinqurabi.com/kits",
  // Named Mini App created via BotFather /newapp on 2026-08-06. A URL button
  // pointing here works in EVERY chat type; a `web_app` button is private-chat
  // only, which is why the group needed this.
  miniAppLink: "https://t.me/HighQualityMotorServiceBot/kits",

  // The Floor Board is a plain public page, so a normal URL button opens it in
  // Telegram's in-app browser. No auth, exactly as on the wall tablet.
  boardUrl: "https://hq.yassinqurabi.com"
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

/**
 * Build an inline_keyboard, or null when there are no buttons.
 *
 * THREE button kinds, and a button is exactly ONE of them — Telegram rejects
 * the whole keyboard if a button carries more than one action field:
 *   { text, data }   -> callback_data, namespaced (the default, unchanged)
 *   { text, url }    -> a link; works in every chat type
 *   { text, webApp } -> launches a Mini App; PRIVATE chats only
 */
function _tgKeyboard(buttons) {
  if (!buttons || !buttons.length) return null;
  var rows = buttons.map(function (row) {
    return row.map(function (b) {
      if (b.url)    return { text: b.text, url: b.url };
      if (b.webApp) return { text: b.text, web_app: { url: b.webApp } };
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

  "/ripple": {
    help: "which parts to restock first to unblock the most kits",
    run: function () { return previewRestockRipple(); }
  },

  // ⚠ THE SIBLING OF /ripple, AND THE DIFFERENCE MATTERS. /ripple is REACTIVE —
  // what is blocked now and what frees it. /critical is STRUCTURAL — what the kit
  // catalogue leans on, blocked or not. A part twelve healthy kits depend on never
  // appears in /ripple (it returns early on any kit that is not already blocked)
  // right up until the morning all twelve stop at once. Both help texts say which
  // is which, because asking the wrong one gives a confidently wrong answer.
  "/critical": {
    help: "which parts the most kits depend on, blocked or not",
    run: function () { return previewKitCriticality(); }
  },

  // Reorder report, scoped to what actually MOVED in the window — not a
  // catalogue sweep. `/lowstock` uses the defaults; `/lowstock 14 10` widens
  // the window to 14 days and the threshold to 10.
  "/lowstock": {
    help:  "what sold recently and is now running out",
    usage: "[days] [qty]",
    run: function (argStr, args) {
      var days = args && args[0] ? parseInt(args[0], 10) : undefined;
      var thr  = args && args[1] ? parseInt(args[1], 10) : undefined;
      return buildLowStockText(analyzeLowStock({ days: days, threshold: thr }));
    }
  },

  // Opens the hosted kit-expansion page — the ONE workflow with no answer on a
  // tablet, because the Sheets mobile app cannot render an Apps Script modal.
  //
  // Reports the queue COUNT in the message on purpose: the picker learns
  // whether it is worth opening before they open it, which is most of the value
  // of the command on a phone.
  "/kits": {
    help: "expand kits from the tablet",
    run: function (argStr, args, msg) {
      var count = null, kitLines = "";
      try {
        var q = getKitQueueForWeb();
        if (q && q.ok) {
          count = q.count;
          var show = (q.kits || []).slice(0, 6);
          for (var i = 0; i < show.length; i++) {
            kitLines += "\n  · " + show[i].kitSku + "  " + (show[i].kitName || "") +
                        (show[i].kitType === "READY" ? "  [READY]" : "");
          }
          if (q.count > show.length) kitLines += "\n  … +" + (q.count - show.length) + " more";
        }
      } catch (e) {
        try { console.log("/kits queue read failed: " + e); } catch (_) {}
      }

      var head = (count === 0)
        ? "📦 KIT EXPANSION\n\nNothing waiting — no kits need a decision."
        : "📦 KIT EXPANSION\n\n" +
          (count == null ? "Open the tablet page to see what's waiting."
                         : count + " kit" + (count === 1 ? "" : "s") + " waiting:" + kitLines);

      // The button is offered even on an EMPTY queue, deliberately. Suppressing
      // it meant the very first /kits after wiring the link up — almost always
      // on a quiet queue — came back bare, which reads as "the button is
      // broken" rather than "there is nothing to do". The queue can also change
      // between this reply and the tap.

      // A URL button to the named Mini App works in EVERY chat type, so prefer
      // it whenever it's configured. web_app buttons are private-chat only.
      if (TG_COMMANDS.miniAppLink) {
        return { text: head, buttons: [[{ text: "🔧 Open kit expansion",
                                          url: TG_COMMANDS.miniAppLink }]] };
      }

      var isPrivate = msg && msg.chat && msg.chat.type === "private";
      if (isPrivate) {
        return { text: head, buttons: [[{ text: "🔧 Open kit expansion",
                                          webApp: TG_COMMANDS.webAppUrl }]] };
      }

      // Group chat with no Mini App link configured. Say exactly why rather
      // than sending a button Telegram will ignore — and note that pasting the
      // URL into a browser genuinely will not work, since the page needs
      // Telegram to sign the identity.
      return head +
        "\n\nOpen this in a PRIVATE chat with me to get the button" +
        "\n(or set up a shareable link once: BotFather → /newapp)." +
        "\n\nA plain browser link won't work — the page needs Telegram to prove who you are.";
    }
  },

  "/pull": {
    help:  "pull a Zoho sales order into the DIRECT table",
    usage: "<SO or INV>",
    run: function (argStr) {
      if (!argStr) return "Usage: /pull <SO or INV>\nExample: /pull SO-23219";
      return _tgPullPreview(argStr);
    }
  },

  // ⭐ THE DOOR. Hand-typing into All Orders was the last major operation in this
  // system with no proper entry point, and on 2026-08-28 that cost a picked and
  // counted row: someone meant to add a missing line and overwrote a live order id
  // instead. These two commands INSERT ONLY — they have no code path that can touch
  // an existing row, so that slip is impossible here by construction.
  "/missing": {
    help:  "add a line for an item missing from a shipment",
    usage: "<original order> <sku> [qty] [note]",
    run: function (argStr) {
      if (!argStr) {
        return "Usage: /missing <original order> <sku> [qty] [note]\n" +
               "Example: /missing 05-15052-93025 212498 1";
      }
      return _tgReplacementPreview("missing", argStr);
    }
  },

  "/replacement": {
    help:  "add a replacement line for a wrong or damaged item",
    usage: "<original order> <sku> [qty] [note]",
    run: function (argStr) {
      if (!argStr) {
        return "Usage: /replacement <original order> <sku> [qty] [note]\n" +
               "Example: /replacement 19-14597-26309 171378 1 gaskets and studs only";
      }
      return _tgReplacementPreview("replacement", argStr);
    }
  },

  "/note": {
    help:  "pin a note to an order — shows on the Floor Board",
    usage: "<order> <text>",
    run: function (argStr) {
      if (!argStr) return "Usage: /note <order> <text>\nExample: /note SO-23219 call before shipping";
      return _tgAddNote(argStr);
    }
  },

  // ⏸ THE PHONE DOOR. The shipping-responsible person is remote, on their
  // phone, and the incident this exists for started with them announcing a
  // change into a chat group where it was attached to nothing and nobody had to
  // answer for it. /hold writes the SAME note the sheet writes, so it lands on
  // the board, colours the cell, and starts the escalation clock — one line
  // instead of hunting for a row in mobile Sheets.
  "/hold": {
    help:  "hold an order — do NOT hand it to the carrier",
    usage: "<order> <reason>",
    run: function (argStr) {
      if (!argStr) {
        return "Usage: /hold <order> <reason>\n" +
               "Example: /hold 24-14979-87359 buyer wants 2-Day, change the service";
      }
      return _tgHold(argStr);
    }
  },

  "/ack": {
    help:  "acknowledge a hold — records who saw it, and when",
    usage: "<order>",
    // ⭐ Routes are invoked as run(argStr, args, msg) — and `msg.from` means this
    // door knows exactly who tapped, so the note can carry a real name instead
    // of a place. The other two doors have to fall back to where the tap
    // happened; this one never does.
    run: function (argStr, args, msg) {
      if (!argStr) return "Usage: /ack <order>\nExample: /ack 24-14979-87359";
      var u = msg && msg.from;
      var who = u ? [u.first_name, u.last_name].filter(Boolean).join(" ").trim() : "";
      return _tgAckHold(argStr.trim().split(/\s+/)[0], who);
    }
  },

  "/unnote": {
    help:  "remove the floor notes from an order (leaves buyer/kit notes)",
    usage: "<order>",
    run: function (argStr) {
      if (!argStr) return "Usage: /unnote <order>\nExample: /unnote SO-23219";
      return _tgClearNote(argStr);
    }
  },

  // Needed once, to allowlist yourself for the hosted kit-expansion page:
  // the web app identifies people by Telegram USER id, which isn't visible
  // anywhere in the client. This is how you read yours.
  "/whoami": {
    help: "your Telegram user id (needed once for web-app access)",
    // Routes are invoked as run(argStr, args, msg) — the third argument is the
    // raw Telegram message, which is where `from` lives.
    run: function (argStr, args, msg) {
      var u = msg && msg.from;
      if (!u || !u.id) return "Couldn't read your user id from this message.";
      var allowed = [];
      try { allowed = listTelegramWebAppUsers(); } catch (e) {}
      var isOn = allowed.indexOf(String(u.id)) !== -1;
      return "👤 " + [u.first_name, u.last_name].filter(Boolean).join(" ") +
             (u.username ? "  @" + u.username : "") + "\n\n" +
             "user id: " + u.id + "\n" +
             "chat id: " + ((msg && msg.chat && msg.chat.id) || "?") + "\n\n" +
             (isOn ? "✅ Already allowed on the web app."
                   : "Not yet allowed on the web app. Run this ONCE from the\n" +
                     "Apps Script editor:\n\n  addTelegramWebAppUser('" + u.id + "')");
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
  "rline": {
    toast: "Adding…",
    run: function (token) { return _tgReplacementApply(token); }
  },
  "rlcancel": {
    toast: "Cancelled",
    run: function () { return "✖ Cancelled — nothing was added to the sheet."; }
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


// =======================================================================================
// /missing  +  /replacement  —  THE INSERT-ONLY DOOR
// =======================================================================================
//
// Both commands share one engine (Replacements.js). The card confirms, the button
// commits. Same two-step shape as /pull, for the same reason: this writes a real pick
// line to the floor's list, so it should take a deliberate second tap.
//
// ⚠ WHY THE PAYLOAD GOES THROUGH THE CACHE AND NOT INTO callback_data
//   Telegram caps callback_data at 64 BYTES. Order id + SKU + qty already spends ~40
//   of those, and an optional note would blow it — silently, because Telegram simply
//   refuses to render the button. Caching under a short token keeps the callback tiny
//   and is the pattern the kit-expansion modal already uses for its queue.
//
// ⚠ THE TOKEN IS SCRIPT-SCOPED, NOT USER-SCOPED. In a group chat the person who taps
//   the button is frequently NOT the person who typed the command; a user cache would
//   hand them an empty token and the tap would die with a confusing "expired".

var TG_REPLACEMENT_CACHE_SEC = 1800;   // 30 min — a card older than that should be re-run

/** Build the confirmation card for /missing or /replacement. */
function _tgReplacementPreview(kind, argStr) {
  var a = _rlParseCommandArgs(argStr);

  if (!a.originalOrder || !a.sku) {
    return "Usage: /" + kind + " <original order> <sku> [qty] [note]\n" +
           "Example: /" + kind + " 05-15052-93025 212498 1";
  }

  var p = previewReplacementLine(kind, a.originalOrder, a.sku, a.qty, a.note);
  if (!p.ok) return "⚠ " + p.error;

  var c = p.clean;
  var token = Utilities.getUuid().replace(/-/g, "").slice(0, 10);
  try {
    CacheService.getScriptCache().put(
      "rl:" + token,
      JSON.stringify({ kind: c.kind, originalOrder: c.originalOrder, sku: c.sku,
                       qty: c.qty, note: c.note }),
      TG_REPLACEMENT_CACHE_SEC);
  } catch (e) {
    return "⚠ Could not stage that line (cache unavailable) — try again.";
  }

  var L = [];
  L.push("➕ " + c.label + " LINE");
  L.push("");
  L.push("  " + c.qty + "× " + c.sku +
         (p.stock.location && p.stock.location !== "NOT FOUND"
            ? "   " + p.stock.location : "   ⚠ no shelf"));
  L.push("  on hand " + p.stock.hand);
  L.push("");
  L.push("For order " + c.originalOrder + "  (found " + p.original.where + ")");
  L.push("Column D will read:  " + c.salesOrder);
  if (c.note) L.push("Note: " + _tgClip(c.note, 60));

  if (p.warnings.length) {
    L.push("");
    p.warnings.forEach(function (w) { L.push("⚠ " + w); });
  }

  L.push("");
  L.push("It lands at the top of the eBay table as PENDING.");

  return {
    text: L.join("\n"),
    buttons: [[
      { text: "✅ Add line", data: "rline:" + token },
      { text: "✖ Cancel",   data: "rlcancel:" + token }
    ]]
  };
}


/** Commit the staged line. Called from the button tap. */
function _tgReplacementApply(token) {
  var raw;
  try {
    raw = CacheService.getScriptCache().get("rl:" + String(token || "").trim());
  } catch (e) {
    return "⚠ Could not read the staged line — re-run the command.";
  }
  if (!raw) {
    return "⏳ That card expired (30 min) — re-run the command to add the line.";
  }

  var d;
  try { d = JSON.parse(raw); }
  catch (e2) { return "⚠ The staged line was unreadable — re-run the command."; }

  // ⚠ addReplacementLine RE-VALIDATES from scratch: the original order, the stock, and
  //   the duplicate guard all run again at apply time. State can move between the card
  //   and the tap — someone may have added the same line on the sheet in between — and
  //   the duplicate refusal is what makes a double-tap safe rather than doubling a
  //   pick line. Same property /pull gets from re-running its diff.
  var r = addReplacementLine(d.kind, d.originalOrder, d.sku, d.qty, d.note, "telegram");
  if (!r || !r.ok) return "⚠ " + ((r && r.message) || "Could not add that line.");

  // Burn the token so a second tap cannot even reach the engine.
  try { CacheService.getScriptCache().remove("rl:" + token); } catch (_) {}

  var out = [r.message];
  if (r.warnings && r.warnings.length) {
    out.push("");
    r.warnings.forEach(function (w) { out.push("⚠ " + w); });
  }
  return out.join("\n");
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
// =======================================================================================
// /note  —  THE PARKED "FLOOR NOTES STEP 3", FINALLY LANDING
// =======================================================================================
//
// The Floor Board has scanned every open order's NOTE for a `**` marker since
// 2026-06-03 and rendered what follows as a 📌 floor note. The missing half was
// always "text a note from your phone → it appears on the wall." This is it.
//
// WHERE IT WRITES: every OPEN (PENDING / PREPARING) row of the order. A picker
// works line by line, so a note living only on row 1 of a 9-line order is a note
// they will miss. Terminal rows are skipped — the board only shows open orders,
// and annotating a shipped row helps nobody standing at a shelf.
//
// UN-PREFIXED, by convention: only BUYER notes carry the "Buyer Note:" tag
// (stamped at doPost). Operator notes stay bare — see the 2026-07-29 Zoho-Pull
// note work, which set that rule.
//
// FORMULA INJECTION is impossible by construction: the appended segment always
// begins with the `**` marker, so a cell can never start with `=`/`+`/`@`.

var TG_NOTE_MAX = 200;   // a floor note is a shout, not an essay

/**
 * Append a floor note to every open row of an order.
 * @param {string} argStr  "<order> <text>"
 * @returns {string} plain-text confirmation for the chat
 */
function _tgAddNote(argStr) {
  var m = String(argStr || "").trim().match(/^(\S+)\s+([\s\S]+)$/);
  if (!m) {
    return "Usage: /note <order> <text>\n" +
           "Example: /note SO-23219 call before shipping\n\n" +
           "(needs BOTH an order and some text)";
  }

  var query = m[1].trim();
  var text  = m[2].trim().replace(/\s+/g, " ");
  var clipped = false;
  if (text.length > TG_NOTE_MAX) { text = text.slice(0, TG_NOTE_MAX - 1) + "…"; clipped = true; }

  // Read-modify-write on shared cells — take the same script lock every other
  // writer here uses, so two notes landing together can't clobber each other.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return "⚠ Sheet is busy right now — try again in a moment."; }

  try {
    var found;
    try { found = lookupOrder(query); }
    catch (e) { return "⚠ Lookup failed: " + (e.message || e); }

    if (!found || !found.rows || !found.rows.length) {
      return "🔍 No rows on the sheet for " + query + ".\n" +
             "Check the id — /order " + query + " shows what the system knows.";
    }

    var open = found.rows.filter(function (r) {
      var s = String(r.status || "").trim().toUpperCase();
      return s === Schema.status.PENDING || s === Schema.status.PREPARING;
    });

    if (!open.length) {
      // Derive the state names from the rows we already hold — lookupOrder's
      // summary exposes `statuses` (an array), not a single `status`.
      var seen = {};
      found.rows.forEach(function (r) {
        var s = String(r.status || "").trim().toUpperCase();
        if (s) seen[s] = 1;
      });
      var states = Object.keys(seen).join(" / ") || "closed";
      return "✋ " + query + " has no open rows — every line is " + states + ".\n" +
             "Floor notes only reach the board while an order is still being picked.";
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "⚠ All Orders sheet not found.";

    var segment = TG_COMMANDS.floorMark + " " + text;
    var logRows = [];
    var written = 0;

    for (var i = 0; i < open.length; i++) {
      var r = open[i];
      var existing = String(r.note || "").trim();
      var next = existing ? (existing + TG_COMMANDS.noteSep + segment) : segment;
      sheet.getRange(r.row, Schema.cols.NOTE).setValue(next);
      written++;
      logRows.push(["NOTE", r.salesOrder || query, r.sku, r.qty, "telegram",
                    "floor note added", "", next]);
    }
    SpreadsheetApp.flush();

    try { logActivityBatch(logRows); }
    catch (e) { console.log("_tgAddNote: activity log failed — " + e); }

    var orderId = open[0].salesOrder || query;
    var out = "📌 Note added · " + orderId + "\n\n" +
              "  " + segment + "\n\n" +
              "On " + written + " open row" + (written === 1 ? "" : "s") +
              " · now showing on the Floor Board.";
    if (clipped) out += "\n\n(trimmed to " + TG_NOTE_MAX + " characters)";
    return out;

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


/**
 * Strip every floor-note segment from a NOTE cell, leaving everything else
 * untouched. PURE, so the surgery is Node-testable.
 *
 * A NOTE cell is a run of segments joined by TG_COMMANDS.noteSep. Only the
 * ones beginning with the floor marker are ours to remove — buyer notes
 * ("Buyer Note: …"), kit-expansion tags ("↳ from KIT-…") and Zoho flags
 * ("⚠️ ZOHO QTY…") are real data written by other parts of the system and
 * MUST survive. This is why /unnote can't just blank the cell.
 *
 * @param {string} note  current cell value
 * @returns {string} the cell value with floor notes removed
 */
function _tgStripFloorNotes(note) {
  var s = String(note == null ? "" : note);
  if (!s.trim()) return "";
  var kept = s.split(TG_COMMANDS.noteSep).filter(function (seg) {
    return seg.trim().indexOf(TG_COMMANDS.floorMark) !== 0;
  });
  return kept.join(TG_COMMANDS.noteSep).trim();
}


/**
 * Remove floor notes from an order.
 *
 * Targets EVERY matching row, not just open ones — a note added before a line
 * shipped would otherwise linger with no way to clear it from the phone.
 *
 * @param {string} argStr  "<order>"
 * @returns {string} plain-text confirmation
 */
function _tgClearNote(argStr) {
  var query = String(argStr || "").trim().split(/\s+/)[0];
  if (!query) return "Usage: /unnote <order>\nExample: /unnote SO-23219";

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return "⚠ Sheet is busy right now — try again in a moment."; }

  try {
    var found;
    try { found = lookupOrder(query); }
    catch (e) { return "⚠ Lookup failed: " + (e.message || e); }

    if (!found || !found.rows || !found.rows.length) {
      return "🔍 No rows on the sheet for " + query + ".";
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "⚠ All Orders sheet not found.";

    var cleared = 0, keptSomething = false;
    var logRows = [];

    for (var i = 0; i < found.rows.length; i++) {
      var r = found.rows[i];
      var before = String(r.note || "");
      if (before.indexOf(TG_COMMANDS.floorMark) < 0) continue;

      var after = _tgStripFloorNotes(before);
      if (after === before.trim()) continue;

      sheet.getRange(r.row, Schema.cols.NOTE).setValue(after);
      cleared++;
      if (after) keptSomething = true;
      logRows.push(["NOTE", r.salesOrder || query, r.sku, r.qty, "telegram",
                    "floor note cleared", "", after]);
    }

    if (!cleared) return "✋ No floor notes on " + query + " — nothing to clear.";

    SpreadsheetApp.flush();
    try { logActivityBatch(logRows); }
    catch (e) { console.log("_tgClearNote: activity log failed — " + e); }

    return "🧹 Floor note cleared · " + query + "\n\n" +
           "Removed from " + cleared + " row" + (cleared === 1 ? "" : "s") +
           " · gone from the Floor Board." +
           (keptSomething ? "\n\nBuyer/kit notes on those rows were left untouched." : "");

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}


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


// =======================================================================================
// PINNED CONTROL PANEL
// =======================================================================================

/**
 * Post a small "HQ CONTROL" message with permanent buttons, then PIN it.
 *
 * WHY THIS EXISTS: Telegram has no persistent per-group bot button. The menu
 * button from BotFather /setmenubutton is PRIVATE-CHAT ONLY, and `web_app`
 * inline buttons are too. The only thing that stays visible at the top of a
 * group is a PINNED MESSAGE — and a pinned message keeps its inline keyboard,
 * so its buttons stay tappable forever.
 *
 * Both buttons are URL buttons on purpose: those work in every chat type, which
 * is the whole reason the named Mini App (BotFather /newapp) was created.
 *
 * Run once from the editor, then pin the message by hand in Telegram
 * (long-press / right-click the message -> Pin).
 *
 * @param {string|number} [chatId] defaults to the admin chat
 */
function sendHqControlPanel(chatId) {
  var target = chatId || TELEGRAM_ADMIN_CHAT_ID;
  var text =
    "\u258C HQ CONTROL\n\n" +
    "Tap to open \u2014 no typing needed.\n" +
    "Pin this message so it stays at the top of the group.\n\n" +
    "Kit expansion asks Telegram who you are, so only allowlisted people get in. " +
    "Not on the list? Send /whoami and pass the id to Yassin.";

  var buttons = [];
  if (TG_COMMANDS.miniAppLink) {
    buttons.push([{ text: "\uD83D\uDD27 Kit Expansion", url: TG_COMMANDS.miniAppLink }]);
  }
  if (TG_COMMANDS.boardUrl) {
    buttons.push([{ text: "\uD83D\uDCFA Floor Board", url: TG_COMMANDS.boardUrl }]);
  }
  if (!buttons.length) return "No links configured - set TG_COMMANDS.miniAppLink first.";

  _tgSend(target, text, buttons);
  return "Sent to " + target + " - now PIN it in Telegram.";
}


/**
 * /hold — write a hold onto EVERY live row of an order.
 *
 * ⚠ IT DELIBERATELY INCLUDES SHIPPED ROWS, which is the difference between this
 * and /note. A floor note is for work still being picked; a hold is most often
 * needed AFTER the label is bought, because buying the label is what tells the
 * buyer their order shipped — which is what prompts "wait, change it". Refusing
 * to write on a shipped row would refuse exactly the case this was built for.
 *
 * ⚠ FORMULA INJECTION IS IMPOSSIBLE BY CONSTRUCTION: the appended segment always
 * begins with the word HOLD, so a cell can never start with = + or @ whatever
 * the sender types.
 */
function _tgHold(argStr) {
  var m = String(argStr || "").trim().match(/^(\S+)\s+([\s\S]+)$/);
  if (!m) {
    return "Usage: /hold <order> <reason>\n" +
           "Example: /hold 24-14979-87359 buyer wants 2-Day\n\n" +
           "(needs BOTH an order and a reason — the reason is what the floor reads)";
  }
  var query = m[1].trim();
  var text  = m[2].trim().replace(/\s+/g, " ");
  if (text.length > TG_NOTE_MAX) text = text.slice(0, TG_NOTE_MAX - 1) + "…";

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (e) { return "⚠ Sheet is busy right now — try again in a moment."; }

  try {
    var found;
    try { found = lookupOrder(query); }
    catch (e) { return "⚠ Lookup failed: " + (e.message || e); }
    if (!found || !found.rows || !found.rows.length) {
      return "🔍 No rows on the sheet for " + query + ".\n" +
             "Check the id — /order " + query + " shows what the system knows.";
    }

    var live = found.rows.filter(function (r) {
      return String(r.status || "").trim().toUpperCase() !== Schema.status.CANCELED;
    });
    if (!live.length) return "✋ " + query + " is CANCELED — there is nothing left to hold.";

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
    if (!sheet) return "⚠ All Orders sheet not found.";

    var segment = "HOLD — " + text;
    var logRows = [], written = 0, shipped = 0;
    for (var i = 0; i < live.length; i++) {
      var r = live[i];
      var existing = String(r.note || "").trim();
      if (holdNoteHasHold(existing)) continue;      // already held — do not stack
      var next = existing ? (existing + " · " + segment) : segment;
      sheet.getRange(r.row, Schema.cols.NOTE).setValue(next);
      written++;
      if (String(r.status || "").trim().toUpperCase() === Schema.status.SHIPPED) shipped++;
      logRows.push(["NOTE", r.salesOrder || query, r.sku, r.qty, "telegram",
                    "HOLD set", "", next]);
    }
    SpreadsheetApp.flush();
    try { logActivityBatch(logRows); } catch (e) { console.log("_tgHold log: " + e); }
    try { _dashBustTickCache(); } catch (e) {}

    if (!written) return "⏸ " + query + " is already held — nothing changed.";

    return "⏸ HOLD SET · " + (live[0].salesOrder || query) + "\n\n" +
           "  " + segment + "\n\n" +
           "On " + written + " row" + (written === 1 ? "" : "s") +
           (shipped ? ("  (" + shipped + " already SHIPPED — the box may still be here)") : "") +
           ".\nThe floor board will sound and take over within a minute.\n" +
           "If nobody acknowledges it in " + HOLDS.escalateAfterMin + " min, you get told.";
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * /ack — the same write the tablet's ✓ Got it makes, from a phone.
 *
 * ⚠ Worth being honest about who should use it: this records that somebody has
 * SEEN the hold. If the person who wrote the hold acknowledges their own hold,
 * the record says the floor answered when it did not — and the escalation that
 * would have told them nobody was listening is exactly what gets silenced.
 */
function _tgAckHold(query, who) {
  query = String(query || "").trim();
  if (!query) return "Usage: /ack <order>";
  var res;
  try { res = boardAckHold(query, "telegram", who); }
  catch (e) { return "⚠ Failed: " + (e.message || e); }

  if (!res || !res.ok) return "⚠ " + ((res && res.error) || "Could not acknowledge that.");
  if (res.already || !res.rows) {
    return "ℹ️ " + query + " — nothing to acknowledge (no unanswered hold on it).";
  }
  return "✓ Acknowledged · " + query + "\n\n  " + res.tag + "\n\n" +
         "Stamped on " + res.rows + " row" + (res.rows === 1 ? "" : "s") +
         ".\n⚠ The hold is STILL ON — this only records that it was seen.";
}
