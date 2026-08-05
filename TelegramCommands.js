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
  componentLimit: 10      // components listed for a kit
};


// =======================================================================================
// ENTRY POINT
// =======================================================================================

/**
 * Handle one Telegram `message` update carrying a text command.
 *
 * Never throws — doPost must always be able to answer n8n. Returns a small
 * status object describing what happened (useful in the n8n execution log).
 *
 * @param {Object} update  the raw Telegram update object
 * @returns {{ok:boolean, handled:boolean, reason:string, command:string, chatId:(string|number)}}
 */
function handleTelegramCommand(update) {
  try {
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

    _tgSend(chatId, reply);
    return { ok: true, handled: true, reason: "", command: parsed.cmd, chatId: chatId };

  } catch (err) {
    try { console.log("handleTelegramCommand error: " + err + "\n" + (err.stack || "")); } catch (_) {}
    return { ok: false, handled: false, reason: String(err.message || err), command: "", chatId: "" };
  }
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

/** Post a plain-text reply. Best-effort; a Telegram failure is logged, not thrown. */
function _tgSend(chatId, text) {
  try {
    var body = _tgTruncate(String(text == null ? "" : text));
    var res = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: body, disable_web_page_preview: true }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      try { console.log("_tgSend HTTP " + code + ": " + res.getContentText()); } catch (_) {}
      return false;
    }
    return true;
  } catch (e) {
    try { console.log("_tgSend error: " + e); } catch (_) {}
    return false;
  }
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
  }
};


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
