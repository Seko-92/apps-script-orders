// =======================================================================================
// TELEGRAMAUTH.gs — verifying who tapped a Telegram Mini App button
// =======================================================================================
//
// THE WHOLE POINT
//   The hosted pages are public URLs. `/api/board` is fine open — the board was
//   already public on /exec and boardSetStatus is narrowed to PENDING/PREPARING.
//   `/api/kits` is NOT: it inserts inventory rows. So it needs to know WHO is
//   asking, without us building login, sessions, password storage or roles.
//
//   Telegram gives us that for free. When a Mini App opens from a button, it
//   hands the page an `initData` string SIGNED with our bot token. Verifying
//   that signature proves both that Telegram sent it and that it wasn't altered.
//   No passphrase, no accounts — the identity is cryptographic.
//
// WHY IN APPS SCRIPT AND NOT n8n
//   `Utilities.computeHmacSha256Signature` is built in, the bot token already
//   lives in Secrets.js, and n8n Code nodes need NODE_FUNCTION_ALLOW_BUILTIN set
//   before they can `require('crypto')` at all. Keeping it here also means the
//   check is version-controlled and testable rather than buried in a node only
//   one person can edit. n8n stays a dumb forwarder, exactly like the board.
//
// THE ALGORITHM (Telegram's Web App spec — get every step exact)
//   1. Parse initData as a query string.
//   2. Pull out `hash`; every OTHER field takes part in the check.
//   3. data_check_string = those fields as "key=value", sorted by key,
//      joined with "\n".
//   4. secret_key = HMAC_SHA256(key: "WebAppData", message: bot_token)
//   5. expected   = HMAC_SHA256(key: secret_key,  message: data_check_string)
//   6. Constant-time compare expected (hex) against `hash`.
//
//   ⚠ Step 4's key/message order is the classic trip-up, and it DIFFERS from the
//   older Login Widget scheme (which uses SHA256(bot_token) as the key). Getting
//   it backwards produces a check that rejects every genuine request.
//   ⚠ Apps Script's signature is computeHmacSha256Signature(VALUE, KEY) —
//   message first, key second. The reverse of how the spec reads.
//
// REPLAY WINDOW
//   A signature stays valid forever, so a captured initData would work forever
//   too. `auth_date` is inside the signed payload and therefore trustworthy —
//   anything older than TG_AUTH.maxAgeSeconds is refused.
//
// WHO IS ALLOWED
//   Signature valid only proves it came from OUR bot. Anyone who can message the
//   bot could get one. So the verified `user.id` must also appear in the
//   allowlist (Script Property TELEGRAM_WEBAPP_USERS). Run /whoami in the chat
//   to get your id, then addTelegramWebAppUser(id) once.
// =======================================================================================


var TG_AUTH = {
  usersPropKey:  "TELEGRAM_WEBAPP_USERS",
  maxAgeSeconds: 24 * 60 * 60   // reject initData older than a day
};


/**
 * Verify a Telegram Mini App initData string.
 *
 * @param {string} initData  raw query string handed to the page by Telegram
 * @returns {{ok:boolean, reason:string, userId:string, name:string, authDate:number}}
 */
function verifyTelegramInitData(initData) {
  var fail = function (reason) {
    return { ok: false, reason: reason, userId: "", name: "", authDate: 0 };
  };

  var raw = String(initData == null ? "" : initData).trim();
  if (!raw) return fail("no initData");

  // ---- 1 + 2: parse, and lift out the hash ---------------------------------
  var hash = "";
  var fields = [];
  var parts = raw.split("&");
  for (var i = 0; i < parts.length; i++) {
    if (!parts[i]) continue;
    var eq = parts[i].indexOf("=");
    if (eq < 0) continue;
    var key = decodeURIComponent(parts[i].slice(0, eq));
    var val = decodeURIComponent(parts[i].slice(eq + 1));
    if (key === "hash") { hash = val; continue; }   // excluded from the check
    fields.push({ k: key, v: val });
  }
  if (!hash) return fail("no hash in initData");
  if (!fields.length) return fail("initData has no fields");

  // ---- 3: data_check_string ------------------------------------------------
  fields.sort(function (a, b) { return a.k < b.k ? -1 : (a.k > b.k ? 1 : 0); });
  var dataCheckString = fields.map(function (f) { return f.k + "=" + f.v; }).join("\n");

  // ---- 4 + 5: the two HMACs ------------------------------------------------
  // computeHmacSha256Signature(VALUE, KEY) — message first, key second.
  var expectedHex;
  try {
    var secretKey = Utilities.computeHmacSha256Signature(TELEGRAM_BOT_TOKEN, "WebAppData");
    var sigBytes  = Utilities.computeHmacSha256Signature(
                      Utilities.newBlob(dataCheckString).getBytes(), secretKey);
    expectedHex = _tgAuthHex(sigBytes);
  } catch (e) {
    try { console.log("verifyTelegramInitData: hmac failed — " + e); } catch (_) {}
    return fail("signature computation failed");
  }

  // ---- 6: constant-time compare -------------------------------------------
  if (!_tgAuthEquals(expectedHex, String(hash).toLowerCase())) return fail("bad signature");

  // ---- replay window -------------------------------------------------------
  var authDate = 0;
  for (var j = 0; j < fields.length; j++) {
    if (fields[j].k === "auth_date") { authDate = parseInt(fields[j].v, 10) || 0; break; }
  }
  if (!authDate) return fail("no auth_date");
  var ageSec = Math.floor(Date.now() / 1000) - authDate;
  if (ageSec > TG_AUTH.maxAgeSeconds) return fail("initData expired (" + ageSec + "s old)");
  // A little clock skew forward is normal; a lot means something is wrong.
  if (ageSec < -300) return fail("auth_date is in the future");

  // ---- identity ------------------------------------------------------------
  var userId = "", name = "";
  for (var u = 0; u < fields.length; u++) {
    if (fields[u].k !== "user") continue;
    try {
      var parsed = JSON.parse(fields[u].v);
      userId = String(parsed.id || "");
      name   = [parsed.first_name, parsed.last_name].filter(Boolean).join(" ") ||
               parsed.username || "";
    } catch (e) { /* malformed user blob — caught by the empty check below */ }
    break;
  }
  if (!userId) return fail("no user in initData");

  return { ok: true, reason: "", userId: userId, name: name, authDate: authDate };
}


/**
 * Verify AND check the allowlist — the gate every writing web action should use.
 *
 * A valid signature only proves the data came from our bot; anyone who can
 * message it could obtain one. Membership is the second half of the check.
 *
 * @returns {{ok:boolean, reason:string, userId:string, name:string}}
 */
function authorizeWebAppUser(initData) {
  var v = verifyTelegramInitData(initData);
  if (!v.ok) return v;

  var allowed = listTelegramWebAppUsers();
  if (!allowed.length) {
    return { ok: false, userId: v.userId, name: v.name,
             reason: "no web-app users configured — run addTelegramWebAppUser('" + v.userId + "') once" };
  }
  if (allowed.indexOf(String(v.userId)) === -1) {
    try { console.log("authorizeWebAppUser: rejected user " + v.userId + " (" + v.name + ")"); } catch (_) {}
    return { ok: false, userId: v.userId, name: v.name, reason: "not authorised" };
  }
  return v;
}


// =======================================================================================
// ALLOWLIST (editor helpers — run once)
// =======================================================================================

/** @returns {Array<string>} allowed Telegram user ids */
function listTelegramWebAppUsers() {
  var raw = PropertiesService.getScriptProperties().getProperty(TG_AUTH.usersPropKey) || "";
  return raw.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
}

/** Add a Telegram user id to the web-app allowlist. Idempotent. */
function addTelegramWebAppUser(userId) {
  var id = String(userId || "").trim();
  if (!id) return "⚠ No user id given. Run /whoami in Telegram to get yours.";
  var list = listTelegramWebAppUsers();
  if (list.indexOf(id) !== -1) return "Already allowed: " + id;
  list.push(id);
  PropertiesService.getScriptProperties().setProperty(TG_AUTH.usersPropKey, list.join(","));
  return "✅ Added " + id + " — now " + list.length + " allowed.";
}

/**
 * ONE-TIME BOOTSTRAP — select this in the editor's function dropdown and Run.
 *
 * WHY IT EXISTS: the Apps Script editor's Run button cannot pass ARGUMENTS, so
 * `addTelegramWebAppUser('123')` is not directly runnable from the UI. Same
 * reason `setMyPricePushPassphraseNow()` exists. Edit the list below, run once,
 * then check the Execution log for the result.
 *
 * Get an id by sending /whoami to the bot.
 */
function addWebAppUsersNow() {
  var ids = [
    "1654742718"   // Yassin
    // , "..."     // add more here, one per line
  ];
  var out = ids.map(addTelegramWebAppUser);
  out.push("Allowed now: " + listTelegramWebAppUsers().join(", "));
  var msg = out.join("\n");
  try { console.log(msg); } catch (_) {}
  return msg;
}

/** Remove a Telegram user id from the web-app allowlist. */
function removeTelegramWebAppUser(userId) {
  var id = String(userId || "").trim();
  var list = listTelegramWebAppUsers().filter(function (x) { return x !== id; });
  PropertiesService.getScriptProperties().setProperty(TG_AUTH.usersPropKey, list.join(","));
  return "Removed " + id + " — now " + list.length + " allowed.";
}


// =======================================================================================
// PRIVATE
// =======================================================================================

/** Byte array (signed, as Apps Script returns) → lowercase hex. */
function _tgAuthHex(bytes) {
  var out = "";
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];   // Apps Script bytes are signed
    var h = b.toString(16);
    out += (h.length === 1 ? "0" : "") + h;
  }
  return out;
}

/**
 * Constant-time string compare. A plain === can leak how much of the hash
 * matched via timing — irrelevant at our scale, but this is a signature check
 * and the correct shape costs three lines.
 */
function _tgAuthEquals(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}
