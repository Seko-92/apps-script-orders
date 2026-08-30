// =======================================================================================
// OWNERBRIDGE.JS — let a locked sheet keep a working sidebar
// =======================================================================================
//
// THE PROBLEM IT SOLVES
//   2026-08-29: hard protection on All Orders was built, installed, verified — and removed
//   within a day, because a company PC could not add a line. `google.script.run` executes
//   AS THE INVOKING USER, so every staff sidebar action that writes All Orders was blocked:
//   kit expansion, Zoho pull, sort, update locations, cleanup, add rows, bulk status.
//
//   ⚠⚠ AND THE TEST THAT "PROVED" THE LOCK SAFE MEASURED THE WRONG POPULATION. An incognito
//   window has no sidebar and no ⚙️ menu at all, so protection appeared to break nothing.
//   The staff are not anonymous — they sign in with a company account and use the sidebar
//   all day. The incognito window answered "what does a stranger see", and only "what do my
//   employees see" ever mattered.
//
// ⭐ THE WAY OUT, AND IT IS ALREADY WHAT THE WAREHOUSE RUNS ON
//   Sheets protection checks WHO, never WHERE FROM — confirmed against Google's docs, not
//   reasoned about. The only documented separation is running the write as the OWNER through
//   a Web App round trip. That was rejected in August on latency, but this project already
//   does exactly that every day: boardSetStatus, boardAdjust, boardLeft and boardMissingLine
//   all write All Orders through doPost, which is precisely why a picker WITH NO GOOGLE
//   ACCOUNT AT ALL can tap ✓ Pick.
//
//   staff clicks a sidebar button
//     → google.script.run       runs as STAFF  → protection would block the write
//     → _postToSelf()           UrlFetchApp → WEB_APP_URL
//     → doPost 'runAsOwner'     runs as OWNER → writes the sheet ✓
//
// ⭐⭐ THE HOP IS SERVER-SIDE, AND THAT IS THE SECURITY OF IT. If the sidebar called /exec
//   directly, APP_SECRET_TOKEN would have to sit in HTML that staff can read — the same
//   token n8n uses for everything. Keeping the hop inside Apps Script means Secrets.js never
//   reaches a browser, and the client changes NOT AT ALL.
//
// ⭐ HOW A GUARDED FUNCTION IS WRITTEN — one line, no rename, no sidebar change:
//
//     function sortDirectTable() {
//       if (!_obIsOwner()) return _asOwner('sortDirectTable', []);
//       ... original body, untouched ...
//     }
//
//   Called by the owner  → runs directly, no latency.
//   Called by staff      → hops; doPost runs the SAME function as the owner, where
//                          _obIsOwner() is true, so it falls through to the body.
//
// ⚠ WHAT MAY CROSS THE HOP: VALUES, NEVER ROW NUMBERS. A row number captured before a round
//   trip is the 2026-05-08 / 2026-08-21 row-shift class, twice bitten here. Pass SKU +
//   SALES_ORDER and let the far side re-resolve, exactly as updateOrderStatus(orderId, sku)
//   and _findKitRowBySkuAndSo already do.
//
// ⚠ doPost HAS NO ACTIVE RANGE, SHEET OR UI. Anything reading getActiveRangeList() must
//   resolve its selection to values BEFORE hopping. Reads are never blocked by protection,
//   so the split is natural: READ AS THE USER, WRITE AS THE OWNER.
// =======================================================================================


var OWNER_BRIDGE = {
  // ⚠⚠ THE ALLOWLIST IS THE SECURITY BOUNDARY. Same posture as DOPOST_LOCK_FREE: anything
  //   unnamed is refused, so a function added later is CLOSED by default rather than
  //   silently exposed. Dispatch is a literal map built at call time — never this[name],
  //   which would hand a caller constructor / __proto__ / any global function.
  //
  // ⚠ Built INSIDE a function on purpose. Apps Script concatenates root files in an
  //   unspecified order, so a top-level object literal referencing functions from other
  //   files can evaluate before they exist. Resolving at call time cannot.
  actionNames: [
    "sortEbayTable", "sortDirectTable",
    "addRowsTableOne", "addRowsTableTwo",
    "runDeleteEmptyRowsTableOne", "runDeleteEmptyRowsTableTwo",
    "runUpdateLocationsTableOne", "runUpdateLocationsTableTwo",
    "refreshKitSkuMarkers", "refreshAllOrdersEnrichment",
    "highlightAllDuplicates", "clearAllDuplicateHighlights",
    "setupHandConditionalFormatting",
    "commitKitFromModal", "applyZohoPullSelection",
    "markPreparingByValues",
    "addReplacementFromSidebar", "recomputeHandFromZohoStock"
  ],

  ownerEmailKey: "OWNER_EMAIL",   // Secrets.js constant; Script Property overrides it
  timeoutMs: 90000
};


// ⚠⚠ SET ONLY BY doPost's runAsOwner HANDLER, and it makes recursion IMPOSSIBLE. Without
//   it, an _obIsOwner() that returned false on the owner side would hop to itself forever
//   and burn the entire runtime quota in minutes.
var _OB_IN_OWNER_CONTEXT = false;


/**
 * Is this execution already running as the owner?
 *
 * ⚠ AN UNKNOWN IDENTITY HOPS. getEffectiveUser() returns "" in several contexts, and the
 *   hop works for EVERYBODY — the direct path is only an optimisation that spares the owner
 *   ~2s. Guessing "probably the owner" would silently reintroduce the exact blocked-write
 *   failure this file exists to remove.
 */
function _obIsOwner() {
  if (_OB_IN_OWNER_CONTEXT) return true;

  var owner = "";
  try {
    owner = String(PropertiesService.getScriptProperties()
              .getProperty(OWNER_BRIDGE.ownerEmailKey) || "").trim();
  } catch (e) {}
  if (!owner && typeof OWNER_EMAIL !== "undefined") owner = String(OWNER_EMAIL || "").trim();
  if (!owner) return false;                       // not configured → hop

  var me = "";
  try { me = String(Session.getEffectiveUser().getEmail() || "").trim(); } catch (e) { me = ""; }
  if (!me) return false;                          // anonymous / unknown → hop

  return me.toLowerCase() === owner.toLowerCase();
}


/**
 * Run an allowlisted action as the owner, by asking our own web app to do it.
 *
 * @param {string} name  must be in OWNER_BRIDGE.actionNames
 * @param {Array}  args
 */
function _asOwner(name, args) {
  if (OWNER_BRIDGE.actionNames.indexOf(name) === -1) {
    throw new Error("_asOwner: '" + name + "' is not an allowlisted owner action.");
  }
  var res = _obPostToSelf({ action: "runAsOwner", fn: name, args: args || [] });
  if (!res.ok) throw new Error(res.error || "The sheet refused that action.");
  return res.result;
}


/**
 * POST to our own /exec and parse the reply HONESTLY.
 *
 * ⚠⚠ APPS SCRIPT ANSWERS HTTP 200 WITH AN ERROR PAGE. /exec returns a 302, UrlFetchApp
 *   follows it, and a 404 or a crash comes back as 200 with Google's HTML. n8n banked
 *   those as successes for weeks (2026-08-28). So the BODY decides, never the status code.
 */
function _obPostToSelf(payload) {
  if (typeof WEB_APP_URL === "undefined" || !WEB_APP_URL) {
    return { ok: false, error: "WEB_APP_URL is not set in Secrets.js." };
  }
  payload.token = APP_SECRET_TOKEN;

  var raw = "";
  try {
    var resp = UrlFetchApp.fetch(WEB_APP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true      // ⚠ followRedirects defaults TRUE — /exec needs it
    });
    raw = resp.getContentText() || "";
  } catch (e) {
    return { ok: false, error: "Could not reach the sheet's web app: " + e };
  }

  var parsed = null;
  try { parsed = JSON.parse(raw); } catch (e) { parsed = null; }
  if (!parsed || typeof parsed !== "object") {
    // Google's error page, or a deployment that no longer exists.
    return { ok: false, error: "The web app returned a page instead of a result — the " +
                              "deployment may be out of date. Cut a New Version." };
  }
  if (parsed.ok === false || parsed.status === "error") {
    return { ok: false, error: parsed.error || parsed.message || "refused" };
  }
  return { ok: true, result: parsed.result };
}


/**
 * doPost's half. Runs an allowlisted action in the OWNER's execution context.
 *
 * ⚠ Dispatch is an explicit literal map. `this[name]` would expose every global function in
 *   the project — and constructor / __proto__ / toString besides — to anyone holding the
 *   token. The allowlist above is checked FIRST, and this map is the second gate.
 */
function _obRunAsOwner(fnName, args) {
  if (OWNER_BRIDGE.actionNames.indexOf(fnName) === -1) {
    return { ok: false, error: "Not an allowlisted owner action: " + fnName };
  }
  var a = args || [];

  var map = {
    sortEbayTable:                 function () { return sortEbayTable(); },
    sortDirectTable:               function () { return sortDirectTable(); },
    addRowsTableOne:               function () { return addRowsTableOne(); },
    addRowsTableTwo:               function () { return addRowsTableTwo(); },
    runDeleteEmptyRowsTableOne:    function () { return runDeleteEmptyRowsTableOne(); },
    runDeleteEmptyRowsTableTwo:    function () { return runDeleteEmptyRowsTableTwo(); },
    runUpdateLocationsTableOne:    function () { return runUpdateLocationsTableOne(); },
    runUpdateLocationsTableTwo:    function () { return runUpdateLocationsTableTwo(); },
    refreshKitSkuMarkers:          function () { return refreshKitSkuMarkers(); },
    refreshAllOrdersEnrichment:    function () { return refreshAllOrdersEnrichment(); },
    highlightAllDuplicates:        function () { return highlightAllDuplicates(); },
    clearAllDuplicateHighlights:   function () { return clearAllDuplicateHighlights(); },
    setupHandConditionalFormatting:function () { return setupHandConditionalFormatting(); },
    commitKitFromModal:            function () { return commitKitFromModal(a[0], a[1], a[2], a[3], a[4]); },
    applyZohoPullSelection:        function () { return applyZohoPullSelection(a[0], a[1], a[2]); },
    markPreparingByValues:         function () { return markPreparingByValues(a[0]); },
    addReplacementFromSidebar:     function () { return addReplacementFromSidebar(a[0], a[1], a[2], a[3], a[4]); },
    recomputeHandFromZohoStock:    function () { return recomputeHandFromZohoStock(); }
  };

  var fn = map[fnName];
  if (typeof fn !== "function") {
    return { ok: false, error: "No implementation registered for: " + fnName };
  }

  var was = _OB_IN_OWNER_CONTEXT;
  _OB_IN_OWNER_CONTEXT = true;
  try {
    return { ok: true, result: fn() };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  } finally {
    _OB_IN_OWNER_CONTEXT = was;
  }
}


/**
 * Report whether the bridge can work, WITHOUT changing anything.
 * ⚠ Zero-arg — the editor Run button cannot pass one. Output goes to the EXECUTION LOG.
 */
function diagnoseOwnerBridge() {
  var L = ["══ OWNER BRIDGE ══", ""];

  var owner = "";
  try { owner = String(PropertiesService.getScriptProperties()
          .getProperty(OWNER_BRIDGE.ownerEmailKey) || "").trim(); } catch (e) {}
  if (!owner && typeof OWNER_EMAIL !== "undefined") owner = String(OWNER_EMAIL || "").trim();

  var me = "";
  try { me = String(Session.getEffectiveUser().getEmail() || "").trim(); } catch (e) {}

  L.push("owner configured:   " + (owner || "⚠ NOT SET — everyone will hop, including you"));
  L.push("running as:         " + (me || "(unknown — this execution would hop)"));
  L.push("treated as owner:   " + (_obIsOwner() ? "YES — direct writes, no latency" : "no — writes hop through /exec"));
  L.push("WEB_APP_URL set:    " + (typeof WEB_APP_URL !== "undefined" && WEB_APP_URL ? "YES" : "⚠ NO"));
  L.push("allowlisted actions: " + OWNER_BRIDGE.actionNames.length);
  L.push("");

  // every allowlisted name must resolve to something callable
  var missing = OWNER_BRIDGE.actionNames.filter(function (n) {
    try { return typeof eval(n) !== "function"; } catch (e) { return true; }
  });
  L.push(missing.length ? "⚠⚠ NAMES WITH NO FUNCTION: " + missing.join(", ")
                        : "every allowlisted name resolves to a function ✓");
  L.push("");
  L.push("── the round trip ──");
  var probe = _obPostToSelf({ action: "runAsOwner", fn: "__probe__", args: [] });
  L.push(probe.ok ? "⚠ unexpected: the probe SUCCEEDED (it should be refused)"
                  : "reply: " + probe.error);
  L.push("");
  L.push("A refusal naming '__probe__' means the hop works AND the allowlist holds —");
  L.push("that is the healthy result. 'returned a page' means the deployment is stale:");
  L.push("cut a New Version.");

  var out = L.join("\n");
  console.log(out);
  return out;
}


/**
 * Refuse an OWNER-ONLY action for anybody else, in words rather than a stack trace.
 *
 * ⚠⚠ THE LOCK CONTROLS MUST NEVER BE ALLOWLISTED IN OWNER_BRIDGE.actionNames. The bridge
 *   exists so staff WRITES run as the owner — pointing it at unprotectAllOrdersSheet would
 *   let any staff member remove the protection, which makes the lock self-defeating. The
 *   allowlist holds the 16 sheet writers and nothing else; keep it that way.
 *
 * ⚠ And a sidebar button must not merely fail. A sheet editor CAN create a protection, so
 *   staff clicking "Lock" would quietly make one they own — a worse state than refusing.
 *
 * @returns {string|null} a refusal to return to the caller, or null to proceed
 */
function _obRequireOwner(what) {
  if (_obIsOwner()) return null;
  return "🔒 " + what + " is owner-only.\n\n" +
         "This control changes who may edit the sheet, so it deliberately does NOT go " +
         "through the owner bridge — otherwise anyone could switch the protection off. " +
         "Ask Yassin to run it.";
}


// =======================================================================================
// THE ROLLOUT — one Run, with the pre-flight enforced rather than remembered
// =======================================================================================

/**
 * Prove the bridge works END TO END, without changing any policy.
 *
 * ⭐ IT DOES A REAL ROUND-TRIPPED WRITE, not just a reachability ping. `_asOwner` always
 *   hops — even for you — so this exercises the exact path a staff member's sidebar click
 *   will take: UrlFetchApp → /exec → doPost → runAsOwner → the function, as the owner.
 *   `refreshKitSkuMarkers` is chosen because it is idempotent and writes only number
 *   formats, so running it twice costs nothing and changes no data.
 *
 * ⚠ Zero-arg — the editor Run button cannot pass one. Output goes to the EXECUTION LOG.
 */
function verifyOwnerBridge() {
  var L = ["══ OWNER BRIDGE · END-TO-END ══", ""];
  var ok = true;

  // 1 · the allowlist, over the wire
  var probe = _obPostToSelf({ action: "runAsOwner", fn: "__probe__", args: [] });
  if (probe.ok) {
    L.push("1 · allowlist   ⚠⚠ FAILED — an unlisted action was ACCEPTED.");
    ok = false;
  } else if (/__probe__/.test(probe.error || "")) {
    L.push("1 · allowlist   ✓ the hop works and the allowlist held");
  } else {
    L.push("1 · allowlist   ⚠⚠ FAILED — " + probe.error);
    L.push("                a page instead of a result means the deployment is stale:");
    L.push("                cut a New Version. 'Could not reach' means WEB_APP_URL is wrong.");
    ok = false;
  }

  // 2 · a real write, forced through the hop
  try {
    var res = _asOwner("refreshKitSkuMarkers", []);
    L.push("2 · real write  ✓ ran as the owner through /exec — " + String(res).slice(0, 60));
  } catch (e) {
    L.push("2 · real write  ⚠⚠ FAILED — " + e.message);
    ok = false;
  }

  // 3 · who am I, and would I hop?
  L.push("3 · identity    " + (_obIsOwner()
        ? "you are recognised as the owner — your own writes stay direct"
        : "⚠ you are NOT recognised as the owner; check OWNER_EMAIL in Secrets.js"));

  L.push("");
  L.push(ok ? "✅ THE BRIDGE WORKS. installAllOrdersLock() will proceed."
            : "❌ DO NOT INSTALL THE LOCK YET — fix the above first.");
  var out = L.join("\n");
  console.log(out);
  return out;
}


/**
 * Install the All Orders lock, but ONLY once the bridge is proven.
 *
 * ⚠⚠ THIS ORDER IS THE WHOLE POINT, AND REVERSING IT IS A FLOOR OUTAGE. Protection without
 *   a working bridge removes every staff sidebar write — kit expansion, Zoho pull, sort,
 *   update locations, cleanup, add rows, bulk status — with no route around it. That is
 *   exactly what happened on 2026-08-29, and it was removed the same day.
 *
 * ⚠ It also refuses without N8N_SHEETS_ACCOUNT, because `E5. Delete SHIPPED Row` writes All
 *   Orders DIRECTLY via the Sheets node as a NAMED account. Locking without that exception
 *   stops the ~1 AM sweep, and the symptom — a sheet filling with shipped rows — shows days
 *   later. It is the only silent failure in the whole design.
 *
 * Rollback at any point: unprotectAllOrdersSheet().
 */
function installAllOrdersLock() {
  var denied = _obRequireOwner("Locking All Orders");
  if (denied) { console.log(denied); return denied; }

  var L = ["══ INSTALL THE ALL ORDERS LOCK ══", ""];

  // ── gate 1 · the bridge must actually work ──────────────────────────────────────────
  var probe = _obPostToSelf({ action: "runAsOwner", fn: "__probe__", args: [] });
  var bridgeOk = !probe.ok && /__probe__/.test(probe.error || "");
  if (!bridgeOk) {
    L.push("❌ REFUSED — the owner bridge is not answering.");
    L.push("   " + (probe.error || "the probe unexpectedly succeeded"));
    L.push("");
    L.push("   Without it, locking this sheet removes EVERY staff sidebar write with no");
    L.push("   route around it. Run verifyOwnerBridge() and fix what it reports first.");
    console.log(L.join("\n")); return L.join("\n");
  }
  L.push("1 · bridge      ✓ answering, allowlist holding");

  // ── gate 2 · n8n's direct writer must have an exception ─────────────────────────────
  var acct = "";
  try { acct = String(PropertiesService.getScriptProperties()
          .getProperty(ALL_ORDERS_LOCK.n8nAccountKey) || "").trim(); } catch (e) {}
  if (!acct) {
    L.push("2 · n8n acct    ❌ NOT SET");
    L.push("");
    L.push("   Set the Script Property '" + ALL_ORDERS_LOCK.n8nAccountKey + "' to the Google");
    L.push("   account on n8n's Sheets credential — or to the literal 'none' if you have");
    L.push("   CONFIRMED in the live n8n UI that nothing writes All Orders outside Apps");
    L.push("   Script. setN8nSheetsAccountNow() does it in one Run.");
    console.log(L.join("\n")); return L.join("\n");
  }
  L.push("2 · n8n acct    ✓ " + acct);

  // ── install ─────────────────────────────────────────────────────────────────────────
  var res = protectAllOrdersSheet();
  L.push("3 · protection  " + res);
  L.push("");
  L.push("── NOW TEST IT THE WAY STAFF ACTUALLY WORK ──");
  L.push("  ⚠⚠ NOT in an incognito window. That is the mistake 2026-08-29 made: an");
  L.push("     anonymous user has no sidebar and no ⚙️ menu at all, so protection looked");
  L.push("     harmless. Your staff sign in with a company account and use the sidebar all");
  L.push("     day — that is the population that matters.");
  L.push("");
  L.push("  1. signed in as staff: a sort, Update Locations, a kit commit, a Zoho pull");
  L.push("  2. typing into SKU / QTY / LOCATION / SALES ORDER is refused");
  L.push("  3. NOTE, STATUS and LEFT still accept, and the F2/H2 dropdowns still work");
  L.push("  4. ⚠ let a real n8n sync land, then confirm the ~1 AM sweep still deletes");
  L.push("     shipped rows — the only failure here that is silent");
  L.push("");
  L.push("  Rollback at any point: unprotectAllOrdersSheet()");

  var out = L.join("\n");
  console.log(out);
  return out;
}


/**
 * Set the n8n Sheets account in one Run — the editor Run button cannot pass arguments,
 * a trap this project has walked into three times.
 *
 * ⚠ EDIT THE VALUE BELOW FIRST. Use the account email from n8n's Google Sheets credential,
 *   or the literal "none" if you have confirmed nothing outside Apps Script writes to All
 *   Orders. Getting this wrong stops the ~1 AM shipped-row sweep, silently.
 *
 * ⚠⚠ THIS IS THE ONLY DEFINITION. A second one lived in BrandTheme.js until 2026-08-30 —
 *   Apps Script concatenates root files into ONE global scope in an unspecified order, so
 *   which body ran was undefined, and someone editing VALUE in one copy could have had the
 *   other execute. Run the duplicate scan before adding any top-level name:
 *   design-lab/test-global-collisions.js
 */
function setN8nSheetsAccountNow() {
  var VALUE = "";   // ← put the n8n Sheets account email here, or "none"

  if (!VALUE) {
    var msg = "❌ Edit VALUE inside setN8nSheetsAccountNow() first — the n8n Sheets " +
              "account email, or the literal '" + ALL_ORDERS_LOCK.noneSentinel + "'.";
    console.log(msg); return msg;
  }

  // ⚠ Delegates to the ONE writer in BrandTheme.js rather than setting the property here,
  //   so the validation that protects the nightly sweep cannot drift between two entry
  //   points — which is exactly how the duplicate above went unnoticed.
  var out = setN8nSheetsAccount(VALUE);
  if (out.indexOf("✅") === 0) out += "\n   Now run installAllOrdersLock().";
  console.log(out); return out;
}
