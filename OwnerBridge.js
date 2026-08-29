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
    "markPreparingByValues"
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
    markPreparingByValues:         function () { return markPreparingByValues(a[0]); }
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
