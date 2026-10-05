// =======================================================================================
// SHEET_DRAWER.gs — folders for the spreadsheet's tabs (2026-10-05)
// =======================================================================================
//
// THE PROBLEM. The file carries ~40 tabs. To keep the floor from getting lost, the
// warehouse hides everything but a handful — so working on Kit Health or Price Audit
// meant digging through Sheets' hidden-tabs menu, unhiding, working, and remembering to
// hide it again. And ⚠ HIDING IS FOR EVERYONE: a tab you unhide to work on shows up on
// every worker's screen until someone hides it back.
//
// Google Sheets has no tab folders and no per-person visibility, so this does not try to
// fake either. It gives the hidden tabs a proper DOOR (the sidebar's folder drawer) and
// makes putting the place back one button — or, once switched on, automatic at 4 AM.
//
// THE PIECES
//   • Folders by job. Every known tab is filed once in SHEET_DRAWER.folders. Click → it is
//     unhidden and opened.
//   • The FLOOR SET — the tabs workers should see. Stored in Document Properties, so it is
//     the same for everyone and editable from the drawer (★ on a tab).
//   • Tidy up — hides every visible tab that is not in the floor set.
//   • Auto-tidy (OFF by default) — the same tidy, run by the 4 AM Pick-ID reset.
//   • Work in progress — any tab the drawer does not know (a temp sheet you or a worker
//     made) is filed here automatically. "+ New work sheet" makes one with who/when stamped
//     on it; "Done" COPIES it to a Drive archive spreadsheet and only then deletes it, so a
//     finished count is never thrown away.
//   • System tabs (`__*`, Debug Log, …) never appear. They stay hidden for good.
//
// ⚠ DOES NOT NEED THE OWNER BRIDGE. Showing, hiding, adding and deleting TABS is structure,
//   not a write to All Orders' protected ranges — any editor can do it. If a tab is ever
//   sheet-protected, the call fails for staff with Google's own message, which is surfaced.
// =======================================================================================

var SHEET_DRAWER = {
  propFloor:     "SHEET_DRAWER_FLOOR",      // JSON array of tab names
  propAutoTidy:  "SHEET_DRAWER_AUTOTIDY",   // "on" | anything else = off
  propArchiveId: "SHEET_DRAWER_ARCHIVE_ID", // Drive spreadsheet that keeps finished work sheets
  metaKey:       "hq.worksheet",            // developer metadata on tabs made by the drawer

  defaultFloor: ["All orders", "Prep Queue", "Out of Stock", "Location Update", "Supplies"],

  // Order here is the order the drawer shows. A tab may appear in ONE folder only.
  folders: [
    { key: "orders",    label: "Orders",       tabs: ["All orders", "Prep Queue", "Pending Sales Orders", "Investigations", "Tracking"] },
    { key: "inventory", label: "Inventory",    tabs: ["Out of Stock", "Supplies", "Location Update", "Zoho Stock", "Master Inventory", "Provision SKUs"] },
    { key: "kits",      label: "Kits",         tabs: ["Kit Health", "Kit Registry"] },
    { key: "prices",    label: "Prices",       tabs: ["Price Audit", "Price Push Log", "Product Data Health"] },
    { key: "reports",   label: "Reports & logs", tabs: ["KPI History", "Order Archive", "Activity Log", "MPN Search Log", "Repeat Buyers", "Traffic History"] }
  ],

  // Never shown. Prefix "__" covers the helper tabs (__SparkData, __Identity, __Published…).
  systemTabs: ["Debug Log", "Telegram_Messages", "Settings", "API Usage", "Copy of All orders"],
  systemPatterns: [/^__/, /_BACKUP_/i]
};


function _sdIsSystem(name) {
  if (SHEET_DRAWER.systemTabs.indexOf(name) > -1) return true;
  for (var i = 0; i < SHEET_DRAWER.systemPatterns.length; i++) {
    if (SHEET_DRAWER.systemPatterns[i].test(name)) return true;
  }
  return false;
}

function _sdKnownSet() {
  var s = {};
  SHEET_DRAWER.folders.forEach(function (f) { f.tabs.forEach(function (t) { s[t] = f.key; }); });
  return s;
}

function _sdGetFloor() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(SHEET_DRAWER.propFloor);
    var arr = raw ? JSON.parse(raw) : null;
    if (arr && arr.length) return arr;
  } catch (_) {}
  // ⚠ FIRST USE: the floor set starts as WHATEVER IS VISIBLE NOW, not a guessed list.
  //   Measured 2026-10-05: workers had "Heads & Cranks - Temp" and "ABDUL TEMP" open —
  //   a hardcoded default would have hidden their work on the very first Tidy up.
  try {
    var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
    var vis = ss.getSheets().filter(function (sh) { return !sh.isSheetHidden() && !_sdIsSystem(sh.getName()); })
                            .map(function (sh) { return sh.getName(); });
    if (vis.length) { _sdSetFloor(vis); return vis; }
  } catch (_) {}
  return SHEET_DRAWER.defaultFloor.slice();
}

function _sdSetFloor(arr) {
  PropertiesService.getDocumentProperties().setProperty(SHEET_DRAWER.propFloor, JSON.stringify(arr));
}

function _sdWorkMeta(sheet) {
  try {
    var md = sheet.getDeveloperMetadata();
    for (var i = 0; i < md.length; i++) {
      if (md[i].getKey() === SHEET_DRAWER.metaKey) return JSON.parse(md[i].getValue() || "{}");
    }
  } catch (_) {}
  return null;
}


/**
 * Everything the drawer paints, in one read.
 * @returns {{ok, folders:[{key,label,tabs:[{name,hidden,floor,active}]}], work:[...],
 *            floor:string[], autoTidy:boolean, active:string}}
 */
function getSheetDrawer() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var active = "";
  try { active = ss.getActiveSheet().getName(); } catch (_) {}
  var floor = _sdGetFloor();
  var known = _sdKnownSet();
  var present = {};
  var work = [];

  sheets.forEach(function (sh) {
    var name = sh.getName();
    present[name] = sh;
    if (_sdIsSystem(name) || known[name]) return;
    var meta = _sdWorkMeta(sh);
    work.push({
      name: name, hidden: sh.isSheetHidden(), floor: floor.indexOf(name) > -1,
      active: name === active, by: meta ? meta.by || "" : "", at: meta ? meta.at || 0 : 0
    });
  });

  var folders = SHEET_DRAWER.folders.map(function (f) {
    return {
      key: f.key, label: f.label,
      tabs: f.tabs.filter(function (t) { return !!present[t]; }).map(function (t) {
        return { name: t, hidden: present[t].isSheetHidden(), floor: floor.indexOf(t) > -1, active: t === active };
      })
    };
  }).filter(function (f) { return f.tabs.length > 0; });

  var autoTidy = PropertiesService.getDocumentProperties().getProperty(SHEET_DRAWER.propAutoTidy) === "on";
  return { ok: true, folders: folders, work: work, floor: floor, autoTidy: autoTidy, active: active };
}


/** Unhide (if needed) and open a tab. */
function openSheetFromDrawer(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(String(name || ""));
  if (!sh) return { ok: false, message: "Tab \"" + name + "\" no longer exists." };
  if (_sdIsSystem(sh.getName())) return { ok: false, message: "That is a system tab." };
  if (sh.isSheetHidden()) sh.showSheet();
  ss.setActiveSheet(sh);
  return { ok: true, message: "Opened " + sh.getName() };
}


/** Hide one tab (the eye on an open tab). Refuses to hide the last visible tab. */
function hideSheetFromDrawer(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(String(name || ""));
  if (!sh) return { ok: false, message: "Tab not found." };
  var visible = ss.getSheets().filter(function (s) { return !s.isSheetHidden(); });
  if (visible.length <= 1) return { ok: false, message: "That is the only visible tab." };
  if (ss.getActiveSheet().getName() === sh.getName()) {
    var home = ss.getSheetByName(MAIN_SHEET_NAME);
    if (home && home.getName() !== sh.getName()) ss.setActiveSheet(home);
  }
  sh.hideSheet();
  return { ok: true, message: "Put away " + sh.getName() };
}


/**
 * TIDY UP — hide every visible tab that is not in the floor set.
 * Pure decision split out (_sdTidyPlan) so the rule is testable without Sheets.
 */
function _sdTidyPlan(tabs, floor) {
  // tabs: [{name, hidden}] ; returns names to hide. Never hides a floor tab, and never
  // empties the window — if nothing in the floor set exists, the first visible tab stays.
  var floorSet = {};
  floor.forEach(function (n) { floorSet[n] = true; });
  var anyFloorPresent = tabs.some(function (t) { return floorSet[t.name]; });
  var hide = [];
  var keptOne = anyFloorPresent;
  tabs.forEach(function (t) {
    if (t.hidden || floorSet[t.name]) return;
    if (!keptOne) { keptOne = true; return; }
    hide.push(t.name);
  });
  return hide;
}

function tidySheets() {
  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID);
  var floor = _sdGetFloor();
  var sheets = ss.getSheets();
  var plan = _sdTidyPlan(sheets.map(function (s) { return { name: s.getName(), hidden: s.isSheetHidden() }; }), floor);
  if (!plan.length) return { ok: true, hidden: 0, message: "Already tidy." };

  // Land on a floor tab first — hiding the tab you are looking at is jarring.
  try {
    var home = ss.getSheetByName(floor[0]) || ss.getSheetByName(MAIN_SHEET_NAME);
    if (home && !home.isSheetHidden()) ss.setActiveSheet(home);
  } catch (_) {}

  var failed = [];
  plan.forEach(function (n) {
    try { ss.getSheetByName(n).hideSheet(); } catch (e) { failed.push(n); }
  });
  var done = plan.length - failed.length;
  return { ok: failed.length === 0, hidden: done,
           message: "Put away " + done + " tab" + (done === 1 ? "" : "s")
                  + (failed.length ? " · could not hide: " + failed.join(", ") : "") };
}


/** ★ on a tab — add it to / remove it from the floor set. */
function setFloorTab(name, inFloor) {
  var floor = _sdGetFloor();
  var i = floor.indexOf(name);
  if (inFloor && i < 0) floor.push(name);
  if (!inFloor && i > -1) floor.splice(i, 1);
  if (!floor.length) return { ok: false, message: "The floor set cannot be empty." };
  _sdSetFloor(floor);
  return { ok: true, floor: floor };
}


/** Auto-tidy switch. Stored only; the 4 AM reset reads it. */
function setAutoTidy(on) {
  PropertiesService.getDocumentProperties().setProperty(SHEET_DRAWER.propAutoTidy, on ? "on" : "off");
  return { ok: true, autoTidy: !!on };
}

/** Called from resetDailyPickIds (4 AM). Does nothing unless switched on. */
function runAutoTidyIfOn() {
  try {
    if (PropertiesService.getDocumentProperties().getProperty(SHEET_DRAWER.propAutoTidy) !== "on") return "auto-tidy off";
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var floor = _sdGetFloor();
    var plan = _sdTidyPlan(ss.getSheets().map(function (s) { return { name: s.getName(), hidden: s.isSheetHidden() }; }), floor);
    plan.forEach(function (n) { try { ss.getSheetByName(n).hideSheet(); } catch (_) {} });
    return "auto-tidy: put away " + plan.length;
  } catch (e) {
    return "auto-tidy error: " + e;
  }
}


/**
 * + NEW WORK SHEET. Names it, stamps who/when (developer metadata — invisible, survives
 * renames), gives it a header row in the house style, optionally puts it on the floor.
 */
function createWorkSheet(name, pinToFloor) {
  var clean = String(name || "").replace(/[\[\]\*\?\/\\:]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
  if (!clean) return { ok: false, message: "Give the sheet a name." };
  if (_sdIsSystem(clean)) return { ok: false, message: "That name is reserved." };
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(clean)) return { ok: false, message: "A tab called \"" + clean + "\" already exists." };

  var sh = ss.insertSheet(clean, ss.getSheets().length);
  var by = "";
  try { by = Session.getActiveUser().getEmail() || ""; } catch (_) {}
  if (!by) { try { by = (typeof getCurrentPicker === "function" ? (getCurrentPicker() || {}).name : "") || ""; } catch (_) {} }
  sh.addDeveloperMetadata(SHEET_DRAWER.metaKey, JSON.stringify({ by: by, at: Date.now() }));

  try {
    var hdr = sh.getRange(1, 1, 1, 6);
    hdr.setValues([["SKU", "QTY", "LOCATION", "NOTE", "DONE", ""]])
       .setBackground("#1d1d1b").setFontColor("#ffd400").setFontFamily("Oswald").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 5, 120);
    sh.setTabColor("#ffd400");
  } catch (_) {}

  if (pinToFloor) {
    var floor = _sdGetFloor();
    if (floor.indexOf(clean) < 0) { floor.push(clean); _sdSetFloor(floor); }
  }
  ss.setActiveSheet(sh);
  return { ok: true, message: "Made \"" + clean + "\"", name: clean };
}


/** The Drive spreadsheet that keeps finished work sheets. Made on first use. */
function _sdArchive() {
  var props = PropertiesService.getDocumentProperties();
  var id = props.getProperty(SHEET_DRAWER.propArchiveId);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (_) {} }
  // ⚠ Created ONLY by the owner. A staff account would create it in THEIR Drive, where
  //   the owner may never see it. Staff get a clear refusal (and nothing is deleted).
  if (typeof _obIsOwner === "function" && !_obIsOwner()) {
    throw new Error("the archive is not set up yet — the owner account must press Done once first");
  }
  var arch = SpreadsheetApp.create("HQ Work Sheets Archive");
  props.setProperty(SHEET_DRAWER.propArchiveId, arch.getId());
  // Everyone who can edit All Orders can add copies to it.
  try {
    var ed = SpreadsheetApp.getActive().getEditors().map(function (u) { return u.getEmail(); }).filter(String);
    if (ed.length) arch.addEditors(ed);
  } catch (_) {}
  try { arch.getSheets()[0].setName("README").getRange("A1")
          .setValue("Finished work sheets from All Orders. Each tab is a copy taken when someone pressed Done."); } catch (_) {}
  return arch;
}


/**
 * DONE — copy the tab into the archive spreadsheet, THEN delete it. If the copy fails,
 * nothing is deleted. Only work-in-progress tabs: a known or system tab is refused, so
 * this button can never delete All Orders, the Activity Log or Kit Registry.
 */
function finishWorkSheet(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(String(name || ""));
  if (!sh) return { ok: false, message: "Tab not found." };
  var n = sh.getName();
  if (_sdIsSystem(n) || _sdKnownSet()[n] || n === MAIN_SHEET_NAME) {
    return { ok: false, message: "\"" + n + "\" is a system tab — Done only works on work sheets." };
  }
  var visible = ss.getSheets().filter(function (s) { return !s.isSheetHidden(); });
  if (visible.length <= 1 && !sh.isSheetHidden()) return { ok: false, message: "That is the only visible tab." };

  var arch, copy, url = "";
  try {
    arch = _sdArchive();
    copy = sh.copyTo(arch);
    var stamp = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd");
    var title = (n + " · " + stamp).slice(0, 95);
    var k = 2;
    while (arch.getSheetByName(title)) title = (n + " · " + stamp + " (" + (k++) + ")").slice(0, 95);
    copy.setName(title);
    url = arch.getUrl() + "#gid=" + copy.getSheetId();
  } catch (e) {
    return { ok: false, message: "Could not save a copy, so nothing was deleted: " + (e.message || e) };
  }

  if (ss.getActiveSheet().getName() === n) {
    var home = ss.getSheetByName(_sdGetFloor()[0]) || ss.getSheetByName(MAIN_SHEET_NAME);
    if (home) ss.setActiveSheet(home);
  }
  ss.deleteSheet(sh);
  var floor = _sdGetFloor(), i = floor.indexOf(n);
  if (i > -1) { floor.splice(i, 1); if (floor.length) _sdSetFloor(floor); }
  return { ok: true, message: "Saved to the archive and removed \"" + n + "\"", url: url };
}
