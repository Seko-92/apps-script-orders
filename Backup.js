// =======================================================================================
// BACKUP.gs — nightly copy of the sheets that cannot be rebuilt (2026-10-05)
// =======================================================================================
//
// WHY. Some tabs hold the only copy of what they record — the Activity Log (who picked
// what), the Order Archive, KPI History, Investigations, the Price Push Log. Nothing
// re-creates them. Version history can roll the WHOLE file back, but that also rolls
// back every order, pick and note since (the Sept-17 Prep Queue incident nearly did
// exactly that). A backup lets you restore ONE tab and leave the rest alone.
//
// HOW. Every night (~3 AM Houston) a NEW spreadsheet "HQ Backup · yyyy-MM-dd" is made in
// the Drive folder "HQ Backups", and each tab in BACKUP.sheets is copied into it.
//   ⚠ A FRESH FILE, NOT A COPY OF THE WHOLE SPREADSHEET. A file copy would carry the
//     bound Apps Script — and opening that copy runs onOpen, whose code opens the LIVE
//     sheet by id. A fresh file has no code, so opening a backup can never touch prod.
//   ⚠ VALUES, NOT FORMULAS. A copied tab keeps its formulas, and any that point at
//     another tab break in a file where that tab is missing. Each copy is overwritten
//     with its source's VALUES, so the backup shows what the sheet showed.
//   • Kept BACKUP.keepDays (30) days; older backup files go to the Drive TRASH (Drive
//     keeps trash 30 more days — nothing is destroyed outright by this code).
//   • A failure sends ONE Telegram message to the admin chat, so the backup can never
//     stop silently. Last result is stored and shown in the sidebar.
//
// RESTORE (also written into each backup's README tab): open the dated backup → right-
// click the tab → "Copy to" → "Existing spreadsheet" → All Orders. Then rename / move the
// rows you need. Never paste over a live tab without looking first.
//
// PUBLIC: runNightlyBackup() (trigger target + "Back up now"), setupNightlyBackup(),
//         removeNightlyBackup(), getBackupStatus(), openBackupsFolderUrl()
// =======================================================================================

var BACKUP = {
  folderName:  "HQ Backups",
  propFolder:  "BACKUP_FOLDER_ID",
  propLast:    "BACKUP_LAST",           // JSON {at, ok, file, url, sheets, message}
  hour:        3,                        // America/Chicago — before the 4 AM Pick-ID reset
  keepDays:    30,
  // The tabs that cannot be rebuilt, plus the live work tabs (worth a nightly snapshot).
  // Master Inventory / Zoho Stock / Out of Stock / Kit Health / Price Audit are left out
  // on purpose: the system rebuilds them, and MI alone is ~780,000 cells.
  sheets: ["All orders", "Activity Log", "Order Archive", "KPI History", "Investigations",
           "Price Push Log", "Kit Registry", "Supplies", "Prep Queue", "Location Update",
           "Pending Sales Orders"]
};


function _bkFolder(createIfMissing) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(BACKUP.propFolder);
  if (id) { try { return DriveApp.getFolderById(id); } catch (_) {} }
  if (!createIfMissing) return null;
  var f = DriveApp.createFolder(BACKUP.folderName);
  props.setProperty(BACKUP.propFolder, f.getId());
  return f;
}

function _bkAlert(text) {
  try {
    if (typeof TELEGRAM_ADMIN_CHAT_ID === "undefined" || !TELEGRAM_ADMIN_CHAT_ID) return;
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage", {
      method: "post", contentType: "application/json", muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: TELEGRAM_ADMIN_CHAT_ID, text: text })
    });
  } catch (e) { try { console.log("_bkAlert: " + e); } catch (_) {} }
}

/** Pure: which backup files are past the keep window. files: [{name, created(ms)}]. */
function _bkExpired(files, nowMs, keepDays) {
  var cut = nowMs - keepDays * 86400000;
  return files.filter(function (f) { return /^HQ Backup · /.test(f.name) && f.created < cut; })
              .map(function (f) { return f.name; });
}


/**
 * Make tonight's backup. Safe to run any time ("Back up now"). Returns a summary.
 */
function runNightlyBackup(e) {
  // From the sidebar ("Back up now") the caller is a person — owner only, or a staff
  // click would create the folder and files in THEIR Drive. The 3 AM trigger passes an
  // event object and runs as the owner who installed it.
  if (!(e && e.triggerUid) && typeof _obIsOwner === "function" && !_obIsOwner()) {
    return { ok: false, message: "Only the owner account can run a backup — it saves into the owner's Drive." };
  }
  var started = Date.now();
  var stamp = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd");
  var res = { at: started, ok: false, file: "", url: "", sheets: 0, message: "" };
  try {
    var src = SpreadsheetApp.openById(SPREADSHEET_ID);
    var folder = _bkFolder(true);

    // Same-day reruns get "(2)", "(3)" — never overwrite a backup.
    var name = "HQ Backup · " + stamp, n = 2;
    while (folder.getFilesByName(name).hasNext()) name = "HQ Backup · " + stamp + " (" + (n++) + ")";

    var dest = SpreadsheetApp.create(name);
    var file = DriveApp.getFileById(dest.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (_) {}

    var done = [], missing = [], failed = [];
    BACKUP.sheets.forEach(function (tab) {
      var sh = src.getSheetByName(tab);
      if (!sh) { missing.push(tab); return; }
      try {
        var copy = sh.copyTo(dest);
        copy.setName(tab);
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        if (lr > 0 && lc > 0) {
          // Values over formulas — see the header. getDisplayValues would freeze dates as
          // text; getValues keeps real dates and numbers.
          copy.getRange(1, 1, lr, lc).setValues(sh.getRange(1, 1, lr, lc).getValues());
        }
        if (sh.isSheetHidden()) copy.showSheet();
        done.push(tab + " (" + lr + " rows)");
      } catch (e) {
        failed.push(tab + ": " + (e.message || e));
      }
    });

    // README first, the blank default tab removed.
    var readme = dest.getSheets()[0];
    readme.setName("README");
    readme.getRange(1, 1, 6, 1).setValues([
      ["HQ nightly backup — " + Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd h:mm a") + " Houston"],
      ["Copied: " + done.join(" · ")],
      [missing.length ? "Not found (skipped): " + missing.join(", ") : "All listed tabs were found."],
      ["Values only — formulas were replaced with what they showed."],
      ["RESTORE ONE TAB: right-click its tab here → Copy to → Existing spreadsheet → All Orders. Then move the rows you need. Never paste over a live tab without looking first."],
      ["Older backups are trashed after " + BACKUP.keepDays + " days (Drive keeps trash 30 more days)."]
    ]);
    readme.setColumnWidth(1, 900);
    dest.setActiveSheet(readme); dest.moveActiveSheet(1);

    // Retention.
    var it = folder.getFiles(), files = [];
    while (it.hasNext()) { var f = it.next(); files.push({ name: f.getName(), created: f.getDateCreated().getTime(), ref: f }); }
    var old = _bkExpired(files, Date.now(), BACKUP.keepDays);
    files.forEach(function (f) { if (old.indexOf(f.name) > -1) { try { f.ref.setTrashed(true); } catch (_) {} } });

    res.ok = failed.length === 0;
    res.file = name; res.url = dest.getUrl(); res.sheets = done.length;
    res.message = (res.ok ? "Backed up " : "Backed up with problems: ") + done.length + " tabs in "
                + Math.round((Date.now() - started) / 1000) + "s"
                + (failed.length ? " · FAILED " + failed.join(" | ") : "")
                + (missing.length ? " · not found " + missing.join(", ") : "")
                + (old.length ? " · trashed " + old.length + " old" : "");
    if (!res.ok) _bkAlert("⚠ HQ nightly backup had problems\n\n" + res.message + "\n\n" + res.url);
  } catch (err) {
    res.ok = false;
    res.message = "Backup FAILED: " + (err.message || err);
    _bkAlert("⚠ HQ nightly backup FAILED\n\n" + res.message + "\n\nNothing was saved tonight. The next run is tomorrow ~3 AM.");
  }
  try { PropertiesService.getScriptProperties().setProperty(BACKUP.propLast, JSON.stringify(res)); } catch (_) {}
  console.log(res.message);
  return res;
}


/** Install the daily trigger (owner — triggers run as whoever installs them). */
function setupNightlyBackup() {
  if (typeof _obIsOwner === "function" && !_obIsOwner()) {
    return { ok: false, message: "Only the owner account can switch on the nightly backup (it runs as you, into your Drive)." };
  }
  removeNightlyBackup();
  // ⚠ .inTimezone is load-bearing: the SCRIPT runs in Asia/Amman, so a bare atHour(3)
  //   would fire at 3 AM Amman = 7 PM Houston, mid-evening.
  ScriptApp.newTrigger("runNightlyBackup").timeBased().everyDays(1)
    .atHour(BACKUP.hour).inTimezone("America/Chicago").create();
  return { ok: true, message: "Nightly backup is ON — every night around " + BACKUP.hour + " AM Houston." };
}

function removeNightlyBackup() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "runNightlyBackup") { ScriptApp.deleteTrigger(t); n++; }
  });
  return { ok: true, removed: n, message: n ? "Nightly backup switched off." : "It was not on." };
}

/** For the sidebar card. */
function getBackupStatus() {
  var last = null;
  try { last = JSON.parse(PropertiesService.getScriptProperties().getProperty(BACKUP.propLast) || "null"); } catch (_) {}
  var armed = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "runNightlyBackup"; });
  var folder = _bkFolder(false);
  return { armed: armed, last: last, folderUrl: folder ? folder.getUrl() : "" , sheets: BACKUP.sheets };
}
