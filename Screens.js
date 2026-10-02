/**
 * Screens.js — "is a floor screen actually on?" (2026-10-02)
 *
 * Every board and wall now has a NAME and a ROLE (Floor / Wall / Office), set once
 * on the device from the ⋯ menu. They send it with every request, and n8n's
 * hq-board workflow remembers when it last heard from each one — n8n, because
 * board polls are answered from n8n's cache and never reach Apps Script. Asking
 * n8n `{action:"boardDevices"}` returns that list without touching Apps Script.
 *
 * THE PROBLEM THIS SOLVES: on 2026-10-02 at 08:30 Houston, in working hours, no
 * tablet was polling the board at all. A screen that is asleep or closed looks
 * EXACTLY like a quiet day — and a new order nobody sees is the failure the whole
 * board exists to prevent.
 *
 * RULES
 *   • Working hours only (Mon–Fri 9–17 Houston, the board's own definition), and
 *     never in the first `graceMin` of the shift — the tablet is being switched on.
 *   • Silent until at least one FLOOR-role screen has ever checked in. Before the
 *     devices are named there is nothing to watch, and an alarm about that would
 *     be noise.
 *   • Alert ONCE when the newest floor check-in passes `quietMin`, and once more
 *     when one comes back. Recorded only after a successful send, so a failed send
 *     retries instead of being swallowed — the same rule as the pulse alarm.
 *   • Off-hours resets the state silently, so tomorrow morning starts fresh.
 */

var SCREENS = {
  quietMin:      15,   // no floor check-in for this long = alert
  graceMin:      15,   // ignore the first minutes of the shift
  everyMin:      5,    // how often the check runs (rides the 1-min publish trigger)
  stateKey:      "SCREENS_STATE",       // "ok" | "quiet"
  lastRunKey:    "SCREENS_LAST_RUN",
  startHour:     9,
  endHour:       17
};


function checkFloorScreens() {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();

  var last = Number(props.getProperty(SCREENS.lastRunKey) || 0);
  if (now - last < SCREENS.everyMin * 60000 - 5000) return "skip";
  props.setProperty(SCREENS.lastRunKey, String(now));

  var hNow = Utilities.formatDate(new Date(now), "America/Chicago", "u H m").split(" ");
  var dow = Number(hNow[0]), hour = Number(hNow[1]), minute = Number(hNow[2]);
  var working = dow >= 1 && dow <= 5 && hour >= SCREENS.startHour && hour < SCREENS.endHour;
  if (!working) {
    props.setProperty(SCREENS.stateKey, "ok");
    return "off-hours";
  }
  var minsIntoShift = (hour - SCREENS.startHour) * 60 + minute;
  if (minsIntoShift < SCREENS.graceMin) return "shift starting";

  var list = _screensFetch();
  if (!list) return "no answer from n8n";          // the pulse alarm covers a dead endpoint
  var floor = list.filter(function (d) { return d.role === "floor"; });
  if (!floor.length) return "no floor screen named yet";

  floor.sort(function (a, b) { return a.ageSec - b.ageSec; });
  var freshest = floor[0];
  var state = freshest.ageSec > SCREENS.quietMin * 60 ? "quiet" : "ok";
  var prev = props.getProperty(SCREENS.stateKey) || "ok";
  if (state === prev) return state + " (" + freshest.name + " " + Math.round(freshest.ageSec / 60) + "m)";

  var msg = state === "quiet"
    ? "📵 NO FLOOR SCREEN IS ON\n\n" +
      "No floor screen has checked in for " + _screensAge(freshest.ageSec) + ".\n" +
      "Last seen: " + floor.map(function (d) { return d.name + " (" + _screensAge(d.ageSec) + " ago)"; }).join(", ") + ".\n\n" +
      "A new order will land with nobody seeing it. Wake the tablet, or open the board on it."
    : "✅ Floor screen back: " + freshest.name + " is checking in again.";

  var sent = _tgSend(TELEGRAM_ADMIN_CHAT_ID, msg);
  if (sent) props.setProperty(SCREENS.stateKey, state);
  return prev + " → " + state + (sent ? " (alerted)" : " (SEND FAILED, will retry)");
}


/** Ask n8n which screens it has heard from. Returns null when it cannot say. */
function _screensFetch() {
  try {
    var r = UrlFetchApp.fetch(HQ_BOARD_API_URL, {
      method: "post", contentType: "application/json", muteHttpExceptions: true,
      payload: JSON.stringify({ action: "boardDevices" })
    });
    if (r.getResponseCode() !== 200) return null;
    var j = JSON.parse(r.getContentText());
    return (j && Array.isArray(j.devices)) ? j.devices : null;
  } catch (e) {
    console.log("_screensFetch: " + e);
    return null;
  }
}

function _screensAge(sec) {
  var m = Math.round(sec / 60);
  return m < 60 ? m + " min" : Math.floor(m / 60) + "h " + (m % 60) + "m";
}

/** Editor helper — what n8n currently knows about the screens. */
function showScreensNow() {
  var list = _screensFetch();
  console.log(list ? JSON.stringify(list, null, 1) : "n8n did not return a device list");
  return list;
}
