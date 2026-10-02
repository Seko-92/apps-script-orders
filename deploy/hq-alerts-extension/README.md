# HQ Alerts — Chrome extension

The `/alerts` tab without the tab. Chrome checks the board every 30 seconds in the
background and shows the same notifications the tab shows: new orders during Houston
work hours, unanswered holds at any time. A badge on the toolbar icon shows the state.
**Read-only. It never changes the sheet.**

| Badge | Means |
|---|---|
| none | all clear |
| yellow number | new orders since you last opened the popup |
| red number | unanswered holds (stays until someone acknowledges on the board) |
| grey `!` | cannot reach the server |

## Install on a PC (developer mode, about 2 minutes)

1. Download **https://hq.yassinqurabi.com/hq-alerts-extension.zip** and unzip it to a
   folder that will stay put, e.g. `C:\HQ\hq-alerts-extension`. Chrome loads it from
   there, so don't delete or move that folder.
2. In Chrome open `chrome://extensions`, turn on **Developer mode** (top right).
3. Click **Load unpacked** and choose the unzipped folder (the one with `manifest.json`).
4. Click the puzzle-piece icon in the toolbar and **pin** HQ Alerts.
5. Click the HQ icon → **Test**. A notification should appear in the corner of the screen.
   If it doesn't: Windows Settings → System → Notifications → make sure **Google Chrome**
   is on, and Focus Assist / Do Not Disturb is off.
6. Chrome Settings → System → turn on **Continue running background apps when Google
   Chrome is closed**, so alerts keep coming after the last window is closed.
7. **Close the `/alerts` tab on this PC.** With both running you'd get every alert twice.

Chrome shows a "developer mode extensions" warning at start-up. That is expected for
this install method; click to dismiss it.

## Update

Download the new zip, unzip it over the same folder, then `chrome://extensions` →
the reload ↻ icon on HQ Alerts.

## Remove

`chrome://extensions` → HQ Alerts → **Remove**.

## For whoever maintains it

* The rules live in `core.js`. **It is the same file the `/alerts` tab loads**, deployed
  to the VPS as `/opt/hq-app/alerts-core.js`. Change it once, then deploy both:
  `scp core.js root@10.8.0.1:/opt/hq-app/alerts-core.js` and rebuild the zip.
* `background.js` is plumbing only: the 30 s alarm, memory in `chrome.storage.local`
  (a service worker forgets everything between wakes), notifications, badge.
* Tests: `design-lab/test-alerts-extension.js` (the real worker against stubbed Chrome
  APIs) and `design-lab/probe-alerts-extension.js` (loads it into real Chromium against
  the live board). `design-lab/test-alerts.js` covers the tab with the same scenarios.
* Bump `version` in `manifest.json` on every change so PCs can tell which one they run.
