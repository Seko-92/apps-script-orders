/**
 * probe-floating-image.gs — PASTE INTO THE APPS SCRIPT EDITOR, RUN, LOOK, DELETE.
 *
 * ⚠ NOT part of the project. It lives in dial/ (a subdirectory, so clasp never pushes it)
 *   precisely so it cannot become permanent by accident. Paste into a scratch file in the
 *   editor, run, answer the questions, then DELETE the file.
 *
 * ⚠⚠ v2 — 2026-09-02. THE v1 PROBE RAN CLEAN AND DREW NOTHING, AND THE REASON MATTERS.
 *
 *    v1 inlined its art as `data:image/...;base64,...` URLs. Both probes returned with no
 *    exception, `setWidth()` succeeded on the returned object, the log said "inserted" —
 *    and the sheet stayed empty. **insertImage() accepts a data URL and silently fails to
 *    render it.** Success status, failure content: the same shape as Zoho's 200-with-an-
 *    error-body, the /exec 404 that answered HTTP 200, and the masthead face that 200s
 *    with the board's HTML.
 *
 *    ⭐ WHERE THE WRONG IDEA CAME FROM. CLAUDE.md states "a Blob throws; a public URL or
 *      DATA URL works." The Blob half was measured (2026-08-30, "blob format is
 *      unsupported"). The data-URL half never was — every insertImage in this project that
 *      has ever rendered used an http(s) URL (GIF_PROBE.url, MASTHEAD.baseUrl). The note
 *      was an assumption wearing a measurement's clothes, and this file believed it.
 *      **A note here is evidence, not proof.**
 *
 *    v2 serves both assets over HTTPS from the directory that already serves the masthead
 *    faces, byte-verified on the server. No Caddy edit — /mast/ is an existing route.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * WHAT TO DO — run them in this order.
 *
 *   0 · diagnoseFloatingImages()   READ-ONLY. Lists every floating image on the sheet with
 *       its anchor, size and source URL. Run it FIRST: v1 reported "removed 1 probe
 *       image(s)" before it had inserted anything, so something was already there and we
 *       do not yet know what. This answers that without changing a thing.
 *
 *   1 · probeFloatingDial()        THE ANCHOR QUESTION.
 *       a · is it MOVING?                                     -> expected yes
 *       b · SCROLL DOWN one full page. Still on the banner, or slid away?  <- THE QUESTION
 *       c · reload the sheet. Still there, still moving?
 *
 *   2 · probeAlphaPassthrough()    THE ALPHA QUESTION.
 *       A magenta frame lands over D1. Its middle is a HOLE; its right third carries a
 *       white swatch, a black swatch and a 50% block.
 *       · D1's headline readable through the hole -> ALPHA PASSES -> the STATIC FRAME
 *         shape is live: zero recurring cost, and the state-drift trap disappears
 *       · flat WHITE (the white swatch vanishes into it)  -> composited on white
 *       · flat BLACK with no text (black swatch vanishes)  -> composited on black
 *       Either composite answer kills the static frame and leaves the animated board.
 *       The 50% block is a free third answer: a real blend means 8-bit alpha, solid-or-
 *       gone means Sheets treats alpha as binary.
 *
 *   3 · removeFloatingDialProbe()  Clears both. DO NOT SKIP.
 *
 * ⚠ It writes NO cells and touches NO formulas. The only thing it changes is the sheet's
 *   floating-image list, and step 3 empties that of anything this probe made.
 */

// Served from /opt/hq-app/mast on the VPS — the SAME directory as the 440 masthead faces,
// so no Caddy route was added and the documented blast radius is untouched.
// Both byte-verified on the server: 359 b PNG (RGBA, 232x56) and 16,389 b GIF.
// ⚠ The ?v= is a cache-buster. Google fetches these server-side and will hold a copy;
//   bump it if you ever replace the art at the same filename.
var PROBE_GIF_URL = 'https://hq.yassinqurabi.com/mast/anchor-probe-v1.gif?v=1';
var PROBE_PNG_URL = 'https://hq.yassinqurabi.com/mast/alpha-probe-v1.png?v=1';

// The two widths this probe owns. WIDTH IS THE DISCRIMINATOR FOR REMOVAL, never the anchor:
// the brand logo also floats at A1 (setupBrandLogo -> insertImage(blob, 1, 1)) and is capped
// at colW-4 = 103px, so it can never be 196 or 232. Matching row+column alone would delete
// it — which is a live defect in BrandTheme.removeMastheadAnimated(), whose comment claims
// the very protection its code omits. Do not copy that pattern.
var PROBE_W_ANCHOR = 196;
var PROBE_W_ALPHA  = 232;


/** READ-ONLY. Changes nothing. Answers "what is actually floating on this sheet?" */
function diagnoseFloatingImages() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  var imgs = sheet.getImages();
  var out = ['Floating images on "' + MAIN_SHEET_NAME + '": ' + imgs.length, ''];
  for (var i = 0; i < imgs.length; i++) {
    var im = imgs[i], anchor = '?', url = '?';
    try { anchor = im.getAnchorCell().getA1Notation(); } catch (e) { anchor = '(unreadable)'; }
    // getUrl() returns null for blob-inserted images — that is itself the tell for the
    // brand logo, which setupBrandLogo inserts from a Drive blob.
    try { url = im.getUrl() || '(no url — inserted from a blob, e.g. the brand logo)'; }
    catch (e) { url = '(no url)'; }
    var w = Math.round(im.getWidth()), h = Math.round(im.getHeight());
    var mine = (w === PROBE_W_ANCHOR || w === PROBE_W_ALPHA) ? '  <- MATCHES THIS PROBE' : '';
    out.push((i + 1) + ' · anchor ' + anchor + '  ' + w + 'x' + h +
             '  offset ' + im.getAnchorCellXOffset() + ',' + im.getAnchorCellYOffset() + mine);
    out.push('     ' + url);
  }
  if (!imgs.length) out.push('(none — nothing is floating over this sheet)');
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}


function probeFloatingDial() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  _removeProbeWidth(sheet, PROBE_W_ANCHOR);          // idempotent, and only ours
  var img = sheet.insertImage(PROBE_GIF_URL, 1, 1, 6, 4);   // anchored A1, inside row 1
  img.setWidth(PROBE_W_ANCHOR).setHeight(85);
  SpreadsheetApp.flush();
  var msg =
    'Anchor probe inserted over A1 (' + PROBE_W_ANCHOR + 'x85).\n\n' +
    '1 · is it MOVING?\n' +
    '2 · SCROLL DOWN one page — does it stay on the banner, or slide away?\n' +
    '3 · reload the sheet — still there, still moving?\n\n' +
    'If you see NOTHING, run diagnoseFloatingImages() — if it lists an image at A1\n' +
    'with this URL, then it was placed and something is drawing on top of it.\n\n' +
    'Then run removeFloatingDialProbe().';
  Logger.log(msg);
  return msg;
}


function probeAlphaPassthrough() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  _removeProbeWidth(sheet, PROBE_W_ALPHA);           // re-runnable; leaves the anchor probe
  // ⚠ D1 is the target on purpose: it is the surface the board would actually cover, it is
  //   the dark band the design sits on, and it holds LIVE TEXT — the only thing that can
  //   separate "alpha passed through" from "composited onto a black ground". A blank or
  //   white cell would leave that ambiguous.
  var img = sheet.insertImage(PROBE_PNG_URL, Schema.cols.SALES_ORDER, 1, 0, 0);
  img.setWidth(PROBE_W_ALPHA).setHeight(56);
  SpreadsheetApp.flush();
  var msg =
    'Alpha probe placed over D1 (' + PROBE_W_ALPHA + 'x56).\n\n' +
    'Look INSIDE the magenta frame:\n' +
    '  · can you READ D1\'s headline through it?  -> alpha PASSES THROUGH\n' +
    '  · flat WHITE (white swatch disappears)     -> composited on white\n' +
    '  · flat BLACK, no text (black swatch gone)  -> composited on black\n\n' +
    'The 50% grey block, lower right: a real blend means 8-bit alpha,\n' +
    'solid-or-gone means Sheets is treating alpha as binary.\n\n' +
    'Then run removeFloatingDialProbe().';
  Logger.log(msg);
  return msg;
}


/** Removes ONLY row-1 images of one exact width. The narrowness is the safety story. */
function _removeProbeWidth(sheet, width) {
  var imgs = sheet.getImages(), gone = 0;
  for (var i = 0; i < imgs.length; i++) {
    try {
      if (imgs[i].getAnchorCell().getRow() === 1 &&
          Math.round(imgs[i].getWidth()) === width) { imgs[i].remove(); gone++; }
    } catch (e) { /* unreadable anchor — not ours, leave it alone */ }
  }
  return gone;
}


function removeFloatingDialProbe() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  var gone = _removeProbeWidth(sheet, PROBE_W_ANCHOR) + _removeProbeWidth(sheet, PROBE_W_ALPHA);
  SpreadsheetApp.flush();
  var msg = 'removed ' + gone + ' probe image(s)';
  Logger.log(msg);
  return msg;
}
