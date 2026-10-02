/**
 * BandGifProbe.js — one question before the animated band marks are built (2026-10-03).
 *
 * The DIRECT/AMAZON band moves down every time an order lands above it. A floating GIF
 * over the band is tied to its anchor cell, so it should move with the band. What we have
 * NOT measured: whether an ALREADY-OPEN tab redraws a moved image without a reload. A newly
 * inserted image did not (2026-09-02, row 1), and only an eye can tell — getImages() reports
 * the anchor either way.
 *
 * Run, in order, from the editor, with the sheet open in another tab:
 *   1. probeBandGif()        → a scratch tab "__BandGifProbe" with a yellow band on row 10 and a
 *                              small moving GIF over it. Look at the tab: is the GIF on the band?
 *                              (If not, reload once — that is the known new-image repaint.)
 *   2. probeBandGifShift()   → inserts 3 rows ABOVE the band, the way an arrival does. Do NOT
 *                              reload. Watch: does the GIF move down with the band to row 13?
 *   3. removeBandGifProbe()  → deletes the scratch tab. All Orders is never touched.
 */
var BAND_GIF_PROBE = {
  sheet: '__BandGifProbe',
  band:  10,
  url:   'https://hq.yassinqurabi.com/mast/probe-f1.gif'
};

function probeBandGif() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var old = ss.getSheetByName(BAND_GIF_PROBE.sheet);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(BAND_GIF_PROBE.sheet);
  var r = BAND_GIF_PROBE.band;
  sh.getRange(r, 1, 1, 8).setBackground('#ffd400');
  sh.getRange(r, 1).setValue('BAND').setFontWeight('bold');
  sh.setRowHeight(r, 60);
  for (var i = 1; i <= 8; i++) sh.setColumnWidth(i, 110);
  sh.getRange(r + 1, 1, 5, 1).setValues([['order below 1'], ['order below 2'], ['order below 3'], ['order below 4'], ['order below 5']]);
  var img = sh.insertImage(BAND_GIF_PROBE.url, 4, r, 20, 10);
  img.setWidth(40).setHeight(40);
  ss.setActiveSheet(sh);
  var msg = 'Probe up. GIF anchored at ' + img.getAnchorCell().getA1Notation() +
            ' (band = row ' + r + '). Look at the "' + BAND_GIF_PROBE.sheet + '" tab: the moving dot should sit on the yellow band. ' +
            'Then run probeBandGifShift() WITHOUT reloading.';
  console.log(msg);
  return msg;
}

function probeBandGifShift() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BAND_GIF_PROBE.sheet);
  if (!sh) return 'Run probeBandGif() first.';
  sh.insertRowsBefore(5, 3);
  sh.getRange(5, 1, 3, 1).setValues([['new arrival 1'], ['new arrival 2'], ['new arrival 3']]);
  var imgs = sh.getImages();
  var at = imgs.length ? imgs[0].getAnchorCell().getA1Notation() : '(no image)';
  var bandNow = BAND_GIF_PROBE.band + 3;
  var msg = '3 rows inserted above. The band is now row ' + bandNow + '; the GIF reports its anchor at ' + at + '. ' +
            'Without reloading: is the moving dot still on the yellow band (YES), or left behind over the "new arrival" rows (NO)? ' +
            'Then run removeBandGifProbe().';
  console.log(msg);
  return msg;
}

function removeBandGifProbe() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(BAND_GIF_PROBE.sheet);
  if (!sh) return 'Nothing to remove.';
  ss.deleteSheet(sh);
  return 'Probe tab removed.';
}
