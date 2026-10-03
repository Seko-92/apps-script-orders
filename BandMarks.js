/**
 * BandMarks.js — the animated marks on the DIRECT / AMAZON table bands (2026-10-03).
 *
 * Owner-approved on the mock "Table band · round 3" — "this is an ART". Two GIFs, drawn and
 * verified in dial/band-marks/ (marks.html → cap.js → enc.py):
 *   band-direct-v2.gif  the hand truck tips back, rolls, the box settles, a warm light glides
 *                       over the word, then a small box TRUNDLES across the empty stretch on
 *                       its own shadow and drops off — hand-carried, unhurried
 *   band-amazon-v2.gif  OUR HQ parcel in the lane right of Amazon's still logo: a scan line
 *                       sweeps it, a label stamps on, it GLIDES out with a soft trail, and a
 *                       fresh one drops in. ⚠ Amazon's mark is never inside our image.
 * Both loops are 22 s and Amazon moves 11 s after Direct, so the bands never move together.
 *
 * ⭐ HOW IT STAYS SAFE
 *   · A floating image is tied to its anchor cell, so it rides down with the band when orders
 *     land above it — MEASURED 2026-10-03 (BandGifProbe: inserted rows, open tab, no reload).
 *   · The GIFs are OPAQUE on the band's exact #ffd400 and sit exactly over the still =IMAGE()
 *     mark, which stays underneath as the fallback. removeBandMarks() is the whole rollback.
 *   · They cover only band cells, which nobody types in, so they block no real click.
 * ⚠ A newly inserted image does not appear in an already-open tab until it is reloaded
 *   (2026-09-02). Hard-reload after installBandMarks().
 * ⚠ The Amazon parcel goes on only when the AMAZON table exists. Re-run installBandMarks()
 *   after setupAmazonTable().
 */
var BAND_MARKS = {
  family: '/mast/band-',                       // how our images are recognised (URL family)
  // v2 (round 7): the Direct GIF spans from the mark to the end of E (the delivery run);
  // the Amazon GIF is ONLY the lane to the right of Amazon's still logo.
  direct: { file: 'band-direct-v2.gif', w: 341, h: 50, markW: 140 },
  amazon: { file: 'band-amazon-v2.gif', w: 197, h: 50, markW: 99 }
};

/** Our band images on a sheet, by URL family — never by position. */
function _bandMarkImages(sheet) {
  return sheet.getImages().filter(function (im) {
    var u = ''; try { u = String(im.getUrl() || ''); } catch (e) {}
    return u.indexOf(BAND_MARKS.family) !== -1;
  });
}

/** Owner-run. Idempotent: removes our old band marks, then places one per band that exists. */
function installBandMarks() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ All Orders not found';
  _bandMarkImages(sheet).forEach(function (im) { im.remove(); });
  var L = getTableLayout(sheet);
  var span = sheet.getColumnWidth(Schema.bandLogoCol) + sheet.getColumnWidth(Schema.bandLogoCol + 1);   // D:E
  var out = [];
  function place(row, spec, offX, label) {
    var rh = sheet.getRowHeight(row);
    var offY = Math.max(0, Math.round((rh - spec.h) / 2));
    var img = sheet.insertImage(MASTHEAD.baseUrl + spec.file, Schema.bandLogoCol, row, Math.max(0, Math.round(offX)), offY);
    img.setWidth(spec.w).setHeight(spec.h);
    out.push('✓ ' + label + ' on row ' + row + ' (anchor ' + img.getAnchorCell().getA1Notation() + ', offset ' + Math.round(offX) + ',' + offY + ')');
  }
  if (L.direct > 0) {
    // the GIF's mark starts 12px in, so this lands it exactly on the centred still mark
    // the GIF's mark starts 12px in → x = the centred still mark's left - 12
    place(L.direct, BAND_MARKS.direct, (span - BAND_MARKS.direct.markW) / 2 - 12, 'DIRECT truck + delivery run');
  } else out.push('✗ DIRECT band not found');
  if (L.amazon > 0) {
    // the lane starts 12px right of Amazon's centred still logo — the logo is never covered
    place(L.amazon, BAND_MARKS.amazon, (span + BAND_MARKS.amazon.markW) / 2 + 12, 'AMAZON parcel lane');
  } else out.push('· AMAZON table not on the sheet — the parcel goes on when it is (re-run this then)');
  var msg = out.join('\n') + '\n⚠ Hard-reload the sheet tab to see a newly inserted image. Rollback: removeBandMarks()';
  console.log(msg);
  return msg;
}

/** The whole rollback: our band GIFs off; the still marks underneath were never touched. */
function removeBandMarks() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return '❌ All Orders not found';
  var imgs = _bandMarkImages(sheet);
  imgs.forEach(function (im) { im.remove(); });
  return '✅ Removed ' + imgs.length + ' band mark(s). The still marks are untouched.';
}

/** Read-only: what band images the sheet really holds (anchor, offset, size, URL). */
function describeBandMarks() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var L = getTableLayout(sheet);
  var lines = ['bands: DIRECT row ' + L.direct + ' · AMAZON row ' + (L.amazon || 'none')];
  sheet.getImages().forEach(function (im) {
    var u = ''; try { u = String(im.getUrl() || ''); } catch (e) {}
    lines.push(im.getAnchorCell().getA1Notation() + ' +' + im.getAnchorCellXOffset() + ',' + im.getAnchorCellYOffset() +
               ' · ' + im.getWidth() + 'x' + im.getHeight() + ' · ' + (u ? u.split('/').pop() : '(no url)'));
  });
  var msg = lines.join('\n'); console.log(msg); return msg;
}
