/**
 * MiAudit.js — Master Inventory pre-flight audit (Phase 0 of the parcel-data plan)
 *
 * ⚠ TEMPORARY DIAGNOSTIC. Delete this file once it has answered its questions.
 *   House ruling (2026-08-21): "a diagnostic that has answered its question is
 *   just a card taking up room."
 *
 * WHY IT EXISTS
 *   Before appending 17 parcel/shipping/health columns to Master Inventory we need
 *   four facts that nothing currently reports:
 *
 *   1. DOES THE SUNDAY FULL SYNC ACTUALLY COMPLETE?  It is the backfill mechanism
 *      for every new column. If it dies partway through the Trading quota, a large
 *      slice of MI never gets weight — and we would ship 17 permanently-empty
 *      columns without knowing.
 *   2. WHERE DOES THE SHEET ACTUALLY END?  The append position for the new headers
 *      must be derived, not assumed. Two n8n Sheets nodes currently disagree with
 *      the sheet about its own width (cached schemas hold 195 and 189 entries).
 *   3. IS THE `error` COLUMN A HIGH-WATER MARK?  Both parsers omit `error` on the
 *      success path and autoMapInputData only writes keys that are present, so the
 *      column is never cleared. If most rows carry a stale failure, "is the sync
 *      healthy" is currently unanswerable.
 *   4. HOW STALE ARE THE PHOTOS?  MAIN (hourly) writes picture1..5; SUB (Sunday)
 *      writes pictureUrl1..5. Four consumers read ONLY pictureUrl* — PhotoQueue,
 *      LowStock, PartConsole, and the orders workflow's Telegram card. This
 *      measures the resulting lag instead of arguing about it.
 *
 * HOW TO RUN
 *   Apps Script editor → function dropdown → auditMasterInventoryFreshness → Run.
 *   Output goes to the EXECUTION LOG (View → Executions), not the return value:
 *   ⚠ the Run button does not display return values — a lesson that has cost this
 *   project an evening twice (getPublishedTick, auditBoardStockAdjustments).
 *
 * ⚠ ZERO-ARG BY DESIGN. The Run button cannot pass arguments; this project has
 *   walked into that three times.
 *
 * READ-ONLY. Writes nothing, anywhere.
 */

/** How many distinct lastUpdated day-buckets to print, newest first. */
var MI_AUDIT_TOP_DAYS = 14;

/** Day key of the biggest lastUpdated bucket — set by the freshness pass, read by the photo pass. */
var _miaLastFullDay = '';


function auditMasterInventoryFreshness() {
  var out = [];
  function say(line) { out.push(line); try { console.log(line); } catch (e) {} }

  say('════════ MASTER INVENTORY AUDIT — ' + _miaNow() + ' ════════');

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mi = ss.getSheetByName(DB_SHEET_NAME);
  if (!mi) { say('❌ Sheet "' + DB_SHEET_NAME + '" not found.'); return out.join('\n'); }

  var lastRow = mi.getLastRow();
  var lastCol = mi.getLastColumn();
  var nRows   = Math.max(0, lastRow - 1);
  if (nRows < 1) { say('❌ No data rows.'); return out.join('\n'); }

  // ── 2. GEOMETRY — the append position for Phase 2 ───────────────────────────
  var headers = mi.getRange(1, 1, 1, lastCol).getValues()[0];
  var namedCols = 0, blankTail = 0;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim() !== '') { namedCols = h + 1; }
  }
  blankTail = lastCol - namedCols;

  say('');
  say('── GEOMETRY ─────────────────────────────────────────────');
  say('  data rows          : ' + nRows);
  say('  getLastColumn()    : ' + lastCol + '  (' + _miaColLetter(lastCol) + ')');
  say('  last NAMED header  : ' + namedCols + '  (' + _miaColLetter(namedCols) + ')  "' + headers[namedCols - 1] + '"');
  if (blankTail > 0) say('  ⚠ ' + blankTail + ' trailing column(s) exist but have no header.');
  say('  ⭐ APPEND THE 17 NEW HEADERS AT COLUMNS ' +
      (namedCols + 1) + '–' + (namedCols + 17) +
      '  (' + _miaColLetter(namedCols + 1) + '..' + _miaColLetter(namedCols + 17) + ')');

  // The one hardcoded MI reader in the codebase. Prove it is still correct
  // BEFORE we touch the sheet — appending is only safe while these four hold.
  say('');
  say('── KitRegistry.js:1118 GUARD (reads cols 2..40 by POSITION) ──');
  var guard = [
    { col: 2,  want: DB_SKU_HEADER,      label: 'sku'          },
    { col: 3,  want: 'title',            label: 'title'        },
    { col: 6,  want: 'quantity',         label: 'quantity'     },
    { col: 40, want: 'C:Model Year',     label: 'aisle'        }
  ];
  var guardOk = true;
  for (var g = 0; g < guard.length; g++) {
    var actual = String(headers[guard[g].col - 1] || '').trim();
    var ok = actual === String(guard[g].want).trim();
    if (!ok) guardOk = false;
    say('  ' + (ok ? '✅' : '❌') + ' col ' + guard[g].col + ' (' + _miaColLetter(guard[g].col) + ') ' +
        guard[g].label + ' — want "' + guard[g].want + '", got "' + actual + '"');
  }
  say(guardOk
    ? '  ✅ SAFE TO APPEND. Never insert a column at or before AN.'
    : '  ❌ ALREADY DRIFTED — kit aisles are reading the wrong field. STOP and fix this first.');

  // ── column resolver (header name, first match, trimmed+lowercased) ──────────
  function colIdx(name) {
    var t = String(name).trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').trim().toLowerCase() === t) return i + 1;
    }
    return -1;
  }
  function readCol(name) {
    var c = colIdx(name);
    if (c < 1) return null;
    return mi.getRange(2, c, nRows, 1).getValues();
  }

  // ── 1. SUNDAY FULL SYNC — does it complete? ─────────────────────────────────
  say('');
  say('── FRESHNESS  (lastUpdated, bucketed in America/Chicago) ─');
  var luCol = colIdx('lastUpdated');
  if (luCol < 1) {
    say('  ❌ no "lastUpdated" header found.');
  } else {
    var lu = mi.getRange(2, luCol, nRows, 1).getValues();
    var buckets = {}, unreadable = 0, empty = 0;
    for (var i = 0; i < lu.length; i++) {
      var key = _miaDayKey(lu[i][0]);
      if (key === '')  { empty++; continue; }
      if (key === '?') { unreadable++; continue; }
      buckets[key] = (buckets[key] || 0) + 1;
    }
    var days = Object.keys(buckets).sort().reverse();
    say('  column           : ' + luCol + ' (' + _miaColLetter(luCol) + ')');
    say('  distinct days    : ' + days.length + '   empty: ' + empty + '   unreadable: ' + unreadable);
    say('');
    var shown = Math.min(days.length, MI_AUDIT_TOP_DAYS), covered = 0;
    for (var d = 0; d < shown; d++) {
      var k = days[d], n = buckets[k];
      covered += n;
      say('    ' + k + ' ' + _miaWeekday(k) + '  ' + _miaPad(n, 5) + '  ' +
          _miaPct(n, nRows) + '  ' + _miaBar(n, nRows));
    }
    if (days.length > shown) {
      var rest = nRows - covered - empty - unreadable;
      say('    … ' + (days.length - shown) + ' older day(s), ' + rest + ' row(s)');
    }
    say('');
    // The verdict: a healthy Sunday sync touches nearly every row on one Sunday.
    var bestDay = null, bestN = 0;
    for (var dd = 0; dd < days.length; dd++) {
      if (buckets[days[dd]] > bestN) { bestN = buckets[days[dd]]; bestDay = days[dd]; }
    }
    var bestPctNum = nRows ? (bestN / nRows) * 100 : 0;
    _miaLastFullDay = bestDay;   // window the photo check below is measuring across
    var ageDays = bestDay ? _miaDaysAgo(bestDay) : -1;
    say('  biggest single day : ' + bestDay + ' ' + _miaWeekday(bestDay) +
        ' — ' + bestN + ' rows (' + bestPctNum.toFixed(1) + '%), ' + ageDays + ' day(s) ago');
    if (bestPctNum >= 80 && ageDays <= 9) {
      say('  ✅ VERDICT: the full sync is completing. Backfill for new columns will work.');
    } else if (bestPctNum >= 80) {
      say('  ⚠ VERDICT: a full pass DID complete, but ' + ageDays + ' days ago.');
      say('    Confirm the Sunday schedule is still enabled before adding columns.');
    } else {
      say('  ❌ VERDICT: no single day covers most of the sheet — the full sync is NOT');
      say('    completing. FIX THIS BEFORE PHASE 2, or the 17 new columns stay empty');
      say('    for the rows the sync never reaches.');
    }
  }

  // ── 3. `error` — high-water mark, or current state? ─────────────────────────
  say('');
  say('── error COLUMN ─────────────────────────────────────────');
  var errVals = readCol('error');
  if (!errVals) {
    say('  (no "error" header found)');
  } else {
    var errCount = 0, errKinds = {};
    for (var e2 = 0; e2 < errVals.length; e2++) {
      var v = String(errVals[e2][0] || '').trim();
      if (!v) continue;
      errCount++;
      var kind = v.length > 60 ? v.slice(0, 60) + '…' : v;
      errKinds[kind] = (errKinds[kind] || 0) + 1;
    }
    say('  column ' + colIdx('error') + ' (' + _miaColLetter(colIdx('error')) + ')  non-empty: ' +
        errCount + ' / ' + nRows + '  (' + _miaPct(errCount, nRows) + ')');
    var kinds = Object.keys(errKinds).sort(function(a, b) { return errKinds[b] - errKinds[a]; });
    for (var kk = 0; kk < Math.min(kinds.length, 5); kk++) {
      say('    ' + _miaPad(errKinds[kinds[kk]], 5) + '  "' + kinds[kk] + '"');
    }
    if (errCount > nRows * 0.5) {
      say('  ⚠ Over half the sheet carries an error string. Because neither parser');
      say('    emits `error` on success, this is "failed at least once, ever" — NOT');
      say('    "currently broken". Phase 1 step 4 makes this column mean something.');
    }
  }

  // ── 4. PHOTO DIVERGENCE — the measured cost of picture1 vs pictureUrl1 ──────
  say('');
  say('── PHOTO KEY DIVERGENCE  (MAIN writes picture*, 4 readers want pictureUrl*) ─');
  // ⚠ FIRST VERSION OF THIS CHECK COMPARED ONLY pictureUrl1 vs picture1 AND
  // UNDERSTATED THE LAG. The Photo Queue does not care which image is first —
  // it counts NON-EMPTY pictureUrl1..5 and flags <= 1. So the transition that
  // matters is "logo only" -> "photographer uploaded 3 more", and on that row
  // image #1 is still the logo. A first-image comparison calls that identical.
  // Compare COUNTS, which is what the reader actually does.
  var puCols = [], pCols = [];
  for (var pi = 1; pi <= 5; pi++) {
    puCols.push(readCol('pictureUrl' + pi));
    pCols.push(readCol('picture' + pi));
  }
  if (puCols.indexOf(null) >= 0 || pCols.indexOf(null) >= 0) {
    say('  (need all of pictureUrl1..5 and picture1..5; at least one header is missing)');
  } else {
    function countAt(cols, r) {
      var n = 0;
      for (var c = 0; c < 5; c++) if (String(cols[c][r][0] || '').trim()) n++;
      return n;
    }
    var MAXI = (typeof PREP_PHOTO !== 'undefined' && PREP_PHOTO.maxImages) ? PREP_PHOTO.maxImages : 1;
    var agree = 0, mainAhead = 0, subAhead = 0, firstDiffers = 0;
    var wouldLeaveQueue = 0, wouldEnterQueue = 0, queueNow = 0, queueTrue = 0;
    for (var r2 = 0; r2 < nRows; r2++) {
      var nSub  = countAt(puCols, r2);   // what the 4 readers see  (SUB, weekly)
      var nMain = countAt(pCols,  r2);   // what eBay said most recently (MAIN, hourly)
      if (nSub  <= MAXI) queueNow++;     // currently flagged "needs photos"
      if (nMain <= MAXI) queueTrue++;    // actually still needs photos
      if (nMain === nSub) agree++;
      else if (nMain > nSub) { mainAhead++; if (nSub <= MAXI && nMain > MAXI) wouldLeaveQueue++; }
      else { subAhead++; if (nSub > MAXI && nMain <= MAXI) wouldEnterQueue++; }
      var a1 = String(puCols[0][r2][0] || '').trim();
      var b1 = String(pCols[0][r2][0]  || '').trim();
      if (a1 && b1 && a1 !== b1) firstDiffers++;
    }
    say('  image COUNTS agree              : ' + _miaPad(agree, 5) + '  (' + _miaPct(agree, nRows) + ')');
    say('  MAIN (hourly) has MORE images   : ' + _miaPad(mainAhead, 5) + '   photos added since the last full sync');
    say('  SUB  (weekly) has MORE images   : ' + _miaPad(subAhead, 5) + '   photos removed since');
    say('  first image URL differs         : ' + _miaPad(firstDiffers, 5));
    say('');
    say('  ── what this costs the Photo Queue (flags <= ' + MAXI + ' image) ──');
    say('  flagged today  (reads pictureUrl*) : ' + _miaPad(queueNow, 5));
    say('  actually needs a photo (picture*)  : ' + _miaPad(queueTrue, 5));
    say('  ⭐ FALSE "needs photos" — already shot : ' + _miaPad(wouldLeaveQueue, 5));
    say('  ⭐ MISSED  — needs one, not flagged    : ' + _miaPad(wouldEnterQueue, 5));
    if (wouldLeaveQueue + wouldEnterQueue === 0) {
      say('  ✅ ZERO rows misclassified right now. The photo half of the Phase 1');
      say('     argument is WEAK at this moment — judge it on the ' + _miaDaysAgo(_miaLastFullDay || '') );
      say('     day(s) since the last full sync, and re-run just before the next one.');
    } else {
      say('  ⚠ ' + (wouldLeaveQueue + wouldEnterQueue) + ' row(s) are misclassified by the key divergence alone.');
    }
  }

  // ── 5. ITEM-SPECIFICS SHREDDING — the plan's headline claim, measured ───────
  // MAIN takes the FIRST <Value> of a NameValueList; SUB joins ALL of them with
  // ", ". Both write the same C: columns. So the two populations below were last
  // written by different parsers, and the multi-value RATE between them is the
  // shredding, in numbers rather than argument.
  //   A = rows last written by SUB   (lastUpdated == the big full-sync day)
  //   B = rows last written by MAIN  (lastUpdated  >  that day)
  // ⚠ CAVEAT, stated because it matters: MAIN only touches items that CHANGED,
  //   so B is not a random sample of the catalogue. Read the gap as strong
  //   evidence of direction and rough size, not as a precise percentage.
  say('');
  say('── ITEM-SPECIFICS SHREDDING  (MAIN keeps 1 value, SUB keeps all) ─');
  if (!_miaLastFullDay) {
    say('  (no full-sync day identified above; skipped)');
  } else {
    var cCols = [];
    for (var ci = 0; ci < headers.length; ci++) {
      if (String(headers[ci] || '').indexOf('C:') === 0) cCols.push(ci);
    }
    var luIdx = colIdx('lastUpdated') - 1;
    if (cCols.length === 0 || luIdx < 0) {
      say('  (no C: columns or no lastUpdated; skipped)');
    } else {
      var all = mi.getRange(2, 1, nRows, lastCol).getValues();   // one read
      var spotIdx = colIdx('C:Compatible Equipment Type') - 1;

      var st = { A: { cells: 0, multi: 0, rows: 0, spot: 0, spotMulti: 0 },
                 B: { cells: 0, multi: 0, rows: 0, spot: 0, spotMulti: 0 } };
      for (var r3 = 0; r3 < nRows; r3++) {
        var day = _miaDayKey(all[r3][luIdx]);
        var pop = (day === _miaLastFullDay) ? 'A' : (day > _miaLastFullDay ? 'B' : null);
        if (!pop) continue;                       // older rows: neither parser recently
        st[pop].rows++;
        for (var cc = 0; cc < cCols.length; cc++) {
          var val = String(all[r3][cCols[cc]] || '').trim();
          if (!val) continue;
          st[pop].cells++;
          if (val.indexOf(', ') >= 0) st[pop].multi++;
        }
        if (spotIdx >= 0) {
          var sv = String(all[r3][spotIdx] || '').trim();
          if (sv) { st[pop].spot++; if (sv.indexOf(', ') >= 0) st[pop].spotMulti++; }
        }
      }
      function rate(o, k, kk) { return o[k] ? ((o[kk] / o[k]) * 100).toFixed(1) + '%' : 'n/a'; }
      say('  C: columns scanned : ' + cCols.length);
      say('');
      say('  A · last written by SUB  (' + _miaLastFullDay + ')  rows ' + _miaPad(st.A.rows, 5));
      say('      filled C: cells ' + _miaPad(st.A.cells, 7) + '   multi-value ' +
          _miaPad(st.A.multi, 6) + '  = ' + rate(st.A, 'cells', 'multi'));
      say('  B · last written by MAIN (since)        rows ' + _miaPad(st.B.rows, 5));
      say('      filled C: cells ' + _miaPad(st.B.cells, 7) + '   multi-value ' +
          _miaPad(st.B.multi, 6) + '  = ' + rate(st.B, 'cells', 'multi'));
      if (spotIdx >= 0) {
        say('');
        say('  spotlight — C:Compatible Equipment Type');
        say('      SUB-written  ' + _miaPad(st.A.spotMulti, 5) + ' / ' + _miaPad(st.A.spot, 5) +
            ' carry >1 value  = ' + rate(st.A, 'spot', 'spotMulti'));
        say('      MAIN-written ' + _miaPad(st.B.spotMulti, 5) + ' / ' + _miaPad(st.B.spot, 5) +
            ' carry >1 value  = ' + rate(st.B, 'spot', 'spotMulti'));
      }
      say('');
      var aR = st.A.cells ? st.A.multi / st.A.cells : 0;
      var bR = st.B.cells ? st.B.multi / st.B.cells : 0;
      if (aR > 0 && bR < aR * 0.5) {
        say('  ⭐ CONFIRMED: MAIN-written rows carry roughly ' +
            (aR / Math.max(bR, 0.0001)).toFixed(1) + '× fewer multi-value cells.');
        say('     Every hourly touch flattens them; the weekend sync restores them.');
        say('     ' + st.B.rows + ' row(s) are shredded RIGHT NOW.');
      } else if (aR > 0) {
        say('  ⚠ NOT confirmed at this sample size — the two rates are close (' +
            rate(st.A, 'cells', 'multi') + ' vs ' + rate(st.B, 'cells', 'multi') + ').');
        say('     Do not lead with this claim until it is re-measured later in the week.');
      } else {
        say('  (no multi-value cells found at all — nothing to shred)');
      }
    }
  }

  say('');
  say('════════ END ════════');
  return out.join('\n');
}


/* ── helpers (MI-audit private; `_mia` prefix keeps them out of the global namespace's way) ── */

function _miaNow() {
  return Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd HH:mm') + ' Houston';
}

/**
 * lastUpdated is written by both parsers as new Date().toISOString() — a STRING.
 * But Sheets coerces date-looking strings into real Dates on write, so getValues()
 * can hand back either. Gotcha #16 in CLAUDE.md, third instance of this class.
 * Returns '' for empty, '?' for unreadable, else a Chicago yyyy-MM-dd key.
 */
function _miaDayKey(v) {
  if (v === '' || v == null) return '';
  var d = null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    d = v;
  } else {
    var s = String(v).trim();
    if (!s) return '';
    var t = Date.parse(s);
    if (!isNaN(t)) d = new Date(t);
  }
  if (!d || isNaN(d.getTime())) return '?';
  return Utilities.formatDate(d, 'America/Chicago', 'yyyy-MM-dd');
}

function _miaWeekday(key) {
  if (!key) return '';
  var p = String(key).split('-');
  // Noon UTC anchor so no DST hop can flip the day (the _oaDayAdd lesson).
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12, 0, 0));
  return '(' + ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getUTCDay()] + ')';
}

function _miaDaysAgo(key) {
  var p = String(key).split('-');
  var then = Date.UTC(+p[0], +p[1] - 1, +p[2], 12, 0, 0);
  var todayKey = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd').split('-');
  var now = Date.UTC(+todayKey[0], +todayKey[1] - 1, +todayKey[2], 12, 0, 0);
  return Math.round((now - then) / 86400000);
}

function _miaPct(n, total) {
  if (!total) return '0.0%';
  return ((n / total) * 100).toFixed(1) + '%';
}

function _miaPad(n, w) {
  var s = String(n);
  while (s.length < w) s = ' ' + s;
  return s;
}

function _miaBar(n, total) {
  var width = total ? Math.round((n / total) * 40) : 0;
  var s = '';
  for (var i = 0; i < width; i++) s += '█';
  return s;
}

function _miaColLetter(i) {
  var s = '';
  while (i > 0) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
  return s;
}
