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
  var pu1 = readCol('pictureUrl1');
  var p1  = readCol('picture1');
  if (!pu1 || !p1) {
    say('  (need both "pictureUrl1" and "picture1" headers; one is missing)');
  } else {
    var bothEmpty = 0, onlyOld = 0, onlyNew = 0, differ = 0, same = 0;
    for (var q = 0; q < nRows; q++) {
      var a = String(pu1[q][0] || '').trim();   // what the readers actually use
      var b = String(p1[q][0]  || '').trim();   // what MAIN's hourly writes
      if (!a && !b) { bothEmpty++; continue; }
      if (!a && b)  { onlyOld++;   continue; }
      if (a && !b)  { onlyNew++;   continue; }
      if (a === b) same++; else differ++;
    }
    say('  both empty                      : ' + _miaPad(bothEmpty, 5));
    say('  pictureUrl1 only (SUB reached)  : ' + _miaPad(onlyNew, 5));
    say('  picture1 only  (MAIN-only row)  : ' + _miaPad(onlyOld, 5) + '   ← INVISIBLE to all 4 readers');
    say('  both, identical                 : ' + _miaPad(same, 5));
    say('  both, DIFFERENT                 : ' + _miaPad(differ, 5) + '   ← photo changed since the last Sunday');
    if (onlyOld + differ > 0) {
      say('  ⚠ ' + (onlyOld + differ) + ' row(s) where the hourly sync has fresher photo data than');
      say('    the four consumers can see. Phase 1 step 3 closes exactly this gap.');
    } else {
      say('  ✅ no divergence right now (expected soon after a Sunday run).');
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
