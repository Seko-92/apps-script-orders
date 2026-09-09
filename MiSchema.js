// =======================================================================================
// MiSchema.js — resolve Master Inventory columns BY HEADER NAME, never by position
// =======================================================================================
//
// WHY THIS EXISTS. MI is 216 columns wide and its column ORDER is not ours to control.
// The `C:` columns are generated from eBay item specifics by n8n's `autoMapInputData`,
// one column per distinct <Name>, appended in whatever order the first written item
// happens to have. So a column's INDEX is an accident of history, not a contract.
//
// Every MI reader in this codebase resolves by header name — except one, and the way it
// failed is the reason this file exists:
//
//     var data = miSheet.getRange(2, 2, lastRow - 1, 39).getValues();
//     location: data[i][38]   // col 40 = C:Model Year = the kit's aisle
//
// `C:Model Year` sitting on column 40 is a coincidence. If it ever moves, that read
// returns a NEIGHBOURING FIELD and `KitRegistry` classifies kits by testing it against
// /^K[-\s]/ — which quietly fails, so EVERY kit becomes MANUAL. No exception, no error,
// no wrong-looking number. A READY kit that should ship as one pre-assembled box gets
// expanded into components instead. That is the failure this file makes impossible.
//
// THE THREE PROPERTIES THAT MATTER:
//
//   1. FAILS LOUD. A missing required column THROWS, naming the column and listing the
//      nearest header matches. The bug being replaced was silent; a resolver that
//      returned -1 and let the caller carry on would just relocate the silence.
//
//   2. BOUNDED READ, WITHOUT A HARDCODED WIDTH. `readColumns()` resolves the names,
//      computes the min..max span that covers them, and issues ONE getRange over
//      exactly that span — then hands back offsets relative to the slice. Callers keep
//      the "don't getDataRange() a 216-column sheet" win without ever naming a width.
//      ⚠ Span width is nearly free anyway: measured on this sheet 2026-08-19, a
//      198-column read vs a 39-column read was 1,905ms vs 1,872ms — the round trip
//      dominates, not the payload. So correctness costs nothing here.
//
//   3. MEMOIZED PER EXECUTION. The header row is read once per sheet per execution and
//      cached in memory. Apps Script executions are short-lived, so there is no
//      cross-run staleness to reason about — and callers that resolve inside a loop
//      (KitRegistry builds its map twice per import) pay one round trip, not N.
//
// ⚠ NOT a persistent cache. Never move this to CacheService or Script Properties: a
// stale column map is exactly the silent-wrong-column bug wearing a different hat.
//
// Sheet-generic by construction (it takes a sheet), but named for MI because MI is the
// concrete problem. Generalise it when a second sheet actually needs it, not before.
// =======================================================================================

/** In-memory, execution-scoped header cache: { sheetName: [header, header, ...] }. */
var _miHeaderCache = {};

var MiSchema = {

  /**
   * Header row for a sheet, read at most once per execution.
   * @param {Sheet} sheet
   * @return {string[]} raw header values (0-based)
   */
  headers: function (sheet) {
    var key = sheet.getSheetName();
    if (_miHeaderCache[key]) return _miHeaderCache[key];

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return (_miHeaderCache[key] = []);

    var row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    return (_miHeaderCache[key] = row);
  },

  /** Drop the memo — for tests, and after any deliberate column change mid-execution. */
  clearCache: function () { _miHeaderCache = {}; },

  /**
   * Resolve header names to 0-based indices.
   *
   * @param {Sheet} sheet
   * @param {string[]} names            required — a miss THROWS
   * @param {string[]=} optionalNames   a miss yields -1, no throw
   * @return {Object} { name: 0-basedIndex }   (-1 only for optional misses)
   */
  resolve: function (sheet, names, optionalNames) {
    var headers = MiSchema.headers(sheet);

    // Build a normalised lookup once: trimmed + lowercased.
    var lookup = {};
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] == null ? "" : headers[i]).trim().toLowerCase();
      // FIRST occurrence wins — a duplicated header name must not silently
      // re-point an already-resolved column to a later copy.
      if (h && !lookup.hasOwnProperty(h)) lookup[h] = i;
    }

    var out = {};
    var missing = [];

    function find(name) {
      var k = String(name).trim().toLowerCase();
      return lookup.hasOwnProperty(k) ? lookup[k] : -1;
    }

    (names || []).forEach(function (n) {
      var idx = find(n);
      if (idx === -1) missing.push(n);
      out[n] = idx;
    });

    (optionalNames || []).forEach(function (n) { out[n] = find(n); });

    if (missing.length) {
      throw new Error(
        "Master Inventory is missing required column(s): " + missing.join(", ") +
        ". " + MiSchema._nearMissHint(headers, missing) +
        " (sheet '" + sheet.getSheetName() + "' has " + headers.length + " columns)"
      );
    }
    return out;
  },

  /**
   * Resolve names, read ONE bounded span covering them, return rows + slice offsets.
   *
   * The returned `idx` is relative to each returned row, so a caller never touches an
   * absolute column number:
   *
   *     var r = MiSchema.readColumns(mi, [DB_SKU_HEADER, DB_LOCATION_HEADER]);
   *     r.rows.forEach(function (row) {
   *       var sku = row[r.idx[DB_SKU_HEADER]];
   *       var loc = row[r.idx[DB_LOCATION_HEADER]];
   *     });
   *
   * @param {Sheet} sheet
   * @param {string[]} names           required columns
   * @param {Object=} opts             { optional: string[], startRow: number }
   * @return {Object} { rows, idx, absCol, headers, span:{first,last,width} }
   */
  readColumns: function (sheet, names, opts) {
    opts = opts || {};
    var startRow = opts.startRow || 2;                 // default: skip the header row
    var resolved = MiSchema.resolve(sheet, names, opts.optional);

    // Span across the columns we actually found (optional misses are -1, excluded).
    var found = [];
    Object.keys(resolved).forEach(function (n) {
      if (resolved[n] >= 0) found.push(resolved[n]);
    });
    if (!found.length) return { rows: [], idx: resolved, absCol: {}, headers: MiSchema.headers(sheet), span: null };

    var first = Math.min.apply(null, found);
    var last = Math.max.apply(null, found);
    var width = last - first + 1;

    var lastRow = sheet.getLastRow();
    var rows = (lastRow < startRow)
      ? []
      : sheet.getRange(startRow, first + 1, lastRow - startRow + 1, width).getValues();

    // Offsets relative to the slice, and absolute 1-based columns for writers.
    var idx = {}, absCol = {};
    Object.keys(resolved).forEach(function (n) {
      var a = resolved[n];
      idx[n] = (a >= 0) ? (a - first) : -1;
      absCol[n] = (a >= 0) ? (a + 1) : -1;
    });

    return {
      rows: rows,
      idx: idx,
      absCol: absCol,
      headers: MiSchema.headers(sheet),
      span: { first: first + 1, last: last + 1, width: width }
    };
  },

  /**
   * Best-effort "did you mean" for a throw message. A renamed column is far more
   * likely than a deleted one, so naming the near-misses turns a dead end into a fix.
   */
  _nearMissHint: function (headers, missing) {
    var hints = [];
    missing.forEach(function (m) {
      var needle = String(m).trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (!needle) return;
      var near = [];
      for (var i = 0; i < headers.length && near.length < 3; i++) {
        var h = String(headers[i] == null ? "" : headers[i]).trim();
        var flat = h.toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!flat) continue;
        if (flat.indexOf(needle) !== -1 || needle.indexOf(flat) !== -1) near.push(h);
      }
      if (near.length) hints.push("'" + m + "' ~ " + near.join(" / "));
    });
    return hints.length ? "Closest headers: " + hints.join("; ") + "." : "No similar header found.";
  }
};
