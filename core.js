/*
 * Tick Mark Manager - core re-lettering logic.
 * Pure functions with no Excel dependency, so they can be unit tested.
 */
(function (root) {
  "use strict";

  var TICK_RE = /^[A-Za-z]{1,2}$/;

  /** 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB ... (like Excel column letters) */
  function letterFor(i) {
    var s = "";
    var n = i + 1;
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function isTickText(v) {
    return typeof v === "string" && TICK_RE.test(v.trim());
  }

  /**
   * Parse the legend block (2 columns: letter, description).
   * Rows whose first cell is not a tick letter (the header, blank rows) are skipped.
   * Returns entries in legend order. Duplicate letters keep the first occurrence.
   */
  function parseLegend(values) {
    var entries = [];
    var seen = {};
    for (var r = 0; r < values.length; r++) {
      var letter = values[r][0];
      if (!isTickText(letter)) continue;
      letter = letter.trim().toUpperCase();
      if (seen[letter]) continue;
      seen[letter] = true;
      var desc = values[r].length > 1 && values[r][1] != null ? String(values[r][1]) : "";
      entries.push({ letter: letter, desc: desc, offset: r });
    }
    return entries;
  }

  /**
   * From a used-range value grid, list cells that look like tick marks
   * (still need a style check). Cells inside the legend block are skipped.
   */
  function findCandidates(values, rowIndex, columnIndex, legendBox) {
    var out = [];
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var v = values[r][c];
        if (!isTickText(v)) continue;
        var row = rowIndex + r;
        var col = columnIndex + c;
        if (legendBox && inBox(row, col, legendBox)) continue;
        var t = v.trim();
        out.push({ row: row, col: col, letter: t.toUpperCase(), lower: t === t.toLowerCase() });
      }
    }
    return out;
  }

  function inBox(row, col, box) {
    return row >= box.rowIndex && row < box.rowIndex + box.rowCount &&
      col >= box.columnIndex && col < box.columnIndex + box.columnCount;
  }

  function countTicks(ticks) {
    var counts = {};
    ticks.forEach(function (t) { counts[t.letter] = (counts[t.letter] || 0) + 1; });
    return counts;
  }

  /** First position (row-major) of each letter, for "re-letter by position". */
  function firstPositions(ticks) {
    var pos = {};
    ticks.forEach(function (t) {
      var p = pos[t.letter];
      if (!p || t.row < p.row || (t.row === p.row && t.col < p.col)) pos[t.letter] = { row: t.row, col: t.col };
    });
    return pos;
  }

  /**
   * Work out which legend entries to keep when running automatically.
   *  - an entry whose count went from >0 to 0 since last time is dropped
   *  - a letter that was in the legend last time but whose legend row is gone is "deleted"
   *    (its remaining cells get cleared)
   * prev = { letters: [...], counts: {...} } or null (no history -> nothing dropped)
   */
  function autoPlan(entries, counts, prev) {
    if (!prev) return { kept: entries.slice(), deleted: [] };
    var kept = entries.filter(function (e) {
      return !((prev.counts[e.letter] || 0) > 0 && (counts[e.letter] || 0) === 0);
    });
    var current = {};
    entries.forEach(function (e) { current[e.letter] = true; });
    var deleted = prev.letters.filter(function (l) { return !current[l]; });
    return { kept: kept, deleted: deleted };
  }

  /**
   * Build the new legend and the old->new letter mapping.
   * oldEntries: current legend entries.
   * keptEntries: entries to keep, in their new order. An entry with letter === null is new.
   * deletedLetters: letters no longer in the legend whose cells should be cleared.
   * mapping[old] = new letter, or null to clear the cell. Letters not in mapping are left alone.
   */
  function buildPlan(oldEntries, keptEntries, deletedLetters) {
    var mapping = {};
    oldEntries.forEach(function (e) { mapping[e.letter] = null; });
    (deletedLetters || []).forEach(function (l) { mapping[l] = null; });
    keptEntries.forEach(function (e, i) { if (e.letter) mapping[e.letter] = letterFor(i); });
    // Descriptions in the legend can mention other tick marks too ("see TM {B}").
    var newEntries = keptEntries.map(function (e, i) {
      return { letter: letterFor(i), desc: remapReferences(e.desc, mapping).text, oldLetter: e.letter };
    });
    var changed = oldEntries.length !== newEntries.length ||
      (deletedLetters || []).length > 0 ||
      oldEntries.some(function (e, i) { return e.letter !== newEntries[i].letter || e.desc !== newEntries[i].desc; });
    return { mapping: mapping, newEntries: newEntries, changed: changed };
  }

  /** True if any letter actually changes or gets cleared (cells need rewriting). */
  function mappingMovesCells(mapping, counts) {
    return Object.keys(mapping).some(function (k) {
      return mapping[k] !== k && (counts[k] || 0) > 0;
    });
  }

  /** Tick cells whose letter is not in the legend (and not being deleted). */
  function orphans(ticks, mapping) {
    return ticks.filter(function (t) { return !Object.prototype.hasOwnProperty.call(mapping, t.letter); });
  }

  /**
   * Would applying this mapping make an orphan look like a real tick mark?
   * e.g. orphan "A" exists and legend B gets re-lettered to A.
   */
  function orphanCollision(ticks, mapping) {
    var targets = {};
    Object.keys(mapping).forEach(function (k) { if (mapping[k] && mapping[k] !== k) targets[mapping[k]] = true; });
    return orphans(ticks, mapping).filter(function (t) { return targets[t.letter]; });
  }

  /** Counts keyed by the new letters, for remembering state after a change. */
  function remapCounts(counts, mapping) {
    var out = {};
    Object.keys(counts).forEach(function (k) {
      var nk = Object.prototype.hasOwnProperty.call(mapping, k) ? mapping[k] : k;
      if (nk) out[nk] = (out[nk] || 0) + counts[k];
    });
    return out;
  }

  function orderByPosition(entries, ticks) {
    var pos = firstPositions(ticks);
    return entries
      .map(function (e, i) { return { e: e, i: i, p: pos[e.letter] }; })
      .sort(function (a, b) {
        if (a.p && b.p) return a.p.row - b.p.row || a.p.col - b.p.col || a.i - b.i;
        if (a.p) return -1;
        if (b.p) return 1;
        return a.i - b.i;
      })
      .map(function (x) { return x.e; });
  }


  /* ------------------------------------------------ references in text */
  /*
   * Recognized (any capitalization, letters A-Z or AA-ZZ, upper or lower case):
   *   tick mark A, tickmark A, tick-mark A, tick marks A and B, TM A, tm a, Tm {A}, TM{a}, {A}
   * A bare letter after "tick mark"/"TM" only counts if it is a known tick mark letter,
   * so ordinary words ("tick mark is", "TM an") are not touched.
   */
  var PREFIX_RE = /\b(?:tick[ \t-]*marks?|tms?)(?![A-Za-z0-9])[ \t]*[:#-]?[ \t]*/gi;
  var BRACED_RE = /^\{[ \t]*([A-Za-z]{1,2})[ \t]*\}/;
  var BARE_RE = /^([A-Za-z]{1,2})(?![A-Za-z0-9])/;
  var SEP_RE = /^(?:[ \t]*,[ \t]*(?:(?:and|or)[ \t]+|&[ \t]*)?|[ \t]+(?:and|or)[ \t]+|[ \t]*[&\/][ \t]*)/i;
  var ANY_BRACED_RE = /\{[ \t]*([A-Za-z]{1,2})[ \t]*\}/g;

  function readLetter(text, i, known) {
    var rest = text.slice(i);
    var m = BRACED_RE.exec(rest);
    if (m) return { start: i, end: i + m[0].length, raw: m[1], braced: true };
    m = BARE_RE.exec(rest);
    if (m && known(m[1].toUpperCase())) return { start: i, end: i + m[0].length, raw: m[1], braced: false };
    return null;
  }

  /** Find tick mark references in a piece of text. known(letter) -> bool for bare letters. */
  function findReferences(text, known) {
    if (typeof text !== "string" || !text) return [];
    var toks = [];
    var m;
    PREFIX_RE.lastIndex = 0;
    while ((m = PREFIX_RE.exec(text))) {
      var i = m.index + m[0].length;
      var tok = readLetter(text, i, known);
      while (tok) {
        toks.push(tok);
        i = tok.end;
        var sep = SEP_RE.exec(text.slice(i));
        if (!sep) break;
        tok = readLetter(text, i + sep[0].length, known);
      }
      PREFIX_RE.lastIndex = Math.max(PREFIX_RE.lastIndex, i);
    }
    ANY_BRACED_RE.lastIndex = 0;
    while ((m = ANY_BRACED_RE.exec(text))) {
      var s0 = m.index, e0 = s0 + m[0].length;
      var taken = toks.some(function (t) { return s0 < t.end && t.start < e0; });
      if (!taken) toks.push({ start: s0, end: e0, raw: m[1], braced: true });
    }
    toks.sort(function (a, b) { return a.start - b.start; });
    toks.forEach(function (t) { t.letter = t.raw.toUpperCase(); t.lower = t.raw === t.raw.toLowerCase(); });
    return toks;
  }

  function withCase(letter, lower) { return lower ? letter.toLowerCase() : letter; }

  /**
   * Rewrite references using mapping (old -> new letter, null = deleted).
   * Deleted references become "[deleted]". Letters not in mapping are left alone.
   */
  function remapReferences(text, mapping) {
    var has = function (l) { return Object.prototype.hasOwnProperty.call(mapping, l); };
    var toks = findReferences(text, has);
    var out = "", last = 0, changed = false, deleted = false;
    toks.forEach(function (t) {
      if (!has(t.letter)) return;
      var nl = mapping[t.letter];
      if (nl === t.letter) return;
      var rep;
      if (nl === null) { rep = "[deleted]"; deleted = true; }
      else rep = t.braced ? "{" + withCase(nl, t.lower) + "}" : withCase(nl, t.lower);
      out += text.slice(last, t.start) + rep;
      last = t.end;
      changed = true;
    });
    if (!changed) return { text: text, changed: false, deleted: false };
    return { text: out + text.slice(last), changed: true, deleted: deleted };
  }

  /** Text cells that might contain references (formulas and the legend are skipped). */
  function findRefCells(values, formulas, rowIndex, columnIndex, legendBox) {
    var out = [];
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var v = values[r][c];
        if (typeof v !== "string" || v.length < 3 || !/tick|tm|\{/i.test(v)) continue;
        var f = formulas && formulas[r] ? formulas[r][c] : "";
        if (typeof f === "string" && f.charAt(0) === "=") continue;
        var row = rowIndex + r, col = columnIndex + c;
        if (legendBox && inBox(row, col, legendBox)) continue;
        out.push({ row: row, col: col, text: v });
      }
    }
    return out;
  }

  /** How many times each known letter is referenced across ref cells. */
  function countReferences(refCells, knownLetters) {
    var set = {};
    knownLetters.forEach(function (l) { set[l] = true; });
    var counts = {};
    refCells.forEach(function (rc) {
      findReferences(rc.text, function (l) { return !!set[l]; }).forEach(function (t) {
        if (set[t.letter]) counts[t.letter] = (counts[t.letter] || 0) + 1;
      });
    });
    return counts;
  }

  /* ------------------------------------- tick marks as legend formulas */

  function colIndex(letters) {
    var n = 0;
    letters = letters.toUpperCase();
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  // =A14, =$A$14, =Sheet1!$A$14, ='Cash Disb'!A14
  var REF_RE = /^=\s*(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)\s*$/;

  /** A formula that is just one reference to a cell on this sheet -> {row, col}; otherwise null. */
  function parseRef(formula, sheetName) {
    if (typeof formula !== "string") return null;
    var m = REF_RE.exec(formula);
    if (!m) return null;
    var sh = m[1] != null ? m[1].replace(/''/g, "'") : m[2];
    if (sh && sheetName && sh.toLowerCase() !== String(sheetName).toLowerCase()) return null;
    return { row: parseInt(m[4], 10) - 1, col: colIndex(m[3]) };
  }

  /** The formula a tick mark cell gets: an absolute reference to its letter in the legend. */
  function refFormula(row, col) {
    return "=$" + letterFor(col) + "$" + (row + 1);
  }

  /**
   * Sort the used range into:
   *  consts       - cells holding a plain letter (need a style check to count as tick marks)
   *  formulaTicks - cells whose formula points at a letter in the legend (=$A$14) -> tick marks
   *  refBroken    - cells showing =#REF! (a tick mark whose legend row was deleted, if styled)
   *  blankRefs    - formulas pointing at an empty row inside the legend
   * entryRows: { rowIndex: letter } for legend rows that hold a tick mark.
   */
  function classifyCells(values, formulas, rowIndex, columnIndex, legendBox, entryRows, sheetName) {
    var out = { consts: [], formulaTicks: [], refBroken: [], blankRefs: [] };
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var row = rowIndex + r, col = columnIndex + c;
        if (legendBox && inBox(row, col, legendBox)) continue;
        var v = values[r][c];
        var f = formulas && formulas[r] ? formulas[r][c] : v;
        if (typeof f === "string" && f.charAt(0) === "=") {
          if (/^=\s*#REF!\s*$/i.test(f)) { out.refBroken.push({ row: row, col: col }); continue; }
          var ref = parseRef(f, sheetName);
          if (ref && legendBox && ref.col === legendBox.columnIndex &&
              ref.row > legendBox.rowIndex && ref.row < legendBox.rowIndex + legendBox.rowCount) {
            var L = entryRows[ref.row];
            if (L) out.formulaTicks.push({ row: row, col: col, letter: L, refRow: ref.row, formula: true });
            else out.blankRefs.push({ row: row, col: col });
          }
          continue;
        }
        if (isTickText(v)) {
          var t = v.trim();
          out.consts.push({ row: row, col: col, letter: t.toUpperCase(), lower: t === t.toLowerCase() });
        }
      }
    }
    return out;
  }

  /** "A1"-style address from zero-based indexes. */
  function cellAddress(row, col) {
    return letterFor(col) + (row + 1);
  }

  var TickCore = {
    letterFor: letterFor,
    isTickText: isTickText,
    parseLegend: parseLegend,
    findCandidates: findCandidates,
    inBox: inBox,
    countTicks: countTicks,
    firstPositions: firstPositions,
    autoPlan: autoPlan,
    buildPlan: buildPlan,
    mappingMovesCells: mappingMovesCells,
    orphans: orphans,
    orphanCollision: orphanCollision,
    remapCounts: remapCounts,
    orderByPosition: orderByPosition,
    cellAddress: cellAddress,
    colIndex: colIndex,
    parseRef: parseRef,
    refFormula: refFormula,
    classifyCells: classifyCells,
    withCase: withCase,
    findReferences: findReferences,
    remapReferences: remapReferences,
    findRefCells: findRefCells,
    countReferences: countReferences
  };

  if (typeof module !== "undefined" && module.exports) module.exports = TickCore;
  else root.TickCore = TickCore;
})(this);
