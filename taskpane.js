/*
 * Tick Mark Manager - Excel add-in (task pane + shared runtime).
 * Each worksheet has its own legend, found through the sheet-level name "TickMarkLegend".
 * A tick mark is a cell whose formula points at its letter in the legend (=$A$14), so Ctrl+[ jumps there.
 * Placed tick marks also get the cell style "Tick Mark" (red, bold, centered).
 * Older tick marks that hold a plain letter with the "Tick Mark" style are still supported.
 */
/* global Office, Excel, TickCore */
"use strict";

var STYLE_TICK = "Tick Mark";
var LEGEND_NAME = "TickMarkLegend";
var LEGEND_HEADER = "Tick Mark Legend";
var TICK_COLOR = "#C00000";
var MAX_FILL = 5000;
var DEBOUNCE_MS = 700;
// Yellow for text that mentions a deleted tick mark. One step off Excel's standard yellow (#FFFF00)
// so the add-in can remove its own highlight once the text is fixed, without touching other yellow cells.
var DELETED_FILL = "#FFFF01";
// Mentions in text use this red. It is one step off the tick mark red (#C00000), which looks the same
// but lets the add-in tell its own formatting apart from red text someone typed.
var MENTION_COLOR = "#C00001";
var MAX_EDIT_CELLS = 200;

var state = {};     // worksheetId -> { letters: [...], counts: {...} } from the last time we looked
var timers = {};
var edited = {};    // worksheetId -> addresses changed since the last automatic pass
var chain = Promise.resolve();

function UserError(msg) { this.message = msg; this.isUser = true; }

/* Run Excel work one job at a time so automatic and manual changes never overlap. */
function enqueue(fn) {
  var p = chain.then(fn);
  chain = p.catch(function () {});
  return p;
}

function lowerOn() {
  try { return localStorage.getItem("tm.lower") === "on"; } catch (e) { return false; }
}

function autoOn() {
  try { return localStorage.getItem("tm.auto") !== "off"; } catch (e) { return true; }
}

/* ---------------------------------------------------------------- startup */

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    var n = document.getElementById("notInExcel");
    if (n) n.style.display = "block";
    return;
  }
  wireUi();
  registerShortcuts();
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    showStatus("This version of Excel is too old for Tick Marks. Please update Excel.", "error");
    return;
  }
  enqueue(function () {
    return Excel.run(function (ctx) {
      ctx.workbook.worksheets.onChanged.add(onSheetChanged);
      ctx.workbook.worksheets.onActivated.add(function () { return enqueue(refreshUi); });
      return ctx.sync();
    });
  })
    .then(function () { return enqueue(recordAllSheets); })
    .then(function () { return enqueue(refreshUi); })
    .catch(showError);
});

/* Remember the current tick marks on every sheet, so a later deletion can be noticed. */
function recordAllSheets() {
  return Excel.run(async function (ctx) {
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/id");
    await ctx.sync();
    var fix = [];
    for (var i = 0; i < sheets.items.length; i++) {
      var data = await recordState(ctx, sheets.items[i], sheets.items[i].id);
      if (data.legend) fix.push(sheets.items[i].id); // tidy up + format mentions
    }
    return fix;
  }).then(async function (fix) {
    for (var i = 0; i < fix.length; i++) await autoSync(fix[i]);
  });
}

function onSheetChanged(evt) {
  if (evt.source && evt.source !== Excel.EventSource.local) return Promise.resolve();
  var id = evt.worksheetId;
  if (evt.address) (edited[id] = edited[id] || []).push(evt.address);
  clearTimeout(timers[id]);
  timers[id] = setTimeout(function () {
    var addrs = edited[id] || [];
    edited[id] = [];
    enqueue(function () { return autoOn() ? autoSync(id, addrs) : null; })
      .then(function () { return enqueue(refreshUi); })
      .catch(showError);
  }, DEBOUNCE_MS);
  return Promise.resolve();
}

/* --------------------------------------------------------- reading a sheet */

async function readSheet(ctx, ws) {
  ws.load("name");
  var legend = null;
  try {
    var name = ws.names.getItemOrNullObject(LEGEND_NAME);
    await ctx.sync();
    if (!name.isNullObject) {
      var r = name.getRangeOrNullObject();
      r.load("values,rowIndex,columnIndex,rowCount,columnCount");
      await ctx.sync();
      if (!r.isNullObject) {
        legend = {
          entries: TickCore.parseLegend(r.values),
          box: { rowIndex: r.rowIndex, columnIndex: r.columnIndex, rowCount: r.rowCount, columnCount: r.columnCount }
        };
        legend.entries.forEach(function (e) { e.row = r.rowIndex + e.offset; });
      }
    }
  } catch (e) {
    legend = null; // name points at deleted cells (#REF!)
  }

  var ticks = [], broken = [], refCells = [];
  var used = ws.getUsedRangeOrNullObject(true);
  used.load("values,formulas,rowIndex,columnIndex");
  await ctx.sync();
  if (!legend && !used.isNullObject) {
    // No named legend: adopt a "Tick Mark Legend" block already on the sheet.
    var found = TickCore.findLegendBlock(used.values, used.rowIndex, used.columnIndex);
    if (found) {
      var vals = used.values.slice(found.rowIndex - used.rowIndex, found.rowIndex - used.rowIndex + found.rowCount)
        .map(function (row) { return row.slice(found.columnIndex - used.columnIndex, found.columnIndex - used.columnIndex + 2); });
      legend = { entries: TickCore.parseLegend(vals), box: found, adopted: true };
      legend.entries.forEach(function (e) { e.row = found.rowIndex + e.offset; });
    }
  }
  if (!used.isNullObject) {
    var entryRows = {};
    if (legend) legend.entries.forEach(function (e) { entryRows[e.row] = e.letter; });
    var cls = TickCore.classifyCells(used.values, used.formulas, used.rowIndex, used.columnIndex,
      legend && legend.box, entryRows, ws.name);
    refCells = TickCore.findRefCells(used.values, used.formulas, used.rowIndex, used.columnIndex, legend && legend.box);

    // Style checks: plain letters and #REF! cells only count if they carry the Tick Mark style.
    var check = cls.consts.concat(cls.refBroken, cls.formulaTicks).map(function (c) {
      var cell = ws.getCell(c.row, c.col);
      cell.load("style");
      return { c: c, cell: cell };
    });
    if (check.length) await ctx.sync();
    var styled = function (x) { return x.cell.style === STYLE_TICK; };
    check.forEach(function (x) {
      x.c.styled = styled(x);
      if (x.c.formula) ticks.push(x.c);
      else if (cls.consts.indexOf(x.c) >= 0) { if (x.c.styled) ticks.push(x.c); }
      else if (x.c.styled) broken.push(x.c); // =#REF! tick mark: its legend row was deleted
    });
    broken = broken.concat(cls.blankRefs);  // point at an emptied legend row
  }
  var letters = legend ? legend.entries.map(function (e) { return e.letter; }) : [];
  return {
    legend: legend, ticks: ticks, broken: broken, counts: TickCore.countTicks(ticks),
    refCells: refCells, refCounts: TickCore.countReferences(refCells, letters)
  };
}

async function recordState(ctx, ws, id) {
  var data = await readSheet(ctx, ws);
  state[id] = data.legend
    ? { letters: data.legend.entries.map(function (e) { return e.letter; }), counts: data.counts }
    : null;
  return data;
}

/* Something on the sheet should be tidied even if no letters change. */
function needsFix(data) {
  if (!data.legend) return false;
  var inLegend = {};
  data.legend.entries.forEach(function (e) { inLegend[e.letter] = true; });
  return data.legend.adopted || data.broken.length > 0 || legendUntidy(data.legend) ||
    data.ticks.some(function (t) { return (t.formula && (!t.styled || t.asText)) || (!t.formula && inLegend[t.letter]); });
}

/* True if the legend has gaps, a missing header or stray rows and should be rewritten. */
function legendUntidy(legend) {
  if (!legend) return false;
  return legend.box.rowCount !== legend.entries.length + 1 ||
    legend.entries.some(function (e, i) { return e.offset !== i + 1; });
}

/* ----------------------------------------------------------------- writing */

async function ensureStyle(ctx) {
  var st = ctx.workbook.styles.getItemOrNullObject(STYLE_TICK);
  await ctx.sync();
  if (st.isNullObject) {
    ctx.workbook.styles.add(STYLE_TICK);
    var s = ctx.workbook.styles.getItem(STYLE_TICK);
    s.font.color = TICK_COLOR;
    s.font.bold = true;
    s.horizontalAlignment = "Center";
    s.numberFormat = "General";
    await ctx.sync();
  } else {
    // A Text number format would make tick mark formulas show as "=$B$19".
    st.load("numberFormat");
    await ctx.sync();
    if (st.numberFormat !== "General") { st.numberFormat = "General"; await ctx.sync(); }
  }
}

/*
 * Update tick mark cells and text for a re-lettering.
 * legendPos: {row, col} of the legend header the new legend will be written at.
 * Formula tick marks are re-pointed at their entry's new legend row; deleted ones are cleared.
 */
function applyChanges(ws, data, plan, legendPos) {
  var mapping = plan.mapping;
  var newRow = {};
  plan.newEntries.forEach(function (e, i) { newRow[e.letter] = legendPos ? legendPos.row + 1 + i : null; });
  data.ticks.forEach(function (t) {
    if (!Object.prototype.hasOwnProperty.call(mapping, t.letter)) return; // plain letter not in legend: leave alone
    var nl = mapping[t.letter];
    var cell = ws.getCell(t.row, t.col);
    if (nl === null) { cell.clear(Excel.ClearApplyTo.contents); return; }
    if (t.formula) {
      if (newRow[nl] !== t.refRow || !t.styled || t.asText) {
        if (!t.styled) cell.style = STYLE_TICK;
        cell.numberFormat = [["General"]];
        cell.formulas = [[TickCore.refFormula(newRow[nl], legendPos.col)]];
      }
    } else if (legendPos) {
      // Older plain-letter tick mark: turn it into a formula pointing at the legend.
      cell.numberFormat = [["General"]];
      cell.formulas = [[TickCore.refFormula(newRow[nl], legendPos.col)]];
    } else if (nl !== t.letter) {
      cell.values = [[TickCore.withCase(nl, t.lower)]];
    }
  });
  data.broken.forEach(function (b) { ws.getCell(b.row, b.col).clear(Excel.ClearApplyTo.contents); });
}

/*
 * Paragraphs that mention tick marks ("see tick mark B", "TM {b}"):
 *  - re-letter the mentions (mapping old -> new, null = deleted -> "[deleted]" + yellow cell)
 *  - show each mention in bold red, the rest of the text unchanged (needs ExcelApi 1.18)
 * Returns how many cells had their text changed.
 */
async function updateMentions(ctx, ws, refCells, mapping, letters, extraCells) {
  var known = {};
  letters.forEach(function (l) { known[l] = true; });
  var isKnown = function (l) { return !!known[l]; };
  var rich = Office.context.requirements.isSetSupported("ExcelApi", "1.18");

  // Cells to look at: every cell whose text might mention a tick mark, plus cells just edited
  // (an edit can turn a mention into plain text that still carries the bold red).
  var items = [], seen = {};
  refCells.forEach(function (rc) {
    seen[rc.row + "," + rc.col] = true;
    items.push({ text: rc.text, cell: ws.getCell(rc.row, rc.col), legacy: true });
  });
  var extras = (extraCells || []).filter(function (x) { return !seen[x.row + "," + x.col]; }).map(function (x) {
    var cell = ws.getCell(x.row, x.col);
    cell.load("values,formulas");
    return { cell: cell, extra: true };
  });
  if (extras.length) {
    await ctx.sync();
    extras.forEach(function (x) {
      var v = x.cell.values[0][0], f = x.cell.formulas[0][0];
      if (typeof v === "string" && v !== "" && !(typeof f === "string" && f.charAt(0) === "=")) items.push({ text: v, cell: x.cell });
    });
  }
  items.forEach(function (it) {
    it.r = mapping ? TickCore.remapReferences(it.text, mapping) : { text: it.text, changed: false, deleted: false };
    it.spans = TickCore.referenceSpans(it.r.text, isKnown);
  });
  if (!items.length) return 0;

  if (rich) {
    items.forEach(function (it) { it.props = it.cell.getCellProperties({ textRuns: true, format: { font: { bold: true, color: true }, fill: { color: true } } }); });
    await ctx.sync();
  }
  var n = 0;
  items.forEach(function (it) {
    if (it.r.deleted) it.cell.format.fill.color = DELETED_FILL;
    if (it.r.changed) n++;
    if (!rich) {
      if (it.r.changed) it.cell.values = [[it.r.text]];
      return;
    }
    var cur = it.props.value[0][0];
    // Text fixed: take away the yellow the add-in put there.
    var fill = cur.format && cur.format.fill && typeof cur.format.fill.color === "string" ? cur.format.fill.color.toUpperCase() : "";
    if ((fill === DELETED_FILL || (it.legacy && fill === "#FFFF00")) && !/\[deleted\]/i.test(it.r.text)) it.cell.format.fill.clear();
    var marked = hasMentionFormat(cur, it.legacy);
    if (!it.spans.length && !marked && !it.r.changed) return; // ordinary text, never touched
    var want = TickCore.textRuns(it.r.text, it.spans);
    if (!it.r.changed && runsMatch(cur.textRuns, want)) return;
    var base = plainFont(cur);
    it.cell.setCellProperties([[{
      textRuns: want.map(function (w) {
        return { text: w.text, font: w.ref ? { bold: true, color: MENTION_COLOR } : base };
      })
    }]]);
  });
  return n;
}

/* Does any part of the cell carry the add-in's mention formatting? */
function hasMentionFormat(cur, legacy) {
  var test = function (f) {
    // legacy: the first release used the tick mark red (#C00000) for mentions; accept it in text
    // that talks about tick marks so those cells get cleaned up too.
    return isRefFont(f) || (legacy && f && f.bold && typeof f.color === "string" && f.color.toUpperCase() === TICK_COLOR);
  };
  if (cur.format && test(cur.format.font)) return true;
  return (cur.textRuns || []).some(function (r) { return test(r.font); });
}

function isRefFont(f) {
  return !!(f && f.bold && typeof f.color === "string" && f.color.toUpperCase() === MENTION_COLOR);
}

/* Does the cell already show exactly these mentions in bold red? */
function runsMatch(curRuns, want) {
  if (!curRuns || !curRuns.length) return false;
  var merged = [];
  curRuns.forEach(function (r) {
    var ref = isRefFont(r.font);
    var last = merged[merged.length - 1];
    if (last && last.ref === ref) last.text += r.text; else merged.push({ text: r.text, ref: ref });
  });
  merged = merged.filter(function (m) { return m.text !== ""; });
  if (merged.length !== want.length) return false;
  return merged.every(function (m, i) { return m.text === want[i].text && m.ref === want[i].ref; });
}

/* Font for the ordinary text around a mention: keep what the cell already uses. */
function plainFont(cur) {
  var runs = cur.textRuns || [];
  for (var i = 0; i < runs.length; i++) {
    if (runs[i].font && !isRefFont(runs[i].font)) {
      return { bold: !!runs[i].font.bold, color: runs[i].font.color || "#000000" };
    }
  }
  var f = cur.format && cur.format.font;
  if (f && !isRefFont(f) && f.color) return { bold: !!f.bold, color: f.color };
  return { bold: false, color: "#000000" };
}

/* Where a new legend goes: two rows below everything on the sheet (and below minRow). */
async function defaultLegendPos(ctx, ws, minRow) {
  var u = ws.getUsedRangeOrNullObject(true);
  u.load("rowIndex,rowCount,columnIndex");
  await ctx.sync();
  var bottom = u.isNullObject ? -1 : u.rowIndex + u.rowCount - 1;
  if (minRow != null) bottom = Math.max(bottom, minRow);
  return { row: bottom < 0 ? 0 : bottom + 2, col: u.isNullObject ? 0 : u.columnIndex };
}

/*
 * Write the legend block (header + one row per tick mark) at pos and point the sheet name at it.
 * oldBox: where the legend is now when it stays in place (null for a new spot).
 */
async function writeLegend(ctx, ws, oldBox, entries, pos) {
  var row = pos.row, col = pos.col;
  var oldRows = oldBox ? oldBox.rowCount : 0;
  var newRows = 1 + entries.length;

  if (newRows > oldRows) {
    // Need more rows: if the cells below are in use, insert whole rows so nothing is overwritten.
    var extra = ws.getRangeByIndexes(row + oldRows, col, newRows - oldRows, 2);
    extra.load("values");
    await ctx.sync();
    var busy = extra.values.some(function (r) { return r.some(function (v) { return v !== ""; }); });
    if (busy) {
      if (!oldBox) throw new UserError("Pick a spot with " + newRows + " empty rows across 2 columns for the legend.");
      ws.getRangeByIndexes(row + oldRows, 0, newRows - oldRows, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
    }
  } else if (newRows < oldRows) {
    ws.getRangeByIndexes(row + newRows, col, oldRows - newRows, 2).clear(Excel.ClearApplyTo.all);
  }

  var values = [[LEGEND_HEADER, ""]];
  var formats = [["@", "@"]];
  entries.forEach(function (e) { values.push([TickCore.withCase(e.letter, lowerOn()), e.desc]); formats.push(["@", "@"]); });

  var block = ws.getRangeByIndexes(row, col, newRows, 2);
  block.numberFormat = formats;
  block.values = values;
  var header = ws.getRangeByIndexes(row, col, 1, 2);
  header.format.font.bold = true;
  header.format.font.color = "#000000";
  if (entries.length) {
    var letters = ws.getRangeByIndexes(row + 1, col, entries.length, 1);
    letters.format.font.color = TICK_COLOR;
    letters.format.font.bold = true;
    letters.format.horizontalAlignment = "Center";
  }

  var old = ws.names.getItemOrNullObject(LEGEND_NAME);
  await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  try {
    ws.names.add(LEGEND_NAME, block, "Tick Mark Manager legend - do not delete");
    await ctx.sync();
  } catch (e) {
    // Not fatal: the add-in also finds the legend by its "Tick Mark Legend" heading.
    console.warn("Could not name the legend range", e);
  }
  return { row: row, col: col };
}

function legendPosOf(data) {
  return data.legend ? { row: data.legend.box.rowIndex, col: data.legend.box.columnIndex } : null;
}

/* Apply a new legend order. kept: entries in new order (letter null = brand new). */
async function applyOrder(ctx, ws, data, kept, deleted) {
  await ensureStyle(ctx);
  var oldEntries = data.legend ? data.legend.entries : [];
  var plan = TickCore.buildPlan(oldEntries, kept, deleted || []);
  var clash = TickCore.orphanCollision(data.ticks, plan.mapping);
  if (clash.length) {
    throw new UserError("Can't re-letter yet: some tick mark cells use a letter that isn't in the legend (" +
      clash.slice(0, 5).map(function (t) { return TickCore.cellAddress(t.row, t.col) + " = " + t.letter; }).join(", ") +
      "). Fix or remove those cells, then try again.");
  }
  var pos = legendPosOf(data);
  applyChanges(ws, data, plan, pos);
  plan.refsUpdated = await updateMentions(ctx, ws, data.refCells, plan.mapping, plan.newEntries.map(function (e) { return e.letter; }));
  await writeLegend(ctx, ws, data.legend ? data.legend.box : null, plan.newEntries, pos);
  return plan;
}

/* --------------------------------------------------------- automatic mode */

/* Cells (row/col) inside the edited addresses, capped so a big paste stays fast. */
function editedCells(addrs) {
  var out = [], seen = {};
  (addrs || []).forEach(function (a) {
    var box = TickCore.parseA1Range(a);
    if (!box || box.rowCount * box.columnCount > MAX_EDIT_CELLS) return;
    for (var r = 0; r < box.rowCount; r++) for (var c = 0; c < box.columnCount; c++) {
      var k = (box.rowIndex + r) + "," + (box.columnIndex + c);
      if (!seen[k] && out.length < MAX_EDIT_CELLS) { seen[k] = true; out.push({ row: box.rowIndex + r, col: box.columnIndex + c }); }
    }
  });
  return out;
}

function autoSync(id, addrs) {
  return Excel.run(async function (ctx) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(id);
    await ctx.sync();
    if (ws.isNullObject) return;
    var prev = state[id] || null;
    var data = await readSheet(ctx, ws);
    if (!data.legend) {
      state[id] = null;
      await updateMentions(ctx, ws, data.refCells, null, [], editedCells(addrs)); // clear stale highlighting
      return;
    }

    var entries = data.legend.entries;
    if (prev && prev.letters.length && !entries.length) {
      // The whole legend was cleared: treat that as an accident, not "delete every tick mark".
      showStatus("The tick mark legend on this sheet is empty. Press Ctrl+Z if that was a mistake.", "warn");
      state[id] = { letters: [], counts: data.counts };
      return;
    }

    var ap = TickCore.autoPlan(entries, data.counts, prev);
    var plan = TickCore.buildPlan(entries, ap.kept, ap.deleted);
    if (plan.changed || needsFix(data)) {
      if (TickCore.orphanCollision(data.ticks, plan.mapping).length) {
        showStatus("Tick marks need re-lettering, but some tick cells use letters that aren't in the legend. Open Tick Marks to fix.", "warn");
      } else {
        await ensureStyle(ctx);
        var pos = legendPosOf(data);
        applyChanges(ws, data, plan, pos);
        var nRefs = await updateMentions(ctx, ws, data.refCells, plan.mapping, plan.newEntries.map(function (e) { return e.letter; }), editedCells(addrs));
        await writeLegend(ctx, ws, data.legend.box, plan.newEntries, pos);
        if (plan.changed) {
          var removed = entries.filter(function (e) { return plan.mapping[e.letter] === null; }).map(function (e) { return e.letter; });
          removed = removed.concat(ap.deleted);
          showStatus((removed.length
            ? "Removed tick mark " + removed.join(", ") + " and re-lettered the rest."
            : "Re-lettered tick marks.") + refsNote(nRefs), "ok");
        }
      }
    }
    else {
      // Nothing to re-letter: just make sure every mention is bold red (e.g. one just typed).
      await updateMentions(ctx, ws, data.refCells, null, entries.map(function (e) { return e.letter; }), editedCells(addrs));
    }
    await recordState(ctx, ws, id);
  });
}

/* ---------------------------------------------------------- pane actions */

function withActiveSheet(fn) {
  return enqueue(function () {
    return Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("id,name");
      await ctx.sync();
      var data = await readSheet(ctx, ws);
      var msg = await fn(ctx, ws, data);
      await ctx.sync();
      await recordState(ctx, ws, ws.id);
      return msg;
    });
  })
    .then(function (msg) { if (msg) showStatus(msg, "ok"); return enqueue(refreshUi); })
    .catch(showError);
}

async function selectionAreas(ctx) {
  var sel = ctx.workbook.getSelectedRanges();
  sel.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
  await ctx.sync();
  return sel.areas.items;
}

function checkAreas(areas, legendBox) {
  var total = areas.reduce(function (n, a) { return n + a.rowCount * a.columnCount; }, 0);
  if (total > MAX_FILL) throw new UserError("Select fewer cells (" + total + " selected).");
  if (legendBox && areas.some(function (a) { return boxesOverlap(a, legendBox); })) {
    throw new UserError("The selection overlaps the legend. Select the cells where the tick mark should go.");
  }
  return total;
}

/* Put a tick mark formula (=$A$14) in every selected cell. */
function fillAreas(areas, formula) {
  areas.forEach(function (a) {
    var f = [];
    for (var r = 0; r < a.rowCount; r++) {
      var row = [];
      for (var c = 0; c < a.columnCount; c++) row.push(formula);
      f.push(row);
    }
    var nf = f.map(function (row) { return row.map(function () { return "General"; }); });
    a.style = STYLE_TICK;
    a.numberFormat = nf;
    a.formulas = f;
  });
}

function refsNote(n) {
  return n ? " Updated " + n + " cell" + (n === 1 ? "" : "s") + " of text that mention tick marks." : "";
}

function boxesOverlap(a, b) {
  return a.rowIndex < b.rowIndex + b.rowCount && b.rowIndex < a.rowIndex + a.rowCount &&
    a.columnIndex < b.columnIndex + b.columnCount && b.columnIndex < a.columnIndex + a.columnCount;
}

function setLoadOnOpen() {
  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }
  } catch (e) { /* older Excel: pane must be opened for auto mode */ }
}

function plural(n, word) { return n + " " + word + (n === 1 ? "" : "s"); }

function addTick(desc) {
  return withActiveSheet(async function (ctx, ws, data) {
    await ensureStyle(ctx);
    var areas = await selectionAreas(ctx);
    var n = checkAreas(areas, data.legend && data.legend.box);
    var entries = data.legend ? data.legend.entries : [];
    var plan = TickCore.buildPlan(entries, entries.concat([{ letter: null, desc: desc }]), []);
    if (TickCore.orphanCollision(data.ticks, plan.mapping).length) {
      throw new UserError("Some tick mark cells use letters that aren't in the legend. Fix those first (see the warning below).");
    }
    var pos = legendPosOf(data);
    if (!pos) {
      var selBottom = Math.max.apply(null, areas.map(function (a) { return a.rowIndex + a.rowCount - 1; }));
      pos = await defaultLegendPos(ctx, ws, selBottom);
    }
    applyChanges(ws, data, plan, pos);
    await updateMentions(ctx, ws, data.refCells, plan.mapping, plan.newEntries.map(function (e) { return e.letter; }));
    await writeLegend(ctx, ws, data.legend ? data.legend.box : null, plan.newEntries, pos);
    var idx = plan.newEntries.length - 1;
    fillAreas(await selectionAreas(ctx), TickCore.refFormula(pos.row + 1 + idx, pos.col));
    setLoadOnOpen();
    return "Added tick mark " + plan.newEntries[idx].letter + " to " + plural(n, "cell") + ".";
  });
}

function placeTick(letter) {
  return withActiveSheet(async function (ctx, ws, data) {
    await ensureStyle(ctx);
    var entry = data.legend && data.legend.entries.filter(function (e) { return e.letter === letter; })[0];
    if (!entry) throw new UserError("Tick mark " + letter + " isn't in the legend.");
    var areas = await selectionAreas(ctx);
    var n = checkAreas(areas, data.legend.box);
    fillAreas(areas, TickCore.refFormula(entry.row, data.legend.box.columnIndex));
    setLoadOnOpen();
    return "Placed " + letter + " in " + plural(n, "cell") + ".";
  });
}

function removeFromSelection() {
  return withActiveSheet(async function (ctx, ws, data) {
    var areas = await selectionAreas(ctx);
    var hit = data.ticks.filter(function (t) {
      return areas.some(function (a) { return TickCore.inBox(t.row, t.col, a); });
    });
    if (!hit.length) throw new UserError("No tick marks in the selected cells.");
    hit.forEach(function (t) { ws.getCell(t.row, t.col).clear(Excel.ClearApplyTo.contents); });
    await ctx.sync();
    var msg = "Removed " + plural(hit.length, "tick mark") + ".";
    if (!autoOn() || !data.legend) return msg;
    // Drop any tick mark that no longer appears anywhere, then re-letter.
    var after = await readSheet(ctx, ws);
    var prev = { letters: data.legend.entries.map(function (e) { return e.letter; }), counts: data.counts };
    var ap = TickCore.autoPlan(after.legend.entries, after.counts, prev);
    var plan = await applyOrder(ctx, ws, after, ap.kept, ap.deleted);
    var gone = after.legend.entries.filter(function (e) { return plan.mapping[e.letter] === null; });
    return gone.length ? msg + " Tick mark " + gone.map(function (e) { return e.letter; }).join(", ") + " is no longer used, so the rest were re-lettered." : msg;
  });
}

function deleteTick(letter) {
  return withActiveSheet(async function (ctx, ws, data) {
    var kept = data.legend.entries.filter(function (e) { return e.letter !== letter; });
    var plan = await applyOrder(ctx, ws, data, kept, []);
    return "Deleted tick mark " + letter + " and re-lettered the rest." + refsNote(plan.refsUpdated);
  });
}

function moveTick(letter, dir) {
  return withActiveSheet(async function (ctx, ws, data) {
    var list = data.legend.entries.slice();
    var i = list.findIndex(function (e) { return e.letter === letter; });
    var j = i + dir;
    if (i < 0 || j < 0 || j >= list.length) return null;
    var tmp = list[i]; list[i] = list[j]; list[j] = tmp;
    var plan = await applyOrder(ctx, ws, data, list, []);
    return "Re-lettered tick marks." + refsNote(plan.refsUpdated);
  });
}

function editDesc(letter, desc) {
  return withActiveSheet(async function (ctx, ws, data) {
    var list = data.legend.entries.map(function (e) {
      return e.letter === letter ? { letter: e.letter, desc: desc, offset: e.offset } : e;
    });
    await applyOrder(ctx, ws, data, list, []);
    return null;
  });
}

function reletterByPosition() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No tick marks on this sheet.");
    var plan = await applyOrder(ctx, ws, data, TickCore.orderByPosition(data.legend.entries, data.ticks), []);
    return "Re-lettered in the order they appear (top to bottom, left to right)." + refsNote(plan.refsUpdated);
  });
}

function removeUnused() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No tick marks on this sheet.");
    var kept = data.legend.entries.filter(function (e) { return (data.counts[e.letter] || 0) > 0; });
    var n = data.legend.entries.length - kept.length;
    var plan = await applyOrder(ctx, ws, data, kept, []);
    return (n ? "Removed " + plural(n, "unused tick mark") + " and re-lettered." : "Re-lettered. No unused tick marks.") + refsNote(plan.refsUpdated);
  });
}

function moveLegend() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No legend on this sheet yet.");
    await ensureStyle(ctx);
    var cell = ctx.workbook.getActiveCell();
    cell.load("rowIndex,columnIndex");
    await ctx.sync();
    var box = data.legend.box;
    ws.getRangeByIndexes(box.rowIndex, box.columnIndex, box.rowCount, Math.max(2, box.columnCount)).clear(Excel.ClearApplyTo.all);
    await ctx.sync();
    // Same letters, new spot: every tick mark formula is re-pointed at the moved legend.
    var entries = data.legend.entries;
    var plan = TickCore.buildPlan(entries, entries, []);
    var pos = { row: cell.rowIndex, col: cell.columnIndex };
    var err = null;
    try {
      await writeLegend(ctx, ws, null, plan.newEntries, pos);
    } catch (e) {
      err = e; // new spot isn't empty: put it back where it was (those cells are empty now)
      pos = { row: box.rowIndex, col: box.columnIndex };
      await writeLegend(ctx, ws, null, plan.newEntries, pos);
    }
    data.ticks.forEach(function (t) { t.refRow = -1; }); // force every formula to be rewritten
    applyChanges(ws, data, plan, pos);
    if (err) throw err;
    return "Moved the legend. Tick marks still point to it.";
  });
}

/* ------------------------------------------- jump to the legend and back */

var lastJump = null; // { sheetId, row, col } of the cell we jumped from

/* Like Excel's Ctrl+[ (which Excel on the web doesn't support): go to the cell a formula points at. */
function goToLegend() {
  return enqueue(function () {
    return Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("id,name");
      var cell = ctx.workbook.getActiveCell();
      cell.load("formulas,rowIndex,columnIndex");
      await ctx.sync();
      var ref = TickCore.parseRef(cell.formulas[0][0], ws.name);
      if (!ref) throw new UserError("Select a tick mark (a cell with a formula like =$B$19) first.");
      lastJump = { sheetId: ws.id, row: cell.rowIndex, col: cell.columnIndex };
      ws.getCell(ref.row, ref.col).select();
      await ctx.sync();
      return "Jumped to " + TickCore.cellAddress(ref.row, ref.col) + ". Use Back to return.";
    });
  }).then(function (msg) { showStatus(msg, "ok"); }).catch(showError);
}

function goBack() {
  return enqueue(function () {
    return Excel.run(async function (ctx) {
      if (!lastJump) throw new UserError("Nothing to go back to yet.");
      var ws = ctx.workbook.worksheets.getItemOrNullObject(lastJump.sheetId);
      await ctx.sync();
      if (ws.isNullObject) throw new UserError("That sheet no longer exists.");
      ws.activate();
      ws.getCell(lastJump.row, lastJump.col).select();
      await ctx.sync();
      return "Back at " + TickCore.cellAddress(lastJump.row, lastJump.col) + ".";
    });
  }).then(function (msg) { showStatus(msg, "ok"); }).catch(showError);
}

/* Keyboard shortcuts (Alt+Shift+G and Alt+Shift+B) declared in shortcuts.json. */
function registerShortcuts() {
  try {
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("GoToLegend", function () { return goToLegend(); });
      Office.actions.associate("GoBack", function () { return goBack(); });
    }
  } catch (e) { /* older Excel: buttons in the pane still work */ }
}

/* --------------------------------------------------------------------- UI */

function $(id) { return document.getElementById(id); }

function wireUi() {
  if (!$("addBtn")) return;
  $("addBtn").onclick = function () {
    var d = $("newDesc").value.trim();
    if (!d) { showStatus("Type a description for the new tick mark first.", "warn"); $("newDesc").focus(); return; }
    addTick(d).then(function () { $("newDesc").value = ""; });
  };
  $("newDesc").onkeydown = function (e) { if (e.key === "Enter") $("addBtn").click(); };
  $("removeSelBtn").onclick = removeFromSelection;
  if ($("jumpBtn")) $("jumpBtn").onclick = goToLegend;
  if ($("backBtn")) $("backBtn").onclick = goBack;
  $("byPosBtn").onclick = reletterByPosition;
  $("unusedBtn").onclick = removeUnused;
  $("moveLegendBtn").onclick = moveLegend;
  $("refreshBtn").onclick = function () { enqueue(recordAllSheets).then(function () { return enqueue(refreshUi); }).catch(showError); };
  $("lowerChk").checked = lowerOn();
  $("lowerChk").onchange = function () {
    try { localStorage.setItem("tm.lower", this.checked ? "on" : "off"); } catch (e) { /* ignore */ }
    enqueue(refreshUi).catch(showError);
  };
  $("autoChk").checked = autoOn();
  $("autoChk").onchange = function () {
    try { localStorage.setItem("tm.auto", this.checked ? "on" : "off"); } catch (e) { /* ignore */ }
  };
}

/* ------------------------------------------------------- needs attention */

var IGNORE_KEY = "tm.ignored";

function ignoredList() {
  try {
    var s = Office.context.document.settings;
    var v = s && s.get(IGNORE_KEY);
    return Array.isArray(v) ? v : [];
  } catch (e) { return []; }
}

function ignoreIssue(issue) {
  try {
    var s = Office.context.document.settings;
    var list = ignoredList();
    list.push(issue.key);
    s.set(IGNORE_KEY, list.slice(-500));
    s.saveAsync(function () {});
  } catch (e) { /* settings unavailable: nothing to persist */ }
  return enqueue(refreshUi).catch(showError);
}

function snippet(text) {
  text = String(text).replace(/\s+/g, " ");
  return text.length > 70 ? text.slice(0, 67) + "…" : text;
}

/* Everything on one sheet that probably needs a person to look at it. */
function computeIssues(data, sheet) {
  var out = [];
  var entries = data.legend ? data.legend.entries : [];
  var known = {};
  entries.forEach(function (e) { known[e.letter] = true; });
  var isKnown = function (l) { return !!known[l]; };
  var where = function (row, col) { return { sheetId: sheet.id, sheetName: sheet.name, row: row, col: col, addr: TickCore.cellAddress(row, col) }; };
  var make = function (kind, row, col, extra) {
    var it = where(row, col);
    Object.keys(extra).forEach(function (k) { it[k] = extra[k]; });
    it.kind = kind;
    it.key = [sheet.id, kind, it.letter || "", extra.text || ""].join("|");
    return it;
  };

  data.refCells.forEach(function (rc) {
    if (/\[deleted\]/i.test(rc.text)) {
      out.push(make("deleted", rc.row, rc.col, {
        title: "Mentions a tick mark that was deleted",
        hint: "Rewrite the sentence; the yellow highlight goes away on its own.",
        text: snippet(rc.text)
      }));
    }
    TickCore.unknownMentions(rc.text, isKnown).forEach(function (L) {
      out.push(make("unknown", rc.row, rc.col, {
        letter: L,
        title: "Mentions tick mark " + TickCore.withCase(L, lowerOn()) + ", which isn't in the legend",
        hint: "It isn't shown in bold red. Fix the letter, add the tick mark, or ignore this.",
        text: snippet(rc.text)
      }));
    });
  });

  var mapping = {};
  entries.forEach(function (e) { mapping[e.letter] = e.letter; });
  TickCore.orphans(data.ticks, mapping).forEach(function (tk) {
    out.push(make("orphan", tk.row, tk.col, {
      letter: tk.letter,
      title: "Tick mark cell shows " + tk.letter + ", but there's no " + tk.letter + " in the legend",
      hint: "Place a legend tick mark here instead, or remove it.",
      text: ""
    }));
  });

  entries.forEach(function (e) {
    if ((data.counts[e.letter] || 0) === 0) {
      out.push(make("unused", e.row, data.legend.box.columnIndex, {
        letter: e.letter,
        title: "Tick mark " + TickCore.withCase(e.letter, lowerOn()) + " is in the legend but not used",
        hint: "Place it on the work paper, or remove it from the legend.",
        text: snippet(e.desc)
      }));
    }
  });

  var ignored = {};
  ignoredList().forEach(function (k) { ignored[k] = true; });
  return out.filter(function (it) { return !ignored[it.key]; });
}

function goToIssue(it) {
  return enqueue(function () {
    return Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(it.sheetId);
      await ctx.sync();
      if (ws.isNullObject) throw new UserError("That sheet no longer exists.");
      ws.activate();
      ws.getCell(it.row, it.col).select();
      await ctx.sync();
    });
  }).catch(showError);
}

function removeOrphan(it) {
  return enqueue(function () {
    return Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(it.sheetId);
      ws.getCell(it.row, it.col).clear(Excel.ClearApplyTo.contents);
      await ctx.sync();
    });
  }).then(function () { showStatus("Removed the tick mark in " + it.addr + ".", "ok"); return enqueue(refreshUi); }).catch(showError);
}

function removeUnusedTick(it) {
  return goToIssue(it).then(function () { return deleteTick(it.letter); });
}

function refreshUi() {
  if (!$("list")) return Promise.resolve();
  return Excel.run(async function (ctx) {
    var active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("id,name");
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/id,items/name");
    await ctx.sync();
    var activeData = null, issues = [];
    for (var i = 0; i < sheets.items.length; i++) {
      var s = sheets.items[i];
      var data = await readSheet(ctx, s);
      if (s.id === active.id) activeData = data;
      issues = issues.concat(computeIssues(data, s));
    }
    if (!activeData) activeData = await readSheet(ctx, active);
    // Active sheet first, then the others in tab order.
    issues.sort(function (a, b) { return (a.sheetId === active.id ? 0 : 1) - (b.sheetId === active.id ? 0 : 1); });
    render(active.name, activeData, issues);
  });
}

function renderIssues(issues) {
  var box = $("attention");
  if (!box) return;
  var list = $("attnList");
  list.innerHTML = "";
  $("attnCount").textContent = issues.length ? String(issues.length) : "";
  $("attnCount").className = "attn-count" + (issues.length ? " on" : "");
  $("attnEmpty").style.display = issues.length ? "none" : "block";
  issues.slice(0, 100).forEach(function (it) {
    var row = document.createElement("div");
    row.className = "issue " + it.kind;
    var head = document.createElement("div");
    head.className = "issue-title";
    head.textContent = it.title;
    row.appendChild(head);
    var loc = document.createElement("div");
    loc.className = "issue-where";
    loc.textContent = it.sheetName + "!" + it.addr + (it.text ? " — “" + it.text + "”" : "");
    row.appendChild(loc);
    var hint = document.createElement("div");
    hint.className = "issue-hint";
    hint.textContent = it.hint;
    row.appendChild(hint);
    var btns = document.createElement("div");
    btns.className = "issue-btns";
    btns.appendChild(button("Go to " + it.addr, "Select this cell", function () { goToIssue(it); }));
    if (it.kind === "orphan") btns.appendChild(button("Remove", "Clear this tick mark cell", function () { removeOrphan(it); }));
    if (it.kind === "unused") btns.appendChild(button("Remove from legend", "Delete this tick mark and re-letter", function () { removeUnusedTick(it); }));
    if (it.kind === "unknown" || it.kind === "unused") btns.appendChild(button("Ignore", "Hide this from the list", function () { ignoreIssue(it); }));
    row.appendChild(btns);
    list.appendChild(row);
  });
  if (issues.length > 100) {
    var more = document.createElement("div");
    more.className = "sub";
    more.textContent = "And " + (issues.length - 100) + " more.";
    list.appendChild(more);
  }
}

function render(sheetName, data, issues) {
  renderIssues(issues || []);
  $("sheetName").textContent = "Sheet: " + sheetName;
  var list = $("list");
  list.innerHTML = "";
  var entries = data.legend ? data.legend.entries : [];
  $("empty").style.display = entries.length ? "none" : "block";

  entries.forEach(function (e, i) {
    var n = data.counts[e.letter] || 0;
    var row = document.createElement("div");
    row.className = "entry" + (n ? "" : " unused");

    var badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = TickCore.withCase(e.letter, lowerOn());
    row.appendChild(badge);

    var mid = document.createElement("div");
    mid.className = "mid";
    var input = document.createElement("input");
    input.value = e.desc;
    input.title = "Edit description";
    input.onchange = function () { editDesc(e.letter, input.value); };
    mid.appendChild(input);
    var count = document.createElement("div");
    count.className = "count";
    var refs = data.refCounts[e.letter] || 0;
    count.textContent = (n ? "Used in " + n + " cell" + (n === 1 ? "" : "s") : "Not used on this sheet") +
      (refs ? " \u00b7 mentioned " + refs + " time" + (refs === 1 ? "" : "s") + " in text" : "");
    mid.appendChild(count);
    row.appendChild(mid);

    var btns = document.createElement("div");
    btns.className = "btns";
    btns.appendChild(button("Place", "Put " + e.letter + " in the selected cells", function () { placeTick(e.letter); }));
    btns.appendChild(button("↑", "Move up", function () { moveTick(e.letter, -1); }, i === 0));
    btns.appendChild(button("↓", "Move down", function () { moveTick(e.letter, 1); }, i === entries.length - 1));
    var del = button("✕", "Delete " + e.letter + " everywhere and re-letter", null);
    del.className = "del";
    var armed = false;
    del.onclick = function () {
      if (armed) { deleteTick(e.letter); return; }
      armed = true;
      del.textContent = "Delete?";
      setTimeout(function () { armed = false; del.textContent = "✕"; }, 3000);
    };
    btns.appendChild(del);
    row.appendChild(btns);
    list.appendChild(row);
  });

}

function button(text, title, onclick, disabled) {
  var b = document.createElement("button");
  b.textContent = text;
  b.title = title;
  if (onclick) b.onclick = onclick;
  if (disabled) b.disabled = true;
  return b;
}

function showStatus(msg, kind) {
  var el = $("status");
  if (!el) return;
  el.textContent = msg;
  el.className = kind || "";
}

function showError(err) {
  if (err && err.isUser) showStatus(err.message, "warn");
  else { showStatus("Something went wrong: " + (err && err.message ? err.message : err), "error"); console.error(err); }
}
