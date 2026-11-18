/*
 * Tick Mark Manager - Excel add-in (task pane + shared runtime).
 * Tick mark cells are tagged with the cell style "Tick Mark".
 * Each worksheet has its own legend, found through the sheet-level name "TickMarkLegend".
 */
/* global Office, Excel, TickCore */
"use strict";

var STYLE_TICK = "Tick Mark";
var LEGEND_NAME = "TickMarkLegend";
var LEGEND_HEADER = "Tick Mark Legend";
var TICK_COLOR = "#C00000";
var MAX_FILL = 5000;
var DEBOUNCE_MS = 700;
var DELETED_FILL = "#FFFF00";

var state = {};     // worksheetId -> { letters: [...], counts: {...} } from the last time we looked
var timers = {};
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
  if (info.host !== Office.HostType.Excel) return;
  wireUi();
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
    for (var i = 0; i < sheets.items.length; i++) {
      await recordState(ctx, sheets.items[i], sheets.items[i].id);
    }
  });
}

function onSheetChanged(evt) {
  if (evt.source && evt.source !== Excel.EventSource.local) return Promise.resolve();
  var id = evt.worksheetId;
  clearTimeout(timers[id]);
  timers[id] = setTimeout(function () {
    enqueue(function () { return autoOn() ? autoSync(id) : null; })
      .then(function () { return enqueue(refreshUi); })
      .catch(showError);
  }, DEBOUNCE_MS);
  return Promise.resolve();
}

/* --------------------------------------------------------- reading a sheet */

async function readSheet(ctx, ws) {
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
      }
    }
  } catch (e) {
    legend = null; // name points at deleted cells (#REF!)
  }

  var ticks = [];
  var used = ws.getUsedRangeOrNullObject(true);
  used.load("values,formulas,rowIndex,columnIndex");
  await ctx.sync();
  var refCells = [];
  if (!used.isNullObject) {
    refCells = TickCore.findRefCells(used.values, used.formulas, used.rowIndex, used.columnIndex, legend && legend.box);
    var cands = TickCore.findCandidates(used.values, used.rowIndex, used.columnIndex, legend && legend.box);
    var cells = cands.map(function (c) {
      var cell = ws.getCell(c.row, c.col);
      cell.load("style");
      return { c: c, cell: cell };
    });
    if (cells.length) await ctx.sync();
    cells.forEach(function (x) { if (x.cell.style === STYLE_TICK) ticks.push(x.c); });
  }
  var letters = legend ? legend.entries.map(function (e) { return e.letter; }) : [];
  return {
    legend: legend, ticks: ticks, counts: TickCore.countTicks(ticks),
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
    await ctx.sync();
  }
}

function applyMappingToCells(ws, data, mapping) {
  data.ticks.forEach(function (t) {
    if (!Object.prototype.hasOwnProperty.call(mapping, t.letter)) return; // not in legend: leave alone
    var nl = mapping[t.letter];
    if (nl === t.letter) return;
    var cell = ws.getCell(t.row, t.col);
    if (nl === null) cell.clear(Excel.ClearApplyTo.contents);
    else cell.values = [[TickCore.withCase(nl, t.lower)]];
  });
  // Paragraphs that mention tick marks ("see tick mark B", "TM {b}").
  var n = 0;
  data.refCells.forEach(function (rc) {
    var r = TickCore.remapReferences(rc.text, mapping);
    if (!r.changed) return;
    var cell = ws.getCell(rc.row, rc.col);
    cell.values = [[r.text]];
    if (r.deleted) cell.format.fill.color = DELETED_FILL;
    n++;
  });
  return n;
}

/*
 * Write the legend block (header + one row per tick mark) and point the sheet name at it.
 * oldBox: where the legend is now (null if none). anchor: optional {row, col} to move it to.
 */
async function writeLegend(ctx, ws, oldBox, entries, anchor) {
  var row, col, oldRows;
  if (anchor) {
    row = anchor.row; col = anchor.col; oldRows = 0;
  } else if (oldBox) {
    row = oldBox.rowIndex; col = oldBox.columnIndex; oldRows = oldBox.rowCount;
  } else {
    var u = ws.getUsedRangeOrNullObject(true);
    u.load("rowIndex,rowCount,columnIndex");
    await ctx.sync();
    row = u.isNullObject ? 0 : u.rowIndex + u.rowCount + 1;
    col = u.isNullObject ? 0 : u.columnIndex;
    oldRows = 0;
  }
  var newRows = 1 + entries.length;

  if (newRows > oldRows) {
    // Need more rows: if the cells below are in use, insert whole rows so nothing is overwritten.
    var extra = ws.getRangeByIndexes(row + oldRows, col, newRows - oldRows, 2);
    extra.load("values");
    await ctx.sync();
    var busy = extra.values.some(function (r) { return r.some(function (v) { return v !== ""; }); });
    if (busy) {
      if (anchor) throw new UserError("Pick a spot with " + newRows + " empty rows across 2 columns for the legend.");
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
  ws.names.add(LEGEND_NAME, block, "Tick Mark Manager legend - do not delete");
  await ctx.sync();
}

/* Apply a new legend order. kept: entries in new order (letter null = brand new). */
async function applyOrder(ctx, ws, data, kept, deleted) {
  var oldEntries = data.legend ? data.legend.entries : [];
  var plan = TickCore.buildPlan(oldEntries, kept, deleted || []);
  var clash = TickCore.orphanCollision(data.ticks, plan.mapping);
  if (clash.length) {
    throw new UserError("Can't re-letter yet: some tick mark cells use a letter that isn't in the legend (" +
      clash.slice(0, 5).map(function (t) { return TickCore.cellAddress(t.row, t.col) + " = " + t.letter; }).join(", ") +
      "). Fix or remove those cells, then try again.");
  }
  plan.refsUpdated = applyMappingToCells(ws, data, plan.mapping);
  await writeLegend(ctx, ws, data.legend ? data.legend.box : null, plan.newEntries);
  return plan;
}

/* --------------------------------------------------------- automatic mode */

function autoSync(id) {
  return Excel.run(async function (ctx) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(id);
    await ctx.sync();
    if (ws.isNullObject) return;
    var prev = state[id] || null;
    var data = await readSheet(ctx, ws);
    if (!data.legend) { state[id] = null; return; }

    var entries = data.legend.entries;
    if (prev && prev.letters.length && !entries.length) {
      // The whole legend was cleared: treat that as an accident, not "delete every tick mark".
      showStatus("The tick mark legend on this sheet is empty. Press Ctrl+Z if that was a mistake.", "warn");
      state[id] = { letters: [], counts: data.counts };
      return;
    }

    var ap = TickCore.autoPlan(entries, data.counts, prev);
    var plan = TickCore.buildPlan(entries, ap.kept, ap.deleted);
    if (plan.changed) {
      if (TickCore.orphanCollision(data.ticks, plan.mapping).length) {
        showStatus("Tick marks need re-lettering, but some tick cells use letters that aren't in the legend. Open Tick Marks to fix.", "warn");
      } else {
        var nRefs = applyMappingToCells(ws, data, plan.mapping);
        await writeLegend(ctx, ws, data.legend.box, plan.newEntries);
        var removed = entries.filter(function (e) { return plan.mapping[e.letter] === null; }).map(function (e) { return e.letter; });
        removed = removed.concat(ap.deleted);
        showStatus((removed.length
          ? "Removed tick mark " + removed.join(", ") + " and re-lettered the rest."
          : "Re-lettered tick marks.") + refsNote(nRefs), "ok");
      }
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

async function fillSelection(ctx, letter, legendBox) {
  letter = TickCore.withCase(letter, lowerOn());
  var areas = await selectionAreas(ctx);
  var total = areas.reduce(function (n, a) { return n + a.rowCount * a.columnCount; }, 0);
  if (total > MAX_FILL) throw new UserError("Select fewer cells (" + total + " selected).");
  if (legendBox && areas.some(function (a) { return boxesOverlap(a, legendBox); })) {
    throw new UserError("The selection overlaps the legend. Select the cells where the tick mark should go.");
  }
  areas.forEach(function (a) {
    var vals = [];
    for (var r = 0; r < a.rowCount; r++) {
      var row = [];
      for (var c = 0; c < a.columnCount; c++) row.push(letter);
      vals.push(row);
    }
    a.style = STYLE_TICK;
    a.values = vals;
  });
  return total;
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

function addTick(desc) {
  return withActiveSheet(async function (ctx, ws, data) {
    await ensureStyle(ctx);
    var entries = data.legend ? data.legend.entries : [];
    var kept = entries.concat([{ letter: null, desc: desc }]);
    var oldEntries = entries;
    var plan = TickCore.buildPlan(oldEntries, kept, []);
    if (TickCore.orphanCollision(data.ticks, plan.mapping).length) {
      throw new UserError("Some tick mark cells use letters that aren't in the legend. Fix those first (see the warning below).");
    }
    var letter = plan.newEntries[plan.newEntries.length - 1].letter;
    applyMappingToCells(ws, data, plan.mapping);
    var n = await fillSelection(ctx, letter, data.legend && data.legend.box);
    await writeLegend(ctx, ws, data.legend ? data.legend.box : null, plan.newEntries);
    setLoadOnOpen();
    return "Added tick mark " + letter + " to " + n + " cell" + (n === 1 ? "" : "s") + ".";
  });
}

function placeTick(letter) {
  return withActiveSheet(async function (ctx, ws, data) {
    await ensureStyle(ctx);
    var n = await fillSelection(ctx, letter, data.legend && data.legend.box);
    setLoadOnOpen();
    return "Placed " + letter + " in " + n + " cell" + (n === 1 ? "" : "s") + ".";
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
    var msg = "Removed " + hit.length + " tick mark" + (hit.length === 1 ? "" : "s") + ".";
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
    await applyOrder(ctx, ws, data, list, []);
    return "Re-lettered tick marks.";
  });
}

function editDesc(letter, desc) {
  return withActiveSheet(async function (ctx, ws, data) {
    var list = data.legend.entries.map(function (e) {
      return e.letter === letter ? { letter: e.letter, desc: desc } : e;
    });
    await applyOrder(ctx, ws, data, list, []);
    return null;
  });
}

function reletterByPosition() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No tick marks on this sheet.");
    await applyOrder(ctx, ws, data, TickCore.orderByPosition(data.legend.entries, data.ticks), []);
    return "Re-lettered in the order they appear (top to bottom, left to right).";
  });
}

function removeUnused() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No tick marks on this sheet.");
    var kept = data.legend.entries.filter(function (e) { return (data.counts[e.letter] || 0) > 0; });
    var n = data.legend.entries.length - kept.length;
    await applyOrder(ctx, ws, data, kept, []);
    return n ? "Removed " + n + " unused tick mark" + (n === 1 ? "" : "s") + " and re-lettered." : "Re-lettered. No unused tick marks.";
  });
}

function moveLegend() {
  return withActiveSheet(async function (ctx, ws, data) {
    if (!data.legend) throw new UserError("No legend on this sheet yet.");
    var cell = ctx.workbook.getActiveCell();
    cell.load("rowIndex,columnIndex");
    await ctx.sync();
    var box = data.legend.box;
    var oldRange = ws.getRangeByIndexes(box.rowIndex, box.columnIndex, box.rowCount, Math.max(2, box.columnCount));
    oldRange.clear(Excel.ClearApplyTo.all);
    await ctx.sync();
    // Keep the current letters: moving the legend never re-letters anything.
    var entries = data.legend.entries.map(function (e) { return { letter: e.letter, desc: e.desc }; });
    try {
      await writeLegend(ctx, ws, null, entries, { row: cell.rowIndex, col: cell.columnIndex });
    } catch (e) {
      // New spot isn't empty: put the legend back where it was (those cells are empty now).
      await writeLegend(ctx, ws, null, entries, { row: box.rowIndex, col: box.columnIndex });
      throw e;
    }
    return "Moved the legend.";
  });
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

function refreshUi() {
  if (!$("list")) return Promise.resolve();
  return Excel.run(async function (ctx) {
    var ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    await ctx.sync();
    var data = await readSheet(ctx, ws);
    render(ws.name, data);
  });
}

function render(sheetName, data) {
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

  var mapping = {};
  entries.forEach(function (e) { mapping[e.letter] = e.letter; });
  var orphans = TickCore.orphans(data.ticks, mapping);
  var w = $("warnings");
  if (orphans.length) {
    w.style.display = "block";
    w.textContent = "These tick mark cells use a letter that isn't in the legend: " +
      orphans.slice(0, 10).map(function (t) { return TickCore.cellAddress(t.row, t.col) + " (" + t.letter + ")"; }).join(", ") +
      (orphans.length > 10 ? " and " + (orphans.length - 10) + " more" : "") +
      ". Change them to a legend letter or remove them.";
  } else {
    w.style.display = "none";
  }
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
