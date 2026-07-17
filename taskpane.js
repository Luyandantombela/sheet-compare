/* ============================================================
   Sheet Compare — Task Pane Logic  (Dark UI, progressive flow)
   ============================================================ */
"use strict";

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  rangeA:     null,   // { address, sheetName, headers, data }
  rangeB:     null,
  lastResult: null,   // stored after a run, used by Highlight + Export
};

// ── DOM helper ────────────────────────────────────────────────────────────────
function $(id) { return document.getElementById(id); }

// ── Office init ───────────────────────────────────────────────────────────────
Office.onReady(function () {
  // Range capture
  $("btnSelectA").onclick  = () => captureRange("A");
  $("btnSelectB").onclick  = () => captureRange("B");
  $("btnReA").onclick      = () => resetCapture("A");
  $("btnReB").onclick      = () => resetCapture("B");

  // Key columns change → refresh diff columns & run button
  $("keyColumns").onchange = () => { rebuildDiffColumns(); refreshRunButton(); };

  // Mode chips → show/hide diff columns panel, refresh run button
  $("modeMissing").onchange   = refreshRunButton;
  $("modeMatching").onchange  = refreshRunButton;
  $("modeDifferent").onchange = () => { toggleDiffCols(); refreshRunButton(); };

  // Options accordion
  $("opts-toggle").onclick = toggleOpts;
  $("optNumericTol").onchange = () => {
    $("tolWrap").style.display = $("optNumericTol").checked ? "" : "none";
  };

  // Run + post-run actions
  $("btnRun").onclick       = runComparison;
  $("btnHighlight").onclick = highlightSources;
  $("btnExportCsv").onclick = exportCSV;

  // Modal
  $("btnModalCancel").onclick  = () => resolveModal(false);
  $("btnModalProceed").onclick = () => resolveModal(true);

  // Init UI state
  setStepState("b",     false);
  setStepState("keys",  false);
  setStepState("modes", false);
  setStepState("opts",  false);
  toggleDiffCols();
});

// ── Step unlock helpers ───────────────────────────────────────────────────────
function setStepState(which, enabled) {
  const el  = $("step-" + which);
  const num = $("num-"  + which);
  if (!el) return;
  if (enabled) {
    el.classList.add("unlocked");
    el.classList.remove("step-locked");
  } else {
    el.classList.remove("unlocked");
    el.classList.add("step-locked");
  }
  // Number bubble
  if (num) {
    num.classList.toggle("active", enabled);
  }
}

function markStepDone(which) {
  const num = $("num-" + which);
  if (num) { num.classList.remove("active"); num.classList.add("done"); num.textContent = "✓"; }
}

// ── Range capture ─────────────────────────────────────────────────────────────
async function captureRange(which) {
  const btn = $("btnSelect" + which);
  btn.disabled = true;
  clearMessages();

  try {
    const result = await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["address", "values"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const values = range.values;
      if (!values || values.length === 0)   throw new Error("Selected range is empty.");
      if (values.length < 2)                throw new Error("Need at least one header row + one data row.");

      const headers = values[0].map((h) => String(h ?? "").trim());
      if (headers.some((h) => h === ""))    throw new Error("Header row has blank cells — fill all column headers first.");

      return { address: range.address, sheetName: ws.name, headers, data: values.slice(1) };
    });

    state["range" + which] = result;
    showCaptureDone(which, result);
    onCaptureComplete();

  } catch (err) {
    btn.disabled = false;
    showMsg("Table " + which + ": " + friendlyError(err), "error");
  }
}

function showCaptureDone(which, result) {
  // FIX: use lowercase for element IDs (idle-a, done-a, done-a-name, done-a-meta)
  const lc = which.toLowerCase();
  $("idle-" + lc).style.display = "none";
  $("done-" + lc).style.display = "flex";

  // Show info
  const addr = result.address.split("!").pop().replace(/\$/g, "");
  $("done-" + lc + "-name").textContent =
    result.sheetName + "  ·  " + addr;
  $("done-" + lc + "-meta").textContent =
    result.headers.length + " columns  ·  " + result.data.length + " rows";

  markStepDone(lc);
}

function resetCapture(which) {
  state["range" + which] = null;
  // FIX: use lowercase for element IDs
  const lc = which.toLowerCase();
  $("idle-" + lc).style.display = "";
  $("done-" + lc).style.display = "none";
  $("btnSelect" + which).disabled = false;

  const num = $("num-" + lc);
  if (num) { num.classList.remove("done"); num.classList.add("active"); num.textContent = which === "A" ? "1" : "2"; }

  // Reset downstream steps
  clearMessages();
  rebuildKeyColumns();
  onCaptureComplete();
}

function onCaptureComplete() {
  const hasA = !!state.rangeA;
  const hasB = !!state.rangeB;

  // Unlock B after A is captured
  setStepState("b", hasA);
  if (hasA) $("btnSelectB").disabled = false;

  // Unlock keys after both captured
  const bothDone = hasA && hasB;
  setStepState("keys",  bothDone);
  setStepState("modes", bothDone);
  setStepState("opts",  bothDone);

  if (bothDone) {
    rebuildKeyColumns();
    rebuildDiffColumns();
    refreshRunButton();
  } else {
    $("run-area").style.display = "none";
    resetKeySelect();
  }
}

// ── Column selects ────────────────────────────────────────────────────────────
function resetKeySelect() {
  const sel = $("keyColumns");
  sel.innerHTML = '<option disabled>— select both tables first —</option>';
}

function rebuildKeyColumns() {
  const sel = $("keyColumns");
  sel.innerHTML = "";

  if (!state.rangeA || !state.rangeB) { resetKeySelect(); return; }

  const common = state.rangeA.headers.filter((h) => state.rangeB.headers.includes(h));
  if (common.length === 0) {
    sel.innerHTML = '<option disabled>— no common column names —</option>';
    showMsg("Tables share no common column headers — key selection is impossible.", "warning");
    return;
  }

  common.forEach((h) => {
    const o = document.createElement("option");
    o.value = h; o.textContent = h;
    sel.appendChild(o);
  });
  rebuildDiffColumns();
}

function rebuildDiffColumns() {
  const sel  = $("diffColumns");
  sel.innerHTML = "";

  const keys = selectedKeys();
  if (!state.rangeA || !state.rangeB || keys.length === 0) {
    sel.innerHTML = '<option disabled>— select key columns first —</option>';
    return;
  }

  const candidates = [...new Set([
    ...state.rangeA.headers.filter((h) => !keys.includes(h)),
    ...state.rangeB.headers.filter((h) => !keys.includes(h)),
  ])];

  if (candidates.length === 0) {
    sel.innerHTML = '<option disabled>— all columns are keys —</option>';
    return;
  }

  candidates.forEach((h) => {
    const o = document.createElement("option");
    o.value = h; o.textContent = h;
    sel.appendChild(o);
  });
}

function toggleDiffCols() {
  const show = $("modeDifferent").checked;
  $("diffColsReveal").style.display = show ? "" : "none";
}

function refreshRunButton() {
  const keys    = selectedKeys();
  const anyMode = $("modeMissing").checked || $("modeMatching").checked || $("modeDifferent").checked;
  const ready   = state.rangeA && state.rangeB && keys.length > 0 && anyMode;
  $("run-area").style.display = ready ? "" : "none";
}

function selectedKeys()     { return Array.from($("keyColumns").selectedOptions).map((o) => o.value); }
function selectedDiffCols() { return Array.from($("diffColumns").selectedOptions).map((o) => o.value); }

// ── Options accordion ─────────────────────────────────────────────────────────
function toggleOpts() {
  const body = $("opts-body");
  const chev = $("opts-chevron");
  const open = body.classList.toggle("open");
  chev.classList.toggle("open", open);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function yieldToUI()      { return new Promise((r) => setTimeout(r, 0)); }
function clearMessages()  { $("messages").innerHTML = ""; }

function showMsg(text, type) {
  const el = document.createElement("div");
  el.className = "msg msg-" + (type || "info");
  el.textContent = text;
  $("messages").appendChild(el);
  el.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function friendlyError(err) {
  if (typeof OfficeExtension !== "undefined" && err instanceof OfficeExtension.Error) return err.message;
  if (err instanceof Error) return err.message;
  return String(err);
}

function setBtnRunLoading(loading) {
  const btn = $("btnRun");
  btn.innerHTML = loading
    ? '<span class="spinner"></span>  Running…'
    : "▶  Run Comparison";
  btn.disabled = loading;
}

function parseRangeTopLeft(address) {
  const clean   = address.replace(/\$/g, "").split("!").pop().split(":")[0];
  const letters = clean.replace(/[^A-Za-z]/g, "").toUpperCase();
  const digits  = clean.replace(/[^0-9]/g, "");
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return { col: col - 1, row: parseInt(digits, 10) - 1 };
}

// ── Modal ─────────────────────────────────────────────────────────────────────
let _modalResolve = null;
function showModal(title, body) {
  $("modal-title").textContent = title;
  $("modal-body").textContent  = body;
  $("modal-overlay").style.display = "flex";
  return new Promise((r) => { _modalResolve = r; });
}
function resolveModal(v) {
  $("modal-overlay").style.display = "none";
  if (_modalResolve) { _modalResolve(v); _modalResolve = null; }
}

// ── Comparison core ───────────────────────────────────────────────────────────
function normalizeVal(val, ignoreCase) {
  let s = val == null ? "" : String(val);
  if (ignoreCase) s = s.trim().toLowerCase();
  return s;
}
function numericClose(a, b, tol) {
  const na = parseFloat(a), nb = parseFloat(b);
  if (isNaN(na) || isNaN(nb)) return false;
  return Math.abs(na - nb) <= tol;
}
function makeKey(row, indices, ignoreCase) {
  return indices.map((i) => normalizeVal(row[i], ignoreCase)).join("\u0000");
}
function buildMap(data, indices, ignoreCase) {
  const map = new Map();
  data.forEach((row, idx) => {
    const k = makeKey(row, indices, ignoreCase);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push({ row, idx });
  });
  return map;
}
function findDuplicates(data, indices, ignoreCase) {
  const counts = new Map();
  data.forEach((row) => {
    const k = makeKey(row, indices, ignoreCase);
    counts.set(k, (counts.get(k) || 0) + 1);
  });
  return [...counts.entries()].filter(([, c]) => c > 1)
    .map(([k]) => k.replace(/\u0000/g, " | "));
}
function computeDiffs(rowA, rowB, headersA, headersB, diffCols, ignoreCase, numericTol) {
  const diffs = [];
  diffCols.forEach((col) => {
    const iA = headersA.indexOf(col); const iB = headersB.indexOf(col);
    if (iA < 0 || iB < 0) return;
    const vA = rowA[iA]; const vB = rowB[iB];
    let equal = normalizeVal(vA, ignoreCase) === normalizeVal(vB, ignoreCase);
    if (!equal && numericTol != null) equal = numericClose(vA, vB, numericTol);
    if (!equal) diffs.push({ col, vA, vB });
  });
  return diffs;
}
async function runCompareBatched(mapA, mapB, headersA, headersB, keys, diffCols, ignoreCase, numericTol, modes) {
  const BATCH = 500;
  const missingFromB = [], missingFromA = [], matching = [], different = [];
  const keysA = [...mapA.keys()];
  const keysB = [...mapB.keys()];

  for (let i = 0; i < keysA.length; i++) {
    if (i > 0 && i % BATCH === 0) await yieldToUI();
    const k = keysA[i]; const entriesA = mapA.get(k);
    if (mapB.has(k)) {
      const rowA = entriesA[0].row; const rowB = mapB.get(k)[0].row;
      if (modes.matching) matching.push({ keyStr: k, rowA, rowB });
      if (modes.different) {
        const d = computeDiffs(rowA, rowB, headersA, headersB, diffCols, ignoreCase, numericTol);
        if (d.length > 0) different.push({ keyStr: k, rowA, rowB, diffs: d });
      }
    } else if (modes.missing) {
      entriesA.forEach(({ row }) => missingFromB.push({ keyStr: k, row }));
    }
  }
  if (modes.missing) {
    for (let i = 0; i < keysB.length; i++) {
      if (i > 0 && i % BATCH === 0) await yieldToUI();
      const k = keysB[i];
      if (!mapA.has(k)) mapB.get(k).forEach(({ row }) => missingFromA.push({ keyStr: k, row }));
    }
  }
  return { missingFromB, missingFromA, matching, different };
}

// ── Run Comparison ────────────────────────────────────────────────────────────
async function runComparison() {
  clearMessages();

  const keys       = selectedKeys();
  const ignoreCase = $("optIgnoreCase").checked;
  const numericTol = $("optNumericTol").checked ? (parseFloat($("tolValue").value) || 0) : null;
  const modes      = {
    missing:   $("modeMissing").checked,
    matching:  $("modeMatching").checked,
    different: $("modeDifferent").checked,
  };
  const headersA = state.rangeA.headers;
  const headersB = state.rangeB.headers;
  const keyIdxA  = keys.map((k) => headersA.indexOf(k));
  const keyIdxB  = keys.map((k) => headersB.indexOf(k));

  if (keyIdxA.some((i) => i < 0)) { showMsg("Key column not found in Table A.", "error"); return; }
  if (keyIdxB.some((i) => i < 0)) { showMsg("Key column not found in Table B.", "error"); return; }

  // Blank key check
  for (const row of state.rangeA.data) {
    if (keyIdxA.some((i) => row[i] == null || String(row[i]).trim() === "")) {
      showMsg("Table A has rows with blank key values.", "error"); return;
    }
  }
  for (const row of state.rangeB.data) {
    if (keyIdxB.some((i) => row[i] == null || String(row[i]).trim() === "")) {
      showMsg("Table B has rows with blank key values.", "error"); return;
    }
  }

  // Duplicate key warning
  const dupA = findDuplicates(state.rangeA.data, keyIdxA, ignoreCase);
  const dupB = findDuplicates(state.rangeB.data, keyIdxB, ignoreCase);
  if (dupA.length > 0 || dupB.length > 0) {
    let msg = "";
    if (dupA.length) msg += "Table A:\n" + dupA.slice(0, 8).join("\n") + (dupA.length > 8 ? "\n…+" + (dupA.length - 8) + " more" : "") + "\n\n";
    if (dupB.length) msg += "Table B:\n" + dupB.slice(0, 8).join("\n") + (dupB.length > 8 ? "\n…+" + (dupB.length - 8) + " more" : "");
    const ok = await showModal("Duplicate Keys Detected", msg.trim() + "\n\nFirst occurrence used for matching. Proceed?");
    if (!ok) return;
  }

  let diffCols = selectedDiffCols();
  if (diffCols.length === 0) {
    diffCols = [...new Set([
      ...headersA.filter((h) => !keys.includes(h)),
      ...headersB.filter((h) => !keys.includes(h)),
    ])];
  }

  const mapA = buildMap(state.rangeA.data, keyIdxA, ignoreCase);
  const mapB = buildMap(state.rangeB.data, keyIdxB, ignoreCase);

  setBtnRunLoading(true);
  await yieldToUI();

  try {
    const result = await runCompareBatched(mapA, mapB, headersA, headersB, keys, diffCols, ignoreCase, numericTol, modes);
    state.lastResult = { result, headersA, headersB, keys, diffCols };
    await writeReportSheet(result, headersA, headersB, keys, diffCols, modes);
    $("action-row").style.display = "flex";
    showMsg(
      "Report ready — Missing from B: " + result.missingFromB.length +
      "  ·  Missing from A: " + result.missingFromA.length +
      "  ·  Matching: " + result.matching.length +
      "  ·  Different: " + result.different.length,
      "success"
    );
  } catch (err) {
    showMsg("Error: " + friendlyError(err), "error");
  } finally {
    setBtnRunLoading(false);
  }
}

// ── Write Report Sheet ────────────────────────────────────────────────────────
const ACCENT = "#B0FF70";
const BLACK  = "#000000";
const DARK   = "#141414";

async function writeReportSheet(result, headersA, headersB, keys, diffCols, modes) {
  await Excel.run(async (ctx) => {
    const wks = ctx.workbook.worksheets;
    wks.load("items/name");
    await ctx.sync();
    const old = wks.items.find((s) => s.name === "Comparison Report");
    if (old) { old.delete(); await ctx.sync(); }

    const rpt = wks.add("Comparison Report");
    rpt.activate();
    let r = 0;

    // Title block
    const titleCell = rpt.getCell(r, 0);
    titleCell.values = [["Sheet Compare — Comparison Report"]];
    titleCell.format.font.bold = true; titleCell.format.font.size = 14; titleCell.format.font.color = ACCENT;
    r++;
    rpt.getCell(r, 0).values = [["Generated: " + new Date().toLocaleString()]];
    rpt.getCell(r, 0).format.font.color = "#888888";
    r += 2;

    // Summary row
    const summaryHeaders = ["Missing from B", "Missing from A", "Matching", "Different"];
    const summaryCounts  = [result.missingFromB.length, result.missingFromA.length, result.matching.length, result.different.length];
    const sumHdrRange = rpt.getRangeByIndexes(r, 0, 1, 4);
    sumHdrRange.values = [summaryHeaders];
    sumHdrRange.format.font.bold = true; sumHdrRange.format.font.color = "#888888"; sumHdrRange.format.font.size = 10;
    r++;
    const sumValRange = rpt.getRangeByIndexes(r, 0, 1, 4);
    sumValRange.values = [summaryCounts];
    sumValRange.format.font.bold = true; sumValRange.format.font.size = 16; sumValRange.format.font.color = ACCENT;
    r += 2;

    async function writeSection(title, cols, dataRows, fillColor) {
      if (dataRows.length === 0) return;
      // Section title
      const tc = rpt.getCell(r, 0);
      tc.values = [[title]]; tc.format.font.bold = true; tc.format.font.color = ACCENT;
      r++;
      // Header row
      const hr = rpt.getRangeByIndexes(r, 0, 1, cols.length);
      hr.values = [cols]; hr.format.font.bold = true; hr.format.font.color = BLACK; hr.format.fill.color = ACCENT;
      r++;
      // Data
      const BATCH = 250;
      for (let i = 0; i < dataRows.length; i += BATCH) {
        const chunk = dataRows.slice(i, i + BATCH);
        const dr = rpt.getRangeByIndexes(r, 0, chunk.length, cols.length);
        dr.values = chunk;
        if (fillColor) dr.format.fill.color = fillColor;
        r += chunk.length;
        await ctx.sync();
      }
      r++;
    }

    if (modes.missing && result.missingFromB.length > 0)
      await writeSection("Missing from B (" + result.missingFromB.length + ")", headersA,
        result.missingFromB.map(({ row }) => headersA.map((_, i) => row[i] ?? "")), "#1a0a0a");

    if (modes.missing && result.missingFromA.length > 0)
      await writeSection("Missing from A (" + result.missingFromA.length + ")", headersB,
        result.missingFromA.map(({ row }) => headersB.map((_, i) => row[i] ?? "")), "#1a0a0a");

    if (modes.matching && result.matching.length > 0)
      await writeSection("Matching (" + result.matching.length + ")", headersA,
        result.matching.map(({ rowA }) => headersA.map((_, i) => rowA[i] ?? "")), "#0a1a0a");

    if (modes.different && result.different.length > 0) {
      const difCols = [...keys, ...diffCols.flatMap((c) => [c + " (A)", c + " (B)"])];
      const difRows = result.different.map(({ rowA, rowB }) => [
        ...keys.map((k) => { const i = headersA.indexOf(k); return i >= 0 ? rowA[i] ?? "" : ""; }),
        ...diffCols.flatMap((col) => {
          const iA = headersA.indexOf(col); const iB = headersB.indexOf(col);
          return [iA >= 0 ? rowA[iA] ?? "" : "", iB >= 0 ? rowB[iB] ?? "" : ""];
        }),
      ]);
      await writeSection("Different (" + result.different.length + ")", difCols, difRows, "#1a1a00");
    }

    rpt.getRangeByIndexes(0, 0, r + 1, Math.max(headersA.length, headersB.length, 10))
       .format.autofitColumns();
    await ctx.sync();
  });
}

// ── Highlight sources ─────────────────────────────────────────────────────────
async function highlightSources() {
  if (!state.lastResult) { showMsg("Run a comparison first.", "error"); return; }
  const { result, headersA, headersB, keys } = state.lastResult;
  const ignoreCase = $("optIgnoreCase").checked;

  try {
    await Excel.run(async (ctx) => {
      const wks   = ctx.workbook.worksheets;
      const sheetA = wks.getItem(state.rangeA.sheetName);
      const sheetB = wks.getItem(state.rangeB.sheetName);
      await ctx.sync();

      const baseA   = parseRangeTopLeft(state.rangeA.address);
      const baseB   = parseRangeTopLeft(state.rangeB.address);
      const nColsA  = state.rangeA.headers.length;
      const nColsB  = state.rangeB.headers.length;

      function colorRow(sheet, base, dataIdx, numCols, color) {
        sheet.getRangeByIndexes(base.row + 1 + dataIdx, base.col, 1, numCols).format.fill.color = color;
      }

      const keyIdxA = keys.map((k) => headersA.indexOf(k));
      const keyIdxB = keys.map((k) => headersB.indexOf(k));
      const rowMapA = new Map(); state.rangeA.data.forEach((row, i) => { const k = makeKey(row, keyIdxA, ignoreCase); if (!rowMapA.has(k)) rowMapA.set(k, i); });
      const rowMapB = new Map(); state.rangeB.data.forEach((row, i) => { const k = makeKey(row, keyIdxB, ignoreCase); if (!rowMapB.has(k)) rowMapB.set(k, i); });

      let ops = 0;
      const SYNC = 300;
      for (const { keyStr } of result.missingFromB) { const i = rowMapA.get(keyStr); if (i != null) colorRow(sheetA, baseA, i, nColsA, "#3d0000"); if (++ops % SYNC === 0) await ctx.sync(); }
      for (const { keyStr } of result.missingFromA) { const i = rowMapB.get(keyStr); if (i != null) colorRow(sheetB, baseB, i, nColsB, "#3d0000"); if (++ops % SYNC === 0) await ctx.sync(); }
      for (const { keyStr } of result.different)    { const iA = rowMapA.get(keyStr); const iB = rowMapB.get(keyStr); if (iA != null) colorRow(sheetA, baseA, iA, nColsA, "#2e2800"); if (iB != null) colorRow(sheetB, baseB, iB, nColsB, "#2e2800"); if (++ops % SYNC === 0) await ctx.sync(); }
      await ctx.sync();
    });
    showMsg("Highlighted — dark red = missing, dark yellow = different.", "success");
  } catch (err) {
    showMsg("Highlight error: " + friendlyError(err), "error");
  }
}

// ── Export CSV ────────────────────────────────────────────────────────────────
function exportCSV() {
  if (!state.lastResult) { showMsg("Run a comparison first.", "error"); return; }
  const { result, headersA, headersB, keys, diffCols } = state.lastResult;

  function esc(v) { const s = String(v == null ? "" : v); return (s.includes(",") || s.includes('"') || s.includes("\n")) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function row2csv(arr) { return arr.map(esc).join(","); }

  const rows = [];
  if (result.missingFromB.length > 0) {
    rows.push(row2csv(["SECTION", ...headersA]));
    result.missingFromB.forEach(({ row }) => rows.push(row2csv(["Missing from B", ...headersA.map((_, i) => row[i] ?? "")])));
    rows.push("");
  }
  if (result.missingFromA.length > 0) {
    rows.push(row2csv(["SECTION", ...headersB]));
    result.missingFromA.forEach(({ row }) => rows.push(row2csv(["Missing from A", ...headersB.map((_, i) => row[i] ?? "")])));
    rows.push("");
  }
  if (result.matching.length > 0) {
    rows.push(row2csv(["SECTION", ...headersA]));
    result.matching.forEach(({ rowA }) => rows.push(row2csv(["Matching", ...headersA.map((_, i) => rowA[i] ?? "")])));
    rows.push("");
  }
  if (result.different.length > 0) {
    rows.push(row2csv(["SECTION", ...keys, ...diffCols.flatMap((c) => [c + " (A)", c + " (B)"])]));
    result.different.forEach(({ rowA, rowB }) => {
      rows.push(row2csv([
        "Different",
        ...keys.map((k) => { const i = headersA.indexOf(k); return i >= 0 ? rowA[i] ?? "" : ""; }),
        ...diffCols.flatMap((col) => { const iA = headersA.indexOf(col); const iB = headersB.indexOf(col); return [iA >= 0 ? rowA[iA] ?? "" : "", iB >= 0 ? rowB[iB] ?? "" : ""]; }),
      ]));
    });
  }

  const blob = new Blob([rows.join("\r\n")], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement("a"), { href: url, download: "comparison_report.csv" });
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showMsg("CSV downloaded.", "success");
}
