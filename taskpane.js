/* ============================================================
   Sheet Compare — Task Pane Logic
   All comparison runs entirely client-side via Office.js.
   ============================================================ */

"use strict";

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  rangeA: null,   // { address, sheetName, headers, data }
  rangeB: null,
  lastResult: null,
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function $(elId) { return document.getElementById(elId); }

function yieldToUI() { return new Promise((r) => setTimeout(r, 0)); }

function showMsg(text, type) {
  const el = document.createElement("div");
  el.className = "msg msg-" + (type || "info");
  el.textContent = text;
  $("messages").appendChild(el);
  el.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function clearMessages() { $("messages").innerHTML = ""; }

function friendlyError(err) {
  if (typeof OfficeExtension !== "undefined" && err instanceof OfficeExtension.Error) return err.message;
  if (err instanceof Error) return err.message;
  return String(err);
}

function setBtnLoading(loading) {
  const btn = $("btnRun");
  if (loading) {
    btn.innerHTML = '<span class="spinner"></span>Running…';
    btn.disabled = true;
  } else {
    btn.innerHTML = "▶ Run Comparison";
    btn.disabled = false;
  }
}

// ── Modal ─────────────────────────────────────────────────────────────────────
let _modalResolve = null;

function showModal(title, body) {
  $("modal-title").textContent = title;
  $("modal-body").textContent  = body;
  $("modal-overlay").style.display = "flex";
  return new Promise((resolve) => { _modalResolve = resolve; });
}

function resolveModal(value) {
  $("modal-overlay").style.display = "none";
  if (_modalResolve) { _modalResolve(value); _modalResolve = null; }
}

// ── Address parsing ───────────────────────────────────────────────────────────
/**
 * Extract 0-based {row, col} for the top-left cell of an Excel range address.
 * Handles formats like "Sheet1!$A$2:$D$100", "A2:D100", "$B$3".
 */
function parseRangeTopLeft(address) {
  const clean = address.replace(/\$/g, "").split("!").pop().split(":")[0];
  const letters = clean.replace(/[^A-Za-z]/g, "").toUpperCase();
  const digits  = clean.replace(/[^0-9]/g, "");
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return { col: col - 1, row: parseInt(digits, 10) - 1 };
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
  data.forEach((row, rowIdx) => {
    const k = makeKey(row, indices, ignoreCase);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push({ row, rowIdx });
  });
  return map;
}

function findDuplicateKeys(data, indices, ignoreCase) {
  const counts = new Map();
  data.forEach((row) => {
    const k = makeKey(row, indices, ignoreCase);
    counts.set(k, (counts.get(k) || 0) + 1);
  });
  return [...counts.entries()]
    .filter(([, c]) => c > 1)
    .map(([k]) => k.replace(/\u0000/g, " | "));
}

function computeDiffs(rowA, rowB, headersA, headersB, diffCols, ignoreCase, numericTol) {
  const diffs = [];
  diffCols.forEach((col) => {
    const iA = headersA.indexOf(col);
    const iB = headersB.indexOf(col);
    if (iA < 0 || iB < 0) return;
    const vA = rowA[iA];
    const vB = rowB[iB];
    let equal = normalizeVal(vA, ignoreCase) === normalizeVal(vB, ignoreCase);
    if (!equal && numericTol != null) equal = numericClose(vA, vB, numericTol);
    if (!equal) diffs.push({ col, vA, vB });
  });
  return diffs;
}

async function runCompareBatched(mapA, mapB, headersA, headersB, keys, diffCols, ignoreCase, numericTol, modes) {
  const BATCH = 500;
  const missingFromB = [], missingFromA = [], matching = [], different = [];
  const allKeysA = [...mapA.keys()];
  const allKeysB = [...mapB.keys()];

  for (let i = 0; i < allKeysA.length; i++) {
    if (i > 0 && i % BATCH === 0) await yieldToUI();
    const k = allKeysA[i];
    const entriesA = mapA.get(k);
    if (mapB.has(k)) {
      const rowA = entriesA[0].row;
      const rowB = mapB.get(k)[0].row;
      if (modes.matching) matching.push({ keyStr: k, rowA, rowB });
      if (modes.different) {
        const diffs = computeDiffs(rowA, rowB, headersA, headersB, diffCols, ignoreCase, numericTol);
        if (diffs.length > 0) different.push({ keyStr: k, rowA, rowB, diffs });
      }
    } else {
      if (modes.missing) entriesA.forEach(({ row }) => missingFromB.push({ keyStr: k, row }));
    }
  }

  if (modes.missing) {
    for (let i = 0; i < allKeysB.length; i++) {
      if (i > 0 && i % BATCH === 0) await yieldToUI();
      const k = allKeysB[i];
      if (!mapA.has(k)) {
        mapB.get(k).forEach(({ row }) => missingFromA.push({ keyStr: k, row }));
      }
    }
  }

  return { missingFromB, missingFromA, matching, different };
}

// ── Report sheet ──────────────────────────────────────────────────────────────
const GREEN  = "#217346";
const WHITE  = "#FFFFFF";
const SECT_BG = "#E9F5EE";
const SECT_FG = "#1A3D2B";

async function writeReportSheet(result, headersA, headersB, keys, diffCols, modes) {
  await Excel.run(async (ctx) => {
    // Remove old report sheet if present
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const old = sheets.items.find((s) => s.name === "Comparison Report");
    if (old) { old.delete(); await ctx.sync(); }

    const rpt = sheets.add("Comparison Report");
    rpt.activate();
    let r = 0; // current 0-based row

    // Summary
    const summaryData = [
      ["Sheet Compare — Comparison Report"],
      ["Generated:", new Date().toLocaleString()],
      [],
      ["Summary"],
      ["Missing from B (in A, not B):", result.missingFromB.length],
      ["Missing from A (in B, not A):", result.missingFromA.length],
      ["Matching:",                      result.matching.length],
      ["Different:",                     result.different.length],
      [],
    ];

    summaryData.forEach((line, i) => {
      if (!line || line.length === 0) { r++; return; }
      rpt.getCell(r, 0).values = [[line[0]]];
      if (line[1] != null) rpt.getCell(r, 1).values = [[line[1]]];
      if (i === 0) {
        const c = rpt.getCell(r, 0);
        c.format.font.bold = true; c.format.font.size = 14; c.format.font.color = GREEN;
      } else if (i === 3) {
        rpt.getCell(r, 0).format.font.bold = true;
      }
      r++;
    });

    // Helper: write a colored section
    async function writeSection(title, cols, dataRows, fillColor) {
      if (dataRows.length === 0) return;

      // Section title row
      const titleCell = rpt.getCell(r, 0);
      titleCell.values = [[title]];
      titleCell.format.font.bold = true;
      titleCell.format.font.color = SECT_FG;
      titleCell.format.fill.color = SECT_BG;
      r++;

      // Header row
      const hdr = rpt.getRangeByIndexes(r, 0, 1, cols.length);
      hdr.values = [cols];
      hdr.format.font.bold = true; hdr.format.font.color = WHITE; hdr.format.fill.color = GREEN;
      r++;

      // Data in batches of 250
      const BATCH = 250;
      for (let i = 0; i < dataRows.length; i += BATCH) {
        const chunk = dataRows.slice(i, i + BATCH);
        const rng = rpt.getRangeByIndexes(r, 0, chunk.length, cols.length);
        rng.values = chunk;
        if (fillColor) rng.format.fill.color = fillColor;
        r += chunk.length;
        await ctx.sync();
      }
      r++; // blank line separator
    }

    if (modes.missing && result.missingFromB.length > 0) {
      await writeSection(
        `Missing from B — ${result.missingFromB.length} row(s) in A but not in B`,
        headersA,
        result.missingFromB.map(({ row }) => headersA.map((_, i) => row[i] ?? "")),
        "#FDE8E8"
      );
    }

    if (modes.missing && result.missingFromA.length > 0) {
      await writeSection(
        `Missing from A — ${result.missingFromA.length} row(s) in B but not in A`,
        headersB,
        result.missingFromA.map(({ row }) => headersB.map((_, i) => row[i] ?? "")),
        "#FDE8E8"
      );
    }

    if (modes.matching && result.matching.length > 0) {
      await writeSection(
        `Matching — ${result.matching.length} row(s) present in both tables`,
        headersA,
        result.matching.map(({ rowA }) => headersA.map((_, i) => rowA[i] ?? "")),
        "#D1FAE5"
      );
    }

    if (modes.different && result.different.length > 0) {
      const difCols = [
        ...keys,
        ...diffCols.flatMap((c) => [`${c} (A)`, `${c} (B)`]),
      ];
      const difRows = result.different.map(({ rowA, rowB }) => [
        ...keys.map((k) => { const i = headersA.indexOf(k); return i >= 0 ? rowA[i] ?? "" : ""; }),
        ...diffCols.flatMap((col) => {
          const iA = headersA.indexOf(col); const iB = headersB.indexOf(col);
          return [iA >= 0 ? rowA[iA] ?? "" : "", iB >= 0 ? rowB[iB] ?? "" : ""];
        }),
      ]);
      await writeSection(
        `Different — ${result.different.length} row(s) with value differences`,
        difCols, difRows, "#FFF9DB"
      );
    }

    // Footer note
    rpt.getCell(r, 0).values = [["Use the 'Export as CSV' button in the task pane to export this report."]];
    rpt.getCell(r, 0).format.font.italic = true;
    rpt.getCell(r, 0).format.font.color  = "#6b7280";

    // Auto-fit columns
    rpt.getRangeByIndexes(0, 0, r + 1, Math.max(headersA.length, headersB.length, 20))
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
      const wks = ctx.workbook.worksheets;

      function getSheet(name) {
        const s = wks.getItem(name);
        s.load("name");
        return s;
      }

      const sheetA = getSheet(state.rangeA.sheetName);
      const sheetB = getSheet(state.rangeB.sheetName);
      await ctx.sync();

      const baseA = parseRangeTopLeft(state.rangeA.address);
      const baseB = parseRangeTopLeft(state.rangeB.address);
      const numColsA = state.rangeA.headers.length;
      const numColsB = state.rangeB.headers.length;

      function colorRow(sheet, base, dataRowIdx, numCols, color) {
        // dataRowIdx is 0-based among data rows (header at base.row, data starts at base.row+1)
        const excelRow = base.row + 1 + dataRowIdx;
        sheet.getRangeByIndexes(excelRow, base.col, 1, numCols).format.fill.color = color;
      }

      // Build key → first data-row index maps
      const keyIdxA = keys.map((k) => headersA.indexOf(k));
      const keyIdxB = keys.map((k) => headersB.indexOf(k));
      const rowIdxA = new Map();
      const rowIdxB = new Map();
      state.rangeA.data.forEach((row, i) => {
        const k = makeKey(row, keyIdxA, ignoreCase);
        if (!rowIdxA.has(k)) rowIdxA.set(k, i);
      });
      state.rangeB.data.forEach((row, i) => {
        const k = makeKey(row, keyIdxB, ignoreCase);
        if (!rowIdxB.has(k)) rowIdxB.set(k, i);
      });

      let ops = 0;
      const SYNC_EVERY = 300;

      for (const { keyStr } of result.missingFromB) {
        const i = rowIdxA.get(keyStr);
        if (i != null) colorRow(sheetA, baseA, i, numColsA, "#FFCCCC");
        if (++ops % SYNC_EVERY === 0) await ctx.sync();
      }
      for (const { keyStr } of result.missingFromA) {
        const i = rowIdxB.get(keyStr);
        if (i != null) colorRow(sheetB, baseB, i, numColsB, "#FFCCCC");
        if (++ops % SYNC_EVERY === 0) await ctx.sync();
      }
      for (const { keyStr } of result.different) {
        const iA = rowIdxA.get(keyStr);
        const iB = rowIdxB.get(keyStr);
        if (iA != null) colorRow(sheetA, baseA, iA, numColsA, "#FFF3B0");
        if (iB != null) colorRow(sheetB, baseB, iB, numColsB, "#FFF3B0");
        if (++ops % SYNC_EVERY === 0) await ctx.sync();
      }

      await ctx.sync();
    });

    showMsg("Source sheets highlighted: red = missing, yellow = different.", "success");
  } catch (err) {
    showMsg("Error highlighting: " + friendlyError(err), "error");
  }
}

// ── Export CSV ────────────────────────────────────────────────────────────────
function exportCSV() {
  if (!state.lastResult) { showMsg("Run a comparison first.", "error"); return; }
  try {
    const { result, headersA, headersB, keys, diffCols } = state.lastResult;
    const rows = [];

    function escapeCsv(v) {
      const s = String(v == null ? "" : v);
      return (s.includes(",") || s.includes('"') || s.includes("\n"))
        ? '"' + s.replace(/"/g, '""') + '"' : s;
    }
    function rowToCsv(arr) { return arr.map(escapeCsv).join(","); }

    if (result.missingFromB.length > 0) {
      rows.push(rowToCsv(["SECTION", ...headersA]));
      result.missingFromB.forEach(({ row }) =>
        rows.push(rowToCsv(["Missing from B", ...headersA.map((_, i) => row[i] ?? "")])));
      rows.push("");
    }
    if (result.missingFromA.length > 0) {
      rows.push(rowToCsv(["SECTION", ...headersB]));
      result.missingFromA.forEach(({ row }) =>
        rows.push(rowToCsv(["Missing from A", ...headersB.map((_, i) => row[i] ?? "")])));
      rows.push("");
    }
    if (result.matching.length > 0) {
      rows.push(rowToCsv(["SECTION", ...headersA]));
      result.matching.forEach(({ rowA }) =>
        rows.push(rowToCsv(["Matching", ...headersA.map((_, i) => rowA[i] ?? "")])));
      rows.push("");
    }
    if (result.different.length > 0) {
      const difHdrs = ["SECTION", ...keys, ...diffCols.flatMap((c) => [`${c} (A)`, `${c} (B)`])];
      rows.push(rowToCsv(difHdrs));
      result.different.forEach(({ rowA, rowB }) => {
        rows.push(rowToCsv([
          "Different",
          ...keys.map((k) => { const i = headersA.indexOf(k); return i >= 0 ? rowA[i] ?? "" : ""; }),
          ...diffCols.flatMap((col) => {
            const iA = headersA.indexOf(col); const iB = headersB.indexOf(col);
            return [iA >= 0 ? rowA[iA] ?? "" : "", iB >= 0 ? rowB[iB] ?? "" : ""];
          }),
        ]));
      });
    }

    const blob = new Blob([rows.join("\r\n")], { type: "text/csv" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = "comparison_report.csv";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showMsg("CSV exported.", "success");
  } catch (err) {
    showMsg("Export error: " + friendlyError(err), "error");
  }
}

// ── Select range ──────────────────────────────────────────────────────────────
async function selectRange(which) {
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
      if (!values || values.length === 0) throw new Error("Selected range is empty.");
      if (values.length < 2) throw new Error("Table must have at least one header row and one data row.");

      const headers = values[0].map((h) => String(h ?? "").trim());
      const blankHeaders = headers.filter((h) => h === "");
      if (blankHeaders.length > 0)
        throw new Error("Header row contains blank cells. Ensure all column headers are filled in.");

      return { address: range.address, sheetName: ws.name, headers, data: values.slice(1) };
    });

    state["range" + which] = result;
    $("range" + which + "-addr").textContent =
      result.sheetName + "!" + result.address + " (" + result.data.length + " rows)";

    rebuildKeyColumns();
    rebuildDiffColumns();
    showMsg(
      "Table " + which + " selected: " + result.headers.length + " columns, " + result.data.length + " rows.",
      "success"
    );
  } catch (err) {
    showMsg("Error selecting Table " + which + ": " + friendlyError(err), "error");
  } finally {
    btn.disabled = false;
  }
}

// ── Column selects ────────────────────────────────────────────────────────────
function rebuildKeyColumns() {
  const sel = $("keyColumns");
  sel.innerHTML = "";

  if (!state.rangeA || !state.rangeB) {
    sel.innerHTML = '<option disabled>— select both tables first —</option>';
    return;
  }

  const common = state.rangeA.headers.filter((h) => state.rangeB.headers.includes(h));
  if (common.length === 0) {
    sel.innerHTML = '<option disabled>— no common column names found —</option>';
    showMsg("Warning: Tables share no common column headers. Key selection is impossible.", "warning");
    return;
  }

  common.forEach((h) => {
    const opt = document.createElement("option");
    opt.value = h; opt.textContent = h;
    sel.appendChild(opt);
  });

  rebuildDiffColumns();
}

function rebuildDiffColumns() {
  const sel = $("diffColumns");
  sel.innerHTML = "";

  if (!state.rangeA || !state.rangeB) {
    sel.innerHTML = '<option disabled>— select tables &amp; keys first —</option>';
    return;
  }

  const keys = selectedKeys();
  const candidates = [
    ...new Set([
      ...state.rangeA.headers.filter((h) => !keys.includes(h)),
      ...state.rangeB.headers.filter((h) => !keys.includes(h)),
    ]),
  ];

  if (candidates.length === 0) {
    sel.innerHTML = '<option disabled>— all columns are keys —</option>';
    return;
  }

  candidates.forEach((h) => {
    const opt = document.createElement("option");
    opt.value = h; opt.textContent = h;
    sel.appendChild(opt);
  });
}

function refreshDiffColumnsVis() {
  $("diffColumnsWrap").style.display = $("modeDifferent").checked ? "" : "none";
}

function selectedKeys() {
  return Array.from($("keyColumns").selectedOptions).map((o) => o.value);
}

function selectedDiffCols() {
  return Array.from($("diffColumns").selectedOptions).map((o) => o.value);
}

// ── Main run ──────────────────────────────────────────────────────────────────
async function runComparison() {
  clearMessages();

  if (!state.rangeA || !state.rangeB) {
    showMsg("Please select both Table A and Table B first.", "error"); return;
  }

  const keys = selectedKeys();
  if (keys.length === 0) {
    showMsg("Please select at least one key column.", "error"); return;
  }

  const modes = {
    missing:   $("modeMissing").checked,
    matching:  $("modeMatching").checked,
    different: $("modeDifferent").checked,
  };
  if (!modes.missing && !modes.matching && !modes.different) {
    showMsg("Please check at least one comparison mode.", "error"); return;
  }

  // Same range overlap guard
  if (
    state.rangeA.sheetName === state.rangeB.sheetName &&
    state.rangeA.address === state.rangeB.address
  ) {
    showMsg("Table A and Table B are the same range. Select different ranges.", "error"); return;
  }

  const ignoreCase = $("optIgnoreCase").checked;
  const numericTol = $("optNumericTol").checked ? (parseFloat($("tolValue").value) || 0) : null;

  const headersA = state.rangeA.headers;
  const headersB = state.rangeB.headers;
  const keyIdxA  = keys.map((k) => headersA.indexOf(k));
  const keyIdxB  = keys.map((k) => headersB.indexOf(k));

  if (keyIdxA.some((i) => i < 0)) {
    showMsg("One or more key columns not found in Table A.", "error"); return;
  }
  if (keyIdxB.some((i) => i < 0)) {
    showMsg("One or more key columns not found in Table B.", "error"); return;
  }

  // Blank key check
  for (const row of state.rangeA.data) {
    if (keyIdxA.some((i) => row[i] == null || String(row[i]).trim() === "")) {
      showMsg("Table A has rows with blank key values. Please fix and re-select.", "error"); return;
    }
  }
  for (const row of state.rangeB.data) {
    if (keyIdxB.some((i) => row[i] == null || String(row[i]).trim() === "")) {
      showMsg("Table B has rows with blank key values. Please fix and re-select.", "error"); return;
    }
  }

  // Duplicate key detection
  const dupA = findDuplicateKeys(state.rangeA.data, keyIdxA, ignoreCase);
  const dupB = findDuplicateKeys(state.rangeB.data, keyIdxB, ignoreCase);

  if (dupA.length > 0 || dupB.length > 0) {
    let msg = "";
    if (dupA.length) {
      msg += "Table A duplicates:\n" + dupA.slice(0, 8).join("\n") +
        (dupA.length > 8 ? "\n…and " + (dupA.length - 8) + " more" : "") + "\n\n";
    }
    if (dupB.length) {
      msg += "Table B duplicates:\n" + dupB.slice(0, 8).join("\n") +
        (dupB.length > 8 ? "\n…and " + (dupB.length - 8) + " more" : "");
    }
    const ok = await showModal("⚠️ Duplicate Keys Detected",
      msg.trim() + "\n\nRow matching will use the first occurrence. Proceed?");
    if (!ok) return;
  }

  // Determine diff columns
  let diffCols = selectedDiffCols();
  if (diffCols.length === 0) {
    diffCols = [...new Set([
      ...headersA.filter((h) => !keys.includes(h)),
      ...headersB.filter((h) => !keys.includes(h)),
    ])];
  }

  // Build maps and run
  const mapA = buildMap(state.rangeA.data, keyIdxA, ignoreCase);
  const mapB = buildMap(state.rangeB.data, keyIdxB, ignoreCase);

  setBtnLoading(true);
  await yieldToUI();

  try {
    const result = await runCompareBatched(
      mapA, mapB, headersA, headersB, keys, diffCols, ignoreCase, numericTol, modes
    );
    state.lastResult = { result, headersA, headersB, keys, diffCols };

    await writeReportSheet(result, headersA, headersB, keys, diffCols, modes);
    $("btnHighlight").style.display = "";
    $("btnExportCsv").style.display = "";
    showMsg(
      "✅ Done! Report written to the 'Comparison Report' sheet. " +
      "Missing from B: " + result.missingFromB.length +
      " | Missing from A: " + result.missingFromA.length +
      " | Matching: " + result.matching.length +
      " | Different: " + result.different.length,
      "success"
    );
  } catch (err) {
    showMsg("Error: " + friendlyError(err), "error");
  } finally {
    setBtnLoading(false);
  }
}

// ── Office init ───────────────────────────────────────────────────────────────
Office.onReady(function () {
  $("btnSelectA").onclick = () => selectRange("A");
  $("btnSelectB").onclick = () => selectRange("B");

  $("modeDifferent").onchange = refreshDiffColumnsVis;
  $("optNumericTol").onchange = () => {
    $("tolWrap").style.display = $("optNumericTol").checked ? "" : "none";
  };

  $("btnRun").onclick       = runComparison;
  $("btnHighlight").onclick = highlightSources;
  $("btnExportCsv").onclick = exportCSV;

  $("btnModalCancel").onclick  = () => resolveModal(false);
  $("btnModalProceed").onclick = () => resolveModal(true);

  refreshDiffColumnsVis();
});
