/* Filter and Export — Excel task pane add-in */
const $ = (id) => document.getElementById(id);
const radio = (name) => document.querySelector(`input[name="${name}"]:checked`).value;
const checkedSheets = () => [...document.querySelectorAll("#sheetList input:checked")].map((i) => i.value);

let sheetNames = [];
const selected = new Set();

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return setStatus("Open this add-in in Excel.", true);
  $("selectAll").onclick = () => { sheetNames.forEach((n) => selected.add(n)); renderSheets(); };
  $("refresh").onclick = () => run(refreshSheets);
  $("fSheet").onchange = () => run(refreshColumns);
  $("fOp").onchange = () => $("val2Wrap").classList.toggle("hidden", $("fOp").value !== "between");
  $("applyFilter").onclick = () => run(applyFilter);
  $("clearFilter").onclick = () => run(clearFilters);
  $("exportBtn").onclick = () => run(exportFile);
  $("copyBtn").onclick = () => run(copyToSheets);
  await run(refreshSheets);
});

/* ---------- helpers ---------- */
function setStatus(msg, isError = false) {
  $("status").textContent = msg;
  $("status").className = "status" + (isError ? " error" : "");
}

async function run(fn) {
  const buttons = document.querySelectorAll("button");
  buttons.forEach((b) => (b.disabled = true));
  try { await fn(); }
  catch (e) { console.error(e); setStatus(e.message || String(e), true); }
  finally { buttons.forEach((b) => (b.disabled = false)); }
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: filename });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
const cleanName = (n) => n.replace(/[\[\]:*?/\\]/g, " ").trim().slice(0, 31) || "Sheet";

function uniqueName(base, taken) {
  let name = cleanName(base), i = 2;
  while (taken.has(name.toLowerCase())) name = cleanName(base).slice(0, 28) + " " + i++;
  taken.add(name.toLowerCase());
  return name;
}

function pad(rows) {
  const w = Math.max(1, ...rows.map((r) => r.length));
  return rows.map((r) => Array.from({ length: w }, (_, i) => (r[i] === undefined ? "" : r[i])));
}

/* ---------- sheet list + columns ---------- */
async function refreshSheets() {
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    sheetNames = sheets.items.map((s) => s.name);
  });
  [...selected].forEach((n) => !sheetNames.includes(n) && selected.delete(n));
  if (!selected.size && sheetNames.length) selected.add(sheetNames[0]);
  renderSheets();
  const current = $("fSheet").value;
  $("fSheet").innerHTML = sheetNames.map((n) => `<option>${esc(n)}</option>`).join("");
  if (sheetNames.includes(current)) $("fSheet").value = current;
  await refreshColumns();
  setStatus("");
}

function renderSheets() {
  $("sheetList").innerHTML = sheetNames
    .map((n) => `<label><input type="checkbox" value="${esc(n)}" ${selected.has(n) ? "checked" : ""}> ${esc(n)}</label>`)
    .join("");
  $("sheetList").querySelectorAll("input").forEach((i) =>
    (i.onchange = () => (i.checked ? selected.add(i.value) : selected.delete(i.value)))
  );
}

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

async function refreshColumns() {
  const name = $("fSheet").value;
  if (!name) { $("fCol").innerHTML = ""; return; }
  let headers = [];
  await Excel.run(async (ctx) => {
    const used = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await ctx.sync();
    if (used.isNullObject) return;
    const hdr = used.getRow(0);
    hdr.load("values");
    await ctx.sync();
    headers = hdr.values[0];
  });
  $("fCol").innerHTML = headers
    .map((h, i) => `<option value="${i}">${esc(h === "" ? `Column ${i + 1}` : h)}</option>`)
    .join("");
}

/* ---------- filtering (uses Excel's own AutoFilter) ---------- */
function criteria(op, v1, v2) {
  switch (op) {
    case "contains": return { criterion1: `=*${v1}*` };
    case "starts":   return { criterion1: `=${v1}*` };
    case "equals":   return { criterion1: `=${v1}` };
    case "gt":       return { criterion1: `>${v1}` };
    case "lt":       return { criterion1: `<${v1}` };
    case "between":  return { criterion1: `>=${v1}`, criterion2: `<=${v2}`, operator: Excel.FilterOperator.and };
  }
}

async function applyFilter() {
  const name = $("fSheet").value, col = Number($("fCol").value), op = $("fOp").value;
  const v1 = $("fVal").value.trim(), v2 = $("fVal2").value.trim();
  if (!name) throw new Error("Pick a sheet to filter.");
  if (!v1 || (op === "between" && !v2)) throw new Error("Enter a value to filter by.");

  let shown = 0, total = 0;
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(name);
    const used = ws.getUsedRange(true);
    ws.autoFilter.apply(used, col, { filterOn: Excel.FilterOn.custom, ...criteria(op, v1, v2) });
    await ctx.sync();
    used.load("rowCount");
    const vis = used.getSpecialCellsOrNullObject(Excel.SpecialCellType.visible);
    vis.load("isNullObject");
    await ctx.sync();
    if (!vis.isNullObject) {
      vis.areas.load("items/rowCount");
      await ctx.sync();
      shown = vis.areas.items.reduce((n, a) => n + a.rowCount, 0) - 1;
    }
    total = used.rowCount - 1;
  });
  setStatus(`Filter applied. Showing ${shown} of ${total} rows on "${name}".`);
}

async function clearFilters() {
  const name = $("fSheet").value;
  if (!name) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(name).autoFilter.clearCriteria();
    await ctx.sync();
  });
  setStatus(`Filters cleared on "${name}".`);
}

/* ---------- reading data ---------- */
async function readSheet(name, visibleOnly) {
  let out = { name, values: [], text: [] };
  await Excel.run(async (ctx) => {
    const used = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await ctx.sync();
    if (used.isNullObject) return;
    if (!visibleOnly) {
      used.load(["values", "text"]);
      await ctx.sync();
      out.values = used.values; out.text = used.text;
      return;
    }
    const vis = used.getSpecialCellsOrNullObject(Excel.SpecialCellType.visible);
    vis.load("isNullObject");
    await ctx.sync();
    if (vis.isNullObject) return;
    vis.areas.load("items/values,items/text");
    await ctx.sync();
    vis.areas.items.forEach((a) => { out.values.push(...a.values); out.text.push(...a.text); });
  });
  return out;
}

/* Returns [{ name, values, text }]; merges into one table when layout = merged */
async function collect() {
  const names = checkedSheets();
  if (!names.length) throw new Error("Select at least one sheet.");
  const visibleOnly = radio("scope") === "visible";
  const sets = [];
  for (const n of names) {
    const d = await readSheet(n, visibleOnly);
    if (d.values.length) sets.push(d);
  }
  if (!sets.length) throw new Error("The selected sheets have no data.");

  if (radio("layout") === "merged" && sets.length > 1) {
    const merge = (key) => {
      const rows = [["Source sheet", ...sets[0][key][0]]];
      sets.forEach((s) => s[key].slice(1).forEach((r) => rows.push([s.name, ...r])));
      return pad(rows);
    };
    return [{ name: "Combined", values: merge("values"), text: merge("text") }];
  }
  return sets;
}

/* ---------- export ---------- */
async function exportFile() {
  const sets = await collect();
  const fmt = radio("fmt");

  if (fmt === "xlsx") {
    const wb = XLSX.utils.book_new(), taken = new Set();
    sets.forEach((s) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.values), uniqueName(s.name, taken)));
    const bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    download(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `export-${stamp()}.xlsx`);
  } else {
    const doc = new window.jspdf.jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    sets.forEach((s, i) => {
      if (i > 0) doc.addPage();
      const [head, ...body] = s.text;
      doc.setFontSize(12);
      doc.text(s.name, 40, 36);
      doc.autoTable({
        head: [head], body, startY: 48, margin: { left: 40, right: 40 },
        styles: { fontSize: head.length > 12 ? 6 : head.length > 8 ? 7 : 8.5, cellPadding: 3 },
        headStyles: { fillColor: [15, 110, 90] },
      });
    });
    doc.save(`export-${stamp()}.pdf`);
  }
  setStatus(`Exported ${sets.length} ${sets.length === 1 ? "sheet" : "sheets"} as ${fmt === "pdf" ? "PDF" : "Excel"}.`);
}

/* ---------- copy results into this workbook ---------- */
async function copyToSheets() {
  const sets = await collect();
  await Excel.run(async (ctx) => {
    const existing = ctx.workbook.worksheets;
    existing.load("items/name");
    await ctx.sync();
    const taken = new Set(existing.items.map((s) => s.name.toLowerCase()));
    let last;
    for (const s of sets) {
      const ws = ctx.workbook.worksheets.add(uniqueName(`Export ${s.name}`, taken));
      const data = pad(s.values);
      const range = ws.getRangeByIndexes(0, 0, data.length, data[0].length);
      range.values = data;
      ws.getRangeByIndexes(0, 0, 1, data[0].length).format.font.bold = true;
      range.format.autofitColumns();
      last = ws;
    }
    last.activate();
    await ctx.sync();
  });
  await refreshSheets();
  setStatus(`Added ${sets.length} new ${sets.length === 1 ? "sheet" : "sheets"} to this workbook.`);
}
