const $ = (id) => document.getElementById(id);
const val = (name) => document.querySelector(`input[name="${name}"]:checked`).value;
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

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

function ready(init) {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) return setStatus("Open this add-in in Excel.", true);
    init();
  });
}

/* selected cells, or the used range of the active sheet */
function getRange(ctx, scope) {
  return scope === "sheet"
    ? ctx.workbook.worksheets.getActiveWorksheet().getUsedRange(true)
    : ctx.workbook.getSelectedRange();
}

const SCOPE_HTML = `
  <section>
    <h2>Apply to</h2>
    <label class="opt"><input type="radio" name="scope" value="selection" checked> Selected cells</label>
    <label class="opt"><input type="radio" name="scope" value="sheet"> Whole active sheet</label>
  </section>`;
