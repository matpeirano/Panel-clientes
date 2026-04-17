
const state = {
  actualData: [],
  anteriorData: [],
  currentActualName: "",
  currentAnteriorName: ""
};

const REQUIRED_COLUMNS = [
  "comitente", "cuenta", "Es Juridica", "arancel",
  "AUM en Dolares", "cv7000", "$ Operables CI",
  "MEP Operables CI", "Comision 180", "Tipo Cbio MEP"
];

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("fileActual").addEventListener("change", (e) => handleFile(e, "actual"));
  document.getElementById("fileAnterior").addEventListener("change", (e) => handleFile(e, "anterior"));
  document.getElementById("btnExportPdf").addEventListener("click", () => window.print());

  ["filtroJuridica", "filtroArancel", "filtroCliente", "filtroTopN"].forEach(id => {
    document.getElementById(id).addEventListener("input", renderDashboard);
    document.getElementById(id).addEventListener("change", renderDashboard);
  });
});

function normalizeText(v) {
  return (v ?? "").toString().trim();
}

function normalizeNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  s = s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

async function handleFile(event, mode) {
  const file = event.target.files?.[0];
  if (!file) return;

  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheet];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  if (!json.length) {
    alert("El archivo no contiene datos.");
    return;
  }

  validateColumns(json[0]);

  const rows = json.map(row => normalizeRow(row));

  if (mode === "actual") {
    state.actualData = rows;
    state.currentActualName = file.name;
    document.getElementById("actualFileName").textContent = `Excel actual: ${file.name}`;
    populateArancelFilter(rows);
  } else {
    state.anteriorData = rows;
    state.currentAnteriorName = file.name;
    document.getElementById("anteriorFileName").textContent = `Excel anterior: ${file.name}`;
  }

  renderDashboard();
}

function validateColumns(exampleRow) {
  const cols = Object.keys(exampleRow);
  const missing = REQUIRED_COLUMNS.filter(c => !cols.includes(c));
  if (missing.length) {
    throw new Error("Faltan columnas obligatorias en el Excel: " + missing.join(", "));
  }
}

function normalizeRow(row) {
  return {
    comitente: normalizeText(row["comitente"]),
    cuenta: normalizeText(row["cuenta"]),
    esJuridica: normalizeText(row["Es Juridica"]),
    arancel: normalizeText(row["arancel"]),
    aum: normalizeNumber(row["AUM en Dolares"]),
    cv7000: normalizeNumber(row["cv7000"]),
    pesosOperables: normalizeNumber(row["$ Operables CI"]),
    mepOperables: normalizeNumber(row["MEP Operables CI"]),
    comision180: normalizeNumber(row["Comision 180"]),
    tipoCbioMep: normalizeNumber(row["Tipo Cbio MEP"])
  };
}

function populateArancelFilter(rows) {
  const select = document.getElementById("filtroArancel");
  const current = select.value;
  const aranceles = [...new Set(rows.map(r => r.arancel).filter(Boolean))].sort();

  select.innerHTML = `<option value="TODOS">Todos</option>` +
    aranceles.map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join("");

  if (aranceles.includes(current)) {
    select.value = current;
  } else {
    select.value = "TODOS";
  }
}

function getFilteredData() {
  const juridica = document.getElementById("filtroJuridica").value;
  const arancel = document.getElementById("filtroArancel").value;
  const cliente = document.getElementById("filtroCliente").value.trim().toUpperCase();

  return state.actualData.filter(r => {
    const okJ = juridica === "TODAS" || r.esJuridica === juridica;
    const okA = arancel === "TODOS" || r.arancel === arancel;
    const text = `${r.cuenta} ${r.comitente}`.toUpperCase();
    const okC = !cliente || text.includes(cliente);
    return okJ && okA && okC;
  });
}

function renderDashboard() {
  if (!state.actualData.length) return;

  const rows = getFilteredData();
  const topN = Number(document.getElementById("filtroTopN").value || 15);

  const enriched = rows.map(r => ({
    ...r,
    pesosOperablesUsd: r.tipoCbioMep > 0 ? (r.pesosOperables / r.tipoCbioMep) : 0,
    liquidezTotal: r.cv7000 + (r.tipoCbioMep > 0 ? (r.pesosOperables / r.tipoCbioMep) : 0) + r.mepOperables
  }));

  updateKpis(enriched);
  renderTopAumTable(enriched, topN);
  renderLiquidezTable(enriched, topN);
  renderInactivosTable(enriched);

  if (state.anteriorData.length) {
    renderComparacionAum(enriched, topN);
  } else {
    document.getElementById("comparacionSection").classList.add("hidden");
  }
}

function updateKpis(rows) {
  const totalClientes = rows.length;
  const juridicas = rows.filter(r => r.esJuridica === "1").length;
  const noJuridicas = rows.filter(r => r.esJuridica === "0").length;
  const aum = sum(rows, "aum");
  const liquidez = rows.reduce((acc, r) => acc + r.liquidezTotal, 0);
  const comision = sum(rows, "comision180");
  const inactivos = rows.filter(r => r.comision180 < 15).length;

  setText("kpiClientes", formatInt(totalClientes));
  setText("kpiJuridicas", formatInt(juridicas));
  setText("kpiNoJuridicas", formatInt(noJuridicas));
  setText("kpiAum", formatUsd(aum));
  setText("kpiLiquidez", formatUsd(liquidez));
  setText("kpiComision", formatUsd(comision));
  setText("kpiInactivos", formatInt(inactivos));
}

function renderLiquidezTable(rows, topN) {
  const top = [...rows]
    .sort((a,b) => b.liquidezTotal - a.liquidezTotal)
    .slice(0, topN);

  renderTable("tablaLiquidez", top, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatUsd(row.cv7000)}</td>
      <td>${formatUsd(row.pesosOperablesUsd)}</td>
      <td>${formatUsd(row.mepOperables)}</td>
      <td>${formatUsd(row.tipoCbioMep)}</td>
      <td>${formatUsd(row.liquidezTotal)}</td>
      <td>${formatUsd(row.comision180)}</td>
    </tr>
  `);
}

function renderTopAumTable(rows, topN) {
  const top = [...rows]
    .sort((a,b) => b.aum - a.aum)
    .slice(0, topN);

  renderTable("tablaTopAum", top, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${row.esJuridica === "1" ? "Sí" : "No"}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatUsd(row.aum)}</td>
    </tr>
  `);
}


function renderInactivosTable(rows) {
  const inactivos = [...rows]
    .filter(r => r.comision180 < 15)
    .sort((a,b) => a.comision180 - b.comision180);

  renderTable("tablaInactivos", inactivos, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${row.esJuridica === "1" ? "Sí" : "No"}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatUsd(row.aum)}</td>
      <td>${formatUsd(row.comision180)}</td>
      <td>${formatUsd(row.liquidezTotal)}</td>
    </tr>
  `);
}

function renderComparacionAum(actualRows, topN) {
  const anteriorMap = new Map();
  state.anteriorData.forEach(r => anteriorMap.set(r.comitente, r));

  const merged = actualRows.map(r => {
    const prev = anteriorMap.get(r.comitente);
    const aumAnterior = prev ? prev.aum : 0;
    const diff = r.aum - aumAnterior;
    const pct = aumAnterior !== 0 ? (diff / aumAnterior) * 100 : null;
    return {
      comitente: r.comitente,
      cuenta: r.cuenta,
      aumAnterior,
      aumActual: r.aum,
      diferencia: diff,
      variacion: pct
    };
  });

  const delta = merged.reduce((acc, r) => acc + r.diferencia, 0);
  const totalAnterior = merged.reduce((acc, r) => acc + r.aumAnterior, 0);
  const deltaPct = totalAnterior !== 0 ? (delta / totalAnterior) * 100 : null;

  setText("kpiDeltaAum", formatUsd(delta));
  setText("kpiDeltaPct", deltaPct === null ? "-" : `${deltaPct.toFixed(2)}%`);

  document.getElementById("comparacionSection").classList.remove("hidden");

  const tabla = [...merged].sort((a,b) => b.diferencia - a.diferencia).slice(0, topN);

  renderTable("tablaComparacion", tabla, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${formatUsd(row.aumAnterior)}</td>
      <td>${formatUsd(row.aumActual)}</td>
      <td>${formatUsd(row.diferencia)}</td>
      <td>${row.variacion === null ? "-" : row.variacion.toFixed(2) + "%"}</td>
    </tr>
  `);
}

function renderTable(tableId, rows, rowRenderer) {
  const tbody = document.querySelector(`#${tableId} tbody`);
  if (!tbody) return;
  tbody.innerHTML = rows.map(rowRenderer).join("");
}

function sum(rows, key) {
  return rows.reduce((acc, r) => acc + (Number(r[key]) || 0), 0);
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function formatUsd(n) {
  return new Intl.NumberFormat("es-AR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number(n || 0));
}

function formatInt(n) {
  return new Intl.NumberFormat("es-AR", {
    maximumFractionDigits: 0
  }).format(Number(n || 0));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
