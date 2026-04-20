const state = {
  actualData: [],
  anteriorData: []
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

async function handleFile(e, mode) {
  try {
    const file = e.target.files?.[0];
    if (!file) return;

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!json.length) {
      alert("El archivo no contiene datos.");
      return;
    }

    validateColumns(json[0]);

    const rows = json.map(r => ({
      comitente: normalizeText(r["comitente"]),
      cuenta: normalizeText(r["cuenta"]),
      esJuridica: normalizeText(r["Es Juridica"]),
      arancel: normalizeText(r["arancel"]),
      aum: normalizeNumber(r["AUM en Dolares"]),
      cv7000: normalizeNumber(r["cv7000"]),
      pesosOperables: normalizeNumber(r["$ Operables CI"]),
      mepOperables: normalizeNumber(r["MEP Operables CI"]),
      comision180: normalizeNumber(r["Comision 180"]),
      tipoCbioMep: normalizeNumber(r["Tipo Cbio MEP"])
    }));

    if (mode === "actual") {
      state.actualData = rows;
      document.getElementById("actualFileName").textContent = `Excel actual: ${file.name}`;
      populateArancelFilter(rows);
    } else {
      state.anteriorData = rows;
      document.getElementById("anteriorFileName").textContent = `Excel anterior: ${file.name}`;
    }

    document.getElementById("statusText").textContent = "Archivo procesado correctamente";
    renderDashboard();
  } catch (err) {
    document.getElementById("statusText").textContent = "Error al procesar archivo";
    alert(err.message || "Error al procesar el archivo.");
    console.error(err);
  }
}

function validateColumns(exampleRow) {
  const cols = Object.keys(exampleRow);
  const missing = REQUIRED_COLUMNS.filter(c => !cols.includes(c));
  if (missing.length) {
    throw new Error("Faltan columnas obligatorias en el Excel: " + missing.join(", "));
  }
}

function calcRoa(row) {
  if (!row || !row.aum || row.aum <= 0) return null;
  return ((row.comision180 * 2) / row.aum) * 100;
}

function benchmarkAnual(aum) {
  return aum * 0.015;
}

function benchmarkMensual(aum) {
  return benchmarkAnual(aum) / 12;
}

function roaSemaforo(roa) {
  if (roa === null || roa === undefined || !Number.isFinite(roa)) return "roa-sin-dato";
  if (roa < 1.1) return "roa-rojo";
  if (roa < 1.3) return "roa-amarillo";
  if (roa <= 2.0) return "roa-verde";
  return "roa-amarillo";
}

function populateArancelFilter(rows) {
  const select = document.getElementById("filtroArancel");
  const current = select.value;
  const aranceles = [...new Set(rows.map(r => r.arancel).filter(Boolean))].sort();

  select.innerHTML = `<option value="TODOS">Todos</option>` +
    aranceles.map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join("");

  select.value = aranceles.includes(current) ? current : "TODOS";
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

  const topN = Number(document.getElementById("filtroTopN").value || 15);

  const rows = getFilteredData().map(r => {
    const pesosOperablesUsd = r.tipoCbioMep > 0 ? (r.pesosOperables / r.tipoCbioMep) : 0;
    const liquidezTotal = r.cv7000 + pesosOperablesUsd + r.mepOperables;
    const roa = calcRoa(r);
    return {
      ...r,
      pesosOperablesUsd,
      liquidezTotal,
      roa,
      benchmarkAnual: benchmarkAnual(r.aum),
      benchmarkMensual: benchmarkMensual(r.aum)
    };
  });

  updateKpis(rows);
  renderLiquidezTable(rows, topN);
  renderTopAumTable(rows, topN);
  renderInactivosTable(rows);
  renderComparacionAum(rows, topN);
}

function updateKpis(rows) {
  const totalClientes = rows.length;
  const juridicas = rows.filter(r => r.esJuridica === "1").length;
  const noJuridicas = rows.filter(r => r.esJuridica === "0").length;
  const aum = sum(rows, "aum");
  const liquidez = rows.reduce((acc, r) => acc + r.liquidezTotal, 0);
  const comision = sum(rows, "comision180");
  const inactivos = rows.filter(r => r.comision180 < 15).length;
  const roaTotal = aum > 0 ? ((comision * 2) / aum) * 100 : null;

  setText("kpiClientes", formatInt(totalClientes));
  setText("kpiJuridicas", formatInt(juridicas));
  setText("kpiNoJuridicas", formatInt(noJuridicas));
  setText("kpiAum", formatMonto(aum));
  setText("kpiLiquidez", formatMonto(liquidez));
  setText("kpiComision", formatMonto(comision));
  setText("kpiInactivos", formatInt(inactivos));
  setText("kpiRoa", formatPct(roaTotal));
  setText("kpiBenchmarkAnual", formatMonto(benchmarkAnual(aum)));
  setText("kpiBenchmarkMensual", formatMonto(benchmarkMensual(aum)));
}

function renderLiquidezTable(rows, topN) {
  const top = [...rows].sort((a, b) => b.liquidezTotal - a.liquidezTotal).slice(0, topN);
  renderTable("tablaLiquidez", top, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatMonto(row.cv7000)}</td>
      <td>${formatMonto(row.pesosOperablesUsd)}</td>
      <td>${formatMonto(row.mepOperables)}</td>
      <td>${formatMonto(row.tipoCbioMep)}</td>
      <td>${formatMonto(row.liquidezTotal)}</td>
      <td>${formatMonto(row.comision180)}</td>
      <td><span class="roa-badge ${roaSemaforo(row.roa)}">${formatPct(row.roa)}</span></td>
      <td>${formatMonto(row.benchmarkMensual)}</td>
    </tr>
  `);
}

function renderTopAumTable(rows, topN) {
  const top = [...rows].sort((a, b) => b.aum - a.aum).slice(0, topN);
  renderTable("tablaTopAum", top, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${row.esJuridica === "1" ? "Sí" : "No"}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatMonto(row.aum)}</td>
      <td><span class="roa-badge ${roaSemaforo(row.roa)}">${formatPct(row.roa)}</span></td>
      <td>${formatMonto(row.benchmarkMensual)}</td>
    </tr>
  `);
}

function renderInactivosTable(rows) {
  const inactivos = [...rows]
    .filter(r => r.comision180 < 15)
    .sort((a, b) => a.comision180 - b.comision180);

  renderTable("tablaInactivos", inactivos, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${row.esJuridica === "1" ? "Sí" : "No"}</td>
      <td>${escapeHtml(row.arancel)}</td>
      <td>${formatMonto(row.aum)}</td>
      <td>${formatMonto(row.comision180)}</td>
      <td>${formatMonto(row.liquidezTotal)}</td>
      <td><span class="roa-badge ${roaSemaforo(row.roa)}">${formatPct(row.roa)}</span></td>
      <td>${formatMonto(row.benchmarkMensual)}</td>
    </tr>
  `);
}

function renderComparacionAum(actualRows, topN) {
  if (!state.anteriorData.length) {
    document.getElementById("comparacionSection").classList.add("hidden");
    setText("kpiDeltaAum", "-");
    setText("kpiDeltaPct", "-");
    return;
  }

  const anteriorMap = new Map();
  state.anteriorData.forEach(r => anteriorMap.set(r.comitente, r));

  const merged = actualRows.map(r => {
    const prev = anteriorMap.get(r.comitente);
    const aumAnterior = prev ? prev.aum : 0;
    const diff = r.aum - aumAnterior;
    const pct = aumAnterior !== 0 ? (diff / aumAnterior) * 100 : null;
    return { comitente: r.comitente, cuenta: r.cuenta, aumAnterior, aumActual: r.aum, diferencia: diff, variacion: pct };
  });

  const delta = merged.reduce((acc, r) => acc + r.diferencia, 0);
  const totalAnterior = merged.reduce((acc, r) => acc + r.aumAnterior, 0);
  const deltaPct = totalAnterior !== 0 ? (delta / totalAnterior) * 100 : null;

  setText("kpiDeltaAum", formatMonto(delta));
  setText("kpiDeltaPct", deltaPct === null ? "-" : `${deltaPct.toFixed(1)}%`);

  document.getElementById("comparacionSection").classList.remove("hidden");

  const tabla = [...merged].sort((a, b) => b.diferencia - a.diferencia).slice(0, topN);
  renderTable("tablaComparacion", tabla, row => `
    <tr>
      <td>${escapeHtml(row.comitente)}</td>
      <td>${escapeHtml(row.cuenta)}</td>
      <td>${formatMonto(row.aumAnterior)}</td>
      <td>${formatMonto(row.aumActual)}</td>
      <td>${formatMonto(row.diferencia)}</td>
      <td>${row.variacion === null ? "-" : row.variacion.toFixed(1) + "%"}</td>
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

function formatMonto(n) {
  return new Intl.NumberFormat("es-AR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(Math.round(Number(n || 0)));
}

function formatPct(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "-";
  return `${n.toFixed(1)}%`;
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
