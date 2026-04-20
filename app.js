document.getElementById("fileActual").addEventListener("change", handleFile);

let data = [];

function handleFile(e) {
 const file = e.target.files[0];
 const reader = new FileReader();

 reader.onload = function(evt) {
  const wb = XLSX.read(evt.target.result, { type: "binary" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet);

  data = json.map(r => ({
    cliente: r["comitente"],
    aum: Number(r["AUM en Dolares"]) || 0,
    comision180: Number(r["Comision 180"]) || 0,
    comision1y: Number(r["Comision 1y"]) || 0,
    pesos: Number(r["$ Operables CI"]) || 0,
    mep: Number(r["MEP Operables CI"]) || 0,
    tc: Number(r["Tipo Cbio MEP"]) || 1
  }));

  render();
 };

 reader.readAsBinaryString(file);
}

function calcROA(r) {
 if (r.aum === 0) return 0;
 return (r.comision1y / r.aum) * 100;
}

function render() {
 const enriched = data.map(r => {
  const liquidez = (r.pesos / r.tc) + r.mep;
  const roa = calcROA(r);
  return {...r, liquidez, roa};
 });

 document.getElementById("kpiClientes").innerText = enriched.length;

 const totalAUM = enriched.reduce((a,b)=>a+b.aum,0);
 const totalCom = enriched.reduce((a,b)=>a+b.comision1y,0);
 const roaTotal = (totalCom / totalAUM) * 100;

 document.getElementById("kpiAum").innerText = Math.round(totalAUM);
 document.getElementById("kpiRoa").innerText = roaTotal.toFixed(1)+"%";

 const inactivos = enriched.filter(r => r.comision180 < 15);
 document.getElementById("kpiInactivos").innerText = inactivos.length;

 document.querySelector("#tablaLiquidez tbody").innerHTML =
 enriched.slice(0,10).map(r => `
 <tr><td>${r.cliente}</td><td>${Math.round(r.liquidez)}</td><td>${r.roa.toFixed(1)}%</td></tr>`).join("");

 document.querySelector("#tablaTopAum tbody").innerHTML =
 enriched.sort((a,b)=>b.aum-a.aum).slice(0,10).map(r => `
 <tr><td>${r.cliente}</td><td>${Math.round(r.aum)}</td><td>${r.roa.toFixed(1)}%</td></tr>`).join("");

 document.querySelector("#tablaInactivos tbody").innerHTML =
 inactivos.map(r => `
 <tr><td>${r.cliente}</td><td>${Math.round(r.aum)}</td><td>${r.roa.toFixed(1)}%</td></tr>`).join("");
}
