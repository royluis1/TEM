// Pegar en Google Sheets > Extensiones > Apps Script
const HOJA = "Resultados";
const NIVELES = ["nivel-1", "nivel-2", "nivel-3"];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) { sh = ss.insertSheet(HOJA); sh.appendRow(["Fecha", "Nombre", "Nivel", "Puntaje", "Total"]); }
  return sh;
}
function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function abrev_(n) {
  const p = String(n).trim().split(/\s+/);
  return p.length > 1 ? p[0] + " " + p[p.length - 1].charAt(0).toUpperCase() + "." : p[0];
}
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const nombre = String(d.name || "").trim().replace(/^[=+\-@]+/, "").slice(0, 40);
    const s = Number(d.s), t = Number(d.t);
    if (!nombre || NIVELES.indexOf(d.k) < 0 || !Number.isInteger(s) || !Number.isInteger(t) || t < 1 || t > 50 || s < 0 || s > t) return salida_({ ok: false });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    hoja_().appendRow([new Date(), nombre, d.k, s, t]);
    lock.releaseLock();
    return salida_({ ok: true });
  } catch (err) { return salida_({ ok: false }); }
}
function doGet() {
  const filas = hoja_().getDataRange().getValues().slice(1);
  const mejor = {};
  filas.forEach(function (r) {
    const nombre = String(r[1]);
    const clave = nombre.toLowerCase() + "|" + r[2];
    const fila = { name: abrev_(nombre), k: r[2], s: r[3], t: r[4], ts: new Date(r[0]).getTime() };
    if (!mejor[clave] || fila.s > mejor[clave].s) mejor[clave] = fila;
  });
  return salida_({ rows: Object.keys(mejor).map(function (k) { return mejor[k]; }) });
}
