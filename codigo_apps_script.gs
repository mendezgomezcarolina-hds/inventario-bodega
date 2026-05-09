// ============================================================
//  GOOGLE APPS SCRIPT - VERSIÓN FINAL
//  Compatible con PC, Android e iPhone (Safari)
//  Implementar → Aplicación web → Cualquier persona
// ============================================================

const SHEET_DATOS   = "Hoja1";
const SHEET_LUGARES = "Hoja2";
const SHEET_ITEMS   = "Hoja3";

function doGet(e) {
  const accion   = e.parameter.accion   || "";
  const callback = e.parameter.callback || ""; // para JSONP (iPhone)

  if (accion === "lugares") return responder(leerColumnaA(SHEET_LUGARES, "lugares", 1), callback);
  if (accion === "items")   return responder(leerColumnaA(SHEET_ITEMS,   "items",   2), callback);

  return escribir(e, callback);
}

function doPost(e) { return escribir(e, ""); }

// ── Responder: soporta JSON normal y JSONP ────────────────────
function responder(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    // JSONP: envuelve en función para que Safari pueda leerlo
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Leer columna A de una hoja ────────────────────────────────
function leerColumnaA(nombreHoja, clave, respaldo) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(nombreHoja) || ss.getSheets()[respaldo];
    const lista = sheet.getDataRange().getValues()
      .map(f => f[0]).filter(v => v !== "" && v != null);
    const r = { status: "ok" };
    r[clave] = lista;
    return r;
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Escribir registro ─────────────────────────────────────────
function escribir(e, callback) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
    const p     = e.parameter || {};

    if (!p.lugar && !p.item && !p.cantidad) throw new Error("Sin datos.");

    if (sheet.getLastRow() === 0)
      sheet.appendRow(["Lugar","Ítem","Cantidad","Fecha/Hora","Usuario"]);

    sheet.appendRow([
      p.lugar || "", p.item || "", p.cantidad || "",
      new Date().toLocaleString("es-CL"), p.usuario || "Anónimo"
    ]);

    return responder({ status: "ok" }, callback);
  } catch(err) {
    return responder({ status: "error", mensaje: err.toString() }, callback);
  }
}

function testWrite() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.appendRow(["TEST","TEST","1", new Date().toLocaleString("es-CL"), "Test"]);
  Logger.log("OK → " + sheet.getName());
}
