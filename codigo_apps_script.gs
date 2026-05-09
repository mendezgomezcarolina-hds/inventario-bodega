// ============================================================
//  GOOGLE APPS SCRIPT - SISTEMA DE GESTIÓN DE INSUMOS
//  Implementar → Aplicación web → Cualquier persona
// ============================================================

const SHEET_DATOS         = "PEDIDOS";
const SHEET_LUGARES       = "LUGAR";
const SHEET_ITEMS         = "INSUMOS";
const SHEET_COLABORADORES = "COLABORADORES";
const SHEET_SOLICITUDES   = "SOLICITUDES";

function doGet(e) {
  const accion   = e.parameter.accion   || "";
  const callback = e.parameter.callback || "";

  if (accion === "lugares")           return responder(leerColumnaA(SHEET_LUGARES,       "lugares",       1), callback);
  if (accion === "items")             return responder(leerColumnaA(SHEET_ITEMS,         "items",         2), callback);
  if (accion === "colaboradores")     return responder(leerColaboradores(),                                   callback);
  if (accion === "solicitud")         return responder(escribirSolicitud(e),                                  callback);
  if (accion === "listarSolicitudes") return responder(listarSolicitudes(e),                                  callback);
  if (accion === "actualizarEstado")  return responder(actualizarEstado(e),                                   callback);

  return responder(escribir(e), callback);
}

function doPost(e) { return responder(escribir(e), ""); }

// ── Responder: JSON normal y JSONP ────────────────────────────
function responder(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Leer colaboradores: columna A (Nombre) + B (ID) ──────────
function leerColaboradores() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_COLABORADORES) || ss.getSheets()[3];
    const datos = sheet.getDataRange().getValues();
    const colaboradores = datos
      .filter(f => f[0] !== "" && f[0] != null)
      .map(f => ({ nombre: f[0], id: f[1] || "" }));
    return { status: "ok", colaboradores };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
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

// ── Formatear fecha ISO → DD/MM/AAAA ─────────────────────────
function formatearFecha(iso) {
  if (!iso) return "";
  const partes = iso.split("-");
  if (partes.length !== 3) return iso;
  return partes[2] + "/" + partes[1] + "/" + partes[0];
}

// ── Escribir inventario (PEDIDOS) ────────────────────────────
function escribir(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
    const p     = e.parameter || {};

    if (!p.lugar && !p.item && !p.cantidad) throw new Error("Sin datos.");

    if (sheet.getLastRow() === 0)
      sheet.appendRow(["Lugar","Ítem","Cantidad","Vencimiento","Fecha/Hora","Responsable","ID"]);

    sheet.appendRow([
      p.lugar       || "",
      p.item        || "",
      p.cantidad    || "",
      formatearFecha(p.vencimiento),
      new Date().toLocaleString("es-CL"),
      p.usuario     || "Anónimo",
      p.usuarioId   || ""
    ]);

    return { status: "ok" };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Escribir solicitud (SOLICITUDES) ─────────────────────────
// Columnas: A=ID | B=Lugar | C=Insumo | D=Cantidad | E=Responsable | F=ID_Resp | G=Fecha | H=Estado | I=Fecha Resol | J=Supervisor
function escribirSolicitud(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) sheet = ss.insertSheet(SHEET_SOLICITUDES);

    const p = e.parameter || {};
    if (!p.item || !p.cantidad) throw new Error("Sin datos de insumo.");

    if (sheet.getLastRow() === 0)
      sheet.appendRow(["ID Solicitud","Lugar","Insumo","Cantidad","Responsable","ID Responsable","Fecha Solicitud","Estado","Fecha Resolución","Supervisor"]);

    sheet.appendRow([
      p.idSolicitud  || "",
      p.lugar        || "",
      p.item         || "",
      p.cantidad     || "",
      p.usuario      || "Anónimo",
      p.usuarioId    || "",
      new Date().toLocaleString("es-CL"),
      "PENDIENTE",
      "",
      ""
    ]);

    return { status: "ok" };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Listar solicitudes (con filtro opcional de estado) ────────
function listarSolicitudes(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "ok", solicitudes: [] };

    const filtroEstado = (e.parameter.estado || "").toUpperCase();
    const datos = sheet.getDataRange().getValues();
    const encab = datos[0];
    const filas = datos.slice(1);

    const solicitudes = filas
      .map((f, i) => ({ f, fila: i + 2 }))
      .filter(({ f }) => f[0] !== "" && (!filtroEstado || f[7] === filtroEstado))
      .map(({ f, fila }) => ({
        fila:         fila,
        id:           f[0],
        lugar:        f[1],
        item:         f[2],
        cantidad:     f[3],
        responsable:  f[4],
        idResp:       f[5],
        fecha:        f[6] instanceof Date ? f[6].toLocaleString("es-CL") : f[6],
        estado:       f[7],
        fechaResol:   f[8] instanceof Date ? f[8].toLocaleString("es-CL") : (f[8] || ""),
        supervisor:   f[9] || ""
      }));

    return { status: "ok", solicitudes };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Actualizar estado de ítem individual (por número de fila) ─
function actualizarEstado(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) throw new Error("Hoja SOLICITUDES no existe.");

    const p           = e.parameter || {};
    const nuevoEstado = (p.estado || "").toUpperCase();
    const supervisor  = p.supervisor || "";
    const fila        = parseInt(p.fila || "0");

    if (!nuevoEstado) throw new Error("Falta parámetro estado.");

    // Modo 1: fila específica (ítem individual)
    if (fila > 1) {
      sheet.getRange(fila, 8).setValue(nuevoEstado);
      sheet.getRange(fila, 9).setValue(new Date().toLocaleString("es-CL"));
      sheet.getRange(fila, 10).setValue(supervisor);
      return { status: "ok", actualizados: 1 };
    }

    // Modo 2: por idSolicitud (todos los ítems del grupo)
    const idBuscado = p.idSolicitud || "";
    if (!idBuscado) throw new Error("Falta fila o idSolicitud.");
    const datos = sheet.getDataRange().getValues();
    let actualizados = 0;
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === idBuscado) {
        sheet.getRange(i + 1, 8).setValue(nuevoEstado);
        sheet.getRange(i + 1, 9).setValue(new Date().toLocaleString("es-CL"));
        sheet.getRange(i + 1, 10).setValue(supervisor);
        actualizados++;
      }
    }
    if (actualizados === 0) throw new Error("ID no encontrado: " + idBuscado);
    return { status: "ok", actualizados };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Test manual ───────────────────────────────────────────────
function testWrite() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATOS)
    || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.appendRow(["TEST","TEST","1","","", new Date().toLocaleString("es-CL"),"Test",""]);
  Logger.log("OK → " + sheet.getName());
}
