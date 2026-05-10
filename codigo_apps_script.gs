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
  if (accion === "items")             return responder(leerItems(),                                          callback);
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

// ── Leer insumos: columna A (Código) + B (Descripción) ──────
function leerItems() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ITEMS) || ss.getSheets()[2];
    const datos = sheet.getDataRange().getValues();
    const items = datos
      .filter(f => f[0] !== "" && f[0] != null)
      .map(f => ({ codigo: String(f[0]), descripcion: String(f[1] || "") }));
    return { status: "ok", items };
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
      sheet.appendRow(["Lugar","Código","Ítem","Cantidad","Vencimiento","Fecha/Hora","Responsable","ID"]);

    sheet.appendRow([
      p.lugar        || "",
      p.codigo       || p.item || "",
      p.descripcion  || p.item || "",
      p.cantidad     || "",
      formatearFecha(p.vencimiento),
      new Date().toLocaleString("es-CL"),
      p.usuario      || "Anónimo",
      p.usuarioId    || ""
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
      sheet.appendRow(["ID Solicitud","Lugar","Código","Insumo","Cantidad","Responsable","ID Responsable","Fecha Solicitud","Estado","Fecha Resolución","Supervisor"]);

    sheet.appendRow([
      p.idSolicitud  || "",
      p.lugar        || "",
      p.codigo       || p.item || "",
      p.descripcion  || p.item || "",
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
      .filter(({ f }) => f[0] !== "" && (!filtroEstado || f[8] === filtroEstado))
      .map(({ f, fila }) => ({
        fila:         fila,
        id:           f[0],
        lugar:        f[1],
        codigo:       f[2],
        item:         f[3],
        cantidad:     f[4],
        responsable:  f[5],
        idResp:       f[6],
        fecha:        f[7] instanceof Date ? f[7].toLocaleString("es-CL") : f[7],
        estado:       f[8],
        fechaResol:   f[9] instanceof Date ? f[9].toLocaleString("es-CL") : (f[9] || ""),
        supervisor:   f[10] || ""
      }));

    return { status: "ok", solicitudes };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Actualizar estado por idSolicitud + item + lugar ─────────
// No depende de número de fila - funciona con cualquier versión
function actualizarEstado(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) throw new Error("Hoja SOLICITUDES no existe.");

    const p           = e.parameter || {};
    const nuevoEstado = (p.estado       || "").toUpperCase();
    const supervisor  = p.supervisor    || "";
    const idBuscado   = p.idSolicitud   || "";
    const itemBuscado = p.item          || "";
    const lugarBuscado= p.lugar         || "";

    if (!nuevoEstado || !idBuscado || !itemBuscado)
      throw new Error("Faltan parámetros: estado, idSolicitud, item.");

    const datos = sheet.getDataRange().getValues();
    let actualizados = 0;

    for (let i = 1; i < datos.length; i++) {
      const coincide = datos[i][0] === idBuscado &&
                       datos[i][3] === itemBuscado &&   // col D = Insumo (descripcion)
                       (lugarBuscado === "" || datos[i][1] === lugarBuscado);
      if (coincide) {
        sheet.getRange(i + 1, 9).setValue(nuevoEstado);                          // col I = Estado
        sheet.getRange(i + 1, 10).setValue(new Date().toLocaleString("es-CL")); // col J = Fecha Resol
        sheet.getRange(i + 1, 11).setValue(supervisor);                          // col K = Supervisor
        actualizados++;
      }
    }

    if (actualizados === 0) throw new Error("Fila no encontrada.");
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
