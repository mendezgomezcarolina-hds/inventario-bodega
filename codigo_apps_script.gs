// ============================================================
//  GOOGLE APPS SCRIPT - SISTEMA DE GESTIÓN DE INSUMOS
//  Implementar → Aplicación web → Cualquier persona
// ============================================================

const SHEET_DATOS         = "INVENTARIO";
const SHEET_LUGARES       = "LUGAR";
const SHEET_ITEMS         = "INSUMOS";
const SHEET_COLABORADORES = "COLABORADORES";
const SHEET_SOLICITUDES   = "SOLICITUDES";
const SHEET_RECEPCION     = "RECEPCION_BODEGA";
const SHEET_MOVIMIENTOS   = "MOVIMIENTOS";

function doGet(e) {
  const accion   = e.parameter.accion   || "";
  const callback = e.parameter.callback || "";

  if (accion === "lugares")             return responder(leerColumnaA(SHEET_LUGARES, "lugares", 1), callback);
  if (accion === "items")               return responder(leerItems(),                               callback);
  if (accion === "colaboradores")       return responder(leerColaboradores(),                       callback);
  if (accion === "solicitud")           return responder(escribirSolicitud(e),                      callback);
  if (accion === "listarSolicitudes")   return responder(listarSolicitudes(e),                      callback);
  if (accion === "actualizarEstado")    return responder(actualizarEstado(e),                       callback);
  if (accion === "listarRecepcion")     return responder(listarRecepcion(e),                        callback);
  if (accion === "actualizarRecepcion") return responder(actualizarRecepcion(e),                    callback);
  if (accion === "listarMovimientos")   return responder(listarMovimientos(),                        callback);

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

// ── Leer insumos: columna A (Código) + B (Descripción) ───────
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

// ── Escribir inventario (INVENTARIO) ─────────────────────────
function escribir(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
    const p     = e.parameter || {};
    if (!p.lugar && !p.item && !p.cantidad) throw new Error("Sin datos.");
    if (sheet.getLastRow() === 0)
      sheet.appendRow(["Lugar","Código","Ítem","Cantidad","Vencimiento","Fecha/Hora","Responsable","ID"]);
    sheet.appendRow([
      p.lugar       || "",
      p.codigo      || p.item || "",
      p.descripcion || p.item || "",
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
function escribirSolicitud(e) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) sheet = ss.insertSheet(SHEET_SOLICITUDES);
    const p = e.parameter || {};
    if (!p.item || !p.cantidad) throw new Error("Sin datos de insumo.");
    if (sheet.getLastRow() === 0)
      sheet.appendRow(["ID Solicitud","Lugar","Código","Insumo","Cantidad","Responsable","ID Responsable","Fecha Solicitud","Estado","Fecha Resolución","Supervisor"]);
    sheet.appendRow([
      p.idSolicitud || "",
      p.lugar       || "",
      p.codigo      || p.item || "",
      p.descripcion || p.item || "",
      p.cantidad    || "",
      p.usuario     || "Anónimo",
      p.usuarioId   || "",
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

// ── Listar solicitudes ────────────────────────────────────────
function listarSolicitudes(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "ok", solicitudes: [] };
    const filtroEstado = (e.parameter.estado || "").toUpperCase();
    const datos  = sheet.getDataRange().getValues();
    const inicio = String(datos[0][0]).indexOf("SOL-") === 0 ? 0 : 1;
    const solicitudes = [];
    for (var i = inicio; i < datos.length; i++) {
      const f = datos[i];
      if (!f[0] || String(f[0]).trim() === "") continue;
      const estado = String(f[8] || "").trim().toUpperCase();
      if (filtroEstado && estado !== filtroEstado) continue;
      const fechaH = f[7] instanceof Date ? f[7].toLocaleString("es-CL") : String(f[7] || "");
      const fechaR = f[9] instanceof Date ? f[9].toLocaleString("es-CL") : String(f[9] || "");
      solicitudes.push({
        fila: i + 1, id: String(f[0]||""), lugar: String(f[1]||""),
        codigo: String(f[2]||""), item: String(f[3]||""), cantidad: String(f[4]||""),
        responsable: String(f[5]||""), idResp: String(f[6]||""),
        fecha: fechaH, estado: String(f[8]||""), fechaResol: fechaR, supervisor: String(f[10]||"")
      });
    }
    return { status: "ok", solicitudes };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Actualizar estado solicitud ───────────────────────────────
function actualizarEstado(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) throw new Error("Hoja SOLICITUDES no existe.");
    const p           = e.parameter || {};
    const nuevoEstado = (p.estado     || "").toUpperCase();
    const supervisor  = p.supervisor  || "";
    const idBuscado   = p.idSolicitud || "";
    const itemBuscado = p.item        || "";
    const lugarBusc   = p.lugar       || "";
    if (!nuevoEstado || !idBuscado) throw new Error("Faltan parámetros.");
    const datos = sheet.getDataRange().getValues();
    let actualizados = 0;
    for (let i = 0; i < datos.length; i++) {
      const f = datos[i];
      const coincide = String(f[0]) === idBuscado &&
                       (!itemBuscado || String(f[3]) === itemBuscado) &&
                       (!lugarBusc   || String(f[1]) === lugarBusc);
      if (coincide) {
        sheet.getRange(i+1, 9).setValue(nuevoEstado);
        sheet.getRange(i+1, 10).setValue(new Date().toLocaleString("es-CL"));
        sheet.getRange(i+1, 11).setValue(supervisor);
        actualizados++;
      }
    }
    if (actualizados === 0) throw new Error("Fila no encontrada: " + idBuscado + " / " + itemBuscado);
    return { status: "ok", actualizados };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Listar RECEPCION_BODEGA por mes ──────────────────────────
// Columnas: A=Mes B=N°Sol C=Lugar D=Codigo E=Descripcion
//           F=CantSolicitada G=CantRecibida H=Estado I=Fecha/Hora
function listarRecepcion(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RECEPCION);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "ok", pedidos: [] };
    const mes   = (e.parameter.mes || "").trim().toUpperCase();
    const datos = sheet.getDataRange().getValues();
    const inicio = String(datos[0][0]).toUpperCase() === "MES" ? 1 : 0;
    const pedidos = [];
    for (var i = inicio; i < datos.length; i++) {
      const f = datos[i];
      if (!f[0] && !f[1] && !f[3] && !f[4]) continue;
      const mesFila    = String(f[0] || "").trim().toUpperCase();
      const estadoFila = String(f[7] || "").trim().toUpperCase() || "PENDIENTE";
      if (mes && mesFila !== mes) continue;
      if (estadoFila !== "PENDIENTE") continue;
      if (!f[3] && !f[4]) continue;
      const fecha = f[8] instanceof Date ? f[8].toLocaleString("es-CL") : String(f[8] || "");
      pedidos.push({
        fila: i+1, mes: mesFila, id: String(f[1]||""), lugar: String(f[2]||""),
        codigo: String(f[3]||""), item: String(f[4]||""),
        cantSolicitada: String(f[5]||"").replace(/-/g,"").trim(),
        cantRecibida:   String(f[6]||"").replace(/-/g,"").trim(),
        estado: estadoFila, fecha: fecha
      });
    }
    return { status: "ok", pedidos };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Actualizar recepcion + escribir INGRESO en MOVIMIENTOS ───
function actualizarRecepcion(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RECEPCION);
    if (!sheet) throw new Error("Hoja RECEPCION_BODEGA no existe.");
    const p    = e.parameter || {};
    const fila = parseInt(p.fila || "0");
    const est  = (p.estado      || "").toUpperCase();
    const cant = p.cantRecibida || "";
    if (!fila || !est) throw new Error("Faltan fila o estado.");

    // Leer datos de la fila
    const filaDatos = sheet.getRange(fila, 1, 1, 9).getValues()[0];
    const mes    = String(filaDatos[0] || "");
    const nSol   = String(filaDatos[1] || "");
    const lugar  = String(filaDatos[2] || "");
    const codigo = String(filaDatos[3] || "");
    const descr  = String(filaDatos[4] || "");

    // Actualizar RECEPCION_BODEGA
    sheet.getRange(fila, 7).setValue(cant);
    sheet.getRange(fila, 8).setValue(est);
    sheet.getRange(fila, 9).setValue(new Date().toLocaleString("es-CL"));

    const fechaVenc = p.fechaVencimiento || "";

    // Si APROBADO → registrar INGRESO en MOVIMIENTOS
    if (est === "APROBADO") {
      let movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
      if (!movSheet) movSheet = ss.insertSheet(SHEET_MOVIMIENTOS);
      if (movSheet.getLastRow() === 0)
        movSheet.appendRow(["Fecha/Hora","Tipo","N° Sol","Mes","Lugar","Código","Descripción","Cantidad","Fecha Vencimiento"]);
      movSheet.appendRow([
        new Date().toLocaleString("es-CL"), "INGRESO",
        nSol, mes, lugar, codigo, descr, cant, fechaVenc
      ]);
    }

    return { status: "ok", fila };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Listar MOVIMIENTOS ────────────────────────────────────────
function listarMovimientos() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    if (!sheet || sheet.getLastRow() === 0) return { status: "ok", movimientos: [], total: 0 };
    const datos  = sheet.getDataRange().getValues();
    const inicio = String(datos[0][0]).toUpperCase().indexOf("FECHA") === 0 ? 1 : 0;
    const movimientos = datos.slice(inicio).map(f => ({
      fecha:       f[0] instanceof Date ? f[0].toLocaleString("es-CL") : String(f[0]||""),
      tipo:        String(f[1]||""),
      nSol:        String(f[2]||""),
      mes:         String(f[3]||""),
      lugar:       String(f[4]||""),
      codigo:      String(f[5]||""),
      descripcion: String(f[6]||""),
      cantidad:    String(f[7]||"")
    }));
    return { status: "ok", total: movimientos.length, movimientos };
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

// ── Test MOVIMIENTOS (ejecutar directo desde el editor) ───────
function testMovimientos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log("1. Spreadsheet: " + ss.getName());
    
    // Intentar obtener o crear hoja MOVIMIENTOS
    let movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    Logger.log("2. Hoja MOVIMIENTOS existe: " + (movSheet !== null));
    
    if (!movSheet) {
      movSheet = ss.insertSheet(SHEET_MOVIMIENTOS);
      Logger.log("3. Hoja creada OK");
    }
    
    // Intentar escribir una fila de prueba
    if (movSheet.getLastRow() === 0) {
      movSheet.appendRow(["Fecha/Hora","Tipo","N° Sol","Mes","Lugar","Código","Descripción","Cantidad"]);
      Logger.log("4. Encabezados escritos");
    }
    
    movSheet.appendRow([
      new Date().toLocaleString("es-CL"),
      "TEST-INGRESO",
      "S_TEST",
      "JUNIO",
      "BODEGA TEST",
      "9999999",
      "ITEM DE PRUEBA",
      "99"
    ]);
    
    Logger.log("5. Fila de prueba escrita OK. Total filas: " + movSheet.getLastRow());
    
  } catch(err) {
    Logger.log("ERROR: " + err.toString());
  }
}
