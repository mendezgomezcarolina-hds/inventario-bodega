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
  if (accion === "itemsPorLugar")       return responder(leerItemsPorLugar(e),                      callback);
  if (accion === "colaboradores")       return responder(leerColaboradores(),                       callback);
  if (accion === "accesos")             return responder(leerAccesos(),                            callback);
  if (accion === "lugaresConInventario") return responder(lugaresConInventario(),                  callback);
  if (accion === "stockLugar")          return responder(stockLugar(e),                           callback);
  if (accion === "registrarEgreso")          return responder(registrarEgreso(e),                      callback);
  if (accion === "registrarCorreccion")        return responder(registrarCorreccion(e),                callback);
  if (accion === "notificarResumenSolicitud")  return responder(notificarResumenSolicitud(e),           callback);
  if (accion === "notificarResumenSolicitud") return responder(notificarResumenSolicitud(e),             callback);
  if (accion === "listarSolicitudesPorLugar") return responder(listarSolicitudesPorLugar(e),             callback);
  if (accion === "login")                     return responder(login(e),                               callback);
  if (accion === "obtenerAccesos")            return responder(obtenerAccesos(e),                      callback);
  if (accion === "guardarAccesos")            return responder(guardarAccesos(e),                      callback);
  if (accion === "recepcionarSolicitud")      return responder(recepcionarSolicitud(e),                  callback);
  if (accion === "solicitud")           return responder(escribirSolicitud(e),                      callback);
  if (accion === "listarSolicitudes")   return responder(listarSolicitudes(e),                      callback);
  if (accion === "actualizarEstado")    return responder(actualizarEstado(e),                       callback);
  if (accion === "listarRecepcion")     return responder(listarRecepcion(e),                        callback);
  if (accion === "actualizarRecepcion") return responder(actualizarRecepcion(e),                    callback);
  if (accion === "listarMovimientos")   return responder(listarMovimientos(),                        callback);
  if (accion === "verificarAcceso")    return responder(verificarAcceso(e),                       callback);
  if (accion === "enviarReporte")      return responder(enviarReporte(e),                          callback);

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
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_COLABORADORES) || ss.getSheets()[3];
    var datos = sheet.getDataRange().getValues();
    var colaboradores = [];
    for (var i = 0; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      colaboradores.push({ nombre: String(datos[i][0]), id: String(datos[i][1] || "") });
    }
    return { status: "ok", colaboradores: colaboradores };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Leer accesos: IDs con acceso=si en hoja ACCESOS (col B=ID, col C=si/no) ──
function leerAccesos() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ACCESOS");
    if (!sheet) return { status: "ok", ids: [] };
    var datos = sheet.getDataRange().getValues();
    var ids   = [];
    for (var i = 0; i < datos.length; i++) {
      var id  = String(datos[i][1] || "").trim();
      var acc = String(datos[i][2] || "").trim().toLowerCase();
      if (id && acc === "si") ids.push(id);
    }
    return { status: "ok", ids: ids };
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

// ── Leer items por lugar (hoja con nombre del lugar) ────────
// Mapa explícito para nombres que difieren del nombre de hoja
var MAPA_LUGAR_SHEET = {
  // Bodegas → usan hoja INSUMOS
  "BODEGA INSUMOS CLINICOS":    "INSUMOS",
  "BODEGA INSUMOS NO CLINICOS": "INSUMOS",
  // Lugares clínicos → hoja propia (nombres exactos del sheet)
  "CURACIONES":                 "CURACIONES",
  "UNACESS":                    "UNACESS",
  "TOMA DE MUESTRAS":           "TOMA_MUESTRAS",
  "LASERTERAPIA":               "LASERTERAPIA",
  "BOX MEDICOS":                "BOX_MEDICOS",
  "FOTOTERAPIA":                "FOTOTERAPIA",
  "PABELLÓN":                   "PABELLON",
  "PABELLON":                   "PABELLON",
  "ÁREA TÉCNICA DERMA":         "AREA_TECNICA",
  "AREA TECNICA DERMA":         "AREA_TECNICA",
  "OFICINA SECRETARIA":         "OFICINA_ADMIN",
  "OFICINA ADMINISTRATIVA":     "OFICINA_ADMIN"
};

function leerItemsPorLugar(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var lugar  = (e.parameter.lugar  || "").trim();
    var bodega = (e.parameter.bodega || "").trim().toUpperCase();
    var sheet  = null;

    if (lugar) {
      var lugarUp    = lugar.toUpperCase();
      var nombreHoja = MAPA_LUGAR_SHEET[lugarUp] || MAPA_LUGAR_SHEET[lugar] || null;
      if (nombreHoja) sheet = ss.getSheetByName(nombreHoja);
      if (!sheet) sheet = ss.getSheetByName(lugar);
      if (!sheet) {
        var hojas = ss.getSheets();
        for (var i = 0; i < hojas.length; i++) {
          if (hojas[i].getName().toUpperCase() === lugarUp) { sheet = hojas[i]; break; }
        }
      }
    }

    if (!sheet) sheet = ss.getSheetByName(SHEET_ITEMS) || ss.getSheets()[2];

    var datos = sheet.getDataRange().getValues();
    var items = [];
    var encIni = String(datos[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;

    for (var j = encIni; j < datos.length; j++) {
      var cod  = String(datos[j][0]||"").trim();
      var desc = String(datos[j][1]||"").trim();
      var col6 = String(datos[j][5]||"").trim().toUpperCase();
      if (!cod) continue;
      if (bodega && col6 !== bodega) continue;
      items.push({ codigo: cod, descripcion: desc });
    }

    return { status: "ok", items: items, fuente: sheet.getName(), total: items.length };
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
      try { actualizarStockCuraciones(); } catch(ex) { Logger.log("Stock bodega no actualizado: " + ex); }
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

// ── Verificar acceso al panel de aprobación ─────────────────
// Lee hoja ACCESOS: col A=Nombre, col B=ID, col C=si/no
function verificarAcceso(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ACCESOS");
    if (!sheet) return { status: "error", mensaje: "Hoja ACCESOS no existe." };
    var p     = e.parameter || {};
    var idBuscado = String(p.id || "").trim();
    if (!idBuscado) return { status: "error", mensaje: "Sin ID." };
    var datos = sheet.getDataRange().getValues();
    for (var i = 0; i < datos.length; i++) {
      var fila = datos[i];
      var id   = String(fila[1] || "").trim();
      if (id === idBuscado) {
        var acceso = String(fila[2] || "").trim().toLowerCase();
        return {
          status:   "ok",
          nombre:   String(fila[0] || ""),
          id:       id,
          acceso:   acceso === "si"
        };
      }
    }
    return { status: "ok", acceso: false, mensaje: "Responsable no encontrado." };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Lugares con inventario registrado ────────────────────────
function lugaresConInventario() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DATOS);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "ok", lugares: [] };
    var datos  = sheet.getDataRange().getValues();
    var inicio = String(datos[0][0]).toUpperCase() === "LUGAR" ? 1 : 0;
    var set    = {};
    for (var i = inicio; i < datos.length; i++) {
      var lugar = String(datos[i][0] || "").trim();
      if (lugar) set[lugar] = true;
    }
    return { status: "ok", lugares: Object.keys(set) };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Registrar corrección en MOVIMIENTOS ──────────────────────
function registrarCorreccion(e) {
  // Registra un AJUSTE trimestral: diferencia entre conteo físico y stock calculado
  // cantidad positiva = sobrante (AJUSTE+), negativa = faltante (AJUSTE-)
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var p        = e.parameter || {};
    var lugar    = String(p.lugar        || "").trim();
    var cod      = String(p.codigo       || "").trim();
    var desc     = String(p.descripcion  || "").trim();
    var stockReal = parseFloat(p.stockReal  || 0);   // conteo físico real
    var stockCalc = parseFloat(p.stockCalc  || 0);   // stock calculado por sistema
    var usuario  = String(p.usuario      || "").trim();
    var usuId    = String(p.usuarioId    || "").trim();
    var motivo   = String(p.motivo       || "Ajuste trimestral").trim();

    if (!lugar || !cod) return { status: "error", mensaje: "Faltan datos." };

    var diferencia = stockReal - stockCalc; // positivo=sobra, negativo=falta
    if (diferencia === 0) return { status: "ok", mensaje: "Sin diferencia, no se registró ajuste." };

    var tipo = diferencia > 0 ? "AJUSTE+" : "AJUSTE-";

    var movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    if (!movSheet) return { status: "error", mensaje: "Sin hoja MOVIMIENTOS." };

    movSheet.appendRow([
      new Date().toLocaleString("es-CL"),
      tipo,
      motivo,       // campo N° Sol reutilizado para motivo del ajuste
      "",
      lugar, cod, desc,
      diferencia,   // positivo o negativo según corresponda
      "",
      usuario, usuId
    ]);

    try { actualizarStockCuraciones(); } catch(e) { Logger.log("Stock no actualizado: " + e); }
    return { status: "ok", diferencia: diferencia, tipo: tipo };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Stock por lugar: inventario + ingresos - egresos ─────────
function stockLugar(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var lugar = (e.parameter.lugar || "").trim();
    if (!lugar) return { status: "error", mensaje: "Falta lugar." };

    // Bodegas con semáforo por DÍAS (leen hoja INSUMOS col C=crítico, D=mínimo, E=máximo)
    var LUGARES_BODEGA = ["BODEGA INSUMOS CLINICOS", "BODEGA INSUMOS NO CLINICOS"];
    var esBodega = LUGARES_BODEGA.indexOf(lugar) > -1;

    // ── Leer umbrales en unidades ─────────────────────────────
    // Todos los lugares: col C=crítico, D=mínimo, E=máximo
    // Bodegas → hoja INSUMOS / Lugares clínicos → hoja propia del lugar
    // Los valores representan: C=stock 7d, D=stock 15d, E=stock 45d
    var mapaUmbrales = {}; // cod → { critico, minimo, maximo }
    var shUmb = esBodega ? ss.getSheetByName("INSUMOS") : ss.getSheetByName(lugar);
    if (shUmb && shUmb.getLastRow() > 1) {
      var umbData = shUmb.getDataRange().getValues();
      var umbIni  = String(umbData[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;
      for (var r = umbIni; r < umbData.length; r++) {
        var uc = String(umbData[r][0]||"").trim();
        if (!uc) continue;
        mapaUmbrales[uc] = {
          critico: parseFloat(umbData[r][2]||0) || 0,
          minimo:  parseFloat(umbData[r][3]||0) || 0,
          maximo:  parseFloat(umbData[r][4]||0) || 0
        };
      }
    }

    // ── Inventario inicial ────────────────────────────────────
    var invSheet = ss.getSheetByName(SHEET_DATOS);
    var invDatos = invSheet ? invSheet.getDataRange().getValues() : [];
    var ini      = String(invDatos[0] && invDatos[0][0] || "").toUpperCase() === "LUGAR" ? 1 : 0;
    var mapaInv  = {}; // cod → { descripcion, cantidad, ingresos, egresos }
    for (var i = ini; i < invDatos.length; i++) {
      var f = invDatos[i];
      if (String(f[0]||"").trim() !== lugar) continue;
      var cod  = String(f[1]||"").trim();
      var desc = String(f[2]||"").trim();
      var qty  = parseFloat(f[3]||0) || 0;
      if (!cod) continue;
      if (mapaInv[cod]) mapaInv[cod].cantidad += qty;
      else mapaInv[cod] = { descripcion: desc, cantidad: qty, ingresos: 0, egresos: 0 };
    }

    // ── Movimientos ───────────────────────────────────────────
    var mapaEgreso90 = {}; // ya no usado, se mantiene por compatibilidad

    var movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    if (movSheet && movSheet.getLastRow() > 1) {
      var movDatos = movSheet.getDataRange().getValues();
      var mIni     = String(movDatos[0][0]||"").toUpperCase().indexOf("FECHA") === 0 ? 1 : 0;
      for (var j = mIni; j < movDatos.length; j++) {
        var m    = movDatos[j];
        var mLug = String(m[4]||"").trim();
        var mTip = String(m[1]||"").trim().toUpperCase();
        var mCod = String(m[5]||"").trim();
        var mQty = parseFloat(m[7]||0) || 0;
        if (mLug !== lugar || !mCod) continue;
        if (!mapaInv[mCod]) mapaInv[mCod] = { descripcion: String(m[6]||mCod), cantidad: 0, ingresos: 0, egresos: 0 };
        if (mTip === "INGRESO")       mapaInv[mCod].ingresos += mQty;
        else if (mTip === "EGRESO")   mapaInv[mCod].egresos  += mQty;

        // Acumular egresos últimos 90 días para consumo promedio (solo bodegas)

      }
    }

    // ── Calcular stock, estado y días ─────────────────────────
    var items = [];
    for (var cod in mapaInv) {
      var it    = mapaInv[cod];
      var stock = it.cantidad + it.ingresos + it.egresos;
      var estado, dias = null;

      // Semáforo por unidades — igual para todos los lugares
      var u = mapaUmbrales[cod] || { critico:0, minimo:0, maximo:0 };
      if (u.critico > 0 || u.minimo > 0) {
        estado = stock <= u.critico ? "CRITICO"
               : stock <= u.minimo  ? "BAJO"
               : u.maximo > 0 && stock > u.maximo ? "SOBRESTOCK"
               : "MAXIMO";
      } else {
        estado = stock <= 0 ? "CRITICO" : "OK";
      }

      items.push({
        codigo:      cod,
        descripcion: it.descripcion,
        stock:       stock,
        estado:      estado,
        dias:        dias !== null ? Math.round(dias) : null,
        esBodega:    esBodega
      });
    }

    items.sort(function(a,b){ return a.descripcion.localeCompare(b.descripcion); });
    return { status: "ok", lugar: lugar, items: items, esBodega: esBodega };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Registrar egreso en MOVIMIENTOS ──────────────────────────
function registrarEgreso(e) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var p   = e.parameter || {};
    var lugar   = String(p.lugar       || "").trim();
    var cod     = String(p.codigo      || "").trim();
    var desc    = String(p.descripcion || "").trim();
    var qty     = parseFloat(p.cantidad || 0);
    var usuario = String(p.usuario     || "").trim();
    var usuId   = String(p.usuarioId   || "").trim();
    var nSol    = String(p.nSol        || "").trim();
    if (!lugar || !cod || !qty) return { status: "error", mensaje: "Faltan datos." };
    var movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
    if (!movSheet) movSheet = ss.insertSheet(SHEET_MOVIMIENTOS);
    if (movSheet.getLastRow() === 0)
      movSheet.appendRow(["Fecha/Hora","Tipo","N° Sol","Mes","Lugar","Código","Descripción","Cantidad","Fecha Vencimiento","Responsable","ID"]);
    movSheet.appendRow([
      new Date().toLocaleString("es-CL"), "EGRESO", nSol, "", lugar, cod, desc, -Math.abs(qty), "", usuario, usuId
    ]);
    try { actualizarStockCuraciones(); } catch(e) { Logger.log("Stock no actualizado: " + e); }
    return { status: "ok" };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Listar solicitudes por lugar ─────────────────────────────
function listarSolicitudesPorLugar(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet || sheet.getLastRow() <= 1) return { status: "ok", solicitudes: [] };
    var lugar = (e.parameter.lugar || "").trim();
    var datos = sheet.getDataRange().getValues();
    var inicio = String(datos[0][0]).indexOf("SOL-") === 0 ? 0 : 1;
    var solicitudes = [];
    for (var i = inicio; i < datos.length; i++) {
      var f = datos[i];
      if (!f[0] || String(f[0]).trim() === "") continue;
      var lugarFila = String(f[1] || "").trim();
      if (lugar && lugarFila.toUpperCase() !== lugar.toUpperCase()) continue;
      var estado = String(f[8] || "PENDIENTE").trim().toUpperCase();
      if (estado === "RECHAZADO") continue;
      var fecha = f[7] instanceof Date ? f[7].toLocaleString("es-CL") : String(f[7] || "");
      solicitudes.push({
        fila: i + 1,
        id:           String(f[0] || ""),
        lugar:        lugarFila,
        codigo:       String(f[2] || ""),
        item:         String(f[3] || ""),
        cantidad:     String(f[4] || ""),
        responsable:  String(f[5] || ""),
        idResp:       String(f[6] || ""),
        fecha:        fecha,
        estado:       estado
      });
    }
    return { status: "ok", solicitudes: solicitudes };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Recepcionar solicitud → actualiza SOLICITUDES + escribe en MOVIMIENTOS ──
function recepcionarSolicitud(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) throw new Error("Hoja SOLICITUDES no existe.");
    var p    = e.parameter || {};
    var fila = parseInt(p.fila || "0");
    var est  = (p.estado      || "").toUpperCase();
    var cant = p.cantRecibida || "";
    var venc = p.fechaVenc    || "";
    if (!fila || !est) throw new Error("Faltan parámetros.");

    if (est !== "APROBADO") {
      // Solo actualizar estado (RECHAZADO no genera movimientos)
      sheet.getRange(fila, 9).setValue("RECHAZADO");
      sheet.getRange(fila, 10).setValue(new Date().toLocaleString("es-CL"));
      return { status: "ok", fila: fila };
    }
    // APROBADO → registrar INGRESO en lugar + EGRESO en bodega origen
    // (estado se actualiza DESPUÉS para no marcar como recepcionado si falla)
    if (est === "APROBADO") {
      var lugar = String(p.lugar       || "").trim();
      var cod   = String(p.codigo      || "").trim();
      var desc  = String(p.descripcion || "").trim();
      var qty   = parseFloat(cant) || 0;
      var nSol  = String(p.idSolicitud || "").trim();
      var bodegaOrigen = String(p.bodegaOrigen || "").trim();

      var movSheet = ss.getSheetByName(SHEET_MOVIMIENTOS);
      if (!movSheet) movSheet = ss.insertSheet(SHEET_MOVIMIENTOS);
      if (movSheet.getLastRow() === 0)
        movSheet.appendRow(["Fecha/Hora","Tipo","N° Sol","Mes","Lugar","Código","Descripción","Cantidad","Fecha Vencimiento","Responsable","ID"]);

      var ahora = new Date().toLocaleString("es-CL");

      // INGRESO en el lugar que recibe
      movSheet.appendRow([ahora, "INGRESO", nSol, "", lugar, cod, desc, Math.abs(qty), venc, "", ""]);

      // EGRESO en la bodega que despacha
      // Si no viene bodegaOrigen desde la app, leer desde SOLICITUDES col 11 o buscar en INSUMOS col F
      if (!bodegaOrigen) {
        var solData = sheet.getDataRange().getValues();
        if (fila > 0 && fila < solData.length + 1) {
          var solRow = solData[fila - 1];
          if (solRow[10]) bodegaOrigen = String(solRow[10]||"").trim();
        }
      }
      // Como fallback final, buscar en INSUMOS col F
      if (!bodegaOrigen) {
        var shIns = ss.getSheetByName("INSUMOS");
        if (shIns) {
          var insData = shIns.getDataRange().getValues();
          for (var bi = 1; bi < insData.length; bi++) {
            if (String(insData[bi][0]||"").trim() === cod) {
              var tipoBodega = String(insData[bi][5]||"").trim().toUpperCase();
              bodegaOrigen = tipoBodega === "CLINICOS" ? "BODEGA INSUMOS CLINICOS" : "BODEGA INSUMOS NO CLINICOS";
              break;
            }
          }
        }
      }
      if (bodegaOrigen) {
        movSheet.appendRow([ahora, "EGRESO", nSol, "", bodegaOrigen, cod, desc, -Math.abs(qty), venc, "", ""]);
      }
    }
    // Actualizar estado en SOLICITUDES (aquí, después de escribir movimientos)
    sheet.getRange(fila, 9).setValue(est === "APROBADO" ? "RECEPCIONADO" : "RECHAZADO");
    sheet.getRange(fila, 10).setValue(new Date().toLocaleString("es-CL"));
    try { actualizarStockCuraciones(); } catch(ex) { Logger.log("Stock no actualizado: " + ex); }
    return { status: "ok", fila: fila };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Login ────────────────────────────────────────────────────
function login(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("COLABORADORES");
    if (!sheet) throw new Error("Hoja COLABORADORES no existe.");
    var p   = e.parameter || {};
    var rut = String(p.rut || "").trim();
    var pwd = String(p.pwd || "").trim();
    if (!rut || !pwd) return { status: "error", mensaje: "Faltan credenciales." };
    var datos = sheet.getDataRange().getValues();
    for (var i = 0; i < datos.length; i++) {
      var fRut = String(datos[i][1] || "").trim();
      var fPwd = String(datos[i][2] || "").trim();
      if (fRut === rut && fPwd === pwd) {
        // Leer accesos desde hoja ACCESOS
        var accesos = leerAccesosUsuario(ss, fRut);
        return { status: "ok", nombre: String(datos[i][0] || ""), id: fRut, accesos: accesos };
      }
    }
    return { status: "error", mensaje: "Credenciales incorrectas." };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Leer accesos de un usuario desde hoja ACCESOS ────────────
function leerAccesosUsuario(ss, id) {
  var ADMIN_ID = "15579172-1";
  if (id === ADMIN_ID) return [1,1,1,1,1,1,1,1]; // Admin siempre con acceso total
  var sheet = ss.getSheetByName("ACCESOS");
  if (!sheet || sheet.getLastRow() <= 1) return [1,1,1,1,1,1,1,1];
  var datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1] || "").trim() === id) {
      var acc = [];
      for (var j = 2; j < 10; j++) acc.push(String(datos[i][j] || "").trim().toLowerCase() === "si" ? 1 : 0);
      return acc;
    }
  }
  return [1,1,1,1,1,1,1,1];
}

// ── Obtener todos los accesos para panel admin ────────────────
function obtenerAccesos(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var colabs = ss.getSheetByName("COLABORADORES");
    var accs   = ss.getSheetByName("ACCESOS");
    if (!colabs) throw new Error("Hoja COLABORADORES no existe.");
    var colDatos = colabs.getDataRange().getValues();
    // Construir mapa de accesos por ID
    var mapaAccesos = {};
    if (accs && accs.getLastRow() > 1) {
      var accDatos = accs.getDataRange().getValues();
      for (var i = 1; i < accDatos.length; i++) {
        var id  = String(accDatos[i][1] || "").trim();
        var acc = [];
        for (var j = 2; j < 10; j++) acc.push(String(accDatos[i][j] || "").trim().toLowerCase() === "si" ? 1 : 0);
        mapaAccesos[id] = acc;
      }
    }
    var colaboradores = [];
    for (var i = 0; i < colDatos.length; i++) {
      if (!colDatos[i][0] || !colDatos[i][1]) continue;
      var id = String(colDatos[i][1]).trim();
      colaboradores.push({
        nombre:  String(colDatos[i][0]).trim(),
        id:      id,
        accesos: mapaAccesos[id] || [1,1,1,1,1,1,1,1]
      });
    }
    return { status: "ok", colaboradores: colaboradores };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}

// ── Guardar accesos desde panel admin ────────────────────────
function guardarAccesos(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ACCESOS");
    if (!sheet) sheet = ss.insertSheet("ACCESOS");
    var p     = e.parameter || {};
    var datos = JSON.parse(decodeURIComponent(p.datos || "[]"));
    // Limpiar y reescribir desde fila 2
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
    // Encabezados si faltan
    if (lastRow < 1) sheet.getRange(1,1,1,8).setValues([["Responsable","ID","Recep. por Lugar","Stock por Lugar","Captura Inventario","Solicitud Insumos","Recep. Bodega","Aprobación Solic.","Ajuste Stock","Reportes Gestión"]]);
    // También cargar nombres desde COLABORADORES para la columna A
    var colabs  = ss.getSheetByName("COLABORADORES");
    var mapaNom = {};
    if (colabs) {
      var cd = colabs.getDataRange().getValues();
      for (var i = 0; i < cd.length; i++) mapaNom[String(cd[i][1]).trim()] = String(cd[i][0]).trim();
    }
    var filas = [];
    for (var i = 0; i < datos.length; i++) {
      var d   = datos[i];
      var acc = d.accesos || [0,0,0,0,0,0,0,0];
      filas.push([
        mapaNom[d.id] || d.id, d.id,
        acc[0] ? "si" : "no", acc[1] ? "si" : "no",
        acc[2] ? "si" : "no", acc[3] ? "si" : "no",
        acc[4] ? "si" : "no", acc[5] ? "si" : "no",
        acc[6] ? "si" : "no", acc[7] ? "si" : "no"
      ]);
    }
    if (filas.length > 0) sheet.getRange(2, 1, filas.length, 10).setValues(filas);
    return { status: "ok" };
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

// ============================================================
//  NOTIFICACIONES POR CORREO — Bodega Dermatología HDS
// ============================================================

function obtenerDestinatarios() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ACCESOS");
    if (!sheet) return ["cmendez@hsalvador.cl"];
    var datos = sheet.getDataRange().getValues();
    var destinos = [];
    for (var i = 0; i < datos.length; i++) {
      var correo = String(datos[i][9] || "").trim();
      var activo = String(datos[i][10] || "").trim().toLowerCase();
      if (correo && correo.indexOf("@") > -1 && activo === "si") {
        destinos.push(correo);
      }
    }
    return destinos.length > 0 ? destinos : ["cmendez@hsalvador.cl"];
  } catch(err) {
    return ["cmendez@hsalvador.cl"];
  }
}

function notificarResumenSolicitud(e) {
  try {
    var p          = e.parameter || {};
    var idSolicitud = String(p.idSolicitud  || "").trim();
    var lugar       = String(p.lugar        || "").trim();
    var responsable = String(p.responsable  || "").trim();
    var idResp      = String(p.idResp       || "").trim();
    if (!idSolicitud) return { status: "error", mensaje: "Sin idSolicitud." };

    // Leer todos los ítems de esta solicitud desde SOLICITUDES
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SOLICITUDES");
    if (!sheet) return { status: "error", mensaje: "Sin hoja SOLICITUDES." };
    var datos = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (String(fila[0]||"").trim() === idSolicitud) {
        items.push({
          codigo:  String(fila[2]||"").trim(),
          desc:    String(fila[3]||"").trim(),
          cantidad: String(fila[4]||"").trim()
        });
      }
    }
    if (!items.length) return { status: "ok", mensaje: "Sin ítems para notificar." };

    var destinos = obtenerDestinatarios();
    var fecha    = new Date().toLocaleString("es-CL");
    var asunto   = "🛒 Solicitud " + idSolicitud + " — " + lugar + " — " + fecha;

    // Construir tabla HTML de ítems
    var filasHtml = "";
    for (var j = 0; j < items.length; j++) {
      var bg = j % 2 === 0 ? "#ffffff" : "#f8fafc";
      filasHtml +=
        "<tr style='background:" + bg + "'>" +
          "<td style='padding:7px 10px;font-family:monospace;font-size:12px;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].codigo + "</td>" +
          "<td style='padding:7px 10px;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].desc + "</td>" +
          "<td style='padding:7px 10px;text-align:center;font-weight:700;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].cantidad + "</td>" +
        "</tr>";
    }

    var cuerpo =
      "<div style='font-family:Arial,sans-serif;max-width:600px;'>" +
      "<div style='background:#1B5FA5;padding:16px 20px;border-radius:8px 8px 0 0;'>" +
        "<h2 style='color:#fff;margin:0;font-size:18px;'>🛒 Nueva solicitud de insumos</h2>" +
        "<p style='color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px;'>Dermatología · Hospital del Salvador</p>" +
      "</div>" +
      "<div style='border:1px solid #d1dce8;border-top:none;padding:16px 20px;background:#fff;'>" +
        "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;font-size:13px;'>" +
          "<tr><td style='padding:5px 8px;color:#64748b;width:140px;'>N° Solicitud:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;font-weight:600;color:#185FA5;'>" + idSolicitud + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Responsable:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#1e293b;'>" + responsable + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>RUT:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;color:#1e293b;'>" + idResp + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Lugar:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#185FA5;'>" + lugar + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>Fecha:</td>" +
              "<td style='padding:5px 8px;color:#1e293b;'>" + fecha + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Total ítems:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#1e293b;'>" + items.length + "</td></tr>" +
        "</table>" +
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
          "<tr style='background:#EBF3FC;'>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Código</th>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Descripción</th>" +
            "<th style='padding:7px 10px;text-align:center;color:#185FA5;'>Cant.</th>" +
          "</tr>" +
          filasHtml +
        "</table>" +
        "<p style='font-size:12px;color:#64748b;margin-top:14px;'>Esta solicitud está en estado <strong>PENDIENTE</strong> de aprobación.</p>" +
      "</div>" +
      "<div style='background:#f8fafc;border:1px solid #d1dce8;border-top:none;padding:10px 20px;border-radius:0 0 8px 8px;text-align:center;font-size:11px;color:#94a3b8;'>" +
        "Sistema de Gestión de Insumos · Dermatología HDS" +
      "</div></div>";

    MailApp.sendEmail({ to: destinos.join(","), subject: asunto, htmlBody: cuerpo });
    Logger.log("✓ Resumen solicitud enviado: " + idSolicitud + " (" + items.length + " ítems)");
    return { status: "ok" };
  } catch(err) {
    Logger.log("Error notificarResumenSolicitud: " + err.toString());
    return { status: "error", mensaje: err.toString() };
  }
}

function notificarResumenSolicitud(e) {
  try {
    var p           = e.parameter || {};
    var idSolicitud = String(p.idSolicitud || "").trim();
    var lugar       = String(p.lugar       || "").trim();
    var responsable = String(p.responsable || "").trim();
    var idResp      = String(p.idResp      || "").trim();
    if (!idSolicitud) return { status: "error", mensaje: "Sin idSolicitud." };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SOLICITUDES");
    if (!sheet) return { status: "error", mensaje: "Sin hoja SOLICITUDES." };
    var datos = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (String(fila[0]||"").trim() === idSolicitud) {
        items.push({ codigo: String(fila[2]||""), desc: String(fila[3]||""), cantidad: String(fila[4]||"") });
      }
    }
    if (!items.length) return { status: "ok", mensaje: "Sin items." };

    var destinos = obtenerDestinatarios();
    var fecha    = new Date().toLocaleString("es-CL");
    var asunto   = "🛒 Solicitud " + idSolicitud + " — " + lugar + " — " + fecha;
    var filasHtml = "";
    for (var j = 0; j < items.length; j++) {
      var bg = j % 2 === 0 ? "#ffffff" : "#f8fafc";
      filasHtml += "<tr style='background:" + bg + "'>" +
        "<td style='padding:7px 10px;font-family:monospace;font-size:12px;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].codigo + "</td>" +
        "<td style='padding:7px 10px;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].desc + "</td>" +
        "<td style='padding:7px 10px;text-align:center;font-weight:700;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + items[j].cantidad + "</td>" +
        "</tr>";
    }
    var cuerpo =
      "<div style='font-family:Arial,sans-serif;max-width:600px;'>" +
      "<div style='background:#1B5FA5;padding:16px 20px;border-radius:8px 8px 0 0;'>" +
        "<h2 style='color:#fff;margin:0;font-size:18px;'>🛒 Nueva solicitud de insumos</h2>" +
        "<p style='color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px;'>Dermatología · Hospital del Salvador</p>" +
      "</div>" +
      "<div style='border:1px solid #d1dce8;border-top:none;padding:16px 20px;background:#fff;'>" +
        "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;font-size:13px;'>" +
          "<tr><td style='padding:5px 8px;color:#64748b;width:140px;'>N° Solicitud:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;font-weight:600;color:#185FA5;'>" + idSolicitud + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Responsable:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#1e293b;'>" + responsable + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>RUT:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;color:#1e293b;'>" + idResp + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Lugar:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#185FA5;'>" + lugar + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>Fecha:</td>" +
              "<td style='padding:5px 8px;color:#1e293b;'>" + fecha + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>Total ítems:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#1e293b;'>" + items.length + "</td></tr>" +
        "</table>" +
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
          "<tr style='background:#EBF3FC;'>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Código</th>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Descripción</th>" +
            "<th style='padding:7px 10px;text-align:center;color:#185FA5;'>Cant.</th>" +
          "</tr>" + filasHtml +
        "</table>" +
        "<p style='font-size:12px;color:#64748b;margin-top:14px;'>Esta solicitud está en estado <strong>PENDIENTE</strong> de aprobación.</p>" +
      "</div>" +
      "<div style='background:#f8fafc;border:1px solid #d1dce8;border-top:none;padding:10px 20px;border-radius:0 0 8px 8px;text-align:center;font-size:11px;color:#94a3b8;'>" +
        "Sistema de Gestión de Insumos · Dermatología HDS" +
      "</div></div>";
    MailApp.sendEmail({ to: destinos.join(","), subject: asunto, htmlBody: cuerpo });
    Logger.log("✓ Resumen solicitud enviado: " + idSolicitud + " (" + items.length + " items)");
    return { status: "ok" };
  } catch(err) {
    Logger.log("Error notificarResumenSolicitud: " + err.toString());
    return { status: "error", mensaje: err.toString() };
  }
}

function notificarSolicitud(datos) {
  try {
    var destinos = obtenerDestinatarios();
    var fecha    = new Date().toLocaleString("es-CL");
    var asunto   = "🛒 Nueva solicitud de insumos — " + (datos.lugar || "") + " — " + fecha;
    var cuerpo =
      "<div style='font-family:Arial,sans-serif;max-width:600px;'>" +
      "<div style='background:#1B5FA5;padding:16px 20px;border-radius:8px 8px 0 0;'>" +
        "<h2 style='color:#fff;margin:0;font-size:18px;'>🛒 Nueva solicitud de insumos</h2>" +
        "<p style='color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px;'>Dermatología · Hospital del Salvador</p>" +
      "</div>" +
      "<div style='border:1px solid #d1dce8;border-top:none;padding:16px 20px;background:#fff;'>" +
        "<table style='width:100%;border-collapse:collapse;margin-bottom:14px;font-size:13px;'>" +
          "<tr><td style='padding:5px 8px;color:#64748b;width:140px;'>Responsable:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#1e293b;'>" + (datos.responsable || "—") + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>RUT:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;color:#1e293b;'>" + (datos.idResp || "—") + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>Lugar:</td>" +
              "<td style='padding:5px 8px;font-weight:600;color:#185FA5;'>" + (datos.lugar || "—") + "</td></tr>" +
          "<tr style='background:#f8fafc;'><td style='padding:5px 8px;color:#64748b;'>N° Solicitud:</td>" +
              "<td style='padding:5px 8px;font-family:monospace;color:#1e293b;'>" + (datos.idSolicitud || "—") + "</td></tr>" +
          "<tr><td style='padding:5px 8px;color:#64748b;'>Fecha:</td>" +
              "<td style='padding:5px 8px;color:#1e293b;'>" + fecha + "</td></tr>" +
        "</table>" +
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
          "<tr style='background:#EBF3FC;'>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Código</th>" +
            "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Descripción</th>" +
            "<th style='padding:7px 10px;text-align:center;color:#185FA5;'>Cantidad</th>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding:7px 10px;font-family:monospace;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + (datos.codigo || "—") + "</td>" +
            "<td style='padding:7px 10px;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + (datos.descripcion || datos.item || "—") + "</td>" +
            "<td style='padding:7px 10px;text-align:center;font-weight:700;color:#1e293b;border-bottom:1px solid #f0f4f8;'>" + (datos.cantidad || "—") + "</td>" +
          "</tr>" +
        "</table>" +
        "<p style='font-size:12px;color:#64748b;margin-top:14px;'>Esta solicitud está en estado <strong>PENDIENTE</strong> de aprobación.</p>" +
      "</div>" +
      "<div style='background:#f8fafc;border:1px solid #d1dce8;border-top:none;padding:10px 20px;border-radius:0 0 8px 8px;text-align:center;font-size:11px;color:#94a3b8;'>" +
        "Sistema de Gestión de Insumos · Dermatología HDS" +
      "</div></div>";
    MailApp.sendEmail({ to: destinos.join(","), subject: asunto, htmlBody: cuerpo });
  } catch(err) {
    Logger.log("Error notificarSolicitud: " + err.toString());
  }
}

function agregarTriggerVencimientos() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "alertaVencimientos") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("alertaVencimientos")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log("✓ Trigger semanal configurado: alertaVencimientos — Lunes 8AM");
}

function alertaVencimientos() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("INVENTARIO");
    if (!sheet || sheet.getLastRow() <= 1) return;
    var hoy    = new Date();
    var en1mes = new Date(hoy); en1mes.setMonth(en1mes.getMonth() + 1);
    var en3mes = new Date(hoy); en3mes.setMonth(en3mes.getMonth() + 3);
    var datos  = sheet.getDataRange().getValues();
    var alerta1 = [], alerta3 = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      var venc = f[4];
      if (!venc || String(venc).toLowerCase() === "no aplica" || String(venc).trim() === "") continue;
      var fechaVenc = null;
      if (venc instanceof Date) {
        fechaVenc = venc;
      } else {
        var partes = String(venc).trim().split("/");
        if (partes.length === 3) fechaVenc = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
      }
      if (!fechaVenc || isNaN(fechaVenc.getTime()) || fechaVenc < hoy) continue;
      var fila = { lugar: String(f[0]||""), cod: String(f[1]||""), desc: String(f[2]||""), cant: String(f[3]||""), venc: fechaVenc.toLocaleDateString("es-CL") };
      if (fechaVenc <= en1mes) alerta1.push(fila);
      else if (fechaVenc <= en3mes) alerta3.push(fila);
    }
    if (alerta1.length === 0 && alerta3.length === 0) { Logger.log("Sin alertas esta semana."); return; }
    var destinos = obtenerDestinatarios();
    var fecha    = hoy.toLocaleDateString("es-CL");
    var asunto   = "⚠ Alerta vencimientos — Bodega Dermatología — " + fecha;
    function filaHtml(f, color, emoji) {
      return "<tr><td style='padding:6px 10px;border-bottom:1px solid #f0f4f8;'>" + f.lugar + "</td>" +
        "<td style='padding:6px 10px;font-family:monospace;font-size:12px;border-bottom:1px solid #f0f4f8;'>" + f.cod + "</td>" +
        "<td style='padding:6px 10px;border-bottom:1px solid #f0f4f8;'>" + f.desc + "</td>" +
        "<td style='padding:6px 10px;text-align:center;border-bottom:1px solid #f0f4f8;'>" + f.cant + "</td>" +
        "<td style='padding:6px 10px;text-align:center;font-weight:700;color:" + color + ";border-bottom:1px solid #f0f4f8;'>" + emoji + " " + f.venc + "</td></tr>";
    }
    var tabla = "";
    if (alerta1.length > 0) {
      tabla += "<tr style='background:#fee2e2;'><td colspan='5' style='padding:6px 10px;font-weight:700;color:#dc2626;'>🔴 Vencen en menos de 1 mes (" + alerta1.length + " ítems)</td></tr>";
      for (var a = 0; a < alerta1.length; a++) tabla += filaHtml(alerta1[a], "#dc2626", "🔴");
    }
    if (alerta3.length > 0) {
      tabla += "<tr style='background:#fef3c7;'><td colspan='5' style='padding:6px 10px;font-weight:700;color:#92400e;'>🟡 Vencen entre 1 y 3 meses (" + alerta3.length + " ítems)</td></tr>";
      for (var b = 0; b < alerta3.length; b++) tabla += filaHtml(alerta3[b], "#d97706", "🟡");
    }
    var cuerpo =
      "<div style='font-family:Arial,sans-serif;max-width:700px;'>" +
      "<div style='background:#92400e;padding:16px 20px;border-radius:8px 8px 0 0;'>" +
        "<h2 style='color:#fff;margin:0;font-size:18px;'>⚠ Alerta de vencimientos próximos</h2>" +
        "<p style='color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px;'>Bodega Dermatología · Hospital del Salvador · " + fecha + "</p>" +
      "</div>" +
      "<div style='border:1px solid #d1dce8;border-top:none;padding:16px 20px;background:#fff;'>" +
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
          "<tr style='background:#EBF3FC;'><th style='padding:7px 10px;text-align:left;color:#185FA5;'>Lugar</th>" +
          "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Código</th>" +
          "<th style='padding:7px 10px;text-align:left;color:#185FA5;'>Descripción</th>" +
          "<th style='padding:7px 10px;text-align:center;color:#185FA5;'>Cant.</th>" +
          "<th style='padding:7px 10px;text-align:center;color:#185FA5;'>Vencimiento</th></tr>" +
          tabla +
        "</table>" +
        "<p style='font-size:12px;color:#64748b;margin-top:14px;'>🔴 menos de 1 mes &nbsp;|&nbsp; 🟡 entre 1 y 3 meses</p>" +
      "</div>" +
      "<div style='background:#f8fafc;border:1px solid #d1dce8;border-top:none;padding:10px 20px;border-radius:0 0 8px 8px;text-align:center;font-size:11px;color:#94a3b8;'>" +
        "Sistema de Gestión de Insumos · Dermatología HDS · Reporte automático semanal" +
      "</div></div>";
    MailApp.sendEmail({ to: destinos.join(","), subject: asunto, htmlBody: cuerpo });
    Logger.log("✓ Alerta enviada a: " + destinos.join(", "));
  } catch(err) {
    Logger.log("Error alertaVencimientos: " + err.toString());
  }
}

// ============================================================
// ============================================================
//  HOJAS RESUMEN DE STOCK — Bodega Dermatología HDS
//  Función genérica actualizarStockLugar(nombre)
//  Crea/actualiza STOCK_<LUGAR> para cualquier lugar clínico
// ============================================================

// Colores de encabezado por lugar
var COLOR_LUGAR = {
  "CURACIONES":    "#185FA5",
  "PABELLON":      "#7c3aed",
  "UNACESS":       "#0F6E56",
  "LASERTERAPIA":  "#b45309",
  "TOMA_MUESTRAS": "#0369a1",
  "BOX_MEDICOS":   "#be185d",
  "FOTOTERAPIA":   "#065f46",
  "AREA_TECNICA":  "#1e3a5f",
  "OFICINA_ADMIN": "#374151"
};

// Lista de todos los lugares clínicos con hoja catálogo propia
var LUGARES_STOCK = [
  "CURACIONES", "PABELLON", "UNACESS", "LASERTERAPIA",
  "TOMA_MUESTRAS", "BOX_MEDICOS", "FOTOTERAPIA", "AREA_TECNICA", "OFICINA_ADMIN"
];

// ── Función genérica ─────────────────────────────────────────
function actualizarStockLugar(nombreLugar) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var lugarUP  = nombreLugar.toUpperCase();
  var shLugar  = ss.getSheetByName(nombreLugar);
  var shInv    = ss.getSheetByName("INVENTARIO");
  var shMov    = ss.getSheetByName("MOVIMIENTOS");
  var stockNom = "STOCK_" + lugarUP;
  var shStock  = ss.getSheetByName(stockNom);

  // Crear hoja STOCK si no existe
  if (!shStock) {
    shStock = ss.insertSheet(stockNom);
    Logger.log("✓ Hoja creada: " + stockNom);
  }

  // 1 — Catálogo del lugar (A=código, B=desc, C=crítico, D=reposición, E=máximo)
  // Los códigos SAP son numéricos; cualquier fila cuya col A no empiece por dígito
  // es encabezado o fila vacía → se ignora automáticamente sin depender de iniLug.
  var items = [];
  if (shLugar && shLugar.getLastRow() > 0) {
    var itemsData = shLugar.getDataRange().getValues();
    for (var i = 0; i < itemsData.length; i++) {
      var cod  = String(itemsData[i][0] || "").trim();
      var desc = String(itemsData[i][1] || "").trim();
      if (!cod || !/^\d/.test(cod)) continue; // omite encabezados y filas vacías
      var critico    = (itemsData[i][2] !== "" && !isNaN(itemsData[i][2])) ? Number(itemsData[i][2]) : "";
      var reposicion = (itemsData[i][3] !== "" && !isNaN(itemsData[i][3])) ? Number(itemsData[i][3]) : "";
      var maximo     = (itemsData[i][4] !== "" && !isNaN(itemsData[i][4])) ? Number(itemsData[i][4]) : "";
      items.push({ codigo: cod, desc: desc || cod, critico: critico, reposicion: reposicion, maximo: maximo });
    }
  }

  // Normaliza texto: quita tildes, reemplaza espacios por _, mayúsculas
  // Permite comparar "PABELLÓN" con "PABELLON" y "ÁREA TÉCNICA DERMA" con "AREA_TECNICA"
  function normalizarLugar(s) {
    return s.toUpperCase()
      .replace(/[ÁÀÂÄ]/g,"A").replace(/[ÉÈÊË]/g,"E")
      .replace(/[ÍÌÎÏ]/g,"I").replace(/[ÓÒÔÖ]/g,"O")
      .replace(/[ÚÙÛÜ]/g,"U").replace(/Ñ/g,"N")
      .replace(/\s+/g,"_");
  }
  var lugarNorm = normalizarLugar(lugarUP);

  // 2 — Inventario inicial filtrado por lugar (comparación normalizada)
  var invMap = {};
  if (shInv && shInv.getLastRow() > 1) {
    var invData = shInv.getDataRange().getValues();
    for (var r = 1; r < invData.length; r++) {
      var lugRaw = String(invData[r][0] || "").trim();
      var lug    = normalizarLugar(lugRaw);
      var cod    = String(invData[r][1] || "").trim();
      var desc   = String(invData[r][2] || "").trim();
      var cant   = Number(invData[r][3]) || 0;
      // Coincidencia exacta normalizada O si el nombre del lugar contiene lugarNorm
      var coincide = lug === lugarNorm ||
                     lug.indexOf(lugarNorm) === 0 ||
                     lugarNorm.indexOf(lug) === 0;
      if (coincide && cod) {
        if (!invMap[cod]) invMap[cod] = { desc: desc, cantidad: 0 };
        invMap[cod].cantidad += cant;
      }
    }
  }

  // 3 — Movimientos filtrados por lugar (comparación normalizada)
  var ingMap = {}, egrMap = {}, descMov = {};
  if (shMov && shMov.getLastRow() > 1) {
    var movData = shMov.getDataRange().getValues();
    var mIni = String(movData[0][0] || "").toUpperCase().indexOf("FECHA") === 0 ? 1 : 0;
    for (var m = mIni; m < movData.length; m++) {
      var tipo   = String(movData[m][1] || "").trim().toUpperCase();
      var lugRaw = String(movData[m][4] || "").trim();
      var lug    = normalizarLugar(lugRaw);
      var cod    = String(movData[m][5] || "").trim();
      var desc   = String(movData[m][6] || "").trim();
      var cant   = Number(movData[m][7]) || 0;
      var coincide = lug === lugarNorm ||
                     lug.indexOf(lugarNorm) === 0 ||
                     lugarNorm.indexOf(lug) === 0;
      if (!coincide || !cod) continue;
      if (desc) descMov[cod] = desc;
      if (tipo === "INGRESO")     ingMap[cod] = (ingMap[cod] || 0) + cant;
      else if (tipo === "EGRESO") egrMap[cod] = (egrMap[cod] || 0) + cant;
    }
  }

  // 4 — Añadir códigos con movimiento que no estén en catálogo
  var codsVistos = {};
  items.forEach(function(it) { codsVistos[it.codigo] = true; });
  Object.keys(ingMap).concat(Object.keys(egrMap)).forEach(function(cod) {
    if (!codsVistos[cod]) {
      codsVistos[cod] = true;
      var d = descMov[cod] || (invMap[cod] ? invMap[cod].desc : cod);
      items.push({ codigo: cod, desc: d, critico: "", reposicion: "", maximo: "" });
    }
  });

  // 5 — Construir filas
  var ahora = new Date().toLocaleString("es-CL");
  var encabezado = [
    "Código", "Descripción", "Inv. Inicial", "+Recibido", "−Egreso",
    "Stock Actual", "Stock Crítico", "Stock Reposición", "Stock Máximo", "Estado", "Actualizado"
  ];
  var filas = [encabezado];

  items.forEach(function(it) {
    var ini    = invMap[it.codigo] ? invMap[it.codigo].cantidad : 0;
    var ing    = ingMap[it.codigo] || 0;
    var egr    = egrMap[it.codigo] || 0;
    var actual = ini + ing + egr;
    var estado = "";
    if (it.critico !== "") {
      if (actual <= it.critico)              estado = "🔴 CRÍTICO";
      else if (actual <= it.reposicion)      estado = "🟡 REPONER";
      else if (it.maximo && actual > it.maximo) estado = "🔵 SOBRE";
      else                                   estado = "🟢 OK";
    } else {
      estado = actual <= 0 ? "🔴 CRÍTICO" : "🟢 OK";
    }
    filas.push([it.codigo, it.desc, ini, ing, egr, actual,
                it.critico, it.reposicion, it.maximo, estado, ahora]);
  });

  // 6 — Escribir y formatear
  shStock.clearContents();
  shStock.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  var color = COLOR_LUGAR[lugarUP] || "#185FA5";
  shStock.getRange(1, 1, 1, encabezado.length)
    .setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
  if (filas.length > 1) {
    shStock.getRange(2, 3, filas.length - 1, 4).setNumberFormat("0");
    shStock.autoResizeColumns(1, encabezado.length);
  }
  Logger.log("✓ " + stockNom + " actualizado: " + (filas.length - 1) + " ítems");
  return filas.length - 1;
}

// ── Alias retrocompatible ────────────────────────────────────
function actualizarStockCuraciones() { actualizarStockLugar("CURACIONES"); }

function archivarMovimientosAnuales() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();

  // Pedir año a archivar (por defecto el año anterior)
  var anioActual  = new Date().getFullYear();
  var anioArch    = anioActual - 1;

  var resp = ui.prompt(
    "📦 Archivar MOVIMIENTOS",
    "¿Qué año deseas archivar?\n(Se moverán los registros de ese año a la hoja MOVIMIENTOS_" + anioArch + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var inputAnio = parseInt(resp.getResponseText().trim());
  if (!isNaN(inputAnio) && inputAnio > 2000 && inputAnio < anioActual) {
    anioArch = inputAnio;
  } else if (!isNaN(inputAnio)) {
    ui.alert("⚠️ Solo puedes archivar años anteriores al año en curso (" + anioActual + ").");
    return;
  }

  var nombreArchivo = "MOVIMIENTOS_" + anioArch;

  // Verificar si ya existe la hoja de archivo
  if (ss.getSheetByName(nombreArchivo)) {
    var conf = ui.alert(
      "⚠️ Ya existe la hoja " + nombreArchivo,
      "¿Deseas reemplazarla con los datos actuales?",
      ui.ButtonSet.YES_NO
    );
    if (conf !== ui.Button.YES) return;
    ss.deleteSheet(ss.getSheetByName(nombreArchivo));
  }

  var shMov = ss.getSheetByName("MOVIMIENTOS");
  if (!shMov) { ui.alert("No se encontró la hoja MOVIMIENTOS."); return; }

  var datos = shMov.getDataRange().getValues();
  if (datos.length <= 1) { ui.alert("No hay datos en MOVIMIENTOS para archivar."); return; }

  var encabezado = datos[0];
  var filasMover = [];
  var filasQuedan = [encabezado];

  // Columna 0 = Fecha/Hora — detectar registros del año a archivar
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var celda = fila[0];
    var anioFila = null;

    if (celda instanceof Date) {
      anioFila = celda.getFullYear();
    } else if (typeof celda === "string" && celda.length >= 4) {
      // Formato "DD/MM/YYYY HH:MM:SS" → extraer año
      var partes = celda.split(/[\s\/\-]/);
      for (var p = 0; p < partes.length; p++) {
        var n = parseInt(partes[p]);
        if (n > 2000 && n < 2100) { anioFila = n; break; }
      }
    }

    if (anioFila === anioArch) {
      filasMover.push(fila);
    } else {
      filasQuedan.push(fila);
    }
  }

  if (filasMover.length === 0) {
    ui.alert("No se encontraron registros del año " + anioArch + " en MOVIMIENTOS.");
    return;
  }

  // Crear hoja de archivo y copiar datos
  var shArch = ss.insertSheet(nombreArchivo);
  var todasArch = [encabezado].concat(filasMover);
  shArch.getRange(1, 1, todasArch.length, encabezado.length).setValues(todasArch);

  // Formato encabezado en hoja de archivo
  shArch.getRange(1, 1, 1, encabezado.length)
    .setBackground("#0f2d52")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  // Nota de archivo
  shArch.getRange(todasArch.length + 2, 1)
    .setValue("📦 Archivado el " + new Date().toLocaleDateString("es-CL") + " · SIIDER · Dermatología HDS");

  // Reescribir MOVIMIENTOS solo con encabezado + año actual
  shMov.clearContents();
  shMov.getRange(1, 1, filasQuedan.length, encabezado.length).setValues(filasQuedan);
  shMov.getRange(1, 1, 1, encabezado.length)
    .setBackground("#0f2d52")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  ui.alert(
    "✅ Archivado exitoso",
    filasMover.length + " registros del año " + anioArch + " movidos a la hoja \"" + nombreArchivo + "\".\n" +
    filasQuedan.length - 1 + " registros quedan en MOVIMIENTOS (año " + anioActual + ").",
    ui.ButtonSet.OK
  );

  Logger.log("✓ Archivado " + filasMover.length + " filas → " + nombreArchivo);
}

// ── Clasificar insumos: escribe CLINICOS/NO CLINICOS en col F ──
// Busca cada código de INSUMOS en las hojas de lugares clínicos
// Si aparece en alguna → CLINICOS, si no → NO CLINICOS
function clasificarInsumos() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  var shIns = ss.getSheetByName("INSUMOS");
  if (!shIns) { ui.alert("No se encontró la hoja INSUMOS."); return; }

  // Hojas de lugares clínicos donde buscar
  var LUGARES_CLINICOS = ["CURACIONES", "PABELLON", "UNACESS", "TOMA_MUESTRAS",
                          "LASERTERAPIA", "BOX_MEDICOS", "FOTOTERAPIA", "AREA_TECNICA"];

  // Construir mapa de códigos clínicos desde todas las hojas de lugares
  var codigosClinicos = {};
  LUGARES_CLINICOS.forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    var ini  = String(data[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;
    for (var i = ini; i < data.length; i++) {
      var cod = String(data[i][0]||"").trim();
      if (cod) codigosClinicos[cod] = true;
    }
  });

  // Leer INSUMOS y escribir col F
  var insData = shIns.getDataRange().getValues();
  var ini     = String(insData[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;
  var clinicos = 0, noClinicos = 0, yaTenia = 0;

  for (var r = ini; r < insData.length; r++) {
    var cod    = String(insData[r][0]||"").trim();
    var actual = String(insData[r][5]||"").trim();
    if (!cod) continue;

    // Si ya tiene valor en col F, no sobreescribir
    if (actual === "CLINICOS" || actual === "NO CLINICOS") { yaTenia++; continue; }

    var valor = codigosClinicos[cod] ? "CLINICOS" : "NO CLINICOS";
    shIns.getRange(r + 1, 6).setValue(valor);
    if (valor === "CLINICOS") clinicos++; else noClinicos++;
  }

  ui.alert(
    "Clasificacion completada",
    clinicos   + " insumos marcados CLINICOS. " +
    noClinicos + " marcados NO CLINICOS. " +
    yaTenia    + " ya tenian valor (no modificados).",
    ui.ButtonSet.OK
  );
}

// ── Escribir col F (Bodega) en hojas de cada lugar ───────────
// Lee col F de INSUMOS y la copia en cada hoja de lugar clínico
// Permite que el sistema sepa de qué bodega sale cada insumo
function clasificarInsumosLugares() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  // 1 — Construir mapa código → bodega desde hoja INSUMOS
  var shIns = ss.getSheetByName("INSUMOS");
  if (!shIns) { ui.alert("No se encontro la hoja INSUMOS."); return; }

  var insData = shIns.getDataRange().getValues();
  var mapaBodega = {}; // cod → "CLINICOS" o "NO CLINICOS"
  var iniIns = String(insData[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;
  for (var i = iniIns; i < insData.length; i++) {
    var cod  = String(insData[i][0]||"").trim();
    var bod  = String(insData[i][5]||"").trim();
    if (cod && bod) mapaBodega[cod] = bod;
  }

  // 2 — Hojas de lugares a procesar
  var HOJAS_LUGARES = [
    "CURACIONES", "PABELLON", "UNACESS", "TOMA_MUESTRAS",
    "LASERTERAPIA", "BOX_MEDICOS", "FOTOTERAPIA",
    "AREA_TECNICA", "OFICINA_ADMIN"
  ];

  var totalEscrito = 0, totalSinCodigo = 0;

  HOJAS_LUGARES.forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return;

    var data = sh.getDataRange().getValues();
    var ini  = String(data[0][0]||"").toUpperCase().indexOf("COD") === 0 ? 1 : 0;

    for (var r = ini; r < data.length; r++) {
      var cod    = String(data[r][0]||"").trim();
      var actual = String(data[r][5]||"").trim();
      if (!cod) continue;

      // No sobreescribir si ya tiene valor
      if (actual === "CLINICOS" || actual === "NO CLINICOS") continue;

      var bod = mapaBodega[cod];
      if (bod) {
        sh.getRange(r + 1, 6).setValue(bod);
        totalEscrito++;
      } else {
        totalSinCodigo++;
      }
    }
  });

  ui.alert(
    "Clasificacion completada",
    totalEscrito    + " insumos clasificados en hojas de lugares. " +
    totalSinCodigo  + " codigos sin match en INSUMOS.",
    ui.ButtonSet.OK
  );
}

// ── Diagnóstico col F en hojas de lugares ──────────────────
function diagnosticarColF() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var HOJAS = ["CURACIONES","PABELLON","UNACESS","TOMA_MUESTRAS",
               "LASERTERAPIA","BOX_MEDICOS","FOTOTERAPIA","AREA_TECNICA","OFICINA_ADMIN"];
  var msg = "";
  HOJAS.forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh) { msg += nombre + ": hoja no encontrada. "; return; }
    var data = sh.getDataRange().getValues();
    var conF = 0, sinF = 0;
    for (var i = 1; i < data.length; i++) {
      var cod = String(data[i][0]||"").trim();
      var f   = String(data[i][5]||"").trim();
      if (!cod) continue;
      if (f) conF++; else sinF++;
    }
    msg += nombre + ": " + conF + " con F, " + sinF + " sin F. ";
  });
  ui.alert("Diagnostico col F", msg, ui.ButtonSet.OK);
}

// ── Corrección manual SOL-260519-094240 ──────────────────────
function corregirSOL260519094240() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var mov = ss.getSheetByName("MOVIMIENTOS");
  if (!mov) { ui.alert("No se encontro hoja MOVIMIENTOS."); return; }

  var ahora = "19-05-2026, 10:15:12 p. m.";
  var sol   = "SOL-260519-094240";
  var cod   = "2002416175";
  var desc  = "TORULA C/VASTAGO DE MADERA ESTERIL";
  var qty   = 20;

  mov.appendRow([ahora, "INGRESO", sol, "", "LASERTERAPIA", cod, desc, qty, "", "", ""]);
  mov.appendRow([ahora, "EGRESO",  sol, "", "BODEGA INSUMOS CLINICOS", cod, desc, -qty, "", "", ""]);

  ui.alert("Listo. Se agregaron 2 registros en MOVIMIENTOS para " + sol);
}

// ── Menú SIIDER en el sheet ───────────────────────────────────
// ── Trigger onEdit: actualiza stock automáticamente al editar MOVIMIENTOS ──
function onEdit(e) {
  try {
    var sheet = e && e.range && e.range.getSheet();
    if (!sheet) return;
    var nombre = sheet.getName();
    // Si se edita MOVIMIENTOS o INVENTARIO → actualizar stock
    if (nombre === SHEET_MOVIMIENTOS || nombre === SHEET_DATOS) {
      actualizarStockCuraciones();
    }
  } catch(err) {
    Logger.log("onEdit error: " + err);
  }
}

// ── Actualizar stock de TODOS los lugares + bodegas ──────────
function actualizarTodoElStock() {
  var ui = SpreadsheetApp.getUi();
  var resumen = [];
  LUGARES_STOCK.forEach(function(lu) {
    try {
      var n = actualizarStockLugar(lu);
      resumen.push("✓ " + lu + ": " + n + " ítems");
    } catch(e) {
      resumen.push("✗ " + lu + ": " + e.message);
    }
  });
  try { actualizarStockBodega(); resumen.push("✓ BODEGA (general)"); } catch(e) {}
  ui.alert("✅ Stock actualizado", resumen.join("\n"), ui.ButtonSet.OK);
}

// ── Funciones individuales por lugar (para el menú) ──────────
function stockCuraciones()   { actualizarStockLugar("CURACIONES");   SpreadsheetApp.getUi().alert("✓ STOCK_CURACIONES actualizado."); }
function stockPabellon()     { actualizarStockLugar("PABELLON");     SpreadsheetApp.getUi().alert("✓ STOCK_PABELLON actualizado."); }
function stockUnacess()      { actualizarStockLugar("UNACESS");      SpreadsheetApp.getUi().alert("✓ STOCK_UNACESS actualizado."); }
function stockLaserterapia() { actualizarStockLugar("LASERTERAPIA"); SpreadsheetApp.getUi().alert("✓ STOCK_LASERTERAPIA actualizado."); }
function stockTomaMuestras() { actualizarStockLugar("TOMA_MUESTRAS");SpreadsheetApp.getUi().alert("✓ STOCK_TOMA_MUESTRAS actualizado."); }
function stockBoxMedicos()   { actualizarStockLugar("BOX_MEDICOS");  SpreadsheetApp.getUi().alert("✓ STOCK_BOX_MEDICOS actualizado."); }
function stockFototerapia()  { actualizarStockLugar("FOTOTERAPIA");  SpreadsheetApp.getUi().alert("✓ STOCK_FOTOTERAPIA actualizado."); }
function stockAreaTecnica()  { actualizarStockLugar("AREA_TECNICA"); SpreadsheetApp.getUi().alert("✓ STOCK_AREA_TECNICA actualizado."); }
function stockOficinaAdmin() { actualizarStockLugar("OFICINA_ADMIN");SpreadsheetApp.getUi().alert("✓ STOCK_OFICINA_ADMIN actualizado."); }

function onOpen() {
  var ui   = SpreadsheetApp.getUi();
  var menu = ui.createMenu("⚙ SIIDER");

  menu.addItem("📊 Actualizar TODO el stock (todos los lugares)", "actualizarTodoElStock")
      .addSeparator()
      .addSubMenu(ui.createMenu("📍 Actualizar lugar individual")
        .addItem("Curaciones",      "stockCuraciones")
        .addItem("Pabellón",        "stockPabellon")
        .addItem("UNACESS",         "stockUnacess")
        .addItem("Laserterapia",    "stockLaserterapia")
        .addItem("Toma de muestras","stockTomaMuestras")
        .addItem("Box médicos",     "stockBoxMedicos")
        .addItem("Fototerapia",     "stockFototerapia")
        .addItem("Área técnica",    "stockAreaTecnica")
        .addItem("Oficina admin",   "stockOficinaAdmin")
      )
      .addSeparator()
      .addItem("📦 Archivar movimientos anuales", "archivarMovimientosAnuales")
      .addSeparator()
      .addItem("🏷 Clasificar insumos bodega (INSUMOS)", "clasificarInsumos")
      .addItem("🏷 Clasificar insumos lugares (hojas)",  "clasificarInsumosLugares")
      .addItem("🔍 Diagnosticar col F en hojas",         "diagnosticarColF")
      .addItem("🛠 Corregir SOL-260519-094240",          "corregirSOL260519094240")
      .addToUi();
}

// ══════════════════════════════════════════════════════════════
//  ENVIAR REPORTE POR CORREO — Panel Reportes de Gestión
//  Parámetros GET:
//    accion=enviarReporte
//    destinatario = email destino
//    asunto       = asunto del correo
//    tipo         = consumo | semaforo | solicitudes | ranking
//    htmlTabla    = contenido HTML de la tabla (URL-encoded)
//    periodo      = texto descriptivo del período (ej: "Mayo 2026")
// ══════════════════════════════════════════════════════════════
function enviarReporte(e) {
  try {
    var p            = e.parameter || {};
    var destinatario = p.destinatario || "";
    var asunto       = p.asunto      || "Reporte SIIDER — Dermatología HDS";
    var htmlTabla    = p.htmlTabla   || "";
    var periodo      = p.periodo     || "";
    var tipo         = p.tipo        || "reporte";

    if (!destinatario) return { status: "error", mensaje: "Falta destinatario." };

    var ahora   = new Date().toLocaleString("es-CL");
    var titulos = {
      consumo:      "📊 Consumo Mensual por Lugar",
      semaforo:     "🚦 Semáforo Consolidado de Stock",
      solicitudes:  "📋 Estado de Solicitudes",
      ranking:      "🏆 Ranking de Insumos más Consumidos"
    };
    var titulo = titulos[tipo] || "Reporte de Gestión";

    var htmlCuerpo = [
      '<html><body style="font-family:Arial,sans-serif;color:#1e293b;max-width:700px;margin:0 auto;">',
      '<div style="background:#185FA5;padding:16px 20px;border-radius:8px 8px 0 0;">',
      '<h2 style="color:#fff;margin:0;font-size:16px;">SIIDER · Dermatología · Hospital del Salvador</h2>',
      '<p style="color:#bfdbfe;margin:4px 0 0;font-size:13px;">' + titulo + '</p>',
      '</div>',
      '<div style="border:1px solid #d1dce8;border-top:none;border-radius:0 0 8px 8px;padding:16px;">',
      periodo ? '<p style="font-size:13px;color:#64748b;margin-bottom:12px;">📅 Período: <strong>' + periodo + '</strong></p>' : '',
      htmlTabla,
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:16px 0;">',
      '<p style="font-size:11px;color:#94a3b8;">Generado el ' + ahora + ' desde SIIDER.<br>',
      'Este es un correo automático — no responder a esta dirección.</p>',
      '</div></body></html>'
    ].join("");

    GmailApp.sendEmail(destinatario, asunto, "Reporte adjunto (ver versión HTML).", {
      htmlBody: htmlCuerpo,
      name: "SIIDER — Dermatología HDS"
    });

    return { status: "ok", mensaje: "Correo enviado a " + destinatario };
  } catch(err) {
    return { status: "error", mensaje: err.toString() };
  }
}
