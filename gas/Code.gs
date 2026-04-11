// ============================================================
// warehouseMos — Code.gs
// Router principal del Web App (GAS)
// Desplegar como Web App: Execute as Me, Anyone with link
// ============================================================

var SS_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

function getSpreadsheet() {
  return SpreadsheetApp.openById(SS_ID);
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

// ============================================================
// CORS + Entry Points
// ============================================================
function doGet(e) {
  var result = _route('GET', e);
  return _respond(result);
}

function doPost(e) {
  var result = _route('POST', e);
  return _respond(result);
}

// ── Deduplicación por localId ──────────────────────────────
function _checkDuplicado(localId) {
  if (!localId || !String(localId).startsWith('L')) return null;
  try {
    var sheet = getSheet('SYNC_LOG');
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === localId) {
        try { return JSON.parse(data[i][2]); } catch { return { ok: true, dedup: true }; }
      }
    }
  } catch(e) {}
  return null;
}

function _logSync(localId, action, resultado) {
  if (!localId || !String(localId).startsWith('L')) return;
  try {
    var sheet = getSheet('SYNC_LOG');
    if (sheet) sheet.appendRow([localId, action, JSON.stringify(resultado), new Date()]);
  } catch(e) {}
}

function _respond(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// Router
// ============================================================
function _route(method, e) {
  try {
    var params = (method === 'GET')
      ? e.parameter
      : JSON.parse(e.postData ? e.postData.contents : '{}');

    var action = params.action || '';

    // Deduplicar operaciones offline (solo para POSTs con localId)
    if (params.localId) {
      var dup = _checkDuplicado(params.localId);
      if (dup) return dup;
    }

    result = (function() { switch (action) {
      // ── Dashboard ──────────────────────────────────────────
      case 'getDashboard':       return getDashboard();

      // ── Productos ──────────────────────────────────────────
      case 'getProductos':       return getProductos(params);
      case 'getProducto':        return getProducto(params.codigo);
      case 'crearProducto':      return crearProducto(params);
      case 'actualizarProducto': return actualizarProducto(params);

      // ── Stock ──────────────────────────────────────────────
      case 'getStock':           return getStock(params);
      case 'getStockProducto':   return getStockProducto(params.codigo);

      // ── Guias ──────────────────────────────────────────────
      case 'getGuias':           return getGuias(params);
      case 'getGuia':            return getGuia(params.idGuia);
      case 'crearGuia':          return crearGuia(params);
      case 'agregarDetalleGuia': return agregarDetalleGuia(params);
      case 'cerrarGuia':         return cerrarGuia(params.idGuia, params.usuario);

      // ── Preingresos ────────────────────────────────────────
      case 'getPreingresos':     return getPreingresos(params);
      case 'crearPreingreso':    return crearPreingreso(params);
      case 'aprobarPreingreso':  return aprobarPreingreso(params);

      // ── Envasados ──────────────────────────────────────────
      case 'getEnvasados':       return getEnvasados(params);
      case 'getPendientesEnvasado': return getPendientesEnvasado();
      case 'registrarEnvasado':  return registrarEnvasado(params);

      // ── Proveedores ────────────────────────────────────────
      case 'getProveedores':     return getProveedores(params);
      case 'crearProveedor':     return crearProveedor(params);
      case 'actualizarProveedor': return actualizarProveedor(params);

      // ── Mermas ─────────────────────────────────────────────
      case 'getMermas':          return getMermas(params);
      case 'registrarMerma':     return registrarMerma(params);

      // ── Auditorias ─────────────────────────────────────────
      case 'getAuditorias':      return getAuditorias(params);
      case 'asignarAuditoria':   return asignarAuditoria(params);
      case 'ejecutarAuditoria':  return ejecutarAuditoria(params);

      // ── Ajustes ────────────────────────────────────────────
      case 'getAjustes':         return getAjustes(params);
      case 'crearAjuste':        return crearAjuste(params);

      // ── Producto Nuevo ─────────────────────────────────────
      case 'getProductosNuevos': return getProductosNuevos(params);
      case 'registrarProductoNuevo': return registrarProductoNuevo(params);
      case 'aprobarProductoNuevo':   return aprobarProductoNuevo(params);

      // ── Lotes ──────────────────────────────────────────────
      case 'getLotesVencimiento': return getLotesVencimiento(params);

      // ── Config (solo lectura desde frontend) ───────────────
      case 'getConfig':          return getConfigAll();

      // ── PrintNode ──────────────────────────────────────────
      case 'imprimirEtiqueta':   return imprimirEtiqueta(params);

      // ── Personal ───────────────────────────────────────────
      case 'loginPersonal':      return loginPersonal(params);
      case 'cerrarTurno':        return cerrarTurno(params);
      case 'getPersonal':        return getPersonal(params);
      case 'getPersonalConPin':  return getPersonalConPin(params);
      case 'getSesionActiva':    return getSesionActiva(params);
      case 'getDesempenoDia':    return getDesempenoDia(params);
      case 'getResumenPersonal': return getResumenPersonal(params);

      default:
        return { ok: false, error: 'Acción no reconocida: ' + action };
    }})();

    // Loguear para deduplicación futura
    if (params.localId && result && result.ok) {
      _logSync(params.localId, action, result);
    }
    return result;
  } catch (err) {
    return { ok: false, error: err.message, stack: err.stack };
  }
}

// ============================================================
// Helpers compartidos
// ============================================================
function _sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v instanceof Date) {
        obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[h] = v;
      }
    });
    return obj;
  }).filter(function(obj) {
    // Filtrar filas completamente vacías
    return Object.values(obj).some(function(v) { return v !== '' && v !== null && v !== undefined; });
  });
}

function _generateId(prefix) {
  return prefix + new Date().getTime();
}

function _getConfigValue(clave) {
  var sheet = getSheet('CONFIG');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clave) return data[i][1];
  }
  return null;
}

function getConfigAll() {
  var sheet = getSheet('CONFIG');
  var rows = _sheetToObjects(sheet);
  var config = {};
  rows.forEach(function(r) { config[r.clave] = r.valor; });
  return { ok: true, data: config };
}

function setConfigValue(clave, valor) {
  var sheet = getSheet('CONFIG');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valor);
      return { ok: true };
    }
  }
  sheet.appendRow([clave, valor, '']);
  return { ok: true };
}

// ============================================================
// Stock helpers (usados por múltiples módulos)
// ============================================================
function _getStockProducto(codigo) {
  var sheet = getSheet('STOCK');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === codigo) {
      return { fila: i + 1, cantidad: parseFloat(data[i][2]) || 0 };
    }
  }
  return { fila: -1, cantidad: 0 };
}

// ============================================================
// Bridge hacia ProyectoMOS
// Cuando MOS_SS_ID esté configurado, warehouseMos lee
// PRODUCTOS_MASTER y PROVEEDORES_MASTER del Sheet de MOS.
// Mientras esté vacío, usa sus propias tablas (sin cambios).
// ============================================================
function getProductosSheet() {
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (mosSsId) {
    try { return SpreadsheetApp.openById(mosSsId).getSheetByName('PRODUCTOS_MASTER'); }
    catch(e) { Logger.log('[WH] getProductosSheet fallback: ' + e.message); }
  }
  return getSheet('PRODUCTOS');
}

function getProveedoresSheet() {
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (mosSsId) {
    try { return SpreadsheetApp.openById(mosSsId).getSheetByName('PROVEEDORES_MASTER'); }
    catch(e) {}
  }
  return getSheet('PROVEEDORES');
}

// Personal: lee PERSONAL_MASTER de MOS (solo tipo=OPERADOR/appOrigen=warehouseMos)
// Si MOS no conectado → tabla PERSONAL local
function getPersonalSheet() {
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (mosSsId) {
    try { return SpreadsheetApp.openById(mosSsId).getSheetByName('PERSONAL_MASTER'); }
    catch(e) { Logger.log('[WH] getPersonalSheet fallback: ' + e.message); }
  }
  return getSheet('PERSONAL');
}

// Devuelve el PrintNode ID de una impresora configurada en MOS para WH.
// tipo: 'ADHESIVO' o 'TICKET'. Fallback a Script Properties (modo standalone).
function getPrinterNodeId(tipo) {
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (mosSsId) {
    try {
      var sheet = SpreadsheetApp.openById(mosSsId).getSheetByName('IMPRESORAS');
      var rows  = _sheetToObjects(sheet);
      var imp   = rows.find(function(r){
        return r.appOrigen === 'warehouseMos' &&
               r.tipo      === tipo           &&
               String(r.activo) === '1'       &&
               r.printNodeId    !== '';
      });
      if (imp) return imp.printNodeId;
    } catch(e) {}
  }
  // Fallback: Script Properties originales
  return tipo === 'ADHESIVO'
    ? (PropertiesService.getScriptProperties().getProperty('PRINTER_ETIQUETAS_ID') || '')
    : (PropertiesService.getScriptProperties().getProperty('PRINTER_TICKETS_ID')   || '');
}

function _actualizarStock(codigo, delta) {
  var sheet = getSheet('STOCK');
  var info = _getStockProducto(codigo);
  var nuevaCantidad = Math.max(0, info.cantidad + delta);
  var now = new Date();

  if (info.fila > 0) {
    sheet.getRange(info.fila, 3, 1, 2).setValues([[nuevaCantidad, now]]);
  } else {
    var id = 'STK' + new Date().getTime();
    sheet.appendRow([id, codigo, nuevaCantidad, now]);
  }
  return nuevaCantidad;
}
