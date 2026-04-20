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

    var result = (function() { switch (action) {
      // ── Descarga maestros (precarga localStorage) ──────────
      case 'descargarMaestros':    return descargarMaestros();
      case 'descargarOperacional': return descargarOperacional();

      // ── Dashboard ──────────────────────────────────────────
      case 'getDashboard':       return getDashboard();

      // ── Productos ──────────────────────────────────────────
      case 'getProductos':       return getProductos(params);
      case 'getProducto':        return getProducto(params.codigo);
      case 'crearProducto':      return crearProducto(params);
      case 'actualizarProducto': return actualizarProducto(params);

      // ── Stock ──────────────────────────────────────────────
      case 'getStock':              return getStock(params);
      case 'getStockProducto':      return getStockProducto(params.codigo);
      case 'getHistorialStock':     return getHistorialStock(params);
      case 'imprimirHistorialStock':return imprimirHistorialStock(params);

      // ── Guias ──────────────────────────────────────────────
      case 'getGuias':           return getGuias(params);
      case 'getGuia':            return getGuia(params.idGuia);
      case 'crearGuia':          return crearGuia(params);
      case 'agregarDetalleGuia':        return agregarDetalleGuia(params);
      case 'actualizarFechaVencimiento':  return actualizarFechaVencimiento(params);
      case 'actualizarCantidadDetalle':   return actualizarCantidadDetalle(params);
      case 'cerrarGuia':                return cerrarGuia(params.idGuia, params.usuario, params.idSesion);
      case 'reabrirGuia':               return reabrirGuia(params);
      case 'anularDetalle':             return anularDetalle(params);
      case 'autoCloseDayGuias':  return autoCloseDayGuias();

      // ── Preingresos ────────────────────────────────────────
      case 'getPreingresos':            return getPreingresos(params);
      case 'crearPreingreso':           return crearPreingreso(params);
      case 'aprobarPreingreso':         return aprobarPreingreso(params);
      case 'subirFotoPreingreso':       return subirFotoPreingreso(params);
      case 'actualizarFotosPreingreso': return actualizarFotosPreingreso(params);
      case 'actualizarPreingreso':      return actualizarPreingreso(params);
      case 'eliminarFotoDrive':         return eliminarFotoDrive(params);

      // ── Guías — foto + comentario ──────────────────────────
      case 'subirFotoGuia':          return subirFotoGuia(params);
      case 'actualizarGuia':         return actualizarGuia(params);
      case 'copiarFotoDePreingreso': return copiarFotoDePreingreso(params);

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
      case 'auditarProducto':    return auditarProducto(params);

      // ── Ajustes ────────────────────────────────────────────
      case 'getAjustes':         return getAjustes(params);
      case 'crearAjuste':        return crearAjuste(params);

      // ── Producto Nuevo ─────────────────────────────────────
      case 'getProductosNuevos': return getProductosNuevos(params);
      case 'registrarProductoNuevo': return registrarProductoNuevo(params);
      case 'aprobarProductoNuevo':   return aprobarProductoNuevo(params);

      // ── Lotes ──────────────────────────────────────────────
      case 'getLotesVencimiento': return getLotesVencimiento(params);

      // ── Config ─────────────────────────────────────────────
      case 'getConfig':          return getConfigAll();
      case 'setConfigValue':     return setConfigValue(params.clave, params.valor);

      // ── PrintNode ──────────────────────────────────────────
      case 'imprimirEtiqueta':    return imprimirEtiqueta(params);
      case 'imprimirBienvenida':  return imprimirBienvenida(params);
      case 'imprimirMembrete':    return imprimirMembrete(params);

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
// Campos que deben forzarse a string (evita pérdida de ceros a la izquierda en Sheets)
var _STRING_FIELDS = { codigoBarra: true, barcode: true, ean: true, pin: true, adminPin: true };

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
      } else if (_STRING_FIELDS[h] && v !== '' && v !== null && v !== undefined) {
        obj[h] = String(v);   // preservar ceros a la izquierda
      } else {
        obj[h] = v;
      }
    });
    return obj;
  }).filter(function(obj) {
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
  var sCod = String(codigo);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === sCod) {
      return { fila: i + 1, cantidad: parseFloat(data[i][2]) || 0 };
    }
  }
  return { fila: -1, cantidad: 0 };
}

// ============================================================
// Bridge hacia ProyectoMOS — siempre lee de MOS_SS_ID
// Sin fallback local: si MOS_SS_ID no está configurado, error claro.
// ============================================================

function _getMosSS() {
  var id = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (!id) throw new Error('MOS_SS_ID no configurado en Script Properties. Contactar administrador.');
  try { return SpreadsheetApp.openById(id); }
  catch(e) { throw new Error('No se pudo abrir el Spreadsheet de MOS (' + id + '): ' + e.message); }
}

function getProductosSheet() {
  var sheet = _getMosSS().getSheetByName('PRODUCTOS_MASTER');
  if (!sheet) throw new Error('Hoja PRODUCTOS_MASTER no encontrada en MOS.');
  return sheet;
}

function getProveedoresSheet() {
  var sheet = _getMosSS().getSheetByName('PROVEEDORES_MASTER');
  if (!sheet) throw new Error('Hoja PROVEEDORES_MASTER no encontrada en MOS.');
  return sheet;
}

function getPersonalSheet() {
  var sheet = _getMosSS().getSheetByName('PERSONAL_MASTER');
  if (!sheet) throw new Error('Hoja PERSONAL_MASTER no encontrada en MOS.');
  return sheet;
}

// Helper: devuelve todo el personal activo de MOS (cualquier appOrigen)
function _getPersonalWH() {
  var rows = _sheetToObjects(getPersonalSheet());
  return rows.filter(function(p) {
    return String(p.estado) === '1';
  });
}

// Devuelve el PrintNode ID de una impresora en MOS.
// tipo: 'ADHESIVO' | 'TICKET'. idZona opcional (ej: 'ALMACEN').
function getPrinterNodeId(tipo, idZona) {
  var sheet = _getMosSS().getSheetByName('IMPRESORAS');
  if (!sheet) throw new Error('Hoja IMPRESORAS no encontrada en MOS.');
  var rows = _sheetToObjects(sheet);
  var imp  = rows.find(function(r) {
    return r.tipo === tipo &&
           String(r.activo) === '1' &&
           String(r.printNodeId || '') !== '' &&
           (!idZona || r.idZona === idZona);
  });
  if (!imp) throw new Error(
    'No hay impresora tipo ' + tipo + (idZona ? ' zona ' + idZona : '') + ' activa en MOS.'
  );
  return imp.printNodeId;
}

// ============================================================
// descargarMaestros — descarga las 4 tablas maestras en un solo
// request para poblar el localStorage del dispositivo.
// ============================================================
function descargarMaestros() {
  var mosSS;
  try { mosSS = _getMosSS(); }
  catch(e) { return { ok: false, error: e.message }; }

  var result = { productos: [], equivalencias: [], proveedores: [], personal: [], impresoras: [], zonas: [] };
  var errores = [];

  // PRODUCTOS_MASTER
  try {
    var sh = mosSS.getSheetByName('PRODUCTOS_MASTER');
    result.productos = sh ? _sheetToObjects(sh) : [];
  } catch(e) { errores.push('productos: ' + e.message); }

  // EQUIVALENCIAS (barcodes de almacén que apuntan al skuBase)
  try {
    var sh = mosSS.getSheetByName('EQUIVALENCIAS');
    result.equivalencias = sh ? _sheetToObjects(sh).filter(function(e) {
      return String(e.activo) === '1';
    }) : [];
  } catch(e) { errores.push('equivalencias: ' + e.message); }

  // PROVEEDORES_MASTER
  try {
    var sh = mosSS.getSheetByName('PROVEEDORES_MASTER');
    result.proveedores = sh ? _sheetToObjects(sh) : [];
  } catch(e) { errores.push('proveedores: ' + e.message); }

  // PERSONAL_MASTER — todo el personal activo (cualquier appOrigen puede usar warehouseMos)
  try {
    var sh = mosSS.getSheetByName('PERSONAL_MASTER');
    result.personal = sh ? _sheetToObjects(sh).filter(function(p) {
      return String(p.estado) === '1';
    }) : [];
  } catch(e) { errores.push('personal: ' + e.message); }

  // IMPRESORAS — solo appOrigen=warehouseMos y activo=1
  try {
    var sh = mosSS.getSheetByName('IMPRESORAS');
    result.impresoras = sh ? _sheetToObjects(sh).filter(function(r) {
      return String(r.appOrigen || '').toLowerCase() === 'warehousemos' &&
             String(r.activo) === '1';
    }) : [];
  } catch(e) { errores.push('impresoras: ' + e.message); }

  // ZONAS (hoja local en warehouseMos)
  try {
    var whSS = SpreadsheetApp.getActiveSpreadsheet();
    var zonasSheet = whSS.getSheetByName('ZONAS');
    result.zonas = zonasSheet ? _sheetToObjects(zonasSheet).filter(function(z) {
      return String(z.estado) === '1';
    }) : [];
  } catch(e) { errores.push('zonas: ' + e.message); }

  // ESTACIONES — adminPin de ALMACEN (para reabrir guías)
  try {
    var estSheet = mosSS.getSheetByName('ESTACIONES');
    if (estSheet) {
      var estaciones = _sheetToObjects(estSheet);
      var almacen = estaciones.find(function(e) {
        return String(e.idEstacion || e.nombre || '').toUpperCase() === 'ALMACEN';
      });
      if (almacen && almacen.adminPin) result.adminPin = String(almacen.adminPin);
    }
  } catch(e) { errores.push('adminPin: ' + e.message); }

  return { ok: true, data: result, errores: errores, generadoEn: new Date().toISOString() };
}

// ============================================================
// descargarOperacional — guías, detalles y preingresos en un
// solo request para poblar el localStorage del dispositivo.
// Se llama antes del login y cada 60s en background.
// ============================================================
function descargarOperacional() {
  var result  = { guias: [], detalles: [], preingresos: [], stock: [], ajustes: [], auditorias: [] };
  var errores = [];

  try {
    var guias = _sheetToObjects(getSheet('GUIAS'));
    result.guias = guias.slice().reverse().slice(0, 200);
  } catch(e) { errores.push('guias: ' + e.message); }

  try {
    result.detalles = _sheetToObjects(getSheet('GUIA_DETALLE'));
  } catch(e) { errores.push('detalles: ' + e.message); }

  try {
    var preingresos = _sheetToObjects(getSheet('PREINGRESOS'));
    var cutoff30 = new Date(); cutoff30.setDate(cutoff30.getDate() - 30);
    result.preingresos = preingresos.filter(function(p) {
      var f = new Date(p.fecha);
      return isNaN(f.getTime()) || f >= cutoff30;
    });
  } catch(e) { errores.push('preingresos: ' + e.message); }

  try {
    result.stock = _sheetToObjects(getSheet('STOCK'));
  } catch(e) { errores.push('stock: ' + e.message); }

  try {
    var ajustes = _sheetToObjects(getSheet('AJUSTES'));
    var cutoff90 = new Date(); cutoff90.setDate(cutoff90.getDate() - 90);
    result.ajustes = ajustes.filter(function(a) {
      var f = new Date(a.fecha);
      return isNaN(f.getTime()) || f >= cutoff90;
    });
  } catch(e) { errores.push('ajustes: ' + e.message); }

  try {
    var auditorias = _sheetToObjects(getSheet('AUDITORIAS'));
    var cutoff60 = new Date(); cutoff60.setDate(cutoff60.getDate() - 60);
    result.auditorias = auditorias.filter(function(a) {
      var f = new Date(a.fechaEjecucion || a.fechaAsignacion);
      return isNaN(f.getTime()) || f >= cutoff60;
    });
  } catch(e) { errores.push('auditorias: ' + e.message); }

  return { ok: true, data: result, errores: errores, generadoEn: new Date().toISOString() };
}

// ============================================================
// imprimirBienvenida — ticket de inicio de turno
// Busca la impresora TICKET de zona ALMACEN y envía ESC/POS
// ============================================================
function imprimirBienvenida(params) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  var printerId;
  try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
  catch(e) { return { ok: false, error: e.message }; }

  var tz      = Session.getScriptTimeZone();
  var ahora   = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var empresa = _getConfigValue('EMPRESA_NOMBRE') || 'InversionMos';
  var nombre  = String(params.nombre   || '');
  var apellido= String(params.apellido || '');
  var rol     = String(params.rol      || '');
  var hora    = String(params.horaInicio || Utilities.formatDate(new Date(), tz, 'HH:mm:ss'));

  var t = '';
  t += '\x1b\x40';                          // Init impresora
  t += '\x1b\x61\x01';                      // Centrar
  t += '\x1b\x21\x30';                      // Doble alto + ancho
  t += empresa + '\n';
  t += '\x1b\x21\x00';                      // Normal
  t += 'warehouseMos \u2014 Alm\u00e1cen\n';
  t += '================================\n\n';
  t += '\x1b\x21\x10';                      // Doble alto
  t += 'INICIO DE TURNO\n';
  t += '\x1b\x21\x00\n';
  t += '\x1b\x61\x01';
  t += '\x1b\x21\x20';                      // Doble ancho
  t += nombre + ' ' + apellido + '\n';
  t += '\x1b\x21\x00';
  t += rol + '\n\n';
  t += '================================\n';
  t += '\x1b\x61\x00';                      // Izquierda
  t += 'Hora inicio : ' + hora + '\n';
  t += 'Impreso     : ' + ahora + '\n';
  t += '\n';
  t += '\x1b\x61\x01';
  t += 'Bienvenido al turno. \u00a1Mucho \u00e9xito!\n';
  t += '\n\n\n\n\n';
  t += '\x1d\x56\x00';                      // Corte

  var payload = {
    printerId:   parseInt(printerId),
    title:       'Bienvenida ' + nombre + ' ' + apellido,
    contentType: 'raw_base64',
    content:     Utilities.base64Encode(t),
    source:      'warehouseMos'
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    return code === 201
      ? { ok: true }
      : { ok: false, error: 'PrintNode ' + code + ': ' + resp.getContentText() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function _actualizarStock(codigo, delta) {
  var sheet = getSheet('STOCK');
  var info = _getStockProducto(codigo);
  var nuevaCantidad = Math.max(0, info.cantidad + delta);
  var now = new Date();

  if (info.fila > 0) {
    sheet.getRange(info.fila, 3).setNumberFormat('0.##');
    sheet.getRange(info.fila, 4).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(info.fila, 3, 1, 2).setValues([[nuevaCantidad, now]]);
  } else {
    // Nueva fila: preservar barcode como texto
    var nextRow = sheet.getLastRow() + 1;
    var stVals  = ['STK' + new Date().getTime(), String(codigo), nuevaCantidad, now];
    sheet.getRange(nextRow, 2).setNumberFormat('@');
    sheet.getRange(nextRow, 4).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(nextRow, 1, 1, stVals.length).setValues([stVals]);
  }
  return nuevaCantidad;
}
