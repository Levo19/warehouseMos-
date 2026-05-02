// ============================================================
// warehouseMos — Auditoria.gs
// Cierre forzado nocturno + auditoría de cuadre stock vs historial.
// Genera alertas para diferencias > 0.5 que se muestran en el dashboard.
//
// Triggers (ejecutar setupTriggersAuditoria una vez desde el editor):
//   21:00 → cerrarGuiasAbiertasGlobal (cierra todas las ABIERTA del día)
//   22:00 → auditarStockGlobal (compara stock real vs teórico)
// ============================================================

// ============================================================
// 1. Cierre forzado de TODAS las guías abiertas (21:00)
// ============================================================
function cerrarGuiasAbiertasGlobal() {
  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId  = hdrs.indexOf('idGuia');
  var idxEst = hdrs.indexOf('estado');
  if (idxId < 0 || idxEst < 0) return { ok: false, error: 'Columnas no encontradas' };

  var abiertas = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxEst]).toUpperCase() === 'ABIERTA') {
      abiertas.push(String(data[i][idxId]));
    }
  }

  var ok = 0, err = 0;
  abiertas.forEach(function(idGuia) {
    try {
      var res = cerrarGuia(idGuia, 'sistema-cierre-21h', null);
      if (res.ok) ok++; else err++;
    } catch(e) { err++; Logger.log('Error cerrando ' + idGuia + ': ' + e.message); }
  });

  Logger.log('cerrarGuiasAbiertasGlobal: ' + ok + ' cerradas, ' + err + ' errores de ' + abiertas.length + ' totales');
  return { ok: true, data: { total: abiertas.length, ok: ok, error: err } };
}

// ============================================================
// 2. Auditoría de cuadre stock vs historial (22:00)
//    Stock teórico = sum(AJUSTES) + sum(detalles INGRESO cerrados)
//                    - sum(detalles SALIDA cerrados)
//    Diff = stock_real - stock_teorico
//    Si |diff| > 0.5 → alerta
// ============================================================
function auditarStockGlobal() {
  var stockSheet = getSheet('STOCK');
  var stockData  = _sheetToObjects(stockSheet);
  var guias      = _sheetToObjects(getSheet('GUIAS'));
  var detalles   = _sheetToObjects(getSheet('GUIA_DETALLE'));
  var ajustes    = [];
  try { ajustes = _sheetToObjects(getSheet('AJUSTES')); } catch(e) {}

  // Solo guías CERRADAS o AUTOCERRADAS aplican stock
  var guiaMap = {};
  guias.forEach(function(g) {
    var est = String(g.estado || '').toUpperCase();
    if (est === 'CERRADA' || est === 'AUTOCERRADA') {
      guiaMap[g.idGuia] = String(g.tipo || '').toUpperCase();
    }
  });

  // Resolver nombre de producto para el reporte
  var prodMap = {};
  try {
    var prods = _sheetToObjects(getProductosSheet());
    prods.forEach(function(p) {
      var name = p.descripcion || p.nombre || '';
      if (!name) return;
      if (p.codigoBarra) prodMap[String(p.codigoBarra).trim()] = name;
      if (p.idProducto)  prodMap[String(p.idProducto)] = name;
    });
    // Equivalencias: indexan al nombre del producto base
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      _sheetToObjects(equivSheet).forEach(function(e) {
        if (!e.codigoBarra || !e.skuBase) return;
        var name = prodMap[String(e.skuBase)] || prodMap[String(e.skuBase).trim().toUpperCase()];
        if (name) prodMap[String(e.codigoBarra).trim()] = name;
      });
    }
  } catch(e) {}

  // Precalcular movimientos teóricos por código
  var teoricos = {};  // codigoProducto → { entradas, salidas }
  function _ensure(cod) {
    if (!teoricos[cod]) teoricos[cod] = { entradas: 0, salidas: 0 };
    return teoricos[cod];
  }

  // AJUSTES — INC/INI suman, DEC resta
  ajustes.forEach(function(a) {
    var cod  = String(a.codigoProducto || '').trim();
    var cant = Math.abs(parseFloat(a.cantidadAjuste || 0));
    var tipo = String(a.tipoAjuste || '').toUpperCase();
    if (!cod || cant <= 0) return;
    var t = _ensure(cod);
    if (tipo === 'INC' || tipo === 'INI') t.entradas += cant;
    else if (tipo === 'DEC') t.salidas += cant;
  });

  // GUIA_DETALLE — solo de guías CERRADAS y no anuladas
  detalles.forEach(function(d) {
    if (String(d.observacion || '').toUpperCase() === 'ANULADO') return;
    var tipoGuia = guiaMap[d.idGuia];
    if (!tipoGuia) return; // guía no cerrada
    var cod  = String(d.codigoProducto || '').trim();
    var cant = Math.abs(parseFloat(d.cantidadRecibida || d.cantidadReal || d.cantidadEsperada || 0));
    if (!cod || cant <= 0) return;
    var t = _ensure(cod);
    if (tipoGuia.indexOf('INGRESO') === 0) t.entradas += cant;
    else if (tipoGuia.indexOf('SALIDA') === 0) t.salidas += cant;
  });

  // Comparar stock real vs teórico
  var alertas = [];
  var fechaAud = new Date();
  stockData.forEach(function(s) {
    var cod  = String(s.codigoProducto || '').trim();
    if (!cod) return;
    var real = parseFloat(s.cantidadDisponible) || 0;
    var t    = teoricos[cod] || { entradas: 0, salidas: 0 };
    var teor = t.entradas - t.salidas;
    var diff = real - teor;
    if (Math.abs(diff) > 0.5) {
      alertas.push({
        codigoProducto: cod,
        descripcion:    prodMap[cod] || cod,
        stockReal:      real,
        stockTeorico:   teor,
        diferencia:     diff
      });
    }
  });

  // Productos sin fila en STOCK pero con movimientos teóricos != 0 — también son alertas
  Object.keys(teoricos).forEach(function(cod) {
    var existeEnStock = stockData.some(function(s) { return String(s.codigoProducto).trim() === cod; });
    if (existeEnStock) return;
    var t = teoricos[cod];
    var teor = t.entradas - t.salidas;
    if (Math.abs(teor) > 0.5) {
      alertas.push({
        codigoProducto: cod,
        descripcion:    prodMap[cod] || cod,
        stockReal:      0,
        stockTeorico:   teor,
        diferencia:     -teor
      });
    }
  });

  // Guardar en hoja ALERTAS_STOCK
  _guardarAlertasStock(alertas, fechaAud);

  Logger.log('auditarStockGlobal: ' + alertas.length + ' alertas generadas (' + stockData.length + ' productos auditados)');
  return { ok: true, data: { alertas: alertas.length, productos: stockData.length } };
}

// ============================================================
// Persistencia de alertas en hoja ALERTAS_STOCK
// Reemplaza las alertas anteriores no revisadas con las nuevas para
// no acumular ruido (las revisadas se preservan como histórico).
// ============================================================
function _guardarAlertasStock(alertas, fechaAud) {
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('ALERTAS_STOCK');
  if (!sheet) {
    sheet = ss.insertSheet('ALERTAS_STOCK');
    sheet.getRange(1, 1, 1, 9).setValues([[
      'idAlerta','fecha','codigoProducto','descripcion',
      'stockReal','stockTeorico','diferencia','revisado','fechaRevision'
    ]]);
  }

  // Borrar las NO revisadas previas — solo conservamos histórico de revisadas
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var idxRev = hdrs.indexOf('revisado');
  var filasABorrar = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxRev] || '').toUpperCase() !== 'SI') {
      filasABorrar.push(i + 1);
    }
  }
  filasABorrar.forEach(function(f) { sheet.deleteRow(f); });

  // Agregar las nuevas
  if (!alertas.length) return;
  var rows = alertas.map(function(a) {
    return [
      'AL' + new Date().getTime() + Math.floor(Math.random() * 1000),
      fechaAud,
      String(a.codigoProducto),
      String(a.descripcion),
      a.stockReal,
      a.stockTeorico,
      a.diferencia,
      'NO',
      ''
    ];
  });
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(nextRow, 2, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(nextRow, 3, rows.length, 1).setNumberFormat('@');
}

// ============================================================
// Endpoints para el frontend
// ============================================================
function getAlertasStock(params) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('ALERTAS_STOCK');
  if (!sheet) return { ok: true, data: [] };
  var rows = _sheetToObjects(sheet);
  var soloPendientes = params && (params.soloPendientes === true || params.soloPendientes === 'true');
  if (soloPendientes) {
    rows = rows.filter(function(r) { return String(r.revisado || '').toUpperCase() !== 'SI'; });
  }
  // Más recientes primero
  rows.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return { ok: true, data: rows };
}

function marcarAlertaRevisada(params) {
  var idAlerta = String(params.idAlerta || '');
  if (!idAlerta) return { ok: false, error: 'idAlerta requerido' };
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('ALERTAS_STOCK');
  if (!sheet) return { ok: false, error: 'Hoja ALERTAS_STOCK no existe' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var idxId   = hdrs.indexOf('idAlerta');
  var idxRev  = hdrs.indexOf('revisado');
  var idxFecR = hdrs.indexOf('fechaRevision');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idAlerta) {
      sheet.getRange(i + 1, idxRev + 1).setValue('SI');
      if (idxFecR >= 0) sheet.getRange(i + 1, idxFecR + 1).setValue(new Date());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Alerta no encontrada' };
}

// ============================================================
// Limpieza one-shot de duplicados en GUIA_DETALLE.
// Recorre todas las guías y, si hay 2+ detalles con mismo codigoProducto
// (no anulados), conserva el PRIMERO (suma sus cantidades) y borra el resto.
// Útil para reparar duplicados históricos. Idempotente: si no hay
// duplicados, no hace nada. Devuelve resumen.
// ============================================================
function limpiarDuplicadosGuiaDetalle() {
  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId    = hdrs.indexOf('idDetalle');
  var idxIdG   = hdrs.indexOf('idGuia');
  var idxCb    = hdrs.indexOf('codigoProducto');
  var idxRec   = hdrs.indexOf('cantidadRecibida');
  var idxObs   = hdrs.indexOf('observacion');

  // Mapa: clave (idGuia + cb) → primera fila vista
  var primeros = {};
  var aBorrar  = []; // lista de filas (1-based) a borrar de abajo hacia arriba
  var sumar    = []; // [{filaPrimero, cantidadAdicional}]

  for (var i = 1; i < data.length; i++) {
    var idG = String(data[i][idxIdG] || '');
    var cb  = String(data[i][idxCb]  || '').toUpperCase();
    var obs = String(data[i][idxObs] || '').toUpperCase();
    if (!idG || !cb || obs === 'ANULADO') continue;
    var key = idG + '|' + cb;
    if (primeros[key] === undefined) {
      primeros[key] = i;  // primera vez: conservar
    } else {
      // Duplicado: sumar al primero, borrar este
      var qtyDup = parseFloat(data[i][idxRec]) || 0;
      sumar.push({ filaPrimero: primeros[key] + 1, cantidad: qtyDup });
      aBorrar.push(i + 1);
    }
  }

  // Aplicar sumas primero (antes de borrar para preservar índices)
  sumar.forEach(function(s) {
    var celda  = sheet.getRange(s.filaPrimero, idxRec + 1);
    var qtyAct = parseFloat(celda.getValue()) || 0;
    celda.setValue(qtyAct + s.cantidad);
  });

  // Borrar de abajo hacia arriba para preservar índices
  aBorrar.sort(function(a, b) { return b - a; }).forEach(function(f) {
    sheet.deleteRow(f);
  });

  Logger.log('limpiarDuplicadosGuiaDetalle: ' + aBorrar.length + ' duplicados eliminados, ' +
             sumar.length + ' sumas aplicadas');
  return { ok: true, data: { duplicadosEliminados: aBorrar.length, sumasAplicadas: sumar.length } };
}

// ============================================================
// Setup de triggers — ejecutar UNA vez desde el editor de Apps Script
// ============================================================
function setupTriggersAuditoria() {
  // Borrar triggers existentes con esos handlers para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var h = t.getHandlerFunction();
    if (h === 'cerrarGuiasAbiertasGlobal' || h === 'auditarStockGlobal') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 21:00 → cierre forzado
  ScriptApp.newTrigger('cerrarGuiasAbiertasGlobal')
    .timeBased()
    .atHour(21)
    .everyDays(1)
    .create();

  // 22:00 → auditoría de stock
  ScriptApp.newTrigger('auditarStockGlobal')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .create();

  return { ok: true, mensaje: 'Triggers configurados: 21:00 cierre + 22:00 auditoría' };
}
