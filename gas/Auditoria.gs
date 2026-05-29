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
      // skipMosSync: evita el UrlFetchApp por cada detalle (lento si hay muchos
      // productos, causa timeout >6min). El sync interactivo de productos
      // proveedor sigue funcionando cuando el usuario cierra guías manualmente.
      var res = cerrarGuia(idGuia, 'sistema-cierre-21h', null, { skipMosSync: true });
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

// ============================================================
// aceptarTeoricoAlerta — corrección one-click: crea AJUSTE para
// que stock real = stock teórico, y marca la alerta como revisada.
// Útil para diferencias chicas donde no hace falta conteo físico.
// ============================================================
function aceptarTeoricoAlerta(params) {
  var idAlerta = String(params.idAlerta || '');
  if (!idAlerta) return { ok: false, error: 'idAlerta requerido' };
  var usuario = String(params.usuario || 'sistema');

  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('ALERTAS_STOCK');
  if (!sheet) return { ok: false, error: 'Hoja ALERTAS_STOCK no existe' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var idxId    = hdrs.indexOf('idAlerta');
  var idxCb    = hdrs.indexOf('codigoProducto');
  var idxReal  = hdrs.indexOf('stockReal');
  var idxTeor  = hdrs.indexOf('stockTeorico');
  var idxRev   = hdrs.indexOf('revisado');
  var idxFecR  = hdrs.indexOf('fechaRevision');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) !== idAlerta) continue;
    var cb       = String(data[i][idxCb]);
    var stockReal = parseFloat(data[i][idxReal]) || 0;
    var stockTeor = parseFloat(data[i][idxTeor]) || 0;
    var diff      = stockTeor - stockReal;  // ajuste para que real → teórico
    if (Math.abs(diff) <= 0.5) {
      // Ya están iguales; solo marcar revisada
      sheet.getRange(i + 1, idxRev + 1).setValue('SI');
      if (idxFecR >= 0) sheet.getRange(i + 1, idxFecR + 1).setValue(new Date());
      return { ok: true, data: { idAlerta: idAlerta, ajusteAplicado: 0 } };
    }
    // Crear AJUSTE INC o DEC para corregir
    var resAj = crearAjuste({
      codigoProducto: cb,
      tipoAjuste:     diff > 0 ? 'INC' : 'DEC',
      cantidadAjuste: Math.abs(diff),
      motivo:         'Aceptar teórico (alerta cuadre stock)',
      usuario:        usuario
    });
    if (!resAj.ok) return { ok: false, error: 'Error creando ajuste: ' + resAj.error };
    // Marcar alerta como revisada
    sheet.getRange(i + 1, idxRev + 1).setValue('SI');
    if (idxFecR >= 0) sheet.getRange(i + 1, idxFecR + 1).setValue(new Date());
    return { ok: true, data: { idAlerta: idAlerta, ajusteAplicado: diff, idAjuste: resAj.data.idAjuste } };
  }
  return { ok: false, error: 'Alerta no encontrada' };
}

// Endpoint para consultar STOCK_MOVIMIENTOS de un producto específico.
// Útil para diagnosticar "¿quién/cuándo movió este stock?".
function getStockMovimientos(params) {
  var codigo = String(params.codigoProducto || '').trim();
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('STOCK_MOVIMIENTOS');
  if (!sheet) return { ok: true, data: [] };
  var rows = _sheetToObjects(sheet);
  if (codigo) rows = rows.filter(function(r) { return String(r.codigoProducto) === codigo; });
  rows.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  if (params.limit) rows = rows.slice(0, parseInt(params.limit));
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
// Limpieza one-shot de códigos numéricos sin ceros (causados por race
// conditions en agregarDetalleGuia antes del LockService).
//
// Si encuentra fila con codigoProducto = "27" y existe producto real "00027",
// reescribe la fila como "00027" preservando el formato texto. Después la
// otra limpieza (limpiarDuplicadosGuiaDetalle) puede mergear cantidades.
// ============================================================
function repararCodigosNumericos() {
  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxCb = hdrs.indexOf('codigoProducto');

  // Construir set de códigos VÁLIDOS desde PRODUCTOS_MASTER (con ceros preservados)
  var prodsValidos = {};
  try {
    var prods = _sheetToObjects(getProductosSheet());
    prods.forEach(function(p) {
      var cb = String(p.codigoBarra || '').trim();
      if (cb) prodsValidos[cb] = true;
    });
    // También incluir equivalencias
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      _sheetToObjects(equivSheet).forEach(function(e) {
        var cb = String(e.codigoBarra || '').trim();
        if (cb) prodsValidos[cb] = true;
      });
    }
  } catch(e) {}

  var reparadas = [];
  for (var i = 1; i < data.length; i++) {
    var cbActual = String(data[i][idxCb] || '').trim();
    // Solo procesar códigos puramente numéricos cortos (probablemente perdieron ceros)
    if (!/^\d+$/.test(cbActual) || cbActual.length >= 12) continue;
    // Si el código ya existe tal cual en productos válidos, no es bug
    if (prodsValidos[cbActual]) continue;

    // Buscar si existe versión con ceros: 00027, 0027, 027
    var candidatos = [
      cbActual.padStart(5, '0'),
      cbActual.padStart(4, '0'),
      cbActual.padStart(6, '0')
    ];
    var match = null;
    for (var c = 0; c < candidatos.length; c++) {
      if (prodsValidos[candidatos[c]]) { match = candidatos[c]; break; }
    }
    if (!match) continue;

    // Reescribir con formato texto + valor correcto
    sheet.getRange(i + 1, idxCb + 1).setNumberFormat('@').setValue(match);
    reparadas.push({ fila: i + 1, antes: cbActual, ahora: match });
  }
  Logger.log('repararCodigosNumericos: ' + reparadas.length + ' filas reparadas');
  return { ok: true, data: { reparadas: reparadas.length, detalles: reparadas } };
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

// ════════════════════════════════════════════════════════════════════
// [v2.13.55] Reconciliación masiva del stock — auto-corrige
// todas las discrepancias mayores a un umbral mediante AJUSTES
// auditados. Cada corrección queda como fila en hoja AJUSTES con
// tipo INC/DEC y motivo trazable.
//
// Política confirmada:
//  - El stock SIEMPRE debe igualar el teórico (= Σ AJUSTES + Σ INGRESOS − Σ SALIDAS)
//  - Si difiere → se crea AJUSTE explícito por la diferencia
//  - cada ajuste es auditado: usuario, motivo, fecha
//  - LockService para evitar 2 ejecuciones concurrentes
// ════════════════════════════════════════════════════════════════════
function reconciliarStockMasivo(params) {
  params = params || {};
  var maxDiffAuto = parseFloat(params.maxDiffAuto || 0);  // 0 = corregir todas. Si >0, solo corrige diff<=N
  var dryRun      = params.dryRun === true || params.dryRun === 'true';
  var autorizadoPor = String(params.autorizadoPor || 'sistema-reconciliacion');
  var motivoLabel   = String(params.motivo || 'Reconciliación masiva auto');

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(eL) {
    return { ok: false, error: 'Lock saturado, retry: ' + eL.message };
  }
  try {
    // 1) Re-correr auditoría para tener datos frescos
    var auditRes = auditarStockGlobal();
    if (!auditRes.ok) return auditRes;

    // 2) Leer alertas pendientes
    var ss = SpreadsheetApp.openById(SS_ID);
    var alertSh = ss.getSheetByName('ALERTAS_STOCK');
    if (!alertSh) return { ok: true, data: { corregidas: 0, omitidas: 0, errores: 0 } };
    var alerts = _sheetToObjects(alertSh).filter(function(r) {
      return String(r.revisado || '').toUpperCase() !== 'SI';
    });

    var corregidas = 0, omitidas = 0, errores = 0;
    var detalles = [];

    alerts.forEach(function(a) {
      var diff = parseFloat(a.diferencia) || 0;
      var absDiff = Math.abs(diff);
      // diff = stockReal − stockTeorico
      // Para igualar real → teórico, hay que aplicar −diff
      var ajusteCantidad = -diff;
      // Si umbral activo y la diff es mayor → omitir (requiere revisión manual)
      if (maxDiffAuto > 0 && absDiff > maxDiffAuto) {
        omitidas++;
        detalles.push({ codigoProducto: a.codigoProducto, diff: diff, accion: 'OMITIDA_UMBRAL' });
        return;
      }
      if (dryRun) {
        detalles.push({ codigoProducto: a.codigoProducto, diff: diff, accion: 'DRY_RUN' });
        corregidas++;
        return;
      }
      try {
        // crearAjuste internamente llama _actualizarStock con el delta correcto
        var resAj = crearAjuste({
          codigoProducto: String(a.codigoProducto),
          tipoAjuste:     ajusteCantidad > 0 ? 'INC' : 'DEC',
          cantidadAjuste: Math.abs(ajusteCantidad),
          motivo:         motivoLabel + ' · diff=' + diff.toFixed(2) +
                          ' · real=' + a.stockReal + ' · teo=' + a.stockTeorico,
          usuario:        autorizadoPor
        });
        if (resAj && resAj.ok) {
          corregidas++;
          detalles.push({ codigoProducto: a.codigoProducto, diff: diff, accion: 'CORREGIDO', idAjuste: resAj.data.idAjuste });
          // Marcar alerta revisada
          var alertData = alertSh.getDataRange().getValues();
          var alertHdr = alertData[0];
          var idxAId  = alertHdr.indexOf('idAlerta');
          var idxARev = alertHdr.indexOf('revisado');
          var idxAFR  = alertHdr.indexOf('fechaRevision');
          for (var i = 1; i < alertData.length; i++) {
            if (String(alertData[i][idxAId]) === String(a.idAlerta)) {
              alertSh.getRange(i + 1, idxARev + 1).setValue('SI');
              if (idxAFR >= 0) alertSh.getRange(i + 1, idxAFR + 1).setValue(new Date());
              break;
            }
          }
        } else {
          errores++;
          detalles.push({ codigoProducto: a.codigoProducto, diff: diff, accion: 'ERROR', error: (resAj && resAj.error) || 'sin info' });
        }
      } catch(e) {
        errores++;
        detalles.push({ codigoProducto: a.codigoProducto, diff: diff, accion: 'EXCEPCION', error: e.message });
      }
    });

    Logger.log('[reconciliarStockMasivo] corregidas=' + corregidas + ' omitidas=' + omitidas + ' errores=' + errores);
    return { ok: true, data: { corregidas: corregidas, omitidas: omitidas, errores: errores, dryRun: dryRun, detalles: detalles } };
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// [v2.13.55] Reconciliar UN solo producto por codigoBarra (canónico o equivalente).
// Útil para tap "Reconciliar" en panel admin MOS por fila.
function reconciliarStockProducto(params) {
  var codigoBarra = String(params.codigoBarra || params.codigoProducto || '').trim();
  if (!codigoBarra) return { ok: false, error: 'codigoBarra requerido' };
  var autorizadoPor = String(params.autorizadoPor || 'admin-mos');

  // Reusamos la matemática de auditarStockGlobal pero solo para este código
  var stockSh = getSheet('STOCK');
  var stockData = _sheetToObjects(stockSh);
  var miFila = null;
  for (var i = 0; i < stockData.length; i++) {
    if (String(stockData[i].codigoProducto).trim() === codigoBarra) {
      miFila = stockData[i];
      break;
    }
  }
  var real = miFila ? (parseFloat(miFila.cantidadDisponible) || 0) : 0;

  // Sumar movimientos
  var teorico = 0;
  try {
    var guiaMap = {};
    _sheetToObjects(getSheet('GUIAS')).forEach(function(g) {
      var est = String(g.estado || '').toUpperCase();
      if (est === 'CERRADA' || est === 'AUTOCERRADA') guiaMap[g.idGuia] = String(g.tipo || '').toUpperCase();
    });
    _sheetToObjects(getSheet('GUIA_DETALLE')).forEach(function(d) {
      if (String(d.codigoProducto).trim() !== codigoBarra) return;
      if (String(d.observacion || '').toUpperCase() === 'ANULADO') return;
      var tipoG = guiaMap[d.idGuia];
      if (!tipoG) return;
      var cant = Math.abs(parseFloat(d.cantidadRecibida || d.cantidadReal || d.cantidadEsperada || 0));
      if (cant <= 0) return;
      if (tipoG.indexOf('INGRESO') === 0) teorico += cant;
      else if (tipoG.indexOf('SALIDA') === 0) teorico -= cant;
    });
    var ajSh = getSheet('AJUSTES');
    if (ajSh) {
      _sheetToObjects(ajSh).forEach(function(a) {
        if (String(a.codigoProducto).trim() !== codigoBarra) return;
        var cant = Math.abs(parseFloat(a.cantidadAjuste || 0));
        var tipoA = String(a.tipoAjuste || '').toUpperCase();
        if (cant <= 0) return;
        if (tipoA === 'INC' || tipoA === 'INI') teorico += cant;
        else if (tipoA === 'DEC') teorico -= cant;
      });
    }
  } catch(e) {
    return { ok: false, error: 'Error calculando teórico: ' + e.message };
  }

  var diff = real - teorico;
  if (Math.abs(diff) <= 0.5) {
    return { ok: true, data: { codigoBarra: codigoBarra, real: real, teorico: teorico, diff: diff, accion: 'YA_CUADRA' } };
  }
  // Aplicar ajuste para igualar real → teórico
  var ajusteCantidad = -diff;
  var resAj = crearAjuste({
    codigoProducto: codigoBarra,
    tipoAjuste:     ajusteCantidad > 0 ? 'INC' : 'DEC',
    cantidadAjuste: Math.abs(ajusteCantidad),
    motivo:         'Reconciliación manual · diff=' + diff.toFixed(2) + ' · real=' + real + ' · teo=' + teorico,
    usuario:        autorizadoPor
  });
  if (!resAj || !resAj.ok) return { ok: false, error: 'Error creando ajuste: ' + (resAj && resAj.error) };
  return {
    ok: true,
    data: {
      codigoBarra: codigoBarra,
      real: real, teorico: teorico, diff: diff,
      ajusteAplicado: ajusteCantidad,
      idAjuste: resAj.data.idAjuste,
      accion: 'CORREGIDO'
    }
  };
}

// [v2.13.55] Cron 21:00 mejorado: LockService + flag idempotente diario.
// Garantiza que aunque se dispare 2x (rare), solo aplica una vez por día.
function cerrarGuiasAbiertasGlobalSafe() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) {
    Logger.log('[cron21] lock saturado: ' + e.message);
    return { ok: false, error: 'lock' };
  }
  try {
    var hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var prop = PropertiesService.getScriptProperties();
    var ultimo = prop.getProperty('cron_cierre_diario_ts') || '';
    if (ultimo === hoy) {
      Logger.log('[cron21] ya corrió hoy ' + hoy);
      return { ok: true, data: { skipped: true, fecha: hoy } };
    }
    var res = cerrarGuiasAbiertasGlobal();
    prop.setProperty('cron_cierre_diario_ts', hoy);
    return res;
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}
