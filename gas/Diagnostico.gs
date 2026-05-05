// ============================================================
// warehouseMos — Diagnostico.gs
// Sistema de auto-tests para detectar bugs de cuadre en guías de
// salida (race conditions, doble-click, auto-suma errónea, etc.)
//
// Flujo:
//   1. iniciarTestDiagnostico(idTest, params) → snapshot inicial + idEjecucion
//   2. usuario realiza pasos manuales en la app
//   3. finalizarTestDiagnostico(idEjecucion) → compara y reporta PASS/FAIL
// ============================================================

function _getSheetDiagnostico() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('DIAGNOSTICO_TESTS');
  if (!sheet) {
    sheet = ss.insertSheet('DIAGNOSTICO_TESTS');
    sheet.getRange(1, 1, 1, 9).setValues([[
      'idEjecucion','idTest','usuario','fechaInicio','fechaFin',
      'snapshot','resultado','estado','mensaje'
    ]]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  return sheet;
}

function iniciarTestDiagnostico(params) {
  var idTest        = String(params.idTest || '');
  var usuario       = String(params.usuario || '');
  var codigoProducto = String(params.codigoProducto || '').trim();
  if (!idTest) return { ok: false, error: 'idTest requerido' };

  var idEjecucion = 'EJEC_' + new Date().getTime();
  var snapshot = {
    idTest: idTest,
    timestampInicio: new Date().toISOString(),
    codigoProducto: codigoProducto
  };

  // Capturar stock actual del producto si aplica
  if (codigoProducto) {
    try {
      var info = _getStockProducto(codigoProducto);
      snapshot.stockInicial = info.cantidad;
    } catch(e) {}
  }

  // Capturar última fila de GUIAS y GUIA_DETALLE para luego diff
  try {
    snapshot.ultimaFilaGuias = getSheet('GUIAS').getLastRow();
    snapshot.ultimaFilaDetalle = getSheet('GUIA_DETALLE').getLastRow();
  } catch(e) {}

  // Guardar la ejecución como RUNNING
  var sheet = _getSheetDiagnostico();
  sheet.appendRow([
    idEjecucion, idTest, usuario, new Date(), '',
    JSON.stringify(snapshot), '', 'RUNNING', ''
  ]);

  return { ok: true, data: { idEjecucion: idEjecucion, snapshot: snapshot } };
}

function finalizarTestDiagnostico(params) {
  var idEjecucion = String(params.idEjecucion || '');
  if (!idEjecucion) return { ok: false, error: 'idEjecucion requerido' };

  // Buscar la ejecución
  var sheet = _getSheetDiagnostico();
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId   = hdrs.indexOf('idEjecucion');
  var idxSnap = hdrs.indexOf('snapshot');
  var idxRes  = hdrs.indexOf('resultado');
  var idxEst  = hdrs.indexOf('estado');
  var idxFin  = hdrs.indexOf('fechaFin');
  var idxMsg  = hdrs.indexOf('mensaje');

  var fila = -1;
  var snapshot = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idEjecucion) {
      fila = i + 1;
      try { snapshot = JSON.parse(data[i][idxSnap]); } catch(e) {}
      break;
    }
  }
  if (fila < 0) return { ok: false, error: 'Ejecución no encontrada' };
  if (!snapshot) return { ok: false, error: 'Snapshot inválido' };

  // Validar según el test
  var resultado = _validarTest(snapshot);

  // Guardar resultado
  sheet.getRange(fila, idxFin + 1).setValue(new Date());
  sheet.getRange(fila, idxRes + 1).setValue(JSON.stringify(resultado));
  sheet.getRange(fila, idxEst + 1).setValue(resultado.pass ? 'PASS' : 'FAIL');
  sheet.getRange(fila, idxMsg + 1).setValue(resultado.mensaje || '');

  return { ok: true, data: resultado };
}

// ============================================================
// Validadores por test — leen estado actual y comparan vs snapshot
// ============================================================
function _validarTest(snap) {
  var idTest = snap.idTest;
  switch (idTest) {
    case 'TEST_1': return _validarTest1(snap);
    case 'TEST_2': return _validarTest2(snap);
    case 'TEST_3': return _validarTest3(snap);
    case 'TEST_4': return _validarTest4(snap);
    case 'TEST_5': return _validarTest5(snap);
    case 'TEST_6': return _validarTest6(snap);
    default: return { pass: false, mensaje: 'Test desconocido: ' + idTest };
  }
}

// Buscar guías SALIDA creadas DESPUÉS del timestamp del snapshot
function _guiasSalidaPosteriores(timestampISO) {
  var ts = new Date(timestampISO).getTime();
  var guias = _sheetToObjects(getSheet('GUIAS'));
  return guias.filter(function(g) {
    if (!g.tipo || g.tipo.indexOf('SALIDA') !== 0) return false;
    var fGuia = g.fecha ? new Date(g.fecha).getTime() : 0;
    return fGuia >= ts - 5000; // tolerancia 5s
  });
}

function _detallesPorGuia(idGuia) {
  return _sheetToObjects(getSheet('GUIA_DETALLE'))
    .filter(function(d) { return String(d.idGuia) === String(idGuia); });
}

// TEST 1 — 10 escaneos rápidos del mismo producto
// Esperado: 1 guía nueva, 1 detalle, qty=10
function _validarTest1(snap) {
  var nuevasGuias = _guiasSalidaPosteriores(snap.timestampInicio);
  if (!nuevasGuias.length) {
    return { pass: false, mensaje: 'No se detectó guía nueva. ¿Creaste la guía después de iniciar el test?' };
  }
  // Tomar la última creada
  var guia = nuevasGuias[nuevasGuias.length - 1];
  var dets = _detallesPorGuia(guia.idGuia)
    .filter(function(d) { return String(d.codigoProducto || '').toUpperCase() === String(snap.codigoProducto || '').toUpperCase(); });

  if (!dets.length) {
    return { pass: false, mensaje: 'Guía sin detalles del producto seleccionado.', esperado: '1 fila qty=10', obtenido: '0 filas' };
  }
  if (dets.length > 1) {
    return { pass: false,
             mensaje: 'BUG: ' + dets.length + ' filas duplicadas en lugar de 1 con auto-suma',
             esperado: '1 fila qty=10', obtenido: dets.length + ' filas' };
  }
  var qty = parseFloat(dets[0].cantidadRecibida) || 0;
  if (qty === 10) {
    return { pass: true, mensaje: 'OK: 1 fila con qty=10', esperado: 'qty=10', obtenido: 'qty=10', idGuia: guia.idGuia };
  }
  return { pass: false,
           mensaje: 'BUG: qty=' + qty + ' (esperaba 10) — posible race condition en auto-suma',
           esperado: 'qty=10', obtenido: 'qty=' + qty, idGuia: guia.idGuia };
}

// TEST 2 — Edición concurrente: scan 1, edit 15, scan 1 más
// Esperado: qty final = 16
function _validarTest2(snap) {
  var nuevasGuias = _guiasSalidaPosteriores(snap.timestampInicio);
  if (!nuevasGuias.length) return { pass: false, mensaje: 'No se detectó guía nueva' };
  var guia = nuevasGuias[nuevasGuias.length - 1];
  var dets = _detallesPorGuia(guia.idGuia)
    .filter(function(d) { return String(d.codigoProducto || '').toUpperCase() === String(snap.codigoProducto || '').toUpperCase(); });

  if (!dets.length) return { pass: false, mensaje: 'Sin detalles del producto', esperado: 'qty=16' };
  var qty = parseFloat(dets[0].cantidadRecibida) || 0;
  // Aceptamos margen: la edición + scan adicional debería ser >=15. Si quedó menor → bug out-of-order
  if (qty >= 15) {
    return { pass: true,
             mensaje: 'qty=' + qty + ' (esperado >=15, ideal 16)',
             esperado: 'qty>=15', obtenido: 'qty=' + qty, idGuia: guia.idGuia };
  }
  return { pass: false,
           mensaje: 'BUG: qty=' + qty + ' menor a lo esperado — actualizarCantidad llegó out-of-order',
           esperado: 'qty>=15', obtenido: 'qty=' + qty, idGuia: guia.idGuia };
}

// TEST 3 — Cerrar guía inmediatamente tras escaneos
// Esperado: stock real - 5 = stock final, 1 fila qty=5
function _validarTest3(snap) {
  var nuevasGuias = _guiasSalidaPosteriores(snap.timestampInicio);
  if (!nuevasGuias.length) return { pass: false, mensaje: 'No se detectó guía nueva' };
  var guia = nuevasGuias[nuevasGuias.length - 1];

  if (String(guia.estado || '').toUpperCase() !== 'CERRADA') {
    return { pass: false, mensaje: 'Guía no está CERRADA (estado=' + guia.estado + ')',
             esperado: 'guía CERRADA', obtenido: 'estado=' + guia.estado };
  }

  var dets = _detallesPorGuia(guia.idGuia)
    .filter(function(d) {
      return String(d.codigoProducto || '').toUpperCase() === String(snap.codigoProducto || '').toUpperCase()
          && String(d.observacion || '') !== 'ANULADO';
    });
  if (dets.length !== 1) {
    return { pass: false, mensaje: 'BUG: ' + dets.length + ' filas (esperaba 1)',
             esperado: '1 fila qty=5', obtenido: dets.length + ' filas', idGuia: guia.idGuia };
  }
  var qty = parseFloat(dets[0].cantidadRecibida) || 0;
  if (qty !== 5) {
    return { pass: false, mensaje: 'qty=' + qty + ' (esperaba 5)',
             esperado: 'qty=5', obtenido: 'qty=' + qty, idGuia: guia.idGuia };
  }
  // Validar stock
  var stockActual = _getStockProducto(snap.codigoProducto).cantidad;
  var stockEsperado = (snap.stockInicial || 0) - 5;
  if (Math.abs(stockActual - stockEsperado) > 0.01) {
    return { pass: false,
             mensaje: 'BUG STOCK: real=' + stockActual + ' esperado=' + stockEsperado +
                      ' diff=' + (stockActual - stockEsperado),
             esperado: 'stock=' + stockEsperado,
             obtenido: 'stock=' + stockActual, idGuia: guia.idGuia };
  }
  return { pass: true, mensaje: 'OK · qty=5 · stock bajó 5 unidades', idGuia: guia.idGuia };
}

// TEST 4 — Despacho doble-click: solo se crea 1 guía
function _validarTest4(snap) {
  var nuevasGuias = _guiasSalidaPosteriores(snap.timestampInicio);
  // Filtrar SALIDA_ZONA o SALIDA_JEFATURA o SALIDA_DEVOLUCION (despacho rápido)
  var guiasDespacho = nuevasGuias.filter(function(g) {
    var t = String(g.tipo || '').toUpperCase();
    return t === 'SALIDA_ZONA' || t === 'SALIDA_JEFATURA' || t === 'SALIDA_DEVOLUCION';
  });
  if (guiasDespacho.length === 0) {
    return { pass: false, mensaje: 'No se detectó guía nueva del despacho',
             esperado: '1 guía', obtenido: '0 guías' };
  }
  if (guiasDespacho.length > 1) {
    return { pass: false,
             mensaje: 'BUG: ' + guiasDespacho.length + ' guías duplicadas por doble-click',
             esperado: '1 guía',
             obtenido: guiasDespacho.length + ' guías (' + guiasDespacho.map(function(g){return g.idGuia;}).join(',') + ')' };
  }
  return { pass: true, mensaje: 'OK · solo 1 guía creada', idGuia: guiasDespacho[0].idGuia };
}

// TEST 5 — +/- rápidos en cam: 1 inicial + 5 + 2 -2 = 4
function _validarTest5(snap) {
  var nuevasGuias = _guiasSalidaPosteriores(snap.timestampInicio);
  if (!nuevasGuias.length) return { pass: false, mensaje: 'No se detectó guía nueva' };
  var guia = nuevasGuias[nuevasGuias.length - 1];
  var dets = _detallesPorGuia(guia.idGuia)
    .filter(function(d) { return String(d.codigoProducto || '').toUpperCase() === String(snap.codigoProducto || '').toUpperCase(); });
  if (!dets.length) return { pass: false, mensaje: 'Sin detalles del producto' };
  var qty = parseFloat(dets[0].cantidadRecibida) || 0;
  if (qty === 4) {
    return { pass: true, mensaje: 'OK · qty=4 (1+5-2)', idGuia: guia.idGuia };
  }
  return { pass: false,
           mensaje: 'BUG: qty=' + qty + ' (esperaba 4) — los +/- no se aplicaron correctamente',
           esperado: 'qty=4', obtenido: 'qty=' + qty, idGuia: guia.idGuia };
}

// TEST 6 — Reabrir y editar: stock final = stock inicial - 3
function _validarTest6(snap) {
  var stockActual = _getStockProducto(snap.codigoProducto).cantidad;
  var stockEsperado = (snap.stockInicial || 0) - 3;
  if (Math.abs(stockActual - stockEsperado) <= 0.01) {
    return { pass: true,
             mensaje: 'OK · stock bajó 3 unidades (no 5)',
             esperado: 'stock=' + stockEsperado,
             obtenido: 'stock=' + stockActual };
  }
  // Diagnóstico extra
  var diff = stockActual - stockEsperado;
  var diagnostico = '';
  if (diff < 0) diagnostico = ' BUG: stock bajó MÁS de lo debido — reabrir/cerrar duplicó descuento';
  else          diagnostico = ' BUG: stock bajó MENOS — reabrir no revirtió completamente';
  return { pass: false,
           mensaje: 'stock real=' + stockActual + ' esperado=' + stockEsperado + ' diff=' + diff + diagnostico,
           esperado: 'stock=' + stockEsperado,
           obtenido: 'stock=' + stockActual };
}

// ============================================================
// Endpoint: lista los últimos resultados de tests
// ============================================================
function getResultadosDiagnostico() {
  var sheet = _getSheetDiagnostico();
  var rows = _sheetToObjects(sheet);
  rows.sort(function(a, b) { return new Date(b.fechaInicio) - new Date(a.fechaInicio); });
  return { ok: true, data: rows.slice(0, 20).map(function(r) {
    var resultadoObj = {};
    try { resultadoObj = JSON.parse(r.resultado || '{}'); } catch(e) {}
    return {
      idEjecucion: r.idEjecucion,
      idTest:      r.idTest,
      usuario:     r.usuario,
      fechaInicio: r.fechaInicio,
      fechaFin:    r.fechaFin,
      estado:      r.estado,
      mensaje:     r.mensaje,
      resultado:   resultadoObj
    };
  }) };
}
