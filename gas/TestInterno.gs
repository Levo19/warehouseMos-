// ============================================================
// warehouseMos — TestInterno.gs
// Suite de tests automatizados que valida los flujos críticos:
//   - auto-suma de detalles
//   - idempotencia de cierre/anulación/reabrir
//   - edición de cantidad
//   - preservación de ceros
//
// IMPORTANTE: el test crea una guía temporal SALIDA_ZONA, ejecuta
// operaciones reales sobre ella, y al final BORRA físicamente todas
// las filas creadas (GUIAS, GUIA_DETALLE, STOCK_MOVIMIENTOS) y revierte
// el stock al valor original. NO deja rastro.
//
// Uso desde editor:
//   runInternalTests({ codigoBarra: 'WHCAABCA001KG' })
// ============================================================

function runInternalTests(params) {
  var codigoBarra = String((params && params.codigoBarra) || '').trim();
  if (!codigoBarra) return { ok: false, error: 'codigoBarra requerido' };

  var resultados   = [];
  var idGuiaTest   = null;
  var stockOriginal = 0;

  function _add(test, pass, detalle, esperado, obtenido) {
    resultados.push({ test: test, pass: !!pass, detalle: detalle || '',
                      esperado: esperado, obtenido: obtenido });
  }

  try {
    // ── Snapshot inicial ──────────────────────────────────────
    var infoStock = _getStockProducto(codigoBarra);
    if (infoStock.fila < 0) {
      return { ok: false, error: 'codigoBarra no encontrado en STOCK: ' + codigoBarra };
    }
    stockOriginal = infoStock.cantidad;
    _add('SETUP', true, 'Stock inicial registrado: ' + stockOriginal);

    // ── Crear guía test ───────────────────────────────────────
    var resG = crearGuia({
      tipo:        'SALIDA_ZONA',
      idZona:      'TEST',
      usuario:     'sistema-test',
      comentario:  '[TEST INTERNO] eliminar al terminar'
    });
    if (!resG.ok) throw new Error('crearGuia: ' + resG.error);
    idGuiaTest = resG.data.idGuia;
    _add('TEST 0: crear guía', true, 'idGuia=' + idGuiaTest);

    // ── TEST 1: auto-suma — 10 llamadas secuenciales del mismo cb ──
    var i;
    for (i = 0; i < 10; i++) {
      var rd = agregarDetalleGuia({
        idGuia:           idGuiaTest,
        codigoProducto:   codigoBarra,
        cantidadEsperada: 1,
        cantidadRecibida: 1
      });
      if (!rd.ok) {
        _add('TEST 1: auto-suma', false, 'falló iter ' + i + ': ' + rd.error);
        throw new Error('aborta');
      }
    }
    var detsT1 = _detallesPorGuiaCb(idGuiaTest, codigoBarra);
    _add('TEST 1: auto-suma 10 escaneos',
         detsT1.length === 1 && parseFloat(detsT1[0].cantidadRecibida) === 10,
         'filas=' + detsT1.length + ' qty=' + (detsT1[0] && detsT1[0].cantidadRecibida),
         '1 fila qty=10', detsT1.length + ' filas qty=' + (detsT1[0] && detsT1[0].cantidadRecibida));

    var idDetTest = detsT1[0] && detsT1[0].idDetalle;

    // ── TEST 2: editar cantidad en GUÍA ABIERTA — stock NO debe cambiar ──
    var stockAntes2 = _getStockProducto(codigoBarra).cantidad;
    var rEd = actualizarCantidadDetalle({ idDetalle: idDetTest, cantidadRecibida: 15 });
    var stockDesp2 = _getStockProducto(codigoBarra).cantidad;
    _add('TEST 2: editar qty en abierta — stock intacto',
         rEd.ok && stockAntes2 === stockDesp2,
         'antes=' + stockAntes2 + ' después=' + stockDesp2,
         'sin cambio en stock', 'cambió en ' + (stockDesp2 - stockAntes2));

    // ── TEST 3: cerrarGuia — stock baja exactamente la cantidad ──
    var stockAntes3 = _getStockProducto(codigoBarra).cantidad;
    var rC = cerrarGuia(idGuiaTest, 'sistema-test', null);
    var stockDesp3 = _getStockProducto(codigoBarra).cantidad;
    var diff3 = stockDesp3 - stockAntes3;
    _add('TEST 3: cerrar guía — stock baja 15',
         rC.ok && Math.abs(diff3 - (-15)) <= 0.01,
         'diff=' + diff3,
         '-15', diff3);

    // ── TEST 4: cerrarGuia idempotente — segunda llamada NO debe restar más ──
    var stockAntes4 = _getStockProducto(codigoBarra).cantidad;
    var rC2 = cerrarGuia(idGuiaTest, 'sistema-test', null);
    var stockDesp4 = _getStockProducto(codigoBarra).cantidad;
    _add('TEST 4: cerrar idempotente',
         rC2.ok && stockAntes4 === stockDesp4 && (rC2.data && rC2.data.yaCerrada === true),
         'yaCerrada=' + (rC2.data && rC2.data.yaCerrada) + ' diff=' + (stockDesp4 - stockAntes4),
         'yaCerrada=true, diff=0',
         'yaCerrada=' + (rC2.data && rC2.data.yaCerrada) + ', diff=' + (stockDesp4 - stockAntes4));

    // ── TEST 5: reabrirGuia — stock vuelve al valor pre-cierre ──
    var stockAntes5 = _getStockProducto(codigoBarra).cantidad;
    var rR = reabrirGuia({ idGuia: idGuiaTest });
    var stockDesp5 = _getStockProducto(codigoBarra).cantidad;
    var diff5 = stockDesp5 - stockAntes5;
    _add('TEST 5: reabrir — stock revertido +15',
         rR.ok && Math.abs(diff5 - 15) <= 0.01,
         'diff=' + diff5,
         '+15', diff5);

    // ── TEST 6: reabrir idempotente — segunda llamada NO debe sumar más ──
    var stockAntes6 = _getStockProducto(codigoBarra).cantidad;
    var rR2 = reabrirGuia({ idGuia: idGuiaTest });
    var stockDesp6 = _getStockProducto(codigoBarra).cantidad;
    _add('TEST 6: reabrir idempotente',
         rR2.ok && stockAntes6 === stockDesp6 && (rR2.data && rR2.yaAbierta === true || rR2.yaAbierta),
         'yaAbierta=' + (rR2.yaAbierta) + ' diff=' + (stockDesp6 - stockAntes6),
         'yaAbierta=true, diff=0',
         'yaAbierta=' + (rR2.yaAbierta) + ', diff=' + (stockDesp6 - stockAntes6));

    // ── TEST 7: cerrar otra vez tras reabrir + edit a 7 → stock debe bajar 7 ──
    var rEd2 = actualizarCantidadDetalle({ idDetalle: idDetTest, cantidadRecibida: 7 });
    var stockAntes7 = _getStockProducto(codigoBarra).cantidad;
    var rC3 = cerrarGuia(idGuiaTest, 'sistema-test', null);
    var stockDesp7 = _getStockProducto(codigoBarra).cantidad;
    var diff7 = stockDesp7 - stockAntes7;
    _add('TEST 7: cerrar tras reabrir+edit a 7',
         rC3.ok && Math.abs(diff7 - (-7)) <= 0.01,
         'edit→7, cerró, diff=' + diff7,
         '-7', diff7);

    // ── TEST 8: anular detalle en guía cerrada — stock se devuelve ──
    var stockAntes8 = _getStockProducto(codigoBarra).cantidad;
    var rAn = anularDetalle({ idDetalle: idDetTest });
    var stockDesp8 = _getStockProducto(codigoBarra).cantidad;
    var diff8 = stockDesp8 - stockAntes8;
    _add('TEST 8: anular tras cierre — stock devuelve +7',
         rAn.ok && Math.abs(diff8 - 7) <= 0.01,
         'diff=' + diff8,
         '+7', diff8);

    // ── TEST 9: anular idempotente — segunda llamada NO devuelve más ──
    var stockAntes9 = _getStockProducto(codigoBarra).cantidad;
    var rAn2 = anularDetalle({ idDetalle: idDetTest });
    var stockDesp9 = _getStockProducto(codigoBarra).cantidad;
    _add('TEST 9: anular idempotente',
         rAn2.ok && stockAntes9 === stockDesp9 && rAn2.yaAnulado === true,
         'yaAnulado=' + rAn2.yaAnulado + ' diff=' + (stockDesp9 - stockAntes9),
         'yaAnulado=true, diff=0',
         'yaAnulado=' + rAn2.yaAnulado + ', diff=' + (stockDesp9 - stockAntes9));

  } catch(eAll) {
    _add('EXCEPTION', false, eAll.message);
  }

  // ────────────────────────────────────────────────────────────
  // LIMPIEZA — borrar todo rastro y revertir stock a original
  // ────────────────────────────────────────────────────────────
  var stockFinal;
  try {
    if (idGuiaTest) {
      // 1. Si la guía aún está cerrada, reabrir para revertir stock pendiente
      var info = _getGuiaInfo(idGuiaTest);
      if (info && String(info.estado).toUpperCase() === 'CERRADA') {
        try { reabrirGuia({ idGuia: idGuiaTest }); } catch(e) {}
      }
      // 2. Borrar movimientos de STOCK_MOVIMIENTOS con origen=idGuiaTest o idDetTest
      _borrarMovimientosTest(idGuiaTest);
      // 3. Borrar detalles
      _borrarDetallesGuia(idGuiaTest);
      // 4. Borrar la guía
      _borrarFilaGuia(idGuiaTest);
    }

    // 5. Verificar que stock está correcto. Si no, ajustar manualmente.
    stockFinal = _getStockProducto(codigoBarra).cantidad;
    var diffFinal = stockFinal - stockOriginal;
    if (Math.abs(diffFinal) > 0.01) {
      // Forzar regreso al valor original (limpieza segura)
      _actualizarStock(codigoBarra, -diffFinal, {
        tipoOperacion: 'TEST_CLEANUP',
        origen:        idGuiaTest || 'TEST',
        usuario:       'sistema-test',
        observacion:   'reverso forzado tras test'
      });
      // Y borrar ese movimiento también
      _borrarMovimientosTest(idGuiaTest);
      stockFinal = _getStockProducto(codigoBarra).cantidad;
    }
    _add('CLEANUP: stock = original',
         Math.abs(stockFinal - stockOriginal) <= 0.01,
         'original=' + stockOriginal + ' final=' + stockFinal,
         stockOriginal, stockFinal);
  } catch(eC) {
    _add('CLEANUP: error', false, eC.message);
  }

  var totalPass = resultados.filter(function(r){return r.pass;}).length;
  var totalFail = resultados.filter(function(r){return !r.pass;}).length;
  return { ok: true, data: {
    resultados:   resultados,
    pass:         totalPass,
    fail:         totalFail,
    total:        resultados.length,
    stockOriginal: stockOriginal,
    stockFinal:    stockFinal,
    idGuiaBorrada: idGuiaTest
  }};
}

// ============================================================
// Helpers internos — bórrado físico de filas
// ============================================================
function _detallesPorGuiaCb(idGuia, codigoBarra) {
  return _sheetToObjects(getSheet('GUIA_DETALLE'))
    .filter(function(d) {
      return String(d.idGuia) === String(idGuia)
          && String(d.codigoProducto).trim().toUpperCase() === String(codigoBarra).trim().toUpperCase()
          && String(d.observacion || '').toUpperCase() !== 'ANULADO';
    });
}

function _borrarDetallesGuia(idGuia) {
  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxIdG = hdrs.indexOf('idGuia');
  var aBorrar = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdG]) === String(idGuia)) aBorrar.push(i + 1);
  }
  aBorrar.sort(function(a, b) { return b - a; }).forEach(function(f) {
    try { sheet.deleteRow(f); } catch(e) {}
  });
}

function _borrarFilaGuia(idGuia) {
  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId = hdrs.indexOf('idGuia');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      try { sheet.deleteRow(i + 1); return; } catch(e) {}
    }
  }
}

function _borrarMovimientosTest(idGuia) {
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('STOCK_MOVIMIENTOS');
  if (!sheet) return;
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxOrigen = hdrs.indexOf('origen');
  var idxOp     = hdrs.indexOf('tipoOperacion');
  var aBorrar = [];
  for (var i = 1; i < data.length; i++) {
    var ori = String(data[i][idxOrigen] || '');
    var op  = String(data[i][idxOp] || '');
    var esTest = (idGuia && ori === String(idGuia))
              || (op === 'TEST_CLEANUP')
              || /test/i.test(op);
    if (esTest) aBorrar.push(i + 1);
  }
  aBorrar.sort(function(a, b) { return b - a; }).forEach(function(f) {
    try { sheet.deleteRow(f); } catch(e) {}
  });
}
