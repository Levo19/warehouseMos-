// ============================================================
// warehouseMos — Guias.gs
// CRUD guías de ingreso y salida + actualización de stock
// Tipos: INGRESO_PROVEEDOR | INGRESO_JEFATURA |
//        SALIDA_DEVOLUCION | SALIDA_ZONA |
//        SALIDA_JEFATURA   | SALIDA_ENVASADO | SALIDA_MERMA
// ============================================================

function getGuias(params) {
  var sheet = getSheet('GUIAS');
  var rows  = _sheetToObjects(sheet);

  if (params.tipo)    rows = rows.filter(function(r){ return r.tipo === params.tipo; });
  if (params.estado)  rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  if (params.usuario) rows = rows.filter(function(r){ return r.usuario === params.usuario; });
  if (params.limit)   rows = rows.slice(0, parseInt(params.limit));

  return { ok: true, data: rows };
}

function getGuia(idGuia) {
  var guias    = _sheetToObjects(getSheet('GUIAS'));
  var detalles = _sheetToObjects(getSheet('GUIA_DETALLE'));

  var guia = guias.find(function(g){ return g.idGuia === idGuia; });
  if (!guia) return { ok: false, error: 'Guía no encontrada: ' + idGuia };

  guia.detalle = detalles.filter(function(d){ return d.idGuia === idGuia; });

  // Enriquecer con nombres de productos (indexar por codigoBarra, idProducto, skuBase y equivalentes)
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p){
    var name = p.descripcion || p.nombre || '';
    if (!name) return;
    if (p.codigoBarra) prodMap[String(p.codigoBarra)] = name;
    if (p.idProducto)  prodMap[String(p.idProducto)] = name;
    // skuBase solo para producto BASE (factor=1) — evita que presentaciones sobreescriban
    var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
    if (esBase && p.skuBase) prodMap[String(p.skuBase).trim().toUpperCase()] = name;
  });
  // Equivalentes → resuelven al producto base via skuBase
  try {
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      _sheetToObjects(equivSheet).forEach(function(e){
        if (!e.codigoBarra || !e.skuBase) return;
        var skuKey = String(e.skuBase).trim().toUpperCase();
        var name   = prodMap[skuKey];
        if (name) prodMap[String(e.codigoBarra)] = name;
      });
    }
  } catch(e) {}
  guia.detalle.forEach(function(d){
    d.descripcionProducto = d.descripcionProducto
      || prodMap[d.codigoProducto]
      || d.codigoProducto;
  });

  return { ok: true, data: guia };
}

function crearGuia(params) {
  var sheet  = getSheet('GUIAS');
  var idGuia = _generateId('G');
  var fecha  = new Date();

  var tipo = params.tipo || '';
  if (!_validarTipoGuia(tipo)) {
    return { ok: false, error: 'Tipo de guía inválido: ' + tipo };
  }

  if (params.idSesion) registrarActividad(params.idSesion, 'GUIA_CREADA', 1);

  sheet.appendRow([
    idGuia,
    tipo,
    fecha,
    params.usuario       || '',
    params.idProveedor   || '',
    params.idZona        || '',
    params.numeroDocumento || '',
    params.comentario    || '',
    0,                       // montoTotal (se calcula al cerrar)
    'ABIERTA',
    params.idPreingreso  || '',
    params.foto          || ''
  ]);

  return { ok: true, data: { idGuia: idGuia, estado: 'ABIERTA' } };
}

function agregarDetalleGuia(params) {
  return _conLock('agregarDetalleGuia', function() {
    return _agregarDetalleGuiaImpl(params);
  });
}

// Helper para envolver una operación con LockService, con timeout largo
// y reintentos automáticos en caso de timeout. Reduce dramáticamente los
// errores "servidor ocupado" en operaciones concurrentes legítimas.
function _conLock(nombre, fn) {
  var lock = LockService.getScriptLock();
  // Intento 1: 30s
  try {
    lock.waitLock(30000);
    try { return fn(); }
    finally { try { lock.releaseLock(); } catch(e) {} }
  } catch(eL) {
    // Intento 2: 15s extra (60s+ acumulado)
    try {
      lock.waitLock(15000);
      try { return fn(); }
      finally { try { lock.releaseLock(); } catch(e) {} }
    } catch(eL2) {
      Logger.log('[' + nombre + '] timeout lock tras 45s: ' + eL2.message);
      return { ok: false, error: 'Sistema saturado, espera unos segundos y reintenta' };
    }
  }
}

function _agregarDetalleGuiaImpl(params) {
  var sheet     = getSheet('GUIA_DETALLE');
  var idDetalle = _generateId('DET');

  // Forzar codigoBarra a string para preservar ceros a la izquierda
  var codigoBuscado = String(params.codigoProducto || '').trim().toUpperCase();

  // Validar que el código de producto existe (comparación case-insensitive)
  var productos = _sheetToObjects(getProductosSheet());
  var prod = productos.find(function(p) {
    return String(p.codigoBarra || '').trim().toUpperCase() === codigoBuscado;
  });

  // Si no está en PRODUCTOS_MASTER, buscar en EQUIVALENCIAS (MOS master SS, igual que el caché del frontend)
  if (!prod) {
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      var equivs = _sheetToObjects(equivSheet);
      var equiv  = equivs.find(function(e) {
        return String(e.codigoBarra || '').trim().toUpperCase() === codigoBuscado && _esActivo(e.activo);
      });
      if (equiv) {
        var skuB = String(equiv.skuBase || '').trim().toUpperCase();
        // Resolver al producto BASE (factor=1, activo). Evita matchear presentaciones
        // (factor>1) que comparten skuBase con su producto base.
        prod = productos.find(function(p) {
          var esBase = parseFloat(p.factorConversion || 1) === 1
                    && String(p.estado || '') !== '0';
          return esBase && (
            String(p.idProducto  || '').trim().toUpperCase() === skuB ||
            String(p.skuBase     || '').trim().toUpperCase() === skuB ||
            String(p.codigoBarra || '').trim().toUpperCase() === skuB
          );
        });
      }
    }
  }

  // Si no existe en ninguna tabla → sugerir ProductoNuevo
  if (!prod) {
    return {
      ok: false,
      error: 'PRODUCTO_NO_ENCONTRADO',
      mensaje: 'Código no registrado. ¿Deseas registrarlo como producto nuevo?',
      codigoBuscado: codigoBuscado
    };
  }

  var cantEsperada  = parseFloat(params.cantidadEsperada)  || 0;
  var cantRecibida  = parseFloat(params.cantidadRecibida !== undefined ? params.cantidadRecibida : cantEsperada);
  var precioUnit    = parseFloat(params.precioUnitario)    || 0;

  // Preservar el código escaneado exacto (puede ser equiv o master) — stock es independiente por barcode
  var cbProd = codigoBuscado;

  // ── AUTO-SUMA: si ya existe detalle (mismo idGuia + cb, no anulado), sumar cantidad ──
  // Previene duplicados por reintentos, doble-click o flujos paralelos.
  // Si la guía está CERRADA, también ajusta el stock por el delta.
  var detData    = sheet.getDataRange().getValues();
  var hdrs       = detData[0];
  var idxIdG     = hdrs.indexOf('idGuia');
  var idxCb      = hdrs.indexOf('codigoProducto');
  var idxRec     = hdrs.indexOf('cantidadRecibida');
  var idxObs     = hdrs.indexOf('observacion');
  var idxIdDet   = hdrs.indexOf('idDetalle');
  for (var dr = 1; dr < detData.length; dr++) {
    if (String(detData[dr][idxIdG]) !== String(params.idGuia)) continue;
    if (String(detData[dr][idxCb]).toUpperCase() !== cbProd) continue;
    if (String(detData[dr][idxObs] || '').toUpperCase() === 'ANULADO') continue;

    // Match: sumar cantidades en el detalle existente
    var existingId  = String(detData[dr][idxIdDet]);
    var qtyAnterior = parseFloat(detData[dr][idxRec]) || 0;
    var qtyNueva    = qtyAnterior + cantRecibida;
    sheet.getRange(dr + 1, idxRec + 1).setValue(qtyNueva);

    // Si la guía ya está CERRADA, ajustar stock por la diferencia agregada
    // Auto-suma sobre guía CERRADA: ajustar stock por la cantidad agregada
    // (excepto en envasados, que manejan stock directo)
    var guiaInfo2 = _getGuiaInfo(params.idGuia);
    if (guiaInfo2 && String(guiaInfo2.estado).toUpperCase() === 'CERRADA'
        && cantRecibida !== 0 && !_esGuiaEnvasado(guiaInfo2.tipo)) {
      var esIngreso2 = String(guiaInfo2.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
      var deltaSum = esIngreso2 ? cantRecibida : -cantRecibida;
      _actualizarStock(cbProd, deltaSum, {
        tipoOperacion: 'AUTO_SUMA_DETALLE',
        origen:        existingId,
        usuario:       String(params.usuario || ''),
        observacion:   'idGuia=' + params.idGuia
      });
    }

    return {
      ok: true,
      autoSumado: true,
      data: {
        idDetalle:          existingId,
        idGuia:             params.idGuia,
        codigoProducto:     cbProd,
        descripcionProducto: prod.descripcion || prod.nombre || prod.idProducto,
        cantidadEsperada:   cantEsperada,
        cantidadRecibida:   qtyNueva,
        precioUnitario:     precioUnit
      }
    };
  }

  // Lote: si viene fechaVencimiento, crear lote inmediatamente
  var idLote = params.idLote || '';
  var fechaVenc = params.fechaVencimiento || '';
  if (fechaVenc && fechaVenc !== '') {
    if (!idLote) idLote = _generateId('LOT');
    var loteSheet = getSheet('LOTES_VENCIMIENTO');
    if (loteSheet) {
      var lotNextRow = loteSheet.getLastRow() + 1;
      loteSheet.appendRow([
        idLote, cbProd, fechaVenc, cantRecibida, cantRecibida,
        params.idGuia, 'ACTIVO', new Date()
      ]);
      loteSheet.getRange(lotNextRow, 2).setNumberFormat('@').setValue(cbProd);
    }
  }
  // FIX REFORZADO: en lugar de appendRow + setNumberFormat (que no garantiza
  // texto porque Sheets infiere tipo del valor), usar setValues directamente
  // en un rango con formato '@' pre-aplicado a las cols clave.
  var nextRow  = sheet.getLastRow() + 1;
  // 1. Pre-formatear toda la nueva fila — col 3 (codigoBarra) e col 7 (idLote) como texto
  sheet.getRange(nextRow, 3).setNumberFormat('@');
  if (idLote) sheet.getRange(nextRow, 7).setNumberFormat('@');
  SpreadsheetApp.flush(); // forzar que el formato se aplique antes de escribir
  // 2. Escribir valores con setValues (respeta el formato de la celda)
  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    idDetalle,
    params.idGuia,
    String(cbProd),
    cantEsperada,
    cantRecibida,
    precioUnit,
    String(idLote || ''),
    params.observacion || ''
  ]]);
  // 3. Re-asegurar formato y valor exacto en col codigoBarra (defensa final)
  sheet.getRange(nextRow, 3).setNumberFormat('@').setValue(String(cbProd));

  return {
    ok: true,
    data: {
      idDetalle:          idDetalle,
      idGuia:             params.idGuia,
      codigoProducto:     cbProd,
      descripcionProducto: prod.descripcion || prod.nombre || prod.idProducto,
      cantidadEsperada:   cantEsperada,
      cantidadRecibida:   cantRecibida,
      precioUnitario:     precioUnit,
      idLote:             idLote,
      fechaVencimiento:   fechaVenc
    }
  };
}

// ── Actualizar cantidad recibida de un detalle existente ────────────
function actualizarCantidadDetalle(params) {
  return _conLock('actualizarCantidadDetalle', function() {
    return _actualizarCantidadDetalleImpl(params);
  });
}

function _actualizarCantidadDetalleImpl(params) {
  var idDetalle = String(params.idDetalle || '');
  var cantidad  = parseFloat(params.cantidadRecibida);
  if (!idDetalle || isNaN(cantidad)) return { ok: false, error: 'idDetalle y cantidadRecibida requeridos' };

  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxRec = hdrs.indexOf('cantidadRecibida');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCod = hdrs.indexOf('codigoProducto');
  if (idxId < 0 || idxRec < 0) return { ok: false, error: 'Columnas no encontradas' };

  var idxObs = hdrs.indexOf('observacion');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idDetalle) {
      var cantidadVieja = parseFloat(data[i][idxRec]) || 0;
      var idGuia        = String(data[i][idxIdG] || '');
      var codigo        = String(data[i][idxCod] || '');
      var obsActual     = idxObs >= 0 ? String(data[i][idxObs] || '').toUpperCase() : '';

      // Ajustar stock por diff SOLO si guía CERRADA y NO es envasado
      // (envasados manejan stock directo, no via cerrarGuia)
      var guiaInfo = _getGuiaInfo(idGuia);
      if (guiaInfo && String(guiaInfo.estado).toUpperCase() === 'CERRADA' && codigo
          && !_esGuiaEnvasado(guiaInfo.tipo)) {
        var diff = cantidad - cantidadVieja;
        if (diff !== 0) {
          var esIngreso = String(guiaInfo.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
          var delta = esIngreso ? diff : -diff;
          _actualizarStock(codigo, delta, {
            tipoOperacion: 'EDICION_CANTIDAD',
            origen:        idDetalle,
            usuario:       String(params.usuario || ''),
            observacion:   'idGuia=' + idGuia + ' diff=' + diff
          });
        }
      }

      // Si estaba ANULADO y ahora se edita a cantidad > 0, des-anularlo
      if (obsActual === 'ANULADO' && cantidad > 0 && idxObs >= 0) {
        sheet.getRange(i + 1, idxObs + 1).setValue('');
      }

      sheet.getRange(i + 1, idxRec + 1).setValue(cantidad);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado: ' + idDetalle };
}

// Actualizar precios unitarios de varias líneas de una guía en batch.
// params: { idGuia, items: [{ idDetalle, precioUnitario }] }
function actualizarPreciosDetalle(params) {
  if (!params.idGuia || !Array.isArray(params.items) || !params.items.length) {
    return { ok: false, error: 'idGuia + items[] requeridos' };
  }
  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxPU  = hdrs.indexOf('precioUnitario');
  if (idxId < 0 || idxPU < 0) return { ok: false, error: 'Columnas no encontradas en GUIA_DETALLE' };

  // Mapa idDetalle → precio nuevo
  var mapPrecio = {};
  params.items.forEach(function(it) {
    if (it.idDetalle && it.precioUnitario !== undefined && it.precioUnitario !== '' && !isNaN(parseFloat(it.precioUnitario))) {
      mapPrecio[String(it.idDetalle)] = parseFloat(it.precioUnitario);
    }
  });

  var actualizados = 0;
  var sumaTotal   = 0;
  var idxCantR    = hdrs.indexOf('cantidadRecibida');
  var idxCantE    = hdrs.indexOf('cantidadEsperada');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdG]) !== String(params.idGuia)) continue;
    var idDet = String(data[i][idxId]);
    if (mapPrecio.hasOwnProperty(idDet)) {
      sheet.getRange(i + 1, idxPU + 1).setValue(mapPrecio[idDet]);
      actualizados++;
    }
    // Sumar para recalcular monto total de la guía
    var cant = (idxCantR >= 0 ? parseFloat(data[i][idxCantR]) : 0) || (idxCantE >= 0 ? parseFloat(data[i][idxCantE]) : 0) || 0;
    var precio = mapPrecio[idDet] !== undefined ? mapPrecio[idDet] : (parseFloat(data[i][idxPU]) || 0);
    sumaTotal += cant * precio;
  }

  // Actualizar montoTotal de la guía
  try {
    var shGuias = getSheet('GUIAS');
    var dataG = shGuias.getDataRange().getValues();
    var hdrsG = dataG[0];
    var idxGid = hdrsG.indexOf('idGuia');
    var idxGmt = hdrsG.indexOf('montoTotal');
    if (idxGid >= 0 && idxGmt >= 0) {
      for (var j = 1; j < dataG.length; j++) {
        if (String(dataG[j][idxGid]) === String(params.idGuia)) {
          shGuias.getRange(j + 1, idxGmt + 1).setValue(Math.round(sumaTotal * 100) / 100);
          break;
        }
      }
    }
  } catch(_){}

  return { ok: true, data: { actualizados: actualizados, montoTotalNuevo: Math.round(sumaTotal * 100) / 100 } };
}

// Helper: lee la guía y retorna {tipo, estado} o null si no existe
function _getGuiaInfo(idGuia) {
  var sheet   = getSheet('GUIAS');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxId   = hdrs.indexOf('idGuia');
  var idxTipo = hdrs.indexOf('tipo');
  var idxEst  = hdrs.indexOf('estado');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      return { tipo: String(data[i][idxTipo] || ''), estado: String(data[i][idxEst] || '') };
    }
  }
  return null;
}

// ── Actualizar solo la fecha de vencimiento de un detalle existente ─
function actualizarFechaVencimiento(params) {
  var idDetalle = String(params.idDetalle || '');
  var fechaVenc = params.fechaVencimiento || '';
  if (!idDetalle) return { ok: false, error: 'idDetalle requerido' };

  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxVenc= hdrs.indexOf('fechaVencimiento');
  if (idxId < 0 || idxVenc < 0) return { ok: false, error: 'Columnas no encontradas' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idDetalle) {
      var val = fechaVenc ? new Date(fechaVenc + 'T12:00:00') : '';
      sheet.getRange(i + 1, idxVenc + 1).setValue(val);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado: ' + idDetalle };
}

// ── Anular un ítem de detalle ────────────────────────────
// Si la guía padre está CERRADA, devuelve el stock que ya se había aplicado.
// Si está ABIERTA, solo marca como anulado (stock aún no descontado).
function anularDetalle(params) {
  return _conLock('anularDetalle', function() {
    return _anularDetalleImpl(params);
  });
}

function _anularDetalleImpl(params) {
  var idDetalle = params.idDetalle;
  if (!idDetalle) return { ok: false, error: 'idDetalle requerido' };

  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxObs = hdrs.indexOf('observacion');
  var idxRec = hdrs.indexOf('cantidadRecibida');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCod = hdrs.indexOf('codigoProducto');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === idDetalle) {
      // Idempotencia: si ya está ANULADO, no devolver stock de nuevo
      if (String(data[i][idxObs] || '') === 'ANULADO') {
        return { ok: true, yaAnulado: true };
      }

      var cantidadActual = idxRec >= 0 ? (parseFloat(data[i][idxRec]) || 0) : 0;
      var idGuia         = String(data[i][idxIdG] || '');
      var codigo         = String(data[i][idxCod] || '');

      // Si la guía está CERRADA y tenía cantidad > 0, devolver stock.
      // EXCEPCIÓN: guías de envasado manejan stock directamente desde Envasados.gs
      // (no via cerrarGuia), por lo tanto anular detalle NO debe revertir stock.
      var guiaInfo = _getGuiaInfo(idGuia);
      if (guiaInfo && String(guiaInfo.estado).toUpperCase() === 'CERRADA'
          && cantidadActual > 0 && codigo
          && !_esGuiaEnvasado(guiaInfo.tipo)) {
        var esIngreso = String(guiaInfo.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
        // Reverso: si era INGRESO, suma se revierte (resta); si SALIDA, devuelve (suma)
        var delta = esIngreso ? -cantidadActual : cantidadActual;
        _actualizarStock(codigo, delta, {
          tipoOperacion: 'ANULACION_DETALLE',
          origen:        idDetalle,
          usuario:       String(params.usuario || ''),
          observacion:   'idGuia=' + idGuia
        });
      }

      // Si era una línea de PN pendiente, marcar el PN correspondiente como ANULADO
      // (de lo contrario MOS lo sigue mostrando para aprobación)
      var obsActual = String(data[i][idxObs] || '').toUpperCase();
      if (obsActual === 'PN_PENDIENTE' && idGuia && codigo) {
        try { _anularPNPorGuiaYCodigo(idGuia, codigo); } catch(e) {
          Logger.log('No se pudo anular PN huérfano: ' + e.message);
        }
      }

      // Marcar como anulado y poner cantidad en 0
      sheet.getRange(i + 1, idxObs + 1).setValue('ANULADO');
      if (idxRec >= 0) sheet.getRange(i + 1, idxRec + 1).setValue(0);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado' };
}

// Marca como ANULADO los PNs PENDIENTE asociados a (idGuia, codigoBarra).
// Idempotente: si no hay PN PENDIENTE coincidente, no hace nada.
function _anularPNPorGuiaYCodigo(idGuia, codigoBarra) {
  var sh = getSheet('PRODUCTO_NUEVO');
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  var hdrs = data[0];
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCb  = hdrs.indexOf('codigoBarra');
  var idxEst = hdrs.indexOf('estado');
  if (idxIdG < 0 || idxCb < 0 || idxEst < 0) return;

  var cb = String(codigoBarra || '').toUpperCase();
  var ig = String(idGuia || '');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxIdG]) !== ig) continue;
    if (String(data[r][idxCb] || '').toUpperCase() !== cb) continue;
    if (String(data[r][idxEst] || '').toUpperCase() !== 'PENDIENTE') continue;
    sh.getRange(r + 1, idxEst + 1).setValue('ANULADO');
  }
}

// ── Reabrir una guía cerrada (requiere adminPin en el cliente) ──
// REVIERTE el stock que se aplicó al cerrar para que al volver a cerrar
// no se descuente dos veces. Es la operación inversa de cerrarGuia.
function reabrirGuia(params) {
  var idGuia = params.idGuia;
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId   = hdrs.indexOf('idGuia');
  var idxEst  = hdrs.indexOf('estado');
  var idxTipo = hdrs.indexOf('tipo');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      var estadoActual = String(data[i][idxEst] || '').toUpperCase();
      // Idempotencia: si ya está ABIERTA, no revertir stock
      if (estadoActual === 'ABIERTA') {
        return { ok: true, yaAbierta: true };
      }
      var tipoGuia = String(data[i][idxTipo] || '');
      var esIngreso = tipoGuia.toUpperCase().indexOf('INGRESO') === 0;

      // Solo revertir si estaba CERRADA Y no es de envasado.
      // Envasados manejan stock directo (no via cerrarGuia), por lo tanto
      // reabrir NO debe revertir stock — eso descuadraría el inventario.
      if (estadoActual === 'CERRADA' && !_esGuiaEnvasado(tipoGuia)) {
        var detalles = _sheetToObjects(getSheet('GUIA_DETALLE')).filter(function(d) {
          return d.idGuia === idGuia && d.observacion !== 'ANULADO';
        });
        detalles.forEach(function(d) {
          var cant = parseFloat(d.cantidadRecibida) || 0;
          if (cant === 0 || !d.codigoProducto) return;
          // Reverso del cierre: si era INGRESO, restar; si SALIDA, sumar
          var deltaReverso = esIngreso ? -cant : cant;
          _actualizarStock(String(d.codigoProducto), deltaReverso, {
            tipoOperacion: 'REABRIR_REVERSO',
            origen:        idGuia,
            usuario:       String(params.usuario || ''),
            observacion:   'tipo=' + tipoGuia
          });
        });
      }

      sheet.getRange(i + 1, idxEst + 1).setValue('ABIERTA');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Guía no encontrada' };
}

// ── Auto-cerrar guías abiertas de días anteriores ────────────
function autoCloseDayGuias() {
  var sheet   = getSheet('GUIAS');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxId   = hdrs.indexOf('idGuia');
  var idxFec  = hdrs.indexOf('fecha');
  var idxEst  = hdrs.indexOf('estado');
  var tz      = Session.getScriptTimeZone();
  var hoy     = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var cerradas = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxEst] !== 'ABIERTA') continue;
    var fechaGuia = '';
    var fv = data[i][idxFec];
    if (fv instanceof Date) {
      fechaGuia = Utilities.formatDate(fv, tz, 'yyyy-MM-dd');
    } else {
      fechaGuia = String(fv || '').substring(0, 10);
    }
    if (fechaGuia && fechaGuia < hoy) {
      sheet.getRange(i + 1, idxEst + 1).setValue('AUTOCERRADA');
      cerradas++;
    }
  }
  return { ok: true, data: { cerradas: cerradas } };
}

function cerrarGuia(idGuia, usuario, idSesion, opts) {
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  return _conLock('cerrarGuia', function() {
    return _cerrarGuiaImpl(idGuia, usuario, idSesion, opts);
  });
}

function _cerrarGuiaImpl(idGuia, usuario, idSesion, opts) {
  opts = opts || {};
  var skipMosSync = opts.skipMosSync === true;  // omitir sync a MOS (uso en cierres masivos para evitar timeout)

  var guiasSheet    = getSheet('GUIAS');
  var guias         = guiasSheet.getDataRange().getValues();
  var headers       = guias[0];
  var idxIdGuia     = headers.indexOf('idGuia');
  var idxEstado     = headers.indexOf('estado');
  var idxTipo       = headers.indexOf('tipo');
  var idxMontoTotal = headers.indexOf('montoTotal');

  var filaGuia = -1;
  var tipoGuia = '';
  var estadoActual = '';
  for (var i = 1; i < guias.length; i++) {
    if (guias[i][idxIdGuia] === idGuia) {
      filaGuia = i + 1;
      tipoGuia = guias[i][idxTipo];
      estadoActual = String(guias[i][idxEstado] || '').toUpperCase();
      break;
    }
  }
  if (filaGuia < 0) return { ok: false, error: 'Guía no encontrada' };

  // Idempotencia: si la guía ya está CERRADA (o AUTOCERRADA), no reaplicar stock.
  // Permite que el frontend reintente sin duplicar descuentos.
  if (estadoActual === 'CERRADA' || estadoActual === 'AUTOCERRADA') {
    var montoExistente = idxMontoTotal >= 0 ? (parseFloat(guias[filaGuia - 1][idxMontoTotal]) || 0) : 0;
    return { ok: true, data: { idGuia: idGuia, estado: estadoActual, montoTotal: montoExistente, yaCerrada: true } };
  }

  // Obtener detalles
  var detalles = _sheetToObjects(getSheet('GUIA_DETALLE')).filter(function(d){
    return d.idGuia === idGuia;
  });

  // Calcular monto total
  var montoTotal = detalles.reduce(function(acc, d) {
    return acc + (parseFloat(d.cantidadRecibida) || 0) * (parseFloat(d.precioUnitario) || 0);
  }, 0);

  var esIngreso = tipoGuia.startsWith('INGRESO');
  var esEnvasado = _esGuiaEnvasado(tipoGuia);

  // Actualizar stock por cada detalle.
  // SALIDA_ENVASADO / INGRESO_ENVASADO: stock ya fue aplicado por Envasados.gs
  // directamente con _actualizarStock — saltarse para no duplicar.
  if (!esEnvasado) {
    detalles.forEach(function(d) {
      var cantidad = parseFloat(d.cantidadRecibida) || 0;
      if (cantidad === 0) return;
      var delta = esIngreso ? cantidad : -cantidad;
      _actualizarStock(d.codigoProducto, delta, {
        tipoOperacion: 'CIERRE_GUIA',
        origen:        idGuia,
        usuario:       String(usuario || ''),
        observacion:   'tipo=' + tipoGuia
      });
      // Si es ingreso → crear/actualizar lote de vencimiento si tiene fecha
      if (esIngreso && d.idLote && d.idLote !== '') {
        _actualizarLote(d.idLote, d.codigoProducto, cantidad, d.fechaVencimiento || '', idGuia);
      }
    });
  }

  if (idSesion) registrarActividad(idSesion, 'GUIA_CERRADA', 1);

  // Marcar guía como cerrada
  guiasSheet.getRange(filaGuia, idxEstado + 1).setValue('CERRADA');
  guiasSheet.getRange(filaGuia, idxMontoTotal + 1).setValue(montoTotal);

  // Si fue SALIDA_MERMA (cierre semanal manual), marcar mermas asociadas como DESECHADA
  // y notificar a MASTER/ADMINISTRADOR vía push
  if (tipoGuia === 'SALIDA_MERMA') {
    try { _cerrarMermasDeGuia(idGuia, detalles); }
    catch(eM) { Logger.log('cerrar mermas: ' + eM.message); }
  }

  // Si fue INGRESO_PROVEEDOR con idProveedor, sincroniza productos a MOS
  // (silencioso: si falla, log y sigue)
  // OMITIR si skipMosSync (cierres masivos para evitar timeout — el sync hace
  // un UrlFetchApp por cada detalle, lento si la guía tiene muchos productos).
  if (!skipMosSync) {
    try {
      if (tipoGuia === 'INGRESO_PROVEEDOR') {
        var idxProv = headers.indexOf('idProveedor');
        var idProveedor = idxProv >= 0 ? String(guias[filaGuia - 1][idxProv] || '').trim() : '';
        if (idProveedor) _syncProductosProvAMos(idProveedor, detalles);
      }
    } catch(eS) { Logger.log('sync productos proveedor: ' + eS.message); }
  }

  return { ok: true, data: { idGuia: idGuia, estado: 'CERRADA', montoTotal: montoTotal } };
}

// Llama al GAS de MOS para upsert de cada producto en PROVEEDORES_PRODUCTOS
function _syncProductosProvAMos(idProveedor, detalles) {
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) return;
  detalles.forEach(function(d){
    var cb = String(d.codigoProducto || '').trim();
    var precio = parseFloat(d.precioUnitario) || 0;
    if (!cb) return;
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'upsertProductoProveedor',
          idProveedor: idProveedor,
          codigoBarra: cb,
          precioUnitario: precio,
          descripcion: d.descripcion || ''
        }),
        muteHttpExceptions: true
      });
    } catch(e) { /* silencioso */ }
  });
}

function _actualizarLote(idLote, codigoProducto, cantidad, fechaVencimiento, idGuia) {
  var sheet   = getSheet('LOTES_VENCIMIENTO');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxId   = headers.indexOf('idLote');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] !== idLote) continue;
    // Actualizar cantidad final y fechaVencimiento en fila existente
    var idxQI = headers.indexOf('cantidadInicial');
    var idxQA = headers.indexOf('cantidadActual');
    var idxFV = headers.indexOf('fechaVencimiento');
    if (idxQI >= 0) sheet.getRange(i + 1, idxQI + 1).setValue(cantidad);
    if (idxQA >= 0) sheet.getRange(i + 1, idxQA + 1).setValue(cantidad);
    if (idxFV >= 0 && fechaVencimiento) sheet.getRange(i + 1, idxFV + 1).setValue(fechaVencimiento);
    return;
  }
  // No existe → crear con cantidad y fecha confirmadas
  sheet.appendRow([
    idLote, codigoProducto, fechaVencimiento || '', cantidad, cantidad, idGuia, 'ACTIVO', new Date()
  ]);
}

// ── PICKUPS (pedidos externos de Cabanossi / n8n) ──────────────
function getPickups(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  var filtroEstado = params.estado || 'PENDIENTE';
  var rows = _sheetToObjects(sheet);
  var result = rows.filter(function(r) {
    return !filtroEstado || r.estado === filtroEstado;
  }).map(function(r) {
    try { r.items = JSON.parse(r.items || '[]'); } catch(e) { r.items = []; }
    return r;
  });

  return { ok: true, data: result };
}

function actualizarPickup(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var idxId   = headers.indexOf('idPickup');
  var idxEst  = headers.indexOf('estado');
  var idxAte  = headers.indexOf('fechaAtendido');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(params.idPickup)) {
      if (idxEst >= 0) sheet.getRange(i + 1, idxEst + 1).setValue(params.estado);
      if (idxAte >= 0 && params.estado === 'COMPLETADO') {
        sheet.getRange(i + 1, idxAte + 1).setValue(new Date());
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pickup no encontrado' };
}

function crearDespachoRapido(params) {
  var idZona   = params.idZona   || '';
  var tipo     = params.tipo     || 'SALIDA_ZONA';
  var items    = params.items    || [];
  var usuario  = params.usuario  || '';
  var nota     = params.nota     || '';
  var imprimir = params.imprimir !== false;

  // idZona solo requerido cuando el tipo es SALIDA_ZONA
  if (tipo === 'SALIDA_ZONA' && !idZona) return { ok: false, error: 'idZona requerido' };
  if (!items.length) return { ok: false, error: 'Carrito vacío' };

  var comentario = nota || ('Despacho rápido' + (tipo !== 'SALIDA_ZONA' ? ' · ' + tipo : ''));

  // 1. Crear guía con el tipo correcto
  var guiaRes = crearGuia({ tipo: tipo, idZona: idZona || null, usuario: usuario, comentario: comentario });
  if (!guiaRes.ok) return guiaRes;
  var idGuia = guiaRes.data.idGuia;

  // 2. Registrar cada ítem
  var errores = [];
  items.forEach(function(item) {
    var qty = parseFloat(item.cantidad) || 0;
    if (qty <= 0) return;
    var det = agregarDetalleGuia({
      idGuia:           idGuia,
      codigoProducto:   String(item.codigoBarra || '').trim(),
      cantidadEsperada: qty,
      cantidadRecibida: qty,
      usuario:          usuario
    });
    if (!det.ok) errores.push(String(item.codigoBarra) + ': ' + det.error);
  });

  // 3. Cerrar guía (descuenta stock)
  var cerrarRes = cerrarGuia(idGuia, usuario, null);
  if (!cerrarRes.ok) return { ok: false, error: 'Error al cerrar guía: ' + cerrarRes.error };

  // 4. Imprimir ticket
  var impresion = { ok: false, error: 'omitido' };
  if (imprimir) {
    try { impresion = imprimirTicketGuia({ idGuia: idGuia }); } catch(e) { impresion = { ok: false, error: e.message }; }
  }

  return { ok: true, data: { idGuia: idGuia, errores: errores, impresion: impresion } };
}

function _validarTipoGuia(tipo) {
  var validos = ['INGRESO_PROVEEDOR','INGRESO_JEFATURA','INGRESO_ENVASADO',
                 'SALIDA_DEVOLUCION','SALIDA_ZONA','SALIDA_JEFATURA',
                 'SALIDA_ENVASADO','SALIDA_MERMA'];
  return validos.indexOf(tipo) >= 0;
}
