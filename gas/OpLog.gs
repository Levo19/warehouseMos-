// ============================================================
// warehouseMos — OpLog.gs
//
// Operation log idempotente para resolver el "problema #1"
// (pérdida de información al escanear con UI optimista).
//
// Cada operación del frontend lleva un `idOp` único generado en
// cliente ('OP-<ts>-<rand>'). El server:
//   - Si idOp ya existe en OPS_LOG con estado APPLIED → retorna
//     el resultado cacheado (idempotente).
//   - Si no existe → ejecuta el tipo de op correspondiente,
//     persiste el row con estado=APPLIED y resultado.
//   - Si falla → persiste estado=FAILED con error.
//
// El frontend mantiene una cola IndexedDB de ops pending; el
// endpoint listarOpsPendientes le permite verificar qué quedó
// efectivamente aplicado tras recargar.
//
// Tipos soportados (extensible):
//   SCAN, EDIT_QTY, DELETE_ITEM, ANULAR_DETALLE, ANULAR_GUIA,
//   PN_REGISTRAR, MERMA_AGREGAR, MERMA_SOLUCIONAR,
//   MERMA_PROCESAR, CARGADOR_ADD, CARGADOR_REMOVE
// ============================================================

function aplicarOp(params) {
  var idOp = String(params.idOp || '').trim();
  if (!idOp) return { ok: false, error: 'idOp requerido' };
  var tipo = String(params.tipo || '').toUpperCase();
  if (!tipo) return { ok: false, error: 'tipo requerido' };

  return _conLock('aplicarOp', function() {
    var sheet = getSheet('OPS_LOG');
    if (!sheet) return { ok: false, error: 'tabla OPS_LOG no existe — ejecutar setupExtenderF0()' };
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0];
    var idxIdOp = hdrs.indexOf('idOp');
    var idxEst  = hdrs.indexOf('estado');
    var idxRes  = hdrs.indexOf('resultado');

    // Idempotencia: si idOp ya APPLIED, retornar resultado cacheado
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxIdOp]) === idOp) {
        var est = String(data[i][idxEst] || '').toUpperCase();
        if (est === 'APPLIED') {
          var cached;
          try { cached = JSON.parse(data[i][idxRes] || '{}'); }
          catch(_) { cached = { ok: true, dedup: true }; }
          if (cached && cached.ok !== false) cached.dedup = true;
          return cached;
        }
        // Si FAILED, dejamos reintentar abajo (sobrescribir)
        break;
      }
    }

    // Ejecutar según tipo
    var payload = params.payload || {};
    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); } catch(_) { payload = {}; }
    }

    var resultado;
    try {
      resultado = _ejecutarOp(tipo, payload, params);
    } catch(e) {
      resultado = { ok: false, error: e.message, stack: e.stack };
    }

    var estadoFinal = (resultado && resultado.ok) ? 'APPLIED' : 'FAILED';
    var errorMsg    = (resultado && !resultado.ok) ? (resultado.error || '') : '';
    var resStr      = JSON.stringify(resultado || {});

    // Escribir/actualizar row
    var existingRow = -1;
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][idxIdOp]) === idOp) { existingRow = j + 1; break; }
    }
    var rowVals = [
      idOp,
      String(params.idGuia || payload.idGuia || ''),
      tipo,
      typeof params.payload === 'string' ? params.payload : JSON.stringify(payload),
      estadoFinal,
      String(params.deviceId || ''),
      String(params.usuario  || ''),
      new Date(),
      new Date(),
      errorMsg,
      resStr
    ];
    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowVals.length).setValues([rowVals]);
    } else {
      sheet.appendRow(rowVals);
    }

    return resultado;
  });
}

// ── Listar ops pendientes/fallidas por guía ──
// El cliente compara contra su cola local para reconciliar.
function listarOpsPendientes(params) {
  var idGuia = String(params.idGuia || '').trim();
  try {
    var rows = _sheetToObjects(getSheet('OPS_LOG'));
    var filtered = rows.filter(function(r) {
      if (idGuia && String(r.idGuia || '') !== idGuia) return false;
      var est = String(r.estado || '').toUpperCase();
      return est === 'FAILED' || est === 'PENDING';
    });
    return { ok: true, data: filtered };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Dispatcher de tipos de op ──
function _ejecutarOp(tipo, payload, meta) {
  switch (tipo) {
    case 'SCAN':
      // payload: { idGuia, codigoProducto, cantidad, observacion? }
      return agregarDetalleGuia({
        idGuia:           payload.idGuia,
        codigoProducto:   payload.codigoProducto,
        cantidadEsperada: payload.cantidad || payload.cantidadEsperada || 0,
        cantidadRecibida: payload.cantidad || payload.cantidadRecibida || 0,
        precioUnitario:   payload.precioUnitario || 0,
        observacion:      payload.observacion || ''
      });

    case 'EDIT_QTY':
      // payload: { idDetalle, cantidadRecibida }
      return actualizarCantidadDetalle({
        idDetalle: payload.idDetalle,
        cantidadRecibida: payload.cantidadRecibida
      });

    case 'DELETE_ITEM':
    case 'ANULAR_DETALLE':
      return anularDetalle({ idDetalle: payload.idDetalle });

    case 'ANULAR_GUIA':
      // No hay endpoint directo, marcamos guía como ANULADA si existe
      return _anularGuiaSimple(payload.idGuia, meta.usuario);

    case 'PN_REGISTRAR':
      return registrarProductoNuevo(payload);

    case 'MERMA_AGREGAR':
      return agregarAMermas(payload);

    case 'MERMA_SOLUCIONAR':
      return solucionarMerma(payload);

    case 'MERMA_PROCESAR':
      return procesarEliminacionMermas({
        claveAdmin: payload.claveAdmin,
        usuario:    meta.usuario
      });

    case 'CARGADOR_ADD':
      return addCargadorDia(payload);

    case 'CARGADOR_REMOVE':
      return removeCargadorDia(payload);

    case 'CREAR_GUIA':
      return crearGuia(payload);

    case 'CERRAR_GUIA':
      return cerrarGuia(payload.idGuia, meta.usuario, meta.idSesion);

    case 'REABRIR_GUIA':
      return reabrirGuia(payload);

    default:
      return { ok: false, error: 'tipo de op no soportado: ' + tipo };
  }
}

// Helper minimal para ANULAR_GUIA — marca estado=ANULADA si la guía existe
function _anularGuiaSimple(idGuia, usuario) {
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId  = hdrs.indexOf('idGuia');
  var idxEst = hdrs.indexOf('estado');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      sheet.getRange(i + 1, idxEst + 1).setValue('ANULADA');
      return { ok: true, data: { idGuia: idGuia, estado: 'ANULADA' } };
    }
  }
  return { ok: false, error: 'guía no encontrada' };
}
