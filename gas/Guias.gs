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

  // Enriquecer con nombres de productos (indexar por codigoBarra y por idProducto)
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p){
    if (p.codigoBarra) prodMap[String(p.codigoBarra)] = p.descripcion;
    prodMap[p.idProducto] = p.descripcion;
  });
  guia.detalle.forEach(function(d){
    d.descripcionProducto = prodMap[d.codigoProducto] || d.codigoProducto;
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
  var sheet     = getSheet('GUIA_DETALLE');
  var idDetalle = _generateId('DET');

  // Forzar codigoBarra a string para preservar ceros a la izquierda
  var codigoBuscado = String(params.codigoProducto || '').trim();

  // Validar que el código de producto existe
  var productos = _sheetToObjects(getProductosSheet());
  var prod = productos.find(function(p) {
    return String(p.codigoBarra).trim() === codigoBuscado;
  });

  // Si no existe → sugerir ProductoNuevo
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

  var cbProd = String(prod.codigoBarra);

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
  var nextRow  = sheet.getLastRow() + 1;
  sheet.appendRow([
    idDetalle,
    params.idGuia,
    cbProd,
    cantEsperada,
    cantRecibida,
    precioUnit,
    idLote,
    params.observacion || ''
  ]);
  sheet.getRange(nextRow, 3).setNumberFormat('@').setValue(cbProd);

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
  var idDetalle = String(params.idDetalle || '');
  var cantidad  = parseFloat(params.cantidadRecibida);
  if (!idDetalle || isNaN(cantidad)) return { ok: false, error: 'idDetalle y cantidadRecibida requeridos' };

  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxRec = hdrs.indexOf('cantidadRecibida');
  if (idxId < 0 || idxRec < 0) return { ok: false, error: 'Columnas no encontradas' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idDetalle) {
      sheet.getRange(i + 1, idxRec + 1).setValue(cantidad);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado: ' + idDetalle };
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
function anularDetalle(params) {
  var idDetalle = params.idDetalle;
  if (!idDetalle) return { ok: false, error: 'idDetalle requerido' };

  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId = hdrs.indexOf('idDetalle');
  var idxObs = hdrs.indexOf('observacion');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === idDetalle) {
      // Marcar como anulado en observacion
      sheet.getRange(i + 1, idxObs + 1).setValue('ANULADO');
      // Poner cantidadRecibida = 0
      var idxRec = hdrs.indexOf('cantidadRecibida');
      if (idxRec >= 0) sheet.getRange(i + 1, idxRec + 1).setValue(0);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado' };
}

// ── Reabrir una guía cerrada (requiere adminPin en el cliente) ──
function reabrirGuia(params) {
  var idGuia = params.idGuia;
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId  = hdrs.indexOf('idGuia');
  var idxEst = hdrs.indexOf('estado');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
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

function cerrarGuia(idGuia, usuario, idSesion) {
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var guiasSheet    = getSheet('GUIAS');
  var guias         = guiasSheet.getDataRange().getValues();
  var headers       = guias[0];
  var idxIdGuia     = headers.indexOf('idGuia');
  var idxEstado     = headers.indexOf('estado');
  var idxTipo       = headers.indexOf('tipo');
  var idxMontoTotal = headers.indexOf('montoTotal');

  var filaGuia = -1;
  var tipoGuia = '';
  for (var i = 1; i < guias.length; i++) {
    if (guias[i][idxIdGuia] === idGuia) {
      filaGuia = i + 1;
      tipoGuia = guias[i][idxTipo];
      break;
    }
  }
  if (filaGuia < 0) return { ok: false, error: 'Guía no encontrada' };

  // Obtener detalles
  var detalles = _sheetToObjects(getSheet('GUIA_DETALLE')).filter(function(d){
    return d.idGuia === idGuia;
  });

  // Calcular monto total
  var montoTotal = detalles.reduce(function(acc, d) {
    return acc + (parseFloat(d.cantidadRecibida) || 0) * (parseFloat(d.precioUnitario) || 0);
  }, 0);

  var esIngreso = tipoGuia.startsWith('INGRESO');

  // Actualizar stock por cada detalle
  detalles.forEach(function(d) {
    var cantidad = parseFloat(d.cantidadRecibida) || 0;
    if (cantidad === 0) return;

    var delta = esIngreso ? cantidad : -cantidad;
    _actualizarStock(d.codigoProducto, delta);

    // Si es ingreso → crear/actualizar lote de vencimiento si tiene fecha
    if (esIngreso && d.idLote && d.idLote !== '') {
      _actualizarLote(d.idLote, d.codigoProducto, cantidad, idGuia);
    }
  });

  if (idSesion) registrarActividad(idSesion, 'GUIA_CERRADA', 1);

  // Marcar guía como cerrada
  guiasSheet.getRange(filaGuia, idxEstado + 1).setValue('CERRADA');
  guiasSheet.getRange(filaGuia, idxMontoTotal + 1).setValue(montoTotal);

  return { ok: true, data: { idGuia: idGuia, estado: 'CERRADA', montoTotal: montoTotal } };
}

function _actualizarLote(idLote, codigoProducto, cantidad, idGuia) {
  var sheet = getSheet('LOTES_VENCIMIENTO');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxId   = headers.indexOf('idLote');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === idLote) return; // ya existe
  }
  // Crear nuevo lote (sin fecha de vencimiento — se llena desde GuiaDetalle si se tiene)
  sheet.appendRow([
    idLote, codigoProducto, '', cantidad, cantidad, idGuia, 'ACTIVO', new Date()
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
