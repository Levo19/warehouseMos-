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
  if (params.estado)  rows = rows.filter(function(r){ return r.estado === params.estado; });
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

  // Enriquecer con nombres de productos
  var productos = _sheetToObjects(getSheet('PRODUCTOS'));
  var prodMap = {};
  productos.forEach(function(p){ prodMap[p.idProducto] = p.descripcion; });
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

  // Validar que el código de producto existe
  var productos = _sheetToObjects(getSheet('PRODUCTOS'));
  var prod = productos.find(function(p){ return p.idProducto === params.codigoProducto || p.codigoBarra === params.codigoProducto; });

  // Si no existe → sugerir ProductoNuevo
  if (!prod) {
    return {
      ok: false,
      error: 'PRODUCTO_NO_ENCONTRADO',
      mensaje: 'Código no registrado. ¿Deseas registrarlo como producto nuevo?',
      codigoBuscado: params.codigoProducto
    };
  }

  var cantEsperada  = parseFloat(params.cantidadEsperada)  || 0;
  var cantRecibida  = parseFloat(params.cantidadRecibida !== undefined ? params.cantidadRecibida : cantEsperada);
  var precioUnit    = parseFloat(params.precioUnitario)    || 0;

  sheet.appendRow([
    idDetalle,
    params.idGuia,
    prod.idProducto,
    cantEsperada,
    cantRecibida,
    precioUnit,
    params.idLote    || '',
    params.observacion || ''
  ]);

  return { ok: true, data: { idDetalle: idDetalle, codigoProducto: prod.idProducto } };
}

function cerrarGuia(idGuia, usuario) {
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

function _validarTipoGuia(tipo) {
  var validos = ['INGRESO_PROVEEDOR','INGRESO_JEFATURA',
                 'SALIDA_DEVOLUCION','SALIDA_ZONA','SALIDA_JEFATURA',
                 'SALIDA_ENVASADO','SALIDA_MERMA'];
  return validos.indexOf(tipo) >= 0;
}
