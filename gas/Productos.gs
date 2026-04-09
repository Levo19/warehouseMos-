// ============================================================
// warehouseMos — Productos.gs
// CRUD productos, stock, lotes y producto nuevo
// ============================================================

function getProductos(params) {
  var rows = _sheetToObjects(getSheet('PRODUCTOS'));
  if (params.categoria) rows = rows.filter(function(r){ return r.idCategoria === params.categoria; });
  if (params.estado)    rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.soloBase)  rows = rows.filter(function(r){ return r.esEnvasable === '1'; });
  if (params.q) {
    var q = params.q.toLowerCase();
    rows = rows.filter(function(r){
      return (r.descripcion || '').toLowerCase().indexOf(q) >= 0 ||
             (r.codigoBarra || '').indexOf(q) >= 0 ||
             (r.idProducto  || '').toLowerCase().indexOf(q) >= 0;
    });
  }
  return { ok: true, data: rows };
}

function getProducto(codigo) {
  var rows = _sheetToObjects(getSheet('PRODUCTOS'));
  var prod = rows.find(function(p){
    return p.idProducto === codigo || p.codigoBarra === codigo;
  });
  if (!prod) return { ok: false, error: 'Producto no encontrado: ' + codigo };

  // Enriquecer con stock actual
  prod.stockActual = _getStockProducto(prod.idProducto).cantidad;

  // Lotes vigentes
  var lotes = _sheetToObjects(getSheet('LOTES_VENCIMIENTO'));
  prod.lotes = lotes.filter(function(l){
    return l.codigoProducto === prod.idProducto && l.estado === 'ACTIVO' &&
           parseFloat(l.cantidadActual) > 0;
  });

  return { ok: true, data: prod };
}

function getStock(params) {
  var rows = _sheetToObjects(getSheet('STOCK'));
  // Enriquecer con info de producto
  var productos = _sheetToObjects(getSheet('PRODUCTOS'));
  var prodMap = {};
  productos.forEach(function(p){ prodMap[p.idProducto] = p; });

  rows = rows.map(function(s) {
    var p = prodMap[s.codigoProducto] || {};
    s.descripcion    = p.descripcion   || s.codigoProducto;
    s.stockMinimo    = p.stockMinimo   || 0;
    s.stockMaximo    = p.stockMaximo   || 0;
    s.unidad         = p.unidad        || '';
    s.alertaMinimo   = parseFloat(s.cantidadDisponible) < parseFloat(p.stockMinimo || 0);
    return s;
  });

  if (params.soloAlertas === 'true') {
    rows = rows.filter(function(r){ return r.alertaMinimo; });
  }

  return { ok: true, data: rows };
}

function getStockProducto(codigo) {
  var info = _getStockProducto(codigo);
  var prods = _sheetToObjects(getSheet('PRODUCTOS'));
  var prod  = prods.find(function(p){ return p.idProducto === codigo; }) || {};
  return {
    ok: true,
    data: {
      codigo:     codigo,
      descripcion: prod.descripcion || codigo,
      cantidad:   info.cantidad,
      unidad:     prod.unidad || '',
      stockMinimo: prod.stockMinimo || 0,
      alerta:     info.cantidad < (parseFloat(prod.stockMinimo) || 0)
    }
  };
}

function crearProducto(params) {
  var sheet = getSheet('PRODUCTOS');
  var id    = params.idProducto || ('P' + new Date().getTime());

  // Verificar duplicado por código de barras
  if (params.codigoBarra) {
    var existing = _sheetToObjects(sheet);
    var dup = existing.find(function(p){ return p.codigoBarra === params.codigoBarra; });
    if (dup) return { ok: false, error: 'Código de barras ya registrado: ' + dup.idProducto };
  }

  sheet.appendRow([
    id,
    params.codigoBarra        || '',
    params.descripcion        || '',
    params.marca              || '',
    params.idCategoria        || '',
    params.unidad             || 'UNIDAD',
    parseFloat(params.stockMinimo)  || 0,
    parseFloat(params.stockMaximo)  || 0,
    parseFloat(params.precioCompra) || 0,
    params.esEnvasable        || '0',
    params.codigoProductoBase || '',
    parseFloat(params.factorConversion)  || '',
    parseFloat(params.mermaEsperadaPct)  || '',
    params.foto               || '',
    '1',
    new Date()
  ]);

  // Inicializar stock en 0
  _actualizarStock(id, 0);

  return { ok: true, data: { idProducto: id } };
}

function actualizarProducto(params) {
  var sheet = getSheet('PRODUCTOS');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === params.idProducto || data[i][1] === params.codigoBarra) {
      var campos = ['codigoBarra','descripcion','marca','idCategoria','unidad',
                    'stockMinimo','stockMaximo','precioCompra','esEnvasable',
                    'codigoProductoBase','factorConversion','mermaEsperadaPct',
                    'foto','estado'];
      campos.forEach(function(campo) {
        if (params[campo] !== undefined) {
          var col = headers.indexOf(campo);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(params[campo]);
        }
      });
      return { ok: true };
    }
  }
  return { ok: false, error: 'Producto no encontrado' };
}

// ── Lotes de vencimiento ────────────────────────────────────
function getLotesVencimiento(params) {
  var rows = _sheetToObjects(getSheet('LOTES_VENCIMIENTO'));
  var hoy  = new Date();

  if (params.codigoProducto) {
    rows = rows.filter(function(r){ return r.codigoProducto === params.codigoProducto; });
  }
  if (params.soloActivos === 'true') {
    rows = rows.filter(function(r){
      return r.estado === 'ACTIVO' && parseFloat(r.cantidadActual) > 0;
    });
  }
  if (params.proximosVencer) {
    var dias = parseInt(params.proximosVencer) || 30;
    var limite = new Date(hoy); limite.setDate(hoy.getDate() + dias);
    rows = rows.filter(function(r){
      return r.fechaVencimiento && new Date(r.fechaVencimiento) <= limite;
    });
  }

  // Calcular días restantes
  rows = rows.map(function(r) {
    if (r.fechaVencimiento) {
      r.diasRestantes = Math.ceil((new Date(r.fechaVencimiento) - hoy) / (1000*60*60*24));
    }
    return r;
  });

  rows.sort(function(a,b){ return (a.diasRestantes || 999) - (b.diasRestantes || 999); });

  return { ok: true, data: rows };
}

// ── Producto Nuevo ──────────────────────────────────────────
function getProductosNuevos(params) {
  var rows = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
  if (params.estado) rows = rows.filter(function(r){ return r.estado === params.estado; });
  return { ok: true, data: rows };
}

function registrarProductoNuevo(params) {
  var sheet = getSheet('PRODUCTO_NUEVO');
  var id    = _generateId('PN');
  sheet.appendRow([
    id,
    params.idGuia            || '',
    params.marca             || '',
    params.descripcion       || '',
    params.codigoBarra       || '',
    params.idCategoria       || '',
    params.unidad            || '',
    parseFloat(params.cantidad) || 0,
    params.fechaVencimiento  || '',
    params.foto              || '',
    'PENDIENTE',
    params.usuario           || '',
    new Date(),
    '',
    ''
  ]);
  return { ok: true, data: { idProductoNuevo: id } };
}

function aprobarProductoNuevo(params) {
  var sheet  = getSheet('PRODUCTO_NUEVO');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idProductoNuevo');
  var idxEst = hdrs.indexOf('estado');
  var idxApb = hdrs.indexOf('aprobadoPor');
  var idxFAp = hdrs.indexOf('fechaAprobacion');

  var fila = -1;
  var row;
  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === params.idProductoNuevo) {
      fila = i + 1;
      row  = data[i];
      break;
    }
  }
  if (fila < 0) return { ok: false, error: 'ProductoNuevo no encontrado' };

  // Crear en PRODUCTOS
  var prodData = {};
  hdrs.forEach(function(h, idx){ prodData[h] = row[idx]; });

  var resultCrear = crearProducto({
    codigoBarra:    prodData.codigoBarra,
    descripcion:    prodData.descripcion,
    marca:          prodData.marca,
    idCategoria:    prodData.idCategoria || params.idCategoria,
    unidad:         prodData.unidad      || params.unidad || 'UNIDAD',
    stockMinimo:    params.stockMinimo   || 0,
    stockMaximo:    params.stockMaximo   || 0,
    precioCompra:   params.precioCompra  || 0,
    esEnvasable:    params.esEnvasable   || '0',
    codigoProductoBase: params.codigoProductoBase || '',
    factorConversion:   params.factorConversion   || '',
    mermaEsperadaPct:   params.mermaEsperadaPct   || ''
  });

  if (!resultCrear.ok) return resultCrear;

  // Actualizar estado del ProductoNuevo
  sheet.getRange(fila, idxEst + 1).setValue('APROBADO');
  sheet.getRange(fila, idxApb + 1).setValue(params.aprobadoPor || params.usuario);
  sheet.getRange(fila, idxFAp + 1).setValue(new Date());

  return { ok: true, data: { idProducto: resultCrear.data.idProducto } };
}

// ── Mermas ──────────────────────────────────────────────────
function getMermas(params) {
  var rows = _sheetToObjects(getSheet('MERMAS'));
  if (params.estado)  rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.codigo)  rows = rows.filter(function(r){ return r.codigoProducto === params.codigo; });
  if (params.limit)   rows = rows.slice(0, parseInt(params.limit));
  return { ok: true, data: rows };
}

function registrarMerma(params) {
  var sheet = getSheet('MERMAS');
  var id    = _generateId('M');
  var cant  = parseFloat(params.cantidadOriginal) || 0;

  sheet.appendRow([
    id,
    new Date(),
    params.origen          || 'ALMACEN',
    params.codigoProducto,
    params.idLote          || '',
    cant,
    cant,                    // pendiente = original al inicio
    params.motivo          || '',
    params.usuario         || '',
    params.idGuia          || '',
    'PENDIENTE'
  ]);

  // Descontar del stock si el origen ya está en almacén
  if (params.descontarStock === true || params.descontarStock === 'true') {
    _actualizarStock(params.codigoProducto, -cant);

    // Crear guía de salida por merma
    var resultG = crearGuia({
      tipo:       'SALIDA_MERMA',
      usuario:    params.usuario,
      comentario: 'Merma: ' + (params.motivo || '')
    });
    if (resultG.ok) {
      agregarDetalleGuia({
        idGuia:           resultG.data.idGuia,
        codigoProducto:   params.codigoProducto,
        cantidadEsperada: cant,
        cantidadRecibida: cant
      });
      cerrarGuia(resultG.data.idGuia, params.usuario);
    }
  }

  return { ok: true, data: { idMerma: id } };
}

// ── Auditorias ──────────────────────────────────────────────
function getAuditorias(params) {
  var rows = _sheetToObjects(getSheet('AUDITORIAS'));
  if (params.estado)  rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.usuario) rows = rows.filter(function(r){ return r.usuario === params.usuario; });
  return { ok: true, data: rows };
}

function asignarAuditoria(params) {
  var sheet = getSheet('AUDITORIAS');
  var id    = _generateId('AUD');
  var stockActual = _getStockProducto(params.codigoProducto).cantidad;

  sheet.appendRow([
    id, new Date(), params.codigoProducto, params.usuario,
    stockActual, 0, 0, '', '', 'ASIGNADA', ''
  ]);
  return { ok: true, data: { idAuditoria: id } };
}

function ejecutarAuditoria(params) {
  var sheet   = getSheet('AUDITORIAS');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxId   = hdrs.indexOf('idAuditoria');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === params.idAuditoria) {
      var stockSis  = parseFloat(data[i][hdrs.indexOf('stockSistema')]) || 0;
      var stockFis  = parseFloat(params.stockFisico) || 0;
      var diff      = stockFis - stockSis;
      var resultado = Math.abs(diff) <= 0.5 ? 'OK' : 'DIFERENCIA';

      sheet.getRange(i + 1, hdrs.indexOf('stockFisico')    + 1).setValue(stockFis);
      sheet.getRange(i + 1, hdrs.indexOf('diferencia')     + 1).setValue(diff);
      sheet.getRange(i + 1, hdrs.indexOf('resultado')      + 1).setValue(resultado);
      sheet.getRange(i + 1, hdrs.indexOf('observacion')    + 1).setValue(params.observacion || '');
      sheet.getRange(i + 1, hdrs.indexOf('estado')         + 1).setValue('EJECUTADA');
      sheet.getRange(i + 1, hdrs.indexOf('fechaEjecucion') + 1).setValue(new Date());

      return { ok: true, data: { idAuditoria: params.idAuditoria, diferencia: diff, resultado: resultado } };
    }
  }
  return { ok: false, error: 'Auditoría no encontrada' };
}

// ── Ajustes ─────────────────────────────────────────────────
function getAjustes(params) {
  var rows = _sheetToObjects(getSheet('AJUSTES'));
  if (params.codigo) rows = rows.filter(function(r){ return r.codigoProducto === params.codigo; });
  return { ok: true, data: rows };
}

function crearAjuste(params) {
  var sheet  = getSheet('AJUSTES');
  var id     = _generateId('AJ');
  var tipo   = params.tipoAjuste === 'INC' ? 'INC' : 'DEC';
  var cant   = parseFloat(params.cantidadAjuste) || 0;
  var delta  = tipo === 'INC' ? cant : -cant;

  sheet.appendRow([
    id, params.codigoProducto, tipo, cant,
    params.motivo || '', params.usuario || '',
    params.idAuditoria || '', new Date()
  ]);

  _actualizarStock(params.codigoProducto, delta);

  return { ok: true, data: { idAjuste: id, stockNuevo: _getStockProducto(params.codigoProducto).cantidad } };
}

// ── Proveedores ─────────────────────────────────────────────
function getProveedores(params) {
  var rows = _sheetToObjects(getSheet('PROVEEDORES'));
  if (params.estado) rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.q) {
    var q = params.q.toLowerCase();
    rows = rows.filter(function(r){
      return (r.nombre || '').toLowerCase().indexOf(q) >= 0 ||
             (r.ruc    || '').indexOf(q) >= 0;
    });
  }
  return { ok: true, data: rows };
}

function crearProveedor(params) {
  var sheet = getSheet('PROVEEDORES');
  var id    = _generateId('PROV');
  sheet.appendRow([
    id, params.nombre, params.ruc || '', params.imagen || '',
    params.telefono || '', params.banco || '', params.numeroCuenta || '',
    params.cci || '', params.email || '',
    params.diaPedido || '', params.diaPago || '', params.diaEntrega || '',
    params.formaPago || 'CONTADO', params.plazoCredito || 0,
    params.responsable || '', params.categoriaProducto || '', '1'
  ]);
  return { ok: true, data: { idProveedor: id } };
}

function actualizarProveedor(params) {
  var sheet = getSheet('PROVEEDORES');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === params.idProveedor) {
      var campos = ['nombre','ruc','telefono','banco','numeroCuenta','cci','email',
                    'diaPedido','diaPago','diaEntrega','formaPago','plazoCredito',
                    'responsable','categoriaProducto','estado'];
      campos.forEach(function(c){
        if (params[c] !== undefined) {
          var col = hdrs.indexOf(c);
          if (col >= 0) sheet.getRange(i+1, col+1).setValue(params[c]);
        }
      });
      return { ok: true };
    }
  }
  return { ok: false, error: 'Proveedor no encontrado' };
}

// ── Preingresos ─────────────────────────────────────────────
function getPreingresos(params) {
  var rows = _sheetToObjects(getSheet('PREINGRESOS'));
  if (params.estado)     rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.idProveedor) rows = rows.filter(function(r){ return r.idProveedor === params.idProveedor; });
  return { ok: true, data: rows };
}

function crearPreingreso(params) {
  var sheet = getSheet('PREINGRESOS');
  var id    = _generateId('PI');
  sheet.appendRow([
    id, new Date(), params.idProveedor || '', params.usuario || '',
    params.numeroFactura || '', parseFloat(params.monto) || 0,
    params.comprobante || '', params.fotos || '', params.comentario || '',
    params.etiqueta || '', 'PENDIENTE', ''
  ]);
  return { ok: true, data: { idPreingreso: id } };
}

function aprobarPreingreso(params) {
  var sheet = getSheet('PREINGRESOS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === params.idPreingreso) {
      // Crear guía de ingreso
      var resultGuia = crearGuia({
        tipo:          'INGRESO_PROVEEDOR',
        usuario:       params.usuario || data[i][hdrs.indexOf('usuario')],
        idProveedor:   data[i][hdrs.indexOf('idProveedor')],
        idPreingreso:  params.idPreingreso,
        comentario:    data[i][hdrs.indexOf('comentario')]
      });

      sheet.getRange(i + 1, hdrs.indexOf('estado') + 1).setValue('PROCESADO');
      if (resultGuia.ok) {
        sheet.getRange(i + 1, hdrs.indexOf('idGuia') + 1).setValue(resultGuia.data.idGuia);
      }

      return { ok: true, data: { idGuia: resultGuia.data ? resultGuia.data.idGuia : null } };
    }
  }
  return { ok: false, error: 'Preingreso no encontrado' };
}
