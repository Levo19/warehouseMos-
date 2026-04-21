// ============================================================
// warehouseMos — Productos.gs
// CRUD productos, stock, lotes y producto nuevo
// ============================================================

function getProductos(params) {
  var rows = _sheetToObjects(getProductosSheet());
  if (params.categoria) rows = rows.filter(function(r){ return r.idCategoria === params.categoria; });
  if (params.estado)    rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
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
  var rows = _sheetToObjects(getProductosSheet());
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
  var productos = _sheetToObjects(getProductosSheet());
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
  var prods = _sheetToObjects(getProductosSheet());
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
  var sheet = getProductosSheet();
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
  var sheet = getProductosSheet();
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === params.idProducto || data[i][1] === params.codigoBarra) {
      var campos = ['codigoBarra','descripcion','marca','idCategoria','unidad',
                    'stockMinimo','stockMaximo','precioCompra','esEnvasable',
                    'codigoProductoBase','factorConversion','factorConversionBase',
                    'mermaEsperadaPct','foto','estado'];
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
  if (params.estado) rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
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

// ── Historial de stock por producto ─────────────────────────
// Devuelve los movimientos (GUIA_DETALLE JOIN GUIAS) del producto,
// enriquecidos con balance corriente, ordenados DESC por fecha.
function getHistorialStock(params) {
  var raw = String(params.codigoProducto || '');
  if (!raw) return { ok: false, error: 'codigoProducto requerido' };

  // Acepta barcode único o comma-separated para grupos multi-barcode
  var codigos = raw.split(',').map(function(c){ return c.trim(); }).filter(Boolean);
  var codSet  = {};
  codigos.forEach(function(c){ codSet[c] = true; });

  var guias    = _sheetToObjects(getSheet('GUIAS'));
  var detalles = _sheetToObjects(getSheet('GUIA_DETALLE'));
  var guiaMap  = {};
  guias.forEach(function(g) { guiaMap[g.idGuia] = g; });

  // Movimientos de guías
  var guiaMovs = detalles
    .filter(function(d) { return codSet[String(d.codigoProducto)]; })
    .map(function(d) {
      var g    = guiaMap[d.idGuia] || {};
      var tipo = String(g.tipo || '').toUpperCase();
      return {
        idGuia:    d.idGuia,
        fecha:     g.fecha  || d.fecha || '',
        tipo:      g.tipo   || '—',
        esIngreso: tipo.indexOf('INGRESO') >= 0 || tipo.indexOf('ENTRADA') >= 0 ||
                   (!tipo.includes('SALIDA') && parseFloat(d.cantidad || 0) > 0),
        cantidad:  Math.abs(parseFloat(d.cantidadReal || d.cantidadEsperada || d.cantidad || 0)),
        usuario:   g.usuario || d.usuario || '—',
        origen:    g.idProveedor || g.destino || '',
        estado:    g.estado || '',
        fuente:    'guia'
      };
    })
    .filter(function(m){ return m.cantidad > 0; });

  // Ajustes
  var ajusteMovs = [];
  try {
    var ajSheet = getSheet('AJUSTES');
    if (ajSheet) {
      _sheetToObjects(ajSheet)
        .filter(function(a){ return codSet[String(a.codigoProducto)]; })
        .forEach(function(a){
          var cant = Math.abs(parseFloat(a.cantidadAjuste || 0));
          var tAj  = String(a.tipoAjuste || '').toUpperCase();
          if (cant > 0) ajusteMovs.push({
            idGuia:    a.idAjuste || '',
            fecha:     a.fecha   || '',
            tipo:      'Ajuste ' + (a.tipoAjuste || ''),
            esIngreso: tAj === 'INC' || tAj === 'INI',
            cantidad:  cant,
            usuario:   a.usuario || '—',
            origen:    a.motivo  || '',
            estado:    '',
            fuente:    'ajuste'
          });
        });
    }
  } catch(e) {}

  var todos = guiaMovs.concat(ajusteMovs)
    .sort(function(a, b){ return new Date(b.fecha) - new Date(a.fecha); });

  return { ok: true, data: todos };
}

// ── Imprimir historial de stock (ticket ESC/POS via PrintNode) ─
function imprimirHistorialStock(params) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  // Siempre resuelve la impresora TICKET de ALMACEN desde MOS — no depende del frontend
  var printerId;
  try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
  catch(e) { return { ok: false, error: e.message }; }

  var tz     = Session.getScriptTimeZone();
  var ahora  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var nombre = String(params.codigoProducto || '');
  var texto  = String(params.texto || '');

  // Si no vino texto pre-formateado, armar uno básico
  if (!texto) {
    texto = [
      '================================',
      '      HISTORIAL DE STOCK',
      '   ALMACEN CENTRAL - MOS',
      '================================',
      'Codigo   : ' + nombre,
      'Generado : ' + ahora,
      '================================',
      ''
    ].join('\n');
  }

  // ESC/POS: init + buzzer (1 beep) + texto + 3 avances + corte automático
  var esc    = '\x1b\x40';              // ESC @ — init impresora
  var buzzer = '\x1b\x42\x01\x01';     // ESC B 1 1 — 1 beep, 100 ms
  var feed   = '\n\n\n';               // 3 avances de papel
  var corte  = '\x1d\x56\x00';        // GS V 0 — corte automático completo
  var rawText = esc + buzzer + texto + feed + corte;

  var payload = {
    printerId:   parseInt(printerId),
    title:       'Historial Stock ' + nombre,
    contentType: 'raw_base64',
    content:     Utilities.base64Encode(rawText),
    source:      'warehouseMos'
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:  'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload:            JSON.stringify(payload),
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

// ── Mermas ──────────────────────────────────────────────────
function getMermas(params) {
  var rows = _sheetToObjects(getSheet('MERMAS'));
  if (params.estado)  rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  if (params.codigo)  rows = rows.filter(function(r){ return r.codigoProducto === params.codigo; });
  if (params.limit)   rows = rows.slice(0, parseInt(params.limit));
  return { ok: true, data: rows };
}

function registrarMerma(params) {
  var sheet = getSheet('MERMAS');
  var id    = _generateId('M');
  var cant  = parseFloat(params.cantidadOriginal) || 0;

  if (params.idSesion) registrarActividad(params.idSesion, 'MERMA_REGISTRADA', 1);

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
  if (params.estado)  rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
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
      if (params.idSesion) registrarActividad(params.idSesion, 'AUDITORIA_EJECUTADA', 1);
      return { ok: true, data: { idAuditoria: params.idAuditoria, diferencia: diff, resultado: resultado } };
    }
  }
  return { ok: false, error: 'Auditoría no encontrada' };
}

// ── Auditoría directa desde módulo Productos ────────────────
// Registra conteo físico, crea STOCK si no existe (texto),
// crea AJUSTE si hay diferencia (texto). Todo preserva codigoBarra como texto.
function auditarProducto(params) {
  var codigoBarra = String(params.codigoBarra || '').trim();
  if (!codigoBarra) return { ok: false, error: 'codigoBarra requerido' };

  var stockFisico = parseFloat(params.stockFisico);
  if (isNaN(stockFisico) || stockFisico < 0)
    return { ok: false, error: 'stockFisico inválido' };

  // ── 1. Stock sistema actual ────────────────────────────────
  var stockInfo    = _getStockProducto(codigoBarra);
  var stockSistema = stockInfo.cantidad;
  var diff         = stockFisico - stockSistema;
  var resultado    = Math.abs(diff) <= 0.5 ? 'OK' : 'DIFERENCIA';

  // ── 2. Registrar en AUDITORIAS (codigoBarra como texto) ───
  var audSheet  = getSheet('AUDITORIAS');
  var audHdrs   = audSheet.getRange(1, 1, 1, audSheet.getLastColumn()).getValues()[0];
  var audColCod = audHdrs.indexOf('codigoProducto') + 1; // 1-based
  var audId     = _generateId('AUD');
  var audVals   = [audId, new Date(), codigoBarra, String(params.usuario || ''),
                   stockSistema, stockFisico, diff, resultado,
                   String(params.observacion || ''), 'EJECUTADA', new Date()];
  var audNext   = audSheet.getLastRow() + 1;
  var audRange  = audSheet.getRange(audNext, 1, 1, audVals.length);
  if (audColCod > 0) audSheet.getRange(audNext, audColCod).setNumberFormat('@');
  // Cols 2 (fechaAsignacion) y 11 (fechaEjecucion) con hora
  audSheet.getRange(audNext, 2).setNumberFormat('dd/MM/yyyy HH:mm');
  audSheet.getRange(audNext, 11).setNumberFormat('dd/MM/yyyy HH:mm');
  audRange.setValues([audVals]);

  // ── 3. STOCK: crear si no existe, actualizar si hay diferencia ──
  var ajSheet  = getSheet('AJUSTES');
  var ajHdrs   = ajSheet.getRange(1, 1, 1, ajSheet.getLastColumn()).getValues()[0];
  var ajColCod = ajHdrs.indexOf('codigoProducto') + 1;

  function _writeAjuste(tipo, cant, motivo) {
    var ajId   = _generateId('AJ');
    var ajVals = [ajId, codigoBarra, tipo, cant, motivo,
                  String(params.usuario || ''), audId, new Date()];
    var ajNext = ajSheet.getLastRow() + 1;
    if (ajColCod > 0) ajSheet.getRange(ajNext, ajColCod).setNumberFormat('@');
    ajSheet.getRange(ajNext, 8).setNumberFormat('dd/MM/yyyy HH:mm');
    ajSheet.getRange(ajNext, 1, 1, ajVals.length).setValues([ajVals]);
    return ajId;
  }

  if (stockInfo.fila < 0) {
    // Sin registro previo → stock inicial, siempre registrar en AJUSTES como INI
    if (stockFisico > 0) _writeAjuste('INI', stockFisico, 'Stock inicial (auditoria)');
    // Crear fila STOCK
    var stSheet = getSheet('STOCK');
    var stNext  = stSheet.getLastRow() + 1;
    var stVals  = ['STK' + new Date().getTime(), codigoBarra, stockFisico, new Date()];
    stSheet.getRange(stNext, 2).setNumberFormat('@');
    stSheet.getRange(stNext, 4).setNumberFormat('dd/MM/yyyy HH:mm');
    stSheet.getRange(stNext, 1, 1, stVals.length).setValues([stVals]);
  } else if (Math.abs(diff) > 0.5) {
    // Diferencia real → AJUSTE INC/DEC + actualizar STOCK
    _writeAjuste(diff > 0 ? 'INC' : 'DEC', Math.abs(diff), 'Auditoria diaria');
    _actualizarStock(codigoBarra, diff);
  }
  // Si diff ≤ 0.5: stock cuadra, solo queda en AUDITORIAS, sin tocar AJUSTES

  return {
    ok: true,
    data: {
      idAuditoria:  audId,
      stockSistema: stockSistema,
      stockFisico:  stockFisico,
      diferencia:   diff,
      resultado:    resultado
    }
  };
}

// ── Ajustes ─────────────────────────────────────────────────
function getAjustes(params) {
  var rows = _sheetToObjects(getSheet('AJUSTES'));
  if (params.codigo) rows = rows.filter(function(r){ return r.codigoProducto === params.codigo; });
  return { ok: true, data: rows };
}

function crearAjuste(params) {
  var sheet     = getSheet('AJUSTES');
  var hdrs      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colCod    = hdrs.indexOf('codigoProducto') + 1;
  var id        = _generateId('AJ');
  var tipo      = params.tipoAjuste === 'INC' ? 'INC' : 'DEC';
  var cant      = parseFloat(params.cantidadAjuste) || 0;
  var delta     = tipo === 'INC' ? cant : -cant;
  var codigoBarra = String(params.codigoProducto || '');
  var ajVals    = [id, codigoBarra, tipo, cant,
                   String(params.motivo || ''), String(params.usuario || ''),
                   String(params.idAuditoria || ''), new Date()];
  var nextRow   = sheet.getLastRow() + 1;
  if (colCod > 0) sheet.getRange(nextRow, colCod).setNumberFormat('@');
  sheet.getRange(nextRow, 8).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(nextRow, 1, 1, ajVals.length).setValues([ajVals]);

  _actualizarStock(codigoBarra, delta);

  return { ok: true, data: { idAjuste: id, stockNuevo: _getStockProducto(codigoBarra).cantidad } };
}

// ── Proveedores ─────────────────────────────────────────────
function getProveedores(params) {
  var rows = _sheetToObjects(getProveedoresSheet());
  if (params.estado) rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
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
  var sheet = getProveedoresSheet();
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
  var sheet = getProveedoresSheet();
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
  if (params.estado)     rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  if (params.idProveedor) rows = rows.filter(function(r){ return r.idProveedor === params.idProveedor; });
  return { ok: true, data: rows };
}

function crearPreingreso(params) {
  var sheet = getSheet('PREINGRESOS');
  // Usar ID del cliente si viene (previene duplicados por retry)
  var id = params.idPreingreso || _generateId('PI');

  // Idempotencia: si ya existe la fila, devolver OK sin insertar
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) return { ok: true, data: { idPreingreso: id } };
  }

  // Escribir por nombre de columna — funciona con 9 o 12 cols (schema viejo/nuevo)
  var hdrs = data[0];
  var row  = new Array(hdrs.length).fill('');
  function _set(col, val) { var i = hdrs.indexOf(col); if (i >= 0) row[i] = val; }
  _set('idPreingreso', id);
  _set('fecha',        new Date());
  _set('idProveedor',  params.idProveedor  || '');
  _set('cargadores',   params.cargadores   || '');
  _set('usuario',      params.usuario      || '');
  _set('monto',        parseFloat(params.monto) || 0);
  _set('fotos',        params.fotos        || '');
  _set('comentario',   params.comentario   || '');
  _set('estado',       'PENDIENTE');
  _set('idGuia',       '');
  sheet.appendRow(row);
  return { ok: true, data: { idPreingreso: id } };
}

// ── Subir foto de preingreso a Drive ────────────────────────
function _getOrCreateFolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

/**
 * Ejecutar una vez desde el editor de GAS:
 * 1. Crea carpetas imagenes/preingresos junto al Spreadsheet
 * 2. Las comparte como ANYONE_WITH_LINK VIEW
 * 3. Guarda los IDs en Script Properties para uso rápido
 */
function setupPreingresosFolders() {
  var ssId     = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ssFile   = DriveApp.getFileById(ssId);
  var parent   = ssFile.getParents().next();

  var imgFolder = _getOrCreateFolder(parent, 'imagenes');
  var preFolder = _getOrCreateFolder(imgFolder, 'preingresos');

  imgFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  preFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  PropertiesService.getScriptProperties().setProperties({
    'FOTOS_IMG_FOLDER_ID': imgFolder.getId(),
    'FOTOS_PRE_FOLDER_ID': preFolder.getId()
  });

  Logger.log('setupPreingresosFolders OK. preFolder ID: ' + preFolder.getId());
  return { ok: true, preFolderId: preFolder.getId() };
}

function subirFotoPreingreso(params) {
  var idPreingreso = String(params.idPreingreso || '');
  var fotoBase64   = String(params.fotoBase64   || '');
  var mimeType     = String(params.mimeType     || 'image/jpeg');
  var indice       = parseInt(params.indice)    || 1;

  if (!idPreingreso || !fotoBase64) return { ok: false, error: 'idPreingreso y fotoBase64 son requeridos' };

  try {
    // Usar carpeta guardada en Properties (más rápido y sin permisos adicionales)
    var preFolderId = PropertiesService.getScriptProperties().getProperty('FOTOS_PRE_FOLDER_ID');
    var preFolder;
    if (preFolderId) {
      preFolder = DriveApp.getFolderById(preFolderId);
    } else {
      // Fallback: recorrer desde el Spreadsheet (requiere que setupPreingresosFolders se haya ejecutado al menos una vez)
      var ssId   = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      var parent = DriveApp.getFileById(ssId).getParents().next();
      preFolder  = _getOrCreateFolder(_getOrCreateFolder(parent, 'imagenes'), 'preingresos');
      preFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    var piFolder = _getOrCreateFolder(preFolder, idPreingreso);
    piFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var ext      = mimeType === 'image/png' ? 'png' : 'jpg';
    var fileName = idPreingreso + '_' + indice + '.' + ext;
    var blob     = Utilities.newBlob(Utilities.base64Decode(fotoBase64), mimeType, fileName);
    var file     = piFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    return { ok: true, data: { url: url, fileId: file.getId() } };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function eliminarFotoDrive(params) {
  var fileId = String(params.fileId || '');
  if (!fileId) return { ok: false, error: 'fileId requerido' };
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function actualizarPreingreso(params) {
  var sheet        = getSheet('PREINGRESOS');
  var idPreingreso = String(params.idPreingreso || '');
  if (!idPreingreso) return { ok: false, error: 'idPreingreso requerido' };

  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== idPreingreso) continue;

    // Campos editables
    var editable = { idProveedor: true, monto: true, comentario: true, fotos: true };
    Object.keys(editable).forEach(function(key) {
      if (params[key] === undefined) return;
      var col = hdrs.indexOf(key);
      if (col < 0) return;
      var val = key === 'monto' ? (parseFloat(params[key]) || 0) : String(params[key]);
      sheet.getRange(i + 1, col + 1).setValue(val);
    });

    // Propagar a guía de ingreso vinculada (si existe)
    var colGuia = hdrs.indexOf('idGuia');
    var idGuia  = colGuia >= 0 ? String(data[i][colGuia] || '') : '';
    if (idGuia) {
      try {
        var gs    = getSheet('GUIAS');
        var gData = gs.getDataRange().getValues();
        var gHdrs = gData[0];
        for (var j = 1; j < gData.length; j++) {
          if (gData[j][0] !== idGuia) continue;
          if (params.idProveedor !== undefined) {
            var cP = gHdrs.indexOf('idProveedor');
            if (cP >= 0) gs.getRange(j + 1, cP + 1).setValue(String(params.idProveedor));
          }
          if (params.comentario !== undefined) {
            var cC = gHdrs.indexOf('comentario');
            if (cC >= 0) gs.getRange(j + 1, cC + 1).setValue(String(params.comentario));
          }
          break;
        }
      } catch(e) { /* non-fatal */ }
    }
    return { ok: true };
  }
  return { ok: false, error: 'Preingreso no encontrado: ' + idPreingreso };
}

function actualizarFotosPreingreso(params) {
  var sheet        = getSheet('PREINGRESOS');
  var idPreingreso = String(params.idPreingreso || '');
  var fotos        = String(params.fotos        || '');

  if (!idPreingreso) return { ok: false, error: 'idPreingreso requerido' };

  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var colFotos = hdrs.indexOf('fotos');
  if (colFotos < 0) return { ok: false, error: 'Columna fotos no encontrada' };

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idPreingreso) {
      sheet.getRange(i + 1, colFotos + 1).setValue(fotos);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Preingreso no encontrado: ' + idPreingreso };
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

// ============================================================
// Fotos + comentario de GUÍAS
// ============================================================

/**
 * Crea carpeta imagenes/guias y guarda ID en Script Properties.
 * Ejecutar una vez desde el editor de GAS.
 */
function setupGuiasFolders() {
  var ssId      = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var parent    = DriveApp.getFileById(ssId).getParents().next();
  var imgFolder = _getOrCreateFolder(parent, 'imagenes');
  var gFolder   = _getOrCreateFolder(imgFolder, 'guias');
  imgFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  gFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  PropertiesService.getScriptProperties().setProperty('FOTOS_GUIA_FOLDER_ID', gFolder.getId());
  Logger.log('setupGuiasFolders OK. gFolder ID: ' + gFolder.getId());
  return { ok: true, guiaFolderId: gFolder.getId() };
}

function _actualizarColumnaGuia(idGuia, campo, valor) {
  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var col   = hdrs.indexOf(campo);
  if (col < 0) return;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idGuia) {
      sheet.getRange(i + 1, col + 1).setValue(valor);
      return;
    }
  }
}

function subirFotoGuia(params) {
  var idGuia     = String(params.idGuia     || '');
  var fotoBase64 = String(params.fotoBase64 || '');
  var mimeType   = String(params.mimeType   || 'image/jpeg');
  if (!idGuia || !fotoBase64) return { ok: false, error: 'idGuia y fotoBase64 requeridos' };
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('FOTOS_GUIA_FOLDER_ID');
    var folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      var ssId   = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      var parent = DriveApp.getFileById(ssId).getParents().next();
      folder = _getOrCreateFolder(_getOrCreateFolder(parent, 'imagenes'), 'guias');
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    var ext  = mimeType === 'image/png' ? 'png' : 'jpg';
    var name = idGuia + '.' + ext;
    // Borrar foto anterior si existe
    var existing = folder.getFilesByName(name);
    while (existing.hasNext()) { existing.next().setTrashed(true); }
    var blob = Utilities.newBlob(Utilities.base64Decode(fotoBase64), mimeType, name);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    _actualizarColumnaGuia(idGuia, 'foto', url);
    return { ok: true, data: { url: url, fileId: file.getId() } };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function copiarFotoDePreingreso(params) {
  var idGuia       = String(params.idGuia       || '');
  var idPreingreso = String(params.idPreingreso || '');
  if (!idGuia || !idPreingreso) return { ok: false, error: 'idGuia e idPreingreso requeridos' };
  try {
    var preFolderId  = PropertiesService.getScriptProperties().getProperty('FOTOS_PRE_FOLDER_ID');
    var guiaFolderId = PropertiesService.getScriptProperties().getProperty('FOTOS_GUIA_FOLDER_ID');
    if (!preFolderId || !guiaFolderId) return { ok: false, error: 'Carpetas no configuradas. Ejecuta setupPreingresosFolders() y setupGuiasFolders()' };
    var piFolder = DriveApp.getFolderById(preFolderId).getFoldersByName(idPreingreso);
    if (!piFolder.hasNext()) return { ok: false, error: 'Carpeta del preingreso no encontrada' };
    var files = piFolder.next().getFiles();
    if (!files.hasNext()) return { ok: false, error: 'Sin fotos en el preingreso' };
    var srcFile    = files.next();
    var guiaFolder = DriveApp.getFolderById(guiaFolderId);
    var copyName   = idGuia + '.jpg';
    // Borrar copia anterior si existe
    var existentes = guiaFolder.getFilesByName(copyName);
    while (existentes.hasNext()) { existentes.next().setTrashed(true); }
    var copy = srcFile.makeCopy(copyName, guiaFolder);
    copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + copy.getId() + '&sz=w800';
    _actualizarColumnaGuia(idGuia, 'foto', url);
    return { ok: true, data: { url: url, fileId: copy.getId() } };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function actualizarGuia(params) {
  var idGuia = String(params.idGuia || '');
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== idGuia) continue;
    ['tipo', 'idProveedor', 'idZona', 'numeroDocumento', 'comentario', 'foto'].forEach(function(key) {
      if (params[key] === undefined) return;
      var col = hdrs.indexOf(key);
      if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(String(params[key]));
    });
    // Propagar comentario al preingreso vinculado
    if (params.comentario !== undefined) {
      var colPre = hdrs.indexOf('idPreingreso');
      var idPre  = colPre >= 0 ? String(data[i][colPre] || '') : '';
      if (idPre) {
        try {
          var ps    = getSheet('PREINGRESOS');
          var pData = ps.getDataRange().getValues();
          var pHdrs = pData[0];
          var cC    = pHdrs.indexOf('comentario');
          for (var j = 1; j < pData.length; j++) {
            if (pData[j][0] === idPre && cC >= 0) {
              ps.getRange(j + 1, cC + 1).setValue(String(params.comentario));
              break;
            }
          }
        } catch(e) { /* non-fatal */ }
      }
    }
    return { ok: true };
  }
  return { ok: false, error: 'Guía no encontrada: ' + idGuia };
}

// ============================================================
// imprimirMembrete — ESC/POS 80mm, layout 2 columnas × 3 filas
//
// Columna izquierda (≈34mm): 1-3 barcodes CODE128 + texto EAN
// Columna derecha (≈37mm):   Nombre grande + SKU barcode (si hay varios EAN)
//
// Usa ESC/POS Page Mode (ESC L) para posicionamiento preciso.
// ============================================================
function imprimirMembrete(params) {
  var idProducto = String(params.idProducto || '');
  if (!idProducto) return { ok: false, error: 'idProducto requerido' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  var printerId;
  try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
  catch(e) { return { ok: false, error: e.message }; }

  // ── Producto ───────────────────────────────────────────────
  var productos = _sheetToObjects(getProductosSheet());
  var prod = productos.find(function(p) {
    return p.idProducto === idProducto || String(p.codigoBarra) === idProducto;
  });
  if (!prod) return { ok: false, error: 'Producto no encontrado: ' + idProducto };

  var sku = String(prod.idProducto);

  // ── EANs (principal + equivalencias activas, max 3) ────────
  var altCodes = [];
  try {
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      _sheetToObjects(equivSheet)
        .filter(function(e) {
          return String(e.idProducto) === sku && String(e.activo) === '1' && e.codigoBarra;
        })
        .forEach(function(e) { altCodes.push(String(e.codigoBarra)); });
    }
  } catch(e) {}

  var allEan = [];
  if (prod.codigoBarra) allEan.push(String(prod.codigoBarra));
  altCodes.forEach(function(c) { if (allEan.indexOf(c) < 0) allEan.push(c); });
  if (allEan.length === 0) allEan.push(sku);  // fallback al SKU si no hay EAN
  var numEan      = Math.min(allEan.length, 3);
  var hasMultiple = allEan.length > 1;

  // ── Nombre (word-wrap sin cortar palabras) ─────────────────
  // Columna derecha ≈37mm con fuente doble ancho: ~11 chars/línea
  var descripcion = String(prod.descripcion || sku).toUpperCase();
  var nameLines   = _membWrap(descripcion, 11);

  // ── Constantes ESC/POS ─────────────────────────────────────
  var ESC = '\x1b';
  var GS  = '\x1d';
  var LF  = '\x0a';
  var FF  = '\x0c';  // FormFeed → imprime página y vuelve al modo estándar

  // ── Dimensiones (dots a 203 dpi; 1mm ≈ 8 dots) ────────────
  //   Papel 80mm, zona imprimible ≈ 72mm = 576 dots
  //   Col izquierda: x=0..271   (272 dots ≈ 34mm)
  //   Col derecha:   x=288..575 (288 dots ≈ 36mm)
  var L_W  = 272;
  var R_X  = 288;
  var R_W  = 272;

  // Altura de cada fila según cuántos EAN hay
  // 1 EAN → fila alta (código grande y visible)
  // 2 EAN → fila media
  // 3 EAN → fila compacta
  var ROW_H = numEan === 1 ? 200 : (numEan === 2 ? 130 : 100);
  var totalH = numEan * ROW_H;

  // Altura del código de barras dentro de la fila
  // = altura fila − 30 (texto HRI) − 14 (márgenes)
  var BC_H = Math.max(50, ROW_H - 44);

  // Área del SKU en col derecha (solo si hay múltiples EAN)
  var SKU_H  = hasMultiple ? 88 : 0;
  var NAME_H = totalH - SKU_H;

  // ── Helpers ────────────────────────────────────────────────
  // ESC W: define zona activa en page mode
  function setArea(x, y, w, h) {
    return ESC + 'W' +
      String.fromCharCode(x & 0xff, (x >> 8) & 0xff,
                          y & 0xff, (y >> 8) & 0xff,
                          w & 0xff, (w >> 8) & 0xff,
                          h & 0xff, (h >> 8) & 0xff);
  }

  // CODE128 auto-select ({B), centrado
  function barcode(code, height) {
    var bd = '{B' + code;
    return ESC + '\x61\x01' +                              // centrar
      GS + '\x68' + String.fromCharCode(Math.min(255, height)) +
      GS + '\x77\x02' +                                    // módulo ancho 2
      GS + '\x48\x02' +                                    // HRI debajo
      GS + '\x66\x01' +                                    // HRI font B (pequeña)
      GS + '\x6b\x49' + String.fromCharCode(bd.length) + bd +
      LF;
  }

  // ── Construir ESC/POS ──────────────────────────────────────
  var t = '';
  t += ESC + '\x40';  // inicializar impresora
  t += ESC + 'L';     // entrar a Page Mode

  // ── Columna izquierda: 1-3 barcodes EAN ───────────────────
  for (var i = 0; i < numEan; i++) {
    t += setArea(0, i * ROW_H, L_W, ROW_H);
    t += barcode(allEan[i], BC_H);
  }

  // ── Columna derecha superior: nombre del producto ──────────
  t += setArea(R_X, 0, R_W, NAME_H);
  t += ESC + '\x61\x01';   // centrar
  // Padding vertical aproximado para centrar el bloque de nombre
  var NAME_LINE_H = 48;    // fuente doble en dots (≈6mm)
  var FEED_H      = 24;    // un LF en spacing estándar ≈3mm
  var nameH = Math.min(nameLines.length, 4) * NAME_LINE_H;
  var padLines = Math.max(0, Math.floor((NAME_H - nameH) / 2 / FEED_H));
  for (var p = 0; p < padLines; p++) { t += LF; }
  t += ESC + '\x21\x38';   // font A + bold + doble alto + doble ancho
  for (var ln = 0; ln < Math.min(nameLines.length, 4); ln++) {
    t += nameLines[ln] + LF;
  }
  t += ESC + '\x21\x00';   // normal

  // ── Columna derecha inferior: SKU barcode (si hay varios EAN)
  if (hasMultiple) {
    t += setArea(R_X, NAME_H, R_W, SKU_H);
    t += ESC + '\x61\x01';                    // centrar
    t += ESC + '\x21\x01';                    // font B pequeña
    t += 'SKU GRUPAL' + LF;
    t += ESC + '\x21\x00';
    t += barcode(sku, Math.max(40, SKU_H - 36));
  }

  // ── Imprimir página y salir de page mode ──────────────────
  t += FF;

  // ── Feed antes del corte (cuchilla ≈20mm sobre cabezal) ───
  t += ESC + 'J' + String.fromCharCode(160); // avanzar ≈20mm
  t += GS  + '\x56\x00';                     // corte completo

  // ── Enviar a PrintNode ─────────────────────────────────────
  var payload = {
    printerId:   parseInt(printerId),
    title:       'Membrete ' + sku,
    contentType: 'raw_base64',
    content:     Utilities.base64Encode(t),
    source:      'warehouseMos'
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:  'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload:            JSON.stringify(payload),
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

// Word-wrap sin cortar palabras
function _membWrap(text, maxLen) {
  var words = String(text || '').trim().split(/\s+/);
  var lines = [];
  var cur   = '';
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (!cur) {
      cur = w;
    } else if ((cur + ' ' + w).length <= maxLen) {
      cur += ' ' + w;
    } else {
      lines.push(cur);
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines;
}
