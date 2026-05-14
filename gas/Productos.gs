// ============================================================
// warehouseMos — Productos.gs
// CRUD productos, stock, lotes y producto nuevo
// ============================================================

function getProductos(params) {
  var rows = _sheetToObjects(getProductosSheet());
  if (params.categoria) rows = rows.filter(function(r){ return r.idCategoria === params.categoria; });
  if (params.estado) {
    var estadoBuscado = String(params.estado);
    rows = rows.filter(function(r) {
      // Normalizar: boolean true → "1", boolean false → "0" (Sheets checkbox)
      var estadoNorm = (r.estado === true) ? '1' : (r.estado === false) ? '0' : String(r.estado);
      return estadoNorm === estadoBuscado;
    });
  }
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

  // Enriquecer con stock actual (STOCK usa codigoBarra como clave)
  prod.stockActual = _getStockProducto(prod.codigoBarra || prod.idProducto).cantidad;

  // Lotes vigentes (LOTES usa codigoBarra)
  var lotes = _sheetToObjects(getSheet('LOTES_VENCIMIENTO'));
  prod.lotes = lotes.filter(function(l){
    return (l.codigoProducto === prod.codigoBarra || l.codigoProducto === prod.idProducto) && l.estado === 'ACTIVO' &&
           parseFloat(l.cantidadActual) > 0;
  });

  return { ok: true, data: prod };
}

function getStock(params) {
  var rows = _sheetToObjects(getSheet('STOCK'));
  // Enriquecer con info de producto (STOCK usa codigoBarra como clave)
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p){ if (p.codigoBarra) prodMap[String(p.codigoBarra)] = p; });

  rows = rows.map(function(s) {
    var p = prodMap[String(s.codigoProducto)] || {};
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

function _nextNLEVCode() {
  var sheet  = getSheet('PRODUCTO_NUEVO');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxCod = hdrs.indexOf('codigoBarra');
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var cod = String(data[i][idxCod] || '');
    if (cod.indexOf('NLEV') === 0) {
      var n = parseInt(cod.slice(4), 10) || 0;
      if (n > maxNum) maxNum = n;
    }
  }
  var s = String(maxNum + 1);
  while (s.length < 5) s = '0' + s;
  return 'NLEV' + s;
}

function _subirFotoProductoNuevo(codigoBarra, fotoBase64, mimeType) {
  mimeType = mimeType || 'image/jpeg';
  var folderId = PropertiesService.getScriptProperties().getProperty('FOTOS_PN_FOLDER_ID');
  var folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    var ssId   = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    var parent = DriveApp.getFileById(ssId).getParents().next();
    folder = _getOrCreateFolder(_getOrCreateFolder(parent, 'imagenes'), 'productoimagenes');
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    PropertiesService.getScriptProperties().setProperty('FOTOS_PN_FOLDER_ID', folder.getId());
  }
  var ext      = mimeType === 'image/png' ? 'png' : 'jpg';
  var fileName = codigoBarra + '.' + ext;
  var existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) { existing.next().setTrashed(true); }
  var blob = Utilities.newBlob(Utilities.base64Decode(fotoBase64), mimeType, fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

function getProductosNuevos(params) {
  var rows = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
  if (params.estado) rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  return { ok: true, data: rows };
}

// Productos nuevos APROBADOS dentro de los últimos N días (default 3)
function getProductosNuevosRecientes(params) {
  var dias = parseInt(params && params.dias) || 3;
  var corte = new Date();
  corte.setDate(corte.getDate() - dias);
  var rows = _sheetToObjects(getSheet('PRODUCTO_NUEVO')).filter(function(r){
    if (String(r.estado).toUpperCase() !== 'APROBADO') return false;
    if (!r.fechaAprobacion) return false;
    var f = r.fechaAprobacion instanceof Date ? r.fechaAprobacion : new Date(r.fechaAprobacion);
    return f >= corte;
  });
  // Mapear con tipoAprobacion derivado de la observacion (NUEVO o "EQUIVALENTE de X")
  var resultado = rows.map(function(r){
    var obs = String(r.observacion || '').trim();
    var tipo = obs.indexOf('EQUIVALENTE') === 0 ? 'EQUIVALENTE' : 'NUEVO';
    return {
      idProductoNuevo:  r.idProductoNuevo,
      codigoBarra:      r.codigoBarra,
      descripcion:      r.descripcion,
      foto:             r.foto,
      cantidad:         r.cantidad,
      fechaCreacion:    r.fechaCreacion,
      fechaAprobacion:  r.fechaAprobacion,
      aprobadoPor:      r.aprobadoPor,
      usuario:          r.usuario,
      tipoAprobacion:   tipo,
      observacion:      obs
    };
  }).sort(function(a, b){
    return new Date(b.fechaAprobacion) - new Date(a.fechaAprobacion);
  });
  return { ok: true, data: resultado };
}

// ── F0: productos cambiados desde un timestamp (para reconciliación caso 4) ──
// Retorna PRODUCTO_NUEVO con estado APROBADO cuya fechaAprobacion >= since.
// Frontend usa esto para detectar cambios desde el último polling y actualizar
// los items con badge "🆕 PENDIENTE" → nombre canónico/equivalente real.
function getProductosCambiadosDesde(params) {
  var sinceTs = String(params.since || '').trim();
  var sinceMs = 0;
  if (sinceTs) {
    var d = new Date(sinceTs);
    if (!isNaN(d.getTime())) sinceMs = d.getTime();
  }
  try {
    var rows = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
    var cambios = rows.filter(function(r) {
      var est = String(r.estado || '').toUpperCase();
      if (est !== 'APROBADO' && est !== 'RECHAZADO') return false;
      if (!sinceMs) return true;
      var f = r.fechaAprobacion ? new Date(r.fechaAprobacion).getTime() : 0;
      return f >= sinceMs;
    });
    return { ok: true, data: cambios, generadoEn: new Date().toISOString() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function registrarProductoNuevo(params) {
  var sheet = getSheet('PRODUCTO_NUEVO');
  var id    = _generateId('PN');

  // Generar NLEV si no viene código de barra
  var codigoBarra = String(params.codigoBarra || '').trim();
  if (!codigoBarra) codigoBarra = _nextNLEVCode();

  // Subir foto a Drive si viene base64
  var fotoUrl    = String(params.foto || '');
  var fotoBase64 = String(params.fotoBase64 || '').trim();
  if (fotoBase64) {
    try {
      fotoUrl = _subirFotoProductoNuevo(codigoBarra, fotoBase64, params.mimeType || 'image/jpeg');
    } catch(e) {
      Logger.log('Error subiendo foto PN: ' + e.message);
    }
  }

  sheet.appendRow([
    id,
    params.idGuia           || '',
    params.marca            || '',
    params.descripcion      || '',
    codigoBarra,
    params.idCategoria      || '',
    params.unidad           || '',
    parseFloat(params.cantidad) || 0,
    params.fechaVencimiento || '',
    fotoUrl,
    'PENDIENTE',
    params.usuario          || '',
    new Date(),
    '',
    ''
  ]);

  // Si tiene guía → escribir directo a GUIA_DETALLE sin validar catálogo
  // (el PN aún no está aprobado → no existe en PRODUCTOS_MASTER)
  // AUTO-SUMA: si ya existe detalle (mismo idGuia + codigoBarra, no anulado),
  // sumar cantidad en lugar de duplicar. Previene reintentos y doble-clicks.
  if (params.idGuia) {
    try {
      var detSheet  = getSheet('GUIA_DETALLE');
      var detData   = detSheet.getDataRange().getValues();
      var hdrs      = detData[0];
      var idxIdG    = hdrs.indexOf('idGuia');
      var idxCb     = hdrs.indexOf('codigoProducto');
      var idxRec    = hdrs.indexOf('cantidadRecibida');
      var idxObs    = hdrs.indexOf('observacion');
      var cantNueva = parseFloat(params.cantidad) || 1;
      var sumado    = false;

      for (var dr = 1; dr < detData.length; dr++) {
        if (String(detData[dr][idxIdG]) !== String(params.idGuia)) continue;
        if (String(detData[dr][idxCb]).toUpperCase() !== String(codigoBarra).toUpperCase()) continue;
        if (String(detData[dr][idxObs] || '').toUpperCase() === 'ANULADO') continue;
        // Match: sumar al detalle existente
        var qtyExist = parseFloat(detData[dr][idxRec]) || 0;
        detSheet.getRange(dr + 1, idxRec + 1).setValue(qtyExist + cantNueva);
        sumado = true;
        break;
      }

      if (!sumado) {
        var idDetalle  = _generateId('DET');
        var nextDetRow = detSheet.getLastRow() + 1;
        detSheet.appendRow([
          idDetalle,
          params.idGuia,
          codigoBarra,
          0,
          cantNueva,
          0,
          '',
          'PN_PENDIENTE'
        ]);
        detSheet.getRange(nextDetRow, 3).setNumberFormat('@').setValue(codigoBarra);
      }
    } catch(e) {
      Logger.log('Error al agregar detalle PN a guia: ' + e.message);
    }
  }

  // Push: notificar a MOS que llegó un producto nuevo para revisar
  try {
    var descCorta = (params.descripcion || '').substring(0, 40);
    var qty = parseFloat(params.cantidad) || 0;
    _notificarMOS(
      '🆕 Producto nuevo pendiente',
      descCorta + ' · ' + qty + ' uds · ' + (params.usuario || 'Operador'),
      params.usuario || null,
      'WH_PRODUCTO_NUEVO'
    );
  } catch(eP) { Logger.log('Push PN nuevo: ' + eP.message); }

  return { ok: true, data: { idProductoNuevo: id, codigoBarra: codigoBarra } };
}

function aprobarProductoNuevo(params) {
  var tipo           = String(params.tipo || 'NUEVO').toUpperCase();
  var idGuia         = params.idGuia;
  var codigoOriginal = String(params.codigoOriginal || '').trim();
  var codigoFinal    = String(params.codigoFinal    || '').trim() || codigoOriginal;
  var cantidadFinal  = parseFloat(params.cantidadFinal) || 0;

  // 1) Localizar PN
  var sheet  = getSheet('PRODUCTO_NUEVO');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idProductoNuevo');
  var idxEst = hdrs.indexOf('estado');
  var idxApb = hdrs.indexOf('aprobadoPor');
  var idxFAp = hdrs.indexOf('fechaAprobacion');

  var fila = -1, row;
  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === params.idProductoNuevo) { fila = i + 1; row = data[i]; break; }
  }
  if (fila < 0) return { ok: false, error: 'ProductoNuevo no encontrado' };

  // 2) Si NUEVO: crear en catálogo local de WH (idempotente)
  var idProductoCreado = '';
  if (tipo === 'NUEVO') {
    var prodDataPN = {};
    hdrs.forEach(function(h, idx){ prodDataPN[h] = row[idx]; });
    var resultCrear = crearProducto({
      codigoBarra:        codigoFinal,
      descripcion:        params.descripcion || prodDataPN.descripcion,
      marca:              prodDataPN.marca,
      idCategoria:        params.idCategoria || prodDataPN.idCategoria,
      unidad:             params.unidad      || prodDataPN.unidad || 'UNIDAD',
      stockMinimo:        params.stockMinimo || 0,
      stockMaximo:        params.stockMaximo || 0,
      precioCompra:       params.precioCompra || 0,
      esEnvasable:        params.esEnvasable  || '0',
      codigoProductoBase: params.codigoProductoBase || '',
      factorConversion:   params.factorConversion   || '',
      mermaEsperadaPct:   params.mermaEsperadaPct   || ''
    });
    if (resultCrear && resultCrear.ok) idProductoCreado = resultCrear.data.idProducto;
    // Si falla por duplicado, seguimos
  }

  // 3) Marcar PN como APROBADO
  sheet.getRange(fila, idxEst + 1).setValue('APROBADO');
  sheet.getRange(fila, idxApb + 1).setValue(params.aprobadoPor || 'MOS');
  sheet.getRange(fila, idxFAp + 1).setValue(new Date());
  // Guardar tipoAprobacion en la columna observacion (col 14, idx 13) si existe
  // o en la última columna libre. Usamos formato: "NUEVO" o "EQUIVALENTE de <skuBase>"
  var tipoLabel = tipo === 'EQUIVALENTE' ? ('EQUIVALENTE de ' + (params.skuBase || '')) : 'NUEVO';
  try {
    var idxObs = hdrs.indexOf('observacion');
    if (idxObs >= 0) sheet.getRange(fila, idxObs + 1).setValue(tipoLabel);
  } catch(_){}

  // 4) Actualizar GUIA_DETALLE: buscar por idGuia + codigoOriginal
  var stockDeltaOrig = 0;  // restar del original (si guía cerrada)
  var stockDeltaFin  = 0;  // sumar al final  (si guía cerrada)
  var cantOriginalEnDetalle = 0;
  var guiaCerrada = false;
  var esIngreso   = false;

  if (idGuia) {
    try {
      // Estado de la guía
      var guiasSh   = getSheet('GUIAS');
      var guiasData = guiasSh.getDataRange().getValues();
      var gHdrs     = guiasData[0];
      var iGId      = gHdrs.indexOf('idGuia');
      var iGEst     = gHdrs.indexOf('estado');
      var iGTip     = gHdrs.indexOf('tipo');
      for (var gi = 1; gi < guiasData.length; gi++) {
        if (String(guiasData[gi][iGId]).trim() === String(idGuia).trim()) {
          guiaCerrada = String(guiasData[gi][iGEst] || '').toUpperCase() === 'CERRADA';
          esIngreso   = String(guiasData[gi][iGTip] || '').toUpperCase().indexOf('INGRESO') === 0;
          break;
        }
      }

      var detSheet = getSheet('GUIA_DETALLE');
      var detData  = detSheet.getDataRange().getValues();
      var detHdrs  = detData[0];
      var iDetGuia = detHdrs.indexOf('idGuia');
      var iDetCod  = detHdrs.indexOf('codigoProducto');
      var iDetCant = detHdrs.indexOf('cantidadRecibida');
      var iDetObs  = detHdrs.indexOf('observacion');

      for (var di = 1; di < detData.length; di++) {
        if (String(detData[di][iDetGuia] || '').trim() !== String(idGuia).trim()) continue;
        var codDet = String(detData[di][iDetCod] || '').trim();
        if (codDet !== codigoOriginal) continue;

        cantOriginalEnDetalle = parseFloat(detData[di][iDetCant]) || 0;
        var cantFinal = cantidadFinal > 0 ? cantidadFinal : cantOriginalEnDetalle;

        // Update código si cambió
        if (codigoFinal !== codigoOriginal) {
          detSheet.getRange(di + 1, iDetCod + 1).setNumberFormat('@').setValue(codigoFinal);
        }
        // Update cantidad si cambió
        if (cantFinal !== cantOriginalEnDetalle) {
          detSheet.getRange(di + 1, iDetCant + 1).setValue(cantFinal);
        }
        // Marcar observacion como APROBADO
        detSheet.getRange(di + 1, iDetObs + 1).setValue('APROBADO');

        // Calcular ajustes de stock si guía cerrada (stock ya se aplicó)
        if (guiaCerrada && esIngreso) {
          if (codigoFinal !== codigoOriginal) {
            stockDeltaOrig = -cantOriginalEnDetalle;  // revertir aplicación al código viejo
            stockDeltaFin  = +cantFinal;              // aplicar al código nuevo
          } else if (cantFinal !== cantOriginalEnDetalle) {
            // Mismo código, solo ajusta delta
            stockDeltaFin = (cantFinal - cantOriginalEnDetalle);
          }
        }
        break;
      }
    } catch(e) {
      Logger.log('aprobarProductoNuevo: GUIA_DETALLE error: ' + e.message);
    }
  }

  // 5) Aplicar deltas
  try {
    if (stockDeltaOrig !== 0) {
      _actualizarStock(codigoOriginal, stockDeltaOrig, {
        tipoOperacion: 'APROBACION_PN',
        origen:        idProductoNuevo,
        usuario:       String(params.usuario || ''),
        observacion:   'reverso codigo original'
      });
    }
    if (stockDeltaFin !== 0) {
      _actualizarStock(codigoFinal, stockDeltaFin, {
        tipoOperacion: 'APROBACION_PN',
        origen:        idProductoNuevo,
        usuario:       String(params.usuario || ''),
        observacion:   'aplicar codigo final'
      });
    }
  } catch(eS) {
    Logger.log('aprobarProductoNuevo: stock error: ' + eS.message);
  }

  return { ok: true, data: { idProducto: idProductoCreado, tipo: tipo, guiaCerrada: guiaCerrada } };
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

  // Ampliar codSet con idProducto de cada codigoBarra para cubrir registros históricos
  // anteriores al fix (que guardaban idProducto en lugar de codigoBarra)
  try {
    var prods = _sheetToObjects(getProductosSheet());
    codigos.forEach(function(cb) {
      var p = prods.find(function(p) { return String(p.codigoBarra) === cb; });
      if (p && p.idProducto) codSet[String(p.idProducto)] = true;
    });
  } catch(e) {}

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
        cantidad:  Math.abs(parseFloat(d.cantidadRecibida || d.cantidadReal || d.cantidadEsperada || d.cantidad || 0)),
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

  // printerIdOverride: el modal de selección (admin/master) puede mandar a
  // otra impresora del ecosistema. Sin override → TICKET de ALMACEN.
  var printerId;
  if (params.printerIdOverride) {
    printerId = String(params.printerIdOverride).trim();
  } else {
    try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
    catch(e) { return { ok: false, error: e.message }; }
  }

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
  if (cant <= 0) return { ok: false, error: 'cantidad inválida' };
  if (!params.codigoProducto) return { ok: false, error: 'codigoProducto requerido' };

  // Foto obligatoria — subirla a Drive si viene base64
  var fotoUrl    = String(params.foto || '');
  var fotoBase64 = String(params.fotoBase64 || '').trim();
  if (fotoBase64) {
    try { fotoUrl = _subirFotoMerma(id, fotoBase64, params.mimeType || 'image/jpeg'); }
    catch(e) { Logger.log('Error subiendo foto merma: ' + e.message); }
  }
  if (!fotoUrl) return { ok: false, error: 'Foto obligatoria al registrar merma' };

  if (params.idSesion) registrarActividad(params.idSesion, 'MERMA_REGISTRADA', 1);

  // Escribir por nombre de columna (la hoja puede tener schema viejo o nuevo)
  // Al primer registro con campos nuevos, agregamos las columnas faltantes.
  _ensureColumnasMerma(sheet);

  var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row  = new Array(hdrs.length).fill('');
  function _set(col, val) { var i = hdrs.indexOf(col); if (i >= 0) row[i] = val; }
  _set('idMerma',           id);
  _set('fechaIngreso',      new Date());
  _set('origen',            params.responsable || params.origen || 'ALMACEN');  // RECEPCION/ALMACEN/idZona
  _set('responsable',       params.responsable || '');
  _set('codigoProducto',    params.codigoProducto);
  _set('idLote',            params.idLote || '');
  _set('cantidadOriginal',  cant);
  _set('cantidadPendiente', cant);
  _set('cantidadReparada',  0);
  _set('cantidadDesechada', 0);
  _set('motivo',            params.motivo || '');
  _set('usuario',           params.usuario || '');
  _set('foto',              fotoUrl);
  _set('estado',            'EN_PROCESO');
  sheet.appendRow(row);

  // Push a MOS — notificar a admins/master de nueva merma EN_PROCESO
  try {
    var titulo = '⚠ Merma registrada';
    var sub    = (params.codigoProducto || '?') + ' · ' + cant + ' uds · ' +
                 (params.responsable || params.origen || 'ALMACEN') +
                 ' · por ' + (params.usuario || 'operador');
    _notificarMOS(titulo, sub, params.usuario || null, 'WH_MERMA_REGISTRADA');
  } catch(eP) { Logger.log('Push merma: ' + eP.message); }

  return { ok: true, data: { idMerma: id, fotoUrl: fotoUrl } };
}

// Asegura que la hoja MERMAS tenga las columnas nuevas. Idempotente.
function _ensureColumnasMerma(sheet) {
  var lastCol  = sheet.getLastColumn();
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var faltantes = [];
  ['responsable','cantidadReparada','cantidadDesechada','foto',
   'fechaResolucion','observacionResolucion','idGuiaSalida'
  ].forEach(function(c) {
    if (existing.indexOf(c) < 0) faltantes.push(c);
  });
  if (!faltantes.length) return;
  sheet.getRange(1, lastCol + 1, 1, faltantes.length).setValues([faltantes]);
}

// Sube foto de merma a Drive con permiso público (similar al de preingresos)
function _subirFotoMerma(idMerma, base64Data, mimeType) {
  var folderId = PropertiesService.getScriptProperties().getProperty('MERMAS_FOLDER_ID');
  var folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = null; }
  }
  if (!folder) {
    // Crear carpeta junto al spreadsheet
    var ssId   = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    var ssFile = DriveApp.getFileById(ssId);
    var parent = ssFile.getParents().next();
    var imgIter = parent.getFoldersByName('imagenes');
    var imgFolder = imgIter.hasNext() ? imgIter.next() : parent.createFolder('imagenes');
    var mIter = imgFolder.getFoldersByName('mermas');
    folder = mIter.hasNext() ? mIter.next() : imgFolder.createFolder('mermas');
    PropertiesService.getScriptProperties().setProperty('MERMAS_FOLDER_ID', folder.getId());
  }
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, idMerma + '.jpg');
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

// ============================================================
// resolverMerma — el operador decide cuántas reparar y cuántas desechar.
// La parte desechada se envía a la guía SALIDA_MERMA semanal (lun-dom).
// La parte reparada NO toca stock (siempre estuvo).
// ============================================================
function resolverMerma(params) {
  var idMerma         = String(params.idMerma || '');
  var cantReparada    = parseFloat(params.cantidadReparada)  || 0;
  var cantDesechada   = parseFloat(params.cantidadDesechada) || 0;
  var observacion     = String(params.observacionResolucion || '');
  var usuario         = String(params.usuario || '');
  if (!idMerma) return { ok: false, error: 'idMerma requerido' };

  var sheet = getSheet('MERMAS');
  _ensureColumnasMerma(sheet);
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId       = hdrs.indexOf('idMerma');
  var idxEstado   = hdrs.indexOf('estado');
  var idxOrig     = hdrs.indexOf('cantidadOriginal');
  var idxPend     = hdrs.indexOf('cantidadPendiente');
  var idxRep      = hdrs.indexOf('cantidadReparada');
  var idxDes      = hdrs.indexOf('cantidadDesechada');
  var idxFechaR   = hdrs.indexOf('fechaResolucion');
  var idxObsR     = hdrs.indexOf('observacionResolucion');
  var idxIdGuia   = hdrs.indexOf('idGuiaSalida');
  var idxCb       = hdrs.indexOf('codigoProducto');
  var idxMotivo   = hdrs.indexOf('motivo');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) !== idMerma) continue;
    var estadoActual = String(data[i][idxEstado] || '').toUpperCase();
    if (estadoActual === 'RESUELTA' || estadoActual === 'DESECHADA') {
      return { ok: false, error: 'Merma ya resuelta' };
    }
    var cantOrig = parseFloat(data[i][idxOrig]) || 0;
    if (Math.abs((cantReparada + cantDesechada) - cantOrig) > 0.001) {
      return { ok: false, error: 'reparada + desechada debe igualar ' + cantOrig };
    }

    var idGuiaSalida = '';
    var codigoProd   = String(data[i][idxCb] || '');
    var motivoOrig   = String(data[i][idxMotivo] || '');

    // Si hay parte para desechar, agregar a la guía SALIDA_MERMA semanal
    if (cantDesechada > 0 && codigoProd) {
      var resGuia = _getOCrearGuiaMermaSemana(usuario);
      if (!resGuia.ok) return { ok: false, error: 'Error guía merma: ' + resGuia.error };
      idGuiaSalida = resGuia.data.idGuia;
      var resDet = agregarDetalleGuia({
        idGuia:           idGuiaSalida,
        codigoProducto:   codigoProd,
        cantidadEsperada: cantDesechada,
        cantidadRecibida: cantDesechada,
        observacion:      'Merma ' + idMerma + (motivoOrig ? ' · ' + motivoOrig : '')
      });
      if (!resDet.ok) return { ok: false, error: 'Detalle guía: ' + resDet.error };
    }

    // Actualizar la fila de merma
    sheet.getRange(i + 1, idxRep    + 1).setValue(cantReparada);
    sheet.getRange(i + 1, idxDes    + 1).setValue(cantDesechada);
    sheet.getRange(i + 1, idxPend   + 1).setValue(0);
    sheet.getRange(i + 1, idxEstado + 1).setValue('RESUELTA');
    if (idxFechaR >= 0)  sheet.getRange(i + 1, idxFechaR + 1).setValue(new Date());
    if (idxObsR >= 0)    sheet.getRange(i + 1, idxObsR   + 1).setValue(observacion);
    if (idxIdGuia >= 0 && idGuiaSalida) sheet.getRange(i + 1, idxIdGuia + 1).setValue(idGuiaSalida);

    return { ok: true, data: { idMerma: idMerma, idGuiaSalida: idGuiaSalida,
                                cantidadReparada: cantReparada, cantidadDesechada: cantDesechada } };
  }
  return { ok: false, error: 'Merma no encontrada' };
}

// Busca o crea la guía SALIDA_MERMA ABIERTA de la semana actual (lun-dom).
// Todas las mermas desechadas durante la semana van a la misma guía.
// El operador la cierra manualmente cuando los productos físicamente se retiran.
function _getOCrearGuiaMermaSemana(usuario) {
  var tz = Session.getScriptTimeZone();
  // Calcular lunes de la semana actual
  var ahora = new Date();
  var diaSem = ahora.getDay(); // 0=domingo, 1=lunes, ..., 6=sábado
  var diasAtrasLunes = diaSem === 0 ? 6 : (diaSem - 1);
  var lunes = new Date(ahora);
  lunes.setDate(ahora.getDate() - diasAtrasLunes);
  lunes.setHours(0, 0, 0, 0);
  var domingo = new Date(lunes);
  domingo.setDate(lunes.getDate() + 7);

  // Buscar guía SALIDA_MERMA ABIERTA cuya fecha caiga en esta semana
  var guias = _sheetToObjects(getSheet('GUIAS'));
  for (var i = 0; i < guias.length; i++) {
    var g = guias[i];
    if (g.tipo !== 'SALIDA_MERMA') continue;
    if (String(g.estado || '').toUpperCase() !== 'ABIERTA') continue;
    var fGuia = g.fecha ? new Date(g.fecha) : null;
    if (!fGuia) continue;
    if (fGuia >= lunes && fGuia < domingo) {
      return { ok: true, data: { idGuia: g.idGuia, esNueva: false } };
    }
  }
  // No existe → crear nueva
  var fmt    = Utilities.formatDate(lunes, tz, 'dd MMM');
  var fmtFin = Utilities.formatDate(new Date(domingo.getTime() - 1), tz, 'dd MMM');
  var res = crearGuia({
    tipo:       'SALIDA_MERMA',
    usuario:    usuario || 'sistema',
    comentario: 'Mermas semana ' + fmt + ' al ' + fmtFin
  });
  if (!res.ok) return res;
  return { ok: true, data: { idGuia: res.data.idGuia, esNueva: true } };
}

// Endpoint: lista mermas EN_PROCESO con flag de vencidas (>3 días)
function getMermasEnProceso(params) {
  var rows = _sheetToObjects(getSheet('MERMAS'));
  rows = rows.filter(function(r) { return String(r.estado || '').toUpperCase() === 'EN_PROCESO'; });
  var ahora = new Date().getTime();
  var TRES_DIAS = 3 * 24 * 60 * 60 * 1000;
  rows.forEach(function(r) {
    var f = r.fechaIngreso ? new Date(r.fechaIngreso).getTime() : ahora;
    r.diasEnProceso = Math.floor((ahora - f) / (24 * 60 * 60 * 1000));
    r.vencida       = (ahora - f) > TRES_DIAS;
  });
  rows.sort(function(a, b) { return new Date(b.fechaIngreso) - new Date(a.fechaIngreso); });
  return { ok: true, data: rows };
}

// Cuenta mermas vencidas (>3 días sin resolver) — para badge en dashboard
function getMermasVencidas() {
  var rows = _sheetToObjects(getSheet('MERMAS'));
  var ahora = new Date().getTime();
  var TRES_DIAS = 3 * 24 * 60 * 60 * 1000;
  var vencidas = rows.filter(function(r) {
    if (String(r.estado || '').toUpperCase() !== 'EN_PROCESO') return false;
    var f = r.fechaIngreso ? new Date(r.fechaIngreso).getTime() : ahora;
    return (ahora - f) > TRES_DIAS;
  });
  return { ok: true, data: { count: vencidas.length, mermas: vencidas } };
}

// Marca mermas asociadas a una guía SALIDA_MERMA cerrada como DESECHADA
// y emite push de resumen a MASTER/ADMINISTRADOR.
function _cerrarMermasDeGuia(idGuia, detalles) {
  var sheet = getSheet('MERMAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId      = hdrs.indexOf('idMerma');
  var idxEstado  = hdrs.indexOf('estado');
  var idxIdGuia  = hdrs.indexOf('idGuiaSalida');
  if (idxId < 0 || idxEstado < 0 || idxIdGuia < 0) return;

  var marcadas = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdGuia]) === String(idGuia)) {
      sheet.getRange(i + 1, idxEstado + 1).setValue('DESECHADA');
      marcadas++;
    }
  }

  // Push resumen a admins
  try {
    var totalUds = (detalles || []).reduce(function(s, d) {
      return s + (parseFloat(d.cantidadRecibida) || 0);
    }, 0);
    var nProds = (detalles || []).length;
    _notificarMOS('🗑 Desecho de mermas',
      marcadas + ' merma(s) · ' + nProds + ' producto(s) · ' + totalUds + ' unid · guía ' + idGuia,
      null, 'WH_MERMA_DESECHO');
  } catch(eP) { Logger.log('Push desecho: ' + eP.message); }
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
    _actualizarStock(codigoBarra, diff, {
      tipoOperacion: 'AUDITORIA',
      origen:        audId,
      usuario:       String(params.usuario || ''),
      observacion:   'físico=' + stockFisico + ' sistema=' + stockSistema
    });
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

  _actualizarStock(codigoBarra, delta, {
    tipoOperacion: 'AJUSTE_MANUAL',
    origen:        id,
    usuario:       String(params.usuario || ''),
    observacion:   String(params.motivo || '')
  });

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

// ── Snapshot de tarifa por carreta ─────────────────────────────
// Recibe el JSON string de cargadores [{id,nombre,carretas,tarifa?}] y retorna
// el mismo string pero con `tarifa` rellenada con la tarifa global vigente
// si no venía. Garantiza que reimpresiones a futuro conserven el monto que
// se pactó al momento, aunque cambie TARIFA_CARRETA global.
function _snapshotTarifaCargadores(cargadoresStr) {
  if (!cargadoresStr) return '';
  try {
    var arr = JSON.parse(cargadoresStr);
    if (!Array.isArray(arr) || !arr.length) return cargadoresStr;
    var tarifaGlobal = parseFloat(_getConfigValue('TARIFA_CARRETA')) || 0;
    var out = arr.map(function(c) {
      if (typeof c !== 'object' || c === null) return c;
      var t = parseFloat(c.tarifa);
      if (isNaN(t) || t < 0) t = tarifaGlobal;
      return { id: c.id, nombre: c.nombre, carretas: c.carretas, tarifa: t };
    });
    return JSON.stringify(out);
  } catch(e) { return cargadoresStr; }
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

  // Snapshot de tarifa por carreta para que reimpresiones futuras conserven
  // el monto pactado al momento, aunque después se cambie la tarifa global.
  var cargSnapshot = _snapshotTarifaCargadores(params.cargadores);

  // Escribir por nombre de columna — funciona con 9 o 12 cols (schema viejo/nuevo)
  var hdrs = data[0];
  var row  = new Array(hdrs.length).fill('');
  function _set(col, val) { var i = hdrs.indexOf(col); if (i >= 0) row[i] = val; }
  _set('idPreingreso', id);
  _set('fecha',        new Date());
  _set('idProveedor',  params.idProveedor  || '');
  _set('cargadores',   cargSnapshot);
  _set('usuario',      params.usuario      || '');
  _set('monto',        parseFloat(params.monto) || 0);
  _set('fotos',        params.fotos        || '');
  _set('comentario',   params.comentario   || '');
  _set('estado',       'PENDIENTE');
  _set('idGuia',       '');
  sheet.appendRow(row);

  // Push: notificar nuevo preingreso a MOS
  try {
    _notificarMOS(
      '📦 Nuevo preingreso',
      (params.usuario || 'Operador') + ' · S/ ' + (parseFloat(params.monto) || 0).toFixed(2),
      params.usuario || null,
      'WH_PREINGRESO'
    );
  } catch(eP) { Logger.log('Push preingreso: ' + eP.message); }

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

    // Campos editables (cargadores se snapshotea con tarifa antes de escribir)
    var editable = { idProveedor: true, monto: true, comentario: true, fotos: true, cargadores: true };
    Object.keys(editable).forEach(function(key) {
      if (params[key] === undefined) return;
      var col = hdrs.indexOf(key);
      if (col < 0) return;
      var val;
      if (key === 'monto')           val = parseFloat(params[key]) || 0;
      else if (key === 'cargadores') val = _snapshotTarifaCargadores(params[key]);
      else                            val = String(params[key]);
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

      if (!resultGuia.ok) return { ok: false, error: 'Error al crear guía: ' + resultGuia.error };

      sheet.getRange(i + 1, hdrs.indexOf('estado') + 1).setValue('PROCESADO');
      sheet.getRange(i + 1, hdrs.indexOf('idGuia') + 1).setValue(resultGuia.data.idGuia);

      return { ok: true, data: { idGuia: resultGuia.data.idGuia } };
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
  var fileId       = String(params.fileId       || '').trim();  // foto específica elegida
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  try {
    var guiaFolderId = PropertiesService.getScriptProperties().getProperty('FOTOS_GUIA_FOLDER_ID');
    if (!guiaFolderId) return { ok: false, error: 'Carpeta guías no configurada. Ejecuta setupGuiasFolders()' };
    var guiaFolder = DriveApp.getFolderById(guiaFolderId);
    var copyName   = idGuia + '.jpg';

    // Resolver srcFile: si viene fileId, usar ese; sino tomar la primera del folder del preingreso (legado)
    var srcFile = null;
    if (fileId) {
      try { srcFile = DriveApp.getFileById(fileId); }
      catch(e) { return { ok: false, error: 'Foto no encontrada (fileId inválido)' }; }
    } else {
      if (!idPreingreso) return { ok: false, error: 'idPreingreso o fileId requerido' };
      var preFolderId = PropertiesService.getScriptProperties().getProperty('FOTOS_PRE_FOLDER_ID');
      if (!preFolderId) return { ok: false, error: 'Carpeta preingresos no configurada' };
      var piFolder = DriveApp.getFolderById(preFolderId).getFoldersByName(idPreingreso);
      if (!piFolder.hasNext()) return { ok: false, error: 'Carpeta del preingreso no encontrada' };
      var files = piFolder.next().getFiles();
      if (!files.hasNext()) return { ok: false, error: 'Sin fotos en el preingreso' };
      srcFile = files.next();
    }

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
// imprimirMembrete — ESC/POS 80mm, layout secuencial (standard mode)
//
// Nombre: bold+doble alto+doble ancho, centrado, máx 2 líneas.
// Barcodes EAN: CODE128, module width=2, height=80 dots (uniforme).
// SKU barcode al final si hay más de 1 EAN.
//
// ENCODING: byte array nativo → Utilities.newBlob → base64
// BARCODES: usa params.barcodes del frontend; fallback a DB.
// ============================================================
function imprimirMembrete(params) {
  var idProducto = String(params.idProducto || '');
  if (!idProducto) return { ok: false, error: 'idProducto requerido' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  // printerIdOverride: el modal de selección (admin/master) puede mandar a
  // otra impresora del ecosistema. Sin override → TICKET de ALMACEN.
  var printerId;
  if (params.printerIdOverride) {
    printerId = String(params.printerIdOverride).trim();
  } else {
    try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
    catch(e) { return { ok: false, error: e.message }; }
  }

  // ── Producto ────────────────────────────────────────────────
  var productos = _sheetToObjects(getProductosSheet());
  var prod = productos.find(function(p) {
    return p.idProducto === idProducto || String(p.codigoBarra) === idProducto;
  });
  if (!prod) return { ok: false, error: 'Producto no encontrado: ' + idProducto };

  var sku = String(prod.skuBase || prod.idProducto);

  // ── Barcodes: frontend envía params.barcodes; fallback a DB ─
  var allEan = [];
  if (params.barcodes) {
    try {
      var parsed = JSON.parse(String(params.barcodes));
      if (Array.isArray(parsed)) allEan = parsed.map(String).filter(Boolean);
    } catch(e) {}
  }
  if (!allEan.length) {
    if (prod.codigoBarra) allEan.push(String(prod.codigoBarra).trim());
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        var skuU = String(sku).trim().toUpperCase();
        _sheetToObjects(equivSheet)
          .filter(function(e) {
            return String(e.skuBase || '').trim().toUpperCase() === skuU
                && _esActivo(e.activo)
                && e.codigoBarra;
          })
          .forEach(function(e) {
            var c = String(e.codigoBarra).trim();
            if (c && allEan.indexOf(c) < 0) allEan.push(c);
          });
      }
    } catch(e) {}
  }
  if (!allEan.length) allEan.push(sku);

  var numEan      = allEan.length;
  var hasMultiple = numEan > 1;

  // ── Nombre: word-wrap, máx 2 líneas, ~20 chars/línea ───────
  // Font A doble ancho en 80mm ≈ 24 chars visibles; 20 deja margen
  var desc      = String(prod.descripcion || sku).toUpperCase();
  var nameLines = _membWrap(desc, 20).slice(0, 2);

  // ── Construir ESC/POS como array de bytes ───────────────────
  var B = [];
  function b1(v)   { B.push(v & 0xff); }
  function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
  function bLn(s)  { bStr(s); b1(0x0a); }

  // Separadores (48 chars = ancho nominal 80mm font A)
  var SEPEQ  = '================================================';
  var SEPDA  = '------------------------------------------------';

  // ESC @ — init
  b1(0x1b); b1(0x40);

  // ── Nombre: centrado, bold + doble alto + doble ancho ───────
  b1(0x1b); b1(0x61); b1(0x01);   // center
  b1(0x1b); b1(0x21); b1(0x38);   // bold + double-height + double-width
  for (var ni = 0; ni < nameLines.length; ni++) { bLn(nameLines[ni]); }
  b1(0x1b); b1(0x21); b1(0x00);   // normal

  // SKU como texto (siempre, debajo del nombre): bold tamaño normal, centrado
  b1(0x1b); b1(0x21); b1(0x08);   // bold
  bLn('SKU: ' + sku);
  b1(0x1b); b1(0x21); b1(0x00);   // normal

  // Separador "===..." debajo del SKU (centrado explícito)
  b1(0x1b); b1(0x61); b1(0x01);   // center
  bLn(SEPEQ);

  // ── EAN barcodes: tamaño uniforme h=80 w=2 ──────────────────
  for (var ei = 0; ei < numEan; ei++) {
    var bd = '{B' + allEan[ei];
    b1(0x1b); b1(0x61); b1(0x01);   // center
    b1(0x1d); b1(0x68); b1(80);     // GS h 80: altura uniforme 80 dots
    b1(0x1d); b1(0x77); b1(0x02);   // GS w 2: module width fijo 2
    b1(0x1d); b1(0x48); b1(0x02);   // GS H 2: HRI debajo
    b1(0x1d); b1(0x66); b1(0x00);   // GS f 0: HRI font A
    b1(0x1d); b1(0x6b); b1(0x49);   // GS k 73: CODE128 func.2
    b1(bd.length & 0xff);
    bStr(bd);
    b1(0x0a);                       // LF tras barcode
    // Separadores también centrados (consistencia visual)
    b1(0x1b); b1(0x61); b1(0x01);   // center
    var isLast = (ei === numEan - 1);
    bLn(isLast ? SEPEQ : SEPDA);
  }

  // ── SKU barcode (solo si hay varios EAN) ────────────────────
  if (hasMultiple) {
    var skuBd = '{B' + sku;
    b1(0x1b); b1(0x61); b1(0x01);   // center
    b1(0x1b); b1(0x21); b1(0x08);   // bold
    bLn('SKU: ' + sku);
    b1(0x1b); b1(0x21); b1(0x00);   // normal
    b1(0x1d); b1(0x68); b1(60);     // GS h 60
    b1(0x1d); b1(0x77); b1(0x02);   // GS w 2
    b1(0x1d); b1(0x48); b1(0x02);   // GS H 2: HRI debajo
    b1(0x1d); b1(0x66); b1(0x00);   // font A
    b1(0x1d); b1(0x6b); b1(0x49);   // CODE128
    b1(skuBd.length & 0xff);
    bStr(skuBd);
    b1(0x0a);
    b1(0x1b); b1(0x61); b1(0x01);   // center (consistencia)
    bLn(SEPEQ);
  }

  // ── Avance ≈20mm + corte completo ───────────────────────────
  b1(0x1b); b1(0x4a); b1(160);   // ESC J 160
  b1(0x1d); b1(0x56); b1(0x00);  // GS V 0

  // ── Base64 vía Blob (raw bytes garantizados) ─────────────────
  var blob = Utilities.newBlob(B, 'application/octet-stream');
  var b64  = Utilities.base64Encode(blob.getBytes());

  var payload = {
    printerId:   parseInt(printerId),
    title:       'Membrete ' + sku,
    contentType: 'raw_base64',
    content:     b64,
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
