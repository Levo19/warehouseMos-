// ============================================================
// warehouseMos — Envasados.gs
// Gestión de envasado + generación automática de guías
// + integración PrintNode para etiquetas adhesivas
// ============================================================

function getEnvasados(params) {
  var rows = _sheetToObjects(getSheet('ENVASADOS'));
  if (params.estado) rows = rows.filter(function(r){ return r.estado === params.estado; });
  if (params.fecha)  rows = rows.filter(function(r){ return r.fecha  === params.fecha; });
  if (params.limit)  rows = rows.slice(0, parseInt(params.limit));
  return { ok: true, data: rows };
}

function getPendientesEnvasado() {
  var productos = _sheetToObjects(getSheet('PRODUCTOS'));
  var stock     = _sheetToObjects(getSheet('STOCK'));

  var stockMap = {};
  stock.forEach(function(s){ stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });

  return { ok: true, data: _calcularPendientesEnvasado(productos, stockMap) };
}

// ============================================================
// registrarEnvasado — corazón del módulo
// Genera automáticamente:
//   1. GuíaSalida (descuenta base)
//   2. GuíaIngreso (agrega derivado)
//   3. Registro en ENVASADOS
//   4. Imprime etiquetas adhesivas vía PrintNode
// ============================================================
function registrarEnvasado(params) {
  var codigoBase       = params.codigoProductoBase;
  var cantBase         = parseFloat(params.cantidadBase) || 0;
  var codigoDerivado   = params.codigoProductoEnvasado;
  var unidadesReales   = parseInt(params.unidadesProducidas) || 0;
  var fechaVencimiento = params.fechaVencimiento || '';
  var usuario          = params.usuario || 'sistema';
  var imprimirEtiq     = params.imprimirEtiquetas !== false; // default true

  if (!codigoBase || !codigoDerivado || cantBase <= 0 || unidadesReales <= 0) {
    return { ok: false, error: 'Faltan datos: codigoProductoBase, codigoProductoEnvasado, cantidadBase, unidadesProducidas' };
  }

  // Cargar producto derivado para calcular factor
  var productos = _sheetToObjects(getSheet('PRODUCTOS'));
  var prodDerivado = productos.find(function(p){ return p.idProducto === codigoDerivado; });
  var prodBase     = productos.find(function(p){ return p.idProducto === codigoBase; });
  if (!prodDerivado) return { ok: false, error: 'Producto derivado no encontrado: ' + codigoDerivado };
  if (!prodBase)     return { ok: false, error: 'Producto base no encontrado: ' + codigoBase };

  // Validar stock base suficiente
  var stockBaseActual = _getStockProducto(codigoBase).cantidad;
  if (stockBaseActual < cantBase) {
    return {
      ok: false,
      error: 'Stock insuficiente. Disponible: ' + stockBaseActual + ' ' + prodBase.unidad +
             ', solicitado: ' + cantBase
    };
  }

  var factor    = parseFloat(prodDerivado.factorConversion) || 1;
  var merma     = parseFloat(prodDerivado.mermaEsperadaPct) || 0;
  var unidadesEsperadas = Math.floor(cantBase * factor * (1 - merma / 100));
  var mermaReal         = Math.max(0, unidadesEsperadas - unidadesReales);
  var eficiencia        = unidadesEsperadas > 0
    ? Math.round((unidadesReales / unidadesEsperadas) * 1000) / 10
    : 100;

  var fecha = new Date();
  var idEnvasado = _generateId('ENV');

  // ── Guía SALIDA (descuenta base) ───────────────────────────
  var resultGS = crearGuia({
    tipo:       'SALIDA_ENVASADO',
    usuario:    usuario,
    comentario: 'Auto: envasado ' + prodDerivado.descripcion + ' (' + unidadesReales + ' uds)'
  });
  if (!resultGS.ok) return { ok: false, error: 'Error al crear guía salida: ' + resultGS.error };
  var idGuiaSalida = resultGS.data.idGuia;

  agregarDetalleGuia({
    idGuia:           idGuiaSalida,
    codigoProducto:   codigoBase,
    cantidadEsperada: cantBase,
    cantidadRecibida: cantBase,
    precioUnitario:   0
  });
  cerrarGuia(idGuiaSalida, usuario);

  // ── Guía INGRESO (agrega derivado) ─────────────────────────
  var resultGI = crearGuia({
    tipo:       'INGRESO_JEFATURA',
    usuario:    usuario,
    comentario: 'Auto: ingreso envasado ' + prodDerivado.descripcion
  });
  if (!resultGI.ok) return { ok: false, error: 'Error al crear guía ingreso: ' + resultGI.error };
  var idGuiaIngreso = resultGI.data.idGuia;

  var idLoteNuevo = 'LOT' + new Date().getTime();
  agregarDetalleGuia({
    idGuia:           idGuiaIngreso,
    codigoProducto:   codigoDerivado,
    cantidadEsperada: unidadesEsperadas,
    cantidadRecibida: unidadesReales,
    precioUnitario:   0,
    idLote:           idLoteNuevo
  });
  cerrarGuia(idGuiaIngreso, usuario);

  // Actualizar fecha de vencimiento en el lote si se proporcionó
  if (fechaVencimiento) {
    var lotesSheet = getSheet('LOTES_VENCIMIENTO');
    var lotesData  = lotesSheet.getDataRange().getValues();
    var hdrs = lotesData[0];
    var idxLoteId = hdrs.indexOf('idLote');
    var idxFV     = hdrs.indexOf('fechaVencimiento');
    for (var i = 1; i < lotesData.length; i++) {
      if (lotesData[i][idxLoteId] === idLoteNuevo) {
        lotesSheet.getRange(i + 1, idxFV + 1).setValue(new Date(fechaVencimiento));
        break;
      }
    }
  }

  // ── Registro ENVASADOS ─────────────────────────────────────
  getSheet('ENVASADOS').appendRow([
    idEnvasado,
    codigoBase,
    cantBase,
    prodBase.unidad,
    codigoDerivado,
    unidadesEsperadas,
    unidadesReales,
    mermaReal,
    eficiencia,
    fecha,
    usuario,
    'COMPLETADO',
    idGuiaSalida,
    idGuiaIngreso,
    params.observacion || ''
  ]);

  // ── Registrar actividad en desempeño ──────────────────────
  if (params.idSesion) {
    registrarActividad(params.idSesion, 'ENVASADO_REGISTRADO', 1);
    registrarActividad(params.idSesion, 'UNIDADES_ENVASADAS', unidadesReales);
  }

  // ── Imprimir etiquetas adhesivas ───────────────────────────
  var resultImpresion = null;
  if (imprimirEtiq) {
    resultImpresion = _imprimirEtiquetasEnvasado({
      codigoDerivado:   codigoDerivado,
      descripcion:      prodDerivado.descripcion,
      codigoBarra:      prodDerivado.codigoBarra,
      cantidad:         prodDerivado.unidad,
      unidades:         unidadesReales,
      fechaVencimiento: fechaVencimiento,
      idLote:           idLoteNuevo
    });
  }

  return {
    ok: true,
    data: {
      idEnvasado:       idEnvasado,
      idGuiaSalida:     idGuiaSalida,
      idGuiaIngreso:    idGuiaIngreso,
      unidadesEsperadas:unidadesEsperadas,
      unidadesProducidas:unidadesReales,
      mermaReal:        mermaReal,
      eficienciaPct:    eficiencia,
      idLote:           idLoteNuevo,
      impresion:        resultImpresion
    }
  };
}

// ============================================================
// Impresión de etiquetas adhesivas vía PrintNode (ZPL)
// Se imprime una etiqueta por unidad producida
// ============================================================
function _getPrintNodeProps() {
  var props = PropertiesService.getScriptProperties();
  return {
    apiKey:           props.getProperty('PRINTNODE_API_KEY')    || '',
    printerEtiquetas: props.getProperty('PRINTER_ETIQUETAS_ID') || '',
    printerTickets:   props.getProperty('PRINTER_TICKETS_ID')   || ''
  };
}

function _imprimirEtiquetasEnvasado(data) {
  var pn        = _getPrintNodeProps();
  var apiKey    = pn.apiKey;
  var printerId = pn.printerEtiquetas;
  if (!apiKey || !printerId) {
    return { ok: false, error: 'PrintNode no configurado (PRINTNODE_API_KEY / PRINTNODE_PRINTER_ID)' };
  }

  var empresaNombre = _getConfigValue('EMPRESA_NOMBRE') || 'InversionMos';
  var fechaImpresion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var fechaVenc = data.fechaVencimiento
    ? Utilities.formatDate(new Date(data.fechaVencimiento), Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : 'SIN VENC.';

  // ZPL para etiqueta adhesiva 50mm×30mm (ajustar según impresora)
  var zpl =
    '^XA' +
    '^PW400' +           // ancho en dots (50mm a 8dpi)
    '^LL240' +           // largo en dots (30mm)
    // Código de barras (si existe)
    (data.codigoBarra
      ? '^FO20,10^BY2^BCN,60,Y,N,N^FD' + data.codigoBarra + '^FS'
      : '') +
    // Nombre del producto
    '^FO20,80^A0N,22,22^FD' + _truncate(data.descripcion, 22) + '^FS' +
    // Vencimiento
    '^FO20,110^A0N,20,20^FDVto: ' + fechaVenc + '^FS' +
    // Lote
    '^FO20,138^A0N,18,18^FDLote: ' + data.idLote + '^FS' +
    // Empresa
    '^FO20,165^A0N,16,16^FD' + empresaNombre + '^FS' +
    // Cantidad de copias
    '^PQ' + data.unidades + ',0,1,Y' +
    '^XZ';

  var payload = {
    printerId: parseInt(printerId),
    title:     'Etiquetas ' + data.descripcion,
    contentType: 'raw_base64',
    content:   Utilities.base64Encode(zpl),
    source:    'warehouseMos'
  };

  try {
    var response = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:  'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code === 201) {
      return { ok: true, jobId: JSON.parse(response.getContentText()), unidades: data.unidades };
    } else {
      return { ok: false, error: response.getContentText(), httpCode: code };
    }
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function imprimirEtiqueta(params) {
  return _imprimirEtiquetasEnvasado(params);
}

function _truncate(str, max) {
  if (!str) return '';
  return str.length > max ? str.substring(0, max) : str;
}
