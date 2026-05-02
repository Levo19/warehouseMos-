// ============================================================
// warehouseMos — Envasados.gs
// Gestión de envasado + generación automática de guías
// + integración PrintNode para etiquetas adhesivas
// ============================================================

function getEnvasados(params) {
  var rows = _sheetToObjects(getSheet('ENVASADOS'));
  if (params.estado)     rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  if (params.fecha)      rows = rows.filter(function(r){ return r.fecha === params.fecha; });
  if (params.fechaDesde) rows = rows.filter(function(r){ return String(r.fecha) >= String(params.fechaDesde); });
  if (params.limit)      rows = rows.slice(0, parseInt(params.limit));
  return { ok: true, data: rows };
}

function getPendientesEnvasado() {
  var productos = _sheetToObjects(getProductosSheet());
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
  var codigoBarra      = String(params.codigoBarra || '').trim();
  var unidadesReales   = parseInt(params.unidadesProducidas) || 0;
  var fechaVencimiento = params.fechaVencimiento || '';
  var usuario          = params.usuario || 'sistema';
  var imprimirEtiq     = params.imprimirEtiquetas !== false;

  if (!codigoBarra || unidadesReales <= 0) {
    return { ok: false, error: 'Faltan datos: codigoBarra, unidadesProducidas' };
  }

  var productos = _sheetToObjects(getProductosSheet());

  // 1. Buscar derivado por codigoBarra
  var prodDerivado = productos.find(function(p) {
    return String(p.codigoBarra).trim() === codigoBarra;
  });
  if (!prodDerivado) return { ok: false, error: 'Producto envasado no encontrado: ' + codigoBarra };

  // 2. Buscar base por skuBase === codigoProductoBase del derivado (fallback idProducto)
  var claveBase = String(prodDerivado.codigoProductoBase || '').trim();
  if (!claveBase) return { ok: false, error: 'El producto no tiene codigoProductoBase configurado' };

  var prodBase = productos.find(function(p) {
    return String(p.skuBase).trim() === claveBase || String(p.idProducto).trim() === claveBase;
  });
  if (!prodBase) return { ok: false, error: 'Producto base no encontrado: ' + claveBase };

  // 3. Calcular cantidad base consumida: unidades × factorConversionBase
  var factorBase = parseFloat(prodDerivado.factorConversionBase) || 0;
  if (factorBase <= 0) {
    return { ok: false, error: 'factorConversionBase no configurado para: ' + prodDerivado.descripcion };
  }
  var cantBase = unidadesReales * factorBase;

  var fecha      = new Date();
  var idEnvasado = _generateId('ENV');
  var idLote     = _generateId('LOT');

  // 5. Guía SALIDA_ENVASADO del día — reutilizar si ya existe ABIERTA hoy
  var gsRes = _getOCrearGuiaDia('SALIDA_ENVASADO', usuario);
  if (!gsRes.ok) return { ok: false, error: 'Error guía salida: ' + gsRes.error };

  var detSalida = agregarDetalleGuia({
    idGuia:           gsRes.data.idGuia,
    codigoProducto:   prodBase.codigoBarra,
    cantidadEsperada: cantBase,
    cantidadRecibida: cantBase,
    precioUnitario:   0
  });
  if (!detSalida.ok) return { ok: false, error: 'Detalle salida: ' + detSalida.error };
  _actualizarStock(prodBase.codigoBarra, -cantBase, {
    tipoOperacion: 'ENVASADO_BASE',
    origen:        idEnvasado,
    usuario:       String(usuario || ''),
    observacion:   'consumo base ' + unidadesReales + ' uds'
  });

  // 6. Guía INGRESO_ENVASADO del día — reutilizar si ya existe ABIERTA hoy
  var giRes = _getOCrearGuiaDia('INGRESO_ENVASADO', usuario);
  if (!giRes.ok) return { ok: false, error: 'Error guía ingreso: ' + giRes.error };

  var detIngreso = agregarDetalleGuia({
    idGuia:           giRes.data.idGuia,
    codigoProducto:   prodDerivado.codigoBarra,
    cantidadEsperada: unidadesReales,
    cantidadRecibida: unidadesReales,
    precioUnitario:   0,
    idLote:           idLote,
    fechaVencimiento: fechaVencimiento
  });
  if (!detIngreso.ok) return { ok: false, error: 'Detalle ingreso: ' + detIngreso.error };
  _actualizarStock(prodDerivado.codigoBarra, unidadesReales, {
    tipoOperacion: 'ENVASADO_DERIVADO',
    origen:        idEnvasado,
    usuario:       String(usuario || ''),
    observacion:   'producción ' + unidadesReales + ' uds'
  });

  // 7. Registro en hoja ENVASADOS
  getSheet('ENVASADOS').appendRow([
    idEnvasado,
    prodBase.codigoBarra || prodBase.idProducto,
    cantBase,
    prodBase.unidad,
    prodDerivado.codigoBarra || prodDerivado.idProducto,
    unidadesReales,
    unidadesReales,
    0,
    100,
    fecha,
    usuario,
    'COMPLETADO',
    gsRes.data.idGuia,
    giRes.data.idGuia,
    ''
  ]);

  // 8. Imprimir etiquetas
  var resultImpresion = null;
  if (imprimirEtiq) {
    resultImpresion = _imprimirEtiquetasEnvasado({
      codigoDerivado:   prodDerivado.idProducto,
      descripcion:      prodDerivado.descripcion,
      codigoBarra:      prodDerivado.codigoBarra,
      cantidad:         prodDerivado.unidad,
      unidades:         unidadesReales,
      fechaVencimiento: fechaVencimiento,
      idLote:           idLote
    });
  }

  return {
    ok: true,
    data: {
      idEnvasado:        idEnvasado,
      idGuiaSalida:      gsRes.data.idGuia,
      idGuiaIngreso:     giRes.data.idGuia,
      cantidadBase:      cantBase,
      unidadesProducidas: unidadesReales,
      idLote:            idLote,
      impresion:         resultImpresion
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

// Devuelve la guía de ese tipo creada hoy (cualquier estado), o crea+cierra una nueva.
// Garantiza una sola guía por tipo por día; detalles se agregan a la existente.
function _getOCrearGuiaDia(tipo, usuario) {
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var guias = _sheetToObjects(getSheet('GUIAS'));
  for (var i = 0; i < guias.length; i++) {
    var g = guias[i];
    if (g.tipo !== tipo) continue;
    var fGuia = g.fecha ? String(g.fecha).substring(0, 10) : '';
    if (fGuia === hoy) return { ok: true, data: { idGuia: g.idGuia } };
  }
  // No existe → crear y cerrar de inmediato (sin tocar stock)
  var res = crearGuia({ tipo: tipo, usuario: usuario, comentario: 'Envasados ' + hoy });
  if (!res.ok) return res;
  _cerrarGuiaSinStock(res.data.idGuia);
  return res;
}
