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
  return _conLock('registrarEnvasado', function() {
    return _registrarEnvasadoImpl(params);
  });
}

function _registrarEnvasadoImpl(params) {
  var codigoBarra      = String(params.codigoBarra || '').trim();
  var unidadesReales   = parseInt(params.unidadesProducidas) || 0;
  var fechaVencimiento = params.fechaVencimiento || '';
  var usuario          = params.usuario || 'sistema';
  var imprimirEtiq     = params.imprimirEtiquetas !== false;
  var idempotencyKey   = String(params.idempotencyKey || '').trim();

  if (!codigoBarra || unidadesReales <= 0) {
    return { ok: false, error: 'Faltan datos: codigoBarra, unidadesProducidas' };
  }

  // ── IDEMPOTENCIA ─────────────────────────────────────────
  // Bug histórico: doble-click + reintento por timeout creaban registros
  // duplicados en ENVASADOS (mismo usuario + producto + cantidad, separados
  // por 3-16s) inflando stock derivado y duplicando consumo base.
  //   - Si llega params.idempotencyKey y ya vimos esa key → retornar el
  //     idEnvasado existente sin re-ejecutar.
  //   - Si no llega, fingerprint = usuario + codigoBarra + unidades + minuto.
  // TTL 120s — cubre reintentos por timeout y dobles clicks. Igual patrón
  // que crearDespachoRapido (v2.3.0).
  var cache = CacheService.getScriptCache();
  var keyEfectiva = idempotencyKey;
  if (!keyEfectiva) {
    var minuto = Math.floor(Date.now() / 60000);
    keyEfectiva = 'env_' + usuario + '_' + codigoBarra + '_' + unidadesReales + '_' + minuto;
    if (keyEfectiva.length > 240) {
      keyEfectiva = 'env_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyEfectiva)
        .map(function(b){ return (b < 0 ? b + 256 : b).toString(16); }).join('');
    }
  } else {
    keyEfectiva = 'env_' + keyEfectiva;
  }
  try {
    var prev = cache.get(keyEfectiva);
    if (prev) {
      Logger.log('registrarEnvasado idempotente: hit cache → ' + prev);
      return { ok: true, data: { idEnvasado: prev, dedup: true }, dedup: true };
    }
  } catch(_) {}

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

  // Reservar la idempotency key INMEDIATAMENTE — cualquier retry concurrent
  // que llegue antes de que terminen las 4 operaciones recibe este idEnvasado
  // y no ejecuta nada. TTL 120s.
  try { cache.put(keyEfectiva, idEnvasado, 120); } catch(_){}

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

// ════════════════════════════════════════════════════════════════════
// ANULAR ENVASADO MANUAL — operador corrige un error de captura
// ════════════════════════════════════════════════════════════════════
// Anula un envasado individual. Reverte stock base (suma) y stock
// derivado (resta) anulando los detalles de las guías correspondientes
// con anularDetalle (que ya tiene _conLock y registra el movimiento).
// Idempotente: si ya estaba anulado retorna yaAnulado:true.
function anularEnvasadoManual(params) {
  return _conLock('anularEnvasadoManual', function() {
    var idEnv = String((params && params.idEnvasado) || '').trim();
    var usuario = String((params && params.usuario) || 'manual');
    var motivo = String((params && params.motivo) || 'corrección manual');
    if (!idEnv) return { ok: false, error: 'Requiere idEnvasado' };

    var sheet = getSheet('ENVASADOS');
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h); });
    var iId    = hdrs.indexOf('idEnvasado');
    var iBase  = hdrs.indexOf('codigoProductoBase');
    var iCnt   = hdrs.indexOf('cantidadBase');
    var iDer   = hdrs.indexOf('codigoProductoEnvasado');
    var iUds   = hdrs.indexOf('unidadesProducidas');
    var iEst   = hdrs.indexOf('estado');
    var iGs    = hdrs.indexOf('idGuiaSalida');
    var iGi    = hdrs.indexOf('idGuiaIngreso');
    var iObs   = hdrs.indexOf('observacion');
    var iFecha = hdrs.indexOf('fecha');

    var rowIdx = -1, env = null;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iId]) === idEnv) {
        rowIdx = r + 1;
        env = {
          base:    String(data[r][iBase] || ''),
          cnt:     parseFloat(data[r][iCnt]) || 0,
          der:     String(data[r][iDer] || ''),
          uds:     parseFloat(data[r][iUds]) || 0,
          estado:  String(data[r][iEst] || '').toUpperCase(),
          idGs:    String(data[r][iGs] || ''),
          idGi:    String(data[r][iGi] || ''),
          fechaTs: data[r][iFecha] instanceof Date ? data[r][iFecha].getTime() : new Date(data[r][iFecha]).getTime()
        };
        break;
      }
    }
    if (rowIdx < 2) return { ok: false, error: 'Envasado no encontrado: ' + idEnv };
    if (env.estado === 'ANULADO' || env.estado === 'ANULADO_DUPLICADO') {
      return { ok: true, data: { idEnvasado: idEnv, yaAnulado: true, estado: env.estado } };
    }

    // Buscar el detalle de SALIDA (idGuiaSalida, codigo base, cantidad)
    // y de INGRESO (idGuiaIngreso, codigo derivado, unidades). Toleramos
    // que codigoProductoBase pueda venir como skuBase o codigoBarra: si
    // no encuentro por el código directo, busco cualquier detalle no
    // anulado con esa cantidad en la guía.
    function _buscarDet(idGuia, codigo, cantidad, refTs) {
      try {
        var dets = _sheetToObjects(getSheet('GUIA_DETALLE'));
        var cand = dets.filter(function(d) {
          if (String(d.idGuia) !== idGuia) return false;
          if (String(d.observacion || '').toUpperCase() === 'ANULADO') return false;
          if ((parseFloat(d.cantidadEsperada) || 0) !== cantidad) return false;
          if (codigo && String(d.codigoProducto) !== codigo) {
            // Si el código no coincide exacto, igual lo consideramos como
            // candidato — la cantidad + idGuia ya filtran bastante.
            return true;
          }
          return true;
        });
        // Ordenar por proximidad temporal al envasado (idDetalle = DET<ts>)
        cand.sort(function(a, b) {
          var ta = parseInt((String(a.idDetalle || '').match(/(\d{10,})/) || [])[1] || 0, 10);
          var tb = parseInt((String(b.idDetalle || '').match(/(\d{10,})/) || [])[1] || 0, 10);
          return Math.abs(ta - refTs) - Math.abs(tb - refTs);
        });
        return cand[0] || null;
      } catch(_) { return null; }
    }

    var resSal = null, resIng = null;
    var detSal = _buscarDet(env.idGs, env.base, env.cnt, env.fechaTs);
    var detIng = _buscarDet(env.idGi, env.der, env.uds, env.fechaTs);
    try { if (detSal) resSal = anularDetalle({ idDetalle: detSal.idDetalle, usuario: usuario }); } catch(e1) {}
    try { if (detIng) resIng = anularDetalle({ idDetalle: detIng.idDetalle, usuario: usuario }); } catch(e2) {}

    var nowStr = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    sheet.getRange(rowIdx, iEst + 1).setValue('ANULADO');
    var prevObs = sheet.getRange(rowIdx, iObs + 1).getValue();
    sheet.getRange(rowIdx, iObs + 1).setValue(
      String(prevObs || '') + ' | ANULADO por ' + usuario + ' · ' + motivo + ' · ' + nowStr
    );

    return {
      ok: true,
      data: {
        idEnvasado:    idEnv,
        estado:        'ANULADO',
        stockBaseRevertido:    !!(resSal && resSal.ok),
        stockDerivRevertido:   !!(resIng && resIng.ok),
        cantidadBaseRevertida: env.cnt,
        unidadesRevertidas:    env.uds
      }
    };
  });
}

// ════════════════════════════════════════════════════════════════════
// SCRIPT ONE-SHOT — limpiar envasados duplicados desde una fecha
// ════════════════════════════════════════════════════════════════════
// Antes de v2.8.0 registrarEnvasado no tenía idempotencia → un click +
// retry automático por timeout creaba 2-3 registros idénticos en ENVASADOS,
// inflaba el stock derivado y duplicaba el consumo de stock base.
//
// Este script:
//   1. Lee ENVASADOS desde fechaDesde (default 2026-05-01)
//   2. Agrupa por (usuario + codigoProductoBase + cantidadBase +
//      codigoProductoEnvasado + unidadesProducidas + idGuiaSalida)
//   3. En cada grupo con >1 registro CERCANOS (< 120 s), el primero queda
//      como bueno; los demás se anulan:
//        - anularDetalle(idDetalleSalida)   ← reverte stock base
//        - anularDetalle(idDetalleIngreso)  ← reverte stock derivado
//        - ENVASADOS.estado = 'ANULADO_DUPLICADO'
//        - ENVASADOS.observacion = 'duplicado de <idOriginal> · limpieza <ts>'
//
// Para encontrar el detalle correcto en GUIA_DETALLE (puede haber varios
// del mismo producto en la guía del día), correlacionamos por idGuia +
// codigoProducto + cantidad + timestamp del idDetalle (DET<timestamp>)
// cercano a la fecha del envasado duplicado.
//
// Ejecutar UNA VEZ desde el editor de Apps Script:
//   limpiarEnvasadosDuplicados()             → desde 2026-05-01
//   limpiarEnvasadosDuplicados('2026-05-10') → desde otra fecha
// Idempotente: si ya se anuló antes, no se anula otra vez (filtro
// estado != 'ANULADO_DUPLICADO' y observacion != 'ANULADO').
function limpiarEnvasadosDuplicados(fechaDesde) {
  var FECHA_DESDE = fechaDesde || '2026-05-01';
  var VENTANA_SEG = 120; // segundos para considerar duplicado del mismo grupo
  var sheet = getSheet('ENVASADOS');
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, data: { revisados: 0, anulados: 0, msg: 'ENVASADOS vacía' } };
  var hdrs  = data[0].map(function(h){ return String(h); });
  var iId    = hdrs.indexOf('idEnvasado');
  var iBase  = hdrs.indexOf('codigoProductoBase');
  var iCnt   = hdrs.indexOf('cantidadBase');
  var iDer   = hdrs.indexOf('codigoProductoEnvasado');
  var iUds   = hdrs.indexOf('unidadesProducidas');
  var iFecha = hdrs.indexOf('fecha');
  var iUsr   = hdrs.indexOf('usuario');
  var iEst   = hdrs.indexOf('estado');
  var iGs    = hdrs.indexOf('idGuiaSalida');
  var iGi    = hdrs.indexOf('idGuiaIngreso');
  var iObs   = hdrs.indexOf('observacion');
  if (iId < 0 || iBase < 0) return { ok: false, error: 'Columnas requeridas faltantes en ENVASADOS' };

  // 1. Recolectar candidatos (desde fechaDesde, estado COMPLETADO o vacío)
  var registros = [];
  for (var r = 1; r < data.length; r++) {
    var f = data[r][iFecha];
    var fStr = f instanceof Date
      ? Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(f || '').substring(0, 10);
    if (fStr < FECHA_DESDE) continue;
    var est = String(data[r][iEst] || '').toUpperCase();
    if (est === 'ANULADO_DUPLICADO' || est === 'ANULADO') continue;
    var fTs = f instanceof Date ? f.getTime() : new Date(f).getTime();
    if (isNaN(fTs)) continue;
    registros.push({
      rowIdx: r + 1, // 1-indexado para setValue
      idEnvasado:  String(data[r][iId] || ''),
      base:        String(data[r][iBase] || ''),
      cantBase:    parseFloat(data[r][iCnt]) || 0,
      derivado:    String(data[r][iDer] || ''),
      uds:         parseFloat(data[r][iUds]) || 0,
      ts:          fTs,
      usuario:     String(data[r][iUsr] || '').trim().toLowerCase(),
      idGs:        String(data[r][iGs] || ''),
      idGi:        String(data[r][iGi] || '')
    });
  }

  // 2. Agrupar por clave
  var grupos = {};
  registros.forEach(function(reg) {
    var key = [reg.usuario, reg.base, reg.cantBase, reg.derivado, reg.uds, reg.idGs].join('|');
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(reg);
  });

  // 3. Construir índice GUIA_DETALLE por (idGuia + codigoProducto) → lista de detalles
  //    para encontrar rápido el detalle correcto sin anular dos veces.
  var detSheet = getSheet('GUIA_DETALLE');
  var detData  = detSheet.getDataRange().getValues();
  var dHdrs    = detData[0].map(function(h){ return String(h); });
  var dIdDet   = dHdrs.indexOf('idDetalle');
  var dIdGuia  = dHdrs.indexOf('idGuia');
  var dCod     = dHdrs.indexOf('codigoProducto');
  var dCantE   = dHdrs.indexOf('cantidadEsperada');
  var dObs     = dHdrs.indexOf('observacion');
  var detIndex = {};
  for (var dr = 1; dr < detData.length; dr++) {
    var idGuia  = String(detData[dr][dIdGuia] || '');
    var codProd = String(detData[dr][dCod] || '');
    var cant    = parseFloat(detData[dr][dCantE]) || 0;
    var obs     = String(detData[dr][dObs] || '').toUpperCase();
    if (obs === 'ANULADO') continue; // ya anulado, no contemplar
    var idDet   = String(detData[dr][dIdDet] || '');
    // Extraer timestamp del idDetalle (formato 'DET<ts>')
    var detTs   = 0;
    var m = idDet.match(/(\d{10,})/);
    if (m) detTs = parseInt(m[1], 10);
    var k = idGuia + '|' + codProd + '|' + cant;
    if (!detIndex[k]) detIndex[k] = [];
    detIndex[k].push({ idDet: idDet, ts: detTs, _consumido: false });
  }

  // 4. Para cada grupo: ordenar por ts, marcar duplicados, anular detalles
  var reporte = {
    fechaDesde:     FECHA_DESDE,
    gruposRevisados: 0,
    duplicadosAnulados: 0,
    detallesAnulados: 0,
    erroresAnulacion: [],
    detalle: []
  };
  var nowStr = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

  Object.keys(grupos).forEach(function(key) {
    var arr = grupos[key];
    if (arr.length < 2) return;
    arr.sort(function(a, b){ return a.ts - b.ts; });
    var primero = arr[0];
    reporte.gruposRevisados++;
    for (var i = 1; i < arr.length; i++) {
      var dup = arr[i];
      var deltaSeg = Math.abs(dup.ts - primero.ts) / 1000;
      // Solo anular si está dentro de la ventana de duplicación
      if (deltaSeg > VENTANA_SEG) continue;

      // Resolver el codigoBarra real del base (puede que ENVASADOS guarde
      // skuBase/idProducto en codigoProductoBase pero GUIA_DETALLE usa el
      // codigoBarra). Buscamos el detalle por la cantidad y la guía.
      // Probamos las claves posibles: tal como está en ENVASADOS, y si no
      // matchea, recorremos los detalles de esa guía buscando por cantidad.
      function _buscarDet(idGuia, cant, refTs, codigoIntento) {
        // 1. Intento directo con el código guardado
        var k1 = idGuia + '|' + codigoIntento + '|' + cant;
        var bucket = detIndex[k1];
        // 2. Si no, buscar todos los detalles de esa guía con esa cantidad
        if (!bucket || !bucket.length) {
          bucket = [];
          Object.keys(detIndex).forEach(function(kk){
            if (kk.indexOf(idGuia + '|') === 0 && kk.lastIndexOf('|' + cant) === kk.length - String('|' + cant).length) {
              detIndex[kk].forEach(function(d){ bucket.push(d); });
            }
          });
        }
        // Filtrar no consumidos y elegir el más cercano por timestamp
        var cands = bucket.filter(function(d){ return !d._consumido; });
        if (!cands.length) return null;
        cands.sort(function(a, b){ return Math.abs(a.ts - refTs) - Math.abs(b.ts - refTs); });
        return cands[0];
      }

      var detSal = _buscarDet(dup.idGs, dup.cantBase, dup.ts, dup.base);
      var detIng = _buscarDet(dup.idGi, dup.uds,      dup.ts, dup.derivado);

      var resultSal = null, resultIng = null;
      try { if (detSal) { resultSal = anularDetalle({ idDetalle: detSal.idDet, usuario: 'limpieza-duplicados' }); if (resultSal && resultSal.ok) { detSal._consumido = true; reporte.detallesAnulados++; } } } catch(e1) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'salida', error: e1.message }); }
      try { if (detIng) { resultIng = anularDetalle({ idDetalle: detIng.idDet, usuario: 'limpieza-duplicados' }); if (resultIng && resultIng.ok) { detIng._consumido = true; reporte.detallesAnulados++; } } } catch(e2) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'ingreso', error: e2.message }); }

      // Marcar ENVASADO como anulado por duplicado
      try {
        sheet.getRange(dup.rowIdx, iEst + 1).setValue('ANULADO_DUPLICADO');
        var prevObs = sheet.getRange(dup.rowIdx, iObs + 1).getValue();
        sheet.getRange(dup.rowIdx, iObs + 1).setValue(
          String(prevObs || '') + ' | duplicado de ' + primero.idEnvasado + ' · limpieza ' + nowStr
        );
      } catch(eMark) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'marcado', error: eMark.message }); }

      reporte.duplicadosAnulados++;
      reporte.detalle.push({
        idEnvasado:    dup.idEnvasado,
        duplicaA:      primero.idEnvasado,
        deltaSeg:      Math.round(deltaSeg),
        producto:      dup.derivado,
        cantBase:      dup.cantBase,
        uds:           dup.uds,
        detSalAnulado: !!(resultSal && resultSal.ok),
        detIngAnulado: !!(resultIng && resultIng.ok)
      });
    }
  });

  Logger.log('limpiarEnvasadosDuplicados ✓ ' + JSON.stringify({
    desde: FECHA_DESDE, grupos: reporte.gruposRevisados,
    duplicadosAnulados: reporte.duplicadosAnulados,
    detallesAnulados: reporte.detallesAnulados,
    errores: reporte.erroresAnulacion.length
  }));
  return { ok: true, data: reporte };
}
