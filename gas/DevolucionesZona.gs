// ============================================================
// warehouseMos — DevolucionesZona.gs
// [v2.13.51] Flujo two-party witness para devoluciones ME → WH.
//
// Modelo de datos:
//   - DEVOLUCIONES_ZONA: tabla intermedia con payload_zona + payload_almacen
//     en JSON. SOLO registra el comparativo, NUNCA toca stock directamente.
//   - GUIAS tipo INGRESO_DEVOLUCION_ZONA: la guía REAL que materializa los
//     items que el operador WH confirma como ingreso. Esta SÍ suma stock.
//
// Flujo:
//   1. Zona ME envía devolución → crearDevolucionZona() → estado EN_TRANSITO
//      + payload_zona con [{codigo, cantidad, estado, motivo, foto}]
//   2. WH operador recibe → confirmarRecepcionDevolucion() → actualiza con
//      payload_almacen + crea guía INGRESO_DEVOLUCION_ZONA con SOLO los
//      items que el operador marca como INGRESAR_BUENO. El resto queda
//      registrado en payload_almacen pero NO entra a stock.
//   3. MOS admin revisa comparativo lado a lado y reconcilia.
// ============================================================

var _DEVZONA_HEADERS = [
  'idDevolucion','fechaInicio','zonaOrigen','vendedor','idDispositivoOrigen',
  'fechaRecepcion','operadorAlmacen','idDispositivoWH',
  'estado',  // EN_TRANSITO | RECEPCIONADO | RECONCILIADO | ANULADA
  'payload_zona',     // JSON: items declarados por zona
  'payload_almacen',  // JSON: items confirmados por almacén
  'diferenciasJson',  // JSON: auto-calculado para pintar comparativo
  'idGuiaIngresoGenerada',  // FK a GUIAS si almacén materializó ingreso
  'fotoZona','fotoAlmacen',
  'revisadoPor','fechaRevision','notaAdminMOS'
];

// Auto-crea la hoja con headers + freeze
function _asegurarHojaDevolucionesZona() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('DEVOLUCIONES_ZONA');
  if (!sh) {
    sh = ss.insertSheet('DEVOLUCIONES_ZONA');
    sh.appendRow(_DEVZONA_HEADERS);
    sh.getRange(1, 1, 1, _DEVZONA_HEADERS.length)
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#fbbf24');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Crea una devolución ME → WH con el payload declarado por la zona.
 * No toca stock (eso lo hace ME al confirmar su guía SALIDA_DEVOLUCION).
 *
 * params: {
 *   zonaOrigen, vendedor, idDispositivoOrigen, fotoZona,
 *   items: [{codigo, descripcion, cantidad, estado, motivo, foto?, notaItem?}]
 * }
 *
 * Estados zona: BUEN_ESTADO | ROTO_MALOGRADO | VENCIDO | DEFECTO_FABRICA |
 *               GORGOJOS_BICHOS | OTRO (con notaItem)
 */
function crearDevolucionZona(params) {
  var sh = _asegurarHojaDevolucionesZona();
  var idDev = _generateId('DV');
  var fechaInicio = new Date();
  var items = Array.isArray(params.items) ? params.items : [];
  if (!items.length) return { ok: false, error: 'Devolución sin items' };

  var payload_zona = {
    items: items.map(function(it) {
      return {
        codigo:      String(it.codigo || '').trim(),
        descripcion: String(it.descripcion || ''),
        cantidad:    parseFloat(it.cantidad) || 0,
        estado:      String(it.estado || 'BUEN_ESTADO').toUpperCase(),
        motivo:      String(it.motivo || ''),
        foto:        String(it.foto || ''),
        notaItem:    String(it.notaItem || '')
      };
    }),
    notaGeneral: String(params.notaGeneral || ''),
    timestamp:   fechaInicio.toISOString()
  };

  sh.appendRow([
    idDev,
    fechaInicio,
    String(params.zonaOrigen || ''),
    String(params.vendedor || ''),
    String(params.idDispositivoOrigen || ''),
    '', // fechaRecepcion vacía
    '', // operadorAlmacen vacío
    '', // idDispositivoWH vacío
    'EN_TRANSITO',
    JSON.stringify(payload_zona),
    '', // payload_almacen vacío
    '', // diferenciasJson vacío
    '', // idGuiaIngresoGenerada vacío
    String(params.fotoZona || ''),
    '', // fotoAlmacen vacía
    '', '', ''
  ]);

  // Notificar a operadores WH
  try {
    _notificarMOS(
      '📦 Devolución llegando',
      String(params.zonaOrigen || '?') + ' · ' + items.length + ' items · ' + String(params.vendedor || '?'),
      params.vendedor || null,
      'WH_DEVOLUCION_LLEGANDO'
    );
  } catch(e) { Logger.log('[crearDevolucionZona] push: ' + e.message); }

  return { ok: true, data: { idDevolucion: idDev, estado: 'EN_TRANSITO' } };
}

/**
 * Lista devoluciones por estado (para banner WH operador).
 */
function getDevolucionesZona(params) {
  var sh = _asegurarHojaDevolucionesZona();
  var d  = sh.getDataRange().getValues();
  if (d.length < 2) return { ok: true, data: [] };
  var h = d[0];
  var rows = [];
  for (var i = 1; i < d.length; i++) {
    var row = {};
    for (var c = 0; c < h.length; c++) row[h[c]] = d[i][c];
    // Filtrar por estado si viene
    if (params && params.estado) {
      var ests = String(params.estado).split(',').map(function(s){return s.trim();});
      if (ests.indexOf(String(row.estado)) < 0) continue;
    }
    if (params && params.zonaOrigen && String(row.zonaOrigen) !== String(params.zonaOrigen)) continue;
    // Parsear JSONs para frontend
    try { row.payload_zona     = row.payload_zona     ? JSON.parse(row.payload_zona)     : null; } catch(_) { row.payload_zona = null; }
    try { row.payload_almacen  = row.payload_almacen  ? JSON.parse(row.payload_almacen)  : null; } catch(_) { row.payload_almacen = null; }
    try { row.diferenciasJson  = row.diferenciasJson  ? JSON.parse(row.diferenciasJson)  : null; } catch(_) { row.diferenciasJson = null; }
    if (row.fechaInicio    instanceof Date) row.fechaInicio    = row.fechaInicio.toISOString();
    if (row.fechaRecepcion instanceof Date) row.fechaRecepcion = row.fechaRecepcion.toISOString();
    if (row.fechaRevision  instanceof Date) row.fechaRevision  = row.fechaRevision.toISOString();
    rows.push(row);
  }
  rows.sort(function(a, b) { return String(b.fechaInicio).localeCompare(String(a.fechaInicio)); });
  if (params && params.limit) rows = rows.slice(0, parseInt(params.limit));
  return { ok: true, data: rows };
}

/**
 * Detalle de una devolución específica.
 */
function getDevolucionDetalle(idDevolucion) {
  if (!idDevolucion) return { ok: false, error: 'idDevolucion requerido' };
  var sh = _asegurarHojaDevolucionesZona();
  var d  = sh.getDataRange().getValues();
  var h  = d[0];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) !== String(idDevolucion)) continue;
    var row = {};
    for (var c = 0; c < h.length; c++) row[h[c]] = d[i][c];
    try { row.payload_zona    = row.payload_zona    ? JSON.parse(row.payload_zona)    : null; } catch(_) {}
    try { row.payload_almacen = row.payload_almacen ? JSON.parse(row.payload_almacen) : null; } catch(_) {}
    try { row.diferenciasJson = row.diferenciasJson ? JSON.parse(row.diferenciasJson) : null; } catch(_) {}
    if (row.fechaInicio    instanceof Date) row.fechaInicio    = row.fechaInicio.toISOString();
    if (row.fechaRecepcion instanceof Date) row.fechaRecepcion = row.fechaRecepcion.toISOString();
    return { ok: true, data: row };
  }
  return { ok: false, error: 'Devolución no encontrada: ' + idDevolucion };
}

/**
 * Operador WH confirma recepción item-by-item.
 *
 * params: {
 *   idDevolucion, operadorAlmacen, idDispositivoWH, fotoAlmacen,
 *   items: [{codigo, cantidadRecibida, estadoAlmacen, comentario?, foto?, notaItem?}]
 * }
 *
 * Estados almacén: INGRESAR_BUENO | REPARABLE | MERMA | RECLAMAR_PROVEEDOR | OTRO
 *
 * Efecto:
 *  1. Actualiza fila DEVOLUCIONES_ZONA con payload_almacen + diferenciasJson
 *  2. Crea guía INGRESO_DEVOLUCION_ZONA con SOLO items estado=INGRESAR_BUENO
 *  3. Esa guía suma stock real (vía mecanismo existente de cerrar/aprobar guía)
 *  4. Items REPARABLE/MERMA/RECLAMAR_PROVEEDOR NO entran a stock, quedan
 *     solo registrados en el JSON del comparativo
 *  5. Marca devolución como RECEPCIONADO
 */
function confirmarRecepcionDevolucion(params) {
  return _conLock('confirmarRecepcionDevolucion', function() {
    var idDev = String(params.idDevolucion || '').trim();
    if (!idDev) return { ok: false, error: 'idDevolucion requerido' };
    var itemsAlmacen = Array.isArray(params.items) ? params.items : [];

    // Localizar fila
    var sh = _asegurarHojaDevolucionesZona();
    var d = sh.getDataRange().getValues();
    var h = d[0];
    var fila = -1, payloadZonaObj = null;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]) === idDev) {
        fila = i + 1;
        try { payloadZonaObj = d[i][h.indexOf('payload_zona')] ? JSON.parse(d[i][h.indexOf('payload_zona')]) : { items: [] }; }
        catch(_) { payloadZonaObj = { items: [] }; }
        var estadoActual = String(d[i][h.indexOf('estado')]);
        if (estadoActual !== 'EN_TRANSITO') {
          return { ok: false, error: 'Devolución ya procesada (estado: ' + estadoActual + ')' };
        }
        break;
      }
    }
    if (fila < 0) return { ok: false, error: 'Devolución no encontrada' };

    // Construir payload_almacen
    var payloadAlmacen = {
      items: itemsAlmacen.map(function(it) {
        return {
          codigo:           String(it.codigo || '').trim(),
          descripcion:      String(it.descripcion || ''),
          cantidadRecibida: parseFloat(it.cantidadRecibida) || 0,
          estado:           String(it.estado || 'INGRESAR_BUENO').toUpperCase(),
          comentario:       String(it.comentario || ''),
          foto:             String(it.foto || ''),
          notaItem:         String(it.notaItem || '')
        };
      }),
      operador:   String(params.operadorAlmacen || ''),
      deviceId:   String(params.idDispositivoWH || ''),
      timestamp:  new Date().toISOString()
    };

    // Calcular diferencias item-por-item (match por codigo)
    var difItems = [];
    var totalDif = 0, totalCambioEstado = 0;
    var zonaItemsByCodigo = {};
    payloadZonaObj.items.forEach(function(zit) { zonaItemsByCodigo[String(zit.codigo).toUpperCase()] = zit; });
    var alm = {};
    payloadAlmacen.items.forEach(function(ait) { alm[String(ait.codigo).toUpperCase()] = ait; });

    // Recorrer todos los codigos (zona ∪ almacén)
    var allCodes = {};
    Object.keys(zonaItemsByCodigo).forEach(function(k) { allCodes[k] = true; });
    Object.keys(alm).forEach(function(k) { allCodes[k] = true; });

    Object.keys(allCodes).forEach(function(codigo) {
      var z = zonaItemsByCodigo[codigo] || {};
      var a = alm[codigo] || {};
      var qZ = parseFloat(z.cantidad) || 0;
      var qA = parseFloat(a.cantidadRecibida) || 0;
      var delta = qA - qZ;
      // Comparar estados — zona usa BUEN_ESTADO/ROTO_*; almacén usa INGRESAR_BUENO/etc.
      // Para el comparativo, mapear ambos a "BUENO/NO_BUENO":
      var zBueno = String(z.estado || '') === 'BUEN_ESTADO';
      var aBueno = String(a.estado || '') === 'INGRESAR_BUENO';
      var cambioEstado = (zBueno !== aBueno);
      var severidad = 'OK';
      if (delta !== 0 && cambioEstado) severidad = 'ALTA';
      else if (delta !== 0 || cambioEstado) severidad = 'MEDIA';
      difItems.push({
        codigo: codigo,
        descripcion: a.descripcion || z.descripcion || '',
        cantidadZona: qZ,
        cantidadAlmacen: qA,
        deltaCantidad: delta,
        estadoZona: String(z.estado || ''),
        estadoAlmacen: String(a.estado || ''),
        cambioEstado: cambioEstado,
        severidad: severidad,
        comentarioZona: z.motivo || z.notaItem || '',
        comentarioAlmacen: a.comentario || a.notaItem || ''
      });
      if (delta !== 0) totalDif++;
      if (cambioEstado) totalCambioEstado++;
    });

    var diferencias = {
      items: difItems,
      resumen: {
        totalItemsConDiferencia:   totalDif,
        totalUnidadesPerdidas:     difItems.reduce(function(s, it) { return s + Math.min(0, it.deltaCantidad); }, 0),
        totalConCambioDeEstado:    totalCambioEstado,
        tieneDiscrepanciaGrave:    difItems.some(function(it) { return it.severidad === 'ALTA'; })
      }
    };

    // Generar guía INGRESO_DEVOLUCION_ZONA solo con items INGRESAR_BUENO
    var idGuia = '';
    var itemsParaIngreso = payloadAlmacen.items.filter(function(it) {
      return it.estado === 'INGRESAR_BUENO' && it.cantidadRecibida > 0;
    });
    if (itemsParaIngreso.length > 0) {
      try {
        var rGuia = crearGuia({
          tipo: 'INGRESO_DEVOLUCION_ZONA',
          usuario: payloadAlmacen.operador,
          idZona: '',  // origen es zona ME pero la guía es de ingreso a almacén
          numeroDocumento: idDev,
          comentario: 'Ingreso por devolución ME — origen ' + (d[fila-1][h.indexOf('zonaOrigen')] || '?')
        });
        if (rGuia.ok) {
          idGuia = rGuia.data.idGuia;
          // Agregar detalle item por item
          itemsParaIngreso.forEach(function(it) {
            try {
              _agregarDetalleGuiaImpl({
                idGuia: idGuia,
                codigoProducto: it.codigo,
                cantidadEsperada: it.cantidadRecibida,
                cantidadRecibida: it.cantidadRecibida,
                precioUnitario: 0,
                observacion: 'DEVOLUCION_ZONA: ' + (it.comentario || it.notaItem || '')
              });
            } catch(eD) { Logger.log('[confirmarRecepcion] agregar detalle: ' + eD.message); }
          });
        }
      } catch(eG) { Logger.log('[confirmarRecepcion] crear guía ingreso: ' + eG.message); }
    }

    // Update fila DEVOLUCIONES_ZONA
    var ts = new Date();
    sh.getRange(fila, h.indexOf('fechaRecepcion')         + 1).setValue(ts);
    sh.getRange(fila, h.indexOf('operadorAlmacen')        + 1).setValue(String(params.operadorAlmacen || ''));
    sh.getRange(fila, h.indexOf('idDispositivoWH')        + 1).setValue(String(params.idDispositivoWH || ''));
    sh.getRange(fila, h.indexOf('estado')                 + 1).setValue('RECEPCIONADO');
    sh.getRange(fila, h.indexOf('payload_almacen')        + 1).setValue(JSON.stringify(payloadAlmacen));
    sh.getRange(fila, h.indexOf('diferenciasJson')        + 1).setValue(JSON.stringify(diferencias));
    sh.getRange(fila, h.indexOf('idGuiaIngresoGenerada')  + 1).setValue(idGuia);
    if (params.fotoAlmacen) {
      sh.getRange(fila, h.indexOf('fotoAlmacen') + 1).setValue(String(params.fotoAlmacen));
    }

    // Push al vendedor + a MOS admins si hay diferencias
    if (diferencias.resumen.totalItemsConDiferencia > 0 || diferencias.resumen.totalConCambioDeEstado > 0) {
      try {
        _notificarMOS(
          '⚠ Devolución con diferencias',
          (d[fila-1][h.indexOf('zonaOrigen')] || '?') + ' · ' +
            diferencias.resumen.totalItemsConDiferencia + ' diferencias · ' +
            diferencias.resumen.totalConCambioDeEstado + ' cambios de estado',
          null,
          'WH_DEVOLUCION_DIFERENCIAS'
        );
      } catch(eP) { Logger.log('[confirmarRecepcion] push: ' + eP.message); }
    }

    return {
      ok: true,
      data: {
        idDevolucion: idDev,
        estado: 'RECEPCIONADO',
        idGuiaIngreso: idGuia,
        itemsIngresados: itemsParaIngreso.length,
        diferencias: diferencias.resumen
      }
    };
  });
}

/**
 * Admin MOS marca como reconciliada (cierre) tras revisar el comparativo.
 */
function reconciliarDevolucionZona(params) {
  var idDev = String(params.idDevolucion || '').trim();
  if (!idDev) return { ok: false, error: 'idDevolucion requerido' };
  var sh = _asegurarHojaDevolucionesZona();
  var d  = sh.getDataRange().getValues();
  var h  = d[0];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) !== idDev) continue;
    var fila = i + 1;
    var estActual = String(d[i][h.indexOf('estado')]);
    if (estActual !== 'RECEPCIONADO') {
      return { ok: false, error: 'Solo se puede reconciliar una devolución RECEPCIONADA (actual: ' + estActual + ')' };
    }
    sh.getRange(fila, h.indexOf('estado')        + 1).setValue('RECONCILIADO');
    sh.getRange(fila, h.indexOf('revisadoPor')   + 1).setValue(String(params.revisadoPor || ''));
    sh.getRange(fila, h.indexOf('fechaRevision') + 1).setValue(new Date());
    sh.getRange(fila, h.indexOf('notaAdminMOS')  + 1).setValue(String(params.nota || ''));
    return { ok: true, data: { idDevolucion: idDev, estado: 'RECONCILIADO' } };
  }
  return { ok: false, error: 'Devolución no encontrada' };
}
