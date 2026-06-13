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

  // Enriquecer con nombres de productos (indexar por codigoBarra, idProducto, skuBase y equivalentes)
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p){
    var name = p.descripcion || p.nombre || '';
    if (!name) return;
    if (p.codigoBarra) prodMap[String(p.codigoBarra)] = name;
    if (p.idProducto)  prodMap[String(p.idProducto)] = name;
    // skuBase solo para producto BASE (factor=1) — evita que presentaciones sobreescriban
    var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
    if (esBase && p.skuBase) prodMap[String(p.skuBase).trim().toUpperCase()] = name;
  });
  // Equivalentes → resuelven al producto base via skuBase
  try {
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      _sheetToObjects(equivSheet).forEach(function(e){
        if (!e.codigoBarra || !e.skuBase) return;
        var skuKey = String(e.skuBase).trim().toUpperCase();
        var name   = prodMap[skuKey];
        if (name) prodMap[String(e.codigoBarra)] = name;
      });
    }
  } catch(e) {}
  guia.detalle.forEach(function(d){
    d.descripcionProducto = d.descripcionProducto
      || prodMap[d.codigoProducto]
      || d.codigoProducto;
  });

  return { ok: true, data: guia };
}

function crearGuia(params) {
  // [v2.13.183] Bajo lock global: serializa contra el resto de operaciones de
  // guías/stock para evitar escrituras intercaladas. (La idempotencia ante
  // doble-click se cubre además en el frontend con guard de re-entrada y en
  // los flujos que crean guía: aprobarPreingreso por estado, crearDespachoRapido
  // por CacheService.)
  return _conLock('crearGuia', function() { return _crearGuiaImpl(params); });
}

function _crearGuiaImpl(params) {
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

  // [WH Fase 2 · PASO 2] dual-write de la guía RECIÉN creada a la sombra (best-effort, upsert por id_guia).
  // Cierra el gap: ahora una guía creada post-batch SÍ está en wh.guias, así el PATCH de estado al cerrarla
  // (cerrarGuia/reabrirGuia) sí aplica. Campos OCR ausentes = null en el INSERT (correcto: guía nueva sin OCR).
  // aprobarPreingreso reusa esta función (lock reentrante) → su guía también queda espejada.
  try {
    if (typeof _dualWriteWH === 'function') {
      _dualWriteWH('guias', {
        idGuia: idGuia, tipo: tipo, fecha: fecha, usuario: params.usuario || '',
        idProveedor: params.idProveedor || '', idZona: params.idZona || '',
        numeroDocumento: params.numeroDocumento || '', comentario: params.comentario || '',
        montoTotal: 0, estado: 'ABIERTA', idPreingreso: params.idPreingreso || '', foto: params.foto || ''
      });
    }
  } catch(_eDW) {}

  return { ok: true, data: { idGuia: idGuia, estado: 'ABIERTA' } };
}

function agregarDetalleGuia(params) {
  return _conLock('agregarDetalleGuia', function() {
    return _agregarDetalleGuiaImpl(params);
  });
}

// Helper para envolver una operación con LockService.
// Timeout 60s — Apps Script Web Apps tiene quota generosa de tiempo total,
// y el LockService es la única forma confiable de evitar race conditions
// entre POSTs paralelos. Operaciones individuales son rápidas (<1s típico),
// así que este timeout solo afecta picos de concurrencia extrema.
// [v2.13.183] REENTRANTE dentro de una misma ejecución. Como el lock es global
// de script, una función envuelta que llama a OTRA envuelta (ej. aprobarPreingreso
// → crearGuia, crearDespachoRapido → cerrarGuia) reusa el lock que ya tiene esta
// ejecución en vez de intentar re-adquirirlo (lo que se auto-bloquearía 60s).
// `_lockHeld` es por-ejecución (cada invocación GAS tiene su propio contexto),
// así que NO rompe la serialización entre POSTs concurrentes.
var _lockHeld = false;
function _conLock(nombre, fn) {
  if (_lockHeld) return fn();   // ya tenemos el lock en esta ejecución → reentrante
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
  } catch(eL) {
    Logger.log('[' + nombre + '] timeout lock tras 60s: ' + eL.message);
    return { ok: false, error: 'Sistema saturado, espera unos segundos y reintenta' };
  }
  _lockHeld = true;
  try { return fn(); }
  finally { _lockHeld = false; try { lock.releaseLock(); } catch(e) {} }
}

function _agregarDetalleGuiaImpl(params) {
  var sheet     = getSheet('GUIA_DETALLE');
  var idDetalle = _generateId('DET');

  // Forzar codigoBarra a string para preservar ceros a la izquierda
  var codigoBuscado = String(params.codigoProducto || '').trim().toUpperCase();

  // Validar que el código de producto existe (comparación case-insensitive)
  var productos = _sheetToObjects(getProductosSheet());
  var prod = productos.find(function(p) {
    return String(p.codigoBarra || '').trim().toUpperCase() === codigoBuscado;
  });

  // Si no está en PRODUCTOS_MASTER, buscar en EQUIVALENCIAS (MOS master SS, igual que el caché del frontend)
  if (!prod) {
    var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
    if (equivSheet) {
      var equivs = _sheetToObjects(equivSheet);
      var equiv  = equivs.find(function(e) {
        return String(e.codigoBarra || '').trim().toUpperCase() === codigoBuscado && _esActivo(e.activo);
      });
      if (equiv) {
        var skuB = String(equiv.skuBase || '').trim().toUpperCase();
        // Resolver al producto BASE (factor=1, activo). Evita matchear presentaciones
        // (factor>1) que comparten skuBase con su producto base.
        prod = productos.find(function(p) {
          var esBase = parseFloat(p.factorConversion || 1) === 1
                    && String(p.estado || '') !== '0';
          return esBase && (
            String(p.idProducto  || '').trim().toUpperCase() === skuB ||
            String(p.skuBase     || '').trim().toUpperCase() === skuB ||
            String(p.codigoBarra || '').trim().toUpperCase() === skuB
          );
        });
      }
    }
  }

  // Si no existe en ninguna tabla → sugerir ProductoNuevo
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
  // [v2.13.183] Cantidades negativas nunca son válidas (corromperían el stock).
  if (cantEsperada < 0 || cantRecibida < 0) {
    return { ok: false, error: 'Las cantidades no pueden ser negativas' };
  }

  // Preservar el código escaneado exacto (puede ser equiv o master) — stock es independiente por barcode
  var cbProd = codigoBuscado;

  // ── AUTO-SUMA: si ya existe detalle (mismo idGuia + cb, no anulado), sumar cantidad ──
  // Previene duplicados por reintentos, doble-click o flujos paralelos.
  // Si la guía está CERRADA, también ajusta el stock por el delta.
  var detData    = sheet.getDataRange().getValues();
  var hdrs       = detData[0];
  var idxIdG     = hdrs.indexOf('idGuia');
  var idxCb      = hdrs.indexOf('codigoProducto');
  var idxRec     = hdrs.indexOf('cantidadRecibida');
  var idxObs     = hdrs.indexOf('observacion');
  var idxIdDet   = hdrs.indexOf('idDetalle');
  for (var dr = 1; dr < detData.length; dr++) {
    if (String(detData[dr][idxIdG]) !== String(params.idGuia)) continue;
    if (String(detData[dr][idxCb]).toUpperCase() !== cbProd) continue;
    if (String(detData[dr][idxObs] || '').toUpperCase() === 'ANULADO') continue;

    // Match: sumar cantidades en el detalle existente
    var existingId  = String(detData[dr][idxIdDet]);
    var qtyAnterior = parseFloat(detData[dr][idxRec]) || 0;
    var qtyNueva    = qtyAnterior + cantRecibida;
    sheet.getRange(dr + 1, idxRec + 1).setValue(qtyNueva);

    // [v2.13.58 BUG A FIX] AUTO_SUMA debe sincronizar lote también.
    // Caso típico: 1er tap agrega sin fecha (idLote=''), 2do tap agrega con fecha.
    // Antes: la fecha se descartaba. Ahora: sync lote con cantidad acumulada.
    // Lee fecha actual del detalle si el call no la trae (preserva fecha previa).
    try {
      var idxLoteRow = hdrs.indexOf('idLote');
      var idxVencRow = hdrs.indexOf('fechaVencimiento');
      var fechaParam = params.fechaVencimiento || '';
      var fechaExist = idxVencRow >= 0 ? detData[dr][idxVencRow] : '';
      var idLotePrev = idxLoteRow >= 0 ? String(detData[dr][idxLoteRow] || '') : '';
      // Política: si el call trae fecha, gana; si no, mantiene la fecha existente del detalle
      var fechaUsar  = fechaParam || fechaExist;
      var fvDateAS;
      if (fechaUsar instanceof Date) fvDateAS = fechaUsar;
      else if (fechaUsar) fvDateAS = new Date(String(fechaUsar).substring(0, 10) + 'T12:00:00');
      else fvDateAS = '';
      // Si el call trae fecha distinta, escribirla en el detalle también
      if (fechaParam && idxVencRow >= 0) {
        sheet.getRange(dr + 1, idxVencRow + 1).setValue(fvDateAS);
      }
      // Sync lote si hay fecha o ya había lote (para actualizar cantidad)
      if ((fvDateAS && !isNaN(fvDateAS.getTime ? fvDateAS.getTime() : NaN)) || idLotePrev) {
        var resLoteAS = _sincronizarLoteDesdeDetalle({
          idLoteActual:   idLotePrev,
          codigoProducto: cbProd,
          cantidad:       qtyNueva,
          fechaVenc:      fvDateAS,
          idGuia:         params.idGuia,
          idDetalle:      existingId,
          usuario:        String(params.usuario || ''),
          motivo:         'auto_suma_detalle'
        });
        if (resLoteAS && resLoteAS.idLote && idxLoteRow >= 0 && resLoteAS.idLote !== idLotePrev) {
          sheet.getRange(dr + 1, idxLoteRow + 1).setNumberFormat('@').setValue(resLoteAS.idLote);
        }
      }
    } catch(eAS) { Logger.log('[autoSuma] sync lote fallo: ' + eAS.message); }

    // [v2.13.55] Política nueva: si guía CERRADA → en vez de AUTO_SUMA silenciosa,
    // crear AJUSTE EXPLÍCITO en hoja AJUSTES. Razones:
    //   1. Visibilidad: la edición queda como movimiento auditable separado
    //   2. Reconciliación: la suma teórica cuadra (AJUSTES + INGRESOS - SALIDAS)
    //   3. Idempotencia: si se reintenta por timeout, el AJUSTE no duplica
    //      stock (crearAjuste verifica duplicados en últimos 10s por usuario+motivo)
    var guiaInfo2 = _getGuiaInfo(params.idGuia);
    if (guiaInfo2 && String(guiaInfo2.estado).toUpperCase() === 'CERRADA'
        && cantRecibida !== 0 && !_esGuiaEnvasado(guiaInfo2.tipo)) {
      var esIngreso2 = String(guiaInfo2.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
      var deltaSum = esIngreso2 ? cantRecibida : -cantRecibida;
      try {
        if (typeof crearAjuste === 'function') {
          crearAjuste({
            codigoProducto: cbProd,
            tipoAjuste:     deltaSum > 0 ? 'INC' : 'DEC',
            cantidadAjuste: Math.abs(deltaSum),
            motivo:         'Edición guía cerrada · idGuia=' + params.idGuia +
                            ' · detalle=' + existingId + ' · +' + cantRecibida + 'u',
            usuario:        String(params.usuario || 'sistema')
          });
        } else {
          // Fallback al comportamiento legacy si crearAjuste no disponible
          _actualizarStock(cbProd, deltaSum, {
            tipoOperacion: 'EDICION_GUIA_CERRADA',
            origen:        existingId,
            usuario:       String(params.usuario || ''),
            observacion:   'idGuia=' + params.idGuia
          });
        }
      } catch(eAj) {
        Logger.log('[agregarDetalle] AJUSTE fallo, fallback a _actualizarStock: ' + eAj.message);
        _actualizarStock(cbProd, deltaSum, {
          tipoOperacion: 'EDICION_GUIA_CERRADA',
          origen:        existingId,
          usuario:       String(params.usuario || ''),
          observacion:   'idGuia=' + params.idGuia
        });
      }
    }

    return {
      ok: true,
      autoSumado: true,
      data: {
        idDetalle:          existingId,
        idGuia:             params.idGuia,
        codigoProducto:     cbProd,
        descripcionProducto: prod.descripcion || prod.nombre || prod.idProducto,
        cantidadEsperada:   cantEsperada,
        cantidadRecibida:   qtyNueva,
        precioUnitario:     precioUnit
      }
    };
  }

  // [v2.13.58 BUG B FIX] Lote desde fecha — usar _sincronizarLoteDesdeDetalle
  // en vez de appendRow directo. Razones:
  //   1. Reutiliza lote existente si ya hay uno (cod+guia+fecha) → no duplica
  //   2. Logea en LOTES_HISTORIAL accion=INSERT
  //   3. Aplica formato '@' a codigoBarra
  //   4. Mismo helper que cierre/edición → un solo camino canónico
  var idLote = String(params.idLote || '');
  var fechaVenc = params.fechaVencimiento || '';
  if (fechaVenc && fechaVenc !== '') {
    try {
      var fvDateNuevo = fechaVenc instanceof Date
        ? fechaVenc
        : new Date(String(fechaVenc).substring(0, 10) + 'T12:00:00');
      if (fvDateNuevo && !isNaN(fvDateNuevo.getTime())) {
        var resLoteNuevo = _sincronizarLoteDesdeDetalle({
          idLoteActual:   idLote,
          codigoProducto: cbProd,
          cantidad:       cantRecibida,
          fechaVenc:      fvDateNuevo,
          idGuia:         params.idGuia,
          idDetalle:      idDetalle,
          usuario:        String(params.usuario || ''),
          motivo:         'agregar_detalle_con_fecha'
        });
        if (resLoteNuevo && resLoteNuevo.idLote) idLote = resLoteNuevo.idLote;
      }
    } catch(eN) { Logger.log('[agregarDetalle] sync lote fallo: ' + eN.message); }
  }
  // FIX REFORZADO: setValues sobre rango pre-formateado (en vez de appendRow)
  // garantiza que codigoBarra preserve ceros a la izquierda. El LockService
  // ya serializa, así que no hay race entre llamadas paralelas.
  var nextRow  = sheet.getLastRow() + 1;
  // Pre-formatear col codigoBarra (3) e idLote (7) como texto
  sheet.getRange(nextRow, 3).setNumberFormat('@');
  if (idLote) sheet.getRange(nextRow, 7).setNumberFormat('@');
  // Escribir todos los valores en una sola operación
  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    idDetalle,
    params.idGuia,
    String(cbProd),
    cantEsperada,
    cantRecibida,
    precioUnit,
    String(idLote || ''),
    params.observacion || ''
  ]]);

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

// ============================================================
// [v2.11.2 Fix #4] _agregarDetallesBatch — versión bulk para
// guías recién creadas (despacho rápido, pickup→despacho).
// ------------------------------------------------------------
// Antes: items.forEach(agregarDetalleGuia) hacía N iteraciones
// con 1 lock + 2 lecturas full-sheet + 1 setValues cada una.
// Para 22 items y PRODUCTOS_MASTER con miles de rows, esto tarda
// 30-60s y el cliente PWA tira timeout → imprimirTicketGuia
// se ejecutaba con la hoja a medio escribir → ticket cortado a
// solo 2-3 productos.
//
// Esta versión hace:
//   1. UNA lectura de PRODUCTOS_MASTER (resuelve todos los códigos)
//   2. UNA lectura de EQUIVALENCIAS (solo si hay códigos no encontrados)
//   3. Consolida duplicados del mismo batch (auto-suma in-memory)
//   4. UN solo setValues con todas las filas
//   5. flush() para garantizar persistencia antes de seguir
//
// Pasa de O(N × lecturas) a O(1 × lecturas). Velocidad ×30 típica.
// Auto-suma con detalles preexistentes NO se soporta (asume guía nueva).
// ============================================================
function _agregarDetallesBatchImpl(idGuia, items, usuario) {
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  if (!items || !items.length) return { ok: true, data: { agregados: 0, errores: [], detalles: [] } };

  var sheet = getSheet('GUIA_DETALLE');

  // 1. Lectura ÚNICA de PRODUCTOS_MASTER → mapa por codigoBarra
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p) {
    var cb = String(p.codigoBarra || '').trim().toUpperCase();
    if (cb) prodMap[cb] = p;
  });

  // 2. Detectar códigos no encontrados → cargar EQUIVALENCIAS solo si hace falta
  var sinMatch = [];
  items.forEach(function(it) {
    var cb = String(it.codigoBarra || '').trim().toUpperCase();
    if (cb && !prodMap[cb]) sinMatch.push(cb);
  });
  if (sinMatch.length) {
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        var equivs = _sheetToObjects(equivSheet);
        // Buscar primero equivalencias activas que matcheen
        sinMatch.forEach(function(cb) {
          var eq = equivs.find(function(e) {
            return String(e.codigoBarra || '').trim().toUpperCase() === cb && _esActivo(e.activo);
          });
          if (!eq) return;
          var skuB = String(eq.skuBase || '').trim().toUpperCase();
          var pBase = productos.find(function(p) {
            var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
            return esBase && (
              String(p.idProducto  || '').trim().toUpperCase() === skuB ||
              String(p.skuBase     || '').trim().toUpperCase() === skuB ||
              String(p.codigoBarra || '').trim().toUpperCase() === skuB
            );
          });
          if (pBase) prodMap[cb] = pBase;
        });
      }
    } catch(_){}
  }

  // 3. Consolidar duplicados del batch (auto-suma in-memory).
  // [v2.13.58 BUG C FIX] Preservar fechaVencimiento. Si varios items del mismo
  // codigoBarra tienen fechas distintas, gana la PRIMERA (rara vez ocurre; el
  // operador agrupa por presentación). Se podría fragmentar pero rompería la
  // política existente de 1 detalle por (idGuia, codigoBarra).
  var consolidado = {};   // cb → { codigoBarra, cantidad, fechaVencimiento }
  var orden       = [];   // mantener orden de aparición
  items.forEach(function(it) {
    var cb  = String(it.codigoBarra || '').trim().toUpperCase();
    var qty = parseFloat(it.cantidad) || 0;
    if (!cb || qty <= 0) return;
    if (!consolidado[cb]) {
      consolidado[cb] = {
        codigoBarra: cb,
        cantidad: 0,
        fechaVencimiento: it.fechaVencimiento || ''
      };
      orden.push(cb);
    }
    consolidado[cb].cantidad += qty;
    // Si el primer item no traía fecha pero un item posterior sí → adoptar
    if (!consolidado[cb].fechaVencimiento && it.fechaVencimiento) {
      consolidado[cb].fechaVencimiento = it.fechaVencimiento;
    }
  });

  // 4. Construir filas en memoria
  var filas    = [];
  var detalles = [];
  var errores  = [];
  orden.forEach(function(cb) {
    var c = consolidado[cb];
    var prod = prodMap[cb];
    if (!prod) {
      errores.push(cb + ': producto no encontrado');
      return;
    }
    var idDetalle = _generateId('DET');
    filas.push([
      idDetalle, idGuia, String(cb),
      c.cantidad, c.cantidad,  // esperada = recibida (despacho rápido)
      0, '', ''                  // precio=0, idLote='', observacion=''
    ]);
    detalles.push({
      idDetalle: idDetalle, idGuia: idGuia,
      codigoProducto: cb, descripcionProducto: prod.descripcion || prod.nombre || prod.idProducto,
      cantidadEsperada: c.cantidad, cantidadRecibida: c.cantidad,
      fechaVencimiento: c.fechaVencimiento || ''     // [v2.13.58] para sync lote post-batch
    });
  });

  if (!filas.length) {
    Logger.log('[batch] idGuia=' + idGuia + ' · sin filas válidas · errores=' + errores.length);
    return { ok: true, data: { agregados: 0, errores: errores, detalles: [] } };
  }

  // 5. UN solo setValues + pre-formato col codigoBarra como texto
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 3, filas.length, 1).setNumberFormat('@');
  sheet.getRange(startRow, 1, filas.length, 8).setValues(filas);

  // 6. FLUSH crítico: garantiza que las filas sean visibles a las siguientes
  // lecturas (cerrarGuia + imprimirTicketGuia) ANTES de continuar.
  try { SpreadsheetApp.flush(); } catch(_){}

  // 7. [v2.13.58 BUG C FIX] Sincronizar lotes para items que trajeron fecha.
  // Bulk path = pickup, despacho rápido, devolución zona, aprobaciones masivas.
  // Antes esto se saltaba completamente y dejaba productos sin trazabilidad.
  // Solo recorre items con fechaVencimiento → costo cero si nadie trae fecha.
  var detallesConFecha = detalles.filter(function(d) {
    return d.fechaVencimiento && String(d.fechaVencimiento).trim() !== '';
  });
  if (detallesConFecha.length) {
    // Map idDetalle → rowSheet para persistir idLote en col 7
    var detRowMap = {};
    filas.forEach(function(_, idx) { detRowMap[detalles[idx].idDetalle] = startRow + idx; });
    detallesConFecha.forEach(function(d) {
      try {
        var fvDate = d.fechaVencimiento instanceof Date
          ? d.fechaVencimiento
          : new Date(String(d.fechaVencimiento).substring(0, 10) + 'T12:00:00');
        if (isNaN(fvDate.getTime())) return;
        var resLoteB = _sincronizarLoteDesdeDetalle({
          idLoteActual:   '',
          codigoProducto: d.codigoProducto,
          cantidad:       d.cantidadRecibida,
          fechaVenc:      fvDate,
          idGuia:         idGuia,
          idDetalle:      d.idDetalle,
          usuario:        String(usuario || ''),
          motivo:         'batch_add_con_fecha'
        });
        if (resLoteB && resLoteB.idLote) {
          var rowB = detRowMap[d.idDetalle];
          if (rowB) {
            sheet.getRange(rowB, 7).setNumberFormat('@').setValue(resLoteB.idLote);
            sheet.getRange(rowB, 8 + 0); // no-op para mantener referencias
          }
          // Persistir también la fecha en col fechaVencimiento si existe en schema
          var hdrsFinal = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          var colFv = hdrsFinal.indexOf('fechaVencimiento');
          if (colFv >= 0 && rowB) sheet.getRange(rowB, colFv + 1).setValue(fvDate);
        }
      } catch(eB) { Logger.log('[batch] sync lote ' + d.codigoProducto + ': ' + eB.message); }
    });
  }

  Logger.log('[batch] idGuia=' + idGuia + ' · escribí ' + filas.length + ' detalles · errores=' + errores.length +
             ' · lotes_sync=' + detallesConFecha.length);
  return { ok: true, data: { agregados: filas.length, errores: errores, detalles: detalles } };
}

// ── Actualizar cantidad recibida de un detalle existente ────────────
function actualizarCantidadDetalle(params) {
  return _conLock('actualizarCantidadDetalle', function() {
    return _actualizarCantidadDetalleImpl(params);
  });
}

function _actualizarCantidadDetalleImpl(params) {
  var idDetalle = String(params.idDetalle || '');
  var cantidad  = parseFloat(params.cantidadRecibida);
  if (!idDetalle || isNaN(cantidad)) return { ok: false, error: 'idDetalle y cantidadRecibida requeridos' };

  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxRec = hdrs.indexOf('cantidadRecibida');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCod = hdrs.indexOf('codigoProducto');
  if (idxId < 0 || idxRec < 0) return { ok: false, error: 'Columnas no encontradas' };

  var idxObs = hdrs.indexOf('observacion');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idDetalle) {
      var cantidadVieja = parseFloat(data[i][idxRec]) || 0;
      var idGuia        = String(data[i][idxIdG] || '');
      var codigo        = String(data[i][idxCod] || '');
      var obsActual     = idxObs >= 0 ? String(data[i][idxObs] || '').toUpperCase() : '';

      // [v2.13.55] Edición de cantidad en guía CERRADA → crear AJUSTE explícito
      // por la diferencia. Reemplaza el EDICION_CANTIDAD silencioso anterior.
      var guiaInfo = _getGuiaInfo(idGuia);
      if (guiaInfo && String(guiaInfo.estado).toUpperCase() === 'CERRADA' && codigo
          && !_esGuiaEnvasado(guiaInfo.tipo)) {
        var diff = cantidad - cantidadVieja;
        if (diff !== 0) {
          var esIngreso = String(guiaInfo.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
          var delta = esIngreso ? diff : -diff;
          try {
            if (typeof crearAjuste === 'function') {
              crearAjuste({
                codigoProducto: codigo,
                tipoAjuste:     delta > 0 ? 'INC' : 'DEC',
                cantidadAjuste: Math.abs(delta),
                motivo:         'Edición cantidad guía cerrada · idGuia=' + idGuia +
                                ' · detalle=' + idDetalle +
                                ' · ' + cantidadVieja + '→' + cantidad + 'u',
                usuario:        String(params.usuario || 'sistema')
              });
            } else {
              _actualizarStock(codigo, delta, {
                tipoOperacion: 'EDICION_CANTIDAD',
                origen:        idDetalle,
                usuario:       String(params.usuario || ''),
                observacion:   'idGuia=' + idGuia + ' diff=' + diff
              });
            }
          } catch(eAj) {
            Logger.log('[actualizarCantidad] AJUSTE fallo, fallback: ' + eAj.message);
            _actualizarStock(codigo, delta, {
              tipoOperacion: 'EDICION_CANTIDAD',
              origen:        idDetalle,
              usuario:       String(params.usuario || ''),
              observacion:   'idGuia=' + idGuia + ' diff=' + diff
            });
          }
        }
      }

      // Si estaba ANULADO y ahora se edita a cantidad > 0, des-anularlo
      if (obsActual === 'ANULADO' && cantidad > 0 && idxObs >= 0) {
        sheet.getRange(i + 1, idxObs + 1).setValue('');
      }

      sheet.getRange(i + 1, idxRec + 1).setValue(cantidad);

      // [v2.13.58 BUG D FIX] Sincronizar cantidad del lote asociado.
      // Antes: detalle 15u + lote 10u → FIFO rompía + reconciliación marcaba ajuste fantasma.
      // Ahora: si el detalle tiene idLote o fechaVencimiento, propaga la nueva cantidad.
      try {
        var idxLoteRowD = hdrs.indexOf('idLote');
        var idxVencRowD = hdrs.indexOf('fechaVencimiento');
        var idLotePrevD = idxLoteRowD >= 0 ? String(data[i][idxLoteRowD] || '') : '';
        var fechaPrevD  = idxVencRowD >= 0 ? data[i][idxVencRowD] : '';
        if ((idLotePrevD || fechaPrevD) && codigo) {
          var fvDateD;
          if (fechaPrevD instanceof Date) fvDateD = fechaPrevD;
          else if (fechaPrevD) fvDateD = new Date(String(fechaPrevD).substring(0, 10) + 'T12:00:00');
          else fvDateD = '';
          // Si no hay fecha pero sí lote previo, _sincronizarLoteDesdeDetalle CASO C anularía.
          // Solo sincronizamos cantidad si hay fecha real (preserva lote activo).
          if (fvDateD && !isNaN(fvDateD.getTime())) {
            var resLoteD = _sincronizarLoteDesdeDetalle({
              idLoteActual:   idLotePrevD,
              codigoProducto: codigo,
              cantidad:       cantidad,
              fechaVenc:      fvDateD,
              idGuia:         idGuia,
              idDetalle:      idDetalle,
              usuario:        String(params.usuario || ''),
              motivo:         'edit_cantidad_detalle ' + cantidadVieja + '→' + cantidad
            });
            if (resLoteD && resLoteD.idLote && idxLoteRowD >= 0 && resLoteD.idLote !== idLotePrevD) {
              sheet.getRange(i + 1, idxLoteRowD + 1).setNumberFormat('@').setValue(resLoteD.idLote);
            }
          }
        }
      } catch(eD) { Logger.log('[actualizarCantidad] sync lote fallo: ' + eD.message); }

      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado: ' + idDetalle };
}

// Actualizar precios unitarios de varias líneas de una guía en batch.
// params: { idGuia, items: [{ idDetalle, precioUnitario }] }
function actualizarPreciosDetalle(params) {
  return _conLock('actualizarPreciosDetalle', function() { return _actualizarPreciosDetalleImpl(params); });
}
function _actualizarPreciosDetalleImpl(params) {
  if (!params.idGuia || !Array.isArray(params.items) || !params.items.length) {
    return { ok: false, error: 'idGuia + items[] requeridos' };
  }
  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxPU  = hdrs.indexOf('precioUnitario');
  if (idxId < 0 || idxPU < 0) return { ok: false, error: 'Columnas no encontradas en GUIA_DETALLE' };

  // Mapa idDetalle → precio nuevo
  var mapPrecio = {};
  params.items.forEach(function(it) {
    if (it.idDetalle && it.precioUnitario !== undefined && it.precioUnitario !== '' && !isNaN(parseFloat(it.precioUnitario))) {
      mapPrecio[String(it.idDetalle)] = parseFloat(it.precioUnitario);
    }
  });

  var actualizados = 0;
  var sumaTotal   = 0;
  var idxCantR    = hdrs.indexOf('cantidadRecibida');
  var idxCantE    = hdrs.indexOf('cantidadEsperada');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdG]) !== String(params.idGuia)) continue;
    var idDet = String(data[i][idxId]);
    if (mapPrecio.hasOwnProperty(idDet)) {
      sheet.getRange(i + 1, idxPU + 1).setValue(mapPrecio[idDet]);
      actualizados++;
    }
    // Sumar para recalcular monto total de la guía
    var cant = (idxCantR >= 0 ? parseFloat(data[i][idxCantR]) : 0) || (idxCantE >= 0 ? parseFloat(data[i][idxCantE]) : 0) || 0;
    var precio = mapPrecio[idDet] !== undefined ? mapPrecio[idDet] : (parseFloat(data[i][idxPU]) || 0);
    sumaTotal += cant * precio;
  }

  // Actualizar montoTotal de la guía
  try {
    var shGuias = getSheet('GUIAS');
    var dataG = shGuias.getDataRange().getValues();
    var hdrsG = dataG[0];
    var idxGid = hdrsG.indexOf('idGuia');
    var idxGmt = hdrsG.indexOf('montoTotal');
    if (idxGid >= 0 && idxGmt >= 0) {
      for (var j = 1; j < dataG.length; j++) {
        if (String(dataG[j][idxGid]) === String(params.idGuia)) {
          shGuias.getRange(j + 1, idxGmt + 1).setValue(Math.round(sumaTotal * 100) / 100);
          break;
        }
      }
    }
  } catch(_){}

  return { ok: true, data: { actualizados: actualizados, montoTotalNuevo: Math.round(sumaTotal * 100) / 100 } };
}

// Helper: lee la guía y retorna {tipo, estado} o null si no existe
function _getGuiaInfo(idGuia) {
  var sheet   = getSheet('GUIAS');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxId   = hdrs.indexOf('idGuia');
  var idxTipo = hdrs.indexOf('tipo');
  var idxEst  = hdrs.indexOf('estado');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      return { tipo: String(data[i][idxTipo] || ''), estado: String(data[i][idxEst] || '') };
    }
  }
  return null;
}

// ── Actualizar fecha vencimiento de detalle + SINCRONIZAR LOTE ─────
// [v2.13.53] Antes solo escribía la fecha en GUIA_DETALLE sin crear lote.
// Resultado: detalles con fecha pero sin fila en LOTES_VENCIMIENTO →
// invisible para alertas/FIFO. Ahora crea/actualiza/anula lote según caso.
//
// Política confirmada por usuario:
//   - Clave única: (codigoProducto, fechaVencimiento) por guía
//   - Si edita fecha → actualiza la misma fila
//   - Si borra fecha → marca lote como ANULADO
//   - Lote NO se duplica si re-aplica misma fecha
function actualizarFechaVencimiento(params) {
  return _conLock('actualizarFechaVencimiento', function() { return _actualizarFechaVencimientoImpl(params); });
}
function _actualizarFechaVencimientoImpl(params) {
  var idDetalle = String(params.idDetalle || '');
  var fechaVencRaw = params.fechaVencimiento || '';
  if (!idDetalle) return { ok: false, error: 'idDetalle requerido' };

  var sheet  = getSheet('GUIA_DETALLE');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId   = hdrs.indexOf('idDetalle');
  var idxVenc = hdrs.indexOf('fechaVencimiento');
  var idxLote = hdrs.indexOf('idLote');
  var idxCod  = hdrs.indexOf('codigoProducto');
  var idxRec  = hdrs.indexOf('cantidadRecibida');
  var idxIdG  = hdrs.indexOf('idGuia');
  if (idxId < 0 || idxVenc < 0) return { ok: false, error: 'Columnas no encontradas' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) !== idDetalle) continue;

    // 1) Escribir la fecha en el detalle (o limpiar si vacía)
    var val = fechaVencRaw ? new Date(fechaVencRaw + 'T12:00:00') : '';
    sheet.getRange(i + 1, idxVenc + 1).setValue(val);

    // 2) Sincronizar LOTES_VENCIMIENTO
    var codigoProducto = String(data[i][idxCod] || '').trim();
    var cantidad       = parseFloat(data[i][idxRec]) || 0;
    var idGuia         = String(data[i][idxIdG] || '');
    var idLoteActual   = idxLote >= 0 ? String(data[i][idxLote] || '') : '';
    var resLote = _sincronizarLoteDesdeDetalle({
      idLoteActual:   idLoteActual,
      codigoProducto: codigoProducto,
      cantidad:       cantidad,
      fechaVenc:      val,
      idGuia:         idGuia,
      idDetalle:      idDetalle,
      usuario:        String(params.usuario || ''),
      motivo:         'edit_fecha_venc_desde_detalle'
    });
    if (resLote && resLote.idLote && idxLote >= 0 && resLote.idLote !== idLoteActual) {
      sheet.getRange(i + 1, idxLote + 1).setNumberFormat('@').setValue(resLote.idLote);
    }
    return { ok: true, data: { idLote: resLote && resLote.idLote, accion: resLote && resLote.accion } };
  }
  return { ok: false, error: 'Detalle no encontrado: ' + idDetalle };
}

// [v2.13.53] Helper canónico: sincroniza un lote desde la info de un detalle.
// Política: clave (codigoProducto, fechaVencimiento) por guía.
//
// Casos:
//   A) idLoteActual vacío + fecha nueva con valor → INSERT nuevo lote
//   B) idLoteActual vacío + fecha vacía           → no-op (nunca tuvo lote)
//   C) idLoteActual existe + fecha vacía          → ANULAR lote
//   D) idLoteActual existe + fecha = misma del lote → no-op
//   E) idLoteActual existe + fecha distinta       → UPDATE fecha+cantidad del lote
//
// Devuelve: {idLote, accion: INSERT|UPDATE|ANULAR|NOOP}
function _sincronizarLoteDesdeDetalle(opts) {
  var sheet = getSheet('LOTES_VENCIMIENTO');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var idxId    = hdrs.indexOf('idLote');
  var idxCod   = hdrs.indexOf('codigoProducto');
  var idxVenc  = hdrs.indexOf('fechaVencimiento');
  var idxQI    = hdrs.indexOf('cantidadInicial');
  var idxQA    = hdrs.indexOf('cantidadActual');
  var idxGuia  = hdrs.indexOf('idGuia');
  var idxEst   = hdrs.indexOf('estado');

  var fechaVenc = opts.fechaVenc;
  var fechaVencStr = (fechaVenc instanceof Date && !isNaN(fechaVenc.getTime()))
    ? Utilities.formatDate(fechaVenc, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : '';
  var idLoteActual = String(opts.idLoteActual || '');

  // CASO C: idLoteActual existe + fecha vacía → ANULAR
  if (idLoteActual && !fechaVencStr) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxId]) !== idLoteActual) continue;
      if (idxEst >= 0) sheet.getRange(i + 1, idxEst + 1).setValue('ANULADO');
      _logMovimientoLote({
        idLote: idLoteActual, codigoProducto: opts.codigoProducto, idGuia: opts.idGuia,
        accion: 'ANULAR', cantidad: 0, motivo: opts.motivo || 'fecha eliminada', usuario: opts.usuario
      });
      return { idLote: idLoteActual, accion: 'ANULAR' };
    }
    return { idLote: idLoteActual, accion: 'NOOP' };
  }

  // CASO B: sin lote y sin fecha → nada
  if (!idLoteActual && !fechaVencStr) return { idLote: '', accion: 'NOOP' };

  // CASO E o D: tiene lote y fecha (puede ser igual o distinta)
  if (idLoteActual) {
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][idxId]) !== idLoteActual) continue;
      var fechaActualLote = data[j][idxVenc];
      var fechaActualStr = fechaActualLote instanceof Date
        ? Utilities.formatDate(fechaActualLote, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaActualLote || '').substring(0, 10);
      // D: fecha igual y cantidad igual → no-op (excepto si lote estaba ANULADO, reactivar)
      var estActual = String(data[j][idxEst] || '').toUpperCase();
      var cantActual = parseFloat(data[j][idxQA]) || 0;
      if (fechaActualStr === fechaVencStr && cantActual === opts.cantidad && estActual === 'ACTIVO') {
        return { idLote: idLoteActual, accion: 'NOOP' };
      }
      // E: actualizar fecha y cantidad
      if (idxVenc >= 0) sheet.getRange(j + 1, idxVenc + 1).setValue(opts.fechaVenc);
      if (idxQI   >= 0) sheet.getRange(j + 1, idxQI + 1).setValue(opts.cantidad);
      if (idxQA   >= 0) sheet.getRange(j + 1, idxQA + 1).setValue(opts.cantidad);
      if (idxEst  >= 0 && estActual !== 'ACTIVO') sheet.getRange(j + 1, idxEst + 1).setValue('ACTIVO');
      _logMovimientoLote({
        idLote: idLoteActual, codigoProducto: opts.codigoProducto, idGuia: opts.idGuia,
        accion: 'UPDATE', cantidad: opts.cantidad,
        motivo: 'fecha=' + fechaVencStr + ' (era ' + fechaActualStr + ')', usuario: opts.usuario
      });
      return { idLote: idLoteActual, accion: 'UPDATE' };
    }
  }

  // CASO A: no tenía lote, hay fecha → buscar lote previo (codigoProducto + idGuia + fecha
  // exacta) por si ya existe otro detalle con misma combinación; reusar idLote.
  for (var k = 1; k < data.length; k++) {
    if (String(data[k][idxCod]).toUpperCase() !== String(opts.codigoProducto).toUpperCase()) continue;
    if (idxGuia >= 0 && String(data[k][idxGuia]) !== String(opts.idGuia)) continue;
    var fLote = data[k][idxVenc];
    var fStr  = fLote instanceof Date
      ? Utilities.formatDate(fLote, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fLote || '').substring(0, 10);
    if (fStr !== fechaVencStr) continue;
    var idReuse = String(data[k][idxId]);
    // Reactivar si estaba ANULADO
    if (idxEst >= 0) sheet.getRange(k + 1, idxEst + 1).setValue('ACTIVO');
    if (idxQI  >= 0) sheet.getRange(k + 1, idxQI + 1).setValue(opts.cantidad);
    if (idxQA  >= 0) sheet.getRange(k + 1, idxQA + 1).setValue(opts.cantidad);
    _logMovimientoLote({
      idLote: idReuse, codigoProducto: opts.codigoProducto, idGuia: opts.idGuia,
      accion: 'REUSE', cantidad: opts.cantidad, motivo: 'mismo (cod,guia,fecha)', usuario: opts.usuario
    });
    return { idLote: idReuse, accion: 'UPDATE' };
  }

  // Crear nuevo lote
  var nuevoId = _generateId('LOT');
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([nuevoId, String(opts.codigoProducto), opts.fechaVenc, opts.cantidad, opts.cantidad,
                   String(opts.idGuia), 'ACTIVO', new Date()]);
  sheet.getRange(nextRow, 2).setNumberFormat('@').setValue(String(opts.codigoProducto));
  _logMovimientoLote({
    idLote: nuevoId, codigoProducto: opts.codigoProducto, idGuia: opts.idGuia,
    accion: 'INSERT', cantidad: opts.cantidad, motivo: opts.motivo || '', usuario: opts.usuario
  });
  return { idLote: nuevoId, accion: 'INSERT' };
}

// [v2.13.53] Log de movimientos del lote en hoja LOTES_HISTORIAL.
// Para auditoría completa: quién, cuándo, qué cambió, en qué guía.
function _logMovimientoLote(opts) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('LOTES_HISTORIAL');
    if (!sh) {
      sh = ss.insertSheet('LOTES_HISTORIAL');
      sh.appendRow(['ts','idLote','codigoProducto','idGuia','accion','cantidad','motivo','usuario']);
      sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0f172a').setFontColor('#22d3ee');
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(),
      String(opts.idLote || ''),
      String(opts.codigoProducto || ''),
      String(opts.idGuia || ''),
      String(opts.accion || ''),
      parseFloat(opts.cantidad) || 0,
      String(opts.motivo || ''),
      String(opts.usuario || '')
    ]);
  } catch(e) { Logger.log('[_logMovimientoLote] ' + e.message); }
}

// [v2.13.54] Consume lotes FIFO real para una SALIDA.
//
// Política FIFO (vence primero → sale primero):
//   1. Carga lotes ACTIVOS con cantidadActual > 0 del producto
//   2. Ordena ASC por fechaVencimiento (los sin fecha al final = más viejos virtuales)
//   3. Va descontando hasta cubrir la cantidad pedida
//   4. Por cada lote consumido: actualiza LOTES_VENCIMIENTO.cantidadActual
//      + registra en LOTES_HISTORIAL accion=CONSUMO con la cantidad consumida
//   5. Si el lote queda en 0 → marca estado=AGOTADO (mantiene historial)
//   6. Si la cantidad pedida > suma de lotes disponibles → devuelve huérfano
//      (no se registra como lote, queda como "consumo sin lote" del producto)
//
// Devuelve: { lotesConsumidos: [{idLote, cantidad, fechaVenc}], huerfano: 0 o N }
function _consumirLotesFIFO(codigoProducto, cantidadPedida, idGuia, usuario, motivo) {
  if (cantidadPedida <= 0) return { lotesConsumidos: [], huerfano: 0 };
  var sheet = getSheet('LOTES_VENCIMIENTO');
  if (!sheet) return { lotesConsumidos: [], huerfano: cantidadPedida };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { lotesConsumidos: [], huerfano: cantidadPedida };
  var hdrs = data[0];
  var idxId    = hdrs.indexOf('idLote');
  var idxCod   = hdrs.indexOf('codigoProducto');
  var idxVenc  = hdrs.indexOf('fechaVencimiento');
  var idxQA    = hdrs.indexOf('cantidadActual');
  var idxEst   = hdrs.indexOf('estado');

  // Recolectar candidatos
  var candidatos = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxCod]).toUpperCase() !== String(codigoProducto).toUpperCase()) continue;
    var est = String(data[i][idxEst] || '').toUpperCase();
    if (est !== 'ACTIVO') continue;
    var cant = parseFloat(data[i][idxQA]) || 0;
    if (cant <= 0) continue;
    var fv = data[i][idxVenc];
    var fvTime = (fv instanceof Date && !isNaN(fv.getTime())) ? fv.getTime() : Number.MAX_SAFE_INTEGER;
    candidatos.push({
      fila:   i + 1,
      idLote: String(data[i][idxId]),
      cantidadDisponible: cant,
      fechaVenc: fv instanceof Date ? fv : null,
      fvTime: fvTime
    });
  }
  // Orden FIFO real: vence antes → primero. Sin fecha al final.
  candidatos.sort(function(a, b) { return a.fvTime - b.fvTime; });

  // Consumir
  var lotesConsumidos = [];
  var restante = cantidadPedida;
  for (var c = 0; c < candidatos.length && restante > 0; c++) {
    var cand = candidatos[c];
    var consumir = Math.min(cand.cantidadDisponible, restante);
    var nuevaCantLote = cand.cantidadDisponible - consumir;
    // Update cantidadActual + estado si queda 0
    sheet.getRange(cand.fila, idxQA + 1).setValue(nuevaCantLote);
    if (nuevaCantLote <= 0 && idxEst >= 0) {
      sheet.getRange(cand.fila, idxEst + 1).setValue('AGOTADO');
    }
    // Log
    _logMovimientoLote({
      idLote: cand.idLote, codigoProducto: codigoProducto, idGuia: idGuia,
      accion: 'CONSUMO', cantidad: consumir,
      motivo: motivo || ('salida FIFO restante=' + restante),
      usuario: usuario
    });
    lotesConsumidos.push({
      idLote: cand.idLote, cantidad: consumir,
      fechaVencimiento: cand.fechaVenc ? cand.fechaVenc.toISOString() : null
    });
    restante -= consumir;
  }
  return { lotesConsumidos: lotesConsumidos, huerfano: restante };
}

// [v2.13.53] Devuelve lotes ACTIVOS de un producto ordenados por fechaVencimiento ASC.
// Útil para FIFO real al despachar: el lote que vence primero sale primero.
// Si cantidadActual > 0 está disponible para consumo.
function getLotesFIFO(params) {
  var codigoProducto = String(params.codigoProducto || '').trim();
  if (!codigoProducto) return { ok: false, error: 'codigoProducto requerido' };
  var sheet = getSheet('LOTES_VENCIMIENTO');
  if (!sheet) return { ok: true, data: [] };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, data: [] };
  var hdrs = data[0];
  var idxId   = hdrs.indexOf('idLote');
  var idxCod  = hdrs.indexOf('codigoProducto');
  var idxVenc = hdrs.indexOf('fechaVencimiento');
  var idxQA   = hdrs.indexOf('cantidadActual');
  var idxEst  = hdrs.indexOf('estado');
  var idxGuia = hdrs.indexOf('idGuia');
  var idxCre  = hdrs.indexOf('fechaCreacion');
  var hoy = new Date();
  var lotes = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxCod]).toUpperCase() !== codigoProducto.toUpperCase()) continue;
    var est = String(data[i][idxEst] || '').toUpperCase();
    if (est !== 'ACTIVO') continue;
    var cant = parseFloat(data[i][idxQA]) || 0;
    if (cant <= 0) continue;
    var fv = data[i][idxVenc];
    var diasRestantes = null;
    if (fv instanceof Date && !isNaN(fv.getTime())) {
      diasRestantes = Math.ceil((fv.getTime() - hoy.getTime()) / (1000 * 60 * 60 * 24));
    }
    lotes.push({
      idLote:           String(data[i][idxId]),
      codigoProducto:   String(data[i][idxCod]),
      fechaVencimiento: fv instanceof Date ? fv.toISOString() : String(fv || ''),
      cantidadActual:   cant,
      idGuia:           String(data[i][idxGuia] || ''),
      diasRestantes:    diasRestantes,
      fechaCreacion:    data[i][idxCre] instanceof Date ? data[i][idxCre].toISOString() : String(data[i][idxCre] || '')
    });
  }
  // Orden FIFO: el que vence primero → primero (lotes sin fecha al final)
  lotes.sort(function(a, b) {
    if (!a.fechaVencimiento && !b.fechaVencimiento) return 0;
    if (!a.fechaVencimiento) return 1;
    if (!b.fechaVencimiento) return -1;
    return new Date(a.fechaVencimiento).getTime() - new Date(b.fechaVencimiento).getTime();
  });
  return { ok: true, data: lotes };
}

// [v2.13.53] Historial completo de un lote — para trazabilidad UI.
function getHistorialLote(params) {
  var idLote = String(params.idLote || '').trim();
  if (!idLote) return { ok: false, error: 'idLote requerido' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOTES_HISTORIAL');
  if (!sh) return { ok: true, data: [] };
  var d = sh.getDataRange().getValues();
  if (d.length < 2) return { ok: true, data: [] };
  var h = d[0];
  var idxTs = h.indexOf('ts');
  var idxId = h.indexOf('idLote');
  var rows = [];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][idxId]) !== idLote) continue;
    var r = {};
    for (var c = 0; c < h.length; c++) r[h[c]] = d[i][c];
    if (r.ts instanceof Date) r.ts = r.ts.toISOString();
    rows.push(r);
  }
  rows.sort(function(a, b) { return String(b.ts).localeCompare(String(a.ts)); });
  return { ok: true, data: rows };
}

// ── Anular un ítem de detalle ────────────────────────────
// Si la guía padre está CERRADA, devuelve el stock que ya se había aplicado.
// Si está ABIERTA, solo marca como anulado (stock aún no descontado).
function anularDetalle(params) {
  return _conLock('anularDetalle', function() {
    return _anularDetalleImpl(params);
  });
}

function _anularDetalleImpl(params) {
  var idDetalle = params.idDetalle;
  if (!idDetalle) return { ok: false, error: 'idDetalle requerido' };

  var sheet = getSheet('GUIA_DETALLE');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId  = hdrs.indexOf('idDetalle');
  var idxObs = hdrs.indexOf('observacion');
  var idxRec = hdrs.indexOf('cantidadRecibida');
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCod = hdrs.indexOf('codigoProducto');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] === idDetalle) {
      // Idempotencia: si ya está ANULADO, no devolver stock de nuevo
      if (String(data[i][idxObs] || '') === 'ANULADO') {
        return { ok: true, yaAnulado: true };
      }

      var cantidadActual = idxRec >= 0 ? (parseFloat(data[i][idxRec]) || 0) : 0;
      var idGuia         = String(data[i][idxIdG] || '');
      var codigo         = String(data[i][idxCod] || '');

      // Si la guía está CERRADA y tenía cantidad > 0, devolver stock.
      // EXCEPCIÓN: guías de envasado manejan stock directamente desde Envasados.gs
      // (no via cerrarGuia), por lo tanto anular detalle NO debe revertir stock.
      var guiaInfo = _getGuiaInfo(idGuia);
      if (guiaInfo && String(guiaInfo.estado).toUpperCase() === 'CERRADA'
          && cantidadActual > 0 && codigo
          && !_esGuiaEnvasado(guiaInfo.tipo)) {
        var esIngreso = String(guiaInfo.tipo || '').toUpperCase().indexOf('INGRESO') === 0;
        // Reverso: si era INGRESO, suma se revierte (resta); si SALIDA, devuelve (suma)
        var delta = esIngreso ? -cantidadActual : cantidadActual;
        _actualizarStock(codigo, delta, {
          tipoOperacion: 'ANULACION_DETALLE',
          origen:        idDetalle,
          usuario:       String(params.usuario || ''),
          observacion:   'idGuia=' + idGuia
        });
      }

      // Si era una línea de PN pendiente, marcar el PN correspondiente como ANULADO
      // (de lo contrario MOS lo sigue mostrando para aprobación)
      var obsActual = String(data[i][idxObs] || '').toUpperCase();
      if (obsActual === 'PN_PENDIENTE' && idGuia && codigo) {
        try { _anularPNPorGuiaYCodigo(idGuia, codigo); } catch(e) {
          Logger.log('No se pudo anular PN huérfano: ' + e.message);
        }
      }

      // Marcar como anulado y poner cantidad en 0
      sheet.getRange(i + 1, idxObs + 1).setValue('ANULADO');
      if (idxRec >= 0) sheet.getRange(i + 1, idxRec + 1).setValue(0);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Detalle no encontrado' };
}

// Marca como ANULADO los PNs PENDIENTE asociados a (idGuia, codigoBarra).
// Idempotente: si no hay PN PENDIENTE coincidente, no hace nada.
function _anularPNPorGuiaYCodigo(idGuia, codigoBarra) {
  var sh = getSheet('PRODUCTO_NUEVO');
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  var hdrs = data[0];
  var idxIdG = hdrs.indexOf('idGuia');
  var idxCb  = hdrs.indexOf('codigoBarra');
  var idxEst = hdrs.indexOf('estado');
  if (idxIdG < 0 || idxCb < 0 || idxEst < 0) return;

  var cb = String(codigoBarra || '').toUpperCase();
  var ig = String(idGuia || '');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxIdG]) !== ig) continue;
    if (String(data[r][idxCb] || '').toUpperCase() !== cb) continue;
    if (String(data[r][idxEst] || '').toUpperCase() !== 'PENDIENTE') continue;
    sh.getRange(r + 1, idxEst + 1).setValue('ANULADO');
  }
}

// ── Reabrir una guía cerrada (requiere adminPin en el cliente) ──
// REVIERTE el stock que se aplicó al cerrar para que al volver a cerrar
// no se descuente dos veces. Es la operación inversa de cerrarGuia.
function reabrirGuia(params) {
  return _conLock('reabrirGuia', function() { return _reabrirGuiaImpl(params); });
}
function _reabrirGuiaImpl(params) {
  var idGuia = params.idGuia;
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var sheet = getSheet('GUIAS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxId   = hdrs.indexOf('idGuia');
  var idxEst  = hdrs.indexOf('estado');
  var idxTipo = hdrs.indexOf('tipo');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idGuia)) {
      var estadoActual = String(data[i][idxEst] || '').toUpperCase();
      // Idempotencia: si ya está ABIERTA, no revertir stock
      if (estadoActual === 'ABIERTA') {
        return { ok: true, yaAbierta: true };
      }
      var tipoGuia = String(data[i][idxTipo] || '');
      var esIngreso = tipoGuia.toUpperCase().indexOf('INGRESO') === 0;

      // Solo revertir si estaba CERRADA Y no es de envasado.
      // Envasados manejan stock directo (no via cerrarGuia), por lo tanto
      // reabrir NO debe revertir stock — eso descuadraría el inventario.
      if (estadoActual === 'CERRADA' && !_esGuiaEnvasado(tipoGuia)) {
        var detalles = _sheetToObjects(getSheet('GUIA_DETALLE')).filter(function(d) {
          return d.idGuia === idGuia && d.observacion !== 'ANULADO';
        });
        detalles.forEach(function(d) {
          var cant = parseFloat(d.cantidadRecibida) || 0;
          if (cant === 0 || !d.codigoProducto) return;
          // Reverso del cierre: si era INGRESO, restar; si SALIDA, sumar
          var deltaReverso = esIngreso ? -cant : cant;
          _actualizarStock(String(d.codigoProducto), deltaReverso, {
            tipoOperacion: 'REABRIR_REVERSO',
            origen:        idGuia,
            usuario:       String(params.usuario || ''),
            observacion:   'tipo=' + tipoGuia
          });
        });
      }

      sheet.getRange(i + 1, idxEst + 1).setValue('ABIERTA');
      // [WH Fase 2 · PASO 2] PATCH del estado a la sombra (best-effort)
      try { if (typeof _dualWritePatchWH === 'function') _dualWritePatchWH('guias', { id_guia: 'eq.' + idGuia }, { estado: 'ABIERTA' }); } catch(_eDW) {}
      return { ok: true };
    }
  }
  return { ok: false, error: 'Guía no encontrada' };
}

// ── Auto-cerrar guías abiertas de días anteriores ────────────
function autoCloseDayGuias() {
  return _conLock('autoCloseDayGuias', function() { return _autoCloseDayGuiasImpl(); });
}
function _autoCloseDayGuiasImpl() {
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

function cerrarGuia(idGuia, usuario, idSesion, opts) {
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };
  return _conLock('cerrarGuia', function() {
    return _cerrarGuiaImpl(idGuia, usuario, idSesion, opts);
  });
}

function _cerrarGuiaImpl(idGuia, usuario, idSesion, opts) {
  opts = opts || {};
  var skipMosSync = opts.skipMosSync === true;  // omitir sync a MOS (uso en cierres masivos para evitar timeout)

  var guiasSheet    = getSheet('GUIAS');
  var guias         = guiasSheet.getDataRange().getValues();
  var headers       = guias[0];
  var idxIdGuia     = headers.indexOf('idGuia');
  var idxEstado     = headers.indexOf('estado');
  var idxTipo       = headers.indexOf('tipo');
  var idxMontoTotal = headers.indexOf('montoTotal');

  var filaGuia = -1;
  var tipoGuia = '';
  var estadoActual = '';
  for (var i = 1; i < guias.length; i++) {
    if (guias[i][idxIdGuia] === idGuia) {
      filaGuia = i + 1;
      tipoGuia = guias[i][idxTipo];
      estadoActual = String(guias[i][idxEstado] || '').toUpperCase();
      break;
    }
  }
  if (filaGuia < 0) return { ok: false, error: 'Guía no encontrada' };

  // Idempotencia: si la guía ya está CERRADA (o AUTOCERRADA), no reaplicar stock.
  // Permite que el frontend reintente sin duplicar descuentos.
  if (estadoActual === 'CERRADA' || estadoActual === 'AUTOCERRADA') {
    var montoExistente = idxMontoTotal >= 0 ? (parseFloat(guias[filaGuia - 1][idxMontoTotal]) || 0) : 0;
    return { ok: true, data: { idGuia: idGuia, estado: estadoActual, montoTotal: montoExistente, yaCerrada: true } };
  }

  // Obtener detalles
  var detalles = _sheetToObjects(getSheet('GUIA_DETALLE')).filter(function(d){
    return d.idGuia === idGuia;
  });

  // Calcular monto total
  var montoTotal = detalles.reduce(function(acc, d) {
    return acc + (parseFloat(d.cantidadRecibida) || 0) * (parseFloat(d.precioUnitario) || 0);
  }, 0);

  // [v2.13.183] Comparación case-insensitive (consistente con el resto del módulo)
  var esIngreso = String(tipoGuia || '').toUpperCase().indexOf('INGRESO') === 0;
  var esEnvasado = _esGuiaEnvasado(tipoGuia);

  // Actualizar stock por cada detalle.
  // SALIDA_ENVASADO / INGRESO_ENVASADO: stock ya fue aplicado por Envasados.gs
  // directamente con _actualizarStock — saltarse para no duplicar.
  //
  // [v2.13.57] Para sincronizar idLote de vuelta al detalle cuando el cierre
  // crea el lote desde fecha, necesitamos el rowIndex original en GUIA_DETALLE.
  // Releemos la hoja una sola vez y armamos un map idDetalle → rowSheet.
  var detSheetRef = getSheet('GUIA_DETALLE');
  var detSheetVals = detSheetRef.getDataRange().getValues();
  var detSheetHdrs = detSheetVals[0];
  var detIdxId    = detSheetHdrs.indexOf('idDetalle');
  var detIdxLote  = detSheetHdrs.indexOf('idLote');
  var detRowByDet = {};
  for (var dsi = 1; dsi < detSheetVals.length; dsi++) {
    var detId = String(detSheetVals[dsi][detIdxId] || '');
    if (detId) detRowByDet[detId] = dsi + 1;
  }

  if (!esEnvasado) {
    detalles.forEach(function(d) {
      var cantidad = parseFloat(d.cantidadRecibida) || 0;
      if (cantidad === 0) return;
      var delta = esIngreso ? cantidad : -cantidad;
      _actualizarStock(d.codigoProducto, delta, {
        tipoOperacion: 'CIERRE_GUIA',
        origen:        idGuia,
        usuario:       String(usuario || ''),
        observacion:   'tipo=' + tipoGuia
      });
      // [v2.13.57] Si es INGRESO y el detalle tiene fechaVencimiento → garantizar lote.
      // ANTES: solo actualizaba si d.idLote ya existía. Esto dejaba detalles con fecha
      // sin lote (ej: el operador editó fecha inline y actualizarFechaVencimiento falló
      // silenciosa por .catch del frontend, o el detalle se sumó por AUTO_SUMA sin lote).
      // AHORA: usa _sincronizarLoteDesdeDetalle (idempotente por cod+guia+fecha) y
      // persiste el idLote resultante en GUIA_DETALLE si no estaba.
      if (esIngreso && d.fechaVencimiento && String(d.fechaVencimiento).trim() !== '') {
        try {
          var fechaVencDate = d.fechaVencimiento instanceof Date
            ? d.fechaVencimiento
            : new Date(String(d.fechaVencimiento).substring(0, 10) + 'T12:00:00');
          var resLote = _sincronizarLoteDesdeDetalle({
            idLoteActual:   String(d.idLote || ''),
            codigoProducto: d.codigoProducto,
            cantidad:       cantidad,
            fechaVenc:      fechaVencDate,
            idGuia:         idGuia,
            idDetalle:      d.idDetalle,
            usuario:        String(usuario || ''),
            motivo:         'cierre_guia tipo=' + tipoGuia
          });
          // Persistir idLote en GUIA_DETALLE si era vacío o cambió
          if (resLote && resLote.idLote && detIdxLote >= 0) {
            var rowDet = detRowByDet[String(d.idDetalle)];
            if (rowDet && String(d.idLote || '') !== resLote.idLote) {
              detSheetRef.getRange(rowDet, detIdxLote + 1)
                         .setNumberFormat('@')
                         .setValue(resLote.idLote);
            }
          }
        } catch(eL) { Logger.log('[cierre] sync lote fallo ' + d.codigoProducto + ': ' + eL.message); }
      } else if (esIngreso && d.idLote && d.idLote !== '') {
        // Sin fecha pero con idLote → legacy path (mantiene compat con guías
        // antiguas que crearon idLote sin fecha por _actualizarLote viejo)
        _actualizarLote(d.idLote, d.codigoProducto, cantidad, d.fechaVencimiento || '', idGuia);
      }
      // [v2.13.54] Si es SALIDA: consumir lotes FIFO real
      //   - El lote más viejo sale primero
      //   - cantidadActual de cada lote baja proporcionalmente
      //   - Si no hay lotes suficientes, queda como "consumo sin lote"
      //     (no se registra, refleja la realidad: el producto entró antes
      //      de que existiera el sistema de lotes)
      if (!esIngreso) {
        try {
          var resFifo = _consumirLotesFIFO(
            d.codigoProducto, cantidad, idGuia,
            String(usuario || ''),
            'cierre_guia tipo=' + tipoGuia
          );
          if (resFifo.huerfano > 0) {
            Logger.log('[FIFO] consumo huerfano ' + resFifo.huerfano + ' de ' +
                       d.codigoProducto + ' (sin lote disponible) en ' + idGuia);
          }
        } catch(eF) { Logger.log('[FIFO] consumo fallo en ' + d.codigoProducto + ': ' + eF.message); }
      }
    });
  }

  if (idSesion) registrarActividad(idSesion, 'GUIA_CERRADA', 1);

  // Marcar guía como cerrada
  guiasSheet.getRange(filaGuia, idxEstado + 1).setValue('CERRADA');
  guiasSheet.getRange(filaGuia, idxMontoTotal + 1).setValue(montoTotal);
  // [WH Fase 2 · PASO 2] PATCH del estado+monto a la sombra en tiempo real (best-effort)
  try { if (typeof _dualWritePatchWH === 'function') _dualWritePatchWH('guias', { id_guia: 'eq.' + idGuia }, { estado: 'CERRADA', monto_total: montoTotal }); } catch(_eDW) {}
  // [WH Fase 2 · PASO 2] re-sincronizar las líneas (ítems) de la guía a la sombra: al cerrar, los ítems son
  // finales y es cuando se leen para despacho/auditoría → wh.guia_detalle queda fresco junto con la cabecera.
  try { if (typeof _dualWriteDetallesGuiaWH === 'function') _dualWriteDetallesGuiaWH(idGuia); } catch(_eDD) {}

  // Si fue SALIDA_MERMA (cierre semanal manual), marcar mermas asociadas como DESECHADA
  // y notificar a MASTER/ADMINISTRADOR vía push
  if (tipoGuia === 'SALIDA_MERMA') {
    try { _cerrarMermasDeGuia(idGuia, detalles); }
    catch(eM) { Logger.log('cerrar mermas: ' + eM.message); }
  }

  // Si fue INGRESO_PROVEEDOR con idProveedor, sincroniza productos a MOS
  // (silencioso: si falla, log y sigue)
  // OMITIR si skipMosSync (cierres masivos para evitar timeout — el sync hace
  // un UrlFetchApp por cada detalle, lento si la guía tiene muchos productos).
  if (!skipMosSync) {
    try {
      if (tipoGuia === 'INGRESO_PROVEEDOR') {
        var idxProv = headers.indexOf('idProveedor');
        var idProveedor = idxProv >= 0 ? String(guias[filaGuia - 1][idxProv] || '').trim() : '';
        if (idProveedor) _syncProductosProvAMos(idProveedor, detalles);
      }
    } catch(eS) { Logger.log('sync productos proveedor: ' + eS.message); }
  }

  return { ok: true, data: { idGuia: idGuia, estado: 'CERRADA', montoTotal: montoTotal } };
}

// Llama al GAS de MOS para upsert de cada producto en PROVEEDORES_PRODUCTOS
function _syncProductosProvAMos(idProveedor, detalles) {
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) return;
  detalles.forEach(function(d){
    var cb = String(d.codigoProducto || '').trim();
    var precio = parseFloat(d.precioUnitario) || 0;
    if (!cb) return;
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'upsertProductoProveedor',
          idProveedor: idProveedor,
          codigoBarra: cb,
          precioUnitario: precio,
          descripcion: d.descripcion || ''
        }),
        muteHttpExceptions: true
      });
    } catch(e) { /* silencioso */ }
  });
}

function _actualizarLote(idLote, codigoProducto, cantidad, fechaVencimiento, idGuia) {
  var sheet   = getSheet('LOTES_VENCIMIENTO');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idxId   = headers.indexOf('idLote');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxId] !== idLote) continue;
    // Actualizar cantidad final y fechaVencimiento en fila existente
    var idxQI = headers.indexOf('cantidadInicial');
    var idxQA = headers.indexOf('cantidadActual');
    var idxFV = headers.indexOf('fechaVencimiento');
    if (idxQI >= 0) sheet.getRange(i + 1, idxQI + 1).setValue(cantidad);
    if (idxQA >= 0) sheet.getRange(i + 1, idxQA + 1).setValue(cantidad);
    if (idxFV >= 0 && fechaVencimiento) sheet.getRange(i + 1, idxFV + 1).setValue(fechaVencimiento);
    return;
  }
  // No existe → crear con cantidad y fecha confirmadas
  sheet.appendRow([
    idLote, codigoProducto, fechaVencimiento || '', cantidad, cantidad, idGuia, 'ACTIVO', new Date()
  ]);
}

// ── PICKUPS (cierres de caja ME + cualquier fuente externa que escriba en hoja PICKUPS) ──
function getPickups(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  var filtroEstado = params.estado || 'PENDIENTE';
  var rows = _sheetToObjects(sheet);
  var result = rows.filter(function(r) {
    if (filtroEstado === 'TODOS') return true;
    // Si filtroEstado tiene comas, OR de varios estados (ej "PENDIENTE,EN_PROCESO")
    var estados = String(filtroEstado).split(',').map(function(s){return s.trim();});
    return estados.indexOf(r.estado) >= 0;
  }).map(function(r) {
    try { r.items = JSON.parse(r.items || '[]'); } catch(e) { r.items = []; }
    return r;
  });
  // Más recientes primero (el operador atiende los nuevos)
  result.sort(function(a, b) {
    var ta = new Date(a.fechaCreado || 0).getTime();
    var tb = new Date(b.fechaCreado || 0).getTime();
    return tb - ta;
  });
  return { ok: true, data: result };
}

// ── RECIBIR PICKUP DESDE MosExpress al cierre de caja ──────────
// Payload: { idGuiaME, idCaja, idZona, cajero, items: [{skuBase,nombre,solicitado,despachado,codigosOriginales}] }
// Idempotente por idGuiaME — si ya existe un pickup con ese origen, no duplica.
function recibirPickupDeME(params) {
  return _conLock('recibirPickupDeME', function() { return _recibirPickupDeMEImpl(params); });
}
function _recibirPickupDeMEImpl(params) {
  var idGuiaME = String(params.idGuiaME || '').trim();
  var idZona   = String(params.idZona   || '').trim();
  if (!idGuiaME) return { ok: false, error: 'Requiere idGuiaME' };
  if (!params.items || !params.items.length) return { ok: false, error: 'Sin items' };

  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };

  // Idempotencia: si ya hay pickup con esta idGuiaME, no duplicar
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){return String(h);});
  var iId  = hdrs.indexOf('idPickup');
  var iSrc = hdrs.indexOf('fuente');
  var iSt  = hdrs.indexOf('estado');
  var iIt  = hdrs.indexOf('items');
  var iZn  = hdrs.indexOf('idZona');
  var iNt  = hdrs.indexOf('notas');
  var iCr  = hdrs.indexOf('creadoPor');
  var iFC  = hdrs.indexOf('fechaCreado');
  // Columnas opcionales — el código las usa si existen, si no degrada limpio
  var iAt  = hdrs.indexOf('atendidoPor');     // lock multi-operador
  var iUa  = hdrs.indexOf('ultimaActividad'); // detector pickups atascados
  for (var r = 1; r < data.length; r++) {
    var notas = String(data[iNt >= 0 ? r : 0][iNt] || '');
    if (notas.indexOf('idGuiaME=' + idGuiaME) >= 0) {
      return { ok: true, data: { idPickup: data[r][iId], dedup: true } };
    }
  }

  // Sanear items: solo skuBase + solicitado > 0
  var itemsLimpios = params.items
    .map(function(it){
      return {
        skuBase:           String(it.skuBase || '').trim(),
        nombre:            String(it.nombre || it.skuBase || '').trim(),
        solicitado:        parseFloat(it.solicitado) || 0,
        despachado:        parseFloat(it.despachado) || 0,
        codigosOriginales: Array.isArray(it.codigosOriginales) ? it.codigosOriginales : []
      };
    })
    .filter(function(it){ return it.skuBase && it.solicitado > 0; });
  if (!itemsLimpios.length) return { ok: false, error: 'Items inválidos' };

  // Ordenar por nombre para que el operador los lea fácil
  itemsLimpios.sort(function(a, b){ return String(a.nombre).localeCompare(String(b.nombre)); });

  var nowIso = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var idPickup = 'PCK-' + new Date().getTime();
  var fila = new Array(hdrs.length).fill('');
  if (iId  >= 0) fila[iId]  = idPickup;
  if (iSrc >= 0) fila[iSrc] = 'ME_CIERRE_CAJA';
  if (iSt  >= 0) fila[iSt]  = 'PENDIENTE';
  if (iIt  >= 0) fila[iIt]  = JSON.stringify(itemsLimpios);
  if (iZn  >= 0) fila[iZn]  = idZona;
  if (iNt  >= 0) fila[iNt]  = 'idGuiaME=' + idGuiaME + ' · idCaja=' + (params.idCaja || '') + ' · cajero=' + (params.cajero || '');
  if (iCr  >= 0) fila[iCr]  = params.cajero || 'ME_AUTO';
  if (iFC  >= 0) fila[iFC]  = nowIso;
  if (iAt  >= 0) fila[iAt]  = '';     // pickup nuevo: nadie lo atiende
  if (iUa  >= 0) fila[iUa]  = nowIso; // primera actividad = creación
  sheet.appendRow(fila);

  // Avisar a MOS que hay un pickup nuevo (push a operadores almacén)
  try {
    var mosUrl = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
    if (mosUrl) {
      var totalUds = itemsLimpios.reduce(function(s, it){ return s + it.solicitado; }, 0);
      UrlFetchApp.fetch(mosUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          action: 'enviarPushNotif',
          titulo: '🚨 Nuevo pickup · ' + (idZona || 'zona'),
          cuerpo: itemsLimpios.length + ' productos · ' + Math.round(totalUds) + ' uds · cajero ' + (params.cajero || ''),
          soloRolesWH: true
        }),
        muteHttpExceptions: true
      });
    }
  } catch(e) { Logger.log('Push pickup falló: ' + e.message); }

  return { ok: true, data: { idPickup: idPickup, items: itemsLimpios.length } };
}

// Normaliza nombre de operador para comparar locks de pickup. Tolera
// dobles espacios / mayúsculas / trims — necesario porque PERSONAL a veces
// guarda nombres con espacios extra y al comparar ===-estricto el mismo
// operador queda bloqueado de sí mismo desde otro device.
function _normUser_(u) {
  return String(u || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function _sameUser_(a, b) {
  var na = _normUser_(a), nb = _normUser_(b);
  return !!na && !!nb && na === nb;
}

// Actualiza el estado de un pickup. Soporta lock optimista por atendidoPor:
// si params.lockUsuario viene, sólo permite cambios si atendidoPor está vacío
// o coincide con lockUsuario (comparación normalizada). Si tomarLock=true,
// marca atendidoPor=lockUsuario.
function actualizarPickup(params) {
  return _conLock('actualizarPickup', function() { return _actualizarPickupImpl(params); });
}
function _actualizarPickupImpl(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var idxId   = headers.indexOf('idPickup');
  var idxEst  = headers.indexOf('estado');
  var idxAte  = headers.indexOf('fechaAtendido');
  var idxAtp  = headers.indexOf('atendidoPor');
  var idxUa   = headers.indexOf('ultimaActividad');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(params.idPickup)) {
      // Lock check — si me piden tomar el pickup y ya lo tiene OTRO usuario
      // (no yo mismo desde otro device), rechazar. Comparación normalizada
      // para tolerar dobles espacios/mayúsculas entre devices.
      if (params.lockUsuario && idxAtp >= 0) {
        var actual = String(data[i][idxAtp] || '').trim();
        if (actual && !_sameUser_(actual, params.lockUsuario)) {
          return { ok: false, error: 'Pickup atendido por ' + actual, atendidoPor: actual, conflicto: true };
        }
      }
      if (idxEst >= 0 && params.estado) sheet.getRange(i + 1, idxEst + 1).setValue(params.estado);
      if (idxAte >= 0 && params.estado === 'COMPLETADO') {
        sheet.getRange(i + 1, idxAte + 1).setValue(new Date());
      }
      // Tomar lock
      if (params.tomarLock && idxAtp >= 0 && params.lockUsuario) {
        sheet.getRange(i + 1, idxAtp + 1).setValue(String(params.lockUsuario));
      }
      // Liberar lock explícito
      if (params.liberarLock === true && idxAtp >= 0) {
        sheet.getRange(i + 1, idxAtp + 1).setValue('');
      }
      // Heartbeat de actividad
      if (idxUa >= 0) {
        sheet.getRange(i + 1, idxUa + 1).setValue(
          Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")
        );
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pickup no encontrado' };
}

// Liberar lock de un pickup (operador "suelta" para que otro lo tome).
// Vuelve estado a PENDIENTE si estaba EN_PROCESO sin progreso.
function liberarPickup(params) {
  return _conLock('liberarPickup', function() { return _liberarPickupImpl(params); });
}
function _liberarPickupImpl(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var idxId   = headers.indexOf('idPickup');
  var idxEst  = headers.indexOf('estado');
  var idxIt   = headers.indexOf('items');
  var idxAtp  = headers.indexOf('atendidoPor');
  var idxUa   = headers.indexOf('ultimaActividad');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) !== String(params.idPickup)) continue;
    // Si hay progreso, sólo limpia atendidoPor (deja EN_PROCESO para que cualquiera continue)
    var items = []; try { items = JSON.parse(String(data[i][idxIt] || '[]')); } catch(_){}
    var hayProgreso = items.some(function(it){ return (parseFloat(it.despachado) || 0) > 0; });
    if (idxAtp >= 0) sheet.getRange(i + 1, idxAtp + 1).setValue('');
    if (idxEst >= 0 && !hayProgreso) sheet.getRange(i + 1, idxEst + 1).setValue('PENDIENTE');
    if (idxUa  >= 0) sheet.getRange(i + 1, idxUa  + 1).setValue(
      Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")
    );
    return { ok: true, data: { hayProgreso: hayProgreso } };
  }
  return { ok: false, error: 'Pickup no encontrado' };
}

// Devuelve un pickup específico (para que el frontend hidrate localStorage
// contra el backend al refrescar y detecte si ya fue cerrado por otro).
function getPickup(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, error: 'Pickup no encontrado' };
  var rows = _sheetToObjects(sheet);
  var p = rows.find(function(r){ return String(r.idPickup) === String(params.idPickup); });
  if (!p) return { ok: false, error: 'Pickup no encontrado' };
  try { p.items = JSON.parse(p.items || '[]'); } catch(_){ p.items = []; }
  return { ok: true, data: p };
}

// Ajusta solicitado del pickup cuando ME anula una venta de la caja origen.
// params: { idCaja, idGuiaME?, itemsAnulados: [{codigoBarra, cantidad}] }
// Si después del descuento todos los items quedan en 0 → pickup CANCELADO.
function pickupDescontarVenta(params) {
  return _conLock('pickupDescontarVenta', function() { return _pickupDescontarVentaImpl(params); });
}
function _pickupDescontarVentaImpl(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };
  var idCaja    = String(params.idCaja || '').trim();
  var idGuiaME  = String(params.idGuiaME || '').trim();
  var itemsAnul = Array.isArray(params.itemsAnulados) ? params.itemsAnulados : [];
  if (!itemsAnul.length) return { ok: false, error: 'Sin itemsAnulados' };
  if (!idCaja && !idGuiaME) return { ok: false, error: 'Requiere idCaja o idGuiaME' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){return String(h);});
  var idxId   = headers.indexOf('idPickup');
  var idxIt   = headers.indexOf('items');
  var idxEst  = headers.indexOf('estado');
  var idxNt   = headers.indexOf('notas');

  for (var i = 1; i < data.length; i++) {
    var notas = String(data[i][idxNt] || '');
    var match = (idGuiaME && notas.indexOf('idGuiaME=' + idGuiaME) >= 0) ||
                (idCaja   && notas.indexOf('idCaja='   + idCaja)   >= 0);
    if (!match) continue;
    var estado = String(data[i][idxEst] || '');
    if (estado === 'COMPLETADO' || estado === 'CANCELADO' || estado === 'PARCIAL') {
      // Ya cerró — no se puede ajustar
      return { ok: true, data: { ajustado: false, motivo: 'Pickup ya cerrado: ' + estado } };
    }
    var items = []; try { items = JSON.parse(String(data[i][idxIt] || '[]')); } catch(_){}
    if (!items.length) return { ok: true, data: { ajustado: false, motivo: 'Sin items' } };

    var ajustes = 0;
    itemsAnul.forEach(function(an){
      var codU = String(an.codigoBarra || '').toUpperCase();
      var qty  = parseFloat(an.cantidad) || 0;
      if (!codU || qty <= 0) return;
      // Buscar item del pickup que tenga ese código en codigosOriginales
      var it = items.find(function(x){
        return Array.isArray(x.codigosOriginales) &&
               x.codigosOriginales.some(function(c){ return String(c).toUpperCase() === codU; });
      });
      if (!it) return;
      it.solicitado = Math.max(0, (parseFloat(it.solicitado) || 0) - qty);
      ajustes++;
    });

    // Quitar items con solicitado=0
    var itemsFinal = items.filter(function(it){ return (parseFloat(it.solicitado) || 0) > 0; });
    if (!itemsFinal.length) {
      sheet.getRange(i + 1, idxEst + 1).setValue('CANCELADO');
      sheet.getRange(i + 1, idxIt  + 1).setValue('[]');
      return { ok: true, data: { ajustado: true, ajustes: ajustes, cancelado: true } };
    }
    sheet.getRange(i + 1, idxIt + 1).setValue(JSON.stringify(itemsFinal));
    return { ok: true, data: { ajustado: true, ajustes: ajustes, cancelado: false } };
  }
  return { ok: false, error: 'Pickup origen no encontrado' };
}

// Job time-driven: pickups EN_PROCESO sin actividad >2h vuelven a PENDIENTE
// y atendidoPor='' para que otro operador los tome. Push aviso a roles WH.
function _jobReabrirPickupsAtascados() {
  var sheet = getSheet('PICKUPS');
  if (!sheet || sheet.getLastRow() < 2) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){return String(h);});
  var idxId   = headers.indexOf('idPickup');
  var idxEst  = headers.indexOf('estado');
  var idxAtp  = headers.indexOf('atendidoPor');
  var idxUa   = headers.indexOf('ultimaActividad');
  var idxZn   = headers.indexOf('idZona');
  if (idxUa < 0) return; // sin esa col, no podemos detectar

  var ahora = Date.now();
  var THRESHOLD_MS = 2 * 60 * 60 * 1000; // 2 horas
  var reabiertos = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxEst]) !== 'EN_PROCESO') continue;
    var ua = data[i][idxUa];
    var t  = ua ? new Date(ua).getTime() : 0;
    if (!t || (ahora - t) < THRESHOLD_MS) continue;
    sheet.getRange(i + 1, idxEst + 1).setValue('PENDIENTE');
    if (idxAtp >= 0) sheet.getRange(i + 1, idxAtp + 1).setValue('');
    sheet.getRange(i + 1, idxUa + 1).setValue(
      Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")
    );
    reabiertos.push({
      idPickup: data[i][idxId],
      idZona:   idxZn >= 0 ? data[i][idxZn] : ''
    });
  }
  // Avisar a operadores WH si hubo reaperturas
  if (reabiertos.length) {
    try {
      var mosUrl = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
      if (mosUrl) {
        UrlFetchApp.fetch(mosUrl, {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify({
            action: 'enviarPushNotif',
            titulo: '⏰ Pickup' + (reabiertos.length>1?'s':'') + ' abandonado' + (reabiertos.length>1?'s':''),
            cuerpo: reabiertos.length + ' pickup' + (reabiertos.length>1?'s':'') + ' sin movimiento >2h · alguien que retome',
            soloRolesWH: true
          }),
          muteHttpExceptions: true
        });
      }
    } catch(e) { Logger.log('Push reapertura falló: ' + e.message); }
  }
  return { ok: true, data: { reabiertos: reabiertos.length } };
}

// Crea el trigger horario (correr 1 vez desde el editor para activar).
function setupPickupTriggers() {
  // Limpiar triggers viejos del mismo handler
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === '_jobReabrirPickupsAtascados') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('_jobReabrirPickupsAtascados').timeBased().everyHours(1).create();
  return { ok: true, mensaje: 'Trigger horario _jobReabrirPickupsAtascados creado' };
}

// Guardar progreso del despacho (autosave optimista mientras el operador trabaja).
// El frontend manda items con despachado actualizado; aquí solo overwrite del JSON
// y marca estado='EN_PROCESO' + actualiza ultimaActividad (heartbeat).
function guardarProgresoPickup(params) {
  return _conLock('guardarProgresoPickup', function() { return _guardarProgresoPickupImpl(params); });
}
function _guardarProgresoPickupImpl(params) {
  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){return String(h);});
  var idxId   = headers.indexOf('idPickup');
  var idxIt   = headers.indexOf('items');
  var idxEst  = headers.indexOf('estado');
  var idxUa   = headers.indexOf('ultimaActividad');
  var idxAtp  = headers.indexOf('atendidoPor');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(params.idPickup)) {
      // Lock check — si lockUsuario viene, debe coincidir con atendidoPor (si está
      // seteado). Comparación normalizada para tolerar dobles espacios/mayúsculas
      // entre devices del mismo operador.
      if (params.lockUsuario && idxAtp >= 0) {
        var actual = String(data[i][idxAtp] || '').trim();
        if (actual && !_sameUser_(actual, params.lockUsuario)) {
          return { ok: false, error: 'Pickup atendido por ' + actual, atendidoPor: actual, conflicto: true };
        }
        // Si no había lock, tomarlo aquí (autosave implica que estoy trabajando)
        if (!actual) sheet.getRange(i + 1, idxAtp + 1).setValue(String(params.lockUsuario));
      }
      var itemsActualizados = Array.isArray(params.items) ? params.items : null;
      if (itemsActualizados && idxIt >= 0) {
        sheet.getRange(i + 1, idxIt + 1).setValue(JSON.stringify(itemsActualizados));
      }
      // Solo cambiar a EN_PROCESO si está PENDIENTE (no degradar COMPLETADO)
      if (idxEst >= 0 && String(data[i][idxEst]) === 'PENDIENTE') {
        sheet.getRange(i + 1, idxEst + 1).setValue('EN_PROCESO');
      }
      // Heartbeat
      if (idxUa >= 0) {
        sheet.getRange(i + 1, idxUa + 1).setValue(
          Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'")
        );
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pickup no encontrado' };
}

// Cerrar pickup: emite GUIA_SALIDA real con códigos de barra escaneados.
// Items NO despachados se anotan en la observación (no en detalle de stock).
// Marca pickup COMPLETADO o PARCIAL según haya despachado todo o no.
// ═══════════════════════════════════════════════════════════════════════
// REGLA DE ORO WH (escrita en piedra) — manejo de codigos en pickups/guías:
//
// 1. MATCHING (al escanear): se acepta cualquiera de:
//      - skuBase del producto
//      - codigoBarra del canónico (factorConversion=1)
//      - codigoBarra de cualquier EQUIVALENCIA activa apuntando al skuBase
//    Todos cuentan como progreso del MISMO item del pickup (agrupado por sku).
//
// 2. TRAZABILIDAD (al cerrar): el frontend acumula despachadoPorCodigo:
//    { '6959749711163': 4, 'EAN-EQUIV-001': 2 }  ← codigoBarra REALES escaneados
//
// 3. GUIA_SALIDA (registro): cada fila de GUIAS_DETALLE lleva el codigoBarra
//    REAL (canónico o equivalente). NUNCA el skuBase como codigoBarra.
//    Razón: STOCK_ZONAS descuenta por codigoBarra específico, no por skuBase.
//    Un producto con 1 canónico + 2 equivalentes activos puede tener 3 rows
//    de stock distintos — el detalle de la guía debe reflejar de cuál se sacó.
//
// 4. skuBase NO es un codigoBarra. Es un agrupador conceptual. Si un sku
//    aparece en despachoDetalle como "codigoBarra" es un bug aguas arriba.
// ═══════════════════════════════════════════════════════════════════════
function cerrarPickupConDespacho(params) {
  // Lock + idempotencia robusta: evita que doble-click o reintentos paralelos
  // generen múltiples GUIA_SALIDA. Estados terminales (COMPLETADO/PARCIAL/CANCELADO)
  // se rechazan; solo PENDIENTE y EN_PROCESO admiten el cierre.
  return _conLock('cerrarPickupConDespacho', function() {
    return _cerrarPickupConDespachoImpl(params);
  });
}

// Helper: ¿ya existe alguna GUIA con [pickup:idPickup] en su comentario que NO
// esté ANULADA? Retorna la más reciente. Solo busca SALIDA_ZONA. Sirve como
// defensa última contra duplicados si la primera llamada creó la guía pero
// falló al marcar el pickup como terminal.
function _buscarGuiaPorPickupReciente(idPickup) {
  try {
    var sheet = getSheet('GUIAS');
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    var hdrs = data[0].map(function(h){ return String(h); });
    var idxId   = hdrs.indexOf('idGuia');
    var idxTipo = hdrs.indexOf('tipo');
    var idxFec  = hdrs.indexOf('fecha');
    var idxCom  = hdrs.indexOf('comentario');
    var idxEst  = hdrs.indexOf('estado');
    if (idxId < 0 || idxCom < 0 || idxEst < 0) return null;

    var marca = '[pickup:' + idPickup + ']';
    var found = null;
    for (var i = data.length - 1; i >= 1; i--) {
      var tipo = String(data[i][idxTipo] || '');
      if (tipo !== 'SALIDA_ZONA') continue;
      var estado = String(data[i][idxEst] || '').toUpperCase();
      if (estado === 'ANULADA') continue;
      var coment = String(data[i][idxCom] || '');
      if (coment.indexOf(marca) < 0) continue;
      // Si tiene el prefijo [ANULADA-DUPLICADO] tampoco la consideramos
      if (coment.indexOf('[ANULADA-DUPLICADO]') === 0) continue;
      found = {
        idGuia:       String(data[i][idxId] || ''),
        fecha:        data[i][idxFec],
        estadoGuia:   estado,
        // Estado a forzar en el pickup. Si el comentario tiene "sin despachar:"
        // significa que fue parcial; si no, completo. (Heurística leve, en caso
        // de duda preferir COMPLETADO para que no se pueda reabrir).
        estadoPickup: coment.indexOf('sin despachar:') >= 0 ? 'PARCIAL' : 'COMPLETADO'
      };
      break;
    }
    return found;
  } catch(e) {
    Logger.log('_buscarGuiaPorPickupReciente error: ' + e.message);
    return null;
  }
}

function _cerrarPickupConDespachoImpl(params) {
  var idPickup = String(params.idPickup || '').trim();
  if (!idPickup) return { ok: false, error: 'Requiere idPickup' };
  var usuario = params.usuario || '';

  var sheet = getSheet('PICKUPS');
  if (!sheet) return { ok: false, error: 'Hoja PICKUPS no existe' };
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){return String(h);});
  var idxId   = headers.indexOf('idPickup');
  var idxIt   = headers.indexOf('items');
  var idxEst  = headers.indexOf('estado');
  var idxAte  = headers.indexOf('fechaAtendido');
  var idxZn   = headers.indexOf('idZona');
  var idxNt   = headers.indexOf('notas');

  var rowIdx = -1, pickup = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idPickup) {
      rowIdx = i + 1;
      pickup = {
        idZona: String(data[i][idxZn] || ''),
        notas:  String(data[i][idxNt] || ''),
        estado: String(data[i][idxEst] || '')
      };
      break;
    }
  }
  if (rowIdx < 2) return { ok: false, error: 'Pickup no encontrado' };
  // ── IDEMPOTENCIA NIVEL 1: por estado del pickup ──
  var estUpper = String(pickup.estado || '').toUpperCase();
  var TERMINALES = { 'COMPLETADO': true, 'PARCIAL': true, 'CANCELADO': true };
  if (TERMINALES[estUpper]) {
    return { ok: false, error: 'El pickup ya fue cerrado (estado=' + estUpper + ')', yaCerrado: true };
  }

  // ── IDEMPOTENCIA NIVEL 2: por DATO en GUIAS ──
  // Antes de crear cualquier guía, buscar si YA existe una guía con
  // [pickup:idPickup] en su comentario. Si existe, el primer cierre ya
  // creó su guía pero algo falló al marcar el pickup como terminal
  // (timeout PrintNode, throw no atrapado, etc). En ese caso reusamos
  // la guía existente y forzamos el estado del pickup, sin crear duplicado.
  //
  // Esta defensa es independiente del estado del pickup y cubre:
  //   - Doble-click cuando el frontend viejo no tiene el lock
  //   - Falla parcial: guía creada pero pickup quedó en EN_PROCESO
  //   - Reintentos automáticos del cliente tras timeout de respuesta
  var guiaExistente = _buscarGuiaPorPickupReciente(idPickup);
  if (guiaExistente) {
    // Forzar el pickup a estado terminal (idempotente: si ya lo está, no cambia)
    var nowFix = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    if (idxEst >= 0) sheet.getRange(rowIdx, idxEst + 1).setValue(guiaExistente.estadoPickup || 'COMPLETADO');
    if (idxAte >= 0) sheet.getRange(rowIdx, idxAte + 1).setValue(nowFix);
    var idxAtpFx = headers.indexOf('atendidoPor');
    var idxUaFx  = headers.indexOf('ultimaActividad');
    if (idxAtpFx >= 0) sheet.getRange(rowIdx, idxAtpFx + 1).setValue('');
    if (idxUaFx  >= 0) sheet.getRange(rowIdx, idxUaFx  + 1).setValue(nowFix);
    Logger.log('cerrarPickupConDespacho IDEMPOTENT: reusing existing guía ' + guiaExistente.idGuia + ' for pickup ' + idPickup);
    return {
      ok: true,
      data: {
        idGuia:        guiaExistente.idGuia,
        estado:        guiaExistente.estadoPickup || 'COMPLETADO',
        yaCerrado:     true,
        idempotente:   true
      }
    };
  }

  // El frontend manda items con despachado y opcionalmente despachoDetalle:
  //   despachoDetalle: [{codigoBarra, cantidad}, ...] — codigoBarras REALES
  //   escaneados (canónico o equivalente). NUNCA skuBase.
  var items = Array.isArray(params.items) ? params.items : [];
  var despachoDetalle = Array.isArray(params.despachoDetalle) ? params.despachoDetalle : [];

  // Si no nos mandaron despachoDetalle, derivar uno mínimo a partir de los
  // codigosOriginales del item (que solo contiene codigoBarra del canónico
  // + equivalentes). NO se debe usar skuBase como codigoBarra (regla de oro).
  if (!despachoDetalle.length) {
    items.forEach(function(it){
      var qty = parseFloat(it.despachado) || 0;
      if (qty <= 0) return;
      // Solo aceptamos codigosOriginales (canónico o equivalente). Si por algún
      // motivo el item no los tiene, registramos warning y saltamos — mejor
      // perder ese item del despacho que generar GUIA_SALIDA con skuBase
      // como codigoBarra (rompe el descuento de stock por codigoBarra).
      var cod = (it.codigosOriginales && it.codigosOriginales[0]) || '';
      if (!cod) {
        Logger.log('cerrarPickupConDespacho: item ' + it.skuBase + ' sin codigosOriginales — skipped');
        return;
      }
      despachoDetalle.push({ codigoBarra: String(cod), cantidad: qty });
    });
  }

  // Validación: ningún codigoBarra del despachoDetalle debe ser un skuBase.
  // Los skuBase tienen formato típico LEVxxx, IDPROxxxx (sin dígitos EAN puros).
  // Aquí solo loggeamos warning si detectamos posible inconsistencia.
  var skusDelPickup = {};
  items.forEach(function(it){ if (it.skuBase) skusDelPickup[String(it.skuBase)] = true; });
  despachoDetalle.forEach(function(d){
    if (skusDelPickup[String(d.codigoBarra)]) {
      Logger.log('⚠ cerrarPickupConDespacho: codigoBarra=' + d.codigoBarra +
                 ' coincide con un skuBase del pickup — verificar regla canónico/equivalente');
    }
  });

  // Items NO despachados (solicitado > despachado) — para observación
  var noDespachados = items.filter(function(it){
    return (parseFloat(it.solicitado) || 0) > (parseFloat(it.despachado) || 0);
  }).map(function(it){
    var falta = (parseFloat(it.solicitado) || 0) - (parseFloat(it.despachado) || 0);
    return it.nombre + ' (' + it.skuBase + ') · faltó ' + falta;
  });

  var totalDespachado = despachoDetalle.reduce(function(s, d){ return s + (parseFloat(d.cantidad) || 0); }, 0);
  var huboDespacho    = totalDespachado > 0;
  var nuevoEstado     = noDespachados.length === 0 ? 'COMPLETADO' : (huboDespacho ? 'PARCIAL' : 'CANCELADO');

  // Crear GUIA_SALIDA si hubo al menos un item despachado
  var idGuia = null;
  if (huboDespacho) {
    // Observación: SOLO la marca [pickup:PCK-X]. El ticket detecta esa marca
    // y reconstruye las 3 secciones (despachado / extras / no despachado)
    // leyendo la hoja PICKUPS. Antes acá se concatenaba la lista completa de
    // "sin despachar" — redundante con la sección NO DESPACHADO del ticket.
    var nota = '[pickup:' + idPickup + ']';
    // [v2.13.59] Propagar imprimir explícito. Antes: undefined !== false === true
    // → backend imprimía Y frontend también imprimía → 2 copias mínimo.
    // Ahora: solo imprimimos si el caller lo pide explícitamente con imprimir:true.
    var guiaRes = crearDespachoRapido({
      tipo:     'SALIDA_ZONA',
      idZona:   pickup.idZona,
      items:    despachoDetalle,    // {codigoBarra, cantidad}
      usuario:  usuario,
      nota:     nota,
      imprimir: params.imprimir === true
    });
    if (!guiaRes.ok) return { ok: false, error: 'Falló GUIA_SALIDA: ' + guiaRes.error };
    idGuia = guiaRes.data && guiaRes.data.idGuia;
  }

  // Actualizar pickup
  var idxAtp = headers.indexOf('atendidoPor');
  var idxUa  = headers.indexOf('ultimaActividad');
  var nowIsoCierre = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  if (idxIt  >= 0) sheet.getRange(rowIdx, idxIt  + 1).setValue(JSON.stringify(items));
  if (idxEst >= 0) sheet.getRange(rowIdx, idxEst + 1).setValue(nuevoEstado);
  if (idxAte >= 0) sheet.getRange(rowIdx, idxAte + 1).setValue(nowIsoCierre);
  if (idxAtp >= 0) sheet.getRange(rowIdx, idxAtp + 1).setValue(''); // libera lock al cerrar
  if (idxUa  >= 0) sheet.getRange(rowIdx, idxUa  + 1).setValue(nowIsoCierre);

  return { ok: true, data: {
    idGuia:        idGuia,
    estado:        nuevoEstado,
    despachados:   despachoDetalle.length,
    noDespachados: noDespachados.length
  }};
}

function crearDespachoRapido(params) {
  return _conLock('crearDespachoRapido', function() {
    return _crearDespachoRapidoImpl(params);
  });
}

function _crearDespachoRapidoImpl(params) {
  var idZona   = params.idZona   || '';
  var tipo     = params.tipo     || 'SALIDA_ZONA';
  var items    = params.items    || [];
  var usuario  = params.usuario  || '';
  var nota     = params.nota     || '';
  var imprimir = params.imprimir !== false;
  var idempotencyKey = String(params.idempotencyKey || '').trim();

  // idZona solo requerido cuando el tipo es SALIDA_ZONA
  if (tipo === 'SALIDA_ZONA' && !idZona) return { ok: false, error: 'idZona requerido' };
  if (!items.length) return { ok: false, error: 'Carrito vacío' };

  // ── IDEMPOTENCIA ──────────────────────────────────────────
  // Bug histórico: triple-click en "Generar guía" creaba 3 GUIA_SALIDA
  // con timestamps separados por <30s (incidente PCK-1778539364666 del
  // 2026-05-12 + repetición 2026-05-13 7:41-7:42). Aunque el frontend
  // ya tenga lock, defendemos también acá:
  //   - Si llega idempotencyKey del cliente y ya vimos esa key → retornar
  //     la guía existente sin re-crear, sin re-descontar stock, sin re-
  //     imprimir.
  //   - Si no llega key, computamos una "fingerprint" del payload (usuario
  //     + tipo + idZona + items canónicos + minuto) y usamos esa.
  // TTL 120s (suficiente para reintentos por timeout o doble-click).
  var cache = CacheService.getScriptCache();
  var keyEfectiva = idempotencyKey;
  if (!keyEfectiva) {
    // Fingerprint: minuto actual + payload normalizado
    var minuto = Math.floor(Date.now() / 60000);
    var itemsCanon = items.map(function(it) {
      return String(it.codigoBarra || '').trim().toUpperCase() + ':' + (parseFloat(it.cantidad) || 0);
    }).sort().join('|');
    keyEfectiva = 'dr_' + usuario + '_' + tipo + '_' + idZona + '_' + minuto + '_' + itemsCanon;
    // Truncar — CacheService key máx 250 chars
    if (keyEfectiva.length > 240) {
      keyEfectiva = 'dr_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyEfectiva)
        .map(function(b){ return (b < 0 ? b + 256 : b).toString(16); }).join('');
    }
  } else {
    keyEfectiva = 'dr_' + keyEfectiva;
  }
  try {
    var prev = cache.get(keyEfectiva);
    if (prev) {
      Logger.log('crearDespachoRapido idempotente: hit cache, retornando ' + prev);
      return { ok: true, data: { idGuia: prev, errores: [], impresion: { ok: true, dedup: true }, dedup: true } };
    }
  } catch(_) {}

  var comentario = nota || ('Despacho rápido' + (tipo !== 'SALIDA_ZONA' ? ' · ' + tipo : ''));

  // 1. Crear guía con el tipo correcto
  var guiaRes = crearGuia({ tipo: tipo, idZona: idZona || null, usuario: usuario, comentario: comentario });
  if (!guiaRes.ok) return guiaRes;
  var idGuia = guiaRes.data.idGuia;

  // Reservar la key INMEDIATAMENTE después de crear la guía (antes de
  // detalles + cerrar), para que cualquier retry concurrent reciba el
  // idGuia. El lock externo previene la race, pero esto es defensa extra.
  try { cache.put(keyEfectiva, idGuia, 120); } catch(_) {}

  // 2. Registrar TODOS los ítems en UN solo batch (Fix #4 v2.11.2).
  // Antes: forEach con agregarDetalleGuia individual (cada uno re-lee productos
  // + toma lock + appendRow). Para 22 items y sheet grande: 30-60s, el cliente
  // tira timeout y la impresión sale con la hoja a medio escribir → ticket
  // cortado a 2-3 items. Ahora es UNA operación atómica + flush garantizado.
  Logger.log('[crearDespachoRapido] idGuia=' + idGuia + ' · solicita escribir ' + items.length + ' items');
  var batchRes = _agregarDetallesBatchImpl(idGuia, items, usuario);
  var errores  = batchRes.data?.errores || [];
  var escritos = batchRes.data?.agregados || 0;
  Logger.log('[crearDespachoRapido] idGuia=' + idGuia + ' · batch escribió ' + escritos + '/' + items.length + ' · errores=' + errores.length);

  // 3. Cerrar guía (descuenta stock)
  var cerrarRes = cerrarGuia(idGuia, usuario, null);
  if (!cerrarRes.ok) return { ok: false, error: 'Error al cerrar guía: ' + cerrarRes.error };

  // [Fix #1 v2.11.2] Flush ANTES de imprimir. cerrarGuia hace varios setValue
  // (estado, stock, etc) y queremos garantizar que la lectura de imprimirTicketGuia
  // vea el estado final de la hoja.
  try { SpreadsheetApp.flush(); } catch(_){}

  // 4. Imprimir ticket — pasamos `esperado` para que valide y reintente si lee menos
  var impresion = { ok: false, error: 'omitido' };
  if (imprimir) {
    try {
      impresion = imprimirTicketGuia({ idGuia: idGuia, esperadoDetalles: escritos });
    } catch(e) {
      impresion = { ok: false, error: e.message };
    }
  }
  Logger.log('[crearDespachoRapido] idGuia=' + idGuia + ' · impresion.ok=' + (impresion && impresion.ok) + ' · vistos=' + (impresion && impresion.data && impresion.data.detallesImpresos));

  return {
    ok: true,
    data: {
      idGuia:    idGuia,
      errores:   errores,
      items:     escritos,
      esperados: items.length,
      impresion: impresion
    }
  };
}

function _validarTipoGuia(tipo) {
  // [v2.13.51] Agregado INGRESO_DEVOLUCION_ZONA — flujo two-party witness:
  // zona ME envía devolución → operador WH valida → solo lo que confirma
  // como BUEN_ESTADO se materializa en esta guía y suma stock real.
  // La tabla intermedia DEVOLUCIONES_ZONA solo guarda el comparativo,
  // NUNCA toca stock directamente.
  var validos = ['INGRESO_PROVEEDOR','INGRESO_JEFATURA','INGRESO_ENVASADO',
                 'INGRESO_DEVOLUCION_ZONA',
                 'SALIDA_DEVOLUCION','SALIDA_ZONA','SALIDA_JEFATURA',
                 'SALIDA_ENVASADO','SALIDA_MERMA'];
  return validos.indexOf(tipo) >= 0;
}
