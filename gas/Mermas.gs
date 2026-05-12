// ============================================================
// warehouseMos — Mermas.gs (V2)
//
// Modelo nuevo "Cesta + Separación":
// 1. agregarAMermas → row en MERMAS estado=EN_PROCESO, foto opcional
// 2. solucionarMerma → suma a cantidadReparada y/o cantidadDesechada
//    (permite múltiples soluciones parciales acumulativas)
// 3. procesarEliminacionMermas → toma TODAS las mermas con desechado>0
//    sin idGuiaSalida, genera UNA guía SALIDA_MERMA agrupada, descuenta
//    stock y marca mermas como ELIMINADAS
// 4. getMermasCesta → retorna agrupado por sección (pendientes/descartado/solucionado)
// 5. contadorMermasPendientes → para badge en topbar
//
// COEXISTE con registrarMerma + resolverMerma (Productos.gs) sin reemplazarlos.
// La diferencia clave: V2 permite solución parcial libre y procesado
// manual de descartados, NO genera guía semanal automática.
// ============================================================

// ── Agregar a mermas (foto opcional, zona responsable, motivo libre) ──
function agregarAMermas(params) {
  var cant = parseFloat(params.cantidadOriginal) || 0;
  if (cant <= 0) return { ok: false, error: 'cantidad inválida' };
  if (!params.codigoProducto) return { ok: false, error: 'codigoProducto requerido' };

  return _conLock('agregarAMermas', function() {
    var sheet = getSheet('MERMAS');
    _ensureColumnasMerma(sheet);
    var id = _generateId('M');

    var fotoUrl = String(params.foto || '');
    var fotoBase64 = String(params.fotoBase64 || '').trim();
    if (fotoBase64) {
      try { fotoUrl = _subirFotoMerma(id, fotoBase64, params.mimeType || 'image/jpeg'); }
      catch(e) { Logger.log('agregarAMermas foto: ' + e.message); }
    }

    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row  = new Array(hdrs.length).fill('');
    function _set(col, val) { var i = hdrs.indexOf(col); if (i >= 0) row[i] = val; }
    _set('idMerma',           id);
    _set('fechaIngreso',      new Date());
    _set('origen',            params.zonaResponsable || params.origen || 'ALMACEN');
    _set('responsable',       params.zonaResponsable || params.responsable || '');
    _set('codigoProducto',    String(params.codigoProducto));
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

    // Forzar codigoProducto como texto (preservar ceros)
    var idxCb = hdrs.indexOf('codigoProducto');
    if (idxCb >= 0) {
      sheet.getRange(sheet.getLastRow(), idxCb + 1).setNumberFormat('@').setValue(String(params.codigoProducto));
    }

    if (params.idSesion) {
      try { registrarActividad(params.idSesion, 'MERMA_REGISTRADA', 1); } catch(_){}
    }

    return { ok: true, data: { idMerma: id, fotoUrl: fotoUrl } };
  });
}

// ── Solucionar parcial o total — suma a recuperado/descartado ──
function solucionarMerma(params) {
  var idMerma     = String(params.idMerma || '');
  var deltaRecup  = parseFloat(params.deltaRecuperado) || 0;
  var deltaDesc   = parseFloat(params.deltaDescartado) || 0;
  var obs         = String(params.observacion || '');
  var usuario     = String(params.usuario || '');
  if (!idMerma) return { ok: false, error: 'idMerma requerido' };
  if (deltaRecup < 0 || deltaDesc < 0) return { ok: false, error: 'deltas deben ser positivos' };
  if (deltaRecup === 0 && deltaDesc === 0) return { ok: false, error: 'nada que asignar' };

  return _conLock('solucionarMerma', function() {
    var sheet = getSheet('MERMAS');
    _ensureColumnasMerma(sheet);
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0];
    var idxId      = hdrs.indexOf('idMerma');
    var idxEst     = hdrs.indexOf('estado');
    var idxOrig    = hdrs.indexOf('cantidadOriginal');
    var idxPend    = hdrs.indexOf('cantidadPendiente');
    var idxRep     = hdrs.indexOf('cantidadReparada');
    var idxDes     = hdrs.indexOf('cantidadDesechada');
    var idxFechaR  = hdrs.indexOf('fechaResolucion');
    var idxObsR    = hdrs.indexOf('observacionResolucion');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxId]) !== idMerma) continue;
      var estadoActual = String(data[i][idxEst] || '').toUpperCase();
      if (estadoActual === 'ELIMINADO' || estadoActual === 'DESECHADA') {
        return { ok: false, error: 'merma ya procesada (' + estadoActual + ')' };
      }
      var orig    = parseFloat(data[i][idxOrig]) || 0;
      var repAct  = parseFloat(data[i][idxRep])  || 0;
      var desAct  = parseFloat(data[i][idxDes])  || 0;
      var newRep  = repAct + deltaRecup;
      var newDes  = desAct + deltaDesc;
      var newPend = orig - newRep - newDes;
      if (newPend < -0.001) {
        return { ok: false, error: 'suma excede original ' + orig };
      }
      newPend = Math.max(0, newPend);

      sheet.getRange(i + 1, idxRep  + 1).setValue(newRep);
      sheet.getRange(i + 1, idxDes  + 1).setValue(newDes);
      sheet.getRange(i + 1, idxPend + 1).setValue(newPend);

      var nuevoEstado;
      if (newPend > 0) nuevoEstado = 'EN_PROCESO';
      else if (newRep === orig) nuevoEstado = 'RESUELTA';
      else if (newDes === orig) nuevoEstado = 'DESCARTADO_TOTAL';
      else nuevoEstado = 'SOLUCIONADA_PARCIAL';
      sheet.getRange(i + 1, idxEst + 1).setValue(nuevoEstado);

      if (newPend === 0 && idxFechaR >= 0) {
        sheet.getRange(i + 1, idxFechaR + 1).setValue(new Date());
      }
      if (obs && idxObsR >= 0) {
        sheet.getRange(i + 1, idxObsR + 1).setValue(obs);
      }

      return { ok: true, data: {
        idMerma: idMerma,
        recuperado: newRep,
        descartado: newDes,
        pendiente:  newPend,
        estado:     nuevoEstado
      }};
    }
    return { ok: false, error: 'merma no encontrada' };
  });
}

// ── Procesar eliminación: genera UNA guía SALIDA_MERMA con todo lo descartado ──
function procesarEliminacionMermas(params) {
  var check = _requireAdmin(params);
  if (!check.ok) return check;
  var usuario = String(params.usuario || '');

  return _conLock('procesarEliminacionMermas', function() {
    var sheet = getSheet('MERMAS');
    _ensureColumnasMerma(sheet);
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0];
    var idxId      = hdrs.indexOf('idMerma');
    var idxEst     = hdrs.indexOf('estado');
    var idxDes     = hdrs.indexOf('cantidadDesechada');
    var idxCb      = hdrs.indexOf('codigoProducto');
    var idxMotivo  = hdrs.indexOf('motivo');
    var idxIdGuia  = hdrs.indexOf('idGuiaSalida');

    var paraProcesar = [];
    for (var i = 1; i < data.length; i++) {
      var des = parseFloat(data[i][idxDes]) || 0;
      var idg = String(data[i][idxIdGuia] || '');
      var est = String(data[i][idxEst] || '').toUpperCase();
      if (des > 0 && !idg && est !== 'ELIMINADO') {
        paraProcesar.push({
          row:     i + 1,
          idMerma: String(data[i][idxId]),
          codigo:  String(data[i][idxCb] || ''),
          desc:    des,
          motivo:  String(data[i][idxMotivo] || '')
        });
      }
    }
    if (!paraProcesar.length) return { ok: false, error: 'no hay descartados para procesar' };

    var resGuia = crearGuia({
      tipo:       'SALIDA_MERMA',
      usuario:    usuario || 'admin',
      comentario: 'Procesamiento de mermas descartadas · ' + paraProcesar.length + ' items'
    });
    if (!resGuia.ok) return resGuia;
    var idGuia = resGuia.data.idGuia;

    var fallidos = [];
    paraProcesar.forEach(function(m) {
      if (!m.codigo) { fallidos.push(m.idMerma); return; }
      var resDet = agregarDetalleGuia({
        idGuia:           idGuia,
        codigoProducto:   m.codigo,
        cantidadEsperada: m.desc,
        cantidadRecibida: m.desc,
        observacion:      'Merma ' + m.idMerma + (m.motivo ? ' · ' + m.motivo : '')
      });
      if (!resDet.ok) { fallidos.push(m.idMerma); return; }
      sheet.getRange(m.row, idxIdGuia + 1).setValue(idGuia);
      sheet.getRange(m.row, idxEst    + 1).setValue('ELIMINADO');
    });

    return { ok: true, data: {
      idGuiaSalida: idGuia,
      procesados:   paraProcesar.length - fallidos.length,
      fallidos:     fallidos
    }};
  });
}

// ── Cesta agrupada por sección (frontend) ──
function getMermasCesta(params) {
  var sheet = getSheet('MERMAS');
  _ensureColumnasMerma(sheet);
  var rows = _sheetToObjects(sheet);
  var pendientes  = [];
  var descartado  = [];
  var solucionado = [];
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 7);

  rows.forEach(function(r) {
    var est = String(r.estado || '').toUpperCase();
    var pend = parseFloat(r.cantidadPendiente) || 0;
    var des  = parseFloat(r.cantidadDesechada) || 0;
    var rep  = parseFloat(r.cantidadReparada)  || 0;
    var idg  = String(r.idGuiaSalida || '');

    if (est === 'EN_PROCESO' && pend > 0) {
      pendientes.push(r);
    } else if (des > 0 && !idg && est !== 'ELIMINADO') {
      descartado.push(r);
    } else {
      var f = r.fechaResolucion ? new Date(r.fechaResolucion)
            : r.fechaIngreso    ? new Date(r.fechaIngreso) : null;
      if (rep > 0 && f && f >= cutoff) solucionado.push(r);
    }
  });

  var ordSc = function(a, b){
    var fa = new Date(a.fechaResolucion || a.fechaIngreso || 0);
    var fb = new Date(b.fechaResolucion || b.fechaIngreso || 0);
    return fb - fa;
  };
  pendientes.sort(ordSc);
  descartado.sort(ordSc);
  solucionado.sort(ordSc);

  return { ok: true, data: {
    pendientes:  pendientes,
    descartado:  descartado,
    solucionado: solucionado,
    totalPendientes: pendientes.length,
    totalDescartado: descartado.length
  }};
}

// ── Contador para badge topbar ──
function contadorMermasPendientes(params) {
  try {
    var rows = _sheetToObjects(getSheet('MERMAS'));
    var n = 0;
    rows.forEach(function(r) {
      var est  = String(r.estado || '').toUpperCase();
      var pend = parseFloat(r.cantidadPendiente) || 0;
      var des  = parseFloat(r.cantidadDesechada) || 0;
      var idg  = String(r.idGuiaSalida || '');
      if (est === 'EN_PROCESO' && pend > 0) n++;
      else if (des > 0 && !idg && est !== 'ELIMINADO') n++;
    });
    return { ok: true, data: { count: n } };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
