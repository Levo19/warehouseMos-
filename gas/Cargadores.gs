// ============================================================
// warehouseMos — Cargadores.gs
// Sistema cargadores INDEPENDIENTE de preingresos.
//
// Modelo:
// - Catálogo: PROVEEDORES_MASTER de MOS con nombre que empieza con "CARGADOR".
// - Log diario: tabla CARGADORES_LOG en WH (un row por cada "+1").
// - Resumen: COUNT(estado='ACTIVO') por idCargador por fecha.
// - Remove: marca el row ACTIVO más reciente del cargador como ELIMINADO.
//
// NO confundir con getCargadoresDelDia (Reporte.gs) que agrupa los
// cargadores embedded en JSON de preingresos. Esa función sigue viva
// para compatibilidad; el sistema nuevo es paralelo.
// ============================================================

// ── Catálogo ────────────────────────────────────────────────
// Lista cargadores del PROVEEDORES_MASTER de MOS cuyo nombre
// empieza con "CARGADOR" (case-insensitive). Estado=1 obligatorio.
function listarCargadoresMaster(params) {
  try {
    var rows = _sheetToObjects(getProveedoresSheet());
    var cargadores = rows
      .filter(function(p) {
        if (!_esActivo(p.estado)) return false;
        var nombre = String(p.nombre || '').trim().toUpperCase();
        return nombre.indexOf('CARGADOR') === 0;
      })
      .map(function(p) {
        return {
          idCargador: String(p.idProveedor || ''),
          nombre:     String(p.nombre || '').replace(/^cargador\s*/i, '').trim()
                       || String(p.nombre || ''),
          nombreCompleto: String(p.nombre || ''),
          telefono:   String(p.telefono || ''),
          foto:       String(p.imagen   || '')
        };
      });
    return { ok: true, data: cargadores };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Agregar cargador al día (+1) ────────────────────────────
function addCargadorDia(params) {
  var idCargador = String(params.idCargador || '').trim();
  var fecha      = String(params.fecha || '').trim() || _hoyStr();
  var nombre     = String(params.nombre || '').trim();
  var addedBy    = String(params.usuario || '').trim();
  var deviceId   = String(params.deviceId || '').trim();
  if (!idCargador) return { ok: false, error: 'idCargador requerido' };

  return _conLock('addCargadorDia', function() {
    var sheet = getSheet('CARGADORES_LOG');
    if (!sheet) return { ok: false, error: 'Tabla CARGADORES_LOG no existe. Ejecutar setupExtenderF0()' };
    var id = _generateId('CLG');
    sheet.appendRow([
      id, fecha, idCargador, nombre, addedBy, deviceId, new Date(), 'ACTIVO'
    ]);
    var conteo = _contarCargadorDia(idCargador, fecha);
    return { ok: true, data: { idLog: id, conteo: conteo, fecha: fecha } };
  });
}

// ── Quitar cargador del día (-1) ────────────────────────────
// Marca el row ACTIVO más reciente del cargador en esa fecha como ELIMINADO.
function removeCargadorDia(params) {
  var idCargador = String(params.idCargador || '').trim();
  var fecha      = String(params.fecha || '').trim() || _hoyStr();
  if (!idCargador) return { ok: false, error: 'idCargador requerido' };

  return _conLock('removeCargadorDia', function() {
    var sheet = getSheet('CARGADORES_LOG');
    if (!sheet) return { ok: false, error: 'Tabla CARGADORES_LOG no existe' };
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0];
    var idxFecha = hdrs.indexOf('fecha');
    var idxIdC   = hdrs.indexOf('idCargador');
    var idxEst   = hdrs.indexOf('estado');
    var idxTs    = hdrs.indexOf('ts');

    var bestRow  = -1;
    var bestTs   = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxIdC]) !== idCargador) continue;
      if (String(data[i][idxFecha]).substring(0, 10) !== fecha) continue;
      if (String(data[i][idxEst]).toUpperCase() !== 'ACTIVO') continue;
      var ts = data[i][idxTs] instanceof Date ? data[i][idxTs].getTime() : 0;
      if (ts > bestTs) { bestTs = ts; bestRow = i + 1; }
    }
    if (bestRow < 0) return { ok: false, error: 'sin entradas ACTIVO para quitar' };
    sheet.getRange(bestRow, idxEst + 1).setValue('ELIMINADO');
    var conteo = _contarCargadorDia(idCargador, fecha);
    return { ok: true, data: { conteo: conteo, fecha: fecha } };
  });
}

// ── Resumen del día ─────────────────────────────────────────
// Retorna { fecha, total, cargadores: [{idCargador, nombre, count}] }
function getResumenCargadoresDia(params) {
  var fecha = String(params && params.fecha || '').trim() || _hoyStr();
  try {
    var sheet = getSheet('CARGADORES_LOG');
    if (!sheet) return { ok: true, data: { fecha: fecha, total: 0, cargadores: [] } };
    var rows = _sheetToObjects(sheet).filter(function(r) {
      return String(r.fecha).substring(0, 10) === fecha &&
             String(r.estado).toUpperCase() === 'ACTIVO';
    });
    var byId = {};
    rows.forEach(function(r) {
      var k = String(r.idCargador);
      if (!byId[k]) byId[k] = { idCargador: k, nombre: String(r.nombre || ''), count: 0 };
      byId[k].count++;
    });
    var arr = Object.keys(byId).map(function(k){ return byId[k]; });
    arr.sort(function(a, b){ return b.count - a.count; });
    var total = arr.reduce(function(s, c){ return s + c.count; }, 0);
    return { ok: true, data: { fecha: fecha, total: total, cargadores: arr } };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Helpers internos ────────────────────────────────────────
function _hoyStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _contarCargadorDia(idCargador, fecha) {
  var sheet = getSheet('CARGADORES_LOG');
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0];
  var idxFecha = hdrs.indexOf('fecha');
  var idxIdC   = hdrs.indexOf('idCargador');
  var idxEst   = hdrs.indexOf('estado');
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdC]) !== idCargador) continue;
    if (String(data[i][idxFecha]).substring(0, 10) !== fecha) continue;
    if (String(data[i][idxEst]).toUpperCase() === 'ACTIVO') count++;
  }
  return count;
}
