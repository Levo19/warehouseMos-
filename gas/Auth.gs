// ============================================================
// warehouseMos — Auth.gs
// Verificación de clave admin para acciones sensibles
// (reabrir guía cerrada, procesar eliminación de mermas, etc.)
//
// La clave admin vive en ESTACIONES.ALMACEN.adminPin de MOS.
// Es la misma clave usada por admin+master en MOS para acciones
// restringidas. 8 dígitos.
// ============================================================

function verificarClaveAdmin(params) {
  var clave = String(params.clave || '').trim();
  if (!clave) return { ok: false, error: 'clave requerida' };

  var adminPin = _getAdminPinAlmacen();
  if (!adminPin) return { ok: false, error: 'adminPin no configurado en MOS' };

  if (clave !== adminPin) {
    return { ok: false, error: 'clave incorrecta' };
  }
  return { ok: true, data: { verificado: true, ts: new Date().toISOString() } };
}

// Internal: recupera adminPin de ESTACIONES.ALMACEN en MOS
function _getAdminPinAlmacen() {
  try {
    var estSheet = _getMosSS().getSheetByName('ESTACIONES');
    if (!estSheet) return null;
    var estaciones = _sheetToObjects(estSheet);
    var almacen = estaciones.find(function(e) {
      var key = String(e.idEstacion || e.nombre || '').toUpperCase();
      return key === 'ALMACEN';
    });
    return almacen && almacen.adminPin ? String(almacen.adminPin) : null;
  } catch(e) {
    Logger.log('_getAdminPinAlmacen error: ' + e.message);
    return null;
  }
}

// Helper para que otros endpoints exijan clave admin sin duplicar lógica.
// Uso: var check = _requireAdmin(params); if (!check.ok) return check;
function _requireAdmin(params) {
  var clave = String(params.claveAdmin || params.clave || '').trim();
  if (!clave) return { ok: false, error: 'clave admin requerida' };
  var adminPin = _getAdminPinAlmacen();
  if (!adminPin) return { ok: false, error: 'adminPin no configurado' };
  if (clave !== adminPin) return { ok: false, error: 'clave incorrecta' };
  return { ok: true };
}
