// ============================================================
// warehouseMos — Auth.gs
// Verificación de clave admin/master mixta de 8 dígitos:
//   - 4 dígitos globales (clave global compartida por todo admin/master)
//   - 4 dígitos del PIN personal del admin/master que autoriza
// La validación se hace contra MOS (única fuente de verdad).
// Misma clave usada para: anular ventas/cambiar moneda (ME),
// autorizar dispositivo in-situ (WH), reabrir guía, anular envasado.
// ============================================================

function verificarClaveAdmin(params) {
  var clave = String(params.clave || params.claveAdmin || '').trim();
  if (!clave) return { ok: false, error: 'clave requerida' };
  return _validarClaveAdminViaMOS(clave, params.accion || 'verificar', params.refDocumento || '');
}

// Helper: delega la validación de la clave mixta al MOS.
// MOS conoce la globalPin actual y los PINs personales de los admin/master,
// además registra la auditoría de cada validación.
function _validarClaveAdminViaMOS(clave, accion, refDocumento) {
  var mosUrl = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!mosUrl) {
    return { ok: false, error: 'MOS_WEB_APP_URL no configurada en Script Properties de WH' };
  }
  try {
    var resp = UrlFetchApp.fetch(mosUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        action:        'verificarClaveAdmin',
        clave:         clave,
        accion:        accion || 'verificar',
        refDocumento:  refDocumento || '',
        origen:        'warehouseMos'
      }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code < 200 || code >= 300) {
      return { ok: false, error: 'MOS respondió HTTP ' + code };
    }
    var j;
    try { j = JSON.parse(body); }
    catch(e) { return { ok: false, error: 'MOS respondió payload no-JSON' }; }
    if (!j || j.ok === false) {
      return { ok: false, error: (j && j.error) || 'clave incorrecta' };
    }
    // MOS retorna { ok: true, data: { validadoPor, idPersonal, ... } }
    return { ok: true, data: j.data || {} };
  } catch(e) {
    return { ok: false, error: 'No se pudo conectar con MOS: ' + e.message };
  }
}

// Helper para que otros endpoints exijan clave admin sin duplicar lógica.
// Uso: var check = _requireAdmin(params); if (!check.ok) return check;
function _requireAdmin(params) {
  var clave = String(params.claveAdmin || params.clave || '').trim();
  if (!clave) return { ok: false, error: 'clave admin requerida' };
  return _validarClaveAdminViaMOS(clave, params.accion || params.action || 'admin', params.refDocumento || '');
}
