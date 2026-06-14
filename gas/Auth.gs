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

// [F2] ¿validar la clave admin DIRECTO contra Supabase (mos.verificar_clave_admin) en vez de HTTP-MOS?
// Flag WH_AUTH_DIRECTO='1' (default off → HTTP-MOS). GAS usa service_role → pasa el gate wh._claim_ok().
// La RPC agrega el chequeo de NIVEL por acción (admin vs master-only); el HTTP-MOS no lo tiene (es la mejora del escalón).
function _authClaveDirecta() {
  try { return PropertiesService.getScriptProperties().getProperty('WH_AUTH_DIRECTO') === '1'; }
  catch (e) { return false; }
}
function _validarClaveAdminDirecto(clave, accion, refDocumento) {
  // Llama la RPC en schema mos vía service_role. Devuelve null si NO se debe usar (→ caller cae a HTTP-MOS).
  try {
    var r = _sbRpc('mos', 'verificar_clave_admin', {
      p_clave: clave, p_accion: accion || 'verificar', p_ref: refDocumento || '', p_app: 'warehouseMos'
    });
    if (!r || !r.ok || !r.data) return null;            // fallo de transporte → fallback HTTP-MOS
    var d = r.data;                                      // {ok, autorizado, validado_por, id_personal, nombre, rol, error}
    if (d.ok === false) return null;                     // error interno de la RPC → fallback
    if (d.autorizado !== true) return { ok: false, error: d.error || 'clave incorrecta' };
    // mapear al shape que espera WH (igual que el HTTP-MOS): { ok, data: { validadoPor, idPersonal, nombre, rol } }
    return { ok: true, data: { validadoPor: d.validado_por, idPersonal: d.id_personal, nombre: d.nombre, rol: d.rol, nivel: d.nivel } };
  } catch (e) { return null; }                           // cualquier excepción → fallback HTTP-MOS
}

// Helper: delega la validación de la clave mixta al MOS.
// MOS conoce la globalPin actual y los PINs personales de los admin/master,
// además registra la auditoría de cada validación.
function _validarClaveAdminViaMOS(clave, accion, refDocumento) {
  // [F2] camino directo a Supabase (inerte por flag; fallback TOTAL a HTTP-MOS ante null/excepción)
  if (_authClaveDirecta()) {
    var directo = _validarClaveAdminDirecto(clave, accion, refDocumento);
    if (directo) return directo;   // autorizado o rechazo explícito (clave incorrecta / nivel insuficiente)
    // null → cae a HTTP-MOS abajo
  }
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
