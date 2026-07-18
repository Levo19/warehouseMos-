// ============================================================
// warehouseMos — Fase 2 Auth (PASO 5 · B1)  ·  [CERO-GAS 2026-07-18] RETIRADO
// Este mint GAS validaba el deviceId contra la hoja DISPOSITIVOS de MOS, que quedó ORFANADA (frozen) al matar
// el reverse-sync. Un device REVOCADO pero aún ACTIVO en la hoja congelada podría obtener un JWT → landmine de
// auth. El mint REAL es la Edge `mint-wh` (valida mos.dispositivos = la SOMBRA, fail-closed, service-role); el
// frontend WH es Edge-only cero-fallback (js/api.js). NACIÓ SIN USO y nunca se cableó → se neutraliza a
// fail-closed en vez de confiar en la hoja congelada. NO borrar la función (la referencia probarMintTokenWH +
// el router case) — solo cortar el path de confianza a la hoja.
// ============================================================

function _b64urlWH_(bytes){ return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function _b64urlStrWH_(str){ return _b64urlWH_(Utilities.newBlob(str).getBytes()); }

function mintSupabaseTokenWH(deviceId){
  // [CERO-GAS] Retirado: el mint va por la Edge `mint-wh` (valida la SOMBRA mos.dispositivos, no la hoja
  // congelada). Fail-closed duro para no acuñar tokens desde datos stale de la hoja orfanada.
  return { ok:false, error:'mintTokenWH GAS retirado — usar Edge mint-wh (valida la sombra, fail-closed)' };
}

// Diagnóstico (editor): confirma secret presente + mintea para el 1er dispositivo warehouseMos ACTIVO.
function probarMintTokenWH(){
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret){ Logger.log('❌ FALTA SUPABASE_JWT_SECRET en Propiedades del script'); return; }
  Logger.log('✅ SUPABASE_JWT_SECRET presente ('+secret.length+' chars)');
  var sh = _getMosSS().getSheetByName('DISPOSITIVOS');
  var rows = _sheetToObjects(sh).filter(function(d){
    return (!d.App || String(d.App)==='warehouseMos') && (d.Estado==='ACTIVO' || d.estado==='ACTIVO');
  });
  if(!rows.length){ Logger.log('sin dispositivos warehouseMos ACTIVOS en DISPOSITIVOS'); return; }
  var dev = String(rows[0].ID_Dispositivo || rows[0].idDispositivo);
  var out = mintSupabaseTokenWH(dev);
  Logger.log('mint para '+dev+' → '+JSON.stringify({ok:out.ok, app:out.app, error:out.error, exp:out.exp}));
}
