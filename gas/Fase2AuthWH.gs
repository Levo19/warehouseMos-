// ============================================================
// warehouseMos — Fase 2 Auth (PASO 5 · B1)
// Mint de JWT Supabase HS256 con claim app='warehouseMos'. Espejo FIEL de Fase2Auth.gs de ME.
// El secreto (SUPABASE_JWT_SECRET) no sale de GAS; el navegador recibe solo el token corto (exp 5min).
// Valida el deviceId contra la hoja DISPOSITIVOS VIVA de MOS (autoritativa, no la sombra). Fail-closed.
// NACE SIN USO: nada lo invoca hasta que el frontend lo pida y las RPCs tengan RLS (B2). Cero riesgo.
// ============================================================

function _b64urlWH_(bytes){ return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function _b64urlStrWH_(str){ return _b64urlWH_(Utilities.newBlob(str).getBytes()); }

function mintSupabaseTokenWH(deviceId){
  var idd = String(deviceId || '').trim();
  if(!idd) return { ok:false, error:'deviceId requerido' };
  var secret = PropertiesService.getScriptProperties().getProperty('SUPABASE_JWT_SECRET');
  if(!secret) return { ok:false, error:'falta SUPABASE_JWT_SECRET en Script Properties' };

  // Validar contra DISPOSITIVOS VIVA de MOS (autoritativa). Fail-closed: no registrado/ACTIVO/warehouseMos → no token.
  var mosSS, dispSheet;
  try { mosSS = _getMosSS(); dispSheet = mosSS.getSheetByName('DISPOSITIVOS'); }
  catch(e){ return { ok:false, error:'no se pudo abrir DISPOSITIVOS de MOS: '+(e&&e.message) }; }
  if(!dispSheet) return { ok:false, error:'DISPOSITIVOS no disponible' };

  var rows = _sheetToObjects(dispSheet), devOk = false;
  for(var di=0; di<rows.length; di++){
    var dd = rows[di];
    var idMatch  = (String(dd.ID_Dispositivo) === idd || String(dd.idDispositivo) === idd);
    var appMatch = (!dd.App || String(dd.App) === 'warehouseMos');
    var actMatch = (dd.Estado === 'ACTIVO' || dd.estado === 'ACTIVO' || dd.activo === 1 || dd.activo === '1');
    if(idMatch && appMatch && actMatch){ devOk = true; break; }
  }
  if(!devOk) return { ok:false, error:'dispositivo no registrado/activo para warehouseMos' };

  var now = Math.floor(Date.now()/1000);
  var header  = { alg:'HS256', typ:'JWT' };
  var payload = {
    iss:'supabase', role:'authenticated', aud:'authenticated', sub:idd,
    app:'warehouseMos',
    iat: now, exp: now + 300   // 5 min (corto; re-mint en heartbeat)
  };
  var signingInput = _b64urlStrWH_(JSON.stringify(header)) + '.' + _b64urlStrWH_(JSON.stringify(payload));
  var sig = Utilities.computeHmacSha256Signature(signingInput, secret);
  var token = signingInput + '.' + _b64urlWH_(sig);
  return { ok:true, token:token, app:'warehouseMos', exp:payload.exp };
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
