/**
 * ============================================================
 * WH · ESCRITURA DIRECTA DE STOCK A SUPABASE (Fase 1)
 * ============================================================
 * Objetivo del dueño: que la ESCRITURA de stock del almacén sea fuente de
 * verdad en Supabase (RPCs wh.*) en vez de la Hoja, para que NO se cruce con
 * el sync Hoja→Supabase (lección "MOS cutover ESCRITURA requiere apagar sync"
 * + "WH escritura directa PASO 4").
 *
 * ⚠️ NACE INERTE. Todo está gateado por el flag WH_ESCRITURA_DIRECTA:
 *    OFF (default) → comportamiento IDÉNTICO a hoy (Hoja + dual-write best-effort).
 *    ON            → la RPC de Supabase es la FUENTE DE VERDAD del stock; si la
 *                    RPC falla, se hace FALLBACK a la lógica vieja (red de
 *                    seguridad: nunca se pierde una escritura).
 *
 * Doble gate (defensa en profundidad):
 *   1. Script Property WH_ESCRITURA_DIRECTA (kill-switch del lado GAS, este archivo).
 *   2. mos.config.WH_*_DIRECTO (kill-switch server-side dentro de cada RPC; cada
 *      RPC devuelve {ok:false,error:'..._OFF'} si su flag no está en '1').
 *   → Para que la escritura directa REALMENTE corra, AMBOS deben estar prendidos.
 *     Eso es a propósito: permite desplegar este GAS inerte sin tocar nada en
 *     producción, y activar gradualmente (primero el server-side por RPC, luego
 *     el master de GAS) bajo control del dueño.
 *
 * IMPORTANTE — este archivo NO apaga el sync ni cambia ninguna lectura.
 * El plan de apagado del sync está documentado al pie (NO ejecutado).
 *
 * Patrón de cada wrapper:
 *   - Si gate OFF  → retorna {handled:false}. El caller corre su lógica de hoy.
 *   - Si gate ON   → llama la RPC vía _sbRpc. Si la RPC dice ok:true →
 *                    {handled:true, data:...}. Si la RPC falla (red/HTTP/_OFF
 *                    server-side/excepción) → {handled:false, fallback:true} para
 *                    que el caller corra la lógica vieja como red de seguridad.
 */

// ── Gate maestro del lado GAS ───────────────────────────────
// Script Property WH_ESCRITURA_DIRECTA: '1'/'true'/'on' = ON. Cualquier otra cosa = OFF.
function _whEscrituraDirectaON() {
  try {
    var v = String(PropertiesService.getScriptProperties().getProperty('WH_ESCRITURA_DIRECTA') || '')
              .trim().toLowerCase();
    return v === '1' || v === 'true' || v === 'on' || v === 'si' || v === 'sí';
  } catch (e) { return false; }
}

// ── Helper: llamar una RPC wh.* y normalizar la respuesta ───
// _sbRpc (Supabase.gs) ya maneja Content-Profile=wh y devuelve {ok,code,data}.
// La RPC devuelve un jsonb {ok:boolean, ...}. Consideramos ÉXITO sólo si:
//   - la llamada HTTP fue ok (r.ok), Y
//   - el cuerpo jsonb trae ok === true (la RPC no rechazó por flag/params).
// Cualquier otro caso → fallback a la lógica vieja.
function _whRpcStock(fn, args) {
  try {
    var r = _sbRpc('wh', fn, args || {});
    if (!r || !r.ok) {
      Logger.log('[whRpcStock ' + fn + '] HTTP no-ok: ' + (r && r.code) + ' ' + (r && r.error || ''));
      return { ok: false, fallback: true, error: 'HTTP', detalle: (r && r.error) || '' };
    }
    var body = r.data;
    // PostgREST devuelve el jsonb directamente (no envuelto). Puede venir como objeto.
    if (body && typeof body === 'object' && body.ok === true) {
      return { ok: true, data: body };
    }
    Logger.log('[whRpcStock ' + fn + '] RPC ok:false → ' + JSON.stringify(body));
    return { ok: false, fallback: true, error: (body && body.error) || 'RPC_NO_OK', detalle: body };
  } catch (e) {
    Logger.log('[whRpcStock ' + fn + '] excepción: ' + (e && e.message));
    return { ok: false, fallback: true, error: String(e && e.message || e) };
  }
}

// ============================================================
// WRAPPERS POR PUNTO DE ESCRITURA
// Cada uno: gate OFF → {handled:false}; gate ON → intenta RPC; si falla → {handled:false,fallback:true}.
// El caller SIEMPRE corre la lógica vieja cuando handled===false (idéntico a hoy / red de seguridad).
// ============================================================

/**
 * CIERRE DE GUÍA → wh.cerrar_guia_idempotente(p_id_guia).
 * Money-safety CRÍTICO: usa la versión IDEMPOTENTE (delta = cant_recibida −
 * cantidad_aplicada). Recerrar/reabrir NO duplica (bug LOPESA −72 ×2).
 * La RPC aplica TODO el detalle de la guía en una sola transacción atómica,
 * incluyendo el envasado-skip y el FIFO de salida lo maneja la propia RPC por tipo.
 * NOTA: cerrar_guia_idempotente NO está gateada por un flag mos.config (a
 * diferencia de las demás): sólo exige service_role. Por eso el único gate es
 * el WH_ESCRITURA_DIRECTA del lado GAS.
 */
function _whCerrarGuiaDirecto(idGuia) {
  if (!_whEscrituraDirectaON()) return { handled: false };
  var r = _whRpcStock('cerrar_guia_idempotente', { p_id_guia: String(idGuia || '') });
  if (r.ok) return { handled: true, data: r.data };
  return { handled: false, fallback: true, error: r.error };
}

/**
 * AJUSTE MANUAL → wh.crear_ajuste(p jsonb).
 * Atómico (UPDATE cantidad+delta, nunca read-modify-write). Idempotente por id_ajuste.
 * Los ids los genera GAS y se pasan → mismos ids que Sheets + idempotencia real.
 */
function _whCrearAjusteDirecto(o) {
  if (!_whEscrituraDirectaON()) return { handled: false };
  var r = _whRpcStock('crear_ajuste', {
    p: {
      id_ajuste:       String(o.idAjuste || ''),
      codigo_producto: String(o.codigoProducto || ''),
      tipo:            String(o.tipo || ''),          // 'INC' | 'DEC'
      cantidad:        o.cantidad,
      motivo:          String(o.motivo || ''),
      usuario:         String(o.usuario || ''),
      id_auditoria:    String(o.idAuditoria || ''),
      id_stock_nuevo:  String(o.idStockNuevo || ''),
      id_mov:          String(o.idMov || ''),
      fecha:           o.fecha || ''
    }
  });
  if (r.ok) return { handled: true, data: r.data };
  return { handled: false, fallback: true, error: r.error };
}

/**
 * AUDITORÍA (set-absoluto) → wh.auditar_producto(p jsonb).
 * Orquestador atómico: registra auditoría EJECUTADA + ajusta stock al físico
 * por la diferencia EN UNA SOLA TRANSACCIÓN. Idempotente por id_auditoria.
 */
function _whAuditarProductoDirecto(o) {
  if (!_whEscrituraDirectaON()) return { handled: false };
  var r = _whRpcStock('auditar_producto', {
    p: {
      id_auditoria:   String(o.idAuditoria || ''),
      codigo_barra:   String(o.codigoBarra || ''),
      stock_fisico:   o.stockFisico,
      usuario:        String(o.usuario || ''),
      observacion:    String(o.observacion || ''),
      id_ajuste:      String(o.idAjuste || ''),
      id_stock_nuevo: String(o.idStockNuevo || ''),
      id_mov:         String(o.idMov || '')
    }
  });
  if (r.ok) return { handled: true, data: r.data };
  return { handled: false, fallback: true, error: r.error };
}

/**
 * ENVASADO → wh.registrar_envasado(p jsonb).
 * Orquestador atómico: consume BASE (granel) y produce DERIVADO (unidades) +
 * guías SALIDA/INGRESO_ENVASADO del día + lote + fila wh.envasados, todo en una
 * transacción. El cliente (GAS) resuelve el catálogo (cod base, cant base) y los pasa.
 * Idempotente por id_envasado.
 */
function _whRegistrarEnvasadoDirecto(o) {
  if (!_whEscrituraDirectaON()) return { handled: false };
  var r = _whRpcStock('registrar_envasado', {
    p: {
      id_envasado:          String(o.idEnvasado || ''),
      cod_producto_base:    String(o.codProductoBase || ''),
      cod_producto_envasado:String(o.codProductoEnvasado || ''),
      cantidad_base:        o.cantidadBase,
      unidades_producidas:  o.unidadesProducidas,
      unidad_base:          String(o.unidadBase || ''),
      fecha_vencimiento:    String(o.fechaVencimiento || ''),
      usuario:              String(o.usuario || 'sistema')
    }
  });
  if (r.ok) return { handled: true, data: r.data };
  return { handled: false, fallback: true, error: r.error };
}

// ════════════════════════════════════════════════════════════════════════════
// [CUTOVER · PASO FINAL] Desactivar los triggers GAS que SINCRONIZAN/PISAN Supabase.
// ----------------------------------------------------------------------------
// Tras activar la escritura directa (WH_ESCRITURA_DIRECTA=1 + flags mos.config) y apagar el
// sync de las tablas que muta (whSyncOffTablas), estos triggers time-based YA NO deben correr:
// reescribirían/sincronizarían la Hoja→Supabase o cerrarían guías sin la idempotencia de las RPCs,
// PISANDO/duplicando la verdad de la BD.
//
// SE BORRAN (solo estos handlers):
//   · syncWHReciente / syncWHCompleto      → sync Hoja→Supabase (15min / 3:30am). El sync es el que pisa.
//   · reconciliarDiarioWH                  → reconciliación nocturna (legacy/embebida en syncWHCompleto).
//   · cerrarGuiasAbiertasGlobal(+Safe)     → cierre masivo nocturno SOBRE LA HOJA (21h). Sin idempotencia de delta.
//   · auditarStockGlobal                   → auditoría nocturna que reescribe ALERTAS_STOCK en la Hoja (22h).
//   · autocerrarGuiasInactivas             → autocierre por inactividad EN LA HOJA. Lo reemplaza el cron
//                                            idempotente de Supabase 'wh-autocierre-inactividad' (*/15), que SE QUEDA.
//
// SE CONSERVAN (NO se tocan):
//   · procesarLotesPendientes              → impresión de etiquetas (PrintNode). Operacional, no sincroniza.
//   · enviarResumenEnvasadosDia / enviarResumenCargadores12 → push/reportes informativos.
//   · _jobReabrirPickupsAtascados          → reabre pickups atascados (operacional; ahora dual-write idempotente).
//   · TODOS los triggers de Seguridad (SeguridadAlerts.gs) y cualquier otro handler no listado.
//
// IDEMPOTENTE: re-correr no falla (si el trigger ya no existe, no hace nada). REVERSIBLE: reinstalar con
// instalarTriggersSyncWH() / setupTriggersAuditoria() / instalarTriggerAutocierre() si hiciera falta.
// El dueño la corre UNA vez, al final, desde el editor GAS.
// ════════════════════════════════════════════════════════════════════════════
function desactivarTriggersPisanSupabaseWH() {
  var MATAR = {
    'syncWHReciente': true,
    'syncWHCompleto': true,
    'reconciliarDiarioWH': true,
    'cerrarGuiasAbiertasGlobal': true,
    'cerrarGuiasAbiertasGlobalSafe': true,
    'auditarStockGlobal': true,
    'autocerrarGuiasInactivas': true
  };
  var borrados = [], conservados = [];
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var fn = '';
    try { fn = t.getHandlerFunction(); } catch(_) {}
    if (MATAR[fn]) {
      try { ScriptApp.deleteTrigger(t); borrados.push(fn); }
      catch(eD) { Logger.log('[desactivarTriggersPisanSupabaseWH] no se pudo borrar ' + fn + ': ' + (eD && eD.message)); }
    } else if (fn) {
      conservados.push(fn);
    }
  });
  var out = {
    ok: true,
    borrados: borrados,
    conservados: conservados,
    nota: 'El cron idempotente de Supabase wh-autocierre-inactividad reemplaza autocerrarGuiasInactivas y SE QUEDA. ' +
          'Asegurate de tener WH_ESCRITURA_DIRECTA=1 y el sync apagado (whSyncOffTablas) ANTES de correr esto.'
  };
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// ════════════════════════════════════════════════════════════════════════════
// [CUTOVER · PASO 1] RECONCILIAR la "landmine" guía ② en la HOJA.
// ----------------------------------------------------------------------------
// La guía SALIDA_ZONA G1781445112212 ya aplicó sus 59 líneas al kardex (en
// Supabase cantidad_aplicada=cant_recibida; el stock real ya las refleja). Para
// que el cierre idempotente dé delta 0 y NUNCA re-aplique — incluso por el
// FALLBACK a Hoja si el gate WH_ESCRITURA_DIRECTA estuviera OFF o la RPC fallara —
// dejamos la HOJA coherente con ese hecho:
//   · GUIA_DETALLE.cantidadAplicada = cantidadRecibida en sus 59 líneas
//   · GUIAS.estado = CERRADA (deja de aparecer ABIERTA si algún path lee la Hoja)
// Es idempotente: re-correr deja exactamente los mismos valores (delta 0).
// Bajo _conLock (toda escritura a hojas críticas). NO toca stock por sí misma:
// solo alinea el marcador para que el cierre posterior NO mueva inventario.
// ════════════════════════════════════════════════════════════════════════════
function reconciliarGuiaSheet(params) {
  params = params || {};
  var idGuia = String(params.idGuia || 'G1781445112212');
  return _conLock('reconciliarGuiaSheet', function() {
    var out = { ok: true, idGuia: idGuia, lineasAlineadas: 0, yaAlineadas: 0, estadoAntes: '', estadoDespues: '' };

    // 1) GUIA_DETALLE: cantidadAplicada = cantidadRecibida por línea de esta guía.
    var detSheet = getSheet('GUIA_DETALLE');
    if (!detSheet) return { ok: false, error: 'GUIA_DETALLE no existe' };
    var idxApl = (typeof _ensureColCantidadAplicada === 'function')
      ? _ensureColCantidadAplicada(detSheet)
      : -1;   // 0-based
    var data = detSheet.getDataRange().getValues();
    var hdrs = data[0];
    var iIdG = hdrs.indexOf('idGuia');
    var iRec = hdrs.indexOf('cantidadRecibida');
    if (iIdG < 0 || iRec < 0 || idxApl < 0) return { ok: false, error: 'Columnas idGuia/cantidadRecibida/cantidadAplicada no encontradas' };
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iIdG]) !== idGuia) continue;
      var rec = parseFloat(data[r][iRec]) || 0;
      var apl = parseFloat(data[r][idxApl]) || 0;
      if (apl === rec) { out.yaAlineadas++; continue; }
      detSheet.getRange(r + 1, idxApl + 1).setValue(rec);
      out.lineasAlineadas++;
    }

    // 2) GUIAS: marcar estado = CERRADA (idempotente).
    var gSheet = getSheet('GUIAS');
    var gData = gSheet.getDataRange().getValues();
    var gHdrs = gData[0];
    var gIid = gHdrs.indexOf('idGuia');
    var gIes = gHdrs.indexOf('estado');
    var encontrada = false;
    for (var g = 1; g < gData.length; g++) {
      if (String(gData[g][gIid]) !== idGuia) continue;
      encontrada = true;
      out.estadoAntes = String(gData[g][gIes] || '');
      if (out.estadoAntes.toUpperCase() !== 'CERRADA') {
        gSheet.getRange(g + 1, gIes + 1).setValue('CERRADA');
      }
      out.estadoDespues = 'CERRADA';
      break;
    }
    if (!encontrada) { out.estadoAntes = '(guía no está en la Hoja GUIAS)'; out.estadoDespues = '(sin cambio)'; }

    try { SpreadsheetApp.flush(); } catch(_) {}
    Logger.log('[reconciliarGuiaSheet] ' + JSON.stringify(out));
    return out;
  });
}

// ── Diagnóstico / control (ejecutar a mano desde el editor) ──
function whEscrituraDirectaEstado() {
  var o = {
    WH_ESCRITURA_DIRECTA: String(PropertiesService.getScriptProperties().getProperty('WH_ESCRITURA_DIRECTA') || '(no set)'),
    activo: _whEscrituraDirectaON()
  };
  Logger.log(JSON.stringify(o, null, 2));
  return o;
}
function whEscrituraDirectaON()  { PropertiesService.getScriptProperties().setProperty('WH_ESCRITURA_DIRECTA', '1'); Logger.log('✅ WH_ESCRITURA_DIRECTA = 1 (recordá prender también los flags mos.config WH_*_DIRECTO server-side)'); return { ok: true, activo: true }; }
function whEscrituraDirectaOFF() { PropertiesService.getScriptProperties().setProperty('WH_ESCRITURA_DIRECTA', '0'); Logger.log('↩️ WH_ESCRITURA_DIRECTA = 0 — rollback instantáneo a Hoja+dual-write'); return { ok: true, activo: false }; }

// ════════════════════════════════════════════════════════════════════════════
// [CUTOVER · PASO 5] Control del apagado del sync Hoja→Supabase por tabla.
// Setea/lee WH_SYNC_OFF_TABLAS (CSV de nombres pg). _syncWHImpl ya respeta esta
// lista (ver _whSyncOffSet en MigracionWH.gs): omite esas tablas en cada pasada,
// así el sync deja de pisar la verdad que escribe la escritura directa. REVERSIBLE:
// whSyncOnTablas() la vacía y reactiva el sync de TODO al instante.
// El set por defecto cubre exactamente las tablas que mutan las RPCs directas.
var _WH_SYNC_OFF_DEFAULT = 'stock,stock_movimientos,guias,guia_detalle,ajustes,envasados,auditorias,lotes_vencimiento';
function whSyncOffTablas(params) {
  params = params || {};
  var csv = String(params.tablas || _WH_SYNC_OFF_DEFAULT)
              .split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean).join(',');
  PropertiesService.getScriptProperties().setProperty('WH_SYNC_OFF_TABLAS', csv);
  Logger.log('🔻 WH_SYNC_OFF_TABLAS = ' + csv + ' — el sync deja de re-upsertear esas tablas desde la Hoja');
  return { ok: true, off: csv.split(',') };
}
function whSyncOnTablas() {
  PropertiesService.getScriptProperties().deleteProperty('WH_SYNC_OFF_TABLAS');
  Logger.log('🔼 WH_SYNC_OFF_TABLAS borrada — sync de TODAS las tablas reactivado (rollback)');
  return { ok: true, off: [] };
}
function whSyncOffEstado() {
  var v = String(PropertiesService.getScriptProperties().getProperty('WH_SYNC_OFF_TABLAS') || '');
  var o = { WH_SYNC_OFF_TABLAS: v || '(vacío → sync de todo)', off: v ? v.split(',') : [] };
  Logger.log(JSON.stringify(o));
  return o;
}

/* ════════════════════════════════════════════════════════════════════════════
 * PLAN DE APAGADO DEL SYNC — NO EJECUTAR. Para revisión del dueño.
 * ════════════════════════════════════════════════════════════════════════════
 * La escritura directa de stock NO puede COEXISTIR con el sync Hoja→Supabase de
 * las tablas que muta (el sync pisa lo que la RPC escribió, y el read-back vería
 * dos verdades → duplicación). Por eso el apagado del sync es un paso aparte,
 * a coordinar por el dueño DESPUÉS de validar la escritura directa con flota 100%.
 *
 * Secuencia segura propuesta:
 *  1. Desplegar este GAS con WH_ESCRITURA_DIRECTA OFF (inerte, idéntico a hoy).
 *  2. Prender los flags server-side por RPC EN mos.config, de a uno, validando:
 *       WH_CREAR_AJUSTE_DIRECTO=1, WH_AUDITAR_PRODUCTO_DIRECTO=1,
 *       WH_REGISTRAR_ENVASADO_DIRECTO=1.
 *     (cerrar_guia_idempotente no tiene flag server-side: sólo service_role.)
 *  3. Prender WH_ESCRITURA_DIRECTA=1 en una sola unidad / horario tranquilo.
 *     Observar wh.stock vs Hoja con la auditoría de cuadre (wh.auditar_cuadre_stock,
 *     supabase/71). Mientras el sync siga vivo, ambos deben converger.
 *  4. Con la escritura directa validada y la flota actualizada al GAS nuevo:
 *     APAGAR el sync SÓLO de las tablas que ahora escribe la RPC, para que el
 *     sync deje de pisar la verdad de Supabase:
 *        WH_SYNC_OFF_TABLAS debe incluir, como mínimo:
 *           wh.stock, wh.stock_movimientos, wh.ajustes,
 *           wh.guias, wh.guia_detalle, wh.envasados, wh.auditorias, wh.lotes_vencimiento
 *     (revisar el nombre EXACTO de la lista de exclusión del sync en MigracionWH.gs
 *      antes de tocarla — este archivo NO la modifica).
 *  5. A partir de ahí, la Hoja deja de ser fuente de verdad de stock. La escritura
 *     a la Hoja puede seguir como respaldo de sólo-lectura/auditoría hasta el corte
 *     final de Sheets, pero ya NO debe re-sincronizarse hacia Supabase.
 *
 * RIESGO si se apaga el sync ANTES de tener la escritura directa 100% en la flota:
 *   un dispositivo con GAS viejo (gate OFF) escribiría sólo la Hoja, y sin sync esa
 *   escritura nunca llegaría a Supabase → desync silencioso. Por eso: flota 100% y
 *   escritura directa validada PRIMERO; sync OFF DESPUÉS.
 * ════════════════════════════════════════════════════════════════════════════ */
