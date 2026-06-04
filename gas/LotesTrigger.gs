// ============================================================
// warehouseMos — LotesTrigger.gs   [v2.13.118]
// ============================================================
//
// FIRE-AND-FORGET orquestación de lotes de etiquetas.
//
// El frontend crea un lote → status ENCOLADO.
// Trigger time-based cada 1 min ejecuta procesarLotesPendientes()
// que toma siguiente sub-job de cada lote y lo manda a PrintNode.
//
// Si el operador cierra la app, el lote SIGUE procesándose porque
// el trigger corre en el backend GAS independiente del frontend.
//
// Setup (1 vez desde editor):
//   instalarTriggerLotesEtiqueta()

// ────────────────────────────────────────────────────────────────────
// procesarLotesPendientes — corre cada 1 min via trigger
// ────────────────────────────────────────────────────────────────────
function procesarLotesPendientes() {
  // LockService evita que 2 ejecuciones se pisen (trigger doble fire).
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('[Trigger] otro proceso en curso, skip');
    return;
  }
  try {
    var sheet = _getSheetLotesAdhesivo();
    var rows = _sheetToObjects(sheet);
    // [AUDIT FIX #3] Filtramos también por fechaUltimoUpdate: si está
    // IMPRIMIENDO pero el update fue hace <90s, asumimos que otro proceso
    // lo está manejando AHORA. Si está IMPRIMIENDO con update >90s atrás,
    // probablemente crasheó — lo re-tomamos.
    var nowMs = new Date().getTime();
    var pendientes = rows.filter(function(r) {
      var s = String(r.status || '').toUpperCase();
      var completadas = parseInt(r.completadas) || 0;
      var total       = parseInt(r.totalEtq)    || 0;
      if (completadas >= total) return false;
      if (s === 'ENCOLADO') return true;
      if (s === 'IMPRIMIENDO') {
        var lastUpd = String(r.fechaUltimoUpdate || '');
        if (!lastUpd) return true;
        var lastMs = new Date(lastUpd).getTime();
        if (isNaN(lastMs)) return true;
        // Si pasó >90s desde el último update, considerarlo abandonado
        return (nowMs - lastMs) > 90000;
      }
      return false;
    });
    Logger.log('[Trigger] ' + pendientes.length + ' lotes pendientes');

    // Por ciclo, procesamos UN sub-job de UN lote.
    // Razón: el polling state de PrintNode puede tomar 20s; si procesamos
    // muchos lotes en 1 ciclo, podríamos exceder los 6 min de GAS.
    // Próximo ciclo de trigger (1 min) toma el siguiente.
    if (pendientes.length === 0) return;

    // Priorizar lotes más viejos (FIFO)
    pendientes.sort(function(a, b) {
      return String(a.fechaCreacion).localeCompare(String(b.fechaCreacion));
    });

    // Procesar primero de la cola
    var lote = pendientes[0];
    var idLote = String(lote.idLote);
    var tipo = String(lote.tipoEtiqueta || 'ADHESIVO_ENVASADO').toUpperCase();

    Logger.log('[Trigger] procesando ' + idLote + ' tipo=' + tipo + ' completadas=' + lote.completadas + '/' + lote.totalEtq);

    var result;
    if (tipo === 'MEMBRETE_ME' || tipo === 'MEMBRETE_WH') {
      result = imprimirSubLoteMembrete(idLote);
    } else {
      // ADHESIVO_ENVASADO (legacy)
      result = imprimirSubLoteAdhesivo({ idLote: idLote });
    }

    if (result && result.ok === false) {
      Logger.log('[Trigger] error en ' + idLote + ': ' + result.error);
    } else {
      Logger.log('[Trigger] OK: ' + idLote + ' completadas=' + (result.data && result.data.completadas));
    }
  } catch(e) {
    Logger.log('[Trigger] excepción: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// [v2.13.130] Lazy installer — chequea si el trigger existe, lo instala si falta.
// Llamado automáticamente desde crearLoteMembrete y crearLoteAdhesivo para
// garantizar que los lotes ENCOLADO se procesen sin requerir setup manual.
function _asegurarTriggerLotes() {
  try {
    var TRG = 'procesarLotesPendientes';
    var existe = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === TRG;
    });
    if (existe) return { ok: true, yaExistia: true };
    // Auto-instalar
    ScriptApp.newTrigger(TRG).timeBased().everyMinutes(1).create();
    Logger.log('[_asegurarTriggerLotes] AUTO-INSTALADO ' + TRG);
    return { ok: true, instaladoAhora: true };
  } catch(e) {
    Logger.log('[_asegurarTriggerLotes] error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// [v2.13.130] Diagnóstico del trigger — endpoint público para el frontend.
function diagnosticoTriggerLotes() {
  try {
    var TRG = 'procesarLotesPendientes';
    var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
      return t.getHandlerFunction() === TRG;
    });
    var sheet = _getSheetLotesAdhesivo();
    var rows = _sheetToObjects(sheet);
    var encolados = rows.filter(function(r) { return String(r.status||'').toUpperCase() === 'ENCOLADO'; });
    var imprimiendo = rows.filter(function(r) { return String(r.status||'').toUpperCase() === 'IMPRIMIENDO'; });
    return { ok: true, data: {
      triggerInstalado: triggers.length > 0,
      cantidadTriggers: triggers.length,
      lotesEncolados: encolados.length,
      lotesImprimiendo: imprimiendo.length,
      idLotesEncolados: encolados.map(function(r){ return String(r.idLote); }).slice(0, 10),
      mensaje: triggers.length === 0
        ? '⚠ TRIGGER NO INSTALADO — los lotes ENCOLADO no se procesan. Llamá auto-fix o instalarTriggerLotesEtiqueta() manualmente.'
        : '✅ Trigger activo — ' + encolados.length + ' lotes en cola se procesarán en próximo minuto.'
    }};
  } catch(e) { return { ok: false, error: e.message }; }
}

// ────────────────────────────────────────────────────────────────────
// instalarTriggerLotesEtiqueta — UNA VEZ desde editor GAS
// ────────────────────────────────────────────────────────────────────
function instalarTriggerLotesEtiqueta() {
  var TRG = 'procesarLotesPendientes';
  // Borrar triggers previos del mismo handler
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === TRG) ScriptApp.deleteTrigger(t);
  });
  // Crear nuevo: cada 1 min
  ScriptApp.newTrigger(TRG)
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('[Trigger] ' + TRG + ' instalado · cada 1 min');
  return { ok: true, mensaje: 'Trigger instalado correctamente. Lotes ENCOLADO se procesan automáticamente cada 1 minuto.' };
}

function desinstalarTriggerLotesEtiqueta() {
  var TRG = 'procesarLotesPendientes';
  var borrados = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === TRG) {
      ScriptApp.deleteTrigger(t);
      borrados++;
    }
  });
  Logger.log('[Trigger] ' + borrados + ' triggers borrados');
  return { ok: true, borrados: borrados };
}

// ────────────────────────────────────────────────────────────────────
// procesarUnLoteAhora — manual desde editor (debugging)
// ────────────────────────────────────────────────────────────────────
function procesarUnLoteAhora() {
  procesarLotesPendientes();
}
