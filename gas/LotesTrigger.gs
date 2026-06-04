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
  // [v2.13.140] Logging persistente + procesamiento múltiple + inferencia de tipo + skip robusto
  var resumen = { intentados: 0, ok: 0, errores: 0, idsConError: [], inicio: new Date().toISOString() };
  try {
    var sheet = _getSheetLotesAdhesivo();
    var rows = _sheetToObjects(sheet);
    var nowMs = new Date().getTime();
    var pendientes = rows.filter(function(r) {
      var s = String(r.status || '').toUpperCase();
      var completadas = parseInt(r.completadas) || 0;
      var total       = parseInt(r.totalEtq)    || 0;
      if (completadas >= total) return false;
      if (s === 'ENCOLADO') return true;
      // [v2.13.140] CREADO también se procesa (membretes y adhesivos abandonados)
      if (s === 'CREADO') return true;
      if (s === 'IMPRIMIENDO' || s === 'CALIBRANDO') {
        var lastUpd = String(r.fechaUltimoUpdate || '');
        if (!lastUpd) return true;
        var lastMs = new Date(lastUpd).getTime();
        if (isNaN(lastMs)) return true;
        return (nowMs - lastMs) > 90000;  // abandonado >90s
      }
      return false;
    });
    Logger.log('[Trigger] ' + pendientes.length + ' lotes pendientes');
    if (pendientes.length === 0) return _persistirResumen(resumen);

    // FIFO (más viejos primero)
    pendientes.sort(function(a, b) {
      return String(a.fechaCreacion).localeCompare(String(b.fechaCreacion));
    });

    // [v2.13.140] Procesar hasta MAX_LOTES_POR_CICLO. PrintNode polling ~25s
    // por lote; con 5 min máximo de ejecución GAS, podemos procesar ~10 lotes
    // seguros. Saltamos los que fallan para no quedar atorados en uno solo.
    var MAX_POR_CICLO = parseInt(
      PropertiesService.getScriptProperties().getProperty('TRIGGER_MAX_LOTES_POR_CICLO')
    ) || 8;
    var tiempoLimiteMs = nowMs + 4 * 60 * 1000;  // máx 4 min por ciclo (margen vs 6min GAS)

    // [v2.13.141] FIRE-AND-FORGET COMPLETO: cada lote se procesa HASTA
    // TERMINAR (todos los sub-jobs en serie) en lugar de 1 sub-job por
    // ciclo. Así un lote de 100 etiquetas se imprime en ~5 min en 1 ciclo,
    // no en 10 ciclos = 10 min. El operador puede cerrar la app después
    // de crear el lote y el backend completa todo solo.
    for (var i = 0; i < pendientes.length && i < MAX_POR_CICLO; i++) {
      if (new Date().getTime() > tiempoLimiteMs) {
        Logger.log('[Trigger] tiempo límite alcanzado, parando en ' + i + '/' + pendientes.length);
        break;
      }
      var lote = pendientes[i];
      var idLote = String(lote.idLote);
      var tipo = String(lote.tipoEtiqueta || '').toUpperCase();
      if (!tipo) {
        var desc = String(lote.descripcion || '');
        if (desc.indexOf('ME:') === 0)      tipo = 'MEMBRETE_ME';
        else if (desc.indexOf('WH:') === 0) tipo = 'MEMBRETE_WH';
        else                                 tipo = 'ADHESIVO_ENVASADO';
      }
      resumen.intentados++;
      Logger.log('[Trigger ' + (i+1) + '/' + Math.min(MAX_POR_CICLO, pendientes.length) + '] procesando ' + idLote + ' tipo=' + tipo);

      // [v2.13.141] Loop interno: procesar TODOS los sub-jobs del lote hasta
      // que termine (COMPLETADO), falle (PAUSADO_*) o se agote tiempo del ciclo.
      var subJobsHechos = 0;
      var MAX_SUBJOBS_POR_LOTE = 30;  // hard cap por defensa (300 etiquetas máx por lote por ciclo)
      var loteOk = true;
      var ultimoResult = null;
      while (loteOk && subJobsHechos < MAX_SUBJOBS_POR_LOTE) {
        if (new Date().getTime() > tiempoLimiteMs) {
          Logger.log('[Trigger] tiempo agotado dentro de lote ' + idLote + ' tras ' + subJobsHechos + ' sub-jobs');
          break;
        }
        try {
          if (tipo === 'MEMBRETE_ME' || tipo === 'MEMBRETE_WH') {
            ultimoResult = imprimirSubLoteMembrete(idLote);
          } else {
            ultimoResult = imprimirSubLoteAdhesivo({ idLote: idLote });
          }
          subJobsHechos++;
          // Si falló → marcar PAUSADO_ERROR y romper
          if (ultimoResult && ultimoResult.ok === false) {
            loteOk = false;
            resumen.errores++;
            resumen.idsConError.push(idLote + ': ' + (ultimoResult.error || 'unknown'));
            Logger.log('[Trigger] error en sub-job de ' + idLote + ': ' + ultimoResult.error);
            try {
              var rowIdx = _findLoteRow(sheet, idLote);
              if (rowIdx >= 0) {
                _patchLote(sheet, rowIdx, {
                  status: 'PAUSADO_ERROR',
                  ultimoError: 'Trigger: ' + String(ultimoResult.error || '').substring(0, 200)
                });
              }
            } catch(_){}
            break;
          }
          // Si terminó el lote → break
          var st = ultimoResult && ultimoResult.data && String(ultimoResult.data.status || '').toUpperCase();
          if (st === 'COMPLETADO' || st === 'CANCELADO' || (st && st.indexOf('PAUSADO') === 0)) {
            Logger.log('[Trigger] lote ' + idLote + ' finalizó con status=' + st + ' (' + subJobsHechos + ' sub-jobs)');
            if (st === 'COMPLETADO') resumen.ok++;
            break;
          }
          // Si qtyImpresa=0 (nada que hacer) → break para no loopear infinito
          if (ultimoResult && ultimoResult.data && ultimoResult.data.qtyImpresa === 0) {
            Logger.log('[Trigger] lote ' + idLote + ' devolvió qtyImpresa=0, parando');
            break;
          }
        } catch (eIter) {
          loteOk = false;
          resumen.errores++;
          resumen.idsConError.push(idLote + ': EXC ' + eIter.message);
          Logger.log('[Trigger] EXCEPCIÓN en sub-job de ' + idLote + ': ' + eIter.message);
          break;
        }
      }
      if (loteOk && subJobsHechos > 0 && !resumen.ok) resumen.ok++;
    }
  } catch(e) {
    Logger.log('[Trigger] excepción global: ' + e.message);
    resumen.errores++;
    resumen.idsConError.push('GLOBAL: ' + e.message);
  } finally {
    try { lock.releaseLock(); } catch(_){}
    resumen.fin = new Date().toISOString();
    _persistirResumen(resumen);
  }
}

// [v2.13.140] Persistir resumen del último ciclo para diagnóstico
function _persistirResumen(resumen) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      'TRIGGER_ULTIMO_RESUMEN',
      JSON.stringify(resumen).substring(0, 8000)
    );
  } catch(_){}
}

// [v2.13.140] Diagnóstico PrintNode adhesivo: dice cuál printerId está
// configurado, su estado en PrintNode y los últimos N jobs enviados.
// Ayuda a detectar cuando el printerId de IMPRESORAS apunta a la impresora
// EQUIVOCADA (ej: una ESC/POS de tickets en vez de la TSC de adhesivos).
function diagnosticoPrintNodeAdhesivo() {
  var diag = { printerId: '', impresoraInfo: null, jobsRecientes: [], errores: [] };
  try {
    var printerId;
    try { printerId = String(getPrinterNodeId('ADHESIVO', 'ALMACEN')); }
    catch (e) {
      diag.errores.push('IMPRESORAS no tiene tipo=ADHESIVO zona=ALMACEN activa: ' + e.message);
      return { ok: true, data: diag };
    }
    diag.printerId = printerId;

    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) {
      diag.errores.push('PRINTNODE_API_KEY no configurada');
      return { ok: true, data: diag };
    }
    var auth = 'Basic ' + Utilities.base64Encode(apiKey + ':');

    // 1) Info de la impresora configurada
    try {
      var rImp = UrlFetchApp.fetch('https://api.printnode.com/printers/' + printerId, {
        headers: { 'Authorization': auth }, muteHttpExceptions: true
      });
      if (rImp.getResponseCode() === 200) {
        var arr = JSON.parse(rImp.getContentText());
        var p = Array.isArray(arr) ? arr[0] : arr;
        if (p) {
          diag.impresoraInfo = {
            id: p.id,
            name: p.name || '',
            description: p.description || '',
            state: (p.computer && p.computer.state) || p.state || 'unknown',
            online: (p.computer && p.computer.state === 'connected'),
            computer: (p.computer && p.computer.name) || ''
          };
        }
      } else {
        diag.errores.push('PrintNode rechazó getPrinter: HTTP ' + rImp.getResponseCode());
      }
    } catch (eP) { diag.errores.push('Error consultando impresora: ' + eP.message); }

    // 2) Últimos 10 jobs enviados a esa impresora
    try {
      var rJobs = UrlFetchApp.fetch('https://api.printnode.com/printers/' + printerId + '/printjobs?limit=10', {
        headers: { 'Authorization': auth }, muteHttpExceptions: true
      });
      if (rJobs.getResponseCode() === 200) {
        var jobs = JSON.parse(rJobs.getContentText());
        diag.jobsRecientes = (jobs || []).map(function(j) {
          return {
            id: j.id,
            title: String(j.title || '').substring(0, 60),
            state: j.state || 'unknown',
            createTimestamp: j.createTimestamp || '',
            source: String(j.source || '').substring(0, 40)
          };
        });
      } else {
        diag.errores.push('PrintNode rechazó getJobs: HTTP ' + rJobs.getResponseCode());
      }
    } catch (eJ) { diag.errores.push('Error consultando jobs: ' + eJ.message); }

  } catch (e) {
    diag.errores.push('Excepción global: ' + e.message);
  }
  return { ok: true, data: diag };
}

// [v2.13.140] Endpoint manual: procesar AHORA todos los pendientes desde el panel.
// Útil cuando hay muchos lotes en cola y el operador no quiere esperar.
function procesarAhoraTodos() {
  procesarLotesPendientes();
  var resumen = {};
  try { resumen = JSON.parse(PropertiesService.getScriptProperties().getProperty('TRIGGER_ULTIMO_RESUMEN') || '{}'); } catch(_){}
  return { ok: true, data: resumen };
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
