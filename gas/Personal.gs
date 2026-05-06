// ============================================================
// warehouseMos — Personal.gs
// Login por PIN, sesiones, desempeño y cierre de turno
// ============================================================

// ── Login ───────────────────────────────────────────────────
function loginPersonal(params) {
  var pin = String(params.pin || '').trim();
  if (!pin) return { ok: false, error: 'PIN requerido' };

  var personal = _getPersonalWH();
  var operador = personal.find(function(p) {
    return String(p.pin) === pin;
  });
  if (!operador) return { ok: false, error: 'PIN incorrecto' };

  // Bloqueo por horario laboral — solo afecta a roles no-admin
  var horario = _horarioPermitido(operador.rol);
  if (!horario.permitido) {
    return {
      ok: false,
      error: 'FUERA_DE_HORARIO',
      data: {
        rol:       operador.rol,
        nombre:    operador.nombre,
        apertura:  horario.apertura,
        cierre:    horario.cierre,
        dia:       horario.dia,
        motivo:    horario.motivo
      }
    };
  }

  var tz       = Session.getScriptTimeZone();
  var ahora    = new Date();
  var fechaStr = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
  var horaStr  = Utilities.formatDate(ahora, tz, 'HH:mm:ss');

  // ── Revisar si ya existe sesión hoy ───────────────────────────
  var todasSesiones  = _sheetToObjects(getSheet('SESIONES'));
  var sesionHoyActiva = null;
  var tuvoPreviaHoy   = false;

  for (var si = 0; si < todasSesiones.length; si++) {
    var s = todasSesiones[si];
    if (String(s.idPersonal) !== String(operador.idPersonal)) continue;
    var sfec = String(s.fechaInicio || '').substring(0, 10);
    if (sfec !== fechaStr) continue;
    tuvoPreviaHoy = true;
    if (s.estado === 'ACTIVA') { sesionHoyActiva = s; break; }
  }

  // Segundo dispositivo o reapertura en el día: devolver sesión existente
  if (sesionHoyActiva) {
    return {
      ok: true,
      data: {
        idSesion:          sesionHoyActiva.idSesion,
        idPersonal:        operador.idPersonal,
        nombre:            operador.nombre,
        apellido:          operador.apellido,
        rol:               operador.rol,
        color:             operador.color,
        horaInicio:        sesionHoyActiva.horaInicio,
        yaEnSesionHoy:     true,
        bienvenidaImpresa: true
      }
    };
  }

  // ── Primera sesión del día (o re-login tras cierre) ───────────
  var huerfanas = _cerrarSesionesHuerfanas(operador.idPersonal);

  var idSesion    = _generateId('SES');
  var idDesempeno = _generateId('DES');

  getSheet('SESIONES').appendRow([
    idSesion, operador.idPersonal, fechaStr, horaStr, '', '', 0, 'ACTIVA'
  ]);

  getSheet('DESEMPENO').appendRow([
    idDesempeno, operador.idPersonal, idSesion, fechaStr,
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '',
    parseFloat(operador.montoBase) || 0, 0, 0, 'ACTIVA'
  ]);

  try {
    _registrarJornadaEnMOS(
      operador.nombre + ' ' + (operador.apellido || ''),
      operador.rol,
      parseFloat(operador.montoBase) || 0
    );
  } catch(eJ) { Logger.log('Auto-jornada MOS: ' + eJ.message); }

  try {
    _notificarMOS(
      '👤 ' + operador.nombre + ' ingresó a Almacén',
      (operador.rol || 'Operador') + ' · ' + horaStr,
      // Solo admins/master reciben + no auto-notificarse si entró un admin
      (operador.nombre + ' ' + (operador.apellido || '')).trim()
    );
  } catch(eP) { Logger.log('Push login WH: ' + eP.message); }

  return {
    ok: true,
    data: {
      idSesion:          idSesion,
      idPersonal:        operador.idPersonal,
      nombre:            operador.nombre,
      apellido:          operador.apellido,
      rol:               operador.rol,
      color:             operador.color,
      horaInicio:        horaStr,
      yaEnSesionHoy:     false,
      bienvenidaImpresa: tuvoPreviaHoy,   // ya hubo sesión hoy → ticket ya fue impreso
      sesionAnterior:    huerfanas.ultimaFecha || null
    }
  };
}

// Registra la jornada del operador en ProyectoMOS al iniciar sesión.
// Idempotente: si ya existe una jornada con el mismo nombre y fecha no inserta duplicados.
// Cubre el caso multi-dispositivo: tablet + celular → solo la primera sesión registra.
function _registrarJornadaEnMOS(nombre, rol, montoBase) {
  nombre = String(nombre || '').trim();
  if (!nombre) return;
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  if (!mosSsId) return;

  var tz    = Session.getScriptTimeZone();
  var fecha = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var ss    = SpreadsheetApp.openById(mosSsId);
  var sheet = ss.getSheetByName('JORNADAS');
  if (!sheet) return;

  var tz2  = Session.getScriptTimeZone();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var fechaFila = data[i][1] instanceof Date
      ? Utilities.formatDate(data[i][1], tz2, 'yyyy-MM-dd')
      : String(data[i][1] || '').substring(0, 10);
    if (String(data[i][3]).toLowerCase() === nombre.toLowerCase() && fechaFila === fecha) return;
  }

  sheet.appendRow([
    'JOR' + new Date().getTime(), fecha, '', nombre,
    rol || 'OPERADOR', 'warehouseMos', '', parseFloat(montoBase) || 0,
    '', 'AUTO', 'AUTO_LOGIN'
  ]);
}

// ── Diagnóstico: corre desde el editor GAS para verificar conexión con MOS ──
function testMosConexion() {
  var mosSsId = PropertiesService.getScriptProperties().getProperty('MOS_SS_ID');
  Logger.log('MOS_SS_ID leído: ' + mosSsId);
  if (!mosSsId) { Logger.log('ERROR: MOS_SS_ID no está en Script Properties'); return; }
  try {
    var ss = SpreadsheetApp.openById(mosSsId);
    Logger.log('Spreadsheet abierto: ' + ss.getName());
    var sheet = ss.getSheetByName('JORNADAS');
    Logger.log('Hoja JORNADAS: ' + (sheet ? 'encontrada (' + sheet.getLastRow() + ' filas)' : 'NO ENCONTRADA'));
  } catch(e) {
    Logger.log('ERROR al abrir spreadsheet: ' + e.message);
  }
}

// ── Cerrar turno ────────────────────────────────────────────
function cerrarTurno(params) {
  var idSesion  = params.idSesion;
  var forzado   = params.forzado === true || params.forzado === 'true';
  if (!idSesion) return { ok: false, error: 'idSesion requerido' };

  var sesSheet = getSheet('SESIONES');
  var sesData  = sesSheet.getDataRange().getValues();
  var sesHdrs  = sesData[0];
  var idxSesId = sesHdrs.indexOf('idSesion');

  var filaS = -1;
  var sesRow;
  for (var i = 1; i < sesData.length; i++) {
    if (sesData[i][idxSesId] === idSesion) {
      filaS  = i + 1;
      sesRow = sesData[i];
      break;
    }
  }
  if (filaS < 0) return { ok: false, error: 'Sesión no encontrada' };

  var ahora       = new Date();
  var horaFin     = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm:ss');
  var fechaFin    = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var horaInicio  = sesRow[sesHdrs.indexOf('horaInicio')];
  var fechaInicio = sesRow[sesHdrs.indexOf('fechaInicio')];

  // Normalizar fechaInicio: Sheets puede devolverla como objeto Date
  var fechaInicioStr = (fechaInicio instanceof Date)
    ? Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(fechaInicio).substring(0, 10);
  // Calcular minutos trabajados
  var inicio = new Date(fechaInicioStr + 'T' + horaInicio);
  var minutos = Math.round((ahora - inicio) / 60000);
  var horas   = Math.round(minutos / 60 * 100) / 100;

  // Actualizar SESIONES
  sesSheet.getRange(filaS, sesHdrs.indexOf('fechaFin')       + 1).setValue(fechaFin);
  sesSheet.getRange(filaS, sesHdrs.indexOf('horaFin')        + 1).setValue(horaFin);
  sesSheet.getRange(filaS, sesHdrs.indexOf('minutosActivos') + 1).setValue(minutos);
  sesSheet.getRange(filaS, sesHdrs.indexOf('estado')         + 1).setValue(forzado ? 'FORZADA' : 'CERRADA');

  // Actualizar DESEMPENO con métricas finales
  var reporte = _calcularYCerrarDesempeno(idSesion, minutos, horas);

  return { ok: true, data: reporte };
}

// ── Registrar actividad (llamado desde otros módulos) ────────
function registrarActividad(idSesion, tipo, cantidad) {
  if (!idSesion) return;
  cantidad = parseInt(cantidad) || 1;

  var sheet   = getSheet('DESEMPENO');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxSes  = hdrs.indexOf('idSesion');
  var idxEst  = hdrs.indexOf('estado');

  var colMap = {
    'GUIA_CREADA':           hdrs.indexOf('guiasCreadas'),
    'GUIA_CERRADA':          hdrs.indexOf('guiasCerradas'),
    'ENVASADO_REGISTRADO':   hdrs.indexOf('envasadosRegistrados'),
    'UNIDADES_ENVASADAS':    hdrs.indexOf('unidadesEnvasadas'),
    'MERMA_REGISTRADA':      hdrs.indexOf('mermasRegistradas'),
    'AUDITORIA_EJECUTADA':   hdrs.indexOf('auditoriaEjecutadas'),
    'PREINGRESO_CREADO':     hdrs.indexOf('preingresoCreados'),
    'AJUSTE_REALIZADO':      hdrs.indexOf('ajustesRealizados')
  };

  var col = colMap[tipo];
  if (col === undefined || col < 0) return;

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxSes] === idSesion && data[i][idxEst] === 'ACTIVA') {
      var actual = parseInt(data[i][col]) || 0;
      sheet.getRange(i + 1, col + 1).setValue(actual + cantidad);

      // Actualizar totalActividades
      var totalCol = hdrs.indexOf('totalActividades');
      var totalActual = parseInt(data[i][totalCol]) || 0;
      sheet.getRange(i + 1, totalCol + 1).setValue(totalActual + cantidad);
      break;
    }
  }
}

// ── Calcular y cerrar desempeño ──────────────────────────────
function _calcularYCerrarDesempeno(idSesion, minutos, horas) {
  var sheet = getSheet('DESEMPENO');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0];
  var idxSes = hdrs.indexOf('idSesion');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxSes] !== idSesion) continue;

    var row = data[i];
    var get = function(col) { return parseFloat(row[hdrs.indexOf(col)]) || 0; };

    var total       = get('totalActividades');
    var actPorHora  = horas > 0 ? Math.round(total / horas * 10) / 10 : 0;
    var puntuacion  = Math.min(10, Math.round(actPorHora * 10) / 10);
    var calificacion = puntuacion >= 9 ? 'EXCELENTE'
                     : puntuacion >= 7 ? 'BUENO'
                     : puntuacion >= 5 ? 'REGULAR'
                     : 'BAJO';

    var montoBase   = get('montoBase');
    var bonusMin    = parseFloat(_getConfigValue('BONUS_PUNTUACION_MIN')) || 8;
    var bonusPct    = parseFloat(_getConfigValue('BONUS_PORCENTAJE'))     || 10;
    var monoBonus   = puntuacion >= bonusMin ? Math.round(montoBase * bonusPct / 100 * 100) / 100 : 0;
    var montoTotal  = montoBase + monoBonus;

    // Actualizar fila
    var updates = {
      'minutosActivos':    minutos,
      'horasTrabajadas':   horas,
      'totalActividades':  total,
      'actividadesPorHora':actPorHora,
      'puntuacion':        puntuacion,
      'calificacion':      calificacion,
      'montoBonus':        monoBonus,
      'montoTotal':        montoTotal,
      'estado':            'CERRADO'
    };
    Object.keys(updates).forEach(function(k) {
      var col = hdrs.indexOf(k);
      if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(updates[k]);
    });

    // Construir reporte para mostrar al operador
    var idPersonal = row[hdrs.indexOf('idPersonal')];
    var personal   = _getPersonalWH();
    var operador   = personal.find(function(p){ return p.idPersonal === idPersonal; }) || {};

    return {
      nombre:           operador.nombre + ' ' + operador.apellido,
      rol:              operador.rol,
      color:            operador.color,
      horasTrabajadas:  horas,
      minutosActivos:   minutos,
      guiasCreadas:     get('guiasCreadas'),
      guiasCerradas:    get('guiasCerradas'),
      envasadosRegistrados: get('envasadosRegistrados'),
      unidadesEnvasadas:    get('unidadesEnvasadas'),
      mermasRegistradas:    get('mermasRegistradas'),
      auditoriaEjecutadas:  get('auditoriaEjecutadas'),
      preingresoCreados:    get('preingresoCreados'),
      ajustesRealizados:    get('ajustesRealizados'),
      totalActividades:  total,
      actividadesPorHora:actPorHora,
      puntuacion:        puntuacion,
      calificacion:      calificacion,
      montoBase:         montoBase,
      montoBonus:        monoBonus,
      montoTotal:        montoTotal
    };
  }
  return null;
}

// ── Getters ──────────────────────────────────────────────────
function getPersonal(params) {
  var rows = _getPersonalWH();
  // No enviar el PIN al frontend
  rows = rows.map(function(r) {
    var safe = Object.assign({}, r);
    delete safe.pin;
    return safe;
  });
  return { ok: true, data: rows };
}

// Solo para precarga offline — incluye PIN para validación local
// Ahora servido por descargarMaestros(); mantenemos este endpoint por compatibilidad
function getPersonalConPin(params) {
  return { ok: true, data: _getPersonalWH() }; // PIN incluido intencionalmente para caché local
}

function getSesionActiva(params) {
  var idSesion = params.idSesion;
  var sesiones = _sheetToObjects(getSheet('SESIONES'));
  var ses = sesiones.find(function(s){ return s.idSesion === idSesion && s.estado === 'ACTIVA'; });
  if (!ses) return { ok: false, error: 'Sesión inválida o expirada' };
  return { ok: true, data: ses };
}

function getDesempenoDia(params) {
  var rows = _sheetToObjects(getSheet('DESEMPENO'));
  if (params.idPersonal) rows = rows.filter(function(r){ return r.idPersonal === params.idPersonal; });
  if (params.fecha)      rows = rows.filter(function(r){ return r.fecha === params.fecha; });
  return { ok: true, data: rows };
}

function getResumenPersonal(params) {
  // Para supervisor: resumen de todos los operadores del día
  var hoy  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var fecha = params.fecha || hoy;
  var desempeno = _sheetToObjects(getSheet('DESEMPENO')).filter(function(r){ return r.fecha === fecha; });
  var personal  = _getPersonalWH();
  var resumen   = personal.map(function(p) {
    var des = desempeno.find(function(d){ return d.idPersonal === p.idPersonal; }) || {};
    return {
      idPersonal:      p.idPersonal,
      nombre:          p.nombre + ' ' + p.apellido,
      rol:             p.rol,
      color:           p.color,
      estado:          des.estado || 'SIN_TURNO',
      horasTrabajadas: des.horasTrabajadas || 0,
      totalActividades:des.totalActividades || 0,
      calificacion:    des.calificacion || '—',
      montoTotal:      des.montoTotal || 0
    };
  });
  return { ok: true, data: resumen };
}

// ── Helpers internos ─────────────────────────────────────────
// Cierra sesiones ACTIVA del operador. Retorna { count, ultimaFecha }
// donde ultimaFecha es la fecha de inicio si era de un día anterior (para aviso al usuario).
function _cerrarSesionesHuerfanas(idPersonal) {
  var sheet   = getSheet('SESIONES');
  var data    = sheet.getDataRange().getValues();
  var hdrs    = data[0];
  var idxIdP  = hdrs.indexOf('idPersonal');
  var idxEst  = hdrs.indexOf('estado');
  var idxFec  = hdrs.indexOf('fechaInicio');
  var tz      = Session.getScriptTimeZone();
  var hoy     = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var count   = 0;
  var ultimaFecha = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxIdP] === idPersonal && data[i][idxEst] === 'ACTIVA') {
      sheet.getRange(i + 1, idxEst + 1).setValue('HUERFANA');
      count++;
      var fec = String(data[i][idxFec] || '').substring(0, 10);
      // Solo reportar si la sesión huérfana era de otro día (no de hoy)
      if (fec && fec !== hoy) ultimaFecha = fec;
    }
  }
  return { count: count, ultimaFecha: ultimaFecha };
}


// ── Notificar a ProyectoMOS vía push (requiere MOS_WEB_APP_URL en Script Properties) ──
// Solo manda a MASTER/ADMIN y excluye al sender si fue un admin (auto-exclusión).
function _notificarMOS(titulo, cuerpo, excluirUsuario) {
  var url = PropertiesService.getScriptProperties().getProperty('MOS_WEB_APP_URL');
  if (!url) { Logger.log('[Push] MOS_WEB_APP_URL no configurada'); return; }
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        action: 'enviarPushNotif',
        titulo: titulo,
        cuerpo: cuerpo,
        soloRolesAdmin: true,
        excluirUsuario: excluirUsuario || null
      }),
      muteHttpExceptions: true
    });
    Logger.log('[Push→MOS] HTTP ' + resp.getResponseCode() + ' | ' + resp.getContentText().substring(0, 120));
  } catch(e) { Logger.log('[Push→MOS] excepcion: ' + e.message); }
}

// ============================================================
// HORARIO LABORAL + CIERRE NOCTURNO
// LUN-SÁB: 07:00 - 19:00
// DOMINGO: 07:00 - 16:00
// MASTER y ADMINISTRADOR: 24/7 sin restricción
// ============================================================
function _horarioPermitido(rol) {
  var rolUp = String(rol || '').toUpperCase();
  if (rolUp === 'MASTER' || rolUp === 'ADMINISTRADOR') return { permitido: true, motivo: 'rol_admin' };

  var tz   = Session.getScriptTimeZone();
  var ahora = new Date();
  var dia  = parseInt(Utilities.formatDate(ahora, tz, 'u'), 10); // 1=lun, 7=dom
  var hora = parseInt(Utilities.formatDate(ahora, tz, 'H'), 10);
  var min  = parseInt(Utilities.formatDate(ahora, tz, 'm'), 10);
  var horaDecimal = hora + (min / 60);

  var apertura = 7;
  var cierre   = (dia === 7) ? 16 : 19;  // domingo cierra 16h, resto 19h

  if (horaDecimal >= apertura && horaDecimal < cierre) {
    return { permitido: true, apertura: apertura, cierre: cierre, dia: dia };
  }
  return {
    permitido: false,
    apertura:  apertura,
    cierre:    cierre,
    dia:       dia,
    motivo:    horaDecimal < apertura ? 'antes_apertura' : 'despues_cierre'
  };
}

// Endpoint público: el frontend lo llama antes de mostrar la pantalla de login
// y antes de desbloquear sesión, para saber si el operador puede entrar.
function verificarHorario(params) {
  var rol = String(params.rol || '');
  var info = _horarioPermitido(rol);
  return { ok: true, data: info };
}

// Cierra todas las SESIONES ACTIVAS de operadores/envasadores. MASTER/ADMINISTRADOR
// se mantienen activos. Llamado por trigger nocturno.
function forzarCierreSesionesNocturno() {
  var sheet  = getSheet('SESIONES');
  var data   = sheet.getDataRange().getValues();
  var hdrs   = data[0];
  var idxId    = hdrs.indexOf('idSesion');
  var idxIdP   = hdrs.indexOf('idPersonal');
  var idxFFin  = hdrs.indexOf('fechaFin');
  var idxHFin  = hdrs.indexOf('horaFin');
  var idxEst   = hdrs.indexOf('estado');

  var personal = _sheetToObjects(getPersonalSheet());
  var rolMap = {};
  personal.forEach(function(p) { rolMap[String(p.idPersonal)] = String(p.rol || '').toUpperCase(); });

  var tz       = Session.getScriptTimeZone();
  var ahora    = new Date();
  var fechaStr = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
  var horaStr  = Utilities.formatDate(ahora, tz, 'HH:mm:ss');

  var cerradas = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxEst] || '').toUpperCase() !== 'ACTIVA') continue;
    var rol = rolMap[String(data[i][idxIdP])];
    if (rol === 'MASTER' || rol === 'ADMINISTRADOR') continue;  // no tocar admins

    sheet.getRange(i + 1, idxFFin + 1).setValue(fechaStr);
    sheet.getRange(i + 1, idxHFin + 1).setValue(horaStr);
    sheet.getRange(i + 1, idxEst  + 1).setValue('CERRADA_AUTO');
    cerradas++;
  }
  Logger.log('forzarCierreSesionesNocturno: ' + cerradas + ' sesiones cerradas');
  return { ok: true, data: { cerradas: cerradas } };
}
