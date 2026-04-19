// ============================================================
// warehouseMos — Personal.gs
// Login por PIN, sesiones, desempeño y cierre de turno
// ============================================================

// ── Login ───────────────────────────────────────────────────
function loginPersonal(params) {
  var pin = String(params.pin || '').trim();
  if (!pin) return { ok: false, error: 'PIN requerido' };

  // Solo operadores de warehouseMos activos (appOrigen=warehouseMos, estado=1)
  var personal = _getPersonalWH();
  var operador = personal.find(function(p) {
    return String(p.pin) === pin;
  });
  if (!operador) return { ok: false, error: 'PIN incorrecto' };

  // Cerrar sesiones anteriores y capturar si había una de otro día
  var huerfanas = _cerrarSesionesHuerfanas(operador.idPersonal);

  // Crear sesión
  var idSesion  = _generateId('SES');
  var ahora     = new Date();
  var horaStr   = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm:ss');
  var fechaStr  = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  getSheet('SESIONES').appendRow([
    idSesion,
    operador.idPersonal,
    fechaStr,
    horaStr,
    '', '', 0, 'ACTIVA'
  ]);

  // Crear fila vacía en DESEMPENO para esta sesión
  var idDesempeno = _generateId('DES');
  getSheet('DESEMPENO').appendRow([
    idDesempeno, operador.idPersonal, idSesion, fechaStr,
    0, 0,       // minutos, horas
    0, 0,       // guiasCreadas, guiasCerradas
    0, 0,       // envasadosRegistrados, unidadesEnvasadas
    0, 0,       // mermasRegistradas, auditoriaEjecutadas
    0, 0,       // preingresoCreados, ajustesRealizados
    0, 0,       // totalActividades, actividadesPorHora
    0, '',      // puntuacion, calificacion
    parseFloat(operador.montoBase) || 0,
    0, 0,       // bonus, total
    'ACTIVA'
  ]);

  return {
    ok: true,
    data: {
      idSesion:       idSesion,
      idPersonal:     operador.idPersonal,
      nombre:         operador.nombre,
      apellido:       operador.apellido,
      rol:            operador.rol,
      color:          operador.color,
      horaInicio:     horaStr,
      sesionAnterior: huerfanas.ultimaFecha || null   // fecha si había sesión de otro día
    }
  };
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

  // Calcular minutos trabajados
  var inicio = new Date(fechaInicio + 'T' + horaInicio);
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
