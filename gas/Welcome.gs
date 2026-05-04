// ============================================================
// warehouseMos — Welcome.gs
// Datos para la pantalla de bienvenida post-login:
// stats personales (racha, sesiones), ayer vs hoy, pendientes.
// ============================================================

function getWelcomeData(params) {
  var idPersonal = String(params.idPersonal || '').trim();
  if (!idPersonal) return { ok: false, error: 'idPersonal requerido' };

  var tz       = Session.getScriptTimeZone();
  var ahora    = new Date();
  var hoy      = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
  var ayer     = new Date(ahora); ayer.setDate(ayer.getDate() - 1);
  var ayerStr  = Utilities.formatDate(ayer, tz, 'yyyy-MM-dd');

  var data = {
    racha:           0,
    totalSesiones:   0,
    horaInicioHoy:   '',
    ayer: {
      guiasCerradas:    0,
      unidadesEnvasadas:0,
      mermasResueltas:  0
    },
    pendientes: {
      mermasVencidas:  0,
      envasesUrgentes: 0,
      auditoriasHoy:   0
    },
    sinResolverAyer: 0
  };

  // ── Stats de sesiones personales ────────────────────
  try {
    var sesiones = _sheetToObjects(getSheet('SESIONES'))
      .filter(function(s) { return String(s.idPersonal) === idPersonal; });
    data.totalSesiones = sesiones.length;

    // Hora de inicio de la sesión activa de hoy (si existe)
    var sesHoy = sesiones.find(function(s) {
      return String(s.fechaInicio || '').substring(0, 10) === hoy && s.estado === 'ACTIVA';
    });
    if (sesHoy) data.horaInicioHoy = String(sesHoy.horaInicio || '');

    // Racha: contar días consecutivos hasta hoy con al menos 1 sesión
    var fechasSet = {};
    sesiones.forEach(function(s) {
      var f = String(s.fechaInicio || '').substring(0, 10);
      if (f) fechasSet[f] = true;
    });
    var racha = 0;
    var d = new Date(ahora);
    while (true) {
      var fStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (fechasSet[fStr]) {
        racha++;
        d.setDate(d.getDate() - 1);
      } else {
        break;
      }
    }
    data.racha = racha;
  } catch(e) { Logger.log('welcome sesiones: ' + e.message); }

  // ── Ayer: guías cerradas por este operador ────────────
  try {
    var personal = _sheetToObjects(getPersonalSheet());
    var op = personal.find(function(p) { return String(p.idPersonal) === idPersonal; });
    var nombreCompleto = op ? (op.nombre + ' ' + (op.apellido || '')).trim() : '';

    if (nombreCompleto) {
      var guias = _sheetToObjects(getSheet('GUIAS'));
      data.ayer.guiasCerradas = guias.filter(function(g) {
        var fGuia = g.fecha ? Utilities.formatDate(new Date(g.fecha), tz, 'yyyy-MM-dd') : '';
        return fGuia === ayerStr
            && String(g.estado || '').toUpperCase() === 'CERRADA'
            && String(g.usuario || '').trim() === nombreCompleto;
      }).length;

      // Envasados de ayer (suma de unidadesProducidas)
      try {
        var envasados = _sheetToObjects(getSheet('ENVASADOS'));
        data.ayer.unidadesEnvasadas = envasados
          .filter(function(e) {
            var f = e.fecha ? Utilities.formatDate(new Date(e.fecha), tz, 'yyyy-MM-dd') : '';
            return f === ayerStr && String(e.usuario || '').trim() === nombreCompleto;
          })
          .reduce(function(s, e) { return s + (parseFloat(e.unidadesProducidas) || 0); }, 0);
      } catch(eE) {}

      // Mermas resueltas ayer
      try {
        var mermas = _sheetToObjects(getSheet('MERMAS'));
        data.ayer.mermasResueltas = mermas.filter(function(m) {
          if (String(m.estado || '').toUpperCase() !== 'RESUELTA' && String(m.estado || '').toUpperCase() !== 'DESECHADA') return false;
          var fr = m.fechaResolucion ? Utilities.formatDate(new Date(m.fechaResolucion), tz, 'yyyy-MM-dd') : '';
          return fr === ayerStr;
        }).length;
      } catch(eM) {}
    }
  } catch(e) { Logger.log('welcome ayer: ' + e.message); }

  // ── Pendientes globales ──────────────────────────────
  try {
    // Mermas vencidas (>3 días sin resolver)
    var mermas = _sheetToObjects(getSheet('MERMAS'));
    var TRES_DIAS = 3 * 24 * 60 * 60 * 1000;
    var nowMs = ahora.getTime();
    data.pendientes.mermasVencidas = mermas.filter(function(m) {
      if (String(m.estado || '').toUpperCase() !== 'EN_PROCESO') return false;
      var f = m.fechaIngreso ? new Date(m.fechaIngreso).getTime() : nowMs;
      return (nowMs - f) > TRES_DIAS;
    }).length;
    data.sinResolverAyer = mermas.filter(function(m) {
      return String(m.estado || '').toUpperCase() === 'EN_PROCESO';
    }).length;
  } catch(e) { Logger.log('welcome pendientes: ' + e.message); }

  try {
    // Auditorías asignadas para hoy
    var auds = _sheetToObjects(getSheet('AUDITORIAS'));
    data.pendientes.auditoriasHoy = auds.filter(function(a) {
      if (String(a.estado || '').toUpperCase() !== 'ASIGNADA') return false;
      var f = a.fechaAsignacion ? Utilities.formatDate(new Date(a.fechaAsignacion), tz, 'yyyy-MM-dd') : '';
      return f === hoy;
    }).length;
  } catch(e) {}

  try {
    // Envases urgentes — productos derivados con stock <= 0 o < min
    var prods = _sheetToObjects(getProductosSheet());
    var stock = _sheetToObjects(getSheet('STOCK'));
    var stockMap = {};
    stock.forEach(function(s) { stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });
    var envUrgentes = 0;
    prods.forEach(function(p) {
      if (!p.codigoProductoBase || p.estado !== '1') return;
      var esWH = String(p.codigoBarra || p.idProducto || '').toUpperCase().indexOf('WH') === 0;
      if (!esWH) return;
      var st = stockMap[p.codigoBarra] !== undefined ? stockMap[p.codigoBarra] : (stockMap[p.idProducto] || 0);
      var mn = parseFloat(p.stockMinimo) || 0;
      if (st <= 0 || (mn > 0 && st < mn / 2)) envUrgentes++;  // críticos: bajo la mitad del min
    });
    data.pendientes.envasesUrgentes = envUrgentes;
  } catch(e) {}

  return { ok: true, data: data };
}
