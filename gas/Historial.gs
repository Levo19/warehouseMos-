// ============================================================
// warehouseMos — Historial.gs
//
// Auditoría completa de una guía: timeline cronológico con todos
// los eventos desde creación hasta último cambio.
//
// Fuentes de eventos:
//   1. GUIAS row              → creación, cambios de estado, foto
//   2. GUIA_DETALLE rows      → items actuales (con flag ANULADO)
//   3. OPS_LOG                → audit trail completo (post v2.1)
//   4. SYNC_LOG               → operaciones legacy (idempotencia)
//   5. STOCK_MOVIMIENTOS      → cambios de stock derivados de la guía
//   6. PRODUCTO_NUEVO         → PN registrados/aprobados en la guía
//   7. MERMAS (idGuiaSalida)  → mermas procesadas que generaron la guía
//
// Gating: solo admin/master según PERSONAL_MASTER.rol o claveAdmin
// válida en params. El admin master ve todo, el operador no.
// ============================================================

function getHistorialGuia(params) {
  var idGuia = String(params.idGuia || '').trim();
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  // Gating: rol del usuario debe ser ADMIN o MASTER, o clave admin válida
  var autorizado = _checkAutorizacionHistorial(params);
  if (!autorizado.ok) return autorizado;

  var eventos = [];

  // 1. GUIA row (creación + estado actual)
  try {
    var guiasSheet = getSheet('GUIAS');
    var guiasData = guiasSheet.getDataRange().getValues();
    var hdrs = guiasData[0];
    var idxId = hdrs.indexOf('idGuia');
    var guia = null;
    for (var i = 1; i < guiasData.length; i++) {
      if (String(guiasData[i][idxId]) === idGuia) {
        guia = {};
        hdrs.forEach(function(h, k) {
          var v = guiasData[i][k];
          guia[h] = (v instanceof Date) ? v.toISOString() : v;
        });
        break;
      }
    }
    if (!guia) return { ok: false, error: 'Guía no encontrada' };

    eventos.push({
      ts:       guia.fecha,
      tipo:     'CREACION',
      icono:    '🆕',
      usuario:  guia.usuario || '',
      deviceId: '',
      titulo:   'Guía creada · ' + (guia.tipo || ''),
      detalle:  {
        tipo:           guia.tipo,
        idProveedor:    guia.idProveedor,
        idZona:         guia.idZona,
        numeroDocumento:guia.numeroDocumento,
        comentario:     guia.comentario,
        idPreingreso:   guia.idPreingreso
      }
    });

    if (String(guia.estado || '').toUpperCase() === 'CERRADA' ||
        String(guia.estado || '').toUpperCase() === 'AUTOCERRADA') {
      eventos.push({
        ts:       guia.fecha,  // aproximado — no tenemos fechaCierre
        tipo:     'CIERRE_ESTADO',
        icono:    '🔒',
        usuario:  guia.usuario || '',
        deviceId: '',
        titulo:   'Estado actual: ' + guia.estado +
                  (guia.montoTotal ? ' · S/. ' + parseFloat(guia.montoTotal).toFixed(2) : ''),
        detalle:  { estado: guia.estado, montoTotal: guia.montoTotal }
      });
    }
    if (guia.foto) {
      eventos.push({
        ts:      guia.fecha,
        tipo:    'FOTO',
        icono:   '📷',
        usuario: guia.usuario || '',
        titulo:  'Foto adjunta',
        detalle: { url: guia.foto }
      });
    }
  } catch(e) {
    return { ok: false, error: 'Error leyendo GUIAS: ' + e.message };
  }

  // 2. GUIA_DETALLE — items actuales (sin timestamp, mostrar al final como "estado actual")
  var detalle = [];
  try {
    var detSheet = getSheet('GUIA_DETALLE');
    var detData = detSheet.getDataRange().getValues();
    var dh = detData[0];
    var iId = dh.indexOf('idGuia');
    for (var j = 1; j < detData.length; j++) {
      if (String(detData[j][iId]) !== idGuia) continue;
      var row = {};
      dh.forEach(function(h, k) {
        var v = detData[j][k];
        row[h] = (v instanceof Date) ? v.toISOString() : v;
      });
      detalle.push(row);
    }
  } catch(e) { /* no-op */ }

  // 3. OPS_LOG — audit trail completo
  try {
    var opsSheet = getSheet('OPS_LOG');
    if (opsSheet) {
      var ops = _sheetToObjects(opsSheet).filter(function(o) {
        return String(o.idGuia || '') === idGuia;
      });
      ops.forEach(function(o) {
        var pl = {};
        try { pl = typeof o.payload === 'string' ? JSON.parse(o.payload) : (o.payload || {}); } catch(_){}
        eventos.push({
          ts:       o.fechaAplicado || o.fechaCreado,
          tipo:     'OP_' + String(o.tipo || ''),
          icono:    _iconoOp(o.tipo),
          usuario:  o.usuario || '',
          deviceId: o.deviceId || '',
          titulo:   _tituloOp(o.tipo, pl),
          estado:   o.estado,
          error:    o.error || '',
          detalle:  { tipo: o.tipo, payload: pl, idOp: o.idOp }
        });
      });
    }
  } catch(e) { /* no-op */ }

  // 4. SYNC_LOG — operaciones legacy (info limitada, solo timestamp + acción)
  try {
    var syncSheet = getSheet('SYNC_LOG');
    if (syncSheet) {
      var syncRows = _sheetToObjects(syncSheet);
      syncRows.forEach(function(s) {
        // Heurística: si el resultado contiene el idGuia, es relevante
        var resStr = String(s.resultado || '');
        if (resStr.indexOf(idGuia) < 0) return;
        eventos.push({
          ts:       s.fecha,
          tipo:     'LEGACY_' + String(s.action || ''),
          icono:    '⚙',
          usuario:  '',  // SYNC_LOG no captura usuario directo
          deviceId: '',
          titulo:   'Operación legacy: ' + s.action,
          detalle:  { localId: s.localId, action: s.action }
        });
      });
    }
  } catch(e) { /* no-op */ }

  // 5. STOCK_MOVIMIENTOS — origen = idGuia
  try {
    var movSheet = getSheet('STOCK_MOVIMIENTOS');
    if (movSheet) {
      var movs = _sheetToObjects(movSheet).filter(function(m) {
        return String(m.origen || '') === idGuia;
      });
      movs.forEach(function(m) {
        eventos.push({
          ts:      m.fecha,
          tipo:    'STOCK_' + String(m.tipoOperacion || ''),
          icono:   parseFloat(m.delta) > 0 ? '📈' : '📉',
          usuario: m.usuario || '',
          titulo:  'Stock ' + (parseFloat(m.delta) > 0 ? '+' : '') + m.delta + ' · ' + (m.codigoProducto || ''),
          detalle: {
            codigoProducto: m.codigoProducto,
            delta:          m.delta,
            stockAntes:     m.stockAntes,
            stockDespues:   m.stockDespues,
            tipoOperacion:  m.tipoOperacion,
            observacion:    m.observacion
          }
        });
      });
    }
  } catch(e) { /* no-op */ }

  // 6. PRODUCTO_NUEVO — registrados/aprobados en la guía
  try {
    var pnSheet = getSheet('PRODUCTO_NUEVO');
    if (pnSheet) {
      var pns = _sheetToObjects(pnSheet).filter(function(p) {
        return String(p.idGuia || '') === idGuia;
      });
      pns.forEach(function(p) {
        eventos.push({
          ts:      p.fechaRegistro || p.fecha,
          tipo:    'PN_REGISTRADO',
          icono:   '🆕',
          usuario: p.usuario || '',
          titulo:  'Producto nuevo registrado: ' + (p.descripcion || p.codigoBarra || ''),
          detalle: { idProductoNuevo: p.idProductoNuevo, codigoBarra: p.codigoBarra, cantidad: p.cantidad, estado: p.estado }
        });
        if (p.fechaAprobacion) {
          eventos.push({
            ts:      p.fechaAprobacion,
            tipo:    'PN_APROBADO',
            icono:   '✓',
            usuario: p.aprobadoPor || '',
            titulo:  'Producto nuevo aprobado: ' + (p.descripcion || ''),
            detalle: { idProductoNuevo: p.idProductoNuevo, aprobadoPor: p.aprobadoPor }
          });
        }
      });
    }
  } catch(e) { /* no-op */ }

  // 7. MERMAS — donde esta guía es la SALIDA generada por procesar mermas
  try {
    var mermaSheet = getSheet('MERMAS');
    if (mermaSheet) {
      var mermas = _sheetToObjects(mermaSheet).filter(function(m) {
        return String(m.idGuiaSalida || m.idGuia || '') === idGuia;
      });
      mermas.forEach(function(m) {
        eventos.push({
          ts:      m.fechaResolucion || m.fechaIngreso,
          tipo:    'MERMA_PROCESADA',
          icono:   '🗑',
          usuario: m.usuario || '',
          titulo:  'Merma procesada: ' + (m.codigoProducto || '') + ' × ' + (m.cantidadDesechada || m.cantidadOriginal || 0),
          detalle: { idMerma: m.idMerma, codigoProducto: m.codigoProducto, motivo: m.motivo }
        });
      });
    }
  } catch(e) { /* no-op */ }

  // Sort cronológicamente
  eventos.sort(function(a, b) {
    var ta = new Date(a.ts || 0).getTime() || 0;
    var tb = new Date(b.ts || 0).getTime() || 0;
    return ta - tb;
  });

  return {
    ok: true,
    data: {
      idGuia:    idGuia,
      eventos:   eventos,
      itemsActuales: detalle,
      generadoEn: new Date().toISOString(),
      totalEventos: eventos.length
    }
  };
}

// ── Helpers ──────────────────────────────────────────────────
function _checkAutorizacionHistorial(params) {
  // Opción 1: clave admin válida
  var clave = String(params.claveAdmin || '').trim();
  if (clave) {
    var adminPin = _getAdminPinAlmacen();
    if (adminPin && clave === adminPin) return { ok: true, via: 'clave' };
  }
  // Opción 2: usuario con rol admin o master
  var idPersonal = String(params.idPersonal || '').trim();
  var usuario    = String(params.usuario || '').trim();
  if (idPersonal || usuario) {
    try {
      var personal = _getPersonalWH();
      var p = personal.find(function(x) {
        return (idPersonal && String(x.idPersonal) === idPersonal) ||
               (usuario && (String(x.nombre || '').toLowerCase() === usuario.toLowerCase() ||
                            String(x.usuario || '').toLowerCase() === usuario.toLowerCase()));
      });
      if (p) {
        var rol = String(p.rol || '').toUpperCase();
        if (rol === 'ADMIN' || rol === 'MASTER') return { ok: true, via: 'rol:' + rol };
      }
    } catch(e) {}
  }
  return { ok: false, error: 'no autorizado · solo admin/master' };
}

function _iconoOp(tipo) {
  var map = {
    SCAN:           '📲',
    EDIT_QTY:       '✏',
    DELETE_ITEM:    '🗑',
    ANULAR_DETALLE: '🚫',
    ANULAR_GUIA:    '❌',
    PN_REGISTRAR:   '🆕',
    MERMA_AGREGAR:  '🗑',
    MERMA_SOLUCIONAR:'♻',
    MERMA_PROCESAR: '⚡',
    CARGADOR_ADD:   '🛺',
    CARGADOR_REMOVE:'➖',
    CREAR_GUIA:     '🆕',
    CERRAR_GUIA:    '🔒',
    REABRIR_GUIA:   '🔓'
  };
  return map[String(tipo || '').toUpperCase()] || '•';
}

function _tituloOp(tipo, payload) {
  var t = String(tipo || '').toUpperCase();
  payload = payload || {};
  switch (t) {
    case 'SCAN':
      return 'Scan: ' + (payload.codigoProducto || '?') +
             ' × ' + (payload.cantidad || payload.cantidadRecibida || 1);
    case 'EDIT_QTY':
      return 'Cantidad cambiada · detalle ' + (payload.idDetalle || '');
    case 'DELETE_ITEM':
    case 'ANULAR_DETALLE':
      return 'Item anulado · detalle ' + (payload.idDetalle || '');
    case 'CREAR_GUIA':
      return 'Guía creada (op-log)';
    case 'CERRAR_GUIA':
      return 'Guía cerrada (op-log)';
    case 'REABRIR_GUIA':
      return 'Guía reabierta (op-log)';
    case 'PN_REGISTRAR':
      return 'Producto nuevo: ' + (payload.descripcion || payload.codigoBarra || '');
    default:
      return t;
  }
}
