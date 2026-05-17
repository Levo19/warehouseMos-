// ============================================================
// warehouseMos — ListasSombra.gs
// Listas sombra COMPARTIDAS entre operadores. Modelo "tomable" tipo pickup:
// - Operador A crea (estado DISPONIBLE) → cualquier operador la ve
// - B toma (estado EN_USO) → queda bloqueada con su nombre
// - B cierra al despachar (estado COMPLETADA) o libera (vuelve a DISPONIBLE)
// ============================================================

var LS_SHEET_NAME = 'LISTAS_SOMBRA';
var LS_HEADERS = ['idLista','fechaCreacion','usuarioCreador','items','estado',
                  'usuarioTomada','fechaTomada','fechaCompletada','nota'];

// Auto-create columnas faltantes (idempotente). Crea la hoja si no existe.
function _ensureListasSombraSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(LS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LS_SHEET_NAME);
    sh.appendRow(LS_HEADERS);
    // [v2.13.15] columna items como texto para evitar truncamiento de JSON
    var colItems = LS_HEADERS.indexOf('items') + 1;
    if (colItems > 0) sh.getRange(1, colItems, sh.getMaxRows()).setNumberFormat('@');
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var faltantes = LS_HEADERS.filter(function(c){ return existing.indexOf(c) < 0; });
  if (faltantes.length) {
    sh.getRange(1, lastCol + 1, 1, faltantes.length).setValues([faltantes]);
  }
  return sh;
}

// ── crearListaSombra: el operador acaba de subir una lista vía IA ─────────
// Por defecto, queda EN_USO del creador (lo más natural: el que la sube
// generalmente la trabaja). Si quiere dejarla DISPONIBLE para que otro la
// tome, pasa params.compartir=true.
function crearListaSombra(params) {
  return _conLock('crearListaSombra', function() {
    var sh = _ensureListasSombraSheet();
    var usuario = String(params.usuario || '').trim();
    if (!usuario) return { ok: false, error: 'usuario requerido' };
    var items = params.items;
    if (typeof items === 'string') {
      try { items = JSON.parse(items); } catch(e) { return { ok: false, error: 'items JSON inválido' }; }
    }
    if (!Array.isArray(items) || !items.length) return { ok: false, error: 'sin items' };

    var idLista = String(params.idLista || ('LS' + Date.now()));
    var compartir = !!params.compartir;
    var ahora = new Date();

    // Idempotencia: si ya existe, devolver tal cual
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idLista) {
        return { ok: true, data: { idLista: idLista, duplicado: true } };
      }
    }

    var row = new Array(hdrs.length).fill('');
    function _set(col, val) { var k = hdrs.indexOf(col); if (k >= 0) row[k] = val; }
    _set('idLista',          idLista);
    _set('fechaCreacion',    ahora);
    _set('usuarioCreador',   usuario);
    _set('items',            JSON.stringify(items));
    _set('estado',           compartir ? 'DISPONIBLE' : 'EN_USO');
    _set('usuarioTomada',    compartir ? '' : usuario);
    _set('fechaTomada',      compartir ? '' : ahora);
    _set('fechaCompletada',  '');
    _set('nota',             String(params.nota || ''));
    sh.appendRow(row);

    // Push a otros operadores si quedó DISPONIBLE
    if (compartir) {
      try {
        _notificarMOS(
          '📋 Nueva lista sombra disponible',
          usuario + ' subió una lista de ' + items.length + ' productos para despachar',
          usuario,
          'WH_LISTA_SOMBRA_NUEVA'
        );
      } catch(eP) {}
    }

    return { ok: true, data: { idLista: idLista, estado: compartir ? 'DISPONIBLE' : 'EN_USO' } };
  });
}

// ── getListasSombra: devuelve activas (DISPONIBLE + EN_USO) + opcionalmente
// las completadas del día (para historial corto)
function getListasSombra(params) {
  var sh = _ensureListasSombraSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, data: { listas: [] } };
  var hdrs = data[0];
  var idx = {};
  hdrs.forEach(function(h, k) { idx[String(h).trim()] = k; });
  var hoy = new Date();
  var hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  var incluirCompletadas = !!params.incluirCompletadas;

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var estado = String(r[idx.estado] || '').toUpperCase();
    if (estado !== 'DISPONIBLE' && estado !== 'EN_USO') {
      if (!incluirCompletadas) continue;
      // Solo completadas de hoy
      var fComp = r[idx.fechaCompletada];
      if (!(fComp instanceof Date) || fComp < hoy0) continue;
    }
    var items = [];
    try { items = JSON.parse(r[idx.items] || '[]'); } catch(e) {}
    var total = items.length;
    var completos = items.filter(function(it) {
      return (parseFloat(it.cantidadEscaneada) || 0) >= (parseFloat(it.cantidad) || 0);
    }).length;
    rows.push({
      idLista:         String(r[idx.idLista] || ''),
      fechaCreacion:   r[idx.fechaCreacion] instanceof Date ? r[idx.fechaCreacion].toISOString() : String(r[idx.fechaCreacion] || ''),
      usuarioCreador:  String(r[idx.usuarioCreador] || ''),
      estado:          estado,
      usuarioTomada:   String(r[idx.usuarioTomada] || ''),
      fechaTomada:     r[idx.fechaTomada] instanceof Date ? r[idx.fechaTomada].toISOString() : String(r[idx.fechaTomada] || ''),
      fechaCompletada: r[idx.fechaCompletada] instanceof Date ? r[idx.fechaCompletada].toISOString() : String(r[idx.fechaCompletada] || ''),
      nota:            String(r[idx.nota] || ''),
      total:           total,
      completos:       completos,
      items:           items  // full items para no requerir segunda llamada al tomarla
    });
  }
  // Más recientes primero
  rows.sort(function(a, b) {
    return String(b.fechaCreacion).localeCompare(String(a.fechaCreacion));
  });
  return { ok: true, data: { listas: rows } };
}

// ── tomarListaSombra: B toma una DISPONIBLE (o le quita a otro si la lista
// está EN_USO pero el usuario tiene rol admin — esto no se valida acá, el
// frontend decide cuándo permitir robar)
function tomarListaSombra(params) {
  return _conLock('tomarListaSombra', function() {
    var sh = _ensureListasSombraSheet();
    var idLista = String(params.idLista || '').trim();
    var usuario = String(params.usuario || '').trim();
    if (!idLista || !usuario) return { ok: false, error: 'idLista y usuario requeridos' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idx = {};
    hdrs.forEach(function(h, k) { idx[String(h).trim()] = k; });
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== idLista) continue;
      var estado = String(data[i][idx.estado] || '').toUpperCase();
      if (estado === 'COMPLETADA') return { ok: false, error: 'YA_COMPLETADA' };
      var dueno = String(data[i][idx.usuarioTomada] || '').trim();
      if (estado === 'EN_USO' && dueno && dueno !== usuario && !params.forzar) {
        return { ok: false, error: 'EN_USO_POR_OTRO', mensaje: 'Tomada por: ' + dueno };
      }
      sh.getRange(i + 1, idx.estado + 1).setValue('EN_USO');
      sh.getRange(i + 1, idx.usuarioTomada + 1).setValue(usuario);
      sh.getRange(i + 1, idx.fechaTomada + 1).setValue(new Date());
      var items = [];
      try { items = JSON.parse(data[i][idx.items] || '[]'); } catch(e) {}
      return { ok: true, data: { idLista: idLista, items: items, dueno: usuario } };
    }
    return { ok: false, error: 'NO_ENCONTRADA' };
  });
}

// ── liberarListaSombra: el operador la suelta sin completarla (alguien más
// la puede tomar)
function liberarListaSombra(params) {
  return _conLock('liberarListaSombra', function() {
    var sh = _ensureListasSombraSheet();
    var idLista = String(params.idLista || '').trim();
    var usuario = String(params.usuario || '').trim();
    if (!idLista) return { ok: false, error: 'idLista requerido' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idx = {};
    hdrs.forEach(function(h, k) { idx[String(h).trim()] = k; });
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== idLista) continue;
      var dueno = String(data[i][idx.usuarioTomada] || '').trim();
      if (usuario && dueno !== usuario && !params.forzar) {
        return { ok: false, error: 'NO_ES_TUYA', mensaje: 'Está tomada por: ' + dueno };
      }
      sh.getRange(i + 1, idx.estado + 1).setValue('DISPONIBLE');
      sh.getRange(i + 1, idx.usuarioTomada + 1).setValue('');
      sh.getRange(i + 1, idx.fechaTomada + 1).setValue('');
      return { ok: true };
    }
    return { ok: false, error: 'NO_ENCONTRADA' };
  });
}

// ── actualizarProgresoListaSombra: sync de cantidades escaneadas
// (el operador sigue trabajando localmente pero cada ~15s pushea al backend
// para que los otros vean el progreso)
function actualizarProgresoListaSombra(params) {
  return _conLock('actualizarProgresoListaSombra', function() {
    var sh = _ensureListasSombraSheet();
    var idLista = String(params.idLista || '').trim();
    var items = params.items;
    if (typeof items === 'string') {
      try { items = JSON.parse(items); } catch(e) { return { ok: false, error: 'items JSON inválido' }; }
    }
    if (!Array.isArray(items)) return { ok: false, error: 'items debe ser array' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idx = {};
    hdrs.forEach(function(h, k) { idx[String(h).trim()] = k; });
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== idLista) continue;
      sh.getRange(i + 1, idx.items + 1).setValue(JSON.stringify(items));
      return { ok: true };
    }
    return { ok: false, error: 'NO_ENCONTRADA' };
  });
}

// ── cerrarListaSombra: marca COMPLETADA (se llamó al confirmar el despacho)
function cerrarListaSombra(params) {
  return _conLock('cerrarListaSombra', function() {
    var sh = _ensureListasSombraSheet();
    var idLista = String(params.idLista || '').trim();
    if (!idLista) return { ok: false, error: 'idLista requerido' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var idx = {};
    hdrs.forEach(function(h, k) { idx[String(h).trim()] = k; });
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== idLista) continue;
      // Si vienen items finales, persistirlos antes de cerrar
      if (params.items) {
        var itemsFinal = params.items;
        if (typeof itemsFinal === 'string') {
          try { itemsFinal = JSON.parse(itemsFinal); } catch(e) {}
        }
        if (Array.isArray(itemsFinal)) {
          sh.getRange(i + 1, idx.items + 1).setValue(JSON.stringify(itemsFinal));
        }
      }
      sh.getRange(i + 1, idx.estado + 1).setValue('COMPLETADA');
      sh.getRange(i + 1, idx.fechaCompletada + 1).setValue(new Date());
      return { ok: true };
    }
    return { ok: false, error: 'NO_ENCONTRADA' };
  });
}
