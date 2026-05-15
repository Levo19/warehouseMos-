// ============================================================
// warehouseMos — Fotos.gs
// Helpers genéricos Drive para fotos de entidades (preingreso,
// guía, merma, productoNuevo).
//
// Estructura Drive: WH_FOTOS/<YYYY-MM>/<idEntidad>/<file>.jpg
// + thumbnail _thumb_<file>.jpg (200x200)
//
// Las funciones _subirFotoPreingreso, _subirFotoMerma, _subirFotoProductoNuevo
// que ya existen (cada una con su carpeta dedicada) siguen funcionando para
// compatibilidad. Este módulo expone un endpoint UNIFICADO subirFotoEntidad
// que el frontend nuevo usa, escogiendo carpeta según `entidad`.
// ============================================================

var _FOTO_ROOT_FOLDER_NAME = 'WH_FOTOS';

// ── Endpoint público: subir foto de cualquier entidad ──
// params: { entidad: 'preingreso'|'guia'|'merma'|'productoNuevo',
//           idEntidad, base64, mimeType, makeThumb }
function subirFotoEntidad(params) {
  var entidad   = String(params.entidad   || '').toLowerCase();
  var idEntidad = String(params.idEntidad || '').trim();
  var base64    = String(params.base64    || params.fotoBase64 || '').trim();
  var mime      = String(params.mimeType  || 'image/jpeg');
  var makeThumb = params.makeThumb !== false;
  if (!entidad)   return { ok: false, error: 'entidad requerida' };
  if (!idEntidad) return { ok: false, error: 'idEntidad requerido' };
  if (!base64)    return { ok: false, error: 'base64 requerido' };

  try {
    var folder = _getCarpetaEntidad(entidad, idEntidad);
    var nombre = idEntidad + '_' + new Date().getTime() + '.jpg';
    var blob   = Utilities.newBlob(Utilities.base64Decode(base64), mime, nombre);
    var file   = folder.createFile(blob);
    // [v2.11.3] Triple-set robusto en file + folder para evitar fotos negras
    // por permiso parcial.
    if (typeof _setSharingPublicoRobusto === 'function') {
      _setSharingPublicoRobusto(file);
      _setSharingPublicoRobusto(folder);
    } else {
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){}
    }

    var fileId   = file.getId();
    var fullUrl  = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600';
    var thumbUrl = '';
    if (makeThumb) {
      thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200';
    }
    return { ok: true, data: {
      fileId:   fileId,
      url:      fullUrl,
      thumb:    thumbUrl,
      nombre:   nombre
    }};
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Endpoint: eliminar foto por fileId ──
function eliminarFotoEntidad(params) {
  var fileId = String(params.fileId || '').trim();
  if (!fileId) {
    // Soporte: si viene URL, extraer fileId
    var url = String(params.url || '');
    var m = url.match(/[?&]id=([^&]+)/);
    if (m) fileId = m[1];
  }
  if (!fileId) return { ok: false, error: 'fileId requerido' };
  try {
    var file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// F1 — Migración one-shot: detecta base64 inline en celdas y
// las mueve a Drive. Idempotente (skipea si ya es URL).
//
// Ejecutar manualmente desde Apps Script editor:
//   migrarFotosABase64Drive()
// ============================================================
function migrarFotosABase64Drive() {
  var pre = _migrarBase64InlineEnHoja('PREINGRESOS', 'fotos', 'preingreso');
  var gui = _migrarBase64InlineEnHoja('GUIAS', 'foto', 'guia');
  var pn  = _migrarBase64InlineEnHoja('PRODUCTO_NUEVO', 'foto', 'productoNuevo');
  var mrm = _migrarBase64InlineEnHoja('MERMAS', 'foto', 'merma');
  Logger.log('Migración F1:');
  Logger.log('  preingresos: ' + JSON.stringify(pre));
  Logger.log('  guías:       ' + JSON.stringify(gui));
  Logger.log('  productoNvo: ' + JSON.stringify(pn));
  Logger.log('  mermas:      ' + JSON.stringify(mrm));
  return { preingresos: pre, guias: gui, productosNuevos: pn, mermas: mrm };
}

function _migrarBase64InlineEnHoja(nombreHoja, columnaFoto, entidad) {
  var sheet = getSheet(nombreHoja);
  if (!sheet) return { skipped: true, motivo: 'hoja no existe' };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { migrados: 0, total: 0 };

  var hdrs = data[0];
  var idxFoto = hdrs.indexOf(columnaFoto);
  if (idxFoto < 0) return { skipped: true, motivo: 'columna ' + columnaFoto + ' no existe' };

  var idCols = ['idPreingreso','idGuia','idProductoNuevo','idMerma'];
  var idxId = -1;
  idCols.forEach(function(c){ if (idxId < 0) idxId = hdrs.indexOf(c); });
  if (idxId < 0) return { skipped: true, motivo: 'columna id no encontrada' };

  var migrados = 0;
  var fallos = 0;
  for (var i = 1; i < data.length; i++) {
    var val = String(data[i][idxFoto] || '').trim();
    if (!val) continue;
    if (val.indexOf('http') === 0) continue;     // ya es URL
    if (val.indexOf('drive.google') >= 0) continue;
    if (val.length < 100) continue;              // probablemente filename placeholder

    var base64 = val.indexOf('base64,') >= 0 ? val.split('base64,')[1] : val;
    var idEntidad = String(data[i][idxId] || ('row' + i));
    try {
      var res = subirFotoEntidad({
        entidad:   entidad,
        idEntidad: idEntidad,
        base64:    base64,
        mimeType:  'image/jpeg',
        makeThumb: true
      });
      if (res && res.ok && res.data && res.data.url) {
        sheet.getRange(i + 1, idxFoto + 1).setValue(res.data.url);
        migrados++;
      } else {
        fallos++;
      }
    } catch(e) {
      Logger.log('Error migrando row ' + i + ': ' + e.message);
      fallos++;
    }
  }
  return { migrados: migrados, fallos: fallos, total: data.length - 1 };
}

// ── Carpeta por entidad/mes ──
function _getCarpetaEntidad(entidad, idEntidad) {
  var rootProp = 'WH_FOTOS_ROOT_FOLDER_ID';
  var rootId   = PropertiesService.getScriptProperties().getProperty(rootProp);
  var root;
  if (rootId) {
    try { root = DriveApp.getFolderById(rootId); } catch(e) { root = null; }
  }
  if (!root) {
    var ssId   = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    var ssFile = DriveApp.getFileById(ssId);
    var parent = ssFile.getParents().next();
    var it     = parent.getFoldersByName(_FOTO_ROOT_FOLDER_NAME);
    root       = it.hasNext() ? it.next() : parent.createFolder(_FOTO_ROOT_FOLDER_NAME);
    PropertiesService.getScriptProperties().setProperty(rootProp, root.getId());
  }

  // YYYY-MM
  var yyyyMM = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  var mensIt = root.getFoldersByName(yyyyMM);
  var mens   = mensIt.hasNext() ? mensIt.next() : root.createFolder(yyyyMM);

  // entidad/idEntidad
  var entIt  = mens.getFoldersByName(entidad);
  var entF   = entIt.hasNext() ? entIt.next() : mens.createFolder(entidad);
  var idIt   = entF.getFoldersByName(idEntidad);
  var idF    = idIt.hasNext() ? idIt.next() : entF.createFolder(idEntidad);
  return idF;
}
