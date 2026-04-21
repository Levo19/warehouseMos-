// ============================================================
// warehouseMos — Reporte.gs
// Endpoint público para páginas de reporte en tiempo real
// Acción: getReporte?tipo=preingreso|guia&id=xxx
// ============================================================

function getReporte(params) {
  var tipo = String(params.tipo || '');
  var id   = String(params.id   || '');
  if (!tipo || !id) return { ok: false, error: 'tipo e id requeridos' };
  if (tipo === 'preingreso') return _reportePreingreso(id);
  if (tipo === 'guia')       return _reporteGuia(id);
  return { ok: false, error: 'tipo invalido: ' + tipo };
}

function _reportePreingreso(id) {
  var rows = _sheetToObjects(getSheet('PREINGRESOS'));
  var pi   = rows.find(function(r) { return r.idPreingreso === id; });
  if (!pi) return { ok: false, error: 'Preingreso no encontrado: ' + id };

  var provName = '';
  try {
    var provs = _sheetToObjects(getProveedoresSheet());
    var prov  = provs.find(function(p) { return p.idProveedor === pi.idProveedor; });
    if (prov) provName = String(prov.nombre || '');
  } catch(e) {}

  var cargadores = [];
  try { cargadores = JSON.parse(pi.cargadores || '[]'); } catch(e) {}

  var guiaResumen = null;
  if (pi.idGuia) {
    try {
      var guias = _sheetToObjects(getSheet('GUIAS'));
      var g     = guias.find(function(x) { return x.idGuia === pi.idGuia; });
      if (g) {
        var dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
          .filter(function(d) { return d.idGuia === g.idGuia && d.observacion !== 'ANULADO'; });
        guiaResumen = {
          idGuia:  g.idGuia,
          tipo:    g.tipo    || '',
          estado:  g.estado  || '',
          usuario: g.usuario || '',
          items:   dets.length
        };
      }
    } catch(e) {}
  }

  return {
    ok: true,
    data: {
      tipo:         'preingreso',
      idPreingreso: pi.idPreingreso,
      fecha:        String(pi.fecha || ''),
      estado:       pi.estado    || '',
      idProveedor:  pi.idProveedor || '',
      proveedor:    provName,
      monto:        pi.monto     || '',
      comentario:   pi.comentario || '',
      fotos:        pi.fotos ? String(pi.fotos).split(',').filter(Boolean) : [],
      cargadores:   cargadores,
      usuario:      pi.usuario   || '',
      idGuia:       pi.idGuia    || '',
      guia:         guiaResumen,
      generado:     new Date().toISOString()
    }
  };
}

function _reporteGuia(id) {
  var guias = _sheetToObjects(getSheet('GUIAS'));
  var g     = guias.find(function(x) { return x.idGuia === id; });
  if (!g) return { ok: false, error: 'Guia no encontrada: ' + id };

  var provName = '';
  try {
    var provs = _sheetToObjects(getProveedoresSheet());
    var prov  = provs.find(function(p) { return p.idProveedor === g.idProveedor; });
    if (prov) provName = String(prov.nombre || '');
  } catch(e) {}

  var dets = [];
  try {
    var prods   = _sheetToObjects(getProductosSheet());
    var prodMap = {};
    prods.forEach(function(p) { prodMap[p.idProducto] = p.descripcion || p.idProducto; });
    dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
      .filter(function(d) { return d.idGuia === id && d.observacion !== 'ANULADO'; })
      .map(function(d) {
        return {
          codigoProducto:   d.codigoProducto,
          descripcion:      prodMap[d.codigoProducto] || d.codigoProducto,
          cantidadEsperada: d.cantidadEsperada || 0,
          cantidadReal:     d.cantidadReal || d.cantidadRecibida || 0,
          fechaVencimiento: String(d.fechaVencimiento || ''),
          observacion:      d.observacion || ''
        };
      });
  } catch(e) {}

  var preResumen = null;
  if (g.idPreingreso) {
    try {
      var pis = _sheetToObjects(getSheet('PREINGRESOS'));
      var pi  = pis.find(function(x) { return x.idPreingreso === g.idPreingreso; });
      if (pi) {
        var nFotos = pi.fotos ? String(pi.fotos).split(',').filter(Boolean).length : 0;
        preResumen = {
          idPreingreso: pi.idPreingreso,
          estado:       pi.estado || '',
          monto:        pi.monto  || '',
          nFotos:       nFotos
        };
      }
    } catch(e) {}
  }

  return {
    ok: true,
    data: {
      tipo:         'guia',
      idGuia:       g.idGuia,
      tipoGuia:     g.tipo       || '',
      estado:       g.estado     || '',
      fecha:        String(g.fecha || ''),
      idProveedor:  g.idProveedor || '',
      proveedor:    provName,
      usuario:      g.usuario    || '',
      comentario:   g.comentario || '',
      foto:         g.foto       || '',
      idPreingreso: g.idPreingreso || '',
      preingreso:   preResumen,
      detalle:      dets,
      generado:     new Date().toISOString()
    }
  };
}
