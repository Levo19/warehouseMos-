// ============================================================
// warehouseMos — Reporte.gs
// Endpoint público para páginas de reporte en tiempo real
// Acción: getReporte?tipo=preingreso|guia&id=xxx
// ============================================================

// ============================================================
// imprimirTicketGuia — ESC/POS 80mm con QR al final
// ============================================================
function imprimirTicketGuia(params) {
  var idGuia = String(params.idGuia || '');
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  var printerId;
  try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
  catch(e) { return { ok: false, error: e.message }; }

  // ── Datos ────────────────────────────────────────────────────
  var guias = _sheetToObjects(getSheet('GUIAS'));
  var g     = guias.find(function(x) { return x.idGuia === idGuia; });
  if (!g) return { ok: false, error: 'Guía no encontrada: ' + idGuia };

  var TIPO_LABELS = {
    INGRESO_PROVEEDOR:'Ingreso Proveedor', INGRESO_JEFATURA:'Ingreso Jefatura',
    SALIDA_ZONA:'Salida Zona', SALIDA_DEVOLUCION:'Devolucion',
    SALIDA_JEFATURA:'Salida Jefatura', SALIDA_ENVASADO:'Envasado', SALIDA_MERMA:'Merma'
  };

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

    // PN map: codigoBarra → { descripcion, estado }
    var pns = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
    var pnMap = {};
    pns.forEach(function(pn) {
      var cod = String(pn.codigoBarra || '');
      if (cod) pnMap[cod] = { desc: String(pn.descripcion || pn.marca || cod), estado: String(pn.estado || '') };
    });

    dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
      .filter(function(d) { return d.idGuia === idGuia && d.observacion !== 'ANULADO'; })
      .map(function(d) {
        var cod   = String(d.codigoProducto || '');
        var esPN  = !!(pnMap[cod]) || cod.indexOf('NLEV') === 0;
        var desc  = esPN
          ? (pnMap[cod] ? pnMap[cod].desc : cod)
          : (prodMap[cod] || cod);
        return {
          descripcion:     desc,
          cantidad:        parseFloat(d.cantidadReal || d.cantidadRecibida || d.cantidadEsperada || 0),
          esProductoNuevo: esPN,
          estadoPN:        esPN && pnMap[cod] ? pnMap[cod].estado : ''
        };
      });
  } catch(e) {}

  var tz     = Session.getScriptTimeZone();
  var fecha  = Utilities.formatDate(new Date(g.fecha || new Date()), tz, 'dd MMM yyyy');
  var hora   = Utilities.formatDate(new Date(), tz, 'HH:mm');
  var tipoLabel = TIPO_LABELS[g.tipo] || String(g.tipo || '—');

  // ── URL del reporte (para QR) ────────────────────────────────
  var reporteUrl = String(params.reporteUrl || '');

  // ── ESC/POS byte array ───────────────────────────────────────
  var B = [];
  function b1(v) { B.push(v & 0xff); }
  function bStr(s) { for (var i = 0; i < s.length; i++) B.push(s.charCodeAt(i) & 0xff); }
  function bLn(s) { bStr(s); b1(0x0a); }

  // 48 chars: tag(2) + nombre(38) + cant(rest) → con etiqueta N o I
  function lineaDet(tag, nombre, cant) {
    var pre = (tag ? tag.substring(0,1) : ' ') + ' ';   // ej: "N " o "I "
    var n   = String(nombre).substring(0, 38);
    var c   = String(cant);
    var pad = 48 - pre.length - n.length - c.length;
    if (pad < 1) pad = 1;
    return pre + n + Array(pad + 1).join(' ') + c;
  }
  // Retro-compat para header (sin tag)
  function lineaProd(nombre, cant) { return lineaDet(' ', nombre, cant); }

  // Línea etiqueta: label fijo 10 chars + valor
  function lineaKV(label, valor) {
    var l = String(label);
    var v = String(valor).substring(0, 48 - l.length);
    return l + v;
  }

  var SEP  = '================================================';
  var SEP2 = '------------------------------------------------';

  // Init
  b1(0x1b); b1(0x40);

  // Header centrado
  b1(0x1b); b1(0x61); b1(0x01);
  b1(0x1b); b1(0x21); b1(0x08);  // bold
  bLn('ALMACEN CENTRAL - MOS');
  b1(0x1b); b1(0x21); b1(0x00);
  b1(0x1b); b1(0x61); b1(0x00);  // left

  bLn(SEP);

  bLn(lineaKV('GUIA    : ', idGuia));
  bLn(lineaKV('TIPO    : ', tipoLabel));
  bLn(lineaKV('FECHA   : ', fecha + '  ' + hora));
  bLn(lineaKV('ESTADO  : ', g.estado || '—'));

  bLn(SEP2);

  if (provName) bLn(lineaKV('PROVEEDOR: ', provName));
  bLn(lineaKV('USUARIO  : ', g.usuario || '—'));
  if (g.comentario) {
    bLn(SEP2);
    bLn(lineaKV('NOTA     : ', String(g.comentario).substring(0, 37)));
  }

  bLn(SEP);

  // Header detalle
  b1(0x1b); b1(0x21); b1(0x08);  // bold
  bLn(lineaDet(' ', 'PRODUCTO', 'CANT'));
  b1(0x1b); b1(0x21); b1(0x00);
  bLn(SEP2);

  // Items
  dets.forEach(function(d) {
    var tag    = d.esProductoNuevo ? 'N' : 'I';
    var nombre = String(d.descripcion || '').toUpperCase();
    var cant   = d.cantidad % 1 === 0 ? String(Math.round(d.cantidad)) : String(d.cantidad);
    if (nombre.length <= 38) {
      bLn(lineaDet(tag, nombre, cant));
    } else {
      bLn(lineaDet(tag, nombre.substring(0, 38), cant));
      bLn('  ' + nombre.substring(38, 78));
    }
  });

  if (!dets.length) bLn('  (sin items registrados)');

  bLn(SEP);

  // QR Code
  if (reporteUrl) {
    b1(0x1b); b1(0x61); b1(0x01);  // centrar

    var qrData = reporteUrl;
    var qrLen  = qrData.length + 3;
    var qrpL   = qrLen & 0xff;
    var qrpH   = (qrLen >> 8) & 0xff;

    // Modelo 2
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x04); b1(0x00); b1(0x31); b1(0x41); b1(0x32); b1(0x00);
    // Tamaño módulo 4
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x43); b1(0x04);
    // Corrección de errores M
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x45); b1(0x31);
    // Almacenar datos
    b1(0x1d); b1(0x28); b1(0x6b); b1(qrpL); b1(qrpH); b1(0x31); b1(0x50); b1(0x30);
    bStr(qrData);
    // Imprimir
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x51); b1(0x30);

    b1(0x1b); b1(0x21); b1(0x00);
    bLn('Escanea para ver el reporte');
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);
  }

  // Feed + corte
  b1(0x1b); b1(0x4a); b1(160);
  b1(0x1d); b1(0x56); b1(0x00);

  // Encode
  var blob = Utilities.newBlob(B, 'application/octet-stream');
  var b64  = Utilities.base64Encode(blob.getBytes());

  var payload = {
    printerId:   parseInt(printerId),
    title:       'Ticket ' + idGuia,
    contentType: 'raw_base64',
    content:     b64,
    source:      'warehouseMos'
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:  'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    return code === 201
      ? { ok: true }
      : { ok: false, error: 'PrintNode ' + code + ': ' + resp.getContentText() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

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

    var pns = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
    var pnMap = {};
    pns.forEach(function(pn) {
      var cod = String(pn.codigoBarra || '');
      if (cod) pnMap[cod] = { desc: String(pn.descripcion || pn.marca || cod), estado: String(pn.estado || '') };
    });

    dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
      .filter(function(d) { return d.idGuia === id && d.observacion !== 'ANULADO'; })
      .map(function(d) {
        var cod  = String(d.codigoProducto || '');
        var esPN = !!(pnMap[cod]) || cod.indexOf('NLEV') === 0;
        var desc = esPN
          ? (pnMap[cod] ? pnMap[cod].desc : cod)
          : (prodMap[cod] || cod);
        return {
          codigoProducto:   cod,
          descripcion:      desc,
          cantidadEsperada: d.cantidadEsperada || 0,
          cantidadReal:     d.cantidadReal || d.cantidadRecibida || 0,
          fechaVencimiento: String(d.fechaVencimiento || ''),
          observacion:      d.observacion || '',
          esProductoNuevo:  esPN,
          estadoPN:         esPN && pnMap[cod] ? pnMap[cod].estado : ''
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
