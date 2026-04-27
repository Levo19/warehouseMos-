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
    prods.forEach(function(p) {
      var desc = p.descripcion || p.idProducto;
      if (!desc) return;
      prodMap[p.idProducto] = desc;
      if (p.codigoBarra) prodMap[String(p.codigoBarra).trim()] = desc;
      // skuBase solo para producto BASE (factor=1) — evita que presentaciones sobreescriban
      var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
      if (esBase && p.skuBase) prodMap[String(p.skuBase).trim().toUpperCase()] = desc;
    });
    // Equivalentes → resuelven al producto base via skuBase
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        _sheetToObjects(equivSheet).forEach(function(e) {
          if (!e.codigoBarra || !e.skuBase) return;
          var skuKey = String(e.skuBase).trim().toUpperCase();
          var desc   = prodMap[skuKey];
          if (desc) prodMap[String(e.codigoBarra).trim()] = desc;
        });
      }
    } catch(e) {}

    var pns = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
    var pnMap = {};
    pns.forEach(function(pn) {
      var cod = String(pn.codigoBarra || '');
      if (cod) pnMap[cod] = { desc: String(pn.descripcion || pn.marca || cod), estado: String(pn.estado || '') };
    });

    dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
      .filter(function(d) { return d.idGuia === idGuia && d.observacion !== 'ANULADO'; })
      .map(function(d) {
        var cod        = String(d.codigoProducto || '');
        var esPN       = !!(pnMap[cod]) || cod.indexOf('NLEV') === 0;
        var enCatalogo = !!prodMap[cod];
        var desc       = esPN ? (pnMap[cod] ? pnMap[cod].desc : cod)
                       : enCatalogo ? prodMap[cod] : cod;
        return {
          codigoProducto:  cod,
          descripcion:     desc,
          cantidad:        parseFloat(d.cantidadReal || d.cantidadRecibida || d.cantidadEsperada || 0),
          fechaVencimiento: String(d.fechaVencimiento || '').split('T')[0],
          esProductoNuevo: esPN,
          esIncompleto:    !esPN && !enCatalogo,
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

  // 48 chars: tag(2) + nombre(38) + cant(rest)
  // tag: 'n'=nuevo, 'i'=incompleto, ' '=exacto (sin prefijo visual)
  function lineaDet(tag, nombre, cant) {
    var pre = (tag && tag !== ' ' ? tag.substring(0,1) : ' ') + ' ';
    var n   = String(nombre).substring(0, 38);
    var c   = String(cant);
    var pad = 48 - pre.length - n.length - c.length;
    if (pad < 1) pad = 1;
    return pre + n + Array(pad + 1).join(' ') + c;
  }
  function lineaProd(nombre, cant) { return lineaDet(' ', nombre, cant); }

  // Formato simple de fecha "YYYY-MM-DD" → "15 ago 2027"
  function fmtVenc(raw) {
    if (!raw) return '';
    var parts = String(raw).split('-');
    if (parts.length !== 3) return raw;
    var meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    var m = parseInt(parts[1], 10) - 1;
    return parts[2] + ' ' + (meses[m] || parts[1]) + ' ' + parts[0];
  }

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

  // ── HEADER: WAREHOUSE / MOS ─────────────────────────────────
  // Centrado + doble alto + doble ancho + bold
  b1(0x1b); b1(0x61); b1(0x01);
  b1(0x1b); b1(0x21); b1(0x38);
  bLn('WAREHOUSE');
  bLn('MOS');
  b1(0x1b); b1(0x21); b1(0x00);
  b1(0x1b); b1(0x61); b1(0x00);

  bLn(SEP);

  // ── CUADRO 1: info compacta de la guía ──────────────────────
  // Quitar prefijo "SALIDA" / "INGRESO" del tipo, dejar solo el destino
  var tipoCorto = tipoLabel.toUpperCase().replace(/^SALIDA\s+/, '').replace(/^INGRESO\s+/, '');

  // Línea 1: tipo (centrado + bold + doble-alto)
  b1(0x1b); b1(0x61); b1(0x01);  // center
  b1(0x1b); b1(0x21); b1(0x10);  // double-height
  b1(0x1b); b1(0x45); b1(0x01);  // bold
  bLn(tipoCorto);
  b1(0x1b); b1(0x45); b1(0x00);
  b1(0x1b); b1(0x21); b1(0x00);

  // Línea 2: fecha + hora juntos (centrado, bold)
  b1(0x1b); b1(0x45); b1(0x01);
  bLn(fecha.toUpperCase() + '  ' + hora);
  b1(0x1b); b1(0x45); b1(0x00);

  // Línea 3: estado (centrado)
  bLn((g.estado || '—').toUpperCase());

  b1(0x1b); b1(0x61); b1(0x00);  // left

  // Nota: si hay comentario, word-wrap inteligente
  if (g.comentario) {
    var notaLines = _wrapPalabras('Nota: ' + String(g.comentario), 48);
    notaLines.forEach(function(ln) { bLn(ln); });
  }

  bLn(SEP);

  // ── CUADRO 2: productos ─────────────────────────────────────
  // Header sección
  b1(0x1b); b1(0x61); b1(0x01);  // center
  b1(0x1b); b1(0x45); b1(0x01);  // bold
  bLn('PRODUCTOS');
  b1(0x1b); b1(0x45); b1(0x00);
  b1(0x1b); b1(0x61); b1(0x00);  // left
  bLn(SEP2);

  // Items con cantidad bold grande, word-wrap inteligente del nombre
  dets.forEach(function(d) {
    var tag    = d.esProductoNuevo ? 'n' : d.esIncompleto ? 'i' : ' ';
    var nombre = String(d.descripcion || '').toUpperCase();
    var cant   = d.cantidad % 1 === 0 ? String(Math.round(d.cantidad)) : String(d.cantidad);
    var marca  = (tag === 'n' ? '[N] ' : tag === 'i' ? '[!] ' : '');
    var prefix = cant + 'x ';
    var anchoP = 48 - prefix.length;     // ancho disponible primera línea
    var anchoR = 48 - 4;                  // ancho con sangría continuaciones

    var lineas = _wrapPalabras(marca + nombre, anchoP, anchoR);

    // Línea 1: cantidad bold + primera porción del nombre
    b1(0x1b); b1(0x45); b1(0x01);
    bStr(prefix);
    b1(0x1b); b1(0x45); b1(0x00);
    bLn(lineas[0] || '');
    // Continuaciones con sangría
    for (var li = 1; li < lineas.length; li++) {
      bLn('    ' + lineas[li]);
    }
    // Código de barra debajo (sangría)
    if (d.codigoProducto) {
      bLn('    ' + String(d.codigoProducto));
    }
    // Fecha de vencimiento (si tiene)
    if (d.fechaVencimiento) {
      bLn('    Venc: ' + fmtVenc(d.fechaVencimiento));
    }
  });

  if (!dets.length) bLn('  (sin items registrados)');

  bLn(SEP);

  // ── CUADRO 3: QR Code para reporte en tiempo real ───────────
  function _imprimirQR(url, titulo, sub1, sub2) {
    b1(0x1b); b1(0x61); b1(0x01);  // centrar
    b1(0x1b); b1(0x45); b1(0x01);
    bLn(titulo);
    b1(0x1b); b1(0x45); b1(0x00);

    var qrLen = url.length + 3;
    var qrpL  = qrLen & 0xff;
    var qrpH  = (qrLen >> 8) & 0xff;

    b1(0x1d); b1(0x28); b1(0x6b); b1(0x04); b1(0x00); b1(0x31); b1(0x41); b1(0x32); b1(0x00);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x43); b1(0x05);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x45); b1(0x31);
    b1(0x1d); b1(0x28); b1(0x6b); b1(qrpL); b1(qrpH); b1(0x31); b1(0x50); b1(0x30);
    bStr(url);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x51); b1(0x30);

    b1(0x1b); b1(0x21); b1(0x00);
    if (sub1) bLn(sub1);
    if (sub2) bLn(sub2);
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);
  }

  if (reporteUrl) {
    _imprimirQR(reporteUrl, 'REPORTE EN TIEMPO REAL',
                'Escanea con la camara',
                'para ver el detalle al instante');
  }

  // ── BLOQUE 2: PREINGRESO (solo si la guía proviene de uno) ─────────
  if (g.idPreingreso) {
    try {
      var preRows = _sheetToObjects(getSheet('PREINGRESOS'));
      var pi      = preRows.find(function(x) { return x.idPreingreso === g.idPreingreso; });
      if (pi) {
        // Espacio entre bloques (~2cm para corte manual)
        b1(0x1b); b1(0x4a); b1(80);   // ~1cm

        // Header: PREINGRESO
        b1(0x1b); b1(0x61); b1(0x01);
        b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);  // double-height + bold
        bLn('PREINGRESO');
        b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
        b1(0x1b); b1(0x61); b1(0x00);
        bLn(SEP);

        // Empresa (nombre del proveedor) — centrado, bold doble alto
        var preProvName = '';
        try {
          var preProv = provs ? provs.find(function(p) { return p.idProveedor === pi.idProveedor; }) : null;
          if (!preProv) {
            var ps = _sheetToObjects(getProveedoresSheet());
            preProv = ps.find(function(p) { return p.idProveedor === pi.idProveedor; });
          }
          if (preProv) preProvName = String(preProv.nombre || '');
        } catch(e) {}

        if (preProvName) {
          b1(0x1b); b1(0x61); b1(0x01);
          b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
          var nameLines = _wrapPalabras(preProvName.toUpperCase(), 24);
          for (var nl = 0; nl < nameLines.length && nl < 2; nl++) bLn(nameLines[nl]);
          b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);

          // Fecha: "14 abril 2pm"
          var fechaPI = '';
          try {
            var d = new Date(pi.fecha || pi.fechaCreacion || new Date());
            var meses = ['enero','febrero','marzo','abril','mayo','junio',
                         'julio','agosto','septiembre','octubre','noviembre','diciembre'];
            var dia    = d.getDate();
            var mes    = meses[d.getMonth()] || '';
            var hh     = d.getHours();
            var mm     = d.getMinutes();
            var ampm   = hh >= 12 ? 'pm' : 'am';
            var hh12   = hh % 12; if (hh12 === 0) hh12 = 12;
            var horaTxt = mm === 0 ? (hh12 + ampm) : (hh12 + ':' + (mm < 10 ? '0' : '') + mm + ampm);
            fechaPI = dia + ' ' + mes + ' ' + horaTxt;
          } catch(e) {}
          if (fechaPI) bLn(fechaPI);
          b1(0x1b); b1(0x61); b1(0x00);
        }

        bLn(SEP);

        // Estado y monto
        if (pi.estado) bLn('Estado:  ' + String(pi.estado).toUpperCase());
        if (pi.monto)  bLn('Monto:   S/. ' + pi.monto);

        // Cargadores: uno por línea
        var cargs = [];
        try { cargs = JSON.parse(pi.cargadores || '[]'); } catch(e) {}
        if (cargs.length) {
          bLn('');
          bLn('Cargadores:');
          cargs.forEach(function(c) {
            var nombre = (typeof c === 'object') ? (c.nombre || c.idPersonal || '') : String(c);
            if (nombre) bLn('  - ' + nombre);
          });
        }

        // Comentario — RESALTADO en doble alto bold
        if (pi.comentario) {
          bLn(SEP2);
          b1(0x1b); b1(0x45); b1(0x01);
          bLn('Comentario:');
          b1(0x1b); b1(0x45); b1(0x00);
          b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);  // double-height + bold
          var comLines = _wrapPalabras(String(pi.comentario), 24);
          for (var ci = 0; ci < comLines.length && ci < 4; ci++) bLn(comLines[ci]);
          b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
        }

        bLn(SEP);

        // Adjuntos (solo cuántos)
        var nFotos = pi.fotos ? String(pi.fotos).split(',').filter(Boolean).length : 0;
        b1(0x1b); b1(0x61); b1(0x01);
        b1(0x1b); b1(0x45); b1(0x01);
        bLn('ADJUNTOS');
        b1(0x1b); b1(0x45); b1(0x00);
        bLn(nFotos + ' imagen' + (nFotos !== 1 ? 'es' : '') + ' adjunta' + (nFotos !== 1 ? 's' : ''));
        bLn('ver en el reporte digital');
        b1(0x1b); b1(0x61); b1(0x00);
        bLn(SEP);

        // QR del preingreso
        var preReporteUrl = '';
        try {
          if (reporteUrl) {
            // Reemplazar tipo=guia&id=XXX por tipo=preingreso&id=YYY conservando dominio
            preReporteUrl = reporteUrl.replace(/tipo=guia/, 'tipo=preingreso')
                                       .replace(/id=[^&]*/, 'id=' + encodeURIComponent(pi.idPreingreso));
          }
        } catch(e) {}
        if (preReporteUrl) {
          _imprimirQR(preReporteUrl, 'PREINGRESO COMPLETO',
                      'Escanea para ver',
                      'fotos y detalles');
        }
      }
    } catch(e) {}
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

// Word-wrap inteligente: parte por palabras sin cortar.
// Si una palabra excede el ancho, la corta como último recurso.
// anchoPrimero = ancho de la 1ª línea; anchoResto = ancho de continuaciones (default = anchoPrimero).
function _wrapPalabras(texto, anchoPrimero, anchoResto) {
  texto = String(texto || '').trim();
  if (!texto) return [''];
  if (anchoResto == null) anchoResto = anchoPrimero;

  var palabras = texto.split(/\s+/);
  var lineas   = [];
  var cur      = '';
  var ancho    = anchoPrimero;

  for (var i = 0; i < palabras.length; i++) {
    var p = palabras[i];
    // Palabra más larga que el ancho — partir la palabra como último recurso
    while (p.length > ancho) {
      if (cur) { lineas.push(cur); cur = ''; ancho = anchoResto; }
      lineas.push(p.substring(0, ancho));
      p = p.substring(ancho);
    }
    var sep = cur ? ' ' : '';
    if ((cur + sep + p).length <= ancho) {
      cur = cur + sep + p;
    } else {
      lineas.push(cur);
      cur = p;
      ancho = anchoResto;
    }
  }
  if (cur) lineas.push(cur);
  return lineas;
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
    prods.forEach(function(p) {
      var desc = p.descripcion || p.idProducto;
      if (!desc) return;
      prodMap[p.idProducto] = desc;
      if (p.codigoBarra) prodMap[String(p.codigoBarra).trim()] = desc;
      // skuBase solo para producto BASE (factor=1) — evita que presentaciones sobreescriban
      var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
      if (esBase && p.skuBase) prodMap[String(p.skuBase).trim().toUpperCase()] = desc;
    });
    // Equivalentes → resuelven al producto base via skuBase
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        _sheetToObjects(equivSheet).forEach(function(e) {
          if (!e.codigoBarra || !e.skuBase) return;
          var skuKey = String(e.skuBase).trim().toUpperCase();
          var desc   = prodMap[skuKey];
          if (desc) prodMap[String(e.codigoBarra).trim()] = desc;
        });
      }
    } catch(e) {}

    var pns = _sheetToObjects(getSheet('PRODUCTO_NUEVO'));
    var pnMap = {};
    pns.forEach(function(pn) {
      var cod = String(pn.codigoBarra || '');
      if (cod) pnMap[cod] = { desc: String(pn.descripcion || pn.marca || cod), estado: String(pn.estado || '') };
    });

    dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
      .filter(function(d) { return d.idGuia === id && d.observacion !== 'ANULADO'; })
      .map(function(d) {
        var cod        = String(d.codigoProducto || '');
        var esPN       = !!(pnMap[cod]) || cod.indexOf('NLEV') === 0;
        var enCatalogo = !!prodMap[cod];
        var desc       = esPN ? (pnMap[cod] ? pnMap[cod].desc : cod)
                       : enCatalogo ? prodMap[cod] : cod;
        return {
          codigoProducto:   cod,
          descripcion:      desc,
          cantidadEsperada: d.cantidadEsperada || 0,
          cantidadReal:     d.cantidadReal || d.cantidadRecibida || 0,
          fechaVencimiento: String(d.fechaVencimiento || '').split('T')[0],
          observacion:      d.observacion || '',
          esProductoNuevo:  esPN,
          esIncompleto:     !esPN && !enCatalogo,
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
