// ============================================================
// warehouseMos — Reporte.gs
// Endpoint público para páginas de reporte en tiempo real
// Acción: getReporte?tipo=preingreso|guia&id=xxx
// ============================================================

// ════════════════════════════════════════════════════════════
// Diagnóstico para entender por qué un ticket no muestra 3 secciones.
// Ejecutar desde el editor: diagnosticarTicketPickup('GXXXXXX')
// Retorna el estado completo del clasificador para esa guía.
// ════════════════════════════════════════════════════════════
function diagnosticarTicketPickup(idGuia) {
  var idG = String(idGuia || '').trim();
  if (!idG) return { ok: false, error: 'idGuia requerido' };

  var guias = _sheetToObjects(getSheet('GUIAS'));
  var g = guias.find(function(x){ return x.idGuia === idG; });
  if (!g) return { ok: false, error: 'Guía no encontrada: ' + idG };

  // Construir dets igual que imprimirTicketGuia
  var dets = _sheetToObjects(getSheet('GUIA_DETALLE'))
    .filter(function(d){ return d.idGuia === idG && d.observacion !== 'ANULADO'; })
    .map(function(d){
      return {
        codigoProducto:  String(d.codigoProducto || ''),
        descripcion:     String(d.descripcion || d.codigoProducto || ''),
        cantidad:        parseFloat(d.cantidadReal || d.cantidadRecibida || d.cantidadEsperada || 0),
        observacion:     String(d.observacion || '')
      };
    });

  var matchPickup = String(g.comentario || '').match(/\[pickup:([^\]]+)\]/);
  var clasif = _clasificarDetallesPorPickup(g, dets);

  var resumen = {
    idGuia:           idG,
    comentario:       String(g.comentario || ''),
    matchRegex:       matchPickup ? matchPickup[1] : null,
    detsCount:        dets.length,
    detsCodigos:      dets.map(function(d){ return d.codigoProducto; }),
    hasPickup:        clasif.hasPickup,
    okCount:          clasif.ok.length,
    extrasCount:      clasif.extras.length,
    faltantesCount:   clasif.faltantes.length,
    okList:           clasif.ok.map(function(d){ return d.codigoProducto + ' x' + d.cantidad; }),
    extrasList:       clasif.extras.map(function(d){ return d.codigoProducto + ' x' + d.cantidad; }),
    faltantesList:    clasif.faltantes.map(function(f){ return f.skuBase + ' falt ' + (f.solicitado - f.despachado); })
  };
  Logger.log(JSON.stringify(resumen, null, 2));
  return resumen;
}

// ════════════════════════════════════════════════════════════
// _clasificarDetallesPorPickup — divide los detalles de una guía
// en 3 secciones cuando la guía proviene de un pickup ME→WH.
// ────────────────────────────────────────────────────────────
// Devuelve { hasPickup, ok[], extras[], faltantes[] } donde:
//   ok        = detalles cuyo total escaneado coincide EXACTO con lo
//               solicitado para ese item del pickup
//   extras    = detalles que aportan MÁS de lo solicitado (sobrante)
//               o detalles cuyo codigoBarra no está en ningún item
//               del pickup (escaneo fuera del pickup)
//   faltantes = items del pickup que se solicitaron y NO se llegaron
//               a despachar completo (incluye despachado=0)
//
// Si la guía no tiene marca de pickup en su comentario, retorna
// { hasPickup:false } y el ticket usa lista única.
// ════════════════════════════════════════════════════════════
function _clasificarDetallesPorPickup(g, dets) {
  var out = { hasPickup: false, ok: [], extras: [], faltantes: [] };
  var comentario = String(g.comentario || '');
  var m = comentario.match(/\[pickup:([^\]]+)\]/);
  if (!m) return out;
  var idPickup = m[1];

  var sheetP = null;
  try { sheetP = getSheet('PICKUPS'); } catch(e) { return out; }
  if (!sheetP) return out;

  var dataP = sheetP.getDataRange().getValues();
  if (!dataP.length) return out;
  var hdrsP   = dataP[0].map(function(h){return String(h);});
  var idxId   = hdrsP.indexOf('idPickup');
  var idxItms = hdrsP.indexOf('items');
  if (idxId < 0 || idxItms < 0) return out;

  var rowJson = null;
  for (var r = 1; r < dataP.length; r++) {
    if (String(dataP[r][idxId]) === idPickup) {
      rowJson = String(dataP[r][idxItms] || '');
      break;
    }
  }
  if (!rowJson) return out;

  var pickupItems;
  try { pickupItems = JSON.parse(rowJson); } catch(e) { return out; }
  if (!Array.isArray(pickupItems)) return out;

  out.hasPickup = true;

  // Mapa codigoBarra (upper) → { item del pickup, índice }
  var codToItem = {};
  pickupItems.forEach(function(it, idx) {
    var codos = (it.codigosOriginales || []);
    codos.forEach(function(c) {
      if (!c) return;
      codToItem[String(c).trim().toUpperCase()] = { item: it, idx: idx };
    });
    // Permitir match por skuBase / idProducto también (la trazabilidad
    // del pickup acepta cualquiera de los códigos del array codigosOriginales,
    // pero a veces el pickup contiene idProducto IDPRO… mezclado).
    if (it.skuBase) codToItem[String(it.skuBase).trim().toUpperCase()] = { item: it, idx: idx };
    // ⚡ FIX: incluir los códigos FÍSICOS realmente escaneados, que viven en
    // despachadoPorCodigo (ej. 'WHANEODO250GR': 6). codigosOriginales viene
    // de MosExpress con códigos de catálogo MOS, pero el operador WH escanea
    // códigos internos del almacén que NO están ahí → sin esto, los detalles
    // de la guía no matcheaban y caían como "no despachado" (bug 2/17).
    if (it.despachadoPorCodigo && typeof it.despachadoPorCodigo === 'object') {
      Object.keys(it.despachadoPorCodigo).forEach(function(c) {
        if (!c) return;
        codToItem[String(c).trim().toUpperCase()] = { item: it, idx: idx };
      });
    }
  });

  // Acumular despachado por item del pickup (suma de cantidades de detalles
  // cuyo codigoBarra mapea a ese item)
  var despPorIdx = {};
  var detsExtra  = []; // detalles cuyo codigoBarra no pertenece al pickup

  dets.forEach(function(d) {
    var cb = String(d.codigoProducto || '').trim().toUpperCase();
    var hit = codToItem[cb];
    if (hit) {
      despPorIdx[hit.idx] = (despPorIdx[hit.idx] || 0) + (parseFloat(d.cantidad) || 0);
      // Lo asociamos para luego decidir si va a "ok" o "extras"
      d._pickupIdx = hit.idx;
    } else {
      detsExtra.push(d);
    }
  });

  // Clasificar cada detalle individual de la guía. Forzar la descripción
  // al NOMBRE DEL CANÓNICO (que viene en it.nombre del pickup) para que el
  // ticket nunca muestre nombre de equivalente / presentación.
  dets.forEach(function(d) {
    if (typeof d._pickupIdx === 'undefined') return; // ya está en detsExtra
    var it = pickupItems[d._pickupIdx];
    if (it && it.nombre) d.descripcion = String(it.nombre);
    var sol  = parseFloat(it.solicitado) || 0;
    var desp = parseFloat(despPorIdx[d._pickupIdx]) || 0;
    if (desp > sol + 1e-9) {
      out.extras.push(d);
    } else {
      out.ok.push(d);
    }
  });

  // Los detalles cuyo código no estaba en el pickup → "extras" (escaneos libres)
  detsExtra.forEach(function(d) { out.extras.push(d); });

  // Items del pickup que NO se despacharon o se despacharon parcial → "faltantes"
  pickupItems.forEach(function(it, idx) {
    var sol  = parseFloat(it.solicitado) || 0;
    var desp = parseFloat(despPorIdx[idx]) || 0;
    if (sol > 0 && desp + 1e-9 < sol) {
      out.faltantes.push({
        skuBase:     it.skuBase     || '',
        nombre:      it.nombre      || it.skuBase || '(sin nombre)',
        solicitado:  sol,
        despachado:  desp
      });
    }
  });

  return out;
}

// ── Helpers de formato compartidos ──────────────────────────
function _padLine48(left, right) {
  var l = String(left || '');
  var r = String(right || '');
  var pad = 48 - l.length - r.length;
  if (pad < 1) pad = 1;
  return l + Array(pad + 1).join(' ') + r;
}

function _fmtMoney(n) {
  var v = parseFloat(n);
  if (isNaN(v)) return '0.00';
  return v.toFixed(2);
}

// ============================================================
// imprimirTicketGuia — ESC/POS 80mm con QR al final
// ============================================================
function imprimirTicketGuia(params) {
  var idGuia = String(params.idGuia || '');
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  // printerIdOverride: lo manda el modal de selección de impresora (admin/
  // master eligió una impresora distinta a la del almacén). Sin override
  // → impresora TICKET de ALMACEN como siempre.
  var printerId;
  if (params.printerIdOverride) {
    printerId = String(params.printerIdOverride).trim();
  } else {
    try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
    catch(e) { return { ok: false, error: e.message }; }
  }

  // [Fix #1+#3 v2.11.2] Garantizar lectura consistente antes de imprimir.
  // Bug: ticket salía cortado con 2-3 items aunque la guía tenía 22 porque
  // imprimirTicketGuia se ejecutaba antes de que Sheets confirmara todas
  // las escrituras del forEach previo. Solución:
  //   1. Flush para forzar persistencia
  //   2. Si el caller pasa `esperadoDetalles`, validar que vemos todos.
  //      Si vemos menos: esperar + flush + reintentar (hasta 3 veces).
  try { SpreadsheetApp.flush(); } catch(_){}
  var esperado = parseInt(params.esperadoDetalles) || 0;
  if (esperado > 0) {
    function _contarDetalles() {
      return _sheetToObjects(getSheet('GUIA_DETALLE'))
        .filter(function(d) { return d.idGuia === idGuia && d.observacion !== 'ANULADO'; })
        .length;
    }
    var actual = _contarDetalles();
    var intentos = 0;
    while (actual < esperado && intentos < 3) {
      Logger.log('[imprimirTicket] guía=' + idGuia + ' visibles=' + actual + '/' + esperado + ' · espero 1.5s (retry ' + (intentos+1) + ')');
      Utilities.sleep(1500);
      try { SpreadsheetApp.flush(); } catch(_){}
      actual = _contarDetalles();
      intentos++;
    }
    Logger.log('[imprimirTicket] guía=' + idGuia + ' final visibles=' + actual + '/' + esperado + ' tras ' + intentos + ' retries');
  }

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

    // PASO 1 — índice skuBase → desc del CANÓNICO (factor=1, activo).
    // Fuente de verdad para resolver presentaciones y equivalentes.
    var canonicoPorSku = {};
    prods.forEach(function(p) {
      var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
      if (!esBase || !p.skuBase) return;
      var sk = String(p.skuBase).trim().toUpperCase();
      canonicoPorSku[sk] = p.descripcion || p.idProducto || sk;
    });

    // PASO 2 — cada producto apunta al canónico de su skuBase.
    // Presentaciones (factor!=1) ya NO ganan sobre el canónico.
    prods.forEach(function(p) {
      var sk = String(p.skuBase || '').trim().toUpperCase();
      var desc = canonicoPorSku[sk] || p.descripcion || p.idProducto;
      if (!desc) return;
      if (p.idProducto)  prodMap[p.idProducto] = desc;
      if (p.codigoBarra) prodMap[String(p.codigoBarra).trim()] = desc;
      if (sk)            prodMap[sk] = desc;
    });

    // PASO 3 — equivalentes: SIEMPRE al canónico de su skuBase.
    // Sobrescriben cualquier mapeo previo (incluyendo presentaciones con
    // el mismo codigoBarra que el equivalente).
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        _sheetToObjects(equivSheet).forEach(function(e) {
          if (!e.codigoBarra) return;
          var sk = String(e.skuBase || '').trim().toUpperCase();
          var desc = canonicoPorSku[sk] || prodMap[sk] || String(e.descripcion || '') || String(e.codigoBarra);
          prodMap[String(e.codigoBarra).trim()] = desc;
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
  // Si la guía proviene de un pickup ([pickup:PCK-X] en comentario), imprimimos
  // en 3 secciones: DESPACHADO (OK) · EXTRAS/SOBRANTES · FALTANTES. Caso contrario,
  // lista única como antes.
  var pickupClasif = _clasificarDetallesPorPickup(g, dets);

  // Helper para renderizar un detalle de la guía con cantidad bold + word-wrap
  function _imprimirItemDetalle(d) {
    var tag    = d.esProductoNuevo ? 'n' : d.esIncompleto ? 'i' : ' ';
    var nombre = String(d.descripcion || '').toUpperCase();
    var cant   = d.cantidad % 1 === 0 ? String(Math.round(d.cantidad)) : String(d.cantidad);
    var marca  = (tag === 'n' ? '[N] ' : tag === 'i' ? '[!] ' : '');
    var prefix = cant + 'x ';
    var anchoP = 48 - prefix.length;
    var anchoR = 48 - 4;
    var lineas = _wrapPalabras(marca + nombre, anchoP, anchoR);
    b1(0x1b); b1(0x45); b1(0x01);
    bStr(prefix);
    b1(0x1b); b1(0x45); b1(0x00);
    bLn(lineas[0] || '');
    for (var li = 1; li < lineas.length; li++) bLn('    ' + lineas[li]);
    if (d.codigoProducto) bLn('    ' + String(d.codigoProducto));
    if (d.fechaVencimiento) bLn('    Venc: ' + fmtVenc(d.fechaVencimiento));
  }

  // Helper para items "faltantes" del pickup (no están en GUIA_DETALLE).
  // Muestra: solicitado vs despachado=0 (o despachado < solicitado).
  function _imprimirItemFaltante(it) {
    var nombre = String(it.nombre || '').toUpperCase();
    var sol    = parseFloat(it.solicitado) || 0;
    var desp   = parseFloat(it.despachado) || 0;
    var falta  = sol - desp;
    var qFalta = falta % 1 === 0 ? String(Math.round(falta)) : String(falta);
    var prefix = '-' + qFalta + ' ';
    var anchoP = 48 - prefix.length;
    var anchoR = 48 - 4;
    var lineas = _wrapPalabras(nombre, anchoP, anchoR);
    b1(0x1b); b1(0x45); b1(0x01);
    bStr(prefix);
    b1(0x1b); b1(0x45); b1(0x00);
    bLn(lineas[0] || '');
    for (var li = 1; li < lineas.length; li++) bLn('    ' + lineas[li]);
    bLn('    (pidio ' + sol + ', llego ' + desp + ')');
  }

  if (pickupClasif.hasPickup) {
    // ─── Sección 1: DESPACHADO OK (solicitado === despachado) ───
    b1(0x1b); b1(0x61); b1(0x01); b1(0x1b); b1(0x45); b1(0x01);
    bLn('DESPACHADO OK (' + pickupClasif.ok.length + ')');
    b1(0x1b); b1(0x45); b1(0x00); b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP2);
    if (pickupClasif.ok.length) {
      pickupClasif.ok.forEach(_imprimirItemDetalle);
    } else {
      bLn('  (ninguno coincidio exacto)');
    }
    bLn(SEP);

    // ─── Sección 2: EXTRAS / SOBRANTES (despachado > solicitado, o no en pickup) ───
    b1(0x1b); b1(0x61); b1(0x01); b1(0x1b); b1(0x45); b1(0x01);
    bLn('EXTRAS / SOBRANTES (' + pickupClasif.extras.length + ')');
    b1(0x1b); b1(0x45); b1(0x00); b1(0x1b); b1(0x61); b1(0x00);
    if (pickupClasif.extras.length) {
      bLn('(producto pedido en mayor cant.');
      bLn(' o no estaba en el pickup)');
      bLn(SEP2);
      pickupClasif.extras.forEach(_imprimirItemDetalle);
    } else {
      bLn(SEP2);
      bLn('  (sin extras)');
    }
    bLn(SEP);

    // ─── Sección 3: FALTANTES (despachado < solicitado o no despachado) ───
    b1(0x1b); b1(0x61); b1(0x01); b1(0x1b); b1(0x45); b1(0x01);
    bLn('NO DESPACHADO (' + pickupClasif.faltantes.length + ')');
    b1(0x1b); b1(0x45); b1(0x00); b1(0x1b); b1(0x61); b1(0x00);
    if (pickupClasif.faltantes.length) {
      bLn('(quedaron pendientes)');
      bLn(SEP2);
      pickupClasif.faltantes.forEach(_imprimirItemFaltante);
    } else {
      bLn(SEP2);
      bLn('  Sin faltantes - PICKUP COMPLETO');
    }
    bLn(SEP);
  } else {
    // Comportamiento clásico — lista única "PRODUCTOS"
    b1(0x1b); b1(0x61); b1(0x01); b1(0x1b); b1(0x45); b1(0x01);
    bLn('PRODUCTOS');
    b1(0x1b); b1(0x45); b1(0x00); b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP2);
    dets.forEach(_imprimirItemDetalle);
    if (!dets.length) bLn('  (sin items registrados)');
    bLn(SEP);
  }

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

  // ── BLOQUE PREINGRESO (solo si la guía proviene de uno) ───────────
  // Va ARRIBA de la guía: redirigimos las escrituras (b1/bStr/bLn/_imprimirQR
  // pushean a la variable B del scope) a un buffer aparte `Bpre`, y al final
  // lo anteponemos al ticket de la guía. Antes el preingreso se imprimía
  // debajo de la guía; ahora va primero.
  var Bpre = [];
  var _Bmain = B;
  B = Bpre;
  if (g.idPreingreso) {
    try {
      var preRows = _sheetToObjects(getSheet('PREINGRESOS'));
      var pi      = preRows.find(function(x) { return x.idPreingreso === g.idPreingreso; });
      if (pi) {

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

        // Cargadores con detalle: nombre · N x tarifa = subtotal
        var cargs = [];
        try { cargs = JSON.parse(pi.cargadores || '[]'); } catch(e) {}
        if (cargs.length) {
          bLn('');
          bLn('Cargadores:');
          var totalCargG = 0;
          var tarifaGlobalG = parseFloat(_getConfigValue('TARIFA_CARRETA')) || 0;
          cargs.forEach(function(c) {
            var nombre = (typeof c === 'object') ? (c.nombre || c.idPersonal || '') : String(c);
            if (!nombre) return;
            var carretas = (typeof c === 'object' && c.carretas) ? parseInt(c.carretas) || 0 : 0;
            var tarifa   = (typeof c === 'object' && c.tarifa !== undefined && c.tarifa !== '')
                           ? (parseFloat(c.tarifa) || 0) : tarifaGlobalG;
            var sub      = carretas * tarifa;
            totalCargG  += sub;
            if (carretas > 0 && tarifa > 0) {
              bLn(_padLine48('  - ' + nombre, carretas + ' x S/' + _fmtMoney(tarifa) + ' = S/ ' + _fmtMoney(sub)));
            } else {
              bLn('  - ' + nombre);
            }
          });
          if (totalCargG > 0) {
            b1(0x1b); b1(0x45); b1(0x01);
            bLn(_padLine48('  TOTAL CARRETAS', 'S/ ' + _fmtMoney(totalCargG)));
            b1(0x1b); b1(0x45); b1(0x00);
          }
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
  // Restaurar el buffer principal (la guía) y anteponer el preingreso si
  // hubo. Separador de ~1cm entre el preingreso (arriba) y la guía (abajo).
  B = _Bmain;
  if (Bpre.length) {
    B = Bpre.concat([0x1b, 0x4a, 60]).concat(_Bmain);
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
    if (code === 201) {
      // [Fix #5 v2.11.2] Devolver conteo de detalles efectivamente impresos
      // para que el caller (frontend / crearDespachoRapido) pueda validar
      // que el ticket salió completo y disparar reimpresión si no.
      Logger.log('[imprimirTicket] guía=' + idGuia + ' OK · detallesImpresos=' + dets.length);
      return { ok: true, data: { detallesImpresos: dets.length, idGuia: idGuia } };
    }
    return { ok: false, error: 'PrintNode ' + code + ': ' + resp.getContentText() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// imprimirAvisoCajeros — manda ticket de aviso de preingreso a
// la(s) impresora(s) donde haya un cajero/vendedor con caja ABIERTA
// en MosExpress. Lee MosExpress.CAJAS WHERE Estado=ABIERTA y para
// cada caja toma su PrintNode_ID directo (lo que MosExpress guardó
// al abrirla).
// ============================================================
function imprimirAvisoCajeros(params) {
  var idPreingreso = String(params.idPreingreso || '');
  if (!idPreingreso) return { ok: false, error: 'idPreingreso requerido' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  // 1. Leer preingreso
  var pi;
  try {
    var preRows = _sheetToObjects(getSheet('PREINGRESOS'));
    pi = preRows.find(function(r) { return r.idPreingreso === idPreingreso; });
  } catch(e) { return { ok: false, error: 'Error leyendo PREINGRESOS: ' + e.message }; }
  if (!pi) return { ok: false, error: 'Preingreso no encontrado: ' + idPreingreso };

  // 2. Resolver proveedor
  var provName = '';
  try {
    var provs = _sheetToObjects(getProveedoresSheet());
    var prov  = provs.find(function(p) { return p.idProveedor === pi.idProveedor; });
    if (prov) provName = String(prov.nombre || '');
  } catch(e) {}

  // 3. Leer cajas ABIERTAS en MosExpress
  var cajasAbiertas = [];
  try {
    var sheetCajas = _getMosExpressSS().getSheetByName('CAJAS');
    if (!sheetCajas) return { ok: false, error: 'Hoja CAJAS no existe en MosExpress' };
    var rows = _sheetToObjects(sheetCajas);
    cajasAbiertas = rows.filter(function(c) {
      return String(c.Estado || '').trim().toUpperCase() === 'ABIERTA' && c.PrintNode_ID;
    });
  } catch(e) { return { ok: false, error: 'No se pudo leer CAJAS de MosExpress: ' + e.message }; }

  if (!cajasAbiertas.length) {
    return { ok: false, error: 'NO_HAY_CAJEROS_ACTIVOS', mensaje: 'No hay cajas abiertas con impresora asignada' };
  }

  // 4. URL del reporte público para QR
  var reporteUrl = String(params.reporteUrl || '');

  // 5. Construir ESC/POS UNA VEZ (mismo ticket para todas las cajas)
  var bytes = _construirAvisoIngresoBytes(pi, provName, reporteUrl);
  var blob  = Utilities.newBlob(bytes, 'application/octet-stream');
  var b64   = Utilities.base64Encode(blob.getBytes());

  // 6. Deduplicar por PrintNode_ID: varias cajas pueden compartir la
  //    misma impresora física. Antes mandábamos un job por caja y eso
  //    causaba 2-3 impresiones idénticas si había 2-3 cajas abiertas
  //    apuntando al mismo printer (bug histórico 2026-05-13).
  //    Ahora: 1 job por printer único, pero el resultado lista TODAS las
  //    cajas que reciben ese aviso (para auditoría/UI).
  var porPrinter = {};
  cajasAbiertas.forEach(function(caja) {
    var pid = String(caja.PrintNode_ID || '').trim();
    if (!pid) return;
    if (!porPrinter[pid]) porPrinter[pid] = { printerId: pid, cajas: [] };
    porPrinter[pid].cajas.push(caja);
  });

  // 7. Enviar UN job por printer único
  var resultados = [];
  Object.keys(porPrinter).forEach(function(pid) {
    var grupo = porPrinter[pid];
    var printerId = parseInt(grupo.printerId);
    var cajasDelPrinter = grupo.cajas;
    if (!printerId) {
      cajasDelPrinter.forEach(function(c) {
        resultados.push({ vendedor: c.Vendedor, zona: c.Zona_ID, estacion: c.Estacion,
                          printNodeId: c.PrintNode_ID, ok: false,
                          error: 'PrintNode_ID inválido', dedupCount: cajasDelPrinter.length });
      });
      return;
    }
    var okJob = false, errMsg = '';
    try {
      var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
        method:  'post',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
          'Content-Type':  'application/json'
        },
        payload: JSON.stringify({
          printerId:   printerId,
          title:       'Aviso ingreso ' + idPreingreso,
          contentType: 'raw_base64',
          content:     b64,
          source:      'warehouseMos'
        }),
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      okJob = (code === 201);
      if (!okJob) errMsg = 'PrintNode ' + code;
    } catch(e) {
      errMsg = e.message;
    }
    // 1 resultado por caja (para la UI), pero todas comparten el mismo job
    cajasDelPrinter.forEach(function(c) {
      resultados.push({
        vendedor: c.Vendedor, zona: c.Zona_ID, estacion: c.Estacion,
        printNodeId: c.PrintNode_ID,
        ok: okJob, error: errMsg,
        dedupCount: cajasDelPrinter.length  // info: cuántas cajas comparten esta impresora
      });
    });
  });

  var algunOk = resultados.some(function(r) { return r.ok; });
  return { ok: algunOk, data: { idPreingreso: idPreingreso, impresiones: resultados } };
}

// Construye los bytes ESC/POS del ticket de aviso de ingreso (sin la palabra "PREINGRESO").
// Header: empresa grande / fecha / hora / monto bien grande / cargadores / comentario destacado / QR
function _construirAvisoIngresoBytes(pi, provName, reporteUrl) {
  var B = [];
  function b1(v)   { B.push(v & 0xff); }
  function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
  function bLn(s)  { bStr(s); b1(0x0a); }

  var SEP  = '================================================';
  var SEP2 = '------------------------------------------------';

  // Init
  b1(0x1b); b1(0x40);

  // ── Header: WAREHOUSE / MOS / AVISO INGRESO ─────────────────
  b1(0x1b); b1(0x61); b1(0x01);
  b1(0x1b); b1(0x21); b1(0x38);
  bLn('WAREHOUSE');
  bLn('MOS');
  b1(0x1b); b1(0x21); b1(0x00);
  b1(0x1b); b1(0x45); b1(0x01);
  bLn('AVISO INGRESO');
  b1(0x1b); b1(0x45); b1(0x00);
  b1(0x1b); b1(0x61); b1(0x00);
  bLn(SEP);

  // ── Empresa: doble alto bold, centrado ──────────────────────
  if (provName) {
    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
    var nameLines = _wrapPalabras(String(provName).toUpperCase(), 24);
    for (var nl = 0; nl < nameLines.length && nl < 2; nl++) bLn(nameLines[nl]);
    b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
  }

  // ── Fecha estilo "2 mayo 1:11pm" — timezone Lima forzado ────
  // Usar Utilities.formatDate con tz explícita para evitar desfases
  // si el script está configurado en otra zona horaria.
  var fechaPI = '';
  try {
    var rawFecha = pi.fecha || pi.fechaCreacion;
    var d = (rawFecha instanceof Date) ? rawFecha : new Date(rawFecha || new Date());
    if (!isNaN(d.getTime())) {
      var tzLima = 'America/Lima';
      var meses = ['enero','febrero','marzo','abril','mayo','junio',
                   'julio','agosto','septiembre','octubre','noviembre','diciembre'];
      var dia    = parseInt(Utilities.formatDate(d, tzLima, 'd'), 10);
      var mesIdx = parseInt(Utilities.formatDate(d, tzLima, 'M'), 10) - 1;
      var hh     = parseInt(Utilities.formatDate(d, tzLima, 'H'), 10);
      var mm     = parseInt(Utilities.formatDate(d, tzLima, 'm'), 10);
      var mes    = meses[mesIdx] || '';
      var ampm   = hh >= 12 ? 'pm' : 'am';
      var hh12   = hh % 12; if (hh12 === 0) hh12 = 12;
      var horaTxt = mm === 0 ? (hh12 + ampm) : (hh12 + ':' + (mm < 10 ? '0' : '') + mm + ampm);
      fechaPI = dia + ' ' + mes + ' ' + horaTxt;
    }
  } catch(e) {}
  if (fechaPI) {
    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn(fechaPI);
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
  }

  bLn(SEP);

  // ── Monto a preparar (BIEN GRANDE, lo más importante) ──────
  if (pi.monto) {
    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('PREPARAR PARA PAGAR');
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x21); b1(0x38); b1(0x1b); b1(0x45); b1(0x01);
    bLn('S/. ' + pi.monto);
    b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);
  }

  // ── Cargadores con detalle: nombre · N x tarifa = subtotal ──
  var cargs = [];
  try { cargs = JSON.parse(pi.cargadores || '[]'); } catch(e) {}
  if (cargs.length) {
    bLn('');
    bLn('Cargadores:');
    var totalCarg = 0;
    var tarifaGlobal = parseFloat(_getConfigValue('TARIFA_CARRETA')) || 0;
    cargs.forEach(function(c) {
      var nombre = (typeof c === 'object') ? (c.nombre || c.idPersonal || '') : String(c);
      if (!nombre) return;
      var carretas = (typeof c === 'object' && c.carretas) ? parseInt(c.carretas) || 0 : 0;
      var tarifa   = (typeof c === 'object' && c.tarifa !== undefined && c.tarifa !== '')
                     ? (parseFloat(c.tarifa) || 0) : tarifaGlobal;
      var sub      = carretas * tarifa;
      totalCarg   += sub;
      if (carretas > 0 && tarifa > 0) {
        bLn(_padLine48('  - ' + nombre, carretas + ' x S/' + _fmtMoney(tarifa) + ' = S/ ' + _fmtMoney(sub)));
      } else {
        bLn('  - ' + nombre);
      }
    });
    if (totalCarg > 0) {
      bLn('                                                ');
      b1(0x1b); b1(0x45); b1(0x01);
      bLn(_padLine48('  TOTAL CARRETAS', 'S/ ' + _fmtMoney(totalCarg)));
      b1(0x1b); b1(0x45); b1(0x00);
    }
  }

  // ── Comentario: doble alto bold (resaltado) ─────────────────
  if (pi.comentario) {
    bLn(SEP2);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('Comentario:');
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
    var comLines = _wrapPalabras(String(pi.comentario), 24);
    for (var ci = 0; ci < comLines.length && ci < 5; ci++) bLn(comLines[ci]);
    b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
  }

  bLn(SEP);

  // ── Adjuntos ────────────────────────────────────────────────
  var nFotos = pi.fotos ? String(pi.fotos).split(',').filter(Boolean).length : 0;
  if (nFotos > 0) {
    b1(0x1b); b1(0x61); b1(0x01);
    bLn(nFotos + ' imagen' + (nFotos !== 1 ? 'es' : '') + ' adjunta' + (nFotos !== 1 ? 's' : ''));
    bLn('ver en el reporte digital');
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);
  }

  // ── QR del reporte del preingreso ───────────────────────────
  if (reporteUrl) {
    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('VER PREINGRESO COMPLETO');
    b1(0x1b); b1(0x45); b1(0x00);

    var qrLen = reporteUrl.length + 3;
    var qrpL  = qrLen & 0xff;
    var qrpH  = (qrLen >> 8) & 0xff;
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x04); b1(0x00); b1(0x31); b1(0x41); b1(0x32); b1(0x00);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x43); b1(0x05);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x45); b1(0x31);
    b1(0x1d); b1(0x28); b1(0x6b); b1(qrpL); b1(qrpH); b1(0x31); b1(0x50); b1(0x30);
    bStr(reporteUrl);
    b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x51); b1(0x30);

    bLn('Escanea para fotos y detalles');
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);
  }

  // Feed + corte
  b1(0x1b); b1(0x4a); b1(160);
  b1(0x1d); b1(0x56); b1(0x00);

  return B;
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
      tarifaCarreta:parseFloat(_getConfigValue('TARIFA_CARRETA')) || 0,
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

    // PASO 1 — índice skuBase → desc del CANÓNICO (factor=1, activo).
    // Fuente de verdad para resolver presentaciones y equivalentes.
    var canonicoPorSku = {};
    prods.forEach(function(p) {
      var esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
      if (!esBase || !p.skuBase) return;
      var sk = String(p.skuBase).trim().toUpperCase();
      canonicoPorSku[sk] = p.descripcion || p.idProducto || sk;
    });

    // PASO 2 — cada producto apunta al canónico de su skuBase.
    // Presentaciones (factor!=1) ya NO ganan sobre el canónico.
    prods.forEach(function(p) {
      var sk = String(p.skuBase || '').trim().toUpperCase();
      var desc = canonicoPorSku[sk] || p.descripcion || p.idProducto;
      if (!desc) return;
      if (p.idProducto)  prodMap[p.idProducto] = desc;
      if (p.codigoBarra) prodMap[String(p.codigoBarra).trim()] = desc;
      if (sk)            prodMap[sk] = desc;
    });

    // PASO 3 — equivalentes: SIEMPRE al canónico de su skuBase.
    // Sobrescriben cualquier mapeo previo (incluyendo presentaciones con
    // el mismo codigoBarra que el equivalente).
    try {
      var equivSheet = _getMosSS().getSheetByName('EQUIVALENCIAS');
      if (equivSheet) {
        _sheetToObjects(equivSheet).forEach(function(e) {
          if (!e.codigoBarra) return;
          var sk = String(e.skuBase || '').trim().toUpperCase();
          var desc = canonicoPorSku[sk] || prodMap[sk] || String(e.descripcion || '') || String(e.codigoBarra);
          prodMap[String(e.codigoBarra).trim()] = desc;
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

// ============================================================
// CARGADORES DEL DÍA — agregado cross-preingresos
// ============================================================
// [v2.13] Sin tarifa. Devuelve cargadores agrupados con conteos por estado
// de carga (LLENA / MEDIA / VACIA). Si un cargador legacy no tiene `estados`,
// se asume todas LLENAS.
function getCargadoresDelDia(params) {
  var fechaStr = String(params && params.fecha || '').trim();
  if (!fechaStr) {
    var hoy = new Date();
    fechaStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var rows = _sheetToObjects(getSheet('PREINGRESOS'));

  // Resolver nombres de proveedores
  var provMap = {};
  try {
    _sheetToObjects(getProveedoresSheet()).forEach(function(p) {
      provMap[String(p.idProveedor)] = String(p.nombre || '');
    });
  } catch(e) {}

  var preingresosDelDia = rows.filter(function(pi) {
    if (!pi.fecha) return false;
    var d  = new Date(pi.fecha);
    if (isNaN(d.getTime())) return false;
    var ds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return ds === fechaStr;
  });

  // Helper: normaliza estados (asume LLENA si falta)
  function _normEstados(c) {
    var n = parseInt(c.carretas) || 1;
    var arr = [];
    if (Array.isArray(c.estados)) {
      arr = c.estados.slice(0, n).map(function(e) {
        return (e === 'MEDIA' || e === 'VACIA') ? e : 'LLENA';
      });
    }
    while (arr.length < n) arr.push('LLENA');
    return arr;
  }

  // Agrupar por cargador
  var byId = {};
  preingresosDelDia.forEach(function(pi) {
    var cargs = [];
    try { cargs = JSON.parse(pi.cargadores || '[]'); } catch(e) {}
    if (!Array.isArray(cargs)) return;
    cargs.forEach(function(c) {
      if (!c || typeof c !== 'object') return;
      var id     = String(c.id || c.idPersonal || c.nombre || '');
      var nombre = String(c.nombre || c.idPersonal || id || '');
      if (!id || !nombre) return;
      var carretas = parseInt(c.carretas) || 0;
      var estados  = _normEstados(c);
      var llenas = 0, medias = 0, vacias = 0;
      estados.forEach(function(e) {
        if (e === 'LLENA') llenas++;
        else if (e === 'MEDIA') medias++;
        else if (e === 'VACIA') vacias++;
      });
      if (!byId[id]) byId[id] = {
        id: id, nombre: nombre,
        carretasTotal: 0, llenasTotal: 0, mediasTotal: 0, vaciasTotal: 0,
        preingresos: []
      };
      byId[id].carretasTotal += carretas;
      byId[id].llenasTotal   += llenas;
      byId[id].mediasTotal   += medias;
      byId[id].vaciasTotal   += vacias;
      byId[id].preingresos.push({
        idPreingreso: pi.idPreingreso,
        proveedor:    provMap[String(pi.idProveedor)] || pi.idProveedor || '',
        carretas:     carretas,
        estados:      estados,
        llenas: llenas, medias: medias, vacias: vacias,
        estado:       String(pi.estado || '')
      });
    });
  });

  var cargadores = Object.keys(byId).map(function(k) { return byId[k]; })
    .sort(function(a, b) { return b.carretasTotal - a.carretasTotal; });

  var totalCarretas = cargadores.reduce(function(s, c) { return s + c.carretasTotal; }, 0);
  var totalLlenas   = cargadores.reduce(function(s, c) { return s + c.llenasTotal;   }, 0);
  var totalMedias   = cargadores.reduce(function(s, c) { return s + c.mediasTotal;   }, 0);
  var totalVacias   = cargadores.reduce(function(s, c) { return s + c.vaciasTotal;   }, 0);

  return {
    ok: true,
    data: {
      fecha:         fechaStr,
      cargadores:    cargadores,
      totalCarretas: totalCarretas,
      totalLlenas:   totalLlenas,
      totalMedias:   totalMedias,
      totalVacias:   totalVacias,
      preingresos:   preingresosDelDia.length
    }
  };
}

// Imprime el consolidado del día en la impresora ALMACEN (TICKET role).
function imprimirCargadoresDia(params) {
  var fechaStr = String(params && params.fecha || '').trim();
  var resumen  = getCargadoresDelDia({ fecha: fechaStr });
  if (!resumen.ok) return resumen;
  var d = resumen.data;
  if (!d.cargadores.length) return { ok: false, error: 'Sin cargadores ese día' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  var printerId;
  if (params.printerIdOverride) {
    printerId = String(params.printerIdOverride).trim();
  } else {
    try { printerId = getPrinterNodeId('TICKET', 'ALMACEN'); }
    catch(e) { return { ok: false, error: e.message }; }
  }

  // Etiqueta de fecha amigable
  var fechaLabel = d.fecha;
  try {
    var fp = new Date(d.fecha + 'T12:00:00');
    var meses = ['enero','febrero','marzo','abril','mayo','junio',
                 'julio','agosto','septiembre','octubre','noviembre','diciembre'];
    fechaLabel = fp.getDate() + ' ' + meses[fp.getMonth()] + ' ' + fp.getFullYear();
  } catch(e) {}

  var B = [];
  function b1(v)   { B.push(v & 0xff); }
  function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
  function bLn(s)  { bStr(s); b1(0x0a); }
  var SEP  = '================================================';
  var SEP2 = '------------------------------------------------';

  b1(0x1b); b1(0x40);

  // Header
  b1(0x1b); b1(0x61); b1(0x01);
  b1(0x1b); b1(0x21); b1(0x38);
  bLn('CARGADORES');
  b1(0x1b); b1(0x21); b1(0x00);
  b1(0x1b); b1(0x45); b1(0x01);
  bLn(fechaLabel.toUpperCase());
  b1(0x1b); b1(0x45); b1(0x00);
  b1(0x1b); b1(0x61); b1(0x00);
  bLn(SEP);

  // [v2.13] Sin tarifa. Resumen + estados por carreta. El cargador
  // pone su precio en caja directamente.
  bLn(_padLine48('Preingresos del dia:', String(d.preingresos)));
  bLn(SEP2);

  // Helper: texto de estados "(2 medias, 1 casi vacia)" o "" si todo lleno
  function _txtEstados(llenas, medias, vacias) {
    if (medias === 0 && vacias === 0) return '';
    var parts = [];
    if (medias > 0) parts.push(medias + (medias === 1 ? ' media' : ' medias'));
    if (vacias > 0) parts.push(vacias + (vacias === 1 ? ' casi vacia' : ' casi vacias'));
    return '(' + parts.join(', ') + ')';
  }

  // Por cargador: nombre · N carretas · estados de carga
  d.cargadores.forEach(function(c) {
    b1(0x1b); b1(0x45); b1(0x01);
    bLn(c.nombre.toUpperCase());
    b1(0x1b); b1(0x45); b1(0x00);
    var resumenTxt = _txtEstados(c.llenasTotal, c.mediasTotal, c.vaciasTotal);
    bLn(_padLine48('  ' + c.carretasTotal + ' carretas', resumenTxt || 'todas llenas'));
    c.preingresos.forEach(function(pi) {
      var lbl = '  - ' + pi.idPreingreso + ' ' + (pi.proveedor || '').substring(0, 22);
      var t = _txtEstados(pi.llenas, pi.medias, pi.vacias);
      bLn(_padLine48(lbl, pi.carretas + ' cart ' + (t || 'OK')));
    });
    bLn('');
  });

  bLn(SEP);
  // Total general bien grande — solo carretas, sin monto
  b1(0x1b); b1(0x61); b1(0x01);
  b1(0x1b); b1(0x45); b1(0x01);
  bLn('TOTAL DEL DIA');
  b1(0x1b); b1(0x45); b1(0x00);
  b1(0x1b); b1(0x21); b1(0x38); b1(0x1b); b1(0x45); b1(0x01);
  bLn(String(d.totalCarretas) + ' carretas');
  b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
  // Desglose de estados si hay medias o vacias
  if (d.totalMedias > 0 || d.totalVacias > 0) {
    var dets = [];
    if (d.totalLlenas > 0) dets.push(d.totalLlenas + ' llenas');
    if (d.totalMedias > 0) dets.push(d.totalMedias + ' medias');
    if (d.totalVacias > 0) dets.push(d.totalVacias + ' casi vacias');
    bLn(dets.join(' / '));
  }
  b1(0x1b); b1(0x61); b1(0x00);
  bLn(SEP);

  // Feed + corte
  b1(0x1b); b1(0x4a); b1(160);
  b1(0x1d); b1(0x56); b1(0x00);

  var blob = Utilities.newBlob(B, 'application/octet-stream');
  var b64  = Utilities.base64Encode(blob.getBytes());

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:  'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type':  'application/json'
      },
      payload: JSON.stringify({
        printerId:   parseInt(printerId),
        title:       'Cargadores ' + d.fecha,
        contentType: 'raw_base64',
        content:     b64,
        source:      'warehouseMos'
      }),
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 201
      ? { ok: true, data: { fecha: d.fecha, totalMonto: d.totalMonto, totalCarretas: d.totalCarretas } }
      : { ok: false, error: 'PrintNode respuesta ' + resp.getResponseCode() };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
