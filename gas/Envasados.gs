// ============================================================
// warehouseMos — Envasados.gs
// Gestión de envasado + generación automática de guías
// + integración PrintNode para etiquetas adhesivas
// ============================================================

function getEnvasados(params) {
  var rows = _sheetToObjects(getSheet('ENVASADOS'));
  if (params.estado)     rows = rows.filter(function(r){ return String(r.estado) === String(params.estado); });
  if (params.fecha)      rows = rows.filter(function(r){ return r.fecha === params.fecha; });
  if (params.fechaDesde) rows = rows.filter(function(r){ return String(r.fecha) >= String(params.fechaDesde); });
  if (params.limit)      rows = rows.slice(0, parseInt(params.limit));

  // [v2.10.5] Enriquecer con descripciones legibles del derivado y del base.
  // El cache de productos del cliente no siempre tiene todos los códigos
  // resueltos (especialmente granel/canónicos). Devolverlas desde acá
  // evita que la UI muestre solo códigos como "WHACXOVO500GR".
  try {
    var prods = _sheetToObjects(getProductosSheet());
    var mapPorCb = {};
    var mapPorSku = {};
    var mapPorId  = {};
    prods.forEach(function(p) {
      var desc = String(p.descripcion || p.nombre || p.idProducto || '');
      if (!desc) return;
      if (p.codigoBarra) mapPorCb[String(p.codigoBarra)]  = desc;
      if (p.skuBase)     mapPorSku[String(p.skuBase)]    = desc;
      if (p.idProducto)  mapPorId[String(p.idProducto)]  = desc;
    });
    function _resolverDesc(codigo) {
      var k = String(codigo || '');
      if (!k) return '';
      return mapPorCb[k] || mapPorSku[k] || mapPorId[k] || '';
    }
    rows = rows.map(function(r) {
      r.descripcionProductoEnvasado = _resolverDesc(r.codigoProductoEnvasado);
      r.descripcionProductoBase     = _resolverDesc(r.codigoProductoBase);
      return r;
    });
  } catch(e) {
    Logger.log('getEnvasados enrich error: ' + e.message);
  }

  return { ok: true, data: rows };
}

function getPendientesEnvasado() {
  var productos = _sheetToObjects(getProductosSheet());
  var stock     = _sheetToObjects(getSheet('STOCK'));

  var stockMap = {};
  stock.forEach(function(s){ stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });

  return { ok: true, data: _calcularPendientesEnvasado(productos, stockMap) };
}

// ============================================================
// registrarEnvasado — corazón del módulo
// Genera automáticamente:
//   1. GuíaSalida (descuenta base)
//   2. GuíaIngreso (agrega derivado)
//   3. Registro en ENVASADOS
//   4. Imprime etiquetas adhesivas vía PrintNode
// ============================================================
function registrarEnvasado(params) {
  return _conLock('registrarEnvasado', function() {
    return _registrarEnvasadoImpl(params);
  });
}

function _registrarEnvasadoImpl(params) {
  var codigoBarra      = String(params.codigoBarra || '').trim();
  var unidadesReales   = parseInt(params.unidadesProducidas) || 0;
  var fechaVencimiento = params.fechaVencimiento || '';
  var usuario          = params.usuario || 'sistema';
  var imprimirEtiq     = params.imprimirEtiquetas !== false;
  var idempotencyKey   = String(params.idempotencyKey || '').trim();

  if (!codigoBarra || unidadesReales <= 0) {
    return { ok: false, error: 'Faltan datos: codigoBarra, unidadesProducidas' };
  }

  // ── IDEMPOTENCIA ─────────────────────────────────────────
  // Bug histórico: doble-click + reintento por timeout creaban registros
  // duplicados en ENVASADOS (mismo usuario + producto + cantidad, separados
  // por 3-16s) inflando stock derivado y duplicando consumo base.
  //   - Si llega params.idempotencyKey y ya vimos esa key → retornar el
  //     idEnvasado existente sin re-ejecutar.
  //   - Si no llega, fingerprint = usuario + codigoBarra + unidades + minuto.
  // TTL 120s — cubre reintentos por timeout y dobles clicks. Igual patrón
  // que crearDespachoRapido (v2.3.0).
  var cache = CacheService.getScriptCache();
  var keyEfectiva = idempotencyKey;
  if (!keyEfectiva) {
    var minuto = Math.floor(Date.now() / 60000);
    keyEfectiva = 'env_' + usuario + '_' + codigoBarra + '_' + unidadesReales + '_' + minuto;
    if (keyEfectiva.length > 240) {
      keyEfectiva = 'env_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyEfectiva)
        .map(function(b){ return (b < 0 ? b + 256 : b).toString(16); }).join('');
    }
  } else {
    keyEfectiva = 'env_' + keyEfectiva;
  }
  try {
    var prev = cache.get(keyEfectiva);
    if (prev) {
      Logger.log('registrarEnvasado idempotente: hit cache → ' + prev);
      return { ok: true, data: { idEnvasado: prev, dedup: true }, dedup: true };
    }
  } catch(_) {}

  var productos = _sheetToObjects(getProductosSheet());

  // 1. Buscar derivado por codigoBarra
  var prodDerivado = productos.find(function(p) {
    return String(p.codigoBarra).trim() === codigoBarra;
  });
  if (!prodDerivado) return { ok: false, error: 'Producto envasado no encontrado: ' + codigoBarra };

  // 2. Buscar base por skuBase === codigoProductoBase del derivado (fallback idProducto)
  var claveBase = String(prodDerivado.codigoProductoBase || '').trim();
  if (!claveBase) return { ok: false, error: 'El producto no tiene codigoProductoBase configurado' };

  var prodBase = productos.find(function(p) {
    return String(p.skuBase).trim() === claveBase || String(p.idProducto).trim() === claveBase;
  });
  if (!prodBase) return { ok: false, error: 'Producto base no encontrado: ' + claveBase };

  // 3. Calcular cantidad base consumida: unidades × factorConversionBase
  var factorBase = parseFloat(prodDerivado.factorConversionBase) || 0;
  if (factorBase <= 0) {
    return { ok: false, error: 'factorConversionBase no configurado para: ' + prodDerivado.descripcion };
  }
  var cantBase = unidadesReales * factorBase;

  var fecha      = new Date();
  var idEnvasado = _generateId('ENV');
  var idLote     = _generateId('LOT');

  // Reservar la idempotency key INMEDIATAMENTE — cualquier retry concurrent
  // que llegue antes de que terminen las 4 operaciones recibe este idEnvasado
  // y no ejecuta nada. TTL 120s.
  try { cache.put(keyEfectiva, idEnvasado, 120); } catch(_){}

  // 5. Guía SALIDA_ENVASADO del día — reutilizar si ya existe ABIERTA hoy
  var gsRes = _getOCrearGuiaDia('SALIDA_ENVASADO', usuario);
  if (!gsRes.ok) return { ok: false, error: 'Error guía salida: ' + gsRes.error };

  var detSalida = agregarDetalleGuia({
    idGuia:           gsRes.data.idGuia,
    codigoProducto:   prodBase.codigoBarra,
    cantidadEsperada: cantBase,
    cantidadRecibida: cantBase,
    precioUnitario:   0
  });
  if (!detSalida.ok) return { ok: false, error: 'Detalle salida: ' + detSalida.error };
  _actualizarStock(prodBase.codigoBarra, -cantBase, {
    tipoOperacion: 'ENVASADO_BASE',
    origen:        idEnvasado,
    usuario:       String(usuario || ''),
    observacion:   'consumo base ' + unidadesReales + ' uds'
  });

  // 6. Guía INGRESO_ENVASADO del día — reutilizar si ya existe ABIERTA hoy
  var giRes = _getOCrearGuiaDia('INGRESO_ENVASADO', usuario);
  if (!giRes.ok) return { ok: false, error: 'Error guía ingreso: ' + giRes.error };

  var detIngreso = agregarDetalleGuia({
    idGuia:           giRes.data.idGuia,
    codigoProducto:   prodDerivado.codigoBarra,
    cantidadEsperada: unidadesReales,
    cantidadRecibida: unidadesReales,
    precioUnitario:   0,
    idLote:           idLote,
    fechaVencimiento: fechaVencimiento
  });
  if (!detIngreso.ok) return { ok: false, error: 'Detalle ingreso: ' + detIngreso.error };
  _actualizarStock(prodDerivado.codigoBarra, unidadesReales, {
    tipoOperacion: 'ENVASADO_DERIVADO',
    origen:        idEnvasado,
    usuario:       String(usuario || ''),
    observacion:   'producción ' + unidadesReales + ' uds'
  });

  // 7. Registro en hoja ENVASADOS
  getSheet('ENVASADOS').appendRow([
    idEnvasado,
    prodBase.codigoBarra || prodBase.idProducto,
    cantBase,
    prodBase.unidad,
    prodDerivado.codigoBarra || prodDerivado.idProducto,
    unidadesReales,
    unidadesReales,
    0,
    100,
    fecha,
    usuario,
    'COMPLETADO',
    gsRes.data.idGuia,
    giRes.data.idGuia,
    ''
  ]);

  // 8. Imprimir etiquetas
  var resultImpresion = null;
  if (imprimirEtiq) {
    resultImpresion = _imprimirEtiquetasEnvasado({
      codigoDerivado:   prodDerivado.idProducto,
      descripcion:      prodDerivado.descripcion,
      codigoBarra:      prodDerivado.codigoBarra,
      cantidad:         prodDerivado.unidad,
      unidades:         unidadesReales,
      fechaVencimiento: fechaVencimiento,
      idLote:           idLote
    });
  }

  return {
    ok: true,
    data: {
      idEnvasado:        idEnvasado,
      idGuiaSalida:      gsRes.data.idGuia,
      idGuiaIngreso:     giRes.data.idGuia,
      cantidadBase:      cantBase,
      unidadesProducidas: unidadesReales,
      idLote:            idLote,
      impresion:         resultImpresion
    }
  };
}

// ============================================================
// IMPRESION DE ETIQUETAS ADHESIVAS — TSPL2 para TSC TTP-244CE
// 50x25mm termica directa · Logo bitmap "Caserito Tony's" +
// smart highlight de palabras diferenciadoras + word-wrap a 2 lineas
// ============================================================
function _getPrintNodeProps() {
  var props = PropertiesService.getScriptProperties();
  return {
    apiKey:           props.getProperty('PRINTNODE_API_KEY')    || '',
    printerEtiquetas: props.getProperty('PRINTER_ETIQUETAS_ID') || '',
    printerTickets:   props.getProperty('PRINTER_TICKETS_ID')   || ''
  };
}

// ── Logo bitmap (180x36 dots = 23 bytes/row x 36) ──
// Generado con Python+Pillow (Pacifico script + Lilita One bold + casita)
// Ver preview-etiquetas/gen-logo.py y logo-preview-4x.png
var LOGO_W_BYTES = 23;
var LOGO_H = 36;
var LOGO_TSPL_HEX =
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFF0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0FFFF5FFFFFFFFF' +
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0FFFE0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'F0FFFC07FFFFFFFFFFFFFFFFFC7FFFFFFFFFFFFFFFFFFFF0FFF803FFFFF03FFFFFFFFFF87FFF' +
'FFFFFFFFFFFFFFFFF0FFF041FFFFE01FFFFFFFFFF87FFFFFFFFFFFFFFFFFFFF0FFE0E01FFFC7' +
'1FFFFFFFFFF27FFFFFFFFFFFFFFFFFFFF0FFC1F01FFF8F9FFFFFFFFC727FFFFFFFFFFFFFFFFF' +
'FFF0FF03F81FFF1F9FFFFFFFFC627FFFFFFFFFFFFFFFFFFFF0FE07FC0FFF1F9FFFFFFFFC64FF' +
'FFFFFFFFFFFFFFFFFFF0FC1FFC07FF3FFFFFFFFFFFE1FFFFFFFFFFFFFFFFFFFFF0F83FFC03FE' +
'3FFE7F3E198C8021FFFFFFFFFFFFFFFFFFF0F07FFC01FE3FF81E3C190C8040FFFFFFFFFFFFFF' +
'FFFFF0E0FFFC00FE3FF89E789008E788FFFFFFFFFFFFFFFFFFF0C1FFFC107E3FF19C399298E7' +
'981FFFFFFFFFFFFFFFFFF0E2000008FF3FF31D191081E79C1FFFFFFFFFFFFFFFFFF0F600000D' +
'FF1FF3199871C1C79CFFFFFFFFFFFFFFFFFFF0FE7FFFCFFF87C31200E1F98318FFFFFFFFFFFF' +
'FFFFFFF0FE7FFFCFFFC000040001F83001FFFFFFFFFFFFFFFFFFF0FE7FFFCFFFE0388E0E13FC' +
'7843FFFFFFFFFFFFFFFFFFF0FE7FFFCFFE00F83E1C41C0783FFFFFFFFFFFFFFFFFFFF0FE7FFF' +
'CFFE00F00E0C6184701FFFFFFFFFFFFFFFFFFFF0FE7C07CFFE00E00E0C6184603FFFFFFFFFFF' +
'FFFFFFFFF0FE7C07CFFFC7C38604700CE1FFFFFFFFFFFFFFFFFFFFF0FE7C07CFFFC7C7C60470' +
'0CE0FFFFFFFFFFFFFFFFFFFFF0FE7C07CFFFC7C7C600781FE03FFFFFFFFFFFFFFFFFFFF0FE7C' +
'07CFFFC7C7C600783FF01FFFFFFFFFFFFFFFFFFFF0FE7C07CFFFC7C7C6207C3FFC1FFFFFFFFF' +
'FFFFFFFFFFF0FE7C07CFFFC7C386207C3FE61FFFFFFFFFFFFFFFFFFFF0FE7C07CFFFC7E00E30' +
'7C3FE01FFFFFFFFFFFFFFFFFFFF0FE00000FFFC7F01E307C3FE03FFFFFFFFFFFFFFFFFFFF0FE' +
'00000FFFC7F83E387C3FF07FFFFFFFFFFFFFFFFFFFF0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFF0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0';

// ── Helpers de texto ──
function _normalizeEtq(s) {
  if (s === null || s === undefined) return '';
  // NFD + strip diacriticos: "CAÑIHUA" → "CANIHUA", "CAFÉ" → "CAFE"
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function _calcVencimientoEtq(fechaEnvasado) {
  var d = fechaEnvasado ? new Date(fechaEnvasado) : new Date();
  d.setFullYear(d.getFullYear() + 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yy');
}

// Detecta tokens diferenciadores: comparando con productos que comparten
// el PRIMER token (minPrefix=1). Siempre destaca el ultimo token (peso).
function _detectHighlightsEtq(targetTokens, allTokenized) {
  var hl = {};
  hl[targetTokens.length - 1] = true;
  for (var i = 0; i < allTokenized.length; i++) {
    var s = allTokenized[i];
    if (s === targetTokens) continue;
    if (s[0] !== targetTokens[0]) continue;
    for (var pos = 1; pos < targetTokens.length; pos++) {
      if (hl[pos]) continue;
      var prior = true;
      for (var k = 0; k < pos; k++) {
        if (s[k] !== targetTokens[k]) { prior = false; break; }
      }
      if (prior && s[pos] !== undefined && s[pos] !== targetTokens[pos]) {
        hl[pos] = true;
        break;
      }
    }
  }
  var out = [];
  for (var key in hl) if (hl[key]) out.push(parseInt(key));
  out.sort(function(a,b){ return a - b; });
  return out;
}

// Font 3 normal (16 wide x 24 tall), Font 4 highlight (24 wide x 32 tall)
function _fontWidthEtq(isHighlight) { return isHighlight ? 24 : 16; }

// Word-wrap a max 2 lineas. Si todo no cabe en 1, intenta cortar ANTES del
// primer highlight (asi los diferenciadores quedan juntos en linea 2).
function _wrapTokensEtq(tokens, highlights) {
  var MAX_W = 370, SPACE = 8;
  var widths = tokens.map(function(t, i) {
    return t.length * _fontWidthEtq(highlights.indexOf(i) >= 0);
  });
  var isHl = function(i) { return highlights.indexOf(i) >= 0; };

  // Total en 1 linea
  var total = 0;
  for (var i = 0; i < widths.length; i++) total += widths[i] + (i > 0 ? SPACE : 0);
  if (total <= MAX_W) {
    return [tokens.map(function(t, i) { return { tok: t, hl: isHl(i), w: widths[i] }; })];
  }

  // Intentar wrap antes del primer highlight (preserva el grupo destacado)
  var firstHl = highlights.length > 0 ? highlights[0] : tokens.length;
  if (firstHl > 0 && firstHl < tokens.length) {
    var l1 = [], w1 = 0;
    for (var a = 0; a < firstHl; a++) {
      w1 += widths[a] + (l1.length > 0 ? SPACE : 0);
      l1.push({ tok: tokens[a], hl: false, w: widths[a] });
    }
    var l2 = [], w2 = 0;
    for (var b = firstHl; b < tokens.length; b++) {
      w2 += widths[b] + (l2.length > 0 ? SPACE : 0);
      l2.push({ tok: tokens[b], hl: isHl(b), w: widths[b] });
    }
    if (w1 <= MAX_W && w2 <= MAX_W) return [l1, l2];
  }

  // Fallback: greedy por palabra
  var lines = [[]], curW = 0;
  for (var c = 0; c < tokens.length; c++) {
    var sep = lines[lines.length - 1].length === 0 ? 0 : SPACE;
    if (curW + sep + widths[c] <= MAX_W) {
      lines[lines.length - 1].push({ tok: tokens[c], hl: isHl(c), w: widths[c] });
      curW += sep + widths[c];
    } else if (lines.length === 1) {
      lines.push([{ tok: tokens[c], hl: isHl(c), w: widths[c] }]);
      curW = widths[c];
    } else {
      lines[1].push({ tok: tokens[c], hl: isHl(c), w: widths[c] });
      curW += sep + widths[c];
    }
  }
  return lines;
}

function _hexToBytesEtq(hex) {
  var arr = [];
  for (var i = 0; i < hex.length; i += 2) arr.push(parseInt(hex.substr(i, 2), 16));
  return arr;
}

function _strToBytesEtq(s) {
  var arr = [];
  for (var i = 0; i < s.length; i++) arr.push(s.charCodeAt(i) & 0xFF);
  return arr;
}

// Lista de envasables tokenizados (cacheable). Filtro: derivados activos WH.
function _getAllEnvasablesTokens() {
  var rows = _sheetToObjects(getProductosSheet());
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var p = rows[i];
    var estadoNorm = (p.estado === true) ? '1' : (p.estado === false) ? '0' : String(p.estado);
    if (estadoNorm !== '1' || !p.codigoProductoBase) continue;
    var c = String(p.codigoBarra || p.idProducto || '').toUpperCase();
    if (c.indexOf('WH') !== 0) continue;
    out.push({
      codigoBarra: p.codigoBarra || p.idProducto,
      descripcion: p.descripcion,
      tokens: _normalizeEtq(p.descripcion).split(/\s+/)
    });
  }
  return out;
}

// Construye los bytes TSPL2 completos (header + bitmap + texto + barcode + PRINT)
function _buildTSPLEtq(producto, fechaEnvasado, unidades, allEnvasables) {
  var descNorm = _normalizeEtq(producto.descripcion);
  var tokens = descNorm.split(/\s+/);
  var allTok = allEnvasables.map(function(p) { return p.tokens; });
  var highlights = _detectHighlightsEtq(tokens, allTok);
  var lines = _wrapTokensEtq(tokens, highlights);
  var vto = _calcVencimientoEtq(fechaEnvasado);

  // [Calibración configurable vía Script Properties]
  //   ADHESIVO_GAP_MM   default 2  (mm de separación entre adhesivos)
  //   ADHESIVO_DENSITY  default 8  (1-15, sube si sale tenue)
  //   ADHESIVO_SPEED    default 4  (1-6, baja si sale corrida)
  //   ADHESIVO_OFFSET_Y default 0  (dots: positivo baja TODO el contenido, negativo lo sube)
  // Si la impresión sale corrida arriba/abajo del adhesivo, ajustá OFFSET_Y.
  var props = PropertiesService.getScriptProperties();
  var gapMm    = parseFloat(props.getProperty('ADHESIVO_GAP_MM'))   || 2;
  var density  = parseInt(props.getProperty('ADHESIVO_DENSITY'))    || 8;
  var speed    = parseInt(props.getProperty('ADHESIVO_SPEED'))      || 4;
  var offsetY  = parseInt(props.getProperty('ADHESIVO_OFFSET_Y'))   || 0;

  var header = [
    'SIZE 50 mm,25 mm',
    'GAP ' + gapMm + ' mm,0 mm',
    'DIRECTION 1',
    'DENSITY ' + density,
    'SPEED ' + speed,
    'CLS',
    'BITMAP 5,' + (2 + offsetY) + ',' + LOGO_W_BYTES + ',' + LOGO_H + ',0,'
  ].join('\r\n');

  var bytes = _strToBytesEtq(header);
  bytes = bytes.concat(_hexToBytesEtq(LOGO_TSPL_HEX));
  bytes = bytes.concat(_strToBytesEtq('\r\n'));

  // [Offset Y aplica a TODO el contenido (texto, separador, barcode)]
  // Vto top-right
  bytes = bytes.concat(_strToBytesEtq('TEXT 280,' + (12 + offsetY) + ',"2",0,1,1,"Vto ' + vto + '"\r\n'));
  // Separador
  bytes = bytes.concat(_strToBytesEtq('BAR 5,' + (42 + offsetY) + ',390,1\r\n'));

  // Descripcion (1-2 lineas con highlights)
  var startY = 46 + offsetY, LINE_H = 38, SPACE = 8;
  for (var li = 0; li < lines.length; li++) {
    var line = lines[li];
    var totalW = 0;
    for (var ti = 0; ti < line.length; ti++) totalW += line[ti].w + (ti > 0 ? SPACE : 0);
    var x = Math.max(5, Math.round((400 - totalW) / 2));
    var y = startY + li * LINE_H;
    for (var tj = 0; tj < line.length; tj++) {
      var o = line[tj];
      var font = o.hl ? '4' : '3';
      var yAdj = o.hl ? y : y + 4;  // baseline align cuando se mezclan tamaños
      var safe = String(o.tok).replace(/"/g, "'");
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + x + ',' + yAdj + ',"' + font + '",0,1,1,"' + safe + '"\r\n'));
      x += o.w + SPACE;
    }
  }

  // [v2.13.96] Cálculo EXACTO del width Code128 + fallback automático a narrow=1
  //
  // Fórmula real Code128:
  //   modules = 11*bcLen + 35  (Start 11 + chars 11×bcLen + Check 11 + Stop 13)
  //   width   = modules × narrow_dots
  //   Con narrow=2 → 22*bcLen + 70
  //   Con narrow=1 → 11*bcLen + 35
  //
  // Bug previo (v2.13.95): subestimaba con 22*bcLen+50 → barcode se salía
  // del adhesivo para bcLen≥15. Aunque pretendía "centrar y dar quiet zones",
  // la fórmula incorrecta hacía que el cálculo de X estuviera mal.
  //
  // Fix: cálculo exacto + si width > 360 dots (no entra cómodo con narrow=2),
  // bajar automático a narrow=1. Code128 a 203dpi con narrow=1 (0.125mm)
  // es legible para cualquier scanner razonable (los baratos también).
  var bc = String(producto.codigoBarra || '').replace(/"/g, '');
  var bcLen = bc.length;
  var modules = 11 * bcLen + 35;
  var narrowBc = 2;
  var barcodeWidth = modules * narrowBc;
  if (barcodeWidth > 360) {
    narrowBc = 1;
    barcodeWidth = modules * narrowBc; // ahora más compacto
  }
  var barcodeHeight = 44;
  var barcodeX = Math.max(20, Math.floor((400 - barcodeWidth) / 2));
  var barcodeY = 128 + offsetY;
  var barcodeEndX = barcodeX + barcodeWidth;
  // Flecha izquierda (font 3 = 16w × 24h) — X=5 para margen físico seguro al borde
  // Solo si la quiet zone izquierda permite tener flecha + 5 dots de gap
  if (barcodeX - 16 >= 7) {
    bytes = bytes.concat(_strToBytesEtq('TEXT 5,' + (barcodeY + 10) + ',"3",0,1,1,">"\r\n'));
  }
  // Barcode
  bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',1,0,' + narrowBc + ',' + narrowBc + ',"' + bc + '"\r\n'));
  // Flecha derecha — necesita ≥21 dots de espacio (5 gap + 16 width) sin tocar borde
  if (barcodeEndX + 21 <= 395) {
    bytes = bytes.concat(_strToBytesEtq('TEXT ' + (barcodeEndX + 5) + ',' + (barcodeY + 10) + ',"3",0,1,1,"<"\r\n'));
  }

  // Print N copias
  bytes = bytes.concat(_strToBytesEtq('PRINT ' + (unidades || 1) + ',1\r\n'));

  return bytes;
}

function _imprimirEtiquetasEnvasado(data) {
  // API key sigue en Script Properties (igual que todo el ecosistema)
  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurado' };

  // PrinterId: si vino override (modal admin/master) se usa, sino se busca
  // en la hoja IMPRESORAS centralizada de MOS (mismo patrón que imprimirBienvenida,
  // imprimirTicketGuia, imprimirHistorialStock). Tipo 'ADHESIVO', zona 'ALMACEN'.
  var printerId;
  if (data.printerIdOverride) {
    printerId = String(data.printerIdOverride).trim();
  } else {
    try { printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN'); }
    catch (e) { return { ok: false, error: e.message }; }
  }

  if (!data.codigoBarra || !data.descripcion) {
    return { ok: false, error: 'Falta codigoBarra o descripcion' };
  }

  var allEnv = _getAllEnvasablesTokens();
  // fechaVencimiento explicita (legacy) tiene prioridad sobre el calculo +1 año
  var fechaParaCalcular = data.fechaVencimiento
    ? null  // si vino fechaVenc explicita usaremos esa abajo
    : (data.fechaEnvasado || data.fechaImpresion || new Date());

  // Si nos pasaron fechaVencimiento explicita la usamos como ENVASADO menos 1 año
  // para que _calcVencimientoEtq la devuelva igual. Mas simple: override _calc.
  var vtoOverride = null;
  if (data.fechaVencimiento) {
    vtoOverride = Utilities.formatDate(
      new Date(data.fechaVencimiento),
      Session.getScriptTimeZone(),
      'dd/MM/yy'
    );
  }

  // Si vto override: monkey-patch via wrapper
  var bytes;
  if (vtoOverride) {
    // Engineamos fechaEnvasado = vto - 1 año para que el calc devuelva el override
    var d = new Date(data.fechaVencimiento);
    d.setFullYear(d.getFullYear() - 1);
    bytes = _buildTSPLEtq(data, d, data.unidades || 1, allEnv);
  } else {
    bytes = _buildTSPLEtq(data, fechaParaCalcular, data.unidades || 1, allEnv);
  }

  var base64 = Utilities.base64Encode(bytes);

  var payload = {
    printerId: parseInt(printerId),
    title: 'Etiqueta ' + _normalizeEtq(data.descripcion),
    contentType: 'raw_base64',
    content: base64,
    source: 'warehouseMos'
  };

  try {
    var response = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':'),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 201) {
      return { ok: true, jobId: JSON.parse(response.getContentText()), unidades: data.unidades };
    } else {
      return { ok: false, error: response.getContentText(), httpCode: code };
    }
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function imprimirEtiqueta(params) {
  return _imprimirEtiquetasEnvasado(params);
}

// ════════════════════════════════════════════════════════════════════
// Calibrar impresora ADHESIVO/ALMACEN — manda GAPDETECT (auto-calibrar
// sensor de gap al rollo actual) seguido de FORMFEED (alinea al próximo
// inicio de adhesivo). Recomendado ejecutar:
//   - Una vez al cambiar el rollo de etiquetas
//   - Si las impresiones empiezan a salir corridas/desalineadas
//
// El GAPDETECT consume ~3 etiquetas en blanco mientras la impresora mide
// el sensor de gap. Después de eso, cada PRINT siguiente cae perfecto en
// la zona del adhesivo (siempre que el SIZE/GAP del TSPL coincida con
// las medidas reales del rollo).
// ════════════════════════════════════════════════════════════════════
function calibrarImpresoraAdhesivo() {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var printerId;
    try { printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN'); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }

    // TSPL puro de calibración:
    //   SIZE/GAP fijan las dimensiones esperadas del rollo
    //   GAPDETECT manda al sensor a remedir
    //   CLS limpia el buffer
    //   FORMFEED avanza al próximo gap (alinea físicamente la próxima impresión)
    var tspl =
      'SIZE 50 mm,25 mm\r\n' +
      'GAP 2 mm,0 mm\r\n' +
      'DIRECTION 1\r\n' +
      'CLS\r\n' +
      'GAPDETECT\r\n' +
      'FORMFEED\r\n';
    var b64 = Utilities.base64Encode(tspl);
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':') },
      contentType: 'application/json',
      payload: JSON.stringify({
        printerId: parseInt(printerId),
        title: 'Calibrar ADHESIVO (' + new Date().toISOString() + ')',
        contentType: 'raw_base64',
        content: b64,
        source: 'warehouseMos-calibrar'
      }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 201) return { ok: false, error: 'PrintNode HTTP ' + code + ': ' + resp.getContentText() };
    return { ok: true, data: {
      jobId: JSON.parse(resp.getContentText()),
      mensaje: 'Calibración enviada. La impresora va a avanzar ~3 etiquetas en blanco mientras mide el sensor de gap. Después ya podés imprimir normal.'
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════
// Estado de la impresora ADHESIVO/ALMACEN — para indicador 🟢/🔴 en MOS.
// Llamado desde el modal de impresión de adhesivos antes de habilitar
// el botón Imprimir. Si está offline → bloquea con tooltip.
// ════════════════════════════════════════════════════════════════════
function estadoImpresoraAdhesivo() {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var printerId;
    try { printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN'); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }
    var auth = 'Basic ' + Utilities.base64Encode(apiKey + ':');
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printers/' + printerId, {
      headers: { 'Authorization': auth },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      return { ok: false, error: 'PrintNode HTTP ' + resp.getResponseCode(), printerId: printerId };
    }
    var pj = JSON.parse(resp.getContentText());
    var p = Array.isArray(pj) ? pj[0] : pj;
    return { ok: true, data: {
      printerId:    printerId,
      nombre:       p.name,
      estado:       p.state,                          // 'online' | 'offline'
      esOnline:     String(p.state).toLowerCase() === 'online',
      computadora:  p.computer && p.computer.name,
      compEstado:   p.computer && p.computer.state
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// PREVIEW EN LOG (sin imprimir) — corre desde el editor
// Uso: cambia el codigoBarra debajo y dale ▶ Run
// ============================================================
function previsualizarEtiqueta() {
  // ⚙ Cambia este codigo para previsualizar otro producto
  var codigoBarra = 'WHCOLAGO250GR';

  var L = function(m) { Logger.log(m); };
  L('═════ PREVIEW ETIQUETA ' + codigoBarra + ' ═════');

  var all = _getAllEnvasablesTokens();
  var prod = null;
  for (var i = 0; i < all.length; i++) {
    if (all[i].codigoBarra === codigoBarra) { prod = all[i]; break; }
  }
  if (!prod) { L('❌ No se encontro envasable con codigo ' + codigoBarra); return { ok:false }; }

  var descNorm = _normalizeEtq(prod.descripcion);
  var tokens = descNorm.split(/\s+/);
  var allTok = all.map(function(p) { return p.tokens; });
  var hl = _detectHighlightsEtq(tokens, allTok);
  var vto = _calcVencimientoEtq(new Date());

  L('Original:    ' + prod.descripcion);
  L('Normalizado: ' + descNorm);
  L('Tokens:      ' + JSON.stringify(tokens));
  L('Highlights:  ' + JSON.stringify(hl) + ' → ' +
    tokens.map(function(t, i) { return hl.indexOf(i) >= 0 ? '['+t+']' : t; }).join(' '));
  L('Vto (hoy+1a): ' + vto);

  var lines = _wrapTokensEtq(tokens, hl);
  L('Lineas:      ' + lines.length);
  for (var li = 0; li < lines.length; li++) {
    var line = lines[li];
    var txt = '  L' + (li+1) + ': ';
    for (var ti = 0; ti < line.length; ti++) {
      txt += (line[ti].hl ? '**' : '') + line[ti].tok + (line[ti].hl ? '**' : '') + ' ';
    }
    L(txt);
  }

  var bytes = _buildTSPLEtq(prod, new Date(), 1, all);
  L('TSPL2 total: ' + bytes.length + ' bytes');
  L('═════ FIN ═════');
  return { ok: true, producto: prod.descripcion, tokens: tokens, highlights: hl, vto: vto, lineas: lines.length, bytes: bytes.length };
}

// ============================================================
// TEST de impresora de etiquetas (run desde el editor)
// Verifica TODO el chain: Script Properties → API key válida →
// impresora online → envía 1 etiqueta TEST. Resultados en Logger.
// ============================================================
function testImpresoraEtiquetas() {
  var L = function(msg){ Logger.log(msg); };
  L('═══════ TEST IMPRESORA ETIQUETAS — ' + new Date().toLocaleString() + ' ═══════');

  // 1. Resolver config: API key (Script Props) + printerId (hoja IMPRESORAS de MOS)
  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  L('1. Configuración:');
  L('   PRINTNODE_API_KEY (Script Props) = ' + (apiKey ? '✓ ('+apiKey.length+' chars)' : '✗ FALTA'));
  if (!apiKey) { L('❌ Falta PRINTNODE_API_KEY en Script Properties.'); return { ok:false, paso:'apikey' }; }

  var printerId;
  try {
    printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN');
    L('   Impresora ADHESIVO/ALMACEN (hoja IMPRESORAS de MOS) = ' + printerId);
  } catch (e) {
    L('❌ No hay impresora ADHESIVO/ALMACEN activa en MOS. ' + e.message);
    L('   💡 Agrega una fila en la hoja IMPRESORAS de MOS con tipo=ADHESIVO, idZona=ALMACEN, activo=1, printNodeId=<ID>');
    return { ok:false, paso:'printer_registry' };
  }

  // 2. API key válida (whoami)
  L('2. Verificando API key con /whoami...');
  var auth = { 'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':') };
  var who;
  try {
    who = UrlFetchApp.fetch('https://api.printnode.com/whoami', { headers: auth, muteHttpExceptions: true });
    var whoCode = who.getResponseCode();
    if (whoCode !== 200) {
      L('❌ API key inválida (HTTP '+whoCode+'): ' + who.getContentText());
      return { ok:false, paso:'auth', httpCode: whoCode };
    }
    var whoJson = JSON.parse(who.getContentText());
    L('   ✓ Cuenta: ' + (whoJson.email || whoJson.firstname || '(sin email)'));
  } catch(e) {
    L('❌ Error consultando /whoami: ' + e.message); return { ok:false, paso:'auth_err' };
  }

  // 3. Estado de la impresora
  L('3. Verificando estado de la impresora #' + printerId + '...');
  var printer;
  try {
    printer = UrlFetchApp.fetch('https://api.printnode.com/printers/' + printerId, { headers: auth, muteHttpExceptions: true });
    var prCode = printer.getResponseCode();
    if (prCode !== 200) {
      L('❌ Impresora no encontrada o sin acceso (HTTP '+prCode+'): ' + printer.getContentText());
      L('   💡 Lista de tus impresoras disponibles:');
      try {
        var all = UrlFetchApp.fetch('https://api.printnode.com/printers', { headers: auth, muteHttpExceptions: true });
        var arr = JSON.parse(all.getContentText());
        arr.forEach(function(p){ L('      • #'+p.id+' — '+p.name+' ('+p.state+') @ '+(p.computer && p.computer.name)); });
      } catch(_) {}
      return { ok:false, paso:'printer', httpCode: prCode };
    }
    var pj = JSON.parse(printer.getContentText());
    var p = Array.isArray(pj) ? pj[0] : pj;
    L('   ✓ Nombre:     ' + p.name);
    L('   ✓ Estado:     ' + p.state + (p.state==='online' ? ' 🟢' : ' 🔴'));
    L('   ✓ Computador: ' + (p.computer && p.computer.name) + ' (' + (p.computer && p.computer.state) + ')');
    if (p.state !== 'online') {
      L('⚠ Impresora NO está online. Revisa: cable encendido, PrintNode Client corriendo en la PC, drivers instalados.');
    }
  } catch(e) {
    L('❌ Error consultando impresora: ' + e.message); return { ok:false, paso:'printer_err' };
  }

  // 4. Enviar test print: 1 etiqueta TSPL2 con TEST + timestamp + barcode
  // [FIX] Antes usaba ZPL (^XA, ^FO, etc.) pero el sistema productivo usa
  // TSPL2 nativo (SIZE, TEXT, BARCODE, PRINT). La TSC TTP-244CE viene
  // configurada en TSPL2 por default. El test con ZPL imprimía basura o
  // nada → daba falso OK. Ahora coincide con producción.
  L('4. Enviando etiqueta TEST (TSPL2)...');
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm:ss');
  var tspl =
    'SIZE 50 mm,25 mm\r\n' +
    'GAP 2 mm,0 mm\r\n' +
    'DIRECTION 1\r\n' +
    'DENSITY 8\r\n' +
    'SPEED 4\r\n' +
    'CLS\r\n' +
    'TEXT 10,10,"4",0,1,1,"TEST IMPRESORA"\r\n' +
    'TEXT 10,50,"3",0,1,1,"Fecha: ' + ts + '"\r\n' +
    'TEXT 10,80,"3",0,1,1,"warehouseMos / Levo.dev"\r\n' +
    'TEXT 10,110,"2",0,1,1,"Si ves esto, funciona OK!"\r\n' +
    'BARCODE 10,140,"128",50,1,0,2,2,"TEST-WH"\r\n' +
    'PRINT 1,1\r\n';

  var payload = {
    printerId:   parseInt(printerId),
    title:       'TEST WH ' + ts,
    contentType: 'raw_base64',
    content:     Utilities.base64Encode(tspl),
    source:      'warehouseMos-test'
  };

  try {
    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method:'post',
      headers: { 'Authorization': auth.Authorization, 'Content-Type':'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 201) {
      var jobId = JSON.parse(resp.getContentText());
      L('   ✓ Print job enviado — jobId: ' + jobId);
      L('═══════ ✅ TEST OK — la etiqueta debería estar saliendo ahora ═══════');
      return { ok:true, jobId: jobId, printer: printerId };
    } else {
      L('❌ PrintNode rechazó el job (HTTP '+code+'): ' + resp.getContentText());
      return { ok:false, paso:'print', httpCode: code, body: resp.getContentText() };
    }
  } catch(e) {
    L('❌ Error enviando print job: ' + e.message);
    return { ok:false, paso:'print_err' };
  }
}

function _truncate(str, max) {
  if (!str) return '';
  return str.length > max ? str.substring(0, max) : str;
}

// Devuelve la guía de ese tipo creada hoy (cualquier estado), o crea+cierra una nueva.
// Garantiza una sola guía por tipo por día; detalles se agregan a la existente.
function _getOCrearGuiaDia(tipo, usuario) {
  var tz  = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var guias = _sheetToObjects(getSheet('GUIAS'));
  for (var i = 0; i < guias.length; i++) {
    var g = guias[i];
    if (g.tipo !== tipo) continue;
    var fGuia = g.fecha ? String(g.fecha).substring(0, 10) : '';
    if (fGuia === hoy) return { ok: true, data: { idGuia: g.idGuia } };
  }
  // No existe → crear y cerrar de inmediato (sin tocar stock)
  var res = crearGuia({ tipo: tipo, usuario: usuario, comentario: 'Envasados ' + hoy });
  if (!res.ok) return res;
  _cerrarGuiaSinStock(res.data.idGuia);
  return res;
}

// ════════════════════════════════════════════════════════════════════
// ANULAR ENVASADO MANUAL — operador corrige un error de captura
// ════════════════════════════════════════════════════════════════════
// Anula un envasado individual. Reverte stock base (suma) y stock
// derivado (resta) anulando los detalles de las guías correspondientes
// con anularDetalle (que ya tiene _conLock y registra el movimiento).
// Idempotente: si ya estaba anulado retorna yaAnulado:true.
function anularEnvasadoManual(params) {
  return _conLock('anularEnvasadoManual', function() {
    var idEnv = String((params && params.idEnvasado) || '').trim();
    var usuario = String((params && params.usuario) || 'manual');
    var motivo = String((params && params.motivo) || 'corrección manual');
    if (!idEnv) return { ok: false, error: 'Requiere idEnvasado' };

    var sheet = getSheet('ENVASADOS');
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h); });
    var iId    = hdrs.indexOf('idEnvasado');
    var iBase  = hdrs.indexOf('codigoProductoBase');
    var iCnt   = hdrs.indexOf('cantidadBase');
    var iDer   = hdrs.indexOf('codigoProductoEnvasado');
    var iUds   = hdrs.indexOf('unidadesProducidas');
    var iEst   = hdrs.indexOf('estado');
    var iGs    = hdrs.indexOf('idGuiaSalida');
    var iGi    = hdrs.indexOf('idGuiaIngreso');
    var iObs   = hdrs.indexOf('observacion');
    var iFecha = hdrs.indexOf('fecha');

    var rowIdx = -1, env = null;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iId]) === idEnv) {
        rowIdx = r + 1;
        env = {
          base:    String(data[r][iBase] || ''),
          cnt:     parseFloat(data[r][iCnt]) || 0,
          der:     String(data[r][iDer] || ''),
          uds:     parseFloat(data[r][iUds]) || 0,
          estado:  String(data[r][iEst] || '').toUpperCase(),
          idGs:    String(data[r][iGs] || ''),
          idGi:    String(data[r][iGi] || ''),
          fechaTs: data[r][iFecha] instanceof Date ? data[r][iFecha].getTime() : new Date(data[r][iFecha]).getTime()
        };
        break;
      }
    }
    if (rowIdx < 2) return { ok: false, error: 'Envasado no encontrado: ' + idEnv };
    if (env.estado === 'ANULADO' || env.estado === 'ANULADO_DUPLICADO') {
      return { ok: true, data: { idEnvasado: idEnv, yaAnulado: true, estado: env.estado } };
    }

    // Buscar el detalle de SALIDA (idGuiaSalida, codigo base, cantidad)
    // y de INGRESO (idGuiaIngreso, codigo derivado, unidades). Toleramos
    // que codigoProductoBase pueda venir como skuBase o codigoBarra: si
    // no encuentro por el código directo, busco cualquier detalle no
    // anulado con esa cantidad en la guía.
    function _buscarDet(idGuia, codigo, cantidad, refTs) {
      try {
        var dets = _sheetToObjects(getSheet('GUIA_DETALLE'));
        var cand = dets.filter(function(d) {
          if (String(d.idGuia) !== idGuia) return false;
          if (String(d.observacion || '').toUpperCase() === 'ANULADO') return false;
          if ((parseFloat(d.cantidadEsperada) || 0) !== cantidad) return false;
          if (codigo && String(d.codigoProducto) !== codigo) {
            // Si el código no coincide exacto, igual lo consideramos como
            // candidato — la cantidad + idGuia ya filtran bastante.
            return true;
          }
          return true;
        });
        // Ordenar por proximidad temporal al envasado (idDetalle = DET<ts>)
        cand.sort(function(a, b) {
          var ta = parseInt((String(a.idDetalle || '').match(/(\d{10,})/) || [])[1] || 0, 10);
          var tb = parseInt((String(b.idDetalle || '').match(/(\d{10,})/) || [])[1] || 0, 10);
          return Math.abs(ta - refTs) - Math.abs(tb - refTs);
        });
        return cand[0] || null;
      } catch(_) { return null; }
    }

    var resSal = null, resIng = null;
    var detSal = _buscarDet(env.idGs, env.base, env.cnt, env.fechaTs);
    var detIng = _buscarDet(env.idGi, env.der, env.uds, env.fechaTs);
    try { if (detSal) resSal = anularDetalle({ idDetalle: detSal.idDetalle, usuario: usuario }); } catch(e1) {}
    try { if (detIng) resIng = anularDetalle({ idDetalle: detIng.idDetalle, usuario: usuario }); } catch(e2) {}

    var nowStr = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    sheet.getRange(rowIdx, iEst + 1).setValue('ANULADO');
    var prevObs = sheet.getRange(rowIdx, iObs + 1).getValue();
    sheet.getRange(rowIdx, iObs + 1).setValue(
      String(prevObs || '') + ' | ANULADO por ' + usuario + ' · ' + motivo + ' · ' + nowStr
    );

    return {
      ok: true,
      data: {
        idEnvasado:    idEnv,
        estado:        'ANULADO',
        stockBaseRevertido:    !!(resSal && resSal.ok),
        stockDerivRevertido:   !!(resIng && resIng.ok),
        cantidadBaseRevertida: env.cnt,
        unidadesRevertidas:    env.uds
      }
    };
  });
}

// ════════════════════════════════════════════════════════════════════
// SCRIPT ONE-SHOT — limpiar envasados duplicados desde una fecha
// ════════════════════════════════════════════════════════════════════
// Antes de v2.8.0 registrarEnvasado no tenía idempotencia → un click +
// retry automático por timeout creaba 2-3 registros idénticos en ENVASADOS,
// inflaba el stock derivado y duplicaba el consumo de stock base.
//
// Este script:
//   1. Lee ENVASADOS desde fechaDesde (default 2026-05-01)
//   2. Agrupa por (usuario + codigoProductoBase + cantidadBase +
//      codigoProductoEnvasado + unidadesProducidas + idGuiaSalida)
//   3. En cada grupo con >1 registro CERCANOS (< 120 s), el primero queda
//      como bueno; los demás se anulan:
//        - anularDetalle(idDetalleSalida)   ← reverte stock base
//        - anularDetalle(idDetalleIngreso)  ← reverte stock derivado
//        - ENVASADOS.estado = 'ANULADO_DUPLICADO'
//        - ENVASADOS.observacion = 'duplicado de <idOriginal> · limpieza <ts>'
//
// Para encontrar el detalle correcto en GUIA_DETALLE (puede haber varios
// del mismo producto en la guía del día), correlacionamos por idGuia +
// codigoProducto + cantidad + timestamp del idDetalle (DET<timestamp>)
// cercano a la fecha del envasado duplicado.
//
// Ejecutar UNA VEZ desde el editor de Apps Script:
//   limpiarEnvasadosDuplicados()             → desde 2026-05-01
//   limpiarEnvasadosDuplicados('2026-05-10') → desde otra fecha
// Idempotente: si ya se anuló antes, no se anula otra vez (filtro
// estado != 'ANULADO_DUPLICADO' y observacion != 'ANULADO').
function limpiarEnvasadosDuplicados(fechaDesde) {
  var FECHA_DESDE = fechaDesde || '2026-05-01';
  var VENTANA_SEG = 120; // segundos para considerar duplicado del mismo grupo
  var sheet = getSheet('ENVASADOS');
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, data: { revisados: 0, anulados: 0, msg: 'ENVASADOS vacía' } };
  var hdrs  = data[0].map(function(h){ return String(h); });
  var iId    = hdrs.indexOf('idEnvasado');
  var iBase  = hdrs.indexOf('codigoProductoBase');
  var iCnt   = hdrs.indexOf('cantidadBase');
  var iDer   = hdrs.indexOf('codigoProductoEnvasado');
  var iUds   = hdrs.indexOf('unidadesProducidas');
  var iFecha = hdrs.indexOf('fecha');
  var iUsr   = hdrs.indexOf('usuario');
  var iEst   = hdrs.indexOf('estado');
  var iGs    = hdrs.indexOf('idGuiaSalida');
  var iGi    = hdrs.indexOf('idGuiaIngreso');
  var iObs   = hdrs.indexOf('observacion');
  if (iId < 0 || iBase < 0) return { ok: false, error: 'Columnas requeridas faltantes en ENVASADOS' };

  // 1. Recolectar candidatos (desde fechaDesde, estado COMPLETADO o vacío)
  var registros = [];
  for (var r = 1; r < data.length; r++) {
    var f = data[r][iFecha];
    var fStr = f instanceof Date
      ? Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(f || '').substring(0, 10);
    if (fStr < FECHA_DESDE) continue;
    var est = String(data[r][iEst] || '').toUpperCase();
    if (est === 'ANULADO_DUPLICADO' || est === 'ANULADO') continue;
    var fTs = f instanceof Date ? f.getTime() : new Date(f).getTime();
    if (isNaN(fTs)) continue;
    registros.push({
      rowIdx: r + 1, // 1-indexado para setValue
      idEnvasado:  String(data[r][iId] || ''),
      base:        String(data[r][iBase] || ''),
      cantBase:    parseFloat(data[r][iCnt]) || 0,
      derivado:    String(data[r][iDer] || ''),
      uds:         parseFloat(data[r][iUds]) || 0,
      ts:          fTs,
      usuario:     String(data[r][iUsr] || '').trim().toLowerCase(),
      idGs:        String(data[r][iGs] || ''),
      idGi:        String(data[r][iGi] || '')
    });
  }

  // 2. Agrupar por clave
  var grupos = {};
  registros.forEach(function(reg) {
    var key = [reg.usuario, reg.base, reg.cantBase, reg.derivado, reg.uds, reg.idGs].join('|');
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(reg);
  });

  // 3. Construir índice GUIA_DETALLE por (idGuia + codigoProducto) → lista de detalles
  //    para encontrar rápido el detalle correcto sin anular dos veces.
  var detSheet = getSheet('GUIA_DETALLE');
  var detData  = detSheet.getDataRange().getValues();
  var dHdrs    = detData[0].map(function(h){ return String(h); });
  var dIdDet   = dHdrs.indexOf('idDetalle');
  var dIdGuia  = dHdrs.indexOf('idGuia');
  var dCod     = dHdrs.indexOf('codigoProducto');
  var dCantE   = dHdrs.indexOf('cantidadEsperada');
  var dObs     = dHdrs.indexOf('observacion');
  var detIndex = {};
  for (var dr = 1; dr < detData.length; dr++) {
    var idGuia  = String(detData[dr][dIdGuia] || '');
    var codProd = String(detData[dr][dCod] || '');
    var cant    = parseFloat(detData[dr][dCantE]) || 0;
    var obs     = String(detData[dr][dObs] || '').toUpperCase();
    if (obs === 'ANULADO') continue; // ya anulado, no contemplar
    var idDet   = String(detData[dr][dIdDet] || '');
    // Extraer timestamp del idDetalle (formato 'DET<ts>')
    var detTs   = 0;
    var m = idDet.match(/(\d{10,})/);
    if (m) detTs = parseInt(m[1], 10);
    var k = idGuia + '|' + codProd + '|' + cant;
    if (!detIndex[k]) detIndex[k] = [];
    detIndex[k].push({ idDet: idDet, ts: detTs, _consumido: false });
  }

  // 4. Para cada grupo: ordenar por ts, marcar duplicados, anular detalles
  var reporte = {
    fechaDesde:     FECHA_DESDE,
    gruposRevisados: 0,
    duplicadosAnulados: 0,
    detallesAnulados: 0,
    erroresAnulacion: [],
    detalle: []
  };
  var nowStr = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

  Object.keys(grupos).forEach(function(key) {
    var arr = grupos[key];
    if (arr.length < 2) return;
    arr.sort(function(a, b){ return a.ts - b.ts; });
    var primero = arr[0];
    reporte.gruposRevisados++;
    for (var i = 1; i < arr.length; i++) {
      var dup = arr[i];
      var deltaSeg = Math.abs(dup.ts - primero.ts) / 1000;
      // Solo anular si está dentro de la ventana de duplicación
      if (deltaSeg > VENTANA_SEG) continue;

      // Resolver el codigoBarra real del base (puede que ENVASADOS guarde
      // skuBase/idProducto en codigoProductoBase pero GUIA_DETALLE usa el
      // codigoBarra). Buscamos el detalle por la cantidad y la guía.
      // Probamos las claves posibles: tal como está en ENVASADOS, y si no
      // matchea, recorremos los detalles de esa guía buscando por cantidad.
      function _buscarDet(idGuia, cant, refTs, codigoIntento) {
        // 1. Intento directo con el código guardado
        var k1 = idGuia + '|' + codigoIntento + '|' + cant;
        var bucket = detIndex[k1];
        // 2. Si no, buscar todos los detalles de esa guía con esa cantidad
        if (!bucket || !bucket.length) {
          bucket = [];
          Object.keys(detIndex).forEach(function(kk){
            if (kk.indexOf(idGuia + '|') === 0 && kk.lastIndexOf('|' + cant) === kk.length - String('|' + cant).length) {
              detIndex[kk].forEach(function(d){ bucket.push(d); });
            }
          });
        }
        // Filtrar no consumidos y elegir el más cercano por timestamp
        var cands = bucket.filter(function(d){ return !d._consumido; });
        if (!cands.length) return null;
        cands.sort(function(a, b){ return Math.abs(a.ts - refTs) - Math.abs(b.ts - refTs); });
        return cands[0];
      }

      var detSal = _buscarDet(dup.idGs, dup.cantBase, dup.ts, dup.base);
      var detIng = _buscarDet(dup.idGi, dup.uds,      dup.ts, dup.derivado);

      var resultSal = null, resultIng = null;
      try { if (detSal) { resultSal = anularDetalle({ idDetalle: detSal.idDet, usuario: 'limpieza-duplicados' }); if (resultSal && resultSal.ok) { detSal._consumido = true; reporte.detallesAnulados++; } } } catch(e1) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'salida', error: e1.message }); }
      try { if (detIng) { resultIng = anularDetalle({ idDetalle: detIng.idDet, usuario: 'limpieza-duplicados' }); if (resultIng && resultIng.ok) { detIng._consumido = true; reporte.detallesAnulados++; } } } catch(e2) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'ingreso', error: e2.message }); }

      // Marcar ENVASADO como anulado por duplicado
      try {
        sheet.getRange(dup.rowIdx, iEst + 1).setValue('ANULADO_DUPLICADO');
        var prevObs = sheet.getRange(dup.rowIdx, iObs + 1).getValue();
        sheet.getRange(dup.rowIdx, iObs + 1).setValue(
          String(prevObs || '') + ' | duplicado de ' + primero.idEnvasado + ' · limpieza ' + nowStr
        );
      } catch(eMark) { reporte.erroresAnulacion.push({ idEnvasado: dup.idEnvasado, lado: 'marcado', error: eMark.message }); }

      reporte.duplicadosAnulados++;
      reporte.detalle.push({
        idEnvasado:    dup.idEnvasado,
        duplicaA:      primero.idEnvasado,
        deltaSeg:      Math.round(deltaSeg),
        producto:      dup.derivado,
        cantBase:      dup.cantBase,
        uds:           dup.uds,
        detSalAnulado: !!(resultSal && resultSal.ok),
        detIngAnulado: !!(resultIng && resultIng.ok)
      });
    }
  });

  Logger.log('limpiarEnvasadosDuplicados ✓ ' + JSON.stringify({
    desde: FECHA_DESDE, grupos: reporte.gruposRevisados,
    duplicadosAnulados: reporte.duplicadosAnulados,
    detallesAnulados: reporte.detallesAnulados,
    errores: reporte.erroresAnulacion.length
  }));
  return { ok: true, data: reporte };
}

// ============================================================
// SCRIPT ONE-SHOT: corregirEnvasadosManuales
// ------------------------------------------------------------
// Ajusta unidadesProducidas de envasados activos cuando el valor
// registrado difiere del real físicamente envasado. Propaga el
// ajuste a stock derivado, stock base y los GUIA_DETALLE asociados.
//
// Se ejecuta MANUALMENTE desde el editor GAS. Lista hardcodeada de
// los 4 casos reportados el 2026-05-15. Idempotente: si ya se corrigió
// un registro, lo detecta y no vuelve a aplicar el ajuste.
// ============================================================
function corregirEnvasadosManuales() {
  var FECHA_DESDE = '2026-05-01';
  // Casos a corregir: el envasado COMPLETADO más reciente de cada
  // codigoDerivado en el período se ajusta a `unidadesCorrectas`.
  var casos = [
    { codigoDerivado: 'WHPACLDO001KG', valorRegistradoEsperado: 250, unidadesCorrectas: 210, descripcion: 'PAN BLANCO MOLIDO 1KG' },
    { codigoDerivado: 'WHPADCRO001KG', valorRegistradoEsperado: 75,  unidadesCorrectas: 80,  descripcion: 'PAN MOLIDO OSCURO 1KG' },
    { codigoDerivado: 'WHAJLGRO250GR', valorRegistradoEsperado: 90,  unidadesCorrectas: 80,  descripcion: 'AJONJOLI NEGRO 250GR' },
    { codigoDerivado: 'WHAJLGRO100GR', valorRegistradoEsperado: 60,  unidadesCorrectas: 50,  descripcion: 'AJONJOLI NEGRO 100GR' }
  ];

  var sheet = getSheet('ENVASADOS');
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: false, error: 'ENVASADOS vacía' };
  var hdrs  = data[0].map(function(h){ return String(h); });
  var iId    = hdrs.indexOf('idEnvasado');
  var iBase  = hdrs.indexOf('codigoProductoBase');
  var iCnt   = hdrs.indexOf('cantidadBase');
  var iDer   = hdrs.indexOf('codigoProductoEnvasado');
  var iUds   = hdrs.indexOf('unidadesProducidas');
  var iFecha = hdrs.indexOf('fecha');
  var iEst   = hdrs.indexOf('estado');
  var iGs    = hdrs.indexOf('idGuiaSalida');
  var iGi    = hdrs.indexOf('idGuiaIngreso');
  var iObs   = hdrs.indexOf('observacion');

  var productos = _sheetToObjects(getProductosSheet());
  var reporte = { fechaDesde: FECHA_DESDE, procesados: [], saltados: [], errores: [] };
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  casos.forEach(function(caso) {
    // 1. Encontrar producto derivado
    var prodDerivado = productos.find(function(p) {
      return String(p.codigoBarra).trim() === caso.codigoDerivado;
    });
    if (!prodDerivado) {
      reporte.errores.push({ caso: caso.codigoDerivado, error: 'derivado no encontrado en PRODUCTOS' });
      return;
    }
    var factorBase = parseFloat(prodDerivado.factorConversionBase) || 0;
    if (factorBase <= 0) {
      reporte.errores.push({ caso: caso.codigoDerivado, error: 'factorConversionBase invalido' });
      return;
    }
    var prodBase = productos.find(function(p) {
      return String(p.skuBase).trim() === String(prodDerivado.codigoProductoBase).trim()
          || String(p.idProducto).trim() === String(prodDerivado.codigoProductoBase).trim();
    });

    // 2. Buscar el envasado COMPLETADO más reciente desde FECHA_DESDE
    //    para este codigoDerivado con el valor original esperado.
    var candidatos = [];
    for (var r = 1; r < data.length; r++) {
      var f = data[r][iFecha];
      var fStr = f instanceof Date
        ? Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(f || '').substring(0, 10);
      if (fStr < FECHA_DESDE) continue;
      var est = String(data[r][iEst] || '').toUpperCase();
      if (est !== 'COMPLETADO') continue;
      var der = String(data[r][iDer] || '');
      if (der !== caso.codigoDerivado) continue;
      var uds = parseFloat(data[r][iUds]) || 0;
      var obs = String(data[r][iObs] || '');
      // Si ya tiene marca de corrección manual previa → saltarlo
      if (obs.indexOf('corregido-manual') >= 0) {
        reporte.saltados.push({ caso: caso.codigoDerivado, motivo: 'ya corregido previamente · row=' + (r + 1) });
        return;
      }
      // Tolerancia: aceptar match exacto del valor esperado (250, 75, 90, 60)
      if (uds === caso.valorRegistradoEsperado) {
        var fTs = f instanceof Date ? f.getTime() : new Date(f).getTime();
        candidatos.push({ rowIdx: r + 1, ts: fTs, uds: uds, idEnv: String(data[r][iId] || '') });
      }
    }
    if (!candidatos.length) {
      reporte.saltados.push({ caso: caso.codigoDerivado, motivo: 'no se encontró envasado COMPLETADO con valor=' + caso.valorRegistradoEsperado });
      return;
    }
    // Si hay varios COMPLETADOS con el mismo valor (no debería pasar tras
    // limpiarEnvasadosDuplicados), corregir el más reciente.
    candidatos.sort(function(a, b){ return b.ts - a.ts; });
    var target = candidatos[0];

    // 3. Calcular diferencias
    var udsViejas    = target.uds;
    var udsNuevas    = caso.unidadesCorrectas;
    var cantBaseVieja= udsViejas * factorBase;
    var cantBaseNueva= udsNuevas * factorBase;
    var deltaUds     = udsNuevas - udsViejas;            // + significa producir más; - significa producir menos
    var deltaBase    = cantBaseNueva - cantBaseVieja;    // + significa consumir más base; - consumir menos

    try {
      // 4. Actualizar ENVASADOS: unidadesProducidas + cantidadBase + observacion
      sheet.getRange(target.rowIdx, iUds + 1).setValue(udsNuevas);
      sheet.getRange(target.rowIdx, iCnt + 1).setValue(cantBaseNueva);
      var obsPrev = sheet.getRange(target.rowIdx, iObs + 1).getValue();
      sheet.getRange(target.rowIdx, iObs + 1).setValue(
        String(obsPrev || '') + ' | corregido-manual ' + nowStr + ' · ' + udsViejas + '→' + udsNuevas + ' uds'
      );

      // 5. Ajustar stock derivado: si udsNuevas < udsViejas, restar diferencia (revertir sobra producida)
      _actualizarStock(prodDerivado.codigoBarra, deltaUds, {
        tipoOperacion: 'CORRECCION_MANUAL_ENVASADO',
        origen:        target.idEnv,
        usuario:       'correccion-manual',
        observacion:   'derivado: ' + udsViejas + '→' + udsNuevas + ' uds'
      });

      // 6. Ajustar stock base: deltaBase > 0 si ahora consume más; < 0 si consume menos
      //    El consumo base original fue -cantBaseVieja. El nuevo debería ser -cantBaseNueva.
      //    Para llegar de stock actual al correcto: ajustar por -deltaBase
      //    (si delta positivo: necesitamos consumir más → restar al stock; viceversa)
      if (prodBase) {
        _actualizarStock(prodBase.codigoBarra, -deltaBase, {
          tipoOperacion: 'CORRECCION_MANUAL_ENVASADO',
          origen:        target.idEnv,
          usuario:       'correccion-manual',
          observacion:   'base: ' + cantBaseVieja + '→' + cantBaseNueva
        });
      }

      // 7. Ajustar GUIA_DETALLE: ingreso (derivado) y salida (base)
      _ajustarDetalleEnvasado(target.idEnv, prodDerivado.codigoBarra, udsNuevas, 'ingreso');
      if (prodBase) {
        _ajustarDetalleEnvasado(target.idEnv, prodBase.codigoBarra, cantBaseNueva, 'salida');
      }

      reporte.procesados.push({
        codigoDerivado: caso.codigoDerivado,
        descripcion:    caso.descripcion,
        idEnvasado:     target.idEnv,
        udsViejas:      udsViejas,
        udsNuevas:      udsNuevas,
        cantBaseVieja:  cantBaseVieja,
        cantBaseNueva:  cantBaseNueva
      });
    } catch(e) {
      reporte.errores.push({ caso: caso.codigoDerivado, error: e.message });
    }
  });

  Logger.log('corregirEnvasadosManuales ✓ ' + JSON.stringify(reporte));
  return { ok: true, data: reporte };
}

// Helper interno: ajusta cantidadRecibida + cantidadEsperada en el detalle
// asociado a un envasado. Busca el detalle por idGuia+codigoProducto y por
// proximidad temporal al idEnvasado (cuando hay múltiples).
function _ajustarDetalleEnvasado(idEnvasado, codigoBarra, nuevaCant, lado) {
  var envSheet = getSheet('ENVASADOS');
  var envData  = envSheet.getDataRange().getValues();
  var hdrs = envData[0].map(function(h){ return String(h); });
  var iId  = hdrs.indexOf('idEnvasado');
  var iGs  = hdrs.indexOf('idGuiaSalida');
  var iGi  = hdrs.indexOf('idGuiaIngreso');
  var idGuia = '';
  for (var r = 1; r < envData.length; r++) {
    if (String(envData[r][iId]) === idEnvasado) {
      idGuia = String(envData[r][lado === 'salida' ? iGs : iGi] || '');
      break;
    }
  }
  if (!idGuia) return;

  var detSheet = getSheet('GUIA_DETALLE');
  var detData  = detSheet.getDataRange().getValues();
  var dHdrs    = detData[0].map(function(h){ return String(h); });
  var dIdGuia  = dHdrs.indexOf('idGuia');
  var dCod     = dHdrs.indexOf('codigoProducto');
  var dCantE   = dHdrs.indexOf('cantidadEsperada');
  var dCantR   = dHdrs.indexOf('cantidadRecibida');
  var dObs     = dHdrs.indexOf('observacion');
  for (var dr = 1; dr < detData.length; dr++) {
    if (String(detData[dr][dIdGuia]) !== idGuia) continue;
    if (String(detData[dr][dCod])    !== String(codigoBarra)) continue;
    var obs = String(detData[dr][dObs] || '').toUpperCase();
    if (obs === 'ANULADO') continue;
    detSheet.getRange(dr + 1, dCantE + 1).setValue(nuevaCant);
    detSheet.getRange(dr + 1, dCantR + 1).setValue(nuevaCant);
    var obsPrev = detSheet.getRange(dr + 1, dObs + 1).getValue();
    detSheet.getRange(dr + 1, dObs + 1).setValue(
      String(obsPrev || '') + ' | corregido-manual ' + (new Date()).toISOString()
    );
    return; // solo el primero matcheante
  }
}

// ============================================================
// corregirUnidadesEnvasado — admin-gated
// ------------------------------------------------------------
// Edita unidadesProducidas de un envasado EXISTENTE y propaga:
//   - ENVASADOS (uds + cantidadBase + observacion)
//   - STOCK derivado (delta uds, signo según suba/baje)
//   - STOCK base (-deltaBase, restituye o consume kg adicionales)
//   - GUIA_DETALLE ingreso + salida (ajusta cantidadEsperada/Recibida)
// Idempotente: calcula delta contra el valor ACTUAL, no contra el
// original. Editar el mismo envasado N veces solo aplica el delta
// neto en cada paso, sin desbalancear stock.
// ============================================================
function corregirUnidadesEnvasado(params) {
  return _conLock('corregirUnidadesEnvasado', function() {
    return _corregirUnidadesEnvasadoImpl(params);
  });
}

function _corregirUnidadesEnvasadoImpl(params) {
  // 1. Validar admin
  var auth = _requireAdmin(params);
  if (!auth.ok) return auth;

  var idEnvasado    = String(params.idEnvasado || '').trim();
  var nuevasUds     = parseFloat(params.nuevasUnidades);
  var motivo        = String(params.motivo || '').trim() || 'sin motivo';
  var usuario       = String(params.usuario || 'admin').trim();
  if (!idEnvasado || !isFinite(nuevasUds) || nuevasUds < 0) {
    return { ok: false, error: 'Faltan datos: idEnvasado, nuevasUnidades' };
  }

  var sheet = getSheet('ENVASADOS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0].map(function(h){ return String(h); });
  var iId   = hdrs.indexOf('idEnvasado');
  var iBase = hdrs.indexOf('codigoProductoBase');
  var iCnt  = hdrs.indexOf('cantidadBase');
  var iDer  = hdrs.indexOf('codigoProductoEnvasado');
  var iUds  = hdrs.indexOf('unidadesProducidas');
  var iEst  = hdrs.indexOf('estado');
  var iObs  = hdrs.indexOf('observacion');

  // 2. Encontrar fila
  var rowIdx = -1, fila = null;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === idEnvasado) { rowIdx = r + 1; fila = data[r]; break; }
  }
  if (rowIdx < 0) return { ok: false, error: 'Envasado no encontrado: ' + idEnvasado };

  var estado = String(fila[iEst] || '').toUpperCase();
  if (estado === 'ANULADO' || estado === 'ANULADO_DUPLICADO' || estado === 'ANULADO_MANUAL') {
    return { ok: false, error: 'No se puede editar un envasado anulado' };
  }

  // 3. Resolver productos y factor
  var productos = _sheetToObjects(getProductosSheet());
  var cbDerivado = String(fila[iDer] || '');
  var cbBase     = String(fila[iBase] || '');
  var prodDer = productos.find(function(p){ return String(p.codigoBarra) === cbDerivado; });
  if (!prodDer) return { ok: false, error: 'Producto derivado no encontrado: ' + cbDerivado };
  var factorBase = parseFloat(prodDer.factorConversionBase) || 0;
  if (factorBase <= 0) return { ok: false, error: 'factorConversionBase invalido' };
  var prodBase = productos.find(function(p){
    return String(p.codigoBarra) === cbBase
        || String(p.skuBase) === cbBase
        || String(p.idProducto) === cbBase;
  });

  // 4. Calcular deltas
  var udsViejas    = parseFloat(fila[iUds]) || 0;
  var cantBaseV    = parseFloat(fila[iCnt]) || udsViejas * factorBase;
  var cantBaseN    = nuevasUds * factorBase;
  var deltaUds     = nuevasUds - udsViejas;
  var deltaBase    = cantBaseN - cantBaseV;
  if (deltaUds === 0) return { ok: false, error: 'No hay cambio: las unidades nuevas son iguales a las actuales' };

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 5. Aplicar
  try {
    sheet.getRange(rowIdx, iUds + 1).setValue(nuevasUds);
    sheet.getRange(rowIdx, iCnt + 1).setValue(cantBaseN);
    var obsPrev = sheet.getRange(rowIdx, iObs + 1).getValue();
    sheet.getRange(rowIdx, iObs + 1).setValue(
      String(obsPrev || '') + ' | editado ' + nowStr + ' · ' + udsViejas + '→' + nuevasUds + ' uds · admin=' + usuario + ' · ' + motivo
    );

    _actualizarStock(cbDerivado, deltaUds, {
      tipoOperacion: 'EDICION_ENVASADO',
      origen:        idEnvasado,
      usuario:       usuario,
      observacion:   'edición derivado: ' + udsViejas + '→' + nuevasUds + ' · ' + motivo
    });
    if (prodBase) {
      _actualizarStock(prodBase.codigoBarra, -deltaBase, {
        tipoOperacion: 'EDICION_ENVASADO',
        origen:        idEnvasado,
        usuario:       usuario,
        observacion:   'edición base: ' + cantBaseV + '→' + cantBaseN + ' · ' + motivo
      });
    }
    _ajustarDetalleEnvasado(idEnvasado, cbDerivado, nuevasUds, 'ingreso');
    if (prodBase) _ajustarDetalleEnvasado(idEnvasado, prodBase.codigoBarra, cantBaseN, 'salida');

    return {
      ok: true,
      data: {
        idEnvasado:     idEnvasado,
        udsViejas:      udsViejas,
        udsNuevas:      nuevasUds,
        deltaUds:       deltaUds,
        deltaBase:      deltaBase,
        descripcion:    prodDer.descripcion || cbDerivado,
        descripcionBase: prodBase ? prodBase.descripcion : ''
      }
    };
  } catch(e) {
    return { ok: false, error: 'Error aplicando edición: ' + e.message };
  }
}

// ============================================================
// anularEnvasadoConClave — admin-gated
// ------------------------------------------------------------
// Anula un envasado COMPLETADO: revierte stock derivado y base,
// anula los detalles de guía ingreso + salida, y marca el envasado
// como ANULADO_MANUAL con trazabilidad de admin + motivo.
// ============================================================
function anularEnvasadoConClave(params) {
  return _conLock('anularEnvasadoConClave', function() {
    return _anularEnvasadoConClaveImpl(params);
  });
}

function _anularEnvasadoConClaveImpl(params) {
  var auth = _requireAdmin(params);
  if (!auth.ok) return auth;

  var idEnvasado = String(params.idEnvasado || '').trim();
  var motivo     = String(params.motivo || '').trim() || 'sin motivo';
  var usuario    = String(params.usuario || 'admin').trim();
  if (!idEnvasado) return { ok: false, error: 'idEnvasado requerido' };

  var sheet = getSheet('ENVASADOS');
  var data  = sheet.getDataRange().getValues();
  var hdrs  = data[0].map(function(h){ return String(h); });
  var iId   = hdrs.indexOf('idEnvasado');
  var iBase = hdrs.indexOf('codigoProductoBase');
  var iCnt  = hdrs.indexOf('cantidadBase');
  var iDer  = hdrs.indexOf('codigoProductoEnvasado');
  var iUds  = hdrs.indexOf('unidadesProducidas');
  var iEst  = hdrs.indexOf('estado');
  var iObs  = hdrs.indexOf('observacion');
  var iGs   = hdrs.indexOf('idGuiaSalida');
  var iGi   = hdrs.indexOf('idGuiaIngreso');

  var rowIdx = -1, fila = null;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === idEnvasado) { rowIdx = r + 1; fila = data[r]; break; }
  }
  if (rowIdx < 0) return { ok: false, error: 'Envasado no encontrado' };
  var estado = String(fila[iEst] || '').toUpperCase();
  if (estado.indexOf('ANULADO') === 0) return { ok: false, error: 'Ya está anulado' };

  var productos = _sheetToObjects(getProductosSheet());
  var cbDerivado = String(fila[iDer] || '');
  var cbBase     = String(fila[iBase] || '');
  var prodDer = productos.find(function(p){ return String(p.codigoBarra) === cbDerivado; });
  var prodBase = productos.find(function(p){
    return String(p.codigoBarra) === cbBase
        || String(p.skuBase) === cbBase
        || String(p.idProducto) === cbBase;
  });
  var udsViejas = parseFloat(fila[iUds]) || 0;
  var cantBaseV = parseFloat(fila[iCnt]) || 0;

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  try {
    // Revertir stock derivado (-uds) y base (+kg)
    _actualizarStock(cbDerivado, -udsViejas, {
      tipoOperacion: 'ANULACION_ENVASADO',
      origen:        idEnvasado,
      usuario:       usuario,
      observacion:   'anulación derivado · ' + motivo
    });
    if (prodBase) {
      _actualizarStock(prodBase.codigoBarra, +cantBaseV, {
        tipoOperacion: 'ANULACION_ENVASADO',
        origen:        idEnvasado,
        usuario:       usuario,
        observacion:   'anulación base · ' + motivo
      });
    }

    // Anular los detalles de las guías
    var idGs = String(fila[iGs] || '');
    var idGi = String(fila[iGi] || '');
    if (idGi) _anularDetallesDeGuiaPorProducto(idGi, cbDerivado, 'anulación envasado ' + idEnvasado);
    if (idGs && prodBase) _anularDetallesDeGuiaPorProducto(idGs, prodBase.codigoBarra, 'anulación envasado ' + idEnvasado);

    // Marcar envasado
    sheet.getRange(rowIdx, iEst + 1).setValue('ANULADO_MANUAL');
    var obsPrev = sheet.getRange(rowIdx, iObs + 1).getValue();
    sheet.getRange(rowIdx, iObs + 1).setValue(
      String(obsPrev || '') + ' | anulado ' + nowStr + ' · ' + udsViejas + ' uds revertidas · admin=' + usuario + ' · ' + motivo
    );

    return {
      ok: true,
      data: {
        idEnvasado:     idEnvasado,
        udsAnuladas:    udsViejas,
        cantBaseRestit: cantBaseV,
        descripcion:    prodDer ? prodDer.descripcion : cbDerivado
      }
    };
  } catch(e) {
    return { ok: false, error: 'Error anulando: ' + e.message };
  }
}

// Helper: anula TODOS los detalles activos de una guía que matcheen
// el codigoProducto. Usado al anular un envasado.
function _anularDetallesDeGuiaPorProducto(idGuia, codigoBarra, motivo) {
  var detSheet = getSheet('GUIA_DETALLE');
  var detData  = detSheet.getDataRange().getValues();
  var dHdrs    = detData[0].map(function(h){ return String(h); });
  var dIdGuia  = dHdrs.indexOf('idGuia');
  var dCod     = dHdrs.indexOf('codigoProducto');
  var dObs     = dHdrs.indexOf('observacion');
  for (var dr = 1; dr < detData.length; dr++) {
    if (String(detData[dr][dIdGuia]) !== String(idGuia)) continue;
    if (String(detData[dr][dCod])    !== String(codigoBarra)) continue;
    var obs = String(detData[dr][dObs] || '').toUpperCase();
    if (obs === 'ANULADO') continue;
    detSheet.getRange(dr + 1, dObs + 1).setValue('ANULADO · ' + motivo);
    return; // solo el primero matcheante
  }
}

// ============================================================
// enviarResumenEnvasadosDia — cron 20:00 diario
// ------------------------------------------------------------
// Arma resumen por operador de TODOS los envasados COMPLETADOS del
// día y manda push al MOS (rol admin/master) con idNotif=WH_RESUMEN_ENVASADOS_DIA
// para que el dueño vea de un vistazo cuánto envasó cada uno.
// Si no hubo envasados, no manda nada.
// ============================================================
function enviarResumenEnvasadosDia() {
  var tz = Session.getScriptTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var rows = _sheetToObjects(getSheet('ENVASADOS'));
  var productos = _sheetToObjects(getProductosSheet());
  var prodMap = {};
  productos.forEach(function(p){ prodMap[String(p.codigoBarra)] = p.descripcion || p.codigoBarra; });

  // Filtrar a hoy + COMPLETADO
  var delDia = rows.filter(function(r){
    var f = r.fecha;
    var fStr = f instanceof Date
      ? Utilities.formatDate(f, tz, 'yyyy-MM-dd')
      : String(f || '').substring(0, 10);
    return fStr === hoy && String(r.estado || '').toUpperCase() === 'COMPLETADO';
  });
  if (!delDia.length) {
    Logger.log('[ResumenEnvasados] sin envasados COMPLETADOS hoy ' + hoy);
    return { ok: true, data: { hoy: hoy, operadores: 0, total: 0 } };
  }

  // Agrupar por usuario, luego por codigoProductoEnvasado
  var porUsr = {};
  delDia.forEach(function(e){
    var u = String(e.usuario || 'desconocido');
    var cb = String(e.codigoProductoEnvasado || '');
    var uds = parseFloat(e.unidadesProducidas) || 0;
    if (!porUsr[u]) porUsr[u] = { total: 0, productos: {} };
    porUsr[u].total += uds;
    porUsr[u].productos[cb] = (porUsr[u].productos[cb] || 0) + uds;
  });

  // Armar cuerpo: "jorgenis 500u · 200 ajinomoto 1kg · 300 pimienta 1kg"
  var lineas = [];
  Object.keys(porUsr).forEach(function(u){
    var p = porUsr[u];
    var partes = [u + ' ' + Math.round(p.total) + 'u'];
    Object.keys(p.productos).forEach(function(cb){
      var nombre = prodMap[cb] || cb;
      partes.push(Math.round(p.productos[cb]) + ' ' + nombre);
    });
    lineas.push(partes.join(' · '));
  });

  var titulo = '📦 Resumen envasados ' + hoy;
  var cuerpo = lineas.join('\n');

  try {
    _notificarMOS(titulo, cuerpo, null, 'WH_RESUMEN_ENVASADOS_DIA');
    Logger.log('[ResumenEnvasados] push enviado · ' + Object.keys(porUsr).length + ' operadores');
  } catch(e) {
    Logger.log('[ResumenEnvasados] error push: ' + e.message);
  }

  return { ok: true, data: { hoy: hoy, operadores: Object.keys(porUsr).length, lineas: lineas } };
}

// Configura el trigger diario para enviarResumenEnvasadosDia
// (ejecutar UNA VEZ desde editor GAS). Mata triggers previos del mismo handler.
function configurarTriggerResumenEnvasados() {
  var TRG = 'enviarResumenEnvasadosDia';
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === TRG) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(TRG).timeBased().atHour(20).everyDays(1).create();
  Logger.log('[Trigger] ' + TRG + ' configurado · diario 20:00');
  return { ok: true };
}
