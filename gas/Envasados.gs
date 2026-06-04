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

// ── Logo bitmap (184x36 dots = 23 bytes/row x 36) ──
// [v2.13.104] Regenerado: TONY'S block bold + casita (sin "Caserito").
// Fuente canónica: /ProyectoMOS/assets/adhesivo/logo-tonys.svg
// Pipeline reproducible: /ProyectoMOS/assets/adhesivo/gen.py
// Para regenerar: cd ProyectoMOS/assets/adhesivo && python gen.py
//   → copia el contenido de logo-tonys-S.hex acá
//   → copia el contenido de logo-tonys-S.b64 a _ADHESIVO_LOGO_DATAURI en MOS app.js
var LOGO_W_BYTES = 23;
var LOGO_H = 36;
var LOGO_TSPL_HEX =
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE3FFFFC000' +
'1FE03FE0780C078083F80FFFFFFFFFFFFFFFC1FFFFC0001F800FE0380C078083C001FFFFFFFF' +
'FFFFFF007FFFC0001E0007E0380C0701838000FFFFFFFFFFFFFE003FFFC0001E0003E0380E03' +
'018380007FFFFFFFFFFFFC001FFFC0001C0001E0180E03018300007FFFFFFFFFFFF8000FFFC0' +
'001C0201E0180E03018300C07FFFFFFFFFFFE00003FFFE01FC0701E0180F03038301C07FFFFF' +
'FFFFFFC00001FFFE01F80701E0080F03038301C07FFFFFFFFFFF800000FFFE01F80701E0080F' +
'0203C301C07FFFFFFFFFFF0000007FFE01F80701E0080F8007FF00FFFFFFFFFFFFFC0000001F' +
'FE01F80701E0080F8007FF007FFFFFFFFFFFF000000007FE01F80701E0000F8007FF001FFFFF' +
'FFFFFFF000000007FE01F80701E0000F800FFF800FFFFFFFFFFFF000000007FE01F80701E000' +
'0FC00FFFC003FFFFFFFFFFF000000007FE01F80701E0000FC00FFFE001FFFFFFFFFFF0000000' +
'07FE01F80701E0000FC00FFFF000FFFFFFFFFFFE0000003FFE01F80701E0000FE01FFFFC007F' +
'FFFFFFFFFE0000003FFE01F80701E0000FE01FFFFE007FFFFFFFFFFE3E003E3FFE01F80701E0' +
'400FE01FFFFF807FFFFFFFFFFE3E003E3FFE01F80701E0400FE01FFF01C03FFFFFFFFFFE3E00' +
'3E3FFE01F80701E0400FE01FFF01C03FFFFFFFFFFE3EFFBE3FFE01F80701E0600FE01FFF01C0' +
'3FFFFFFFFFFE00FF803FFE01F80701E0600FE01FFF01C03FFFFFFFFFFE00FF803FFE01FC0701' +
'E0600FE01FFF01C03FFFFFFFFFFE00FF803FFE01FC0601E0700FE01FFF00C03FFFFFFFFFFE00' +
'FF803FFE01FC0001E0700FE01FFF80007FFFFFFFFFFE00FF803FFE01FE0003E0700FE01FFF80' +
'007FFFFFFFFFFE00FF803FFE01FE0007E0700FE01FFFC000FFFFFFFFFFFE00FF803FFE01FF80' +
'0FE0780FE01FFFE001FFFFFFFFFFFE00FF803FFE01FFE03FE0780FE01FFFF807FFFFFFFFFFFE' +
'00FF803FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE00FF803FFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFE00FF803FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';

// [v2.13.104] Tabla manual de meses ES — no depende de la locale del Script.
// Antes: Utilities.formatDate(d, tz, 'MMM') devolvía "Jan" si la locale era EN.
var MESES_ES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

// ── Helpers de texto ──
function _normalizeEtq(s) {
  if (s === null || s === undefined) return '';
  // NFD + strip diacriticos: "CAÑIHUA" → "CANIHUA", "CAFÉ" → "CAFE"
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function _calcVencimientoEtq(fechaEnvasado) {
  // [v2.13.104] Formato MES/yyyy (ej "ENE/2027"). Tabla manual MESES_ES
  // para no depender de la locale del Script GAS. El cálculo sigue siendo
  // fechaEnvasado + 1 año exacto.
  var d = fechaEnvasado ? new Date(fechaEnvasado) : new Date();
  d.setFullYear(d.getFullYear() + 1);
  return MESES_ES[d.getMonth()] + '/' + d.getFullYear();
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
  //   ADHESIVO_OFFSET_Y default 0  (dots offset manual base)
  //   ADHESIVO_DRIFT_DOTS_POR_PRINT  drift acumulativo (auto-detect)
  //   ADHESIVO_PRINTS_DESDE_CAL      contador automático
  // [v2.13.118] _calcularOffsetEfectivoParaPrint() devuelve offsetBase
  // MENOS la compensación acumulada de drift, así cada print se va
  // corriendo "hacia abajo" para compensar el drift natural de la impresora.
  var props = PropertiesService.getScriptProperties();
  var gapMm    = parseFloat(props.getProperty('ADHESIVO_GAP_MM'))   || 2;
  var density  = parseInt(props.getProperty('ADHESIVO_DENSITY'))    || 8;
  var speed    = parseInt(props.getProperty('ADHESIVO_SPEED'))      || 4;
  var offsetY  = _calcularOffsetEfectivoParaPrint();

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
  // [v2.13.106] Vto con margen derecho de 2 letras (~24 dots).
  //   Font "2" = 12×20 dots. "Vto ENE/2027" = 144 dots wide.
  //   X=232 → end=376, margen 24 dots = 2 letras de aire (antes era 16 = 1 letra,
  //   se veía pegado al borde).
  bytes = bytes.concat(_strToBytesEtq('TEXT 232,' + (12 + offsetY) + ',"2",0,1,1,"Vto ' + vto + '"\r\n'));
  // Separador
  bytes = bytes.concat(_strToBytesEtq('BAR 5,' + (42 + offsetY) + ',390,1\r\n'));

  // [v2.13.106] Descripción: centrar VERTICAL+HORIZONTAL si es 1 línea.
  // Área disponible: Y=46 a Y=118 (72 dots), antes del frame del barcode.
  //   - Si 1 línea: startY ajustado para centrar vertical en esos 72 dots.
  //   - Si 2 líneas: mantener startY=46 con LINE_H=38 (queda como estaba).
  var DESC_AREA_Y0 = 46, DESC_AREA_H = 72;
  var LINE_H = 38, SPACE = 8;
  var startY;
  if (lines.length === 1) {
    var lineHasHl = lines[0].some(function(t) { return t.hl; });
    var lineHeight = lineHasHl ? 32 : 24;  // font4 vs font3
    // [v2.13.107 AUDIT FIX] El loop aplica yAdj = y + 4 para no-hl (baseline
    // align). Si no compensamos, la línea queda 4 dots abajo del centro real.
    var baselineOffset = lineHasHl ? 0 : 4;
    startY = DESC_AREA_Y0 + Math.floor((DESC_AREA_H - lineHeight) / 2) - baselineOffset + offsetY;
  } else {
    startY = DESC_AREA_Y0 + offsetY;
  }
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

  // [v2.13.104] Barcode minimalista: angosto + quiet zone amplio + sin flechas.
  //
  // Cambios respecto a versiones anteriores:
  //   1. Umbral narrow=1: 360 → 300 dots.
  //      Antes el barcode ocupaba hasta 360 dots (90% del adhesivo).
  //      Ahora con códigos de 11+ chars usa narrow=1 → barcode mucho más
  //      angosto. Códigos cortos (<11) mantienen narrow=2.
  //   2. Quiet zone: 10×narrow → 15×narrow.
  //      ANSI Code128 mínimo es 10×narrow, pero la práctica recomienda 15
  //      para reducir lectura parcial cuando el operario escanea rápido.
  //   3. Sin flechas guía — el quiet zone amplio reemplaza esa función
  //      visual + es lo que técnicamente importa.
  //
  // Fórmula Code128 (Start 11 + chars 11×bcLen + Check 11 + Stop 13):
  //   modules = 11*bcLen + 35
  //   width   = modules × narrow_dots
  var bc = String(producto.codigoBarra || '').replace(/"/g, '');
  var bcLen = bc.length;
  var modules = 11 * bcLen + 35;
  var narrowBc = 2;
  var barcodeWidth = modules * narrowBc;
  if (barcodeWidth > 300) {
    narrowBc = 1;
    barcodeWidth = modules * narrowBc;
  }
  var barcodeHeight = 44;
  var barcodeX = Math.max(20, Math.floor((400 - barcodeWidth) / 2));
  // [v2.13.106] Barcode dentro de frame con corner marks.
  //   Frame box: X=10..390, Y=118..196 (aprovecha margen inferior antes desperdiciado).
  //   Corner marks tipo cámara/visor QR — 4 esquinas tipo "L" de 12 dots.
  //   Barcode Y=124 (6 dots dentro del top del frame).
  //   Texto código centrado horizontal Y=174 (6 dots después del barcode bottom).
  var barcodeY = 124 + offsetY;
  var frameX1 = 10, frameX2 = 389;
  var frameY1 = 118 + offsetY, frameY2 = 196 + offsetY;
  var cmL = 12;  // largo de cada corner mark (L-shape)
  // Top-left corner ┐ (rotated)
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  // Top-right corner ┌ (rotated)
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  // Bottom-left corner ┘
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  // Bottom-right corner └
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  // Barcode SIN texto auto (5° param = 0). El código va con TEXT centrado abajo.
  bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + bc + '"\r\n'));
  // Texto código: font 1 (8×12 dots), centrado horizontal en el adhesivo.
  var codigoFontW = 8;
  var codigoWidth = bc.length * codigoFontW;
  var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));
  var codigoY = barcodeY + barcodeHeight + 8;  // 8 dots después del barcode (barcodeY ya incluye offsetY)
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + bc + '"\r\n'));

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
  // [v2.13.104] Tabla manual MESES_ES — consistente con _calcVencimientoEtq.
  var vtoOverride = null;
  if (data.fechaVencimiento) {
    var dOverride = new Date(data.fechaVencimiento);
    vtoOverride = MESES_ES[dOverride.getMonth()] + '/' + dOverride.getFullYear();
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
// ════════════════════════════════════════════════════════════════════
// [v2.13.100] Diagnóstico: devuelve el TSPL2 EXACTO que se enviaría
// a la impresora para un código dado, SIN imprimir. Útil para verificar
// que los cambios de v2.13.96-97 (barcode centrado + height 44 + flechas)
// efectivamente se están aplicando.
// ════════════════════════════════════════════════════════════════════
function previsualizarTSPLEtq(params) {
  try {
    var codigoBarra = String((params && params.codigoBarra) || 'WHCOLAGO250GR').trim();
    var all = _getAllEnvasablesTokens();
    var prod = null;
    for (var i = 0; i < all.length; i++) {
      if (all[i].codigoBarra === codigoBarra) { prod = all[i]; break; }
    }
    if (!prod) {
      // fallback: producto sintético para testear sin necesidad de uno real
      prod = { codigoBarra: codigoBarra, descripcion: codigoBarra };
    }
    var bytes = _buildTSPLEtq(prod, new Date(), 1, all);
    // Convertir bytes a texto (excepto el bitmap binario que dejamos como hex)
    var txt = '';
    var i2;
    for (i2 = 0; i2 < bytes.length; i2++) {
      var b = bytes[i2];
      if (b >= 32 && b <= 126) txt += String.fromCharCode(b);
      else if (b === 13) txt += '\\r';
      else if (b === 10) txt += '\\n\n';
      else txt += '<' + b.toString(16).toUpperCase().padStart(2, '0') + '>';
    }
    return {
      ok: true,
      data: {
        codigoBarra: prod.codigoBarra,
        descripcion: prod.descripcion,
        totalBytes: bytes.length,
        tsplPreview: txt
      }
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function calibrarImpresoraAdhesivo() {
  // [v2.13.118] Calibración manual al cambiar rollo:
  //   1. GAPDETECT físico (impresora mide el GAP real)
  //   2. FORMFEED (avanza una etiqueta limpia)
  //   3. Reset del contador de prints y drift (rollo nuevo = drift desconocido)
  //   4. Guardar fecha de calibración
  // Gasta ~3 etiquetas pero solo se hace UNA VEZ por rollo.
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var printerId;
    try { printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN'); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }

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

    // [v2.13.118] Reset del estado de calibración: rollo nuevo, contador a 0,
    // drift desconocido (operador puede correr auto-detect después).
    var props = PropertiesService.getScriptProperties();
    props.setProperty('ADHESIVO_ROLLO_CALIBRADO', 'true');
    props.setProperty('ADHESIVO_PRINTS_DESDE_CAL', '0');
    props.setProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT', '0');
    props.setProperty('ADHESIVO_FECHA_CALIBRADO', new Date().toISOString());

    return { ok: true, data: {
      jobId: JSON.parse(resp.getContentText()),
      mensaje: 'Calibración enviada. Después de las ~3 etiquetas blancas, podés ejecutar Auto-detectar drift para fine-tunear, o imprimir directo.',
      estado: estadoCalibracionRollo().data
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════
// [v2.13.118] CALIBRACIÓN INTELIGENTE — drift compensation por print
// ════════════════════════════════════════════════════════════════════
//
// Properties relevantes:
//   ADHESIVO_ROLLO_CALIBRADO       "true"/"false"
//   ADHESIVO_GAP_MM_MEDIDO         "2.05" (informativo, futuro)
//   ADHESIVO_DRIFT_DOTS_POR_PRINT  "4"     (dots a compensar por print)
//   ADHESIVO_PRINTS_DESDE_CAL      "0"     (contador automático)
//   ADHESIVO_FECHA_CALIBRADO       ISO timestamp
//   ADHESIVO_OFFSET_Y              "0"     (offset manual fino base)
//
// Estrategia:
//   - Cada print incrementa PRINTS_DESDE_CAL
//   - El offset efectivo = OFFSET_Y_BASE - (DRIFT * PRINTS_DESDE_CAL)
//     (compensa el drift acumulado moviendo TODO el contenido hacia abajo)
//   - Cuando supera 500 sin recalibrar, frontend muestra alerta amarilla.

// Lee el offsetY a aplicar para el PRÓXIMO print, incluyendo compensación
// acumulada de drift. NO incrementa el contador (eso lo hace _incrementarPrintsCount
// después del print exitoso).
function _calcularOffsetEfectivoParaPrint() {
  var props = PropertiesService.getScriptProperties();
  var offsetBase  = parseFloat(props.getProperty('ADHESIVO_OFFSET_Y'))             || 0;
  var driftDots   = parseFloat(props.getProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT')) || 0;
  var printsCount = parseInt  (props.getProperty('ADHESIVO_PRINTS_DESDE_CAL'))     || 0;
  // [v2.13.144 BUGFIX SIGNO INVERTIDO]
  // En TSPL2 Y=0 = TOPE del label y Y aumenta hacia abajo. Para compensar un
  // drift que SUBE el contenido cada print (caso típico, lo más común), hay
  // que AUMENTAR Y → SUMAR compensación, no restarla.
  // Convención: driftDots > 0 = "imagen sube cada print" (operador ingresa mm
  // positivos en el modal de drift cuando observa subida en CAL #10).
  // Antes la función restaba → amplificaba el drift en lugar de cancelarlo,
  // y cuando offsetY se volvía muy negativo elementos con Y baseline chico
  // (BITMAP Y=2, Vto Y=12, etc.) caían en Y<0 y el firmware los wrappeaba
  // al fondo del label — efecto "unos suben otros bajan" reportado.
  var compensacion = Math.round(driftDots * printsCount);
  var offset = offsetBase + compensacion;
  // Clamp defensivo: nunca permitir que offset+Y baseline mínima (=2 del BITMAP)
  // caiga bajo 0. Si la compensación es tan agresiva que rompería el layout
  // hacia abajo (Y > 200 = fondo del label), se trunca a 50 dots (~6mm) máximo —
  // a partir de ahí hay que recalibrar el rollo en serio.
  if (offset < -1) offset = -1;
  if (offset > 50) offset = 50;
  return offset;
}

// Incrementa el contador después de un print exitoso. Llamar 1 vez por print.
function _incrementarPrintsCount(qty) {
  qty = qty || 1;
  var props = PropertiesService.getScriptProperties();
  var actual = parseInt(props.getProperty('ADHESIVO_PRINTS_DESDE_CAL')) || 0;
  props.setProperty('ADHESIVO_PRINTS_DESDE_CAL', String(actual + qty));
}

// Estado actual de calibración — para frontend.
function estadoCalibracionRollo() {
  try {
    var props = PropertiesService.getScriptProperties();
    var calibrado   = props.getProperty('ADHESIVO_ROLLO_CALIBRADO') === 'true';
    var driftDots   = parseFloat(props.getProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT')) || 0;
    var printsCount = parseInt  (props.getProperty('ADHESIVO_PRINTS_DESDE_CAL'))     || 0;
    var offsetBase  = parseFloat(props.getProperty('ADHESIVO_OFFSET_Y'))             || 0;
    var fechaCal    = props.getProperty('ADHESIVO_FECHA_CALIBRADO') || '';
    var compensacionActual = Math.round(driftDots * printsCount);
    // [v2.13.129] Umbrales ajustados según uso real:
    //   Rollo = 1000 adhesivos, consumo ~400/día → rollo dura ~2.5 días.
    //   Aviso al 80% (800) = "rollo casi terminándose, cambia pronto"
    //   Alerta al 95% (950) = "rollo a punto de acabarse"
    var necesitaRecal      = printsCount > 800;
    var rolloCasiAgotado   = printsCount > 950;
    return { ok: true, data: {
      calibrado:              calibrado,
      driftDotsPorPrint:      driftDots,
      printsDesdeCal:         printsCount,
      offsetBase:             offsetBase,
      compensacionAcumulada:  compensacionActual,
      offsetEfectivoProximoPrint: offsetBase + compensacionActual,  // [v2.13.144] suma, no resta
      fechaCalibrado:         fechaCal,
      necesitaRecalibrar:     necesitaRecal,
      rolloCasiAgotado:       rolloCasiAgotado,
      driftMmPorPrint:        +(driftDots / 8).toFixed(3),
      driftConfigurado:       driftDots > 0,
      capacidadRollo:         1000,
      umbralCasiAgotado:      950
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// imprimirCalibradoresAdhesivo — manda N adhesivos de prueba con
// regla vertical. El operador después medirá el desvío del print #N
// y llamará aplicarDriftDetectado.
//
// Cada calibrador tiene:
//   - Regla vertical IZQ y DER con marcas cada 1mm (8 dots)
//   - Marcas más largas cada 5mm
//   - Número "CAL #N" centrado para identificarlo
//   - Borde superior e inferior con "tick" para alinear visualmente
// ────────────────────────────────────────────────────────────────────
function imprimirCalibradoresAdhesivo(params) {
  try {
    params = params || {};
    var cantidad = parseInt(params.cantidad) || 10;
    if (cantidad < 1 || cantidad > 30) {
      return { ok: false, error: 'cantidad debe estar entre 1 y 30' };
    }
    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var printerId;
    try { printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN'); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }

    var auth = 'Basic ' + Utilities.base64Encode(apiKey + ':');
    var resultados = [];

    // Mandar 1 job por calibrador (para que el sistema OFFSET acumulativo
    // los compense correctamente cuando se midan).
    for (var i = 1; i <= cantidad; i++) {
      var bytes = _buildTSPLCalibrador(i, cantidad);
      var b64 = Utilities.base64Encode(bytes);
      var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
        method: 'post',
        headers: { 'Authorization': auth },
        contentType: 'application/json',
        payload: JSON.stringify({
          printerId: parseInt(printerId),
          title: 'Calibrador #' + i + '/' + cantidad,
          contentType: 'raw_base64',
          content: b64,
          source: 'warehouseMos-calibrador'
        }),
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      if (code === 201) {
        resultados.push({ i: i, ok: true, jobId: JSON.parse(resp.getContentText()) });
        // [AUDIT FIX] Calibradores NO incrementan contador — están "fuera de cuenta".
        // El operador después aplicará drift y eso resetea el contador para que
        // los próximos prints reales arranquen limpios.
      } else {
        resultados.push({ i: i, ok: false, error: 'HTTP ' + code });
      }
    }
    var okCount = resultados.filter(function(r) { return r.ok; }).length;
    return { ok: true, data: {
      enviados: okCount,
      total: cantidad,
      detalle: resultados,
      mensaje: 'Mirá el calibrador #' + cantidad + '. ¿Cuántos mm se subió la regla? Ingresalo en "Aplicar drift detectado".'
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Construye el TSPL del calibrador #N con regla vertical en ambos lados.
// El calibrador NO incluye OFFSET ni drift compensation — es PRECISAMENTE
// para medir cuánto se mueve sin compensación. Por eso usa offsetBase
// directo (NO _calcularOffsetEfectivoParaPrint que incluye drift).
function _buildTSPLCalibrador(numero, total) {
  var props = PropertiesService.getScriptProperties();
  var gapMm   = parseFloat(props.getProperty('ADHESIVO_GAP_MM'))  || 2;
  var density = parseInt  (props.getProperty('ADHESIVO_DENSITY')) || 8;
  var speed   = parseInt  (props.getProperty('ADHESIVO_SPEED'))   || 4;
  // [AUDIT FIX] Solo offset manual base. SIN compensación de drift acumulado.
  var offsetY = parseFloat(props.getProperty('ADHESIVO_OFFSET_Y')) || 0;

  var header = [
    'SIZE 50 mm,25 mm',
    'GAP ' + gapMm + ' mm,0 mm',
    'DIRECTION 1',
    'DENSITY ' + density,
    'SPEED ' + speed,
    'CLS'
  ].join('\r\n') + '\r\n';
  var bytes = _strToBytesEtq(header);

  // ── REGLAS VERTICALES (IZQ X=0..8, DER X=392..400) ──
  // Cada 1mm = 8 dots. Adhesivo 25mm de alto = 200 dots → 25 marcas.
  // Marca normal: 4 dots ancho, 1 dot alto
  // Marca cada 5mm: 8 dots ancho
  // Marca cada 10mm: 12 dots ancho + número
  for (var mm = 0; mm <= 25; mm++) {
    var y = 2 + mm * 8;  // y=2 para mm=0 (tope), y=2+200=202 (cerca del fondo)
    if (y > 198) break;
    var anchoMarca;
    if (mm % 10 === 0)      anchoMarca = 12;
    else if (mm % 5 === 0)  anchoMarca = 8;
    else                    anchoMarca = 4;
    // Marca izquierda
    bytes = bytes.concat(_strToBytesEtq('BAR 0,' + y + ',' + anchoMarca + ',1\r\n'));
    // Marca derecha
    bytes = bytes.concat(_strToBytesEtq('BAR ' + (400 - anchoMarca) + ',' + y + ',' + anchoMarca + ',1\r\n'));
    // Números cada 5mm (font 1 = 8 dots ancho)
    if (mm % 5 === 0 && mm > 0 && mm < 25) {
      bytes = bytes.concat(_strToBytesEtq('TEXT 14,' + (y - 4) + ',"1",0,1,1,"' + mm + '"\r\n'));
      bytes = bytes.concat(_strToBytesEtq('TEXT 370,' + (y - 4) + ',"1",0,1,1,"' + mm + '"\r\n'));
    }
  }

  // ── INDICADOR "0mm" en el tope: marca BLANCA prominente ──
  // Una caja blanca con "0" en el centro arriba para que el operador
  // sepa que ESE es el punto de referencia "donde debería empezar".
  bytes = bytes.concat(_strToBytesEtq('BAR 30,2,20,12\r\n'));     // caja negra
  bytes = bytes.concat(_strToBytesEtq('BAR 32,4,16,8\r\n'));      // hueco blanco
  bytes = bytes.concat(_strToBytesEtq('TEXT 33,5,"1",0,1,1,"0mm"\r\n'));
  // Igual al final
  bytes = bytes.concat(_strToBytesEtq('BAR 350,2,32,12\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR 352,4,28,8\r\n'));
  bytes = bytes.concat(_strToBytesEtq('TEXT 353,5,"1",0,1,1,"0mm"\r\n'));

  // ── CAJA CENTRAL con "CAL #N/T" ──
  // Posición centro: X≈120..280, Y≈85..120
  // Font 3 (16 wide × 24 tall) bien visible.
  var label = 'CAL #' + numero + '/' + total;
  var labelW = label.length * 16;
  var labelX = Math.floor((400 - labelW) / 2);
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + labelX + ',88,"3",0,1,1,"' + label + '"\r\n'));

  var sub = 'mide el desvio mm';
  var subW = sub.length * 8;
  var subX = Math.floor((400 - subW) / 2);
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + subX + ',120,"1",0,1,1,"' + sub + '"\r\n'));

  // ── Flecha SUR indicando "abajo es el final del adhesivo" ──
  // Útil para que el operador sepa por dónde mirar el desvío.
  bytes = bytes.concat(_strToBytesEtq('TEXT 180,150,"3",0,1,1,"v"\r\n'));

  bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
  return bytes;
}

// ────────────────────────────────────────────────────────────────────
// aplicarDriftDetectado — operador midió el desvío del calibrador #N
// con la regla vertical. Calculamos drift dots/print y guardamos.
//
// params: { mmDesviados: <número>, basadoEnPrints: <número, default 10> }
// ────────────────────────────────────────────────────────────────────
function aplicarDriftDetectado(params) {
  try {
    params = params || {};
    var mm     = parseFloat(params.mmDesviados);
    var prints = parseInt(params.basadoEnPrints) || 10;
    // [v2.13.144] Aceptar parámetro `direccion` ('arriba'|'abajo') para que la
    // UI guíe al operador en qué signo aplicar. Default 'arriba' = caso típico.
    var direccion = String(params.direccion || 'arriba').toLowerCase();
    if (isNaN(mm)) {
      return { ok: false, error: 'mmDesviados debe ser un número' };
    }
    if (prints < 1) {
      return { ok: false, error: 'basadoEnPrints debe ser >= 1' };
    }
    // Convertir mm a dots: 1mm = 8 dots a 203 DPI
    // mm SIEMPRE positivo (operador ingresa magnitud, no signo).
    // El SENTIDO lo marca `direccion`:
    //   - 'arriba' = drift natural sube → driftDots POSITIVO → compensación SUMA a Y (baja)
    //   - 'abajo'  = drift natural baja → driftDots NEGATIVO → compensación RESTA a Y (sube)
    var mmAbs = Math.abs(mm);
    var driftDotsPorPrint = (mmAbs / prints) * 8;
    if (direccion === 'abajo') driftDotsPorPrint = -driftDotsPorPrint;
    driftDotsPorPrint = Math.round(driftDotsPorPrint * 10) / 10;

    var props = PropertiesService.getScriptProperties();
    props.setProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT', String(driftDotsPorPrint));
    // [AUDIT FIX] Resetear contador. La compensación se aplica DESDE EL
    // PRÓXIMO print real, no a los calibradores ya impresos.
    props.setProperty('ADHESIVO_PRINTS_DESDE_CAL', '0');

    return { ok: true, data: {
      mmDesviados:      mmAbs,
      direccion:        direccion,
      basadoEnPrints:   prints,
      driftDotsPorPrint: driftDotsPorPrint,
      driftMmPorPrint:  +(driftDotsPorPrint / 8).toFixed(3),
      mensaje: 'Drift ' + direccion + ' aplicado: ' + driftDotsPorPrint
               + ' dots/print. Próximos prints se compensan automáticamente.'
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// [v2.13.144] Reset de emergencia — limpia drift, offset y contador.
// Útil cuando el drift quedó configurado mal (ej. signo invertido) y los
// adhesivos salen completamente fuera de lugar. NO hace GAPDETECT — solo
// resetea las properties del software.
function resetearDriftEmergencia() {
  try {
    var props = PropertiesService.getScriptProperties();
    var antes = {
      drift:   parseFloat(props.getProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT')) || 0,
      offset:  parseFloat(props.getProperty('ADHESIVO_OFFSET_Y'))             || 0,
      prints:  parseInt  (props.getProperty('ADHESIVO_PRINTS_DESDE_CAL'))     || 0
    };
    props.setProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT', '0');
    props.setProperty('ADHESIVO_OFFSET_Y',             '0');
    props.setProperty('ADHESIVO_PRINTS_DESDE_CAL',     '0');
    return { ok: true, data: {
      antes: antes,
      ahora: { drift: 0, offset: 0, prints: 0 },
      mensaje: 'Drift y offset reseteados a 0. Próximos prints saldrán sin compensación. '
             + 'Si querés re-detectar drift correctamente: imprimí 10 calibradores y aplicá la dirección correcta.'
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// ajustarDriftManual — operador setea manualmente el drift sin auto-detect
// ────────────────────────────────────────────────────────────────────
function ajustarDriftManual(params) {
  try {
    params = params || {};
    var driftDots = parseFloat(params.driftDotsPorPrint);
    if (isNaN(driftDots)) {
      return { ok: false, error: 'driftDotsPorPrint debe ser un número' };
    }
    PropertiesService.getScriptProperties()
      .setProperty('ADHESIVO_DRIFT_DOTS_POR_PRINT', String(driftDots));
    return { ok: true, data: { driftDotsPorPrint: driftDots } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// resetearContadorPrints — uso administrativo, sin recalibrar.
// ────────────────────────────────────────────────────────────────────
function resetearContadorPrints() {
  try {
    PropertiesService.getScriptProperties().setProperty('ADHESIVO_PRINTS_DESDE_CAL', '0');
    return { ok: true, data: { printsDesdeCal: 0 } };
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


// ════════════════════════════════════════════════════════════════════
// [v2.13.108] SISTEMA DE LOTES DE IMPRESIÓN DE ADHESIVOS
// ════════════════════════════════════════════════════════════════════
//
// MOTIVACIÓN — Drift acumulativo en impresoras térmicas TSC:
//   Cuando se imprimen muchos adhesivos seguidos en jobs separados, cada
//   feed acumula un error de ~0.1mm por la diferencia entre el GAP
//   configurado y el GAP real del rollo. En 200 prints = ~20mm de drift,
//   suficiente para que la última etiqueta salga por la mitad.
//
// ARQUITECTURA — Lotes con sub-jobs:
//   Un "lote" representa un pedido del usuario (ej "200 adhesivos de COCO").
//   Se divide en sub-jobs de N=10 etiquetas cada uno (config en spec.json).
//   Cada sub-job es un job PrintNode independiente, lo que permite:
//     1. Tracking exacto de progreso (10/200, 20/200, ...)
//     2. Detección de OUT_OF_PAPER por error en sub-job
//     3. Reanudación limpia con GAPDETECT al cambiar rollo
//     4. Drift entre sub-jobs ~0 (mismo rollo, no se necesita re-cal)
//
// GAPDETECT — Cuándo se aplica:
//   - Al inicio del primer sub-job de cada lote (calibra al rollo actual)
//   - Al reanudar después de OUT_OF_PAPER (rollo nuevo)
//   - NO entre sub-jobs del mismo rollo (innecesario, cuesta 1 etiqueta)
//
// FRONTEND ORQUESTA — Backend stateless:
//   El frontend manda 1 sub-job a la vez, espera respuesta, manda el
//   siguiente. Esto evita los timeouts de GAS (max 6min) y da control
//   fino del flujo. La sheet LOTES_ADHESIVO actúa como source of truth.
//
// SHEET LOTES_ADHESIVO — Columnas:
//   idLote · fechaCreacion · fechaUltimoUpdate · usuario · origen
//   codigoBarra · descripcion · vto · totalEtq · completadas · subJobSize
//   status · ultimoError · ultimoPrintNodeJobId · printerId
//
// ESTADOS:
//   CREADO · CALIBRANDO · IMPRIMIENDO · PAUSADO_USUARIO
//   PAUSADO_OUT_PAPER · PAUSADO_ERROR · COMPLETADO · CANCELADO

// [v2.13.118] tipoEtiqueta agregado — sheet ahora soporta ADHESIVO_ENVASADO,
// MEMBRETE_ME, MEMBRETE_WH, CALIBRADOR. itemsJson permite payloads complejos
// (cola múltiple, lista de códigos del WH, precio del ME).
var LOTES_ADHESIVO_HEADERS = [
  'idLote', 'fechaCreacion', 'fechaUltimoUpdate', 'usuario', 'origen',
  'codigoBarra', 'descripcion', 'vto', 'totalEtq', 'completadas',
  'subJobSize', 'status', 'ultimoError', 'ultimoPrintNodeJobId', 'printerId',
  'tipoEtiqueta', 'itemsJson'
];

// Setup idempotente — crea sheet si no existe, agrega columnas faltantes.
// Llamar UNA VEZ desde editor GAS o cuando se cambien los headers.
function setupLotesAdhesivo() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('LOTES_ADHESIVO');
  if (!sheet) {
    sheet = ss.insertSheet('LOTES_ADHESIVO');
    sheet.getRange(1, 1, 1, LOTES_ADHESIVO_HEADERS.length)
         .setValues([LOTES_ADHESIVO_HEADERS])
         .setFontWeight('bold').setBackground('#1e293b').setFontColor('#fbbf24');
    sheet.setFrozenRows(1);
    // Cols string que deben preservar ceros a la izq (codigoBarra)
    sheet.getRange('F:F').setNumberFormat('@');  // codigoBarra
    sheet.getRange('A:A').setNumberFormat('@');  // idLote
    Logger.log('[setupLotesAdhesivo] sheet creada');
  } else {
    // Verificar que tenga todas las columnas; si faltan, agregarlas al final
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var missing = LOTES_ADHESIVO_HEADERS.filter(function(h) { return existing.indexOf(h) < 0; });
    if (missing.length) {
      var startCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
      Logger.log('[setupLotesAdhesivo] agregadas columnas: ' + missing.join(', '));
    }
  }
  return { ok: true, headers: LOTES_ADHESIVO_HEADERS };
}

function _getSheetLotesAdhesivo() {
  var sheet = getSheet('LOTES_ADHESIVO');
  if (!sheet) {
    setupLotesAdhesivo();
    sheet = getSheet('LOTES_ADHESIVO');
  }
  return sheet;
}

// Construye un objeto fila desde un patch.
// [v2.13.136 FIX] Lee los headers REALES de la sheet para respetar el orden
// real. Antes usaba LOTES_ADHESIVO_HEADERS constante; si la sheet tenía las
// columnas en otro orden (ej: tipoEtiqueta agregada al final), los valores
// se escribían en celdas equivocadas. Bug síntoma: lotes MEMBRETE_WH se
// guardaban con tipoEtiqueta vacía → filtro 'ADHESIVO_ENVASADO' los incluía.
function _filaLoteAdhesivo(patch, sheet) {
  var realHeaders = LOTES_ADHESIVO_HEADERS;
  if (sheet) {
    try {
      var hdrRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      realHeaders = hdrRow.map(function(h) { return String(h || '').trim(); }).filter(Boolean);
    } catch(_){}
  }
  return realHeaders.map(function(h) {
    var v = patch[h];
    return v === undefined ? '' : v;
  });
}

// Busca el rowIndex (1-based) de un lote por idLote. -1 si no existe.
// [v2.13.109 AUDIT FIX #4] Protección sheet vacía: getRange('A2:A') falla
// con "Row 2 invalid" en algunos casos cuando lastRow=1. Verificamos primero.
function _findLoteRow(sheet, idLote) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idLote)) return i + 2;
  }
  return -1;
}

function _readLote(sheet, rowIndex) {
  var values = sheet.getRange(rowIndex, 1, 1, LOTES_ADHESIVO_HEADERS.length).getValues()[0];
  var obj = {};
  LOTES_ADHESIVO_HEADERS.forEach(function(h, i) { obj[h] = values[i]; });
  // Numéricos como Number
  obj.totalEtq    = parseInt(obj.totalEtq) || 0;
  obj.completadas = parseInt(obj.completadas) || 0;
  obj.subJobSize  = parseInt(obj.subJobSize) || 10;
  return obj;
}

function _patchLote(sheet, rowIndex, patch) {
  patch.fechaUltimoUpdate = new Date().toISOString();
  // Actualizar solo las columnas que vienen en patch
  Object.keys(patch).forEach(function(k) {
    var colIdx = LOTES_ADHESIVO_HEADERS.indexOf(k);
    if (colIdx >= 0) sheet.getRange(rowIndex, colIdx + 1).setValue(patch[k]);
  });
}

// ────────────────────────────────────────────────────────────────────
// 1) crearLoteAdhesivo
//    POST { codigoBarra, total, usuario, origen, vto?, idempotencyKey, descripcion? }
//    → { ok, data: { idLote, total, completadas, subJobSize, status, vto } }
// ────────────────────────────────────────────────────────────────────
function crearLoteAdhesivo(params) {
  try {
    var codigoBarra = String(params.codigoBarra || '').trim();
    var total       = parseInt(params.total) || 0;
    var usuario     = String(params.usuario || '').trim();
    var origen      = String(params.origen || 'MOS').trim().toUpperCase();
    if (!codigoBarra) return { ok: false, error: 'codigoBarra requerido' };
    if (total <= 0)   return { ok: false, error: 'total debe ser > 0' };

    // [v2.13.115 BUG FIX CRÍTICO] DEDUPE por idempotencyKey.
    // El usuario reportó que se imprimieron 2 lotes de 10 etiquetas (20 total)
    // cuando había pedido solo uno. Causa: si el modal demora y el usuario
    // reintenta, O si el primer intento falló parcialmente y el segundo
    // tuvo éxito, se crearon 2 lotes distintos.
    // Fix: si llega un idempotencyKey que ya existe en la sheet, retornar
    // el lote existente sin crear duplicado.
    var idempotencyKey = String(params.idempotencyKey || '').trim();
    if (idempotencyKey) {
      // [v2.13.117 AUDIT FIX] endsWith en lugar de indexOf para evitar
      // false positives con idempotencyKeys que coincidan parcialmente
      // con sufijos aleatorios de otro idLote.
      // Patrón: idLote = LA<timestamp>_<idempotencyKey>
      var sheetCheck = _getSheetLotesAdhesivo();
      var allRows = _sheetToObjects(sheetCheck);
      var sufijo = '_' + idempotencyKey;
      var found = allRows.find(function(r) {
        var id = String(r.idLote || '');
        return id.length >= sufijo.length && id.substring(id.length - sufijo.length) === sufijo;
      });
      if (found) {
        // Lote ya existe — retornarlo en lugar de crear duplicado.
        return { ok: true, data: {
          idLote:      String(found.idLote),
          total:       parseInt(found.totalEtq) || 0,
          completadas: parseInt(found.completadas) || 0,
          subJobSize:  parseInt(found.subJobSize) || 10,
          status:      String(found.status || 'CREADO'),
          vto:         String(found.vto || ''),
          descripcion: String(found.descripcion || ''),
          printerId:   String(found.printerId || ''),
          deduped:     true
        }};
      }
    }

    // Resolver descripción (lookup productos)
    var descripcion = String(params.descripcion || '').trim();
    if (!descripcion) {
      try {
        var prods = _sheetToObjects(getProductosSheet());
        var p = prods.find(function(x) {
          return String(x.codigoBarra) === codigoBarra || String(x.idProducto) === codigoBarra;
        });
        if (p) descripcion = String(p.descripcion || p.nombre || '');
      } catch(_) {}
    }

    // Vto: si vino explícito, formatear; sino calcular auto +1 año
    var vto;
    if (params.vto) {
      vto = String(params.vto);
    } else if (params.fechaVencimiento) {
      var d = new Date(params.fechaVencimiento);
      vto = MESES_ES[d.getMonth()] + '/' + d.getFullYear();
    } else {
      vto = _calcVencimientoEtq(new Date());
    }

    // Resolver printerId
    var printerId;
    try { printerId = String(getPrinterNodeId('ADHESIVO', 'ALMACEN')); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }

    // Cargar config sub-job size desde Script Properties (override) o default 10
    var subJobSize = parseInt(
      PropertiesService.getScriptProperties().getProperty('ADHESIVO_SUB_JOB_SIZE')
    ) || 10;

    // [v2.13.118] tipoEtiqueta para diferenciar adhesivo de envasado vs
    // membrete ME vs membrete WH. Default ADHESIVO_ENVASADO (back-compat).
    var tipoEtiqueta = String(params.tipoEtiqueta || 'ADHESIVO_ENVASADO').toUpperCase();
    var itemsJson    = params.itemsJson ? String(params.itemsJson) : '';

    var sheet = _getSheetLotesAdhesivo();
    // [v2.13.115] idLote incluye idempotencyKey para dedupe efectivo.
    // Si no hay idempotencyKey, fallback al patrón viejo (más débil).
    var idLote = idempotencyKey
      ? ('LA' + new Date().getTime() + '_' + idempotencyKey)
      : ('LA' + new Date().getTime() + Math.random().toString(36).substr(2, 4).toUpperCase());
    var now = new Date().toISOString();
    var fila = _filaLoteAdhesivo({
      idLote:              idLote,
      fechaCreacion:       now,
      fechaUltimoUpdate:   now,
      usuario:             usuario,
      origen:              origen,
      codigoBarra:         codigoBarra,
      descripcion:         descripcion,
      vto:                 vto,
      totalEtq:            total,
      completadas:         0,
      subJobSize:          subJobSize,
      // [v2.13.141] FIRE-AND-FORGET: nace ENCOLADO para que el trigger backend
      // procese todo solo. Antes era CREADO y dependía del frontend orquestando
      // (WhLoteAdhesivo._orquestar) — si el operador cerraba la app, quedaba
      // a medio imprimir. Ahora la app puede cerrarse sin perder nada.
      status:              'ENCOLADO',
      ultimoError:         '',
      ultimoPrintNodeJobId: '',
      printerId:           printerId,
      tipoEtiqueta:        tipoEtiqueta,
      itemsJson:           itemsJson
    }, sheet);  // [v2.13.136] pasar sheet para respetar orden REAL de columnas
    sheet.appendRow(fila);

    // [v2.13.141] Auto-instalar trigger es crítico ahora — sin trigger los
    // lotes se quedan en ENCOLADO infinitos.
    try { _asegurarTriggerLotes(); } catch(_){}

    return { ok: true, data: {
      idLote:       idLote,
      total:        total,
      completadas:  0,
      subJobSize:   subJobSize,
      status:       'ENCOLADO',  // [v2.13.141] cambio: era CREADO
      vto:          vto,
      descripcion:  descripcion,
      printerId:    printerId,
      tipoEtiqueta: tipoEtiqueta
    }};
  } catch (e) {
    return { ok: false, error: e.message, stack: e.stack };
  }
}

// ────────────────────────────────────────────────────────────────────
// 2) imprimirSubLoteAdhesivo
//    POST { idLote, requireGapDetect? }
//    → { ok, data: { idLote, completadas, total, status, qtyImpresa, printNodeJobId } }
//
//    Lógica:
//      a) Lee lote por idLote
//      b) Calcula qty = min(subJobSize, total - completadas)
//      c) Si requireGapDetect O (status==CREADO) O (status==PAUSADO_OUT_PAPER):
//           agregar GAPDETECT al TSPL2
//      d) Construye TSPL2 con qty etiquetas
//      e) Manda a PrintNode
//      f) Si OK: completadas += qty, status = IMPRIMIENDO (o COMPLETADO si fin)
//      g) Si error: status = PAUSADO_ERROR / PAUSADO_OUT_PAPER
// ────────────────────────────────────────────────────────────────────
function imprimirSubLoteAdhesivo(params) {
  try {
    var idLote = String(params.idLote || '').trim();
    if (!idLote) return { ok: false, error: 'idLote requerido' };

    // [v2.13.129 FIX CRÍTICO] LockService para evitar doble sub-job concurrente.
    // Bug previo: 2 llamadas concurrentes (timeout+retry o 2 instancias WH) leían
    // el mismo `completadas`, ambas enviaban un sub-job de qty=10 a PrintNode,
    // y ambas actualizaban completadas al mismo valor. Resultado: 20 etiquetas
    // físicas pero el sheet solo registra 10. Cliente reportó "22 piden, salen 34"
    // = duplicación del último sub-job.
    // [v2.13.131] _lock reasignable a null cuando se libera anticipadamente
    // antes del polling PrintNode (para no bloquear toda la app 25s).
    var _lock = LockService.getScriptLock();
    try { _lock.waitLock(8000); } catch(e) { return { ok: false, error: 'Sistema ocupado, reintenta en 5s' }; }

    var sheet = _getSheetLotesAdhesivo();
    var rowIdx = _findLoteRow(sheet, idLote);
    if (rowIdx < 0) { try { _lock.releaseLock(); } catch(_){} return { ok: false, error: 'Lote no encontrado: ' + idLote }; }

    var lote = _readLote(sheet, rowIdx);

    // Validar estado: no se puede imprimir si cancelado/completado
    if (lote.status === 'CANCELADO')  { try { _lock.releaseLock(); } catch(_){} return { ok: false, error: 'Lote cancelado' }; }
    if (lote.status === 'COMPLETADO') { try { _lock.releaseLock(); } catch(_){} return { ok: true, data: { idLote: idLote, completadas: lote.completadas, total: lote.totalEtq, status: 'COMPLETADO', qtyImpresa: 0 } }; }

    // [v2.13.129 FIX] Si el lote está IMPRIMIENDO/CALIBRANDO con update reciente
    // (<35s), otro proceso lo está manejando. Rechazar para no duplicar.
    if (lote.status === 'IMPRIMIENDO' || lote.status === 'CALIBRANDO') {
      var lastUpdMs = 0;
      try { lastUpdMs = new Date(lote.fechaUltimoUpdate || '').getTime(); } catch(_){}
      var ageS = (Date.now() - lastUpdMs) / 1000;
      if (lastUpdMs && ageS < 35) {
        try { _lock.releaseLock(); } catch(_){}
        return { ok: true, data: {
          idLote: idLote, completadas: lote.completadas, total: lote.totalEtq,
          status: lote.status, qtyImpresa: 0,
          skipped: 'sub_job_concurrente_en_curso_' + Math.round(ageS) + 's'
        }};
      }
    }

    var qty = Math.min(lote.subJobSize, lote.totalEtq - lote.completadas);
    if (qty <= 0) {
      _patchLote(sheet, rowIdx, { status: 'COMPLETADO' });
      try { _lock.releaseLock(); } catch(_){}
      return { ok: true, data: { idLote: idLote, completadas: lote.completadas, total: lote.totalEtq, status: 'COMPLETADO', qtyImpresa: 0 } };
    }

    // [v2.13.143] GAPDETECT solo cuando se pide EXPLÍCITAMENTE.
    // Antes se hacía automático al primer sub-job (status=CREADO) o al
    // reanudar OUT_OF_PAPER → consumía 3 etiquetas blancas inútiles cada vez.
    // Según la estrategia acordada: calibrar UNA VEZ al cambiar rollo via
    // botón "🔧 Calibrar rollo nuevo" del modal calibrador → drift compensa
    // automático en cada print posterior. Sin más GAPDETECTs gratis.
    var requireGapDetect = params.requireGapDetect === true;

    // Construir el producto para el TSPL (cargar tokens y allEnv)
    var producto;
    try {
      var prods = _sheetToObjects(getProductosSheet());
      var p = prods.find(function(x) {
        return String(x.codigoBarra) === String(lote.codigoBarra) || String(x.idProducto) === String(lote.codigoBarra);
      });
      if (!p) return { ok: false, error: 'Producto no encontrado: ' + lote.codigoBarra };
      producto = {
        codigoBarra: String(p.codigoBarra || p.idProducto),
        descripcion: String(lote.descripcion || p.descripcion || p.nombre || '')
      };
    } catch (e) {
      return { ok: false, error: 'Error cargando producto: ' + e.message };
    }

    // Marcar como CALIBRANDO si vamos a GAPDETECT, sino IMPRIMIENDO
    _patchLote(sheet, rowIdx, {
      status: requireGapDetect ? 'CALIBRANDO' : 'IMPRIMIENDO'
    });

    // Generar TSPL2 — pasamos la fecha del lote (no recalcular)
    // Para que _buildTSPLEtq use el vto del lote, hay que pasarle una
    // fechaEnvasado = vto - 1año.
    var fechaParaCalc = _vtoStringAFechaEnvasado(lote.vto);
    var allEnv = _getAllEnvasablesTokens();
    var bytes = _buildTSPLEtqConGapDetect(producto, fechaParaCalc, qty, allEnv, requireGapDetect);

    // [v2.13.109 AUDIT FIX #1] Pasar byte array DIRECTO a base64Encode.
    // El intento previo con String.fromCharCode + base64Encode(string)
    // expandía bytes >127 como UTF-8 (2 bytes), corrompiendo el bitmap.
    // Utilities.base64Encode(Byte[]) acepta el array crudo. Mismo método
    // que usa el _imprimirEtiquetasEnvasado legacy que funciona.
    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var b64 = Utilities.base64Encode(bytes);
    var auth = 'Basic ' + Utilities.base64Encode(apiKey + ':');

    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: { 'Authorization': auth },
      contentType: 'application/json',
      payload: JSON.stringify({
        printerId: parseInt(lote.printerId),
        title: 'Adhesivo ' + producto.codigoBarra + ' (' + qty + ') lote=' + idLote,
        contentType: 'raw_base64',
        content: b64,
        source: 'warehouseMos-lote-' + idLote
      }),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code !== 201) {
      var errTxt = resp.getContentText();
      _patchLote(sheet, rowIdx, {
        status:      'PAUSADO_ERROR',
        ultimoError: 'PrintNode HTTP ' + code + ': ' + errTxt.substring(0, 200)
      });
      return { ok: false, error: 'PrintNode HTTP ' + code, detalle: errTxt, status: 'PAUSADO_ERROR' };
    }

    var printNodeJobId = JSON.parse(resp.getContentText());

    // [v2.13.131 FIX CRÍTICO] Liberar lock ANTES del polling PrintNode (25s).
    // Antes el lock se mantenía durante todo el polling → bloqueaba TODOS los
    // endpoints WH durante 25-33s por cada sub-job. Operarios reportaban "app
    // congelada" durante impresiones. La protección anti-concurrencia ya hizo
    // su trabajo (marcamos IMPRIMIENDO + check 35s); el polling no necesita lock.
    try { _lock.releaseLock(); } catch(_){}
    _lock = null;  // flag para que finally no la libere de nuevo

    // [v2.13.109 AUDIT FIX #3] Polling al estado del job para detectar
    // OUT_OF_PAPER REAL. PrintNode retorna 201 apenas el job entra a la
    // cola; el OUT_OF_PAPER se descubre cuando la impresora intenta
    // ejecutar. Sin este polling el backend creía que TODO se imprimió.
    //
    // Strategy: polling cada 2s hasta done/error/timeout 25s.
    // Estados PrintNode: new → downloading → printing → done / error / expired
    var pollResult = _esperarFinJobPrintNode(printNodeJobId, auth, 25000, 2000);

    if (pollResult.estado === 'error') {
      var msg = String(pollResult.mensaje || '').toLowerCase();
      var esOutOfPaper = msg.indexOf('paper') >= 0
                      || msg.indexOf('media') >= 0
                      || msg.indexOf('label') >= 0
                      || msg.indexOf('out of') >= 0;
      _patchLote(sheet, rowIdx, {
        status:               esOutOfPaper ? 'PAUSADO_OUT_PAPER' : 'PAUSADO_ERROR',
        ultimoError:          pollResult.mensaje || 'Error en job ' + printNodeJobId,
        ultimoPrintNodeJobId: String(printNodeJobId)
      });
      return {
        ok:     false,
        error:  pollResult.mensaje || 'Error de impresión',
        status: esOutOfPaper ? 'PAUSADO_OUT_PAPER' : 'PAUSADO_ERROR'
      };
    }

    if (pollResult.estado === 'timeout') {
      // No tuvimos confirmación en 25s. Asumimos que sigue procesando.
      // Marcamos como impreso optimista pero registramos warning.
      // Si el siguiente sub-job falla por OUT_OF_PAPER, el operario reanuda.
      Logger.log('[Lote ' + idLote + '] timeout polling job ' + printNodeJobId + ', estado: ' + pollResult.ultimoEstado);
    }

    // Estado = done (o timeout asumido OK) → contar como impreso
    var nuevasCompletadas = lote.completadas + qty;
    var nuevoStatus = nuevasCompletadas >= lote.totalEtq ? 'COMPLETADO' : 'IMPRIMIENDO';

    _patchLote(sheet, rowIdx, {
      completadas:          nuevasCompletadas,
      status:               nuevoStatus,
      ultimoPrintNodeJobId: String(printNodeJobId),
      ultimoError:          ''
    });

    // [v2.13.118] Drift compensation: incrementar contador de prints
    // desde la última calibración. Próximo TSPL aplicará más offset.
    try { _incrementarPrintsCount(qty); } catch(_) {}

    return { ok: true, data: {
      idLote:            idLote,
      completadas:       nuevasCompletadas,
      total:             lote.totalEtq,
      status:            nuevoStatus,
      qtyImpresa:        qty,
      printNodeJobId:    printNodeJobId,
      gapDetectAplicado: requireGapDetect,
      pollEstado:        pollResult.estado
    }};
  } catch (e) {
    return { ok: false, error: e.message, stack: e.stack };
  } finally {
    // [v2.13.131] Solo liberar si aún tenemos el lock (puede haberse liberado
    // antes del polling). _lock=null indica que ya fue liberado.
    if (_lock) { try { _lock.releaseLock(); } catch(_){} }
  }
}

// Helper [v2.13.109] — polling al estado de un job PrintNode.
// Retorna { estado: 'done'|'error'|'timeout', mensaje, ultimoEstado }
// 'done'    → impresión confirmada exitosa
// 'error'   → PrintNode/impresora reportó error (mensaje contiene detalles)
// 'timeout' → no tuvimos confirmación en maxWaitMs (asumir impresión OK provisional)
function _esperarFinJobPrintNode(jobId, auth, maxWaitMs, pollIntervalMs) {
  var startMs = new Date().getTime();
  var ultimoEstado = 'new';
  var ultimoMensaje = '';
  while (new Date().getTime() - startMs < maxWaitMs) {
    Utilities.sleep(pollIntervalMs);
    try {
      var resp = UrlFetchApp.fetch(
        'https://api.printnode.com/printjobs/' + jobId + '/states',
        { headers: { 'Authorization': auth }, muteHttpExceptions: true }
      );
      if (resp.getResponseCode() !== 200) continue;
      var body = JSON.parse(resp.getContentText());
      // El endpoint devuelve array de arrays: [[ { state, message, ... }, ... ]]
      // o array de objetos según versión. Aplanamos.
      var events = [];
      if (Array.isArray(body)) {
        body.forEach(function(b) {
          if (Array.isArray(b)) events = events.concat(b);
          else events.push(b);
        });
      }
      if (events.length === 0) continue;
      var ultimo = events[events.length - 1];
      ultimoEstado  = String(ultimo.state || '').toLowerCase();
      ultimoMensaje = String(ultimo.message || ultimo.data || '');
      if (ultimoEstado === 'done') {
        return { estado: 'done', mensaje: '', ultimoEstado: ultimoEstado };
      }
      if (ultimoEstado === 'error' || ultimoEstado === 'expired') {
        return { estado: 'error', mensaje: ultimoMensaje, ultimoEstado: ultimoEstado };
      }
      // new/queued/downloading/printing → seguir esperando
    } catch (e) {
      Logger.log('[_esperarFinJobPrintNode] error: ' + e.message);
    }
  }
  return { estado: 'timeout', mensaje: 'sin confirmación en ' + maxWaitMs + 'ms', ultimoEstado: ultimoEstado };
}

// Helper — convierte vto string ("ENE/2027") a Date de envasado (= vto - 1 año)
// para que _buildTSPLEtq lo use directamente.
function _vtoStringAFechaEnvasado(vtoStr) {
  if (!vtoStr) return new Date();
  var parts = String(vtoStr).split('/');
  if (parts.length !== 2) return new Date();
  var mesIdx = MESES_ES.indexOf(parts[0].toUpperCase());
  var anio = parseInt(parts[1]);
  if (mesIdx < 0 || !anio) return new Date();
  // vto = primer día del mes N del año Y → fecha envasado = mismo día año - 1
  var d = new Date(anio - 1, mesIdx, 1);
  return d;
}

// Wrapper de _buildTSPLEtq que prepende GAPDETECT al header si corresponde.
// Igual al _buildTSPLEtq original, pero acepta el flag.
function _buildTSPLEtqConGapDetect(producto, fechaEnvasado, unidades, allEnvasables, withGapDetect) {
  var bytes = _buildTSPLEtq(producto, fechaEnvasado, unidades, allEnvasables);
  if (!withGapDetect) return bytes;
  // Prepend "GAPDETECT\r\n" al comienzo. Hay que reconstruir desde el header.
  // _buildTSPLEtq retorna bytes empezando con "SIZE 50 mm,25 mm\r\nGAP...\r\n..."
  // GAPDETECT debe ir ANTES de CLS (para que el sensor se re-mida en este job).
  // Insertamos "GAPDETECT\r\n" antes de "CLS\r\n".
  var prefix = _strToBytesEtq('GAPDETECT\r\n');
  var clsBytes = _strToBytesEtq('CLS\r\n');
  // Buscar la posición del CLS en bytes (es secuencia 'C','L','S','\r','\n')
  for (var i = 0; i < bytes.length - 5; i++) {
    if (bytes[i] === 67 && bytes[i+1] === 76 && bytes[i+2] === 83
        && bytes[i+3] === 13 && bytes[i+4] === 10) {
      return bytes.slice(0, i).concat(prefix).concat(bytes.slice(i));
    }
  }
  // Fallback: prepend al inicio si no encontramos CLS (no debería pasar)
  return prefix.concat(bytes);
}

// ────────────────────────────────────────────────────────────────────
// 3) getEstadoLoteAdhesivo
//    GET ?idLote=...
//    → { ok, data: { idLote, total, completadas, status, ultimoError, vto, descripcion, codigoBarra } }
// ────────────────────────────────────────────────────────────────────
function getEstadoLoteAdhesivo(params) {
  try {
    var idLote = String(params.idLote || '').trim();
    if (!idLote) return { ok: false, error: 'idLote requerido' };
    var sheet = _getSheetLotesAdhesivo();
    var rowIdx = _findLoteRow(sheet, idLote);
    if (rowIdx < 0) return { ok: false, error: 'Lote no encontrado' };
    var lote = _readLote(sheet, rowIdx);
    return { ok: true, data: {
      idLote:        lote.idLote,
      total:         lote.totalEtq,
      completadas:   lote.completadas,
      status:        lote.status,
      ultimoError:   lote.ultimoError || '',
      vto:           lote.vto,
      descripcion:   lote.descripcion,
      codigoBarra:   String(lote.codigoBarra || ''),
      subJobSize:    lote.subJobSize,
      fechaCreacion: lote.fechaCreacion,
      fechaUltimoUpdate: lote.fechaUltimoUpdate,
      usuario:       lote.usuario,
      origen:        lote.origen
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// 4) pausarLoteAdhesivo (manual desde UI)
//    POST { idLote, motivo? }
// ────────────────────────────────────────────────────────────────────
function pausarLoteAdhesivo(params) {
  try {
    var idLote = String(params.idLote || '').trim();
    var sheet = _getSheetLotesAdhesivo();
    var rowIdx = _findLoteRow(sheet, idLote);
    if (rowIdx < 0) return { ok: false, error: 'Lote no encontrado' };
    _patchLote(sheet, rowIdx, {
      status:      'PAUSADO_USUARIO',
      ultimoError: String(params.motivo || '')
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// 5) cancelarLoteAdhesivo
//    POST { idLote }
// ────────────────────────────────────────────────────────────────────
function cancelarLoteAdhesivo(params) {
  try {
    var idLote = String(params.idLote || '').trim();
    var sheet = _getSheetLotesAdhesivo();
    var rowIdx = _findLoteRow(sheet, idLote);
    if (rowIdx < 0) return { ok: false, error: 'Lote no encontrado' };
    _patchLote(sheet, rowIdx, { status: 'CANCELADO' });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────
// diagnosticoLotesAdhesivo — corre desde editor GAS
// Verifica todo lo que necesita el sistema de lotes. Imprime un informe
// en Logger con check ✓/✗ para cada componente.
// ────────────────────────────────────────────────────────────────────
function diagnosticoLotesAdhesivo() {
  var L = function(m) { Logger.log(m); };
  var ok = true;
  L('═══════ DIAGNOSTICO LOTES ADHESIVO ═══════');
  L('Fecha: ' + new Date().toLocaleString());
  L('');

  // 1. Sheet LOTES_ADHESIVO
  L('1. Sheet LOTES_ADHESIVO');
  var sheet = getSpreadsheet().getSheetByName('LOTES_ADHESIVO');
  if (sheet) {
    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var faltan = LOTES_ADHESIVO_HEADERS.filter(function(h){ return hdrs.indexOf(h) < 0; });
    if (faltan.length === 0) {
      L('   [OK] sheet existe con ' + LOTES_ADHESIVO_HEADERS.length + ' headers');
      L('   filas existentes: ' + Math.max(0, sheet.getLastRow() - 1));
    } else {
      L('   [WARN] faltan headers: ' + faltan.join(', '));
      L('   solucion: ejecutar setupLotesAdhesivo()');
      ok = false;
    }
  } else {
    L('   [FAIL] sheet no existe');
    L('   solucion: ejecutar setupLotesAdhesivo()');
    ok = false;
  }
  L('');

  // 2. PRINTNODE_API_KEY
  L('2. PRINTNODE_API_KEY (Script Properties)');
  var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
  if (apiKey) {
    L('   [OK] configurada (' + apiKey.length + ' chars)');
  } else {
    L('   [FAIL] no configurada');
    L('   solucion: Editor GAS > Configuracion del Proyecto > Properties');
    L('             agregar PRINTNODE_API_KEY con la API key de PrintNode');
    ok = false;
  }
  L('');

  // 3. Printer ADHESIVO/ALMACEN
  L('3. Printer ADHESIVO/ALMACEN (sheet IMPRESORAS de MOS)');
  try {
    var printerId = getPrinterNodeId('ADHESIVO', 'ALMACEN');
    L('   [OK] printerId = ' + printerId);
  } catch (e) {
    L('   [FAIL] no encontrada: ' + e.message);
    L('   solucion: en hoja IMPRESORAS de MOS, agregar fila con');
    L('             tipo=ADHESIVO  idZona=ALMACEN  activo=1  printNodeId=<el id>');
    ok = false;
  }
  L('');

  // 4. Properties opcionales con sus valores
  L('4. Properties opcionales (defaults entre parentesis)');
  var props = PropertiesService.getScriptProperties();
  var opcionales = [
    ['ADHESIVO_SUB_JOB_SIZE', '10'],
    ['ADHESIVO_GAP_MM',       '2'],
    ['ADHESIVO_DENSITY',      '8'],
    ['ADHESIVO_SPEED',        '4'],
    ['ADHESIVO_OFFSET_Y',     '0']
  ];
  opcionales.forEach(function(p) {
    var v = props.getProperty(p[0]);
    if (v === null || v === '') {
      L('   ' + p[0] + ' = (default ' + p[1] + ')');
    } else {
      L('   ' + p[0] + ' = ' + v + '   (override del default ' + p[1] + ')');
    }
  });
  L('');

  // 5. Lotes recientes (ultimos 5)
  L('5. Lotes en sheet (ultimos 5)');
  try {
    var rows = _sheetToObjects(_getSheetLotesAdhesivo()).slice(-5).reverse();
    if (rows.length === 0) {
      L('   (vacia, no se ha creado ningun lote todavia)');
    } else {
      rows.forEach(function(r) {
        L('   ' + r.idLote + ' | ' + r.status + ' | ' + r.completadas + '/' + r.totalEtq +
          ' | ' + r.codigoBarra + ' | ' + (r.fechaCreacion || '').substring(0, 16));
      });
    }
  } catch (e) {
    L('   (sheet vacia o error: ' + e.message + ')');
  }
  L('');

  L('═══════ RESULTADO: ' + (ok ? 'TODO OK ✓' : 'HAY ITEMS PENDIENTES — ver arriba') + ' ═══════');
  return { ok: ok };
}

// ────────────────────────────────────────────────────────────────────
// 6) getLotesAdhesivoPendientes
//    GET ?usuario=...&limit=20
//    Devuelve lotes en estado PAUSADO_* o IMPRIMIENDO para reanudar.
// ────────────────────────────────────────────────────────────────────
function getLotesAdhesivoPendientes(params) {
  try {
    var rows = _sheetToObjects(_getSheetLotesAdhesivo());
    var pendientes = rows.filter(function(r) {
      var s = String(r.status || '').toUpperCase();
      return s === 'IMPRIMIENDO' || s === 'CALIBRANDO' || s.indexOf('PAUSADO') === 0;
    });
    if (params.usuario) {
      pendientes = pendientes.filter(function(r) { return String(r.usuario) === String(params.usuario); });
    }
    pendientes.sort(function(a, b) {
      return String(b.fechaUltimoUpdate || '').localeCompare(String(a.fechaUltimoUpdate || ''));
    });
    if (params.limit) pendientes = pendientes.slice(0, parseInt(params.limit));
    return { ok: true, data: pendientes };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// [v2.13.129] Historial de lotes por tipoEtiqueta — para botones "Cola WH/ME/Envasados".
// Devuelve últimos N lotes ordenados por fechaCreacion desc, separados en
// pendientes (en curso/pausados) e historial (completados/cancelados).
function getLotesAdhesivoHistorial(params) {
  try {
    params = params || {};
    var tipoFiltro = String(params.tipoEtiqueta || '').toUpperCase();
    var limit = parseInt(params.limit) || 50;
    var rows = _sheetToObjects(_getSheetLotesAdhesivo());

    // [v2.13.136 FIX] Inferir tipoEtiqueta desde la descripción para lotes
    // viejos que tienen ese campo vacío (antes del fix _filaLoteAdhesivo).
    var _inferirTipo = function(r) {
      var t = String(r.tipoEtiqueta || '').toUpperCase();
      if (t) return t;
      var d = String(r.descripcion || '');
      if (d.indexOf('ME: ') === 0 || d.indexOf('ME:') === 0) return 'MEMBRETE_ME';
      if (d.indexOf('WH: ') === 0 || d.indexOf('WH:') === 0) return 'MEMBRETE_WH';
      return 'ADHESIVO_ENVASADO';
    };
    // Filtrar por tipoEtiqueta si viene
    if (tipoFiltro) {
      rows = rows.filter(function(r) { return _inferirTipo(r) === tipoFiltro; });
    }
    // Ordenar por fecha desc (más recientes primero)
    rows.sort(function(a, b) {
      return String(b.fechaCreacion || '').localeCompare(String(a.fechaCreacion || ''));
    });

    // Separar pendientes vs historial
    var pendientes = [], completados = [];
    rows.forEach(function(r) {
      var s = String(r.status || '').toUpperCase();
      var item = {
        idLote:            String(r.idLote || ''),
        fechaCreacion:     String(r.fechaCreacion || ''),
        fechaUltimoUpdate: String(r.fechaUltimoUpdate || ''),
        usuario:           String(r.usuario || ''),
        origen:            String(r.origen || ''),
        codigoBarra:       String(r.codigoBarra || ''),
        descripcion:       String(r.descripcion || ''),
        vto:               String(r.vto || ''),
        totalEtq:          parseInt(r.totalEtq) || 0,
        completadas:       parseInt(r.completadas) || 0,
        status:            s,
        ultimoError:       String(r.ultimoError || ''),
        tipoEtiqueta:      _inferirTipo(r)  // [v2.13.136] usa el helper que infiere de descripción
      };
      if (s === 'COMPLETADO' || s === 'CANCELADO') completados.push(item);
      else pendientes.push(item);
    });

    // Aplicar limit al historial (pendientes siempre se muestran todos)
    completados = completados.slice(0, limit);

    return { ok: true, data: {
      pendientes: pendientes,
      historial:  completados,
      totalPendientes: pendientes.length,
      totalHistorial:  completados.length
    }};
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
