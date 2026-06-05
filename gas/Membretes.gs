// ============================================================
// warehouseMos — Membretes.gs   [v2.13.118]
// ============================================================
//
// MEMBRETE = adhesivo informativo NO de envasado, sino para pegar
// en góndolas (tienda) o andamios (almacén).
//
// Dos tipos:
//   MEMBRETE_ME: para góndola tienda. Precio MEGA prominente. Sin Vto.
//                Logo "TIENDA". Soles peruanos.
//                Si producto tiene equivalentes → muestra skuBase con icono ⬢
//                Si no → muestra codigoBarra con icono ▌
//
//   MEMBRETE_WH: para andamio almacén. Nombre del producto MEGA. Sin precio.
//                Logo "ALMACEN".
//                Si producto tiene N códigos → genera N+1 membretes:
//                  1 cabecera con skuBase (⬢)
//                  N individuales con cada codigoBarra (▌)
//
// REUTILIZA del sistema de adhesivos:
//   - Sistema de lotes con sub-jobs (sheet LOTES_ADHESIVO con tipoEtiqueta)
//   - Calibración inteligente OFFSET acumulativo (_calcularOffsetEfectivoParaPrint)
//   - Detección OUT_OF_PAPER + GAPDETECT condicional
//   - Modal de progreso unificado en frontend
//   - Helpers TSPL: _strToBytesEtq, _hexToBytesEtq, _wrapTokensEtq, etc

// ────────────────────────────────────────────────────────────────────
// LOGOS de membrete — se generan por gen.py. Por ahora, fallback al
// logo Tony's existente si los nuevos no están cargados todavía.
// Cuando F6 termine, estas constantes contendrán los hex reales.
// ────────────────────────────────────────────────────────────────────
var LOGO_TIENDA_ME_HEX  = 
  'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
  'FFFFFFFFFFFFFFFFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC1FFFFC000' +
  '1C07800381E03C001FF800FFFFFFFFFFF000000007C0001C07800380E03C0007F800FFFFFFFF' +
  'FFF000000007C0001C07800380E03C0003F000FFFFFFFFFFF000000007C0001C07800380E03C' +
  '0001F0007FFFFFFFFFF000000007C0001C07800380603C0001F0007FFFFFFFFFF000000007C0' +
  '001C07800380603C0601F0007FFFFFFFFFF000000007FE01FC0780FF80603C0700F0007FFFFF' +
  'FFFFF000000007FE01FC0780FF80203C0700F0007FFFFFFFFFFC0000001FFE01FC0780FF8020' +
  '3C0700F0207FFFFFFFFFFC0000001FFE01FC0780FF80203C0700E0207FFFFFFFFFFC0000001F' +
  'FE01FC0780FF80203C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFF' +
  'FFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780' +
  '003C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFFFFFFFFFC7FE3FF' +
  '1FFE01FC07800780003C0700E0203FFFFFFFFFFC0000001FFE01FC07800780003C0700E0301F' +
  'FFFFFFFFFC0000001FFE01FC0780FF80003C0700C0701FFFFFFFFFFC0000001FFE01FC0780FF' +
  '81003C0700C0701FFFFFFFFFFC7FE3FF1FFE01FC0780FF81003C0700C0701FFFFFFFFFFC7FE3' +
  'FF1FFE01FC0780FF81003C0700C0001FFFFFFFFFFC7FE3FF1FFE01FC0780FF81803C0700C000' +
  '1FFFFFFFFFFC7FE3FF1FFE01FC0780FF81803C0700C0001FFFFFFFFFFC7FE3FF1FFE01FC0780' +
  'FF81803C0700C0000FFFFFFFFFFC0000001FFE01FC07800381C03C0600C0000FFFFFFFFFFC00' +
  '7F001FFE01FC07800381C03C000180700FFFFFFFFFFC007F001FFE01FC07800381C03C000180' +
  '700FFFFFFFFFFC007F001FFE01FC07800381C03C000180700FFFFFFFFFFC007F001FFE01FC07' +
  '800381E03C000380700FFFFFFFFFFC007F001FFE01FC07800381E03C000F80700FFFFFFFFFFC' +
  '007F001FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC007F001FFFFFFFFFFFFFFFFFFFFFFF' +
  'FFFFFFFFFFFFFFFC007F001FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';  // generado por gen.py
var LOGO_ALMACEN_WH_HEX = 
  'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
  'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC00' +
  '7E03F801C007C007FF01FF000703C07FFFE0020007FC007E03F801C007C007FC007F000701C0' +
  '7FFFA0020007F8007E03F801C0078007F8001F000701C07FFF20020007F8003E03F801C00780' +
  '03F0000F000701C07FFE20020007F8003E03F801C0078003E0000F000700C07FFC20020007F8' +
  '003E03F800C0078003E0180F000700C07FF820020007F8003E03F800C0078003E0380701FF00' +
  'C07FF0203E0007F8003E03F800C0078003C0380701FF00407FF0203E0007F8103E03F8008007' +
  '8103C0380701FF00407FF0203E0007F0103E03F80080070103C0380701FF00407FF0203E0007' +
  'F0101E03F80080070101C0380701FF00407FF0203E0007F0101E03F80080070101C03807000F' +
  '00007FF0203E0007F0101E03F80000070101C03807000F00007FF0203E0007F0101E03F81004' +
  '070101C03FFF000F00007FF0203E0007F0101E03F81004070101C03FFF000F00007FF0203E00' +
  '07F0101E03F81004070101C03FFF000F00007FF0203E0007F0180E03F81004070180C03FFF00' +
  '0F00007FFFFFFFFFFFE0380E03F81004060380C0380701FF00007FFFFFFFFFFFE0380E03F818' +
  '04060380C0380701FF02007FFFFFFFFFFFE0380E03F81804060380C0380701FF02007FFFFFFF' +
  'FFFFE0000E03F8180C060000C0380701FF02007FFFFFFFFFFFE0000E03F8180C060000C03807' +
  '01FF03007FF0203E0007E0000E03F8180C060000C0380701FF03007FF0203E0007E0000603F8' +
  '180C06000060380F01FF03007FF0203E0007E000060018180C06000060180F000703807FF020' +
  '3E0007C0380600181C0C04038060000F000703807FF0203E0007C0380600181C0C0403807000' +
  '1F000703807FF0203E0007C0380600181C1C04038070003F000703807FF0203E0007C0380600' +
  '181C1C0403807C007F000703C07FF0203E0007C0380600181C1C0403807F01FF000703C07FF0' +
  '203E0007FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0203E0007FFFFFFFFFFFFFFFFFFFFFF' +
  'FFFFFFFFFFFFFFF0203E0007FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';  // generado por gen.py

function _getLogoHexParaTipo(tipo) {
  if (tipo === 'MEMBRETE_ME' && LOGO_TIENDA_ME_HEX)  return LOGO_TIENDA_ME_HEX;
  if (tipo === 'MEMBRETE_WH' && LOGO_ALMACEN_WH_HEX) return LOGO_ALMACEN_WH_HEX;
  // Fallback al logo Tony's (envasado)
  return LOGO_TSPL_HEX;
}

// ────────────────────────────────────────────────────────────────────
// _buildTSPLMembreteMe — adhesivo para góndola tienda
//
// Layout 50×25mm = 400×200 dots:
//   Y=2-38     logo TIENDA (184×36)
//   Y=42-66    descripción (Font 3 → 24 dots alto)
//   Y=70-114   barcode con corner marks + código texto
//   Y=120-184  PRECIO Font 5 MEGA (S/ XX.XX, 48 dots alto + caja)
//
// params del producto:
//   { codigoBarra, descripcion, precio, skuBase?, esSkuBase? }
//   esSkuBase=true → muestra skuBase con icono ⬢, no codigoBarra
// ────────────────────────────────────────────────────────────────────
// [v2026-06-05] Acepta offsetOverride opcional para batch de N etiquetas
// en un solo job (cada una con su drift compensado incremental).
function _buildTSPLMembreteMe(producto, allEnvasables, offsetOverride) {
  // [v2.13.142] REDISEÑO completo según feedback:
  // a) Barcode tamaño ESTÁNDAR (44 dots height, igual que adhesivo envasado).
  // b) Defensa codigoBarra undefined → fallback skuBase → idProducto → 'SIN-CODIGO'
  //    (antes salía literalmente "undefined" en el adhesivo).
  // c) Precio en esquina sup. derecha en NEGRITA (no mega central).
  //    Hace lugar para descripción 2 líneas sin sobre-escribir el barcode.
  // d) Mismo layout que adhesivo envasado: logo + desc + frame con corner marks + barcode.
  var descNorm = _normalizeEtq(producto.descripcion || '');
  var tokens = descNorm.split(/\s+/);
  var allTok = (allEnvasables || []).map(function(p) { return p.tokens; });
  var highlights = _detectHighlightsEtq(tokens, allTok);
  var lines = _wrapTokensEtq(tokens, highlights);

  var props = PropertiesService.getScriptProperties();
  var gapMm    = parseFloat(props.getProperty('ADHESIVO_GAP_MM'))   || 2;
  var density  = parseInt  (props.getProperty('ADHESIVO_DENSITY'))  || 8;
  var speed    = parseInt  (props.getProperty('ADHESIVO_SPEED'))    || 4;
  // [v2026-06-05] Si vino offsetOverride (batch), usarlo. Sino calcular.
  var offsetY  = (typeof offsetOverride === 'number')
    ? offsetOverride
    : _calcularOffsetEfectivoParaPrint();

  var logoHex = _getLogoHexParaTipo('MEMBRETE_ME');

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
  bytes = bytes.concat(_hexToBytesEtq(logoHex));
  bytes = bytes.concat(_strToBytesEtq('\r\n'));

  // [v2.13.142] PRECIO en esquina sup. derecha — negrita compacta.
  // Font 4 = 16×24 dots. Ej: "S/ 8.50" = 7 chars × 16 = 112 dots wide.
  var precioStr = 'S/ ' + (parseFloat(producto.precio) || 0).toFixed(2);
  var precioFontW = 16;
  var precioWidth = precioStr.length * precioFontW;
  var precioX = 400 - precioWidth - 4;
  var precioY = 4 + offsetY;
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + precioX + ',' + precioY + ',"4",0,1,1,"' + precioStr + '"\r\n'));
  // Línea decorativa debajo del precio
  bytes = bytes.concat(_strToBytesEtq('BAR ' + precioX + ',' + (precioY + 28) + ',' + precioWidth + ',2\r\n'));

  // [v2.13.142] Descripción centrada en área Y=46-118 (mismo layout que envasado)
  var DESC_AREA_Y0 = 46, DESC_AREA_H = 72;
  var LINE_H = 38, SPACE = 8;
  var startY;
  if (lines.length === 1) {
    var lineHasHl = lines[0].some(function(t) { return t.hl; });
    var lineHeight = lineHasHl ? 32 : 24;
    var baselineOffset = lineHasHl ? 0 : 4;
    startY = DESC_AREA_Y0 + Math.floor((DESC_AREA_H - lineHeight) / 2) - baselineOffset + offsetY;
  } else {
    startY = DESC_AREA_Y0 + offsetY;
  }
  for (var li = 0; li < Math.min(lines.length, 2); li++) {
    var line = lines[li];
    var totalW = 0;
    for (var ti = 0; ti < line.length; ti++) totalW += line[ti].w + (ti > 0 ? SPACE : 0);
    var x = Math.max(5, Math.round((400 - totalW) / 2));
    var y = startY + li * LINE_H;
    for (var tj = 0; tj < line.length; tj++) {
      var o = line[tj];
      var font = o.hl ? '4' : '3';
      var yAdj = o.hl ? y : y + 4;
      var safe = String(o.tok).replace(/"/g, "'");
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + x + ',' + yAdj + ',"' + font + '",0,1,1,"' + safe + '"\r\n'));
      x += o.w + SPACE;
    }
  }

  // [v2026-06-05 FIX scanner pistola] Cálculo adaptativo del barcode →
  // helper _calcBarcodeAdaptativo (Envasados.gs). ANTES: narrowBc=1 si código
  // largo → ilegible para scanner pistola CCD. AHORA: nunca <=2, altura 48.
  var codigo = String(
    (producto.esSkuBase ? producto.skuBase : producto.codigoBarra)
    || producto.codigoBarra
    || producto.skuBase
    || producto.idProducto
    || ''
  ).replace(/"/g, '');
  if (!codigo) codigo = 'SIN-CODIGO';
  var _bcCalc = _calcBarcodeAdaptativo(codigo);
  var narrowBc = _bcCalc.narrowBc;
  var barcodeWidth = _bcCalc.barcodeWidth;
  var barcodeHeight = _bcCalc.barcodeHeight;
  var barcodeX = _bcCalc.barcodeX;
  var barcodeY = 124 + offsetY;
  // Frame con corner marks (igual que envasado y WH)
  var frameX1 = 10, frameX2 = 389;
  var frameY1 = 118 + offsetY, frameY2 = 196 + offsetY, cmL = 12;
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + codigo + '"\r\n'));

  // [v2.13.142] Texto del código LIMPIO (sin ícono ">", "*", "(SKU)" — son sucios)
  var codigoFontW = 8;
  var codigoWidth = codigo.length * codigoFontW;
  var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));
  var codigoY = barcodeY + barcodeHeight + 8;
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + codigo + '"\r\n'));

  bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
  return bytes;
}

// ────────────────────────────────────────────────────────────────────
// _buildTSPLMembreteWh — adhesivo para andamio almacén
//
// Layout:
//   Y=2-38     logo ALMACEN (chico)
//   Y=42-90    NOMBRE producto Font 5 MEGA (24-48 dots tall)
//   Y=95       línea horizontal separadora
//   Y=100-138  barcode + corner marks
//   Y=142-160  texto del código con icono
//   Y=170-185  "SKU: XXX" info extra abajo
//
// params:
//   producto: { descripcion, codigo, skuBase }
//   esCabecera: bool — si true, muestra skuBase como código principal
//                       y agrega indicador "📋 CABECERA"
//   indice: número del adhesivo en la serie (1, 2, 3...)
//   total: total de adhesivos en la serie (para "1/4", "2/4", etc)
// ────────────────────────────────────────────────────────────────────
// [v2026-06-05] Acepta offsetOverride opcional para batch de N etiquetas
// en un solo job (cada una con su drift compensado incremental).
function _buildTSPLMembreteWh(producto, esCabecera, indice, total, allEnvasables, offsetOverride) {
  // [v2.13.142] REDISEÑO completo según feedback:
  // a) Barcode tamaño ESTÁNDAR (44 dots height, igual que adhesivo envasado).
  // b) NO mostrar SKU si solo hay un código (info inútil). Si es parte de
  //    grupo multi-código, mostrar "1/5" o "CABECERA" en esquina sup. der.
  // c) Sin íconos "> " "*"  — se ven sucios. Solo el código limpio.
  var descNorm = _normalizeEtq(producto.descripcion || '');
  var tokens = descNorm.split(/\s+/);
  var allTok = (allEnvasables || []).map(function(p) { return p.tokens; });
  var highlights = _detectHighlightsEtq(tokens, allTok);
  var lines = _wrapTokensEtq(tokens, highlights);

  var props = PropertiesService.getScriptProperties();
  var gapMm    = parseFloat(props.getProperty('ADHESIVO_GAP_MM'))   || 2;
  var density  = parseInt  (props.getProperty('ADHESIVO_DENSITY'))  || 8;
  var speed    = parseInt  (props.getProperty('ADHESIVO_SPEED'))    || 4;
  // [v2026-06-05] Si vino offsetOverride (batch), usarlo. Sino calcular.
  var offsetY  = (typeof offsetOverride === 'number')
    ? offsetOverride
    : _calcularOffsetEfectivoParaPrint();

  var logoHex = _getLogoHexParaTipo('MEMBRETE_WH');

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
  bytes = bytes.concat(_hexToBytesEtq(logoHex));
  bytes = bytes.concat(_strToBytesEtq('\r\n'));

  // [v2.13.142] Indicador esquina sup. derecha SOLO si es serie multi-código.
  // Para producto con 1 solo código: nada (ocultar info inútil).
  if (total > 1) {
    var tagTexto = esCabecera ? 'CAB' : (indice + '/' + total);
    var tagX = 400 - tagTexto.length * 8 - 8;
    bytes = bytes.concat(_strToBytesEtq('TEXT ' + tagX + ',' + (4 + offsetY) + ',"2",0,1,1,"' + tagTexto + '"\r\n'));
  }

  // [v2.13.142] Descripción en área Y=46 a Y=118 (centrada igual que envasado)
  var DESC_AREA_Y0 = 46, DESC_AREA_H = 72;
  var LINE_H = 38, SPACE = 8;
  var startY;
  if (lines.length === 1) {
    var lineHasHl = lines[0].some(function(t) { return t.hl; });
    var lineHeight = lineHasHl ? 32 : 24;
    var baselineOffset = lineHasHl ? 0 : 4;
    startY = DESC_AREA_Y0 + Math.floor((DESC_AREA_H - lineHeight) / 2) - baselineOffset + offsetY;
  } else {
    startY = DESC_AREA_Y0 + offsetY;
  }
  for (var li = 0; li < Math.min(lines.length, 2); li++) {
    var line = lines[li];
    var totalW = 0;
    for (var ti = 0; ti < line.length; ti++) totalW += line[ti].w + (ti > 0 ? SPACE : 0);
    var x = Math.max(5, Math.round((400 - totalW) / 2));
    var y = startY + li * LINE_H;
    for (var tj = 0; tj < line.length; tj++) {
      var o = line[tj];
      var font = o.hl ? '4' : '3';
      var yAdj = o.hl ? y : y + 4;
      var safe = String(o.tok).replace(/"/g, "'");
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + x + ',' + yAdj + ',"' + font + '",0,1,1,"' + safe + '"\r\n'));
      x += o.w + SPACE;
    }
  }

  // [v2026-06-05 FIX scanner pistola] Cálculo adaptativo del barcode →
  // helper _calcBarcodeAdaptativo (Envasados.gs). ANTES: narrowBc=1 si código
  // largo → ilegible para scanner pistola CCD. AHORA: nunca <=2, altura 48.
  var codigo = String(producto.codigo || producto.codigoBarra || producto.skuBase || producto.idProducto || '').replace(/"/g, '');
  if (!codigo) codigo = 'SIN-CODIGO';
  var _bcCalc = _calcBarcodeAdaptativo(codigo);
  var narrowBc = _bcCalc.narrowBc;
  var barcodeWidth = _bcCalc.barcodeWidth;
  var barcodeHeight = _bcCalc.barcodeHeight;
  var barcodeX = _bcCalc.barcodeX;
  var barcodeY = 124 + offsetY;
  // Frame con corner marks (igual que envasado)
  var frameX1 = 10, frameX2 = 389;
  var frameY1 = 118 + offsetY, frameY2 = 196 + offsetY, cmL = 12;
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
  bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + codigo + '"\r\n'));

  // [v2.13.142] Texto del código SIN íconos sucios — solo el código limpio
  var codigoFontW = 8;
  var codigoWidth = codigo.length * codigoFontW;
  var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));
  var codigoY = barcodeY + barcodeHeight + 8;
  bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + codigo + '"\r\n'));

  bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
  return bytes;
}

// ────────────────────────────────────────────────────────────────────
// crearLoteMembrete — endpoint del frontend para crear el lote.
//
// POST {
//   tipo: 'MEMBRETE_ME' | 'MEMBRETE_WH',
//   items: [{ codigoBarra, descripcion, precio?, skuBase?, esSkuBase? }, ...],
//   usuario, origen, idempotencyKey
// }
//
// Para ME: items = lista de productos a imprimir (cada uno 1 adhesivo).
// Para WH: items = lista de productos. Cada producto puede generar N+1
//          adhesivos si tiene equivalentes. Los items expandidos se
//          guardan en itemsJson del lote (con sus codigos individuales).
//
// total del lote = sum de adhesivos (expandidos para WH).
// ────────────────────────────────────────────────────────────────────
function crearLoteMembrete(params) {
  try {
    var tipo = String(params.tipo || '').toUpperCase();
    if (tipo !== 'MEMBRETE_ME' && tipo !== 'MEMBRETE_WH') {
      return { ok: false, error: 'tipo debe ser MEMBRETE_ME o MEMBRETE_WH' };
    }
    var items = params.items || [];
    if (!Array.isArray(items) || items.length === 0) {
      return { ok: false, error: 'items debe ser array no vacío' };
    }
    var usuario = String(params.usuario || '').trim();
    var origen  = String(params.origen  || 'MOS').trim().toUpperCase();
    var idempotencyKey = String(params.idempotencyKey || '').trim();

    // Dedupe por idempotencyKey (igual que crearLoteAdhesivo)
    if (idempotencyKey) {
      var sheetCheck = _getSheetLotesAdhesivo();
      var allRows = _sheetToObjects(sheetCheck);
      var sufijo = '_' + idempotencyKey;
      var found = allRows.find(function(r) {
        var id = String(r.idLote || '');
        return id.length >= sufijo.length && id.substring(id.length - sufijo.length) === sufijo;
      });
      if (found) {
        return { ok: true, data: {
          idLote:       String(found.idLote),
          total:        parseInt(found.totalEtq) || 0,
          completadas:  parseInt(found.completadas) || 0,
          status:       String(found.status || 'CREADO'),
          tipoEtiqueta: String(found.tipoEtiqueta || tipo),
          deduped:      true
        }};
      }
    }

    // Resolver printerId
    var printerId;
    try { printerId = String(getPrinterNodeId('ADHESIVO', 'ALMACEN')); }
    catch (e) { return { ok: false, error: 'Sin impresora ADHESIVO/ALMACEN: ' + e.message }; }

    // Expandir items: para MEMBRETE_ME, 1 producto = 1 adhesivo.
    // Para MEMBRETE_WH, 1 producto con N codigos = 1 cabecera + N codigos
    // (N+1 adhesivos). Si producto tiene 1 solo código (sin equivalentes),
    // genera solo 1 adhesivo (sin cabecera).
    var expandidos = [];
    items.forEach(function(item) {
      if (tipo === 'MEMBRETE_ME') {
        expandidos.push({
          codigo:      String(item.codigoBarra || ''),
          descripcion: String(item.descripcion || ''),
          precio:      parseFloat(item.precio) || 0,
          skuBase:     String(item.skuBase || ''),
          esSkuBase:   !!item.esSkuBase,
          esCabecera:  false,
          tipoItem:    'ME_UNIT'
        });
      } else {
        // MEMBRETE_WH
        // [v2.13.146 FIX] Si item.codigos = [] (array vacío) el Array.isArray
        // pasaba TRUE pero codigos quedaba vacío → forEach 0 veces → expandidos
        // = [] → lote con totalEtq=0 que el trigger marca COMPLETADO sin imprimir.
        // Bug reportado tras 2026-06-04 15:30 cuando productos WH sin equivalentes
        // empezaron a llegar con `codigos:[]` desde el frontend.
        // Defensa: si array vacío, caer al fallback codigoBarra.
        var codigos = (Array.isArray(item.codigos) && item.codigos.length > 0)
          ? item.codigos
          : (item.codigoBarra ? [item.codigoBarra] : []);
        if (codigos.length === 0) {
          return;  // skipear este item — nada que imprimir
        }
        if (codigos.length > 1 && item.skuBase) {
          // Cabecera
          expandidos.push({
            codigo:      String(item.skuBase),
            descripcion: String(item.descripcion || ''),
            skuBase:     String(item.skuBase || ''),
            esCabecera:  true,
            tipoItem:    'WH_CAB'
          });
        }
        codigos.forEach(function(c) {
          expandidos.push({
            codigo:      String(c),
            descripcion: String(item.descripcion || ''),
            skuBase:     String(item.skuBase || ''),
            esCabecera:  false,
            tipoItem:    'WH_COD'
          });
        });
      }
    });

    var total = expandidos.length;
    // [v2.13.146] Validación HARD: si expandidos quedó vacío (todos los items
    // sin codigoBarra ni codigos válidos) → rechazar. Antes se creaba lote
    // con totalEtq=0 que el trigger marcaba COMPLETADO sin imprimir nada.
    if (total === 0) {
      return { ok: false, error: 'Sin items válidos para imprimir — cada item debe traer codigoBarra o codigos no vacío' };
    }
    var subJobSize = 1;  // membretes siempre PRINT 1,1
    var sheet = _getSheetLotesAdhesivo();
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
      codigoBarra:         '',
      descripcion:         tipo === 'MEMBRETE_ME' ? ('ME: ' + items.length + ' productos') : ('WH: ' + items.length + ' productos · ' + total + ' adhesivos'),
      vto:                 '',
      totalEtq:            total,
      completadas:         0,
      subJobSize:          subJobSize,
      status:              'ENCOLADO',
      ultimoError:         '',
      ultimoPrintNodeJobId: '',
      printerId:           printerId,
      tipoEtiqueta:        tipo,
      itemsJson:           JSON.stringify(expandidos)
    }, sheet);  // [v2.13.136] pasar sheet para respetar orden REAL de columnas
    sheet.appendRow(fila);
    // [v2.13.130 FIX] Auto-instalar trigger si falta. Sin trigger, los lotes
    // ENCOLADO nunca se procesan → membretes se quedan en cola para siempre.
    var triggerCheck = { ok: true };
    try { triggerCheck = _asegurarTriggerLotes(); } catch(_){}
    return { ok: true, data: {
      idLote:       idLote,
      total:        total,
      completadas:  0,
      subJobSize:   subJobSize,
      status:       'ENCOLADO',
      tipoEtiqueta: tipo,
      itemsExpandidos: expandidos.length,
      triggerAutoInstalado: !!triggerCheck.instaladoAhora
    }};
  } catch (e) {
    return { ok: false, error: e.message, stack: e.stack };
  }
}

// ────────────────────────────────────────────────────────────────────
// imprimirSubLoteMembrete — procesa SIGUIENTE adhesivo del lote.
// Llamado por procesarLotesPendientes() del trigger time-based, NO
// directamente desde frontend.
//
// Lee itemsJson del lote, toma item[completadas], lo imprime.
// ────────────────────────────────────────────────────────────────────
function imprimirSubLoteMembrete(idLote) {
  // [v2026-06-05 SENIOR AUDIT FIX] LockService para evitar doble impresión.
  // Sin esto: si el frontend dispara fire-and-forget Y el trigger backend
  // se ejecuta a los pocos segundos, AMBAS calls procesarían el mismo lote
  // → mismo TSPL mandado a PrintNode dos veces → adhesivos duplicados.
  // Lock de 15s suficiente para que la 1era termine antes que la 2da pase.
  var _lock = LockService.getScriptLock();
  try { _lock.waitLock(15000); } catch(eLock) {
    return { ok: true, data: { skipped: 'lock_no_obtenido_otro_proceso_en_curso', idLote: idLote } };
  }
  try {
    var sheet = _getSheetLotesAdhesivo();
    var rowIdx = _findLoteRow(sheet, idLote);
    if (rowIdx < 0) return { ok: false, error: 'Lote no encontrado: ' + idLote };
    var lote = _readLote(sheet, rowIdx);
    if (lote.status === 'CANCELADO')  return { ok: false, error: 'Lote cancelado' };
    if (lote.status === 'COMPLETADO') return { ok: true, data: { idLote: idLote, completadas: lote.completadas, status: 'COMPLETADO' } };
    if (lote.completadas >= lote.totalEtq) {
      _patchLote(sheet, rowIdx, { status: 'COMPLETADO' });
      return { ok: true, data: { idLote: idLote, completadas: lote.completadas, status: 'COMPLETADO' } };
    }

    // [v2026-06-05 SENIOR AUDIT FIX] Skipear si otro proceso lo está manejando
    // ahora mismo (status=IMPRIMIENDO con lastUpdate <90s). Mismo patrón que
    // imprimirSubLoteAdhesivo en Envasados.gs línea 2705.
    if (lote.status === 'IMPRIMIENDO' || lote.status === 'CALIBRANDO') {
      var lastUpdMs = 0;
      try { lastUpdMs = new Date(lote.fechaUltimoUpdate || '').getTime(); } catch(_){}
      var ageS = (Date.now() - lastUpdMs) / 1000;
      if (lastUpdMs && ageS < 90) {
        return { ok: true, data: {
          idLote: idLote, completadas: lote.completadas, total: lote.totalEtq,
          status: lote.status,
          skipped: 'sub_job_concurrente_en_curso_' + Math.round(ageS) + 's'
        }};
      }
      // Si han pasado >90s, asumimos que el proceso anterior se colgó y
      // este puede retomarlo. La sheet sigue siendo source of truth.
    }

    // Decodificar items
    var items;
    try { items = JSON.parse(lote.itemsJson || '[]'); }
    catch(e) { return { ok: false, error: 'itemsJson inválido: ' + e.message }; }
    if (!items || items.length <= lote.completadas) {
      return { ok: false, error: 'Sin item para completar #' + lote.completadas };
    }

    // [v2026-06-05 BATCH] Procesar TODOS los items restantes en UN solo TSPL
    // + UN solo job PrintNode. Antes: 1 item por job → 4 adhesivos = 4 calls
    // PrintNode con 20s polling cada uno (60-80s total). Ahora: 1 call con
    // todos los TSPL concatenados (5-10s total).
    //
    // Drift incremental: cada etiqueta DENTRO del lote usa su propio offsetY
    // calculado via _calcularOffsetParaIndiceDentroDeLote(i). El contador
    // de prints se incrementa por N al final (no 1×N veces).
    var tipo = String(lote.tipoEtiqueta || 'MEMBRETE_ME').toUpperCase();
    var indiceInicio = lote.completadas;
    var qtyEnBatch = items.length - indiceInicio;

    // Marcar IMPRIMIENDO
    _patchLote(sheet, rowIdx, { status: 'IMPRIMIENDO' });

    // Construir TSPL para CADA item con drift incremental
    var allEnv = [];
    try { allEnv = _getAllEnvasablesTokens(); } catch(_) {}
    var bytesTotal = [];
    for (var i = indiceInicio; i < items.length; i++) {
      var itemActual = items[i];
      // Offset incremental: la primera del batch usa drift actual + 0,
      // la siguiente + 1*driftDots, etc.
      var offsetParaEste = _calcularOffsetParaIndiceDentroDeLote(i - indiceInicio);
      var bytesItem;
      if (tipo === 'MEMBRETE_ME') {
        bytesItem = _buildTSPLMembreteMe(itemActual, allEnv, offsetParaEste);
      } else if (tipo === 'MEMBRETE_WH') {
        bytesItem = _buildTSPLMembreteWh(itemActual, !!itemActual.esCabecera, i + 1, items.length, allEnv, offsetParaEste);
      } else {
        return { ok: false, error: 'Tipo no soportado: ' + tipo };
      }
      bytesTotal = bytesTotal.concat(bytesItem);
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty('PRINTNODE_API_KEY') || '';
    if (!apiKey) return { ok: false, error: 'PRINTNODE_API_KEY no configurada' };
    var b64 = Utilities.base64Encode(bytesTotal);
    var auth = 'Basic ' + Utilities.base64Encode(apiKey + ':');

    var resp = UrlFetchApp.fetch('https://api.printnode.com/printjobs', {
      method: 'post',
      headers: { 'Authorization': auth },
      contentType: 'application/json',
      payload: JSON.stringify({
        printerId: parseInt(lote.printerId),
        title: tipo + ' lote ' + qtyEnBatch + ' adh · ' + idLote,
        contentType: 'raw_base64',
        content: b64,
        source: 'warehouseMos-membrete-batch-' + idLote
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
      return { ok: false, error: 'PrintNode HTTP ' + code, status: 'PAUSADO_ERROR' };
    }
    var printNodeJobId = JSON.parse(resp.getContentText());

    // Polling de UN solo job (en lugar de N). Timeout proporcional al batch.
    var pollResult = _esperarFinJobPrintNode(printNodeJobId, auth, 30000 + qtyEnBatch * 5000, 1500);
    if (pollResult.estado === 'error') {
      var msg = String(pollResult.mensaje || '').toLowerCase();
      var esOutOfPaper = msg.indexOf('paper') >= 0 || msg.indexOf('media') >= 0 || msg.indexOf('label') >= 0;
      _patchLote(sheet, rowIdx, {
        status:               esOutOfPaper ? 'PAUSADO_OUT_PAPER' : 'PAUSADO_ERROR',
        ultimoError:          pollResult.mensaje || 'Error en job ' + printNodeJobId,
        ultimoPrintNodeJobId: String(printNodeJobId)
      });
      return { ok: false, error: pollResult.mensaje, status: esOutOfPaper ? 'PAUSADO_OUT_PAPER' : 'PAUSADO_ERROR' };
    }

    // OK — incrementar completadas POR N + contador de drift POR N
    var nuevasCompletadas = lote.completadas + qtyEnBatch;
    _patchLote(sheet, rowIdx, {
      completadas:          nuevasCompletadas,
      status:               'COMPLETADO',
      ultimoPrintNodeJobId: String(printNodeJobId),
      ultimoError:          ''
    });
    try { _incrementarPrintsCount(qtyEnBatch); } catch(_) {}

    return { ok: true, data: {
      idLote:         idLote,
      completadas:    nuevasCompletadas,
      total:          items.length,
      status:         'COMPLETADO',
      printNodeJobId: printNodeJobId,
      pollEstado:     pollResult.estado,
      qtyEnBatch:     qtyEnBatch
    }};
  } catch (e) {
    return { ok: false, error: e.message, stack: e.stack };
  } finally {
    // [v2026-06-05 SENIOR AUDIT FIX] Liberar lock SIEMPRE
    try { _lock.releaseLock(); } catch(_){}
  }
}

// ────────────────────────────────────────────────────────────────────
// PREVIEWS para diagnóstico desde el editor GAS
// ────────────────────────────────────────────────────────────────────
function previsualizarMembreteMe(codigoBarra) {
  try {
    var prods = _sheetToObjects(getProductosSheet());
    var p = prods.find(function(x) {
      return String(x.codigoBarra) === codigoBarra || String(x.idProducto) === codigoBarra;
    });
    if (!p) return { ok: false, error: 'Producto no encontrado' };
    var producto = {
      codigoBarra: String(p.codigoBarra || p.idProducto),
      descripcion: String(p.descripcion || p.nombre || ''),
      precio:      parseFloat(p.precio) || 0,
      skuBase:     String(p.skuBase || ''),
      esSkuBase:   false
    };
    var allEnv = _getAllEnvasablesTokens();
    var bytes = _buildTSPLMembreteMe(producto, allEnv);
    var preview = '';
    for (var i = 0; i < Math.min(bytes.length, 800); i++) {
      var b = bytes[i];
      if (b >= 32 && b < 127) preview += String.fromCharCode(b);
      else if (b === 13) preview += '\\r';
      else if (b === 10) preview += '\\n\n';
      else preview += '·';
    }
    Logger.log('═════ MEMBRETE ME preview ' + codigoBarra + ' ═════');
    Logger.log(preview);
    return { ok: true, data: { totalBytes: bytes.length, tsplPreview: preview } };
  } catch(e) { return { ok: false, error: e.message }; }
}

function previsualizarMembreteWh(codigoBarra) {
  try {
    var prods = _sheetToObjects(getProductosSheet());
    var p = prods.find(function(x) {
      return String(x.codigoBarra) === codigoBarra || String(x.idProducto) === codigoBarra;
    });
    if (!p) return { ok: false, error: 'Producto no encontrado' };
    var producto = {
      codigo:      String(p.codigoBarra || p.idProducto),
      descripcion: String(p.descripcion || p.nombre || ''),
      skuBase:     String(p.skuBase || p.codigoBarra)
    };
    var allEnv = _getAllEnvasablesTokens();
    var bytes = _buildTSPLMembreteWh(producto, false, 1, 1, allEnv);
    Logger.log('═════ MEMBRETE WH preview ' + codigoBarra + ' ═════');
    Logger.log('totalBytes: ' + bytes.length);
    return { ok: true, data: { totalBytes: bytes.length } };
  } catch(e) { return { ok: false, error: e.message }; }
}
