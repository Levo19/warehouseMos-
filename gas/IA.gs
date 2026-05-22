// ============================================================
// warehouseMos — IA.gs
// Integraciones con Claude API (Anthropic) para features inteligentes
//
// SETUP:
//   1. PropertiesService → ANTHROPIC_API_KEY = 'sk-ant-...'
//   2. Modelo por defecto: Haiku 4.5 (rápido + barato).
// ============================================================

var IA_MODELO_DEFAULT = 'claude-haiku-4-5-20251001';
var IA_ENDPOINT       = 'https://api.anthropic.com/v1/messages';

// ── Wrapper genérico para llamadas a Claude ────────────────
// Devuelve {ok, text, raw} o {ok:false, error}.
function _llamarClaude(opts) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'KEY_NOT_SET', mensaje: 'ANTHROPIC_API_KEY no configurada en Script Properties' };

  var payload = {
    model:       opts.model || IA_MODELO_DEFAULT,
    max_tokens:  opts.max_tokens || 2048,
    system:      opts.system || '',
    messages:    opts.messages
  };

  try {
    var resp = UrlFetchApp.fetch(IA_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code !== 200) {
      return { ok: false, error: 'API_ERROR_' + code, mensaje: body.substring(0, 300) };
    }
    var data = JSON.parse(body);
    var text = (data.content && data.content[0] && data.content[0].text) || '';
    return { ok: true, text: text, raw: data };
  } catch(e) {
    return { ok: false, error: 'NETWORK', mensaje: e.message };
  }
}

// ============================================================
// analizarListaSombra(params.texto) → {ok, items: [{nombre, cantidad, codigoVisto?}]}
// ------------------------------------------------------------
// Recibe texto pegado de cualquier formato (Excel, WhatsApp, email)
// y devuelve la lista limpia para usar como "lista sombra" en
// Despacho Rápido. Cantidades a 1 decimal.
// ============================================================
function analizarListaSombra(params) {
  var texto = String(params && params.texto || '').trim();
  if (!texto) return { ok: false, error: 'TEXTO_VACIO' };
  // [v2.13.25] Sin chunking en backend — el frontend ya trozó la lista en
  // bloques de ~30 productos. Cada llamada es chica.
  if (texto.length > 50000) return { ok: false, error: 'TEXTO_MUY_LARGO', mensaje: 'Chunk demasiado grande — el frontend debería haberlo trozado más' };

  var system = [
    'Eres un asistente que limpia listas de productos de almacén.',
    'Recibes texto pegado de cualquier formato (Excel, WhatsApp, email, ticket impreso).',
    'Tu trabajo: extraer SOLO los productos reales con su cantidad pedida.',
    '',
    'IGNORA:',
    '- Cabeceras (Código, Descripción, Pedido, etc.)',
    '- Totales / subtotales',
    '- Líneas separadoras (---, ===, ...)',
    '- Comentarios o notas que no sean producto',
    '',
    'POR CADA PRODUCTO devuelve:',
    '- nombre: descripción del producto en MAYÚSCULAS, limpia, sin códigos pegados',
    '- cantidad: número decimal con 1 decimal (ej: 5.0, 18.0, 0.5)',
    '- codigoVisto: opcional — si el texto traía un código/sku al lado, ponlo (string), si no, omite el campo',
    '',
    'RESPONDE EXCLUSIVAMENTE con JSON válido en este formato (sin markdown, sin comentarios):',
    '{"items":[{"nombre":"...","cantidad":N.N,"codigoVisto":"..."}]}'
  ].join('\n');

  // [v2.13.24] max_tokens 8192 (Haiku 4.5 acepta hasta 8192 output)
  // para soportar listas grandes sin truncar el JSON de respuesta.
  var ia = _llamarClaude({
    max_tokens: 8192,
    system: system,
    messages: [{
      role: 'user',
      content: 'Limpia esta lista y devuelve solo JSON:\n\n' + texto
    }]
  });

  if (!ia.ok) return ia;

  // Parsear la respuesta — la IA puede a veces meter texto extra antes/después
  var jsonStr = ia.text.trim();
  // Extraer entre primer { y último }
  var first = jsonStr.indexOf('{');
  var last  = jsonStr.lastIndexOf('}');
  if (first < 0 || last < 0) return { ok: false, error: 'PARSE_FAIL', mensaje: ia.text.substring(0, 200) };
  jsonStr = jsonStr.substring(first, last + 1);

  var parsed;
  try { parsed = JSON.parse(jsonStr); }
  catch(e) { return { ok: false, error: 'PARSE_FAIL', mensaje: e.message + ' · raw: ' + ia.text.substring(0, 200) }; }

  var items = Array.isArray(parsed.items) ? parsed.items : [];
  // Normalizar
  var limpios = items.map(function(it) {
    return {
      nombre:      String(it.nombre || '').toUpperCase().trim(),
      cantidad:    Math.round((parseFloat(it.cantidad) || 0) * 10) / 10,
      codigoVisto: it.codigoVisto ? String(it.codigoVisto) : ''
    };
  }).filter(function(it) { return it.nombre && it.cantidad > 0; });

  return { ok: true, data: { items: limpios, total: limpios.length } };
}

// ============================================================
// testAnthropic() — ejecutar desde el editor GAS para verificar
// que la key esté bien configurada y funcione end-to-end.
//   ✅ → Logger.log dice "OK · modelo: ... · respondió: ..."
//   ❌ KEY_NOT_SET → falta setear ANTHROPIC_API_KEY en Script Properties
//   ❌ API_ERROR_401 → key inválida o expirada
//   ❌ API_ERROR_429 → sin saldo / rate limit
// ============================================================
function testAnthropic() {
  Logger.log('— Test Anthropic API —');
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) {
    Logger.log('❌ ANTHROPIC_API_KEY no configurada en Script Properties');
    return { ok: false, mensaje: 'Configura la key en ⚙ → Script Properties' };
  }
  Logger.log('🔑 Key encontrada (prefijo: ' + key.substring(0, 12) + '...)');
  Logger.log('📡 Llamando al modelo...');

  var res = _llamarClaude({
    max_tokens: 100,
    messages: [{
      role: 'user',
      content: 'Responde solo "PONG" en mayúsculas. Nada más.'
    }]
  });

  if (!res.ok) {
    Logger.log('❌ ERROR: ' + res.error + ' · ' + (res.mensaje || ''));
    if (res.error === 'API_ERROR_401') Logger.log('   → Key inválida. Verifica que sea sk-ant-...');
    if (res.error === 'API_ERROR_429') Logger.log('   → Sin saldo o rate limit. Revisa billing en console.anthropic.com');
    if (res.error === 'API_ERROR_400') Logger.log('   → Request malformado. Avisa al dev.');
    return res;
  }

  Logger.log('✅ OK · modelo respondió: "' + res.text.trim() + '"');
  Logger.log('   tokens usados: input=' + (res.raw.usage?.input_tokens || '?') +
             ' · output=' + (res.raw.usage?.output_tokens || '?'));
  Logger.log('   → Listo. Ya puedes usar "Subir Lista" en Despacho Rápido.');
  return { ok: true, text: res.text };
}

// ============================================================
// [v2.13.41] analizarFacturaProveedor(params) → extrae datos de boleta/factura
// del proveedor desde una foto subida a la guía. Llama a Claude con visión.
//
// Trigger automático: al subir/actualizar foto en una guía INGRESO_PROVEEDOR,
// MosExpress y MOS dependen del campo IGV_Recuperable en GUIAS para el
// "Centro Tributario". El proveedor a veces manda la factura una semana
// después de la mercadería → re-procesar si la foto cambia.
//
// params: { idGuia }  (la foto se lee de columna `foto` de hoja GUIAS)
// Devuelve: {ok, data: {tipoComprobante, ruc, serie, numero, fecha, total,
//                       igvRecuperable, items[], estado, confidence}}
//   estado: 'PROCESADO' (con IGV) | 'SIN_IGV' (ticket/boleta s/RUC)
//           | 'ILEGIBLE' (no se pudo leer) | 'NO_COMPROBANTE' (no es factura)
// ============================================================
function analizarFacturaProveedor(params) {
  var idGuia = String(params && params.idGuia || '').trim();
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  // 1. Leer la foto de la guía (columna `foto` de hoja GUIAS)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet('GUIAS');
  if (!sheet) return { ok: false, error: 'Hoja GUIAS no encontrada' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var idxId  = hdrs.indexOf('idGuia');
  var idxFoto = hdrs.indexOf('foto');
  if (idxId < 0 || idxFoto < 0) return { ok: false, error: 'Columnas idGuia/foto no encontradas' };

  var filaIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idGuia) { filaIdx = i; break; }
  }
  if (filaIdx < 0) return { ok: false, error: 'Guía ' + idGuia + ' no encontrada' };
  var urlFoto = String(data[filaIdx][idxFoto] || '').trim();
  if (!urlFoto) return { ok: false, error: 'La guía no tiene foto' };

  // 2. Descargar la imagen y convertir a base64
  var imgBase64, imgMime;
  try {
    // URLs típicas: https://drive.google.com/thumbnail?id=XXX&sz=w1600
    // Necesitamos descargar la imagen original
    var match = urlFoto.match(/[?&]id=([^&]+)/);
    if (match) {
      var fileId = match[1];
      var file   = DriveApp.getFileById(fileId);
      var blob   = file.getBlob();
      imgMime    = blob.getContentType() || 'image/jpeg';
      imgBase64  = Utilities.base64Encode(blob.getBytes());
    } else {
      // Si no es Drive, intentar fetch directo
      var resp = UrlFetchApp.fetch(urlFoto, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) return { ok: false, error: 'No se pudo descargar la foto' };
      imgMime   = resp.getHeaders()['Content-Type'] || 'image/jpeg';
      imgBase64 = Utilities.base64Encode(resp.getBlob().getBytes());
    }
  } catch(e) {
    return { ok: false, error: 'Descarga foto: ' + e.message };
  }

  // 3. Llamar a Claude con visión
  var system = [
    'Eres un asistente experto en lectura de comprobantes de pago peruanos (SUNAT).',
    'Recibes una imagen y debes extraer los datos del documento.',
    '',
    'TIPOS DE COMPROBANTE:',
    '- FACTURA: tiene RUC del emisor + IGV desglosado (18%) → IGV es recuperable',
    '- BOLETA_VENTA con RUC: emisor identificado pero sin IGV recuperable',
    '- TICKET o NOTA_DE_VENTA: sin IGV → NO recuperable',
    '- NO_COMPROBANTE: la imagen no es un documento fiscal (es un producto, escena, etc.)',
    '- ILEGIBLE: la imagen está borrosa, oscura o no se ve el documento',
    '',
    'Si extraes IGV, debe coincidir con el formato peruano (18% del subtotal gravado).',
    'Si el total es S/ 118 y se ve "IGV 18" o "IGV S/ 18.00", entonces igvRecuperable=18.',
    '',
    'RESPONDE EXCLUSIVAMENTE con JSON válido (sin markdown, sin comentarios):',
    '{',
    '  "tipoComprobante": "FACTURA" | "BOLETA_VENTA" | "TICKET" | "NO_COMPROBANTE" | "ILEGIBLE",',
    '  "rucEmisor": "20XXXXXXXXX" (11 dígitos) o "",',
    '  "razonSocial": "string" o "",',
    '  "serie": "F001" o "B001" o "" (la serie del documento),',
    '  "numero": "0000123" o "" (el número del documento),',
    '  "fecha": "DD/MM/YYYY" o "",',
    '  "total": número o 0 (total del documento en soles),',
    '  "subtotal": número o 0 (gravada sin IGV — solo si es FACTURA),',
    '  "igvRecuperable": número o 0 (solo > 0 si es FACTURA con IGV discriminado),',
    '  "confidence": 0-100 (qué tan seguro estás de los datos extraídos),',
    '  "estado": "PROCESADO" | "SIN_IGV" | "ILEGIBLE" | "NO_COMPROBANTE",',
    '  "notas": "string corto explicando el caso si aplica"',
    '}'
  ].join('\n');

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'KEY_NOT_SET' };

  var payload = {
    model: IA_MODELO_DEFAULT,
    max_tokens: 1024,
    system: system,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: imgMime, data: imgBase64 }
        },
        {
          type: 'text',
          text: 'Analiza este comprobante de proveedor y devuelve el JSON con la estructura indicada.'
        }
      ]
    }]
  };

  try {
    var resp = UrlFetchApp.fetch(IA_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 200) {
      return { ok: false, error: 'API_ERROR_' + code, mensaje: resp.getContentText().substring(0, 300) };
    }
    var apiData = JSON.parse(resp.getContentText());
    var text = (apiData.content && apiData.content[0] && apiData.content[0].text) || '';

    // Extraer JSON entre primer { y último }
    var first = text.indexOf('{');
    var last  = text.lastIndexOf('}');
    if (first < 0 || last < 0) {
      return { ok: false, error: 'PARSE_FAIL', mensaje: text.substring(0, 200) };
    }
    var parsed;
    try { parsed = JSON.parse(text.substring(first, last + 1)); }
    catch(e) { return { ok: false, error: 'PARSE_FAIL', mensaje: e.message + ' · raw: ' + text.substring(0, 200) }; }

    var resultado = {
      tipoComprobante: String(parsed.tipoComprobante || 'NO_COMPROBANTE'),
      rucEmisor:       String(parsed.rucEmisor || ''),
      razonSocial:     String(parsed.razonSocial || ''),
      serie:           String(parsed.serie || ''),
      numero:          String(parsed.numero || ''),
      fecha:           String(parsed.fecha || ''),
      total:           parseFloat(parsed.total) || 0,
      subtotal:        parseFloat(parsed.subtotal) || 0,
      igvRecuperable:  parseFloat(parsed.igvRecuperable) || 0,
      confidence:      parseInt(parsed.confidence, 10) || 0,
      estado:          String(parsed.estado || 'NO_COMPROBANTE'),
      notas:           String(parsed.notas || '')
    };

    // 4. Persistir en la hoja GUIAS — auto-crear columnas si no existen
    _persistirOCRFacturaEnGuia(sheet, filaIdx, resultado);

    return { ok: true, data: resultado };
  } catch(e) {
    return { ok: false, error: 'NETWORK', mensaje: e.message };
  }
}

// Helper: persiste el resultado del OCR en columnas dedicadas en GUIAS.
// Auto-crea las columnas si no existen (idempotente).
function _persistirOCRFacturaEnGuia(sheet, filaIdx, r) {
  var COLS_OCR = [
    'OCR_Estado',           // PROCESADO / SIN_IGV / ILEGIBLE / NO_COMPROBANTE
    'OCR_Tipo',             // FACTURA / BOLETA_VENTA / TICKET / etc
    'OCR_RUC_Emisor',
    'OCR_Razon_Social',
    'OCR_Serie',
    'OCR_Numero',
    'OCR_Fecha_Comprobante',
    'OCR_Total',
    'OCR_Subtotal',
    'IGV_Recuperable',      // KPI clave del Centro Tributario
    'OCR_Confidence',
    'OCR_Notas',
    'OCR_Fecha_Proceso'
  ];
  var hdrRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1));
  var hdrs = hdrRange.getValues()[0].map(function(h){ return String(h).trim(); });
  // Auto-crear columnas faltantes
  COLS_OCR.forEach(function(col){
    if (hdrs.indexOf(col) < 0) {
      var nuevaCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nuevaCol).setValue(col).setFontWeight('bold');
      hdrs.push(col);
    }
  });

  var fila = filaIdx + 1; // 1-indexed
  var valores = {
    OCR_Estado:            r.estado,
    OCR_Tipo:              r.tipoComprobante,
    OCR_RUC_Emisor:        r.rucEmisor,
    OCR_Razon_Social:      r.razonSocial,
    OCR_Serie:             r.serie,
    OCR_Numero:            r.numero,
    OCR_Fecha_Comprobante: r.fecha,
    OCR_Total:             r.total,
    OCR_Subtotal:          r.subtotal,
    IGV_Recuperable:       r.igvRecuperable,
    OCR_Confidence:        r.confidence,
    OCR_Notas:             r.notas,
    OCR_Fecha_Proceso:     new Date().toISOString()
  };
  Object.keys(valores).forEach(function(col){
    var idx = hdrs.indexOf(col);
    if (idx >= 0) sheet.getRange(fila, idx + 1).setValue(valores[col]);
  });
}

// ============================================================
// [v2.13.41] igvFavorMes — devuelve guías del mes con IGV recuperable
// usado por el Centro Tributario en MOS para sumar el IGV a favor.
// params: { mes: 1-12, año: 2026 }
// ============================================================
function igvFavorMes(params) {
  var mes = parseInt(params.mes, 10);
  var anio = parseInt(params.anio || params.año || params.year, 10);
  if (!mes || !anio) {
    var hoy = new Date(); mes = hoy.getMonth() + 1; anio = hoy.getFullYear();
  }
  var sheet = getSheet('GUIAS');
  if (!sheet) return { ok: false, error: 'GUIAS no encontrada' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var idxId    = hdrs.indexOf('idGuia');
  var idxTipo  = hdrs.indexOf('tipo');
  var idxFecha = hdrs.indexOf('fecha');
  var idxProv  = hdrs.indexOf('idProveedor');
  var idxFoto  = hdrs.indexOf('foto');
  var idxOCRE  = hdrs.indexOf('OCR_Estado');
  var idxRUC   = hdrs.indexOf('OCR_RUC_Emisor');
  var idxRS    = hdrs.indexOf('OCR_Razon_Social');
  var idxSerie = hdrs.indexOf('OCR_Serie');
  var idxNum   = hdrs.indexOf('OCR_Numero');
  var idxFchC  = hdrs.indexOf('OCR_Fecha_Comprobante');
  var idxTotal = hdrs.indexOf('OCR_Total');
  var idxIGV   = hdrs.indexOf('IGV_Recuperable');
  var idxConf  = hdrs.indexOf('OCR_Confidence');

  var lista = [];
  var totalIGV = 0;
  var totalGuiasConFoto = 0;
  var totalGuiasSinFoto = 0;
  var totalSinIGV = 0;
  var totalIlegibles = 0;

  for (var i = 1; i < data.length; i++) {
    var tipo = String(data[i][idxTipo] || '');
    if (tipo !== 'INGRESO_PROVEEDOR') continue;
    var fecha = data[i][idxFecha];
    var d = fecha instanceof Date ? fecha : new Date(fecha);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== anio || (d.getMonth() + 1) !== mes) continue;

    var hayFoto  = !!String(data[i][idxFoto] || '').trim();
    var ocrEstado = idxOCRE >= 0 ? String(data[i][idxOCRE] || '') : '';
    var igvRec    = idxIGV >= 0 ? parseFloat(data[i][idxIGV]) || 0 : 0;

    if (!hayFoto) totalGuiasSinFoto++;
    else if (ocrEstado === 'PROCESADO' && igvRec > 0) { totalIGV += igvRec; totalGuiasConFoto++; }
    else if (ocrEstado === 'SIN_IGV') totalSinIGV++;
    else if (ocrEstado === 'ILEGIBLE' || ocrEstado === 'NO_COMPROBANTE') totalIlegibles++;
    else if (hayFoto && !ocrEstado) totalIlegibles++; // foto sin procesar aún

    lista.push({
      idGuia:        String(data[i][idxId] || ''),
      fecha:         d.toISOString(),
      idProveedor:   String(data[i][idxProv] || ''),
      tieneFoto:     hayFoto,
      // [v2.13.43] URL de la foto para mostrar thumbnail + abrir comprobante grande
      urlFoto:       hayFoto ? String(data[i][idxFoto] || '') : '',
      ocrEstado:     ocrEstado,
      rucEmisor:     idxRUC >= 0 ? String(data[i][idxRUC] || '') : '',
      razonSocial:   idxRS  >= 0 ? String(data[i][idxRS] || '')  : '',
      serie:         idxSerie >= 0 ? String(data[i][idxSerie] || '') : '',
      numero:        idxNum >= 0 ? String(data[i][idxNum] || '') : '',
      fechaComprobante: idxFchC >= 0 ? String(data[i][idxFchC] || '') : '',
      total:         idxTotal >= 0 ? parseFloat(data[i][idxTotal]) || 0 : 0,
      igvRecuperable: igvRec,
      confidence:    idxConf >= 0 ? parseInt(data[i][idxConf], 10) || 0 : 0
    });
  }
  // Ordenar por fecha desc
  lista.sort(function(a, b) { return new Date(b.fecha) - new Date(a.fecha); });

  return {
    ok: true,
    data: {
      mes: mes, anio: anio,
      totalIGVFavor: Math.round(totalIGV * 100) / 100,
      totalGuias: lista.length,
      totalGuiasConIGV: totalGuiasConFoto,
      totalGuiasSinFoto: totalGuiasSinFoto,
      totalGuiasSinIGV: totalSinIGV,
      totalGuiasIlegibles: totalIlegibles,
      guias: lista
    }
  };
}

// Re-procesar OCR de una guía (admin desde MOS si la foto se actualizó
// o si quiere refrescar el análisis).
function reprocesarOCRGuia(params) {
  return analizarFacturaProveedor(params);
}

// ============================================================
// [v2.13.42] procesarOCRMasivoMes — corre el OCR sobre TODAS las guías
// INGRESO_PROVEEDOR del mes que tengan foto pero NO tengan OCR_Estado
// (o tengan estado vacío). Útil para arranque del Centro Tributario:
// procesa todo el histórico de fotos sin OCR para que el IGV a favor
// del mes se pueble retroactivamente.
//
// params: { mes, anio, soloSinProcesar }
//   soloSinProcesar=true (default): solo guías sin OCR_Estado o con estado vacío
//   soloSinProcesar=false: re-procesa TODAS las del mes (cuidado: usa créditos)
//
// Procesa secuencial con delay 500ms entre llamadas (no saturar API).
// 30 guías ≈ 60-90s total. Devuelve resumen con stats.
// ============================================================
function procesarOCRMasivoMes(params) {
  var mes = parseInt(params.mes, 10);
  var anio = parseInt(params.anio || params.año || params.year, 10);
  if (!mes || !anio) { var hoy = new Date(); mes = hoy.getMonth() + 1; anio = hoy.getFullYear(); }
  var soloSinProcesar = params.soloSinProcesar !== false; // default true

  var sheet = getSheet('GUIAS');
  if (!sheet) return { ok: false, error: 'GUIAS no encontrada' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var idxId    = hdrs.indexOf('idGuia');
  var idxTipo  = hdrs.indexOf('tipo');
  var idxFecha = hdrs.indexOf('fecha');
  var idxFoto  = hdrs.indexOf('foto');
  var idxOCRE  = hdrs.indexOf('OCR_Estado');

  // Recolectar candidatas
  var candidatas = [];
  for (var i = 1; i < data.length; i++) {
    var tipo = String(data[i][idxTipo] || '');
    if (tipo !== 'INGRESO_PROVEEDOR') continue;
    var fecha = data[i][idxFecha];
    var d = fecha instanceof Date ? fecha : new Date(fecha);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== anio || (d.getMonth() + 1) !== mes) continue;
    var hayFoto = !!String(data[i][idxFoto] || '').trim();
    if (!hayFoto) continue;
    var ocrEstado = idxOCRE >= 0 ? String(data[i][idxOCRE] || '') : '';
    if (soloSinProcesar && ocrEstado) continue;
    candidatas.push({ idGuia: String(data[i][idxId] || ''), fila: i + 1 });
  }

  var stats = {
    ok: true,
    mes: mes, anio: anio,
    soloSinProcesar: soloSinProcesar,
    total: candidatas.length,
    procesadas: 0,
    conIGV: 0,
    sinIGV: 0,
    ilegibles: 0,
    no_comprobante: 0,
    errores: 0,
    totalIGVRecuperado: 0,
    detalles: []
  };

  // Apps Script timeout = 6 min. A ~2s por OCR + 500ms delay = ~2.5s/guía.
  // Permite ~120 guías como máximo en una corrida. 30 está holgado.
  // Si hay más de 120 candidatas, partir en batches.
  var maxPorRun = 120;
  if (candidatas.length > maxPorRun) {
    candidatas = candidatas.slice(0, maxPorRun);
    stats.batched = true;
    stats.mensaje = 'Procesando primeras ' + maxPorRun + ' guías — corre de nuevo para continuar';
  }

  for (var j = 0; j < candidatas.length; j++) {
    var g = candidatas[j];
    try {
      var r = analizarFacturaProveedor({ idGuia: g.idGuia });
      if (r.ok && r.data) {
        stats.procesadas++;
        var est = r.data.estado;
        if (est === 'PROCESADO' && r.data.igvRecuperable > 0) {
          stats.conIGV++;
          stats.totalIGVRecuperado += r.data.igvRecuperable;
        } else if (est === 'SIN_IGV')       stats.sinIGV++;
        else if (est === 'ILEGIBLE')        stats.ilegibles++;
        else if (est === 'NO_COMPROBANTE')  stats.no_comprobante++;
        stats.detalles.push({
          idGuia: g.idGuia, estado: est,
          igv: r.data.igvRecuperable, conf: r.data.confidence
        });
      } else {
        stats.errores++;
        stats.detalles.push({ idGuia: g.idGuia, error: r.error || 'sin detalle' });
      }
    } catch(e) {
      stats.errores++;
      stats.detalles.push({ idGuia: g.idGuia, error: e.message });
    }
    Utilities.sleep(500); // throttle API
  }

  stats.totalIGVRecuperado = Math.round(stats.totalIGVRecuperado * 100) / 100;
  Logger.log('[OCR masivo] ' + JSON.stringify(stats));
  return stats;
}

// Variante que también testea el parser de listas con un ejemplo real
function testAnalizarLista() {
  Logger.log('— Test analizarListaSombra —');
  var ejemplo = 'Codigo  Descripcion              Pedido\n' +
                '111111  AJI PANCA POLVO 250GR    5.0\n' +
                '222222  AJINOMOTO 1KG BOLSA      18\n' +
                '333333  AZUCAR RUBIA SACO        2';
  var res = analizarListaSombra({ texto: ejemplo });
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// ============================================================
// [v2.13.42] extraerCostosFactura — OCR ITEMS LÍNEA POR LÍNEA
// del comprobante del proveedor (FACTURA / BOLETA / TICKET / nota
// manuscrita). El usuario sube la foto al ingresar guía, este OCR
// pre-puebla los costos para reducir tipeo del admin.
// ------------------------------------------------------------
// params: { idGuia }   // toma la foto del campo `foto` de GUIAS
// Retorna: {
//   ok: true,
//   data: {
//     items: [
//       { descripcion, cantidad, precioUnitario, subtotal, confidence }
//     ],
//     totalDocumento: número (header),
//     confidenceGlobal: 0-100 (peor del lote),
//     notas: string
//   }
// }
// ============================================================
function extraerCostosFactura(params) {
  var idGuia = String(params && params.idGuia || '').trim();
  if (!idGuia) return { ok: false, error: 'idGuia requerido' };

  // 1. Reusar la lógica de descarga de foto de analizarFacturaProveedor
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheet('GUIAS');
  if (!sheet) return { ok: false, error: 'Hoja GUIAS no encontrada' };
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h).trim(); });
  var idxId  = hdrs.indexOf('idGuia');
  var idxFoto = hdrs.indexOf('foto');
  if (idxId < 0 || idxFoto < 0) return { ok: false, error: 'Columnas idGuia/foto no encontradas' };
  var filaIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === idGuia) { filaIdx = i; break; }
  }
  if (filaIdx < 0) return { ok: false, error: 'Guía ' + idGuia + ' no encontrada' };
  var urlFoto = String(data[filaIdx][idxFoto] || '').trim();
  if (!urlFoto) return { ok: false, error: 'La guía no tiene foto del comprobante' };

  var imgBase64, imgMime;
  try {
    var match = urlFoto.match(/[?&]id=([^&]+)/);
    if (match) {
      var file = DriveApp.getFileById(match[1]);
      var blob = file.getBlob();
      imgMime   = blob.getContentType() || 'image/jpeg';
      imgBase64 = Utilities.base64Encode(blob.getBytes());
    } else {
      var resp = UrlFetchApp.fetch(urlFoto, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) return { ok: false, error: 'No se pudo descargar la foto' };
      imgMime   = resp.getHeaders()['Content-Type'] || 'image/jpeg';
      imgBase64 = Utilities.base64Encode(resp.getBlob().getBytes());
    }
  } catch(e) { return { ok: false, error: 'Descarga foto: ' + e.message }; }

  // 2. Prompt enfocado a items
  var system = [
    'Eres un experto en lectura de documentos comerciales peruanos.',
    'Recibes una imagen de un comprobante (factura, boleta, ticket impreso, ',
    'o incluso una hoja manuscrita) y debes extraer la LISTA DE PRODUCTOS',
    'con sus precios unitarios.',
    '',
    'POR CADA LÍNEA DE PRODUCTO devuelve:',
    '- descripcion: nombre del producto tal como aparece (MAYÚSCULAS, limpio)',
    '- cantidad: número decimal (ej: 2.0, 5.5)',
    '- precioUnitario: precio por unidad en soles (no el subtotal)',
    '- subtotal: cantidad × precio unitario, o el valor que aparece en la línea',
    '- confidence: 0-100, qué tan seguro estás de los 3 valores anteriores juntos',
    '',
    'INTERPRETACIÓN DE CONFIDENCE:',
    '- 95-100: texto impreso claro, valores nítidos',
    '- 80-94:  texto algo borroso o letra manuscrita legible',
    '- 60-79:  letra difícil, posibles errores en algún dígito',
    '- <60:    no estoy seguro, posible error grave (números ambiguos)',
    '',
    'REGLAS IMPORTANTES:',
    '- Si solo ves total sin desglose por línea, retorna items: []',
    '- Si el comprobante es de un producto sin detalle, retorna 1 item con la info',
    '- Ignora encabezados, totales finales, IGV, propinas',
    '- Si una línea no tiene precio claro, OMÍTELA (no inventes)',
    '- Si una cantidad no se ve, asume 1',
    '',
    'RESPONDE EXCLUSIVAMENTE con JSON válido (sin markdown, sin comentarios):',
    '{',
    '  "items": [',
    '    {"descripcion": "...", "cantidad": N, "precioUnitario": N, "subtotal": N, "confidence": N},',
    '    ...',
    '  ],',
    '  "totalDocumento": N,',
    '  "confidenceGlobal": N,',
    '  "notas": "string corto"',
    '}'
  ].join('\n');

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'KEY_NOT_SET' };

  var payload = {
    model:      IA_MODELO_DEFAULT,
    max_tokens: 4096,                        // listas largas
    system:     system,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: imgMime, data: imgBase64 } },
        { type: 'text', text: 'Analiza este comprobante y extrae la lista de productos con sus precios. Devuelve solo el JSON.' }
      ]
    }]
  };

  try {
    var resp2 = UrlFetchApp.fetch(IA_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (resp2.getResponseCode() !== 200) {
      return { ok: false, error: 'API_ERROR_' + resp2.getResponseCode(), mensaje: resp2.getContentText().substring(0, 300) };
    }
    var apiData = JSON.parse(resp2.getContentText());
    var text = (apiData.content && apiData.content[0] && apiData.content[0].text) || '';
    var first = text.indexOf('{');
    var last  = text.lastIndexOf('}');
    if (first < 0 || last < 0) return { ok: false, error: 'PARSE_FAIL', mensaje: text.substring(0, 200) };
    var parsed = JSON.parse(text.substring(first, last + 1));

    var items = (parsed.items || []).map(function(it) {
      return {
        descripcion:    String(it.descripcion || '').toUpperCase().trim(),
        cantidad:       parseFloat(it.cantidad) || 1,
        precioUnitario: parseFloat(it.precioUnitario) || 0,
        subtotal:       parseFloat(it.subtotal) || 0,
        confidence:     Math.max(0, Math.min(100, parseInt(it.confidence, 10) || 0))
      };
    }).filter(function(it){ return it.descripcion && it.precioUnitario > 0; });

    return { ok: true, data: {
      items:            items,
      totalDocumento:   parseFloat(parsed.totalDocumento) || 0,
      confidenceGlobal: Math.max(0, Math.min(100, parseInt(parsed.confidenceGlobal, 10) || 0)),
      notas:            String(parsed.notas || '')
    }};
  } catch(eA) { return { ok: false, error: 'NETWORK', mensaje: eA.message }; }
}

// ============================================================
// [v2.13.42] extraerCorreccionesJefa — OCR DEL TICKET LLENO POR
// LA JEFA. La foto viene en base64 del cliente (no de Drive).
// El admin saca foto desde su tablet/celular.
// ------------------------------------------------------------
// Estrategia: pasamos a Claude la foto + un CONTEXTO con la lista
// de items que ya estaban en el ticket impreso (skuBase, descripcion,
// costo, venta actual, margen objetivo). Claude debe matchear lo que
// jefa escribió a cada item del contexto y extraer:
//   - skuBase del item original (matching)
//   - ventaNueva (si escribió un nuevo precio venta)
//   - margenNuevoPct (si escribió un nuevo margen, 0..1)
//   - tachó (true si parece tachado/anulado)
//   - confidence (qué tan seguro está del valor extraído)
//
// params:
//   fotoBase64:    string (data:image/jpeg;base64,...)
//   contextoItems: array de { skuBase, descripcion, costo, ventaActual, margenActualPct }
// ============================================================
function extraerCorreccionesJefa(params) {
  if (!params || !params.fotoBase64) return { ok: false, error: 'fotoBase64 requerida' };
  if (!params.contextoItems || !Array.isArray(params.contextoItems) || !params.contextoItems.length) {
    return { ok: false, error: 'contextoItems[] requerido' };
  }
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'KEY_NOT_SET' };

  // Limpiar prefijo data:image/...;base64,
  var fotoBase64 = String(params.fotoBase64);
  var imgMime = 'image/jpeg';
  var m = fotoBase64.match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (m) { imgMime = m[1]; fotoBase64 = m[2]; }

  // Construir contexto compacto para el prompt
  var ctxList = params.contextoItems.map(function(it, idx) {
    return (idx + 1) + '. SKU:' + it.skuBase
         + ' · ' + String(it.descripcion || '').substring(0, 50)
         + ' · COSTO:' + (parseFloat(it.costo) || 0).toFixed(2)
         + ' · VENTA-ACT:' + (parseFloat(it.ventaActual) || 0).toFixed(2)
         + ' · MARGEN-OBJ:' + ((parseFloat(it.margenActualPct) || 0) * 100).toFixed(0) + '%';
  }).join('\n');

  var system = [
    'Eres un asistente que lee correcciones manuscritas en un ticket impreso.',
    'El ticket fue impreso desde un sistema con cuadros por producto.',
    'La jefa lo recibió, lo llenó a mano (escribió nuevos precios o márgenes),',
    'lo devolvió, y ahora tienes la foto del ticket lleno.',
    '',
    'TU TAREA: por cada producto del CONTEXTO de abajo, identifica si la jefa',
    'escribió algo nuevo y extrae los valores.',
    '',
    'CONTEXTO (productos que estaban en el ticket impreso):',
    ctxList,
    '',
    'POR CADA producto donde detectes escritura de la jefa, devuelve:',
    '- skuBase: del contexto (numerado arriba)',
    '- ventaNueva: número si escribió nuevo precio venta, null si no',
    '- margenNuevoPct: número 0..1 si escribió porcentaje (60% = 0.60), null si no',
    '- tachado: true si tachó algún campo (ignorar producto)',
    '- confidence: 0-100',
    '- notas: opcional (string corto)',
    '',
    'REGLAS:',
    '- Si jefa NO escribió nada en un producto, OMÍTELO del resultado',
    '- Si escribió "100" sin contexto, probablemente es precio venta (no margen)',
    '- Si escribió "60%" o "60 %", es margen objetivo',
    '- Si escribió ambos (venta + margen), retorna ambos',
    '- Letra manuscrita es ambigua: si dudas entre 8 y 3, baja confidence',
    '- Confidence ≥95: muy seguro · 80-94: probable · <80: dudoso',
    '',
    'RESPONDE EXCLUSIVAMENTE con JSON válido:',
    '{',
    '  "correcciones": [',
    '    {"skuBase": "...", "ventaNueva": N|null, "margenNuevoPct": N|null, "tachado": false, "confidence": N, "notas": ""}',
    '  ],',
    '  "confidenceGlobal": N,',
    '  "notas": "string corto"',
    '}'
  ].join('\n');

  var payload = {
    model:      IA_MODELO_DEFAULT,
    max_tokens: 4096,
    system:     system,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: imgMime, data: fotoBase64 } },
        { type: 'text',  text: 'Lee el ticket lleno por la jefa y devuelve las correcciones detectadas.' }
      ]
    }]
  };

  try {
    var resp3 = UrlFetchApp.fetch(IA_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (resp3.getResponseCode() !== 200) {
      return { ok: false, error: 'API_ERROR_' + resp3.getResponseCode(), mensaje: resp3.getContentText().substring(0, 300) };
    }
    var apiData2 = JSON.parse(resp3.getContentText());
    var text2 = (apiData2.content && apiData2.content[0] && apiData2.content[0].text) || '';
    var first2 = text2.indexOf('{');
    var last2  = text2.lastIndexOf('}');
    if (first2 < 0 || last2 < 0) return { ok: false, error: 'PARSE_FAIL', mensaje: text2.substring(0, 200) };
    var parsed2 = JSON.parse(text2.substring(first2, last2 + 1));

    var correcciones = (parsed2.correcciones || []).map(function(c) {
      return {
        skuBase:        String(c.skuBase || '').trim(),
        ventaNueva:     (c.ventaNueva !== null && c.ventaNueva !== undefined) ? parseFloat(c.ventaNueva) : null,
        margenNuevoPct: (c.margenNuevoPct !== null && c.margenNuevoPct !== undefined) ? parseFloat(c.margenNuevoPct) : null,
        tachado:        !!c.tachado,
        confidence:     Math.max(0, Math.min(100, parseInt(c.confidence, 10) || 0)),
        notas:          String(c.notas || '')
      };
    }).filter(function(c){
      return c.skuBase && (c.ventaNueva !== null || c.margenNuevoPct !== null || c.tachado);
    });

    return { ok: true, data: {
      correcciones:     correcciones,
      confidenceGlobal: Math.max(0, Math.min(100, parseInt(parsed2.confidenceGlobal, 10) || 0)),
      notas:            String(parsed2.notas || '')
    }};
  } catch(eC) { return { ok: false, error: 'NETWORK', mensaje: eC.message }; }
}
