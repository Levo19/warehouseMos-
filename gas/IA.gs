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
  if (texto.length > 30000) return { ok: false, error: 'TEXTO_MUY_LARGO', mensaje: 'Max 30000 caracteres (~500 productos)' };

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
