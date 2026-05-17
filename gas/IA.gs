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
  if (texto.length > 8000) return { ok: false, error: 'TEXTO_MUY_LARGO', mensaje: 'Max 8000 caracteres' };

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

  var ia = _llamarClaude({
    max_tokens: 2048,
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
