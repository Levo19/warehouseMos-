// warehouseMos — api.js  Comunicación con GAS + soporte offline
'use strict';

const API = (() => {
  function _gasUrl() { return window.WH_CONFIG?.gasUrl || ''; }

  // [cero-GAS · WH 2026-07-14] Acciones PROPIAS de WH que NO deben tocar GAS jamás. Si un POST de estas llega
  // al fallback GAS es porque su directo no commiteó (no cableada / Edge caída / rama muerta detrás de OpLog).
  // Fail-closed con error claro en vez de caer a GAS. Reads ya son 100% cero-GAS (call()→caché, nunca GAS).
  // NO incluye los módulos EXTERNOS Membrete/Seguridad (se cargan de assets MOS y conservan su fallback hasta
  // migrarse en su repo): agregarlos acá los rompería. Auditoría: solo estas 7 quedaban SOLO-GAS y propias.
  const _WH_NO_GAS = new Set([
    'registrarMerma', 'resolverMerma',   // [cero-rastro 2026-07-19] ramas retiradas: v2 = mermaAltaManualV2/procesarMerma; un replay viejo NUNCA cae a GAS
    'iniciarTestDiagnostico', 'finalizarTestDiagnostico', 'runInternalTests',  // panel de diagnóstico QA (marginal)
    'agregarAMermas', 'solucionarMerma',                                        // ramas muertas detrás de OpLog
    'subirFotoEntidad', 'eliminarFotoEntidad'                                   // sin callers
  ]);

  // ════════════════════════════════════════════════════════════════════
  // [PASO 5 · B3-frontend] Lectura DIRECTA a Supabase (navegador→PostgREST).
  // INERTE por defecto: solo se activa con localStorage 'wh_lectura_navegador'='1'
  // (o window.WH_CONFIG.lecturaNavegador===true). Ante CUALQUIER fallo cae a GAS.
  // url + anon key son PÚBLICOS (van en el cliente; la RLS protege en el server vía
  // el claim app=warehouseMos del JWT que mintea GAS en /mintTokenWH — B1).
  // Backend RLS LISTO para: getStock (wh.stock_enriquecido_rls) y getRotacionSemanal
  // (wh.rotacion_semanal_rls). El resto de lecturas espera sus wrappers RLS (sesión
  // backend futura) → _callDirecto devuelve null y siguen yendo por GAS aunque on.
  // ════════════════════════════════════════════════════════════════════
  const _SB_URL  = 'https://rzbzdeipbtqkzjqdchqk.supabase.co';
  const _SB_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ6YnpkZWlwYnRxa3pqcWRjaHFrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA4NzYwMDQsImV4cCI6MjA5NjQ1MjAwNH0.MAlSdz_ugGUZoaU5st6dA_gb_x_IiUL0TXxH176kY9k';
  const _sbTok = { token: null, exp: 0 };

  function _whLecturaDirecta() {
    try { return localStorage.getItem('wh_lectura_navegador') === '1' || window.WH_CONFIG?.lecturaNavegador === true; }
    catch (_) { return window.WH_CONFIG?.lecturaNavegador === true; }
  }

  function _whFetchTimeout(url, opts, ms) {
    const ctrl = new AbortController();
    const tid  = setTimeout(() => ctrl.abort(), ms || 12000);
    return fetch(url, { ...opts, signal: ctrl.signal }).finally(() => clearTimeout(tid));
  }

  // Pide el JWT (app=warehouseMos) y lo cachea. PRIMARIO: Edge Function `mint-wh` (HS256, exp ~30min, ~50-150ms)
  // → corta la dependencia de GAS y, con TTL 30min + refresh proactivo (abajo), saca el mint del hilo de navegación.
  // FALLBACK: si la Edge falla/timeout/{ok:false} (incl. 404 si aún no está viva), cae a GAS `mintTokenWH` como hoy
  //   → red de seguridad: nunca rompe el login aunque la Edge no exista todavía.
  // Re-mintea 30s antes de expirar (camino sincrónico, último recurso); el refresh proactivo debería adelantarse.
  // _mintInFlight dedup: si varias lecturas salen juntas (arranque) → 1 solo POST, no ráfaga.
  let _mintInFlight = null;

  function _whDeviceId() {
    try { return localStorage.getItem('wh_device_id') || ''; } catch (_) { return ''; }
  }

  // Edge `mint-wh`: verify_jwt=false → va con `apikey` (anon, público), SIN Authorization (es quien EMITE el token).
  // Devuelve {ok,token,exp} igual que GAS. Lanza si la Edge no devuelve un token válido → el caller cae a GAS.
  async function _mintViaEdge(deviceId) {
    const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/mint-wh`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Content-Type': 'application/json' },
      body: JSON.stringify({ deviceId })
    }, 6000);
    const d = await res.json().catch(() => null);
    if (!d || !d.ok || !d.token) throw new Error('mint-wh edge: ' + ((d && d.error) || res.status));
    return d;
  }

  async function _mintTokenWH() {
    const now = Math.floor(Date.now() / 1000);
    if (_sbTok.token && (_sbTok.exp - now) > 30) { _agendarRefresh(); return _sbTok.token; }
    if (_mintInFlight) return _mintInFlight;
    _mintInFlight = (async () => {
      const deviceId = _whDeviceId();
      // [CERO-GAS / CERO-FALLBACK] Solo Edge mint-wh. Si falla, propaga el throw → el caller reintenta
      // (el finally de abajo limpia _mintInFlight); ya no cae al GAS mintTokenWH.
      const d = await _mintViaEdge(deviceId);
      const n = Math.floor(Date.now() / 1000);
      _sbTok.token = d.token; _sbTok.exp = d.exp || (n + 1800);
      _agendarRefresh();
      return d.token;
    })();
    try { return await _mintInFlight; }
    finally { _mintInFlight = null; }
  }

  // Refresh PROACTIVO en background: re-mintea ~120s ANTES de expirar, fuera del camino crítico, para que una
  // navegación NUNCA dispare el mint sincrónico (la causa del "se congela al rato"). Fire-and-forget: si falla,
  // el camino sincrónico (Edge mint-wh, cero-GAS) reintenta en la próxima lectura.
  let _refreshTid = null;
  function _agendarRefresh() {
    if (_refreshTid) return;                                   // ya hay un refresh agendado
    const now = Math.floor(Date.now() / 1000);
    const margen = 120;                                        // re-mintear 2 min antes de exp
    let enMs = (_sbTok.exp - now - margen) * 1000;
    if (!isFinite(enMs) || enMs < 1000) enMs = 1000;          // mínimo 1s (token casi vencido)
    if (enMs > 1800000) enMs = 1800000;                       // tope 30min (defensivo)
    _refreshTid = setTimeout(async () => {
      _refreshTid = null;
      try {
        const deviceId = _whDeviceId();
        // [CERO-GAS] Solo Edge mint-wh. Si falla, el outer catch deja que el camino sincrónico re-mintee bajo
        // demanda (antes había un catch → _mintViaGAS(deviceId), función INEXISTENTE = ReferenceError muerto).
        const d = await _mintViaEdge(deviceId);
        const n = Math.floor(Date.now() / 1000);
        _sbTok.token = d.token; _sbTok.exp = d.exp || (n + 1800);
        _agendarRefresh();                                     // reencadena para el próximo ciclo
      } catch (_) { /* el camino sincrónico re-minteará bajo demanda */ }
    }, enMs);
    try { if (_refreshTid && _refreshTid.unref) _refreshTid.unref(); } catch (_) {}
  }

  // [PASO 5 · B5] Impresión DIRECTA vía Edge Function `imprimir` (reemplaza el salto a GAS→PrintNode).
  // El módulo arma el ESC/POS/ZPL (content base64) y lo manda acá; la Edge reenvía a PrintNode con la key (secret).
  // Devuelve true si PrintNode aceptó (status success). El llamador cae a GAS si false/excepción.
  // Convierte un string ESC/POS o ZPL a base64 de sus BYTES UTF-8 — idéntico a Utilities.base64Encode(t) de GAS.
  // Maneja los bytes de control (\x1b, \x1d < 128) y los acentos (multibyte UTF-8). Para que el módulo arme el ticket
  // (el mismo string que GAS) y lo mande con _imprimirDirecto sin diferencias de encoding.
  function _escposB64(str) {
    return btoa(unescape(encodeURIComponent(String(str))));
  }
  async function _imprimirDirecto(printerId, content, title) {
    try {
      const token = await _mintTokenWH();
      const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/imprimir`, {
        method: 'POST',
        headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ printerId: parseInt(printerId), content, title: title || 'warehouseMos' })
      }, 12000);
      const d = await res.json().catch(() => ({}));
      return !!(d && d.status === 'success');
    } catch (_) { return false; }
  }

  // [CERO-GAS F2] Resuelve el printerId de ticket para reportes admin (historial/cargadores):
  //   override explícito (selector admin) → WH_TICKET_PRINTER_ID de config → default conocido.
  let _whTicketPrinterCache = '';
  async function _resolverPrinterTicketWH(override) {
    if (override) return String(override).trim();
    if (_whTicketPrinterCache) return _whTicketPrinterCache;
    try {
      const cfg = await _sbRpcWH('get_config', {}, 'wh');
      const v = cfg && cfg.data && (cfg.data.WH_TICKET_PRINTER_ID || cfg.data.wh_ticket_printer_id);
      if (v) { _whTicketPrinterCache = String(v).trim(); return _whTicketPrinterCache; }
    } catch (_) {}
    return '75247847';   // impresora del almacén (fallback conocido)
  }
  // [CERO-GAS F2] Imprime un texto ya formateado (ESC/POS 80mm) vía la Edge `imprimir`. Envuelve con init+feed+corte.
  async function _imprimirTextoTicketWH(texto, title, printerIdOverride) {
    const printerId = await _resolverPrinterTicketWH(printerIdOverride);
    if (!printerId) return { ok: false, error: 'sin impresora de ticket' };
    const escpos = '\x1b\x40' + String(texto || '') + '\n\x1b\x4a\x60\x1d\x56\x00';   // init + texto + feed 96 + corte
    const ok = await _imprimirDirecto(printerId, _escposB64(escpos), title || 'warehouseMos');
    return ok ? { ok: true, data: { impreso: true } } : { ok: false, error: 'PrintNode rechazó' };
  }
  // [CERO-GAS F2] Arma el ESC/POS del reporte de cargadores del día (port del layout GAS). Defensivo ante shape vacío.
  function _armarCargadoresEscPos(d) {
    const SEP = '================================================', SEP2 = '------------------------------------------------';
    const pad = (a, b) => { a = String(a || ''); b = String(b || ''); const n = 48 - a.length - b.length; return a + (n > 0 ? ' '.repeat(n) : ' ') + b; };
    let t = '\x1b\x61\x01\x1b\x21\x38CARGADORES\n\x1b\x21\x00\x1b\x45\x01' + String(d.fecha || '').toUpperCase() + '\x1b\x45\x00\n\x1b\x61\x00' + SEP + '\n';
    t += pad('Preingresos del dia:', String(d.preingresos != null ? d.preingresos : (d.total || 0))) + '\n';
    t += pad('Carretas totales:', String(d.totalCarretas != null ? d.totalCarretas : '')) + '\n' + SEP2 + '\n';
    const cargs = Array.isArray(d.cargadores) ? d.cargadores : [];
    if (!cargs.length) t += '  (sin cargadores registrados ese dia)\n';
    cargs.forEach(c => {
      t += '\x1b\x45\x01' + String(c.nombre || '').toUpperCase() + '\x1b\x45\x00\n';
      t += pad('  ' + (c.carretasTotal != null ? c.carretasTotal : (c.carretas || 0)) + ' carretas',
               'L' + (c.llenasTotal || 0) + ' M' + (c.mediasTotal || 0) + ' CV' + (c.vaciasTotal || 0)) + '\n';
      (Array.isArray(c.preingresos) ? c.preingresos : []).forEach(pi => {
        t += pad('  - ' + (pi.idPreingreso || '') + ' ' + String(pi.proveedor || '').substring(0, 20),
                 (pi.carretas || 0) + 'c') + '\n';
      });
      t += '\n';
    });
    t += SEP + '\n\x1b\x61\x01\x1b\x45\x01TOTAL DEL DIA\x1b\x45\x00\n\x1b\x21\x38\x1b\x45\x01' + (d.totalCarretas || 0) + ' CARRETAS\x1b\x21\x00\x1b\x45\x00\n\x1b\x61\x00' + SEP;
    return t;
  }

  // [CERO-GAS push] Envía push a una AUDIENCIA vía Edge `push` (resuelve tokens de mos.push_tokens server-side).
  // Reemplaza el fetch a GAS notificarInicioSesionVendedor. Fire-and-forget, sin fallback GAS.
  async function _pushEdgeWH(audiencia, titulo, cuerpo, data) {
    try {
      const token = await _mintTokenWH();
      await _whFetchTimeout(`${_SB_URL}/functions/v1/push`, {
        method: 'POST',
        headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ op: 'send', audiencia, title: titulo, body: cuerpo, data: data || null })
      }, 8000);
    } catch (_) { /* informativo, no bloquea */ }
  }

  // [PASO 5 · B5] Subir foto a Supabase Storage (bucket wh-fotos) en MÁXIMA calidad. path: <tipo>/<id>/<único>.
  // Devuelve {url} (original, full quality) + {preview} (render on-the-fly liviano para listas) + {path}.
  async function _subirFotoStorage(tipo, id, base64, mime, nombreSeed) {
    const token = await _mintTokenWH();
    const ext = (mime || '').includes('png') ? 'png' : (mime || '').includes('webp') ? 'webp' : 'jpg';
    // [40x #2] limpiar prefijo data-URI (FileReader.readAsDataURL lo agrega) — sin esto atob() lanza.
    const b64 = String(base64 || '').replace(/^data:[^;]+;base64,/, '');
    // [40x #1] nombre DETERMINÍSTICO por localId (nombreSeed) → reintento = mismo path → no duplica fotos.
    const nombre = (nombreSeed != null && String(nombreSeed) !== '' ? String(nombreSeed) : (Date.now() + '_' + Math.random().toString(36).slice(2, 7))) + '.' + ext;
    const path = `${encodeURIComponent(tipo)}/${encodeURIComponent(id)}/${nombre}`;
    const bin = Uint8Array.from(atob(b64), ch => ch.charCodeAt(0));   // base64 → binario
    // [v2.13.240] CAUSA RAÍZ del 400 "la foto no aparece": el header `x-upsert: true` hacía que Storage
    // ejecutara INSERT … ON CONFLICT DO UPDATE. Ese camino evalúa la policy de UPDATE (USING) y necesita
    // LEER la fila en conflicto, pero NO existe policy SELECT para `authenticated` en este bucket → la RLS
    // rechaza el row → cuerpo {"statusCode":"403", "message":"new row violates row-level security policy"}
    // que Storage envuelve como HTTP 400 (NO 403 visible) → parecía "payload mal formado". Verificado con curl:
    // POST sin x-upsert = 200; POST con x-upsert = 400 RLS. El INSERT puro SÍ pasa (policy wh_fotos_insert).
    // La idempotencia se conserva SIN upsert: el nombre es DETERMINÍSTICO por localId → un reintento al mismo
    // path devuelve {"statusCode":"409","error":"Duplicate"} → lo tratamos como ÉXITO (la foto ya está ahí).
    const res = await _whFetchTimeout(`${_SB_URL}/storage/v1/object/wh-fotos/${path}`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': mime || 'image/jpeg' },
      body: bin
    }, 30000);   // 30s — fotos de alta resolución pesan
    if (!res.ok) {
      // El "status" REAL viene en el CUERPO JSON (Storage envuelve casi todo como HTTP 400). Leerlo para decidir.
      const body = await res.json().catch(() => null);
      const bodyCode = parseInt((body && body.statusCode), 10) || res.status;
      // 409 Duplicate = el objeto YA existe en este path determinístico (reintento idempotente) → ÉXITO, no error.
      if (bodyCode === 409 || (body && /duplicate/i.test(String(body.error || '')))) {
        return { ok: true, path, url: `${_SB_URL}/storage/v1/object/public/wh-fotos/${path}`, preview: `${_SB_URL}/storage/v1/render/image/public/wh-fotos/${path}?width=800&quality=72` };
      }
      // [400-loop fix] 4xx (≠429) = rechazo DEFINITIVO de Storage (RLS, mime no permitido, path inválido, payload).
      // Reintentarlo eternamente solo spamea. Marcamos `.permanente` → post()/cola lo DESCARTAN (no loop infinito).
      // 429 y 5xx/red = transitorio → error normal (la cola reintenta).
      const err = new Error('storage upload ' + bodyCode + (body && body.message ? ': ' + body.message : ''));
      if (bodyCode >= 400 && bodyCode < 500 && bodyCode !== 429) err.permanente = true;
      throw err;
    }
    return {
      ok: true, path,
      url:     `${_SB_URL}/storage/v1/object/public/wh-fotos/${path}`,                          // original (ver detalle/zoom)
      preview: `${_SB_URL}/storage/v1/render/image/public/wh-fotos/${path}?width=800&quality=72` // liviano (listas)
    };
  }

  // [PASO 5 · B5] Llama la Edge `ia` (proxy a Claude). body = {messages, system?, model?, max_tokens?}. Devuelve el JSON de Claude.
  async function _llamarEdgeIA(body, timeoutMs) {
    const token = await _mintTokenWH();
    const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/ia`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }, timeoutMs || 40000);   // la IA puede tardar; visión/PDF necesita MÁS (subida del archivo en red móvil)
    if (!res.ok) throw new Error('ia ' + res.status);
    return res.json();   // {content:[{text}], ...} de Claude
  }

  // [PASO 5 · B5] OCR de comprobante de proveedor — réplica FIEL de analizarFacturaProveedor (IA.gs:178-375).
  // Llama Claude (visión) vía Edge `ia`, parsea los 12 campos SUNAT y los persiste en wh.guias (guardar_ocr_guia).
  // Best-effort / fire-and-forget: NO bloquea la subida de la foto (igual que el disparo automático del GAS, Fotos.gs:59).
  async function _ocrComprobanteGuia(idGuia, base64, mime) {
    if (!idGuia || !base64) return null;
    // [40x] kill-switch fino en el cliente: permite activar las fotos directas SIN disparar OCR (no gastar IA) durante
    // el cutover. `localStorage 'wh_ocr_off'='1'` apaga la LLAMADA a Claude (el flag SQL solo apaga la persistencia).
    try { if (localStorage.getItem('wh_ocr_off') === '1') return null; } catch (_) {}
    const data = String(base64).replace(/^data:[^;]+;base64,/, '');   // base64 PURO (Claude rechaza el prefijo data-URI)
    const m = /^data:([^;]+);/.exec(String(base64));
    const mediaType = (m && m[1]) || mime || 'image/jpeg';
    // SYSTEM PROMPT literal de IA.gs:224-253 (réplica exacta — datos fiscales SUNAT, no alterar).
    const system = [
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
    let resp;
    try {
      resp = await _llamarEdgeIA({
        max_tokens: 1536, system,   // 1536 > 1024 del GAS: margen para JSON con notas largas (cap Edge=8192)
        messages: [{ role: 'user', content: [
          { type: 'image', source: { type: 'base64', media_type: mediaType, data } },
          { type: 'text', text: 'Analiza este comprobante de proveedor y devuelve el JSON con la estructura indicada.' }
        ] }]
      });
    } catch (_) { return null; }
    const text = (resp && resp.content && resp.content[0] && resp.content[0].text) || '';
    const first = text.indexOf('{'), last = text.lastIndexOf('}');
    if (first < 0 || last < 0) return null;
    let r;
    try { r = JSON.parse(text.substring(first, last + 1)); } catch (_) { return null; }
    // Normalización fiel (IA.gs:300-314): tipado/defaults idénticos a los que persistía el GAS.
    const p = {
      id_guia: idGuia,
      tipo_comprobante: String(r.tipoComprobante || 'NO_COMPROBANTE'),
      ruc_emisor:       String(r.rucEmisor || ''),
      razon_social:     String(r.razonSocial || ''),
      serie:            String(r.serie || ''),
      numero:           String(r.numero || ''),
      fecha:            String(r.fecha || ''),
      total:            parseFloat(r.total) || 0,
      subtotal:         parseFloat(r.subtotal) || 0,
      igv_recuperable:  parseFloat(r.igvRecuperable) || 0,
      confidence:       parseInt(r.confidence, 10) || 0,
      estado:           String(r.estado || 'NO_COMPROBANTE'),
      notas:            String(r.notas || '')
    };
    try { await _sbRpcWH('guardar_ocr_guia', { p }); } catch (_) {}
    return p;
  }

  // ── [perf v2.13.242] Dedup in-flight + micro-cache para LECTURAS pesadas ──
  // Causa raíz de la ráfaga de ~17 leer_tabla_rls: getDashboard (7 RPCs),
  // descargarOperacional (6 RPCs) y getStock (1 RPC) se disparaban en paralelo
  // al navegar/arrancar/timer-60s SIN compartir round-trips, repitiendo las
  // mismas tablas (guias/preingresos/auditorias) y stock_enriquecido varias
  // veces en la misma ventana de tiempo. _dedupRead colapsa llamadas idénticas:
  //   1. si hay una in-flight con la misma clave → devuelve ESA promesa (no abre red).
  //   2. si terminó hace < ttl ms → devuelve el resultado cacheado (sin red).
  // Es transparente: misma firma {ok,data}. NO toca escrituras ni el catálogo.
  const _readInflight = new Map();   // key -> Promise
  const _readCache    = new Map();   // key -> { ts, val }
  // [perf v2.13.242 · FIX stale-tras-escritura] Tras una ESCRITURA directa que muta
  // datos (guía creada/cerrada, ajuste, envasado, merma, preingreso…), el micro-cache
  // de 4s DEBE invalidarse: si no, el refresh inmediato (precargarOperacional(true) /
  // getStock) caía dentro de la ventana de 4s de una lectura PREVIA y devolvía el
  // stock/guías VIEJOS (sin el cambio recién hecho) → el operador NO veía su cambio.
  // Vacía solo el RESULTADO cacheado; NO toca las in-flight (esas ya están en red y se
  // resuelven solas; vaciar el cache hace que la PRÓXIMA lectura abra un round-trip
  // fresco). Barato (Map chico) y correcto: justo tras escribir querés datos frescos.
  function _invalidarLecturas() {
    try { _readCache.clear(); } catch (_) {}
  }
  function _dedupRead(key, ttlMs, fn) {
    const inf = _readInflight.get(key);
    if (inf) return inf;                                   // (1) comparte la in-flight
    const hit = _readCache.get(key);
    if (hit && (Date.now() - hit.ts) < ttlMs) {            // (2) micro-cache fresco
      return Promise.resolve(hit.val);
    }
    const p = Promise.resolve().then(fn).then(val => {
      _readCache.set(key, { ts: Date.now(), val });
      _readInflight.delete(key);
      return val;
    }).catch(err => {
      _readInflight.delete(key);                            // no cachear errores
      throw err;
    });
    _readInflight.set(key, p);
    return p;
  }

  // Llama una RPC de LECTURA directo a PostgREST (apikey + Bearer + Profile). profile='wh' (default) o 'mos' (catálogo).
  // [perf 500x] `ms` configurable: las RPC PESADAS (catálogo ~1.9MB, dashboard) legítimamente tardan >12s en
  // redes lentas o bajo cola; con 12s abortaban y caían a GAS (doble-pago directo-abortado + cold-start GAS).
  async function _sbRpcWH(fn, args, profile, ms) {
    const prof = profile || 'wh';
    const token = await _mintTokenWH();
    const res = await _whFetchTimeout(`${_SB_URL}/rest/v1/rpc/${fn}`, {
      method: 'POST',
      headers: {
        'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token,
        'Accept-Profile': prof, 'Content-Profile': prof, 'Content-Type': 'application/json'
      },
      body: JSON.stringify(args || {})
    }, ms || 12000);
    if (!res.ok) {
      // [400-loop fix] Distinguir un RECHAZO definitivo del servidor (HTTP 4xx:
      // request mal formado / función con firma que no existe / args inválidos)
      // de un fallo transitorio (timeout/red/5xx). Un 4xx NO commitea (PostgREST
      // corre la función en una sola transacción que hace rollback ante error) →
      // es SEGURO que la cola descarte el ítem en vez de reintentarlo para siempre.
      // 408 (timeout) y 429 (rate limit) son transitorios → NO se marcan permanentes.
      const e = new Error('rpc directo HTTP ' + res.status);
      e.status = res.status;
      e.permanente = (res.status >= 400 && res.status < 500 && res.status !== 408 && res.status !== 429);
      throw e;
    }
    return res.json();
  }

  // ── Transformador shape-hoja (portado FIEL de _sbRowsToObjsWH/_sbValToSheet de GAS) ──
  // Specs de LECTURA: subset de _WH_SPECS (las tablas que el navegador lee). [pg_col, claveShape, tipo].
  // Son los MISMOS specs que cuadraron al centavo en PASO 3 → mapeo idéntico a lo que GAS produce.
  const _WH_SPECS_LEC = {
    mermas: [['id_merma','idMerma','text'],['fecha_ingreso','fechaIngreso','date'],['origen','origen','text'],['cod_producto','codigoProducto','text'],['id_lote','idLote','text'],['cantidad_original','cantidadOriginal','num'],['cantidad_pendiente','cantidadPendiente','num'],['motivo','motivo','text'],['usuario','usuario','text'],['id_guia','idGuia','text'],['estado','estado','text'],['responsable','responsable','text'],['cantidad_reparada','cantidadReparada','num'],['cantidad_desechada','cantidadDesechada','num'],['foto','foto','text'],['fecha_resolucion','fechaResolucion','date'],['observacion_resolucion','observacionResolucion','text'],['id_guia_salida','idGuiaSalida','text']],
    auditorias: [['id_auditoria','idAuditoria','text'],['fecha_asignacion','fechaAsignacion','date'],['cod_producto','codigoProducto','text'],['usuario','usuario','text'],['stock_sistema','stockSistema','num'],['stock_fisico','stockFisico','num'],['diferencia','diferencia','num'],['resultado','resultado','text'],['observacion','observacion','text'],['estado','estado','text'],['fecha_ejecucion','fechaEjecucion','date']],
    ajustes: [['id_ajuste','idAjuste','text'],['cod_producto','codigoProducto','text'],['tipo_ajuste','tipoAjuste','text'],['cantidad_ajuste','cantidadAjuste','num'],['motivo','motivo','text'],['usuario','usuario','text'],['id_auditoria','idAuditoria','text'],['fecha','fecha','date']],
    producto_nuevo: [['id_producto_nuevo','idProductoNuevo','text'],['id_guia','idGuia','text'],['marca','marca','text'],['descripcion','descripcion','text'],['codigo_barra','codigoBarra','text'],['id_categoria','idCategoria','text'],['unidad','unidad','text'],['cantidad','cantidad','num'],['fecha_vencimiento','fechaVencimiento','date'],['foto','foto','text'],['estado','estado','text'],['usuario','usuario','text'],['fecha_registro','fechaRegistro','date'],['aprobado_por','aprobadoPor','text'],['fecha_aprobacion','fechaAprobacion','date'],['observacion','observacion','text']],
    preingresos: [['id_preingreso','idPreingreso','text'],['fecha','fecha','date'],['id_proveedor','idProveedor','text'],['cargadores','cargadores','text'],['usuario','usuario','text'],['monto','monto','num'],['fotos','fotos','text'],['comentario','comentario','text'],['estado','estado','text'],['id_guia','idGuia','text'],['snapshot_aviso','snapshotAviso','json']],
    guias: [['id_guia','idGuia','text'],['tipo','tipo','text'],['fecha','fecha','date'],['usuario','usuario','text'],['id_proveedor','idProveedor','text'],['id_zona','idZona','text'],['numero_documento','numeroDocumento','text'],['comentario','comentario','text'],['monto_total','montoTotal','num'],['estado','estado','text'],['id_preingreso','idPreingreso','text'],['foto','foto','text'],['ocr_estado','OCR_Estado','text'],['ocr_tipo','OCR_Tipo','text'],['ocr_ruc_emisor','OCR_RUC_Emisor','text'],['ocr_razon_social','OCR_Razon_Social','text'],['ocr_serie','OCR_Serie','text'],['ocr_numero','OCR_Numero','text'],['ocr_fecha_comprobante','OCR_Fecha_Comprobante','date'],['ocr_total','OCR_Total','num'],['ocr_subtotal','OCR_Subtotal','num'],['igv_recuperable','IGV_Recuperable','num'],['ocr_confidence','OCR_Confidence','num'],['ocr_notas','OCR_Notas','text'],['ocr_fecha_proceso','OCR_Fecha_Proceso','date']],
    lotes_vencimiento: [['id_lote','idLote','text'],['cod_producto','codigoProducto','text'],['fecha_vencimiento','fechaVencimiento','date'],['cantidad_inicial','cantidadInicial','num'],['cantidad_actual','cantidadActual','num'],['id_guia','idGuia','text'],['estado','estado','text'],['fecha_creacion','fechaCreacion','date']],
    stock_movimientos: [['id_mov','idMov','text'],['fecha','fecha','date'],['cod_producto','codigoProducto','text'],['delta','delta','num'],['stock_antes','stockAntes','num'],['stock_despues','stockDespues','num'],['tipo_operacion','tipoOperacion','text'],['origen','origen','text'],['usuario','usuario','text']],
    pickups: [['id_pickup','idPickup','text'],['fuente','fuente','text'],['estado','estado','text'],['items','items','json'],['id_zona','idZona','text'],['notas','notas','text'],['creado_por','creadoPor','text'],['fecha_creado','fechaCreado','ts'],['fecha_atendido','fechaAtendido','date'],['atendido_por','atendidoPor','text'],['ultima_actividad','ultimaActividad','ts']],
    guia_detalle: [['id_guia','idGuia','text'],['linea','linea','int'],['cod_producto','codigoProducto','text'],['cant_esperada','cantidadEsperada','num'],['cant_recibida','cantidadRecibida','num'],['precio_unitario','precioUnitario','num'],['id_lote','idLote','text'],['observacion','observacion','text'],['id_producto_nuevo','idProductoNuevo','text'],['id_detalle','idDetalle','text'],['fecha_vencimiento','fechaVencimiento','date']],
    envasados: [['id_envasado','idEnvasado','text'],['cod_producto_base','codigoProductoBase','text'],['cantidad_base','cantidadBase','num'],['unidad_base','unidadBase','text'],['cod_producto_envasado','codigoProductoEnvasado','text'],['unidades_esperadas','unidadesEsperadas','num'],['unidades_producidas','unidadesProducidas','num'],['merma_real','mermaReal','num'],['eficiencia_pct','eficienciaPct','num'],['fecha','fecha','date'],['usuario','usuario','text'],['estado','estado','text'],['id_guia_salida','idGuiaSalida','text'],['id_guia_ingreso','idGuiaIngreso','text'],['observacion','observacion','text'],['colaborador','colaborador','text']]
  };
  // Resolver de descripciones para ENVASADOS — RÉPLICA FIEL de getEnvasados (GAS): mapPorCb>mapPorSku>mapPorId, fallback ''.
  // (distinto de _prodMapWH: acá skuBase indexa TODOS sin filtro de factor/estado, y el fallback es '' no el código.)
  function _resolverDescEnvasado() {
    const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
    const cb = {}, sku = {}, id = {};
    prods.forEach(p => {
      const d = String(p.descripcion || p.nombre || p.idProducto || '');
      if (!d) return;
      if (p.codigoBarra) cb[String(p.codigoBarra)] = d;
      if (p.skuBase)     sku[String(p.skuBase)] = d;
      if (p.idProducto)  id[String(p.idProducto)] = d;
    });
    return (codigo) => { const k = String(codigo || ''); if (!k) return ''; return cb[k] || sku[k] || id[k] || ''; };
  }
  // Pendientes de envasar — RÉPLICA FIEL de _calcularPendientesEnvasado (GAS). productos=catálogo, stockMap={codigoProducto:cantidad}.
  function _calcularPendientesEnvasadoFront(productos, stockMap) {
    const pendientes = [];
    (productos || []).forEach(p => {
      if (!p.codigoProductoBase || p.codigoProductoBase === '') return;
      if (p.estado !== '1') return;
      const esWH = String(p.codigoBarra || p.idProducto || '').toUpperCase().indexOf('WH') === 0;
      if (!esWH) return;
      const stockDerivado = stockMap[p.codigoBarra] !== undefined ? stockMap[p.codigoBarra] : (stockMap[p.idProducto] || 0);
      const minDerivado = parseFloat(p.stockMinimo) || 0;
      const maxDerivado = parseFloat(p.stockMaximo) || 0;
      let estaPendiente;
      if (maxDerivado > 0) estaPendiente = stockDerivado < maxDerivado;
      else if (minDerivado > 0) estaPendiente = stockDerivado < minDerivado;
      else estaPendiente = stockDerivado <= 0;
      if (!estaPendiente) return;
      const stockBase = stockMap[p.codigoProductoBase] || 0;
      const factor = parseFloat(p.factorConversionBase) || parseFloat(p.factorConversion) || 1;
      const merma = parseFloat(p.mermaEsperadaPct) || 0;
      const maxProducibles = Math.floor(stockBase * factor * (1 - merma / 100));
      const necesita = maxDerivado > 0 ? Math.max(0, maxDerivado - stockDerivado) : (minDerivado > 0 ? Math.max(0, minDerivado - stockDerivado) : Math.max(1, -stockDerivado));
      pendientes.push({
        codigoDerivado: p.idProducto, descripcion: p.descripcion, codigoBase: p.codigoProductoBase,
        stockDerivado, stockMinimo: minDerivado, stockMaximo: maxDerivado, necesitaProducir: necesita,
        stockBase, factorConversionBase: factor, mermaEsperadaPct: merma, maxProducibles,
        granelNecesario: maxDerivado > 0 ? Math.ceil(necesita / factor) : null,
        urgencia: stockDerivado === 0 ? 'CRITICA' : (stockDerivado < minDerivado ? 'ALTA' : 'MEDIA')
      });
    });
    pendientes.sort((a, b) => {
      if (a.urgencia === 'CRITICA' && b.urgencia !== 'CRITICA') return -1;
      if (b.urgencia === 'CRITICA' && a.urgencia !== 'CRITICA') return 1;
      return b.necesitaProducir - a.necesitaProducir;
    });
    return pendientes;
  }
  // 'yyyy-MM-dd HH:mm:ss' en TZ Lima (= Utilities.formatDate(now, scriptTZ, ...) de GAS).
  function _ahoraLima() {
    const p = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(new Date());
    const g = t => (p.find(x => x.type === t) || {}).value || '';
    return `${g('year')}-${g('month')}-${g('day')} ${g('hour')}:${g('minute')}:${g('second')}`;
  }
  // Dashboard — RÉPLICA FIEL de getDashboard (GAS). Días-calendario en TZ Lima; guías>48h por timestamp (horas). Config 30/7 default.
  function _buildDashboardFront(productos, stockMap, lotes, guias, preingresos, mermas, auditorias, envasados) {
    const cfg = (typeof OfflineManager !== 'undefined' && OfflineManager.getConfigCache) ? (OfflineManager.getConfigCache() || {}) : {};
    const diasAlerta = parseInt(cfg.DIAS_ALERTA_VENC) || 30;
    const diasCrit   = parseInt(cfg.DIAS_ALERTA_VENC_CRITICO) || 7;
    const hoyMs = _diaLimaMs(new Date());
    const limAlertaMs = hoyMs + diasAlerta * _UN_DIA;
    const limCritMs   = hoyMs + diasCrit * _UN_DIA;
    // stockMap: indexar también por skuBase (= getDashboard)
    productos.forEach(p => {
      if (!p.skuBase || p.skuBase === '') return;
      if (stockMap[p.skuBase] !== undefined) return;
      const val = stockMap[p.idProducto] || stockMap[p.codigoBarra] || 0;
      if (val > 0) stockMap[p.skuBase] = val;
    });
    const productosMap = {}; productos.forEach(p => { productosMap[p.idProducto] = p; });
    // 1. vencimientos
    const vencCriticos = [], vencAlertas = [];
    (lotes || []).forEach(l => {
      if (l.estado !== 'ACTIVO') return;
      const cant = parseFloat(l.cantidadActual) || 0;
      if (cant <= 0) return;
      if (!l.fechaVencimiento) return;
      const fvMs = _diaLimaMs(l.fechaVencimiento);
      const prod = productosMap[l.codigoProducto] || {};
      const entry = { idLote: l.idLote, codigoProducto: l.codigoProducto, descripcion: prod.descripcion || l.codigoProducto, fechaVencimiento: l.fechaVencimiento, cantidadActual: cant, diasRestantes: Math.ceil((fvMs - hoyMs) / _UN_DIA) };
      if (fvMs <= limCritMs) vencCriticos.push(entry);
      else if (fvMs <= limAlertaMs) vencAlertas.push(entry);
    });
    vencCriticos.sort((a, b) => a.diasRestantes - b.diasRestantes);
    vencAlertas.sort((a, b) => a.diasRestantes - b.diasRestantes);
    // 2. stock bajo mínimo
    const bajominimo = [];
    productos.forEach(p => {
      if (p.estado !== '1') return;
      const actual = stockMap[p.idProducto] || 0;
      const minimo = parseFloat(p.stockMinimo) || 0;
      if (actual < minimo) bajominimo.push({ codigo: p.idProducto, descripcion: p.descripcion, stockActual: actual, stockMinimo: minimo, diferencia: minimo - actual });
    });
    bajominimo.sort((a, b) => (b.diferencia / b.stockMinimo) - (a.diferencia / a.stockMinimo));
    // 3. pendientes envasado
    const pendientesEnvasado = _calcularPendientesEnvasadoFront(productos, stockMap);
    // 4. preingresos pendientes
    const presPendientes = (preingresos || []).filter(p => p.estado === 'PENDIENTE');
    // 5. guías abiertas > 48h (timestamp, horas)
    const hace48Ms = Date.now() - 48 * 60 * 60 * 1000;
    const guiasAbiertas = (guias || []).filter(g => g.estado === 'ABIERTA' && new Date(g.fecha).getTime() < hace48Ms);
    // 6/7. mermas/auditorías pendientes
    // [524] canon por CANTIDADES: los strings legacy (PENDIENTE/PROCESADA) ya no existen tras la normalización
    const mermasPendientes = (mermas || []).filter(m => (parseFloat(m.cantidadPendiente) || 0) > 0);
    const audPendientes = (auditorias || []).filter(a => a.estado === 'PENDIENTE' || a.estado === 'ASIGNADA');
    // 8. KPIs (mes = últimos 30 días-calendario Lima)
    const hace30Ms = hoyMs - 30 * _UN_DIA;
    const mermasMes = (mermas || []).filter(m => _diaLimaMs(m.fechaIngreso) >= hace30Ms).reduce((acc, m) => acc + (parseFloat(m.cantidadOriginal) || 0), 0);
    const envasadosMes = (envasados || []).filter(e => _diaLimaMs(e.fecha) >= hace30Ms && e.estado === 'COMPLETADO');
    let eficienciaPromedio = 0;
    if (envasadosMes.length > 0) { const s = envasadosMes.reduce((acc, e) => acc + (parseFloat(e.eficienciaPct) || 0), 0); eficienciaPromedio = Math.round(s / envasadosMes.length * 10) / 10; }
    const salidasMes = (guias || []).filter(g => _diaLimaMs(g.fecha) >= hace30Ms && String(g.tipo || '').startsWith('SALIDA')).length;
    // 9. totales
    const totalProductos = productos.filter(p => p.estado === '1').length;
    const totalBases = productos.filter(p => p.esEnvasable === '1' && p.estado === '1').length;
    const totalDerivados = productos.filter(p => p.codigoProductoBase !== '' && p.estado === '1').length;
    return {
      alertas: { vencimientosCriticos: vencCriticos, vencimientosAlertas: vencAlertas, stockBajoMinimo: bajominimo, pendientesEnvasado: pendientesEnvasado, preingresosPendientes: presPendientes, guiasAbiertasTardias: guiasAbiertas, mermasPendientes: mermasPendientes, auditoriasPendientes: audPendientes },
      kpis: { totalProductosActivos: totalProductos, totalProductosBase: totalBases, totalProductosDerivados: totalDerivados, mermasTotalMes: mermasMes, eficienciaEnvasadoPct: eficienciaPromedio, salidasUltimos30dias: salidasMes, lotesCriticos: vencCriticos.length, lotesEnAlerta: vencAlertas.length },
      contadores: { alertasTotal: vencCriticos.length + bajominimo.length + pendientesEnvasado.length + presPendientes.length + mermasPendientes.length, criticos: vencCriticos.length + bajominimo.filter(b => b.stockActual === 0).length },
      generadoEn: _ahoraLima()
    };
  }
  // prodMap para enriquecer descripciones — RÉPLICA FIEL de getGuia (GAS): indexa por codigoBarra, idProducto, y
  // skuBase (solo producto BASE: factor=1 y activo), + equivalencias (codigoBarra→nombre del skuBase base). Usa cache LOCAL.
  function _prodMapWH() {
    const map = {};
    const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
    prods.forEach(p => {
      const name = p.descripcion || p.nombre || '';
      if (!name) return;
      if (p.codigoBarra) map[String(p.codigoBarra)] = name;
      if (p.idProducto)  map[String(p.idProducto)] = name;
      const esBase = parseFloat(p.factorConversion || 1) === 1 && String(p.estado || '') !== '0';
      if (esBase && p.skuBase) map[String(p.skuBase).trim().toUpperCase()] = name;
    });
    const equivs = (typeof OfflineManager !== 'undefined' && OfflineManager.getEquivalenciasCache) ? (OfflineManager.getEquivalenciasCache() || []) : [];
    equivs.forEach(e => {
      if (!e.codigoBarra || !e.skuBase) return;
      const name = map[String(e.skuBase).trim().toUpperCase()];
      if (name) map[String(e.codigoBarra)] = name;
    });
    return map;
  }
  // yyyy-MM-dd en TZ Lima (en-CA da ese formato). Idéntico a Utilities.formatDate(d,'America/Lima','yyyy-MM-dd').
  function _fmtFechaLima(v) {
    const d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return '';
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
  }
  // [Fechas mismo idioma · TZ Lima] Medianoche (00:00 Lima) del día de un valor fecha, en ms epoch.
  // Perú = UTC-5 fijo (sin DST). Permite comparar/restar DÍAS-CALENDARIO en Lima sin importar la TZ del
  // dispositivo → diasRestantes/diasEnProceso consistentes en todo el ecosistema. NaN si la fecha es inválida.
  const _LIMA_OFF = '-05:00';
  function _diaLimaMs(v) {
    const ymd = (v == null || v === '') ? '' : _fmtFechaLima(v);
    if (!ymd) return NaN;
    return new Date(ymd + 'T00:00:00' + _LIMA_OFF).getTime();
  }
  const _UN_DIA = 24 * 60 * 60 * 1000;
  // Inverso de _sbValToSheet (GAS) — MISMA lógica: null→'', num→Number|'', date→yyyy-MM-dd|'', json→string, text→String.
  function _sbValFront(v, t) {
    // bool10 ANTES del null-check: bool de catálogo → '1'/'0' (el front compara estado!=='0', esEnvasable==='1'). null/false→'0'.
    if (t === 'bool10') return (v === true || v === 't' || v === 'true' || v === 1 || v === '1') ? '1' : '0';
    if (v == null) return '';
    if (t === 'num' || t === 'int') { const n = (typeof v === 'number') ? v : parseFloat(v); return isNaN(n) ? '' : n; }
    if (t === 'ts') return (v instanceof Date) ? v.toISOString() : String(v);   // [fix] timestamp COMPLETO (con hora). 'date' (date-only) hacía que el "hace" diera "ayer" de noche.
    if (t === 'date') return _fmtFechaLima(v);
    if (t === 'json') return (typeof v === 'object') ? JSON.stringify(v) : String(v);
    return String(v);
  }
  function _sbRowsToObjsFront(tabla, rows) {
    const spec = _WH_SPECS_LEC[tabla]; if (!spec) return [];
    return (rows || []).map(row => {
      const o = {};
      for (let s = 0; s < spec.length; s++) o[spec[s][1]] = _sbValFront(row[spec[s][0]], spec[s][2]);
      return o;
    });
  }
  // Lee una tabla wh.* completa directo (RPC genérica leer_tabla_rls, 1 request, sin límite) → shape-hoja.
  // [perf v2.13.242] Dedup + micro-cache 4s: si getDashboard y descargarOperacional
  // piden 'guias' en la misma ráfaga, se hace UNA sola leer_tabla_rls compartida.
  // 4s << ciclo de refresh (60s) → no afecta frescura percibida.
  function _sbLeerTablaWH(tabla, force) {
    // [v2.13.372] force=true → invalida el micro-caché (4s) antes de leer, para lecturas que
    // necesitan frescura inmediata (poll de pickups + realtime), sin afectar la ráfaga del dashboard.
    if (force) { try { _readCache.delete('tabla:' + tabla); } catch (_) {} }
    return _dedupRead('tabla:' + tabla, 4000, async () => {
      const out = await _sbRpcWH('leer_tabla_rls', { p_tabla: tabla });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'leer_tabla error');
      return _sbRowsToObjsFront(tabla, out.data);
    });
  }

  // [CARGA INTELIGENTE Guías] Detalle operacional FILTRADO server-side (wh.guia_detalle_operacional):
  // SOLO líneas de guías ABIERTA (cualquier edad → la proyección las necesita) + guías de los últimos
  // _GD_DIAS días (cubren rotación 30d, último-mov, conteo de líneas de card y el detalle instantáneo
  // al abrir una guía reciente). El detalle de guías CERRADAS viejas NO se precarga: verDetalle cae a
  // getGuia (get_guia_rls per-guía) al abrirla; el Historial real sale de stock_movimientos
  // (getHistorialStock, autoritativo); el chat usa getGuia per-guía. Acota el payload global (que crecía
  // sin techo con el histórico) a una ventana rodante. Mismo shape que leer_tabla_rls('guia_detalle')
  // → _sbRowsToObjsFront('guia_detalle', ...) sin cambios. Dedup 4s como el resto.
  const _GD_DIAS = 30;   // [perf] 60→30: alineado con el límite de 30d de la lista (todo lo mostrado tiene su
                         // detalle; las ABIERTAS de cualquier edad siguen incluidas) → ~mitad del payload del refresco.
  function _sbGuiaDetalleOperacional() {
    return _dedupRead('guia_detalle_op:' + _GD_DIAS, 4000, async () => {
      const out = await _sbRpcWH('guia_detalle_operacional', { p_dias: _GD_DIAS });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'guia_detalle_op error');
      return _sbRowsToObjsFront('guia_detalle', out.data);
    });
  }

  // [perf v2.13.242] stock_enriquecido_rls (solo_alertas:false) es el RPC MÁS pesado
  // (~48KB) y lo piden a la vez getStock, getDashboard, descargarOperacional y
  // ProductosView._refrescarStockVivo (hasta 3x al entrar a Productos). Dedup+cache
  // 4s colapsa esa ráfaga en UN solo round-trip compartido. Devuelve la respuesta
  // CRUDA del RPC {ok,data:[...]} (igual que _sbRpcWH) para no cambiar a los callers.
  function _sbStockEnriquecidoFull() {
    return _dedupRead('stock_full', 4000, async () => {
      return await _sbRpcWH('stock_enriquecido_rls', { solo_alertas: false });
    });
  }

  // [BUG A · cutover] descargarOperacional DIRECTO a Supabase. Cierra el agujero del cutover:
  // un dispositivo con ESCRITURA directa crea guías 'G_L...' en Supabase, pero el cache operacional
  // (que alimenta el listado de Guías) seguía leyéndose por GAS (descargarOperacional). Si el GAS no
  // espejaba esas guías directas, NUNCA aparecían en el listado (ni tras F5). Acá traemos el operacional
  // directo de Supabase reusando los MISMOS lectores ya probados (_sbLeerTablaWH + stock_enriquecido_rls),
  // con el shape EXACTO que devuelve descargarOperacional de GAS: { ok, data:{ guias, detalles,
  // preingresos, stock, ajustes, auditorias } }. Cualquier fallo LANZA → el llamador cae a GAS (seguro).
  async function _descargarOperacionalDirecto() {
    const [guias, detalles, preingresos, ajustes, auditorias, stockR] = await Promise.all([
      _sbLeerTablaWH('guias'),
      _sbGuiaDetalleOperacional(),  // [CARGA INTELIGENTE] detalle filtrado (abiertas + últimos _GD_DIAS), NO el global
      _sbLeerTablaWH('preingresos'),
      _sbLeerTablaWH('ajustes'),
      _sbLeerTablaWH('auditorias'),
      _sbStockEnriquecidoFull()
    ]);
    // stock: mismo shape vivo que API.getStock (stock_enriquecido_rls) — gana al Sheet congelado.
    // [40x] LANZA si el RPC de stock fallo (igual que _sbLeerTablaWH lanza para las otras 5) -> el llamador
    // cae a GAS. NUNCA devolver stock=[] con ok:true: pisaria el cache bueno con vacio.
    if (!stockR || stockR.ok === false || !Array.isArray(stockR.data)) {
      throw new Error('stock_enriquecido_rls fallo: ' + ((stockR && stockR.error) || 'shape invalido'));
    }
    const stock = stockR.data;
    return { ok: true, data: { guias, detalles, preingresos, stock, ajustes, auditorias } };
  }

  // ── CATÁLOGO directo (mos.catalogo_wh_rls) — reemplaza descargarMaestros. Specs invertidos de _CAT_SPECS (MigracionCatalogo.gs).
  // bools de catálogo → 'bool10' (el front compara estado!=='0', esEnvasable==='1'). Sin pin/pin_hash/numero_cuenta/cci (la RPC ya los excluye).
  // adminPin NO viene acá (va por mos.verificar_clave_admin, F2). offline-first: el USO por-operación es contra el cache local.
  const _CAT_SPECS_LEC = {
    productos: [['id_producto','idProducto','text'],['sku_base','skuBase','text'],['codigo_barra','codigoBarra','text'],['descripcion','descripcion','text'],['marca','marca','text'],['id_categoria','idCategoria','text'],['unidad','unidad','text'],['precio_venta','precioVenta','num'],['precio_costo','precioCosto','num'],['cod_tributo','Cod_Tributo','text'],['igv_porcentaje','IGV_Porcentaje','num'],['cod_sunat','Cod_SUNAT','text'],['tipo_igv','Tipo_IGV','int'],['unidad_medida','Unidad_Medida','text'],['estado','estado','bool10'],['es_envasable','esEnvasable','bool10'],['codigo_producto_base','codigoProductoBase','text'],['factor_conversion','factorConversion','num'],['factor_conversion_base','factorConversionBase','num'],['merma_esperada_pct','mermaEsperadaPct','num'],['stock_minimo','stockMinimo','num'],['stock_maximo','stockMaximo','num'],['zona','zona','text'],['fecha_creacion','fechaCreacion','date'],['creado_por','creadoPor','text'],['modo_venta','modoVenta','text'],['margen_pct','margenPct','num'],['precio_tope','precioTope','num'],['foto_url','fotoUrl','text'],['historial_cambios','historialCambios','json'],['segmentos_precio','segmentos_precio','json'],['tipo_producto','tipoProducto','text']],
    equivalencias: [['id_equiv','idEquiv','text'],['sku_base','skuBase','text'],['codigo_barra','codigoBarra','text'],['descripcion','descripcion','text'],['activo','activo','bool10']],
    proveedores: [['id_proveedor','idProveedor','text'],['nombre','nombre','text'],['ruc','ruc','text'],['imagen','imagen','text'],['telefono','telefono','text'],['banco','banco','text'],['email','email','text'],['dia_pedido','diaPedido','text'],['dia_pago','diaPago','text'],['dia_entrega','diaEntrega','text'],['forma_pago','formaPago','text'],['plazo_credito','plazoCredito','text'],['responsable','responsable','text'],['categoria_producto','categoriaProducto','text'],['estado','estado','text']],
    personal: [['id_personal','idPersonal','text'],['nombre','nombre','text'],['apellido','apellido','text'],['tipo','tipo','text'],['app_origen','appOrigen','text'],['rol','rol','text'],['color','color','text'],['tarifa_hora','tarifaHora','num'],['monto_base','montoBase','num'],['estado','estado','bool10'],['fecha_ingreso','fechaIngreso','date'],['foto','foto','text'],['ultima_conexion','Ultima_Conexion','date']],
    impresoras: [['id_impresora','idImpresora','text'],['nombre','nombre','text'],['printnode_id','printNodeId','text'],['tipo','tipo','text'],['id_estacion','idEstacion','text'],['id_zona','idZona','text'],['app_origen','appOrigen','text'],['activo','activo','bool10'],['descripcion','descripcion','text']],
    zonas: [['id_zona','idZona','text'],['nombre','nombre','text'],['descripcion','descripcion','text'],['direccion','direccion','text'],['responsable','responsable','text'],['estado','estado','bool10'],['politica_json','politicaJSON','json']]
  };
  function _mapCat(tabla, rows) {
    const spec = _CAT_SPECS_LEC[tabla]; if (!spec) return [];
    return (rows || []).map(row => { const o = {}; for (let s = 0; s < spec.length; s++) o[spec[s][1]] = _sbValFront(row[spec[s][0]], spec[s][2]); return o; });
  }
  // Descarga maestros directo de Supabase (catálogo). Mismo shape que descargarMaestros de GAS (sin adminPin → F2).
  async function _sbDescargarMaestros() {
    const out = await _sbRpcWH('catalogo_wh_rls', {}, 'mos', 25000);   // [perf 500x] 25s: el catálogo es ~1.9MB
    if (!out || out.ok === false) throw new Error((out && out.error) || 'catalogo error');
    return { ok: true, server_ts: out.server_ts || null, data: {
      productos:     _mapCat('productos',     out.productos),
      equivalencias: _mapCat('equivalencias', out.equivalencias),
      proveedores:   _mapCat('proveedores',   out.proveedores),
      personal:      _mapCat('personal',      out.personal),
      impresoras:    _mapCat('impresoras',    out.impresoras),
      zonas:         _mapCat('zonas',         out.zonas)
    } };
  }
  // [perf 500x · CD1 DELTA] Descarga INCREMENTAL: solo productos cambiados desde `desdeTs` (+ tablas chicas
  // completas). ~KB en vez de ~1.9MB por bump de versión. El caller (offline.js) mergea los productos por id.
  async function _sbDescargarMaestrosDelta(desdeTs) {
    const out = await _sbRpcWH('catalogo_wh_delta', { p: { desde: desdeTs } }, 'mos', 25000);
    if (!out || out.ok === false) throw new Error((out && out.error) || 'catalogo_delta error');
    return { ok: true, delta: true, server_ts: out.server_ts || null, productos_cambiados: out.productos_cambiados || 0, data: {
      productos:     _mapCat('productos',     out.productos),   // SOLO los cambiados
      equivalencias: _mapCat('equivalencias', out.equivalencias),
      proveedores:   _mapCat('proveedores',   out.proveedores),
      personal:      _mapCat('personal',      out.personal),
      impresoras:    _mapCat('impresoras',    out.impresoras),
      zonas:         _mapCat('zonas',         out.zonas)
    } };
  }

  // ── Versión del catálogo (mos.catalogo_version) — poller barato de cambios del maestro ──
  // El trigger en mos.productos/mos.equivalencias incrementa esa versión (bigint) cada vez que
  // cambia el catálogo. WH la sondea (1 query liviana) para decidir CUÁNDO re-descargar el maestro
  // completo, en vez de re-bajarlo a ciegas. Se llama con profile 'mos' (igual que catalogo_wh_rls).
  // Devuelve un Number monotónico (>=0). LANZA ante fallo → el llamador (poller) lo trata como
  // "no pude leer la versión" y NO toca el baseline (no re-descarga por las dudas).
  async function _catalogoVersion() {
    const out = await _sbRpcWH('catalogo_version', {}, 'mos');
    if (!out || out.ok === false) throw new Error((out && out.error) || 'catalogo_version error');
    const v = Number(out.version);
    if (!Number.isFinite(v)) throw new Error('catalogo_version: version inválida');
    return v;
  }

  // Mapea una acción de lectura → su RPC directa. Devuelve la respuesta {ok,data}
  // o null si la acción NO tiene backend RLS listo (→ el llamador cae a GAS).
  async function _callDirecto(params) {
    const action = params.action;
    if (action === 'descargarMaestros') return await _sbDescargarMaestros();  // catálogo directo (mos.catalogo_wh_rls)
    if (action === 'getDashboard') {
      // Agregado de 8 datasets + KPIs. Trae todo en paralelo + cálculo réplica fiel de getDashboard (GAS).
      const [seR, lotes, guias, preingresos, mermas, auditorias, envasados] = await Promise.all([
        _sbStockEnriquecidoFull(),
        _sbLeerTablaWH('lotes_vencimiento'), _sbLeerTablaWH('guias'), _sbLeerTablaWH('preingresos'),
        _sbLeerTablaWH('mermas'), _sbLeerTablaWH('auditorias'), _sbLeerTablaWH('envasados')
      ]);
      if (!seR || seR.ok === false) throw new Error((seR && seR.error) || 'stock error');
      const stockMap = {};
      (seR.data || []).forEach(s => { stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });
      const productos = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      return { ok: true, data: _buildDashboardFront(productos, stockMap, lotes, guias, preingresos, mermas, auditorias, envasados) };
    }
    if (action === 'getStock') {
      // [perf v2.13.242] El full-stock (solo_alertas:false) pasa por el dedup compartido
      // (ProductosView lo pide hasta 3x al entrar + descargarOperacional + getDashboard
      // lo piden en la misma ráfaga). solo_alertas:true es chico y específico → directo.
      const soloAlertas = String(params.soloAlertas) === 'true';
      const out = soloAlertas
        ? await _sbRpcWH('stock_enriquecido_rls', { solo_alertas: true })
        : await _sbStockEnriquecidoFull();
      if (!out || out.ok === false) throw new Error((out && out.error) || 'rpc stock error');
      return out;  // {ok:true, data:[...]} — mismo shape que getStock por GAS
    }
    if (action === 'getRotacionSemanal') {
      const out = await _sbRpcWH('rotacion_semanal_rls', { semanas: Number(params.semanas) || 8, codigos_producto: params.codigos || null });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'rpc rotacion error');
      return out;  // {ok:true, data:{etiquetas,semanas,productos}}
    }
    if (action === 'getStockProducto') {
      // Stock EN VIVO (RPC dedicada) + enriquecer con catálogo cache. Réplica fiel de getStockProducto (GAS).
      const codigo = String(params.codigo || '');
      const out = await _sbRpcWH('stock_producto_rls', { p_cod: codigo });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'stock_producto error');
      const cantidad = parseFloat(out.cantidad) || 0;
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const prod = prods.find(p => p.idProducto === codigo) || {};
      return { ok: true, data: {
        codigo, descripcion: prod.descripcion || codigo, cantidad, unidad: prod.unidad || '',
        stockMinimo: prod.stockMinimo || 0, alerta: cantidad < (parseFloat(prod.stockMinimo) || 0)
      } };
    }
    if (action === 'getStockProyectado') {
      // [Stock teórico/proyectado] Overlay DERIVADO al vuelo (NO persistido): por producto con guías ABIERTAS,
      // proyectado = real + Σ(ingresos abiertos) − Σ(salidas abiertas). El real NO se toca (se aplica al cerrar).
      // La RPC ya enriquece (descripcion/min/max/unidad). Solo devuelve productos CON movimiento pendiente.
      const out = await _sbRpcWH('stock_proyectado_rls', {});
      if (!out || out.ok === false) throw new Error((out && out.error) || 'rpc proyectado error');
      return out;  // {ok:true, data:[{codigoProducto,cantidadDisponible,porRecibir,porSalir,proyectado,...}]}
    }
    // [Grupo A · asimetría] getAlertasStock: las escrituras (marcar/aceptar) ya son directas, pero la LECTURA
    // iba a GAS (Hoja). RPC dedicada wh.get_alertas_stock (mismo shape {ok,data}). Cierra la asimetría.
    if (action === 'getAlertasStock') {
      return await _sbRpcWH('get_alertas_stock', { p: { soloPendientes: !!params.soloPendientes } }, 'wh');
    }
    // [Grupo A] getHistorialLote: trazabilidad de lote desde Supabase (RPC wh.get_historial_lote) en vez de GAS.
    if (action === 'getHistorialLote') {
      return await _sbRpcWH('get_historial_lote', { p: { idLote: params.idLote || '', codigos: params.codigos } }, 'wh');
    }
    // [Frente 4] getConfig desde Supabase (wh.get_config, filtra secretos). Reemplaza la lectura de la Hoja CONFIG.
    if (action === 'getConfig') {
      return await _sbRpcWH('get_config', {}, 'wh');
    }
    // [Frente 1 · sesión] getSesionActiva desde Supabase (wh.get_sesion_activa). Login/horario sigue GAS (aparte).
    if (action === 'getSesionActiva') {
      return await _sbRpcWH('get_sesion_activa', { p: { idSesion: params.idSesion } }, 'wh');
    }
    // [Frente 4] getHistorialGuia desde Supabase (wh.historial_guia, read-only). Compone 6 fuentes (guias/ops_log/
    // stock_mov/producto_nuevo/mermas) + gating admin/master server-side. APP_NO_AUTORIZADA (misconfig) → GAS;
    // la denegación legítima (no admin) y "no encontrada" SÍ se devuelven (GAS daría lo mismo).
    if (action === 'getHistorialGuia') {
      const out = await _sbRpcWH('historial_guia', { p: {
        idGuia: params.idGuia || '', idPersonal: params.idPersonal || '',
        usuario: params.usuario || '', claveAdmin: params.claveAdmin || ''
      } }, 'wh');
      if (!out || (out.ok === false && String(out.error || '') === 'APP_NO_AUTORIZADA')) return null;
      return out;
    }
    // [Frente 4] getImpresorasEcosistema → Edge `printers` {op:'list'} (PrintNode), no GAS. Read-only.
    if (action === 'getImpresorasEcosistema') {
      try {
        const token = await _mintTokenWH();
        const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/printers`, {
          method: 'POST', headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ op: 'list' })
        }, 20000);
        const d = await res.json().catch(() => null);
        if (res.ok && d && d.ok === true && Array.isArray(d.data)) return { ok: true, data: d.data };
      } catch (_) { /* → GAS */ }
      return null;
    }
    // [Frente 4 · cargadores] getCargadoresDelDia 100% directo: lee wh.preingresos + cache proveedores y
    // consolida cargadores del día (carretas/llenas/medias/vacías) — réplica fiel de Reporte.gs. El spec ya
    // convirtió preingresos.fecha a yyyy-MM-dd Lima → comparo directo (NO re-convertir: 'yyyy-MM-dd' parseado
    // como Date sería UTC y restaría un día). Read-only.
    if (action === 'getCargadoresDelDia') {
      const fechaStr = (String(params.fecha || '').trim().substring(0, 10)) || _fmtFechaLima(new Date());
      const rows = await _sbLeerTablaWH('preingresos');
      const provMap = {};
      const provs = (typeof OfflineManager !== 'undefined' && OfflineManager.getProveedoresCache) ? (OfflineManager.getProveedoresCache() || []) : [];
      provs.forEach(p => { if (p && p.idProveedor != null) provMap[String(p.idProveedor)] = String(p.nombre || ''); });
      const delDia = (rows || []).filter(pi => pi.fecha && String(pi.fecha).substring(0, 10) === fechaStr);
      const _normEstados = (c) => {
        const n = parseInt(c.carretas) || 1;
        let arr = [];
        if (Array.isArray(c.estados)) arr = c.estados.slice(0, n).map(e => (e === 'MEDIA' || e === 'VACIA') ? e : 'LLENA');
        while (arr.length < n) arr.push('LLENA');
        return arr;
      };
      const byId = {};
      delDia.forEach(pi => {
        let cargs = [];
        try { cargs = JSON.parse(pi.cargadores || '[]'); } catch (_) {}
        if (!Array.isArray(cargs)) return;
        cargs.forEach(c => {
          if (!c || typeof c !== 'object') return;
          const id     = String(c.id || c.idPersonal || c.nombre || '');
          const nombre = String(c.nombre || c.idPersonal || id || '');
          if (!id || !nombre) return;
          const carretas = parseInt(c.carretas) || 0;
          const estados  = _normEstados(c);
          let llenas = 0, medias = 0, vacias = 0;
          estados.forEach(e => { if (e === 'LLENA') llenas++; else if (e === 'MEDIA') medias++; else if (e === 'VACIA') vacias++; });
          if (!byId[id]) byId[id] = { id, nombre, carretasTotal: 0, llenasTotal: 0, mediasTotal: 0, vaciasTotal: 0, preingresos: [] };
          byId[id].carretasTotal += carretas; byId[id].llenasTotal += llenas; byId[id].mediasTotal += medias; byId[id].vaciasTotal += vacias;
          byId[id].preingresos.push({
            idPreingreso: pi.idPreingreso, proveedor: provMap[String(pi.idProveedor)] || pi.idProveedor || '',
            carretas, estados, llenas, medias, vacias, estado: String(pi.estado || '')
          });
        });
      });
      const cargadores = Object.keys(byId).map(k => byId[k]).sort((a, b) => b.carretasTotal - a.carretasTotal);
      return { ok: true, data: {
        fecha: fechaStr, cargadores,
        totalCarretas: cargadores.reduce((s, c) => s + c.carretasTotal, 0),
        totalLlenas:   cargadores.reduce((s, c) => s + c.llenasTotal, 0),
        totalMedias:   cargadores.reduce((s, c) => s + c.mediasTotal, 0),
        totalVacias:   cargadores.reduce((s, c) => s + c.vaciasTotal, 0),
        preingresos:   delDia.length
      } };
    }
    // [Frente 4 · cargadores] getResumenCargadoresDia directo (wh.resumen_cargadores_dia). Cierra asimetría:
    // add/remove_cargador_dia YA escriben directo; el resumen leía la Hoja. Comparación de día en TZ Lima.
    if (action === 'getResumenCargadoresDia') {
      return await _sbRpcWH('resumen_cargadores_dia', { p: { fecha: params.fecha || '' } }, 'wh');
    }
    // [Frente 4 · cesta] getMermasCesta 100% directo: lee wh.mermas y agrupa (pendientes/descartado/
    // solucionado) — réplica fiel de getMermas.gs:getMermasCesta. El GAS ya leía la sombra Supabase;
    // esto saca el round-trip a GAS del navegador. Read-only.
    if (action === 'getMermasCesta' || action === 'contadorMermasPendientes') {
      const rows = await _sbLeerTablaWH('mermas');
      const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 7);
      const pendientes = [], descartado = [], solucionado = [];
      (rows || []).forEach(r => {
        const est = String(r.estado || '').toUpperCase();
        const pend = parseFloat(r.cantidadPendiente) || 0;
        const des  = parseFloat(r.cantidadDesechada) || 0;
        const rep  = parseFloat(r.cantidadReparada)  || 0;
        const idg  = String(r.idGuiaSalida || '');
        if (est === 'EN_PROCESO' && pend > 0) pendientes.push(r);
        else if (des > 0 && !idg && est !== 'ELIMINADO') descartado.push(r);
        else {
          const f = r.fechaResolucion ? new Date(r.fechaResolucion)
                  : r.fechaIngreso    ? new Date(r.fechaIngreso) : null;
          if (rep > 0 && f && f >= cutoff) solucionado.push(r);
        }
      });
      if (action === 'contadorMermasPendientes') {
        return { ok: true, data: { count: pendientes.length + descartado.length } };
      }
      const ordSc = (a, b) => new Date(b.fechaResolucion || b.fechaIngreso || 0) - new Date(a.fechaResolucion || a.fechaIngreso || 0);
      pendientes.sort(ordSc); descartado.sort(ordSc); solucionado.sort(ordSc);
      return { ok: true, data: {
        pendientes, descartado, solucionado,
        totalPendientes: pendientes.length, totalDescartado: descartado.length
      } };
    }
    // Lecturas SIMPLES (filas + filtros, sin lógica derivada) — filtros REPLICAN exacto el getXxx de GAS.
    if (action === 'getMermas') {
      let rows = await _sbLeerTablaWH('mermas');
      if (params.estado) rows = rows.filter(r => String(r.estado) === String(params.estado));
      if (params.codigo) rows = rows.filter(r => r.codigoProducto === params.codigo);
      if (params.limit)  rows = rows.slice(0, parseInt(params.limit));
      return { ok: true, data: rows };
    }
    if (action === 'getAuditorias') {
      let rows = await _sbLeerTablaWH('auditorias');
      if (params.estado)  rows = rows.filter(r => String(r.estado) === String(params.estado));
      if (params.usuario) rows = rows.filter(r => r.usuario === params.usuario);
      return { ok: true, data: rows };
    }
    if (action === 'getAjustes') {
      let rows = await _sbLeerTablaWH('ajustes');
      if (params.codigo) rows = rows.filter(r => r.codigoProducto === params.codigo);
      return { ok: true, data: rows };
    }
    if (action === 'getProductosNuevos') {
      let rows = await _sbLeerTablaWH('producto_nuevo');
      if (params.estado) rows = rows.filter(r => String(r.estado) === String(params.estado));
      return { ok: true, data: rows };
    }
    if (action === 'getPreingresos') {
      let rows = await _sbLeerTablaWH('preingresos');
      if (params.estado)      rows = rows.filter(r => String(r.estado) === String(params.estado));
      if (params.idProveedor) rows = rows.filter(r => r.idProveedor === params.idProveedor);
      return { ok: true, data: rows };
    }
    if (action === 'getEnvasados') {
      let rows = await _sbLeerTablaWH('envasados');
      if (params.estado)     rows = rows.filter(r => String(r.estado) === String(params.estado));
      if (params.fecha)      rows = rows.filter(r => r.fecha === params.fecha);
      if (params.fechaDesde) rows = rows.filter(r => String(r.fecha) >= String(params.fechaDesde));
      if (params.limit)      rows = rows.slice(0, parseInt(params.limit));
      const resolver = _resolverDescEnvasado();  // descripciones desde cache local (réplica fiel de GAS)
      rows = rows.map(r => { r.descripcionProductoEnvasado = resolver(r.codigoProductoEnvasado); r.descripcionProductoBase = resolver(r.codigoProductoBase); return r; });
      return { ok: true, data: rows };
    }
    if (action === 'getPendientesEnvasado') {
      // stock EN VIVO (stock_enriquecido) → stockMap {codigoProducto:cantidadDisponible} + productos cache → cálculo réplica fiel.
      const se = await _sbStockEnriquecidoFull();
      if (!se || se.ok === false) throw new Error((se && se.error) || 'stock error');
      const stockMap = {};
      (se.data || []).forEach(s => { stockMap[s.codigoProducto] = parseFloat(s.cantidadDisponible) || 0; });
      const productos = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      return { ok: true, data: _calcularPendientesEnvasadoFront(productos, stockMap) };
    }
    if (action === 'getGuias') {
      // La agrupación Hoy/Ayer (TZ Perú) la hace el FRONT; backend solo filtra (igual que getGuias de GAS).
      let rows = await _sbLeerTablaWH('guias');
      if (params.tipo)    rows = rows.filter(r => r.tipo === params.tipo);
      if (params.estado)  rows = rows.filter(r => String(r.estado) === String(params.estado));
      if (params.usuario) rows = rows.filter(r => r.usuario === params.usuario);
      if (params.limit)   rows = rows.slice(0, parseInt(params.limit));
      return { ok: true, data: rows };
    }
    if (action === 'getGuia') {
      // 1 guía + detalle (RPC dedicada con join) + enriquecer descripciones con cache local (réplica fiel de getGuia GAS).
      const out = await _sbRpcWH('get_guia_rls', { p_id: params.idGuia });
      if (!out) throw new Error('get_guia sin respuesta');
      // [40x #4] 'no encontrada' = respuesta legítima (no GAS); cualquier otro error (auth/transitorio) → cae a GAS.
      if (out.ok === false) return /no encontrada/i.test(out.error || '') ? out : null;
      const guia = _sbRowsToObjsFront('guias', [out.guia])[0];
      guia.detalle = _sbRowsToObjsFront('guia_detalle', out.detalle);
      const pm = _prodMapWH();
      guia.detalle.forEach(d => { d.descripcionProducto = d.descripcionProducto || pm[d.codigoProducto] || d.codigoProducto; });
      return { ok: true, data: guia };
    }
    // ── Lecturas con LÓGICA DERIVADA (replican exacto el post-proceso del getXxx de GAS) ──
    if (action === 'getLotesVencimiento') {
      let rows = await _sbLeerTablaWH('lotes_vencimiento');
      const hoyMs = _diaLimaMs(new Date());   // [TZ Lima] medianoche de hoy en Lima
      if (params.codigoProducto) rows = rows.filter(r => r.codigoProducto === params.codigoProducto);
      if (params.soloActivos === 'true') rows = rows.filter(r => r.estado === 'ACTIVO' && parseFloat(r.cantidadActual) > 0);
      if (params.proximosVencer) {
        const dias = parseInt(params.proximosVencer) || 30;
        const limiteMs = hoyMs + dias * _UN_DIA;
        rows = rows.filter(r => r.fechaVencimiento && _diaLimaMs(r.fechaVencimiento) <= limiteMs);
      }
      rows = rows.map(r => { if (r.fechaVencimiento) r.diasRestantes = Math.round((_diaLimaMs(r.fechaVencimiento) - hoyMs) / _UN_DIA); return r; });
      rows.sort((a, b) => (a.diasRestantes || 999) - (b.diasRestantes || 999));
      return { ok: true, data: rows };
    }
    if (action === 'getLotesFIFO') {
      // Lotes ACTIVOS con stock de un producto, orden FIFO (vence primero). Réplica fiel de getLotesFIFO (GAS):
      // shape propio (fechaVencimiento ISO), diasRestantes en TZ Lima (blindaje). getHistorialLote NO migrado → sigue GAS.
      const codigoProducto = String(params.codigoProducto || '').trim();
      if (!codigoProducto) return { ok: false, error: 'codigoProducto requerido' };
      const out = await _sbRpcWH('leer_tabla_rls', { p_tabla: 'lotes_vencimiento' });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'lotes error');
      const hoyMs = _diaLimaMs(new Date());
      const lotes = [];
      (out.data || []).forEach(d => {
        if (String(d.cod_producto).toUpperCase() !== codigoProducto.toUpperCase()) return;
        if (String(d.estado || '').toUpperCase() !== 'ACTIVO') return;
        const cant = parseFloat(d.cantidad_actual) || 0;
        if (cant <= 0) return;
        const fv = d.fecha_vencimiento ? new Date(d.fecha_vencimiento) : null;
        const valida = fv && !isNaN(fv.getTime());
        lotes.push({
          idLote: String(d.id_lote), codigoProducto: String(d.cod_producto),
          fechaVencimiento: valida ? fv.toISOString() : '',
          cantidadActual: cant, idGuia: String(d.id_guia || ''),
          diasRestantes: valida ? Math.ceil((_diaLimaMs(d.fecha_vencimiento) - hoyMs) / _UN_DIA) : null,
          fechaCreacion: d.fecha_creacion ? new Date(d.fecha_creacion).toISOString() : ''
        });
      });
      lotes.sort((a, b) => {
        if (!a.fechaVencimiento && !b.fechaVencimiento) return 0;
        if (!a.fechaVencimiento) return 1;
        if (!b.fechaVencimiento) return -1;
        return new Date(a.fechaVencimiento).getTime() - new Date(b.fechaVencimiento).getTime();
      });
      return { ok: true, data: lotes };
    }
    // [cero-rastro] getMermasEnProceso/getMermasVencidas eliminados: sin callers — la vista usa mermas_lista (v2).
    if (action === 'getProductosNuevosRecientes') {
      const dias = parseInt(params && params.dias) || 3;
      const corteMs = _diaLimaMs(new Date()) - dias * _UN_DIA;   // [TZ Lima] corte en días-calendario
      let rows = (await _sbLeerTablaWH('producto_nuevo')).filter(r => {
        if (String(r.estado).toUpperCase() !== 'APROBADO') return false;
        if (!r.fechaAprobacion) return false;
        return _diaLimaMs(r.fechaAprobacion) >= corteMs;
      });
      const resultado = rows.map(r => {
        const obs = String(r.observacion || '').trim();
        const tipo = obs.indexOf('EQUIVALENTE') === 0 ? 'EQUIVALENTE' : 'NUEVO';
        return { idProductoNuevo: r.idProductoNuevo, codigoBarra: r.codigoBarra, descripcion: r.descripcion, foto: r.foto, cantidad: r.cantidad, fechaCreacion: r.fechaCreacion, fechaAprobacion: r.fechaAprobacion, aprobadoPor: r.aprobadoPor, usuario: r.usuario, tipoAprobacion: tipo, observacion: obs };
      }).sort((a, b) => _diaLimaMs(b.fechaAprobacion) - _diaLimaMs(a.fechaAprobacion));
      return { ok: true, data: resultado };
    }
    if (action === 'getPickups') {
      let rows = await _sbLeerTablaWH('pickups', params.force);   // [v2.13.372] force → lectura fresca
      const filtroEstado = params.estado || 'PENDIENTE';
      rows = rows.filter(r => {
        if (filtroEstado === 'TODOS') return true;
        const estados = String(filtroEstado).split(',').map(s => s.trim());
        return estados.indexOf(r.estado) >= 0;
      }).map(r => { try { r.items = JSON.parse(r.items || '[]'); } catch (e) { r.items = []; } return r; });
      rows.sort((a, b) => (_diaLimaMs(b.fechaCreado) || 0) - (_diaLimaMs(a.fechaCreado) || 0));  // recientes primero
      return { ok: true, data: rows };
    }
    if (action === 'getPickup') {
      const rows = await _sbLeerTablaWH('pickups');
      const p = rows.find(r => String(r.idPickup) === String(params.idPickup));
      if (!p) return { ok: false, error: 'Pickup no encontrado' };
      try { p.items = JSON.parse(p.items || '[]'); } catch (_) { p.items = []; }
      return { ok: true, data: p };
    }
    if (action === 'getListasSombra') {
      // Listas DISPONIBLE/EN_USO + completadas de HOY (TZ Lima). Réplica fiel de getListasSombra (GAS).
      const out = await _sbRpcWH('leer_tabla_rls', { p_tabla: 'listas_sombra' });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'listas error');
      const hoyMs = _diaLimaMs(new Date());
      const incluirCompletadas = !!params.incluirCompletadas;
      const rows = [];
      (out.data || []).forEach(r => {
        const estado = String(r.estado || '').toUpperCase();
        if (estado !== 'DISPONIBLE' && estado !== 'EN_USO') {
          if (!incluirCompletadas) return;
          if (!r.fecha_completada || _diaLimaMs(r.fecha_completada) < hoyMs) return;  // solo completadas de HOY (Lima)
        }
        let items = r.items;  // jsonb de PostgREST ya viene parseado (objeto), NO string como la hoja
        if (typeof items === 'string') { try { items = JSON.parse(items || '[]'); } catch (e) { items = []; } }
        if (!Array.isArray(items)) items = [];
        const total = items.length;
        const completos = items.filter(it => (parseFloat(it.cantidadEscaneada) || 0) >= (parseFloat(it.cantidad) || 0)).length;
        rows.push({
          idLista: String(r.id_lista || ''),
          fechaCreacion: r.fecha_creacion ? new Date(r.fecha_creacion).toISOString() : '',
          usuarioCreador: String(r.usuario_creador || ''),
          estado, usuarioTomada: String(r.usuario_tomada || ''),
          fechaTomada: r.fecha_tomada ? new Date(r.fecha_tomada).toISOString() : '',
          fechaCompletada: r.fecha_completada ? new Date(r.fecha_completada).toISOString() : '',
          nota: String(r.nota || ''), total, completos, items
        });
      });
      rows.sort((a, b) => String(b.fechaCreacion).localeCompare(String(a.fechaCreacion)));
      return { ok: true, data: { listas: rows } };
    }
    if (action === 'getStockMovimientos') {
      // RPC dedicada con filtro server-side (tabla grande). Luego sort+limit en cliente (idéntico a GAS).
      const cod = String(params.codigoProducto || '').trim();
      const out = await _sbRpcWH('stock_movimientos_rls', { p_cod: cod || null, p_limit: params.limit ? parseInt(params.limit) : null });
      if (!out || out.ok === false) throw new Error((out && out.error) || 'rpc stock_mov error');
      let rows = _sbRowsToObjsFront('stock_movimientos', out.data);
      rows.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
      if (params.limit) rows = rows.slice(0, parseInt(params.limit));
      return { ok: true, data: rows };
    }
    if (action === 'getHistorialStock') {
      // [BUG 2 · cutover Supabase] Historial REAL del producto: movimientos APLICADOS de
      // wh.stock_movimientos (con stock_antes/stock_despues ya calculados → saldo = DATO).
      // Las guías ABIERTAS no escriben movimiento → nunca aparecen acá (no más +15 fantasma
      // ni saldos negativos). Lee directo de Supabase (dato fresco, no depende del flip GAS).
      const codigos = String(params.codigoProducto || '').split(',').map(s => s.trim()).filter(Boolean);
      if (!codigos.length) return { ok: false, error: 'codigoProducto requerido' };
      const codSet = new Set(codigos.map(c => c.toUpperCase()));
      // Movimientos por cada código (RPC con filtro server-side).
      const lotsArr = await Promise.all(codigos.map(cod =>
        _sbRpcWH('stock_movimientos_rls', { p_cod: cod, p_limit: null })
          .then(o => (o && o.ok !== false && Array.isArray(o.data)) ? _sbRowsToObjsFront('stock_movimientos', o.data) : [])
          .catch(() => [])));
      const movsRaw = [].concat.apply([], lotsArr);
      if (!movsRaw.length) return { ok: true, data: [] };
      // Lotes (para enriquecer el INGRESO con su fecha de vencimiento).
      let lotes = [];
      try {
        const lo = await _sbRpcWH('leer_tabla_rls', { p_tabla: 'lotes_vencimiento' });
        if (lo && lo.ok !== false && Array.isArray(lo.data)) lotes = lo.data;
      } catch (_) {}
      // [HISTORIAL ENRIQUECIDO] Zona DESTINO de cada salida + tipo de guía (autoritativo desde wh.guias.id_zona).
      // El cache local de guías (OfflineManager) sólo tiene las recientes → para que TODA salida histórica
      // muestre su zona, leemos wh.guias directo (read-only, whitelisted en leer_tabla_rls). Sólo necesitamos
      // id_guia → {id_zona, tipo}; si falla (offline) caemos al cache local sin romper nada.
      const zonaGuiaMap = {};   // idGuia → { zona: 'ZONA-02'|..., tipo: 'SALIDA_ZONA'|... }
      try {
        const gz = await _sbRpcWH('leer_tabla_rls', { p_tabla: 'guias' });
        if (gz && gz.ok !== false && Array.isArray(gz.data)) {
          gz.data.forEach(g => {
            const gid = String(g.id_guia || '');
            if (gid) zonaGuiaMap[gid] = { zona: String(g.id_zona || ''), tipo: String(g.tipo || ''), usuario: String(g.usuario || '') };
          });
        }
      } catch (_) {}
      // Normaliza el código de zona crudo a una etiqueta legible y estable (ZONA-01/ZONA-02/JEFATURA/VENTAS).
      // Tolera variantes históricas (z001/Z001/z002/ALMACEN/vacío). Devuelve '' si no hay zona conocida.
      const _normZona = (raw) => {
        const z = String(raw || '').trim().toUpperCase();
        if (!z || z === 'ALMACEN') return '';
        if (z === 'Z001') return 'ZONA-01';
        if (z === 'Z002') return 'ZONA-02';
        return z;   // ya viene 'ZONA-01' / 'ZONA-02' / 'JEFATURA' / 'VENTAS'
      };
      const guiaMap = {};
      const guias = (typeof OfflineManager !== 'undefined' && OfflineManager.getGuiasCache) ? (OfflineManager.getGuiasCache() || []) : [];
      guias.forEach(g => { guiaMap[String(g.idGuia)] = g; });
      const _clasificar = (op, delta) => {
        const o = String(op || '').toUpperCase();
        const esIng = parseFloat(delta) > 0;
        if (o === 'AJUSTE_MANUAL' || o === 'AUDITORIA' || o.indexOf('AJUSTE') >= 0)
          return { esIngreso: esIng, fuente: 'ajuste', tipo: 'Ajuste ' + (esIng ? 'INC' : 'DEC') };
        if (o.indexOf('INICIAL') >= 0 || o === 'INI') return { esIngreso: esIng, fuente: 'guia', tipo: 'INICIAL' };
        if (o.indexOf('ENVASADO') >= 0) return { esIngreso: esIng, fuente: 'guia', tipo: esIng ? 'INGRESO ENVASADO' : 'SALIDA ENVASADO' };
        if (o === 'APROBACION_PN') return { esIngreso: esIng, fuente: 'guia', tipo: 'INGRESO (aprobación PN)' };
        return { esIngreso: esIng, fuente: 'guia', tipo: esIng ? 'INGRESO' : 'SALIDA' };
      };
      const data = movsRaw
        .filter(m => codSet.has(String(m.codigoProducto).toUpperCase()))
        .map(m => {
          const delta = parseFloat(m.delta || 0);
          const cls   = _clasificar(m.tipoOperacion, delta);
          const idGuia = String(m.origen || '');
          const g = guiaMap[idGuia] || {};
          const gz = zonaGuiaMap[idGuia] || {};
          // Zona DESTINO sólo aplica a SALIDAS hacia una zona (no a ingresos ni envasados/ajustes).
          const zona = (!cls.esIngreso) ? _normZona(gz.zona || g.idZona || '') : '';
          const mov = {
            idGuia, fecha: m.fecha || '', tipo: cls.tipo, tipoOperacion: String(m.tipoOperacion || ''),
            esIngreso: cls.esIngreso, cantidad: Math.abs(delta),
            saldo: parseFloat(m.stockDespues), stockAntes: parseFloat(m.stockAntes),
            usuario: m.usuario || g.usuario || gz.usuario || '—', origen: g.idProveedor || g.destino || '',
            zona, estado: g.estado || 'CERRADA', fuente: cls.fuente, aplicado: true
          };
          if (cls.esIngreso) {
            const codUp = String(m.codigoProducto).toUpperCase();
            const l = lotes.find(x => String(x.id_guia || '') === idGuia &&
                                       String(x.cod_producto || '').toUpperCase() === codUp);
            if (l) mov.lote = {
              idLote: String(l.id_lote),
              fechaVencimiento: l.fecha_vencimiento ? new Date(l.fecha_vencimiento).toISOString() : '',
              estado: String(l.estado || 'ACTIVO')
            };
          }
          return mov;
        })
        .filter(m => m.cantidad > 0)
        .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
      return { ok: true, data };
    }
    return null;  // sin backend directo → GAS
  }

  // ════════════════════════════════════════════════════════════════════
  // [PASO 5 · B4] ESCRITURA directa a Supabase — compone las RPCs atómicas del PASO 4 en el cliente.
  // INERTE por defecto: gate `_whEscrituraDirecta()` (localStorage 'wh_escritura_navegador'). Además cada RPC nace
  // con su flag server-side WH_*_DIRECTO=0 → devuelve *_OFF → _postDirecto retorna null → FALLBACK TOTAL a GAS.
  // Idempotencia: el localId del POST siembra los ids (mismo reintento → mismos ids → la RPC dedup).
  // ════════════════════════════════════════════════════════════════════
  function _whEscrituraDirecta() {
    try { return localStorage.getItem('wh_escritura_navegador') === '1' || window.WH_CONFIG?.escrituraNavegador === true; }
    catch (_) { return window.WH_CONFIG?.escrituraNavegador === true; }
  }

  // [PASO 5 · B5] Gate PROPIO de la impresión directa por navegador (independiente de la escritura directa).
  // INERTE por defecto: solo se activa con localStorage 'wh_impresion_navegador'='1' (o WH_CONFIG.impresionNavegador===true).
  // Con el flag OFF, los casos de impresión de _postDirecto devuelven null → la impresión sigue 100% por GAS.
  function _whImpresionDirecta() {
    try { return localStorage.getItem('wh_impresion_navegador') === '1' || window.WH_CONFIG?.impresionNavegador === true; }
    catch (_) { return window.WH_CONFIG?.impresionNavegador === true; }
  }

  // [PASO 5 · B5] Resuelve printerHint ('TICKET'|'ADHESIVO'|'ZPL') → printNodeId real en el CLIENTE.
  // El GAS usa getPrinterNodeId(tipo,'ALMACEN') leyendo la hoja IMPRESORAS; acá replicamos eligiendo de los
  // caches locales de impresoras. Orden de prioridad:
  //   1) printerIdOverride explícito (PrintHub: el admin eligió una impresora concreta del ecosistema).
  //   2) Impresora del cache (offline 'wh_impresoras' o ecosistema 'wh_printers_cache') que coincida en tipo,
  //      preferiendo la de la zona del operador y la app warehouseMos; descarta las offline si hay estado.
  // Devuelve String(printNodeId) o '' si no se puede resolver (→ el caller cae a GAS).
  function _resolvePrinterId(hint, printerIdOverride) {
    if (printerIdOverride != null && String(printerIdOverride).trim() !== '') {
      return String(printerIdOverride).trim();
    }
    const tipo = String(hint || '').toUpperCase();
    const miZona = String(window.WH_CONFIG?.zona || '').trim();

    // Junta candidatos de ambos caches (sin duplicar por printNodeId).
    const candidatos = [];
    const _push = (arr) => {
      (Array.isArray(arr) ? arr : []).forEach(p => {
        const pid = String((p && (p.printNodeId || p.printnode_id)) || '').trim();
        if (!pid) return;
        candidatos.push({
          pid,
          tipo:    String(p.tipo || '').toUpperCase(),
          idZona:  String(p.idZona || p.id_zona || '').trim(),
          app:     String(p.appOrigen || p.app || '').toLowerCase(),
          activo:  (p.activo == null) ? true : !!p.activo,
          // estado solo lo trae el cache de ecosistema; el offline no → tratar como desconocido (no descartar)
          online:  (p.state ? p.state === 'ONLINE' : (p.online == null ? null : !!p.online))
        });
      });
    };
    try {
      if (typeof OfflineManager !== 'undefined' && OfflineManager.getImpresorasCache) _push(OfflineManager.getImpresorasCache());
    } catch (_) {}
    try { _push((window._whPrintersCache && window._whPrintersCache.data) || []); } catch (_) {}

    if (!candidatos.length) return '';

    const _coincideTipo = (c) => !tipo || c.tipo === tipo || (!c.tipo);
    const elegibles = candidatos.filter(c => c.activo !== false && _coincideTipo(c) && c.online !== false);
    if (!elegibles.length) return '';

    // Puntaje: zona del operador > app warehouseMos > online conocido. El primero con mejor puntaje gana.
    const _score = (c) => {
      let s = 0;
      if (miZona && c.idZona === miZona) s += 4;
      if (!c.idZona) s += 1;                 // sin zona = comodín, mejor que zona ajena
      if (c.app.indexOf('warehouse') >= 0) s += 2;
      if (c.online === true) s += 1;
      return s;
    };
    elegibles.sort((a, b) => _score(b) - _score(a));
    return elegibles[0].pid;
  }
  // Dispatcher: mapea una acción de escritura → su RPC wh.* (PASO 4). Devuelve {ok,...} o null (→ fallback GAS:
  // tanto si la acción no está cableada como si la RPC responde *_OFF / error).
  // [v2.13.300] Gate PROPIO de la impresión de adhesivos por LOTE vía Supabase (Edge print-adhesivo +
  // RPCs atómicas wh.lote_adhesivo_*). Independiente de escritura/impresión single-job. INERTE por
  // defecto (espeja el flag server-side WH_LOTE_ADHESIVO_DIRECTO). OFF → null → la cola sigue en GAS.
  function _whLoteAdhesivoDirecto() {
    // [cutover 2026-06-21] DEFAULT ON → 100% Supabase en todos los equipos. El KILL-SWITCH real es el
    // flag server-side WH_LOTE_ADHESIVO_DIRECTO (mos.config): si está '0', las RPCs devuelven *_OFF →
    // _postDirectoLoteAdhesivo retorna null → cae a GAS AL INSTANTE, sin re-deploy ni tocar dispositivos.
    // Opt-out por dispositivo (debug): localStorage 'wh_lote_adhesivo_navegador'='0' o WH_CONFIG.loteAdhesivoNavegador=false.
    try {
      const v = localStorage.getItem('wh_lote_adhesivo_navegador');
      if (v === '0') return false;
      if (v === '1') return true;
    } catch (_) {}
    if (window.WH_CONFIG && window.WH_CONFIG.loteAdhesivoNavegador === false) return false;
    return true;
  }
  // [cero-GAS G1] Gate de lectura del estado de bloqueo + heartbeat (mos.estado_bloqueo_usuario). Default ON
  // en cliente; el KILL-SWITCH real es server-side WH_BLOQUEO_DIRECTO (mos.config): si != '1' el RPC devuelve
  // WH_BLOQUEO_DIRECTO_OFF → estadoBloqueoUsuarioDirecto retorna null → cae a GAS al instante, sin redeploy.
  // Opt-out por dispositivo (debug): localStorage 'wh_bloqueo_navegador'='0'.
  function _whBloqueoDirecto() {
    try {
      const v = localStorage.getItem('wh_bloqueo_navegador');
      if (v === '0') return false;
      if (v === '1') return true;
    } catch (_) {}
    if (window.WH_CONFIG && window.WH_CONFIG.bloqueoNavegador === false) return false;
    return true;
  }
  const _MESES_ADH = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  function _vtoDesdeFechaAdh(fechaEnv) {
    try { const d = new Date(fechaEnv); d.setFullYear(d.getFullYear() + 1); return _MESES_ADH[d.getMonth()] + '/' + d.getFullYear(); }
    catch (_) { return ''; }
  }
  // Dispara la Edge en mode:'lote' (fire-and-forget: la Edge completa el lote server-side hasta su
  // presupuesto; el frontend solo poll-ea el estado). Re-disparable: reserve-first es idempotente.
  async function _fireEdgePrintAdhesivo(idLote) {
    try {
      const token = await _mintTokenWH();
      _whFetchTimeout(`${_SB_URL}/functions/v1/print-adhesivo`, {
        method: 'POST',
        headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ mode: 'lote', idLote })
      }, 120000).catch(() => {});   // no await: corre server-side
    } catch (_) {}
  }
  // [v2.13.310] Llamada GENÉRICA y AWAITED al Edge print-adhesivo (la usa el modal de membretes
  // vía edgeCall). Pasa el body tal cual (mode:'crear'|'crear-membrete'|...). Devuelve el JSON del
  // Edge {ok,data}|{ok:false,error}. Timeout amplio (lote grande imprime server-side hasta ~140s).
  async function _printAdhesivoEdge(body) {
    const token = await _mintTokenWH();
    const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/print-adhesivo`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {})
    }, 150000);
    const d = await res.json().catch(() => null);
    if (!res.ok || !d) return { ok: false, error: 'print-adhesivo HTTP ' + res.status };
    return d;
  }
  // Rutea las acciones de LOTE de adhesivos a Supabase. Devuelve undefined (no es lote), null (→ GAS), o el resultado.
  async function _postDirectoLoteAdhesivo(params) {
    const LOTE_ACTIONS = ['crearLoteAdhesivo', 'imprimirSubLoteAdhesivo', 'getEstadoLoteAdhesivo', 'cancelarLoteAdhesivo'];
    if (LOTE_ACTIONS.indexOf(params.action) < 0) return undefined;   // no es acción de lote
    if (!_whLoteAdhesivoDirecto()) return null;                       // flag OFF → GAS

    if (params.action === 'crearLoteAdhesivo') {
      // Solo adhesivos de envasado (los membretes ME/WH siguen su propio camino GAS por ahora).
      const tipo = String(params.tipoEtiqueta || 'ADHESIVO_ENVASADO').toUpperCase();
      if (tipo !== 'ADHESIVO_ENVASADO') return null;
      const printerId = _resolvePrinterId('ADHESIVO', params.printerIdOverride);
      if (!printerId) return null;                                    // sin impresora resoluble → GAS
      let vto = String(params.vto || '');
      if (!vto && params.fechaEnvasado) vto = _vtoDesdeFechaAdh(params.fechaEnvasado);
      const out = await _sbRpcWH('lote_adhesivo_crear', { p: {
        codigoBarra: String(params.codigoBarra || ''), descripcion: params.descripcion || '',
        total: params.total, vto, tipoEtiqueta: tipo, usuario: params.usuario || '',
        origen: params.origen || 'WH', printerId: String(printerId),
        idempotencyKey: String(params.idempotencyKey || '')   // estable → dedup de lote
      } });
      if (!out || out.ok === false) {
        // Tope 500: NO caer a GAS (saltaría el límite). Devolver el error para mostrarlo.
        if (out && /EXCEDE_MAX/.test(String(out.error || ''))) return { ok: false, error: out.error, mensaje: out.mensaje };
        return null;                                                 // *_OFF / otro error → GAS
      }
      const idLote = out.data && out.data.idLote;
      if (!idLote) return null;
      if (!out.dedup) _fireEdgePrintAdhesivo(idLote);                // imprime server-side (solo lote nuevo)
      return { ok: true, data: out.data };                           // shape paritario con GAS crearLoteAdhesivo
    }

    if (params.action === 'imprimirSubLoteAdhesivo') {
      // Migrado: el front YA NO orquesta sub-jobs (lo hace la Edge). Aquí solo devolvemos el ESTADO
      // (poll del modal de progreso). Self-heal: si el lote quedó sin arrancar (ENCOLADO/CREADO),
      // re-disparamos la Edge (idempotente). OJO: NO auto-reanudar PAUSADO_OUT_PAPER (requiere cambiar rollo).
      const idLote = String(params.idLote || '');
      if (!idLote) return null;
      const out = await _sbRpcWH('lote_adhesivo_get', { p: { idLote } });
      if (!out || out.ok === false) return null;
      const d = out.data || {};
      if ((d.status === 'ENCOLADO' || d.status === 'CREADO') && (d.completadas || 0) < (d.total || 0)) {
        _fireEdgePrintAdhesivo(idLote);
      }
      return { ok: true, data: { idLote, completadas: d.completadas, total: d.total, status: d.status, ultimoError: d.ultimoError, qtyImpresa: 0 } };
    }

    if (params.action === 'getEstadoLoteAdhesivo') {
      const idLote = String(params.idLote || '');
      if (!idLote) return null;
      const out = await _sbRpcWH('lote_adhesivo_get', { p: { idLote } });
      if (!out || out.ok === false) return null;
      return { ok: true, data: out.data };
    }

    if (params.action === 'cancelarLoteAdhesivo') {
      const idLote = String(params.idLote || '');
      if (!idLote) return null;
      const out = await _sbRpcWH('lote_adhesivo_cancelar', { p: { idLote } });
      if (!out || out.ok === false) return null;
      return { ok: true, data: out.data };
    }
    return null;
  }

  async function _postDirecto(params) {
    const lid = params.localId || _genLocalId();

    // [PASO 5 · B5] IMPRESIÓN DIRECTA primero (gate PROPIO _whImpresionDirecta, independiente de la escritura).
    // Se evalúa ANTES de las RPCs de escritura para que el flag de impresión pueda entrar a _postDirecto
    // SIN activar la escritura directa. Cada caso revalida su flag; devuelve {ok:true} o null (→ GAS).
    const _printResult = await _postDirectoImpresion(params, lid);
    if (_printResult !== undefined) return _printResult;

    // [v2.13.300] LOTE de adhesivos vía Supabase (gate PROPIO _whLoteAdhesivoDirecto). También antes de
    // la escritura directa: la cola de adhesivos tiene su propio flag y no requiere la escritura ON.
    const _loteResult = await _postDirectoLoteAdhesivo(params);
    if (_loteResult !== undefined) return _loteResult;

    // Las RPCs de escritura SOLO corren si la escritura directa está activa (flag propio). Con la escritura OFF
    // pero la impresión ON, llegar acá significa "no era impresión cableada" → null → GAS (no se toca Supabase).
    // [100x rollback-fix] EXCEPCIÓN: un ítem `_viaDirecta` (encolado por timeout de escritura directa, que PUDO
    // commitear en Supabase) SIEMPRE debe reintentarse por las RPCs aunque el flag global ya esté apagado por
    // rollback. La RPC es idempotente (dedup por el id sembrado del localId) → reintento seguro, no duplica.
    if (!_whEscrituraDirecta() && !params._viaDirecta) return null;

    // [kill-GAS · Mermas V2] motor OpLog → wh.aplicar_op (387). Idempotente por idOp. MERMA_AGREGAR (registra,
    // no mueve stock) / MERMA_SOLUCIONAR (reusa wh.resolver_merma). OpLog lee {ok,data}. Cero-GAS.
    if (params.action === 'aplicarOp') {
      const out = await _sbRpcWH('aplicar_op', { p: {
        idOp: String(params.idOp || ''), tipo: params.tipo, payload: params.payload,
        usuario: params.usuario || '', idGuia: params.idGuia || ''
      } });
      if (!out) return null;   // sin respuesta (red) → cola reintenta; nunca duplica (idempotente por idOp)
      return out;              // {ok:true,data:{...}} | {ok:false,error} — OpLog decide saved/failed
    }

    // [kill-GAS] procesar mermas descartadas → wh.procesar_eliminacion_mermas (392): guía SALIDA_MERMA + marca.
    if (params.action === 'procesarEliminacionMermas') {
      const out = await _sbRpcWH('procesar_eliminacion_mermas', { p: {
        claveAdmin: params.claveAdmin || '', usuario: params.usuario || ''
      } });
      if (!out) return null;
      return out;   // {ok:true,data:{autorizado,idGuiaSalida,procesados,fallidos}}
    }

    if (params.action === 'crearAjuste') {
      const out = await _sbRpcWH('crear_ajuste', { p: {
        id_ajuste: 'AJ_' + lid, codigo_producto: String(params.codigoProducto || ''),
        tipo: params.tipoAjuste === 'INC' ? 'INC' : 'DEC', cantidad: params.cantidadAjuste,
        motivo: params.motivo || '', usuario: params.usuario || '', id_auditoria: params.idAuditoria || '',
        id_stock_nuevo: 'STK_' + lid, id_mov: 'MOV_' + lid
      } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS
      return out;
    }
    if (params.action === 'registrarEnvasado') {
      // ORQUESTADOR ATÓMICO. El cliente resuelve el catálogo (derivado→base+factor) del cache; la RPC mueve stock en 1 tx.
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const cod = String(params.codigoBarra || '').trim();
      const der = prods.find(p => String(p.codigoBarra) === cod);
      // [418 · review MED2] Un registro 🤝 COLABORATIVO jamás cae a GAS en silencio:
      // el GAS no conoce `colaborador` → se registraría como normal y el compañero
      // perdería su mitad sin que nadie lo vea. Mejor fallar VISIBLE (el rollback
      // optimista del front avisa "volvé a intentar").
      const _esColab = !!String(params.colaborador || '').trim();
      if (!der || !der.codigoProductoBase) {
        if (_esColab) throw new Error('Envasado colaborativo: producto no resuelto en el catálogo local — sincroniza y reintenta');
        return null;  // no resoluble → GAS
      }
      // [FIX envasado-directo] El catálogo guarda codigoProductoBase como skuBase del granel
      // (ej "LEV1155"), NO como idProducto/codigoBarra. Antes solo se buscaba por idProducto/codigoBarra
      // → base=undefined SIEMPRE → return null → TODO el envasado caía a GAS (0 filas directas). Mismo
      // orden de resolución que GAS (_registrarEnvasadoImpl: skuBase || idProducto) + codigoBarra como red.
      const claveBase = String(der.codigoProductoBase).trim();
      const base = prods.find(p => String(p.skuBase) === claveBase || String(p.idProducto) === claveBase || String(p.codigoBarra) === claveBase);
      const factor = parseFloat(der.factorConversionBase) || 0;
      const unidades = parseInt(params.unidadesProducidas) || 0;
      if (!base || factor <= 0 || unidades <= 0) {
        if (_esColab) throw new Error('Envasado colaborativo: datos del producto incompletos — no se puede registrar por la vía directa; reintenta');
        return null;  // datos incompletos → GAS
      }
      const out = await _sbRpcWH('registrar_envasado', { p: {
        id_envasado: 'ENV_' + lid, cod_producto_base: String(base.codigoBarra), cod_producto_envasado: cod,
        cantidad_base: unidades * factor, unidades_producidas: unidades, unidad_base: base.unidad || '',
        fecha_vencimiento: params.fechaVencimiento || '', usuario: params.usuario || '',
        // [418] 🤝 colaborativo: el pago se divide 50/50 (SQL 418). '' = registro normal.
        colaborador: String(params.colaborador || '')
      } });
      if (!out || out.ok === false) {
        // [418 · MED2] error de la RPC en un 🤝 (p.ej. COLABORADOR_NO_ENCONTRADO/AMBIGUO)
        // → error VISIBLE, nunca GAS silencioso que registre sin colaborador.
        if (_esColab) throw new Error('Envasado colaborativo rechazado: ' + ((out && out.error) || 'sin conexión directa') + ' — revisa el compañero elegido y reintenta');
        return null;
      }
      if (!out.dedup) {
        const ses = (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '';
        _sbRpcWH('registrar_actividad', { p_id_sesion: ses, p_tipo: 'ENVASADO_REGISTRADO', p_cantidad: 1 }).catch(function(){});   // [fix red frágil] fire-and-forget
        _sbRpcWH('registrar_actividad', { p_id_sesion: ses, p_tipo: 'UNIDADES_ENVASADAS', p_cantidad: unidades }).catch(function(){});   // [fix red frágil] fire-and-forget
      }
      // [shape-fix] el front lee res.data?.idEnvasado (app.js:8916). La RPC devuelve {ok,id_envasado,dedup} al nivel raíz →
      // envolver al shape GAS {ok:true, data:{idEnvasado, unidadesProducidas, dedup?}, dedup?}. id determinista 'ENV_'+lid.
      const _dataEnv = { idEnvasado: 'ENV_' + lid, unidadesProducidas: unidades };
      if (out.dedup) _dataEnv.dedup = true;
      return { ok: true, data: _dataEnv, dedup: !!out.dedup };  // NOTA: la impresión de etiquetas se dispara aparte (API.imprimirDirecto) — efecto secundario
    }
    if (params.action === 'aprobarPreingreso') {
      // ORQUESTADOR ATÓMICO (crear guía desde preingreso + marcar PROCESADO en 1 tx). Idempotente.
      const out = await _sbRpcWH('aprobar_preingreso', { p: {
        id_preingreso: String(params.idPreingreso || ''), id_guia: 'G_' + lid, usuario: params.usuario || ''
      } });
      if (!out || out.ok === false) return null;
      // [shape-fix] el front lee res.data.idGuia SIN optional-chaining (app.js:14645 y 15581) → si data falta, lanza.
      // La RPC devuelve {ok,id_guia,dedup}. En dedup la guía ya existía con OTRO id → usar el que la RPC devuelve
      // (out.id_guia/out.idGuia); si no, el determinista 'G_'+lid. Envolver al shape GAS {ok:true, data:{idGuia, dedup?}}.
      const _idGuia = String(out.id_guia || out.idGuia || ('G_' + lid));
      const _data = { idGuia: _idGuia };
      if (out.dedup) _data.dedup = true;
      return { ok: true, data: _data, dedup: !!out.dedup };
    }
    if (params.action === 'auditarProducto') {
      // ORQUESTADOR ATÓMICO en server (auditoría+ajuste en 1 tx — hallazgo 40x #4). Idempotente por id_auditoria.
      const out = await _sbRpcWH('auditar_producto', { p: {
        codigo_barra: String(params.codigoBarra || ''), stock_fisico: params.stockFisico,
        usuario: params.usuario || '', observacion: params.observacion || '',
        id_auditoria: 'AUD_' + lid, id_ajuste: 'AJ_' + lid, id_stock_nuevo: 'STK_' + lid, id_mov: 'MOV_' + lid
      } });
      if (!out || out.ok === false) return null;
      if (!out.dedup) { _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'AUDITORIA_EJECUTADA', p_cantidad: 1 }).catch(function(){}); }
      return out;
    }
    if (params.action === 'asignarAuditoria') {
      // Asigna una auditoría (fila AUDITORIAS estado ASIGNADA con el stock_sistema VIGENTE de wh.stock).
      // NO mueve stock. Idempotente por id_auditoria. El id lo siembra el localId → reintento no duplica fila.
      const out = await _sbRpcWH('asignar_auditoria', { p: {
        id_auditoria: 'AUD_' + lid, codigo_producto: String(params.codigoProducto || ''),
        usuario: params.usuario || ''
      } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS
      // [shape] el front lee res.data?.idAuditoria (paridad con GAS asignarAuditoria → data:{idAuditoria}).
      return { ok: true, data: { idAuditoria: String(out.idAuditoria || out.id_auditoria || ('AUD_' + lid)) }, dedup: !!out.dedup };
    }
    if (params.action === 'ejecutarAuditoria') {
      // Registra el conteo físico sobre una auditoría ASIGNADA (estado→EJECUTADA, diff/resultado). NO mueve stock
      // (el path que mueve stock es auditarProducto, ya directo). Idempotente NATURAL por estado (EJECUTADA→dedup).
      const out = await _sbRpcWH('ejecutar_auditoria', { p: {
        id_auditoria: String(params.idAuditoria || ''), stock_fisico: params.stockFisico,
        observacion: params.observacion || '', usuario: params.usuario || ''
      } });
      if (!out || out.ok === false) return null;
      if (!out.dedup) { _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'AUDITORIA_EJECUTADA', p_cantidad: 1 }).catch(function(){}); }
      // [shape] el front lee res.data?.diferencia/resultado (paridad con GAS ejecutarAuditoria → data:{idAuditoria,diferencia,resultado}).
      return { ok: true, data: { idAuditoria: String(params.idAuditoria || ''), diferencia: (out.diferencia != null ? out.diferencia : 0), resultado: String(out.resultado || '') }, dedup: !!out.dedup };
    }
    if (params.action === 'actualizarFotosPreingreso' || params.action === 'actualizarPreingreso') {
      // PATCH del preingreso (whitelist: id_proveedor/monto/comentario/fotos/cargadores) — solo los campos presentes.
      const p = { id_preingreso: String(params.idPreingreso || '') };
      if (!p.id_preingreso) return null;
      if ('idProveedor' in params) p.id_proveedor = params.idProveedor;
      if ('monto' in params)       p.monto = params.monto;
      if ('comentario' in params)  p.comentario = params.comentario;
      if ('fotos' in params)       p.fotos = String(params.fotos || '');
      if ('cargadores' in params)  p.cargadores = params.cargadores;
      // [aviso-directo] snapshot_aviso: lo persiste imprimirAvisoCajeros tras avisar a cajas.
      // La RPC wh.actualizar_preingreso ya lo acepta (33_wh_actualizar_preingreso.sql). Se envía
      // SOLO si el front lo manda explícito (no se toca en patches normales del preingreso). Acepta
      // string JSON o el snapshot ya stringificado — se normaliza a string para la columna json.
      if ('snapshotAviso' in params) {
        p.snapshot_aviso = (typeof params.snapshotAviso === 'string')
          ? params.snapshotAviso
          : JSON.stringify(params.snapshotAviso || {});
      }
      const out = await _sbRpcWH('actualizar_preingreso', { p });
      if (!out || out.ok === false) return null;
      return out;
    }
    if (params.action === 'eliminarFotoDrive') {
      // foto NUEVA (Storage) → DELETE por path; foto VIEJA (Drive fileId) → null (la borra GAS).
      const ref = String(params.path || params.fileId || '').trim();
      if (!ref) return null;
      // las de Storage son URL/path con 'wh-fotos/' o un path con '/'; las de Drive son fileIds sin '/'
      if (!ref.includes('/') && !ref.includes('wh-fotos')) return null;   // fileId de Drive → GAS
      const path = ref.replace(/^.*wh-fotos\//, '').replace(/^https?:\/\/[^/]+\/storage\/v1\/object\/(public\/)?wh-fotos\//, '');
      try {
        const token = await _mintTokenWH();
        const res = await _whFetchTimeout(`${_SB_URL}/storage/v1/object/wh-fotos/${path}`, {
          method: 'DELETE', headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token }
        }, 12000);
        return { ok: res.ok };
      } catch (_) { return null; }
    }
    if (params.action === 'analizarListaSombra') {
      // IA directa vía Edge `ia` (Claude). Acepta TEXTO pegado, o UNA o VARIAS fotos/imágenes, o un PDF (visión, Sonnet).
      // [multi] archivos = [{b64,mime}] (varias imágenes = partes de la misma lista, o 1 PDF). Back-compat: archivoB64 único.
      let archivos = Array.isArray(params.archivos) ? params.archivos.filter(a => a && a.b64) : [];
      if (!archivos.length && params.archivoB64) archivos = [{ b64: params.archivoB64, mime: params.mimeType }];
      const esArchivo = archivos.length > 0;
      const texto = String(params.texto || '').trim();
      if (!esArchivo && !texto) return { ok: false, error: 'TEXTO_VACIO' };
      if (!esArchivo && texto.length > 50000) return { ok: false, error: 'TEXTO_MUY_LARGO', mensaje: 'Chunk demasiado grande' };
      // System prompt COLUMNA-CONSCIENTE: la IA nunca confunde "solicitado" con mín/máx/stock/precio.
      const system = [
        'Eres un asistente experto en leer listas/tablas de PEDIDOS de almacén y extraer SOLO los productos con la cantidad SOLICITADA (lo que hay que despachar).',
        'La lista puede venir como texto pegado, foto(s), captura de Excel, tabla o PDF (WhatsApp, email, ticket impreso).',
        'Si recibes VARIAS imágenes, son PARTES DE LA MISMA lista (varias fotos o varias hojas): combínalas en UN SOLO resultado, sin duplicar productos ni sumar cantidades repetidas. Lee con cuidado incluso fotos inclinadas, con sombra o de baja calidad; si un renglón es ilegible, omítelo antes que inventar.', '',
        'REGLA DE ORO — COLUMNAS DE CANTIDAD:',
        'Muchas listas traen VARIAS columnas numéricas por producto. Ejemplo:',
        '  Producto            Solicitado   Cant.Min   Cant.Max   Stock   Precio',
        '  AJINOMOTO 1KG          15           10         40       120    12.50',
        'Debes usar SIEMPRE la cantidad SOLICITADA/PEDIDA (aquí 15) y NUNCA el mínimo, máximo, stock, precio ni el código.',
        '- Identifica la columna correcta por su encabezado: "solicitado","pedido","cantidad","cant","pedir","despachar","requerido","req","a despachar".',
        '- "min/minimo/mín" y "max/maximo/máx" son límites de reposición, NO el pedido: IGNÓRALOS.',
        '- "stock/saldo/existencia" y "precio/costo/P.U./importe" (suelen llevar decimales o S/) NO son el pedido: IGNÓRALOS.',
        '- Si NO hay encabezados claros y hay varios números, elige el que representa lo pedido; ante duda, el más plausible como pedido, NUNCA el precio ni el stock.',
        '- Si solo hay UNA cantidad por producto, úsala.', '',
        'IGNORA: cabeceras de tabla, totales/subtotales, separadores (---, ===), notas y columnas de código/stock/precio/mín/máx.', '',
        'POR CADA PRODUCTO devuelve:',
        '- nombre: descripción del producto en MAYÚSCULAS, limpia, sin códigos pegados',
        '- cantidad: número decimal con 1 decimal (ej: 15.0, 80.0, 0.5)',
        '- codigoVisto: opcional — si trae un código/sku al lado, ponlo (string), si no, omite el campo', '',
        'RESPONDE EXCLUSIVAMENTE con JSON válido en este formato (sin markdown, sin comentarios):',
        '{"items":[{"nombre":"...","cantidad":N.N,"codigoVisto":"..."}]}'
      ].join('\n');
      // Mensaje: archivo(s) (visión, Sonnet — no confunde columnas) o texto (default Haiku, ya chunkeado por el front).
      let body;
      if (esArchivo) {
        const bloques = [];
        for (const a of archivos.slice(0, 8)) {   // tope 8 bloques (varias fotos de la misma lista)
          const raw = String(a.b64 || '').trim(); if (!raw) continue;
          const data = raw.replace(/^data:[^;]+;base64,/, '');   // Claude rechaza el prefijo data-URI
          const m = /^data:([^;]+);/.exec(raw);
          const mime = (m && m[1]) || String(a.mime || '') || 'image/jpeg';
          bloques.push(/pdf/i.test(mime)
            ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data } }
            : { type: 'image',    source: { type: 'base64', media_type: mime, data } });
        }
        if (!bloques.length) return { ok: false, error: 'TEXTO_VACIO' };
        const nImg = bloques.length;
        bloques.push({ type: 'text', text: (nImg > 1
          ? 'Estas ' + nImg + ' imágenes son partes de la MISMA lista: combínalas en un solo resultado, sin duplicar. '
          : '') + 'Extrae cada producto con su cantidad SOLICITADA (no el mín/máx/stock/precio). Devuelve solo el JSON indicado.' });
        // [fix 500x S5] thinking desactivado: baja latencia y evita el bloque `thinking` que arriesgaba el timeout de visión.
        body = { model: 'claude-sonnet-5', max_tokens: 8192, thinking: { type: 'disabled' }, system, messages: [{ role: 'user', content: bloques }] };
      } else {
        body = { max_tokens: 8192, system, messages: [{ role: 'user', content: 'Limpia esta lista y devuelve solo JSON:\n\n' + texto }] };
      }
      let resp;
      // [CERO-FALLBACK] Sin fallback GAS: si el Edge `ia` falla, se devuelve error → el front reintenta/avisa.
      // Visión/PDF: timeout amplio (subida del/los archivo(s) en red móvil + procesamiento). Texto: 40s.
      try { resp = await _llamarEdgeIA(body, esArchivo ? 120000 : 40000); }
      catch (e) { return { ok: false, error: 'IA_EDGE_FAIL', mensaje: 'IA no disponible (Edge): ' + ((e && e.message) || 'red') }; }
      // [robusto] Sonnet-5 puede emitir un bloque `thinking` ANTES del texto → tomar el primer bloque type==='text',
      // no content[0] (que podría ser el thinking con .text vacío → PARSE_FAIL falso).
      const text = (resp && Array.isArray(resp.content)
        ? ((resp.content.find(b => b && b.type === 'text' && b.text) || {}).text || '')
        : '') || '';
      const first = text.indexOf('{'), last = text.lastIndexOf('}');
      if (first < 0 || last < 0) return { ok: false, error: 'PARSE_FAIL', mensaje: text.substring(0, 200) };
      let parsed;
      try { parsed = JSON.parse(text.substring(first, last + 1)); } catch (e) { return { ok: false, error: 'PARSE_FAIL', mensaje: String(e.message) }; }
      const items = Array.isArray(parsed.items) ? parsed.items : [];
      const limpios = items.map(it => ({
        nombre: String(it.nombre || '').toUpperCase().trim(),
        cantidad: Math.round((parseFloat(it.cantidad) || 0) * 10) / 10,
        codigoVisto: it.codigoVisto ? String(it.codigoVisto) : ''
      })).filter(it => it.nombre && it.cantidad > 0);
      return { ok: true, data: { items: limpios, total: limpios.length } };
    }
    if (params.action === 'subirFotoPreingreso') {
      // solo sube la foto a Storage y devuelve la URL (el front acumula la lista y la guarda con actualizarFotos).
      const b64 = String(params.fotoBase64 || '').trim();
      if (!b64) return null;
      try {
        const up = await _subirFotoStorage('preingresos', String(params.idPreingreso || ''), b64, params.mimeType || 'image/jpeg', lid + '_' + (parseInt(params.indice) || 1));
        return { ok: true, data: { url: up.url, fileId: up.path } };  // fileId = path en Storage (para eliminar)
      } catch (_) { return null; }
    }
    if (params.action === 'subirFotoGuia') {
      // sube la foto a Storage (máxima calidad) y setea guias.foto con la URL. Réplica del flujo GAS.
      // [fix] el front envía `fotoBase64` (app.js:8662, paridad con GAS subirFotoGuia). Leer `params.fotoBase`
      // (typo) dejaba b64 vacío → return null → fallback a GAS, que escribe la foto en el Sheet (no en Supabase,
      // donde vive la guía G_L… directa) → foto nunca se persistía/mostraba. Todas las 7 guías directas tenían foto=''.
      const b64 = String(params.fotoBase64 || params.fotoBase || '').trim();
      if (!b64) return null;
      // [v2.13.239] Una guía 'G_L…' nace DIRECTO en Supabase (no existe en el Sheet GUIAS). Para ella, el
      // fallback a GAS de subirFotoGuia es INÚTIL y DAÑINO: GAS sube la foto a Drive y hace
      // _actualizarColumnaGuia que busca la fila en el Sheet → no la encuentra → no escribe nada en guias.foto
      // (de Supabase). El front mostraría la URL de Drive UNA vez y el siguiente getGuia (lectura directa) la
      // borraría (foto='' en Supabase) = "la foto no aparece / desaparece". Por eso, si la subida a Storage
      // FALLA (timeout/red) en una guía directa, NO devolvemos null (que iría a GAS): lanzamos un error
      // TRANSITORIO → post() la encola para REINTENTO DIRECTO (idempotente por el id sembrado del localId →
      // mismo path en Storage → x-upsert no duplica). Las guías legadas (no 'G_L…') sí caen a GAS como antes.
      const _esDirecta = String(params.idGuia || '').indexOf('G_L') === 0;
      let url;
      try { url = (await _subirFotoStorage('guias', String(params.idGuia || ''), b64, params.mimeType || 'image/jpeg', lid)).url; }
      catch (e) { if (_esDirecta) throw (e instanceof Error ? e : new Error('storage upload falló')); return null; }
      const out = await _sbRpcWH('actualizar_foto_guia', { p: { id_guia: String(params.idGuia || ''), foto: url } });
      if (!out || out.ok === false) {
        // [v2.13.239] La subida a Storage SÍ ocurrió (la foto está en wh-fotos) pero la persistencia en
        // guias.foto falló. En una guía directa, ir a GAS no la persistiría en Supabase → devolver error
        // explícito (no null) para que el front avise "Error al subir foto" en vez de mostrar un fantasma
        // que el próximo refresh borra. _OFF (flag apagado) sí debe caer a GAS (return null) — distinguir.
        if (_esDirecta && out && out.ok === false && !/_OFF$/.test(String(out.error || ''))) {
          return { ok: false, error: String(out.error || 'no se pudo guardar la foto') };
        }
        return null;
      }
      // [B5] OCR del comprobante en BACKGROUND (fire-and-forget, igual que el disparo automático de GAS, Fotos.gs:59):
      // no bloquea el ok de la foto; persiste los campos SUNAT cuando Claude responde. Inerte si el flag OCR está OFF.
      _ocrComprobanteGuia(String(params.idGuia || ''), b64, params.mimeType || 'image/jpeg').catch(() => {});
      return { ok: true, data: { url } };
    }
    if (params.action === 'copiarFotoDePreingreso') {
      // Copia la foto elegida del preingreso a la guía. La foto del preingreso YA está en Storage (wh-fotos) o en
      // Drive (legado), siempre como URL pública en wh.preingresos.fotos. NO re-subimos el archivo: la guía referencia
      // la MISMA URL (la imagen es pública y compartida). Solo persistimos esa URL en guias.foto vía actualizar_foto_guia.
      // [fix guías directas] sin este cableo el POST caía a GAS, que para una guía 'G_L…' (en Supabase, NO en el Sheet)
      // no encontraba la fila → la foto copiada nunca persistía. Además FotoPicker._extractFileId('?id=') devuelve '' para
      // URLs de Storage → el fileId que recibía GAS venía vacío → fallaba doble. El front nos pasa la URL completa en `fotoUrl`.
      const idGuia  = String(params.idGuia || '');
      const fotoUrl = String(params.fotoUrl || '').trim();
      if (!idGuia || !fotoUrl) return null;   // sin la URL completa (p.ej. flujo viejo) → que GAS lo intente
      const out = await _sbRpcWH('actualizar_foto_guia', { p: { id_guia: idGuia, foto: fotoUrl } });
      if (!out || out.ok === false) return null;
      // [shape] el front lee res.data?.url (app.js:8693). Réplica del shape GAS {ok:true,data:{url}}.
      return { ok: true, data: { url: fotoUrl } };
    }
    // [cero-rastro] branch registrarMerma eliminado — alta de merma = API.mermaAltaManualV2 / mermaDesdeGuia (RPCs v2).
    if (params.action === 'setConfigValue') {
      // [Frente 4] guardar config directo a wh.config (gate WH_CONFIG_DIRECTO; rechaza secretos). *_OFF → GAS.
      const out = await _sbRpcWH('set_config', { clave: params.clave || '', valor: params.valor, descripcion: params.descripcion || '' });
      if (!out || (out.ok === false && /_OFF$/.test(String(out.error || '')))) return null;
      return out;
    }
    if (params.action === 'registrarProductoNuevo') {
      let fotoUrl = String(params.foto || '');
      const b64 = String(params.fotoBase64 || '').trim();
      if (b64) { try { fotoUrl = (await _subirFotoStorage('producto_nuevo', 'PN_' + lid, b64, params.mimeType || 'image/jpeg', lid)).url; } catch (_) { /* sin foto no bloquea el PN */ } }
      const out = await _sbRpcWH('registrar_producto_nuevo', { p: {
        codigoBarra: params.codigoBarra || '', idGuia: params.idGuia || '', marca: params.marca || '',
        descripcion: params.descripcion || '', idCategoria: params.idCategoria || '', unidad: params.unidad || '',
        cantidad: params.cantidad, fechaVencimiento: params.fechaVencimiento || '', foto: fotoUrl,
        usuario: params.usuario || ''
      } });
      if (!out || (out.ok === false && /_OFF$/.test(String(out.error || '')))) return null;  // flag off → GAS
      return out;  // {ok,data:{idProductoNuevo,codigoBarra,idempotente}}
    }
    if (params.action === 'crearGuia') {
      // si viene fotoBase64, súbela a Storage y pásala como URL (ya no fallback por foto)
      if (String(params.fotoBase64 || '').trim()) {
        try { params = { ...params, foto: (await _subirFotoStorage('guias', 'G_' + lid, params.fotoBase64, params.mimeType || 'image/jpeg', lid)).url }; }
        catch (_) { return null; }
      }
      const out = await _sbRpcWH('crear_guia', { p: {
        id_guia: 'G_' + lid, tipo: params.tipo, usuario: params.usuario || '',
        id_proveedor: params.idProveedor || '', id_zona: params.idZona || '',
        numero_documento: params.numeroDocumento || '', comentario: params.comentario || '',
        id_preingreso: params.idPreingreso || '', foto: String(params.foto || '')
      } });
      if (!out || out.ok === false) return null;
      // tracking de actividad (best-effort) — solo si NO fue dedup (no contar reintentos)
      if (!out.dedup) { _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'GUIA_CREADA', p_cantidad: 1 }).catch(function(){}); }
      // [shape-fix] el front lee res.data?.idGuia (app.js:7634) y res.offline. La RPC devuelve {ok,id_guia,dedup}
      // al nivel raíz → envolver al shape GAS {ok:true, data:{idGuia, estado:'ABIERTA'}}. idGuia = id determinista 'G_'+lid.
      return { ok: true, data: { idGuia: 'G_' + lid, estado: 'ABIERTA' }, dedup: !!out.dedup };
    }
    if (params.action === 'cerrarGuia') {
      // [152 · FIX DOBLE-CONTEO] Usa la RPC IDEMPOTENTE wh.cerrar_guia_idempotente(p_id_guia):
      // lee wh.guia_detalle, aplica SOLO delta = cant_recibida − cantidad_aplicada por línea,
      // SETEA cantidad_aplicada y escribe kardex único (MOVID_<guia>#<linea>). Recerrar tras
      // un reabrir (que ya NO resetea cantidad_aplicada) da delta 0 → no dobla. Antes esta
      // ruta usaba wh.cerrar_guia(jsonb) que dejaba cantidad_aplicada=0 → el cron de inactividad
      // re-aplicaba el total = DOBLE CONTEO. La firma es texto plano (no envoltura {p}).
      // Idempotente: si ya CERRADA/AUTOCERRADA, la RPC devuelve lineasSaltadas (no reaplica).
      const out = await _sbRpcWH('cerrar_guia_idempotente', { p_id_guia: String(params.idGuia || '') }, 'wh', 30000);   // [fix] 30s > 20s base
      if (!out || out.ok === false) return null;
      const _yaCerrada = String(out.eraEstado || '').toUpperCase() === 'CERRADA' || String(out.eraEstado || '').toUpperCase() === 'AUTOCERRADA' || !!(out.yaCerrada || out.ya_cerrada);
      if (!_yaCerrada) { _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'GUIA_CERRADA', p_cantidad: 1 }).catch(function(){}); }
      // [shape-fix] el front lee res.data?.montoTotal (app.js:7571). La RPC devuelve el monto al nivel raíz →
      // envolver al shape GAS {ok:true, data:{idGuia, estado:'CERRADA', montoTotal, yaCerrada?}}. Tolerante al nombre de campo.
      const _monto = (out.montoTotal != null ? out.montoTotal : (out.monto_total != null ? out.monto_total : (out.monto != null ? out.monto : 0)));
      const _data = { idGuia: String(params.idGuia || ''), estado: 'CERRADA', montoTotal: parseFloat(_monto) || 0 };
      if (_yaCerrada) _data.yaCerrada = true;
      return { ok: true, data: _data };
    }
    if (params.action === 'crearDespachoRapido') {
      // ORQUESTADOR ATÓMICO de la PARTE DE DATOS del despacho rápido (guía + detalle + cierre de stock + kardex)
      // en UNA transacción server-side (wh.crear_despacho_rapido). Replica _crearDespachoRapidoImpl SIN la impresión:
      // el front dispara imprimirTicketGuia aparte (queda en GAS→PrintNode, irreducible). Idempotente por id_guia
      // sembrado del localId → un reintento/doble-tap re-envía el MISMO id → la RPC dedupea (no duplica guía/stock/kardex).
      // El stock se aplica con UPDATE atómico (cantidad += signo·delta), money-safe. items: [{codigoBarra, cantidad}].
      const items = (Array.isArray(params.items) ? params.items : []).map(it => ({
        codigo_barra: String(it.codigoBarra || '').trim(), cantidad: parseFloat(it.cantidad) || 0
      }));
      const gid = 'G_' + lid;
      const out = await _sbRpcWH('crear_despacho_rapido', { p: {
        id_guia: gid, tipo: params.tipo || 'SALIDA_ZONA', id_zona: params.idZona || '',
        usuario: params.usuario || '', comentario: params.nota || params.comentario || '',
        items, local_id: lid
      } }, 'wh', 30000);   // [fix 2026-07-08] timeout cliente 30s > 20s de la base: despachos grandes ya NO se cortan a los 12s ni se falsean como "offline". Cliente > base = nunca aborta con la tx viva (sin commit-fantasma).
      // [fix diagnóstico 2026-07-08] Un `EXCEPCION` de la RPC (ej. statement timeout en despachos grandes) hacía
      //   ROLLBACK atómico (no commiteó nada) y acá se tragaba como null → caía a GAS y el operador solo veía
      //   "Error al generar guía". Ahora: *_OFF (kill-switch) sigue cayendo a GAS; un error REAL se SURFACE con su
      //   detalle (SQLERRM) para que el front lo muestre y quede diagnosticable — cero-GAS. El timeout ya subió a 20s.
      if (!out) return null;
      if (out.ok === false) {
        const _e = String(out.error || '');
        if (/_OFF$/.test(_e)) return null;   // kill-switch server apagado → GAS
        try { console.warn('[crearDespachoRapido] RPC error:', _e, '·', out.detalle || ''); } catch (_) {}
        return { ok: false, error: _e || 'error', detalle: String(out.detalle || '') };
      }
      // tracking (best-effort) — solo en la creación real, no en reintentos dedupados
      if (!out.dedup) { _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'GUIA_CREADA', p_cantidad: 1 }).catch(function(){}); }
      // [shape-fix] el front (app.js ~13275) lee res.data.idGuia, res.data.errores, res.data.esperados/items.
      // La RPC devuelve {ok,idGuia,items,errores} al nivel raíz → envolver al shape GAS. impresion:{ok:false,omitido}
      // porque la impresión la dispara el front por separado (imprimir:false). dedup propagado para coherencia.
      const _idGuia = String(out.idGuia || out.id_guia || gid);
      return { ok: true, data: {
        idGuia: _idGuia,
        errores: Array.isArray(out.errores) ? out.errores : [],
        items: (out.items != null ? out.items : items.length),
        esperados: items.length,
        impresion: { ok: false, error: 'omitido' }
      }, dedup: !!out.dedup };
    }
    if (params.action === 'cerrarPickupConDespacho') {
      // CIERRE DE PICKUP 100% SUPABASE (wh.cerrar_pickup_con_despacho, SQL 210).
      // La lista de pickups (🛒) ya se LEE de wh.pickups (getPickups → _sbLeerTablaWH).
      // Antes el cierre seguía en GAS (Hoja PICKUPS) → los pickups RIZ (que solo
      // viven en Supabase) daban "Pickup no encontrado" y no se podían despachar.
      // Ahora el cierre orquesta en Supabase: lee el pickup, idempotencia por estado,
      // crea la GUIA_SALIDA vía wh.crear_despacho_rapido (motor de dinero ya vivo) y
      // marca el pickup terminal — TODO atómico (all-or-nothing). El front dispara la
      // impresión del ticket aparte (imprimir:false acá). idGuia estable GPCK_<idPickup>
      // → un reintento/doble-tap dedupea (no duplica guía/stock/kardex).
      const out = await _sbRpcWH('cerrar_pickup_con_despacho', { p: {
        id_pickup:        String(params.idPickup || ''),
        usuario:          params.usuario || '',
        items:            Array.isArray(params.items) ? params.items : [],
        despacho_detalle: Array.isArray(params.despachoDetalle) ? params.despachoDetalle : []
      } }, 'wh', 30000);   // [fix] 30s > 20s base
      // SOLO *_OFF (flag server apagado) cae a GAS — kill-switch instantáneo. Cualquier
      // otra respuesta (ok:true / yaCerrado / error real) viene de Supabase, que es el
      // master de pickups → devolverla tal cual (NO doble-path a GAS, que para RIZ daría
      // "no encontrado" y para cierre-caja crearía una guía duplicada por el otro camino).
      // Un throw (timeout/5xx) lo maneja post(): cerrarPickupConDespacho ∈ _IDEMPOTENT_ACTIONS
      // → NO cae a GAS (pudo commitear) → reintenta directo. La RPC es atómica: ok:false ⇒ rollback.
      if (!out || out.error === 'WH_CERRAR_PICKUP_DIRECTO_OFF') return null;
      return out;   // {ok:true,data:{idGuia,estado,despachados,noDespachados}} | {ok:false,error,yaCerrado}
    }
    if (params.action === 'actualizarPickup') {
      // Estado + lock (atendido_por) del pickup en Supabase — MISMO store que lee
      // la lista (getPickups). Antes iba a GAS (Hoja) → el lock no viajaba a la
      // lista → 2 equipos del mismo operador no veían lo mismo + RIZ fallaba.
      // Lock POR USUARIO (normalizado): otro device del mismo operador no entra en
      // conflicto; otro operador sí. *_OFF → GAS; conflicto/ok se devuelven tal cual.
      const out = await _sbRpcWH('actualizar_pickup', { p: {
        id_pickup:    String(params.idPickup || ''),
        estado:       params.estado || '',
        lock_usuario: params.lockUsuario || '',
        tomar_lock:   params.tomarLock === true,
        liberar_lock: params.liberarLock === true,
        claveAdmin:   params.claveAdmin || ''   // [SQL 440] re-verificada server-side (non-strict) en el path ELIMINAR
      } });
      if (!out || out.error === 'WH_PICKUP_ESTADO_DIRECTO_OFF') return null;
      return out;   // {ok:true} | {ok:false,error,atendidoPor,conflicto}
    }
    if (params.action === 'guardarProgresoPickup') {
      // Autosave (cada 4s) del avance escaneado → wh.pickups.items (jsonb). Esto es
      // lo que hace que el OTRO equipo del mismo operador vea el progreso al
      // "↻ Continuar" (hidrata despachadoPorCodigo del server). Toma el lock si está libre.
      const out = await _sbRpcWH('guardar_progreso_pickup', { p: {
        id_pickup:    String(params.idPickup || ''),
        items:        Array.isArray(params.items) ? params.items : [],
        lock_usuario: params.lockUsuario || ''
      } });
      if (!out || out.error === 'WH_PICKUP_ESTADO_DIRECTO_OFF') return null;
      return out;   // {ok:true} | {ok:false,error,atendidoPor,conflicto}
    }
    if (params.action === 'liberarPickup') {
      // "Soltar" el pickup: limpia atendido_por. Si hay progreso queda EN_PROCESO
      // (cualquiera del equipo continúa); si no, vuelve a PENDIENTE.
      const out = await _sbRpcWH('liberar_pickup', { p: { id_pickup: String(params.idPickup || '') } });
      if (!out || out.error === 'WH_PICKUP_ESTADO_DIRECTO_OFF') return null;
      return out;   // {ok:true,data:{hayProgreso}}
    }
    if (params.action === 'reabrirGuia') {
      // [152 · INVARIANTE] reabrir NUNCA toca stock y NUNCA resetea cantidad_aplicada (solo estado=ABIERTA +
      // ultima_actividad). El stock aplicado se preserva → al recerrar (idempotente) delta = cant_recibida −
      // cantidad_aplicada = 0 sin editar → NO dobla. Idempotente por estado (FOR UPDATE). La autorización admin
      // (REABRIR_GUIA) se valida ANTES en el flujo. id_guia real (no generado).
      const out = await _sbRpcWH('reabrir_guia', { p: { id_guia: String(params.idGuia || ''), usuario: params.usuario || '', claveAdmin: params.claveAdmin || '' } });
      // [cero-caída · fix bypass] un rechazo de re-verificación (autorizado:false) NO debe caer a GAS
      // (el GAS reabrirGuia no valida clave → bypass). Solo caemos a GAS por transporte caído (null).
      if (!out) return null;
      if (out.autorizado === false) return out;   // clave rechazada → propagar, NO GAS
      if (out.ok === false) return null;           // otros errores → comportamiento previo
      return out;
    }
    if (params.action === 'agregarDetalleGuia') {
      // idempotente por local_id (dedup en la RPC — hallazgo 40x #2). Sin foto/actividad → directo seguro.
      const out = await _sbRpcWH('agregar_detalle_guia', { p: {
        id_guia: String(params.idGuia || ''), codigo_producto: String(params.codigoProducto || ''),
        cantidad_esperada: params.cantidadEsperada, cantidad_recibida: params.cantidadRecibida,
        precio_unitario: params.precioUnitario, id_lote: params.idLote || '',
        observacion: params.observacion || '', fecha_vencimiento: params.fechaVencimiento || '',
        id_detalle: 'DET_' + lid, id_lote_nuevo: 'LOTE_' + lid, id_mov: 'MOV_' + lid,
        usuario: params.usuario || '', local_id: lid
      } });
      if (!out || out.ok === false) return null;
      // [shape-fix] consumidor PESADO: el front hace itemFinal = {...res.data} y lee res.data.idDetalle,
      // res.data.idGuia, res.data.cantidadRecibida, res.data.descripcionProducto (app.js:6903). La RPC devuelve
      // {ok,...} al nivel raíz → reconstruir el shape GAS {ok:true, data:{...9 campos...}}. id determinista 'DET_'+lid;
      // descripción desde el cache (mismo prodMap que getGuia); valores numéricos = los enviados (tolerante al echo RPC).
      const _cod = String(params.codigoProducto || '');
      const _pm = _prodMapWH();
      return { ok: true, data: {
        idDetalle: String(out.id_detalle || out.idDetalle || ('DET_' + lid)),
        idGuia: String(params.idGuia || ''),
        codigoProducto: _cod,
        descripcionProducto: _pm[_cod] || _cod,
        cantidadEsperada: (params.cantidadEsperada != null && params.cantidadEsperada !== '') ? (parseFloat(params.cantidadEsperada) || 0) : 0,
        cantidadRecibida: (params.cantidadRecibida != null && params.cantidadRecibida !== '') ? (parseFloat(params.cantidadRecibida) || 0) : 0,
        precioUnitario: (params.precioUnitario != null && params.precioUnitario !== '') ? (parseFloat(params.precioUnitario) || 0) : 0,
        idLote: String(out.id_lote || params.idLote || ''),
        fechaVencimiento: params.fechaVencimiento || ''
      } };
    }
    if (params.action === 'crearPreingreso') {
      // El frontend genera idPreingreso estable ('PI'+ts) y lo pasa a ambos backends → idempotente en el cruce.
      // Fotos NO van acá (se suben aparte: subirFotoPreingreso → Storage → actualizarFotosPreingreso). RPC dedup por id_preingreso.
      const out = await _sbRpcWH('crear_preingreso', { p: {
        id_preingreso: String(params.idPreingreso || ''),
        id_proveedor:  params.idProveedor || '',
        cargadores:    typeof params.cargadores === 'string' ? params.cargadores : JSON.stringify(params.cargadores || []),
        usuario:       params.usuario || '',
        monto:         params.monto,
        comentario:    params.comentario || '',
        fecha:         params.fecha || ''
      } });
      if (!out || out.ok === false) return null;
      // [shape-fix] el front lee res.data?.idPreingreso (app.js:15827). La RPC devuelve {ok,id_preingreso,dedup} →
      // envolver al shape GAS {ok:true, data:{idPreingreso}}. id = el front-generado que pasamos (idempotente en el cruce).
      return { ok: true, data: { idPreingreso: String(out.id_preingreso || out.idPreingreso || params.idPreingreso || '') }, dedup: !!out.dedup };
    }
    if (params.action === 'actualizarCantidadDetalle') {
      // Edita cant_recibida de una línea. Si la guía está CERRADA, la RPC ajusta stock por el DELTA → NO idempotente
      // natural → dedup por local_id (la acción ya está en _IDEMPOTENT_ACTIONS → lid estable en reintentos).
      // El front solo manda idDetalle/cantidadRecibida; la RPC resuelve idGuia/cod/cant_vieja/lote desde la fila.
      const out = await _sbRpcWH('actualizar_cantidad_detalle', { p: {
        id_detalle: String(params.idDetalle || ''), cantidad_recibida: params.cantidadRecibida,
        usuario: params.usuario || '', id_mov: 'MOV_' + lid, id_lote_nuevo: 'LOTE_' + lid,
        id_ajuste: 'AJ_' + lid, local_id: lid   // [FIX #1] id determinista del ajuste (guía CERRADA crea fila en wh.ajustes)
      } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS
      return out;
    }
    if (params.action === 'anularDetalle') {
      // Anula una línea. Idempotente NATURAL por estado (si ya ANULADO → yaAnulado, no re-devuelve stock); FOR UPDATE
      // en la RPC serializa contra anulación concurrente. Solo devuelve stock si la guía está CERRADA (no envasado).
      const out = await _sbRpcWH('anular_detalle', { p: {
        id_detalle: String(params.idDetalle || ''), usuario: params.usuario || '', id_mov: 'MOV_' + lid
      } });
      if (!out || out.ok === false) return null;
      return out;
    }
    if (params.action === 'actualizarFechaVencimiento') {
      // Edita la fecha de vencimiento de una línea + sincroniza el lote. NO toca stock. Idempotente natural por valor.
      const out = await _sbRpcWH('actualizar_fecha_vencimiento', { p: {
        id_detalle: String(params.idDetalle || ''), fecha_vencimiento: params.fechaVencimiento || '',
        usuario: params.usuario || '', id_lote_nuevo: 'LOTE_' + lid
      } });
      if (!out || out.ok === false) return null;
      return out;
    }
    if (params.action === 'actualizarGuia') {
      // Edita campos de CABECERA (solo los presentes). NO toca stock ni lotes. Idempotente natural (UPDATE concreto).
      // Mapeo fiel a la whitelist GAS; solo se incluye la clave en p si el front la mandó (mandar '' SÍ limpia).
      const p = { id_guia: String(params.idGuia || '') };
      if (!p.id_guia) return null;
      if ('tipo' in params)            p.tipo = String(params.tipo);
      if ('idProveedor' in params)     p.id_proveedor = String(params.idProveedor);
      if ('idZona' in params)          p.id_zona = String(params.idZona);
      if ('numeroDocumento' in params) p.numero_documento = String(params.numeroDocumento);
      if ('comentario' in params)      p.comentario = String(params.comentario);
      if ('foto' in params)            p.foto = String(params.foto);
      const out = await _sbRpcWH('actualizar_guia', { p });
      if (!out || out.ok === false) return null;
      return out;
    }
    // [cero-rastro] branch resolverMerma eliminado — resolver = RPC procesar_merma (API.procesarMerma).
    if (params.action === 'corregirUnidadesEnvasado') {
      // Corrige unidades de un envasado → mueve stock derivado (+deltaUds) y base (-deltaBase). El cliente resuelve
      // base/factor del cache (la RPC no tiene catálogo). Mueve stock por DELTA → NO idempotente → dedup por local_id.
      // La autorización admin (claveAdmin) se valida ANTES en el flujo; la RPC no la chequea.
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const env = (typeof OfflineManager !== 'undefined' && OfflineManager.getEnvasadosCache) ? (OfflineManager.getEnvasadosCache() || []).find(e => String(e.idEnvasado) === String(params.idEnvasado)) : null;
      const cbDer = env ? String(env.codigoProductoEnvasado || '') : '';
      const der = prods.find(p => String(p.codigoBarra) === cbDer);
      if (!der) return null;   // no resoluble → GAS
      const factor = parseFloat(der.factorConversionBase) || 0;
      if (factor <= 0) return null;
      const cbBaseEnv = env ? String(env.codigoProductoBase || '') : '';
      const base = prods.find(p => String(p.codigoBarra) === cbBaseEnv || String(p.skuBase) === cbBaseEnv || String(p.idProducto) === cbBaseEnv);
      const out = await _sbRpcWH('corregir_unidades_envasado', { p: {
        id_envasado: String(params.idEnvasado || ''), nuevas_unidades: params.nuevasUnidades,
        cod_producto_envasado: cbDer, cod_producto_base: base ? String(base.codigoBarra) : '',
        factor_base: factor, motivo: params.motivo || '', usuario: params.usuario || '',
        id_mov_der: 'MOVEDD_' + lid, id_mov_base: 'MOVEDB_' + lid, local_id: lid
      } });
      if (!out || out.ok === false) return null;
      // [shape-fix] el front lee res.data.udsNuevas, .udsViejas, .descripcion (app.js:9366). La RPC devuelve los
      // valores al nivel raíz → envolver al shape GAS {ok:true, data:{...}}. Tolerante al nombre del echo RPC;
      // fallback a lo que sabemos del cache (der.descripcion, env.unidadesProducidas) y el input (nuevasUnidades).
      const _udsNuevas = (out.udsNuevas != null ? out.udsNuevas : (out.uds_nuevas != null ? out.uds_nuevas : (parseInt(params.nuevasUnidades) || 0)));
      const _udsViejas = (out.udsViejas != null ? out.udsViejas : (out.uds_viejas != null ? out.uds_viejas : (env ? (parseFloat(env.unidadesProducidas) || 0) : 0)));
      return { ok: true, data: {
        idEnvasado: String(params.idEnvasado || ''),
        udsViejas: parseFloat(_udsViejas) || 0,
        udsNuevas: parseFloat(_udsNuevas) || 0,
        descripcion: String(out.descripcion || (der && der.descripcion) || '')
      } };
    }
    if (params.action === 'anularEnvasadoConClave') {
      // Anula un envasado (reverso EXACTO de registrar_envasado): -uds derivado, +cantBase base, anula lote+detalles.
      // Idempotente NATURAL por estado (si ya ANULADO → yaAnulado; FOR UPDATE serializa). La RPC usa los cod de la FILA
      // como fuente de verdad; igual le pasamos los del cache como respaldo. La autorización admin se valida ANTES.
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const env = (typeof OfflineManager !== 'undefined' && OfflineManager.getEnvasadosCache) ? (OfflineManager.getEnvasadosCache() || []).find(e => String(e.idEnvasado) === String(params.idEnvasado)) : null;
      const cbDer = env ? String(env.codigoProductoEnvasado || '') : '';
      const cbBaseEnv = env ? String(env.codigoProductoBase || '') : '';
      const base = prods.find(p => String(p.codigoBarra) === cbBaseEnv || String(p.skuBase) === cbBaseEnv || String(p.idProducto) === cbBaseEnv);
      const out = await _sbRpcWH('anular_envasado', { p: {
        id_envasado: String(params.idEnvasado || ''), cod_producto_envasado: cbDer,
        cod_producto_base: base ? String(base.codigoBarra) : '', motivo: params.motivo || '', usuario: params.usuario || ''
      } });
      if (!out || out.ok === false) return null;
      // [shape-fix] el front lee res.data.cantBaseRestit, .udsAnuladas, .descripcion (app.js:9374). La RPC devuelve
      // los valores al nivel raíz → envolver al shape GAS {ok:true, data:{...}}. Tolerante al echo; fallback al cache.
      const _cantBase = (out.cantBaseRestit != null ? out.cantBaseRestit : (out.cant_base_restit != null ? out.cant_base_restit : (env ? (parseFloat(env.cantidadBase) || 0) : 0)));
      const _udsAnul  = (out.udsAnuladas != null ? out.udsAnuladas : (out.uds_anuladas != null ? out.uds_anuladas : (env ? (parseFloat(env.unidadesProducidas) || 0) : 0)));
      const _descEnv  = _resolverDescEnvasado();
      return { ok: true, data: {
        idEnvasado: String(params.idEnvasado || ''),
        udsAnuladas: parseFloat(_udsAnul) || 0,
        cantBaseRestit: parseFloat(_cantBase) || 0,
        descripcion: String(out.descripcion || _descEnv(cbDer) || '')
      } };
    }
    if (params.action === 'anularEnvasadoManual') {
      // [envasado-directo] ANULACIÓN RÁPIDA (sin clave admin) desde la lista/historial y desde el "↺ Deshacer
      // inmediato". GAS la resuelve contra la HOJA ENVASADOS, pero con el sync OFF la hoja está CONGELADA y NO
      // contiene los envasados creados DIRECTO ('ENV_L…') → GAS devolvía "Envasado no encontrado" → el front
      // hacía rollback del stock revertido y el envasado quedaba COMPLETADO con stock inflado (no anulable).
      // Misma RPC IDEMPOTENTE que anularEnvasadoConClave (wh.anular_envasado): reverso EXACTO de registrar_envasado
      // (-uds derivado, +cantBase base, anula lote+detalles), idempotente NATURAL por estado (yaAnulado → no
      // re-revierte; FOR UPDATE serializa). La autorización de esta vía es el propio confirm del operador (no clave).
      // La RPC usa los cod de la FILA como fuente de verdad; le pasamos los del cache como respaldo.
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const env = (typeof OfflineManager !== 'undefined' && OfflineManager.getEnvasadosCache) ? (OfflineManager.getEnvasadosCache() || []).find(e => String(e.idEnvasado) === String(params.idEnvasado)) : null;
      const cbDer = env ? String(env.codigoProductoEnvasado || '') : '';
      const cbBaseEnv = env ? String(env.codigoProductoBase || '') : '';
      const base = prods.find(p => String(p.codigoBarra) === cbBaseEnv || String(p.skuBase) === cbBaseEnv || String(p.idProducto) === cbBaseEnv);
      const out = await _sbRpcWH('anular_envasado', { p: {
        id_envasado: String(params.idEnvasado || ''), cod_producto_envasado: cbDer,
        cod_producto_base: base ? String(base.codigoBarra) : '', motivo: params.motivo || 'anulación manual', usuario: params.usuario || 'manual'
      } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS (legado, solo si el flag está apagado)
      // [shape] el front solo lee res.ok (app.js:10091) / fire-and-forget (app.js:9958). Réplica del shape GAS
      // {ok:true, data:{idEnvasado, yaAnulado?}}. La RPC devuelve {ok,yaAnulado?,id_envasado} al nivel raíz.
      const _data = { idEnvasado: String(params.idEnvasado || '') };
      if (out.yaAnulado || out.ya_anulado) _data.yaAnulado = true;
      return { ok: true, data: _data };
    }
    if (params.action === 'marcarAlertaRevisada') {
      // [Tanda 3] Marca una alerta de stock como revisada. NO toca stock. Idempotente natural (UPDATE a valor fijo).
      const out = await _sbRpcWH('marcar_alerta_revisada', { p: { id_alerta: String(params.idAlerta || '') } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS
      return out;
    }
    if (params.action === 'aceptarTeoricoAlerta') {
      // [Tanda 3] Corrección one-click: crea AJUSTE (stock real → teórico) + marca revisada. SÍ TOCA STOCK por DELTA →
      // NO idempotente natural → dedup por local_id + id_ajuste determinista. La RPC lee real/teórico de la propia alerta.
      const out = await _sbRpcWH('aceptar_teorico_alerta', { p: {
        id_alerta: String(params.idAlerta || ''), usuario: params.usuario || '',
        id_ajuste: 'AJ_' + lid, id_stock_nuevo: 'STK_' + lid, id_mov: 'MOV_' + lid, local_id: lid
      } });
      if (!out || out.ok === false) return null;
      // [paridad GAS] el GAS devuelve data:{idAlerta, ajusteAplicado, idAjuste?}; la RPC los pone en el nivel raíz →
      // normalizar a la forma que espera app.js (res.data.ajusteAplicado).
      if (out.ok && !out.data) {
        out.data = { idAlerta: out.id_alerta, ajusteAplicado: out.ajusteAplicado || 0 };
        if (out.idAjuste) out.data.idAjuste = out.idAjuste;
      }
      return out;
    }
    if (params.action === 'addCargadorDia') {
      // [Tanda 3] +1 cargador del día. NO toca stock. Append no idempotente natural → dedup por id_log determinista.
      const out = await _sbRpcWH('add_cargador_dia', { p: {
        id_cargador: String(params.idCargador || ''), fecha: params.fecha || '', nombre: params.nombre || '',
        usuario: params.usuario || '', device_id: params.deviceId || '', id_log: 'CLG_' + lid
      } });
      if (!out || out.ok === false) return null;
      return out;   // ya viene { ok, data:{ idLog, conteo, fecha } } como el GAS
    }
    if (params.action === 'removeCargadorDia') {
      // [Tanda 3] -1 cargador del día (marca el ACTIVO más reciente como ELIMINADO). NO toca stock. Es un -1 real →
      // dedup por local_id (un reintento del mismo POST no debe quitar dos).
      const out = await _sbRpcWH('remove_cargador_dia', { p: {
        id_cargador: String(params.idCargador || ''), fecha: params.fecha || '', local_id: lid
      } });
      if (!out || out.ok === false) return null;
      return out;   // { ok, data:{ conteo, fecha } } como el GAS
    }
    return null;  // acción de escritura no cableada aún → GAS
  }

  // ════════════════════════════════════════════════════════════════════
  // [PASO 5 · B5] IMPRESIÓN DIRECTA — arma el ticket/etiqueta en el navegador (ImpresionDirecta) y lo manda a la
  // Edge `imprimir` (vía _imprimirDirecto), en vez de saltar a GAS→PrintNode. INERTE por _whImpresionDirecta().
  // Devuelve:
  //   • undefined → la acción NO es de impresión cableada (el dispatcher sigue con las RPCs de escritura/GAS).
  //   • { ok:true } → PrintNode aceptó (éxito directo).
  //   • null       → flag OFF, módulo ausente, datos no resolubles, sin impresora, o PrintNode rechazó → FALLBACK a GAS.
  // Solo se cablean las acciones autocontenidas en el front (o resolubles del cache). Las que el GAS resolvía de
  // Sheets / fan-out multi-impresora / lotes orquestados en server NO se cablean (ver REPORTE).
  // ════════════════════════════════════════════════════════════════════
  async function _postDirectoImpresion(params, lid) {
    // Solo se cablean acciones de impresión de UN job ESC/POS / TSPL autocontenido en el front (o resoluble del
    // cache). NO se cablean:
    //   • imprimirAvisoCajeros → fan-out CRUZA-DB: el GAS lee MosExpress.CAJAS (cajas abiertas) y manda 1 job por
    //     impresora de cajero. El navegador WH no puede leer las cajas de MosExpress → la IMPRESIÓN del aviso queda
    //     en GAS. La RAÍZ del bug del comentario "(vacío)" se cierra aparte, sin impresora (ver app.js
    //     _dispararAvisoCajeros: persiste snapshot directo a Supabase + manda los campos REALES al GAS para el ticket).
    //   • imprimirMembrete góndola ME / almacén WH y las etiquetas Caserito en LOTE → orquestadas por el sistema de
    //     LOTES (crearLoteMembrete/crearLoteAdhesivo: cola en Sheet + trigger que imprime sub-lotes con compensación
    //     de DRIFT térmico por-print desde Script Properties del rollo). Esa cola+drift NO existe en el navegador
    //     (impresion-directa.js usa offset base = drift 0) → se quedan en GAS. Los builders TSPL2 (Caserito/ME/WH)
    //     SÍ están listos en ImpresionDirecta para el día que la cola de lotes se porte a Supabase (Edge cron).
    // Lo cableado (single-job): imprimirBienvenida + imprimirMembrete estándar (ESC/POS) + imprimirEtiqueta
    // (Caserito TSPL2, ruta single-job legacy: drift 0 = offset base, válido para 1 etiqueta).
    const PRINT_ACTIONS = ['imprimirBienvenida', 'imprimirMembrete', 'imprimirEtiqueta'];
    if (PRINT_ACTIONS.indexOf(params.action) < 0) return undefined;   // no es impresión cableada
    if (!_whImpresionDirecta()) return null;                          // flag OFF → GAS
    if (typeof ImpresionDirecta === 'undefined') return null;         // módulo no cargado → GAS

    if (params.action === 'imprimirBienvenida') {
      if (!ImpresionDirecta.armarBienvenida) return null;
      let armado;
      try {
        armado = ImpresionDirecta.armarBienvenida({
          nombre: params.nombre, apellido: params.apellido, rol: params.rol,
          horaInicio: params.horaInicio, empresa: params.empresa
        });
      } catch (_) { return null; }
      if (!armado || !armado.base64) return null;
      const pid = _resolvePrinterId(armado.printerHint, params.printerIdOverride);
      if (!pid) return null;                                          // sin impresora resoluble → GAS
      const ok = await _imprimirDirecto(pid, armado.base64, armado.title);
      return ok ? { ok: true } : null;                               // PrintNode rechazó → GAS
    }

    if (params.action === 'imprimirMembrete') {
      if (!ImpresionDirecta.armarMembreteStd) return null;
      // El front manda idProducto + barcodes (JSON), NO nombre/sku (el GAS los leía de Sheets).
      // Resolver descripción + skuBase del cache de productos (mismo patrón que registrarEnvasado).
      const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
      const idp  = String(params.idProducto || '').trim();
      const prod = idp ? prods.find(p => String(p.idProducto) === idp || String(p.codigoBarra) === idp || String(p.skuBase) === idp) : null;
      const sku  = String((prod && (prod.skuBase || prod.idProducto)) || params.skuGrupal || idp || '').trim();
      if (!sku) return null;                                          // sin sku resoluble → GAS
      const nombre = String((prod && (prod.descripcion || prod.nombre)) || '').trim();
      let armado;
      try { armado = ImpresionDirecta.armarMembreteStd({ nombre, sku, barcodes: params.barcodes }); }
      catch (_) { return null; }
      if (!armado || !armado.base64) return null;
      const pid = _resolvePrinterId(armado.printerHint, params.printerIdOverride);
      if (!pid) return null;
      const ok = await _imprimirDirecto(pid, armado.base64, armado.title);
      return ok ? { ok: true } : null;
    }

    if (params.action === 'imprimirEtiqueta') {
      // Etiqueta Caserito/envasado (TSPL2) — ruta SINGLE-JOB (el GAS imprimirEtiqueta manda 1 job a PrintNode).
      // El smart-highlight necesita el catálogo de envasables tokenizado: el GAS lo lee de Sheets
      // (_getAllEnvasablesTokens); acá lo replicamos del cache de productos con el MISMO filtro (estado=1,
      // tiene codigoProductoBase, código empieza con 'WH'). Si el cache está vacío, allEnv=[] → sin highlight
      // diferenciador (el último token = peso igual se resalta), comportamiento degradado pero válido.
      if (!ImpresionDirecta.armarEtiquetaCaserito) return null;
      const allEnv = _envasablesTokensCache();
      let armado;
      try {
        armado = ImpresionDirecta.armarEtiquetaCaserito({
          codigoBarra:      params.codigoBarra,
          descripcion:      params.descripcion,
          unidades:         params.unidades,
          fechaVencimiento: params.fechaVencimiento,
          fechaEnvasado:    params.fechaEnvasado || params.fechaImpresion,
          allEnvasables:    allEnv
        });
      } catch (_) { return null; }
      if (!armado || !armado.base64) return null;
      const pid = _resolvePrinterId(armado.printerHint, params.printerIdOverride);
      if (!pid) return null;                                          // sin impresora ADHESIVO → GAS
      const ok = await _imprimirDirecto(pid, armado.base64, armado.title);
      return ok ? { ok: true } : null;
    }

    return null;
  }

  // Réplica del catálogo de _getAllEnvasablesTokens (GAS) desde el cache de productos del navegador, para el
  // smart-highlight de las etiquetas TSPL2. Mismo filtro: estado='1' + codigoProductoBase no vacío + código (WH...).
  function _envasablesTokensCache() {
    const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache) ? (OfflineManager.getProductosCache() || []) : [];
    const out = [];
    prods.forEach(p => {
      if (String(p.estado) !== '1') return;
      if (!p.codigoProductoBase) return;
      const c = String(p.codigoBarra || p.idProducto || '').toUpperCase();
      if (c.indexOf('WH') !== 0) return;
      out.push({ descripcion: p.descripcion });   // el builder tokeniza con _normalizeEtq (igual que el GAS)
    });
    return out;
  }

  // GET: network first, cache como fallback
  async function call(params) {
    // [0% GAS / 0% FALLBACK 2026-07-04] Lectura SIEMPRE directa a Supabase (sin gate, sin GAS). Si el directo
    // devuelve null/lanza (acción sin RPC directo o token minteándose), cae a la CACHÉ local offline — NUNCA a GAS.
    if (navigator.onLine) {
      try {
        const directo = await _callDirecto(params);
        if (directo) return directo;
      } catch (_) { /* → caché */ }
    }
    return _fromCache(params);
  }

  // Fallback GET desde caché offline
  function _fromCache(params) {
    const action = params.action;
    if (action === 'getProductos')   return { ok: true, data: OfflineManager.getProductosCache() };
    if (action === 'getStock')       return { ok: true, data: OfflineManager.getStockCache() };
    if (action === 'getProveedores') return { ok: true, data: OfflineManager.getProveedoresCache() };
    if (action === 'getPersonal')    return { ok: true, data: OfflineManager.getPersonalCache().map(p => { const s = {...p}; delete s.pin; return s; }) };
    // [423] tarifa de envasado (lectura mínima; fallback silencioso → el front usa 0.10)
    if (action === 'getTarifaEnvasado') {
      return _sbRpcWH('get_tarifa_envasado', {})
        .then(out => (out && out.ok) ? out : { ok: false })
        .catch(() => ({ ok: false }));
    }
    if (action === 'getConfig')      return { ok: true, data: OfflineManager.getConfigCache() };
    if (action === 'getGuias')       return { ok: true, data: OfflineManager.getGuiasCache() };
    if (action === 'getPreingresos') return { ok: true, data: OfflineManager.getPreingresosCache() };
    if (action === 'getGuia') {
      const guias    = OfflineManager.getGuiasCache();
      const detalles = OfflineManager.getGuiaDetalleCache();
      const prods    = OfflineManager.getProductosCache();
      const guia     = guias.find(g => g.idGuia === params.idGuia);
      if (!guia) return { ok: false, error: 'Guía no en caché' };
      const prodMap = {};
      prods.forEach(p => { prodMap[p.idProducto] = p.descripcion || p.nombre || p.idProducto; });
      const detalle = detalles
        .filter(d => d.idGuia === params.idGuia)
        .map(d => ({ ...d, descripcionProducto: prodMap[d.codigoProducto] || d.codigoProducto }));
      return { ok: true, data: { ...guia, detalle } };
    }
    return { ok: false, error: 'Sin conexión y sin caché disponible' };
  }

  // ── Contador de operaciones en vuelo ─────────────────────────
  // El cliente envía POSTs en paralelo (rápido). GAS LockService serializa
  // internamente para evitar race conditions. El cliente solo cuenta cuántas
  // están esperando respuesta para mostrar feedback visual.
  let _opsEnVuelo = 0;
  function _emitOpsState() {
    try {
      window.dispatchEvent(new CustomEvent('wh:opqueue', { detail: { count: _opsEnVuelo } }));
    } catch(e) {}
  }

  // ── Idempotencia: localId único por operación ────────────────
  // GAS guarda en SYNC_LOG el localId de cada POST procesado. Si llega un POST
  // con el mismo localId (reintento por timeout, doble click, etc.), GAS retorna
  // la respuesta cacheada sin re-ejecutar. Garantiza 0 duplicados.
  const _IDEMPOTENT_ACTIONS = new Set([
    'agregarDetalleGuia', 'actualizarCantidadDetalle', 'actualizarFechaVencimiento',
    'anularDetalle', 'cerrarGuia', 'reabrirGuia', 'crearGuia', 'crearDespachoRapido',
    'registrarEnvasado', 'corregirUnidadesEnvasado', 'anularEnvasadoConClave', 'anularEnvasadoManual',
    'registrarProductoNuevo',   // (aprobarProductoNuevo removido: la aprobación es de MOS, no de WH)
    'crearPreingreso', 'aprobarPreingreso', 'actualizarPreingreso',
    'subirFotoGuia', 'subirFotoPreingreso',   // [40x #1] localId estable → nombre de foto determinístico (idempotente)
    // (crear/actualizar producto y proveedor removidos: WH no escribe catálogo, solo lo usa)
    'crearAjuste', 'auditarProducto',
    'asignarAuditoria', 'ejecutarAuditoria',
    'marcarAlertaRevisada', 'aceptarTeoricoAlerta',
    'addCargadorDia', 'removeCargadorDia',   // [Tanda 3] dedup directo: id_log determinista (add) / -1 real (remove)
    'actualizarGuia', 'actualizarPickup', 'guardarProgresoPickup',
    'cerrarPickupConDespacho', 'liberarPickup', 'cerrarTurno'
  ]);
  function _genLocalId() {
    return 'L' + Date.now() + Math.random().toString(36).substr(2, 7);
  }

  // Timeout por intento — protege al UI de GAS colgado (ej. _notificarMOS lento).
  // Si pasa, abortamos y vamos al siguiente intento; si es el último, encolamos
  // offline (el localId garantiza idempotencia al sincronizar).
  const _FETCH_TIMEOUT_MS = 15000;

  async function _doFetchWithRetry(GAS_URL, params) {
    // 3 intentos: rápido si todo OK, recuperación si red flaquea o lock saturado.
    // Backoff: 600ms, 1500ms (total ~2.1s antes de rendirse).
    // params.localId se mantiene constante en los reintentos → GAS deduplica.
    const delays = [600, 1500];
    for (let intento = 0; intento < 3; intento++) {
      const ctrl = new AbortController();
      const tId  = setTimeout(() => ctrl.abort(), _FETCH_TIMEOUT_MS);
      try {
        const res  = await fetch(GAS_URL, {
          method:  'POST',
          headers: { 'Content-Type': 'text/plain' },
          body:    JSON.stringify(params),
          signal:  ctrl.signal
        });
        clearTimeout(tId);
        const json = await res.json();
        if (json && json.ok === false && /saturado|ocupado/i.test(json.error || '')) {
          if (intento < 2) { await new Promise(r => setTimeout(r, delays[intento])); continue; }
        }
        return json;
      } catch {
        clearTimeout(tId);
        if (intento < 2) { await new Promise(r => setTimeout(r, delays[intento])); continue; }
        // Red falló o timeout tras 3 intentos → encolar offline (reusa localId si ya existe).
        // [40x cruce] Si esta llamada YA viene de la cola (_fromQueue), NO re-encolar: devolver fallo
        // para que sincronizar() lo marque 'error' y reintente luego (evita encolado doble/loop).
        if (params._fromQueue) return { ok: false, error: 'red', _retry: true };
        const localId = OfflineManager.encolar(params.action, params);
        return { ok: true, offline: true, localId, data: { idLocal: localId } };
      }
    }
  }

  // POST: si offline → encolar offline. GAS LockService maneja la serialización.
  async function post(params) {
    const GAS_URL = _gasUrl();
    const idSesion = window.WH_CONFIG?.idSesion;
    if (idSesion && !params.idSesion) params = { ...params, idSesion };

    // [CERO-GAS F2] Prints de reporte admin → Edge `imprimir` (antes GAS). Historial: el front ya arma `texto`.
    //   Cargadores: se arma el ESC/POS client-side. Sin GAS.
    if (params.action === 'imprimirHistorialStock') {
      return await _imprimirTextoTicketWH(params.texto || '', 'Historial ' + (params.codigoProducto || ''), params.printerIdOverride);
    }
    if (params.action === 'imprimirCargadoresDia') {
      // [FIX 500x] _armarCargadoresEscPos espera el shape RICO (carretasTotal/llenasTotal/mediasTotal/vaciasTotal +
      // preingresos[]). `resumen_cargadores_dia` solo trae {fecha,total,cargadores:[{nombre,count}]} → el ticket
      // salía en 0/blanco. La data rica la produce `getCargadoresDelDia` (lee wh.preingresos y consolida carretas).
      const rd = await call({ action: 'getCargadoresDelDia', fecha: String(params.fecha || '') });
      const d = (rd && rd.data) ? rd.data : null;
      if (!d) return { ok: false, error: 'No se pudo leer cargadores del día' };
      return await _imprimirTextoTicketWH(_armarCargadoresEscPos(d), 'Cargadores ' + (d.fecha || ''), params.printerIdOverride);
    }

    // Inyectar localId para acciones idempotentes (si no viene ya)
    if (_IDEMPOTENT_ACTIONS.has(params.action) && !params.localId) {
      params = { ...params, localId: _genLocalId() };
    }

    // [PASO 5 · B4] escritura directa a Supabase (inerte mientras el flag esté OFF).
    // [40x cruce] Distinción CRÍTICA para no duplicar stock en el cruce de fallback:
    //   • _postDirecto devuelve un objeto → ÉXITO directo, listo.
    //   • _postDirecto devuelve null → la RPC respondió *_OFF / error donde NO commiteó → caer a GAS es SEGURO.
    //   • _postDirecto LANZA (timeout/red) → la RPC PUDO commitear y perderse la respuesta. Caer a GAS DUPLICARÍA
    //     (GAS dedupea por su SYNC_LOG, que Supabase no comparte). → NO ir a GAS: encolar para reintento DIRECTO
    //     (la cola reintenta vía post() directo, idempotente por el id sembrado del localId → la RPC dedupea).
    // [PASO 5 · B5] Entra a _postDirecto si CUALQUIERA de los dos flags está activo: escritura directa
    // (RPCs) o impresión directa (Edge `imprimir`). Cada bloque dentro revalida su propio flag, así que
    // con solo el de impresión ON, las RPCs de escritura NO corren (devuelven null → GAS, sin tocar Supabase).
    // [100x rollback-fix] TAMBIÉN entra si el ÍTEM nació de escritura directa (`_viaDirecta`, sellado al
    // encolar por timeout). Un ítem así PUDO commitear en Supabase → SIEMPRE debe reintentarse directo
    // (la RPC dedupea por el id sembrado del localId), aunque el flag global ya esté apagado por rollback.
    // Si fuera a GAS, GAS no tiene su localId en SYNC_LOG → lo ejecutaría → DOBLE STOCK / DOBLE GUÍA.
    if ((_whEscrituraDirecta() || _whImpresionDirecta() || _whLoteAdhesivoDirecto() || params._viaDirecta) && navigator.onLine) {
      _opsEnVuelo++; _emitOpsState();
      let timeoutDirecto = false, rechazoPermanente = false;
      // [FIX stale-tras-escritura] Un éxito directo MUTÓ datos en Supabase → invalidar el
      // micro-cache de lecturas ANTES de devolver, para que el refresh inmediato del front
      // (precargarOperacional(true)/getStock/getDashboard) lea FRESCO y el operador vea su
      // cambio al instante (guía/ajuste/envasado/merma/preingreso…), no la versión de hace <4s.
      try { const d = await _postDirecto(params); if (d) { _invalidarLecturas(); return d; } }
      catch (e) { timeoutDirecto = true; rechazoPermanente = !!(e && e.permanente); }
      finally { _opsEnVuelo--; _emitOpsState(); }
      // [100x rollback-fix] Si llegamos acá con un ítem `_viaDirecta`, _postDirecto NO devolvió éxito
      // (devolvió null o lanzó). Ir a GAS NUNCA es seguro para estos ítems (duplicaría: GAS no comparte el
      // SYNC_LOG de Supabase). Lo dejamos en la cola para reintento DIRECTO posterior (sincronizar() lo verá
      // 'error' por el _retry y volverá a llamar API._postCola → _postDirecto). Nunca cae al GAS de abajo.
      // [400-loop fix] EXCEPTO si el servidor lo RECHAZÓ con un 4xx definitivo (no commiteó): reintentarlo
      // eternamente solo spamea la consola. _descartar:true → la cola lo da por terminado sin mandarlo a GAS.
      if (params._viaDirecta) {
        if (rechazoPermanente) return { ok: false, error: 'rechazo-directo', _descartar: true };
        return { ok: false, error: timeoutDirecto ? 'timeout-directo' : 'no-routeable-directo', _retry: true };
      }
      // [v2.13.240 · 400-loop fix] Rechazo PERMANENTE en el PRIMER intento (no `_viaDirecta`): un 4xx definitivo
      // de Storage (RLS/mime/path) NO commiteó nada (no hay riesgo de duplicado vía GAS, a diferencia de las RPC
      // de stock). Encolarlo solo gastaría un reintento garantizado a fallar. Devolvemos el fallo limpio ya: el UI
      // muestra "Error al subir foto, reintenta" en vez de un fantasma. NUNCA cae al fallback GAS para guías directas.
      if (rechazoPermanente) return { ok: false, error: 'rechazo-directo', _descartar: true };
      if (timeoutDirecto) {
        // [FIX stale-tras-escritura] El timeout NO garantiza que NO commiteó: la RPC PUDO
        // aplicar el cambio y perderse la respuesta. Invalidar el micro-cache para que el
        // refresh del front no sirva datos de hace <4s que omitan ese posible cambio.
        _invalidarLecturas();
        if (params._fromQueue) return { ok: false, error: 'timeout-directo', _retry: true };
        // [100x rollback-fix] La RPC PUDO commitear en Supabase y perderse la respuesta.
        // Sellamos la VÍA de reintento POR ÍTEM (no por flag global al sincronizar): este
        // item DEBE reintentarse SIEMPRE vía _postDirecto (idempotente por el id sembrado del
        // localId → la RPC dedupea), aunque la fase de escritura directa se apague (rollback,
        // limpieza de localStorage, cambio de dispositivo) antes de que la cola se vacíe.
        // Si fuera a GAS, GAS no tiene ese localId en su SYNC_LOG → lo ejecutaría → DOBLE STOCK.
        const localId = OfflineManager.encolar(params.action, { ...params, _viaDirecta: true });
        return { ok: true, offline: true, localId, data: { idLocal: localId } };
      }
      // (si llegamos acá, _postDirecto devolvió null → seguir abajo)
    }

    // [cero-GAS · WH] Acción propia de WH cuyo directo no commiteó → FAIL-CLOSED: ni GAS ni encolar-a-GAS.
    // Error claro para el caller (el panel de diagnóstico / la rama muerta de mermas). Nunca toca gasUrl.
    if (_WH_NO_GAS.has(params.action)) {
      return { ok: false, error: 'Acción no disponible sin conexión directa (cero-GAS)', _ceroGas: true };
    }

    if (!GAS_URL || !navigator.onLine) {
      if (params._fromQueue) return { ok: false, error: 'sin-conexion', _retry: true };
      // [FIX pre-corte GAS · auditoría 2026-07-08] Encolar OFFLINE sella la VÍA DIRECTA (si la app
      // opera en escritura directa, que en prod es SIEMPRE): sin el sello, sincronizar() replayaba
      // estos ítems a GAS al reconectar → escrituras de dinero/stock hechas sin señal MORÍAN con el
      // corte de GAS (quedaban 'error' eternos con el operador creyendo que se guardó). Con el sello,
      // el replay va SIEMPRE por API._postCola → _postDirecto (idempotente por el id del localId).
      // Acciones no cableadas quedan 'error' visibles en cola — jamás van a GAS.
      const _sellado = (_whEscrituraDirecta() || _whImpresionDirecta() || _whLoteAdhesivoDirecto())
        ? { ...params, _viaDirecta: true } : params;
      const localId = OfflineManager.encolar(params.action, _sellado);
      return { ok: true, offline: true, localId, data: { idLocal: localId } };
    }

    _opsEnVuelo++;
    _emitOpsState();
    try {
      return await _doFetchWithRetry(GAS_URL, params);
    } finally {
      _opsEnVuelo--;
      _emitOpsState();
    }
  }

  // ════════════════════════════════════════════════════════════════════
  // [Realtime catálogo] Suscripción a mos.catalogo_meta (UPDATE) por WebSocket.
  // ADITIVO: NO reemplaza al poller de ~50s (offline.js). Su único trabajo es bajar
  // la latencia de propagación del catálogo de ~50s a ~0s cuando MOS cambia productos/
  // proveedores. Si el cliente realtime no carga o el WS cae, NO rompe nada: el poller
  // sigue siendo la red de seguridad.
  //
  // Cliente: @supabase/realtime-js (ESM por CDN) — maneja heartbeat/reconnect/auth.
  //   • Singleton + guard anti-reentrada: NUNCA abre dos conexiones/canales.
  //   • access_token = el MISMO JWT que mintea WH (mint-wh, role authenticated). Al rotar
  //     (~30min) re-aplicamos con setAuth() y, defensivo, re-suscribimos.
  //   • On UPDATE: payload.new.version (o .record.version) → OfflineManager.notificarVersionCatalogo
  //     (compara vs baseline + re-descarga money-safe vía precargar('manual'), igual que el poller).
  //   • Re-chequeo on visible/focus/online por si el WS perdió un evento mientras dormía.
  //   • Cierre limpio en logout (detenerRealtimeCatalogo).
  // CARGA DEFENSIVA: todo el arranque va en try/catch async; cualquier excepción se traga
  // (log y salir) → la app y el poller siguen intactos.
  // ════════════════════════════════════════════════════════════════════
  const _RT = {
    client:    null,   // RealtimeClient
    channel:   null,   // canal de postgres_changes
    starting:  false,  // guard anti-reentrada del arranque
    started:   false,  // hay (o hubo) un intento exitoso de canal
    libPromise:null,   // promesa de import del cliente (1 sola vez)
    listeners: false,  // listeners visible/focus/online ya cableados
    lastResync:0,      // [FIX #5/#6] timestamp del último resync por SUBSCRIBED (throttle anti-flapping)
    gen:       0       // [anti-orphan] generación: _detener la incrementa → un arranque en vuelo se aborta tras sus awaits
  };

  function _rtImportarLib() {
    if (_RT.libPromise) return _RT.libPromise;
    // ESM por CDN. +esm fuerza el bundle ESM de jsDelivr. Si falla → null (poller fallback).
    _RT.libPromise = import('https://cdn.jsdelivr.net/npm/@supabase/realtime-js@2/+esm')
      .then(mod => (mod && (mod.RealtimeClient || (mod.default && mod.default.RealtimeClient))) || null)
      .catch(err => { try { console.warn('[Realtime] import falló (poller sigue):', err); } catch (_) {} return null; });
    return _RT.libPromise;
  }

  // Extrae la versión nueva del payload de postgres_changes (cubre variantes de shape).
  function _rtVersionDePayload(payload) {
    try {
      const rec = (payload && (payload.new || (payload.data && payload.data.record))) || null;
      if (!rec) return null;
      const v = Number(rec.version);
      return Number.isFinite(v) ? v : null;
    } catch (_) { return null; }
  }

  // [FIX #5/#6 2026-06-26] DEBOUNCE anti-loop. Un burst de UPDATEs de mos.catalogo_meta
  // (MOS bumpeando la versión repetidamente) o un WS flapeando (CLOSED↔SUBSCRIBED) disparaba
  // una RE-DESCARGA del maestro POR CADA evento → loop de `descargarMaestros` que satura el
  // main thread (la pantalla de bloqueo laguea, los efectos se congelan, el giro crashea).
  // Coalescemos: nos quedamos con la versión MÁS ALTA del burst y disparamos UNA sola descarga
  // (trailing edge, 1.5s). El poller de ~50s sigue como red de seguridad.
  let _rtNotifTimer = null, _rtNotifPend = null, _rtNotifMotivo = '';
  function _rtNotificar(v, motivo) {
    if (v == null) return;
    _rtNotifPend = (_rtNotifPend == null) ? v : Math.max(_rtNotifPend, v);
    _rtNotifMotivo = motivo || 'realtime';
    if (_rtNotifTimer) return;                 // ya hay un disparo trailing programado → coalescer
    _rtNotifTimer = setTimeout(() => {
      const vv = _rtNotifPend, mm = _rtNotifMotivo;
      _rtNotifTimer = null; _rtNotifPend = null;
      try {
        if (typeof OfflineManager !== 'undefined' && OfflineManager.notificarVersionCatalogo) {
          OfflineManager.notificarVersionCatalogo(vv, mm);
        }
      } catch (_) {}
    }, 1500);
  }

  async function _iniciarRealtimeCatalogo() {
    // Guards: una sola conexión, solo navegador con red.
    if (typeof window === 'undefined') return;
    if (_RT.starting || _RT.channel) return;            // singleton + anti-reentrada
    if (!navigator.onLine) return;                      // sin red no hay WS; el poller cubrirá
    _RT.starting = true;
    const _gen = _RT.gen;   // [anti-orphan] si un logout (detener) ocurre durante los awaits, _RT.gen cambia → abortamos
    try {
      const RealtimeClient = await _rtImportarLib();
      if (!RealtimeClient) return;                       // lib no cargó → poller fallback
      // Re-chequeo: otra llamada pudo crear el canal mientras importábamos, o un logout cerró el canal.
      if (_RT.channel || _gen !== _RT.gen) return;

      const token = await _mintTokenWH().catch(() => null);  // JWT authenticated (mismo que las RPC)
      if (!token) return;                                     // sin token no hay canal; el poller cubre
      if (_RT.channel || _gen !== _RT.gen) return;             // logout/otra apertura durante el mint → abortar

      // URL WebSocket explícita (wss://<ref>.supabase.co/realtime/v1) — coincide con el
      // endpoint verificado. apikey (anon, público) va en params; el access_token (JWT
      // authenticated) lo aplica setAuth → viaja en el phx_join, igual que supabase-js.
      const wsUrl = _SB_URL.replace(/^http/i, 'ws') + '/realtime/v1';
      const client = new RealtimeClient(wsUrl, {
        params: { apikey: _SB_ANON }
      });
      try { client.setAuth(token); } catch (_) {}
      _RT.client = client;

      const channel = client.channel('wh-catalogo-meta');
      channel.on('postgres_changes',
        { event: 'UPDATE', schema: 'mos', table: 'catalogo_meta' },
        (payload) => { _rtNotificar(_rtVersionDePayload(payload), 'realtime'); }
      );
      // [v2.13.344] Realtime de la LISTA DE PICKUP: wh.ops_meta dominio 'pickups' sube cada vez
      // que cambia wh.pickups (cierre de caja, consolidación, despacho). La vista de despacho
      // escucha 'wh:pickups-realtime' y re-pollea al instante (antes esperaba el poller de 30s).
      channel.on('postgres_changes',
        { event: 'UPDATE', schema: 'wh', table: 'ops_meta' },
        (payload) => {
          try {
            const rec = (payload && (payload.new || (payload.data && payload.data.record))) || {};
            if (String(rec.dominio || '') === 'pickups') {
              window.dispatchEvent(new CustomEvent('wh:pickups-realtime'));
            }
          } catch (_) {}
        }
      );
      channel.subscribe((status) => {
        try { console.log('[Realtime] canal catalogo_meta:', status); } catch (_) {}
        // Al (re)suscribir, leer la versión actual del catálogo y notificarla por si
        // perdimos un UPDATE mientras el WS estaba caído/dormido. Money-safe: notificar
        // pasa por el mismo núcleo que el poller (no re-descarga si la versión no subió).
        // El poller de ~50s igual lo cubriría; esto solo acelera la convergencia.
        if (status === 'SUBSCRIBED') {
          // [FIX #5/#6 2026-06-26] Si el WS flapea, throttle: máx 1 resync cada 15s → evita
          // disparar N RPCs de versión + N descargas (el debounce de _rtNotificar ya coalesce,
          // pero esto ahorra las RPCs redundantes). El poller de ~50s cubre el resto.
          const now = Date.now();
          if (now - (_RT.lastResync || 0) < 15000) return;
          _RT.lastResync = now;
          _catalogoVersion().then(v => _rtNotificar(v, 'realtime-resync')).catch(() => {});
        }
      });
      _RT.channel = channel;
      _RT.started = true;

      // [PRESENCIA ecosistema] Canal presence compartido 'ecos-presencia' (mismo WS).
      // WH se ANUNCIA {deviceId, nombre, rol, app:'WH'} → el panel MOS (Configuración)
      // muestra este equipo/operador EN LÍNEA AL SEGUNDO (pisa el heartbeat de 10 min).
      // Best-effort: si falla, nada cambia (el heartbeat sigue siendo la base).
      try {
        const devId = String(_whDeviceId() || '').trim();
        if (devId) {
          const usuario = String((window.WH_CONFIG && window.WH_CONFIG.usuario) || localStorage.getItem('wh_usuario') || '').trim();
          const chP = client.channel('ecos-presencia', { config: { presence: { key: devId } } });
          chP.subscribe((status) => {
            try { console.log('[Realtime] canal ecos-presencia:', status); } catch (_) {}
            if (status === 'SUBSCRIBED') {
              try { chP.track({ deviceId: devId, nombre: usuario, rol: '', app: 'WH' }); } catch (_) {}
            }
          });
          _RT.chPres = chP;
        }
      } catch (e) { try { console.warn('[Realtime] presencia no abrió (heartbeat sigue):', e); } catch (_) {} }

      _rtCablearListeners();
    } catch (err) {
      try { console.warn('[Realtime] arranque falló (poller sigue):', err); } catch (_) {}
    } finally {
      _RT.starting = false;
    }
  }

  // Re-aplica el token (rotación ~30min) al cliente realtime. Fire-and-forget.
  async function _rtRefrescarToken() {
    if (!_RT.client) return;
    try {
      const token = await _mintTokenWH().catch(() => null);
      if (token && _RT.client && _RT.client.setAuth) _RT.client.setAuth(token);
    } catch (_) {}
  }

  function _rtCablearListeners() {
    if (_RT.listeners || typeof window === 'undefined') return;
    _RT.listeners = true;
    // Volver a primer plano / recuperar foco / reconectar: re-asegurar canal + token frescos.
    const reasegurar = () => {
      if (!navigator.onLine) return;
      if (!_RT.channel) { _iniciarRealtimeCatalogo(); return; }   // canal cerrado/caído → re-abrir
      _rtRefrescarToken();                                        // canal vivo → solo refrescar token
    };
    try {
      document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reasegurar(); });
      window.addEventListener('focus',  reasegurar);
      window.addEventListener('online', reasegurar);
    } catch (_) {}
  }

  function _detenerRealtimeCatalogo() {
    _RT.gen++;   // [anti-orphan] invalida cualquier arranque en vuelo (post-await abortará en vez de abrir un canal huérfano)
    try { if (_RT.channel && _RT.client && _RT.client.removeChannel) _RT.client.removeChannel(_RT.channel); } catch (_) {}
    try { if (_RT.chPres && _RT.client && _RT.client.removeChannel) _RT.client.removeChannel(_RT.chPres); } catch (_) {}
    try { if (_RT.client && _RT.client.disconnect) _RT.client.disconnect(); } catch (_) {}
    _RT.channel = null;
    _RT.chPres  = null;
    _RT.client  = null;
    _RT.started = false;
    // [Review R1 fix] limpiar el debounce trailing de _rtNotificar (no dejar un fire-and-forget colgando tras
    // el logout/cierre de turno) + resetear el throttle del resync para que el próximo login no quede gateado.
    if (_rtNotifTimer) { clearTimeout(_rtNotifTimer); _rtNotifTimer = null; }
    _rtNotifPend = null; _rtNotifMotivo = '';
    _RT.lastResync = 0;
  }

  return {
    // Descarga maestros MOS → localStorage
    descargarMaestros:    ()     => call({ action: 'descargarMaestros' }),
    // [perf 500x · CD1] Descarga DELTA (solo productos cambiados desde desdeTs). Sin fallback GAS: si falla,
    // el caller (offline.js) cae a descargarMaestros() completo (que sí tiene fallback GAS).
    descargarMaestrosDelta: (desdeTs) => _sbDescargarMaestrosDelta(desdeTs),
    // [poller catálogo] versión monotónica del maestro (mos.catalogo_version, profile 'mos').
    // Devuelve Number; LANZA ante fallo (el poller lo captura y no toca el baseline).
    catalogoVersion:      ()     => _catalogoVersion(),
    // [Realtime catálogo] Suscripción WebSocket a mos.catalogo_meta (UPDATE) → propagación ~0s.
    // ADITIVA: el poller de ~50s sigue como fallback. Singleton + carga defensiva (no rompe si falla).
    iniciarRealtimeCatalogo: ()  => _iniciarRealtimeCatalogo(),
    detenerRealtimeCatalogo: ()  => _detenerRealtimeCatalogo(),
    // [BUG A · cutover] Si el dispositivo escribe Y/O lee directo a Supabase, el operacional
    // (que alimenta el listado de Guías) DEBE leerse directo — si no, nunca vería sus propias
    // guías directas 'G_L...' a través del GAS (cache stale). Ante cualquier fallo → GAS.
    descargarOperacional: async () => {
      if ((_whLecturaDirecta() || _whEscrituraDirecta()) && navigator.onLine) {
        try { return await _descargarOperacionalDirecto(); }
        catch (_) { /* cae a GAS abajo */ }
      }
      return call({ action: 'descargarOperacional' });
    },

    // Dashboard
    getDashboard:       ()       => call({ action: 'getDashboard' }),

    // [PASO 5 · B5] Impresión directa vía Edge (el módulo arma el content base64). Devuelve bool (éxito PrintNode).
    imprimirDirecto:    (printerId, content, title) => _imprimirDirecto(printerId, content, title),
    escposB64:          (str) => _escposB64(str),   // string ESC/POS/ZPL → base64 (= Utilities.base64Encode de GAS)
    // [PASO 5 · B5] Subir foto a Storage (máxima calidad). Devuelve {url, preview, path}. Para subirFotoGuia/Preingreso/etc.
    subirFotoStorage:   (tipo, id, base64, mime) => _subirFotoStorage(tipo, id, base64, mime),
    // [Pregúntale a tu almacén] Acceso directo a la Edge `ia` (proxy a Claude). body={messages, system?, model?, max_tokens?, tools?}.
    // READ-ONLY: el chat de almacén lo usa para el loop tool-use. Devuelve el JSON de Claude tal cual.
    llamarEdgeIA:       (body) => _llamarEdgeIA(body),

    // Productos
    getProductos:       (p={})   => call({ action: 'getProductos', ...p }),
    getProducto:        (cod)    => call({ action: 'getProducto', codigo: cod }),
    getStock:           (p={})   => call({ action: 'getStock', ...p }),
    getStockProducto:   (cod)    => call({ action: 'getStockProducto', codigo: cod }),
    // [Stock teórico/proyectado] real + Σ(ingresos abiertos) − Σ(salidas abiertos), DERIVADO al vuelo (no persiste).
    getStockProyectado: (p={})   => call({ action: 'getStockProyectado', ...p }),
    getLotes:           (p={})   => call({ action: 'getLotesVencimiento', ...p }),
    // [WH nivel-inferior] crear/editar PRODUCTO NO existe en WH: el catálogo se crea/edita SOLO en MOS
    // (admin) y se propaga. WH lo USA, no lo escribe. (Funciones muertas removidas — antes iban a GAS/Hoja.)

    // Guias
    getGuias:           (p={})   => call({ action: 'getGuias', ...p }),
    getGuia:            (id)     => call({ action: 'getGuia', idGuia: id }),
    crearGuia:          (p)      => post({ action: 'crearGuia', ...p }),
    crearDespachoRapido:(p)     => post({ action: 'crearDespachoRapido', ...p }),
    // [v2.13.8] IA — parser de listas pegadas para Despacho Rápido (lista sombra)
    analizarListaSombra:(p)     => post({ action: 'analizarListaSombra', ...p }),
    // [v2.13.15] Listas sombra compartidas
    crearListaSombra:             (p) => post({ action: 'crearListaSombra', ...p }),
    getListasSombra:              (p={}) => call({ action: 'getListasSombra', ...p }),
    tomarListaSombra:             (p) => post({ action: 'tomarListaSombra', ...p }),
    liberarListaSombra:           (p) => post({ action: 'liberarListaSombra', ...p }),
    actualizarProgresoListaSombra:(p) => post({ action: 'actualizarProgresoListaSombra', ...p }),
    cerrarListaSombra:            (p) => post({ action: 'cerrarListaSombra', ...p }),
    anularListaSombra:            (p) => post({ action: 'anularListaSombra', ...p }),
    getPickups:         (p={})  => call({ action: 'getPickups', ...p }),
    actualizarPickup:   (p)     => post({ action: 'actualizarPickup', ...p }),
    guardarProgresoPickup:    (p) => post({ action: 'guardarProgresoPickup',    ...p }),
    cerrarPickupConDespacho:  (p) => post({ action: 'cerrarPickupConDespacho',  ...p }),
    liberarPickup:            (p) => post({ action: 'liberarPickup',            ...p }),
    getPickup:                (p) => call({ action: 'getPickup', ...p }),
    agregarDetalle:           (p) => post({ action: 'agregarDetalleGuia',        ...p }),
    actualizarFechaVencimiento: (p) => post({ action: 'actualizarFechaVencimiento',  ...p }),
    actualizarCantidadDetalle:  (p) => post({ action: 'actualizarCantidadDetalle',   ...p }),
    // [v2.13.53] Lotes — FIFO + historial trazable
    getLotesFIFO:               (p) => call({ action: 'getLotesFIFO',                ...p }),
    getHistorialLote:           (p) => call({ action: 'getHistorialLote',            ...p }),
    cerrarGuia:         (id, u)  => post({ action: 'cerrarGuia', idGuia: id, usuario: u }),

    // Preingresos
    getPreingresos:           (p={}) => call({ action: 'getPreingresos', ...p }),
    crearPreingreso:          (p)    => post({ action: 'crearPreingreso', ...p }),
    aprobarPreingreso:        (p)    => post({ action: 'aprobarPreingreso', ...p }),
    subirFotoPreingreso:      (p)    => post({ action: 'subirFotoPreingreso', ...p }),
    actualizarFotosPreingreso:(p)    => post({ action: 'actualizarFotosPreingreso', ...p }),
    actualizarPreingreso:     (p)    => post({ action: 'actualizarPreingreso', ...p }),
    eliminarFotoDrive:        (p)    => post({ action: 'eliminarFotoDrive', ...p }),

    // Guías — foto + comentario
    subirFotoGuia:        (p)    => post({ action: 'subirFotoGuia',        ...p }),
    actualizarGuia:       (p)    => post({ action: 'actualizarGuia',       ...p }),
    copiarFotoDePreingreso:(p)   => post({ action: 'copiarFotoDePreingreso',...p }),

    // Envasados
    getEnvasados:       (p={})   => call({ action: 'getEnvasados', ...p }),
    getPendientes:      ()       => call({ action: 'getPendientesEnvasado' }),
    registrarEnvasado:  (p)      => post({ action: 'registrarEnvasado', ...p }),
    anularEnvasadoManual: (p)    => post({ action: 'anularEnvasadoManual', ...p }),
    corregirUnidadesEnvasado: (p) => post({ action: 'corregirUnidadesEnvasado', ...p }),
    anularEnvasadoConClave:   (p) => post({ action: 'anularEnvasadoConClave',   ...p }),

    // Mermas
    getMermas:          (p={})   => call({ action: 'getMermas', ...p }),
    // ── [🎯 Sorpresas + ♻️ Mermas v2 · SQL 516/517] directo a wh.* (cero-GAS) ──
    sorpresasLista:     (p={})   => _sbRpcWH('sorpresas_lista', { p }),
    registrarSorpresa:  (p)      => _sbRpcWH('registrar_sorpresa', { p }),
    mermasV2Lista:      ()       => _sbRpcWH('mermas_lista', { p: { alcance: 'wh' } }),
    vencimientosLista:  ()       => _sbRpcWH('vencimientos_lista', { p: {} }),   // [526] semáforo server-side unificado WH+MOS
    procesarMerma:      (p)      => _sbRpcWH('procesar_merma', { p }),
    mermasEliminarBatch:(p)      => _sbRpcWH('mermas_eliminar_batch', { p }),
    mermaDesdeGuia:     async (p) => {
      if (p && p.fotoBase64) {
        try { p = { ...p, foto: (await _subirFotoStorage('mermas', 'MRM_' + (p.id_merma || Date.now()), p.fotoBase64, p.mimeType || 'image/jpeg', p.id_merma || '')).url }; delete p.fotoBase64; }
        catch (_) { return { ok: false, error: 'FOTO_UPLOAD' }; }
      }
      return _sbRpcWH('merma_desde_guia', { p });
    },
    mermaAltaManualV2:  async (p) => {
      if (p && p.fotoBase64) {
        try { p = { ...p, foto: (await _subirFotoStorage('mermas', 'MRM_' + (p.id_merma || Date.now()), p.fotoBase64, p.mimeType || 'image/jpeg', p.id_merma || '')).url }; delete p.fotoBase64; }
        catch (_) { return { ok: false, error: 'FOTO_UPLOAD' }; }
      }
      return _sbRpcWH('merma_alta_manual', { p });
    },

    // Auditorias
    getAuditorias:      (p={})   => call({ action: 'getAuditorias', ...p }),
    asignarAuditoria:   (p)      => post({ action: 'asignarAuditoria', ...p }),
    ejecutarAuditoria:  (p)      => post({ action: 'ejecutarAuditoria', ...p }),

    // Ajustes
    getAjustes:         (p={})   => call({ action: 'getAjustes', ...p }),
    crearAjuste:        (p)      => post({ action: 'crearAjuste', ...p }),

    // [v2.13.310] Edge print-adhesivo genérico (membretes vía edgeCall del modal compartido)
    printAdhesivoEdge:  (body)    => _printAdhesivoEdge(body),

    // Proveedores — [WH nivel-inferior] WH solo LEE proveedores (los crea/edita el admin en MOS y se
    // propagan). crear/actualizar proveedor removidos (eran código muerto que iba a GAS/Hoja).
    getProveedores:     (p={})   => call({ action: 'getProveedores', ...p }),

    // Producto Nuevo — WH solo EMITE el PN (registrarPN); la APROBACIÓN es de MOS (aprobarPN removido).
    getProductosNuevos:          (p={}) => call({ action: 'getProductosNuevos', ...p }),
    getProductosNuevosRecientes: (p={}) => call({ action: 'getProductosNuevosRecientes', ...p }),
    registrarPN:                 (p)    => post({ action: 'registrarProductoNuevo', ...p }),

    // Config
    getConfig:          ()                  => call({ action: 'getConfig' }),
    setConfig:          (clave, valor)      => post({ action: 'setConfigValue', clave, valor }),

    // Etiquetas / Tickets
    imprimirEtiqueta:   (p)      => post({ action: 'imprimirEtiqueta', ...p }),
    imprimirMembrete:   (p)      => post({ action: 'imprimirMembrete', ...p }),
    imprimirTicketGuia: async (p) => {
      p = p || {};
      // [3x dedup ATÓMICO] Reserva en Supabase ANTES de imprimir (wh.reservar_ticket,
      // FOR UPDATE) — salvo copia manual explícita (fuerzaCopia). 3 llamadas paralelas
      // → solo 1 reserva (primera=true) → solo 1 imprime. Fail-open: si la RPC falla,
      // imprime igual (nunca bloquea la impresión por un fallo de dedup).
      if (!p.fuerzaCopia && p.idGuia) {
        try {
          const rsv = await _sbRpcWH('reservar_ticket', { p: { id_guia: String(p.idGuia), usuario: p.usuario || '' } });
          if (rsv && rsv.ok === true && rsv.primera === false) {
            return { ok: true, ya_impresa: true, dedup: true, data: { idGuia: p.idGuia } };
          }
        } catch (_) { /* fail-open → imprime */ }
      }
      // [100% SUPABASE · CERO GAS] El ticket de guía imprime EXCLUSIVAMENTE por la Edge `ticket-guia` (lee la guía
      // de Postgres, arma el ESC/POS validado BYTE-IDÉNTICO al viejo GAS, y manda a PrintNode con el secret).
      // printerId por mos.config WH_TICKET_PRINTER_ID (ID EXPLÍCITO, nunca por nombre → no confunde la XP-80C
      // duplicada de ME). Se ELIMINÓ todo fallback a GAS: el GAS→PrintNode era lo que DUPLICABA el ticket (su dedup
      // por Hoja era frágil + el front reintentaba al timeoutear → 2 jobs). Ahora el único dedup es wh.reservar_ticket
      // (atómico FOR UPDATE, evaluado arriba salvo fuerzaCopia) → doble imposible.
      if (navigator.onLine && p.idGuia) {
        let _token = null;
        try { _token = await _mintTokenWH(); } catch (_) { _token = null; }
        if (!_token) return { ok: false, error: 'No se pudo obtener token de impresión' };
        try {
          const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/ticket-guia`, {
            method: 'POST',
            headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + _token, 'Content-Type': 'application/json' },
            // reporteUrl → la Edge imprime el QR del reporte (sin esto el ticket sale sin QR).
            body: JSON.stringify({ idGuia: String(p.idGuia), reporteUrl: p.reporteUrl || '', printerIdOverride: p.printerIdOverride || null, fuerzaCopia: !!p.fuerzaCopia, usuario: p.usuario || '' })
          }, 15000);
          const d = await res.json().catch(() => ({}));
          if (d && d.ok) return { ok: true, data: { idGuia: p.idGuia, via: 'edge-ticket', jobId: d.jobId, detallesImpresos: d.detallesImpresos } };
          // Respuesta DEFINITIVA ok:false (auth/printerId/guía/PrintNode rechazó): el ticket NO salió → error claro.
          return { ok: false, error: (d && (d.error || d.mensaje)) || 'No se pudo imprimir el ticket (Supabase)' };
        } catch (_) {
          // [ANTI DOBLE] El fetch lanzó (timeout 15s / red): el job PUDO entrar a PrintNode → NO reimprimir.
          // Optimista; si no salió, "🖨 imprimir copia" (fuerzaCopia salta el dedup) lo reimprime. _imprimirConReintento
          // trata este caso como "posiblemente impreso" y NO reintenta → cero duplicados.
          return { ok: true, data: { idGuia: p.idGuia, via: 'edge-ticket', incierto: true } };
        }
      }
      // Sin red → el caller (_imprimirConReintento) encola; la cola se drena por la Edge al volver online (la
      // reserva permanente de wh.reservar_ticket evita reimprimir lo ya impreso).
      return { ok: false, offline: true, error: 'sin conexión' };
    },
    // [Adhesivo granel despacho] imprime 1 adhesivo por ítem KGM vía Edge print-adhesivo (mode granel-despacho).
    // items = [{codigo, nombre, peso(kg)}]. Fire-and-forget desde el despacho. Cero GAS. {ok,data:{impresos}} | {ok:false}.
    imprimirAdhesivoGranel: async (items, printerId) => {
      try {
        if (!navigator.onLine || !Array.isArray(items) || !items.length) return { ok: false };
        const token = await _mintTokenWH();
        if (!token) return { ok: false, error: 'sin token' };
        const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/print-adhesivo`, {
          method: 'POST',
          headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: 'granel-despacho', items, printerId: printerId || null })
        }, 15000);
        return await res.json().catch(() => ({ ok: false }));
      } catch (e) { return { ok: false, error: e.message }; }
    },
    // [cero-GAS 2026-07-14] aviso a cajas: Edge `aviso-cajas` (lee preingreso+cajas de Postgres, arma ESC/POS,
    // PrintNode). El preingreso ya vive en Supabase (crearPreingreso es directo) → la Edge lo cubre. Ya NO cae
    // a GAS: si la Edge falla/timeout/*_OFF, DEGRADA con aviso claro (el operador reimprime desde el preingreso).
    // Flag server WH_AVISO_DIRECTO=1 en prod (verificado). Sin doble impresión (idemKey en la Edge).
    imprimirAvisoCajeros: async (p) => {
      p = p || {};
      try {
        const token = await _mintTokenWH();
        if (token) {
          const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/aviso-cajas`, {
            method: 'POST',
            headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
            body: JSON.stringify({ idPreingreso: String(p.idPreingreso || ''), idemKey: String(p.idemKey || ''), reporteUrl: String(p.reporteUrl || '') })
          }, 15000);
          const d = await res.json().catch(() => null);
          if (res.ok && d && d.ok === true) {
            return { ok: true, data: { impresiones: Array.isArray(d.impresiones) ? d.impresiones : [], enviados: d.enviados || 0, cajas: d.cajas || 0, yaImpreso: !!d.yaImpreso, via: 'edge' } };
          }
          // d.ok===false (AVISO_OFF / preingreso no en Supabase / error) → degradar (NO GAS).
        }
      } catch (_) { /* degradar (NO GAS) */ }
      // [cero-GAS] Antes caía a GAS `imprimirAvisoCajeros`. Ahora el preingreso ya vive en Supabase (crearPreingreso
      // es directo), así que la Edge `aviso-cajas` cubre el caso; si aún así falla es transitorio → degradamos con
      // aviso claro (el operador reimprime desde el preingreso), NUNCA a GAS.
      return { ok: false, error: 'Aviso a cajas: servicio no disponible ahora · reintenta o reimprime desde el preingreso', via: 'edge-fail' };
    },
    getImpresorasEcosistema: () => call({ action: 'getImpresorasEcosistema' }),
    // [F6 push] Registro de token FCM directo a Supabase (mos.registrar_push_token). Aditivo al registro GAS
    // durante la transición; cuando los disparadores migren, la audiencia ya está en mos.push_tokens.
    registrarPushTokenSB: (p={}) => _sbRpcWH('registrar_push_token', { p }, 'mos'),
    // [F6 espía] Señalización WebRTC directo a Supabase (mos.espia_*). El device WH la usa Supabase-first;
    // *_OFF/APP_NO_AUTORIZADA o fallo de transporte → null → el caller cae a GAS.
    espiaRpc: async (rpc, p={}) => {
      const out = await _sbRpcWH(rpc, { p }, 'mos');
      if (!out || (out.ok === false && String(out.error || '') === 'APP_NO_AUTORIZADA')) return null;
      return out;
    },
    // [CERO-GAS F4] Sube un chunk de audio/video del espía a la Edge `espia-chunk` (Storage + registro).
    espiaSubirChunkEdge: async (params) => {
      try {
        const token = await _mintTokenWH();
        const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/espia-chunk`, {
          method: 'POST',
          headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify(params || {})
        }, 12000);
        return await res.json().catch(() => null);
      } catch (_) { return null; }
    },

    // [cero-GAS] El device cierra su sesión de audio en la sombra (mos.espia_audio_detener, SQL 508) al
    // auto-detenerse. Idempotente; sin fallback GAS. Best-effort (la escucha ya paró localmente).
    espiaAudioDetener: async (idSesion, deviceId) => {
      try { return await _sbRpcWH('espia_audio_detener', { p: { idSesion: idSesion || '', deviceId: deviceId || '' } }, 'mos'); }
      catch (_) { return null; }
    },
    getCargadoresDelDia:  (p={}) => call({ action: 'getCargadoresDelDia', ...p }),
    imprimirCargadoresDia:(p)    => post({ action: 'imprimirCargadoresDia', ...p }),
    getAlertasStock:    (p={})   => call({ action: 'getAlertasStock', ...p }),
    marcarAlertaRevisada:(p)     => post({ action: 'marcarAlertaRevisada', ...p }),
    aceptarTeoricoAlerta:(p)     => post({ action: 'aceptarTeoricoAlerta', ...p }),
    getStockMovimientos: (p={})  => call({ action: 'getStockMovimientos', ...p }),
    getWelcomeData:     (p={})   => call({ action: 'getWelcomeData', ...p }),
    verificarHorario:   (p={})   => call({ action: 'verificarHorario', ...p }),
    iniciarTestDiagnostico:   (p)      => post({ action: 'iniciarTestDiagnostico', ...p }),
    finalizarTestDiagnostico: (p)      => post({ action: 'finalizarTestDiagnostico', ...p }),
    getResultadosDiagnostico: ()       => call({ action: 'getResultadosDiagnostico' }),
    runInternalTests:         (p)      => post({ action: 'runInternalTests', ...p }),
    imprimirBienvenida: (p)      => post({ action: 'imprimirBienvenida', ...p }),

    // Guías — acciones extra
    reabrirGuia:        (p)      => post({ action: 'reabrirGuia', ...p }),
    anularDetalle:      (p)      => post({ action: 'anularDetalle', ...p }),
    // [cero-GAS] `autoCloseDayGuias` (post→GAS) ELIMINADO: era dead code (ningún caller) y un rastro de GAS.

    // Personal / Sesiones
    // [F1 login · cero-GAS + seguro] Valida el PIN SERVER-SIDE (mos.login_pin_wh; el pin nunca sale del server) +
    // crea/reusa sesión wh.sesiones. APP_NO_AUTORIZADA/sin-token → null → el caller muestra "⚠ Sin conexión ·
    // reintenta" (login 100% Supabase). El GAS `loginPersonal` se ELIMINÓ de este api: era dead code (ningún
    // caller lo invocaba) y un rastro de GAS. FIX del login caído: catalogo_wh_rls dejó de mandar el pin.
    loginPersonalSB: async (pin) => {
      // [accesos unificados] mandamos deviceId → el hook server-side registra el ingreso
      // en liquidaciones_dia y el cierre 11pm puede forzar logout de este dispositivo.
      let _dev = ''; try { _dev = _whDeviceId() || ''; } catch (_) {}
      const out = await _sbRpcWH('login_pin_wh', { p: { pin: String(pin || ''), deviceId: String(_dev) } }, 'mos');
      if (!out || (out.ok === false && String(out.error || '') === 'APP_NO_AUTORIZADA')) return null;
      return out;   // {ok:true,data:{...}} | {ok:false,error:'PIN incorrecto'/'PIN requerido'}
    },
    // [accesos unificados] heartbeat de asistencia (última conexión). Inerte si el flag
    // server MOS_ACCESOS_DIRECTO está OFF (la RPC devuelve _OFF). Fire-and-forget.
    heartbeatPersonalSB: async (idPersonal) => {
      if (!idPersonal) return null;
      try { return await _sbRpcWH('heartbeat_personal', { p: { idPersonal: String(idPersonal) } }, 'mos'); }
      catch (_) { return null; }
    },
    cerrarTurno:        (p)      => post({ action: 'cerrarTurno', ...p }),
    getPersonal:        ()       => call({ action: 'getPersonal' }),
    // [423] tarifa por unidad de envasado (mos.config vía wh.get_tarifa_envasado) — para el "Tu pago" del modal
    getTarifaEnvasado:  ()       => call({ action: 'getTarifaEnvasado' }),
    getSesionActiva:    (id)     => call({ action: 'getSesionActiva', idSesion: id }),
    getDesempenoDia:    (p={})   => call({ action: 'getDesempenoDia', ...p }),
    getResumenPersonal: (fecha)  => call({ action: 'getResumenPersonal', fecha: fecha || '' }),

    // Stock — historial de movimientos por producto
    getHistorialStock:    (cod)  => call({ action: 'getHistorialStock', codigoProducto: cod }),
    imprimirHistorialStock: (p)  => post({ action: 'imprimirHistorialStock', ...p }),

    // Auditoría diaria de stock (desde módulo Productos)
    auditarProducto:      (p)   => post({ action: 'auditarProducto', ...p }),

    // ── F0+ : genéricos para módulos nuevos ─────────────────────
    get:  (action, p={}) => call({ action, ...p }),
    post: (action, p={}) => post({ action, ...p }),
    pushEdge: (audiencia, titulo, cuerpo, data) => _pushEdgeWH(audiencia, titulo, cuerpo, data),

    // [40x cruce] usados SOLO por la cola offline (offline.js:sincronizar) para evitar el duplicado del cruce:
    //   _escrituraDirectaActiva() → si está ON, la cola reintenta vía _postCola (directo, dedupea por id sembrado)
    //   en vez de pegarle a GAS (que no comparte el dedup → duplicaría). _fromQueue evita el re-encolado.
    _escrituraDirectaActiva: () => _whEscrituraDirecta(),
    _postCola: (params) => post({ ...params, _fromQueue: true }),

    // ── F0: Auth — Supabase-first (RPC central mos.verificar_clave_admin, bcrypt+cascada+auditoría),
    // GAS solo kill-switch. ⚠ SEGURIDAD: el consumidor WH lee `res.ok` COMO "autorizado" → la RPC
    // devuelve {ok:true, autorizado:false} para clave mala, así que normalizamos ok = ok && autorizado
    // (sin esto, clave incorrecta autorizaría). Offline/sin token → cae a GAS (igual que hoy).
    verificarClaveAdmin: async (p) => {
      try {
        const r = await _sbRpcWH('verificar_clave_admin', {
          p_clave: String(p.clave || ''), p_accion: String(p.accion || 'GENERICA'),
          p_ref: String(p.refDocumento || ''), p_app: String(p.appOrigen || 'warehouseMos'),
          p_device: String(p.deviceId || p.dispositivo || ''), p_detalle: String(p.detalle || ''),
          p_tier: (p.tier != null ? parseInt(p.tier, 10) : null), p_cliente_meta: null
        }, 'mos');
        if (r && typeof r.ok !== 'undefined') {
          return { ok: !!(r.ok && r.autorizado), autorizado: !!r.autorizado, error: r.error || '',
                   nombre: r.nombre || '', rol: r.rol || '', nivel: r.nivel,
                   validadoPor: r.validado_por || '', idPersonal: r.id_personal || '', idAccion: r.id_accion || '' };
        }
      } catch (_) {}
      // [CERO-GAS / CERO-FALLBACK] Sin fallback GAS: si el directo falla, fail-closed (autorizado:false).
      // La RPC mos.verificar_clave_admin ya trae bcrypt + lockout + auditoría server-side. Nunca autoriza por error.
      return { ok: false, autorizado: false, error: 'No se pudo verificar — reintenta' };
    },

    // ── [cero-GAS G1] Estado de bloqueo + heartbeat (dispositivo+personal) en 1 RPC ──
    // Reemplaza el poll GAS getEstadoBloqueoUsuario. Devuelve {ok:true,data:{...}} (shape idéntico al GAS)
    // o null si: flag server-side WH_BLOQUEO_DIRECTO != '1', opt-out de dispositivo, offline o error →
    // el caller (BloqueoRemoto._check) cae a GAS al instante. El RPC hace el heartbeat (ultima_conexion en
    // mos.dispositivos + mos.personal) cuando responde ok, igual que el side-effect del endpoint GAS.
    estadoBloqueoUsuarioDirecto: async (params) => {
      if (!_whBloqueoDirecto() || !navigator.onLine) return null;
      try {
        const r = await _sbRpcWH('estado_bloqueo_usuario', { p: {
          nombre:     (params && params.nombre)     || '',
          idPersonal: (params && params.idPersonal) || '',
          appOrigen:  (params && params.appOrigen)  || 'warehouseMos',
          deviceId:   (params && params.deviceId)   || '',
          idZona:     (params && params.idZona)     || '',
          idEstacion: (params && params.idEstacion) || ''
        } }, 'mos', 8000);
        if (r && r.ok === true && r.data) return r;   // éxito
        return null;                                   // WH_BLOQUEO_DIRECTO_OFF u otro → GAS
      } catch (_) { return null; }
    },

    // ── [NIVEL 1 corte-GAS] Desbloqueo temporal 15 min 100% Supabase (mos.desbloquear_usuario_temporal, SQL 363) ──
    // Cero-GAS/cero-fallback: valida la clave admin + setea unlock_hasta en mos.bloqueos_usuario. Devuelve
    // {ok, data:{autorizado, unlockHasta(ms), msRestantes(ms), validadoPor, error?}} = shape que consume el modal.
    desbloquearUsuarioDirecto: async (params) => {
      return await _sbRpcWH('desbloquear_usuario_temporal', { p: {
        idPersonal: (params && params.idPersonal) || '',
        nombre:     (params && params.nombre)     || '',
        appOrigen:  (params && params.appOrigen)  || 'warehouseMos',
        claveAdmin: (params && params.claveAdmin) || ''
      } }, 'mos', 12000);
    },

    // ── [cero-GAS G2] Registro de ubicación GPS directo a Supabase (mos.registrar_ubicacion, flag GPS_DIRECTO) ──
    // Devuelve true si escribió en Supabase; false si flag OFF / offline / error → el caller cae al write GAS.
    registrarUbicacionDirecto: async (params) => {
      if (!navigator.onLine) return false;
      try {
        const r = await _sbRpcWH('registrar_ubicacion', { p: {
          deviceId:        (params && params.deviceId)        || '',
          lat:             (params && params.lat),
          lng:             (params && params.lng),
          accuracy:        (params && params.accuracy),
          bateria:         (params && params.bateria),
          usuarioLogueado: (params && params.usuarioLogueado) || ''
        } }, 'mos', 8000);
        return !!(r && r.ok === true);   // ok:false (GPS_DIRECTO_OFF) → false → GAS
      } catch (_) { return false; }
    },

    // [G4 online-only] Se eliminó adminPinsCacheDirecto / el caché de PINs admin: la verificación de clave admin
    // es siempre online (verificarClaveAdmin → mos.verificar_clave_admin). Sin conexión las acciones se bloquean.

    // ── F0: Cargadores independientes ───────────────────────────
    listarCargadoresMaster:  ()    => call({ action: 'listarCargadoresMaster' }),
    addCargadorDia:          (p)   => post({ action: 'addCargadorDia', ...p }),
    removeCargadorDia:       (p)   => post({ action: 'removeCargadorDia', ...p }),
    getResumenCargadoresDia: (p={})=> call({ action: 'getResumenCargadoresDia', ...p }),

    // ── F0: Mermas V2 ───────────────────────────────────────────
    getMermasCesta:           ()   => call({ action: 'getMermasCesta' }),
    agregarAMermas:           (p)  => post({ action: 'agregarAMermas', ...p }),
    solucionarMerma:          (p)  => post({ action: 'solucionarMerma', ...p }),
    procesarEliminacionMermas:(p)  => post({ action: 'procesarEliminacionMermas', ...p }),
    contadorMermasPendientes: ()   => call({ action: 'contadorMermasPendientes' }),

    // ── F0: Fotos genérico ──────────────────────────────────────
    subirFotoEntidad:    (p) => post({ action: 'subirFotoEntidad', ...p }),
    eliminarFotoEntidad: (p) => post({ action: 'eliminarFotoEntidad', ...p }),

    // ── F0: Op-log ──────────────────────────────────────────────
    aplicarOp:           (p) => post({ action: 'aplicarOp', ...p }),
    listarOpsPendientes: (p={}) => call({ action: 'listarOpsPendientes', ...p }),

    // ── F0: Productos extra (reconciliación caso 4) ─────────────
    getProductosCambiadosDesde: (p={}) => call({ action: 'getProductosCambiadosDesde', ...p }),

    // ── F0: Personal — rol del usuario ──────────────────────────
    getRolUsuario:       (p={}) => call({ action: 'getRolUsuario', ...p }),

    // ── Historial completo de guía (admin/master only) ──────────
    getHistorialGuia:    (p={}) => call({ action: 'getHistorialGuia', ...p })
  };
})();
