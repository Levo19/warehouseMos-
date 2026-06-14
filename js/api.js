// warehouseMos — api.js  Comunicación con GAS + soporte offline
'use strict';

const API = (() => {
  function _gasUrl() { return window.WH_CONFIG?.gasUrl || ''; }

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

  // Pide el JWT (app=warehouseMos, exp ~5min) a GAS (endpoint mintTokenWH, B1) y lo cachea.
  // Re-mintea 30s antes de expirar. Timeout 6s: un GAS colgado no debe trabar la cadena directa.
  // _mintInFlight dedup: si varias lecturas salen juntas (arranque) → 1 solo POST a GAS, no ráfaga.
  let _mintInFlight = null;
  async function _mintTokenWH() {
    const now = Math.floor(Date.now() / 1000);
    if (_sbTok.token && (_sbTok.exp - now) > 30) return _sbTok.token;
    if (_mintInFlight) return _mintInFlight;
    _mintInFlight = (async () => {
      const GAS_URL = _gasUrl();
      if (!GAS_URL) throw new Error('sin gasUrl');
      let deviceId = '';
      try { deviceId = localStorage.getItem('wh_device_id') || ''; } catch (_) {}
      const res = await _whFetchTimeout(GAS_URL, {
        method: 'POST', headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ action: 'mintTokenWH', deviceId })
      }, 6000);
      const d = await res.json();
      if (!d || !d.ok || !d.token) throw new Error('mint-token: ' + ((d && d.error) || 'sin token'));
      _sbTok.token = d.token; _sbTok.exp = d.exp || (now + 300);
      return d.token;
    })();
    try { return await _mintInFlight; }
    finally { _mintInFlight = null; }
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
    const res = await _whFetchTimeout(`${_SB_URL}/storage/v1/object/wh-fotos/${path}`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': mime || 'image/jpeg', 'x-upsert': 'true' },  // [40x #1] upsert: reintento sobreescribe (idempotente)
      body: bin
    }, 30000);   // 30s — fotos de alta resolución pesan
    if (!res.ok) throw new Error('storage upload ' + res.status);
    return {
      ok: true, path,
      url:     `${_SB_URL}/storage/v1/object/public/wh-fotos/${path}`,                          // original (ver detalle/zoom)
      preview: `${_SB_URL}/storage/v1/render/image/public/wh-fotos/${path}?width=800&quality=72` // liviano (listas)
    };
  }

  // [PASO 5 · B5] Llama la Edge `ia` (proxy a Claude). body = {messages, system?, model?, max_tokens?}. Devuelve el JSON de Claude.
  async function _llamarEdgeIA(body) {
    const token = await _mintTokenWH();
    const res = await _whFetchTimeout(`${_SB_URL}/functions/v1/ia`, {
      method: 'POST',
      headers: { 'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }, 40000);   // la IA puede tardar
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

  // Llama una RPC de LECTURA directo a PostgREST (apikey + Bearer + Profile). profile='wh' (default) o 'mos' (catálogo).
  async function _sbRpcWH(fn, args, profile) {
    const prof = profile || 'wh';
    const token = await _mintTokenWH();
    const res = await _whFetchTimeout(`${_SB_URL}/rest/v1/rpc/${fn}`, {
      method: 'POST',
      headers: {
        'apikey': _SB_ANON, 'Authorization': 'Bearer ' + token,
        'Accept-Profile': prof, 'Content-Profile': prof, 'Content-Type': 'application/json'
      },
      body: JSON.stringify(args || {})
    }, 12000);
    if (!res.ok) throw new Error('rpc directo HTTP ' + res.status);
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
    pickups: [['id_pickup','idPickup','text'],['fuente','fuente','text'],['estado','estado','text'],['items','items','json'],['id_zona','idZona','text'],['notas','notas','text'],['creado_por','creadoPor','text'],['fecha_creado','fechaCreado','date'],['fecha_atendido','fechaAtendido','date'],['atendido_por','atendidoPor','text'],['ultima_actividad','ultimaActividad','date']],
    guia_detalle: [['id_guia','idGuia','text'],['linea','linea','int'],['cod_producto','codigoProducto','text'],['cant_esperada','cantidadEsperada','num'],['cant_recibida','cantidadRecibida','num'],['precio_unitario','precioUnitario','num'],['id_lote','idLote','text'],['observacion','observacion','text'],['id_producto_nuevo','idProductoNuevo','text'],['id_detalle','idDetalle','text'],['fecha_vencimiento','fechaVencimiento','date']],
    envasados: [['id_envasado','idEnvasado','text'],['cod_producto_base','codigoProductoBase','text'],['cantidad_base','cantidadBase','num'],['unidad_base','unidadBase','text'],['cod_producto_envasado','codigoProductoEnvasado','text'],['unidades_esperadas','unidadesEsperadas','num'],['unidades_producidas','unidadesProducidas','num'],['merma_real','mermaReal','num'],['eficiencia_pct','eficienciaPct','num'],['fecha','fecha','date'],['usuario','usuario','text'],['estado','estado','text'],['id_guia_salida','idGuiaSalida','text'],['id_guia_ingreso','idGuiaIngreso','text'],['observacion','observacion','text']]
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
    const mermasPendientes = (mermas || []).filter(m => m.estado === 'PENDIENTE');
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
  async function _sbLeerTablaWH(tabla) {
    const out = await _sbRpcWH('leer_tabla_rls', { p_tabla: tabla });
    if (!out || out.ok === false) throw new Error((out && out.error) || 'leer_tabla error');
    return _sbRowsToObjsFront(tabla, out.data);
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
      _sbLeerTablaWH('guia_detalle'),
      _sbLeerTablaWH('preingresos'),
      _sbLeerTablaWH('ajustes'),
      _sbLeerTablaWH('auditorias'),
      _sbRpcWH('stock_enriquecido_rls', { solo_alertas: false })
    ]);
    // stock: mismo shape vivo que API.getStock (stock_enriquecido_rls) — gana al Sheet congelado.
    const stock = (stockR && stockR.ok !== false && Array.isArray(stockR.data)) ? stockR.data : [];
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
    const out = await _sbRpcWH('catalogo_wh_rls', {}, 'mos');
    if (!out || out.ok === false) throw new Error((out && out.error) || 'catalogo error');
    return { ok: true, data: {
      productos:     _mapCat('productos',     out.productos),
      equivalencias: _mapCat('equivalencias', out.equivalencias),
      proveedores:   _mapCat('proveedores',   out.proveedores),
      personal:      _mapCat('personal',      out.personal),
      impresoras:    _mapCat('impresoras',    out.impresoras),
      zonas:         _mapCat('zonas',         out.zonas)
    } };
  }

  // Mapea una acción de lectura → su RPC directa. Devuelve la respuesta {ok,data}
  // o null si la acción NO tiene backend RLS listo (→ el llamador cae a GAS).
  async function _callDirecto(params) {
    const action = params.action;
    if (action === 'descargarMaestros') return await _sbDescargarMaestros();  // catálogo directo (mos.catalogo_wh_rls)
    if (action === 'getDashboard') {
      // Agregado de 8 datasets + KPIs. Trae todo en paralelo + cálculo réplica fiel de getDashboard (GAS).
      const [seR, lotes, guias, preingresos, mermas, auditorias, envasados] = await Promise.all([
        _sbRpcWH('stock_enriquecido_rls', { solo_alertas: false }),
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
      const out = await _sbRpcWH('stock_enriquecido_rls', { solo_alertas: String(params.soloAlertas) === 'true' });
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
      const se = await _sbRpcWH('stock_enriquecido_rls', { solo_alertas: false });
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
    if (action === 'getMermasEnProceso') {
      let rows = await _sbLeerTablaWH('mermas');
      rows = rows.filter(r => String(r.estado || '').toUpperCase() === 'EN_PROCESO');
      const hoyMs = _diaLimaMs(new Date());   // [TZ Lima] días-calendario
      rows.forEach(r => { const fMs = r.fechaIngreso ? _diaLimaMs(r.fechaIngreso) : hoyMs; const dd = Math.floor((hoyMs - fMs) / _UN_DIA); r.diasEnProceso = dd; r.vencida = dd > 3; });
      rows.sort((a, b) => (_diaLimaMs(b.fechaIngreso) || 0) - (_diaLimaMs(a.fechaIngreso) || 0));   // [40x #3] || 0: fecha inválida no rompe el orden
      return { ok: true, data: rows };
    }
    if (action === 'getMermasVencidas') {
      let rows = await _sbLeerTablaWH('mermas');
      const hoyMs = _diaLimaMs(new Date());   // [TZ Lima] días-calendario
      const vencidas = rows.filter(r => {
        if (String(r.estado || '').toUpperCase() !== 'EN_PROCESO') return false;
        const fMs = r.fechaIngreso ? _diaLimaMs(r.fechaIngreso) : hoyMs;
        return Math.floor((hoyMs - fMs) / _UN_DIA) > 3;
      });
      return { ok: true, data: { count: vencidas.length, mermas: vencidas } };
    }
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
      let rows = await _sbLeerTablaWH('pickups');
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
  async function _postDirecto(params) {
    const lid = params.localId || _genLocalId();

    // [PASO 5 · B5] IMPRESIÓN DIRECTA primero (gate PROPIO _whImpresionDirecta, independiente de la escritura).
    // Se evalúa ANTES de las RPCs de escritura para que el flag de impresión pueda entrar a _postDirecto
    // SIN activar la escritura directa. Cada caso revalida su flag; devuelve {ok:true} o null (→ GAS).
    const _printResult = await _postDirectoImpresion(params, lid);
    if (_printResult !== undefined) return _printResult;

    // Las RPCs de escritura SOLO corren si la escritura directa está activa (flag propio). Con la escritura OFF
    // pero la impresión ON, llegar acá significa "no era impresión cableada" → null → GAS (no se toca Supabase).
    // [100x rollback-fix] EXCEPCIÓN: un ítem `_viaDirecta` (encolado por timeout de escritura directa, que PUDO
    // commitear en Supabase) SIEMPRE debe reintentarse por las RPCs aunque el flag global ya esté apagado por
    // rollback. La RPC es idempotente (dedup por el id sembrado del localId) → reintento seguro, no duplica.
    if (!_whEscrituraDirecta() && !params._viaDirecta) return null;

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
      if (!der || !der.codigoProductoBase) return null;  // no resoluble → GAS
      const base = prods.find(p => String(p.idProducto) === der.codigoProductoBase || String(p.codigoBarra) === der.codigoProductoBase);
      const factor = parseFloat(der.factorConversionBase) || 0;
      const unidades = parseInt(params.unidadesProducidas) || 0;
      if (!base || factor <= 0 || unidades <= 0) return null;  // datos incompletos → GAS
      const out = await _sbRpcWH('registrar_envasado', { p: {
        id_envasado: 'ENV_' + lid, cod_producto_base: String(base.codigoBarra), cod_producto_envasado: cod,
        cantidad_base: unidades * factor, unidades_producidas: unidades, unidad_base: base.unidad || '',
        fecha_vencimiento: params.fechaVencimiento || '', usuario: params.usuario || ''
      } });
      if (!out || out.ok === false) return null;
      if (!out.dedup) {
        const ses = (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '';
        try { await _sbRpcWH('registrar_actividad', { p_id_sesion: ses, p_tipo: 'ENVASADO_REGISTRADO', p_cantidad: 1 }); } catch (_) {}
        try { await _sbRpcWH('registrar_actividad', { p_id_sesion: ses, p_tipo: 'UNIDADES_ENVASADAS', p_cantidad: unidades }); } catch (_) {}
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
      if (!out.dedup) { try { await _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'AUDITORIA_EJECUTADA', p_cantidad: 1 }); } catch (_) {} }
      return out;
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
      // IA directa vía Edge `ia` (Claude). Réplica fiel de analizarListaSombra (IA.gs): mismo system + parse + normalización.
      const texto = String(params.texto || '').trim();
      if (!texto) return { ok: false, error: 'TEXTO_VACIO' };
      if (texto.length > 50000) return { ok: false, error: 'TEXTO_MUY_LARGO', mensaje: 'Chunk demasiado grande' };
      const system = [
        'Eres un asistente que limpia listas de productos de almacén.',
        'Recibes texto pegado de cualquier formato (Excel, WhatsApp, email, ticket impreso).',
        'Tu trabajo: extraer SOLO los productos reales con su cantidad pedida.', '',
        'IGNORA:', '- Cabeceras (Código, Descripción, Pedido, etc.)', '- Totales / subtotales',
        '- Líneas separadoras (---, ===, ...)', '- Comentarios o notas que no sean producto', '',
        'POR CADA PRODUCTO devuelve:',
        '- nombre: descripción del producto en MAYÚSCULAS, limpia, sin códigos pegados',
        '- cantidad: número decimal con 1 decimal (ej: 5.0, 18.0, 0.5)',
        '- codigoVisto: opcional — si el texto traía un código/sku al lado, ponlo (string), si no, omite el campo', '',
        'RESPONDE EXCLUSIVAMENTE con JSON válido en este formato (sin markdown, sin comentarios):',
        '{"items":[{"nombre":"...","cantidad":N.N,"codigoVisto":"..."}]}'
      ].join('\n');
      let resp;
      try { resp = await _llamarEdgeIA({ max_tokens: 8192, system, messages: [{ role: 'user', content: 'Limpia esta lista y devuelve solo JSON:\n\n' + texto }] }); }
      catch (_) { return null; }   // fallback GAS
      const text = (resp && resp.content && resp.content[0] && resp.content[0].text) || '';
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
      const b64 = String(params.fotoBase || '').trim();
      if (!b64) return null;
      let url;
      try { url = (await _subirFotoStorage('guias', String(params.idGuia || ''), b64, params.mimeType || 'image/jpeg', lid)).url; } catch (_) { return null; }
      const out = await _sbRpcWH('actualizar_foto_guia', { p: { id_guia: String(params.idGuia || ''), foto: url } });
      if (!out || out.ok === false) return null;
      // [B5] OCR del comprobante en BACKGROUND (fire-and-forget, igual que el disparo automático de GAS, Fotos.gs:59):
      // no bloquea el ok de la foto; persiste los campos SUNAT cuando Claude responde. Inerte si el flag OCR está OFF.
      _ocrComprobanteGuia(String(params.idGuia || ''), b64, params.mimeType || 'image/jpeg').catch(() => {});
      return { ok: true, data: { url } };
    }
    if (params.action === 'registrarMerma') {
      // foto OBLIGATORIA → sube a Storage (máxima calidad) y registra la merma directo (ya no depende de GAS/Drive).
      let fotoUrl = String(params.foto || '');
      const b64 = String(params.fotoBase64 || '').trim();
      if (b64) { try { fotoUrl = (await _subirFotoStorage('mermas', 'M_' + lid, b64, params.mimeType || 'image/jpeg', lid)).url; } catch (_) { return null; } }
      if (!fotoUrl) return null;   // sin foto → que GAS valide (foto obligatoria)
      const out = await _sbRpcWH('registrar_merma', { p: {
        id_merma: 'M_' + lid, codigo_producto: String(params.codigoProducto || ''), cantidad: params.cantidadOriginal,
        motivo: params.motivo || '', usuario: params.usuario || '', responsable: params.responsable || '',
        origen: params.origen || '', id_lote: params.idLote || '', foto: fotoUrl
      } });
      if (!out || out.ok === false) return null;
      // [100x] guard de dedup (paridad con crearGuia/cerrarGuia/auditarProducto): en un reintento desde la
      // cola la merma se dedupea (out.dedup=true) pero SIN este guard el contador de actividad se incrementaba
      // de nuevo (inflaba la métrica). registrar_merma devuelve dedup:true cuando la fila ya existía. Verificado.
      if (!out.dedup) { try { await _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'MERMA_REGISTRADA', p_cantidad: 1 }); } catch (_) {} }
      return out;
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
      if (!out.dedup) { try { await _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'GUIA_CREADA', p_cantidad: 1 }); } catch (_) {} }
      // [shape-fix] el front lee res.data?.idGuia (app.js:7634) y res.offline. La RPC devuelve {ok,id_guia,dedup}
      // al nivel raíz → envolver al shape GAS {ok:true, data:{idGuia, estado:'ABIERTA'}}. idGuia = id determinista 'G_'+lid.
      return { ok: true, data: { idGuia: 'G_' + lid, estado: 'ABIERTA' }, dedup: !!out.dedup };
    }
    if (params.action === 'cerrarGuia') {
      // orquestador: arma detalles desde la guía (get_guia_rls) → cerrar_guia (aplica stock+FIFO) → actividad.
      // Idempotente: si ya CERRADA, la RPC devuelve yaCerrada (no reaplica stock).
      const g = await _sbRpcWH('get_guia_rls', { p_id: String(params.idGuia || '') });
      if (!g || g.ok === false) return null;
      const detalle = _sbRowsToObjsFront('guia_detalle', g.detalle || []);
      const detalles = detalle.map(d => ({
        codigo_producto: d.codigoProducto, cantidad_recibida: d.cantidadRecibida,
        precio_unitario: d.precioUnitario, id_lote: d.idLote, fecha_vencimiento: d.fechaVencimiento
      }));
      const out = await _sbRpcWH('cerrar_guia', { p: { id_guia: String(params.idGuia || ''), usuario: params.usuario || '', detalles } });
      if (!out || out.ok === false) return null;
      const _yaCerrada = !!(out.yaCerrada || out.ya_cerrada);
      if (!_yaCerrada) { try { await _sbRpcWH('registrar_actividad', { p_id_sesion: (window.WH_CONFIG && window.WH_CONFIG.idSesion) || '', p_tipo: 'GUIA_CERRADA', p_cantidad: 1 }); } catch (_) {} }
      // [shape-fix] el front lee res.data?.montoTotal (app.js:7571). La RPC devuelve el monto al nivel raíz →
      // envolver al shape GAS {ok:true, data:{idGuia, estado:'CERRADA', montoTotal, yaCerrada?}}. Tolerante al nombre de campo.
      const _monto = (out.montoTotal != null ? out.montoTotal : (out.monto_total != null ? out.monto_total : (out.monto != null ? out.monto : 0)));
      const _data = { idGuia: String(params.idGuia || ''), estado: 'CERRADA', montoTotal: parseFloat(_monto) || 0 };
      if (_yaCerrada) _data.yaCerrada = true;
      return { ok: true, data: _data };
    }
    if (params.action === 'reabrirGuia') {
      // Reverso de stock/lotes. Idempotente por estado (FOR UPDATE + anti doble-reverso en la RPC). La autorización
      // admin (REABRIR_GUIA) se valida ANTES en el flujo. id_guia real (no generado).
      const out = await _sbRpcWH('reabrir_guia', { p: { id_guia: String(params.idGuia || ''), usuario: params.usuario || '' } });
      if (!out || out.ok === false) return null;
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
    if (params.action === 'resolverMerma') {
      // Resuelve una merma (rep no toca stock; des → línea en guía SALIDA_MERMA semanal ABIERTA, que descuenta stock al
      // CERRARSE). NO mueve stock acá. Idempotente: guard de estado en la RPC (RESUELTA→yaResuelta) + dedup por local_id
      // (inserta línea de guía). El front solo manda idMerma/rep/des/obs; la RPC resuelve cod/original desde la fila.
      const out = await _sbRpcWH('resolver_merma', { p: {
        id_merma: String(params.idMerma || ''),
        cantidad_reparada: params.cantidadReparada, cantidad_desechada: params.cantidadDesechada,
        observacion_resolucion: params.observacionResolucion || '', usuario: params.usuario || '',
        id_detalle: 'MRMDET_' + lid, local_id: lid
      } });
      if (!out || out.ok === false) return null;   // *_OFF o error → GAS
      return out;
    }
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
    const PRINT_ACTIONS = ['imprimirBienvenida', 'imprimirMembrete'];
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

    return null;
  }

  // GET: network first, cache como fallback
  async function call(params) {
    const GAS_URL = _gasUrl();
    // [PASO 5 · B3] lectura directa a Supabase (inerte; fallback TOTAL a GAS ante cualquier fallo)
    if (_whLecturaDirecta() && navigator.onLine) {
      try {
        const directo = await _callDirecto(params);
        if (directo) return directo;
      } catch (_) { /* cae a GAS abajo */ }
    }
    if (!GAS_URL) return _fromCache(params);
    try {
      const url = new URL(GAS_URL);
      Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
      const res  = await fetch(url.toString(), { redirect: 'follow' });
      const data = await res.json();
      return data;
    } catch {
      return _fromCache(params);
    }
  }

  // Fallback GET desde caché offline
  function _fromCache(params) {
    const action = params.action;
    if (action === 'getProductos')   return { ok: true, data: OfflineManager.getProductosCache() };
    if (action === 'getStock')       return { ok: true, data: OfflineManager.getStockCache() };
    if (action === 'getProveedores') return { ok: true, data: OfflineManager.getProveedoresCache() };
    if (action === 'getPersonal')    return { ok: true, data: OfflineManager.getPersonalCache().map(p => { const s = {...p}; delete s.pin; return s; }) };
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
    'registrarEnvasado', 'corregirUnidadesEnvasado', 'anularEnvasadoConClave',
    'registrarMerma', 'resolverMerma',
    'registrarProductoNuevo', 'aprobarProductoNuevo',
    'crearPreingreso', 'aprobarPreingreso', 'actualizarPreingreso',
    'subirFotoGuia', 'subirFotoPreingreso',   // [40x #1] localId estable → nombre de foto determinístico (idempotente)
    'crearProducto', 'actualizarProducto',
    'crearProveedor', 'actualizarProveedor',
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
    if ((_whEscrituraDirecta() || _whImpresionDirecta() || params._viaDirecta) && navigator.onLine) {
      _opsEnVuelo++; _emitOpsState();
      let timeoutDirecto = false;
      try { const d = await _postDirecto(params); if (d) return d; }
      catch (_) { timeoutDirecto = true; }
      finally { _opsEnVuelo--; _emitOpsState(); }
      // [100x rollback-fix] Si llegamos acá con un ítem `_viaDirecta`, _postDirecto NO devolvió éxito
      // (devolvió null o lanzó). Ir a GAS NUNCA es seguro para estos ítems (duplicaría: GAS no comparte el
      // SYNC_LOG de Supabase). Lo dejamos en la cola para reintento DIRECTO posterior (sincronizar() lo verá
      // 'error' por el _retry y volverá a llamar API._postCola → _postDirecto). Nunca cae al GAS de abajo.
      if (params._viaDirecta) {
        return { ok: false, error: timeoutDirecto ? 'timeout-directo' : 'no-routeable-directo', _retry: true };
      }
      if (timeoutDirecto) {
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
      // (si llegamos acá, _postDirecto devolvió null → seguir a GAS abajo, seguro)
    }

    if (!GAS_URL || !navigator.onLine) {
      if (params._fromQueue) return { ok: false, error: 'sin-conexion', _retry: true };
      const localId = OfflineManager.encolar(params.action, params);
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

  return {
    // Descarga maestros MOS → localStorage
    descargarMaestros:    ()     => call({ action: 'descargarMaestros' }),
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
    getLotes:           (p={})   => call({ action: 'getLotesVencimiento', ...p }),
    crearProducto:      (p)      => post({ action: 'crearProducto', ...p }),
    actualizarProducto: (p)      => post({ action: 'actualizarProducto', ...p }),

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
    registrarMerma:     (p)      => post({ action: 'registrarMerma', ...p }),
    resolverMerma:      (p)      => post({ action: 'resolverMerma', ...p }),
    getMermasEnProceso: (p={})   => call({ action: 'getMermasEnProceso', ...p }),
    getMermasVencidas:  ()       => call({ action: 'getMermasVencidas' }),

    // Auditorias
    getAuditorias:      (p={})   => call({ action: 'getAuditorias', ...p }),
    asignarAuditoria:   (p)      => post({ action: 'asignarAuditoria', ...p }),
    ejecutarAuditoria:  (p)      => post({ action: 'ejecutarAuditoria', ...p }),

    // Ajustes
    getAjustes:         (p={})   => call({ action: 'getAjustes', ...p }),
    crearAjuste:        (p)      => post({ action: 'crearAjuste', ...p }),

    // Proveedores
    getProveedores:     (p={})   => call({ action: 'getProveedores', ...p }),
    crearProveedor:     (p)      => post({ action: 'crearProveedor', ...p }),
    actualizarProveedor:(p)      => post({ action: 'actualizarProveedor', ...p }),

    // Producto Nuevo
    getProductosNuevos:          (p={}) => call({ action: 'getProductosNuevos', ...p }),
    getProductosNuevosRecientes: (p={}) => call({ action: 'getProductosNuevosRecientes', ...p }),
    registrarPN:                 (p)    => post({ action: 'registrarProductoNuevo', ...p }),
    aprobarPN:                   (p)    => post({ action: 'aprobarProductoNuevo', ...p }),

    // Config
    getConfig:          ()                  => call({ action: 'getConfig' }),
    setConfig:          (clave, valor)      => post({ action: 'setConfigValue', clave, valor }),

    // Etiquetas / Tickets
    imprimirEtiqueta:   (p)      => post({ action: 'imprimirEtiqueta', ...p }),
    imprimirMembrete:   (p)      => post({ action: 'imprimirMembrete', ...p }),
    imprimirTicketGuia: (p)      => post({ action: 'imprimirTicketGuia', ...p }),
    imprimirAvisoCajeros:(p)     => post({ action: 'imprimirAvisoCajeros', ...p }),
    getImpresorasEcosistema: () => call({ action: 'getImpresorasEcosistema' }),
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
    autoCloseDayGuias:  ()       => post({ action: 'autoCloseDayGuias' }),

    // Personal / Sesiones
    loginPersonal:      (pin)    => post({ action: 'loginPersonal', pin }),
    cerrarTurno:        (p)      => post({ action: 'cerrarTurno', ...p }),
    getPersonal:        ()       => call({ action: 'getPersonal' }),
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

    // [40x cruce] usados SOLO por la cola offline (offline.js:sincronizar) para evitar el duplicado del cruce:
    //   _escrituraDirectaActiva() → si está ON, la cola reintenta vía _postCola (directo, dedupea por id sembrado)
    //   en vez de pegarle a GAS (que no comparte el dedup → duplicaría). _fromQueue evita el re-encolado.
    _escrituraDirectaActiva: () => _whEscrituraDirecta(),
    _postCola: (params) => post({ ...params, _fromQueue: true }),

    // ── F0: Auth ────────────────────────────────────────────────
    verificarClaveAdmin: (p)   => post({ action: 'verificarClaveAdmin', ...p }),

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
