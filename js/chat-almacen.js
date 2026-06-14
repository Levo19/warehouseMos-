// ============================================================
// warehouseMos — "Pregúntale a tu almacén"
// Chat en lenguaje natural (español neutral peruano) sobre los datos
// REALES del almacén InversionMos, usando Claude vía la Edge `ia`
// con patrón TOOL-USE.
//
// ADITIVO y READ-ONLY: las tools solo LEEN datos ya disponibles
// (stock_enriquecido / cache local / lecturas API). NUNCA escriben,
// NUNCA generan SQL libre. Si algo falla → error limpio, no afecta
// el resto de la app.
//
// Arquitectura:
//   - System prompt acota el rol y obliga a responder solo con datos reales.
//   - tools[] son las herramientas que Claude puede invocar.
//   - Loop tool-use con cap de iteraciones (MAX_ITER) para no loopear.
//   - Fechas resueltas en TZ Lima (America/Lima).
//   - UI autocontenida: FAB flotante + modal (sin dvh, custom, responsive).
// ============================================================
'use strict';

const ChatAlmacen = (() => {

  // ── Límites duros (costo IA / anti-loop) ──────────────────
  const MODEL      = 'claude-haiku-4-5';   // default de la Edge; barato y soporta tool-use
  const MAX_TOKENS = 1024;                 // respuestas concisas
  const MAX_ITER   = 5;                    // tope de rondas tool-use por pregunta
  const MAX_HIST   = 8;                    // mensajes de historial que se reenvían a Claude
  const MAX_ROWS   = 60;                   // filas máx. que una tool devuelve a Claude (acota tokens)

  // ── Fechas en TZ Lima (Perú = UTC-5 fijo, sin DST) ────────
  function _hoyLima() {
    return new Intl.DateTimeFormat('en-CA', {
      timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit'
    }).format(new Date());   // yyyy-MM-dd
  }
  // Normaliza cualquier valor fecha a yyyy-MM-dd en TZ Lima ('' si inválida).
  function _fechaLima(v) {
    if (v == null || v === '') return '';
    const d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) {
      // ¿ya viene como yyyy-MM-dd? respétala
      const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(v));
      return m ? m[0] : '';
    }
    return new Intl.DateTimeFormat('en-CA', {
      timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit'
    }).format(d);
  }

  // ── Rol (anti-costo): solo admin/master pueden usar el chat IA ──
  // Misma fuente que el resto de la app: window.WH_CONFIG.rol (UPPER).
  // El FAB ya se oculta por CSS para no-admin, pero abrir()/preguntar()
  // son invocables por consola → revalidamos acá también.
  function _esAdmin() {
    try {
      const rol = String((window.WH_CONFIG && window.WH_CONFIG.rol) || '').toUpperCase();
      return rol === 'MASTER' || rol === 'ADMINISTRADOR';
    } catch (_) { return false; }
  }

  // ── Helpers de catálogo (cache local) ─────────────────────
  function _productos() {
    try {
      return (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache)
        ? (OfflineManager.getProductosCache() || []) : [];
    } catch (_) { return []; }
  }
  // Mapa código → producto (indexado por idProducto y codigoBarra).
  function _prodMap() {
    const m = {};
    _productos().forEach(p => {
      if (p.idProducto)  m[String(p.idProducto)]  = p;
      if (p.codigoBarra) m[String(p.codigoBarra)] = p;
    });
    return m;
  }
  function _descDe(codigo, fallback) {
    const p = _prodMap()[String(codigo)];
    return (p && (p.descripcion || p.nombre)) || fallback || String(codigo || '');
  }
  // ¿el texto de búsqueda matchea el producto? (por descripción o código, sin acentos)
  function _norm(s) {
    return String(s || '').toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
  }
  function _matchProducto(prod, q) {
    if (!q) return true;
    const nq = _norm(q);
    return _norm(prod.descripcion).includes(nq)
        || _norm(prod.nombre).includes(nq)
        || _norm(prod.idProducto).includes(nq)
        || _norm(prod.codigoBarra).includes(nq)
        || _norm(prod.skuBase).includes(nq);
  }

  // ════════════════════════════════════════════════════════════════════
  // DEFINICIÓN DE TOOLS (esquema Claude). READ-ONLY, acotadas.
  // ════════════════════════════════════════════════════════════════════
  function _tools() {
    return [
      {
        name: 'consultar_stock',
        description: 'Devuelve el stock actual (cantidad disponible) de los productos del almacén. Si se da "producto" filtra por nombre o código. Usa esto para preguntas de existencias, "cuánto hay de X", "qué está por agotarse", stock bajo mínimo.',
        input_schema: {
          type: 'object',
          properties: {
            producto:    { type: 'string', description: 'Nombre o código del producto a buscar (opcional). Vacío = todo el stock relevante.' },
            solo_bajos:  { type: 'boolean', description: 'Si true, solo productos por debajo de su stock mínimo.' }
          }
        }
      },
      {
        name: 'consultar_ingresos',
        description: 'Suma las cantidades que INGRESARON al almacén (guías de tipo INGRESO) en un rango de fechas, opcionalmente filtrando por producto. Usa esto para "cuánto entró de X esta semana", "qué se recibió ayer".',
        input_schema: {
          type: 'object',
          properties: {
            fecha_desde: { type: 'string', description: 'Fecha inicio en formato YYYY-MM-DD (TZ Lima).' },
            fecha_hasta: { type: 'string', description: 'Fecha fin en formato YYYY-MM-DD (TZ Lima, inclusive).' },
            producto:    { type: 'string', description: 'Nombre o código del producto (opcional).' }
          },
          required: ['fecha_desde', 'fecha_hasta']
        }
      },
      {
        name: 'consultar_vencimientos',
        description: 'Devuelve los lotes con stock que vencen en un rango de fechas. Usa esto para "qué vence el lunes", "qué se vence esta semana", "productos por caducar".',
        input_schema: {
          type: 'object',
          properties: {
            fecha_desde: { type: 'string', description: 'Fecha inicio YYYY-MM-DD (opcional; por defecto hoy).' },
            fecha_hasta: { type: 'string', description: 'Fecha fin YYYY-MM-DD (opcional; por defecto hoy + 30 días).' }
          }
        }
      },
      {
        name: 'consultar_envasados',
        description: 'Devuelve los envasados (producción de presentaciones) registrados en un rango de fechas. Usa esto para "qué se envasó esta semana", "cuántas unidades se produjeron".',
        input_schema: {
          type: 'object',
          properties: {
            fecha_desde: { type: 'string', description: 'Fecha inicio YYYY-MM-DD (opcional; por defecto últimos 7 días).' },
            fecha_hasta: { type: 'string', description: 'Fecha fin YYYY-MM-DD (opcional; por defecto hoy).' }
          }
        }
      }
    ];
  }

  // ════════════════════════════════════════════════════════════════════
  // EJECUTORES de tools (locales, READ-ONLY). Devuelven objeto JSON-able
  // pequeño que se pasa de vuelta a Claude como tool_result.
  // ════════════════════════════════════════════════════════════════════

  // Stock: prioriza lectura API (auto directo/GAS); cae a cache local.
  async function _ejecConsultarStock(input) {
    let filas = [];
    try {
      const r = await API.getStock({ soloAlertas: input && input.solo_bajos ? 'true' : 'false' });
      if (r && r.ok && Array.isArray(r.data)) filas = r.data;
    } catch (_) { /* cae a cache */ }
    if (!filas.length) {
      // fallback: cache local de stock + enriquecer descripción del catálogo
      try {
        const sc = (typeof OfflineManager !== 'undefined' && OfflineManager.getStockCache)
          ? (OfflineManager.getStockCache() || []) : [];
        filas = sc.map(s => {
          const p = _prodMap()[String(s.codigoProducto)] || {};
          return {
            codigoProducto: s.codigoProducto,
            descripcion:    s.descripcion || p.descripcion || s.codigoProducto,
            cantidadDisponible: parseFloat(s.cantidadDisponible) || 0,
            stockMinimo:    s.stockMinimo != null ? s.stockMinimo : (p.stockMinimo || 0),
            unidad:         s.unidad || p.unidad || ''
          };
        });
        if (input && input.solo_bajos) {
          filas = filas.filter(s => (parseFloat(s.cantidadDisponible) || 0) < (parseFloat(s.stockMinimo) || 0));
        }
      } catch (_) {}
    }
    const q = input && input.producto ? String(input.producto) : '';
    if (q) {
      filas = filas.filter(s => {
        const p = _prodMap()[String(s.codigoProducto)] || { descripcion: s.descripcion };
        return _matchProducto(p, q) || _norm(s.descripcion).includes(_norm(q));
      });
      // [FIX chat] Incluir productos del CATÁLOGO que matchean la búsqueda pero NO tienen fila de stock
      // (cantidad 0 / sin stock). Antes el chat decía "no tengo registro" de un producto que SÍ existe pero
      // está en 0 — porque solo miraba la tabla de stock. Ahora reporta "0 unidades, sin stock".
      try {
        const prods = (typeof OfflineManager !== 'undefined' && OfflineManager.getProductosCache)
          ? (OfflineManager.getProductosCache() || []) : [];
        const yaCods = new Set(filas.map(s => String(s.codigoProducto)));
        prods.forEach(p => {
          const cod = String(p.codigoBarra || p.idProducto || '');
          if (cod && !yaCods.has(cod) && _matchProducto(p, q)) {
            yaCods.add(cod);
            filas.push({
              codigoProducto: cod,
              descripcion: p.descripcion || p.nombre || cod,
              cantidadDisponible: 0,
              stockMinimo: parseFloat(p.stockMinimo) || 0,
              unidad: p.unidad || ''
            });
          }
        });
      } catch (_) {}
    }
    const items = filas.slice(0, MAX_ROWS).map(s => ({
      producto: s.descripcion || _descDe(s.codigoProducto),
      codigo:   String(s.codigoProducto || ''),
      cantidad: parseFloat(s.cantidadDisponible) || 0,
      unidad:   s.unidad || '',
      minimo:   parseFloat(s.stockMinimo) || 0,
      bajo_minimo: (parseFloat(s.cantidadDisponible) || 0) < (parseFloat(s.stockMinimo) || 0)
    }));
    return { total_encontrados: filas.length, mostrados: items.length, items };
  }

  // Ingresos: guías INGRESO* en el rango + sus detalles (cantidad recibida).
  async function _ejecConsultarIngresos(input) {
    const desde = _fechaLima(input && input.fecha_desde) || _hoyLima();
    const hasta = _fechaLima(input && input.fecha_hasta) || _hoyLima();
    let guias = [];
    try {
      const r = await API.getGuias({});
      if (r && r.ok && Array.isArray(r.data)) guias = r.data;
    } catch (_) {}
    if (!guias.length) {
      try {
        guias = (typeof OfflineManager !== 'undefined' && OfflineManager.getGuiasCache)
          ? (OfflineManager.getGuiasCache() || []) : [];
      } catch (_) {}
    }
    // Solo guías de INGRESO en el rango de fechas (TZ Lima).
    const guiasIng = guias.filter(g => {
      const tipo = String(g.tipo || '').toUpperCase();
      if (tipo.indexOf('INGRESO') !== 0) return false;
      const f = _fechaLima(g.fecha);
      return f && f >= desde && f <= hasta;
    });
    if (!guiasIng.length) {
      return { rango: { desde, hasta }, total_guias: 0, mensaje: 'No hay guías de ingreso en ese rango.' };
    }
    // Traer detalle de cada guía (acotado: máx. 30 guías para no abusar de la red/IA).
    const guiasUsar = guiasIng.slice(0, 30);
    const q = input && input.producto ? String(input.producto) : '';
    const acum = {};   // codigo → { producto, cantidad, unidad }
    for (const g of guiasUsar) {
      let det = [];
      try {
        const rg = await API.getGuia(g.idGuia);
        if (rg && rg.ok && rg.data && Array.isArray(rg.data.detalle)) det = rg.data.detalle;
      } catch (_) {}
      det.forEach(d => {
        const cod = String(d.codigoProducto || '');
        if (!cod) return;
        const prod = _prodMap()[cod] || {};
        const nombre = d.descripcionProducto || prod.descripcion || cod;
        if (q && !_matchProducto(prod, q) && !_norm(nombre).includes(_norm(q))) return;
        const cant = parseFloat(d.cantidadRecibida) || 0;
        if (cant <= 0) return;
        if (!acum[cod]) acum[cod] = { producto: nombre, codigo: cod, cantidad: 0, unidad: prod.unidad || '' };
        acum[cod].cantidad += cant;
      });
    }
    const items = Object.values(acum)
      .sort((a, b) => b.cantidad - a.cantidad)
      .slice(0, MAX_ROWS);
    return {
      rango: { desde, hasta },
      total_guias_ingreso: guiasIng.length,
      guias_analizadas: guiasUsar.length,
      productos: items,
      nota: guiasIng.length > guiasUsar.length ? 'Se analizaron las primeras ' + guiasUsar.length + ' guías del rango.' : undefined
    };
  }

  // Vencimientos: lotes con stock que vencen en el rango.
  async function _ejecConsultarVencimientos(input) {
    const hoy = _hoyLima();
    const desde = _fechaLima(input && input.fecha_desde) || hoy;
    let hasta = _fechaLima(input && input.fecha_hasta);
    if (!hasta) {
      const d = new Date(hoy + 'T00:00:00-05:00');
      d.setDate(d.getDate() + 30);
      hasta = _fechaLima(d);
    }
    let lotes = [];
    try {
      const r = await API.getLotes({ soloActivos: 'true' });
      if (r && r.ok && Array.isArray(r.data)) lotes = r.data;
    } catch (_) {}
    const enRango = lotes.filter(l => {
      if (String(l.estado || '').toUpperCase() !== 'ACTIVO') return false;
      if ((parseFloat(l.cantidadActual) || 0) <= 0) return false;
      const fv = _fechaLima(l.fechaVencimiento);
      return fv && fv >= desde && fv <= hasta;
    });
    const items = enRango
      .map(l => ({
        producto: _descDe(l.codigoProducto),
        codigo:   String(l.codigoProducto || ''),
        vence:    _fechaLima(l.fechaVencimiento),
        cantidad: parseFloat(l.cantidadActual) || 0,
        dias_restantes: (typeof l.diasRestantes === 'number') ? l.diasRestantes : undefined
      }))
      .sort((a, b) => String(a.vence).localeCompare(String(b.vence)))
      .slice(0, MAX_ROWS);
    return { rango: { desde, hasta }, total: enRango.length, lotes: items };
  }

  // Envasados en el rango.
  async function _ejecConsultarEnvasados(input) {
    const hoy = _hoyLima();
    let desde = _fechaLima(input && input.fecha_desde);
    if (!desde) {
      const d = new Date(hoy + 'T00:00:00-05:00');
      d.setDate(d.getDate() - 7);
      desde = _fechaLima(d);
    }
    const hasta = _fechaLima(input && input.fecha_hasta) || hoy;
    let env = [];
    try {
      const r = await API.getEnvasados({ fechaDesde: desde });
      if (r && r.ok && Array.isArray(r.data)) env = r.data;
    } catch (_) {}
    const enRango = env.filter(e => {
      const f = _fechaLima(e.fecha);
      return f && f >= desde && f <= hasta;
    });
    const items = enRango.slice(0, MAX_ROWS).map(e => ({
      producto:           e.descripcionProductoEnvasado || _descDe(e.codigoProductoEnvasado),
      unidades_producidas: parseFloat(e.unidadesProducidas) || 0,
      fecha:              _fechaLima(e.fecha),
      eficiencia_pct:     parseFloat(e.eficienciaPct) || 0,
      estado:             e.estado || ''
    }));
    const totalUds = items.reduce((a, b) => a + (b.unidades_producidas || 0), 0);
    return { rango: { desde, hasta }, total_envasados: enRango.length, total_unidades: totalUds, envasados: items };
  }

  async function _ejecutarTool(name, input) {
    try {
      if (name === 'consultar_stock')         return await _ejecConsultarStock(input || {});
      if (name === 'consultar_ingresos')      return await _ejecConsultarIngresos(input || {});
      if (name === 'consultar_vencimientos')  return await _ejecConsultarVencimientos(input || {});
      if (name === 'consultar_envasados')     return await _ejecConsultarEnvasados(input || {});
      return { error: 'Herramienta desconocida: ' + name };
    } catch (e) {
      return { error: 'No se pudo consultar: ' + (e && e.message ? e.message : String(e)) };
    }
  }

  // ════════════════════════════════════════════════════════════════════
  // LOOP TOOL-USE
  // ════════════════════════════════════════════════════════════════════
  function _systemPrompt() {
    return [
      'Eres el asistente del almacén central InversionMos (warehouseMos).',
      'Respondes preguntas del DUEÑO sobre stock, ingresos, vencimientos y envasados.',
      'Hoy es ' + _hoyLima() + ' (zona horaria de Perú, America/Lima).',
      '',
      'REGLAS ESTRICTAS:',
      '- Responde SOLO con datos reales obtenidos de las herramientas. NUNCA inventes cifras, productos ni fechas.',
      '- Si las herramientas no devuelven datos para responder, dilo claramente ("No tengo ese dato" / "No encontré nada").',
      '- Usa las herramientas para obtener los datos antes de responder. Resuelve expresiones de tiempo a fechas concretas (YYYY-MM-DD) en TZ Lima: "esta semana" = lunes a domingo de la semana actual; "el lunes" = el próximo lunes; "ayer", "hoy", etc.',
      '- Responde en ESPAÑOL neutral (peruano), de forma CONCISA y directa. Sin saludos largos.',
      '- Cuando des cantidades, incluye la unidad si está disponible. Si listas varios productos, usa una lista corta.',
      '- No expongas códigos internos a menos que el usuario los pida; usa los nombres de los productos.'
    ].join('\n');
  }

  // historial: [{role:'user'|'assistant', text}] (de la UI). Devuelve string con la respuesta final.
  async function preguntar(pregunta, historial) {
    if (!_esAdmin()) {
      throw new Error('Solo administradores pueden usar el asistente del almacén.');
    }
    if (typeof API === 'undefined' || !API.llamarEdgeIA) {
      throw new Error('La IA no está disponible en esta sesión.');
    }
    const q = String(pregunta || '').trim();
    if (!q) return '';

    // Reconstruir mensajes para Claude desde el historial de la UI (acotado).
    const messages = [];
    (historial || []).slice(-MAX_HIST).forEach(m => {
      if (!m || !m.text) return;
      messages.push({ role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.text) });
    });
    messages.push({ role: 'user', content: q });

    const system = _systemPrompt();
    const tools  = _tools();

    for (let iter = 0; iter < MAX_ITER; iter++) {
      const resp = await API.llamarEdgeIA({
        model: MODEL, max_tokens: MAX_TOKENS, system, tools, messages
      });
      if (!resp || !Array.isArray(resp.content)) {
        throw new Error('Respuesta inesperada de la IA.');
      }
      const content = resp.content;

      // ¿Claude pidió usar herramientas?
      if (resp.stop_reason === 'tool_use') {
        // Asegurar que el turno del assistant (con los tool_use blocks) entre al historial tal cual.
        messages.push({ role: 'assistant', content });
        const toolResults = [];
        for (const block of content) {
          if (block.type !== 'tool_use') continue;
          const out = await _ejecutarTool(block.name, block.input);
          toolResults.push({
            type: 'tool_result',
            tool_use_id: block.id,
            content: JSON.stringify(out)
          });
        }
        messages.push({ role: 'user', content: toolResults });
        continue;   // siguiente ronda: Claude ve los resultados
      }

      // Respuesta final en texto.
      const texto = content.filter(b => b.type === 'text').map(b => b.text).join('').trim();
      return texto || 'No tengo una respuesta para eso.';
    }
    return 'La consulta resultó muy compleja. Intenta preguntar algo más específico.';
  }

  // ════════════════════════════════════════════════════════════════════
  // UI — FAB flotante + modal autocontenido. Sin dvh, custom, responsive.
  // ════════════════════════════════════════════════════════════════════
  const _hist = [];   // [{role, text}]
  let _pensando = false;

  function _esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function _inyectarUI() {
    if (document.getElementById('chatAlmacenModal')) return;
    const css = `
      /* ── FAB inteligente ────────────────────────────────────── */
      #chatAlmacenFab{position:fixed;right:calc(16px + env(safe-area-inset-right,0px));
        bottom:calc(84px + env(safe-area-inset-bottom,0px));z-index:9600;width:56px;height:56px;border:none;border-radius:50%;
        background:linear-gradient(135deg,#0ea5e9,#2563eb);color:#fff;font-size:24px;cursor:pointer;
        box-shadow:0 8px 24px rgba(37,99,235,.45);display:flex;align-items:center;justify-content:center;
        opacity:.55;transform:scale(.94);
        transition:transform .28s cubic-bezier(.34,1.56,.64,1),opacity .3s ease,width .3s ease,height .3s ease,box-shadow .3s ease;
        animation:camFabFloat 4.2s ease-in-out infinite, camFabGlow 6s ease-in-out infinite;
        touch-action:none;-webkit-user-select:none;user-select:none}
      #chatAlmacenFab .cam-fab-ico{transition:transform .28s cubic-bezier(.34,1.56,.64,1);display:inline-block;will-change:transform}
      #chatAlmacenFab:hover,#chatAlmacenFab:focus-visible{opacity:1;transform:scale(1.1);box-shadow:0 12px 32px rgba(37,99,235,.6);outline:none}
      #chatAlmacenFab:active{transform:scale(.9)}
      #chatAlmacenFab.cam-mini{width:42px;height:42px;font-size:18px;opacity:.4;animation:none}
      #chatAlmacenFab.cam-dragging{transition:none;opacity:1;animation:none;cursor:grabbing}
      #chatAlmacenFab .cam-fab-badge{position:absolute;top:-2px;right:-2px;width:14px;height:14px;border-radius:50%;
        background:#f43f5e;border:2px solid #0f172a;display:none;animation:camBadgePulse 1.4s ease-in-out infinite}
      #chatAlmacenFab.cam-has-new .cam-fab-badge{display:block}
      @keyframes camFabFloat{0%,100%{transform:scale(.94) translateY(0)}50%{transform:scale(.94) translateY(-6px)}}
      @keyframes camFabGlow{0%,100%{box-shadow:0 8px 24px rgba(37,99,235,.45)}50%{box-shadow:0 8px 30px rgba(14,165,233,.7)}}
      @keyframes camBadgePulse{0%,100%{transform:scale(1);opacity:1}50%{transform:scale(1.25);opacity:.7}}

      /* ── Overlay + modal (glassmorphism) ────────────────────── */
      #chatAlmacenOverlay{position:fixed;left:0;top:0;right:0;bottom:0;z-index:9601;
        background:rgba(2,6,23,.55);-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px);
        display:none;align-items:flex-end;justify-content:center;opacity:0;transition:opacity .26s ease}
      #chatAlmacenOverlay.open{display:flex;opacity:1}
      #chatAlmacenModal{position:relative;color:#e2e8f0;width:100%;max-width:560px;max-height:88vh;
        background:linear-gradient(160deg,rgba(15,23,42,.92),rgba(15,23,42,.82));
        -webkit-backdrop-filter:blur(22px) saturate(140%);backdrop-filter:blur(22px) saturate(140%);
        border-radius:22px 22px 0 0;display:flex;flex-direction:column;
        box-shadow:0 -12px 50px rgba(0,0,0,.6),inset 0 1px 0 rgba(255,255,255,.06);
        border:1px solid rgba(148,163,184,.16);border-bottom:none;
        transform:translateY(100%);opacity:0;transition:transform .42s cubic-bezier(.22,1.2,.36,1),opacity .3s ease}
      #chatAlmacenOverlay.open #chatAlmacenModal{transform:translateY(0);opacity:1}
      @media(min-width:640px){
        #chatAlmacenOverlay{align-items:center}
        #chatAlmacenModal{border-radius:22px;border-bottom:1px solid rgba(148,163,184,.16);max-height:80vh;
          transform:translateY(24px) scale(.92);transform-origin:bottom right}
        #chatAlmacenOverlay.open #chatAlmacenModal{transform:translateY(0) scale(1)}
      }
      #chatAlmacenModal .cam-head{display:flex;align-items:center;gap:10px;padding:14px 16px;
        border-bottom:1px solid rgba(148,163,184,.12);background:linear-gradient(90deg,rgba(37,99,235,.14),transparent)}
      #chatAlmacenModal .cam-head .cam-avatar{font-size:22px;display:inline-block;animation:camAvatarIdle 5s ease-in-out infinite;transform-origin:50% 70%}
      #chatAlmacenModal .cam-head .cam-ttl{font-weight:700;font-size:15px;flex:1}
      #chatAlmacenModal .cam-head .cam-sub{font-size:11px;color:#7c92ad;font-weight:400;display:block;margin-top:1px}
      #chatAlmacenModal .cam-mute{background:rgba(30,41,59,.7);border:1px solid rgba(148,163,184,.14);color:#94a3b8;
        width:32px;height:32px;border-radius:9px;cursor:pointer;font-size:15px;line-height:1;transition:transform .15s,background .15s;flex:none}
      #chatAlmacenModal .cam-mute:active{transform:scale(.88)}
      #chatAlmacenModal .cam-mute:hover{background:rgba(51,65,85,.8)}
      #chatAlmacenModal .cam-x{background:rgba(30,41,59,.7);border:1px solid rgba(148,163,184,.14);color:#94a3b8;
        width:32px;height:32px;border-radius:9px;cursor:pointer;font-size:18px;line-height:1;transition:transform .15s,background .15s;flex:none}
      #chatAlmacenModal .cam-x:active{transform:scale(.88) rotate(90deg)}
      #chatAlmacenModal .cam-x:hover{background:rgba(127,29,29,.55);color:#fecaca}
      @keyframes camAvatarIdle{0%,88%,100%{transform:rotate(0) scale(1)}92%{transform:rotate(-8deg) scale(1.08)}96%{transform:rotate(6deg) scale(1.08)}}
      @keyframes camAvatarThink{0%,100%{transform:translateY(0) rotate(-3deg)}50%{transform:translateY(-3px) rotate(3deg)}}

      #chatAlmacenBody{flex:1;overflow-y:auto;padding:14px 14px 4px;display:flex;flex-direction:column;gap:10px;min-height:120px;-webkit-overflow-scrolling:touch}
      #chatAlmacenBody .cam-msg{max-width:84%;padding:9px 13px;border-radius:16px;font-size:14px;line-height:1.45;
        white-space:pre-wrap;word-break:break-word;box-shadow:0 2px 8px rgba(0,0,0,.18);
        animation:camMsgIn .38s cubic-bezier(.22,1.2,.36,1) both;will-change:transform,opacity}
      #chatAlmacenBody .cam-user{align-self:flex-end;background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;border-bottom-right-radius:5px;
        animation-name:camMsgInRight}
      #chatAlmacenBody .cam-bot{align-self:flex-start;background:rgba(30,41,59,.85);color:#e2e8f0;border-bottom-left-radius:5px;
        border:1px solid rgba(148,163,184,.1);animation-name:camMsgInLeft}
      #chatAlmacenBody .cam-bot.cam-celebra{animation-name:camMsgInLeft,camBotPop}
      #chatAlmacenBody .cam-err{align-self:flex-start;background:linear-gradient(135deg,#991b1b,#7f1d1d);color:#fecaca;border-radius:14px;animation-name:camMsgInLeft}
      @keyframes camMsgIn{from{opacity:0;transform:translateY(12px) scale(.96)}to{opacity:1;transform:translateY(0) scale(1)}}
      @keyframes camMsgInRight{from{opacity:0;transform:translateX(26px) translateY(6px) scale(.96)}to{opacity:1;transform:translateX(0) translateY(0) scale(1)}}
      @keyframes camMsgInLeft{from{opacity:0;transform:translateX(-26px) translateY(6px) scale(.96)}to{opacity:1;transform:translateX(0) translateY(0) scale(1)}}
      @keyframes camBotPop{0%{}55%{transform:scale(1)}68%{transform:scale(1.04)}82%{transform:scale(.99)}100%{transform:scale(1)}}

      /* sparkle burst al recibir respuesta */
      #chatAlmacenBody .cam-spark{position:relative}
      #chatAlmacenBody .cam-spark::before,#chatAlmacenBody .cam-spark::after,#chatAlmacenBody .cam-spark>.cam-sp3{
        content:'';position:absolute;top:-4px;left:8px;width:6px;height:6px;border-radius:50%;background:#fbbf24;
        opacity:0;pointer-events:none;box-shadow:0 0 6px #fbbf24}
      #chatAlmacenBody .cam-spark::before{animation:camSpark1 .9s ease-out forwards}
      #chatAlmacenBody .cam-spark::after{left:auto;right:10px;background:#7dd3fc;box-shadow:0 0 6px #7dd3fc;animation:camSpark2 .95s ease-out forwards}
      #chatAlmacenBody .cam-spark>.cam-sp3{position:absolute;left:50%;background:#a78bfa;box-shadow:0 0 6px #a78bfa;animation:camSpark3 .85s ease-out forwards}
      @keyframes camSpark1{0%{opacity:0;transform:translate(0,0) scale(.4)}20%{opacity:1}100%{opacity:0;transform:translate(-14px,-22px) scale(1.1)}}
      @keyframes camSpark2{0%{opacity:0;transform:translate(0,0) scale(.4)}20%{opacity:1}100%{opacity:0;transform:translate(14px,-20px) scale(1.1)}}
      @keyframes camSpark3{0%{opacity:0;transform:translate(-50%,0) scale(.4)}20%{opacity:1}100%{opacity:0;transform:translate(-50%,-26px) scale(1.2)}}

      #chatAlmacenBody .cam-think{align-self:flex-start;display:flex;gap:9px;align-items:center;
        background:rgba(30,41,59,.7);border:1px solid rgba(148,163,184,.1);padding:9px 13px;border-radius:16px;border-bottom-left-radius:5px;
        animation:camMsgInLeft .38s cubic-bezier(.22,1.2,.36,1) both}
      #chatAlmacenBody .cam-think .cam-think-ava{font-size:18px;display:inline-block;animation:camAvatarThink 1.2s ease-in-out infinite}
      #chatAlmacenBody .cam-think .cam-dots{display:flex;gap:4px;align-items:center}
      #chatAlmacenBody .cam-think .cam-dot{width:7px;height:7px;border-radius:50%;background:#7dd3fc;animation:camBounce 1.1s infinite ease-in-out}
      #chatAlmacenBody .cam-think .cam-dot:nth-child(2){animation-delay:.16s}
      #chatAlmacenBody .cam-think .cam-dot:nth-child(3){animation-delay:.32s}
      @keyframes camBounce{0%,60%,100%{transform:translateY(0);opacity:.45}30%{transform:translateY(-6px);opacity:1}}
      #chatAlmacenBody .cam-hint{align-self:center;color:#5b708a;font-size:12px;text-align:center;padding:18px 8px;line-height:1.5;animation:camMsgIn .4s ease both}
      #chatAlmacenChips{display:flex;flex-wrap:wrap;gap:6px;padding:6px 14px 0}
      #chatAlmacenChips button{background:rgba(30,41,59,.8);border:1px solid rgba(148,163,184,.18);color:#cbd5e1;font-size:12px;
        padding:6px 11px;border-radius:14px;cursor:pointer;transition:transform .15s,background .15s,border-color .15s}
      #chatAlmacenChips button:hover{background:rgba(37,99,235,.25);border-color:#3b82f6;transform:translateY(-1px)}
      #chatAlmacenChips button:active{transform:scale(.94)}
      #chatAlmacenFoot{display:flex;gap:8px;padding:10px 12px;border-top:1px solid rgba(148,163,184,.12);align-items:flex-end}
      #chatAlmacenInput{flex:1;background:rgba(30,41,59,.8);border:1px solid rgba(148,163,184,.2);color:#e2e8f0;border-radius:14px;
        padding:10px 13px;font-size:14px;resize:none;max-height:96px;outline:none;font-family:inherit;transition:border-color .2s,box-shadow .2s}
      #chatAlmacenInput:focus{border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.18)}
      #chatAlmacenSend{position:relative;overflow:hidden;background:linear-gradient(135deg,#3b82f6,#2563eb);border:none;color:#fff;
        width:44px;height:44px;border-radius:14px;cursor:pointer;font-size:18px;flex:none;transition:transform .14s,box-shadow .2s;
        box-shadow:0 4px 14px rgba(37,99,235,.4)}
      #chatAlmacenSend:hover{box-shadow:0 6px 18px rgba(37,99,235,.6)}
      #chatAlmacenSend:active{transform:scale(.86)}
      #chatAlmacenSend:disabled{opacity:.4;cursor:default;box-shadow:none}
      #chatAlmacenSend .cam-ripple{position:absolute;border-radius:50%;background:rgba(255,255,255,.5);transform:scale(0);
        animation:camRipple .6s ease-out;pointer-events:none}
      @keyframes camRipple{to{transform:scale(2.6);opacity:0}}

      /* ── Respeto a prefers-reduced-motion ──────────────────── */
      @media (prefers-reduced-motion: reduce){
        #chatAlmacenFab{animation:none!important}
        #chatAlmacenFab .cam-fab-ico,#chatAlmacenModal .cam-head .cam-avatar,
        #chatAlmacenBody .cam-think .cam-think-ava,#chatAlmacenBody .cam-think .cam-dot{animation:none!important}
        #chatAlmacenModal,#chatAlmacenOverlay,#chatAlmacenBody .cam-msg,#chatAlmacenBody .cam-think,
        #chatAlmacenBody .cam-hint{animation:none!important;transition-duration:.12s!important}
        #chatAlmacenBody .cam-spark::before,#chatAlmacenBody .cam-spark::after,#chatAlmacenBody .cam-spark>.cam-sp3{animation:none!important;display:none}
        #chatAlmacenModal{transform:none!important;opacity:1}
      }
    `;
    const style = document.createElement('style');
    style.id = 'chatAlmacenCSS';
    style.textContent = css;
    document.head.appendChild(style);

    const wrap = document.createElement('div');
    wrap.innerHTML = `
      <button id="chatAlmacenFab" title="Pregúntale a tu almacén" aria-label="Pregúntale a tu almacén"><span class="cam-fab-ico">🤖</span><span class="cam-fab-badge" aria-hidden="true"></span></button>
      <div id="chatAlmacenOverlay" role="dialog" aria-modal="true" aria-label="Pregúntale a tu almacén">
        <div id="chatAlmacenModal">
          <div class="cam-head">
            <span class="cam-avatar">🤖</span>
            <div class="cam-ttl">Pregúntale a tu almacén<span class="cam-sub">Stock, ingresos, vencimientos y envasados</span></div>
            <button class="cam-mute" id="chatAlmacenMute" aria-label="Activar o silenciar sonidos" title="Sonidos">🔊</button>
            <button class="cam-x" id="chatAlmacenClose" aria-label="Cerrar">&times;</button>
          </div>
          <div id="chatAlmacenBody"></div>
          <div id="chatAlmacenChips"></div>
          <div id="chatAlmacenFoot">
            <textarea id="chatAlmacenInput" rows="1" placeholder="Escribe tu pregunta…" autocomplete="off"></textarea>
            <button id="chatAlmacenSend" title="Enviar">➤</button>
          </div>
        </div>
      </div>`;
    document.body.appendChild(wrap);

    // Restaurar posición arrastrada del FAB (si la hay) y enganchar drag.
    _restaurarPosFab();
    _engancharDragFab();
    _engancharScrollFab();

    document.getElementById('chatAlmacenFab').addEventListener('click', () => {
      if (_fabArrastrado) { _fabArrastrado = false; return; }  // un drag no debe abrir
      abrir();
    });
    document.getElementById('chatAlmacenClose').addEventListener('click', cerrar);
    document.getElementById('chatAlmacenOverlay').addEventListener('click', e => {
      if (e.target && e.target.id === 'chatAlmacenOverlay') cerrar();
    });
    // Toggle de sonido (persistente).
    _syncMuteBtn();
    document.getElementById('chatAlmacenMute').addEventListener('click', () => {
      _muted = !_muted;
      try { localStorage.setItem('cam_muted', _muted ? '1' : '0'); } catch (_) {}
      _syncMuteBtn();
      if (!_muted) _sfx('beep');   // confirmación al reactivar
    });
    const input = document.getElementById('chatAlmacenInput');
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); _enviar(); }
    });
    input.addEventListener('input', () => {
      input.style.height = 'auto';
      input.style.height = Math.min(96, input.scrollHeight) + 'px';
    });
    document.getElementById('chatAlmacenSend').addEventListener('click', _enviar);
  }

  // ── Sonido (reusa SoundFX del ecosistema, iOS-safe vía su _getCtx) ──
  // SoundFX ya hace `new (AudioContext||webkitAudioContext)` + ctx.resume()
  // en cada toque, así que se desbloquea solo en el primer gesto del usuario.
  let _muted = (() => { try { return localStorage.getItem('cam_muted') === '1'; } catch (_) { return false; } })();
  function _sfx(name) {
    if (_muted) return;
    try {
      const fx = window.SoundFX;
      if (fx && typeof fx[name] === 'function') fx[name]();
    } catch (_) {}
  }
  function _syncMuteBtn() {
    const b = document.getElementById('chatAlmacenMute');
    if (b) { b.textContent = _muted ? '🔇' : '🔊'; b.style.opacity = _muted ? '.55' : '1'; }
  }
  // Vibración háptica (Android; iOS no-op). Guard de soporte.
  function _vibrar(p) {
    try { if (navigator && typeof navigator.vibrate === 'function') navigator.vibrate(p); } catch (_) {}
  }

  // ── FAB: drag opcional + persistencia de posición ──────────
  let _fabArrastrado = false;
  function _restaurarPosFab() {
    try {
      const raw = localStorage.getItem('cam_fab_pos');
      if (!raw) return;
      const p = JSON.parse(raw);
      const fab = document.getElementById('chatAlmacenFab');
      if (!fab || !p || typeof p.left !== 'number') return;
      fab.style.left = p.left + 'px';
      fab.style.top  = p.top + 'px';
      fab.style.right = 'auto';
      fab.style.bottom = 'auto';
    } catch (_) {}
  }
  function _engancharDragFab() {
    const fab = document.getElementById('chatAlmacenFab');
    if (!fab || fab._camDrag) return;
    fab._camDrag = true;
    let sx = 0, sy = 0, ox = 0, oy = 0, moved = false, dragging = false;
    function onDown(e) {
      const pt = e.touches ? e.touches[0] : e;
      const r = fab.getBoundingClientRect();
      sx = pt.clientX; sy = pt.clientY; ox = r.left; oy = r.top;
      moved = false; dragging = false;
      document.addEventListener('mousemove', onMove, { passive: false });
      document.addEventListener('touchmove', onMove, { passive: false });
      document.addEventListener('mouseup', onUp);
      document.addEventListener('touchend', onUp);
    }
    function onMove(e) {
      const pt = e.touches ? e.touches[0] : e;
      const dx = pt.clientX - sx, dy = pt.clientY - sy;
      if (!dragging && (Math.abs(dx) > 6 || Math.abs(dy) > 6)) { dragging = true; fab.classList.add('cam-dragging'); }
      if (!dragging) return;
      moved = true;
      if (e.cancelable) e.preventDefault();
      const sz = fab.offsetWidth;
      let nl = Math.max(6, Math.min(window.innerWidth - sz - 6, ox + dx));
      let nt = Math.max(6, Math.min(window.innerHeight - sz - 6, oy + dy));
      fab.style.left = nl + 'px'; fab.style.top = nt + 'px';
      fab.style.right = 'auto'; fab.style.bottom = 'auto';
    }
    function onUp() {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('touchmove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.removeEventListener('touchend', onUp);
      fab.classList.remove('cam-dragging');
      if (moved) {
        _fabArrastrado = true;   // suprime el click-abre inmediato
        try {
          const r = fab.getBoundingClientRect();
          localStorage.setItem('cam_fab_pos', JSON.stringify({ left: Math.round(r.left), top: Math.round(r.top) }));
        } catch (_) {}
      }
    }
    fab.addEventListener('mousedown', onDown);
    fab.addEventListener('touchstart', onDown, { passive: true });
  }

  // ── FAB se "esconde sin esconderse": al scrollear se encoge a mini-burbuja ──
  let _scrollTO = null;
  function _engancharScrollFab() {
    if (window._camScrollHooked) return;
    window._camScrollHooked = true;
    window.addEventListener('scroll', () => {
      const fab = document.getElementById('chatAlmacenFab');
      const ov  = document.getElementById('chatAlmacenOverlay');
      if (!fab || (ov && ov.classList.contains('open'))) return;
      fab.classList.add('cam-mini');
      if (_scrollTO) clearTimeout(_scrollTO);
      _scrollTO = setTimeout(() => { fab.classList.remove('cam-mini'); }, 950);
    }, { passive: true });
  }

  // Marca para celebrar (sparkle) la próxima respuesta del bot renderizada.
  let _celebrarIdx = -1;

  function _render() {
    const body = document.getElementById('chatAlmacenBody');
    if (!body) return;
    let html = '';
    if (!_hist.length && !_pensando) {
      html += '<div class="cam-hint">Pregunta en lenguaje natural.<br>Ejemplo: «¿cuánto azúcar entró esta semana?»</div>';
    }
    _hist.forEach((m, i) => {
      const cls = m.role === 'user' ? 'cam-user' : (m.role === 'error' ? 'cam-err' : 'cam-bot');
      const celebra = (m.role === 'assistant' && i === _celebrarIdx) ? ' cam-celebra cam-spark' : '';
      const extra   = celebra ? '<span class="cam-sp3"></span>' : '';
      html += `<div class="cam-msg ${cls}${celebra}">${_esc(m.text)}${extra}</div>`;
    });
    if (_pensando) {
      html += '<div class="cam-think"><span class="cam-think-ava">🤖</span>'
            + '<span class="cam-dots"><span class="cam-dot"></span><span class="cam-dot"></span><span class="cam-dot"></span></span></div>';
    }
    body.innerHTML = html;
    body.scrollTop = body.scrollHeight;
    _renderChips();
  }

  function _renderChips() {
    const chips = document.getElementById('chatAlmacenChips');
    if (!chips) return;
    if (_hist.length || _pensando) { chips.innerHTML = ''; return; }
    const ejemplos = [
      'Stock de aceite',
      '¿Qué vence esta semana?',
      '¿Qué se envasó esta semana?',
      '¿Qué está por agotarse?'
    ];
    chips.innerHTML = ejemplos.map(e =>
      `<button data-q="${_esc(e)}">${_esc(e)}</button>`).join('');
    chips.querySelectorAll('button').forEach(b => {
      b.addEventListener('click', () => {
        const inp = document.getElementById('chatAlmacenInput');
        inp.value = b.getAttribute('data-q');
        _enviar();
      });
    });
  }

  // Ripple visual en el botón enviar (microinteracción táctil).
  function _ripple() {
    try {
      const b = document.getElementById('chatAlmacenSend');
      if (!b) return;
      const r = document.createElement('span');
      r.className = 'cam-ripple';
      const s = Math.max(b.offsetWidth, b.offsetHeight);
      r.style.width = r.style.height = s + 'px';
      r.style.left = (b.offsetWidth / 2 - s / 2) + 'px';
      r.style.top  = (b.offsetHeight / 2 - s / 2) + 'px';
      b.appendChild(r);
      setTimeout(() => { try { r.remove(); } catch (_) {} }, 620);
    } catch (_) {}
  }

  async function _enviar() {
    if (_pensando) return;
    const inp = document.getElementById('chatAlmacenInput');
    const q = (inp.value || '').trim();
    if (!q) return;
    inp.value = '';
    inp.style.height = 'auto';
    // Feedback al enviar: tick + vibración corta + ripple.
    _ripple();
    _sfx('click');
    _vibrar(10);
    _celebrarIdx = -1;
    _hist.push({ role: 'user', text: q });
    _pensando = true;
    _setSendEnabled(false);
    _render();
    try {
      const resp = await preguntar(q, _hist.slice(0, -1));   // historial SIN la pregunta recién agregada
      _hist.push({ role: 'assistant', text: resp });
      _celebrarIdx = _hist.length - 1;     // sparkle + bounce en esta burbuja
      _sfx('done');                        // chime agradable de 2 notas
      _vibrar([15]);
    } catch (e) {
      _hist.push({ role: 'error', text: 'No pude responder: ' + (e && e.message ? e.message : 'error de conexión') + '.' });
      _sfx('warn');                        // tono descendente suave
      _vibrar([30, 20, 30]);
    } finally {
      _pensando = false;
      _setSendEnabled(true);
      _render();
    }
  }

  function _setSendEnabled(on) {
    const b = document.getElementById('chatAlmacenSend');
    if (b) b.disabled = !on;
  }

  // Listener Escape (se ata al abrir, se limpia al cerrar; nunca se acumula).
  let _onKeydown = null;

  function abrir() {
    // Anti-costo: aunque el FAB esté oculto por CSS, abrir() es invocable por
    // consola. Revalidamos rol contra la misma fuente que el resto de la app.
    if (!_esAdmin()) {
      try { if (typeof toast === 'function') toast('Solo administradores pueden usar el asistente del almacén.', 'warn'); } catch (_) {}
      return;
    }
    _inyectarUI();
    const ov = document.getElementById('chatAlmacenOverlay');
    if (ov) ov.classList.add('open');
    // Sonido de apertura (pop/blip suave) + limpiar badge de "nuevo".
    _sfx('scanReady');
    const fab = document.getElementById('chatAlmacenFab');
    if (fab) { fab.classList.remove('cam-has-new', 'cam-mini'); }
    // Cerrar con Escape mientras el modal está abierto.
    if (!_onKeydown) {
      _onKeydown = function (e) {
        if (e.key === 'Escape' || e.key === 'Esc') { e.preventDefault(); cerrar(); }
      };
      document.addEventListener('keydown', _onKeydown);
    }
    _render();
    setTimeout(() => { const i = document.getElementById('chatAlmacenInput'); if (i) i.focus(); }, 80);
  }
  function cerrar() {
    const ov = document.getElementById('chatAlmacenOverlay');
    if (ov) ov.classList.remove('open');
    if (_onKeydown) { document.removeEventListener('keydown', _onKeydown); _onKeydown = null; }
  }

  // Muestra/oculta el FAB (p.ej. solo admin/master). Por defecto se inyecta en init().
  function mostrarFab(visible) {
    _inyectarUI();
    const fab = document.getElementById('chatAlmacenFab');
    if (fab) fab.style.display = visible === false ? 'none' : 'flex';
  }

  // init: inyecta el FAB. Llamar tras login (idealmente solo admin/master).
  function init(opts) {
    try {
      _inyectarUI();
      mostrarFab(!opts || opts.fab !== false);
    } catch (_) {}
  }

  return { init, abrir, cerrar, mostrarFab, preguntar };
})();

if (typeof window !== 'undefined') window.ChatAlmacen = ChatAlmacen;
