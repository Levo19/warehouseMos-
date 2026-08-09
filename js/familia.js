// ============================================================
// warehouseMos — js/familia.js
// DESAMBIGUACIÓN DE CÓDIGOS DUPLICADOS  ·  window.FamiliaCB
// ------------------------------------------------------------
// EL PROBLEMA (palabras del dueño):
//   Los fabricantes reusan el MISMO código de barras físico para varias
//   presentaciones distintas. Logística lo resuelve agregando una LETRA al
//   código interno:
//     7758725000036A  CABALLO DE ORO AZUL PASTA WANTAN 500GR
//     7758725000036B  CABALLO DE ORO DORADO WANTAN 500GR
//   A veces el código PELADO también es un producto y las letras cuelgan de él:
//     7750464444799   LA CHINA TAMARINDO      ← pelado
//     7750464444799A..D  4 productos FOCH     ← 5 en la familia
//   Las letras NO son correlativas: son SEMÁNTICAS (O de orégano, R de romero,
//   F de fina, G de gruesa). Por eso "el siguiente sufijo" es la siguiente
//   letra LIBRE del alfabeto, no la que sigue a la última.
//
// LA SOLUCIÓN (cita del dueño):
//   "al escanear, este se encuentra pero con opción A B C D o E, entonces ahí
//    debe dar para escoger… cuando WH registra un producto este le da a escoger
//    entre los que existen, es más inteligente".
//
// REGLA DE FAMILIA (en piedra):
//   raíz    = código normalizado sin las letras finales  →  /[A-Za-z]+$/
//   familia = TODOS los productos cuyo código sea la raíz pelada o raíz+letras
//             (case-insensitive, trim, sin caracteres de control).
//   Un código de 14 dígitos que EMPIECE con la raíz NO es de la familia (es otro
//   EAN). Esto es más estricto —y más correcto— que la búsqueda por prefijo que
//   ya existía en WH.
//
// CUÁNDO SE PREGUNTA:
//   · Se escaneó/tecleó la raíz PELADA (el lector físico siempre entrega eso)
//     y la familia tiene MÁS DE UN miembro → selector.
//   · Se tecleó el código CON su letra (…036A) → el operador YA decidió: se
//     resuelve directo, cero latencia, cero preguntas.
//   · Familia de UNO → comportamiento anterior intacto.
//
// 100% CLIENT-SIDE: solo lee el catálogo cacheado de OfflineManager. El operador
// escanea sin red y el selector abre igual.
// ============================================================

(function () {
  'use strict';

  // Raíz mínima: un EAN mutilado de 4 dígitos no debe arrastrar media bodega.
  const MIN_RAIZ  = 6;
  const RE_SUFIJO = /[A-Za-z]+$/;
  const ALFABETO  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const Z_SHEET   = 2147481000;   // encima de todo lo de la app, debajo del lightbox de foto

  // ── helpers resilientes (app.js carga DESPUÉS; se resuelven en runtime) ──
  function _norm(s) {
    if (typeof window.normCb === 'function') return window.normCb(s);
    return String(s || '')
      .replace(/[\x00-\x1F\x7F-\x9F]/g, '')
      .replace(/['’‘´]/g, '-')
      .trim().toUpperCase();
  }
  function _esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function _qty(n) {
    if (typeof window.fmtQty === 'function') return window.fmtQty(n);
    const v = parseFloat(n) || 0;
    return String(Math.round(v * 1000) / 1000);
  }
  function _prods()  { try { return window.OfflineManager.getProductosCache()      || []; } catch (_) { return []; } }
  function _equivs() { try { return window.OfflineManager.getEquivalenciasCache()  || []; } catch (_) { return []; } }
  function _stocks() { try { return window.OfflineManager.getStockCache()          || []; } catch (_) { return []; } }
  function _activo(p) { return p && p.estado !== '0' && p.estado !== 0; }
  function _canonico(p) { return parseFloat((p && p.factorConversion) || 1) === 1; }

  // ══════════════════════════════════════════════════════════════
  // REGLA: raíz y sufijo
  // ══════════════════════════════════════════════════════════════
  function raiz(cod) {
    const c = _norm(cod);
    if (!c) return '';
    const r = c.replace(RE_SUFIJO, '');
    if (r.length < MIN_RAIZ) return c;    // raíz demasiado corta → la regla no aplica
    if (!/[0-9]$/.test(r))   return c;    // la raíz de un EAN termina en dígito; "ABC-XY" no es familia
    return r;
  }

  function sufijo(cod) {
    const c = _norm(cod);
    const r = raiz(cod);
    return (r && c.length > r.length) ? c.slice(r.length) : '';
  }

  // ══════════════════════════════════════════════════════════════
  // ÍNDICE raíz → miembros. Se reconstruye solo cuando cambia la
  // REFERENCIA del cache (mismo truco que _idxPorCodigo en app.js):
  // un escaneo normal cuesta un Map.get(), no un recorrido del catálogo.
  // ══════════════════════════════════════════════════════════════
  let _idx = null, _refProds = null, _refEquivs = null;

  function _indice() {
    const prods  = _prods();
    const equivs = _equivs();
    if (_idx && prods === _refProds && equivs === _refEquivs) return _idx;

    const m = new Map();
    const push = (r, item) => {
      let a = m.get(r);
      if (!a) { a = []; m.set(r, a); }
      a.push(item);
    };

    for (const p of prods) {
      const cb = _norm(p.codigoBarra);
      if (!cb) continue;
      push(raiz(cb), p);
    }

    // Equivalencias: un código físico de la familia puede estar registrado como
    // equivalente y apuntar al canónico. Se resuelve y se marca _scannedCb para
    // que el flujo consumidor sepa qué código leyó realmente el lector.
    const bySku = new Map();
    for (const p of prods) {
      if (!_canonico(p) || !_activo(p)) continue;
      const ks = [p.idProducto, p.skuBase, p.codigoBarra];
      for (const k of ks) {
        const kk = _norm(k);
        if (kk && !bySku.has(kk)) bySku.set(kk, p);
      }
    }
    for (const e of equivs) {
      const cb = _norm(e.codigoBarra);
      if (!cb) continue;
      const base = bySku.get(_norm(e.skuBase));
      if (!base) continue;
      push(raiz(cb), Object.assign({}, base, { _scannedCb: cb }));
    }

    // Dedup por idProducto dentro de cada raíz + orden estable (pelado, A, B, C…)
    m.forEach((arr, r) => {
      const vistos = new Set();
      const out = [];
      for (const p of arr) {
        const id = String(p.idProducto || p.codigoBarra || '');
        if (vistos.has(id)) continue;
        vistos.add(id);
        out.push(p);
      }
      out.sort((a, b) => {
        const sa = sufijo(a._scannedCb || a.codigoBarra);
        const sb = sufijo(b._scannedCb || b.codigoBarra);
        if (sa === sb) return String(a.descripcion || '').localeCompare(String(b.descripcion || ''));
        if (!sa) return -1;
        if (!sb) return 1;
        return sa < sb ? -1 : 1;
      });
      m.set(r, out);
    });

    _idx = m; _refProds = prods; _refEquivs = equivs;
    return _idx;
  }

  // ══════════════════════════════════════════════════════════════
  // API de consulta
  // ══════════════════════════════════════════════════════════════
  // opts: { soloCanonicos:bool (salidas: factor=1), incluirInactivos:bool }
  function familia(cod, opts) {
    opts = opts || {};
    const r = raiz(cod);
    if (!r || r.length < MIN_RAIZ) return [];
    let lista = (_indice().get(r) || []).slice();
    if (!opts.incluirInactivos) lista = lista.filter(_activo);
    if (opts.soloCanonicos)     lista = lista.filter(_canonico);
    return lista;
  }

  // Devuelve null si NO hay ambigüedad (el 99.9% de los escaneos).
  // Devuelve { raiz, miembros, ultima } si el operador tiene que elegir.
  function ambigua(cod, opts) {
    const c = _norm(cod);
    if (!c) return null;
    const r = raiz(c);
    if (!r || r.length < MIN_RAIZ) return null;
    // Vino CON letra → el operador ya fue explícito. No se le pregunta nada.
    if (c !== r) return null;
    const miembros = familia(r, opts);
    if (miembros.length < 2) return null;
    return { raiz: r, miembros: miembros, ultima: ultima(r) };
  }

  // ══════════════════════════════════════════════════════════════
  // MEMORIA DE ELECCIÓN — dentro de la MISMA operación (una guía, un
  // despacho, una sesión de catálogo). Si el operador vuelve a escanear
  // el mismo código, su última elección aparece ARRIBA y marcada.
  // ══════════════════════════════════════════════════════════════
  let _opKey = '';
  const _mem = new Map();

  function nuevaOperacion(key) {
    const k = String(key == null ? '' : key);
    if (k === _opKey) return;
    _opKey = k;
    _mem.clear();
  }
  function recordar(r, cb) {
    const rr = _norm(r);
    if (rr) _mem.set(rr, _norm(cb));
  }
  function ultima(r) { return _mem.get(_norm(r)) || ''; }
  function olvidarTodo() { _mem.clear(); }

  // ══════════════════════════════════════════════════════════════
  // SUFIJO LIBRE — para el alta de producto nuevo.
  // Las letras son SEMÁNTICAS (O=orégano, F=fina), así que no se propone
  // "la siguiente a la última": se propone la primera letra NO USADA.
  // ══════════════════════════════════════════════════════════════
  function sufijoLibre(cod) {
    const r = raiz(cod);
    if (!r || r.length < MIN_RAIZ) return '';
    const usados = new Set();
    familia(r, { incluirInactivos: true }).forEach(p => {
      const s = sufijo(p._scannedCb || p.codigoBarra);
      if (s) usados.add(s.toUpperCase());
    });
    for (const ch of ALFABETO) if (!usados.has(ch)) return ch;
    for (const a of ALFABETO) {
      for (const b of ALFABETO) {
        const s = a + b;
        if (!usados.has(s)) return s;
      }
    }
    return '';
  }
  // Código completo sugerido para un producto nuevo de esta familia.
  function siguienteCodigo(cod) {
    const r = raiz(cod);
    const s = sufijoLibre(cod);
    return (r && s) ? (r + s) : _norm(cod);
  }

  // ══════════════════════════════════════════════════════════════
  // FEEDBACK: sonido + vibración PROPIOS del caso ambiguo.
  // Distinto de beep (ok), warn (no existe) y scanIncompleto (prefijo).
  // ══════════════════════════════════════════════════════════════
  function sonarAmbiguo() {
    try {
      if (window.SoundFX && SoundFX.scanAmbiguo)          SoundFX.scanAmbiguo();
      else if (window.SoundFX && SoundFX.scanIncompleto)  SoundFX.scanIncompleto();
    } catch (_) {}
    try { if (typeof window.vibrate === 'function') window.vibrate([20, 45, 20, 45, 60]); } catch (_) {}
  }

  // ══════════════════════════════════════════════════════════════
  // SELECTOR — sheet de variantes. Prioridad: VELOCIDAD.
  // El operador tiene la caja en la mano: filas grandes, foto, nombre,
  // chip del sufijo y stock. Un tap y el flujo sigue donde estaba.
  // ══════════════════════════════════════════════════════════════
  function _stockMap() {
    const m = {};
    _stocks().forEach(s => { m[String(s.codigoProducto || s.idProducto || '')] = s; });
    return m;
  }
  function _stockDe(map, p) {
    const e = map[String(p.codigoBarra || '')] || map[String(p.idProducto || '')] || null;
    return e ? (parseFloat(e.cantidadDisponible) || 0) : 0;
  }

  const _TINTES = [
    ['rgba(99,102,241,.18)',  '#a5b4fc'], ['rgba(16,185,129,.18)', '#6ee7b7'],
    ['rgba(245,158,11,.18)',  '#fcd34d'], ['rgba(14,165,233,.18)', '#7dd3fc'],
    ['rgba(236,72,153,.18)',  '#f9a8d4'], ['rgba(168,85,247,.18)', '#d8b4fe'],
    ['rgba(20,184,166,.18)',  '#5eead4'], ['rgba(248,113,113,.18)','#fca5a5']
  ];
  function _tinte(key) {
    const s = String(key || '');
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return _TINTES[h % _TINTES.length];
  }
  function _iniciales(txt) {
    const partes = String(txt || '?').trim().split(/[\s\-_/.]+/).filter(Boolean);
    const a = (partes[0] || '?').charAt(0) || '?';
    const b = (partes[1] || '').charAt(0) || (partes[0] || '').charAt(1) || '';
    return (a + b).toUpperCase();
  }
  function _fotoSrc(url, ancho) {
    const u = String(url || '').trim();
    if (!u) return '';
    return (window.Photos && Photos.thumb) ? Photos.thumb(u, ancho) : u;
  }
  function _tileHTML(p) {
    const url = String(p.fotoUrl || p.foto || '').trim();
    const t   = _tinte(p.skuBase || p.idProducto || p.codigoBarra);
    const ini = _iniciales(p.descripcion || p.skuBase);
    const img = url
      ? `<img class="wh-foto-img" loading="lazy" decoding="async" src="${_esc(_fotoSrc(url, 160))}" alt="" onerror="this.remove()">`
      : '';
    return `<div class="wh-foto famcb-tile" style="--wf-bg:${t[0]};--wf-fg:${t[1]}">
      <span class="wh-foto-ini">${_esc(ini)}</span>${img}
    </div>`;
  }

  let _cssListo = false;
  function _css() {
    if (_cssListo || document.getElementById('famCbCss')) { _cssListo = true; return; }
    const st = document.createElement('style');
    st.id = 'famCbCss';
    st.textContent = `
      .famcb-ov{position:fixed;inset:0;z-index:${Z_SHEET};display:flex;align-items:flex-end;
        justify-content:center;background:rgba(2,6,23,.72);backdrop-filter:blur(3px);
        animation:famcbFade .12s ease-out}
      @keyframes famcbFade{from{opacity:0}to{opacity:1}}
      .famcb-sheet{width:100%;max-width:560px;max-height:88vh;display:flex;flex-direction:column;
        background:linear-gradient(180deg,#0f1d33,#0b1526);border:1px solid #26375a;
        border-bottom:none;border-radius:20px 20px 0 0;
        box-shadow:0 -18px 50px rgba(2,6,23,.7);overflow:hidden;
        padding-bottom:env(safe-area-inset-bottom);
        animation:famcbUp .16s cubic-bezier(.22,1,.36,1)}
      @keyframes famcbUp{from{transform:translateY(26px);opacity:.4}to{transform:translateY(0);opacity:1}}
      @media (min-width:720px){
        .famcb-ov{align-items:center}
        .famcb-sheet{border-radius:20px;border-bottom:1px solid #26375a;max-height:82vh}
      }
      .famcb-hdr{padding:14px 16px 10px;border-bottom:1px solid rgba(51,65,85,.55);
        display:flex;align-items:flex-start;gap:11px;flex-shrink:0}
      .famcb-ico{font-size:22px;line-height:1;flex-shrink:0;margin-top:1px}
      .famcb-ttl{font-size:14px;font-weight:900;color:#fbbf24;letter-spacing:.2px}
      .famcb-sub{font-size:11px;color:#94a3b8;margin-top:3px;line-height:1.4}
      .famcb-cod{font-family:'Consolas',monospace;color:#7dd3fc;font-weight:800}
      .famcb-x{flex-shrink:0;width:30px;height:30px;border-radius:9px;border:1px solid #26375a;
        background:rgba(15,23,42,.6);color:#94a3b8;font-size:15px;font-weight:800;cursor:pointer;
        -webkit-tap-highlight-color:transparent;touch-action:manipulation}
      .famcb-list{overflow-y:auto;-webkit-overflow-scrolling:touch;padding:10px 12px 4px;flex:1}
      .famcb-row{width:100%;display:flex;align-items:center;gap:11px;text-align:left;
        padding:9px 11px;margin-bottom:8px;border-radius:14px;cursor:pointer;
        background:#16233c;border:1px solid #26375a;color:inherit;
        -webkit-tap-highlight-color:transparent;touch-action:manipulation;
        transition:transform .1s ease,border-color .1s ease,background .1s ease}
      .famcb-row:active{transform:scale(.975);border-color:#818cf8;background:#1b2a48}
      .famcb-row.is-ultima{border-color:rgba(251,191,36,.65);background:rgba(251,191,36,.07)}
      .famcb-tile{width:52px;height:52px;border-radius:13px;flex-shrink:0}
      .famcb-tile .wh-foto-ini{font-size:16px}
      .famcb-body{flex:1;min-width:0}
      .famcb-name{font-size:13.5px;font-weight:800;color:#f1f5f9;line-height:1.25;
        display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
      .famcb-meta{display:flex;align-items:center;gap:6px;margin-top:5px;flex-wrap:wrap}
      .famcb-chip{font-family:'Consolas',monospace;font-size:10.5px;font-weight:900;
        padding:2px 8px;border-radius:999px;background:rgba(129,140,248,.16);
        color:#a5b4fc;border:1px solid rgba(129,140,248,.35);white-space:nowrap}
      .famcb-chip b{color:#fbbf24}
      .famcb-stock{font-size:10.5px;font-weight:800;color:#6ee7b7;white-space:nowrap}
      .famcb-stock.is-cero{color:#f87171}
      .famcb-ultima-tag{font-size:9.5px;font-weight:900;color:#0a1424;background:#fbbf24;
        border-radius:999px;padding:2px 7px;letter-spacing:.3px;white-space:nowrap}
      .famcb-num{flex-shrink:0;width:26px;height:26px;border-radius:8px;background:rgba(15,23,42,.75);
        border:1px solid #334155;color:#94a3b8;font-size:12px;font-weight:900;
        display:flex;align-items:center;justify-content:center}
      .famcb-foot{padding:4px 12px 12px;flex-shrink:0}
      .famcb-new{width:100%;display:flex;align-items:center;gap:10px;text-align:left;
        padding:10px 12px;border-radius:13px;cursor:pointer;background:rgba(124,58,237,.09);
        border:1.5px dashed rgba(167,139,250,.6);color:inherit;
        -webkit-tap-highlight-color:transparent;touch-action:manipulation}
      .famcb-new:active{transform:scale(.98)}
      .famcb-new-t{font-size:12.5px;font-weight:900;color:#c084fc}
      .famcb-new-s{font-size:10.5px;color:#94a3b8;margin-top:2px}
      .famcb-hint{text-align:center;font-size:10px;color:#5f7290;padding:2px 0 6px}
      @media (prefers-reduced-motion:reduce){
        .famcb-ov,.famcb-sheet{animation:none}
        .famcb-row{transition:none}
      }
      /* ── Aviso de familia dentro del modal de PRODUCTO NUEVO ── */
      .famcb-pn{background:rgba(251,191,36,.07);border:1px solid rgba(251,191,36,.42);
        border-radius:11px;padding:10px 12px;margin-bottom:12px}
      .famcb-pn-t{font-size:12px;font-weight:900;color:#fbbf24;line-height:1.35}
      .famcb-pn-s{font-size:10.5px;color:#94a3b8;margin-top:3px;line-height:1.45}
      .famcb-pn-list{margin-top:9px;max-height:172px;overflow-y:auto}
      .famcb-pn-item{width:100%;display:flex;align-items:center;gap:9px;text-align:left;
        padding:7px 9px;margin-bottom:6px;border-radius:10px;background:#16233c;
        border:1px solid #26375a;color:inherit;cursor:pointer;
        -webkit-tap-highlight-color:transparent;touch-action:manipulation}
      .famcb-pn-item:active{transform:scale(.98);border-color:#818cf8}
      .famcb-pn-item .famcb-tile{width:38px;height:38px;border-radius:10px}
      .famcb-pn-item .famcb-tile .wh-foto-ini{font-size:13px}
      .famcb-pn-name{font-size:12px;font-weight:800;color:#f1f5f9;line-height:1.25;
        display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
      .famcb-pn-btn{width:100%;margin-top:4px;padding:9px 12px;border-radius:10px;
        background:linear-gradient(135deg,#7c3aed,#6d28d9);border:1px solid #8b5cf6;
        color:#fff;font-size:12px;font-weight:900;cursor:pointer;
        -webkit-tap-highlight-color:transparent;touch-action:manipulation}
      .famcb-pn-btn:active{transform:scale(.985)}
    `;
    document.head.appendChild(st);
    _cssListo = true;
  }

  let _abierto = null;   // { ov, onPick, resuelto }

  function cerrar(resultado) {
    const st = _abierto;
    _abierto = null;
    if (!st) return;
    try { document.removeEventListener('keydown', st.onKey, true); } catch (_) {}
    try { st.ov.remove(); } catch (_) {}
    if (typeof st.onPick === 'function') { try { st.onPick(resultado || null); } catch (e) { console.error('[FamiliaCB] onPick', e); } }
  }

  function abierto() { return !!_abierto; }

  /**
   * Muestra el selector de variantes.
   * @param {string}   cod    código escaneado (la raíz pelada)
   * @param {object}   opts   { miembros, soloCanonicos, permitirNuevo, textoNuevo, subtitulo }
   * @param {function} onPick recibe el producto elegido, o null si canceló,
   *                          o { _nuevo:true, codigo, sugerido } si eligió "es otro".
   */
  function elegir(cod, opts, onPick) {
    opts = opts || {};
    const r   = raiz(cod);
    const mem = opts.miembros || familia(r, opts);
    if (!mem.length)  { if (onPick) onPick(null); return; }
    if (mem.length === 1 && !opts.forzar) { if (onPick) onPick(mem[0]); return; }

    _css();
    if (_abierto) cerrar(null);

    const smap = _stockMap();
    const ult  = ultima(r);

    // La última elegida en esta operación va PRIMERO y marcada.
    const orden = mem.slice();
    if (ult) {
      const i = orden.findIndex(p => _norm(p._scannedCb || p.codigoBarra) === ult);
      if (i > 0) orden.unshift(orden.splice(i, 1)[0]);
    }

    const filas = orden.map((p, i) => {
      const cb   = String(p._scannedCb || p.codigoBarra || '');
      const suf  = sufijo(cb);
      const cola = r.length > 4 ? r.slice(-4) : r;
      const chip = suf
        ? `…${_esc(cola)}<b>${_esc(suf)}</b>`
        : `…${_esc(cola)} <span style="color:#94a3b8">sin letra</span>`;
      const st   = _stockDe(smap, p);
      const esU  = ult && _norm(cb) === ult;
      return `<button type="button" class="famcb-row${esU ? ' is-ultima' : ''}" data-famcb-cb="${_esc(cb)}">
        ${_tileHTML(p)}
        <div class="famcb-body">
          <div class="famcb-name">${_esc(p.descripcion || cb)}</div>
          <div class="famcb-meta">
            <span class="famcb-chip">${chip}</span>
            <span class="famcb-stock${st > 0 ? '' : ' is-cero'}">${st > 0 ? 'Stock ' + _esc(_qty(st)) : 'Sin stock'}</span>
            ${esU ? '<span class="famcb-ultima-tag">LA ÚLTIMA</span>' : ''}
          </div>
        </div>
        <span class="famcb-num">${i < 9 ? (i + 1) : '·'}</span>
      </button>`;
    }).join('');

    const pie = opts.permitirNuevo
      ? `<div class="famcb-foot">
           <button type="button" class="famcb-new" data-famcb-nuevo="1">
             <span style="font-size:19px">🆕</span>
             <span style="flex:1;min-width:0">
               <span class="famcb-new-t" style="display:block">${_esc(opts.textoNuevo || 'Ninguno · es un producto nuevo')}</span>
               <span class="famcb-new-s" style="display:block">Se registra como <b style="color:#c084fc">${_esc(siguienteCodigo(r))}</b></span>
             </span>
           </button>
         </div>`
      : '';

    const ov = document.createElement('div');
    ov.className = 'famcb-ov';
    ov.setAttribute('role', 'dialog');
    ov.setAttribute('aria-modal', 'true');
    ov.innerHTML = `
      <div class="famcb-sheet" role="document">
        <div class="famcb-hdr">
          <span class="famcb-ico">🔀</span>
          <div style="flex:1;min-width:0">
            <div class="famcb-ttl">Este código tiene ${orden.length} variantes</div>
            <div class="famcb-sub">${_esc(opts.subtitulo || 'El proveedor usa el mismo código físico para varios productos.')}
              <br><span class="famcb-cod">${_esc(r)}</span> · toca cuál tienes en la mano</div>
          </div>
          <button type="button" class="famcb-x" data-famcb-cerrar="1" aria-label="Cerrar">✕</button>
        </div>
        <div class="famcb-list">${filas}</div>
        ${pie}
        <div class="famcb-hint">Teclas 1-${Math.min(9, orden.length)} para elegir · Esc cancela</div>
      </div>`;

    const pick = (cb) => {
      const p = orden.find(x => String(x._scannedCb || x.codigoBarra || '') === cb);
      if (!p) return;
      recordar(r, cb);
      try { if (window.SoundFX && SoundFX.beep) SoundFX.beep(); } catch (_) {}
      try { if (typeof window.vibrate === 'function') window.vibrate(15); } catch (_) {}
      cerrar(p);
    };

    ov.addEventListener('click', (e) => {
      if (e.target === ov) { cerrar(null); return; }
      const cerrarBtn = e.target.closest('[data-famcb-cerrar]');
      if (cerrarBtn) { cerrar(null); return; }
      const nuevo = e.target.closest('[data-famcb-nuevo]');
      if (nuevo) { cerrar({ _nuevo: true, codigo: r, sugerido: siguienteCodigo(r) }); return; }
      const row = e.target.closest('[data-famcb-cb]');
      if (row) pick(row.getAttribute('data-famcb-cb'));
    });

    const onKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cerrar(null); return; }
      // El lector HID teclea dígitos: se aceptan solo cuando NO hay input enfocado.
      const act = document.activeElement;
      if (act && /^(INPUT|TEXTAREA|SELECT)$/.test(act.tagName)) return;
      const n = parseInt(e.key, 10);
      if (n >= 1 && n <= Math.min(9, orden.length)) {
        e.preventDefault(); e.stopPropagation();
        const p = orden[n - 1];
        pick(String(p._scannedCb || p.codigoBarra || ''));
      }
    };
    document.addEventListener('keydown', onKey, true);

    document.body.appendChild(ov);
    _abierto = { ov: ov, onPick: onPick, onKey: onKey };
    sonarAmbiguo();
  }

  /**
   * Atajo para los flujos de escaneo: si el código es ambiguo abre el selector
   * y devuelve true (el flujo continúa dentro del callback). Si no, devuelve
   * false y el llamador sigue con su camino de siempre — CERO latencia añadida.
   */
  function interceptar(cod, opts, onElegido) {
    if (_abierto) return true;   // ya está eligiendo: no reprocesar el mismo escaneo
    const amb = ambigua(cod, opts);
    if (!amb) return false;
    elegir(amb.raiz, Object.assign({ miembros: amb.miembros }, opts || {}), onElegido);
    return true;
  }

  // ══════════════════════════════════════════════════════════════
  // HTML del aviso para el modal de PRODUCTO NUEVO.
  // Devuelve '' si el código no pertenece a ninguna familia existente.
  // ══════════════════════════════════════════════════════════════
  function avisoPNHtml(cod, opts) {
    opts = opts || {};
    const c = _norm(cod);
    const r = raiz(c);
    if (!r || r.length < MIN_RAIZ) return '';
    const mem = familia(r, {});
    if (!mem.length) return '';
    // Si ya viene con una letra LIBRE, no hay nada que avisar: el operador ya desambiguó.
    if (c !== r && !mem.some(p => _norm(p._scannedCb || p.codigoBarra) === c)) return '';

    const smap = _stockMap();
    const sugerido = siguienteCodigo(r);
    const chips = mem.map(p => {
      const cb  = String(p._scannedCb || p.codigoBarra || '');
      const suf = sufijo(cb);
      const cola = r.length > 4 ? r.slice(-4) : r;
      return '…' + _esc(cola) + (suf ? _esc(suf) : '');
    }).join(', ');

    const items = mem.map(p => {
      const cb  = String(p._scannedCb || p.codigoBarra || '');
      const suf = sufijo(cb);
      const cola = r.length > 4 ? r.slice(-4) : r;
      const st  = _stockDe(smap, p);
      return `<button type="button" class="famcb-pn-item" data-famcb-pn-cb="${_esc(cb)}">
        ${_tileHTML(p)}
        <span style="flex:1;min-width:0">
          <span class="famcb-pn-name" style="display:block">${_esc(p.descripcion || cb)}</span>
          <span class="famcb-meta" style="display:flex">
            <span class="famcb-chip">…${_esc(cola)}${suf ? '<b>' + _esc(suf) + '</b>' : ''}</span>
            <span class="famcb-stock${st > 0 ? '' : ' is-cero'}">${st > 0 ? 'Stock ' + _esc(_qty(st)) : 'Sin stock'}</span>
          </span>
        </span>
      </button>`;
    }).join('');

    return `<div class="famcb-pn" id="pnFamiliaBox">
      <div class="famcb-pn-t">⚠ Este código ya tiene ${mem.length} variante${mem.length === 1 ? '' : 's'} (${chips})</div>
      <div class="famcb-pn-s">El proveedor reusa el mismo código físico. Elige cuál es, o registra uno nuevo con su letra.</div>
      <div class="famcb-pn-list">${items}</div>
      <button type="button" class="famcb-pn-btn" data-famcb-pn-nuevo="${_esc(sugerido)}">
        ➕ Es uno NUEVO · usar ${_esc(sugerido)}
      </button>
    </div>`;
  }

  window.FamiliaCB = {
    raiz: raiz,
    sufijo: sufijo,
    familia: familia,
    ambigua: ambigua,
    elegir: elegir,
    interceptar: interceptar,
    cerrar: cerrar,
    abierto: abierto,
    recordar: recordar,
    ultima: ultima,
    nuevaOperacion: nuevaOperacion,
    olvidarTodo: olvidarTodo,
    sufijoLibre: sufijoLibre,
    siguienteCodigo: siguienteCodigo,
    sonarAmbiguo: sonarAmbiguo,
    avisoPNHtml: avisoPNHtml,
    _MIN_RAIZ: MIN_RAIZ
  };
})();
