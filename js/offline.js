// warehouseMos — offline.js
// Caché local, cola offline y sincronización sin duplicados
'use strict';

const OfflineManager = (() => {
  const KEYS = {
    PERSONAL:      'wh_personal',
    PRODUCTOS:     'wh_productos',
    EQUIVALENCIAS: 'wh_equivalencias',
    STOCK:         'wh_stock',
    PROVEEDORES:   'wh_proveedores',
    IMPRESORAS:    'wh_impresoras',
    ZONAS:         'wh_zonas',
    CONFIG:      'wh_config',
    QUEUE:       'wh_queue',
    LAST_SYNC:   'wh_last_sync',
    // Datos operacionales (guías, preingresos, stock, ajustes, auditorías)
    GUIAS:        'wh_guias',
    GUIA_DETALLE: 'wh_guia_detalle',
    PREINGRESOS:  'wh_preingresos',
    AJUSTES:      'wh_ajustes',
    AUDITORIAS_C: 'wh_auditorias_c',
    ADMIN_PIN:    'wh_admin_pin',          // legacy — único PIN local (se eliminará)
    ADMIN_CACHE:  'wh_admin_cache',        // nuevo — { globalPin, adminPins[], sincronizadoEn }
    LAST_MASTER:  'wh_last_master',
    PN:           'wh_pn',
    ENVASADOS:    'wh_envasados'
  };

  // ── Estado ────────────────────────────────────────────────
  let _syncing       = false;
  let _opLoading     = false;  // guard: evitar llamadas concurrent a precargarOperacional
  let _lastOpTs      = 0;      // timestamp de la última llamada a precargarOperacional
  const OP_MIN_MS    = 15000;  // mínimo 15s entre llamadas operacionales
  let _masterLoading = false;  // guard: evitar llamadas concurrent a precargar (maestros)
  let _lastMasterTs  = 0;      // timestamp de la última llamada a precargar
  const MASTER_MIN_MS = 60000; // maestros: mínimo 60s (cambian poco)
  let _onStatusChange = null;
  let _opRefreshTimer = null;

  function onStatusChange(fn) { _onStatusChange = fn; }

  function _notificar() {
    if (_onStatusChange) _onStatusChange({
      online:   navigator.onLine,
      pending:  getQueue().filter(i => i.status === 'pending').length,
      syncing:  _syncing,
      lastSync: localStorage.getItem(KEYS.LAST_SYNC)
    });
  }

  window.addEventListener('online',  () => { _notificar(); sincronizar(); });
  window.addEventListener('offline', () => _notificar());

  // ── Cache helpers ─────────────────────────────────────────
  function guardar(key, data) {
    try { localStorage.setItem(key, JSON.stringify({ data, ts: Date.now() })); }
    catch(e) { console.warn('[Offline] localStorage lleno:', e); }
  }

  function cargar(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw).data : null;
    } catch { return null; }
  }

  // ── Precarga de datos de referencia ──────────────────────
  // Intenta descargarMaestros (endpoint nuevo que trae las 4 tablas de MOS).
  // Si el GAS desplegado no lo conoce aún, degrada al endpoint legacy
  // getPersonalConPin para que el login siempre funcione.
  async function precargar(forzar = false) {
    if (!navigator.onLine) return;
    if (_masterLoading) return;
    if (!forzar && Date.now() - _lastMasterTs < MASTER_MIN_MS) return;
    _masterLoading = true;
    _lastMasterTs  = Date.now();
    try {
      const [maestros, stock, config] = await Promise.all([
        fetch(_gasUrl('descargarMaestros')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getStock')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getConfig')).then(r => r.json()).catch(() => null)
      ]);

      console.log('[Offline] descargarMaestros respuesta:', maestros);
      const maestrosChanged = [];
      if (maestros?.ok) {
        // GAS nuevo — recibe las 4 tablas de MOS
        const d = maestros.data;
        console.log('[Offline] personal recibido:', d?.personal?.length, 'registros');
        if (d.personal      != null) guardar(KEYS.PERSONAL,      d.personal);
        if (d.productos     != null) { guardar(KEYS.PRODUCTOS,     d.productos);     maestrosChanged.push('productos'); }
        if (d.equivalencias != null) { guardar(KEYS.EQUIVALENCIAS, d.equivalencias); maestrosChanged.push('equivalencias'); }
        if (d.proveedores   != null) guardar(KEYS.PROVEEDORES,   d.proveedores);
        if (d.impresoras    != null) guardar(KEYS.IMPRESORAS,    d.impresoras);
        if (d.zonas         != null) guardar(KEYS.ZONAS,         d.zonas);
        if (d.adminPin)            localStorage.setItem(KEYS.ADMIN_PIN, d.adminPin);
        if (maestros.errores?.length) console.warn('[Offline] descargarMaestros errores:', maestros.errores);
      } else {
        // GAS antiguo o MOS no configurado — degradar a endpoints individuales
        console.warn('[Offline] descargarMaestros no disponible, usando endpoints legacy');
        const [personal, productos, proveedores] = await Promise.all([
          fetch(_gasUrl('getPersonalConPin')).then(r => r.json()).catch(() => null),
          fetch(_gasUrl('getProductos&estado=1')).then(r => r.json()).catch(() => null),
          fetch(_gasUrl('getProveedores&estado=1')).then(r => r.json()).catch(() => null)
        ]);
        console.log('[Offline] legacy getPersonalConPin:', personal);
        if (personal?.ok)    guardar(KEYS.PERSONAL,    personal.data);
        if (productos?.ok)   { guardar(KEYS.PRODUCTOS,   productos.data);   maestrosChanged.push('productos'); }
        if (proveedores?.ok) guardar(KEYS.PROVEEDORES, proveedores.data);
      }

      if (stock?.ok)  guardar(KEYS.STOCK,  stock.data);
      if (config?.ok) guardar(KEYS.CONFIG, config.data);

      localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
      if (maestrosChanged.length) {
        window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: maestrosChanged } }));
      }
      _notificar();
    } catch(e) {
      console.warn('[Offline] Error en precarga:', e);
    } finally {
      _masterLoading = false;
    }
  }

  function _gasUrl(action) {
    return `${window.WH_CONFIG.gasUrl}?action=${action}`;
  }

  // ── Validación PIN local (instantánea) ───────────────────
  function validarPinLocal(pin) {
    const personal = cargar(KEYS.PERSONAL) || [];
    return personal.find(p =>
      String(p.pin) === String(pin) && String(p.estado) === '1'
    ) || null;
  }

  // ── Cola offline ──────────────────────────────────────────
  function getQueue() {
    return cargar(KEYS.QUEUE) || [];
  }

  function encolar(action, params) {
    // Reutilizar localId si ya viene (idempotencia: api.js inyecta uno antes
    // del primer POST, y si la red falla y caemos en cola offline, debemos
    // preservarlo para que GAS deduplique al sincronizar).
    const localId = params?.localId || ('L' + Date.now() + Math.random().toString(36).substr(2, 5));
    const item = { localId, action, params: { ...params, localId }, ts: Date.now(), status: 'pending' };
    const queue = getQueue();
    queue.push(item);
    guardar(KEYS.QUEUE, queue);
    _notificar();
    return localId;
  }

  function _actualizarItemQueue(localId, status) {
    const queue = getQueue().map(i => i.localId === localId ? { ...i, status } : i);
    guardar(KEYS.QUEUE, queue);
  }

  function limpiarSincronizados() {
    const queue = getQueue().filter(i => i.status === 'pending' || i.status === 'error');
    guardar(KEYS.QUEUE, queue);
  }

  // ── Sincronización ────────────────────────────────────────
  async function sincronizar() {
    if (!navigator.onLine || _syncing) return;
    // [Fix v2.9.1] Antes solo procesábamos status='pending'. Pero
    // limpiarSincronizados() preservaba items con status='error' y nunca
    // los reintentábamos → quedaban infinitamente en la cola disparando
    // "X operaciones por sincronizar". Ahora reintentamos también los error.
    const queue = getQueue().filter(i => i.status === 'pending' || i.status === 'error');
    if (!queue.length) return;

    _syncing = true;
    _notificar();

    var huboEnvasado = false;
    for (const item of queue) {
      try {
        const res = await fetch(window.WH_CONFIG.gasUrl, {
          method:  'POST',
          headers: { 'Content-Type': 'text/plain' },
          body:    JSON.stringify(item.params)
        }).then(r => r.json());

        _actualizarItemQueue(item.localId, res.ok ? 'synced' : 'error');
        if (res.ok && item.action === 'registrarEnvasado') huboEnvasado = true;
      } catch {
        _actualizarItemQueue(item.localId, 'error');
      }
    }

    limpiarSincronizados();
    _syncing = false;
    localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
    _notificar();

    // [Fix v2.9.1] Si se sincronizó un envasado optimistic, avisar a la UI
    // para que recargue desde el backend y reemplace los ENV_OPT_* por reales.
    if (huboEnvasado) {
      try { window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['envasados'] } })); } catch(_){}
    }
  }

  // ── Precarga operacional (guías, preingresos, stock, ajustes, auditorías) ──
  // [Fix #2 v2.11.1] Flag global "subiendo fotos en background".
  // Mientras esté en true, precargarOperacional() salta el refresh de
  // preingresos para que la respuesta del backend (que aún no tiene las
  // fotos asociadas porque están subiéndose) no pise el cache local con
  // fotos vacías. Lo activa/desactiva PreingresosView durante el subir
  // las fotos al Drive.
  let _subiendoFotos = false;
  function setSubiendoFotos(on) { _subiendoFotos = !!on; }
  function isSubiendoFotos() { return _subiendoFotos; }

  async function precargarOperacional(forzar = false) {
    if (!navigator.onLine) return;
    if (_opLoading) return;
    if (!forzar && Date.now() - _lastOpTs < OP_MIN_MS) return;
    _opLoading = true;
    _lastOpTs  = Date.now();
    try {
      const r = await fetch(_gasUrl('descargarOperacional')).then(r => r.json()).catch(() => null);
      if (!r?.ok) return;
      const d = r.data;
      const changed = [];

      function _hayDiff(newArr, key) {
        if (!newArr?.length) return false;
        const old = cargar(key);
        if (!old || old.length !== newArr.length) return true;
        return JSON.stringify(newArr) !== JSON.stringify(old);
      }

      // [Fix #1 v2.11.1] Merge-on-refresh para preingresos.
      // Antes el polling sobrescribía las fotos del cache local con un
      // valor vacío si el backend aún no había sincronizado la subida
      // background. Resultado: el operador subía 8 fotos, esperaba unos
      // segundos y veía la lista vacía aunque las fotos sí estaban en
      // Drive. Ahora preservamos el campo `fotos` del cache local si la
      // versión del backend lo trae vacío pero el local ya tenía algo.
      function _mergePreingresos(neuvos, viejos) {
        const oldMap = {};
        (viejos || []).forEach(v => { oldMap[String(v.idPreingreso)] = v; });
        return neuvos.map(n => {
          const old = oldMap[String(n.idPreingreso)];
          if (!old) return n;
          // Preservar fotos si llegan vacías y locales tienen contenido.
          // Mismo principio para fotosFileIds si se usara en el futuro.
          const merged = { ...n };
          if ((!merged.fotos || merged.fotos === '') && old.fotos) {
            merged.fotos = old.fotos;
            if (window.__WH_DEBUG_FOTOS) console.log('[Offline merge] preservé fotos de', n.idPreingreso, '→', old.fotos.substring(0, 60));
          }
          return merged;
        });
      }

      if (d.guias       != null) { if (_hayDiff(d.guias,       KEYS.GUIAS))        { guardar(KEYS.GUIAS,        d.guias);       changed.push('guias'); }       else guardar(KEYS.GUIAS, d.guias); }
      if (d.detalles    != null) { if (_hayDiff(d.detalles,    KEYS.GUIA_DETALLE)) { guardar(KEYS.GUIA_DETALLE, d.detalles);    changed.push('detalles'); }   else guardar(KEYS.GUIA_DETALLE, d.detalles); }
      if (d.preingresos != null) {
        // Si hay subida de fotos en curso, omitir refresh de preingresos
        // (Fix #2). Si no, aplicar merge defensivo (Fix #1) y guardar.
        if (_subiendoFotos) {
          if (window.__WH_DEBUG_FOTOS) console.log('[Offline] skip refresh preingresos: subida en curso');
        } else {
          const viejos  = cargar(KEYS.PREINGRESOS) || [];
          const merged  = _mergePreingresos(d.preingresos, viejos);
          if (_hayDiff(merged, KEYS.PREINGRESOS)) { guardar(KEYS.PREINGRESOS, merged); changed.push('preingresos'); }
          else                                    { guardar(KEYS.PREINGRESOS, merged); }
        }
      }
      if (d.stock       != null) { if (_hayDiff(d.stock,       KEYS.STOCK))        { guardar(KEYS.STOCK,        d.stock);       changed.push('stock'); }       else guardar(KEYS.STOCK, d.stock); }
      if (d.ajustes     != null) { if (_hayDiff(d.ajustes,     KEYS.AJUSTES))      { guardar(KEYS.AJUSTES,      d.ajustes);     changed.push('ajustes'); }     else guardar(KEYS.AJUSTES, d.ajustes); }
      if (d.auditorias  != null) { if (_hayDiff(d.auditorias,  KEYS.AUDITORIAS_C)) { guardar(KEYS.AUDITORIAS_C, d.auditorias);  changed.push('auditorias'); }  else guardar(KEYS.AUDITORIAS_C, d.auditorias); }

      window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed } }));
    } catch(e) { console.warn('[Offline] Error en precarga operacional:', e); }
    finally { _opLoading = false; }
  }

  // Inicia el refresh automático cada 60s (llamar desde App.init, antes del login)
  function iniciarRefreshOperacional() {
    if (_opRefreshTimer) return;
    // Carga inmediata: maestros (si no hay caché) + operacional
    // precargar() ya tiene throttle propio (MASTER_MIN_MS=60s) así que es seguro llamarlo
    if (!cargar(KEYS.PERSONAL)?.length) {
      precargar(true).catch(() => {}); // forzar: primer arranque sin caché
    } else {
      precargar().catch(() => {});     // respetará throttle de 60s
    }
    precargarOperacional();
    _opRefreshTimer = setInterval(() => {
      precargar().catch(() => {});     // throttled internamente
      precargarOperacional();          // throttled internamente
    }, 60000);
  }

  function detenerRefreshOperacional() {
    if (_opRefreshTimer) { clearInterval(_opRefreshTimer); _opRefreshTimer = null; }
  }

  // ── Getters cache ─────────────────────────────────────────
  // _guardarPersonalConPin mantenido por compatibilidad con llamada puntual al iniciar sesión
  function _guardarPersonalConPin(data) { guardar(KEYS.PERSONAL, data); }

  const getPersonalCache      = () => cargar(KEYS.PERSONAL)      || [];
  const getProductosCache     = () => cargar(KEYS.PRODUCTOS)     || [];
  const getEquivalenciasCache = () => cargar(KEYS.EQUIVALENCIAS) || [];
  const getStockCache         = () => cargar(KEYS.STOCK)         || [];
  const getProveedoresCache   = () => cargar(KEYS.PROVEEDORES)   || [];
  const getImpresorasCache    = () => cargar(KEYS.IMPRESORAS)    || [];
  const getZonasCache         = () => cargar(KEYS.ZONAS)         || [];
  const getConfigCache        = () => cargar(KEYS.CONFIG)        || {};
  const getGuiasCache         = () => cargar(KEYS.GUIAS)         || [];
  const getGuiaDetalleCache   = () => cargar(KEYS.GUIA_DETALLE)  || [];
  const getPreingresosCache   = () => cargar(KEYS.PREINGRESOS)   || [];
  const getAjustesCache       = () => cargar(KEYS.AJUSTES)       || [];
  const getAuditoriasCache    = () => cargar(KEYS.AUDITORIAS_C)  || [];
  const getAdminPin           = () => localStorage.getItem(KEYS.ADMIN_PIN) || null;
  const getAdminCache         = () => cargar(KEYS.ADMIN_CACHE) || null;

  async function sincronizarAdminCache() {
    const url = window.WH_CONFIG?.mosGasUrl;
    if (!url || !navigator.onLine) return false;
    try {
      const r = await fetch(url + '?action=getAdminPinsCache');
      const j = await r.json();
      if (!j?.ok || !j.data?.globalPin) return false;
      guardar(KEYS.ADMIN_CACHE, {
        globalPin: String(j.data.globalPin),
        adminPins: Array.isArray(j.data.adminPins) ? j.data.adminPins : [],
        sincronizadoEn: Date.now()
      });
      return true;
    } catch(e) { return false; }
  }
  const getPNCache            = () => cargar(KEYS.PN)            || [];
  const setPNCache            = (v) => guardar(KEYS.PN, v);
  const getEnvasadosCache     = () => cargar(KEYS.ENVASADOS)     || [];
  const guardarEnvasadosCache = (v) => guardar(KEYS.ENVASADOS, v);

  function inyectarEnvasadoCache(item) {
    const cache = getEnvasadosCache();
    if (!cache.find(x => x.idEnvasado === item.idEnvasado)) {
      cache.unshift(item);
      guardar(KEYS.ENVASADOS, cache);
    }
  }

  // Quita un envasado del cache (rollback optimista cuando el backend falla)
  function removerEnvasadoCache(idEnvasado) {
    const cache = getEnvasadosCache();
    const filtrado = cache.filter(x => x.idEnvasado !== idEnvasado);
    if (filtrado.length !== cache.length) {
      guardar(KEYS.ENVASADOS, filtrado);
      return true;
    }
    return false;
  }

  // ── Patch de un preingreso existente en caché ───────────────
  function patchPreingresosCache(id, changes) {
    const cache = cargar(KEYS.PREINGRESOS) || [];
    const idx   = cache.findIndex(x => x.idPreingreso === id);
    if (idx >= 0) { Object.assign(cache[idx], changes); guardar(KEYS.PREINGRESOS, cache); }
  }

  // ── Inyectar un preingreso recién creado en caché ────────────
  function inyectarPreingreso(item) {
    const cache = getPreingresosCache();
    if (!cache.find(x => x.idPreingreso === item.idPreingreso)) {
      cache.unshift(item);
      guardar(KEYS.PREINGRESOS, cache);
    }
  }

  // ── Actualizar cache de detalle para una guía específica ─────
  // Reemplaza todas las entradas de idGuia con los nuevos detalles
  function actualizarDetallesGuia(idGuia, nuevosDetalles) {
    const cache = getGuiaDetalleCache();
    const otros = cache.filter(d => d.idGuia !== idGuia);
    const estos = nuevosDetalles.map(d => ({ ...d, idGuia: d.idGuia || idGuia }));
    guardar(KEYS.GUIA_DETALLE, [...otros, ...estos]);
  }

  // ── Agregar o reemplazar una entrada en cache de detalle ─────
  function addDetalleCache(detalle) {
    const cache = getGuiaDetalleCache();
    const idx = cache.findIndex(d => d.idDetalle === detalle.idDetalle);
    if (idx >= 0) cache[idx] = detalle;
    else cache.push(detalle);
    guardar(KEYS.GUIA_DETALLE, cache);
  }

  // ── Patch optimista de stock local ───────────────────────────
  // Aplica delta a cantidadDisponible sin esperar a GAS.
  // Si el producto no tiene fila en STOCK aún, crea una entrada temporal.
  function patchStockCache(codigoBarra, delta) {
    const stock = cargar(KEYS.STOCK) || [];
    const cb  = String(codigoBarra);
    const idx = stock.findIndex(s => String(s.codigoProducto) === cb);
    if (idx >= 0) {
      stock[idx] = {
        ...stock[idx],
        cantidadDisponible: (parseFloat(stock[idx].cantidadDisponible) || 0) + delta
      };
    } else {
      stock.push({ idStock: 'STK_L' + Date.now(), codigoProducto: cb, cantidadDisponible: delta });
    }
    guardar(KEYS.STOCK, stock);
  }

  return {
    precargar, sincronizar, encolar, getQueue,
    validarPinLocal, onStatusChange,
    _guardarPersonalConPin,
    getPersonalCache, getProductosCache, getEquivalenciasCache,
    getStockCache, getProveedoresCache,
    getImpresorasCache, getZonasCache, getConfigCache,
    getGuiasCache, getGuiaDetalleCache, getPreingresosCache,
    getAjustesCache, getAuditoriasCache,
    getAdminPin, getAdminCache, sincronizarAdminCache,
    actualizarDetallesGuia, addDetalleCache, inyectarPreingreso, patchPreingresosCache, patchStockCache,
    getPNCache, setPNCache,
    getEnvasadosCache, guardarEnvasadosCache, inyectarEnvasadoCache, removerEnvasadoCache,
    precargarOperacional, iniciarRefreshOperacional, detenerRefreshOperacional,
    setSubiendoFotos, isSubiendoFotos,
    estaOnline: () => navigator.onLine
  };
})();
