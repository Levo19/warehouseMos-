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
    CONFIG:      'wh_config',
    QUEUE:       'wh_queue',
    LAST_SYNC:   'wh_last_sync',
    // Datos operacionales (guías, preingresos, stock, ajustes, auditorías)
    GUIAS:        'wh_guias',
    GUIA_DETALLE: 'wh_guia_detalle',
    PREINGRESOS:  'wh_preingresos',
    AJUSTES:      'wh_ajustes',
    AUDITORIAS_C: 'wh_auditorias_c',
    ADMIN_PIN:    'wh_admin_pin',
    LAST_MASTER:  'wh_last_master'
  };

  // ── Estado ────────────────────────────────────────────────
  let _syncing = false;
  let _onStatusChange = null;
  let _opRefreshTimer = null;

  function onStatusChange(fn) { _onStatusChange = fn; }

  function _notificar() {
    if (_onStatusChange) _onStatusChange({
      online:   navigator.onLine,
      pending:  getQueue().length,
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
  async function precargar() {
    if (!navigator.onLine) return;
    try {
      const [maestros, stock, config] = await Promise.all([
        fetch(_gasUrl('descargarMaestros')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getStock')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getConfig')).then(r => r.json()).catch(() => null)
      ]);

      console.log('[Offline] descargarMaestros respuesta:', maestros);
      if (maestros?.ok) {
        // GAS nuevo — recibe las 4 tablas de MOS
        const d = maestros.data;
        console.log('[Offline] personal recibido:', d?.personal?.length, 'registros');
        if (d.personal      != null) guardar(KEYS.PERSONAL,      d.personal);
        if (d.productos     != null) guardar(KEYS.PRODUCTOS,     d.productos);
        if (d.equivalencias != null) guardar(KEYS.EQUIVALENCIAS, d.equivalencias);
        if (d.proveedores   != null) guardar(KEYS.PROVEEDORES,   d.proveedores);
        if (d.impresoras    != null) guardar(KEYS.IMPRESORAS,    d.impresoras);
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
        if (productos?.ok)   guardar(KEYS.PRODUCTOS,   productos.data);
        if (proveedores?.ok) guardar(KEYS.PROVEEDORES, proveedores.data);
      }

      if (stock?.ok)  guardar(KEYS.STOCK,  stock.data);
      if (config?.ok) guardar(KEYS.CONFIG, config.data);

      localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
      _notificar();
    } catch(e) {
      console.warn('[Offline] Error en precarga:', e);
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
    const localId = 'L' + Date.now() + Math.random().toString(36).substr(2, 5);
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
    const queue = getQueue().filter(i => i.status === 'pending');
    if (!queue.length) return;

    _syncing = true;
    _notificar();

    for (const item of queue) {
      try {
        const res = await fetch(window.WH_CONFIG.gasUrl, {
          method:  'POST',
          headers: { 'Content-Type': 'text/plain' },
          body:    JSON.stringify(item.params)
        }).then(r => r.json());

        _actualizarItemQueue(item.localId, res.ok ? 'synced' : 'error');
      } catch {
        _actualizarItemQueue(item.localId, 'error');
      }
    }

    limpiarSincronizados();
    _syncing = false;
    localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
    _notificar();
  }

  // ── Precarga operacional (guías, preingresos, stock, ajustes, auditorías) ──
  async function precargarOperacional() {
    if (!navigator.onLine) return;
    try {
      const r = await fetch(_gasUrl('descargarOperacional')).then(r => r.json()).catch(() => null);
      if (!r?.ok) return;
      const d = r.data;
      const changed = [];

      function _hayDiff(newArr, key) {
        if (!newArr?.length) return false;
        const old = cargar(key);
        return !old || old.length !== newArr.length;
      }

      if (d.guias       != null) { if (_hayDiff(d.guias,       KEYS.GUIAS))        { guardar(KEYS.GUIAS,        d.guias);       changed.push('guias'); }       else guardar(KEYS.GUIAS, d.guias); }
      if (d.detalles    != null) { if (_hayDiff(d.detalles,    KEYS.GUIA_DETALLE)) { guardar(KEYS.GUIA_DETALLE, d.detalles);    changed.push('detalles'); }   else guardar(KEYS.GUIA_DETALLE, d.detalles); }
      if (d.preingresos != null) { if (_hayDiff(d.preingresos, KEYS.PREINGRESOS))  { guardar(KEYS.PREINGRESOS,  d.preingresos); changed.push('preingresos'); } else guardar(KEYS.PREINGRESOS, d.preingresos); }
      if (d.stock       != null) { if (_hayDiff(d.stock,       KEYS.STOCK))        { guardar(KEYS.STOCK,        d.stock);       changed.push('stock'); }       else guardar(KEYS.STOCK, d.stock); }
      if (d.ajustes     != null) { if (_hayDiff(d.ajustes,     KEYS.AJUSTES))      { guardar(KEYS.AJUSTES,      d.ajustes);     changed.push('ajustes'); }     else guardar(KEYS.AJUSTES, d.ajustes); }
      if (d.auditorias  != null) { if (_hayDiff(d.auditorias,  KEYS.AUDITORIAS_C)) { guardar(KEYS.AUDITORIAS_C, d.auditorias);  changed.push('auditorias'); }  else guardar(KEYS.AUDITORIAS_C, d.auditorias); }

      window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed } }));
    } catch(e) { console.warn('[Offline] Error en precarga operacional:', e); }
  }

  // ── Precarga maestros con throttle 5 min ─────────────────
  async function _precargarMaestrosThrottled() {
    const last = parseInt(localStorage.getItem(KEYS.LAST_MASTER) || '0', 10);
    if (Date.now() - last < 5 * 60 * 1000) return; // menos de 5 min → skip
    await precargar().catch(() => {});
    localStorage.setItem(KEYS.LAST_MASTER, String(Date.now()));
  }

  // Inicia el refresh automático cada 60s (llamar desde App.init, antes del login)
  function iniciarRefreshOperacional() {
    if (_opRefreshTimer) return;
    // Carga inmediata: maestros (si no hay caché) + operacional
    if (!cargar(KEYS.PERSONAL)?.length) {
      precargar().catch(() => {});
    } else {
      _precargarMaestrosThrottled();
    }
    precargarOperacional();
    _opRefreshTimer = setInterval(() => {
      _precargarMaestrosThrottled();
      precargarOperacional();
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
  const getConfigCache        = () => cargar(KEYS.CONFIG)        || {};
  const getGuiasCache         = () => cargar(KEYS.GUIAS)         || [];
  const getGuiaDetalleCache   = () => cargar(KEYS.GUIA_DETALLE)  || [];
  const getPreingresosCache   = () => cargar(KEYS.PREINGRESOS)   || [];
  const getAjustesCache       = () => cargar(KEYS.AJUSTES)       || [];
  const getAuditoriasCache    = () => cargar(KEYS.AUDITORIAS_C)  || [];
  const getAdminPin           = () => localStorage.getItem(KEYS.ADMIN_PIN) || null;

  return {
    precargar, sincronizar, encolar, getQueue,
    validarPinLocal, onStatusChange,
    _guardarPersonalConPin,
    getPersonalCache, getProductosCache, getEquivalenciasCache,
    getStockCache, getProveedoresCache,
    getImpresorasCache, getConfigCache,
    getGuiasCache, getGuiaDetalleCache, getPreingresosCache,
    getAjustesCache, getAuditoriasCache,
    getAdminPin,
    precargarOperacional, iniciarRefreshOperacional, detenerRefreshOperacional,
    estaOnline: () => navigator.onLine
  };
})();
