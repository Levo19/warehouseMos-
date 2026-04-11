// warehouseMos — offline.js
// Caché local, cola offline y sincronización sin duplicados
'use strict';

const OfflineManager = (() => {
  const KEYS = {
    PERSONAL:    'wh_personal',
    PRODUCTOS:   'wh_productos',
    STOCK:       'wh_stock',
    PROVEEDORES: 'wh_proveedores',
    CONFIG:      'wh_config',
    QUEUE:       'wh_queue',
    LAST_SYNC:   'wh_last_sync'
  };

  // ── Estado ────────────────────────────────────────────────
  let _syncing = false;
  let _onStatusChange = null;

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
  async function precargar() {
    if (!navigator.onLine) return;
    try {
      // Personal CON pins (solo para validación local)
      const [personal, productos, stock, proveedores, config] = await Promise.all([
        fetch(_gasUrl('getPersonalConPin')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getProductos&estado=1')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getStock')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getProveedores&estado=1')).then(r => r.json()).catch(() => null),
        fetch(_gasUrl('getConfig')).then(r => r.json()).catch(() => null)
      ]);

      if (personal?.ok)    guardar(KEYS.PERSONAL,    personal.data);
      if (productos?.ok)   guardar(KEYS.PRODUCTOS,   productos.data);
      if (stock?.ok)       guardar(KEYS.STOCK,        stock.data);
      if (proveedores?.ok) guardar(KEYS.PROVEEDORES,  proveedores.data);
      if (config?.ok)      guardar(KEYS.CONFIG,       config.data);

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

  // ── Getters cache ─────────────────────────────────────────
  function _guardarPersonalConPin(data) { guardar(KEYS.PERSONAL, data); }

  const getPersonalCache    = () => cargar(KEYS.PERSONAL)    || [];
  const getProductosCache   = () => cargar(KEYS.PRODUCTOS)   || [];
  const getStockCache       = () => cargar(KEYS.STOCK)       || [];
  const getProveedoresCache = () => cargar(KEYS.PROVEEDORES) || [];
  const getConfigCache      = () => cargar(KEYS.CONFIG)      || {};

  return {
    precargar, sincronizar, encolar, getQueue,
    validarPinLocal, onStatusChange,
    _guardarPersonalConPin,
    getPersonalCache, getProductosCache,
    getStockCache, getProveedoresCache, getConfigCache,
    estaOnline: () => navigator.onLine
  };
})();
