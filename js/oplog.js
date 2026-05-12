// ============================================================
// warehouseMos — js/oplog.js
// Operation Log cliente: cola persistente IndexedDB + retry
// exponencial + reconciliación al recargar.
//
// API pública:
//   OpLog.enqueue(op) — agrega op a la cola, dispara intento
//   OpLog.flush()     — fuerza intento inmediato
//   OpLog.subscribe(fn)— recibe eventos {type, op, result}
//   OpLog.pendingFor(idGuia)— ops aún no APPLIED para una guía
//   OpLog.deviceId()  — id estable de este device
//
// Cada op:
//   { idOp, tipo, payload, ts, estado:'pending|saving|saved|failed' }
//
// Backend: GAS endpoint 'aplicarOp' idempotente por idOp.
// ============================================================

(function() {
  'use strict';
  const DB_NAME = 'wh_oplog';
  const DB_VER  = 1;
  const STORE   = 'ops';
  const DEVICE_KEY = 'wh_device_id';

  let db = null;
  let subscribers = [];
  let inFlight = new Set();
  let backoffMs = 1000;
  const BACKOFF_MAX = 27000;

  function _emit(type, op, result) {
    subscribers.forEach(fn => { try { fn({ type, op, result }); } catch(_){} });
  }

  function _openDB() {
    if (db) return Promise.resolve(db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VER);
      req.onupgradeneeded = () => {
        const d = req.result;
        if (!d.objectStoreNames.contains(STORE)) {
          const s = d.createObjectStore(STORE, { keyPath: 'idOp' });
          s.createIndex('idGuia', 'idGuia', { unique: false });
          s.createIndex('estado', 'estado', { unique: false });
        }
      };
      req.onsuccess = () => { db = req.result; resolve(db); };
      req.onerror   = () => reject(req.error);
    });
  }

  function _put(op) {
    return _openDB().then(d => new Promise((res, rej) => {
      const tx = d.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(op);
      tx.oncomplete = () => res(op);
      tx.onerror    = () => rej(tx.error);
    }));
  }

  function _delete(idOp) {
    return _openDB().then(d => new Promise((res) => {
      const tx = d.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(idOp);
      tx.oncomplete = () => res();
      tx.onerror    = () => res();
    }));
  }

  function _all() {
    return _openDB().then(d => new Promise((res) => {
      const tx = d.transaction(STORE, 'readonly');
      const out = [];
      tx.objectStore(STORE).openCursor().onsuccess = e => {
        const c = e.target.result;
        if (c) { out.push(c.value); c.continue(); }
        else res(out);
      };
      tx.onerror = () => res([]);
    }));
  }

  function _getDeviceId() {
    let id = localStorage.getItem(DEVICE_KEY);
    if (!id) {
      id = 'dev-' + Math.random().toString(36).slice(2, 10) + '-' + Date.now().toString(36);
      localStorage.setItem(DEVICE_KEY, id);
    }
    return id;
  }

  function _genIdOp() {
    return 'OP-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  }

  async function enqueue(opPartial) {
    const op = Object.assign({
      idOp:    opPartial.idOp || _genIdOp(),
      estado:  'pending',
      ts:      new Date().toISOString(),
      retries: 0
    }, opPartial);
    await _put(op);
    _emit('enqueued', op);
    flush();
    return op;
  }

  let flushing = false;
  async function flush() {
    if (flushing) return;
    if (!navigator.onLine) return;
    flushing = true;
    try {
      const all = await _all();
      const candidates = all.filter(o => o.estado === 'pending' || o.estado === 'failed');
      candidates.sort((a, b) => new Date(a.ts) - new Date(b.ts));
      for (const op of candidates) {
        if (inFlight.has(op.idOp)) continue;
        await _attempt(op);
      }
    } finally {
      flushing = false;
    }
  }

  async function _attempt(op) {
    inFlight.add(op.idOp);
    op.estado = 'saving';
    await _put(op);
    _emit('saving', op);

    try {
      const result = await _callServer(op);
      if (result && result.ok) {
        op.estado = 'saved';
        op.result = result.data || result;
        await _put(op);
        _emit('saved', op, result);
        backoffMs = 1000;
        setTimeout(() => _delete(op.idOp), 5000);
      } else {
        op.estado  = 'failed';
        op.error   = (result && result.error) || 'unknown';
        op.retries = (op.retries || 0) + 1;
        await _put(op);
        _emit('failed', op, result);
        _scheduleRetry();
      }
    } catch(e) {
      op.estado  = 'failed';
      op.error   = e.message;
      op.retries = (op.retries || 0) + 1;
      await _put(op);
      _emit('failed', op, { ok: false, error: e.message });
      _scheduleRetry();
    } finally {
      inFlight.delete(op.idOp);
    }
  }

  function _scheduleRetry() {
    setTimeout(() => { flush(); }, backoffMs);
    backoffMs = Math.min(backoffMs * 3, BACKOFF_MAX);
  }

  async function _callServer(op) {
    if (typeof API === 'undefined') throw new Error('API global no cargada');
    return API.post('aplicarOp', {
      idOp:     op.idOp,
      idGuia:   op.idGuia || (op.payload && op.payload.idGuia) || '',
      tipo:     op.tipo,
      payload:  JSON.stringify(op.payload || {}),
      deviceId: _getDeviceId(),
      usuario:  (window.App && App.getUsuario && App.getUsuario()) || ''
    });
  }

  async function pendingFor(idGuia) {
    const all = await _all();
    return all.filter(o => (o.idGuia === idGuia) && (o.estado !== 'saved'));
  }

  function subscribe(fn) {
    subscribers.push(fn);
    return () => { subscribers = subscribers.filter(s => s !== fn); };
  }

  window.addEventListener('online',  () => { backoffMs = 1000; flush(); });
  window.addEventListener('focus',   () => { flush(); });
  setInterval(() => flush(), 30000);

  window.OpLog = {
    enqueue, flush, subscribe, pendingFor,
    deviceId: _getDeviceId,
    genIdOp:  _genIdOp,
    _all:     _all
  };
})();
