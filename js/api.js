// warehouseMos — api.js  Comunicación con GAS + soporte offline
'use strict';

const API = (() => {
  function _gasUrl() { return window.WH_CONFIG?.gasUrl || ''; }

  // GET: network first, cache como fallback
  async function call(params) {
    const GAS_URL = _gasUrl();
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
    'registrarEnvasado', 'registrarMerma', 'resolverMerma',
    'registrarProductoNuevo', 'aprobarProductoNuevo',
    'crearPreingreso', 'aprobarPreingreso', 'actualizarPreingreso',
    'crearProducto', 'actualizarProducto',
    'crearProveedor', 'actualizarProveedor',
    'crearAjuste', 'auditarProducto',
    'asignarAuditoria', 'ejecutarAuditoria',
    'marcarAlertaRevisada', 'aceptarTeoricoAlerta',
    'actualizarGuia', 'actualizarPickup', 'cerrarTurno'
  ]);
  function _genLocalId() {
    return 'L' + Date.now() + Math.random().toString(36).substr(2, 7);
  }

  async function _doFetchWithRetry(GAS_URL, params) {
    // 3 intentos: rápido si todo OK, recuperación si red flaquea o lock saturado.
    // Backoff: 600ms, 1500ms (total ~2.1s antes de rendirse).
    // params.localId se mantiene constante en los reintentos → GAS deduplica.
    const delays = [600, 1500];
    for (let intento = 0; intento < 3; intento++) {
      try {
        const res  = await fetch(GAS_URL, {
          method:  'POST',
          headers: { 'Content-Type': 'text/plain' },
          body:    JSON.stringify(params)
        });
        const json = await res.json();
        if (json && json.ok === false && /saturado|ocupado/i.test(json.error || '')) {
          if (intento < 2) { await new Promise(r => setTimeout(r, delays[intento])); continue; }
        }
        return json;
      } catch {
        if (intento < 2) { await new Promise(r => setTimeout(r, delays[intento])); continue; }
        // Red falló tras 3 intentos → encolar offline (reusa localId si ya existe)
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

    if (!GAS_URL || !navigator.onLine) {
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
    descargarOperacional: ()     => call({ action: 'descargarOperacional' }),

    // Dashboard
    getDashboard:       ()       => call({ action: 'getDashboard' }),

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
    getPickups:         (p={})  => call({ action: 'getPickups', ...p }),
    actualizarPickup:   (p)     => post({ action: 'actualizarPickup', ...p }),
    agregarDetalle:           (p) => post({ action: 'agregarDetalleGuia',        ...p }),
    actualizarFechaVencimiento: (p) => post({ action: 'actualizarFechaVencimiento',  ...p }),
    actualizarCantidadDetalle:  (p) => post({ action: 'actualizarCantidadDetalle',   ...p }),
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
    auditarProducto:      (p)   => post({ action: 'auditarProducto', ...p })
  };
})();
