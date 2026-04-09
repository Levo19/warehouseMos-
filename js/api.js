// warehouseMos — api.js  Comunicación con GAS
const API = (() => {
  const GAS_URL = window.WH_CONFIG?.gasUrl || '';

  async function call(params) {
    if (!GAS_URL) {
      console.warn('[API] GAS_URL no configurada. Modo offline.');
      return { ok: false, error: 'Sin conexión al servidor' };
    }
    try {
      const url = new URL(GAS_URL);
      Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
      const res  = await fetch(url.toString(), { mode: 'cors' });
      const data = await res.json();
      return data;
    } catch (err) {
      console.error('[API GET]', err);
      return { ok: false, error: err.message };
    }
  }

  async function post(params) {
    if (!GAS_URL) return { ok: false, error: 'Sin conexión al servidor' };
    try {
      const res  = await fetch(GAS_URL, {
        method:  'POST',
        mode:    'cors',
        headers: { 'Content-Type': 'text/plain' },
        body:    JSON.stringify(params)
      });
      const data = await res.json();
      return data;
    } catch (err) {
      console.error('[API POST]', err);
      return { ok: false, error: err.message };
    }
  }

  return {
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
    agregarDetalle:     (p)      => post({ action: 'agregarDetalleGuia', ...p }),
    cerrarGuia:         (id, u)  => post({ action: 'cerrarGuia', idGuia: id, usuario: u }),

    // Preingresos
    getPreingresos:     (p={})   => call({ action: 'getPreingresos', ...p }),
    crearPreingreso:    (p)      => post({ action: 'crearPreingreso', ...p }),
    aprobarPreingreso:  (p)      => post({ action: 'aprobarPreingreso', ...p }),

    // Envasados
    getEnvasados:       (p={})   => call({ action: 'getEnvasados', ...p }),
    getPendientes:      ()       => call({ action: 'getPendientesEnvasado' }),
    registrarEnvasado:  (p)      => post({ action: 'registrarEnvasado', ...p }),

    // Mermas
    getMermas:          (p={})   => call({ action: 'getMermas', ...p }),
    registrarMerma:     (p)      => post({ action: 'registrarMerma', ...p }),

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
    getProductosNuevos: (p={})   => call({ action: 'getProductosNuevos', ...p }),
    registrarPN:        (p)      => post({ action: 'registrarProductoNuevo', ...p }),
    aprobarPN:          (p)      => post({ action: 'aprobarProductoNuevo', ...p }),

    // Config
    getConfig:          ()       => call({ action: 'getConfig' }),
    setConfig:          (k, v)   => post({ action: 'setConfig', clave: k, valor: v }),

    // Etiquetas
    imprimirEtiqueta:   (p)      => post({ action: 'imprimirEtiqueta', ...p })
  };
})();
