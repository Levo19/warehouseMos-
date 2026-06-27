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
    ENVASADOS:    'wh_envasados',
    CAT_VERSION:  'wh_cat_version'        // [poller] baseline de mos.catalogo_version visto por este equipo
  };

  // ── Estado ────────────────────────────────────────────────
  let _syncing       = false;
  let _opLoading     = false;  // guard: evitar llamadas concurrent a precargarOperacional
  let _opInflight    = null;   // [perf v2.13.242] promesa operacional en vuelo → callers concurrentes la REUSAN
  let _lastOpTs      = 0;      // timestamp de la última llamada a precargarOperacional
  const OP_MIN_MS    = 15000;  // mínimo 15s entre llamadas operacionales
  let _masterInflight = null;  // [perf] promesa en vuelo → callers concurrentes la REUSAN (no dispara N fetch)
  let _lastMasterTs  = 0;      // timestamp de la última llamada (que SÍ ejecutó) a precargar
  const MASTER_MIN_MS = 60000; // maestros: mínimo 60s entre refresh background (cambian poco)
  // [v2.13.236 perf] Piso DURO que incluso forzar=true respeta. Mata la "ráfaga" de
  // descargarMaestros: login + welcome + init + reconexión disparaban precargar(true)
  // espalda-con-espalda y cada uno bajaba el catálogo completo (lento + parseo que
  // traba la nav). Solo el sync MANUAL del usuario (forzar==='manual') lo ignora.
  const MASTER_HARD_MS = 8000;
  // Firmas para detectar si productos/equivalencias REALMENTE cambiaron entre refreshes.
  // Sin esto, cada precarga marcaba 'productos' como cambiado aunque el catálogo fuera
  // idéntico → wh:data-refresh → ProductosView.silentRefresh → flash/parpadeo inútil.
  const _masterSig = {};
  let _onStatusChange = null;
  let _opRefreshTimer = null;

  // ── Poller de versión del catálogo ───────────────────────────
  // El maestro (productos/equivalencias) cambia en MOS; cuando eso pasa, un trigger
  // incrementa mos.catalogo_version. Este equipo guarda la versión vista (baseline) y la
  // sondea barato cada ~50s (+ al volver a foreground/foco). Si subió → re-descarga el
  // catálogo (precargar 'manual') y avanza el baseline. La RE-DESCARGA es de DATOS de
  // referencia: pasa por _guardarSiCambia (no flashea si no cambió) y dispara
  // wh:data-refresh → silentRefresh, que NO resetea guías/envasados/ventas en armado.
  let _catVersionBaseline = null;   // null = aún sin baseline (no comparar todavía)
  let _catPollTimer       = null;
  let _catPollBusy        = false;  // guard anti-reentrada del propio chequeo
  const CAT_POLL_MS       = 50000;  // ~50s; solo corre con la pestaña visible
  // [perf 500x] coalescing de re-descargas: muchos bumps de versión en ráfaga (ej. ediciones en lote en MOS)
  // NO deben disparar una re-descarga de ~1.9MB cada uno. Se difiere y coalesce a 1 sola descarga por ventana.
  let _catRedownloadTimer = null;
  let _catPendingVersion  = null;
  let _catLastCheck       = 0;      // throttle del chequeo de versión (foco/visibility no deben spamear)
  const CAT_REDOWNLOAD_DEBOUNCE_MS = 20000;  // ventana de quietud antes de re-descargar (coalesce)
  const CAT_CHECK_THROTTLE_MS      = 30000;  // no chequear versión más de 1 vez cada 30s (foco/visibility/timer)

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

  // [v2.13.74] Auto-cleanup AGRESIVO al cambio de versión. Al actualizar la
  // app, todos los caches grandes (productos/stock/etc) se regeneran del
  // backend en la próxima precarga. NO tiene sentido conservarlos viejos —
  // mejor borrar todos los wh_* excepto los esenciales del usuario:
  //   - wh_sesion (su login)
  //   - wh_device_id (UUID del equipo)
  //   - wh_app_version (control de versión)
  //   - wh_perms_done_v* (wizard de permisos completado)
  //   - wh_audio_ok (legacy)
  // Esto resuelve QuotaExceededError definitivamente — el localStorage
  // queda casi vacío después de cada actualización.
  (async function _autoCleanup() {
    try {
      const verActual = await fetch('./version.json?t=' + Date.now())
        .then(r => r.json()).then(j => j.version).catch(() => null);
      if (!verActual) return;
      const verAnterior = localStorage.getItem('wh_app_version');
      if (verAnterior && verAnterior !== verActual) {
        console.log('[Offline] cambio versión ' + verAnterior + ' → ' + verActual + ' · cleanup agresivo caches');
        // [v2.13.103] AUDITORÍA SENIOR — bug pre-existente: TIER0 NO estaba en
        // PRESERVAR. Cuando se actualizaba la app, wh_personal/wh_admin_cache/
        // wh_queue se borraban → mismo bug v2.13.99 (login no funciona offline)
        // pero disparado por el upgrade en vez de quota cleanup.
        // Ahora se preservan TODOS los TIER0 (críticos para que la PWA funcione).
        const PRESERVAR = /^(wh_sesion|wh_device_id|wh_app_version|wh_audio_ok|wh_perms_done_v.*|wh_personal|wh_admin_cache|wh_queue|wh_gas_url)$/;
        let borrados = 0;
        Object.keys(localStorage).forEach(k => {
          if (!k.startsWith('wh_')) return;
          if (PRESERVAR.test(k)) return;
          localStorage.removeItem(k);
          borrados++;
        });
        console.log('[Offline cleanup] ' + borrados + ' caches borrados · se regeneran en próxima precarga');
      }
      localStorage.setItem('wh_app_version', verActual);
    } catch(_){}
  })();

  // ── Cache helpers ─────────────────────────────────────────
  // [v2.13.102] Compresión LZ-String para entradas grandes (>50KB) +
  // cleanup smart por antigüedad (anti round-robin destructivo).
  //
  // CAUSA RAÍZ DEL BUG ORIGINAL:
  //   Cada precarga guarda 8+ caches secuenciales. Cuando storage está cerca
  //   al límite (~5MB), cada guardar() entra al cleanup que borra TODO TIER1
  //   excepto el key actual. El SIGUIENTE guardar() encuentra wh_productos
  //   recién guardado y lo borra para hacer espacio. Cascada destructiva:
  //   solo el último cache sobrevive cada ciclo.
  //
  // FIX A — Compresión: wh_productos plain ~600KB → comprimido ~150KB (-75%).
  //   Después del primer guardado todo cabe holgadamente y el cleanup no se
  //   dispara más. Prefijo 'LZ:' identifica entradas comprimidas (back-compat
  //   con datos viejos plain).
  //
  // FIX B — Cleanup por edad: en vez de barrer TODO TIER1 a la vez, borra
  //   el más viejo (timestamp ascendente), reintenta. Caches recién guardados
  //   sobreviven.
  const _COMPRESS_PREFIX    = 'LZ:';
  const _COMPRESS_THRESHOLD = 50_000;  // solo comprimir JSON >50KB

  function _serializar(payload) {
    const json = JSON.stringify(payload);
    if (json.length < _COMPRESS_THRESHOLD || typeof LZString === 'undefined') {
      return json;
    }
    try {
      return _COMPRESS_PREFIX + LZString.compressToUTF16(json);
    } catch(_) { return json; }
  }

  function _deserializar(raw) {
    if (!raw) return null;
    try {
      if (raw.startsWith(_COMPRESS_PREFIX)) {
        if (typeof LZString === 'undefined') return null;  // lib no cargada todavía
        return JSON.parse(LZString.decompressFromUTF16(raw.slice(_COMPRESS_PREFIX.length)));
      }
      return JSON.parse(raw);
    } catch(_) { return null; }
  }

  // Lee el timestamp de una entrada sin descomprimir el data completo cuando
  // se puede (heurística: el ts queda al final del JSON envuelto). Si no
  // podemos extraerlo barato, descomprimimos.
  function _leerTs(key) {
    const raw = localStorage.getItem(key);
    if (!raw) return 0;
    // Datos plain: regex rápido al final
    if (!raw.startsWith(_COMPRESS_PREFIX)) {
      const m = raw.match(/"ts":(\d+)\}\s*$/);
      if (m) return parseInt(m[1], 10);
    }
    // Comprimido o regex fail: descomprimir
    const obj = _deserializar(raw);
    return obj?.ts || 0;
  }

  function guardar(key, data) {
    // [v2.13.110] Invalidar cache de parseo ANTES de escribir — garantiza
    // que cualquier read concurrente que ocurra entre la invalidación y
    // el setItem caiga al miss y refresque.
    _invalidarParseCache(key);
    try { localStorage.setItem(key, _serializar({ data, ts: Date.now() })); }
    catch(e) {
      // [v2.13.98 BUG CRITICO FIX] Cleanup en 2 fases preserva trabajo en progreso.
      //
      // BUG ORIGINAL (v2.13.68 - v2.13.97):
      //   wh_guia_detalle y wh_preingresos estaban en GRANDES, se borraban
      //   junto a los caches puros. Si el usuario modificaba un detalle de
      //   guía cerrada (write optimista al cache local + fire-and-forget
      //   al backend), el siguiente cleanup borraba wh_guia_detalle ANTES
      //   de que el backend confirmara. Si el backend timeoutaba, los
      //   cambios se perdían silenciosamente. Además el round-robin entre
      //   wh_productos ↔ wh_stock hacía que el catálogo se viera
      //   constantemente "Sin productos".
      //
      // FIX:
      //   TIER 1 (cache puro): se puede borrar libremente, se redescarga del backend
      //   TIER 2 (trabajo en progreso): SOLO borrar como último recurso
      //   Borra TIER 1 primero, reintenta. Solo si sigue fallando, TIER 2.
      if (e.name === 'QuotaExceededError' || /quota/i.test(e.message)) {
        console.warn('[Offline] storage lleno · cleanup emergency para guardar', key);
        // [v2.13.100 BUG FIX] wh_personal estaba en TIER1 → cleanup lo borraba →
        // validarPinLocal devolvía null → "no me deja entrar con mi contraseña".
        //
        // Mismo problema con wh_admin_cache (validar clave admin global).
        //
        // Reorganización en 3 tiers:
        //
        // TIER 0 — INTOCABLE (críticos para que la PWA funcione):
        //   wh_personal       → validarPinLocal (login)
        //   wh_admin_cache    → clave admin global + tiers
        //   wh_queue          → operaciones offline sin sincronizar
        //   wh_sesion         → sesión actual del usuario logueado
        //   wh_device_id      → identidad de la tablet/PC
        //   wh_app_version    → tracking de updates
        //   wh_gas_url        → URL del backend (sin esto no hay nada)
        //
        // TIER 1 — CACHE PURO (libre de borrar, redescarga en próxima sync):
        //   wh_productos, wh_stock, wh_proveedores, wh_ajustes,
        //   wh_auditorias_c, wh_ubicaciones, wh_equivalencias, wh_zonas,
        //   wh_impresoras, wh_pn, wh_config, wh_guias
        //
        // TIER 2 — TRABAJO EN PROGRESO (último recurso):
        //   wh_guia_detalle → addDetalleCache (mods optimistas detalles)
        //   wh_preingresos  → inyectarPreingreso / patchPreingresosCache
        //   wh_envasados    → inyectarEnvasadoCache
        const TIER1 = [
          'wh_productos','wh_stock','wh_proveedores','wh_ajustes',
          'wh_auditorias_c','wh_ubicaciones','wh_equivalencias','wh_zonas',
          'wh_impresoras','wh_pn','wh_config','wh_guias'
        ];
        const TIER2 = ['wh_guia_detalle','wh_preingresos','wh_envasados'];

        // [v2.13.102] Cleanup por antigüedad — anti round-robin destructivo.
        //
        // Antes: borrábamos TODO TIER1 de golpe. Si el next guardar() también
        // fallaba, borraba el recién guardado. Cascada: solo el último cache
        // sobrevivía.
        //
        // Ahora: borramos UNO por UNO, el más viejo primero, reintentando
        // después de cada borrado. Los caches recién guardados (más recientes)
        // sobreviven porque el cleanup ataca primero los viejos.
        //
        // [v2.13.103] PERF FIX (auditoría senior): serializamos UNA vez fuera
        // del loop. Antes cada iteración re-corría JSON.stringify + LZString
        // compress (~50-100ms × N iteraciones = pause perceptible en wh_productos).
        const _payloadFinal = _serializar({ data, ts: Date.now() });
        const _intentarGuardar = () => {
          try {
            localStorage.setItem(key, _payloadFinal);
            return true;
          } catch(_) { return false; }
        };

        const _ordenarPorEdad = (tier) => {
          return tier
            .filter(k => k !== key && localStorage.getItem(k))
            .map(k => ({ k, ts: _leerTs(k) }))
            .sort((a, b) => a.ts - b.ts);  // más viejo primero
        };

        // Fase 1: TIER1 por edad
        let liberadosT1 = 0;
        const candidatosT1 = _ordenarPorEdad(TIER1);
        // [v2.13.111 AUDIT FIX A] Invalidar también el cache de parseo
        // cuando borramos del localStorage en el cleanup. Sin esto las
        // próximas cargar(k) devolvían dato viejo hasta TTL 15s, aunque
        // localStorage ya estuviera vacío para esa key.
        for (const { k } of candidatosT1) {
          try {
            localStorage.removeItem(k);
            if (typeof _invalidarParseCache === 'function') _invalidarParseCache(k);
            liberadosT1++;
          } catch(_){}
          if (_intentarGuardar()) {
            console.log('[Offline] ✓ guardado tras liberar ' + liberadosT1 + ' TIER1 viejos:', key);
            return;
          }
        }
        // Fase 2: TIER2 por edad (último recurso)
        console.warn('[Offline] TIER1 insuficiente · borrando TIER2 (trabajo del usuario en riesgo)');
        let liberadosT2 = 0;
        const candidatosT2 = _ordenarPorEdad(TIER2);
        for (const { k } of candidatosT2) {
          try {
            localStorage.removeItem(k);
            if (typeof _invalidarParseCache === 'function') _invalidarParseCache(k);
            liberadosT2++;
          } catch(_){}
          if (_intentarGuardar()) {
            console.log('[Offline] ✓ guardado tras liberar ' + liberadosT2 + ' TIER2:', key);
            return;
          }
        }
        // Nada funcionó
        console.warn('[Offline] localStorage IMPOSIBLE · omitido:', key);
        if (typeof window !== 'undefined' && !window._whStorageWarned) {
          window._whStorageWarned = true;
          setTimeout(() => {
            if (typeof toast === 'function') toast('⚠ Almacenamiento del navegador lleno. Cerrá y reabrí la PWA para limpiar.', 'error', 10000);
          }, 500);
        }
      } else { console.warn('[Offline] localStorage error:', e); }
    }
  }

  // [v2.13.110] CACHE de parseo. Antes cada cargar() descomprimía LZ-String
  // + JSON.parse cada vez → para wh_productos (~150KB comprimido), eso es
  // 50-150ms por llamada. Si una vista llama getProductosCache() + getStockCache()
  // + getEquivalenciasCache() varias veces en un cambio de módulo, se acumulan
  // a >1s + lag visible.
  //
  // Estrategia: cachear el resultado parseado con su timestamp. La entrada se
  // invalida cuando guardar() escribe esa key (writes son monotónicos).
  // TTL de 15s como seguro adicional contra staleness en edge cases.
  const _parseCache = new Map();   // key → { data, ts }
  const _PARSE_TTL_MS = 15000;

  function cargar(key) {
    // 1. Hit en cache de parseo
    const cached = _parseCache.get(key);
    if (cached && (Date.now() - cached.ts) < _PARSE_TTL_MS) {
      return cached.data;
    }
    // 2. Miss → descomprimir + parsear
    const raw = localStorage.getItem(key);
    const obj = _deserializar(raw);
    const data = obj ? obj.data : null;
    if (data !== null) _parseCache.set(key, { data, ts: Date.now() });
    return data;
  }

  // Invalidar cache de parseo cuando guardar() escribe. Esto es el contrato
  // que garantiza que cargar() después de guardar() vea el dato nuevo.
  function _invalidarParseCache(key) {
    _parseCache.delete(key);
  }

  // [v2.13.111 AUDIT FIX B] Cross-tab invalidation.
  // Si otra pestaña/ventana de la app escribe localStorage, el `storage` event
  // dispara en TODAS las demás (no en la que escribió). Invalidamos su cache
  // para que la próxima cargar() vea el dato nuevo escrito por la otra tab.
  if (typeof window !== 'undefined') {
    window.addEventListener('storage', (e) => {
      if (e.key && e.key.startsWith('wh_')) _invalidarParseCache(e.key);
    });
  }

  // ── Precarga de datos de referencia ──────────────────────
  // Intenta descargarMaestros (endpoint nuevo que trae las 4 tablas de MOS).
  // Si el GAS desplegado no lo conoce aún, degrada al endpoint legacy
  // getPersonalConPin para que el login siempre funcione.
  // Firma barata y estable de un dataset (longitud + hash rodante del JSON). Evita
  // guardar JSON.stringify gigante en memoria; suficiente para detectar "no cambió".
  function _firma(arr) {
    if (!arr) return 'null';
    const s = JSON.stringify(arr);
    let h = 5381;
    for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return arr.length + ':' + (h >>> 0);
  }
  // Marca cambio en `changed` SOLO si la firma del dataset difiere de la última vista.
  function _guardarSiCambia(key, sigKey, arr, label, changed) {
    if (arr == null) return;
    const sig = _firma(arr);
    // Idéntico a lo último visto Y ya hay algo persistido → no reescribir ni avisar a
    // la UI (evita el flash + el costo de recomprimir un array grande en cada refresh).
    if (_masterSig[sigKey] === sig && (cargar(key) || []).length) return;
    _masterSig[sigKey] = sig;
    guardar(key, arr);
    if (label) changed.push(label);
  }

  // forzar: false = refresh background (throttle 60s) · true = forzado (respeta piso
  // duro de 8s, mata ráfagas) · 'manual' = sync explícito del usuario (ignora todo piso).
  async function precargar(forzar = false) {
    if (!navigator.onLine) return;
    // [perf] Dedup in-flight: si ya hay una precarga corriendo, los callers
    // concurrentes (login + welcome + init + reconexión) REUSAN esa promesa en
    // vez de disparar otro descargarMaestros. Esto solo es la causa raíz de la ráfaga.
    if (_masterInflight) return _masterInflight;
    const ahora = Date.now();
    if (forzar === 'manual') {
      /* sync manual del usuario: sin piso */
    } else if (forzar) {
      if (ahora - _lastMasterTs < MASTER_HARD_MS) return; // piso duro: aún forzado, no en ráfaga
    } else if (ahora - _lastMasterTs < MASTER_MIN_MS) {
      return; // refresh background normal: throttle generoso (maestros cambian poco)
    }
    _lastMasterTs  = ahora;
    _masterInflight = (async () => {
    try {
      // [FIX asimetría catálogo] Maestros vía API.descargarMaestros → call() usa Supabase directo
      // (mos.catalogo_wh_rls) cuando _whLecturaDirecta está ON, con fallback a GAS. ANTES era fetch crudo
      // a GAS = leía la HOJA; un proveedor/estación escrito DIRECTO a Supabase (que NO toca la Hoja) era
      // INVISIBLE para WH (el dato estaba en mos.* pero WH cacheaba la Hoja). Mismo arreglo que ya se hizo
      // para operacional (API.descargarOperacional). getStock/getConfig siguen por GAS (fuera de scope).
      const [maestros, stock, config] = await Promise.all([
        ((typeof API !== 'undefined' && API.descargarMaestros)
          ? API.descargarMaestros().catch(() => null)
          : fetch(_gasUrl('descargarMaestros')).then(r => r.json()).catch(() => null)),
        // [FIX #4 asimetría stock] getStock vía API → Supabase (stock_enriquecido_rls) cuando lectura
        // directa ON, fallback GAS. ANTES: fetch crudo a la Hoja STOCK = stale si el stock se escribe
        // directo a Supabase. Cierra la ventana junto con precargarOperacional (que ya es Supabase-first).
        ((typeof API !== 'undefined' && API.getStock)
          ? API.getStock().catch(() => null)
          : fetch(_gasUrl('getStock')).then(r => r.json()).catch(() => null)),
        // [Frente 4] config vía API.getConfig → Supabase (wh.get_config), fallback GAS.
        ((typeof API !== 'undefined' && API.getConfig)
          ? API.getConfig().catch(() => null)
          : fetch(_gasUrl('getConfig')).then(r => r.json()).catch(() => null))
      ]);

      console.log('[Offline] descargarMaestros respuesta:', maestros);
      const maestrosChanged = [];
      if (maestros?.ok) {
        // GAS nuevo — recibe las 4 tablas de MOS
        const d = maestros.data;
        console.log('[Offline] personal recibido:', d?.personal?.length, 'registros');
        if (d.personal      != null) guardar(KEYS.PERSONAL,      d.personal);
        // [perf] Solo avisar 'productos'/'equivalencias' si REALMENTE cambiaron →
        // sin esto, cada refresh repintaba+flasheaba ProductosView con datos idénticos.
        _guardarSiCambia(KEYS.PRODUCTOS,     'productos',     d.productos,     'productos',     maestrosChanged);
        _guardarSiCambia(KEYS.EQUIVALENCIAS, 'equivalencias', d.equivalencias, 'equivalencias', maestrosChanged);
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
        if (productos?.ok)   _guardarSiCambia(KEYS.PRODUCTOS, 'productos', productos.data, 'productos', maestrosChanged);
        if (proveedores?.ok) guardar(KEYS.PROVEEDORES, proveedores.data);
      }

      if (stock?.ok)  guardar(KEYS.STOCK,  stock.data);
      if (config?.ok) guardar(KEYS.CONFIG, config.data);

      localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
      if (maestrosChanged.length) {
        window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: maestrosChanged } }));
      }
      // [poller catálogo] Sembrar el baseline de versión la PRIMERA vez que bajamos el maestro.
      // Así el poller compara contra la versión que efectivamente quedó en cache (no re-descarga
      // de inmediato). Si ya hay baseline (de localStorage o de un chequeo previo), no lo tocamos:
      // el avance del baseline es responsabilidad exclusiva de _chequearVersionCatalogo.
      if (_catVersionBaseline == null && typeof API !== 'undefined' && typeof API.catalogoVersion === 'function') {
        API.catalogoVersion().then(v => { if (_catVersionBaseline == null) _setBaselineCatalogo(v); }).catch(() => {});
      }
      _notificar();
    } catch(e) {
      console.warn('[Offline] Error en precarga:', e);
    } finally {
      _masterInflight = null;
    }
    })();
    return _masterInflight;
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

    // [100x rollback-fix] La vía de reintento se decide POR ÍTEM, NO por un flag global al sincronizar.
    //
    // BUG corregido: antes mirábamos `API._escrituraDirectaActiva()` en tiempo de sync. Si un ítem se
    // encolaba bajo escritura directa por TIMEOUT (la RPC PUDO commitear en Supabase y perderse la
    // respuesta) y LUEGO se hacía rollback de la fase (apagar el flag / limpiar localStorage / cambiar
    // de dispositivo) antes de vaciar la cola, `directoOn` veía false y mandaba ese ítem a GAS. GAS no
    // tiene ese localId en su SYNC_LOG (la op nunca pasó por GAS) → lo EJECUTABA → DOBLE STOCK / DOBLE GUÍA.
    //
    // FIX: api.js sella el ítem al encolar (`params._viaDirecta = true`) SOLO cuando nace del timeout de
    // escritura directa. Acá, un ítem así SIEMPRE se reintenta vía _postDirecto (idempotente por el id
    // sembrado del localId → la RPC dedupea), aunque el flag global ya esté apagado por rollback.
    //   • viaDirecta + API disponible → API._postCola(item.params) (dedup en Supabase).
    //   • viaDirecta + API NO disponible (módulo no cargado) → NO ir a GAS (duplicaría): dejar 'error' y reintentar luego.
    //   • legacy (sin la marca) → GAS, exactamente como hoy.
    // Con la escritura directa nunca activada, ningún ítem lleva la marca → 100% GAS = comportamiento actual (INERTE).
    var huboEnvasado = false;
    for (const item of queue) {
      try {
        const viaDirecta = !!(item._viaDirecta || (item.params && item.params._viaDirecta));
        let res;
        if (viaDirecta) {
          // Ítem que pudo commitear en Supabase → SIEMPRE reintento directo (dedup), nunca GAS.
          if (typeof API === 'undefined' || !API._postCola) {
            // Módulo de escritura directa no disponible: no podemos reintentar directo y mandarlo a
            // GAS duplicaría. Lo dejamos pendiente (marcado 'error') para reintentar cuando cargue.
            _actualizarItemQueue(item.localId, 'error');
            continue;
          }
          res = await API._postCola(item.params);
        } else {
          res = await fetch(window.WH_CONFIG.gasUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'text/plain' },
            body:    JSON.stringify(item.params)
          }).then(r => r.json());
        }

        // [400-loop fix] Un ítem `_viaDirecta` rechazado por el servidor con un 4xx
        // definitivo (no commiteó) llega con `_descartar:true`. NO tiene sentido
        // reintentarlo: lo damos por terminado ('synced' lo purga en limpiarSincronizados)
        // para que no spamee la consola con POST .../rpc/... 400 en cada ciclo de sync.
        // Solo descartamos ante un rechazo explícito del servidor, nunca ante timeout/red.
        if (res && res._descartar) {
          try { console.warn('[cola] ítem descartado por rechazo definitivo del servidor:', item.action, item.localId, res.error); } catch (_) {}
          _actualizarItemQueue(item.localId, 'synced');
          continue;
        }
        _actualizarItemQueue(item.localId, (res && res.ok) ? 'synced' : 'error');
        if (res && res.ok && item.action === 'registrarEnvasado') huboEnvasado = true;
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

  // [v2.13.173 BUG FIX] Registro de CAMPOS de preingreso con escritura local en
  // vuelo (sin confirmar en el Sheet). Mientras un campo esté pendiente,
  // _mergePreingresos preserva su valor LOCAL para que el polling de 60s no
  // revierta lo que el backend aún no terminó de persistir. Patrón hermano de
  // _subiendoFotos (Fix #2 v2.11.1), pero scoped por id+campo y con TTL para
  // que un fallo de red no congele el campo para siempre. Corrige la pérdida
  // tanto de `cargadores` como de `comentario`/`monto` al refrescar.
  const _preingPendientes = new Map(); // idPreingreso -> { campo: expiraTs }
  function marcarPreingresoPendiente(id, campos, ttlMs) {
    if (!id || !campos) return;
    const k = String(id);
    let m = _preingPendientes.get(k);
    if (!m) { m = {}; _preingPendientes.set(k, m); }
    const exp = Date.now() + (ttlMs || 15000);
    (Array.isArray(campos) ? campos : [campos]).forEach(c => { m[c] = exp; });
  }
  function _preingCampoPendiente(id, campo) {
    const m = _preingPendientes.get(String(id));
    if (!m || !m[campo]) return false;
    if (Date.now() > m[campo]) {
      delete m[campo];
      if (!Object.keys(m).length) _preingPendientes.delete(String(id));
      return false;
    }
    return true;
  }
  // Atajo para el caso más común (estado de carretas).
  function marcarCargadoresPendiente(id, ttlMs) { marcarPreingresoPendiente(id, 'cargadores', ttlMs); }

  async function precargarOperacional(forzar = false) {
    if (!navigator.onLine) return;
    // [perf v2.13.242] Dedup in-flight: si ya hay una precarga operacional corriendo,
    // los callers concurrentes (nav rápido entre módulos + timer 60s + visibilitychange +
    // cada View.cargar) REUSAN esa promesa. Antes el guard `_opLoading` devolvía
    // undefined → el caller leía cache STALE y, peor, "saltar rápido" encolaba intentos.
    // Ahora coalescen en UNA descarga; sus .then leen el MISMO cache fresco.
    if (_opInflight) return _opInflight;
    if (!forzar && Date.now() - _lastOpTs < OP_MIN_MS) return;
    _opLoading = true;
    _lastOpTs  = Date.now();
    _opInflight = (async () => {
    try {
      // [BUG A · cutover] Pasar por API.descargarOperacional (no fetch crudo a GAS): así,
      // si el dispositivo escribe/lee directo a Supabase, el operacional se trae DIRECTO y el
      // listado de Guías ve sus propias guías directas 'G_L...' (antes el GAS stale nunca las
      // mostraba). API.descargarOperacional cae a GAS solo ante fallo. Fallback al fetch crudo
      // si API no está cargada (orden de scripts / arranque temprano).
      const r = (typeof API !== 'undefined' && API.descargarOperacional)
        ? await API.descargarOperacional().catch(() => null)
        : await fetch(_gasUrl('descargarOperacional')).then(r => r.json()).catch(() => null);
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
          // [v2.13.173 BUG FIX] Preservar campos editables con escritura local
          // en vuelo. Sin esto, el poll de 60s revertía lo que el operador
          // acababa de cambiar pero que el backend aún no había grabado
          // (síntoma: "el cambio se pierde / vuelve al inicial" — afectaba a
          // cargadores Y a comentario/monto).
          ['cargadores', 'comentario', 'monto', 'idProveedor'].forEach(campo => {
            if (_preingCampoPendiente(n.idPreingreso, campo) && old[campo] != null) {
              merged[campo] = old[campo];
            }
          });
          return merged;
        });
      }

      // [40x] No pisar cache bueno con un dataset VACÍO: un poll/RPC que devolvió [] no debe borrar
      // datos previos. Solo persiste si trae filas, o si el cache ya estaba vacío.
      const _persist = (key, arr, label) => {
        if (arr == null) return;
        if (!arr.length && (cargar(key) || []).length) return;   // vacío sobre lleno → skip
        if (_hayDiff(arr, key)) { guardar(key, arr); changed.push(label); }
        else guardar(key, arr);
      };
      _persist(KEYS.GUIAS,        d.guias,    'guias');
      _persist(KEYS.GUIA_DETALLE, d.detalles, 'detalles');
      if (d.preingresos != null) {
        // Si hay subida de fotos en curso, omitir refresh de preingresos
        // (Fix #2). Si no, aplicar merge defensivo (Fix #1) y guardar.
        if (_subiendoFotos) {
          if (window.__WH_DEBUG_FOTOS) console.log('[Offline] skip refresh preingresos: subida en curso');
        } else {
          const viejos  = cargar(KEYS.PREINGRESOS) || [];
          const merged  = _mergePreingresos(d.preingresos, viejos);
          // [40x] no pisar con vacío si había datos (el merge ya preserva, guard defensivo)
          if (!(merged.length === 0 && viejos.length)) {
            if (_hayDiff(merged, KEYS.PREINGRESOS)) { guardar(KEYS.PREINGRESOS, merged); changed.push('preingresos'); }
            else                                    { guardar(KEYS.PREINGRESOS, merged); }
          }
        }
      }
      _persist(KEYS.STOCK,        d.stock,      'stock');
      _persist(KEYS.AJUSTES,      d.ajustes,    'ajustes');
      _persist(KEYS.AUDITORIAS_C, d.auditorias, 'auditorias');

      // [perf v2.13.242] Solo notificar si REALMENTE cambió algo. Antes se
      // disparaba wh:data-refresh en cada poll aunque `changed` fuera []; aunque
      // los listeners filtran por dataset, el evento igual despierta a todos los
      // handlers en cada ciclo. Sin cambios → no se molesta a nadie.
      if (changed.length) {
        window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed } }));
      }
    } catch(e) { console.warn('[Offline] Error en precarga operacional:', e); }
    finally { _opLoading = false; _opInflight = null; }
    })();
    return _opInflight;
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

  // ── Poller de versión del catálogo ───────────────────────────
  // Sondea mos.catalogo_version (1 query liviana, profile 'mos'). Si la versión subió respecto
  // al baseline → re-descarga el catálogo completo y avanza el baseline. EFICIENTE: solo corre con
  // la app visible; NO re-descarga si la versión es igual; ante cualquier fallo deja el baseline
  // intacto (no re-descarga "por las dudas"). MONEY-SAFE: la re-descarga es del catálogo (datos de
  // referencia) vía precargar('manual') → _guardarSiCambia + wh:data-refresh + silentRefresh, que
  // NO toca formularios/carritos en armado (guía/envasado/venta). No es un reload de la app.
  function _setBaselineCatalogo(v) {
    _catVersionBaseline = v;
    try { localStorage.setItem(KEYS.CAT_VERSION, String(v)); } catch (_) {}
  }

  async function _chequearVersionCatalogo(motivo) {
    // Guardas de eficiencia: solo con red, app visible y la API disponible.
    if (!navigator.onLine) return;
    if (typeof document !== 'undefined' && document.visibilityState && document.visibilityState !== 'visible') return;
    if (typeof API === 'undefined' || typeof API.catalogoVersion !== 'function') return;
    // [perf 500x] throttle: foco/visibility/timer pueden coincidir → no spamear el round-trip de versión.
    // 'init' (1er chequeo del arranque) se exime para sembrar el baseline de inmediato.
    const _now = Date.now();
    if (motivo !== 'init' && (_now - _catLastCheck) < CAT_CHECK_THROTTLE_MS) return;
    _catLastCheck = _now;
    if (_catPollBusy) return;                     // un chequeo a la vez (timer + foco + visibility coinciden)
    _catPollBusy = true;
    try {
      const v = await API.catalogoVersion();      // LANZA ante fallo → catch → baseline intacto
      // Baseline aún no fijado (1er chequeo si no se sembró tras la 1ra descarga): adoptar y salir.
      if (_catVersionBaseline == null) { _setBaselineCatalogo(v); return; }
      if (v <= _catVersionBaseline) return;       // sin cambios → NO re-descargar
      await _aplicarVersionCatalogo(v, motivo || 'poll');
    } catch (_) {
      /* fallo de red/RPC: dejar baseline intacto → reintenta en el próximo ciclo */
    } finally {
      _catPollBusy = false;
    }
  }

  // Núcleo compartido por el POLLER (_chequearVersionCatalogo) y el REALTIME
  // (notificarVersionCatalogo): dada una versión NUEVA (> baseline), re-descarga
  // el catálogo y avanza el baseline. MONEY-SAFE: la re-descarga es vía
  // precargar('manual') → _guardarSiCambia + wh:data-refresh + silentRefresh, que
  // NO resetea formularios/carritos en armado (guía/envasado/venta) — no es un
  // reload de la app. NO toca el baseline si la re-descarga lanzó (se reintenta).
  async function _aplicarVersionCatalogo(v, motivo) {
    // [perf 500x] COALESCING: en vez de re-descargar ~1.9MB por CADA bump (las versiones suben en ráfaga),
    // diferimos y agrupamos: tomamos la versión más alta y descargamos UNA sola vez tras una ventana de
    // quietud. Si ya hay una descarga programada, solo actualizamos el objetivo (no apilamos descargas).
    _catPendingVersion = Math.max(Number(_catPendingVersion) || 0, Number(v) || 0);
    if (_catRedownloadTimer) return;
    console.log('[Offline] catálogo ' + _catVersionBaseline + ' → ' + _catPendingVersion + ' (' + (motivo || 'evento') + ') · re-descarga diferida ' + (CAT_REDOWNLOAD_DEBOUNCE_MS / 1000) + 's (coalesce)');
    _catRedownloadTimer = setTimeout(async () => {
      _catRedownloadTimer = null;
      const target = _catPendingVersion;
      await precargar('manual').catch(() => {});
      _setBaselineCatalogo(target);              // avanzar baseline SOLO tras intentar la re-descarga
      if (typeof toast === 'function') toast('Catálogo actualizado', 'info', 2200);
    }, CAT_REDOWNLOAD_DEBOUNCE_MS);
  }

  // [Realtime] Llamado por la suscripción Realtime de api.js al recibir un UPDATE de
  // mos.catalogo_meta con record.version. Comparte el guard anti-reentrada + el núcleo
  // money-safe del poller. Si la versión NO subió respecto al baseline → no hace nada
  // (el poller ya cubría ese caso). Si el baseline aún no estaba sembrado, lo adopta sin
  // re-descargar (la 1ra precarga ya trajo ese estado). Es ADITIVO: el poller de ~50s
  // sigue como red de seguridad si el WebSocket cae.
  async function notificarVersionCatalogo(v, motivo) {
    const nv = Number(v);
    if (!Number.isFinite(nv)) return;
    if (!navigator.onLine) return;
    if (_catPollBusy) return;                     // poll/foco/visibility en curso → ese ciclo lo cubre
    _catPollBusy = true;
    try {
      if (_catVersionBaseline == null) { _setBaselineCatalogo(nv); return; }
      if (nv <= _catVersionBaseline) return;      // ya estamos a esa versión o más → nada que hacer
      await _aplicarVersionCatalogo(nv, motivo || 'realtime');
    } catch (_) {
      /* fallo de red/RPC: baseline intacto → el poller reintenta */
    } finally {
      _catPollBusy = false;
    }
  }

  // Inicia el poller de versión del catálogo. Idempotente. Llamar tras arrancar el refresh
  // operacional (App.init). Siembra el baseline desde localStorage si existe (continuidad entre
  // recargas) y hace un primer chequeo que lo adopta/actualiza.
  function iniciarPollerCatalogo() {
    if (_catPollTimer) return;
    if (_catVersionBaseline == null) {
      const guardado = (() => { try { return localStorage.getItem(KEYS.CAT_VERSION); } catch (_) { return null; } })();
      if (guardado != null && guardado !== '' && Number.isFinite(Number(guardado))) _catVersionBaseline = Number(guardado);
    }
    // Chequeo inmediato: fija baseline si no había, o detecta cambios ocurridos mientras la app estaba cerrada.
    _chequearVersionCatalogo('init');
    _catPollTimer = setInterval(() => { _chequearVersionCatalogo('timer'); }, CAT_POLL_MS);
    // Volver a foreground / recuperar el foco → chequear ya (no esperar al timer).
    if (typeof document !== 'undefined') {
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') _chequearVersionCatalogo('visible');
      });
    }
    if (typeof window !== 'undefined') {
      window.addEventListener('focus', () => { _chequearVersionCatalogo('focus'); });
    }
  }

  function detenerPollerCatalogo() {
    if (_catPollTimer) { clearInterval(_catPollTimer); _catPollTimer = null; }
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

  // ── Patch de una guía existente en caché ────────────────────
  // [v2.13.186 BUG reabrir] Reabrir/cerrar solo actualizaba la guía en memoria
  // (todas[idx] + _guiaActual). El cache wh_guias quedaba con el estado VIEJO →
  // cualquier silentRefresh (que lee getGuiasCache) revertía visualmente el
  // estado y la guía reabierta volvía a verse CERRADA = no se podía editar
  // cantidad. Mismo patrón que patchPreingresosCache.
  function patchGuiaCache(idGuia, changes) {
    const cache = cargar(KEYS.GUIAS) || [];
    const idx   = cache.findIndex(g => g.idGuia === idGuia);
    if (idx >= 0) { Object.assign(cache[idx], changes); guardar(KEYS.GUIAS, cache); }
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
    patchGuiaCache,
    getPNCache, setPNCache,
    getEnvasadosCache, guardarEnvasadosCache, inyectarEnvasadoCache, removerEnvasadoCache,
    precargarOperacional, iniciarRefreshOperacional, detenerRefreshOperacional,
    iniciarPollerCatalogo, detenerPollerCatalogo, notificarVersionCatalogo,
    setSubiendoFotos, isSubiendoFotos,
    marcarPreingresoPendiente, marcarCargadoresPendiente,
    estaOnline: () => navigator.onLine
  };
})();
