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
    CAT_VERSION:  'wh_cat_version',       // [poller] baseline de mos.catalogo_version visto por este equipo
    CAT_SYNC_TS:  'wh_cat_sync_ts'        // [delta] server_ts del último catálogo bajado (punto de corte del delta)
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

  // [2.13.604] Al reconectar también se retoman los ↺ Deshacer de envasados encolados (wh_env_deshacer): sincronizar()
  // sale temprano si la cola está vacía, y una intención persistida (p.ej. tras recargar) igual debe ejecutarse.
  window.addEventListener('online',  () => { _notificar(); sincronizar(); setTimeout(() => { try { _envDeshacerProcesar(); } catch (_) {} }, 1500); });
  window.addEventListener('offline', () => _notificar());

  // [FIX lag sync] Reintento PERIÓDICO de la cola. Antes sincronizar() solo corría en el evento 'online',
  // al aplicar sesión o con el botón manual → un ítem que falló (timeout de escritura directa) quedaba
  // pegado mostrando "N operaciones por sincronizar" hasta que la red parpadeara ("demora mucho").
  // Cada 20s drenamos la cola si hay red e ítems pendientes/error. Idempotente (dedup por localId) y
  // reentrante-seguro (el guard _syncing evita solapes). No hace nada si la cola está vacía.
  setInterval(() => {
    if (!navigator.onLine || _syncing) return;
    if (getQueue().some(i => i.status === 'pending' || i.status === 'error')) sincronizar();
    // [2.13.604] Intenciones de deshacer pendientes (reintentos con backoff propio; no hace nada si no hay).
    else { try { _envDeshacerProcesar(); } catch (_) {} }
  }, 20000);

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
        // [BLINDAJE carrito · data-loss] Todo el namespace `wh_despacho_*` (carrito, pickup activo, zona,
        // historial y CUALQUIER estado futuro de despacho) + la lista sombra son TRABAJO EN CURSO del operador,
        // NO caches regenerables. Estaban fuera de la allowlist → el cleanup por cambio de versión los borraba y
        // se perdía la lista al actualizar. Se preservan por PREFIJO (no por key puntual) para que un estado
        // nuevo de despacho quede protegido automáticamente sin tener que acordarse de agregarlo aquí.
        // [2.13.603] wh_lote_auto = adhesivos pendientes de envasados en cola: es estado ligado a wh_queue, no cache.
        // [2.13.604] wh_env_deshacer = ↺ Deshacer de envasados encolados pendientes de anular: también ligado a wh_queue.
        const PRESERVAR = /^(wh_sesion|wh_device_id|wh_app_version|wh_audio_ok|wh_perms_done_v.*|wh_personal|wh_admin_cache|wh_queue|wh_gas_url|wh_despacho_.*|wh_lista_sombra|wh_lote_auto|wh_env_deshacer)$/;
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

  // [2.13.604] Tiers de desalojo por cuota en constantes de módulo: los usa guardar() (abajo, misma lógica de
  // siempre) y también wh_env_deshacer, que se escribe en crudo pero debe poder liberar espacio igual.
  const _STORAGE_TIER1 = [
    'wh_productos','wh_stock','wh_proveedores','wh_ajustes',
    'wh_auditorias_c','wh_ubicaciones','wh_equivalencias','wh_zonas',
    'wh_impresoras','wh_pn','wh_config','wh_guias'
  ];
  const _STORAGE_TIER2 = ['wh_guia_detalle','wh_preingresos','wh_envasados'];

  function guardar(key, data) {
    // [v2.13.110] Invalidar cache de parseo ANTES de escribir — garantiza
    // que cualquier read concurrente que ocurra entre la invalidación y
    // el setItem caiga al miss y refresque.
    _invalidarParseCache(key);
    // [2.13.596] Respaldo en memoria: si el storage está lleno (origen levo19.github.io compartido con
    // MOS/ME) el cleanup desaloja caches TIER1 entre sí y wh_guias quedaba vacío → "no cargan las guías".
    _memFallback.set(key, data);
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
        const TIER1 = _STORAGE_TIER1;   // [2.13.604] mismas listas, ahora en constantes de módulo
        const TIER2 = _STORAGE_TIER2;

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
  const _memFallback = new Map();  // key → último data escrito por guardar() (sobrevive al desalojo por cuota)

  function cargar(key) {
    // 1. Hit en cache de parseo
    const cached = _parseCache.get(key);
    if (cached && (Date.now() - cached.ts) < _PARSE_TTL_MS) {
      return cached.data;
    }
    // 2. Miss → descomprimir + parsear
    const raw = localStorage.getItem(key);
    const obj = _deserializar(raw);
    let data = obj ? obj.data : null;
    // [2.13.596] desalojado por cuota → lo último escrito en esta sesión (nunca peor que "vacío")
    if (data === null && !raw && _memFallback.has(key)) data = _memFallback.get(key);
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
  // [534] Firma PERSISTIDA. `_masterSig` vivía solo en memoria, así que después de CADA
  // recarga (F5, update de SW, la recarga de mantenimiento, o simplemente reabrir la PWA)
  // la firma arrancaba vacía y el catálogo se volvía a comprimir ENTERO con
  // LZString.compressToUTF16 — cientos de ms en PC y VARIOS SEGUNDOS de hilo principal
  // bloqueado en un celular de gama baja, aunque el catálogo fuera idéntico al que ya
  // estaba en localStorage. Guardando la firma, una recarga con el catálogo sin cambios
  // se salta por completo la compresión.
  const _SIG_KEY = 'wh_master_sig';
  function _cargarSigs() {
    try { return JSON.parse(localStorage.getItem(_SIG_KEY)) || {}; } catch (_) { return {}; }
  }
  function _persistirSig(sigKey, sig) {
    try {
      const all = _cargarSigs();
      all[sigKey] = sig;
      localStorage.setItem(_SIG_KEY, JSON.stringify(all));
    } catch (_) {}
  }
  // Marca cambio en `changed` SOLO si la firma del dataset difiere de la última vista.
  function _guardarSiCambia(key, sigKey, arr, label, changed) {
    if (arr == null) return;
    const sig = _firma(arr);
    // Idéntico a lo último visto Y ya hay algo persistido → no reescribir ni avisar a
    // la UI (evita el flash + el costo de recomprimir un array grande en cada refresh).
    const sigPrev = (_masterSig[sigKey] !== undefined) ? _masterSig[sigKey] : _cargarSigs()[sigKey];
    if (sigPrev === sig && (cargar(key) || []).length) { _masterSig[sigKey] = sig; return; }
    _masterSig[sigKey] = sig;
    guardar(key, arr);
    _persistirSig(sigKey, sig);
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
      // [CERO-GAS 2026-07-19] SOLO Supabase directo (API.*). Los fetch crudos a GAS se
      // eliminaron: en el boot frío el primer intento directo puede perder la carrera del
      // mint → antes degradaba a GAS (3 GETs a la Hoja stale); ahora reintenta DIRECTO
      // (ver _retryMaestros abajo) y mientras tanto sirve el cache local.
      const [maestros, stock, config] = await Promise.all([
        (typeof API !== 'undefined' && API.descargarMaestros) ? API.descargarMaestros().catch(() => null) : null,
        (typeof API !== 'undefined' && API.getStock)          ? API.getStock().catch(() => null)          : null,
        (typeof API !== 'undefined' && API.getConfig)         ? API.getConfig().catch(() => null)         : null
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
        // [CERO-GAS · audit 2026-07-13] El servidor ya NO emite d.adminPin (PIN plano de Sheet).
        // Purgamos activamente cualquier valor que haya quedado cacheado en equipos.
        try { localStorage.removeItem(KEYS.ADMIN_PIN); } catch (_) {}
        // [delta] guardar el punto de corte del servidor → el próximo refresh por versión baja solo el delta.
        if (maestros.server_ts) { try { localStorage.setItem(KEYS.CAT_SYNC_TS, String(maestros.server_ts)); } catch (_) {} }
        if (maestros.errores?.length) console.warn('[Offline] descargarMaestros errores:', maestros.errores);
        // [534] Bajó bien → resetear el backoff. Sin esto el contador solo subía y, tras
        // una racha de fallos, un equipo YA recuperado seguía esperando 60s entre refrescos.
        _retryMaestrosN = 0;
        if (_retryMaestrosTid) { clearTimeout(_retryMaestrosTid); _retryMaestrosTid = null; }
      } else {
        // [CERO-GAS 2026-07-19] El directo no respondió (boot frío: carrera del mint / red).
        // NADA de GAS: el login sirve del cache local y programamos UN reintento directo.
        console.warn('[Offline] descargarMaestros directo no respondió — reintento directo en 6s (cero-GAS)');
        _retryMaestros();
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
      return { ok: true };               // [500x HIGH] éxito explícito → el caller puede avanzar el baseline
    } catch(e) {
      console.warn('[Offline] Error en precarga:', e);
      return { ok: false };              // [500x HIGH] fallo → NO avanzar el baseline (reintentar próximo ciclo)
    } finally {
      _masterInflight = null;
    }
    })();
    return _masterInflight;
  }

  // [perf 500x · CD1 DELTA] Refresco INCREMENTAL del catálogo: baja SOLO los productos cambiados desde el
  // último corte (server_ts) y los mergea por idProducto en el cache, en vez de re-bajar ~1.9MB. Las tablas
  // chicas viajan completas. Sin punto de corte o ante CUALQUIER error → cae a la descarga completa (que
  // además tiene fallback GAS). Lo usa _aplicarVersionCatalogo en cada bump de versión.
  async function _refrescarCatalogoDelta() {
    if (!navigator.onLine) return;
    const desdeTs = (() => { try { return localStorage.getItem(KEYS.CAT_SYNC_TS); } catch (_) { return null; } })();
    if (!desdeTs || typeof API === 'undefined' || typeof API.descargarMaestrosDelta !== 'function') {
      return precargar('manual');                 // sin corte → descarga completa (siembra el corte + server_ts)
    }
    if (_masterInflight) return _masterInflight;   // una descarga completa ya en vuelo → reusarla
    let delta;
    try { delta = await API.descargarMaestrosDelta(desdeTs); }
    catch (e) { console.warn('[Offline] delta catálogo falló → full:', e && e.message); return precargar('manual'); }
    if (!delta || !delta.ok) return precargar('manual');
    const d = delta.data || {};
    const changed = [];
    // [500x HIGH delete] ids borrados desde el corte → sacar del cache (un delete no viaja por updated_at).
    const elim = Array.isArray(delta.eliminados) ? delta.eliminados : [];
    const elimSet = elim.length ? new Set(elim) : null;
    // productos: MERGE por idProducto sobre una COPIA (no mutar la referencia del parse-cache), + quitar borrados.
    if ((Array.isArray(d.productos) && d.productos.length) || elimSet) {
      let cache = (cargar(KEYS.PRODUCTOS) || []).slice();   // [500x MED] copia: no mutar el array que retiene la UI
      const idx = {};
      for (let i = 0; i < cache.length; i++) { const id = cache[i] && cache[i].idProducto; if (id != null) idx[id] = i; }
      (d.productos || []).forEach(p => {
        const id = p && p.idProducto;
        if (id != null && idx[id] != null) cache[idx[id]] = p; else cache.push(p);
      });
      if (elimSet) cache = cache.filter(p => !(p && elimSet.has(p.idProducto)));
      _guardarSiCambia(KEYS.PRODUCTOS, 'productos', cache, 'productos', changed);
    }
    // [500x LOW] tablas chicas por _guardarSiCambia (no reescribir/repintar si no cambiaron).
    if (d.equivalencias != null) _guardarSiCambia(KEYS.EQUIVALENCIAS, 'equivalencias', d.equivalencias, 'equivalencias', changed);
    if (d.proveedores   != null) _guardarSiCambia(KEYS.PROVEEDORES,   'proveedores',   d.proveedores,   'proveedores',   changed);
    if (d.personal      != null) _guardarSiCambia(KEYS.PERSONAL,      'personal',      d.personal,      'personal',      changed);
    if (d.impresoras    != null) _guardarSiCambia(KEYS.IMPRESORAS,    'impresoras',    d.impresoras,    'impresoras',    changed);
    if (d.zonas         != null) _guardarSiCambia(KEYS.ZONAS,         'zonas',         d.zonas,         'zonas',         changed);
    if (delta.server_ts) { try { localStorage.setItem(KEYS.CAT_SYNC_TS, String(delta.server_ts)); } catch (_) {} }
    localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
    console.log('[Offline] catálogo DELTA: ' + (delta.productos_cambiados || 0) + ' cambiados · ' + elim.length + ' borrados' + (changed.length ? ' (' + changed.join(',') + ')' : ''));
    if (changed.length) { try { window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed } })); } catch (_) {} }
    _notificar();
    return { ok: true, delta: true };
  }

  // [CERO-GAS] Reintento directo del maestro tras un boot frío fallido. Máx 3 intentos
  // espaciados; si todos fallan, los ciclos normales (login/welcome/reconexión) lo retoman.
  let _retryMaestrosN = 0;
  // [534] Antes: 3 reintentos y SILENCIO ABSOLUTO para siempre.
  //
  // Síntoma real que producía: si el maestro no bajaba en el arranque (mint-wh que
  // 401ea, edge en cold start, red a medias, dispositivo recién reactivado), el equipo
  // se quedaba con el catálogo VACÍO o viejo y NADIE se enteraba: la app se veía normal,
  // pero Productos salía vacío, los modales que dependen del catálogo abrían en blanco y
  // el ajuste de stock no encontraba el producto. Eso explica "a unos dispositivos no les
  // carga tal modal" y "no pueden ajustar stock" — nunca fue el modal, era el dato.
  //
  // Ahora: backoff que NO se rinde (6s,12s,18s,30s,60s… tope 60s) y, si tras 3 fallos
  // seguidos seguimos sin catálogo, aviso VISIBLE y accionable al operador.
  const _RETRY_MAX_MS = 60000;
  let _retryMaestrosTid = null;
  function _retryMaestros() {
    if (_retryMaestrosTid) return;                 // ya hay uno agendado → no apilar
    _retryMaestrosN++;
    const espera = Math.min(6000 * _retryMaestrosN, _RETRY_MAX_MS);

    if (_retryMaestrosN === 3) {
      // ¿Es degradación (hay cache viejo) o apagón total (no hay nada que mostrar)?
      let hayCatalogo = 0;
      try { hayCatalogo = (cargar(KEYS.PRODUCTOS) || []).length; } catch (_) {}
      try {
        if (typeof toast === 'function') {
          toast(hayCatalogo
            ? '⚠ No se pudo actualizar el catálogo — trabajando con datos guardados. Revisa la señal.'
            : '❌ Sin catálogo: este equipo no pudo descargar los productos. Revisa la señal o avisa al admin.',
            hayCatalogo ? 'warn' : 'danger', 10000);
        }
      } catch (_) {}
      try { console.warn('[Offline] catálogo NO descargado tras 3 intentos · productos en cache:', hayCatalogo); } catch (_) {}
    }

    _retryMaestrosTid = setTimeout(() => {
      _retryMaestrosTid = null;
      try { _lastMasterTs = 0; precargar(true); } catch (_) {}
    }, espera);
  }

  // [CERO-GAS 2026-07-19] _gasUrl ELIMINADO: no queda NINGÚN camino a GAS en este módulo.

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
    // [2.13.604] Al fijar el resultado se limpia la marca de "en vuelo" (enVueloTs): el envío ya terminó.
    // Lectura FRESCA (QA L2): otra pestaña (↺ Deshacer) pudo cambiar la cola durante el await del envío.
    _invalidarParseCache(KEYS.QUEUE);
    const queue = getQueue().map(i => {
      if (i.localId !== localId) return i;
      const { enVueloTs, ...resto } = i;
      return { ...resto, status };
    });
    guardar(KEYS.QUEUE, queue);
  }

  // [v2.13.376] Parchea el fechaVencimiento de un agregarDetalleGuia aún PENDIENTE en la
  // cola offline. Caso: el operador agrega una línea estando OFFLINE y luego le pone el
  // vencimiento inline — el alta se encoló con fechaVencimiento:'' y la cola reaplica el
  // payload tal cual → el venc se perdía al sincronizar. Esto actualiza la op encolada para
  // que el alta lleve el venc actual (un solo INSERT con el dato correcto). Devuelve true si
  // parcheó. El localId de la op = 'DET_'+idLocal (catálogo) o 'PNDET_'+idLocal (producto nuevo).
  function patchPendingDetalleVenc(lineLocalId, fechaVencimiento) {
    if (!lineLocalId) return false;
    let patched = false;
    const upd = getQueue().map(it => {
      if (it.status !== 'pending' || it.action !== 'agregarDetalleGuia') return it;
      if (it.localId === 'DET_' + lineLocalId || it.localId === 'PNDET_' + lineLocalId) {
        patched = true;
        return { ...it, params: { ...it.params, fechaVencimiento: fechaVencimiento || '' } };
      }
      return it;
    });
    if (patched) guardar(KEYS.QUEUE, upd);
    return patched;
  }

  // ── [2.13.603] Adhesivos de envasados encolados offline ─────────────────────
  // Problema (QA tras 2.13.600): crearLoteAdhesivo ya NUNCA se encola (evita lotes duplicados), así que un
  // ENVASADO registrado sin red subía al reconectar pero sus adhesivos no salían solos (había que tocar
  // "Reintentar"). Este registro persistido recuerda "este envasado (clave ENV-…) quiere adhesivos" y su
  // estado, para que app.js (WhLoteAuto) los imprima UNA vez cuando la cola confirme el envasado.
  // Va en localStorage CRUDO (sin el cache de parseo de cargar(), que es por pestaña y con TTL 15s): dos
  // pestañas deben ver el mismo estado al instante. Es chico (unas pocas entradas) y se purga solo.
  // Estados: PENDIENTE (envasado en cola) → SINCRONIZADO (la cola lo confirmó) → TOMADO (una pestaña lo
  // está creando) → HECHO. Laterales: AVISO (>12h, se pregunta), DESCARTADO (el operador dijo que no),
  // DESHECHO (↺ Deshacer sobre el optimista: jamás imprimir).
  const LOTE_AUTO_KEY = 'wh_lote_auto';
  const _LOTE_AUTO_FINALES = ['HECHO', 'DESCARTADO', 'DESHECHO'];
  function _loteAutoLeer() {
    try { const o = JSON.parse(localStorage.getItem(LOTE_AUTO_KEY) || '{}'); return (o && typeof o === 'object') ? o : {}; }
    catch (_) { return {}; }
  }
  function _loteAutoEscribir(mapa) {
    // Purga: finales >3 días y cualquier entrada >7 días (un PENDIENTE cuyo envasado la cola descartó).
    const ahora = Date.now();
    Object.keys(mapa).forEach(k => {
      const e = mapa[k] || {};
      const edad = ahora - (e.tsRegistro || e.tsEstado || 0);
      if (edad > 7 * 864e5 || (_LOTE_AUTO_FINALES.indexOf(e.estado) >= 0 && edad > 3 * 864e5)) delete mapa[k];
    });
    try { localStorage.setItem(LOTE_AUTO_KEY, JSON.stringify(mapa)); return true; }
    catch (_) { return false; }   // storage lleno: el peor caso es volver al "Reintentar" manual de hoy
  }
  // Alta SOLO si no existe: si el lote ya se confirmó (HECHO) antes de que llegue el .then del envasado, no se pisa.
  function loteAutoRegistrar(clave, datos) {
    if (!clave) return false;
    const m = _loteAutoLeer();
    if (m[clave]) return false;
    m[clave] = { ...datos, clave, estado: 'PENDIENTE', tsEstado: Date.now(), tsRegistro: (datos && datos.tsRegistro) || Date.now() };
    return _loteAutoEscribir(m);
  }
  function loteAutoGet(clave) { return clave ? (_loteAutoLeer()[clave] || null) : null; }
  function loteAutoListar() { const m = _loteAutoLeer(); return Object.keys(m).map(k => m[k]); }
  // Parche de estado. `crear:true` permite crear la entrada mínima (p.ej. HECHO cuando el lote se confirmó
  // en línea antes de que existiera el registro). `soloSi` = lista de estados desde los que se permite el cambio.
  function loteAutoSet(clave, patch, opts) {
    if (!clave) return false;
    const m = _loteAutoLeer();
    const prev = m[clave];
    if (!prev && !(opts && opts.crear)) return false;
    if (prev && opts && opts.soloSi && opts.soloSi.indexOf(prev.estado) < 0) return false;
    m[clave] = { ...(prev || { clave, tsRegistro: Date.now() }), ...patch, tsEstado: Date.now() };
    return _loteAutoEscribir(m);
  }
  // ↺ Deshacer sobre el optimista (ENV_OPT_*): la entrada se ubica por ese id.
  function loteAutoDeshacerPorOpt(idOpt) { return _loteAutoDeshacerPor('idOpt', idOpt); }
  // [2.13.603 · QA H3] Anular un envasado ya con id real (ENV_L…): la entrada guarda idEnvasado al sincronizar.
  function loteAutoDeshacerPorId(idEnvasado) { return _loteAutoDeshacerPor('idEnvasado', idEnvasado); }
  function _loteAutoDeshacerPor(campo, valor) {
    if (!valor) return false;
    const m = _loteAutoLeer();
    let hit = false;
    Object.keys(m).forEach(k => {
      if (m[k] && m[k][campo] === valor && _LOTE_AUTO_FINALES.indexOf(m[k].estado) < 0) {
        m[k] = { ...m[k], estado: 'DESHECHO', tsEstado: Date.now() }; hit = true;
      }
    });
    return hit ? _loteAutoEscribir(m) : false;
  }

  // ── [2.13.604] ↺ Deshacer de un ENVASADO registrado sin red (QA H4) ─────────────────────────────────
  // BUG: el Deshacer sobre un envasado ENCOLADO (id optimista ENV_OPT_*, ítem registrarEnvasado en wh_queue)
  // solo revertía el stock local; el .then del registro ya había corrido con res.offline (idReal === optimista,
  // nada que anular) y el ítem seguía en la cola → al reconectar el envasado SE CREABA igual y movía stock.
  // Diseño (siempre por la CLAVE del envasado = idempotencyKey 'ENV-…', jamás por nombre/producto):
  //   1) Ítem en la cola y NO en vuelo → se QUITA de wh_queue (síncrono → atómico frente a sincronizar() de esta
  //      pestaña, que re-chequea la cola viva antes de mandar cada envasado). Queda además una intención "por
  //      anular" SIN confirmar sobre su id determinista ('ENV_' + localId): si el ítem nació de un TIMEOUT, la RPC
  //      pudo haber commiteado → se anula; si nunca llegó, el servidor responde ENVASADO_NO_ENCONTRADO y se da
  //      por hecho (tras una ventana de gracia, por si el commit del timeout llega tarde).
  //   2) Ítem en vuelo (o ya enviado y purgado) → intención persistida; cuando la cola lo CONFIRMA con su id real
  //      se anula con la MISMA operación del Deshacer en línea (API.anularEnvasadoManual → wh.anular_envasado,
  //      idempotente por estado: un segundo intento devuelve yaAnulado, nunca revierte dos veces).
  //   3) Envasado que PUDO crearse por la cola con >12 h desde el registro, o rechazo definitivo del servidor →
  //      no se fuerza nada: aviso al operador (modal "Entendido"). Una intención de un ítem QUITADO de la cola no
  //      tiene ese tope (QA M2): lo esperable es NO_ENCONTRADO → hecho. APP_NO_AUTORIZADA / 401 / 403 / red son
  //      transitorias con backoff; si se repiten por más de 12 h → VENCIDO con aviso (QA M3/L1).
  //   4) Sin espacio para guardar la intención (QA H1) → NO se toca la cola y el Deshacer se informa como fallido.
  // Va en localStorage CRUDO (como wh_lote_auto): debe sobrevivir a una recarga y verse entre pestañas.
  // Estados: ESPERA_COLA (ítem en vuelo) → POR_ANULAR → ANULANDO (lease 60 s: una pestaña lo está anulando) →
  // HECHO. Finales laterales: RECHAZADO (el servidor no lo anuló), VENCIDO (>12 h, no se tocó).
  const ENV_DESHACER_KEY = 'wh_env_deshacer';
  const _ED_FINALES   = ['HECHO', 'RECHAZADO', 'VENCIDO'];
  const _ED_MAX_MS    = 12 * 3600e3;    // ventana para anular solo (mismo criterio que los adhesivos en cola)
  const _ED_GRACIA_MS = 120000;         // NO_ENCONTRADO sin confirmar se acepta recién tras 2 min (commit tardío)
  const _ED_LEASE_MS  = 60000;          // ANULANDO huérfano (pestaña cerrada a mitad) se retoma tras 60 s
  const _ED_VUELO_MS  = 60000;          // enVueloTs persistido más viejo que esto = envío muerto (pestaña cerrada)
  let _enVueloLocalId = null;           // localId del envasado que ESTA pestaña está enviando ahora mismo
  let _edBusy = false;                  // guard anti-reentrada del procesador
  function _edLeer() {
    try { const o = JSON.parse(localStorage.getItem(ENV_DESHACER_KEY) || '{}'); return (o && typeof o === 'object') ? o : {}; }
    catch (_) { return {}; }
  }
  function _edEscribir(mapa) {
    // Purga: finales >3 días (ya avisados) y cualquier entrada >7 días.
    const ahora = Date.now();
    Object.keys(mapa).forEach(k => {
      const e = mapa[k] || {};
      const edad = ahora - (e.tsDeshacer || e.tsEstado || 0);
      if (edad > 7 * 864e5 || (_ED_FINALES.indexOf(e.estado) >= 0 && !e.avisoPend && edad > 3 * 864e5)) delete mapa[k];
    });
    const json = JSON.stringify(mapa);
    try { localStorage.setItem(ENV_DESHACER_KEY, json); return true; }
    catch (_) {}
    // [2.13.604 · QA H1] Storage lleno (origen compartido con MOS/ME): liberar con los MISMOS tiers que guardar()
    // (caches puros primero, el más viejo primero; trabajo en progreso solo como último recurso) y reintentar.
    try {
      const _porEdad = tier => tier.filter(k => localStorage.getItem(k)).map(k => ({ k, ts: _leerTs(k) })).sort((a, b) => a.ts - b.ts);
      for (const tier of [_STORAGE_TIER1, _STORAGE_TIER2]) {
        for (const { k } of _porEdad(tier)) {
          try { localStorage.removeItem(k); _invalidarParseCache(k); } catch (_) {}
          try { localStorage.setItem(ENV_DESHACER_KEY, json); return true; } catch (_) {}
        }
      }
    } catch (_) {}
    return false;
  }
  // [2.13.604 · QA H1] Devuelve la entrada actualizada, o null si no existe o NO se pudo guardar (storage lleno).
  function _edPatch(clave, patch) {
    const m = _edLeer();
    if (!m[clave]) return null;
    m[clave] = { ...m[clave], ...patch, tsEstado: Date.now() };
    return _edEscribir(m) ? m[clave] : null;
  }
  // Ítem registrarEnvasado de la cola por su clave (lectura FRESCA: otra pestaña pudo cambiar la cola).
  function _edItemCola(clave) {
    _invalidarParseCache(KEYS.QUEUE);
    return getQueue().find(i => i && i.action === 'registrarEnvasado' && i.params && String(i.params.idempotencyKey || '') === String(clave)) || null;
  }
  function _edEnVuelo(item) {
    if (!item) return false;
    if (_enVueloLocalId && _enVueloLocalId === item.localId) return true;
    return !!(item.enVueloTs && (Date.now() - item.enVueloTs) < _ED_VUELO_MS);
  }
  // Quita el ítem de la cola (solo si sigue ahí y no está en vuelo). Síncrono de punta a punta.
  function _edQuitarDeCola(localId) {
    _invalidarParseCache(KEYS.QUEUE);
    const q = getQueue();
    const it = q.find(i => i && i.localId === localId);
    if (!it || _edEnVuelo(it)) return false;
    guardar(KEYS.QUEUE, q.filter(i => i && i.localId !== localId));
    try { _notificar(); } catch (_) {}   // [QA L4] el contador "N por sincronizar" baja al instante (también en la retoma)
    return true;
  }
  function _edIdDe(localId) { return localId ? 'ENV_' + localId : ''; }   // = id que siembra api.js (registrar_envasado)

  // Punto de entrada del ↺ Deshacer (app.js). info: { localId?, idOpt?, tsRegistro?, descripcion?, unidades? }.
  // Devuelve 'QUITADO' | 'EN_VUELO' | 'YA_ENVIADO' | 'YA_REGISTRADO' | 'SIN_ID'.
  function envDeshacerEncolado(clave, info) {
    if (!clave) return 'SIN_ID';
    info = info || {};
    const m = _edLeer();
    if (m[clave]) return 'YA_REGISTRADO';                 // doble tap / otra pestaña: una sola intención por envasado
    const item = _edItemCola(clave);
    const localId = (item && item.localId) || info.localId || '';
    const base = {
      clave: String(clave), localId, idOpt: info.idOpt || '', idEnvasado: _edIdDe(localId),
      descripcion: String(info.descripcion || ''), unidades: info.unidades || 0,
      tsRegistro: info.tsRegistro || (item && item.ts) || Date.now(), tsDeshacer: Date.now(),
      intentos: 0, proximoTs: 0, confirmado: false
    };
    let modo, quitar = false;
    if (item && item.status === 'synced') {
      // Ya enviado (sincronizar() aún no purgó la cola): anular por su id determinista. Sin 'confirmado': un ítem
      // descartado por el servidor también queda 'synced' y ahí NO_ENCONTRADO es lo correcto (→ hecho tras la gracia).
      base.estado = 'POR_ANULAR'; modo = 'YA_ENVIADO';
    } else if (item && !_edEnVuelo(item)) {
      base.estado = 'POR_ANULAR'; base.tsQuitado = Date.now(); modo = 'QUITADO'; quitar = true;
    } else if (item) {
      base.estado = 'ESPERA_COLA'; modo = 'EN_VUELO';
    } else {
      if (!localId) return 'SIN_ID';                      // sin cola ni id determinista: no se adivina nada
      base.estado = 'POR_ANULAR'; modo = 'YA_ENVIADO';
    }
    // [2.13.604 · QA H1] La intención se escribe PRIMERO. Si no se pudo guardar (ni liberando espacio), NO se toca
    // la cola: quitar el ítem sin intención persistida perdería la anulación de un posible commit por timeout.
    m[clave] = { ...base, tsEstado: Date.now() };
    if (!_edEscribir(m)) return 'SIN_ESPACIO';
    if (quitar) {
      // Síncrono: entre la lectura de arriba y esto no corre nada de sincronizar() en esta pestaña.
      _edQuitarDeCola(item.localId);
      // Si la cola en disco aún lo tiene (guardar() no pudo persistir por cuota), esperar como "en vuelo": la retoma
      // lo quita cuando haya espacio, o lo anula cuando la cola lo confirme.
      if (_edItemCola(clave) && _edPatch(clave, { estado: 'ESPERA_COLA', tsQuitado: 0 })) modo = 'EN_VUELO';
    }
    if (navigator.onLine) setTimeout(() => { try { _envDeshacerProcesar(); } catch (_) {} }, 0);
    return modo;
  }

  // Hook de sincronizar() tras enviar un registrarEnvasado. El caller igual lo envuelve en try/catch.
  function _envDeshacerTrasEnvio(item, res) {
    const clave = item && item.params && item.params.idempotencyKey;
    if (!clave) return;
    const e = _edLeer()[clave];
    if (!e || e.estado !== 'ESPERA_COLA') return;
    if (res && res.ok && !res._descartar && !res.offline) {
      // La cola CONFIRMÓ el envasado: ahora sí existe → anularlo por su id real.
      _edPatch(clave, { estado: 'POR_ANULAR', confirmado: true, idEnvasado: (res.data && res.data.idEnvasado) || _edIdDe(item.localId) });
    } else {
      // Falló / descartado: el ítem ya no está en vuelo → sacarlo de la cola para que no se reintente nunca.
      // Queda POR_ANULAR sin confirmar (si un timeout sí commiteó, se anula; si no, NO_ENCONTRADO → hecho).
      _edQuitarDeCola(item.localId);
      _edPatch(clave, { estado: 'POR_ANULAR', tsQuitado: Date.now() });
    }
  }

  // tipo 'peligro' (RECHAZADO/VENCIDO): app.js lo muestra en modal y solo se marca visto con "Entendido" (QA M1).
  // tipo 'info' (reintento en curso): toast; se marca visto al mostrarse.
  function _edAviso(clave, mensaje, tipo) {
    _edPatch(clave, { avisoPend: mensaje, avisoTipo: tipo || 'peligro' });
    try { window.dispatchEvent(new CustomEvent('wh:env-deshacer-aviso', { detail: { clave: String(clave) } })); } catch (_) {}
  }

  // Ejecuta las intenciones pendientes, una a la vez, con lease y backoff. Solo toca la cola para quitar un ítem
  // que dejó de estar en vuelo (retoma tras recarga).
  async function _envDeshacerProcesar() {
    if (_edBusy || !navigator.onLine) return;
    if (typeof API === 'undefined' || !API.anularEnvasadoManual) return;
    _edBusy = true;
    let huboCambio = false;
    try {
      const claves = Object.keys(_edLeer());
      for (const clave of claves) {
        let e = _edLeer()[clave];                           // re-leer: otra pestaña pudo avanzarla
        if (!e || _ED_FINALES.indexOf(e.estado) >= 0) continue;
        const ahora = Date.now();
        if (e.estado === 'ESPERA_COLA') {
          // Retoma tras recarga o pestaña muerta a mitad del envío: si el ítem ya no está en vuelo, decidir.
          const it = _edItemCola(clave);
          if (it && _edEnVuelo(it)) continue;
          if (it) { _edQuitarDeCola(it.localId); e = _edPatch(clave, { estado: 'POR_ANULAR', tsQuitado: ahora }); }
          else    { e = _edPatch(clave, { estado: 'POR_ANULAR' }); }   // ya se envió y purgó: anular por id determinista
          if (!e) continue;
        }
        if (e.estado === 'ANULANDO' && (ahora - (e.tsEstado || 0)) < _ED_LEASE_MS) continue;
        if (e.proximoTs && ahora < e.proximoTs) continue;
        if (!e.idEnvasado) { _edPatch(clave, { estado: 'RECHAZADO' }); continue; }
        // Tope de 12 h solo si el envasado PUDO crearse por la cola (no quitado). [QA M2] Un ítem QUITADO se intenta
        // igual: lo esperable es NO_ENCONTRADO → hecho; y si un timeout sí lo creó, anularlo es lo correcto.
        if (!e.tsQuitado && ahora - (e.tsRegistro || ahora) > _ED_MAX_MS) {
          _edPatch(clave, { estado: 'VENCIDO' });
          _edAviso(clave, '⚠ Un envasado deshecho sin conexión (' + (e.descripcion || e.idEnvasado) + ') tiene más de 12 h: no se anuló automáticamente. Revísalo en el historial y anúlalo a mano si aparece.');
          huboCambio = true;
          continue;
        }
        if (!_edPatch(clave, { estado: 'ANULANDO' })) break;   // [QA H1] sin espacio para el lease: no anular a ciegas
        let res = null;
        try {
          // MISMA operación del Deshacer en línea. _fromQueue: un timeout NO se re-encola en wh_queue (este
          // procesador ya reintenta). _detalleError: devolver el código del servidor en vez del genérico cero-GAS.
          res = await API.anularEnvasadoManual({
            idEnvasado: e.idEnvasado,
            usuario:    (window.WH_CONFIG && window.WH_CONFIG.usuario) || 'manual',
            motivo:     'deshacer de envasado registrado sin conexión',
            _fromQueue: true, _detalleError: true
          });
        } catch (err) { res = { ok: false, error: (err && err.message) || 'error', _retry: true }; }
        const t = Date.now();
        if (res && res.ok) {
          _edPatch(clave, { estado: 'HECHO', yaAnulado: !!(res.data && res.data.yaAnulado), avisoPend: '' });
          huboCambio = true;
          continue;
        }
        // [QA M3] APP_NO_AUTORIZADA (token/claim del equipo) es transitoria → cae al backoff de abajo.
        if (res && res._rechazoServidor && res.error !== 'APP_NO_AUTORIZADA') {
          if (res.error === 'ENVASADO_NO_ENCONTRADO' && !e.confirmado) {
            // Nunca llegó al servidor. Se acepta pasada la ventana de gracia (commit tardío de un timeout).
            const desde = e.tsQuitado || e.tsDeshacer || 0;
            if (t - desde >= _ED_GRACIA_MS) { _edPatch(clave, { estado: 'HECHO', sinRegistro: true, avisoPend: '' }); huboCambio = true; }
            else _edPatch(clave, { estado: 'POR_ANULAR', proximoTs: desde + _ED_GRACIA_MS });
            continue;
          }
          _edPatch(clave, { estado: 'RECHAZADO', ultimoError: String(res.error || '') });
          _edAviso(clave, '⚠ No se pudo deshacer en el servidor el envasado ' + (e.descripcion || e.idEnvasado) + ' (' + (res.error || 'rechazado') + '). Revísalo en el historial y anúlalo a mano.');
          huboCambio = true;
          continue;
        }
        // Falla transitoria: red/timeout/servicio, APP_NO_AUTORIZADA, o HTTP 401/403 que api.js entrega como
        // 'rechazo-directo' [QA L1]. Reintentable con backoff 20 s → 5 min y aviso informativo una sola vez.
        // Si sigue fallando más de 12 h desde el PRIMER fallo → VENCIDO con aviso de peligro (no se insiste más).
        const tsPrimerFallo = e.tsPrimerFallo || t;
        const ultimoError = String((res && res.error) || '');
        if (t - tsPrimerFallo > _ED_MAX_MS) {
          _edPatch(clave, { estado: 'VENCIDO', ultimoError });
          _edAviso(clave, '⚠ No se pudo anular en el servidor el envasado deshecho (' + (e.descripcion || e.idEnvasado) + ') tras 12 h de reintentos (' + (ultimoError || 'sin respuesta') + '). Revísalo en el historial y anúlalo a mano.');
          huboCambio = true;
          continue;
        }
        const intentos = (e.intentos || 0) + 1;
        const espera = Math.min(300000, 20000 * Math.pow(2, Math.min(intentos - 1, 4)));
        if (!_edPatch(clave, { estado: 'POR_ANULAR', intentos, tsPrimerFallo, proximoTs: t + espera, ultimoError })) break;
        if (!e.avisoError) {
          _edPatch(clave, { avisoError: true });
          _edAviso(clave, '⚠ Aún no se pudo anular en el servidor el envasado deshecho (' + (e.descripcion || e.idEnvasado) + '). Se reintentará solo; no lo registres de nuevo.', 'info');
        }
      }
    } catch (err) {
      try { console.warn('[envDeshacer] procesar:', err && err.message); } catch (_) {}
    } finally {
      _edBusy = false;
    }
    if (huboCambio) {
      try { window.dispatchEvent(new CustomEvent('wh:envasado-reconciliado')); } catch (_) {}
      try { const p = precargarOperacional(true); if (p && p.catch) p.catch(() => {}); } catch (_) {}
      try { window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['envasados'] } })); } catch (_) {}
    }
  }

  // Ids reales que la UI oculta: deshechos (pendientes o hechos). RECHAZADO/VENCIDO se muestran para que el
  // operador pueda anularlos a mano.
  function envDeshacerIdsOcultos() {
    const out = new Set();
    try {
      const m = _edLeer();
      Object.keys(m).forEach(k => {
        const e = m[k];
        if (e && e.idEnvasado && e.estado !== 'RECHAZADO' && e.estado !== 'VENCIDO') out.add(String(e.idEnvasado));
      });
    } catch (_) {}
    return out;
  }
  // [QA M1] Avisos pendientes SIN marcarlos: app.js marca cada uno con envDeshacerMarcarAvisado() recién cuando el
  // operador lo vio (peligro: tocó "Entendido"; info: se mostró el toast).
  function envDeshacerAvisosPendientes() {
    const m = _edLeer();
    return Object.keys(m).filter(k => m[k] && m[k].avisoPend)
      .map(k => ({ clave: k, mensaje: String(m[k].avisoPend), tipo: m[k].avisoTipo || 'peligro' }));
  }
  // Solo limpia si el aviso sigue siendo el mismo (no borra uno más nuevo de la misma clave).
  function envDeshacerMarcarAvisado(clave, mensaje) {
    const m = _edLeer();
    if (!m[clave] || String(m[clave].avisoPend || '') !== String(mensaje || '')) return false;
    m[clave] = { ...m[clave], avisoPend: '', avisoTipo: '' };
    return _edEscribir(m);
  }
  function envDeshacerGet(clave) { return clave ? (_edLeer()[clave] || null) : null; }

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
    try {
    for (const item of queue) {
      let res, enviado = false;   // [2.13.604] para el hook de deshacer (finally)
      try {
        // [2.13.604] La cola se tomó como FOTO al empezar: un ↺ Deshacer pudo quitar este envasado mientras se
        // drenaban los anteriores. Re-chequeo síncrono contra la cola VIVA justo antes de enviarlo.
        if (item.action === 'registrarEnvasado') {
          let _sigue = true;
          try { _invalidarParseCache(KEYS.QUEUE); _sigue = getQueue().some(i => i && i.localId === item.localId); } catch (_) {}
          if (!_sigue) continue;
        }
        const viaDirecta = !!(item._viaDirecta || (item.params && item.params._viaDirecta));
        if (viaDirecta) {
          // Ítem que pudo commitear en Supabase → SIEMPRE reintento directo (dedup), nunca GAS.
          if (typeof API === 'undefined' || !API._postCola) {
            // Módulo de escritura directa no disponible: no podemos reintentar directo y mandarlo a
            // GAS duplicaría. Lo dejamos pendiente (marcado 'error') para reintentar cuando cargue.
            _actualizarItemQueue(item.localId, 'error');
            continue;
          }
          // [2.13.604] Marca EN VUELO (memoria + persistida en el ítem, para otras pestañas): el Deshacer no lo
          // quita de la cola mientras viaja; espera la confirmación y lo anula por su id real.
          if (item.action === 'registrarEnvasado') {
            _enVueloLocalId = item.localId;
            try { guardar(KEYS.QUEUE, getQueue().map(i => i.localId === item.localId ? { ...i, enVueloTs: Date.now() } : i)); } catch (_) {}
          }
          enviado = true;
          res = await API._postCola(item.params);
        } else {
          // [CERO-GAS Rep#1] Ítem legacy SIN sello _viaDirecta. En prod NO existe (todo se sella al encolar bajo
          // escritura directa, que está permanentemente ON). Antes se replayaba a GAS (window.WH_CONFIG.gasUrl);
          // eso era el último rastro de GAS de la cola. Ahora fail-closed: NO va a GAS (evita reejecutar/duplicar).
          // Queda visible como 'error' — si algún día apareciera uno, se ve en la cola en vez de duplicarse en GAS.
          res = { ok: false, error: 'cero-GAS: ítem sin vía directa (no se replaya a GAS)', _ceroGas: true };
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
        if (res && res.ok && item.action === 'registrarEnvasado') {
          huboEnvasado = true;
          // [2.13.603] El envasado (clave ENV-…) quedó confirmado en el servidor: si esperaba adhesivos, se marca
          // SINCRONIZADO AQUÍ MISMO (persistido, no solo un evento) para que una recarga entre el sync y la
          // impresión no lo pierda: WhLoteAuto barre los SINCRONIZADO al iniciar sesión. Solo PENDIENTE avanza.
          const _claveEnv = item.params && item.params.idempotencyKey;
          if (_claveEnv && loteAutoSet(_claveEnv, { estado: 'SINCRONIZADO', idEnvasado: (res.data && res.data.idEnvasado) || '' }, { soloSi: ['PENDIENTE'] })) {
            try { window.dispatchEvent(new CustomEvent('wh:envasado-sincronizado', { detail: { clave: String(_claveEnv) } })); } catch (_) {}
          }
        }
        // [FIX Rep#1 · auto-print] Al drenar una GUÍA creada por red lenta, disparar la impresión del ticket con el
        // idGuia REAL. Sin esto la guía se sincronizaba en Supabase pero el ticket nunca salía → el operador debía
        // "imprimir copia" a mano. El dedup atómico wh.reservar_ticket evita cualquier doble ticket (idempotente).
        if (res && res.ok && res.data && res.data.idGuia &&
            (item.action === 'crearDespachoRapido' || item.action === 'cerrarPickupConDespacho')) {
          try { window.dispatchEvent(new CustomEvent('wh:guia-sincronizada', { detail: { idGuia: String(res.data.idGuia), action: item.action } })); } catch (_) {}
        }
      } catch {
        _actualizarItemQueue(item.localId, 'error');
      } finally {
        // [2.13.604] Fin del vuelo + ¿este envasado tenía un ↺ Deshacer esperando? Envuelto: jamás traba la cola.
        if (item.action === 'registrarEnvasado') {
          if (_enVueloLocalId === item.localId) _enVueloLocalId = null;
          if (enviado) { try { _envDeshacerTrasEnvio(item, res); } catch (_) {} }
        }
      }
    }

    limpiarSincronizados();
    localStorage.setItem(KEYS.LAST_SYNC, new Date().toLocaleTimeString('es-PE'));
    } finally {
      // [F1 · fix revisión 500x] _syncing SIEMPRE se resetea, aunque el loop / limpiarSincronizados() / setItem
      // lancen (ej. QuotaExceededError con la cola congestionada). Sin esto _syncing quedaba pegado en true y el
      // drainer periódico (20s) + el sync por evento 'online' morían hasta un reload completo — justo el bug
      // que el drainer venía a resolver.
      _syncing = false;
    }
    _notificar();

    // [Fix v2.9.1] Si se sincronizó un envasado optimistic, avisar a la UI
    // para que recargue desde el backend y reemplace los ENV_OPT_* por reales.
    if (huboEnvasado) {
      try { window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['envasados'] } })); } catch(_){}
    }
    // [2.13.604] Anular los envasados deshechos que la cola acaba de confirmar (fuera del loop: no frena la cola).
    try { _envDeshacerProcesar(); } catch (_) {}
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
      // [CERO-GAS 2026-07-19] SOLO API (Supabase directo); el fetch crudo a GAS se eliminó.
      const r = (typeof API !== 'undefined' && API.descargarOperacional)
        ? await API.descargarOperacional().catch(() => null)
        : null;
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
          // [v2.13.376] 'fotos' QUITADO del field-protect: causaba que un borrado de foto
          // hecho en OTRO dispositivo se revirtiera por ≤15s. Ya no hace falta protegerlo acá
          // porque (a) durante la subida el refresh se SALTA por _subiendoFotos, y (b) el
          // incremental persist deja las fotos en el backend al terminar → el poll las trae.
          // El preserve por-vacío (arriba, líneas 731-734) sigue cubriendo el caso legítimo.
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
      // [534 perf/memoria] Si NO cambió, NO se reescribe.
      // Antes el `else guardar(key, arr)` reescribía igual: cada ciclo de 60s hacía, por
      // cada uno de los 6 datasets, JSON.stringify + LZString.compressToUTF16 + setItem
      // SINCRÓNICOS en el hilo principal, y además invalidaba el _parseCache, forzando
      // descomprimir+parsear otra vez en la siguiente lectura. 720 ciclos por turno de 12h
      // sobre datasets que crecen todo el día = el "se va poniendo lenta" de las tablets.
      const _persist = (key, arr, label) => {
        if (arr == null) return;
        if (!arr.length && (cargar(key) || []).length) return;   // vacío sobre lleno → skip
        if (_hayDiff(arr, key)) { guardar(key, arr); changed.push(label); }
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
            // [534] idem _persist: sin cambios → sin reescribir (ver comentario arriba).
            if (_hayDiff(merged, KEYS.PREINGRESOS)) { guardar(KEYS.PREINGRESOS, merged); changed.push('preingresos'); }
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
    // [perf 500x] El timer de 60s YA NO re-baja el catálogo completo (1.9MB/min = el descargarMaestros
    // repetido). Es redundante: el poller de versión (50s) + realtime detectan cambios del catálogo y bajan
    // SOLO el delta (~42KB) cuando la versión sube. Aquí solo refrescamos lo operacional (guías/stock).
    _opRefreshTimer = setInterval(() => {
      if (document.hidden) return;     // [perf] no precargar con la pestaña oculta (el handler visible refresca al volver)
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
      // [CD1] refresco INCREMENTAL (delta); cae a full si no hay corte. [500x HIGH] el baseline avanza SOLO
      // si la descarga tuvo éxito → si falla (red/RPC), NO se marca la versión como consumida y se reintenta.
      let r = null;
      try { r = await _refrescarCatalogoDelta(); } catch (_) { r = { ok: false }; }
      if (r && r.ok !== false) {
        _setBaselineCatalogo(target);
        if (typeof toast === 'function') toast('Catálogo actualizado', 'info', 2200);
      } else {
        console.warn('[Offline] re-descarga de catálogo falló → baseline intacto, se reintenta en el próximo ciclo');
      }
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
  // [G4 online-only] Se eliminó el caché de PINs admin (getAdminCache + sincronizarAdminCache): la verificación
  // de clave admin es siempre online (mos.verificar_clave_admin). Ya no se baja ni almacena material de PIN.
  // [CERO-GAS · audit 2026-07-13] getAdminPin ELIMINADO (sin consumidor; era el PIN plano de Sheet).
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

  // ── Patch de un producto del catálogo en caché ──────────────
  // [v2.13.540] Al cambiar la foto desde WH, el cambio se ve YA (sin esperar el próximo
  // delta del catálogo). El delta posterior trae el mismo fotoUrl → converge, no pisa nada.
  function patchProductoCache(idProducto, changes) {
    const cache = cargar(KEYS.PRODUCTOS) || [];
    const idx   = cache.findIndex(p => String(p.idProducto) === String(idProducto));
    if (idx >= 0) { Object.assign(cache[idx], changes); guardar(KEYS.PRODUCTOS, cache); }
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
    actualizarDetallesGuia, addDetalleCache, inyectarPreingreso, patchPreingresosCache, patchStockCache,
    patchGuiaCache, patchProductoCache,
    getPNCache, setPNCache,
    getEnvasadosCache, guardarEnvasadosCache, inyectarEnvasadoCache, removerEnvasadoCache,
    precargarOperacional, iniciarRefreshOperacional, detenerRefreshOperacional,
    iniciarPollerCatalogo, detenerPollerCatalogo, notificarVersionCatalogo,
    setSubiendoFotos, isSubiendoFotos,
    patchPendingDetalleVenc,
    loteAutoRegistrar, loteAutoGet, loteAutoListar, loteAutoSet, loteAutoDeshacerPorOpt, loteAutoDeshacerPorId,   // [2.13.603]
    envDeshacerEncolado, envDeshacerIdsOcultos, envDeshacerAvisosPendientes, envDeshacerMarcarAvisado, envDeshacerGet,   // [2.13.604]
    envDeshacerProcesar: () => _envDeshacerProcesar(),   // [2.13.604] barrido al iniciar sesión
    marcarPreingresoPendiente, marcarCargadoresPendiente,
    estaOnline: () => navigator.onLine
  };
})();
