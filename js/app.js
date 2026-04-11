// warehouseMos — app.js  Lógica de la aplicación
'use strict';

// ════════════════════════════════════════════════
// Helpers UI globales
// ════════════════════════════════════════════════
function toast(msg, tipo = 'info', dur = 3000) {
  const el = document.getElementById('toast');
  const colors = { ok: '#166534|#86efac', danger: '#7f1d1d|#fca5a5', info: '#1e3a8a|#93c5fd', warn: '#854d0e|#fde68a' };
  const [bg, fg] = (colors[tipo] || colors.info).split('|');
  el.style.background = bg;
  el.style.color = fg;
  el.textContent = msg;
  el.style.opacity = '1';
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.opacity = '0'; }, dur);
}

function loading(parentId, show) {
  const el = document.getElementById(parentId);
  if (!el) return;
  if (show) {
    el.innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';
  }
}

function abrirSheet(id) {
  const overlay = document.getElementById('overlay' + id.replace('sheet', ''));
  const sheet = document.getElementById(id);
  overlay?.classList.add('open');
  sheet?.classList.add('open');
}

function cerrarSheet(id) {
  const overlay = document.getElementById('overlay' + id.replace('sheet', ''));
  const sheet = document.getElementById(id);
  overlay?.classList.remove('open');
  sheet?.classList.remove('open');
}

function abrirModal(id) {
  document.getElementById(id)?.classList.add('open');
}

function cerrarModal(id) {
  document.getElementById(id)?.classList.remove('open');
}

let _scannerCallback = null;
function abrirScanner(onResult) {
  document.getElementById('scannerModal').classList.add('open');
  Scanner.start('scanVideo', code => {
    cerrarScanner();
    if (_scannerCallback) _scannerCallback(code);
    else if (onResult) onResult(code);
  }, err => { toast('Error cámara: ' + err, 'danger'); cerrarScanner(); });
}

function abrirScannerPara(inputId, callback) {
  _scannerCallback = callback || (code => {
    const el = document.getElementById(inputId);
    if (el) { el.value = code; el.dispatchEvent(new Event('input')); }
    _scannerCallback = null;
  });
  abrirScanner();
}

function cerrarScanner() {
  Scanner.stop();
  document.getElementById('scannerModal').classList.remove('open');
}

function fmt(n, dec = 0) {
  return (parseFloat(n) || 0).toFixed(dec).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function fmtFecha(s) {
  if (!s) return '—';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return d.toLocaleDateString('es-PE', { day:'2-digit', month:'short', year:'numeric' });
}

function diasColor(dias) {
  if (dias <= 7)  return 'tag-danger';
  if (dias <= 30) return 'tag-warn';
  return 'tag-ok';
}

// ════════════════════════════════════════════════
// SESSION — Login, bloqueo, cierre de turno
// ════════════════════════════════════════════════
const Session = (() => {
  let pinBuffer = '';
  let lockPinBuffer = '';
  let operadorSeleccionado = null;
  let sesionActual = null;
  let lockTimer = null;
  let lockInterval = null;
  let cierreReporte = null;
  const MIN_INACTIVIDAD = 5; // minutos (se sobreescribe desde config)

  async function init() {
    // Verificar sesión guardada
    const saved = _cargarSesion();
    if (saved) {
      const res = await API.getSesionActiva(saved.idSesion).catch(() => ({ ok: false }));
      if (res.ok) {
        sesionActual = saved;
        _aplicarSesion();
        return;
      } else {
        _limpiarSesion();
      }
    }
    // Sin sesión → mostrar login
    await mostrarLogin();
  }

  async function mostrarLogin() {
    _ocultarApp();
    document.getElementById('loginScreen').style.display = 'flex';
    document.getElementById('loginPaso1').classList.remove('hidden');
    document.getElementById('loginPaso2').classList.add('hidden');

    const container = document.getElementById('loginOperadores');

    // Cargar personal CON pins para validación local + sin pins para mostrar botones
    container.innerHTML = '<div class="flex justify-center py-6"><div class="spinner"></div></div>';

    // Siempre intentar red para tener PINs frescos en caché
    if (navigator.onLine) {
      const GAS = window.WH_CONFIG.gasUrl;
      if (GAS) {
        fetch(`${GAS}?action=getPersonalConPin`)
          .then(r => r.json())
          .then(res => { if (res.ok) OfflineManager._guardarPersonalConPin(res.data); })
          .catch(() => {});
      }
    }

    let personal = OfflineManager.getPersonalCache().map(p => { const s={...p}; delete s.pin; return s; });
    if (!personal.length) {
      const res = await API.getPersonal().catch(() => ({ ok: false }));
      personal = res.ok ? res.data : [];
    }

    if (!personal.length) {
      container.innerHTML = `
        <div class="text-center py-4 space-y-4">
          <p class="text-red-400 text-sm">No se pudo cargar el personal.</p>
          <p class="text-slate-500 text-xs">Verifica que el Web App de GAS esté desplegado<br>y que la tabla PERSONAL tenga datos.</p>
          <button onclick="Session.mostrarLogin()" class="btn btn-outline w-full py-3">🔄 Reintentar</button>
        </div>`;
      return;
    }

    container.innerHTML = personal.map(p => `
      <button onclick="Session.seleccionarOperador('${p.idPersonal}','${p.nombre}','${p.apellido}','${p.rol}','${p.color}')"
              class="w-full flex items-center gap-4 p-4 rounded-2xl border-2 border-slate-700 active:border-blue-500"
              style="background:#1e293b;">
        <div class="w-14 h-14 rounded-full flex items-center justify-center text-white font-black text-xl flex-shrink-0"
             style="background:${p.color}">${p.nombre[0]}${p.apellido[0]}</div>
        <div class="text-left">
          <p class="text-white font-bold text-lg">${p.nombre} ${p.apellido}</p>
          <p class="text-slate-400 text-sm">${p.rol}</p>
        </div>
        <span class="ml-auto text-slate-500 text-2xl">›</span>
      </button>`).join('');
  }

  function seleccionarOperador(id, nombre, apellido, rol, color) {
    operadorSeleccionado = { idPersonal: id, nombre, apellido, rol, color };
    pinBuffer = '';
    _actualizarPuntos('pin', 0);
    document.getElementById('loginNombreText').textContent = nombre + ' ' + apellido;
    document.getElementById('loginRolText').textContent = rol;
    const av = document.getElementById('loginAvatar');
    av.textContent = nombre[0] + apellido[0];
    av.style.background = color;
    document.getElementById('loginError').textContent = '';
    document.getElementById('loginPaso1').classList.add('hidden');
    document.getElementById('loginPaso2').classList.remove('hidden');
  }

  function pinTecla(d) {
    if (pinBuffer.length >= 4) return;
    pinBuffer += d;
    _actualizarPuntos('pin', pinBuffer.length);
    if (pinBuffer.length === 4) setTimeout(() => _intentarLogin(), 150);
  }

  function pinAtras(volverPaso1 = false) {
    if (volverPaso1) {
      pinBuffer = '';
      document.getElementById('loginPaso1').classList.remove('hidden');
      document.getElementById('loginPaso2').classList.add('hidden');
      return;
    }
    if (pinBuffer.length > 0) {
      pinBuffer = pinBuffer.slice(0, -1);
      _actualizarPuntos('pin', pinBuffer.length);
    }
  }

  async function _intentarLogin() {
    const pinIntento = pinBuffer;
    pinBuffer = '';
    _actualizarPuntos('pin', 0);

    // 1. Validar PIN localmente (instantáneo si hay caché)
    let localOp = OfflineManager.validarPinLocal(pinIntento);

    // Sin caché → validar directo en GAS
    if (!localOp && navigator.onLine) {
      const res = await API.loginPersonal(pinIntento);
      if (!res.ok || res.offline) {
        document.getElementById('loginError').textContent = '❌ PIN incorrecto';
        setTimeout(() => { document.getElementById('loginError').textContent = ''; }, 2000);
        return;
      }
      // Login exitoso vía GAS
      sesionActual = res.data;
      _guardarSesion(sesionActual);
      document.getElementById('loginScreen').style.display = 'none';
      _aplicarSesion();
      toast(`¡Hola ${sesionActual.nombre}! 👋`, 'ok', 2500);
      return;
    }

    if (!localOp) {
      document.getElementById('loginError').textContent = '❌ PIN incorrecto';
      setTimeout(() => { document.getElementById('loginError').textContent = ''; }, 2000);
      return;
    }

    // 2. Sesión optimista inmediata
    sesionActual = {
      idSesion:   'LOCAL_' + Date.now(),
      idPersonal: localOp.idPersonal,
      nombre:     localOp.nombre,
      apellido:   localOp.apellido,
      rol:        localOp.rol,
      color:      localOp.color,
      horaInicio: new Date().toLocaleTimeString('es-PE')
    };
    _guardarSesion(sesionActual);
    document.getElementById('loginScreen').style.display = 'none';
    _aplicarSesion();
    toast(`¡Hola ${sesionActual.nombre}! 👋`, 'ok', 2500);

    // 3. Confirmar sesión con GAS en segundo plano
    API.loginPersonal(pinIntento).then(res => {
      if (res.ok && !res.offline) {
        sesionActual = { ...sesionActual, ...res.data };
        window.WH_CONFIG.idSesion = sesionActual.idSesion;
        _guardarSesion(sesionActual);
      }
    }).catch(() => {});
  }

  function _aplicarSesion() {
    window.WH_CONFIG.usuario   = sesionActual.nombre + ' ' + sesionActual.apellido;
    window.WH_CONFIG.idSesion  = sesionActual.idSesion;

    // Avatar header
    const av = document.getElementById('topAvatar');
    av.textContent   = sesionActual.nombre[0] + sesionActual.apellido[0];
    av.style.background = sesionActual.color;
    document.getElementById('usuarioNombre').textContent = sesionActual.nombre;

    _mostrarApp();
    _iniciarTimerBloqueo();

    // Conectar indicador de estado online/offline/sync
    OfflineManager.onStatusChange(_actualizarEstadoHeader);
    _actualizarEstadoHeader({
      online:  navigator.onLine,
      pending: OfflineManager.getQueue().length,
      syncing: false
    });

    // Precargar datos en background
    OfflineManager.precargar().then(() => {
      App.cargarDashboard();
      App.cargarProductosMaestro();
      App.cargarProveedoresMaestro();
    });

    // Si hay cola pendiente y hay red, sincronizar
    if (navigator.onLine) OfflineManager.sincronizar();
  }

  function _actualizarEstadoHeader({ online, pending, syncing }) {
    const dot    = document.getElementById('statusDot');
    const dotM   = document.getElementById('statusDotMobile');
    const lbl    = document.getElementById('statusLabel');

    let color, texto;
    if (syncing) {
      color = '#f59e0b'; texto = 'Sincronizando...';
    } else if (!online) {
      color = '#ef4444'; texto = pending > 0 ? `${pending} pendientes` : 'Sin conexión';
    } else if (pending > 0) {
      color = '#f59e0b'; texto = `${pending} por sync`;
    } else {
      color = '#22c55e'; texto = 'En línea';
    }

    if (dot)  { dot.style.background  = color; }
    if (dotM) { dotM.style.background = color; }
    if (lbl)  { lbl.textContent = texto; }
  }

  // ── Bloqueo por inactividad ────────────────────────────────
  function _iniciarTimerBloqueo() {
    _resetTimerBloqueo();
    ['touchstart','click','keydown','scroll'].forEach(ev =>
      document.addEventListener(ev, _resetTimerBloqueo, { passive: true })
    );
  }

  function _resetTimerBloqueo() {
    clearTimeout(lockTimer);
    const min = parseInt(localStorage.getItem('wh_min_inactividad')) || MIN_INACTIVIDAD;
    lockTimer = setTimeout(() => bloquear(), min * 60 * 1000);
  }

  function bloquear() {
    if (!sesionActual) return;
    lockPinBuffer = '';
    _actualizarPuntos('lpin', 0);
    document.getElementById('lockError').textContent = '';

    const av = document.getElementById('lockAvatar');
    av.textContent = sesionActual.nombre[0] + sesionActual.apellido[0];
    av.style.background = sesionActual.color;
    document.getElementById('lockNombre').textContent = sesionActual.nombre + ' ' + sesionActual.apellido;

    const inicio = parseInt(localStorage.getItem('wh_lock_inicio') || Date.now());
    localStorage.setItem('wh_lock_inicio', Date.now());

    document.getElementById('lockScreen').style.display = 'flex';

    clearInterval(lockInterval);
    lockInterval = setInterval(() => {
      const seg = Math.floor((Date.now() - parseInt(localStorage.getItem('wh_lock_inicio'))) / 1000);
      const m = Math.floor(seg / 60), s = seg % 60;
      document.getElementById('lockTiempo').textContent =
        `Bloqueado hace ${m > 0 ? m + 'm ' : ''}${s}s`;
    }, 1000);
  }

  function lockTecla(d) {
    if (lockPinBuffer.length >= 4) return;
    lockPinBuffer += d;
    _actualizarPuntos('lpin', lockPinBuffer.length);
    if (lockPinBuffer.length === 4) setTimeout(() => _intentarDesbloqueo(), 150);
  }

  function lockAtras() {
    if (lockPinBuffer.length > 0) {
      lockPinBuffer = lockPinBuffer.slice(0, -1);
      _actualizarPuntos('lpin', lockPinBuffer.length);
    }
  }

  async function _intentarDesbloqueo() {
    // Verificar contra sesión activa
    const res = await API.loginPersonal(lockPinBuffer);
    lockPinBuffer = '';
    _actualizarPuntos('lpin', 0);
    if (res.ok && res.data.idPersonal === sesionActual.idPersonal) {
      clearInterval(lockInterval);
      document.getElementById('lockScreen').style.display = 'none';
      _resetTimerBloqueo();
    } else {
      document.getElementById('lockError').textContent = '❌ PIN incorrecto';
      setTimeout(() => { document.getElementById('lockError').textContent = ''; }, 2000);
    }
  }

  // ── Cierre de turno ────────────────────────────────────────
  function confirmarCierre() {
    // Mostrar reporte preliminar antes de confirmar
    _mostrarReportePreliminar();
  }

  async function _mostrarReportePreliminar() {
    const overlay = document.getElementById('reporteTurnoOverlay');
    overlay.style.display = 'flex';
    overlay.style.flexDirection = 'column';

    // Calcular tiempo transcurrido localmente
    const sesGuardada = _cargarSesion();
    const inicioSes = new Date(sesGuardada?.fechaGuardado || Date.now());
    const minutos = Math.round((Date.now() - inicioSes) / 60000);
    const horas = (minutos / 60).toFixed(1);

    // Traer desempeño actual del GAS
    const res = await API.getDesempenoDia({ idPersonal: sesionActual.idPersonal })
                         .catch(() => ({ ok: false }));
    const des = (res.ok && res.data.length) ? res.data[res.data.length - 1] : {};

    cierreReporte = { minutos, horas };

    // Llenar reporte
    const av = document.getElementById('reporteAvatar');
    av.textContent = sesionActual.nombre[0] + sesionActual.apellido[0];
    av.style.background = sesionActual.color;
    document.getElementById('reporteNombre').textContent = sesionActual.nombre + ' ' + sesionActual.apellido;
    document.getElementById('reporteRol').textContent = sesionActual.rol;
    document.getElementById('reporteFecha').textContent = new Date().toLocaleDateString('es-PE', { weekday:'long', day:'numeric', month:'long' });

    const total = parseInt(des.totalActividades) || 0;
    const actPH = horas > 0 ? (total / parseFloat(horas)).toFixed(1) : 0;
    const punt  = Math.min(10, parseFloat(actPH)).toFixed(1);
    const calif = punt >= 9 ? ['EXCELENTE','text-emerald-400','🏆']
                : punt >= 7 ? ['BUENO','text-blue-400','⭐']
                : punt >= 5 ? ['REGULAR','text-amber-400','👍']
                : ['BAJO','text-red-400','📉'];

    document.getElementById('reportePuntos').textContent = punt + '/10';
    document.getElementById('reportePuntos').className = 'text-6xl font-black mb-1 ' + calif[1];
    document.getElementById('reporteCalifTexto').textContent = calif[2] + ' ' + calif[0];
    document.getElementById('reporteCalifTexto').className = 'text-xl font-bold mb-1 ' + calif[1];
    document.getElementById('reporteHoras').textContent = `${horas}h trabajadas · ${total} actividades · ${actPH}/h`;

    document.getElementById('rGuias').textContent     = des.guiasCreadas || 0;
    document.getElementById('rEnvasados').textContent = des.envasadosRegistrados || 0;
    document.getElementById('rUnidades').textContent  = des.unidadesEnvasadas || 0;
    document.getElementById('rMermas').textContent    = des.mermasRegistradas || 0;
    document.getElementById('rAuditorias').textContent= des.auditoriaEjecutadas || 0;
    document.getElementById('rTotal').textContent     = total;

    const montoBase = parseFloat(des.montoBase) || 0;
    const bonusMin  = 8;
    const bonusPct  = 10;
    const bonus     = parseFloat(punt) >= bonusMin ? Math.round(montoBase * bonusPct / 100 * 100) / 100 : 0;
    const montoTot  = montoBase + bonus;

    document.getElementById('rMontoBase').textContent  = 'S/. ' + fmt(montoBase, 2);
    document.getElementById('rBonus').textContent      = bonus > 0 ? '+S/. ' + fmt(bonus, 2) : 'S/. 0.00';
    document.getElementById('rMontoTotal').textContent = 'S/. ' + fmt(montoTot, 2);
  }

  async function cerrarTurnoFinal() {
    const res = await API.cerrarTurno({ idSesion: sesionActual.idSesion }).catch(() => ({ ok: false }));
    _limpiarSesion();
    sesionActual = null;
    clearTimeout(lockTimer);
    clearInterval(lockInterval);

    document.getElementById('reporteTurnoOverlay').style.display = 'none';
    _ocultarApp();
    toast('Turno cerrado. ¡Hasta mañana! 👋', 'ok', 3000);
    setTimeout(() => mostrarLogin(), 2000);
  }

  // ── Cierre forzado al final del día ───────────────────────
  function _verificarCierreForzado() {
    if (!sesionActual) return;
    const horaConfig = localStorage.getItem('wh_hora_cierre') || '22:00';
    const [hh, mm] = horaConfig.split(':').map(Number);
    const ahora = new Date();
    if (ahora.getHours() === hh && ahora.getMinutes() === mm) {
      toast('⏰ Fin de turno — cerrando automáticamente', 'warn', 5000);
      setTimeout(() => confirmarCierre(), 5000);
    }
  }

  // ── Helpers ────────────────────────────────────────────────
  function _actualizarPuntos(prefix, n) {
    for (let i = 0; i < 4; i++) {
      const el = document.getElementById(prefix + i);
      if (el) {
        el.className = i < n
          ? 'w-4 h-4 rounded-full bg-blue-500'
          : 'w-4 h-4 rounded-full border-2 border-slate-600';
      }
    }
  }

  function _guardarSesion(ses) {
    localStorage.setItem('wh_sesion', JSON.stringify({ ...ses, fechaGuardado: new Date().toISOString() }));
  }

  function _cargarSesion() {
    try { return JSON.parse(localStorage.getItem('wh_sesion')); }
    catch { return null; }
  }

  function _limpiarSesion() { localStorage.removeItem('wh_sesion'); }

  function _mostrarApp() {
    document.getElementById('topBar').style.display = '';
    document.querySelector('main').style.display = '';
    document.querySelector('nav').style.display = '';
  }

  function _ocultarApp() {
    document.getElementById('topBar').style.display = 'none';
    document.querySelector('main').style.display = 'none';
    document.querySelector('nav').style.display = 'none';
  }

  function getSesion() { return sesionActual; }

  // Verificar cierre forzado cada minuto
  setInterval(_verificarCierreForzado, 60000);

  return { init, mostrarLogin, seleccionarOperador,
           pinTecla, pinAtras, lockTecla, lockAtras,
           bloquear, confirmarCierre, cerrarTurnoFinal,
           getSesion };
})();

// ════════════════════════════════════════════════
// DASHBOARD — paneles expandibles
// ════════════════════════════════════════════════
const Dashboard = (() => {
  let panelActivo = null;
  const panelMap = { venc: 'panelVenc', env: 'panelEnv', stock: 'panelStock', mermas: 'panelMermas' };

  function toggle(key) {
    const id = panelMap[key];
    if (!id) return;
    if (panelActivo === key) {
      document.getElementById(id)?.classList.add('hidden');
      panelActivo = null;
      return;
    }
    if (panelActivo) document.getElementById(panelMap[panelActivo])?.classList.add('hidden');
    document.getElementById(id)?.classList.remove('hidden');
    panelActivo = key;
    setTimeout(() => document.getElementById(id)?.scrollIntoView({ behavior: 'smooth', block: 'nearest' }), 50);
  }

  return { toggle };
})();

// ════════════════════════════════════════════════
// App principal — navegación y estado global
// ════════════════════════════════════════════════
const App = (() => {
  let currentView = 'dashboard';
  let modoEnvasador = false;
  let dashboardData = null;
  let todosProductos = [];
  let todosProveedores = [];

  function init() {
    // Restaurar GAS URL si fue guardada localmente
    const gasUrl = localStorage.getItem('wh_gas_url');
    if (gasUrl) window.WH_CONFIG.gasUrl = gasUrl;

    // Ocultar app hasta login
    document.getElementById('topBar').style.display = 'none';
    document.querySelector('main').style.display = 'none';
    document.querySelector('nav').style.display = 'none';

    // Tipo guía → mostrar/ocultar zona
    document.getElementById('guiaTipo')?.addEventListener('change', e => {
      const isZona = e.target.value === 'SALIDA_ZONA';
      document.getElementById('guiaZonaRow').classList.toggle('hidden', !isZona);
      const isIngProv = e.target.value === 'INGRESO_PROVEEDOR';
      document.getElementById('guiaProvRow').classList.toggle('hidden', !isIngProv);
    });

    // Conteo auditoría → mostrar diferencia en tiempo real
    document.getElementById('audConteoFisico')?.addEventListener('input', e => {
      const sis = parseFloat(document.getElementById('audStockSis')?.textContent) || 0;
      const fis = parseFloat(e.target.value) || 0;
      const diff = fis - sis;
      document.getElementById('audDifValor').textContent = (diff >= 0 ? '+' : '') + fmt(diff, 2);
      document.getElementById('audDifValor').className = 'font-bold ' + (Math.abs(diff) < 0.5 ? 'text-emerald-400' : 'text-red-400');
      document.getElementById('audDiferenciaInfo').classList.remove('hidden');
    });

    // Registrar SW
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(() => {});
    }

    // Iniciar sesión (muestra login si no hay sesión activa)
    Session.init();
  }

  function nav(viewName) {
    // Si activa modo envasador, redirigir a vista envasador
    if (modoEnvasador && viewName === 'envasados') {
      viewName = 'envasador';
    }

    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const el = document.getElementById('view-' + viewName);
    if (el) { el.classList.add('active'); el.classList.add('slide-up'); }

    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    // 3-button nav: 0=Inicio, 1=Guías, 2=Más (envasador has no nav highlight)
    const navBtnIdx = viewName === 'dashboard' ? 0
                    : viewName === 'guias'     ? 1
                    : viewName === 'envasador' ? -1
                    : 2;
    if (navBtnIdx >= 0) {
      document.querySelectorAll('.nav-btn')[navBtnIdx]?.classList.add('active');
    }

    document.getElementById('pageTitle').textContent = {
      dashboard:    'Dashboard',
      envasador:    'Modo Envasador',
      guias:        'Guías',
      envasados:    'Envasados',
      preingresos:  'Preingresos',
      mermas:       'Mermas',
      auditorias:   'Auditorías',
      productos:    'Productos',
      proveedores:  'Proveedores',
      config:       'Configuración'
    }[viewName] || viewName;

    currentView = viewName;

    // Lazy-load de cada vista
    switch (viewName) {
      case 'guias':       GuiasView.cargar(); break;
      case 'envasados':   EnvasadosView.cargar(); break;
      case 'envasador':   EnvasadorView.cargar(); break;
      case 'preingresos': PreingresosView.cargar(); break;
      case 'mermas':      MermasView.cargar(); break;
      case 'auditorias':  AuditoriasView.cargar(); break;
      case 'productos':   ProductosView.cargar(); break;
      case 'proveedores': ProveedoresView.cargar(); break;
    }
  }

  function toggleModoEnvasador() {
    modoEnvasador = !modoEnvasador;
    const ind = document.getElementById('modoIndicador');
    const btn = document.getElementById('btnModo');
    ind.classList.toggle('hidden', !modoEnvasador);
    btn.textContent = modoEnvasador ? '✕ Salir modo' : '⚡ Modo';
    if (modoEnvasador) {
      nav('envasador');
      toast('Modo Envasador activado', 'ok');
    } else {
      nav('dashboard');
      toast('Modo normal', 'info');
    }
  }

  async function cargarDashboard() {
    loading('kpiGrid', false);
    const res = await API.getDashboard().catch(() => ({ ok: false }));
    if (!res.ok) {
      toast('Sin conexión — datos de muestra', 'warn');
      return;
    }
    dashboardData = res.data;
    renderDashboard(res.data);
  }

  function renderDashboard(d) {
    if (!d) return;
    const { alertas = {}, kpis = {}, contadores = {} } = d;

    const criticos = alertas.vencimientosCriticos || [];
    const enAlerta = alertas.vencimientosAlerta || alertas.vencimientosEnAlerta || [];
    const pendEnv  = alertas.pendientesEnvasado  || [];
    const stockBajo = alertas.stockBajoMinimo    || [];
    const mermasPend = alertas.mermasPendientes  || [];

    // KPIs principales
    document.getElementById('kpiCriticos').textContent   = contadores.criticos ?? criticos.length;
    document.getElementById('kpiPendEnv').textContent    = pendEnv.length;
    document.getElementById('kpiStockBajo').textContent  = stockBajo.length;
    document.getElementById('kpiMermas').textContent     = fmt(kpis.mermasTotalMes ?? 0, 1);

    // KPIs secundarios
    document.getElementById('kpiEficiencia').textContent = (kpis.eficienciaEnvasadoPct ?? '—') + '%';
    document.getElementById('kpiSalidas').textContent    = contadores.salidasMes ?? '—';

    // Badge navbar
    const badge = document.getElementById('navAlertaBadge');
    const totalAlertas = contadores.alertasTotal ?? 0;
    if (totalAlertas > 0) {
      badge.textContent = totalAlertas > 9 ? '9+' : totalAlertas;
      badge.classList.remove('hidden');
    }

    // Panel Vencimientos
    document.getElementById('listVencCrit').innerHTML = criticos.map(v => `
      <div class="card-sm flex items-center justify-between">
        <div>
          <p class="font-semibold text-sm">${v.descripcion}</p>
          <p class="text-xs text-slate-400">Lote ${v.idLote} — ${fmt(v.cantidadActual)} uds</p>
        </div>
        <span class="${diasColor(v.diasRestantes)} font-bold">${v.diasRestantes}d</span>
      </div>`).join('');
    document.getElementById('listVencAlerta').innerHTML = enAlerta.map(v => `
      <div class="card-sm flex items-center justify-between">
        <div>
          <p class="font-semibold text-sm">${v.descripcion}</p>
          <p class="text-xs text-slate-400">Lote ${v.idLote} — ${fmt(v.cantidadActual)} uds</p>
        </div>
        <span class="${diasColor(v.diasRestantes)} font-bold">${v.diasRestantes}d</span>
      </div>`).join('');
    document.getElementById('vencVacio')?.classList.toggle('hidden', criticos.length + enAlerta.length > 0);

    // Panel Pendientes envasado
    document.getElementById('listPendEnvDash').innerHTML = pendEnv.map(p => `
      <div class="card-sm flex items-center justify-between cursor-pointer" onclick="App.toggleModoEnvasador()">
        <div>
          <p class="font-semibold text-sm">${p.descripcion}</p>
          <p class="text-xs text-slate-400">Stock: ${fmt(p.stockDerivado)} / Mín: ${fmt(p.stockMinimoDerivado)}</p>
          <p class="text-xs text-emerald-400">Base disp: ${fmt(p.stockBase)} → max ${fmt(p.maxProducibles)} uds</p>
        </div>
        <span class="tag-${p.urgencia === 'CRITICA' ? 'danger' : 'warn'}">${p.urgencia}</span>
      </div>`).join('');

    // Panel Stock bajo
    document.getElementById('listStockBajoDash').innerHTML = stockBajo.slice(0, 8).map(s => `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="text-sm font-semibold">${s.descripcion}</span>
          <span class="text-xs ${s.stockActual === 0 ? 'tag-danger' : 'tag-warn'}">
            ${s.stockActual === 0 ? 'SIN STOCK' : fmt(s.stockActual)}
          </span>
        </div>
        <div class="bar-bg"><div class="bar-fill bg-amber-500"
          style="width:${Math.min(100, (s.stockActual / s.stockMinimo * 100)).toFixed(0)}%"></div></div>
        <p class="text-xs text-slate-500 mt-1">Mínimo: ${fmt(s.stockMinimo)} — Faltan: ${fmt(s.diferencia)}</p>
      </div>`).join('');

    // Panel Mermas pendientes
    document.getElementById('listMermasDash').innerHTML = mermasPend.map(m => `
      <div class="card-sm flex items-center justify-between">
        <div>
          <p class="font-semibold text-sm">${m.codigoProducto || m.descripcion || '—'}</p>
          <p class="text-xs text-slate-400">${fmtFecha(m.fechaIngreso)} · ${m.origen || ''}</p>
        </div>
        <span class="tag-warn text-xs">${fmt(m.cantidadOriginal, 1)}</span>
      </div>`).join('');
  }

  async function cargarProductosMaestro() {
    const res = await API.getProductos({ estado: '1' }).catch(() => ({ ok: false }));
    if (res.ok) todosProductos = res.data;
  }

  async function cargarProveedoresMaestro() {
    const res = await API.getProveedores({ estado: '1' }).catch(() => ({ ok: false }));
    if (res.ok) {
      todosProveedores = res.data;
      // Llenar selects de proveedores
      ['guiaProveedor', 'preProvSelect'].forEach(id => {
        const sel = document.getElementById(id);
        if (!sel) return;
        todosProveedores.forEach(p => {
          const opt = document.createElement('option');
          opt.value = p.idProveedor;
          opt.textContent = p.nombre;
          sel.appendChild(opt);
        });
      });
    }
  }

  function showUsuarioDialog() {
    const nombre = prompt('Nombre del operador:', window.WH_CONFIG.usuario);
    if (nombre) {
      window.WH_CONFIG.usuario = nombre;
      localStorage.setItem('wh_usuario', nombre);
      document.getElementById('usuarioNombre').textContent = nombre;
      document.getElementById('cfgUsuario').value = nombre;
    }
  }

  function getProductosMaestro() { return todosProductos; }
  function getProveedoresMaestro() { return todosProveedores; }

  function abrirMas() { abrirSheet('sheetMas'); }
  function navMas(viewName) { cerrarSheet('sheetMas'); nav(viewName); }

  return { init, nav, abrirMas, navMas, toggleModoEnvasador, cargarDashboard, showUsuarioDialog,
           cargarProductosMaestro, cargarProveedoresMaestro,
           getProductosMaestro, getProveedoresMaestro };
})();

// ════════════════════════════════════════════════
// GUIAS VIEW
// ════════════════════════════════════════════════
const GuiasView = (() => {
  let todas = [];
  let filtroActual = '';

  async function cargar() {
    loading('listGuias', true);
    const res = await API.getGuias({ limit: 50 }).catch(() => ({ ok: false }));
    todas = res.ok ? res.data : [];
    render(todas);
  }

  function filtrar(f) {
    filtroActual = f;
    document.querySelectorAll('.guia-tab').forEach(b => b.classList.toggle('active-tab', b.textContent.toLowerCase().includes(f.toLowerCase())));
    let list = todas;
    if (f === 'INGRESO') list = todas.filter(g => g.tipo?.startsWith('INGRESO'));
    else if (f === 'SALIDA') list = todas.filter(g => g.tipo?.startsWith('SALIDA'));
    else if (f === 'ABIERTA') list = todas.filter(g => g.estado === 'ABIERTA');
    render(list);
  }

  function render(list) {
    const container = document.getElementById('listGuias');
    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">No hay guías</p>'; return; }

    const tipoLabels = {
      INGRESO_PROVEEDOR: '🚚 Proveedor', INGRESO_JEFATURA: '🏢 Jefatura',
      SALIDA_ZONA: '📍 Zona', SALIDA_DEVOLUCION: '↩️ Devolución',
      SALIDA_JEFATURA: '🏢 Jefatura', SALIDA_ENVASADO: '📦 Envasado', SALIDA_MERMA: '⚠️ Merma'
    };

    container.innerHTML = list.map(g => `
      <div class="card-sm cursor-pointer" onclick="GuiasView.verDetalle('${g.idGuia}')">
        <div class="flex items-center justify-between mb-1">
          <span class="text-xs ${g.tipo?.startsWith('INGRESO') ? 'tag-ok' : 'tag-blue'}">
            ${tipoLabels[g.tipo] || g.tipo}
          </span>
          <span class="text-xs ${g.estado === 'ABIERTA' ? 'tag-warn' : 'tag-ok'}">${g.estado}</span>
        </div>
        <p class="font-semibold text-sm">${g.idGuia}</p>
        <p class="text-xs text-slate-400">${fmtFecha(g.fecha)} · ${g.usuario || '—'}</p>
        ${g.comentario ? `<p class="text-xs text-slate-500 mt-1">${g.comentario}</p>` : ''}
        ${g.montoTotal && parseFloat(g.montoTotal) > 0
          ? `<p class="text-xs text-emerald-400 mt-1 font-semibold">S/. ${fmt(g.montoTotal, 2)}</p>` : ''}
      </div>`).join('');
  }

  async function verDetalle(idGuia) {
    toast('Cargando guía...', 'info');
    const res = await API.getGuia(idGuia);
    if (!res.ok) { toast('Error al cargar guía', 'danger'); return; }
    const g = res.data;
    const items = (g.detalle || []).map(d =>
      `<div class="flex justify-between text-sm py-1 border-b border-slate-700">
        <span>${d.descripcionProducto || d.codigoProducto}</span>
        <span class="font-mono">${fmt(d.cantidadRecibida)}</span>
      </div>`
    ).join('');
    toast(`${g.idGuia}: ${g.detalle?.length || 0} ítems — ${g.estado}`, 'info', 4000);
  }

  async function crearGuia() {
    const tipo = document.getElementById('guiaTipo').value;
    const params = {
      tipo,
      usuario:        window.WH_CONFIG.usuario,
      idProveedor:    document.getElementById('guiaProveedor').value,
      idZona:         document.getElementById('guiaZona').value,
      numeroDocumento:document.getElementById('guiaNumDoc').value,
      comentario:     document.getElementById('guiaComentario').value
    };
    const res = await API.crearGuia(params);
    if (res.ok) {
      toast(`Guía ${res.data.idGuia} creada`, 'ok');
      cerrarSheet('sheetGuia');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  function nueva() { abrirSheet('sheetGuia'); }

  return { cargar, filtrar, verDetalle, crearGuia, nueva };
})();

// ════════════════════════════════════════════════
// ENVASADOS VIEW
// ════════════════════════════════════════════════
const EnvasadosView = (() => {
  let derivados = [];
  let productosMaestro = [];

  async function cargar() {
    loading('listEnvasados', true);
    const res = await API.getEnvasados({ limit: 30 }).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    const container = document.getElementById('listEnvasados');

    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin envasados hoy</p>'; return; }

    container.innerHTML = list.map(e => {
      const efPct = parseFloat(e.eficienciaPct) || 0;
      const efColor = efPct >= 95 ? 'text-emerald-400' : efPct >= 85 ? 'text-amber-400' : 'text-red-400';
      return `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="text-xs tag-blue">${e.codigoProductoBase} → ${e.codigoProductoEnvasado}</span>
          <span class="${efColor} font-bold text-sm">${efPct}%</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(e.fecha)} · ${e.usuario}</p>
        <div class="flex gap-4 mt-1 text-xs text-slate-300">
          <span>Base: ${fmt(e.cantidadBase, 1)} ${e.unidadBase}</span>
          <span>Prod: ${fmt(e.unidadesProducidas)} uds</span>
          <span class="text-amber-400">Merma: ${fmt(e.mermaReal)}</span>
        </div>
      </div>`;
    }).join('');
  }

  async function nuevo() {
    // Cargar derivados
    productosMaestro = App.getProductosMaestro();
    derivados = productosMaestro.filter(p => p.codigoProductoBase && p.codigoProductoBase !== '');

    const sel = document.getElementById('envProductoDerivado');
    sel.innerHTML = '<option value="">— Seleccionar producto —</option>';
    derivados.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.idProducto;
      opt.textContent = p.descripcion + ' (base: ' + p.codigoProductoBase + ')';
      sel.appendChild(opt);
    });

    abrirSheet('sheetEnvasado');
  }

  async function onDerivadoChange(idDerivado) {
    const prod = derivados.find(p => p.idProducto === idDerivado);
    if (!prod) {
      document.getElementById('envasadoFactorInfo').classList.add('hidden');
      return;
    }

    const factor  = parseFloat(prod.factorConversion)  || 1;
    const merma   = parseFloat(prod.mermaEsperadaPct)   || 0;

    // Cargar stock base en tiempo real
    const stockRes = await API.getStockProducto(prod.codigoProductoBase).catch(() => ({ ok: false }));
    const stockBase = stockRes.ok ? stockRes.data.cantidad : 0;
    const unidadBase = stockRes.ok ? stockRes.data.unidad : '';

    document.getElementById('envFactor').textContent       = `${factor} uds/${unidadBase || 'unidad base'}`;
    document.getElementById('envMerma').textContent        = merma + '%';
    document.getElementById('envStockBase').textContent    = `${fmt(stockBase, 1)} ${unidadBase}`;
    document.getElementById('envUnidadBase').textContent   = unidadBase || 'unidades';
    const maxProd = Math.floor(stockBase * factor * (1 - merma / 100));
    document.getElementById('envMaxProd').textContent      = `${maxProd} uds`;
    document.getElementById('envasadoFactorInfo').classList.remove('hidden');

    // Sugerir cantidad base
    document.getElementById('envCantBase').placeholder = `Máx. ${fmt(stockBase, 1)} ${unidadBase}`;
  }

  function calcularProyeccion() {
    const idDerivado = document.getElementById('envProductoDerivado').value;
    const prod = derivados.find(p => p.idProducto === idDerivado);
    if (!prod) return;

    const cantBase = parseFloat(document.getElementById('envCantBase').value) || 0;
    const factor   = parseFloat(prod.factorConversion)  || 1;
    const merma    = parseFloat(prod.mermaEsperadaPct)   || 0;
    const esperadas = Math.floor(cantBase * factor * (1 - merma / 100));

    document.getElementById('envUnidades').value = esperadas;
    actualizarResumen();
  }

  function actualizarResumen() {
    const idDerivado  = document.getElementById('envProductoDerivado').value;
    const prod = derivados.find(p => p.idProducto === idDerivado);
    if (!prod) return;

    const cantBase    = parseFloat(document.getElementById('envCantBase').value)   || 0;
    const producidas  = parseInt(document.getElementById('envUnidades').value)     || 0;
    const factor      = parseFloat(prod.factorConversion)  || 1;
    const merma       = parseFloat(prod.mermaEsperadaPct)   || 0;
    const esperadas   = Math.floor(cantBase * factor * (1 - merma / 100));
    const mermaReal   = Math.max(0, esperadas - producidas);
    const efic        = esperadas > 0 ? (producidas / esperadas * 100).toFixed(1) : '—';

    document.getElementById('rEsperadas').textContent = esperadas;
    document.getElementById('rProducidas').textContent = producidas;
    document.getElementById('rMerma').textContent = mermaReal;
    document.getElementById('rEficiencia').textContent = efic + '%';
    document.getElementById('rEficiencia').className = parseFloat(efic) >= 95 ? 'text-emerald-400 font-bold' : 'text-amber-400 font-bold';
    document.getElementById('envResumen').classList.remove('hidden');
  }

  function ajustarUnidades(delta) {
    const el = document.getElementById('envUnidades');
    el.value = Math.max(0, (parseInt(el.value) || 0) + delta);
    actualizarResumen();
  }

  function setUnidades(n) {
    document.getElementById('envUnidades').value = n;
    actualizarResumen();
  }

  async function registrar() {
    const idDerivado = document.getElementById('envProductoDerivado').value;
    const cantBase   = parseFloat(document.getElementById('envCantBase').value)   || 0;
    const producidas = parseInt(document.getElementById('envUnidades').value)     || 0;
    const fechaVenc  = document.getElementById('envFechaVenc').value;
    const imprimir   = document.getElementById('envImprimirEtiq').checked;

    if (!idDerivado || cantBase <= 0 || producidas <= 0) {
      toast('Completa todos los campos', 'warn');
      return;
    }

    const prod = derivados.find(p => p.idProducto === idDerivado);
    const btn  = document.getElementById('btnRegistrarEnvasado');
    btn.disabled = true;
    btn.textContent = '⏳ Guardando...';

    const res = await API.registrarEnvasado({
      codigoProductoBase:    prod.codigoProductoBase,
      cantidadBase:          cantBase,
      codigoProductoEnvasado: idDerivado,
      unidadesProducidas:    producidas,
      fechaVencimiento:      fechaVenc,
      imprimirEtiquetas:     imprimir,
      usuario:               window.WH_CONFIG.usuario
    });

    btn.disabled = false;
    btn.textContent = '📦 Registrar + Imprimir etiquetas';

    if (res.ok) {
      const d = res.data;
      const impMsg = imprimir
        ? (d.impresion?.ok ? ` · ${producidas} etiquetas enviadas` : ' · ⚠️ Impresora no lista')
        : '';
      toast(`✅ ${producidas} uds registradas. Efic: ${d.eficienciaPct}%${impMsg}`, 'ok', 5000);
      cerrarSheet('sheetEnvasado');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger', 5000);
    }
  }

  return { cargar, nuevo, onDerivadoChange, calcularProyeccion, ajustarUnidades, setUnidades, registrar };
})();

// ════════════════════════════════════════════════
// MODO ENVASADOR — vista rápida urgentes
// ════════════════════════════════════════════════
const EnvasadorView = (() => {
  let _historialCargado = false;

  function setTab(tab) {
    const isUrg = tab === 'urgentes';
    document.getElementById('tabPanelUrgentes').classList.toggle('hidden', !isUrg);
    document.getElementById('tabPanelHistorial').classList.toggle('hidden', isUrg);
    const actCls = 'flex-1 py-2 text-sm font-semibold border-b-2 -mb-px text-blue-400 border-blue-400';
    const inaCls = 'flex-1 py-2 text-sm font-semibold border-b-2 -mb-px text-slate-500 border-transparent';
    document.getElementById('tabBtnUrgentes').className  = isUrg  ? actCls : inaCls;
    document.getElementById('tabBtnHistorial').className = !isUrg ? actCls : inaCls;
    if (!isUrg && !_historialCargado) _cargarHistorial();
  }

  async function _cargarHistorial() {
    _historialCargado = true;
    const container = document.getElementById('listHistorialEnvasado');
    container.innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';
    const res = await API.getEnvasados({ limit: 30 }).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin historial</p>'; return; }
    container.innerHTML = list.map(e => {
      const efPct = parseFloat(e.eficienciaPct) || 0;
      const efColor = efPct >= 95 ? 'text-emerald-400' : efPct >= 85 ? 'text-amber-400' : 'text-red-400';
      return `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="text-xs tag-blue">${e.codigoProductoBase} → ${e.codigoProductoEnvasado}</span>
          <span class="${efColor} font-bold text-sm">${efPct}%</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(e.fecha)} · ${e.usuario}</p>
        <div class="flex gap-4 mt-1 text-xs text-slate-300">
          <span>Base: ${fmt(e.cantidadBase, 1)}</span>
          <span>Prod: ${fmt(e.unidadesProducidas)} uds</span>
          <span class="text-amber-400">Merma: ${fmt(e.mermaReal)}</span>
        </div>
      </div>`;
    }).join('');
  }

  async function cargar() {
    _historialCargado = false;
    // Asegurar que estamos en tab urgentes al entrar
    setTab('urgentes');
    const container = document.getElementById('listEnvasadorUrgentes');
    container.innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';

    const res = await API.getPendientes().catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];

    if (!list.length) {
      container.innerHTML = `
        <div class="card text-center py-8">
          <p class="text-2xl mb-2">✅</p>
          <p class="font-semibold">¡Todo en orden!</p>
          <p class="text-xs text-slate-400 mt-1">No hay productos urgentes para envasar</p>
        </div>`;
      return;
    }

    container.innerHTML = list.map(p => `
      <div class="card">
        <div class="flex items-center justify-between mb-2">
          <span class="font-bold text-base">${p.descripcion}</span>
          <span class="tag-${p.urgencia === 'CRITICA' ? 'danger' : 'warn'} font-bold">${p.urgencia}</span>
        </div>

        <div class="grid grid-cols-2 gap-2 text-xs mb-3">
          <div class="card-sm text-center">
            <div class="text-lg font-bold ${p.stockDerivado === 0 ? 'text-red-400' : 'text-amber-400'}">${fmt(p.stockDerivado)}</div>
            <div class="text-slate-400">Stock actual</div>
          </div>
          <div class="card-sm text-center">
            <div class="text-lg font-bold text-slate-300">${fmt(p.stockMinimoDerivado)}</div>
            <div class="text-slate-400">Mínimo requerido</div>
          </div>
          <div class="card-sm text-center">
            <div class="text-lg font-bold text-blue-400">${fmt(p.stockBase, 1)}</div>
            <div class="text-slate-400">Base disponible</div>
          </div>
          <div class="card-sm text-center">
            <div class="text-lg font-bold text-emerald-400">${fmt(p.maxProducibles)}</div>
            <div class="text-slate-400">Máx. producibles</div>
          </div>
        </div>

        <!-- Registro rápido inline -->
        <div class="flex gap-2 items-center">
          <button onclick="EnvasadorView.ajustar('${p.codigoDerivado}', -10)" class="btn btn-outline w-10 h-10 text-lg">−</button>
          <input id="quickEnv_${p.codigoDerivado}" type="number" value="${Math.min(p.maxProducibles, p.faltan)}"
                 class="input text-center text-xl font-bold flex-1" min="1" max="${p.maxProducibles}"/>
          <button onclick="EnvasadorView.ajustar('${p.codigoDerivado}', +10)" class="btn btn-outline w-10 h-10 text-lg">+</button>
        </div>
        <div class="text-xs text-slate-500 text-center mt-1">
          Factor: ${p.factorConversion} uds/unidad · Merma: ${p.mermaEsperadaPct}%
        </div>
        <button onclick="EnvasadorView.registrarRapido('${p.codigoDerivado}','${p.codigoBase}',${p.factorConversion},${p.mermaEsperadaPct})"
                class="btn btn-primary w-full mt-2 btn-lg">
          📦 Envasar + 🖨️ Etiquetas
        </button>
      </div>`).join('');
  }

  function ajustar(codigoDerivado, delta) {
    const el = document.getElementById('quickEnv_' + codigoDerivado);
    if (el) el.value = Math.max(1, (parseInt(el.value) || 0) + delta);
  }

  async function registrarRapido(codigoDerivado, codigoBase, factor, merma) {
    const el = document.getElementById('quickEnv_' + codigoDerivado);
    const unidades = parseInt(el?.value) || 0;
    if (unidades <= 0) { toast('Ingresa la cantidad', 'warn'); return; }

    // Calcular cantidad base necesaria (inverso del factor)
    const cantBase = Math.ceil(unidades / (factor * (1 - merma / 100)));

    const res = await API.registrarEnvasado({
      codigoProductoBase:     codigoBase,
      cantidadBase:           cantBase,
      codigoProductoEnvasado: codigoDerivado,
      unidadesProducidas:     unidades,
      imprimirEtiquetas:      true,
      usuario:                window.WH_CONFIG.usuario
    });

    if (res.ok) {
      const d = res.data;
      const impMsg = d.impresion?.ok ? ` · ${unidades} etiquetas 🖨️` : ' · Impresora no lista';
      toast(`✅ ${unidades} uds envasadas · ${d.eficienciaPct}%${impMsg}`, 'ok', 5000);
      cargar(); // Recargar lista
    } else {
      toast('Error: ' + res.error, 'danger', 5000);
    }
  }

  return { cargar, setTab, ajustar, registrarRapido };
})();

// ════════════════════════════════════════════════
// PREINGRESOS VIEW
// ════════════════════════════════════════════════
const PreingresosView = (() => {
  async function cargar(estado = '') {
    loading('listPreingresos', true);
    const res = await API.getPreingresos(estado ? { estado } : {}).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    const container = document.getElementById('listPreingresos');

    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin preingresos</p>'; return; }

    container.innerHTML = list.map(p => `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="font-semibold text-sm">${p.idPreingreso}</span>
          <span class="tag-${p.estado === 'PENDIENTE' ? 'warn' : p.estado === 'PROCESADO' ? 'ok' : 'blue'}">${p.estado}</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(p.fecha)} · ${p.idProveedor}</p>
        ${p.numeroFactura ? `<p class="text-xs text-slate-300">Fact: ${p.numeroFactura}</p>` : ''}
        <p class="text-sm font-bold text-emerald-400 mt-1">S/. ${fmt(p.monto, 2)}</p>
        ${p.estado === 'PENDIENTE'
          ? `<button onclick="PreingresosView.aprobar('${p.idPreingreso}')"
                     class="btn btn-primary w-full mt-2 text-sm py-2">
               Aprobar → Crear Guía
             </button>`
          : ''}
      </div>`).join('');
  }

  function filtrar(estado) { cargar(estado); }

  async function crear() {
    const params = {
      idProveedor:    document.getElementById('preProvSelect').value,
      numeroFactura:  document.getElementById('preNumFact').value,
      monto:          document.getElementById('preMonto').value,
      comentario:     document.getElementById('preComentario').value,
      usuario:        window.WH_CONFIG.usuario
    };
    if (!params.idProveedor) { toast('Selecciona un proveedor', 'warn'); return; }

    const res = await API.crearPreingreso(params);
    if (res.ok) {
      toast('Preingreso registrado', 'ok');
      cerrarSheet('sheetPreingreso');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  async function aprobar(id) {
    const res = await API.aprobarPreingreso({ idPreingreso: id, usuario: window.WH_CONFIG.usuario });
    if (res.ok) {
      toast(`Guía ${res.data.idGuia} creada`, 'ok');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  function nuevo() { abrirSheet('sheetPreingreso'); }

  return { cargar, filtrar, crear, aprobar, nuevo };
})();

// ════════════════════════════════════════════════
// MERMAS VIEW
// ════════════════════════════════════════════════
const MermasView = (() => {
  async function cargar() {
    loading('listMermas', true);
    const res = await API.getMermas({ limit: 30 }).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    const container = document.getElementById('listMermas');

    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin mermas registradas</p>'; return; }

    container.innerHTML = list.map(m => `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="font-semibold text-sm">${m.codigoProducto}</span>
          <span class="tag-${m.estado === 'PENDIENTE' ? 'warn' : 'ok'}">${m.estado}</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(m.fechaIngreso)} · ${m.origen}</p>
        <p class="text-sm font-bold text-red-400">${fmt(m.cantidadOriginal, 2)} uds</p>
        ${m.motivo ? `<p class="text-xs text-slate-500 mt-1">${m.motivo}</p>` : ''}
      </div>`).join('');
  }

  async function crear() {
    const params = {
      codigoProducto:  document.getElementById('mermaCodigoProd').value,
      origen:          document.getElementById('mermaOrigen').value,
      cantidadOriginal:document.getElementById('mermaCantidad').value,
      motivo:          document.getElementById('mermaMotivo').value,
      descontarStock:  document.getElementById('mermaDescontar').checked,
      usuario:         window.WH_CONFIG.usuario
    };
    if (!params.codigoProducto || !params.cantidadOriginal) {
      toast('Completa producto y cantidad', 'warn');
      return;
    }
    const res = await API.registrarMerma(params);
    if (res.ok) {
      toast('Merma registrada', 'ok');
      cerrarSheet('sheetMerma');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  function nueva() { abrirSheet('sheetMerma'); }
  return { cargar, crear, nueva };
})();

// ════════════════════════════════════════════════
// AUDITORIAS VIEW
// ════════════════════════════════════════════════
const AuditoriasView = (() => {
  let auditoriaActiva = null;

  async function cargar(estado = '') {
    loading('listAuditorias', true);
    const res = await API.getAuditorias(estado ? { estado } : {}).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    const container = document.getElementById('listAuditorias');

    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin auditorías</p>'; return; }

    container.innerHTML = list.map(a => `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="font-semibold text-sm">${a.codigoProducto}</span>
          <span class="tag-${a.estado === 'EJECUTADA' ? (a.resultado === 'OK' ? 'ok' : 'danger') : 'warn'}">${a.resultado || a.estado}</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(a.fechaAsignacion)} · ${a.usuario || '—'}</p>
        ${a.diferencia !== undefined && a.diferencia !== ''
          ? `<p class="text-xs ${parseFloat(a.diferencia) === 0 ? 'text-emerald-400' : 'text-red-400'}">
               Diferencia: ${parseFloat(a.diferencia) >= 0 ? '+' : ''}${fmt(a.diferencia, 2)}
             </p>` : ''}
        ${a.estado !== 'EJECUTADA'
          ? `<button onclick="AuditoriasView.abrirEjecucion('${a.idAuditoria}','${a.codigoProducto}',${a.stockSistema})"
                     class="btn btn-outline w-full mt-2 text-sm py-2">Contar físico</button>` : ''}
      </div>`).join('');
  }

  function filtrar(estado) { cargar(estado); }

  function abrirEjecucion(id, codigo, stockSis) {
    auditoriaActiva = id;
    document.getElementById('audModalSubtitle').textContent = `Producto: ${codigo}`;
    document.getElementById('audStockSis').textContent = fmt(stockSis, 2);
    document.getElementById('audConteoFisico').value = '';
    document.getElementById('audDiferenciaInfo').classList.add('hidden');
    abrirModal('modalAuditoria');
  }

  async function ejecutar() {
    const fisico = parseFloat(document.getElementById('audConteoFisico').value);
    const obs    = document.getElementById('audObservacion').value;
    if (isNaN(fisico)) { toast('Ingresa el conteo físico', 'warn'); return; }

    const res = await API.ejecutarAuditoria({ idAuditoria: auditoriaActiva, stockFisico: fisico, observacion: obs });
    if (res.ok) {
      const d = res.data;
      const msg = d.resultado === 'OK' ? '✅ Sin diferencias' : `⚠️ Diferencia: ${d.diferencia}`;
      toast(msg, d.resultado === 'OK' ? 'ok' : 'warn', 4000);
      cerrarModal('modalAuditoria');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  async function nueva() {
    const codigo = prompt('Código del producto a auditar:');
    if (!codigo) return;
    const res = await API.asignarAuditoria({ codigoProducto: codigo, usuario: window.WH_CONFIG.usuario });
    if (res.ok) {
      toast('Auditoría asignada', 'ok');
      cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  return { cargar, filtrar, abrirEjecucion, ejecutar, nueva };
})();

// ════════════════════════════════════════════════
// PRODUCTOS VIEW
// ════════════════════════════════════════════════
const ProductosView = (() => {
  let todos = [];

  async function cargar() {
    loading('listProductos', true);
    const res = await API.getStock().catch(() => ({ ok: false }));
    todos = res.ok ? res.data : [];
    renderLista(todos);
  }

  function buscar(q) {
    if (!q) { renderLista(todos); return; }
    const qL = q.toLowerCase();
    renderLista(todos.filter(p =>
      (p.descripcion || '').toLowerCase().includes(qL) ||
      (p.codigoProducto || '').toLowerCase().includes(qL)
    ));
  }

  function renderLista(list) {
    const container = document.getElementById('listProductos');
    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin resultados</p>'; return; }

    container.innerHTML = list.map(p => {
      const pct = p.stockMaximo > 0 ? Math.min(100, p.cantidadDisponible / p.stockMaximo * 100) : 0;
      const barColor = p.alertaMinimo ? 'bg-red-500' : pct < 50 ? 'bg-amber-500' : 'bg-emerald-500';
      return `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <p class="font-semibold text-sm">${p.descripcion}</p>
          <span class="font-mono font-bold ${p.alertaMinimo ? 'text-red-400' : 'text-emerald-400'}">
            ${fmt(p.cantidadDisponible)}
          </span>
        </div>
        <p class="text-xs text-slate-400 mb-1">${p.codigoProducto} · ${p.unidad}</p>
        <div class="bar-bg"><div class="bar-fill ${barColor}" style="width:${pct.toFixed(0)}%"></div></div>
        <p class="text-xs text-slate-500 mt-1">Mín: ${p.stockMinimo} · Máx: ${p.stockMaximo}</p>
        ${p.alertaMinimo ? '<span class="tag-danger text-xs mt-1 inline-block">⚠️ BAJO MÍNIMO</span>' : ''}
      </div>`;
    }).join('');
  }

  return { cargar, buscar };
})();

// ════════════════════════════════════════════════
// PROVEEDORES VIEW
// ════════════════════════════════════════════════
const ProveedoresView = (() => {
  async function cargar() {
    loading('listProveedores', true);
    const res = await API.getProveedores({ estado: '1' }).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    const container = document.getElementById('listProveedores');

    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin proveedores</p>'; return; }

    const dias = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
    container.innerHTML = list.map(p => `
      <div class="card-sm">
        <p class="font-bold text-sm">${p.nombre}</p>
        <p class="text-xs text-slate-400">RUC: ${p.ruc || '—'} · ${p.telefono || ''}</p>
        <div class="flex gap-3 mt-2 text-xs text-slate-300">
          <span>📋 Pedido: ${dias[p.diaPedido] || p.diaPedido || '—'}</span>
          <span>💳 Pago: ${dias[p.diaPago] || p.diaPago || '—'}</span>
          <span>🚚 Entrega: ${dias[p.diaEntrega] || p.diaEntrega || '—'}</span>
        </div>
        <p class="text-xs text-slate-400 mt-1">
          ${p.formaPago} ${p.plazoCredito > 0 ? '· Crédito ' + p.plazoCredito + 'd' : ''}
          ${p.responsable ? '· ' + p.responsable : ''}
        </p>
      </div>`).join('');
  }

  function nuevo() { toast('Formulario proveedor próximamente', 'info'); }
  return { cargar, nuevo };
})();

// ════════════════════════════════════════════════
// CONFIG VIEW
// ════════════════════════════════════════════════
const ConfigView = (() => {
  async function guardar() {
    const gasUrl = document.getElementById('cfgGasUrl').value.trim();
    const printKey = document.getElementById('cfgPrintKey').value.trim();
    const printId = document.getElementById('cfgPrintId').value.trim();
    const diasAlerta = document.getElementById('cfgDiasAlerta').value;

    if (gasUrl) {
      window.WH_CONFIG.gasUrl = gasUrl;
      localStorage.setItem('wh_gas_url', gasUrl);
    }

    if (printKey) await API.setConfig('PRINTNODE_API_KEY', printKey);
    if (printId)  await API.setConfig('PRINTNODE_PRINTER_ID', printId);
    if (diasAlerta) await API.setConfig('DIAS_ALERTA_VENC', diasAlerta);

    toast('Configuración guardada', 'ok');
  }

  function guardarUsuario() {
    const nombre = document.getElementById('cfgUsuario').value.trim();
    if (!nombre) return;
    window.WH_CONFIG.usuario = nombre;
    localStorage.setItem('wh_usuario', nombre);
    document.getElementById('usuarioNombre').textContent = nombre;
    toast('Usuario actualizado', 'ok');
  }

  return { guardar, guardarUsuario };
})();

// ════════════════════════════════════════════════
// Init
// ════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => App.init());
