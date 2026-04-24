// warehouseMos — app.js  Lógica de la aplicación
'use strict';

// ════════════════════════════════════════════════
// Vibración táctil (feedback)
// ════════════════════════════════════════════════
let _userActivated = false;
document.addEventListener('pointerdown', () => { _userActivated = true; }, { once: true });

function vibrate(ms = 10) {
  if (_userActivated && navigator.vibrate) navigator.vibrate(ms);
}

// ════════════════════════════════════════════════
// Long press handler
// ════════════════════════════════════════════════
function addLongPress(el, cb, ms = 500) {
  let timer = null;
  const start = e => {
    timer = setTimeout(() => {
      vibrate(30);
      // Ripple visual
      const rect = el.getBoundingClientRect();
      const x = (e.touches?.[0]?.clientX ?? e.clientX) - rect.left;
      const y = (e.touches?.[0]?.clientY ?? e.clientY) - rect.top;
      const rip = document.createElement('div');
      rip.className = 'lp-ripple';
      rip.style.cssText = `left:${x}px;top:${y}px;`;
      el.style.position = 'relative';
      el.style.overflow = 'hidden';
      el.appendChild(rip);
      setTimeout(() => rip.remove(), 600);
      cb(e);
    }, ms);
  };
  const cancel = () => clearTimeout(timer);
  el.addEventListener('touchstart', start, { passive: true });
  el.addEventListener('touchend',   cancel);
  el.addEventListener('touchmove',  cancel, { passive: true });
  el.addEventListener('mousedown',  start);
  el.addEventListener('mouseup',    cancel);
  el.addEventListener('mouseleave', cancel);
}

// ════════════════════════════════════════════════
// Double-tap confirmation para acciones destructivas
// ════════════════════════════════════════════════
function dblTapConfirm(btn, action) {
  if (btn._armed) {
    btn._armed = false;
    btn.classList.remove('armed');
    clearTimeout(btn._armTimer);
    vibrate(20);
    action();
    return;
  }
  btn._armed = true;
  btn.classList.add('armed', 'dbl-btn');
  vibrate(15);
  btn._armTimer = setTimeout(() => {
    btn._armed = false;
    btn.classList.remove('armed');
  }, 2500);
}

// ════════════════════════════════════════════════
// Pull-to-refresh
// ════════════════════════════════════════════════
const PullToRefresh = (() => {
  let startY = 0, pulling = false, threshold = 70;

  function init(scrollEl, onRefresh) {
    const indicator = document.getElementById('ptr-indicator');
    if (!scrollEl || !indicator) return;

    scrollEl.addEventListener('touchstart', e => {
      if (scrollEl.scrollTop === 0) startY = e.touches[0].clientY;
    }, { passive: true });

    scrollEl.addEventListener('touchmove', e => {
      if (!startY) return;
      const dy = e.touches[0].clientY - startY;
      if (dy > 20 && scrollEl.scrollTop === 0) {
        pulling = true;
        if (dy > threshold) indicator.classList.add('visible');
      }
    }, { passive: true });

    scrollEl.addEventListener('touchend', () => {
      if (pulling && indicator.classList.contains('visible')) {
        vibrate(15);
        onRefresh();
        setTimeout(() => indicator.classList.remove('visible'), 800);
      }
      startY = 0; pulling = false;
    });
  }

  return { init };
})();

// ════════════════════════════════════════════════
// Badge counts — actualiza badges en nav
// ════════════════════════════════════════════════
function actualizarBadges({ guiasAbiertas = 0 } = {}) {
  const _setBadge = (id, n) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = n > 9 ? '9+' : String(n);
    el.dataset.hidden = n === 0 ? 'true' : 'false';
  };
  _setBadge('badgeGuias',     guiasAbiertas);
  _setBadge('sideBadgeGuias', guiasAbiertas);
}

// ════════════════════════════════════════════════
// PrintNode status indicator
// ════════════════════════════════════════════════
const PrintNodeStatus = (() => {
  function set(estado) {
    const dot = document.getElementById('pnDot');
    if (!dot) return;
    dot.className = 'pn-dot ' + (estado || '');
  }
  function busy()  { set('busy');  setTimeout(() => set('ok'), 4000); }
  function ok()    { set('ok');    }
  function error() { set('error'); }
  function hide()  { set('');      }
  return { ok, error, busy, hide };
})();

// ════════════════════════════════════════════════
// Global Search
// ════════════════════════════════════════════════
const GlobalSearch = (() => {
  let _debounce = null;

  function abrir() {
    document.getElementById('globalSearchOverlay')?.classList.add('open');
    setTimeout(() => document.getElementById('gSearchInput')?.focus(), 80);
    vibrate(8);
  }

  function cerrar() {
    document.getElementById('globalSearchOverlay')?.classList.remove('open');
    const inp = document.getElementById('gSearchInput');
    if (inp) inp.value = '';
    document.getElementById('gSearchResults').innerHTML =
      '<p class="text-slate-500 text-sm text-center pt-8">Escribe para buscar en toda la app</p>';
  }

  function buscar(q) {
    clearTimeout(_debounce);
    if (!q || q.trim().length < 2) {
      document.getElementById('gSearchResults').innerHTML =
        '<p class="text-slate-500 text-sm text-center pt-8">Escribe para buscar en toda la app</p>';
      return;
    }
    _debounce = setTimeout(() => _ejecutar(q.trim()), 250);
  }

  function _ejecutar(q) {
    const qL = q.toLowerCase();
    const res = document.getElementById('gSearchResults');

    // Buscar en guías (caché)
    const guias = (OfflineManager.getGuiasCache() || []).filter(g =>
      (g.idGuia || '').toLowerCase().includes(qL) ||
      (g.tipo   || '').toLowerCase().includes(qL)
    ).slice(0, 5);

    // Buscar en productos (caché)
    const productos = (OfflineManager.getProductosCache() || []).filter(p =>
      (p.descripcion  || '').toLowerCase().includes(qL) ||
      (p.codigoBarra  || '').toLowerCase().includes(qL) ||
      (p.idProducto   || '').toLowerCase().includes(qL)
    ).slice(0, 5);

    // Buscar en preingresos (caché)
    const preingresos = (OfflineManager.getPreingresosCache?.() || []).filter(pi =>
      (pi.idPreingreso || '').toLowerCase().includes(qL) ||
      (pi.proveedor    || '').toLowerCase().includes(qL)
    ).slice(0, 5);

    if (!guias.length && !productos.length && !preingresos.length) {
      res.innerHTML = `<p class="text-slate-400 text-sm text-center pt-8">Sin resultados para "<strong>${escHtml(q)}</strong>"</p>`;
      return;
    }

    let html = '';
    if (guias.length) {
      html += `<div class="gs-section"><div class="gs-hdr">Guías</div>`;
      html += guias.map(g => `
        <div class="gs-item" onclick="GlobalSearch.cerrar();App.nav('guias')">
          <div class="gs-item-ico" style="background:#1e3a8a">
            <svg width="14" height="14" viewBox="0 0 16 16" fill="#93c5fd">
              <path d="M4 0h5.293A1 1 0 0 1 10 .293L13.707 4a1 1 0 0 1 .293.707V14a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2z"/>
            </svg>
          </div>
          <div class="gs-item-main">
            <div class="gs-item-title">${escHtml(g.idGuia || '—')}</div>
            <div class="gs-item-sub">${escHtml(g.tipo || '')} · ${escHtml(g.estado || '')}</div>
          </div>
        </div>`).join('');
      html += '</div>';
    }
    if (productos.length) {
      html += `<div class="gs-section"><div class="gs-hdr">Productos</div>`;
      html += productos.map(p => `
        <div class="gs-item" onclick="GlobalSearch.cerrar();App.nav('productos')">
          <div class="gs-item-ico" style="background:#14532d">
            <svg width="14" height="14" viewBox="0 0 16 16" fill="#86efac">
              <path d="M1 2a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v2a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V2zm5 0a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v2a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V2z"/>
            </svg>
          </div>
          <div class="gs-item-main">
            <div class="gs-item-title">${escHtml(p.descripcion || '—')}</div>
            <div class="gs-item-sub">${escHtml(p.codigoBarra || p.idProducto || '')}</div>
          </div>
        </div>`).join('');
      html += '</div>';
    }
    if (preingresos.length) {
      html += `<div class="gs-section"><div class="gs-hdr">Pre-Ingresos</div>`;
      html += preingresos.map(pi => `
        <div class="gs-item" onclick="GlobalSearch.cerrar();App.nav('preingresos')">
          <div class="gs-item-ico" style="background:#78350f">
            <svg width="14" height="14" viewBox="0 0 16 16" fill="#fde68a">
              <path d="M4 0h5.293A1 1 0 0 1 10 .293L13.707 4a1 1 0 0 1 .293.707V14a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2z"/>
            </svg>
          </div>
          <div class="gs-item-main">
            <div class="gs-item-title">${escHtml(pi.idPreingreso || pi.proveedor || '—')}</div>
            <div class="gs-item-sub">${escHtml(pi.estado || '')}</div>
          </div>
        </div>`).join('');
      html += '</div>';
    }
    res.innerHTML = html;
  }

  // Cerrar con Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') cerrar();
  });

  return { abrir, cerrar, buscar };
})();

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
    // Skeleton cards en lugar de spinner central
    el.innerHTML = [1,2,3].map(() =>
      '<div class="skel skel-card"></div>'
    ).join('');
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

// Convierte string de fecha del sheet (yyyy-MM-dd) a Date local sin drift UTC
function _parseLocalDate(s) {
  if (!s) return new Date(NaN);
  const str = String(s);
  // Si ya tiene componente de hora, parsear directo; si solo tiene fecha yyyy-MM-dd, añadir mediodia local
  return str.length <= 10 ? new Date(str + 'T12:00:00') : new Date(str);
}

function fmtFecha(s) {
  if (!s) return '—';
  const d = _parseLocalDate(s);
  if (isNaN(d)) return s;
  return d.toLocaleDateString('es-PE', { day:'2-digit', month:'short', year:'numeric' });
}

// Fecha corta "19 abr" para cards
function _fmtCorta(s) {
  if (!s) return '—';
  const months = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const d = _parseLocalDate(s);
  if (isNaN(d)) return s;
  return `${d.getDate()} ${months[d.getMonth()]}`;
}

// Hora desde timestamp embebido en ID (ej: PI1745123456789 → "10:28")
function _horaDesdeId(id) {
  const ts = parseInt((id || '').replace(/\D/g, ''));
  if (!ts || ts < 1e12) return '';
  return new Date(ts).toLocaleTimeString('es-PE', { hour:'2-digit', minute:'2-digit', hour12: false });
}

// Escapa para insertar en atributos onclick="..." (evita romper comillas)
function escAttr(s) { return String(s || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '&quot;'); }
function escHtml(s)  { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// Hora desde campo fecha de guía — solo si tiene componente de hora explícito
function _horaDesdeGuia(g) {
  const f = String(g.fecha || '');
  // Solo usar si contiene hora (ISO 'T' o string con ':' y largo > 10)
  const tieneHora = f.includes('T') || (f.length > 10 && f.includes(':'));
  if (tieneHora) {
    const d = new Date(f);
    if (!isNaN(d)) return d.toLocaleTimeString('es-PE', { hour:'2-digit', minute:'2-digit', hour12: false });
  }
  return _horaDesdeId(g.idGuia);
}

// Parsea comentario → { comp: 'si'|'no'|null, compl: 'si'|'no'|null }
function _tagsFromComentario(comentario) {
  const s = String(comentario || '');
  const tags = { comp: null, compl: null };
  if (/comprobante:\s*sí/i.test(s))        tags.comp  = 'si';
  else if (/comprobante:\s*no/i.test(s))    tags.comp  = 'no';
  if (/completo:\s*sí/i.test(s))            tags.compl = 'si';
  else if (/completo:\s*no/i.test(s))       tags.compl = 'no';
  return tags;
}

// Extrae el texto libre quitando los prefijos de tags
function _textoLibreFromComentario(comentario) {
  return (comentario || '')
    .replace(/Comprobante:\s*(Sí|No)\s*\|?\s*/gi, '')
    .replace(/Completo:\s*(Sí|No)\s*\|?\s*/gi, '')
    .replace(/^\s*\|\s*/, '').replace(/\s*\|\s*$/, '')
    .trim();
}

// Construye string de comentario desde tags + texto libre
function _buildComentario(tags, textoExtra) {
  const partes = [];
  if (tags.comp)  partes.push(`Comprobante: ${tags.comp === 'si' ? 'Sí' : 'No'}`);
  if (tags.compl) partes.push(`Completo: ${tags.compl === 'si' ? 'Sí' : 'No'}`);
  const txt = (textoExtra || '').trim();
  if (txt) partes.push(txt);
  return partes.join(' | ');
}

function diasColor(dias) {
  if (dias <= 7)  return 'tag-danger';
  if (dias <= 30) return 'tag-warn';
  return 'tag-ok';
}

// ════════════════════════════════════════════════
// Carrusel de fotos (global — usado por Preingresos y futuras vistas)
// ════════════════════════════════════════════════
let _carFotos = [];
let _carIdx   = 0;

// Convierte URLs de Drive al formato de embed público con tamaño
// Extrae el fileId de cualquier formato de URL de Drive
function _driveFileId(url) {
  if (!url) return null;
  // lh3.googleusercontent.com/d/FILE_ID o FILE_ID=wXXX
  const lh3 = url.match(/lh3\.googleusercontent\.com\/d\/([^=?&/\s]+)/);
  if (lh3) return lh3[1];
  // ?id= o &id= (thumbnail, uc, etc.)
  const qid = url.match(/[?&]id=([^&\s]+)/);
  if (qid) return qid[1];
  // /file/d/FILE_ID/
  const fid = url.match(/\/file\/d\/([^/?&\s]+)/);
  if (fid) return fid[1];
  return null;
}

function _normalizeDriveUrl(url) {
  if (!url) return url;
  const id = _driveFileId(url);
  if (id) return `https://drive.google.com/thumbnail?id=${id}&sz=w800`;
  return url;
}

function abrirCarrusel(fotos, titulo, startIdx) {
  _carFotos = Array.isArray(fotos) ? fotos : String(fotos).split(',').filter(Boolean);
  _carIdx   = startIdx || 0;
  document.getElementById('carId').textContent = titulo || '';
  document.getElementById('photoCarousel').classList.remove('hidden');
  _renderCarrusel();
}

function cerrarCarrusel() {
  document.getElementById('photoCarousel').classList.add('hidden');
  _carFotos = [];
}

function carruselNav(dir) {
  if (!_carFotos.length) return;
  _carIdx = (_carIdx + dir + _carFotos.length) % _carFotos.length;
  _renderCarrusel();
}

function carruselGoTo(idx) {
  _carIdx = idx;
  _renderCarrusel();
}

function _renderCarrusel() {
  document.getElementById('carImg').src        = _normalizeDriveUrl(_carFotos[_carIdx] || '');
  document.getElementById('carIdx').textContent  = _carIdx + 1;
  document.getElementById('carTotal').textContent = _carFotos.length;
  const multi = _carFotos.length > 1;
  document.getElementById('carPrev').style.display = multi ? '' : 'none';
  document.getElementById('carNext').style.display = multi ? '' : 'none';
  document.getElementById('carThumbs').innerHTML = _carFotos.map((url, i) => `
    <div onclick="carruselGoTo(${i})"
         class="w-14 h-14 rounded-lg overflow-hidden flex-shrink-0 cursor-pointer border-2 transition-all"
         style="border-color:${i === _carIdx ? '#3b82f6' : 'transparent'};background:#1e293b">
      <img src="${_normalizeDriveUrl(url)}" class="w-full h-full object-cover" loading="lazy"/>
    </div>`).join('');
}

// ════════════════════════════════════════════════
// SESSION — Login, bloqueo, cierre de turno
// ════════════════════════════════════════════════
const Session = (() => {
  let pinBuffer = '';
  let lockPinBuffer = '';
  let sesionActual = null;
  let lockTimer = null;
  let lockInterval = null;
  let cierreReporte = null;
  const MIN_INACTIVIDAD = 5; // minutos (se sobreescribe desde config)

  function _hoy() {
    return new Date().toISOString().split('T')[0];
  }

  async function init() {
    const saved = _cargarSesion();

    if (saved) {
      const hoy      = _hoy();
      const fechaDia = saved.fechaDia || null;

      if (fechaDia === hoy) {
        // Misma jornada en este dispositivo → restaurar y bloquear
        sesionActual = saved;
        _aplicarSesion();
        bloquear();
        return;
      }

      if (fechaDia) {
        // Sesión de otro día quedó abierta → cerrar en GAS y pedir login
        const nombre = saved.nombre || 'usuario';
        if (navigator.onLine && saved.idSesion && !saved.idSesion.startsWith('LOCAL_')) {
          API.cerrarTurno({ idSesion: saved.idSesion, forzado: true }).catch(() => {});
        }
        _limpiarSesion();
        await mostrarLogin();
        setTimeout(() => toast(`⚠ La sesión de ${nombre} del ${fechaDia} quedó abierta. Se cerró automáticamente.`, 'warn', 7000), 800);
        return;
      }

      // Sesión legacy sin fechaDia → validar con GAS
      const res = await API.getSesionActiva(saved.idSesion).catch(() => ({ ok: false }));
      if (res.ok) {
        sesionActual = saved;
        _aplicarSesion();
        return;
      }
      _limpiarSesion();
    }

    await mostrarLogin();
  }

  async function mostrarLogin() {
    _ocultarApp();
    pinBuffer = '';
    _actualizarPuntos('pin', 0);
    document.getElementById('loginError').textContent = '';
    document.getElementById('loginScreen').style.display = 'flex';

    const loadingEl = document.getElementById('loginLoading');
    const padEl     = document.getElementById('pinPadArea');

    // Siempre ocultar el teclado al inicio y mostrar spinner
    padEl.style.display    = 'none';
    loadingEl.style.display = 'flex';

    const yaHayCache = OfflineManager.getPersonalCache().length > 0;
    if (!yaHayCache && navigator.onLine && window.WH_CONFIG.gasUrl) {
      // Sin caché: forzar descarga antes de mostrar teclado
      await OfflineManager.precargar(true).catch(() => {});
    }
    // Con caché: el timer de 60s se encarga del refresh — no disparar call extra aquí

    // Ocultar spinner, revelar teclado con animación
    loadingEl.style.display = 'none';
    padEl.classList.remove('fade-in-up');
    // Trigger reflow so animation replays even if already had the class
    void padEl.offsetWidth;
    padEl.classList.add('fade-in-up');
    padEl.style.display = 'flex';
  }

  function _setPinEnabled(on) {
    document.querySelectorAll('#pinPadArea .pin-btn').forEach(b => b.disabled = !on);
  }

  function pinTecla(d) {
    if (pinBuffer.length >= 4) return;
    pinBuffer += d;
    _actualizarPuntos('pin', pinBuffer.length);
    if (pinBuffer.length === 4) setTimeout(() => _intentarLogin(), 150);
  }

  function pinAtras() {
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
    const _dbgPersonal = OfflineManager.getPersonalCache();
    console.log('[Login] PIN ingresado:', pinIntento, '| Personal en caché:', _dbgPersonal.length, 'registros');
    if (_dbgPersonal.length > 0) console.log('[Login] Primer registro de muestra:', JSON.stringify(_dbgPersonal[0]).substring(0, 200));
    let localOp = OfflineManager.validarPinLocal(pinIntento);
    console.log('[Login] validarPinLocal result:', localOp ? localOp.nombre : null);

    // Sin caché → esperar GAS (nuevo dispositivo o sin datos locales)
    if (!localOp && navigator.onLine) {
      const res = await API.loginPersonal(pinIntento);
      if (!res.ok || res.offline) {
        document.getElementById('loginError').textContent = '❌ PIN incorrecto';
        setTimeout(() => { document.getElementById('loginError').textContent = ''; }, 2000);
        return;
      }
      sesionActual = res.data;
      _guardarSesion(sesionActual);
      document.getElementById('loginScreen').style.display = 'none';
      _aplicarSesion();
      if (res.data.yaEnSesionHoy) {
        // Continuación en nuevo dispositivo → pantalla de bloqueo
        bloquear();
      } else {
        _postLogin(res.data.sesionAnterior || null, res.data.bienvenidaImpresa === true);
      }
      return;
    }

    if (!localOp) {
      document.getElementById('loginError').textContent = '❌ PIN incorrecto';
      setTimeout(() => { document.getElementById('loginError').textContent = ''; }, 2000);
      return;
    }

    // 2. Sesión optimista inmediata con caché local
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
    _postLogin(null, true); // ticket se decide solo tras respuesta GAS

    // 3. Confirmar con GAS en segundo plano
    API.loginPersonal(pinIntento).then(res => {
      if (res.ok && !res.offline) {
        sesionActual = { ...sesionActual, ...res.data };
        window.WH_CONFIG.idSesion = sesionActual.idSesion;
        _guardarSesion(sesionActual);
        if (res.data.yaEnSesionHoy) {
          bloquear();
        } else {
          if (res.data.sesionAnterior) _mostrarAvisoSesionAnterior(res.data.sesionAnterior);
          if (!res.data.bienvenidaImpresa) _imprimirTicketBienvenida();
        }
      }
    }).catch(() => {});
  }

  // bienvenidaImpresa: true = ya se imprimió el ticket hoy → no imprimir
  function _postLogin(sesionAnterior, bienvenidaImpresa) {
    toast(`¡Hola ${sesionActual.nombre}! 👋`, 'ok', 2500);
    if (sesionAnterior) _mostrarAvisoSesionAnterior(sesionAnterior);
    if (!bienvenidaImpresa) _imprimirTicketBienvenida();
  }

  function _imprimirTicketBienvenida() {
    if (!navigator.onLine) return;
    API.imprimirBienvenida({
      nombre:     sesionActual.nombre,
      apellido:   sesionActual.apellido,
      rol:        sesionActual.rol,
      horaInicio: sesionActual.horaInicio
    }).catch(() => {});
  }

  function _mostrarAvisoSesionAnterior(fecha) {
    // Toast de advertencia + modal breve (igual que MosExpress)
    toast(`⚠ Sesión anterior del ${fecha} no fue cerrada`, 'warn', 6000);
  }

  function _aplicarSesion() {
    window.WH_CONFIG.usuario   = sesionActual.nombre + ' ' + sesionActual.apellido;
    window.WH_CONFIG.idSesion  = sesionActual.idSesion;

    // Avatar header (top bar)
    const av = document.getElementById('topAvatar');
    av.textContent   = sesionActual.nombre[0] + sesionActual.apellido[0];
    av.style.background = sesionActual.color;
    document.getElementById('usuarioNombre').textContent = sesionActual.nombre;
    // User menu v2: avatar + nombre completo
    const umAv = document.getElementById('umAvatar');
    if (umAv) { umAv.textContent = sesionActual.nombre[0] + sesionActual.apellido[0]; umAv.style.background = sesionActual.color; }

    // Avatar sidebar (tablet)
    const sideAv = document.getElementById('sideAvatar');
    if (sideAv) {
      sideAv.textContent   = sesionActual.nombre[0] + sesionActual.apellido[0];
      sideAv.style.background = sesionActual.color;
    }
    const sideNm = document.getElementById('sideUserName');
    if (sideNm) sideNm.textContent = sesionActual.nombre;
    const sideMnNm = document.getElementById('sideUserMenuName');
    if (sideMnNm) sideMnNm.textContent = sesionActual.nombre + ' ' + sesionActual.apellido;

    _mostrarApp();
    _iniciarTimerBloqueo();

    // Conectar indicador de estado online/offline/sync
    OfflineManager.onStatusChange(_actualizarEstadoHeader);
    _actualizarEstadoHeader({
      online:  navigator.onLine,
      pending: OfflineManager.getQueue().length,
      syncing: false
    });

    // Dashboard y maestros: usar caché si está disponible
    // precargar() ya está corriendo vía iniciarRefreshOperacional — no disparar otra llamada
    App.cargarDashboard();
    App.cargarProductosMaestro();
    App.cargarProveedoresMaestro();
    DespachoView.startPoll();
    DespachoView.badgeUpdate();

    // Si hay cola pendiente y hay red, sincronizar
    if (navigator.onLine) OfflineManager.sincronizar();
  }

  function _actualizarEstadoHeader({ online, pending, syncing }) {
    const dot    = document.getElementById('statusDot');
    const lbl    = document.getElementById('statusLabel');
    const bar    = document.getElementById('syncBar');
    const barLbl = document.getElementById('syncBarLabel');
    const umDot  = document.getElementById('umOnlineDot');
    const umLbl  = document.getElementById('umStatusLabel');

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

    if (dot) dot.style.background = color;
    if (lbl) lbl.textContent = texto;
    if (umDot) umDot.style.background = color;
    if (umLbl) umLbl.textContent = texto;

    // Sync bar (barra bajo topbar)
    if (bar) {
      if (syncing) {
        bar.className = 'show syncing';
        if (barLbl) barLbl.textContent = 'Sincronizando datos…';
      } else if (!online) {
        bar.className = 'show offline';
        if (barLbl) barLbl.textContent = pending > 0 ? `Sin conexión · ${pending} ops pendientes` : 'Sin conexión';
      } else if (pending > 0) {
        bar.className = 'show pending';
        if (barLbl) barLbl.textContent = `${pending} operaciones por sincronizar`;
      } else {
        bar.className = '';
      }
    }
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

  function _intentarDesbloqueo() {
    const pin = lockPinBuffer;
    lockPinBuffer = '';
    _actualizarPuntos('lpin', 0);

    // Validación 100% local — solo acepta el PIN del usuario activo
    const personal = OfflineManager.getPersonalCache();
    const ok = personal.find(p =>
      String(p.pin) === String(pin) && p.idPersonal === sesionActual.idPersonal
    );

    if (ok) {
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
    const isLock = prefix === 'lpin';
    for (let i = 0; i < 4; i++) {
      const el = document.getElementById(prefix + i);
      if (el) {
        el.className = i < n
          ? (isLock ? 'w-4 h-4 rounded-full bg-amber-500 transition-all' : 'w-4 h-4 rounded-full bg-sky-400 transition-all')
          : (isLock ? 'w-4 h-4 rounded-full border-2 border-amber-900 transition-all' : 'w-4 h-4 rounded-full border-2 border-slate-600 transition-all');
      }
    }
  }

  function cerrarSesionDesdeLock() {
    if (!confirm('¿Cerrar sesión? Se perderá el reporte de turno.')) return;
    clearTimeout(lockTimer);
    clearInterval(lockInterval);
    _limpiarSesion();
    sesionActual = null;
    document.getElementById('lockScreen').style.display = 'none';
    _ocultarApp();
    mostrarLogin();
  }

  function _guardarSesion(ses) {
    localStorage.setItem('wh_sesion', JSON.stringify({
      ...ses,
      fechaDia:      _hoy(),
      fechaGuardado: new Date().toISOString()
    }));
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

  return { init, mostrarLogin,
           pinTecla, pinAtras, lockTecla, lockAtras,
           bloquear, confirmarCierre, cerrarTurnoFinal,
           cerrarSesionDesdeLock,
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
    if (gasUrl) {
      console.log('[App] wh_gas_url desde localStorage:', gasUrl);
      window.WH_CONFIG.gasUrl = gasUrl;
    }
    console.log('[App] GAS URL activa:', window.WH_CONFIG.gasUrl);

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
    document.getElementById('auditConteo')?.addEventListener('input', e => {
      const sis = parseFloat(document.getElementById('audStockSis')?.textContent) || 0;
      const fis = parseFloat(e.target.value) || 0;
      const diff = fis - sis;
      document.getElementById('audDifValor').textContent = (diff >= 0 ? '+' : '') + fmt(diff, 2);
      document.getElementById('audDifValor').className = 'font-bold ' + (Math.abs(diff) < 0.5 ? 'text-emerald-400' : 'text-red-400');
      document.getElementById('audDiferenciaInfo').classList.remove('hidden');
    });

    // Cerrar filter dropdowns al tocar fuera
    document.addEventListener('click', e => {
      const guiaBtn  = document.getElementById('guiaFilterBtn');
      const guiaMenu = document.getElementById('guiaFilterMenu');
      if (guiaMenu && guiaBtn && !guiaBtn.contains(e.target) && !guiaMenu.contains(e.target)) {
        guiaMenu.style.display = 'none';
      }
      const preBtn  = document.getElementById('preFilterBtn');
      const preMenu = document.getElementById('preFilterMenu');
      if (preMenu && preBtn && !preBtn.contains(e.target) && !preMenu.contains(e.target)) {
        preMenu.style.display = 'none';
      }
      // Cerrar dropdown de proveedores al hacer click fuera
      const provDrop  = document.getElementById('preProvDrop');
      const provInput = document.getElementById('preProvInput');
      if (provDrop && provInput && !provInput.contains(e.target) && !provDrop.contains(e.target)) {
        provDrop.classList.add('hidden');
      }
    });

    // Precarga universal en background ANTES del login (30s cycle)
    OfflineManager.iniciarRefreshOperacional();

    // Escuchar refresh silencioso → actualizar vista activa sin flicker
    window.addEventListener('wh:data-refresh', e => {
      const changed = e.detail?.changed || [];
      const guiasChanged       = changed.includes('guias') || changed.includes('detalles');
      const preingresosChanged = changed.includes('preingresos');
      const stockChanged       = changed.includes('stock') || changed.includes('ajustes') || changed.includes('auditorias');
      const productosChanged   = changed.includes('productos') || changed.includes('equivalencias') || stockChanged;
      if (currentView === 'guias'       && guiasChanged)       GuiasView.silentRefresh();
      if (currentView === 'preingresos' && preingresosChanged) PreingresosView.silentRefresh();
      if (currentView === 'productos'   && productosChanged)   ProductosView.silentRefresh();
      if (currentView === 'envasador'   && stockChanged)       EnvasadorView.silentRefresh();
    });

    // Pull-to-refresh en la vista principal
    const mainContent = document.getElementById('mainContent');
    PullToRefresh.init(mainContent, () => {
      if (currentView === 'guias')        GuiasView.cargar();
      else if (currentView === 'productos')    ProductosView.cargar();
      else if (currentView === 'dashboard')    cargarDashboard();
      else if (currentView === 'preingresos')  PreingresosView.cargar?.();
    });

    // Vibración en navegación entre módulos
    document.querySelectorAll('.nav-btn').forEach(btn => {
      btn.addEventListener('touchstart', () => vibrate(6), { passive: true });
    });

    // Iniciar sesión (muestra login si no hay sesión activa)
    Session.init();
  }

  function nav(viewName) {
    // Si modo envasador activo y va a envasados → redirigir a envasador
    if (modoEnvasador && viewName === 'envasados') {
      viewName = 'envasador';
    }

    // Pausar cámara de despacho al salir de esa vista
    if (currentView === 'despacho' && viewName !== 'despacho') {
      DespachoView.pauseCamera();
    }

    closeUserMenu();

    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const el = document.getElementById('view-' + viewName);
    if (el) { el.classList.add('active'); el.classList.add('slide-up'); }

    // Marcar botón de nav activo por data-view
    document.querySelectorAll('.nav-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.view === viewName);
    });

    const titles = {
      dashboard:   'Dashboard',
      envasador:   'Modo Envasador',
      despacho:    'Despacho Rápido',
      guias:       'Guías',
      envasados:   'Envasados',
      preingresos: 'Pre-Ingresos',
      mermas:      'Mermas',
      auditorias:  'Auditorías',
      productos:   'Productos',
      proveedores: 'Proveedores',
      tools:       'Tools',
    };
    document.getElementById('pageTitle').textContent = titles[viewName] || viewName;

    currentView = viewName;

    // Lazy-load de cada vista
    switch (viewName) {
      case 'guias':       GuiasView.cargar(); break;
      case 'envasados':   EnvasadosView.cargar(); break;
      case 'envasador':   EnvasadorView.cargar(); break;
      case 'despacho':    DespachoView.cargar(); break;
      case 'preingresos': PreingresosView.cargar(); break;
      case 'mermas':      MermasView.cargar(); break;
      case 'productos':   ProductosView.cargar(); break;
      case 'tools':       _loadTools(); break;
    }
  }

  function toggleModoEnvasador() {
    modoEnvasador = !modoEnvasador;
    const ind = document.getElementById('modoIndicador');
    ind?.classList.toggle('hidden', !modoEnvasador);
    // Actualizar botón dentro de view-envasados
    const btn = document.getElementById('btnModo');
    if (btn) btn.innerHTML = modoEnvasador
      ? '✕ Salir modo'
      : '⚡ Modo Envasador';
    if (modoEnvasador) {
      nav('envasador');
      toast('Modo Envasador activado', 'ok');
    } else {
      nav('envasados');
      toast('Modo normal', 'info');
    }
  }

  // ── User menu (avatar dropdown) ───────────────────────────
  function toggleUserMenu() {
    const m = document.getElementById('userMenu');
    if (!m) return;
    m.classList.toggle('hidden');
    if (!m.classList.contains('hidden')) {
      setTimeout(() => document.addEventListener('click', _closeMenuOutside, { once: true }), 10);
    }
  }
  function closeUserMenu() {
    document.getElementById('userMenu')?.classList.add('hidden');
  }
  function _closeMenuOutside(e) {
    if (!document.getElementById('userMenu')?.contains(e.target)) closeUserMenu();
  }

  // ── Menú de usuario del sidebar (tablet) ─────────────────
  function toggleSideUserMenu() {
    const m = document.getElementById('sideUserMenu');
    if (!m) return;
    const isOpen = m.classList.contains('open');
    m.classList.toggle('open', !isOpen);
    // Girar el chevron: ↑ cuando cerrado, ↓ cuando abierto
    const ch = document.getElementById('sideChevron');
    if (ch) ch.style.transform = !isOpen ? 'rotate(180deg)' : '';
    if (!isOpen) {
      setTimeout(() => document.addEventListener('click', _closeSideMenuOutside, { once: true }), 10);
    }
  }
  function closeSideUserMenu() {
    document.getElementById('sideUserMenu')?.classList.remove('open');
    const ch = document.getElementById('sideChevron');
    if (ch) ch.style.transform = '';
  }
  function _closeSideMenuOutside(e) {
    const card = document.getElementById('sideUserCard');
    if (card && !card.contains(e.target)) closeSideUserMenu();
  }

  // ── Tools view ────────────────────────────────────────────
  function _loadTools() {
    fetch('./version.json').then(r => r.json()).then(v => {
      const el = document.getElementById('toolsVersion');
      if (el) el.textContent = v.version + ' (' + (v.build || '') + ')';
    }).catch(() => {});
    const gasEl = document.getElementById('toolsGasUrl');
    if (gasEl) gasEl.textContent = window.WH_CONFIG?.gasUrl || '—';
  }

  async function syncForzado() {
    const btn = document.getElementById('btnSyncForzado');
    const st  = document.getElementById('syncStatus');
    if (btn) { btn.disabled = true; btn.textContent = 'Sincronizando…'; }
    if (st)  st.textContent = '';
    try {
      await OfflineManager.sincronizar();
      await OfflineManager.precargar();
      if (st) st.textContent = '✅ Sincronizado ' + new Date().toLocaleTimeString('es-PE');
      toast('Sincronización completada', 'ok');
    } catch(e) {
      if (st) st.textContent = '❌ Error: ' + e.message;
      toast('Error al sincronizar', 'danger');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor"><path fill-rule="evenodd" d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2v1z"/><path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466z"/></svg> Sincronizar ahora'; }
    }
  }

  async function checkUpdate() {
    const st = document.getElementById('syncStatus');
    if (st) st.textContent = 'Buscando actualización…';
    if (window._SWCheck) {
      await window._SWCheck();
      if (st) st.textContent = 'Verificado ' + new Date().toLocaleTimeString('es-PE');
    } else {
      if (st) st.textContent = 'Service Worker no disponible en este entorno';
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
    const enAlerta = alertas.vencimientosAlertas || [];
    const pendEnv  = alertas.pendientesEnvasado  || [];
    const stockBajo = alertas.stockBajoMinimo    || [];
    const mermasPend = alertas.mermasPendientes  || [];

    // KPIs principales (reemplaza skeletons con valores reales)
    document.getElementById('kpiCriticos').textContent   = contadores.criticos ?? criticos.length;
    document.getElementById('kpiPendEnv').textContent    = pendEnv.length;
    document.getElementById('kpiStockBajo').textContent  = stockBajo.length;
    document.getElementById('kpiMermas').textContent     = fmt(kpis.mermasTotalMes ?? 0, 1);

    // KPIs secundarios
    document.getElementById('kpiEficiencia').textContent = (kpis.eficienciaEnvasadoPct ?? '—') + '%';
    document.getElementById('kpiSalidas').textContent    = kpis.salidasUltimos30dias ?? '—';

    // Logo alert dot (topbar + sidebar)
    const totalAlertas = contadores.alertasTotal ?? 0;
    document.getElementById('logoAlertDot')?.classList.toggle('hidden', totalAlertas === 0);
    const sideLogoAlert = document.getElementById('sideLogoAlertDot');
    const sideArrow     = document.getElementById('sideAlertArrow');
    if (sideLogoAlert) sideLogoAlert.classList.toggle('visible', totalAlertas > 0);
    if (sideArrow)     sideArrow.classList.toggle('visible',     totalAlertas > 0);

    // Badges en nav: guías abiertas
    const guiasAbiertas = (alertas.guiasAbiertas ?? contadores.guiasAbiertas ?? 0);
    actualizarBadges({ guiasAbiertas });

    // Historial rápido (últimas guías del caché)
    const histEl = document.getElementById('historialRapido');
    if (histEl) {
      const guiasRecientes = (OfflineManager.getGuiasCache() || []).slice(0, 6);
      if (guiasRecientes.length) {
        const TIPO_SHORT = {
          INGRESO_PROVEEDOR:'Ingreso Prov.', INGRESO_JEFATURA:'Ing. Jefatura',
          SALIDA_ZONA:'Salida Zona', SALIDA_DEVOLUCION:'Devolución',
          SALIDA_JEFATURA:'Sal. Jefatura', SALIDA_ENVASADO:'Envasado', SALIDA_MERMA:'Merma'
        };
        histEl.innerHTML = guiasRecientes.map(g => `
          <div class="hist-item" onclick="App.nav('guias')">
            <div class="hist-type">${TIPO_SHORT[g.tipo] || g.tipo || 'Guía'}</div>
            <div class="hist-desc">${escHtml(g.idGuia || '—')}</div>
            <div class="hist-time">${escHtml(g.estado || '')} · ${_fmtCorta(g.fecha)}</div>
          </div>`).join('');
      } else {
        histEl.innerHTML = '<p class="text-slate-500 text-xs py-2 px-2">Sin registros recientes</p>';
      }
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
    // Usar caché primero — precargar() ya trae productos en background
    const cached = OfflineManager.getProductosCache();
    if (cached.length) { todosProductos = cached; return; }
    const res = await API.getProductos({ estado: '1' }).catch(() => ({ ok: false }));
    if (res.ok) todosProductos = res.data;
  }

  async function cargarProveedoresMaestro() {
    // Usar caché primero — precargar() ya trae proveedores en background
    const cached = OfflineManager.getProveedoresCache();
    const lista  = cached.length ? cached : await API.getProveedores({ estado: '1' })
                     .then(r => r.ok ? r.data : []).catch(() => []);
    if (!lista.length) return;
    todosProveedores = lista;
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

  function showUsuarioDialog() {
    const nombre = prompt('Nombre del operador:', window.WH_CONFIG.usuario);
    if (nombre) {
      window.WH_CONFIG.usuario = nombre;
      localStorage.setItem('wh_usuario', nombre);
      document.getElementById('usuarioNombre').textContent = nombre;
    }
  }

  function getProductosMaestro() { return todosProductos; }
  function getProveedoresMaestro() { return todosProveedores; }

  function abrirMas() { abrirSheet('sheetMas'); }
  function navMas(viewName) { cerrarSheet('sheetMas'); nav(viewName); }

  return { init, nav, abrirMas, navMas,
           toggleModoEnvasador,
           toggleUserMenu, closeUserMenu,
           toggleSideUserMenu, closeSideUserMenu,
           syncForzado, checkUpdate,
           instalarPWA: () => window._installPWA?.(),
           cargarDashboard, showUsuarioDialog,
           cargarProductosMaestro, cargarProveedoresMaestro,
           getProductosMaestro, getProveedoresMaestro,
           getView: () => currentView };
})();

// ════════════════════════════════════════════════
// GUIAS VIEW
// ════════════════════════════════════════════════
const GuiasView = (() => {
  let todas = [];
  let filtroActual = '';
  let _busquedaQ   = '';
  let _guiaActual  = null;   // guía abierta en el sheet de detalle
  let _refreshDot  = null;   // indicador visual de refresh
  // Foto guía (una sola foto por guía)
  let _fotoGuiaNueva = null; // { file, objectUrl }
  // Comentario + tags guía (estado del sheet de detalle — se guardan al cerrar)
  let _tagsGuia    = { comp: null, compl: null };
  // Tags para CREAR/EDITAR guía
  let _tagsNueva   = { comp: null, compl: null };
  // Paneles del header de detalle
  let _fotoOpen    = false;
  let _notasOpen   = false;
  // Modo edición de guía existente
  let _guiaModoEdicion = false;
  // Agregar ítem: estado del scanner+form
  let _itemProd    = null;   // product object seleccionado
  let _itemQty     = 1;
  let _itemVenc    = '';

  const TIPO_LABELS = {
    INGRESO_PROVEEDOR: 'Proveedor', INGRESO_JEFATURA: 'Jefatura',
    SALIDA_ZONA: 'Zona',  SALIDA_DEVOLUCION: 'Devolución',
    SALIDA_JEFATURA: 'Jefatura', SALIDA_ENVASADO: 'Envasado', SALIDA_MERMA: 'Merma'
  };

  // Carga inicial: primero desde caché (instantáneo), luego refresca en bg
  async function cargar() {
    const cached = OfflineManager.getGuiasCache();
    if (cached.length) {
      todas = cached;
      render(_filtrarYBuscar());
    } else {
      loading('listGuias', true);
    }
    // Refresca en background (la precarga operacional ya está corriendo,
    // pero aquí forzamos un fetch inmediato para la primera entrada a la vista)
    OfflineManager.precargarOperacional().then(() => {
      const fresh = OfflineManager.getGuiasCache();
      if (fresh.length) { todas = fresh; render(_filtrarYBuscar()); }
    });
  }

  // Refresh silencioso desde el evento 60s — no muestra spinner
  function silentRefresh() {
    const fresh = OfflineManager.getGuiasCache();
    if (!fresh.length) return;
    todas = fresh;
    render(_filtrarYBuscar());
    // Parpadeo sutil del indicador
    const dot = document.getElementById('guiasRefreshDot');
    if (dot) { dot.style.opacity = '1'; setTimeout(() => { dot.style.opacity = '0'; }, 1200); }
  }

  function _filtrar(list, f) {
    if (!f || f === 'TODAS') return list;
    if (f === 'INGRESO') return list.filter(g => g.tipo?.startsWith('INGRESO'));
    if (f === 'SALIDA')  return list.filter(g => g.tipo?.startsWith('SALIDA'));
    if (f === 'ABIERTA') return list.filter(g => g.estado === 'ABIERTA');
    return list;
  }

  function _filtrarYBuscar() {
    let r = _filtrar(todas, filtroActual);
    if (_busquedaQ) {
      const qL = _busquedaQ.toLowerCase();
      r = r.filter(g => {
        const provNombre = _getProvNombre(g.idProveedor).toLowerCase();
        return (g.idGuia         || '').toLowerCase().includes(qL) ||
               (g.idProveedor    || '').toLowerCase().includes(qL) ||
               provNombre.includes(qL) ||
               (g.numeroDocumento|| '').toLowerCase().includes(qL) ||
               (TIPO_LABELS[g.tipo] || g.tipo || '').toLowerCase().includes(qL);
      });
    }
    return r;
  }

  function buscar(q) {
    _busquedaQ = (q || '').trim();
    const cl = document.getElementById('clearBuscarGuia');
    if (cl) cl.style.display = _busquedaQ ? 'flex' : 'none';
    const lista = _filtrarYBuscar();
    render(lista);

    if (!_busquedaQ) return;
    const qL = _busquedaQ.toLowerCase();

    // Detectar coincidencia exacta: idGuia, idProveedor o nombre exacto de proveedor
    const exacto = lista.find(g =>
      (g.idGuia      || '').toLowerCase() === qL ||
      (g.idProveedor || '').toLowerCase() === qL ||
      _getProvNombre(g.idProveedor).toLowerCase() === qL
    );

    requestAnimationFrame(() => {
      if (exacto) {
        const cardId = 'guia-' + (exacto.idGuia || '').replace(/[^a-zA-Z0-9_-]/g, '_');
        const card   = document.getElementById(cardId);
        if (card) {
          card.classList.add('card-exact-match');
          card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        lista.forEach(g => {
          if (g !== exacto) {
            const el = document.getElementById('guia-' + (g.idGuia || '').replace(/[^a-zA-Z0-9_-]/g, '_'));
            if (el) el.classList.add('card-dim');
          }
        });
      } else {
        lista.forEach(g => {
          const el = document.getElementById('guia-' + (g.idGuia || '').replace(/[^a-zA-Z0-9_-]/g, '_'));
          if (el) el.classList.add('card-hi');
        });
      }
    });
  }

  function buscarClear() {
    _busquedaQ = '';
    const inp = document.getElementById('inputBuscarGuia');
    if (inp) inp.value = '';
    const cl = document.getElementById('clearBuscarGuia');
    if (cl) cl.style.display = 'none';
    render(_filtrarYBuscar());
  }

  function toggleFiltro() {
    const menu = document.getElementById('guiaFilterMenu');
    menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
  }

  function _cerrarFiltroMenu() {
    const menu = document.getElementById('guiaFilterMenu');
    if (menu) menu.style.display = 'none';
  }

  const FILTRO_LABELS = { TODAS: 'TODAS', INGRESO: '↓ INGRESOS', SALIDA: '↑ SALIDAS', ABIERTA: '◌ ABIERTAS' };

  function filtrar(f) {
    filtroActual = f || 'TODAS';
    // Update label
    const lbl = document.getElementById('guiaFilterLabel');
    if (lbl) lbl.textContent = FILTRO_LABELS[filtroActual] || 'TODAS';
    // Update active state in dropdown
    document.querySelectorAll('.guia-fopt').forEach(b =>
      b.classList.toggle('sel', b.dataset.filtro === filtroActual));
    _cerrarFiltroMenu();
    render(_filtrarYBuscar());
  }

  function _getProvNombre(idProveedor) {
    if (!idProveedor) return '';
    const p = OfflineManager.getProveedoresCache().find(x => x.idProveedor === idProveedor);
    return p ? (p.nombre || idProveedor) : idProveedor;
  }

  function _renderGuiaCard(g) {
    const isEnvasado  = g.tipo === 'SALIDA_ENVASADO' || g.tipo === 'INGRESO_ENVASADO';
    const isIngreso   = g.tipo?.startsWith('INGRESO');
    const isAbierta   = g.estado === 'ABIERTA';
    const borderColor = isEnvasado ? '#475569' : isAbierta ? '#f59e0b' : isIngreso ? '#22c55e' : '#3b82f6';
    const tipoLabel   = TIPO_LABELS[g.tipo] || g.tipo || '—';
    const provNombre  = _getProvNombre(g.idProveedor) || g.usuario || '—';
    const hora        = _horaDesdeGuia(g);
    const fechaCorta  = _fmtCorta(g.fecha);

    // Guías de envasado: card gris, solo lectura, sin botones de acción
    if (isEnvasado) {
      const detCache = OfflineManager.getGuiaDetalleCache();
      const numItems = detCache.filter(d => d.idGuia === g.idGuia && d.observacion !== 'ANULADO').length;
      const itemsTag = numItems > 0 ? ` <span style="color:#64748b">[${numItems}]</span>` : '';
      return `
      <div class="guia-card" id="guia-${(g.idGuia||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
           style="border-left-color:#475569;opacity:.7;cursor:default"
           onclick="GuiasView.verDetalle('${escAttr(g.idGuia)}')">
        <div class="flex items-center justify-between gap-1 overflow-hidden">
          <span class="text-xs font-bold truncate" style="color:#64748b">${tipoLabel}</span>
          <span style="font-size:10px;color:#475569;flex-shrink:0">🔒 sistema</span>
        </div>
        <p class="text-sm font-semibold truncate" style="color:#64748b">${escAttr(g.usuario || '—')}</p>
        <p class="text-xs" style="color:#475569">${fechaCorta}${itemsTag}</p>
      </div>`;
    }

    const estadoDot  = isAbierta
      ? `<span style="width:6px;height:6px;border-radius:50%;background:#f59e0b;display:inline-block;flex-shrink:0"></span>`
      : `<span style="width:6px;height:6px;border-radius:50%;background:#22c55e;display:inline-block;flex-shrink:0"></span>`;
    const fotoTag = g.foto ? `<span class="pre-qtag pre-qtag-slate">📷</span>` : '';
    // Icono de preingreso vinculado — al tap navega al preingreso
    const preTag  = g.idPreingreso
      ? `<span onclick="event.stopPropagation();GuiasView.irAPreingreso('${escAttr(g.idPreingreso)}')"
               class="pre-qtag pre-qtag-blue" title="Ver preingreso"
               style="cursor:pointer;user-select:none">📋</span>`
      : '';
    // Cantidad de ítems desde la caché de detalle
    const detCache = OfflineManager.getGuiaDetalleCache();
    const numItems = detCache.filter(d => d.idGuia === g.idGuia && d.observacion !== 'ANULADO').length;
    const itemsTag = numItems > 0 ? ` <span class="text-slate-500">[${numItems}]</span>` : '';
    const pnPend   = (OfflineManager.getPNCache() || []).filter(p => p.idGuia === g.idGuia && p.estado === 'PENDIENTE').length;
    const pnBadge  = pnPend ? `<span style="background:#78350f;color:#fde68a;font-size:9px;font-weight:800;
      padding:1px 5px;border-radius:4px;flex-shrink:0;letter-spacing:.04em;cursor:pointer"
      onclick="event.stopPropagation();GuiasView.abrirModalPN('',${JSON.stringify(g.idGuia)})"
      title="${pnPend} producto(s) nuevo(s) pendiente(s)">N${pnPend > 1 ? ' ' + pnPend : ''}</span>` : '';
    const waGuiaBtn = `<button onclick="event.stopPropagation();GuiasView.compartirWA('${escAttr(g.idGuia)}')"
      style="background:none;border:none;cursor:pointer;padding:2px;color:#25d366;flex-shrink:0;line-height:0"
      title="Compartir por WhatsApp">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
    </button>`;
    const printGuiaBtn = `<button onclick="event.stopPropagation();GuiasView.imprimirTicket('${escAttr(g.idGuia)}')"
      style="background:none;border:none;cursor:pointer;padding:2px;color:#94a3b8;flex-shrink:0;line-height:0"
      title="Imprimir ticket">
      <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M2.5 8a.5.5 0 1 0 0-1 .5.5 0 0 0 0 1z"/><path d="M5 1a2 2 0 0 0-2 2v2H2a2 2 0 0 0-2 2v3a2 2 0 0 0 2 2h1v1a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2v-1h1a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-1V3a2 2 0 0 0-2-2H5zM4 3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2H4V3zm1 5a2 2 0 0 0-2 2v1H2a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v-1a2 2 0 0 0-2-2H5zm7 2v3a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1z"/></svg>
    </button>`;
    return `
    <div class="guia-card" id="guia-${(g.idGuia||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
         style="border-left-color:${borderColor}"
         onclick="GuiasView.verDetalle('${escAttr(g.idGuia)}')">
      <div class="flex items-center justify-between gap-1 overflow-hidden">
        <span class="text-xs font-bold truncate" style="color:${isIngreso ? '#4ade80' : '#60a5fd'}">${tipoLabel}</span>
        <div class="flex items-center gap-1 flex-shrink-0">${pnBadge}${preTag}${fotoTag}${waGuiaBtn}${printGuiaBtn}${estadoDot}</div>
      </div>
      <p class="text-sm font-bold text-slate-100 truncate">${escAttr(provNombre)}</p>
      <p class="text-xs text-slate-400">${fechaCorta}${hora ? ' · ' + hora : ''}${itemsTag}</p>
    </div>`;
  }

  function render(list) {
    const container = document.getElementById('listGuias');
    if (!container) return;
    const optCards = Array.from(container.querySelectorAll('.card-optimistic'));
    if (!list.length) {
      container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">No hay guías</p>';
      optCards.forEach(c => container.insertBefore(c, container.firstChild));
      _renderTabletPrePanel();
      return;
    }

    const sorted = [...list].sort((a, b) => {
      const da = _parseLocalDate(a.fecha), db = _parseLocalDate(b.fecha);
      const td = db - da;
      if (td !== 0) return td;
      const na = parseInt((a.idGuia || '').replace(/\D/g, '')) || 0;
      const nb = parseInt((b.idGuia || '').replace(/\D/g, '')) || 0;
      return nb - na;
    });

    const today     = new Date(); today.setHours(0,0,0,0);
    const yesterday = new Date(today.getTime()); yesterday.setDate(today.getDate() - 1);
    const months    = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    function _gKey(g) {
      if (!g.fecha) return '0000-00-00';
      const d = _parseLocalDate(g.fecha);
      if (isNaN(d)) return '0000-00-00';
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }
    function _gLabel(key) {
      if (!key || key === '0000-00-00') return 'Sin fecha';
      const d = new Date(key + 'T12:00:00');
      const dMid = new Date(d); dMid.setHours(0,0,0,0);
      if (dMid.getTime() === today.getTime())     return 'Hoy';
      if (dMid.getTime() === yesterday.getTime()) return 'Ayer';
      return d.getFullYear() === today.getFullYear()
        ? `${d.getDate()} ${months[d.getMonth()]}`
        : `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    }

    const groupMap = {};
    sorted.forEach(g => {
      const k = _gKey(g);
      if (!groupMap[k]) groupMap[k] = [];
      groupMap[k].push(g);
    });

    container.innerHTML = Object.entries(groupMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([key, items]) =>
        `<div class="pre-date-hdr">${_gLabel(key)}</div>
         <div class="pre-date-group">${items.map(_renderGuiaCard).join('')}</div>`
      ).join('');

    // Preservar solo cards optimistas cuyo ID aún no está en la lista real
    optCards.forEach(c => {
      const rid = c.getAttribute('data-real-id') || c.id.replace('optguia_', '');
      if (!sorted.find(g => g.idGuia === rid)) {
        container.insertBefore(c, container.firstChild);
      }
    });

    _renderTabletPrePanel();
  }

  // ── Tablet pre-ingreso panel — delega al renderer real de PreingresosView ──
  function _renderTabletPrePanel() {
    PreingresosView.renderTppList();
  }

  // ── Optimistic guía card ──────────────────────────────────
  function injectOptimisticGuia({ tempId, idProveedor, provNombre }) {
    const container = document.getElementById('listGuias');
    if (!container) return;
    const div = document.createElement('div');
    div.id = 'optguia_' + tempId;
    div.className = 'guia-card card-optimistic';
    div.style.borderLeftColor = '#22c55e';
    div.innerHTML = `
      <div class="flex items-center gap-2">
        <span class="text-xs font-bold text-emerald-400">Proveedor</span>
      </div>
      <div class="flex items-center gap-2">
        <div class="spinner" style="width:12px;height:12px;border-width:2px"></div>
        <span class="text-xs text-slate-400 italic">Creando guía…</span>
      </div>
      <p class="text-xs text-slate-400 truncate">${escAttr(provNombre || 'Sin proveedor')}</p>`;
    container.insertBefore(div, container.firstChild);
  }

  function finalizeOptimisticGuia(tempId, idGuia, tipo, provNombre) {
    const el = document.getElementById('optguia_' + tempId);
    if (el) {
      el.setAttribute('data-real-id', idGuia || tempId);
      el.style.animation = 'none';
      el.style.borderLeftColor = '#22c55e';
      const esI = (tipo || '').startsWith('INGRESO');
      const label = TIPO_LABELS[tipo] || tipo || '';
      el.innerHTML = `
        <div class="flex items-center justify-between gap-1">
          <span class="text-xs ${esI ? 'tag-ok' : 'tag-blue'}">${escAttr(label)}</span>
          <span class="text-xs text-emerald-400 font-bold">ABIERTA</span>
        </div>
        <p class="text-sm font-bold text-slate-100 truncate mt-1">${escAttr(provNombre || 'Sin proveedor')}</p>
        <p class="text-xs text-slate-500 font-mono">${escAttr(idGuia || '')}</p>`;
    }
    // Refrescar en background — render() descartará el opt card cuando el ID real llegue
    setTimeout(() => {
      OfflineManager.precargarOperacional().then(() => {
        const fresh = OfflineManager.getGuiasCache();
        if (fresh.length) { todas = fresh; render(_filtrarYBuscar()); }
      });
    }, 800);
  }

  function removeOptimisticGuia(tempId) {
    document.getElementById('optguia_' + tempId)?.remove();
  }

  // Abre el detalle desde caché instantáneamente
  function verDetalle(idGuia) {
    // 1. Buscar en caché local (instantáneo)
    const guias    = OfflineManager.getGuiasCache();
    const detalles = OfflineManager.getGuiaDetalleCache();
    const prods    = OfflineManager.getProductosCache();
    const prodMap  = {};
    prods.forEach(p => { prodMap[p.idProducto] = p.descripcion || p.nombre || p.idProducto; });

    let guia = guias.find(g => g.idGuia === idGuia);
    if (!guia) {
      // Fallback: mostrar loading y pedir a GAS
      _abrirDetalleConGAS(idGuia);
      return;
    }

    const detalle = detalles
      .filter(d => d.idGuia === idGuia)
      .map(d => ({ ...d, descripcionProducto: prodMap[d.codigoProducto] || d.codigoProducto }));

    _guiaActual = { ...guia, detalle };
    _mostrarDetalleSheet(_guiaActual);

    // 2. Refrescar desde GAS en background (actualiza si hay cambios)
    if (navigator.onLine) {
      API.getGuia(idGuia).then(res => {
        if (res.ok && !res.offline) {
          _guiaActual = res.data;
          _mostrarDetalleSheet(_guiaActual, false); // re-render sin animación
          // Mantener cache local en sync para reaperturas rápidas
          if (Array.isArray(_guiaActual.detalle)) {
            OfflineManager.actualizarDetallesGuia(idGuia, _guiaActual.detalle);
          }
        }
      }).catch(() => {});
    }
  }

  async function _abrirDetalleConGAS(idGuia) {
    document.getElementById('guiaDetHeader').innerHTML =
      '<div class="flex justify-center py-4"><div class="spinner"></div></div>';
    abrirSheet('sheetGuiaDetalle');
    const res = await API.getGuia(idGuia);
    if (!res.ok) { toast('Error al cargar guía', 'danger'); cerrarSheet('sheetGuiaDetalle'); return; }
    _guiaActual = res.data;
    _mostrarDetalleSheet(_guiaActual, false);
  }

  // SVG lock icons
  const SVG_LOCK_OPEN   = `<svg width="18" height="18" viewBox="0 0 16 16" fill="currentColor"><path d="M11 1a2 2 0 0 0-2 2v4a2 2 0 0 0 2 2h3a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2h5V3a3 3 0 0 1 6 0v4a.5.5 0 0 1-1 0V3a2 2 0 0 0-2-2zM5 9a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1v-5a1 1 0 0 0-1-1H5z"/></svg>`;
  const SVG_LOCK_CLOSED = `<svg width="18" height="18" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a2 2 0 0 1 2 2v4H6V3a2 2 0 0 1 2-2zm3 6V3a3 3 0 0 0-6 0v4a2 2 0 0 0-2 2v5a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2zM5 8h6v5H5z"/></svg>`;

  function _mostrarDetalleSheet(g, conAnimacion = true) {
    const esIngreso   = g.tipo?.startsWith('INGRESO');
    const esEnvasado  = g.tipo === 'SALIDA_ENVASADO' || g.tipo === 'INGRESO_ENVASADO';
    const abierta     = g.estado === 'ABIERTA';
    const esDiaAnterior = g.fecha && g.fecha < new Date().toISOString().split('T')[0];

    // Guías de envasado: header bloqueado, sin acciones
    if (esEnvasado) {
      document.getElementById('guiaDetHeader').innerHTML = `
        <div class="flex items-start justify-between gap-2 mb-1">
          <span class="text-xs" style="color:#64748b;background:#1e293b;border-radius:4px;padding:2px 6px">
            ${TIPO_LABELS[g.tipo] || g.tipo}
          </span>
          <span style="font-size:11px;color:#475569;background:#1e293b;border-radius:4px;padding:2px 8px">
            🔒 Generado por sistema
          </span>
        </div>
        <p class="font-bold text-base leading-tight" style="color:#64748b">${escAttr(g.usuario || '—')}</p>
        <p class="text-xs mt-0.5" style="color:#475569">${fmtFecha(g.fecha)} · Solo lectura</p>`;
      // Ocultar panel de foto/notas/editar y el add-item
      document.getElementById('guiaDetFotoPanel')?.classList.add('hidden');
      document.getElementById('guiaDetNotasPanel')?.classList.add('hidden');
      document.getElementById('guiaDetAddItem')?.classList.add('hidden');
    } else {

    // Lock button
    const lockBtn = `
      <button onclick="GuiasView.toggleEstadoGuia()"
              class="flex items-center gap-1 px-3 py-1 rounded-lg border font-bold text-xs tracking-wide transition-colors
                     ${abierta ? 'border-amber-700 text-amber-300 bg-amber-900/30 hover:bg-amber-800/40'
                               : 'border-slate-600 text-slate-400 hover:bg-slate-700'}"
              title="${abierta ? 'Cerrar guía' : 'Reabrir (admin)'}">
        ${abierta ? SVG_LOCK_OPEN : SVG_LOCK_CLOSED}
        ${abierta ? 'ABIERTA' : 'CERRADA'}
      </button>`;

    const provNombreHdr = (() => {
      if (!g.idProveedor) return 'Sin proveedor';
      const pv = OfflineManager.getProveedoresCache().find(p => p.idProveedor === g.idProveedor);
      return pv ? (pv.nombre || g.idProveedor) : g.idProveedor;
    })();

    document.getElementById('guiaDetHeader').innerHTML = `
      <div class="flex items-start justify-between gap-2 mb-1" onclick="GuiasView.deselectItem()">
        <span class="text-xs ${esIngreso ? 'tag-ok' : 'tag-blue'}">${TIPO_LABELS[g.tipo] || g.tipo}</span>
        <span onclick="event.stopPropagation()">${lockBtn}</span>
      </div>
      <p class="font-black text-lg text-white leading-tight" onclick="GuiasView.deselectItem()">${escAttr(provNombreHdr)}</p>
      <p class="text-xs text-slate-500 mt-0.5" onclick="GuiasView.deselectItem()">${fmtFecha(g.fecha)} · ${g.usuario || '—'}</p>
      ${esDiaAnterior && abierta ? `<p class="text-xs text-red-400 mt-1 font-bold">⚠ Guía del día anterior — ciérrala pronto</p>` : ''}
      <div class="flex gap-2 mt-2 mb-1">
        <button onclick="GuiasView.toggleFotoPanel()" id="btnHdrFoto"
                class="flex items-center gap-1.5 py-1.5 px-3 rounded-lg text-xs transition-colors ${_fotoOpen ? 'bg-blue-700/60 text-blue-200' : 'bg-slate-800 text-slate-400'}">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
            <path d="M10.5 8.5a2.5 2.5 0 1 1-5 0 2.5 2.5 0 0 1 5 0z"/>
            <path d="M2 4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2h-1.172a2 2 0 0 1-1.414-.586l-.828-.828A2 2 0 0 0 9.172 2H6.828a2 2 0 0 0-1.414.586l-.828.828A2 2 0 0 1 3.172 4H2zm.5 2a.5.5 0 1 1 0-1 .5.5 0 0 1 0 1zm9 2.5a3.5 3.5 0 1 1-7 0 3.5 3.5 0 0 1 7 0z"/>
          </svg>
          Foto${g.foto ? ' ✓' : ''}
        </button>
        <button onclick="GuiasView.toggleNotasPanel()" id="btnHdrNotas"
                class="flex items-center gap-1.5 py-1.5 px-3 rounded-lg text-xs transition-colors ${_notasOpen ? 'bg-blue-700/60 text-blue-200' : 'bg-slate-800 text-slate-400'}">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
            <path d="M0 2a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H4.414a1 1 0 0 0-.707.293L.854 15.146A.5.5 0 0 1 0 14.793V2zm3.5 1a.5.5 0 0 0 0 1h9a.5.5 0 0 0 0-1h-9zm0 2.5a.5.5 0 0 0 0 1h9a.5.5 0 0 0 0-1h-9zm0 2.5a.5.5 0 0 0 0 1h5a.5.5 0 0 0 0-1h-5z"/>
          </svg>
          Notas${g.comentario ? ' ✓' : ''}
        </button>
        ${abierta ? `
        <button onclick="GuiasView.editarGuia()" class="flex items-center gap-1.5 py-1.5 px-3 rounded-lg text-xs bg-slate-800 text-slate-400 transition-colors ml-auto">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
            <path d="M12.146.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1 0 .708l-10 10a.5.5 0 0 1-.168.11l-5 2a.5.5 0 0 1-.65-.65l2-5a.5.5 0 0 1 .11-.168l10-10zM11.207 2.5 13.5 4.793 14.793 3.5 12.5 1.207 11.207 2.5zm1.586 3L10.5 3.207 4 9.707V10h.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.5h.293l6.5-6.5zm-9.761 5.175-.106.106-1.528 3.821 3.821-1.528.106-.106A.5.5 0 0 1 5 12.5V12h-.5a.5.5 0 0 1-.5-.5V11h-.5a.5.5 0 0 1-.468-.325z"/>
          </svg>
          Editar
        </button>` : ''}
      </div>`;

    // ── Foto panel (toggle) ────────────────────────────────
    const fotoEl = document.getElementById('guiaDetFotoSection');
    if (fotoEl) {
      if (!_fotoOpen) {
        fotoEl.innerHTML = '';
      } else if (g.foto) {
        fotoEl.innerHTML = `
          <div class="relative rounded-lg overflow-hidden mb-3" style="height:110px">
            <img src="${escAttr(_normalizeDriveUrl(g.foto))}" class="w-full h-full object-cover cursor-pointer" loading="lazy"
                 onclick="GuiasView.verFotoGuia()" onerror="this.style.opacity='.3'"/>
            ${abierta ? `<div class="absolute top-2 right-2 flex gap-1">
              <label class="bg-slate-900/80 rounded-lg px-2 py-1 cursor-pointer text-xs text-slate-300" title="Cambiar - Galería">
                <input type="file" accept="image/*" class="hidden" onchange="GuiasView.onFotoGuiaSeleccionada(event)"/>🖼
              </label>
              <label class="bg-blue-900/80 rounded-lg px-2 py-1 cursor-pointer text-xs text-blue-200" title="Cambiar - Cámara">
                <input type="file" accept="image/*" capture="environment" class="hidden" onchange="GuiasView.onFotoGuiaSeleccionada(event)"/>📷
              </label>
              <button onclick="GuiasView.eliminarFotoGuia()" title="Eliminar foto"
                      class="bg-red-900/80 rounded-lg px-2 py-1 text-xs text-red-300 font-bold">✕</button>
            </div>` : ''}
          </div>`;
      } else {
        fotoEl.innerHTML = `
          <div class="flex gap-2 mb-3">
            <label class="flex-1 flex items-center justify-center gap-1 bg-slate-800 rounded-xl cursor-pointer text-slate-300 text-xs" style="min-height:48px">
              <input type="file" accept="image/*" class="hidden" onchange="GuiasView.onFotoGuiaSeleccionada(event)"/>
              🖼 Galería
            </label>
            <label class="flex-1 flex items-center justify-center gap-1 bg-slate-800 rounded-xl cursor-pointer text-blue-300 text-xs" style="min-height:48px">
              <input type="file" accept="image/*" capture="environment" class="hidden" onchange="GuiasView.onFotoGuiaSeleccionada(event)"/>
              📷 Cámara
            </label>
            ${g.idPreingreso ? `<button onclick="GuiasView.copiarFotoDePreingreso()"
                    class="flex-1 flex items-center justify-center bg-slate-800 rounded-xl text-blue-400 text-xs" style="min-height:48px">
              📋 Preingreso
            </button>` : ''}
          </div>`;
      }
    }

    // ── Notas panel (toggle) ───────────────────────────────
    _tagsGuia = _tagsFromComentario(g.comentario);
    const textoLibre = _textoLibreFromComentario(g.comentario);
    const cEl = document.getElementById('guiaDetComentarioSection');
    if (cEl) {
      if (!_notasOpen) {
        cEl.innerHTML = '';
      } else if (abierta) {
        const _tb = (id, label, grupo, val, colorA, colorI) =>
          `<button id="${id}" onclick="GuiasView.toggleTagGuia('${grupo}','${val}')"
                   class="flex-1 py-2 rounded-lg text-xs font-bold border transition-all
                          ${_tagsGuia[grupo]===val ? colorA : colorI}">${label}</button>`;
        cEl.innerHTML = `
          <div class="space-y-1 mb-2">
            <div class="flex gap-1">
              ${_tb('gTagComp1','Comprobante','comp','si',
                    'bg-blue-900/70 border-blue-500 text-blue-200',
                    'border-slate-700 text-slate-500')}
              ${_tb('gTagComp0','Sin comprobante','comp','no',
                    'bg-amber-900/70 border-amber-500 text-amber-200',
                    'border-slate-700 text-slate-500')}
            </div>
            <div class="flex gap-1">
              ${_tb('gTagCompl1','Completo','compl','si',
                    'bg-green-900/70 border-green-500 text-green-200',
                    'border-slate-700 text-slate-500')}
              ${_tb('gTagCompl0','Incompleto','compl','no',
                    'bg-amber-900/70 border-amber-500 text-amber-200',
                    'border-slate-700 text-slate-500')}
            </div>
          </div>
          <textarea id="guiaComentarioEdit" class="input text-xs" rows="2"
                    placeholder="Notas adicionales…">${textoLibre}</textarea>
          <p class="text-xs text-slate-600 mt-1">Se guarda al cerrar.</p>`;
      } else if (g.comentario) {
        cEl.innerHTML = `<p class="text-xs text-slate-400 italic mb-3">${escAttr(g.comentario)}</p>`;
      } else {
        cEl.innerHTML = '';
      }
    }

    } // end else (no esEnvasado)

    const items = (g.detalle || []).filter(d => d.observacion !== 'ANULADO');
    document.getElementById('guiaDetCount').textContent = `${items.length} ítem${items.length !== 1 ? 's' : ''}`;

    // Resetear selección si la guía cambió
    if (_selGuiaId !== g.idGuia) { _selIdx = -1; _selGuiaId = g.idGuia; }

    document.getElementById('guiaDetItems').innerHTML = items.length
      ? items.map((d, idx) => {
          const isSelected = abierta && idx === _selIdx;
          const pendiente  = d._local ? ' opacity-50' : '';
          const vencTxt    = isSelected ? _selVenc : (d.fechaVencimiento || '');

          if (isSelected) {
            // ── Tarjeta expandida ──────────────────────────────
            return `
            <div class="rounded-xl bg-slate-700/35 ring-1 ring-blue-500/30 shadow-lg px-3 pt-3 pb-3 mb-1.5${pendiente}"
                 onclick="event.stopPropagation()">
              <div class="flex items-start gap-3 mb-2.5">
                <div class="flex-1">
                  <p class="text-base font-bold text-white leading-snug">${escAttr(d.descripcionProducto || d.codigoProducto)}</p>
                  <p class="text-xs text-slate-500 font-mono mt-0.5">${escAttr(d.codigoProducto)}</p>
                </div>
                <div class="flex items-center gap-1 flex-shrink-0 mt-0.5">
                  <button onclick="GuiasView.inlineQtyDelta(-1)"
                          class="text-slate-300 text-xl leading-none w-7 h-7 flex items-center justify-center rounded-md active:bg-slate-600 select-none">−</button>
                  <input id="inlineQtyInput" type="number" step="any" inputmode="decimal"
                         value="${_selQty}"
                         class="text-base font-black text-white bg-transparent border-b border-slate-500 text-center w-14 focus:outline-none focus:border-blue-400"
                         oninput="GuiasView.inlineQtyInput(this.value)"
                         onblur="GuiasView.inlineQtyBlur(this.value)"
                         onfocus="this.select()"/>
                  <button onclick="GuiasView.inlineQtyDelta(1)"
                          class="text-blue-400 text-xl leading-none w-7 h-7 flex items-center justify-center rounded-md active:bg-slate-600 select-none">+</button>
                </div>
              </div>
              <div class="flex gap-2">
                ${esIngreso ? `
                <button onclick="GuiasView.inlinePickVenc()" id="inlineVencBtn"
                        class="flex-1 py-2 rounded-lg border ${vencTxt ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-xs flex items-center justify-center gap-1.5">
                  <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/></svg>
                  ${escAttr(vencTxt || 'Vencimiento')}
                </button>` : ''}
                <button onclick="GuiasView.inlineDelete(${idx})"
                        class="py-2 px-3 rounded-lg border border-red-800/60 text-red-400 text-xs flex items-center justify-center">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/><path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3h11V2h-11z"/></svg>
                </button>
              </div>
            </div>`;
          }

          // ── Tarjeta colapsada ──────────────────────────────
          const venc  = d.fechaVencimiento ? `<span class="text-xs text-amber-400 block mt-0.5">Venc: ${d.fechaVencimiento}</span>` : '';
          const iTag  = d._indirect
            ? `<span style="font-size:9px;font-weight:800;padding:1px 4px;border-radius:3px;
                            background:rgba(124,58,237,.18);color:#a78bfa;
                            border:1px solid rgba(124,58,237,.4);margin-right:4px;flex-shrink:0;vertical-align:middle">i</span>`
            : '';
          return `
          <div class="flex items-center gap-3 py-3 px-1 border-b border-slate-700/50 cursor-pointer active:bg-slate-700/20 rounded-lg${pendiente}"
               data-det-id="${d.idDetalle || ''}" onclick="GuiasView.selectItem(${idx})">
            <div class="flex-1 min-w-0">
              <p class="text-sm font-semibold text-slate-100 leading-snug">${iTag}${escAttr(d.descripcionProducto || d.codigoProducto)}</p>
              <p class="text-xs text-slate-500 font-mono mt-0.5">${escAttr(d.codigoProducto)}${d._local ? ' · guardando…' : ''}</p>
              ${venc}
            </div>
            <span class="text-base font-black text-white flex-shrink-0">${fmt(d.cantidadRecibida)}</span>
          </div>`;
        }).join('')
      : '<p class="text-slate-500 text-sm text-center py-4">Sin ítems registrados</p>';

    const monto = parseFloat(g.montoTotal) || 0;
    document.getElementById('guiaDetMontoVal').textContent = monto > 0 ? `S/. ${fmt(monto, 2)}` : '—';
    document.getElementById('guiaDetMonto').style.display = monto > 0 ? 'block' : 'none';

    const acciones = document.getElementById('guiaDetAcciones');
    acciones.innerHTML = abierta ? `
      <div style="display:flex;gap:0;border-radius:14px;overflow:hidden;border:1px solid rgba(124,58,237,.25);width:100%">
        <button onclick="GuiasView.abrirCamaraItem()"
                style="flex:3;min-height:52px;background:rgba(124,58,237,.12);
                       border:none;border-right:1px solid rgba(124,58,237,.2);cursor:pointer;
                       display:flex;align-items:center;justify-content:center;gap:8px;
                       color:#c4b5fd;font-weight:800;font-size:.82em;letter-spacing:.04em;
                       -webkit-tap-highlight-color:transparent"
                ontouchstart="this.style.background='rgba(124,58,237,.25)'" ontouchend="this.style.background='rgba(124,58,237,.12)'">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="opacity:.8">
            <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
            <circle cx="12" cy="13" r="4"/>
          </svg>
          CÁMARA
        </button>
        <button onclick="GuiasView.abrirScannerItem()"
                style="flex:1;min-height:52px;background:rgba(251,191,36,.06);
                       border:none;border-right:1px solid rgba(124,58,237,.15);cursor:pointer;
                       display:flex;align-items:center;justify-content:center;flex-direction:column;gap:3px;
                       color:#fbbf24;
                       -webkit-tap-highlight-color:transparent"
                ontouchstart="this.style.background='rgba(251,191,36,.18)'" ontouchend="this.style.background='rgba(251,191,36,.06)'">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M2 6h2v12H2zm4 0h1v12H6zm3 0h2v12H9zm4 0h1v12h-1zm3 0h2v12h-2zm3 0h1v12h-1z"/>
          </svg>
          <span style="font-size:.55em;font-weight:700;letter-spacing:.05em">SCANNER</span>
        </button>
        <button onclick="GuiasView.abrirPNSinCodigo()"
                style="flex:1;min-height:52px;background:rgba(245,158,11,.06);
                       border:none;cursor:pointer;
                       display:flex;align-items:center;justify-content:center;flex-direction:column;gap:3px;
                       color:#f59e0b;
                       -webkit-tap-highlight-color:transparent"
                ontouchstart="this.style.background='rgba(245,158,11,.22)'" ontouchend="this.style.background='rgba(245,158,11,.06)'"
                title="Registrar producto nuevo (sin barcode)">
          <span style="font-weight:900;font-size:1.15em;line-height:1">N</span>
          <span style="font-size:.55em;font-weight:700;letter-spacing:.05em">NUEVO</span>
        </button>
      </div>` : '';

    if (conAnimacion) abrirSheet('sheetGuiaDetalle');
  }

  // ── Edición inline de ítems ──────────────────────────────
  let _selIdx     = -1;   // índice del ítem seleccionado (-1 = ninguno)
  let _selQty     = 0;
  let _selVenc    = '';
  let _selOrigQty = 0;
  let _selOrigVenc = '';
  let _selGuiaId  = '';

  function selectItem(newIdx) {
    if (_selIdx === newIdx) {
      // Doble-tap: guardar y colapsar
      _commitInline();
      _selIdx = -1;
      _mostrarDetalleSheet(_guiaActual, false);
      return;
    }
    _commitInline(); // guardar anterior si cambió
    const items = (_guiaActual?.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[newIdx];
    if (!d) return;
    _selIdx      = newIdx;
    _selQty      = parseFloat(d.cantidadRecibida) || 0;
    _selVenc     = d.fechaVencimiento || '';
    _selOrigQty  = _selQty;
    _selOrigVenc = _selVenc;
    _mostrarDetalleSheet(_guiaActual, false);
    setTimeout(() => {
      const el = document.getElementById('inlineQtyInput');
      el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      el?.focus();
    }, 60);
  }

  function deselectItem() {
    if (_selIdx < 0) return;
    _commitInline();
    _selIdx = -1;
    _mostrarDetalleSheet(_guiaActual, false);
  }

  function _commitInline() {
    if (_selIdx < 0 || !_guiaActual) return;
    // Leer qty del DOM por si el usuario escribió sin disparar oninput
    const inputEl = document.getElementById('inlineQtyInput');
    if (inputEl) _selQty = parseFloat(inputEl.value) || _selQty;

    const items = (_guiaActual.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[_selIdx];
    if (!d) return;

    const qtyChanged  = _selQty !== _selOrigQty;
    const vencChanged = _selVenc !== _selOrigVenc;
    if (!qtyChanged && !vencChanged) return;

    const idDetalle = d.idDetalle;
    if (_selQty <= 0) {
      d.observacion = 'ANULADO';
      API.anularDetalle({ idDetalle }).catch(() => {});
      toast('Ítem eliminado', 'warn', 1200);
      return;
    }
    d.cantidadRecibida = _selQty;
    d.fechaVencimiento = _selVenc;
    if (qtyChanged)  API.actualizarCantidadDetalle({ idDetalle, cantidadRecibida: _selQty }).catch(() => {});
    if (vencChanged) API.actualizarFechaVencimiento({ idDetalle, fechaVencimiento: _selVenc }).catch(() => {});
    toast('Guardado', 'ok', 1000);
  }

  function inlineQtyDelta(delta) {
    _selQty = Math.max(0, parseFloat(_selQty || 0) + delta);
    const el = document.getElementById('inlineQtyInput');
    if (el) el.value = _selQty;
  }

  function inlineQtyInput(val) {
    const n = parseFloat(val);
    _selQty = isNaN(n) ? 0 : Math.max(0, n);
  }

  function inlineQtyBlur(val) {
    const n = parseFloat(val);
    if (!isNaN(n)) _selQty = Math.max(0, n);
  }

  function inlinePickVenc() {
    const el = document.getElementById('inlineVencHidden');
    if (!el) return;
    el.value = _selVenc;
    el.min = new Date().toISOString().split('T')[0];
    if (typeof el.showPicker === 'function') { try { el.showPicker(); } catch { el.click(); } }
    else el.click();
  }

  function inlineVencChanged(val) {
    _selVenc = val || '';
    const btn = document.getElementById('inlineVencBtn');
    if (btn) {
      btn.className = `flex-1 py-2 rounded-lg border ${_selVenc ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-xs flex items-center justify-center gap-1.5`;
      btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/></svg> ${_selVenc ? escAttr(_selVenc) : 'Vencimiento'}`;
    }
  }

  function inlineDelete(idx) {
    const items = (_guiaActual?.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[idx];
    if (!d) return;
    d.observacion = 'ANULADO';
    _selIdx = -1;
    _mostrarDetalleSheet(_guiaActual, false);
    API.anularDetalle({ idDetalle: d.idDetalle }).catch(() => {});
    toast('Ítem eliminado', 'warn', 1200);
  }

  // (funciones antiguas del sheet — eliminadas en v1.0.40)
  function abrirEditarItem(idx) {
    if (!_guiaActual) return;
    const items = (_guiaActual.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[idx];
    if (!d) return;
    _editItemIdx      = idx;
    _editItemQty      = parseFloat(d.cantidadRecibida) || 0;
    _editItemVenc     = d.fechaVencimiento || '';
    _editItemId       = d.idDetalle;
    _editItemOrigQty  = _editItemQty;
    _editItemOrigVenc = _editItemVenc;

    const abierta   = _guiaActual.estado === 'ABIERTA';
    const esIngreso = (_guiaActual.tipo || '').startsWith('INGRESO');
    const qtyDisplay = Number.isInteger(_editItemQty) ? String(_editItemQty) : String(_editItemQty);

    document.getElementById('editItemContent').innerHTML = `
      <div class="mb-5">
        <p class="font-bold text-white text-base">${escAttr(d.descripcionProducto || d.codigoProducto)}</p>
        <p class="text-xs text-slate-500 font-mono">${escAttr(d.codigoProducto)}</p>
      </div>
      ${abierta ? `
      <div class="flex items-center justify-center gap-5 mb-5">
        <button onclick="GuiasView.itemEditQtyChange(-1)"
                class="w-14 h-14 rounded-full bg-slate-700 text-3xl font-black text-white active:scale-95 select-none">−</button>
        <input id="editItemQtyInput" type="number" step="any" inputmode="decimal"
               value="${qtyDisplay}"
               class="text-4xl font-black text-white bg-transparent border-b-2 border-slate-500 text-center w-28 focus:outline-none focus:border-blue-400"
               onchange="GuiasView.itemEditSetQty(this.value)"
               onfocus="this.select()"/>
        <button onclick="GuiasView.itemEditQtyChange(1)"
                class="w-14 h-14 rounded-full bg-blue-600 text-3xl font-black text-white active:scale-95 select-none">+</button>
      </div>
      ${esIngreso ? `
      <button onclick="GuiasView.itemEditPickVenc()" id="editItemVencBtn"
              class="w-full py-3 rounded-xl border ${_editItemVenc ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-sm font-bold mb-3 flex items-center justify-center gap-2">
        <svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
          <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/>
        </svg>
        ${_editItemVenc ? _editItemVenc : 'Agregar vencimiento'}
      </button>` : ''}
      <button onclick="GuiasView.eliminarItemEdit()"
              class="w-full py-3 rounded-xl border border-red-800/60 text-red-400 text-sm font-bold flex items-center justify-center gap-2">
        <svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
          <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/>
          <path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3h11V2h-11z"/>
        </svg>
        Eliminar ítem
      </button>` : `<p class="text-center text-sm text-slate-500 py-4">Guía cerrada — solo lectura</p>`}`;

    // Inicializar hidden date input
    const hiddenInp = document.getElementById('editItemVencHidden');
    if (hiddenInp) {
      hiddenInp.value = _editItemVenc;
      hiddenInp.min   = new Date().toISOString().split('T')[0];
    }
    abrirSheet('sheetEditItem');
  }

  function cerrarEditItem() {
    cerrarSheet('sheetEditItem');
    if (!_guiaActual || _editItemIdx < 0) return;
    const items = (_guiaActual.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[_editItemIdx];
    if (!d) return;

    // Leer qty del input en caso de edición manual directa sin onchange
    const inputEl = document.getElementById('editItemQtyInput');
    if (inputEl) _editItemQty = parseFloat(inputEl.value) || _editItemQty;
    const qtyFinal = _editItemQty;
    const qtyChanged  = qtyFinal !== _editItemOrigQty;
    const vencChanged = _editItemVenc !== _editItemOrigVenc;
    if (!qtyChanged && !vencChanged) return;

    if (qtyFinal <= 0) {
      // Eliminar
      d.observacion = 'ANULADO';
      _mostrarDetalleSheet(_guiaActual, false);
      API.anularDetalle({ idDetalle: _editItemId }).catch(() => {});
      toast('Ítem eliminado', 'warn', 1500);
      return;
    }

    d.cantidadRecibida = qtyFinal;
    d.fechaVencimiento = _editItemVenc;
    _mostrarDetalleSheet(_guiaActual, false);

    if (qtyChanged)  API.actualizarCantidadDetalle({ idDetalle: _editItemId, cantidadRecibida: qtyFinal }).catch(() => {});
    if (vencChanged) API.actualizarFechaVencimiento({ idDetalle: _editItemId, fechaVencimiento: _editItemVenc }).catch(() => {});
    toast('Ítem guardado', 'ok', 1500);
  }

  function itemEditQtyChange(delta) {
    _editItemQty = Math.max(0, (_editItemQty || 0) + delta);
    const el = document.getElementById('editItemQtyInput');
    if (el) { el.value = _editItemQty; }
  }

  function itemEditSetQty(val) {
    const n = parseFloat(val);
    _editItemQty = isNaN(n) ? 0 : Math.max(0, n);
  }

  function itemEditPickVenc() {
    const el = document.getElementById('editItemVencHidden');
    if (!el) return;
    if (typeof el.showPicker === 'function') { try { el.showPicker(); } catch { el.click(); } }
    else el.click();
  }

  function itemEditOnVencChanged(val) {
    _editItemVenc = val || '';
    const btn = document.getElementById('editItemVencBtn');
    if (btn) {
      btn.textContent = '';
      btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
        <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/>
      </svg> ${val ? escAttr(val) : 'Agregar vencimiento'}`;
      btn.className = `w-full py-3 rounded-xl border ${val ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-sm font-bold mb-3 flex items-center justify-center gap-2`;
    }
  }

  function eliminarItemEdit() {
    const d = (_guiaActual?.detalle || []).filter(x => x.observacion !== 'ANULADO')[_editItemIdx];
    if (!d) return;
    d.observacion = 'ANULADO';
    cerrarSheet('sheetEditItem');
    _mostrarDetalleSheet(_guiaActual, false);
    API.anularDetalle({ idDetalle: _editItemId }).catch(() => {});
    toast('Ítem eliminado', 'warn', 1500);
  }

  // Toggle estado guía: abierta → cerrar; cerrada → pedir adminPin
  function toggleEstadoGuia() {
    if (!_guiaActual) return;
    if (_guiaActual.estado === 'ABIERTA') {
      // Double-tap confirm para cierre (acción irreversible sin admin)
      const lockBtnEl = document.querySelector('[onclick="GuiasView.toggleEstadoGuia()"]');
      if (lockBtnEl) {
        if (!lockBtnEl._tapArmed) {
          lockBtnEl._tapArmed = true;
          lockBtnEl.style.background = 'rgba(245,158,11,.2)';
          lockBtnEl.style.borderColor = '#f59e0b';
          lockBtnEl.style.color = '#fde68a';
          lockBtnEl.title = 'Toca de nuevo para confirmar cierre';
          vibrate(15);
          lockBtnEl._tapTimer = setTimeout(() => {
            lockBtnEl._tapArmed = false;
            lockBtnEl.style.cssText = '';
            lockBtnEl.title = 'Cerrar guía';
          }, 2500);
          return;
        }
        clearTimeout(lockBtnEl._tapTimer);
        lockBtnEl._tapArmed = false;
        lockBtnEl.style.cssText = '';
      }
      confirmarCerrarGuia();
    } else {
      _pedirAdminPin(_guiaActual.idGuia);
    }
  }

  // Local dot-indicator updater for admin PIN (3 dots only)
  function _updAdminDots(n) {
    for (let i = 0; i < 3; i++) {
      const el = document.getElementById('apn' + i);
      if (el) el.className = i < n
        ? 'w-4 h-4 rounded-full bg-amber-400'
        : 'w-4 h-4 rounded-full border-2 border-slate-600';
    }
  }

  // Admin PIN dialog para reabrir guía
  let _pinGuiaTarget = null;
  let _adminPinBuf   = '';

  function _pedirAdminPin(idGuia) {
    _pinGuiaTarget = idGuia;
    _adminPinBuf   = '';
    _updAdminDots(0);
    document.getElementById('adminPinError').textContent = '';
    document.getElementById('adminPinModal').style.display = 'flex';
  }

  function adminPinTecla(d) {
    if (_adminPinBuf.length >= 3) return;
    _adminPinBuf += d;
    _updAdminDots(_adminPinBuf.length);
    if (_adminPinBuf.length === 3) setTimeout(_verificarAdminPin, 150);
  }

  function adminPinAtras() {
    _adminPinBuf = _adminPinBuf.slice(0, -1);
    _updAdminDots(_adminPinBuf.length);
  }

  async function _verificarAdminPin() {
    const cached = OfflineManager.getAdminPin();
    if (!cached) {
      // Sin caché: enviar a GAS de todos modos (GAS verifica desde MOS)
    } else if (String(_adminPinBuf) !== String(cached)) {
      document.getElementById('adminPinError').textContent = 'PIN incorrecto';
      _adminPinBuf = '';
      _updAdminDots(0);
      setTimeout(() => { document.getElementById('adminPinError').textContent = ''; }, 1500);
      return;
    }
    document.getElementById('adminPinModal').style.display = 'none';
    const res = await API.reabrirGuia({ idGuia: _pinGuiaTarget });
    if (res.ok || res.offline) {
      if (_guiaActual?.idGuia === _pinGuiaTarget) {
        _guiaActual.estado = 'ABIERTA';
        _mostrarDetalleSheet(_guiaActual, false);
      }
      // Update in list
      const idx = todas.findIndex(g => g.idGuia === _pinGuiaTarget);
      if (idx >= 0) { todas[idx].estado = 'ABIERTA'; render(_filtrar(todas, filtroActual)); }
      toast('Guía reabierta', 'ok');
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  // ════════════════════════════════════════════════════════════
  // AGREGAR ÍTEM — modo cámara (continua) y modo scanner HID
  // ════════════════════════════════════════════════════════════

  // Compat: abrirAgregarItem legacy → usa cámara
  function abrirAgregarItem() { abrirCamaraItem(); }

  // ── MODO CÁMARA ──────────────────────────────────────────────
  function abrirCamaraItem() {
    if (!_guiaActual) return;
    document.getElementById('scannerModal').classList.add('open');
    // Ocultar picker y toast previos
    document.getElementById('camPicker').style.display = 'none';
    document.getElementById('camToast').style.display  = 'none';
    // Iniciar scanner en modo continuo
    Scanner.start('scanVideo', _onCamResult, err => {
      toast('Error cámara: ' + err, 'danger');
      document.getElementById('scannerModal').classList.remove('open');
    }, { continuous: true, cooldown: 1500 });
  }

  function cerrarCamara() {
    Scanner.stop();
    document.getElementById('scannerModal').classList.remove('open');
    document.getElementById('camPicker').style.display = 'none';
    document.getElementById('camToast').style.display  = 'none';
  }

  // Callback del scanner continuo — procesa código sin cerrar cámara
  function _onCamResult(cod) {
    const codStr = String(cod || '').trim();
    if (!codStr) return;

    // Ocultar picker anterior
    document.getElementById('camPicker').style.display = 'none';

    const candidatos = _buscarCandidatos(codStr);

    if (!candidatos.length) {
      _camToast('⚠ No encontrado: ' + codStr, '', '#ef4444', '#fca5a5');
      return;
    }
    if (candidatos.length === 1 || candidatos[0]._exacto) {
      _agregarProductoDirecto(candidatos[0], false);
      _camToast('✓ ' + String(candidatos[0].descripcion || candidatos[0].idProducto),
                candidatos[0].idProducto, 'rgba(16,185,129,.94)', '#fff');
      return;
    }
    // Múltiples → mostrar picker (cámara sigue activa detrás)
    _mostrarCamPicker(candidatos, codStr);
  }

  function _camToast(desc, id, bg, col) {
    const t = document.getElementById('camToast');
    const d = document.getElementById('camToastDesc');
    const i = document.getElementById('camToastId');
    if (!t) return;
    d.textContent = desc;
    i.textContent = id;
    t.style.background = bg;
    d.style.color = col;
    t.style.display = 'flex';
    clearTimeout(_camToastTimer);
    _camToastTimer = setTimeout(() => { if (t) t.style.display = 'none'; }, 2500);
  }
  let _camToastTimer = null;

  function _mostrarCamPicker(candidatos, codStr) {
    document.getElementById('camPickerCod').textContent = codStr;
    document.getElementById('camPickerList').innerHTML = candidatos.map(p => `
      <button onclick="GuiasView.seleccionarItemCamara('${escAttr(p.idProducto)}')"
              style="width:100%;text-align:left;padding:9px 11px;border-radius:10px;
                     border:1px solid rgba(100,116,139,.3);margin-bottom:6px;
                     background:rgba(15,23,42,.8);display:flex;align-items:center;gap:8px;
                     cursor:pointer;-webkit-tap-highlight-color:transparent"
              ontouchstart="this.style.borderColor='#a78bfa'" ontouchend="this.style.borderColor='rgba(100,116,139,.3)'">
        <div style="flex:1;min-width:0">
          <p style="font-size:.82em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || p.idProducto))}</p>
          <p style="font-size:.69em;color:#64748b;font-family:monospace">${escAttr(p.idProducto)}${p.codigoBarra ? ' · ' + p.codigoBarra : ''}</p>
        </div>
      </button>`).join('');
    document.getElementById('camPicker').style.display = 'block';
  }

  function cerrarCamPicker() {
    document.getElementById('camPicker').style.display = 'none';
    // Scanner ya está corriendo — no hace falta reiniciar
  }

  function seleccionarItemCamara(idProducto) {
    document.getElementById('camPicker').style.display = 'none';
    const prod = OfflineManager.getProductosCache().find(p => p.idProducto === idProducto);
    if (!prod) return;
    _agregarProductoDirecto(prod, true);
    _camToast('✓ ' + String(prod.descripcion || prod.idProducto), prod.idProducto, 'rgba(16,185,129,.94)', '#fff');
  }

  // ── MODO SCANNER HID ─────────────────────────────────────────
  let _hidBuffer   = '';
  let _hidBufTs    = 0;
  let _hidTimer    = null;

  function abrirScannerItem() {
    if (!_guiaActual) return;
    _hidBuffer = '';
    abrirSheet('sheetScanInput');
    // Limpiar estado visual
    const disp = document.getElementById('hidCodeText');
    if (disp) disp.textContent = '— listo para escanear —';
    document.getElementById('hidPicker').style.display   = 'none';
    document.getElementById('hidProductList').innerHTML  = '';
    document.getElementById('hidReadingDot').style.display = 'inline-block';
    setTimeout(() => _enfocarHid(), 260);
  }

  function cerrarScannerItem() {
    clearTimeout(_hidTimer);
    _hidBuffer = '';
    document.getElementById('hidReadingDot').style.display = 'none';
    cerrarSheet('sheetScanInput');
  }

  function _enfocarHid() {
    const inp = document.getElementById('hidScanInput');
    if (!inp) return;
    inp.value = '';
    // Registrar listener de teclado (una sola vez)
    if (!inp._hidListenerSet) {
      inp._hidListenerSet = true;
      inp.addEventListener('keydown', _hidKeydown);
    }
    inp.focus();
  }

  function _hidKeydown(e) {
    const now = Date.now();

    if (e.key === 'Enter') {
      e.preventDefault();
      clearTimeout(_hidTimer);
      const buf = _hidBuffer.trim();
      _hidBuffer = '';
      _updateHidDisplay('— listo para escanear —');
      if (buf) _procesarCodigoHid(buf);
      return;
    }

    if (e.key === 'Backspace' || e.key === 'Delete') {
      _hidBuffer = _hidBuffer.slice(0, -1);
      _updateHidDisplay(_hidBuffer || '— listo para escanear —');
      e.preventDefault();
      return;
    }

    // Solo aceptar caracteres imprimibles
    if (e.key.length !== 1) { e.preventDefault(); return; }

    // Detección de velocidad: si el buffer no está vacío y el intervalo
    // es > 80 ms → probable escritura humana → rechazar y limpiar
    if (_hidBuffer.length > 0 && (now - _hidBufTs) > 80) {
      _hidBuffer = '';
      _updateHidDisplay('— listo para escanear —');
      // No retornar: aceptar este primer carácter del scanner
    }

    _hidBufTs = now;
    _hidBuffer += e.key;
    _updateHidDisplay(_hidBuffer);
    e.preventDefault(); // no alterar input.value

    // Auto-procesar si el scanner no envía Enter (fallback 400ms)
    clearTimeout(_hidTimer);
    _hidTimer = setTimeout(() => {
      const buf = _hidBuffer.trim();
      _hidBuffer = '';
      _updateHidDisplay('— listo para escanear —');
      if (buf) _procesarCodigoHid(buf);
    }, 400);
  }

  function _updateHidDisplay(text) {
    const el = document.getElementById('hidCodeText');
    if (el) el.textContent = text || '— listo para escanear —';
  }

  function _procesarCodigoHid(codStr) {
    const candidatos = _buscarCandidatos(codStr);

    document.getElementById('hidPicker').style.display = 'none';

    if (!candidatos.length) {
      _updateHidDisplay('⚠ No encontrado: ' + codStr);
      _ofrecerPNEnScanner(codStr);
      return;
    }

    if (candidatos.length === 1 || candidatos[0]._exacto) {
      const prod = candidatos[0];
      _agregarProductoDirecto(prod, false);
      _agregarItemHidList(prod);
      _updateHidDisplay('— listo para escanear —');
      setTimeout(() => _enfocarHid(), 100);
      return;
    }

    // Múltiples → picker dentro del sheet
    document.getElementById('hidPickerCod').textContent = codStr;
    document.getElementById('hidPickerList').innerHTML = candidatos.map(p => `
      <button onclick="GuiasView.seleccionarItemHid('${escAttr(p.idProducto)}')"
              style="width:100%;text-align:left;padding:8px 10px;border-radius:9px;
                     border:1px solid rgba(124,58,237,.2);margin-bottom:5px;
                     background:rgba(124,58,237,.06);display:flex;align-items:center;gap:8px;
                     cursor:pointer;-webkit-tap-highlight-color:transparent">
        <div style="flex:1;min-width:0">
          <p style="font-size:.82em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || p.idProducto))}</p>
          <p style="font-size:.69em;color:#64748b;font-family:monospace">${escAttr(p.idProducto)}${p.codigoBarra ? ' · ' + p.codigoBarra : ''}</p>
        </div>
      </button>`).join('');
    document.getElementById('hidPicker').style.display = 'block';
  }

  function seleccionarItemHid(idProducto) {
    const prod = OfflineManager.getProductosCache().find(p => p.idProducto === idProducto);
    document.getElementById('hidPicker').style.display = 'none';
    if (!prod) return;
    _agregarProductoDirecto(prod, true);
    _agregarItemHidList(prod);
    _updateHidDisplay('— listo para escanear —');
    setTimeout(() => _enfocarHid(), 100);
  }

  // Muestra card del producto agregado en la lista del sheet HID
  function _agregarItemHidList(prod) {
    const list = document.getElementById('hidProductList');
    if (!list) return;
    const div = document.createElement('div');
    div.className = 'item-slide-in';
    div.style.cssText = 'display:flex;align-items:center;gap:9px;padding:9px 11px;' +
      'border-radius:11px;background:#1e293b;border:1px solid #334155;margin-bottom:6px';
    div.innerHTML = `<span style="color:#10b981;font-size:1.1em;flex-shrink:0">✓</span>
      <div style="flex:1;min-width:0">
        <p style="font-size:.82em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(prod.descripcion || prod.idProducto))}</p>
        <p style="font-size:.69em;color:#64748b;font-family:monospace">${escAttr(prod.idProducto)}</p>
      </div>`;
    list.insertBefore(div, list.firstChild);
  }

  // ── Búsqueda de candidatos (shared entre cámara y HID) ───────
  function _buscarCandidatos(codStr) {
    const prods = OfflineManager.getProductosCache();
    const cLow  = codStr.toLowerCase();
    const cUp   = codStr.toUpperCase();

    // 1. Exacta
    const exacto = prods.find(p =>
      p.idProducto === codStr ||
      String(p.codigoBarra || '') === codStr ||
      (p.idProducto || '').toLowerCase() === cLow ||
      String(p.codigoBarra || '').toLowerCase() === cLow
    );
    if (exacto) { exacto._exacto = true; return [exacto]; }

    // 2. Prefijo (ej: "12345" → "12345A","12345B")
    const porPrefijo = prods.filter(p =>
      (p.idProducto || '').toUpperCase().startsWith(cUp) ||
      String(p.codigoBarra || '').startsWith(codStr)
    );

    // 3. Parcial (descripción / código)
    const porParcial = prods.filter(p =>
      !porPrefijo.includes(p) && (
        (p.idProducto || '').toLowerCase().includes(cLow) ||
        String(p.codigoBarra || '').toLowerCase().includes(cLow) ||
        String(p.descripcion || '').toLowerCase().includes(cLow)
      )
    ).slice(0, 6);

    return [...porPrefijo, ...porParcial].slice(0, 10);
  }

  // ── Legacy: procesarHidInput (compat con sheetScanInput anterior) ─
  function procesarHidInput() { /* obsoleto, no se usa */ }
  function onHidKey(e) { /* obsoleto */ }
  function usarCamara() { cerrarScannerItem(); setTimeout(abrirCamaraItem, 220); }
  function cerrarScanInput() { cerrarScannerItem(); }

  function _agregarProductoDirecto(prod, indirecto) {
    if (!_guiaActual) return;
    const cod          = prod.idProducto;
    const descCapturada = prod.descripcion || prod.nombre || cod;
    const localId      = 'DL' + Date.now();

    const itemOptimista = {
      idDetalle: localId, idGuia: _guiaActual.idGuia,
      codigoProducto: cod,
      descripcionProducto: descCapturada,
      cantidadEsperada: 0, cantidadRecibida: 1,
      precioUnitario: 0, fechaVencimiento: '', observacion: '',
      _local: true,
      _indirect: !!indirecto
    };
    if (!_guiaActual.detalle) _guiaActual.detalle = [];
    _guiaActual.detalle.push(itemOptimista);
    _mostrarDetalleSheet(_guiaActual, false);
    // Animar el nuevo ítem en la lista
    requestAnimationFrame(() => {
      const el = document.querySelector(`[data-det-id="${localId}"]`);
      if (el) el.classList.add('item-slide-in');
    });
    toast((indirecto ? '↕ ' : '✓ ') + descCapturada, 'ok', 1500);

    const _idGuiaParaDetalle = _guiaActual.idGuia;
    API.agregarDetalle({
      idGuia: _idGuiaParaDetalle,
      codigoProducto: cod,
      cantidadEsperada: 0, cantidadRecibida: 1,
      precioUnitario: 0, fechaVencimiento: ''
    }).then(res => {
      if (res.ok && !res.offline) {
        const idx = _guiaActual.detalle?.findIndex(d => d.idDetalle === localId);
        if (idx >= 0) {
          const itemFinal = {
            ...res.data,
            idGuia: res.data.idGuia || _idGuiaParaDetalle,
            descripcionProducto: res.data.descripcionProducto || descCapturada,
            _local: false,
            _indirect: !!indirecto
          };
          _guiaActual.detalle[idx] = itemFinal;
          _mostrarDetalleSheet(_guiaActual, false);
          // Guardar en cache local para reaperturas rápidas
          OfflineManager.addDetalleCache(itemFinal);
        }
      } else if (!res.offline) {
        _guiaActual.detalle = _guiaActual.detalle.filter(d => d.idDetalle !== localId);
        _mostrarDetalleSheet(_guiaActual, false);
        toast(res.error === 'PRODUCTO_NO_ENCONTRADO'
          ? 'Producto no registrado en el sistema' : 'Error: ' + (res.error || res.mensaje),
          res.error === 'PRODUCTO_NO_ENCONTRADO' ? 'warn' : 'danger', 4000);
      }
    }).catch(() => {});
  }

  function _rescanear() {
    abrirCamaraItem();
  }

  // ── Producto Nuevo ────────────────────────────────────────────
  let _pnCodigo     = '';
  let _pnFotoBase64 = '';
  let _pnFotoMime   = 'image/jpeg';

  function _ofrecerPNEnScanner(cod) {
    _pnCodigo = cod;
    const offer = document.getElementById('pnOffer');
    if (offer) {
      document.getElementById('pnOfferCod').textContent = cod;
      offer.style.display = 'block';
    }
    setTimeout(() => _enfocarHid(), 100);
  }

  function _ocultarPNOffer() {
    const offer = document.getElementById('pnOffer');
    if (offer) offer.style.display = 'none';
    _updateHidDisplay('— listo para escanear —');
    setTimeout(() => _enfocarHid(), 100);
  }

  function abrirModalPN(codigoBarra, idGuia) {
    _pnCodigo     = codigoBarra !== undefined ? codigoBarra : _pnCodigo;
    const guiaId  = idGuia !== undefined ? idGuia : (_guiaActual?.idGuia || null);
    _pnFotoBase64 = '';
    _pnFotoMime   = 'image/jpeg';

    const dispEl = document.getElementById('pnCodigoDisplay');
    if (dispEl) dispEl.textContent = _pnCodigo || '(se generará automáticamente · NLEV...)';

    const cant = document.getElementById('pnCantidad');
    if (cant) cant.value = '1';
    const fv = document.getElementById('pnFechaVenc');
    if (fv) fv.value = '';
    const obs = document.getElementById('pnObservaciones');
    if (obs) obs.value = '';
    const prev = document.getElementById('pnFotoPreview');
    if (prev) prev.style.display = 'none';
    const inp = document.getElementById('pnFotoInput');
    if (inp) inp.value = '';
    const btnTxt = document.getElementById('pnFotoBtn');
    if (btnTxt) btnTxt.textContent = 'Tomar / elegir foto';

    // Guardar idGuia en dataset del modal para recuperarlo al confirmar
    const modal = document.getElementById('modalProductoNuevo');
    if (modal) {
      modal.dataset.pnGuia = guiaId || '';
      modal.classList.add('open');
    }
  }

  function cerrarModalPN() {
    const modal = document.getElementById('modalProductoNuevo');
    if (modal) modal.classList.remove('open');
  }

  function _pnFotoSeleccionada(input) {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      const dataUrl = e.target.result;
      _pnFotoBase64 = dataUrl.split(',')[1] || '';
      _pnFotoMime   = file.type || 'image/jpeg';
      const img = document.getElementById('pnFotoImg');
      if (img) img.src = dataUrl;
      const prev = document.getElementById('pnFotoPreview');
      if (prev) prev.style.display = 'block';
      const btnTxt = document.getElementById('pnFotoBtn');
      if (btnTxt) btnTxt.textContent = '✓ Foto lista — toca para cambiar';
    };
    reader.readAsDataURL(file);
  }

  async function confirmarRegistrarPN() {
    const modal    = document.getElementById('modalProductoNuevo');
    const idGuia   = modal?.dataset.pnGuia || '';
    const cantidad = parseFloat(document.getElementById('pnCantidad')?.value) || 1;
    const fechaVenc = document.getElementById('pnFechaVenc')?.value || '';
    const obs      = (document.getElementById('pnObservaciones')?.value || '').trim();

    const params = {
      codigoBarra:     _pnCodigo || '',
      idGuia,
      cantidad,
      fechaVencimiento: fechaVenc,
      descripcion:     obs,
      usuario:         window.WH_CONFIG?.usuario || '',
      idSesion:        window.WH_CONFIG?.idSesion || ''
    };
    if (_pnFotoBase64) {
      params.fotoBase64 = _pnFotoBase64;
      params.mimeType   = _pnFotoMime;
    }

    cerrarModalPN();
    _ocultarPNOffer();
    toast('Registrando producto nuevo...', 'ok', 2000);

    try {
      const res = await API.registrarPN(params);
      if (res.ok || res.offline) {
        const cod  = res.data?.codigoBarra || _pnCodigo || '';
        const idPN = res.data?.idProductoNuevo || ('PN_L_' + Date.now());

        // Persistir en cache local
        const pnCache = OfflineManager.getPNCache();
        pnCache.unshift({ idProductoNuevo: idPN, idGuia, codigoBarra: cod,
          cantidad, fechaVencimiento: fechaVenc, descripcion: obs, estado: 'PENDIENTE',
          fechaRegistro: new Date().toISOString() });
        OfflineManager.setPNCache(pnCache);

        toast(`✓ Producto nuevo registrado · ${cod || idPN}`, 'ok', 3500);
        vibrate(20);

        // Si la guía está abierta en detalle → agregar ítem optimista con badge N
        if (idGuia && _guiaActual?.idGuia === idGuia) {
          const itemOpt = {
            idDetalle: 'DL_PN_' + Date.now(), idGuia,
            codigoProducto: cod, descripcionProducto: (obs || cod || '(nuevo)') + ' ⬡N',
            cantidadEsperada: 0, cantidadRecibida: cantidad,
            precioUnitario: 0, fechaVencimiento: fechaVenc, observacion: '',
            _local: true, _esPN: true
          };
          if (!_guiaActual.detalle) _guiaActual.detalle = [];
          _guiaActual.detalle.push(itemOpt);
          _mostrarDetalleSheet(_guiaActual, false);
        }

        // Refrescar lista de guías para actualizar badge N
        if (typeof render === 'function' && typeof _filtrar === 'function') {
          render(_filtrar(todas, filtroActual));
        }

        // Mantener foco en scanner si estaba abierto
        if (document.getElementById('sheetScanInput')?.classList.contains('open')) {
          setTimeout(() => _enfocarHid(), 250);
        }
      } else {
        toast('Error: ' + (res.error || 'No se pudo registrar'), 'warn', 4000);
      }
    } catch(e) {
      toast('Error al registrar producto nuevo', 'warn', 3000);
    }
  }

  function abrirPNSinCodigo() {
    abrirModalPN('', _guiaActual?.idGuia || null);
  }

  async function confirmarCerrarGuia() {
    if (!_guiaActual) return;
    const det = (_guiaActual.detalle || []).filter(d => d.observacion !== 'ANULADO');
    if (!det.length) { toast('Agrega al menos un ítem antes de cerrar', 'warn'); return; }

    // Advertir si hay productos nuevos pendientes
    const pnPend = (OfflineManager.getPNCache() || []).filter(p => p.idGuia === _guiaActual.idGuia && p.estado === 'PENDIENTE').length;
    if (pnPend) toast(`⚠ ${pnPend} producto(s) nuevo(s) sin aprobar en esta guía`, 'warn', 4000);

    // Optimista: actualizar estado en UI inmediatamente
    _guiaActual.estado = 'CERRADA';
    _mostrarDetalleSheet(_guiaActual, false);
    const idx = todas.findIndex(g => g.idGuia === _guiaActual.idGuia);
    if (idx >= 0) { todas[idx].estado = 'CERRADA'; render(_filtrar(todas, filtroActual)); }

    const res = await API.cerrarGuia(_guiaActual.idGuia, window.WH_CONFIG.usuario);
    if (res.ok || res.offline) {
      const monto = res.data?.montoTotal;
      toast(`Guía cerrada${monto ? ` · S/. ${fmt(monto, 2)}` : ''}`, 'ok', 3000);
      if (res.ok && !res.offline) {
        _guiaActual.montoTotal = monto || 0;
        if (idx >= 0) todas[idx].montoTotal = monto || 0;
        _mostrarDetalleSheet(_guiaActual, false);
        render(_filtrar(todas, filtroActual));
      }
    } else {
      // Revertir si GAS rechazó
      _guiaActual.estado = 'ABIERTA';
      if (idx >= 0) todas[idx].estado = 'ABIERTA';
      _mostrarDetalleSheet(_guiaActual, false);
      render(_filtrar(todas, filtroActual));
      toast('Error: ' + res.error, 'danger');
    }
  }

  async function crearGuia() {
    const tipo        = document.getElementById('guiaTipo').value;
    const idProveedor = document.getElementById('guiaProveedor').value;
    const textoExtra  = (document.getElementById('guiaComentario').value || '').trim();
    const comentario  = _buildComentario(_tagsNueva, textoExtra);
    const params = {
      tipo,
      usuario:         window.WH_CONFIG.usuario,
      idProveedor,
      idZona:          document.getElementById('guiaZona').value,
      numeroDocumento: document.getElementById('guiaNumDoc').value,
      comentario
    };

    // Optimista con animación pulsante
    const tempId     = 'G_opt_' + Date.now();
    const provNombre = _getProvNombre(idProveedor);
    injectOptimisticGuia({ tempId, idProveedor, provNombre });
    cerrarSheet('sheetGuia');

    const res = await API.crearGuia(params);
    if (res.ok) {
      finalizeOptimisticGuia(tempId, res.data?.idGuia, tipo, provNombre);
      toast(`Guía ${res.data?.idGuia || 'nueva'} creada`, 'ok');
    } else if (!res.offline) {
      removeOptimisticGuia(tempId);
      toast('Error: ' + res.error, 'danger');
    }
  }

  function nueva() {
    _guiaModoEdicion = false;
    _resetSheetGuiaZIndex();
    // Reset título y botón a modo creación
    const titleEl = document.getElementById('guiaFormTitle');
    if (titleEl) titleEl.textContent = '📋 Nueva Guía';
    const btnEl = document.getElementById('btnGuiaSubmit');
    if (btnEl) { btnEl.textContent = 'Crear guía'; btnEl.onclick = () => GuiasView.crearGuia(); }
    // Reset tags de creación
    _tagsNueva = { comp: null, compl: null };
    ['nTagComp1','nTagComp0','nTagCompl1','nTagCompl0'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.className = 'flex-1 py-2 rounded-lg text-xs font-bold border transition-all border-slate-700 text-slate-500';
    });
    const comInput = document.getElementById('guiaComentario');
    if (comInput) comInput.value = '';
    // Poblar proveedor select (solo la primera vez)
    const provSel = document.getElementById('guiaProveedor');
    if (provSel) {
      if (provSel.options.length <= 1) {
        const provs = OfflineManager.getProveedoresCache();
        provs.forEach(p => {
          const opt = document.createElement('option');
          opt.value = p.idProveedor;
          opt.textContent = p.nombre || p.idProveedor;
          provSel.appendChild(opt);
        });
      }
      provSel.value = ''; // siempre limpiar selección
    }
    // Poblar zonas select (dinámico desde caché)
    const zonaEl = document.getElementById('guiaZona');
    if (zonaEl) {
      const zonas = OfflineManager.getZonasCache();
      // Reconstruir siempre para reflejar cambios en Sheets
      zonaEl.innerHTML = '<option value="">— Seleccionar —</option>';
      zonas.forEach(z => {
        const opt = document.createElement('option');
        opt.value = z.idZona;
        opt.textContent = z.nombre || z.idZona;
        zonaEl.appendChild(opt);
      });
      zonaEl.value = '';
    }
    // Reset tipo/prov/zona rows
    const tipoEl = document.getElementById('guiaTipo');
    if (tipoEl) {
      tipoEl.value = 'INGRESO_PROVEEDOR';
      document.getElementById('guiaZonaRow')?.classList.add('hidden');
      document.getElementById('guiaProvRow')?.classList.remove('hidden');
    }
    abrirSheet('sheetGuia');
  }

  // ── Foto guía (sube foto + actualiza columna en GAS) ─────
  function onFotoGuiaSeleccionada(event) {
    const file = event.target.files?.[0];
    if (!file || !_guiaActual) return;
    event.target.value = '';
    const btn = document.getElementById('guiaDetFotoSection');
    if (btn) btn.innerHTML = '<div class="flex justify-center py-3"><div class="spinner"></div></div>';

    _prepararFotoGuia(file).then(({ b64, mime }) =>
      API.subirFotoGuia({ idGuia: _guiaActual.idGuia, fotoBase64: b64, mimeType: mime })
    ).then(res => {
      if (res.ok && !res.offline && res.data?.url) {
        _guiaActual.foto = res.data.url;
        _mostrarDetalleSheet(_guiaActual, false);
        toast('Foto guardada', 'ok', 1500);
      } else {
        toast('Error al subir foto', 'danger');
        _mostrarDetalleSheet(_guiaActual, false);
      }
    }).catch(() => { toast('Error al subir foto', 'danger'); _mostrarDetalleSheet(_guiaActual, false); });
  }

  async function copiarFotoDePreingreso() {
    if (!_guiaActual?.idPreingreso) { toast('Sin preingreso vinculado', 'warn'); return; }
    const fotoEl = document.getElementById('guiaDetFotoSection');
    if (fotoEl) fotoEl.innerHTML = '<div class="flex justify-center py-3"><div class="spinner"></div></div>';
    const res = await API.copiarFotoDePreingreso({ idGuia: _guiaActual.idGuia, idPreingreso: _guiaActual.idPreingreso })
      .catch(() => ({ ok: false, error: 'Sin conexión' }));
    if (res.ok && res.data?.url) {
      _guiaActual.foto = res.data.url;
      _mostrarDetalleSheet(_guiaActual, false);
      toast('Foto copiada', 'ok', 1500);
    } else {
      toast('Error: ' + (res.error || 'no se pudo copiar'), 'danger');
      _mostrarDetalleSheet(_guiaActual, false);
    }
  }

  function _prepararFotoGuia(file) {
    return new Promise((resolve, reject) => {
      const MAX = 1280;
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = ev => {
        const img = new Image();
        img.onerror = reject;
        img.onload = () => {
          let w = img.width, h = img.height;
          if (w > MAX || h > MAX) {
            if (w >= h) { h = Math.round(h * MAX / w); w = MAX; }
            else        { w = Math.round(w * MAX / h); h = MAX; }
          }
          const canvas = document.createElement('canvas');
          canvas.width = w; canvas.height = h;
          canvas.getContext('2d').drawImage(img, 0, 0, w, h);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.82);
          resolve({ b64: dataUrl.split(',')[1], mime: 'image/jpeg' });
        };
        img.src = ev.target.result;
      };
      reader.readAsDataURL(file);
    });
  }

  // ── Comentario + tags guía ────────────────────────────────
  function toggleTagGuia(grupo, valor) {
    _tagsGuia[grupo] = (_tagsGuia[grupo] === valor) ? null : valor;
    // Actualizar clases de los 4 botones sin re-renderizar el sheet
    const configs = [
      { id:'gTagComp1',  g:'comp',  v:'si',  a:'bg-blue-900/70 border-blue-500 text-blue-200',   i:'border-slate-700 text-slate-500' },
      { id:'gTagComp0',  g:'comp',  v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
      { id:'gTagCompl1', g:'compl', v:'si',  a:'bg-green-900/70 border-green-500 text-green-200', i:'border-slate-700 text-slate-500' },
      { id:'gTagCompl0', g:'compl', v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
    ];
    const base = 'flex-1 py-2 rounded-lg text-xs font-bold border transition-all';
    configs.forEach(({ id, g, v, a, i }) => {
      const el = document.getElementById(id);
      if (el) el.className = `${base} ${_tagsGuia[g] === v ? a : i}`;
    });
  }

  // Tags en NUEVA guía
  function toggleTagNueva(grupo, valor) {
    _tagsNueva[grupo] = (_tagsNueva[grupo] === valor) ? null : valor;
    const cfgs = [
      { id:'nTagComp1',  g:'comp',  v:'si',  a:'bg-blue-900/70 border-blue-500 text-blue-200',   i:'border-slate-700 text-slate-500' },
      { id:'nTagComp0',  g:'comp',  v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
      { id:'nTagCompl1', g:'compl', v:'si',  a:'bg-green-900/70 border-green-500 text-green-200', i:'border-slate-700 text-slate-500' },
      { id:'nTagCompl0', g:'compl', v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
    ];
    const base = 'flex-1 py-2 rounded-lg text-xs font-bold border transition-all';
    cfgs.forEach(({ id, g, v, a, i }) => {
      const el = document.getElementById(id);
      if (el) el.className = `${base} ${_tagsNueva[g] === v ? a : i}`;
    });
  }

  // ── Toggle paneles header ────────────────────────────────
  function toggleFotoPanel() {
    _fotoOpen = !_fotoOpen;
    if (_fotoOpen) _notasOpen = false;
    _mostrarDetalleSheet(_guiaActual, false);
  }

  function toggleNotasPanel() {
    _notasOpen = !_notasOpen;
    if (_notasOpen) _fotoOpen = false;
    _mostrarDetalleSheet(_guiaActual, false);
  }

  // ── Eliminar foto de guía ─────────────────────────────────
  async function eliminarFotoGuia() {
    if (!_guiaActual?.foto) return;
    const url = _guiaActual.foto;
    // Extraer fileId de la URL de Drive (ej. ?id=FILE_ID&...)
    const match = url.match(/[?&]id=([^&]+)/);
    // Optimista: limpiar foto localmente
    _guiaActual.foto = '';
    const idx = todas.findIndex(g => g.idGuia === _guiaActual.idGuia);
    if (idx >= 0) todas[idx].foto = '';
    _mostrarDetalleSheet(_guiaActual, false);
    toast('Foto eliminada', 'warn', 1500);
    // Background: eliminar archivo + limpiar columna en sheet
    if (match) API.eliminarFotoDrive({ fileId: match[1] }).catch(() => {});
    API.actualizarGuia({ idGuia: _guiaActual.idGuia, foto: '' }).catch(() => {});
  }

  // ── Editar guía existente ────────────────────────────────
  function editarGuia() {
    if (!_guiaActual) return;
    _guiaModoEdicion = true;
    const g = _guiaActual;

    // Poblar proveedor select si hace falta
    const provSel = document.getElementById('guiaProveedor');
    if (provSel && provSel.options.length <= 1) {
      const provs = OfflineManager.getProveedoresCache();
      provs.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.idProveedor;
        opt.textContent = p.nombre || p.idProveedor;
        provSel.appendChild(opt);
      });
    }

    // Pre-llenar campos
    const tipoEl = document.getElementById('guiaTipo');
    if (tipoEl) {
      tipoEl.value = g.tipo || 'INGRESO_PROVEEDOR';
      // Disparar cambio visual de filas proveedor/zona
      const isZona   = tipoEl.value === 'SALIDA_ZONA';
      const isIngProv = tipoEl.value === 'INGRESO_PROVEEDOR';
      document.getElementById('guiaZonaRow')?.classList.toggle('hidden', !isZona);
      document.getElementById('guiaProvRow')?.classList.toggle('hidden', !isIngProv);
    }
    if (provSel) provSel.value = g.idProveedor || '';
    const zonaEl = document.getElementById('guiaZona');
    if (zonaEl) {
      const zonas = OfflineManager.getZonasCache();
      zonaEl.innerHTML = '<option value="">— Seleccionar —</option>';
      zonas.forEach(z => {
        const opt = document.createElement('option');
        opt.value = z.idZona;
        opt.textContent = z.nombre || z.idZona;
        zonaEl.appendChild(opt);
      });
      zonaEl.value = g.idZona || '';
    }
    const numDocEl = document.getElementById('guiaNumDoc');
    if (numDocEl) numDocEl.value = g.numeroDocumento || '';

    // Tags y texto libre del comentario
    _tagsNueva = { ..._tagsFromComentario(g.comentario) };
    const textoLibre = _textoLibreFromComentario(g.comentario);
    const comInput = document.getElementById('guiaComentario');
    if (comInput) comInput.value = textoLibre;

    // Sincronizar clases de botones de tag
    const cfgs = [
      { id:'nTagComp1',  g:'comp',  v:'si',  a:'bg-blue-900/70 border-blue-500 text-blue-200',   i:'border-slate-700 text-slate-500' },
      { id:'nTagComp0',  g:'comp',  v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
      { id:'nTagCompl1', g:'compl', v:'si',  a:'bg-green-900/70 border-green-500 text-green-200', i:'border-slate-700 text-slate-500' },
      { id:'nTagCompl0', g:'compl', v:'no',  a:'bg-amber-900/70 border-amber-500 text-amber-200', i:'border-slate-700 text-slate-500' },
    ];
    const base = 'flex-1 py-2 rounded-lg text-xs font-bold border transition-all';
    cfgs.forEach(({ id, g: grp, v, a, i }) => {
      const el = document.getElementById(id);
      if (el) el.className = `${base} ${_tagsNueva[grp] === v ? a : i}`;
    });

    // Cambiar título y botón
    const titleEl = document.getElementById('guiaFormTitle');
    if (titleEl) titleEl.textContent = '✏️ Editar Guía';
    const btnEl = document.getElementById('btnGuiaSubmit');
    if (btnEl) { btnEl.textContent = 'GUARDAR CAMBIOS'; btnEl.onclick = () => GuiasView.guardarCambiosGuia(); }

    // Elevar z-index para aparecer encima del sheet de detalle
    const shG  = document.getElementById('sheetGuia');
    const ovG  = document.getElementById('overlayGuia');
    if (shG)  shG.style.zIndex  = '60';
    if (ovG) {
      ovG.style.zIndex = '59';
      ovG.onclick = () => { _resetSheetGuiaZIndex(); cerrarSheet('sheetGuia'); };
    }
    abrirSheet('sheetGuia');
  }

  function _resetSheetGuiaZIndex() {
    const shG = document.getElementById('sheetGuia');
    const ovG = document.getElementById('overlayGuia');
    if (shG) shG.style.zIndex  = '';
    if (ovG) { ovG.style.zIndex = ''; ovG.onclick = () => cerrarSheet('sheetGuia'); }
  }

  async function guardarCambiosGuia() {
    if (!_guiaActual) return;
    const tipo        = document.getElementById('guiaTipo').value;
    const idProveedor = document.getElementById('guiaProveedor').value;
    const idZona      = document.getElementById('guiaZona').value;
    const numDoc      = document.getElementById('guiaNumDoc').value;
    const textoExtra  = (document.getElementById('guiaComentario')?.value || '').trim();
    const comentario  = _buildComentario(_tagsNueva, textoExtra);

    // Actualizar optimistamente
    _guiaActual.tipo            = tipo;
    _guiaActual.idProveedor     = idProveedor;
    _guiaActual.idZona          = idZona;
    _guiaActual.numeroDocumento = numDoc;
    _guiaActual.comentario      = comentario;

    _resetSheetGuiaZIndex();
    cerrarSheet('sheetGuia');
    _mostrarDetalleSheet(_guiaActual, false);

    const idx = todas.findIndex(g => g.idGuia === _guiaActual.idGuia);
    if (idx >= 0) Object.assign(todas[idx], { tipo, idProveedor, idZona, numeroDocumento: numDoc, comentario });

    API.actualizarGuia({ idGuia: _guiaActual.idGuia, tipo, idProveedor, idZona, numeroDocumento: numDoc, comentario })
      .catch(() => toast('Error al guardar cambios', 'danger'));

    toast('Guía actualizada', 'ok', 1500);
  }

  // Auto-guardar comentario al cerrar el sheet de detalle
  function cerrarGuiaDetalle() {
    if (_guiaActual && _guiaActual.estado === 'ABIERTA') {
      const textoExtra = document.getElementById('guiaComentarioEdit')?.value || '';
      const nuevoComentario = _buildComentario(_tagsGuia, textoExtra);
      if (nuevoComentario !== (_guiaActual.comentario || '')) {
        _guiaActual.comentario = nuevoComentario;
        API.actualizarGuia({ idGuia: _guiaActual.idGuia, comentario: nuevoComentario }).catch(() => {});
        // Refrescar card en lista
        const idx = todas.findIndex(g => g.idGuia === _guiaActual.idGuia);
        if (idx >= 0) { todas[idx].comentario = nuevoComentario; }
      }
    }
    cerrarSheet('sheetGuiaDetalle');
  }

  // Navegar al preingreso vinculado
  function irAPreingreso(idPreingreso) {
    cerrarGuiaDetalle();
    App.nav('preingresos');
    setTimeout(() => {
      const cached = OfflineManager.getPreingresosCache();
      if (cached.find(p => p.idPreingreso === idPreingreso)) {
        PreingresosView.abrirDetalle(idPreingreso);
      } else {
        // Cargar y luego abrir
        PreingresosView.cargar().then(() => PreingresosView.abrirDetalle(idPreingreso)).catch(() => {});
      }
    }, 380);
  }

  // Abrir foto guía en carrusel (usa _guiaActual)
  function verFotoGuia() {
    if (!_guiaActual?.foto) { toast('Sin foto', 'info'); return; }
    abrirCarrusel([_normalizeDriveUrl(_guiaActual.foto)], _guiaActual.idGuia);
  }

  return {
    cargar, filtrar, toggleFiltro, silentRefresh, verDetalle,
    buscar, buscarClear,
    abrirAgregarItem, abrirCamaraItem, abrirScannerItem,
    cerrarCamara, cerrarCamPicker, seleccionarItemCamara,
    cerrarScannerItem, _enfocarHid,
    seleccionarItemHid,
    // compat stubs
    cerrarScanInput, onHidKey, procesarHidInput, usarCamara,
    _procesarCodigoEscaneado: _buscarCandidatos, _rescanear,
    toggleEstadoGuia, adminPinTecla, adminPinAtras,
    confirmarCerrarGuia, crearGuia, nueva,
    toggleTagNueva,
    onFotoGuiaSeleccionada, copiarFotoDePreingreso, verFotoGuia,
    toggleTagGuia, cerrarGuiaDetalle, irAPreingreso,
    injectOptimisticGuia, finalizeOptimisticGuia, removeOptimisticGuia,
    selectItem, deselectItem,
    inlineQtyDelta, inlineQtyInput, inlineQtyBlur,
    inlinePickVenc, inlineVencChanged, inlineDelete,
    toggleFotoPanel, toggleNotasPanel, editarGuia, guardarCambiosGuia,
    eliminarFotoGuia,
    filtrarChip,
    compartirWA,
    imprimirTicket,
    abrirModalPN, cerrarModalPN, _pnFotoSeleccionada, confirmarRegistrarPN,
    abrirPNSinCodigo, _ocultarPNOffer
  };

  function imprimirTicket(idGuia) {
    const base       = location.href.split('?')[0].replace(/index\.html$/, '').replace(/\/$/, '');
    const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(idGuia)}`;
    vibrate(10);
    toast('Enviando a impresora...', 'ok', 2000);
    API.imprimirTicketGuia({ idGuia, reporteUrl })
      .then(res => {
        if (!res.ok) toast('Error al imprimir: ' + (res.error || ''), 'warn', 4000);
      })
      .catch(() => toast('Error al imprimir', 'warn', 3000));
  }

  function compartirWA(idGuia) {
    const g = (OfflineManager.getGuiasCache() || []).find(x => x.idGuia === idGuia);
    if (!g) return;
    const prov      = (OfflineManager.getProveedoresCache() || []).find(x => x.idProveedor === g.idProveedor);
    const provNombre = prov?.nombre || g.idProveedor || g.usuario || '—';
    const detCache  = OfflineManager.getGuiaDetalleCache() || [];
    const numItems  = detCache.filter(d => d.idGuia === idGuia && d.observacion !== 'ANULADO').length;
    const TIPO_LABELS_WA = {
      INGRESO_PROVEEDOR:'Ingreso Proveedor', INGRESO_JEFATURA:'Ingreso Jefatura',
      SALIDA_ZONA:'Salida Zona', SALIDA_DEVOLUCION:'Devolución',
      SALIDA_JEFATURA:'Salida Jefatura', SALIDA_ENVASADO:'Envasado', SALIDA_MERMA:'Merma'
    };
    const url    = `${location.href.split('?')[0].replace(/index\.html$/, '').replace(/\/$/, '')}/reporte.html?tipo=guia&id=${encodeURIComponent(idGuia)}`;
    const lineas = [
      `*📋 GUÍA ${idGuia}*`,
      `─────────────────────`,
      `📦 *Tipo:* ${TIPO_LABELS_WA[g.tipo] || g.tipo || '—'}`,
      `🏪 *Proveedor:* ${provNombre}`,
      `📅 *Fecha:* ${_fmtCorta(g.fecha)}`,
      `📊 *Estado:* ${g.estado || '—'}`,
      `👤 *Usuario:* ${g.usuario || '—'}`,
      `📝 *Ítems:* ${numItems}`,
    ];
    if (g.comentario)   lineas.push(`💬 *Comentario:* ${g.comentario}`);
    if (g.idPreingreso) lineas.push(`🔗 *Preingreso:* ${g.idPreingreso}`);
    lineas.push(
      ``,
      `━━━━━━━━━━━━━━━━━━━━━━━━`,
      `👇 *TOCA AQUÍ PARA VER EL REPORTE COMPLETO*`,
      url,
      `━━━━━━━━━━━━━━━━━━━━━━━━`,
      `_InversionMos Warehouse_`
    );
    window.open('https://wa.me/?text=' + encodeURIComponent(lineas.join('\n')), '_blank');
  }

  // Método para chips de filtro (actualiza estado visual de chips + llama filtrar())
  function filtrarChip(chipEl, filtro) {
    document.querySelectorAll('#guiaChips .chip').forEach(c => c.classList.remove('active'));
    chipEl.classList.add('active');
    vibrate(8);
    filtrar(filtro);
  }
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

  function nuevo(preIdBase, preIdDerivado) {
    productosMaestro = App.getProductosMaestro();
    derivados = productosMaestro.filter(p => p.codigoProductoBase && p.codigoProductoBase !== '');

    const sel = document.getElementById('envProductoDerivado');
    sel.innerHTML = '<option value="">— Seleccionar producto a envasar —</option>';
    derivados.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.idProducto;
      opt.textContent = p.descripcion;
      sel.appendChild(opt);
    });

    // Auto-fill fecha vencimiento = hoy + 12 meses
    const fv = new Date();
    fv.setFullYear(fv.getFullYear() + 1);
    document.getElementById('envFechaVenc').value = fv.toISOString().split('T')[0];

    // Reset
    document.getElementById('envUnidades').value = 1;
    document.getElementById('envNombreProducto').textContent = '';
    document.getElementById('envasadoFactorInfo').classList.add('hidden');

    if (preIdDerivado) {
      sel.value = preIdDerivado;
      sel.disabled = true;
      sel.classList.add('hidden');
      onDerivadoChange(preIdDerivado);
    } else {
      sel.disabled = false;
      sel.classList.remove('hidden');
    }

    abrirSheet('sheetEnvasado');
  }

  function onDerivadoChange(idDerivado) {
    const prod = derivados.find(p => p.idProducto === idDerivado);
    if (!prod) {
      document.getElementById('envNombreProducto').textContent = '';
      document.getElementById('envasadoFactorInfo').classList.add('hidden');
      return;
    }

    const factorBase = parseFloat(prod.factorConversionBase) || 0;

    const prodBase = productosMaestro.find(p =>
      (p.skuBase && p.skuBase === prod.codigoProductoBase) ||
      p.idProducto === prod.codigoProductoBase
    );

    // Stock desde caché local usando codigoBarra del base
    const cbBase     = prodBase ? String(prodBase.codigoBarra) : '';
    const stockEntry = OfflineManager.getStockCache().find(s => String(s.codigoProducto) === cbBase);
    const stockBase  = stockEntry ? (parseFloat(stockEntry.cantidadDisponible) || 0) : 0;
    const unidadBase = prodBase ? (prodBase.unidad || '') : '';
    const maxProd    = factorBase > 0 ? Math.floor(stockBase / factorBase) : 0;

    document.getElementById('envNombreProducto').textContent = prod.descripcion;
    document.getElementById('envInfoStock').textContent      = `${fmt(stockBase, 1)} ${unidadBase}`;
    document.getElementById('envInfoMax').textContent        = `${maxProd} uds`;
    document.getElementById('envasadoFactorInfo').classList.remove('hidden');
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
  }

  function setUnidades(n) {
    document.getElementById('envUnidades').value = n;
  }

  async function registrar() {
    const idDerivado = document.getElementById('envProductoDerivado').value;
    const producidas = parseInt(document.getElementById('envUnidades').value) || 0;
    const fechaVenc  = document.getElementById('envFechaVenc').value;
    const imprimir   = document.getElementById('envImprimirEtiq').checked;

    if (!idDerivado || producidas <= 0) {
      toast('Selecciona producto e ingresa las unidades producidas', 'warn');
      return;
    }

    const prod     = derivados.find(p => p.idProducto === idDerivado);
    const prodBase = productosMaestro.find(p =>
      (p.skuBase && p.skuBase === prod.codigoProductoBase) ||
      p.idProducto === prod.codigoProductoBase
    );
    const btn = document.getElementById('btnRegistrarEnvasado');
    btn.disabled = true;
    btn.textContent = 'Registrando...';

    // Optimistic: cerrar modal y avisar de inmediato
    toast(`${producidas} uds registradas${imprimir ? ' · enviando etiquetas...' : ''}`, 'ok', 4000);
    cerrarSheet('sheetEnvasado');
    cargar();

    // Patch stock cache local para actualizar UI de inmediato
    const factorBase = parseFloat(prod.factorConversionBase) || 0;
    const cantBase   = producidas * factorBase;
    if (prodBase?.codigoBarra) OfflineManager.patchStockCache(String(prodBase.codigoBarra), -cantBase);
    OfflineManager.patchStockCache(String(prod.codigoBarra), producidas);
    window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['stock'] } }));

    // GAS en segundo plano
    API.registrarEnvasado({
      codigoBarra:        prod.codigoBarra,
      unidadesProducidas: producidas,
      fechaVencimiento:   fechaVenc,
      imprimirEtiquetas:  imprimir,
      usuario:            window.WH_CONFIG.usuario
    }).then(res => {
      if (!res.ok) {
        toast('Error al guardar envasado: ' + res.error, 'danger', 7000);
      } else if (imprimir && res.data?.impresion && !res.data.impresion.ok) {
        toast('Impresora: ' + res.data.impresion.error, 'warn', 5000);
      }
      OfflineManager.precargarOperacional(true).catch(() => {});
    }).catch(() => {
      toast('Sin conexion — reintenta en un momento', 'warn', 5000);
    }).finally(() => {
      btn.disabled = false;
      btn.textContent = 'Registrar envasado';
    });
  }

  return { cargar, nuevo, onDerivadoChange, calcularProyeccion, ajustarUnidades, setUnidades, registrar };
})();

// ════════════════════════════════════════════════
// MODO ENVASADOR — catálogo por producto base
// ════════════════════════════════════════════════
const EnvasadorView = (() => {
  let _filtroUrg  = false;
  let _catalog    = [];
  let _timer      = null;

  function _buildStockMap() {
    const map = {};
    OfflineManager.getStockCache().forEach(s => {
      map[s.codigoProducto || s.idProducto] = s;
    });
    return map;
  }

  function _urgencia(stockActual, stockMin) {
    if (!stockMin || stockMin <= 0) return null;
    if (stockActual <= 0) return 'CRITICA';
    if (stockActual < stockMin) return 'ALTA';
    return null;
  }

  function _buildCatalog() {
    const master   = App.getProductosMaestro();
    const stockMap = _buildStockMap();

    const bases = master.filter(p =>
      String(p.esEnvasable) === '1' &&
      p.estado !== '0' && p.estado !== 0
    );

    return bases.map(base => {
      const bs        = stockMap[String(base.codigoBarra)] || {};
      const stockBase = parseFloat(bs.cantidadDisponible || 0);
      const unidad    = base.unidad || bs.unidad || '';

      const derivados = master
        .filter(d =>
          (d.codigoProductoBase === base.idProducto || d.codigoProductoBase === base.skuBase) &&
          d.estado !== '0' && d.estado !== 0
        )
        .map(d => {
          const factorBase = parseFloat(d.factorConversionBase || 0);
          const merma      = parseFloat(d.mermaEsperadaPct || 0);
          const posibles   = factorBase > 0 ? Math.floor(stockBase / factorBase) : 0;
          const ds         = stockMap[String(d.codigoBarra)] || {};
          const stockD     = parseFloat(ds.cantidadDisponible || 0);
          const stockMin = parseFloat(ds.stockMinimo || d.stockMinimo || 0);
          return { ...d, stockD, stockMin, posibles, factorBase, merma,
                   urgencia: _urgencia(stockD, stockMin) };
        });

      const worstUrg = derivados.some(d => d.urgencia === 'CRITICA') ? 'CRITICA'
                     : derivados.some(d => d.urgencia === 'ALTA')    ? 'ALTA' : null;

      return { ...base, stockBase, unidad, derivados, worstUrg };
    }).filter(b => b.derivados.length > 0);
  }

  function _urgIcon(urg) {
    if (urg === 'CRITICA') return '<span style="color:#ef4444">⚡</span>';
    if (urg === 'ALTA')    return '<span style="color:#f59e0b">⚡</span>';
    return '';
  }

  // Extrae solo la parte diferenciadora del nombre del derivado respecto al base.
  // Ej: base="AJO EN POLVO GRANEL", derivado="AJO EN POLVO 250GR" → "250GR"
  function _sufijo(baseNomUp, derivDesc) {
    const derUp  = derivDesc.toUpperCase();
    const bWords = baseNomUp.replace(/GRANEL|BULK|BASE/g, '').trim().split(/\s+/).filter(Boolean);
    const dWords = derUp.split(/\s+/).filter(Boolean);
    // Eliminar palabras del base que aparezcan en el derivado (orden no importa)
    const bSet   = new Set(bWords);
    const resto  = dWords.filter(w => !bSet.has(w));
    return resto.length > 0 ? resto.join(' ') : derivDesc;
  }

  function _render() {
    const container = document.getElementById('listEnvasadorCatalog');
    if (!container) return;
    const list = _filtroUrg ? _catalog.filter(b => b.worstUrg) : _catalog;

    if (!list.length) {
      container.innerHTML = _filtroUrg
        ? '<div class="card text-center py-8"><p class="text-2xl mb-2">✅</p><p class="font-semibold">¡Sin urgentes!</p></div>'
        : '<p class="text-slate-500 text-center py-8 text-sm">Sin productos envasables configurados</p>';
      return;
    }

    container.innerHTML = list.map(base => {
      const urgHdr  = base.worstUrg ? _urgIcon(base.worstUrg) + ' ' : '';
      const baseNom = String(base.descripcion || '').toUpperCase();

      const derivRows = base.derivados
        .filter(d => !_filtroUrg || d.urgencia)
        .map(d => {
          const urg     = d.urgencia ? _urgIcon(d.urgencia) + ' ' : '';
          const label   = _sufijo(baseNom, String(d.descripcion || d.idProducto));
          const stockEl = d.stockD > 0
            ? `<span class="font-bold text-slate-200" style="font-size:13px">${fmt(d.stockD,1)} ${escHtml(d.unidad||'')}</span>`
            : `<span class="font-bold text-red-400" style="font-size:13px">0 ${escHtml(d.unidad||'')}</span>`;
          const posiblesHtml = d.posibles > 0
            ? `<span class="text-emerald-400 text-xs">~${fmt(d.posibles)} posibles</span>`
            : `<span class="text-slate-500 text-xs">Base insuficiente</span>`;
          const urgTag = d.urgencia
            ? `<span class="tag-${d.urgencia === 'CRITICA' ? 'danger' : 'warn'} text-xs">${d.urgencia}</span>`
            : '';
          return `
          <div class="env-deriv-row">
            <div class="flex items-center gap-2 flex-wrap">
              <span class="font-semibold text-sm text-slate-100 flex-1">${urg}${escHtml(label)}</span>
              ${stockEl}
              ${urgTag}
            </div>
            <div class="flex items-center gap-2 mt-1">
              ${posiblesHtml}
              <button onclick="EnvasadorView.envasar('${escAttr(base.idProducto)}','${escAttr(d.idProducto)}')"
                      class="btn btn-sm btn-primary ml-auto" style="flex-shrink:0">Envasar</button>
            </div>
          </div>`;
        }).join('');

      return `
      <div class="card env-base-card">
        <div class="flex items-start justify-between cursor-pointer mb-1 gap-2"
             onclick="this.closest('.env-base-card').classList.toggle('env-collapsed')">
          <div class="flex-1 min-w-0">
            <p class="font-bold text-base leading-snug">${urgHdr}${escHtml(base.descripcion)}</p>
            <p class="text-xs text-slate-400 mt-0.5">${fmt(base.stockBase,1)} ${escHtml(base.unidad)} disponible</p>
          </div>
          <svg class="env-chevron w-4 h-4 text-slate-400 flex-shrink-0 mt-1" viewBox="0 0 16 16" fill="currentColor">
            <path d="M1.646 4.646a.5.5 0 0 1 .708 0L8 10.293l5.646-5.647a.5.5 0 0 1 .708.708l-6 6a.5.5 0 0 1-.708 0l-6-6a.5.5 0 0 1 0-.708z"/>
          </svg>
        </div>
        <div class="env-deriv-list">${derivRows}</div>
      </div>`;
    }).join('');
  }

  function _updateUrgBtn() {
    const btn = document.getElementById('envUrgBtn');
    if (!btn) return;
    const urgN = _catalog.filter(b => b.worstUrg).length;
    document.getElementById('envUrgCount').textContent = urgN;
    btn.style.display = urgN > 0 ? 'flex' : 'none';
    btn.classList.toggle('active', _filtroUrg);
  }

  function toggleUrgFilter() {
    _filtroUrg = !_filtroUrg;
    _updateUrgBtn();
    _render();
  }

  function cargar() {
    if (_timer) { clearInterval(_timer); _timer = null; }
    _filtroUrg = false;
    document.getElementById('envCatalogPanel').classList.remove('hidden');
    document.getElementById('envHistorialPanel').classList.add('hidden');
    document.getElementById('listEnvasadorCatalog').innerHTML =
      '<div class="flex justify-center py-8"><div class="spinner"></div></div>';

    _catalog = _buildCatalog();
    if (_catalog.some(b => b.worstUrg)) _filtroUrg = true;
    _updateUrgBtn();
    _render();

    _timer = setInterval(() => {
      _catalog = _buildCatalog();
      _updateUrgBtn();
      _render();
    }, 120_000);
  }

  async function verHistorial() {
    document.getElementById('envCatalogPanel').classList.add('hidden');
    document.getElementById('envHistorialPanel').classList.remove('hidden');
    const container = document.getElementById('listEnvasadorHistorial');
    container.innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';
    const res = await API.getEnvasados({ limit: 30 }).catch(() => ({ ok: false }));
    const list = res.ok ? res.data : [];
    if (!list.length) { container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin historial</p>'; return; }
    container.innerHTML = list.map(e => {
      const efPct   = parseFloat(e.eficienciaPct) || 0;
      const efColor = efPct >= 95 ? 'text-emerald-400' : efPct >= 85 ? 'text-amber-400' : 'text-red-400';
      return `
      <div class="card-sm">
        <div class="flex items-center justify-between mb-1">
          <span class="text-xs tag-blue">${escHtml(e.codigoProductoBase)} → ${escHtml(e.codigoProductoEnvasado)}</span>
          <span class="${efColor} font-bold text-sm">${efPct}%</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(e.fecha)} · ${escHtml(e.usuario)}</p>
        <div class="flex gap-4 mt-1 text-xs text-slate-300">
          <span>Base: ${fmt(e.cantidadBase, 1)}</span>
          <span>Prod: ${fmt(e.unidadesProducidas)} uds</span>
          <span class="text-amber-400">Merma: ${fmt(e.mermaReal)}</span>
        </div>
      </div>`;
    }).join('');
  }

  function verCatalogo() {
    document.getElementById('envHistorialPanel').classList.add('hidden');
    document.getElementById('envCatalogPanel').classList.remove('hidden');
  }

  function envasar(idBase, idDerivado) {
    EnvasadosView.nuevo(idBase, idDerivado);
  }

  function silentRefresh() {
    _catalog = _buildCatalog();
    _updateUrgBtn();
    _render();
  }

  return { cargar, toggleUrgFilter, verHistorial, verCatalogo, envasar, silentRefresh };
})();

// ════════════════════════════════════════════════
// DESPACHO RÁPIDO
// ════════════════════════════════════════════════
const DespachoView = (() => {
  const CART_KEY = 'wh_despacho_cart';
  let _cart     = [];
  let _camHidden = false;

  function _saveCart() { localStorage.setItem(CART_KEY, JSON.stringify(_cart)); }
  function _loadCart() {
    try { return JSON.parse(localStorage.getItem(CART_KEY) || '[]'); } catch { return []; }
  }

  // ── Cámara continua ──────────────────────────────────────────
  function _startCamera() {
    const status = document.getElementById('despCamStatus');
    Scanner.start('despVideo', cod => _procesarCodigo(String(cod).trim()),
      err => { if (status) status.textContent = '⚠ ' + err; },
      { continuous: true, cooldown: 1000 }
    );
    if (status) status.textContent = '● Escaneando';
  }

  function pauseCamera() { Scanner.stop(); }

  function toggleCamera() {
    _camHidden = !_camHidden;
    document.getElementById('despCamPanel').style.display   = _camHidden ? 'none' : 'block';
    document.getElementById('despCamShowBtn').style.display = _camHidden ? 'block' : 'none';
    if (_camHidden) {
      Scanner.stop();
    } else {
      _startCamera();
    }
  }

  function cargar() {
    _cart = _loadCart();
    _renderCart();
    _updateFooter();
    badgeUpdate();
    // Iniciar cámara si no estaba oculta
    if (!_camHidden && !Scanner.isActive()) _startCamera();
  }

  // ── Procesar código escaneado ────────────────────────────────
  function _procesarCodigo(cod) {
    if (!cod) return;
    const master = App.getProductosMaestro();
    const prod   = master.find(p =>
      String(p.codigoBarra || '').trim() === cod ||
      String(p.idProducto  || '').trim() === cod
    );
    if (!prod) { toast(`"${cod}" no encontrado`, 'warn', 3000); return; }

    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });
    const stockD = parseFloat((stockMap[prod.idProducto] || {}).cantidadDisponible || 0);

    const idx = _cart.findIndex(c => c.codigoBarra === prod.idProducto);
    if (idx >= 0) {
      // Mismo producto → sumar 1
      _cart[idx].cantidad = parseFloat((_cart[idx].cantidad || 0)) + 1;
    } else {
      _cart.push({
        codigoBarra: prod.idProducto,
        descripcion: prod.descripcion || prod.idProducto,
        unidad:      prod.unidad || '',
        cantidad:    1,
        stockDisp:   stockD
      });
    }
    _saveCart();
    _renderCart();
    _updateFooter();
    badgeUpdate();

    // Pulsar fila y focus en input de cantidad
    requestAnimationFrame(() => {
      const row = document.getElementById('despRow-' + CSS.escape(prod.idProducto));
      if (row) {
        row.classList.remove('desp-item-pulse');
        void row.offsetWidth; // fuerza reflow
        row.classList.add('desp-item-pulse');
      }
      const inp = document.getElementById('despQty-' + CSS.escape(prod.idProducto));
      if (inp) inp.focus();
    });
  }

  // ── Render carrito ───────────────────────────────────────────
  function _renderCart() {
    const el = document.getElementById('despCartList');
    if (!el) return;
    if (!_cart.length) {
      el.innerHTML = `
        <div class="card text-center py-10">
          <p class="text-4xl mb-2">📦</p>
          <p class="text-slate-400 text-sm">Apunta la cámara al código de barras</p>
          <p class="text-slate-600 text-xs mt-1">El carrito se guarda automáticamente</p>
        </div>`;
      return;
    }
    el.innerHTML = _cart.map((item, i) => {
      const over     = item.stockDisp > 0 && item.cantidad > item.stockDisp;
      const safeId   = escAttr(item.codigoBarra);
      const overWarn = over ? `<span class="text-xs text-red-400">⚠ stock: ${fmt(item.stockDisp,1)}</span>` : '';
      return `
      <div class="card-sm desp-cart-item flex items-center gap-3"
           id="despRow-${safeId}" style="animation-delay:${i*.03}s">
        <div class="flex-1 min-w-0">
          <p class="font-semibold text-sm truncate">${escHtml(item.descripcion)}</p>
          <p class="text-xs text-slate-400">${escHtml(item.unidad)} ${overWarn}</p>
        </div>
        <div class="desp-qty-wrap">
          <button class="desp-qty-btn" onclick="DespachoView.decQty('${safeId}')">−</button>
          <input  id="despQty-${safeId}" type="number" inputmode="decimal" step="0.1" min="0"
                  value="${item.cantidad}"
                  class="desp-qty-input${over?' over':''}"
                  onblur="DespachoView.blurQty('${safeId}',this.value)"
                  onfocus="this.select()">
          <button class="desp-qty-btn" onclick="DespachoView.incQty('${safeId}')">+</button>
          <button class="desp-qty-btn" style="border-color:rgba(239,68,68,.4);color:#f87171;margin-left:2px"
                  onclick="DespachoView.quitarItem('${safeId}')">✕</button>
        </div>
      </div>`;
    }).join('');
  }

  // ── Controles de cantidad ────────────────────────────────────
  function incQty(cod) { _adjustQty(cod, q => q + 1); }
  function decQty(cod) { _adjustQty(cod, q => Math.max(0, q - 1)); }

  function _adjustQty(cod, fn) {
    const idx = _cart.findIndex(c => c.codigoBarra === cod);
    if (idx < 0) return;
    const newQty = fn(parseFloat(_cart[idx].cantidad) || 0);
    if (newQty <= 0) { quitarItem(cod); return; }
    _cart[idx].cantidad = newQty;
    _saveCart();
    // Actualizar solo el input, sin re-render completo (evita perder foco)
    const inp = document.getElementById('despQty-' + CSS.escape(cod));
    if (inp) {
      inp.value = newQty;
      const over = _cart[idx].stockDisp > 0 && newQty > _cart[idx].stockDisp;
      inp.classList.toggle('over', over);
    }
    _updateFooter();
    badgeUpdate();
  }

  function blurQty(cod, val) {
    const idx = _cart.findIndex(c => c.codigoBarra === cod);
    if (idx < 0) return;
    const newQty = parseFloat(val) || 0;
    if (newQty <= 0) { quitarItem(cod); return; }
    _cart[idx].cantidad = newQty;
    _saveCart();
    _updateFooter();
    badgeUpdate();
    // Actualizar clase over
    const inp = document.getElementById('despQty-' + CSS.escape(cod));
    if (inp) {
      const over = _cart[idx].stockDisp > 0 && newQty > _cart[idx].stockDisp;
      inp.classList.toggle('over', over);
    }
  }

  function quitarItem(cod) {
    _cart = _cart.filter(c => c.codigoBarra !== cod);
    _saveCart();
    _renderCart();
    _updateFooter();
    badgeUpdate();
  }

  function _updateFooter() {
    const n   = _cart.length;
    const uds = _cart.reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    document.getElementById('despProgress').textContent = `${n} producto${n !== 1 ? 's' : ''}`;
    document.getElementById('despTotalUds').textContent = `${fmt(uds,2)} unidades total`;
    const btn = document.getElementById('despBtnFinalizar');
    if (btn) { btn.disabled = n === 0; btn.style.opacity = n === 0 ? '.4' : '1'; }
  }

  function finalizar() {
    if (!_cart.length) return;
    const zonas   = OfflineManager.getZonasCache();
    const sel     = document.getElementById('despZonaSelect');
    sel.innerHTML = '<option value="">— Seleccionar zona —</option>' +
      zonas.map(z => `<option value="${escAttr(z.idZona)}">${escHtml(z.nombre || z.idZona)}</option>`).join('');

    // Pre-seleccionar zona si viene del pickup
    const preZona = _pickupActivo?.idZona;
    if (preZona) sel.value = preZona;

    const n   = _cart.length;
    const uds = _cart.reduce((s, c) => s + c.cantidad, 0);
    document.getElementById('despResumenProds').textContent = `${n} producto${n !== 1 ? 's' : ''}`;
    document.getElementById('despResumenUds').textContent   = `${uds} unidades`;

    abrirSheet('sheetDespFinalizar');
  }

  async function confirmarDespacho() {
    const idZona = document.getElementById('despZonaSelect').value;
    if (!idZona) { toast('Selecciona la zona de destino', 'warn'); return; }

    const btn = document.getElementById('btnConfirmarDespacho');
    btn.disabled     = true;
    btn.textContent  = '⏳ Generando guía...';

    const res = await API.crearDespachoRapido({
      idZona,
      usuario: window.WH_CONFIG?.usuario || '',
      items:   _cart.map(c => ({ codigoBarra: c.codigoBarra, cantidad: c.cantidad }))
    }).catch(() => ({ ok: false, error: 'Sin conexión' }));

    btn.disabled    = false;
    btn.textContent = '📦 Generar guía de salida + Imprimir';

    if (res.ok) {
      const d = res.data;
      const impMsg = d.impresion?.ok ? ' · Imprimiendo...' : ' · Sin impresora';
      toast(`✅ Guía ${d.idGuia} generada${impMsg}`, 'ok', 6000);
      if (d.errores?.length) toast(`⚠ ${d.errores.length} ítem(s) con error`, 'warn', 5000);
      // Si vino de un pickup, marcarlo como COMPLETADO
      if (_pickupActivo) {
        API.actualizarPickup({ idPickup: _pickupActivo.idPickup, estado: 'COMPLETADO' }).catch(() => {});
        _pickupsPendientes = _pickupsPendientes.filter(p => p.idPickup !== _pickupActivo.idPickup);
        _pickupActivo = null;
      }
      _cart = [];
      _saveCart();
      cerrarSheet('sheetDespFinalizar');
      _renderCart();
      _updateFooter();
      badgeUpdate();
    } else {
      toast('Error: ' + (res.error || 'No se pudo crear la guía'), 'danger', 6000);
    }
  }

  function cancelar() {
    if (!_cart.length) return;
    if (!confirm('¿Vaciar el carrito de despacho?')) return;
    _cart = [];
    _saveCart();
    _renderCart();
    _updateFooter();
  }

  // ── Badge global (carrito + pickups pendientes) ─────────────
  let _pickupsPendientes = [];
  let _pollTimer = null;

  function badgeUpdate() {
    _cart = _loadCart();
    const n   = _cart.length;
    const hay = _pickupsPendientes.length > 0;

    // Badge topbar 🛒DSP — visible cuando hay items en el carrito
    const despBadge = document.getElementById('despModoIndicador');
    if (despBadge) despBadge.classList.toggle('hidden', n === 0);

    // Botón FAB en vista Productos
    const fabN = document.getElementById('despCartFabN');
    const dot  = document.getElementById('despPickAlertDot');
    if (fabN) { fabN.textContent = n; fabN.style.display = n > 0 ? 'inline-flex' : 'none'; }
    if (dot)  { dot.style.display = hay ? 'block' : 'none'; }

    // Banner en view-despacho
    const banner    = document.getElementById('despPickupBanner');
    const bannerTxt = document.getElementById('despPickupBannerTxt');
    if (banner && _pickupsPendientes.length > 0) {
      const p = _pickupsPendientes[0];
      if (bannerTxt) bannerTxt.textContent = p.idPickup + (p.notas ? ' · ' + p.notas : '');
      banner.style.display = 'flex';
    } else if (banner) {
      banner.style.display = 'none';
    }
  }

  async function _pollPickups() {
    const res = await API.getPickups().catch(() => ({ ok: false }));
    if (res.ok) _pickupsPendientes = res.data || [];
    badgeUpdate();
  }

  function startPoll() {
    if (_pollTimer) return;
    _pollPickups();
    _pollTimer = setInterval(_pollPickups, 120_000);
  }

  // ── Matching engine (Fase 2) ────────────────────────────────
  function _norm(s) {
    return String(s || '').toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function _score(query, target) {
    const qw = _norm(query).split(' ').filter(w => w.length > 1);
    const tn = _norm(target);
    if (!qw.length) return 0;
    return qw.filter(w => tn.includes(w)).length / qw.length;
  }

  function _bestMatch(nombre) {
    const master = App.getProductosMaestro();
    let best = null, bestScore = 0;
    master.forEach(p => {
      const s = _score(nombre, p.descripcion);
      if (s > bestScore) { bestScore = s; best = p; }
    });
    return { producto: best, score: bestScore };
  }

  // ── Pickup pendiente ────────────────────────────────────────
  let _pickupActivo = null;
  let _matchResults = []; // [{nombre, qty, producto, score, status, accepted}]

  function abrirPickupPendiente() {
    if (!_pickupsPendientes.length) return;
    _pickupActivo = _pickupsPendientes[0];
    _runMatching(_pickupActivo);
    abrirSheet('sheetDespPickup');
  }

  function _runMatching(pickup) {
    const items = pickup.items || [];
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });

    _matchResults = items.map(item => {
      const { producto, score } = _bestMatch(item.nombre);
      const stockDisp = producto ? parseFloat((stockMap[producto.idProducto] || {}).cantidadDisponible || 0) : 0;
      const status = score >= 0.75 ? 'exacto' : score >= 0.35 ? 'parcial' : 'nofound';
      return { nombre: item.nombre, qty: parseInt(item.qty) || 0, producto, score, status, accepted: status === 'exacto', stockDisp };
    });

    // Render las 3 secciones
    document.getElementById('despPickTitulo').textContent = 'Pedido ' + pickup.idPickup;
    document.getElementById('despPickSub').textContent =
      `${pickup.fuente || 'Externo'} · ${items.length} producto${items.length !== 1 ? 's' : ''}`;

    const exactos  = _matchResults.filter(r => r.status === 'exacto');
    const parciales = _matchResults.filter(r => r.status === 'parcial');
    const nofound  = _matchResults.filter(r => r.status === 'nofound');

    const secEx = document.getElementById('despPickSecExactos');
    const secPa = document.getElementById('despPickSecParciales');
    const secNF = document.getElementById('despPickSecNoFound');

    secEx.classList.toggle('hidden', !exactos.length);
    secPa.classList.toggle('hidden', !parciales.length);
    secNF.classList.toggle('hidden', !nofound.length);

    document.getElementById('despPickListExactos').innerHTML = exactos.map(r => `
      <div class="card-sm flex items-center justify-between">
        <div class="min-w-0">
          <p class="text-xs text-slate-400 truncate">${escHtml(r.nombre)}</p>
          <p class="font-semibold text-sm truncate">${escHtml(r.producto?.descripcion || '—')}</p>
        </div>
        <span class="font-bold text-emerald-400 flex-shrink-0 ml-2">×${r.qty}</span>
      </div>`).join('');

    document.getElementById('despPickListParciales').innerHTML = parciales.map((r, i) => {
      const idx = _matchResults.indexOf(r);
      const master = App.getProductosMaestro();
      // Top 3 candidates
      const cands = master
        .map(p => ({ p, s: _score(r.nombre, p.descripcion) }))
        .filter(x => x.s >= 0.2)
        .sort((a, b) => b.s - a.s)
        .slice(0, 4);
      const opts = cands.map(x =>
        `<option value="${escAttr(x.p.idProducto)}" ${x.p === r.producto ? 'selected' : ''}>
          ${escHtml(x.p.descripcion)}
        </option>`).join('');
      return `
      <div class="card-sm">
        <p class="text-xs text-slate-400 mb-1">Pedido: <span class="text-amber-300">"${escHtml(r.nombre)}"</span> ×${r.qty}</p>
        <select class="input text-sm mb-2" onchange="DespachoView.elegirMatchParcial(${idx}, this.value)">
          <option value="">— No despachar —</option>
          ${opts}
        </select>
      </div>`;
    }).join('');

    document.getElementById('despPickListNoFound').innerHTML = nofound.map(r => `
      <div class="card-sm flex items-center justify-between opacity-60">
        <p class="text-sm truncate">${escHtml(r.nombre)}</p>
        <span class="text-xs text-red-400 flex-shrink-0 ml-2">×${r.qty}</span>
      </div>`).join('');

    // Update confirm button label
    const n = exactos.length + parciales.filter(r => r.accepted).length;
    document.getElementById('btnConfirmarPickup').textContent =
      `Cargar ${n} ítem${n !== 1 ? 's' : ''} al carrito`;
  }

  function elegirMatchParcial(idx, codigoBarra) {
    const r = _matchResults[idx];
    if (!r) return;
    if (!codigoBarra) {
      r.accepted = false;
      r.productoOverride = null;
    } else {
      const prod = App.getProductosMaestro().find(p => p.idProducto === codigoBarra);
      r.accepted = true;
      r.productoOverride = prod || r.producto;
    }
    const n = _matchResults.filter(r => r.accepted).length;
    document.getElementById('btnConfirmarPickup').textContent =
      `Cargar ${n} ítem${n !== 1 ? 's' : ''} al carrito`;
  }

  async function confirmarPickup() {
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });

    const aceptados = _matchResults.filter(r => r.accepted && (r.productoOverride || r.producto));
    const noEncontrados = _matchResults.filter(r => !r.accepted || r.status === 'nofound');

    aceptados.forEach(r => {
      const prod = r.productoOverride || r.producto;
      const cod  = prod.idProducto;
      const st   = parseFloat((stockMap[cod] || {}).cantidadDisponible || 0);
      const idx  = _cart.findIndex(c => c.codigoBarra === cod);
      const item = { codigoBarra: cod, descripcion: prod.descripcion, unidad: prod.unidad || '', cantidad: r.qty, stockDisp: st };
      if (idx >= 0) _cart[idx] = item; else _cart.push(item);
    });

    // Pre-select zone from pickup
    if (_pickupActivo?.idZona) {
      const zonas = OfflineManager.getZonasCache();
      if (zonas.find(z => z.idZona === _pickupActivo.idZona)) {
        _cart._preZona = _pickupActivo.idZona;
      }
    }

    // Build nota for not-found items
    if (noEncontrados.length) {
      const nota = 'Inexistencias: ' + noEncontrados.map(r => `${r.nombre} ×${r.qty}`).join(', ');
      _cart._nota = nota;
    }

    _saveCart();
    cerrarSheet('sheetDespPickup');
    _renderCart();
    _updateFooter();

    // Marcar pickup como EN_PROCESO
    if (_pickupActivo) {
      API.actualizarPickup({ idPickup: _pickupActivo.idPickup, estado: 'EN_PROCESO' }).catch(() => {});
      _pickupsPendientes = _pickupsPendientes.filter(p => p.idPickup !== _pickupActivo.idPickup);
      _pickupActivo = null;
    }
    badgeUpdate();

    const n = aceptados.length;
    toast(`📥 ${n} producto${n !== 1 ? 's' : ''} cargados del pedido`, 'ok', 3000);
  }

  return { cargar, pauseCamera, toggleCamera,
           incQty, decQty, blurQty, quitarItem,
           finalizar, confirmarDespacho, cancelar,
           badgeUpdate, startPoll,
           abrirPickupPendiente, elegirMatchParcial, confirmarPickup };
})();

// ════════════════════════════════════════════════
// PREINGRESOS VIEW
// ════════════════════════════════════════════════
const PreingresosView = (() => {
  let _filtroEstado      = '';
  let _busquedaQ         = '';
  let _tppBusq           = '';   // búsqueda del panel tablet
  let _tags              = { comp: null, compl: null };   // 'si' | 'no' | null
  let _fotosSeleccionadas = [];                           // [{ file, objectUrl }]
  // Edit modal state
  let _editItem          = null;
  let _tagsEdit          = { comp: null, compl: null };
  let _cargadoresEdit    = [];   // [{ id, nombre, carretas }]
  let _fotosEdit         = [];   // [{ url }] existing Drive URLs kept
  let _fotosNuevas       = [];   // [{ file, objectUrl }] new files to upload

  function silentRefresh() { cargar(_filtroEstado, true); }

  function _aplicarBusqueda(list) {
    if (!_busquedaQ) return list;
    const qL = _busquedaQ.toLowerCase();
    return list.filter(p => {
      const provNombre = _getProveedorNombre(p.idProveedor).toLowerCase();
      return (p.idProveedor  || '').toLowerCase().includes(qL) ||
             provNombre.includes(qL);
    });
  }

  function buscar(q) {
    _busquedaQ = (q || '').trim();
    const cl = document.getElementById('clearBuscarPre');
    if (cl) cl.style.display = _busquedaQ ? 'flex' : 'none';
    const cached = OfflineManager.getPreingresosCache();
    const f = _filtroEstado ? cached.filter(p => p.estado === _filtroEstado) : cached;
    const lista = _aplicarBusqueda(f);
    _renderPreingresos(lista);

    if (!_busquedaQ) return;
    const qL = _busquedaQ.toLowerCase();

    const exacto = lista.find(p =>
      (p.idProveedor || '').toLowerCase() === qL ||
      _getProveedorNombre(p.idProveedor).toLowerCase() === qL
    );

    requestAnimationFrame(() => {
      if (exacto) {
        const cardId = 'pre-' + (exacto.idPreingreso || '').replace(/[^a-zA-Z0-9_-]/g, '_');
        const card   = document.getElementById(cardId);
        if (card) {
          card.classList.add('card-exact-match');
          card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        lista.forEach(p => {
          if (p !== exacto) {
            const el = document.getElementById('pre-' + (p.idPreingreso || '').replace(/[^a-zA-Z0-9_-]/g, '_'));
            if (el) el.classList.add('card-dim');
          }
        });
      } else {
        lista.forEach(p => {
          const el = document.getElementById('pre-' + (p.idPreingreso || '').replace(/[^a-zA-Z0-9_-]/g, '_'));
          if (el) el.classList.add('card-hi');
        });
      }
    });
  }

  function buscarClear() {
    _busquedaQ = '';
    const inp = document.getElementById('inputBuscarPre');
    if (inp) inp.value = '';
    const cl = document.getElementById('clearBuscarPre');
    if (cl) cl.style.display = 'none';
    const cached = OfflineManager.getPreingresosCache();
    const f = _filtroEstado ? cached.filter(p => p.estado === _filtroEstado) : cached;
    _renderPreingresos(f);
  }

  // ── Proveedor nombre desde caché ─────────────────────────
  function _getProveedorNombre(idProveedor) {
    if (!idProveedor) return 'Sin proveedor';
    const prov = OfflineManager.getProveedoresCache().find(p => p.idProveedor === idProveedor);
    return prov ? (prov.nombre || idProveedor) : idProveedor;
  }

  // ── Un card individual (altura fija, grid 3 filas) ───────
  function _renderCard(p) {
    const tieneGuia   = !!(p.idGuia && String(p.idGuia).trim());
    const nFotos      = p.fotos ? String(p.fotos).split(',').filter(Boolean).length : 0;
    const tags        = _tagsFromComentario(p.comentario);
    const borderColor = tieneGuia ? '#22c55e' : '#f59e0b';
    const provNombre  = _getProveedorNombre(p.idProveedor);
    const hora        = _horaDesdeId(p.idPreingreso);
    const fechaCorta  = _fmtCorta(p.fecha);

    // Tags top-right (compactos)
    let nCargadores = 0;
    try { const c = JSON.parse(p.cargadores || '[]'); nCargadores = Array.isArray(c) ? c.length : 0; } catch {}
    const tagHtml = [
      tags.compl === 'si' ? '<span class="pre-qtag pre-qtag-green">Completo</span>'   : '',
      tags.compl === 'no' ? '<span class="pre-qtag pre-qtag-amber">Incompleto</span>' : '',
      tags.comp  === 'si' ? '<span class="pre-qtag pre-qtag-blue">Comprobante</span>' : '',
      nFotos > 0          ? `<span class="pre-qtag pre-qtag-slate">📷${nFotos}</span>` : '',
      nCargadores > 0     ? `<span class="pre-qtag" style="background:#451a03;color:#fbbf24">🛺${nCargadores}</span>` : '',
    ].filter(Boolean).join('');

    // Bottom-right: crear guía OR guía icon
    const actionHtml = tieneGuia
      ? `<svg width="14" height="14" viewBox="0 0 16 16" fill="#22c55e" title="${escAttr(p.idGuia)}">
           <path d="M14 4.5V14a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2h5.5L14 4.5zm-3 0A1.5 1.5 0 0 1 9.5 3V1H4a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V4.5h-2z"/>
         </svg>`
      : `<button onclick="event.stopPropagation();PreingresosView.crearGuiaRapido('${escAttr(p.idPreingreso)}')"
                 class="pre-guia-btn">+ Guía</button>`;

    const waPreBtn = `<button onclick="event.stopPropagation();PreingresosView.compartirWA('${escAttr(p.idPreingreso)}')"
      style="background:none;border:none;cursor:pointer;padding:3px 2px;color:#25d366;flex-shrink:0;line-height:0"
      title="Compartir por WhatsApp">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
    </button>`;

    return `
    <div class="pre-card" id="pre-${(p.idPreingreso||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
         style="border-left-color:${borderColor}"
         onclick="PreingresosView.abrirDetalle('${escAttr(p.idPreingreso)}')">
      <div class="flex items-center justify-between gap-1 overflow-hidden">
        <span class="text-sm font-bold text-slate-100 truncate">${provNombre}</span>
        <div class="flex items-center gap-1 flex-shrink-0">${tagHtml}</div>
      </div>
      <p class="text-xs text-slate-400">${fechaCorta}${hora ? ' · ' + hora : ''}</p>
      <div class="flex items-center justify-between gap-1">
        <p class="text-sm font-bold ${p.monto ? 'text-emerald-400' : 'opacity-0'} leading-none">${p.monto ? 'S/. ' + fmt(p.monto, 2) : '—'}</p>
        <div class="flex items-center gap-1">${waPreBtn}${actionHtml}</div>
      </div>
    </div>`;
  }

  function _renderPreingresos(list) {
    const container = document.getElementById('listPreingresos');
    if (!container) return;
    const optCards = Array.from(container.querySelectorAll('.card-optimistic'));
    if (!list.length) {
      container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin preingresos</p>';
      optCards.forEach(c => container.insertBefore(c, container.firstChild));
      return;
    }

    // Ordenar descendente: por fecha local, luego por timestamp en ID
    const sorted = [...list].sort((a, b) => {
      const da = _parseLocalDate(a.fecha), db = _parseLocalDate(b.fecha);
      const td = db - da;
      if (td !== 0) return td;
      const na = parseInt((a.idPreingreso || '').replace(/\D/g, '')) || 0;
      const nb = parseInt((b.idPreingreso || '').replace(/\D/g, '')) || 0;
      return nb - na;
    });
    const today     = new Date(); today.setHours(0,0,0,0);
    const yesterday = new Date(today.getTime()); yesterday.setDate(today.getDate() - 1);
    const months    = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    function _dateKey(p) {
      if (!p.fecha) return '0000-00-00';
      const d = _parseLocalDate(p.fecha);
      if (isNaN(d)) return '0000-00-00';
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }
    function _dateLabel(key) {
      if (!key || key === '0000-00-00') return 'Sin fecha';
      const d = new Date(key + 'T12:00:00');
      const dMid = new Date(d); dMid.setHours(0,0,0,0);
      if (dMid.getTime() === today.getTime())     return 'Hoy';
      if (dMid.getTime() === yesterday.getTime()) return 'Ayer';
      return d.getFullYear() === today.getFullYear()
        ? `${d.getDate()} ${months[d.getMonth()]}`
        : `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    }

    const groupMap = {};
    sorted.forEach(p => {
      const k = _dateKey(p);
      if (!groupMap[k]) groupMap[k] = [];
      groupMap[k].push(p);
    });

    container.innerHTML = Object.entries(groupMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([key, items]) =>
        `<div class="pre-date-hdr">${_dateLabel(key)}</div>
         <div class="pre-date-group">${items.map(_renderCard).join('')}</div>`
      ).join('');

    // Preservar solo cards optimistas cuyo ID aún no está en la lista real
    optCards.forEach(c => {
      const rid = c.getAttribute('data-real-id') || c.id.replace('optcard_', '');
      if (!sorted.find(p => p.idPreingreso === rid)) {
        container.insertBefore(c, container.firstChild);
      }
    });

    // Sincronizar panel tablet (columna derecha de Guías)
    renderTppList();
  }

  async function cargar(estado = '', silencioso = false) {
    _filtroEstado = estado;
    // Mostrar desde caché primero (instantáneo)
    const cached = OfflineManager.getPreingresosCache();
    const filtrados = estado ? cached.filter(p => p.estado === estado) : cached;
    if (filtrados.length) {
      _renderPreingresos(_aplicarBusqueda(filtrados));
    } else if (!silencioso) {
      loading('listPreingresos', true);
    }
    // Refrescar vía endpoint compartido (throttled, no dispara si recién se llamó)
    await OfflineManager.precargarOperacional().catch(() => {});
    const fresh = OfflineManager.getPreingresosCache();
    const freshFiltrados = estado ? fresh.filter(p => p.estado === estado) : fresh;
    if (freshFiltrados.length) _renderPreingresos(_aplicarBusqueda(freshFiltrados));
  }

  function toggleFiltro() {
    const menu = document.getElementById('preFilterMenu');
    if (!menu) return;
    menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
    if (menu.style.display !== 'none') {
      setTimeout(() => { document.addEventListener('click', _closeFiltroOutside, { once: true }); }, 0);
    }
  }

  function _closeFiltroOutside(e) {
    if (!e.target.closest('#preFilterMenu') && !e.target.closest('#preFilterBtn')) {
      const menu = document.getElementById('preFilterMenu');
      if (menu) menu.style.display = 'none';
    }
  }

  function filtrar(estado) {
    // Actualizar UI del dropdown
    const menu = document.getElementById('preFilterMenu');
    if (menu) menu.style.display = 'none';
    const label = document.getElementById('preFilterLabel');
    if (label) {
      const map = { '': 'TODOS', 'PENDIENTE': 'SIN GUÍA', 'PROCESADO': 'CON GUÍA' };
      label.textContent = map[estado] || 'TODOS';
    }
    document.querySelectorAll('.pre-fopt').forEach(b => {
      b.classList.toggle('sel', b.dataset.pfiltro === estado);
    });
    cargar(estado);
  }

  // ── Panel de preingresos (accesible desde Guías) ────────
  let _panelFiltro = '';

  function abrirPanel() {
    _panelFiltro = '';
    ['preFiltAll','preFiltPend','preFiltProc'].forEach(id =>
      document.getElementById(id)?.classList.remove('active-tab'));
    document.getElementById('preFiltAll')?.classList.add('active-tab');
    _renderPanel('');
    abrirSheet('sheetPreingresosPanel');
  }

  function filtrarPanel(estado) {
    _panelFiltro = estado;
    ['preFiltAll','preFiltPend','preFiltProc'].forEach(id =>
      document.getElementById(id)?.classList.remove('active-tab'));
    const activeId = estado === 'PENDIENTE' ? 'preFiltPend' : estado === 'PROCESADO' ? 'preFiltProc' : 'preFiltAll';
    document.getElementById(activeId)?.classList.add('active-tab');
    _renderPanel(estado);
  }

  function _renderPanel(estado) {
    const cached = OfflineManager.getPreingresosCache();
    const list   = estado ? cached.filter(p => p.estado === estado) : cached;
    const container = document.getElementById('listPreingresosPanel');
    if (!container) return;
    const html = (items) => {
      if (!items.length) return '<p class="text-slate-500 text-sm text-center py-6">Sin preingresos</p>';
      return items.map(p => `
        <div class="card-sm">
          <div class="flex items-center justify-between mb-1">
            <span class="font-bold text-sm font-mono">${p.idPreingreso}</span>
            <span class="tag-${p.estado === 'PENDIENTE' ? 'warn' : p.estado === 'PROCESADO' ? 'ok' : 'blue'} text-xs">${p.estado}</span>
          </div>
          <p class="text-xs text-slate-400">${fmtFecha(p.fecha)} · ${p.idProveedor || '—'}</p>
          <p class="text-sm font-bold text-emerald-400 mt-1">S/. ${fmt(p.monto, 2)}</p>
          ${p.estado === 'PENDIENTE'
            ? `<button onclick="PreingresosView.aprobarDesdePanel('${p.idPreingreso}')"
                       class="btn btn-primary w-full mt-2 py-2 text-xs font-bold tracking-wide">
                 APROBAR → CREAR GUÍA
               </button>` : ''}
        </div>`).join('');
    };
    container.innerHTML = html(list);
    // Datos actualizados llegan por precargarOperacional (60s timer) → sin llamada extra aquí
  }

  async function aprobarDesdePanel(id) {
    const res = await API.aprobarPreingreso({ idPreingreso: id, usuario: window.WH_CONFIG.usuario });
    if (res.ok) {
      toast(`Guía ${res.data.idGuia} creada`, 'ok');
      filtrarPanel(_panelFiltro);
      GuiasView.cargar();
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  // ── Etiquetas toggle (formulario nuevo) ──────────────────
  function toggleTag(grupo, valor) {
    _tags[grupo] = (_tags[grupo] === valor) ? null : valor;
    const map = { comp: { si: 'tagComp1', no: 'tagComp0' }, compl: { si: 'tagCompl1', no: 'tagCompl0' } };
    ['si','no'].forEach(v => document.getElementById(map[grupo][v])?.classList.toggle('active', _tags[grupo] === v));
    // Monto solo visible cuando "con comprobante"
    document.getElementById('preMontoRow')?.classList.toggle('hidden', _tags.comp !== 'si');
  }

  // ── Etiquetas toggle (modal edición) ─────────────────────
  function toggleTagModal(grupo, valor) {
    _tagsEdit[grupo] = (_tagsEdit[grupo] === valor) ? null : valor;
    const map = { comp: { si: 'piTagComp1', no: 'piTagComp0' }, compl: { si: 'piTagCompl1', no: 'piTagCompl0' } };
    ['si','no'].forEach(v => document.getElementById(map[grupo][v])?.classList.toggle('active', _tagsEdit[grupo] === v));
    document.getElementById('piMontoRow')?.classList.toggle('hidden', _tagsEdit.comp !== 'si');
  }

  // ── Fotos seleccionadas ───────────────────────────────────
  function onFotosSeleccionadas(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    const MAX = 6;
    const restantes = MAX - _fotosSeleccionadas.length;
    if (restantes <= 0) { toast(`Máximo ${MAX} fotos por preingreso`, 'warn'); return; }
    files.slice(0, restantes).forEach(file => {
      _fotosSeleccionadas.push({ file, objectUrl: URL.createObjectURL(file) });
    });
    if (files.length > restantes) toast(`Solo se agregaron ${restantes} fotos (máximo ${MAX})`, 'warn');
    event.target.value = ''; // reset para poder seleccionar las mismas fotos otra vez
    _renderFotosPrev();
  }

  function quitarFoto(idx) {
    URL.revokeObjectURL(_fotosSeleccionadas[idx]?.objectUrl);
    _fotosSeleccionadas.splice(idx, 1);
    _renderFotosPrev();
  }

  function _renderFotosPrev() {
    const container = document.getElementById('preFotosPrev');
    const emptyMsg  = document.getElementById('preFotosEmpty');
    const countEl   = document.getElementById('preFotosCount');
    container.querySelectorAll('.foto-thumb').forEach(el => el.remove());
    if (!_fotosSeleccionadas.length) {
      emptyMsg.style.display = 'block';
      countEl.classList.add('hidden');
      return;
    }
    emptyMsg.style.display = 'none';
    countEl.classList.remove('hidden');
    countEl.textContent = `${_fotosSeleccionadas.length} foto${_fotosSeleccionadas.length !== 1 ? 's' : ''} seleccionada${_fotosSeleccionadas.length !== 1 ? 's' : ''}`;
    _fotosSeleccionadas.forEach((f, i) => {
      const div = document.createElement('div');
      div.className = 'foto-thumb';
      div.onclick = () => abrirCarrusel(_fotosSeleccionadas.map(x => x.objectUrl), 'Vista previa', i);
      div.innerHTML = `
        <img src="${f.objectUrl}" loading="lazy"/>
        <span class="foto-num">${i + 1}</span>
        <button class="foto-rm" onclick="event.stopPropagation();PreingresosView.quitarFoto(${i})">×</button>`;
      container.appendChild(div);
    });
  }

  function verFotos(fotosStr, titulo) {
    const fotos = (fotosStr || '').split(',').filter(Boolean);
    if (!fotos.length) { toast('Sin fotos registradas', 'info'); return; }
    abrirCarrusel(fotos, titulo || '');
  }

  // ── Abrir detalle / edición ───────────────────────────────
  function abrirDetalle(idPreingreso) {
    const cached = OfflineManager.getPreingresosCache();
    const p = cached.find(x => x.idPreingreso === idPreingreso);
    if (p) {
      _editItem = { ...p };
      _renderModal(p);
      abrirSheet('sheetDetallePI');
    } else {
      // Aún no está en caché (recién creado) → buscar en GAS
      toast('Cargando...', 'info', 1500);
      API.getPreingresos({ idPreingreso }).then(res => {
        const item = res.ok && res.data?.find ? res.data.find(x => x.idPreingreso === idPreingreso) : null;
        if (item) {
          _editItem = { ...item };
          _renderModal(item);
          abrirSheet('sheetDetallePI');
        } else {
          toast('Preingreso no encontrado', 'warn');
        }
      }).catch(() => toast('Sin conexión', 'warn'));
    }
  }

  function _renderModal(p) {
    // Proveedor dropdown
    const provSel = document.getElementById('piEditProv');
    if (provSel) {
      const provs = OfflineManager.getProveedoresCache();
      provSel.innerHTML = '<option value="">— Seleccionar —</option>' +
        provs.map(pv => `<option value="${escAttr(pv.idProveedor)}"${pv.idProveedor === p.idProveedor ? ' selected' : ''}>${escAttr(pv.nombre || pv.idProveedor)}</option>`).join('');
    }
    // Header
    const idEl = document.getElementById('piDetId');
    const fEl  = document.getElementById('piDetFecha');
    const eEl  = document.getElementById('piDetEstado');
    if (idEl) idEl.textContent = p.idPreingreso;
    if (fEl)  fEl.textContent  = fmtFecha(p.fecha);
    if (eEl)  { eEl.textContent = p.estado || 'PENDIENTE'; eEl.className = `tag-${p.estado === 'PENDIENTE' ? 'warn' : p.estado === 'PROCESADO' ? 'ok' : 'blue'} text-xs`; }

    // Tags
    _tagsEdit = _tagsFromComentario(p.comentario);
    const mapM = { comp: { si: 'piTagComp1', no: 'piTagComp0' }, compl: { si: 'piTagCompl1', no: 'piTagCompl0' } };
    ['comp','compl'].forEach(g => ['si','no'].forEach(v => document.getElementById(mapM[g][v])?.classList.toggle('active', _tagsEdit[g] === v)));
    document.getElementById('piMontoRow')?.classList.toggle('hidden', _tagsEdit.comp !== 'si');

    // Monto
    const montoInp = document.getElementById('piEditMonto');
    if (montoInp) montoInp.value = p.monto || '';

    // Comentario libre
    const comEl = document.getElementById('piEditComentario');
    if (comEl) comEl.value = _textoLibreFromComentario(p.comentario);

    // Cargadores
    try { _cargadoresEdit = JSON.parse(p.cargadores || '[]'); } catch { _cargadoresEdit = []; }
    if (!Array.isArray(_cargadoresEdit)) _cargadoresEdit = [];
    _renderCargadoresEdit();

    // Fotos
    _fotosEdit   = (p.fotos || '').split(',').filter(Boolean).map(url => ({ url }));
    _fotosNuevas.forEach(f => URL.revokeObjectURL(f.objectUrl));
    _fotosNuevas = [];
    _renderFotosEdit();

    // Botón guía
    document.getElementById('btnCrearGuiaPI')?.classList.toggle('hidden', p.estado !== 'PENDIENTE');
    const btnG = document.getElementById('btnCrearGuiaPI');
    if (btnG) { btnG.disabled = false; btnG.textContent = 'Crear Guía de Ingreso'; }
    const btnS = document.getElementById('btnGuardarPI');
    if (btnS) { btnS.disabled = false; btnS.textContent = 'Guardar cambios'; }
  }

  // ── Cargadores edit modal ────────────────────────────────
  function _renderCargadoresEdit() {
    const list = document.getElementById('piCargadoresList');
    if (!list) return;
    if (!_cargadoresEdit.length) {
      list.innerHTML = '<p class="text-xs text-slate-600 italic">Sin cargadores asignados</p>';
      return;
    }
    list.innerHTML = _cargadoresEdit.map((c, i) => `
      <div class="flex items-center gap-2 px-2 py-1.5 rounded-lg" style="background:#1a1505;border:1px solid #854d0e">
        <span class="text-amber-400 text-xs">🛺</span>
        <span class="text-amber-200 text-xs flex-1">${escAttr(c.nombre)}</span>
        <div class="flex items-center gap-1">
          <button onclick="PreingresosView.cambiarCarretasEdit(${i},-1)"
                  class="w-5 h-5 rounded text-center text-slate-300 hover:text-white text-sm leading-none"
                  style="background:#334155">−</button>
          <span class="text-amber-300 text-xs font-bold min-w-[1.2rem] text-center">${c.carretas}</span>
          <button onclick="PreingresosView.cambiarCarretasEdit(${i},1)"
                  class="w-5 h-5 rounded text-center text-slate-300 hover:text-white text-sm leading-none"
                  style="background:#334155">+</button>
        </div>
        <button onclick="PreingresosView.quitarCargadorEdit(${i})"
                class="text-slate-500 hover:text-red-400 text-sm leading-none ml-1">×</button>
      </div>`).join('');
  }

  async function _autoguardarCargadores() {
    if (!_editItem) return;
    const cargadores = JSON.stringify(_cargadoresEdit);
    _editItem.cargadores = cargadores;
    await API.actualizarPreingreso({ idPreingreso: _editItem.idPreingreso, cargadores })
      .catch(() => {});
    // Actualizar en caché local
    const cache = OfflineManager.getPreingresosCache();
    const idx = cache.findIndex(x => x.idPreingreso === _editItem.idPreingreso);
    if (idx >= 0) { cache[idx].cargadores = cargadores; OfflineManager.inyectarPreingreso && null; }
  }

  function cambiarCarretasEdit(idx, delta) {
    if (!_cargadoresEdit[idx]) return;
    _cargadoresEdit[idx].carretas = Math.max(1, _cargadoresEdit[idx].carretas + delta);
    _renderCargadoresEdit();
    _autoguardarCargadores();
  }

  function quitarCargadorEdit(idx) {
    _cargadoresEdit.splice(idx, 1);
    _renderCargadoresEdit();
    _autoguardarCargadores();
  }

  function abrirPickerCargadorEdit() {
    const todos = OfflineManager.getProveedoresCache()
      .filter(p => (p.nombre || '').toLowerCase().startsWith('cargador'));
    if (!todos.length) { toast('No hay cargadores registrados', 'warn'); return; }
    const yaIds = _cargadoresEdit.map(c => c.id);
    const disponibles = todos.filter(p => !yaIds.includes(p.idProveedor));
    if (!disponibles.length) { toast('Ya están todos los cargadores', 'info'); return; }
    const existing = document.getElementById('sheetCargadoresEdit');
    if (existing) existing.remove();
    const sheet = document.createElement('div');
    sheet.id = 'sheetCargadoresEdit';
    sheet.style.cssText = 'position:fixed;bottom:0;left:0;right:0;z-index:9999;padding:1.25rem;background:#0f172a;border-top:1px solid #1e293b;border-radius:1rem 1rem 0 0;max-height:55vh;overflow-y:auto';
    sheet.innerHTML = `
      <div class="w-10 h-1 bg-slate-600 rounded-full mx-auto mb-4"></div>
      <p class="font-bold text-sm mb-3 text-amber-300">🛺 Agregar Cargador</p>
      <div class="space-y-2">
        ${disponibles.map(c => `
          <button onclick="PreingresosView.agregarCargadorEdit('${c.idProveedor}','${escAttr(c.nombre)}')"
                  class="w-full text-left px-3 py-2.5 rounded-lg text-sm text-slate-200"
                  style="background:#1e293b;border:1px solid #334155">
            ${c.nombre}
          </button>`).join('')}
      </div>
      <button onclick="document.getElementById('sheetCargadoresEdit').remove()"
              class="mt-4 w-full text-xs text-slate-500 py-2">Cancelar</button>`;
    document.body.appendChild(sheet);
  }

  function agregarCargadorEdit(id, nombre) {
    if (_cargadoresEdit.find(c => c.id === id)) return;
    _cargadoresEdit.push({ id, nombre, carretas: 1 });
    _renderCargadoresEdit();
    _autoguardarCargadores();
    document.getElementById('sheetCargadoresEdit')?.remove();
  }

  // ── Fotos edit modal ──────────────────────────────────────
  function _renderFotosEdit() {
    const container = document.getElementById('piEditFotosPrev');
    const emptyMsg  = document.getElementById('piEditFotosEmpty');
    if (!container) return;
    container.querySelectorAll('.foto-thumb').forEach(el => el.remove());
    const allUrls = [..._fotosEdit.map(f => _normalizeDriveUrl(f.url)), ..._fotosNuevas.map(f => f.objectUrl)];
    emptyMsg.style.display = allUrls.length ? 'none' : 'block';
    _fotosEdit.forEach((f, i) => {
      const div = document.createElement('div');
      div.className = 'foto-thumb';
      div.onclick = () => abrirCarrusel(allUrls, 'Fotos', i);
      div.innerHTML = `<img src="${_normalizeDriveUrl(f.url)}" loading="lazy"/>
        <span class="foto-num">${i + 1}</span>
        <button class="foto-rm" onclick="event.stopPropagation();PreingresosView.quitarFotoEdit('exist',${i})">×</button>`;
      container.appendChild(div);
    });
    _fotosNuevas.forEach((f, i) => {
      const div = document.createElement('div');
      div.className = 'foto-thumb';
      div.onclick = () => abrirCarrusel(allUrls, 'Fotos', _fotosEdit.length + i);
      div.innerHTML = `<img src="${f.objectUrl}" loading="lazy"/>
        <span class="foto-num">${_fotosEdit.length + i + 1}</span>
        <button class="foto-rm" onclick="event.stopPropagation();PreingresosView.quitarFotoEdit('new',${i})">×</button>`;
      container.appendChild(div);
    });
  }

  function onFotosEditSeleccionadas(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    const MAX = 6;
    const restantes = MAX - _fotosEdit.length - _fotosNuevas.length;
    if (restantes <= 0) { toast(`Máximo ${MAX} fotos`, 'warn'); return; }
    files.slice(0, restantes).forEach(file => _fotosNuevas.push({ file, objectUrl: URL.createObjectURL(file) }));
    if (files.length > restantes) toast(`Solo se agregaron ${restantes} fotos (máx. ${MAX})`, 'warn');
    event.target.value = '';
    _renderFotosEdit();
  }

  function quitarFotoEdit(tipo, idx) {
    if (tipo === 'exist') {
      const url = _fotosEdit[idx]?.url || '';
      const match = url.match(/[?&]id=([^&]+)/);
      if (match) API.eliminarFotoDrive({ fileId: match[1] }).catch(() => {});
      _fotosEdit.splice(idx, 1);
    } else {
      URL.revokeObjectURL(_fotosNuevas[idx]?.objectUrl);
      _fotosNuevas.splice(idx, 1);
    }
    _renderFotosEdit();
  }

  // ── Guardar edición (optimista) ──────────────────────────
  async function guardarEdicion() {
    if (!_editItem) return;

    // Capturar datos y estado antes de cerrar
    const idPreingreso  = _editItem.idPreingreso;
    const idProveedor   = document.getElementById('piEditProv').value;
    const textoExtra    = (document.getElementById('piEditComentario').value || '').trim();
    const partes = [];
    if (_tagsEdit.comp)  partes.push(`Comprobante: ${_tagsEdit.comp === 'si' ? 'Sí' : 'No'}`);
    if (_tagsEdit.compl) partes.push(`Completo: ${_tagsEdit.compl === 'si' ? 'Sí' : 'No'}`);
    if (textoExtra) partes.push(textoExtra);
    const comentario    = partes.join(' | ');
    const monto         = _tagsEdit.comp === 'si' ? (parseFloat(document.getElementById('piEditMonto').value) || 0) : 0;
    const fotosExistentes = [..._fotosEdit];
    const fotosNuevasCaptura = [..._fotosNuevas];

    // Cerrar sheet y mostrar toast inmediatamente — sin esperar red
    cerrarSheet('sheetDetallePI');
    toast('Preingreso actualizado', 'ok');
    cargar(_filtroEstado, true);

    // Limpiar estado de edición
    _fotosNuevas = [];
    _fotosEdit   = [];

    // Subir fotos nuevas + actualizar en segundo plano
    (async () => {
      const todasFotos = [...fotosExistentes];
      for (let i = 0; i < fotosNuevasCaptura.length; i++) {
        try {
          const { b64, mime } = await _prepararFoto(fotosNuevasCaptura[i].file);
          const up = await API.subirFotoPreingreso({ idPreingreso, fotoBase64: b64, mimeType: mime, indice: fotosExistentes.length + i + 1 });
          if (up.ok && !up.offline && up.data?.url) todasFotos.push({ url: up.data.url });
          else console.warn('[FotosEdit] Error', i + 1, up.error || (up.offline ? 'sin conexión' : 'sin URL'));
        } catch(e) { console.warn('[FotosEdit]', e); }
      }
      fotosNuevasCaptura.forEach(f => f.objectUrl && URL.revokeObjectURL(f.objectUrl));
      const fotos = todasFotos.map(f => f.url).join(',');
      await API.actualizarPreingreso({ idPreingreso, idProveedor, monto, comentario, fotos, usuario: window.WH_CONFIG.usuario })
        .catch(e => console.warn('[EditPreingreso]', e));
    })();
  }

  // ── Crear Guía de Ingreso — optimista (modal) ────────────
  async function crearGuiaDesde() {
    if (!_editItem) return;
    const p = _editItem;
    const btn = document.getElementById('btnCrearGuiaPI');
    btn.disabled = true; btn.textContent = 'Creando…';
    cerrarSheet('sheetDetallePI');
    _lanzarCrearGuia(p.idPreingreso, p.idProveedor);
  }

  // ── Crear Guía de Ingreso — optimista (botón en card) ────
  async function crearGuiaRapido(idPreingreso) {
    const cached = OfflineManager.getPreingresosCache();
    const p = cached.find(x => x.idPreingreso === idPreingreso);
    if (!p) { toast('Preingreso no encontrado', 'warn'); return; }
    _lanzarCrearGuia(p.idPreingreso, p.idProveedor);
  }

  async function _lanzarCrearGuia(idPreingreso, idProveedor) {
    const tempId     = 'G_tmp_' + Date.now();
    const provNombre = _getProveedorNombre(idProveedor);
    GuiasView.injectOptimisticGuia({ tempId, idProveedor, provNombre });
    App.nav('guias');

    const res = await API.aprobarPreingreso({ idPreingreso, usuario: window.WH_CONFIG.usuario })
      .catch(() => ({ ok: false, error: 'Sin conexión' }));

    if (res.ok) {
      toast(`Guía ${res.data.idGuia} creada`, 'ok');
      GuiasView.finalizeOptimisticGuia(tempId, res.data.idGuia, 'INGRESO_PROVEEDOR', provNombre);
    } else {
      toast('Error al crear guía: ' + res.error, 'danger');
      GuiasView.removeOptimisticGuia(tempId);
    }
    cargar(_filtroEstado, true);
  }

  // Comprime imagen a max 1280px y quality 0.82 — devuelve {b64, mime}
  function _prepararFoto(file) {
    return new Promise((resolve, reject) => {
      const MAX = 1280;
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = ev => {
        const img = new Image();
        img.onerror = reject;
        img.onload = () => {
          let w = img.width, h = img.height;
          if (w > MAX || h > MAX) {
            if (w >= h) { h = Math.round(h * MAX / w); w = MAX; }
            else        { w = Math.round(w * MAX / h); h = MAX; }
          }
          const canvas = document.createElement('canvas');
          canvas.width = w; canvas.height = h;
          canvas.getContext('2d').drawImage(img, 0, 0, w, h);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.82);
          resolve({ b64: dataUrl.split(',')[1], mime: 'image/jpeg' });
        };
        img.src = ev.target.result;
      };
      reader.readAsDataURL(file);
    });
  }

  // ── Helpers tarjeta optimista ─────────────────────────────
  function _injectOptimisticCard(tempId, idProveedor, monto) {
    const container = document.getElementById('listPreingresos');
    if (!container) return;
    const div = document.createElement('div');
    div.id = 'optcard_' + tempId;
    div.className = 'card-sm ca ca-amber card-optimistic';
    div.innerHTML = `
      <div class="flex items-center justify-between mb-1">
        <span class="font-semibold text-sm text-slate-400 italic text-xs">Registrando…</span>
        <span class="tag-warn">PENDIENTE</span>
      </div>
      <p class="text-xs text-slate-400">${idProveedor}</p>
      ${monto ? `<p class="text-sm font-bold text-emerald-400 mt-1">S/. ${fmt(monto, 2)}</p>` : ''}
      <div class="flex items-center gap-2 mt-2">
        <div class="spinner" style="width:13px;height:13px;border-width:2px"></div>
        <span class="text-xs text-slate-500">Subiendo fotos…</span>
      </div>`;
    container.insertBefore(div, container.firstChild);
  }

  function _updateOptimisticId(tempId, realId) {
    const el = document.getElementById('optcard_' + tempId);
    if (el) el.id = 'optcard_' + realId;
  }

  function _finalizeOptimisticCard(realId, data = {}) {
    const el = document.getElementById('optcard_' + realId);
    if (el) {
      // Parar animación — mantener clase para que renders no destruyan el card
      el.style.animation = 'none';
      el.setAttribute('data-real-id', realId);
      const provNombre = _getProveedorNombre(data.idProveedor || '');
      el.innerHTML = `
        <div class="flex items-center justify-between gap-1 overflow-hidden">
          <span class="text-sm font-bold text-slate-100 truncate">${escAttr(provNombre)}</span>
          <span class="pre-qtag pre-qtag-amber">PENDIENTE</span>
        </div>
        <p class="text-xs text-slate-400">${fmtFecha(new Date())}</p>
        <div class="flex items-center justify-between gap-1 mt-1">
          <p class="text-sm font-bold text-emerald-400">S/. ${fmt(data.monto ?? 0, 2)}</p>
          <button onclick="event.stopPropagation();PreingresosView.crearGuiaRapido('${escAttr(realId)}')"
                  class="pre-guia-btn">+ Guía</button>
        </div>`;
    }
    // Refrescar via endpoint compartido (throttled) — actualizará la lista cuando llegue
    OfflineManager.precargarOperacional().then(() => {
      const fresh = OfflineManager.getPreingresosCache();
      const filtrados = _filtroEstado ? fresh.filter(p => p.estado === _filtroEstado) : fresh;
      if (filtrados.length) _renderPreingresos(_aplicarBusqueda(filtrados));
    }).catch(() => {});
  }

  // ── Búsqueda/filtrado de proveedores (excluye cargadores) ───
  function filtrarProveedores(q) {
    const drop = document.getElementById('preProvDrop');
    if (!drop) return;
    const provs = OfflineManager.getProveedoresCache()
      .filter(p => !(p.nombre || '').toLowerCase().startsWith('cargador'));
    const ql = (q || '').trim().toLowerCase();
    const matches = ql
      ? provs.filter(p => (p.nombre || '').toLowerCase().includes(ql) || (p.idProveedor || '').toLowerCase().includes(ql))
      : provs.slice(0, 12);
    if (!matches.length) { drop.classList.add('hidden'); return; }
    drop.innerHTML = matches.map(p =>
      `<div class="px-3 py-2 text-sm text-slate-200 hover:bg-slate-700 cursor-pointer"
            onclick="PreingresosView.seleccionarProveedor('${p.idProveedor}','${escAttr(p.nombre || p.idProveedor)}')"
       >${p.nombre || p.idProveedor}</div>`
    ).join('');
    drop.classList.remove('hidden');
  }

  function seleccionarProveedor(id, nombre) {
    document.getElementById('preProvSelect').value    = id;
    document.getElementById('preProvInput').value     = '';
    document.getElementById('preProvSelNombre').textContent = nombre;
    document.getElementById('preProvSelBox').classList.remove('hidden');
    document.getElementById('preProvDrop').classList.add('hidden');
  }

  function limpiarProveedor() {
    document.getElementById('preProvSelect').value = '';
    document.getElementById('preProvInput').value  = '';
    document.getElementById('preProvSelBox').classList.add('hidden');
  }

  // ── Cargadores: lista con contador de carretas ───────────
  let _cargadores = []; // [{ id, nombre, carretas }]

  function _renderCargadores() {
    const list = document.getElementById('preCargadoresList');
    if (!list) return;
    if (!_cargadores.length) { list.innerHTML = ''; return; }
    list.innerHTML = _cargadores.map((c, i) => `
      <div class="flex items-center gap-2 px-2 py-1.5 rounded-lg" style="background:#1a1505;border:1px solid #854d0e">
        <span class="text-amber-400 text-xs">🛺</span>
        <span class="text-amber-200 text-xs flex-1">${escAttr(c.nombre)}</span>
        <div class="flex items-center gap-1">
          <button onclick="PreingresosView.cambiarCarretas(${i},-1)"
                  class="w-5 h-5 rounded text-center text-slate-300 hover:text-white text-sm leading-none"
                  style="background:#334155">−</button>
          <span class="text-amber-300 text-xs font-bold min-w-[1.2rem] text-center">${c.carretas}</span>
          <button onclick="PreingresosView.cambiarCarretas(${i},1)"
                  class="w-5 h-5 rounded text-center text-slate-300 hover:text-white text-sm leading-none"
                  style="background:#334155">+</button>
        </div>
        <button onclick="PreingresosView.quitarCargador(${i})"
                class="text-slate-500 hover:text-red-400 text-sm leading-none ml-1">×</button>
      </div>`).join('');
  }

  function abrirPickerCargador() {
    const todos = OfflineManager.getProveedoresCache()
      .filter(p => (p.nombre || '').toLowerCase().startsWith('cargador'));
    if (!todos.length) { toast('No hay cargadores registrados', 'warn'); return; }
    // Excluir los ya agregados
    const yaIds = _cargadores.map(c => c.id);
    const disponibles = todos.filter(p => !yaIds.includes(p.idProveedor));
    if (!disponibles.length) { toast('Ya agregaste todos los cargadores disponibles', 'info'); return; }
    const existing = document.getElementById('sheetCargadores');
    if (existing) existing.remove();
    const sheet = document.createElement('div');
    sheet.id = 'sheetCargadores';
    sheet.style.cssText = 'position:fixed;bottom:0;left:0;right:0;z-index:9999;padding:1.25rem;background:#0f172a;border-top:1px solid #1e293b;border-radius:1rem 1rem 0 0;max-height:60vh;overflow-y:auto';
    sheet.innerHTML = `
      <div class="w-10 h-1 bg-slate-600 rounded-full mx-auto mb-4"></div>
      <p class="font-bold text-sm mb-3 text-amber-300">🛺 Agregar Cargador</p>
      <div class="space-y-2">
        ${disponibles.map(c => `
          <button onclick="PreingresosView.agregarCargador('${c.idProveedor}','${escAttr(c.nombre)}')"
                  class="w-full text-left px-3 py-2.5 rounded-lg text-sm text-slate-200 transition-colors"
                  style="background:#1e293b;border:1px solid #334155">
            ${c.nombre}
          </button>`).join('')}
      </div>
      <button onclick="document.getElementById('sheetCargadores').remove()"
              class="mt-4 w-full text-xs text-slate-500 py-2">Cancelar</button>`;
    document.body.appendChild(sheet);
  }

  function agregarCargador(id, nombre) {
    if (_cargadores.find(c => c.id === id)) return; // evitar duplicado
    _cargadores.push({ id, nombre, carretas: 1 });
    _renderCargadores();
    document.getElementById('sheetCargadores')?.remove();
  }

  function cambiarCarretas(idx, delta) {
    if (!_cargadores[idx]) return;
    _cargadores[idx].carretas = Math.max(1, _cargadores[idx].carretas + delta);
    _renderCargadores();
  }

  function quitarCargador(idx) {
    _cargadores.splice(idx, 1);
    _renderCargadores();
  }

  function limpiarCargador() {
    _cargadores = [];
    _renderCargadores();
  }

  // ── Crear preingreso (optimista) ─────────────────────────
  async function crear() {
    const idProveedor = document.getElementById('preProvSelect').value;
    if (!idProveedor) { toast('Selecciona un proveedor', 'warn'); return; }
    if (!_fotosSeleccionadas.length) { toast('Agrega al menos una foto', 'warn'); return; }

    // Armar comentario: etiquetas + texto libre
    const partes = [];
    if (_tags.comp)  partes.push(`Comprobante: ${_tags.comp === 'si' ? 'Sí' : 'No'}`);
    if (_tags.compl) partes.push(`Completo: ${_tags.compl === 'si' ? 'Sí' : 'No'}`);
    const textoExtra = (document.getElementById('preComentario').value || '').trim();
    if (textoExtra) partes.push(textoExtra);
    const comentario = partes.join(' | ');
    const monto = _tags.comp === 'si' ? (parseFloat(document.getElementById('preMonto').value) || 0) : 0;

    const btn = document.getElementById('btnCrearPre');
    btn.disabled = true; btn.textContent = 'Registrando...';

    // OPTIMISTIC: mostrar tarjeta inmediatamente
    const tempId = 'tmp_' + Date.now();
    _injectOptimisticCard(tempId, idProveedor, monto);
    cerrarSheet('sheetPreingreso');

    // 1. Crear preingreso con ID generado en cliente (evita duplicados por retry)
    const idPreingreso = 'PI' + Date.now();
    const cargadores   = _cargadores.length ? JSON.stringify(_cargadores) : '';
    const res = await API.crearPreingreso({ idPreingreso, idProveedor, cargadores, monto, comentario, usuario: window.WH_CONFIG.usuario })
      .catch(() => ({ ok: false, error: 'Sin conexión' }));

    if (!res.ok) {
      document.getElementById('optcard_' + tempId)?.remove();
      toast('Error: ' + res.error, 'danger');
      btn.disabled = false; btn.textContent = 'Registrar preingreso';
      return;
    }

    const idPreingresoReal = res.data?.idPreingreso || idPreingreso;
    _updateOptimisticId(tempId, idPreingresoReal);

    // 2. Finalizar card inmediatamente — el usuario ya puede ver y usar el preingreso
    _finalizeOptimisticCard(idPreingresoReal, { idProveedor, monto });
    toast(`Preingreso ${idPreingresoReal} registrado`, 'ok');

    // Inyectar en caché para que abrirDetalle no tenga que ir al GAS
    OfflineManager.inyectarPreingreso({
      idPreingreso: idPreingresoReal, idProveedor, cargadores, monto,
      comentario, estado: 'PENDIENTE',
      fecha: new Date().toISOString(), fotos: '',
      usuario: window.WH_CONFIG.usuario
    });
    btn.disabled = false; btn.textContent = 'Registrar preingreso';

    // 3. Subir fotos en segundo plano (no bloquea UI)
    const fotosCaptura = [..._fotosSeleccionadas];
    _fotosSeleccionadas = [];
    _subirFotosEnBackground(idPreingresoReal, fotosCaptura);
  }

  async function _subirFotosEnBackground(idPreingreso, fotos) {
    if (!fotos.length) return;
    const urls = [];
    for (let i = 0; i < fotos.length; i++) {
      try {
        const { b64, mime } = await _prepararFoto(fotos[i].file);
        const up = await API.subirFotoPreingreso({ idPreingreso, fotoBase64: b64, mimeType: mime, indice: i + 1 });
        if (up.ok && !up.offline && up.data?.url) urls.push(up.data.url);
        else console.warn('[Fotos] Error foto', i + 1, up.error || (up.offline ? 'sin conexión' : 'sin URL'));
      } catch(e) { console.warn('[Fotos]', i, e); }
    }
    if (urls.length) {
      await API.actualizarFotosPreingreso({ idPreingreso, fotos: urls.join(',') }).catch(() => {});
      fotos.forEach(f => f.objectUrl && URL.revokeObjectURL(f.objectUrl));
    } else {
      toast('Fotos no pudieron subirse — reintenta desde edición', 'warn', 5000);
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

  function nuevo() {
    // Reset completo del formulario
    _tags = { comp: null, compl: null };
    _fotosSeleccionadas.forEach(f => URL.revokeObjectURL(f.objectUrl));
    _fotosSeleccionadas = [];
    ['tagComp1','tagComp0','tagCompl1','tagCompl0'].forEach(id =>
      document.getElementById(id)?.classList.remove('active'));
    const montoInp = document.getElementById('preMonto');
    const com      = document.getElementById('preComentario');
    const fi       = document.getElementById('preFileInput');
    const btn      = document.getElementById('btnCrearPre');
    if (montoInp) montoInp.value = '';
    if (com)      com.value = '';
    if (fi)       fi.value  = '';
    if (btn)      { btn.disabled = false; btn.textContent = 'Registrar preingreso'; }
    document.getElementById('preMontoRow')?.classList.add('hidden');
    document.getElementById('preFotosEmpty').style.display = 'block';
    document.getElementById('preFotosPrev')?.querySelectorAll('.foto-thumb').forEach(el => el.remove());
    document.getElementById('preFotosCount')?.classList.add('hidden');
    // Reset proveedor/cargador
    limpiarProveedor();
    limpiarCargador();
    document.getElementById('sheetCargadores')?.remove();
    // Cerrar dropdown si quedó abierto
    document.getElementById('preProvDrop')?.classList.add('hidden');
    abrirSheet('sheetPreingreso');
  }

  // ── Panel tablet (columna derecha en vista Guías) ────────
  function renderTppList() {
    const container = document.getElementById('tppList');
    if (!container || !container.offsetParent) return;
    const todos = OfflineManager.getPreingresosCache();
    const list = _tppBusq
      ? todos.filter(p =>
          _getProveedorNombre(p.idProveedor).toLowerCase().includes(_tppBusq) ||
          (p.idPreingreso || '').toLowerCase().includes(_tppBusq) ||
          (p.idProveedor  || '').toLowerCase().includes(_tppBusq)
        )
      : todos;
    if (!list.length) {
      container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin preingresos</p>';
      return;
    }
    const sorted = [...list].sort((a, b) => {
      const da = _parseLocalDate(a.fecha), db = _parseLocalDate(b.fecha);
      const td = db - da;
      if (td !== 0) return td;
      const na = parseInt((a.idPreingreso || '').replace(/\D/g, '')) || 0;
      const nb = parseInt((b.idPreingreso || '').replace(/\D/g, '')) || 0;
      return nb - na;
    });
    container.innerHTML = sorted.slice(0, 40).map(_renderCard).join('');
  }

  function buscarEnPanel(q) {
    _tppBusq = (q || '').toLowerCase().trim();
    const btn = document.getElementById('clearTppSearch');
    if (btn) btn.style.display = _tppBusq ? 'flex' : 'none';
    renderTppList();
  }

  function limpiarBuscarPanel() {
    _tppBusq = '';
    const inp = document.getElementById('tppSearch');
    if (inp) inp.value = '';
    const btn = document.getElementById('clearTppSearch');
    if (btn) btn.style.display = 'none';
    renderTppList();
  }

  function compartirWA(idPreingreso) {
    const pi = (OfflineManager.getPreingresosCache() || []).find(x => x.idPreingreso === idPreingreso);
    if (!pi) return;
    const prov     = (OfflineManager.getProveedoresCache() || []).find(x => x.idProveedor === pi.idProveedor);
    const provNombre = prov?.nombre || pi.idProveedor || '—';
    let cargadores = [];
    try { cargadores = JSON.parse(pi.cargadores || '[]'); } catch {}
    const cargStr  = cargadores.length ? cargadores.map(c => c.nombre || c.id || String(c)).join(', ') : '—';
    const url      = `${location.href.split('?')[0].replace(/index\.html$/, '').replace(/\/$/, '')}/reporte.html?tipo=preingreso&id=${encodeURIComponent(idPreingreso)}`;
    const lineas   = [
      `*📦 PREINGRESO ${idPreingreso}*`,
      `─────────────────────`,
      `🏪 *Proveedor:* ${provNombre}`,
      `💰 *Monto:* ${pi.monto ? 'S/ ' + fmt(pi.monto, 2) : '—'}`,
      `📅 *Fecha:* ${_fmtCorta(pi.fecha)}`,
      `📊 *Estado:* ${pi.estado || '—'}`,
      `👥 *Cargadores:* ${cargStr}`,
    ];
    if (pi.comentario) lineas.push(`💬 *Comentario:* ${pi.comentario}`);
    if (pi.idGuia)     lineas.push(`📋 *Guía:* ${pi.idGuia}`);
    lineas.push(
      ``,
      `━━━━━━━━━━━━━━━━━━━━━━━━`,
      `👇 *TOCA AQUÍ PARA VER EL REPORTE COMPLETO*`,
      url,
      `━━━━━━━━━━━━━━━━━━━━━━━━`,
      `_InversionMos Warehouse_`
    );
    window.open('https://wa.me/?text=' + encodeURIComponent(lineas.join('\n')), '_blank');
  }

  return { cargar, filtrar, toggleFiltro, silentRefresh, buscar, buscarClear, crear, aprobar, nuevo,
           abrirPanel, filtrarPanel, aprobarDesdePanel,
           toggleTag, toggleTagModal,
           onFotosSeleccionadas, quitarFoto, verFotos,
           onFotosEditSeleccionadas, quitarFotoEdit,
           abrirDetalle, guardarEdicion, crearGuiaDesde, crearGuiaRapido,
           filtrarProveedores, seleccionarProveedor, limpiarProveedor,
           abrirPickerCargador, agregarCargador, cambiarCarretas, quitarCargador, limpiarCargador,
           abrirPickerCargadorEdit, agregarCargadorEdit, cambiarCarretasEdit, quitarCargadorEdit,
           renderTppList, buscarEnPanel, limpiarBuscarPanel, compartirWA };
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
// ════════════════════════════════════════════════
// PRODUCTOS VIEW — catálogo maestro MOS agrupado por SKU
// ════════════════════════════════════════════════
const ProductosView = (() => {
  'use strict';
  let _grupos       = [];
  let _filtrados    = [];
  let _stockMap     = {};
  let _histTarget   = null;  // { codigo, nombre }
  let _queryActual  = '';    // búsqueda activa (para sobrevivir bg-refresh)

  // ── Estado de ajuste manual ───────────────────────────────
  let _ajusteTarget = null; // { codigoBarra, nombre }

  // ── Estado de auditoría diaria ────────────────────────────
  const _AUDIT_KEY  = 'wh_audit_dia';
  let _auditDia     = null;  // { fecha, skus:[...30], auditados:{sku:[cods]} }
  let _auditModo    = false; // modo filtro activo
  let _auditTarget  = null;  // barcode actualmente en auditoría

  // ── helpers ────────────────────────────────────
  function _s(id)          { return _stockMap[id] || { cantidadDisponible: 0, stockMinimo: 0, stockMaximo: 0 }; }
  function _buildMap(list) { list.forEach(s => { _stockMap[s.codigoProducto || s.idProducto] = s; }); }

  // Rotación y último mov sobre un array de códigos (equivalencias del grupo)
  function _rotacionMulti(codigos) {
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const hace30   = Date.now() - 30 * 86400000;
    const gMap     = {};
    guias.forEach(g => { gMap[g.idGuia] = g; });
    const set = new Set(codigos);
    const n = detalles.filter(d => set.has(d.codigoProducto) && gMap[d.idGuia] && new Date(gMap[d.idGuia].fecha) >= hace30).length;
    if (n >= 10) return { nivel: 'ALTA',  color: 'text-emerald-400', dot: 'bg-emerald-400' };
    if (n >= 4)  return { nivel: 'MEDIA', color: 'text-amber-400',   dot: 'bg-amber-400'   };
    return             { nivel: 'BAJA',  color: 'text-slate-500',   dot: 'bg-slate-600'   };
  }

  function _ultimoMovMulti(codigos) {
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const gMap     = {};
    guias.forEach(g => { gMap[g.idGuia] = g; });
    const set    = new Set(codigos);
    const fechas = detalles.filter(d => set.has(d.codigoProducto)).map(d => gMap[d.idGuia]?.fecha).filter(Boolean);
    return fechas.sort((a, b) => new Date(b) - new Date(a))[0] || null;
  }

  // ── Agrupar por skuBase ─────────────────────────────────────
  // Hijos = todos los factor=1 de PRODUCTOS_MASTER con ese skuBase
  //       + todos los registros de EQUIVALENCIAS con ese skuBase
  // Stock de cada hijo: lookup por codigoBarra en _stockMap (WH STOCK)
  function _agrupar(prods, equivs) {
    // 1. Solo productos base del almacén: factorConversion === 1 (o vacío)
    // Las presentaciones POS tienen factorConversion ≠ 1 y no se manejan en WH
    const f1 = prods.filter(p => {
      if (p.estado === '0' || p.estado === 0) return false;
      return parseFloat(p.factorConversion || 1) === 1;
    });

    // 2. Agrupar f1 por skuBase
    const grp = {};
    f1.forEach(p => {
      const key = String(p.skuBase || p.idProducto || '').trim();
      if (!key) return;
      if (!grp[key]) grp[key] = { skuBase: key, prods: [] };
      grp[key].prods.push(p);
    });

    // 3. Agregar equivalencias a cada grupo
    equivs.forEach(e => {
      const key = String(e.skuBase || '').trim();
      if (!key || !grp[key]) return;
      grp[key].equivs = grp[key].equivs || [];
      grp[key].equivs.push(e);
    });

    // 4. Construir grupos finales
    return Object.values(grp).map(g => {
      // Header = producto con menor idProducto (el primero registrado)
      g.prods.sort((a, b) => Number(a.idProducto || 0) - Number(b.idProducto || 0));
      const base    = g.prods[0];
      const equivs2 = g.equivs || [];

      // Hijos unificados: primero los de PRODUCTOS_MASTER, luego los de EQUIVALENCIAS
      const children = [
        ...g.prods.map(p => ({
          codigoBarra: p.codigoBarra || p.idProducto,
          descripcion: p.descripcion,
          origen: 'prod'
        })),
        ...equivs2.map(e => ({
          codigoBarra: e.codigoBarra,
          descripcion: e.descripcion || e.codigoBarra,
          origen: 'equiv'
        }))
      ];

      // Stock total = suma de todos los hijos
      let stockTotal = 0, bajoMin = false;
      children.forEach(c => {
        const s  = _s(c.codigoBarra);
        const st = s.cantidadDisponible || 0;
        const mn = parseFloat(s.stockMinimo || base.stockMinimo || 0);
        stockTotal += st;
        if (mn > 0 && st <= mn) bajoMin = true;
      });

      return { skuBase: g.skuBase, base, children, stockTotal, bajoMin };
    }).sort((a, b) => String(a.base.descripcion || '').localeCompare(String(b.base.descripcion || ''), 'es'));
  }

  // ── Render lista de grupos ──────────────────────
  function _render(grupos) {
    const el = document.getElementById('listProductos');
    if (!grupos.length) { el.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin productos</p>'; return; }
    el.innerHTML = grupos.map(_cardGrupo).join('');
  }

  function _cardGrupo(g) {
    const codigos = g.children.map(c => c.codigoBarra).filter(Boolean);
    const rot  = _rotacionMulti(codigos);
    const ulti = _ultimoMovMulti(codigos);
    const mn   = parseFloat(g.base.stockMinimo || 0);
    const mx   = parseFloat(g.base.stockMaximo || 0);
    const pct  = mx > 0 ? Math.min(100, g.stockTotal / mx * 100) : 0;
    const barC = g.bajoMin ? 'bg-red-500' : pct < 40 ? 'bg-amber-500' : 'bg-emerald-500';
    const accentCls = g.bajoMin ? 'ca-red'
                    : g.stockTotal === 0 ? 'ca-slate'
                    : (pct < 40 && mx > 0) ? 'ca-amber' : 'ca-green';
    // Un solo barcode → no hay nada que desplegar
    const hasChildren = g.children.length > 1;
    const safe = escAttr(g.base.descripcion);
    const sid  = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');

    // Auditoría
    const isAudit     = _auditModo && (_auditDia?.skus.includes(g.skuBase));
    const isAuditDone = isAudit && _esGrupoCompleto(g.skuBase);
    const auditCls    = isAuditDone ? 'audit-card-done' : isAudit ? 'audit-card' : '';

    return `
    <div class="prod-card ca ${accentCls} ${auditCls}" id="grp-${sid}">
      <!-- Cabecera -->
      <div class="flex items-start gap-2">
        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-2 flex-wrap">
            <p class="font-bold text-sm leading-snug">${g.base.descripcion || g.skuBase}</p>
            ${g.bajoMin ? '<span class="tag-danger text-xs flex-shrink-0">⚠️ MÍN</span>' : ''}
            ${g.stockTotal === 0 ? '<span class="tag-danger text-xs flex-shrink-0">SIN STOCK</span>' : ''}
          </div>
          ${hasChildren
            ? `<button onclick="event.stopPropagation();ProductosView.toggleGrupo('${sid}')"
                       class="flex items-center gap-1 mt-0.5 text-xs font-mono text-slate-400 hover:text-slate-200 transition-colors">
                <svg class="transition-transform flex-shrink-0" id="chev-${sid}" width="11" height="11" viewBox="0 0 16 16" fill="currentColor">
                  <path d="M1.646 4.646a.5.5 0 0 1 .708 0L8 10.293l5.646-5.647a.5.5 0 0 1 .708.708l-6 6a.5.5 0 0 1-.708 0l-6-6a.5.5 0 0 1 0-.708z"/>
                </svg>
                ${g.children.length} barcodes
              </button>`
            : `<p class="text-xs text-slate-500 font-mono mt-0.5">${g.children[0]?.codigoBarra || g.skuBase}</p>`
          }
        </div>
        <span class="font-black text-base flex-shrink-0 pt-0.5 ${g.bajoMin ? 'text-red-400' : g.stockTotal === 0 ? 'text-slate-500' : 'text-emerald-400'}">${fmt(g.stockTotal)}</span>
      </div>

      <!-- Barra de nivel de stock -->
      ${mx > 0 ? `
        <div class="bar-bg mt-2 mb-1"><div class="bar-fill ${barC}" style="width:${pct.toFixed(0)}%"></div></div>
        <div class="flex justify-between text-xs text-slate-600 mb-1">
          <span>Mín: ${fmt(mn)}</span><span>Máx: ${fmt(mx)}</span>
        </div>` : '<div class="mt-1.5"></div>'}

      <!-- Métricas almacenero -->
      <div class="flex items-center gap-3 text-xs mt-0.5">
        <span class="flex items-center gap-1">
          <span class="w-2 h-2 rounded-full flex-shrink-0 ${rot.dot}"></span>
          <span class="${rot.color} font-semibold">ROT. ${rot.nivel}</span>
        </span>
        <span class="text-slate-500">${ulti ? 'Últ: ' + fmtFecha(ulti) : 'Sin movs.'}</span>
      </div>

      <!-- Acciones -->
      <div class="flex gap-2 mt-2 pt-2 border-t flex-wrap" style="border-color:#334155">
        <button onclick="event.stopPropagation();ProductosView.verHistorial('${escAttr(g.children.map(c=>c.codigoBarra).join('|'))}','${safe}')"
                class="btn btn-outline text-xs py-1 px-2 flex items-center gap-1">
          <svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor">
            <path d="M8 3.5a.5.5 0 0 0-1 0V9a.5.5 0 0 0 .252.434l3.5 2a.5.5 0 0 0 .496-.868L8 8.71V3.5z"/>
            <path d="M8 16A8 8 0 1 0 8 0a8 8 0 0 0 0 16zm7-8A7 7 0 1 1 1 8a7 7 0 0 1 14 0z"/>
          </svg>
          Historial
        </button>
        ${/* Botón ojo: solo barcode único en modo auditoría */
          isAudit && !hasChildren ? (() => {
            const c0   = g.children[0];
            const cod0 = escAttr(String(c0?.codigoBarra || ''));
            const nom0 = escAttr(c0?.descripcion || g.base.descripcion || '');
            const done = _esBarcodeAuditado(g.skuBase, c0?.codigoBarra);
            return `<button onclick="event.stopPropagation();${done ? '' : `ProductosView.abrirAuditBarcode('${cod0}','${nom0}','${escAttr(g.skuBase)}')`}"
              class="btn-eye${done ? ' done' : ''} ml-auto">
              <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor">
                ${done
                  ? '<path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/>'
                  : '<path d="M16 8s-3-5.5-8-5.5S0 8 0 8s3 5.5 8 5.5S16 8 16 8zM1.173 8a13.133 13.133 0 0 1 1.66-2.043C4.12 4.668 5.88 3.5 8 3.5c2.12 0 3.879 1.168 5.168 2.457A13.133 13.133 0 0 1 14.828 8c-.058.087-.122.183-.195.288-.335.48-.83 1.12-1.465 1.755C11.879 11.332 10.119 12.5 8 12.5c-2.12 0-3.879-1.168-5.168-2.457A13.134 13.134 0 0 1 1.172 8z"/><path d="M8 5.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5zM4.5 8a3.5 3.5 0 1 1 7 0 3.5 3.5 0 0 1-7 0z"/>'}
              </svg>
              ${done ? 'OK' : 'Auditar'}
            </button>`;
          })() : ''}
      </div>

      <!-- Panel hijos (colapsado, solo si hay más de uno) -->
      ${hasChildren ? `
      <div id="eqs-${sid}" class="hidden mt-3 space-y-2 border-t pt-3" style="border-color:#334155">
        ${g.children.map(c => _rowChild(c, g.base, g.skuBase)).join('')}
      </div>` : ''}
    </div>`;
  }

  // c = { codigoBarra, descripcion, origen: 'prod'|'equiv' }
  // base = producto master, skuBase del grupo padre
  function _rowChild(c, base, skuBase) {
    const s   = _s(c.codigoBarra);
    const st  = s.cantidadDisponible || 0;
    const mn  = parseFloat(s.stockMinimo  || base.stockMinimo  || 0);
    const mx  = parseFloat(s.stockMaximo  || base.stockMaximo  || 0);
    const baj = mn > 0 && st <= mn;
    const pct = mx > 0 ? Math.min(100, st / mx * 100) : 0;
    const tagOrigen = c.origen === 'equiv'
      ? '<span class="tag-blue text-xs flex-shrink-0" style="font-size:9px">EQUIV</span>'
      : '';
    // Botón ojo de auditoría
    const sku    = skuBase || base.skuBase || base.idProducto || '';
    const isDone = _esBarcodeAuditado(sku, c.codigoBarra);
    const eyeBtn = _auditModo ? `
      <button onclick="event.stopPropagation();${isDone ? '' : `ProductosView.abrirAuditBarcode('${escAttr(String(c.codigoBarra))}','${escAttr(c.descripcion)}','${escAttr(sku)}')`}"
              class="btn-eye${isDone ? ' done' : ''}">
        <svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor">
          ${isDone
            ? '<path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/>'
            : '<path d="M16 8s-3-5.5-8-5.5S0 8 0 8s3 5.5 8 5.5S16 8 16 8zM1.173 8a13.133 13.133 0 0 1 1.66-2.043C4.12 4.668 5.88 3.5 8 3.5c2.12 0 3.879 1.168 5.168 2.457A13.133 13.133 0 0 1 14.828 8c-.058.087-.122.183-.195.288-.335.48-.83 1.12-1.465 1.755C11.879 11.332 10.119 12.5 8 12.5c-2.12 0-3.879-1.168-5.168-2.457A13.134 13.134 0 0 1 1.172 8z"/><path d="M8 5.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5zM4.5 8a3.5 3.5 0 1 1 7 0 3.5 3.5 0 0 1-7 0z"/>'}
        </svg>
        ${isDone ? 'OK' : 'Auditar'}
      </button>` : '';

    return `
    <div class="rounded-lg p-2.5" style="background:#0f172a">
      <div class="flex items-center justify-between gap-2">
        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-1.5 flex-wrap">
            <p class="text-sm font-semibold truncate">${c.descripcion}</p>
            ${tagOrigen}
          </div>
          <p class="text-xs text-slate-500 font-mono mt-0.5">${c.codigoBarra}</p>
        </div>
        <div class="text-right flex-shrink-0">
          <span class="font-black text-base ${baj ? 'text-red-400' : st === 0 ? 'text-slate-500' : 'text-emerald-400'}">${fmt(st)}</span>
          ${baj ? '<p class="text-xs text-red-500 mt-0.5">⚠️ mín.</p>' : ''}
        </div>
      </div>
      ${mx > 0 ? `
        <div class="bar-bg mt-1.5 mb-1"><div class="bar-fill ${baj ? 'bg-red-500' : pct < 40 ? 'bg-amber-500' : 'bg-emerald-500'}" style="width:${pct.toFixed(0)}%"></div></div>
        <p class="text-xs text-slate-600">Mín: ${fmt(mn)} · Máx: ${fmt(mx)}</p>` : ''}
      <div class="flex gap-2 mt-2 flex-wrap">
        <button onclick="ProductosView.verHistorial('${escAttr(c.codigoBarra)}','${escAttr(c.descripcion)}')"
                class="btn btn-outline text-xs py-1 px-2 flex items-center gap-1">
          <svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor">
            <path d="M8 3.5a.5.5 0 0 0-1 0V9a.5.5 0 0 0 .252.434l3.5 2a.5.5 0 0 0 .496-.868L8 8.71V3.5z"/>
            <path d="M8 16A8 8 0 1 0 8 0a8 8 0 0 0 0 16zm7-8A7 7 0 1 1 1 8a7 7 0 0 1 14 0z"/>
          </svg>
          Historial
        </button>
        ${eyeBtn}
      </div>
    </div>`;
  }

  function toggleGrupo(sid) {
    const panel = document.getElementById('eqs-' + sid);
    const chev  = document.getElementById('chev-' + sid);
    if (!panel) return;
    const closed = panel.classList.toggle('hidden');
    if (chev) chev.style.transform = closed ? '' : 'rotate(180deg)';
  }

  // ── Búsqueda inteligente ─────────────────────────
  function buscar(q) {
    const query = (q || '').trim();
    _queryActual = query;   // persistir para el bg-refresh
    const cl = document.getElementById('clearBuscarProd');
    if (cl) cl.style.display = query ? 'flex' : 'none';

    if (!query) {
      _filtrados = [..._grupos];
      _render(_filtrados);
      return;
    }

    const qL     = query.toLowerCase();
    const tokens = qL.split(/\s+/).filter(Boolean);

    // ── Detectar coincidencia exacta (barcode / SKU / idProducto) ──
    let exactGrupo = null;
    for (const g of _grupos) {
      if (String(g.skuBase || '').toLowerCase() === qL ||
          String(g.base.idProducto || '').toLowerCase() === qL) {
        exactGrupo = g; break;
      }
      if (g.children.some(c => String(c.codigoBarra || '').toLowerCase() === qL)) {
        exactGrupo = g; break;
      }
    }

    // ── Filtrar: todos los tokens deben aparecer en algún campo ─────
    _filtrados = _grupos.filter(g => {
      const haystack = [
        String(g.base.descripcion || ''),
        String(g.skuBase || ''),
        String(g.base.idProducto || ''),
        ...g.children.map(c => String(c.descripcion  || '') + ' ' + String(c.codigoBarra || ''))
      ].join(' ').toLowerCase();
      return tokens.every(t => haystack.includes(t));
    });

    _render(_filtrados);

    // ── Señales visuales post-render ────────────────────────────────
    requestAnimationFrame(() => {
      if (exactGrupo) {
        const sid  = exactGrupo.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
        const card = document.getElementById('grp-' + sid);
        if (card) {
          card.classList.add('card-exact-match');
          // Auto-expandir si tiene hijos
          const panel = document.getElementById('eqs-' + sid);
          const chev  = document.getElementById('chev-' + sid);
          if (panel && panel.classList.contains('hidden')) {
            panel.classList.remove('hidden');
            if (chev) chev.style.transform = 'rotate(180deg)';
          }
          // Scroll suave al centro
          card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        // Atenuar los demás resultados parciales
        _filtrados.forEach(g => {
          if (g !== exactGrupo) {
            const s = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
            const el = document.getElementById('grp-' + s);
            if (el) el.classList.add('card-dim');
          }
        });
      } else {
        // Solo coincidencias parciales: leve realce
        _filtrados.forEach(g => {
          const sid = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
          const el  = document.getElementById('grp-' + sid);
          if (el) el.classList.add('card-hi');
        });
      }
    });
  }

  function buscarClear() {
    _queryActual = '';
    const inp = document.getElementById('inputBuscarProd');
    if (inp) inp.value = '';
    buscar('');
  }

  // ── Historial ───────────────────────────────────
  // codigos = array de barcodes del grupo (puede tener 1 o varios)
  function _movimientosLocal(codigos) {
    const set      = new Set(codigos.map(String));
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const gMap     = {};
    guias.forEach(g => { gMap[g.idGuia] = g; });

    // Movimientos de guías
    const guiaMovs = detalles
      .filter(d => set.has(String(d.codigoProducto)))
      .map(d => {
        const g    = gMap[d.idGuia] || {};
        const tipo = (g.tipo || '').toUpperCase();
        return {
          idGuia:    d.idGuia,
          fecha:     g.fecha  || d.fecha || '',
          tipo:      g.tipo   || '—',
          esIngreso: tipo.includes('INGRESO') || tipo.includes('ENTRADA') ||
                     (!tipo.includes('SALIDA') && parseFloat(d.cantidad || 0) > 0),
          cantidad:  Math.abs(parseFloat(d.cantidadRecibida || d.cantidadEsperada || d.cantidad || 0)),
          usuario:   g.usuario || d.usuario || '—',
          origen:    g.idProveedor || g.destino || '',
          estado:    g.estado || '',
          fuente:    'guia'
        };
      });

    // Ajustes del cache (si están disponibles)
    let ajusteMovs = [];
    try {
      const ajustes = OfflineManager.getAjustesCache ? OfflineManager.getAjustesCache() : [];
      ajusteMovs = ajustes
        .filter(a => set.has(String(a.codigoProducto)))
        .map(a => {
          const t = (a.tipoAjuste || '').toUpperCase();
          return {
            idGuia:    a.idAjuste || '',
            fecha:     a.fecha || '',
            tipo:      `Ajuste ${a.tipoAjuste || ''}`,
            esIngreso: t === 'INC' || t === 'INI',
            cantidad:  Math.abs(parseFloat(a.cantidadAjuste || 0)),
            usuario:   a.usuario || '—',
            origen:    a.motivo  || '',
            estado:    '',
            fuente:    'ajuste'
          };
        });
    } catch(e) {}

    return [...guiaMovs, ...ajusteMovs]
      .filter(m => m.cantidad > 0)
      .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
  }

  async function verHistorial(codigosStr, nombre) {
    // Acepta barcode único o pipe-separated para grupos multi-barcode
    const codigos = String(codigosStr).split('|').map(s => s.trim()).filter(Boolean);
    _histTarget = { codigos, codigo: codigos[0], nombre };

    document.getElementById('histNombre').textContent   = nombre;
    document.getElementById('histCodigo').textContent   = codigos.length > 1
      ? `${codigos[0]} +${codigos.length - 1} más`
      : codigos[0];
    document.getElementById('histPrintPanel')?.classList.add('hidden');
    document.getElementById('histList').innerHTML = '<div class="flex justify-center py-6"><div class="spinner"></div></div>';
    abrirSheet('sheetHistorial');

    // Stock total del grupo
    const stockTotal = codigos.reduce((sum, c) => sum + (_s(c).cantidadDisponible || 0), 0);
    const stockMin   = _s(codigos[0]).stockMinimo || 0;
    document.getElementById('histStockActual').textContent = fmt(stockTotal);
    document.getElementById('histStockMin').textContent    = fmt(stockMin);

    const local = _movimientosLocal(codigos);
    if (local.length) _renderHistorial(local, true, stockTotal);

    const res = await API.getHistorialStock(codigos.join(',')).catch(() => ({ ok: false }));
    if (res.ok && res.data?.length) {
      _renderHistorial(res.data, false, stockTotal);
    } else if (!local.length) {
      document.getElementById('histList').innerHTML =
        '<p class="text-slate-500 text-sm text-center py-6">Sin movimientos registrados</p>';
    }
  }

  function _renderHistorial(movs, esLocal, stockActual) {
    if (!movs.length && esLocal) return;
    const stock = stockActual ?? codigos.reduce((s, c) => s + (_s(c).cantidadDisponible || 0), 0);

    // Balance corriente hacia atrás desde stock actual
    let bal = stock;
    const conBal = movs.map(m => {
      const row = { ...m, bal };
      bal = m.esIngreso ? bal - m.cantidad : bal + m.cantidad;
      return row;
    });

    const fBadge = document.getElementById('histFuenteBadge');
    fBadge.textContent = esLocal ? '📦 Local' : '☁️ GAS';
    fBadge.className   = esLocal ? 'tag-blue text-xs' : 'tag-ok text-xs';

    // Tipos de movimiento para colores bancarios
    const _tipoMov = (m) => {
      const t = (m.tipo || '').toUpperCase();
      if (t.includes('INI'))     return 'ini';    // stock inicial
      if (m.fuente === 'ajuste' || t.includes('AJUSTE')) return m.esIngreso ? 'ajuste_inc' : 'ajuste_dec';
      return m.esIngreso ? 'ingreso' : 'salida';
    };
    const _estilos = {
      ingreso:    { bg:'bg-emerald-950', fg:'text-emerald-400', icon:'↑', sign:'+', label:'Ingreso'  },
      salida:     { bg:'bg-red-950',     fg:'text-red-400',     icon:'↓', sign:'−', label:'Salida'   },
      ajuste_inc: { bg:'bg-blue-950',    fg:'text-blue-400',    icon:'⊕', sign:'+', label:'Ajuste ▲' },
      ajuste_dec: { bg:'bg-amber-950',   fg:'text-amber-400',   icon:'⊖', sign:'−', label:'Ajuste ▼' },
      ini:        { bg:'bg-violet-950',  fg:'text-violet-400',  icon:'★', sign:'+', label:'Inicial'  }
    };

    document.getElementById('histList').innerHTML = conBal.map(m => {
      const tipo = _tipoMov(m);
      const e    = _estilos[tipo] || _estilos.ingreso;
      return `
      <div class="flex items-start gap-2.5 py-2.5 border-b" style="border-color:#1e293b">
        <div class="flex flex-col items-center gap-0.5 flex-shrink-0 mt-0.5">
          <span class="w-7 h-7 rounded-full flex items-center justify-center text-xs font-black ${e.bg} ${e.fg}">${e.icon}</span>
          <span class="text-xs ${e.fg} font-semibold" style="font-size:9px">${e.label}</span>
        </div>
        <div class="flex-1 min-w-0">
          <div class="flex items-baseline justify-between gap-1">
            <span class="font-black text-base ${e.fg}">${e.sign}${fmt(m.cantidad)}</span>
            <span class="text-xs text-slate-400 font-mono">${fmtFecha(m.fecha)}</span>
          </div>
          <p class="text-xs text-slate-400 truncate">${m.tipo}${m.idGuia ? ' · ' + m.idGuia : ''}</p>
          ${m.origen ? `<p class="text-xs text-slate-500 truncate">${m.origen}</p>` : ''}
          <div class="flex items-center gap-1 mt-0.5">
            <span class="text-xs text-slate-600">Saldo:</span>
            <span class="text-xs text-slate-200 font-bold">${fmt(m.bal)}</span>
            <span class="text-slate-700 text-xs">·</span>
            <span class="text-xs text-slate-500">${m.usuario}</span>
          </div>
        </div>
      </div>`;
    }).join('') || '<p class="text-slate-500 text-sm text-center py-6">Sin movimientos registrados</p>';
  }

  // ── Imprimir historial ──────────────────────────
  // 80mm ticket = 48 chars por línea (fuente estándar)
  async function imprimirHistorial() {
    if (!_histTarget) return;

    const W       = 48;
    const SEP     = '='.repeat(W);
    const SEP2    = '-'.repeat(W);
    const codigos = _histTarget.codigos || [_histTarget.codigo];
    const stock   = codigos.reduce((sum, c) => sum + (_s(c).cantidadDisponible || 0), 0);
    const stockMin = _s(_histTarget.codigo).stockMinimo || 0;

    toast('Generando ticket...', 'info', 2500);

    // Siempre pedir al GAS para tener datos completos; cache como fallback offline
    let movs = [];
    try {
      const res = await API.getHistorialStock(codigos.join(','));
      if (res.ok && res.data?.length) movs = res.data;
      else movs = _movimientosLocal(codigos);
    } catch { movs = _movimientosLocal(codigos); }

    const pad2 = n => String(n).padStart(2, '0');
    const now  = new Date();
    const fechaImpresion = `${pad2(now.getDate())}/${pad2(now.getMonth()+1)}/${now.getFullYear()}  ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;

    function _ddmm(s) {
      if (!s) return '  /  ';
      const d = new Date(s);
      return isNaN(d) ? String(s).slice(0, 5) : `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}`;
    }

    function _tipoLegible(m) {
      const t = (m.tipo || '').toUpperCase();
      if (t.includes('INI'))                                return 'Stock Inicial';
      if (m.fuente === 'ajuste' || t.includes('AJUSTE'))    return m.esIngreso ? 'Ajuste Ingreso' : 'Ajuste Salida';
      if (t.includes('INGRESO') || t.includes('ENTRADA'))   return 'Ingreso Guia';
      if (t.includes('SALIDA'))                             return 'Salida Guia';
      if (t.includes('ENVASADO'))                           return 'Envasado';
      if (t.includes('MERMA'))                              return 'Merma';
      return (m.tipo || '—').slice(0, 20);
    }

    // Columnas: fecha(5) + 2sp + tipo(padEnd 26) + monto(padStart 15) = 48
    const COL_TIPO  = 26;
    const COL_MON   = 15;
    const header = 'FECHA'.padEnd(7) + 'MOVIMIENTO'.padEnd(COL_TIPO) + 'CANTIDAD'.padStart(COL_MON);

    const movLines = movs.length
      ? movs.slice(0, 50).map(m => {
          const fecha  = _ddmm(m.fecha).padEnd(7);
          const tipo   = _tipoLegible(m).padEnd(COL_TIPO);
          const monto  = ((m.esIngreso ? '+' : '-') + fmt(m.cantidad, 0)).padStart(COL_MON);
          return fecha + tipo + monto;
        })
      : ['  Sin movimientos registrados.'];

    // Centrar título en W chars
    const center = s => s.padStart(Math.floor((W + s.length) / 2)).padEnd(W);
    const nombre = String(_histTarget.nombre || '').slice(0, W);

    const lines = [
      SEP,
      center('HISTORIAL DE STOCK'),
      center('ALMACEN CENTRAL - MOS'),
      SEP,
      `Producto : ${nombre}`,
      `Codigo   : ${_histTarget.codigo}`,
      `Impreso  : ${fechaImpresion}`,
      SEP2,
      header,
      SEP2,
      ...movLines,
      SEP2,
      `${'Stock actual'.padEnd(W - 10)}:${fmt(stock,    0).padStart(9)}`,
      `${'Minimo requerido'.padEnd(W - 10)}:${fmt(stockMin, 0).padStart(9)}`,
      `${'Estado'.padEnd(W - 10)}:${(stock > stockMin ? 'OK' : '! BAJO MINIMO').padStart(9)}`,
      SEP,
      ''
    ];
    const texto = lines.join('\n');

    const res = await API.imprimirHistorialStock({
      texto,
      codigoProducto: _histTarget.codigo
    }).catch(() => ({ ok: false }));
    toast(res.ok ? 'Impreso ✓' : 'No se pudo imprimir — revisa config GAS', res.ok ? 'ok' : 'warn');
  }

  // ── Aplicar query activa ─────────────────────────
  function _aplicarQuery() {
    if (_auditModo) {
      _filtrados = _grupos.filter(g => _auditDia?.skus.includes(g.skuBase));
      _render(_filtrados);
    } else if (_queryActual) {
      buscar(_queryActual);
    } else {
      _filtrados = [..._grupos];
      _render(_filtrados);
    }
  }

  // ── Auditoría diaria — helpers ────────────────────────────
  function _initAuditDia() {
    const hoy = new Date().toISOString().slice(0, 10);
    try {
      const raw = localStorage.getItem(_AUDIT_KEY);
      if (raw) {
        const d = JSON.parse(raw);
        if (d.fecha === hoy) { _auditDia = d; _actualizarBadge(); return; }
      }
    } catch(e) {}
    // Nuevo día → seleccionar 30 al azar
    const skus = _grupos.map(g => g.skuBase);
    const n = Math.min(30, skus.length);
    const shuffled = skus.slice().sort(() => Math.random() - 0.5).slice(0, n);
    _auditDia = { fecha: hoy, skus: shuffled, auditados: {} };
    _guardarAuditDia();
    _actualizarBadge();
  }

  function _guardarAuditDia() {
    try { localStorage.setItem(_AUDIT_KEY, JSON.stringify(_auditDia)); } catch(e) {}
  }

  function _esBarcodeAuditado(sku, cod) {
    return (_auditDia?.auditados?.[sku] || []).includes(String(cod));
  }

  function _esGrupoCompleto(sku) {
    if (!_auditDia) return false;
    const g = _grupos.find(g => g.skuBase === sku);
    if (!g) return false;
    const auditados = _auditDia.auditados?.[sku] || [];
    return g.children.every(c => auditados.includes(String(c.codigoBarra)));
  }

  function _pendientesCount() {
    if (!_auditDia) return 0;
    return _auditDia.skus.filter(sku => !_esGrupoCompleto(sku)).length;
  }

  function _actualizarBadge() {
    const btn = document.getElementById('btnAuditoriaDia');
    if (!btn || !_auditDia) return;
    btn.style.display = '';
    const pend = _pendientesCount();
    const total = _auditDia.skus.length;
    if (pend === 0) {
      btn.innerHTML = `<svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/></svg> Listo`;
      btn.className = 'audit-badge audit-badge-done' + (_auditModo ? ' active' : '');
    } else {
      btn.innerHTML = `<svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M5.5 2A3.5 3.5 0 0 0 2 5.5v5A3.5 3.5 0 0 0 5.5 14h5a3.5 3.5 0 0 0 3.5-3.5V8a.5.5 0 0 1 1 0v2.5a4.5 4.5 0 0 1-4.5 4.5h-5A4.5 4.5 0 0 1 1 10.5v-5A4.5 4.5 0 0 1 5.5 1H8a.5.5 0 0 1 0 1H5.5z"/><path d="M16 3a3 3 0 1 1-6 0 3 3 0 0 1 6 0z"/></svg> ${pend}&nbsp;<span style="opacity:.6;font-weight:400">/ ${total}</span>`;
      btn.className = 'audit-badge audit-badge-pending' + (_auditModo ? ' active' : '');
    }
    // Punto naranja en el nav (bottom + sidebar)
    const dot = document.getElementById('navDotAudit');
    if (dot) dot.style.display = pend > 0 ? '' : 'none';
    const sideDot = document.getElementById('sideNavDotAudit');
    if (sideDot) sideDot.style.display = pend > 0 ? '' : 'none';
  }

  function toggleAuditoriaDia() {
    if (!_auditDia) return;
    _auditModo = !_auditModo;
    // Limpiar búsqueda de texto si entramos en modo auditoría
    if (_auditModo) {
      _queryActual = '';
      const inp = document.getElementById('inputBuscarProd');
      if (inp) inp.value = '';
      const cl = document.getElementById('clearBuscarProd');
      if (cl) cl.style.display = 'none';
      _filtrados = _grupos.filter(g => _auditDia.skus.includes(g.skuBase));
    } else {
      _filtrados = [..._grupos];
    }
    _actualizarBadge();
    _render(_filtrados);
  }

  function abrirAuditBarcode(codigoBarra, nombre, skuBase) {
    const cod = String(codigoBarra);
    const s   = _s(cod);
    _auditTarget = { codigoBarra: cod, nombre, skuBase };
    document.getElementById('auditNombre').textContent   = nombre;
    document.getElementById('auditCodigo').textContent   = cod;
    document.getElementById('auditStockSis').textContent = fmt(s.cantidadDisponible || 0);
    document.getElementById('auditConteo').value         = '';
    document.getElementById('auditObs').value            = '';
    abrirSheet('sheetAudit');
  }

  function confirmarAuditoria() {
    if (!_auditTarget) return;
    const fisico = parseFloat(document.getElementById('auditConteo').value);
    if (isNaN(fisico) || fisico < 0) { toast('Ingresa el conteo físico', 'warn'); return; }
    const obs    = document.getElementById('auditObs').value.trim();
    const target = { ..._auditTarget };

    // ── Optimistic: calcular diff antes de tocar stockMap ──────
    const stockSistema = _s(target.codigoBarra).cantidadDisponible || 0;
    const diff         = fisico - stockSistema;

    // Marcar en localStorage inmediatamente
    const sku = target.skuBase;
    if (_auditDia) {
      if (!_auditDia.auditados[sku]) _auditDia.auditados[sku] = [];
      const cod = String(target.codigoBarra);
      if (!_auditDia.auditados[sku].includes(cod)) _auditDia.auditados[sku].push(cod);
      _guardarAuditDia();
    }

    // Actualizar stock local optimisticamente
    _stockMap[target.codigoBarra] = {
      ...(_stockMap[target.codigoBarra] || {}),
      cantidadDisponible: fisico
    };

    // Cerrar y re-render sin esperar al servidor
    cerrarSheet('sheetAudit');
    _auditTarget = null;
    const msg = Math.abs(diff) <= 0.5
      ? '✅ Sin diferencias'
      : `⚠️ Diferencia: ${diff > 0 ? '+' : ''}${fmt(diff, 2)}`;
    toast(msg, Math.abs(diff) <= 0.5 ? 'ok' : 'warn', 4000);
    _actualizarBadge();
    _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
    _aplicarQuery();

    // ── Enviar al servidor en segundo plano ────────────────────
    API.auditarProducto({
      codigoBarra: String(target.codigoBarra),
      stockFisico: fisico,
      observacion: obs,
      usuario:     window.WH_CONFIG?.usuario || ''
    }).then(res => {
      if (!res.ok) toast('Error al guardar en servidor: ' + (res.error || ''), 'danger', 5000);
    }).catch(() => {
      toast('Sin conexión — auditoría en cola', 'warn', 4000);
    });
  }

  // ── Ajuste manual ───────────────────────────────
  function abrirAjusteDesdeHistorial() {
    if (!_histTarget) return;
    cerrarSheet('sheetHistorial');
    abrirAjuste(_histTarget.codigo, _histTarget.nombre);
  }

  function abrirAjuste(codigoBarra, nombre) {
    _ajusteTarget = { codigoBarra: String(codigoBarra), nombre };
    document.getElementById('ajusteNombre').textContent   = nombre;
    document.getElementById('ajusteCodigo').textContent   = String(codigoBarra);
    const s = _s(String(codigoBarra));
    document.getElementById('ajusteStockSis').textContent = fmt(s.cantidadDisponible || 0);
    document.getElementById('ajusteCant').value           = '';
    document.getElementById('ajusteMotivo').value         = '';
    document.getElementById('ajustePreview').textContent  = '';
    abrirSheet('sheetAjuste');
  }

  function previewAjuste() {
    if (!_ajusteTarget) return;
    const stockReal = parseFloat(document.getElementById('ajusteCant').value);
    const el        = document.getElementById('ajustePreview');
    if (isNaN(stockReal) || document.getElementById('ajusteCant').value === '') {
      el.textContent = ''; return;
    }
    const stockActual = _s(_ajusteTarget.codigoBarra).cantidadDisponible || 0;
    const diff        = stockReal - stockActual;
    if (Math.abs(diff) < 0.01) {
      el.className   = 'text-xs text-center mb-3 h-5 text-slate-400';
      el.textContent = 'Sin diferencia';
    } else if (diff > 0) {
      el.className   = 'text-xs text-center mb-3 h-5 text-emerald-400 font-bold';
      el.textContent = `▲ +${fmt(diff, 2)} unidades`;
    } else {
      el.className   = 'text-xs text-center mb-3 h-5 text-red-400 font-bold';
      el.textContent = `▼ ${fmt(diff, 2)} unidades`;
    }
  }

  function confirmarAjuste() {
    if (!_ajusteTarget) return;
    const stockReal = parseFloat(document.getElementById('ajusteCant').value);
    if (isNaN(stockReal) || stockReal < 0) { toast('Ingresa el stock real', 'warn'); return; }
    const motivo      = document.getElementById('ajusteMotivo').value.trim();
    const target      = { ..._ajusteTarget };
    const stockActual = _s(target.codigoBarra).cantidadDisponible || 0;
    const diff        = stockReal - stockActual;

    if (Math.abs(diff) < 0.01) {
      toast('El stock real coincide con el sistema, no hay cambio', 'info', 3000);
      cerrarSheet('sheetAjuste');
      return;
    }

    const tipo = diff > 0 ? 'INC' : 'DEC';

    // Optimistic
    _stockMap[target.codigoBarra] = {
      ...(_stockMap[target.codigoBarra] || {}),
      cantidadDisponible: stockReal
    };
    cerrarSheet('sheetAjuste');
    _ajusteTarget = null;
    toast(`${diff > 0 ? '▲' : '▼'} Ajuste ${diff > 0 ? '+' : ''}${fmt(diff, 2)} — stock: ${fmt(stockReal)}`, 'ok', 3000);
    _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
    _aplicarQuery();

    // Background sync
    API.crearAjuste({
      codigoProducto: target.codigoBarra,
      tipoAjuste:     tipo,
      cantidadAjuste: Math.abs(diff),
      motivo:         motivo || 'Ajuste manual',
      usuario:        window.WH_CONFIG?.usuario || ''
    }).then(res => {
      if (!res.ok) toast('Error al guardar ajuste: ' + (res.error || ''), 'danger', 5000);
    }).catch(() => {
      toast('Sin conexión — ajuste en cola', 'warn', 4000);
    });
  }

  // ── Cargar ──────────────────────────────────────
  async function cargar() {
    loading('listProductos', true);
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    _buildMap(OfflineManager.getStockCache());
    _grupos = _agrupar(prods, equivs);
    _initAuditDia();
    _aplicarQuery();

    // Refrescar stock via endpoint compartido (throttled) — evita llamada individual a getStock
    OfflineManager.precargarOperacional().then(() => {
      _buildMap(OfflineManager.getStockCache());
      _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
      _actualizarBadge();
      _aplicarQuery();
    }).catch(() => {});
  }

  // Re-render desde caché sin spinner ni API call (llamado por wh:data-refresh)
  function silentRefresh() {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    _buildMap(OfflineManager.getStockCache());
    _grupos = _agrupar(prods, equivs);
    _aplicarQuery();
  }

  return { cargar, silentRefresh, buscar, buscarClear, toggleGrupo, toggleAuditoriaDia,
           abrirAuditBarcode, confirmarAuditoria,
           abrirAjuste, abrirAjusteDesdeHistorial, previewAjuste, confirmarAjuste,
           verHistorial, imprimirHistorial };
})();


// ════════════════════════════════════════════════
// MEMBRETE VIEW — cola de impresión de membretes
// ════════════════════════════════════════════════
const MembreteView = (() => {
  let _cola = []; // [{ prod, allEan }]

  function _buildEan(prod) {
    const equivs = OfflineManager.getEquivalenciasCache();
    const altCodes = equivs
      .filter(e => e.skuBase === (prod.skuBase || prod.idProducto) && String(e.activo) === '1' && e.codigoBarra)
      .map(e => String(e.codigoBarra));
    const allEan = [];
    if (prod.codigoBarra) allEan.push(String(prod.codigoBarra));
    altCodes.forEach(c => { if (!allEan.includes(c)) allEan.push(c); });
    return allEan;
  }

  function _renderCola() {
    const colaEl = document.getElementById('memCola');
    const listEl = document.getElementById('memColaList');
    const btn    = document.getElementById('btnImprimirMembrete');
    if (!listEl) return;

    if (!_cola.length) {
      if (colaEl) colaEl.style.display = 'none';
      if (btn) { btn.disabled = true; btn.innerHTML = _btnHTML('Imprimir membretes'); }
      return;
    }

    if (colaEl) colaEl.style.display = '';
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = _btnHTML(`Imprimir ${_cola.length} membrete${_cola.length > 1 ? 's' : ''}`);
    }

    listEl.innerHTML = _cola.map((item, idx) => {
      const n = item.allEan.length;
      const tagColor = n > 1 ? '#3b82f6' : '#475569';
      const tagBg    = n > 1 ? 'rgba(59,130,246,.15)' : 'rgba(71,85,105,.15)';
      return `<div style="display:flex;align-items:center;gap:8px;padding:9px 10px;
                          background:#1e293b;border-radius:8px;margin-bottom:5px;">
        <div style="flex:1;min-width:0;display:flex;align-items:center;gap:8px;">
          <span style="font-size:13px;font-weight:700;color:#e2e8f0;
                       white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(item.prod.descripcion || item.prod.idProducto)}</span>
          <span style="flex-shrink:0;font-size:10px;font-weight:600;padding:2px 7px;
                       border-radius:99px;color:${tagColor};background:${tagBg};">
            ${n} codebar${n !== 1 ? 's' : ''}</span>
        </div>
        <button onclick="MembreteView.remover(${idx})"
                style="background:none;border:none;cursor:pointer;color:#475569;
                       font-size:18px;line-height:1;padding:2px 4px;flex-shrink:0;"
                title="Quitar">×</button>
      </div>`;
    }).join('');
  }

  function _btnHTML(label) {
    return `<svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
      <path d="M2.5 8a.5.5 0 1 0 0-1 .5.5 0 0 0 0 1z"/>
      <path d="M5 1a2 2 0 0 0-2 2v2H2a2 2 0 0 0-2 2v3a2 2 0 0 0 2 2h1v1a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2v-1h1a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-1V3a2 2 0 0 0-2-2H5zM4 3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2H4V3zm1 5a2 2 0 0 0-2 2v1H2a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v-1a2 2 0 0 0-2-2H5zm7 2v3a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1z"/>
    </svg> ${label}`;
  }

  function buscar(q) {
    const val  = (q || '').trim();
    const sEl  = document.getElementById('memSugerencias');
    if (!sEl) return;
    if (val.length < 2) { sEl.style.display = 'none'; sEl.innerHTML = ''; return; }

    const ql      = val.toLowerCase();
    const prods   = OfflineManager.getProductosCache();
    const equivs  = OfflineManager.getEquivalenciasCache();
    const matches = prods.filter(p =>
      String(p.descripcion || '').toLowerCase().includes(ql) ||
      String(p.idProducto  || '').toLowerCase().includes(ql) ||
      String(p.skuBase     || '').toLowerCase().includes(ql) ||
      String(p.codigoBarra || '').includes(val) ||
      equivs.some(e => e.skuBase === (p.skuBase || p.idProducto) && String(e.codigoBarra || '').includes(val))
    ).slice(0, 8);

    if (!matches.length) {
      sEl.innerHTML = `<div style="padding:10px 12px;font-size:12px;color:#64748b;">Sin resultados</div>`;
    } else {
      sEl.innerHTML = matches.map(p => {
        const yaEnCola = _cola.some(i => i.prod.idProducto === p.idProducto);
        return `<button onclick="MembreteView.seleccionar('${escAttr(p.idProducto)}')"
                style="display:flex;align-items:center;gap:8px;width:100%;text-align:left;
                       padding:10px 12px;background:none;border:none;cursor:pointer;
                       border-bottom:1px solid #1e293b;transition:background .1s;"
                onmouseenter="this.style.background='#1e293b'"
                onmouseleave="this.style.background='none'">
          <span style="flex:1;font-size:13px;font-weight:600;color:${yaEnCola ? '#475569' : '#e2e8f0'};
                       white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${escHtml(p.descripcion || p.idProducto)}
          </span>
          ${yaEnCola ? '<span style="font-size:10px;color:#22c55e;flex-shrink:0;">✓ en cola</span>' : ''}
          <span style="font-size:11px;color:#475569;font-family:monospace;flex-shrink:0;">${escHtml(p.idProducto)}</span>
        </button>`;
      }).join('');
    }
    sEl.style.display = 'block';
  }

  function seleccionar(idProducto) {
    const prods = OfflineManager.getProductosCache();
    const prod  = prods.find(p => p.idProducto === idProducto);
    if (!prod) return;

    if (_cola.some(i => i.prod.idProducto === idProducto)) {
      toast('Ya está en la cola de impresión', 'warn');
    } else {
      _cola.push({ prod, allEan: _buildEan(prod) });
      vibrate(8);
    }

    // Limpiar búsqueda y devolver foco para agregar otro
    const inp = document.getElementById('memBuscar');
    if (inp) { inp.value = ''; inp.focus(); }
    const sEl = document.getElementById('memSugerencias');
    if (sEl) { sEl.style.display = 'none'; sEl.innerHTML = ''; }
    const st = document.getElementById('memStatus');
    if (st) st.style.display = 'none';

    _renderCola();
  }

  function remover(idx) {
    _cola.splice(idx, 1);
    vibrate(8);
    _renderCola();
  }

  function vaciarCola() {
    _cola = [];
    _renderCola();
    const st = document.getElementById('memStatus');
    if (st) st.style.display = 'none';
  }

  function limpiar() {
    const inp = document.getElementById('memBuscar');
    if (inp) inp.value = '';
    const sEl = document.getElementById('memSugerencias');
    if (sEl) { sEl.style.display = 'none'; sEl.innerHTML = ''; }
  }

  function abrirScanner() {
    abrirScannerPara('memBuscar', code => {
      const inp = document.getElementById('memBuscar');
      if (inp) inp.value = code;
      buscar(code);
      const prods  = OfflineManager.getProductosCache();
      const equivs = OfflineManager.getEquivalenciasCache();
      const exact  = prods.find(p =>
        p.idProducto === code || p.codigoBarra === code ||
        equivs.some(e => e.idProducto === p.idProducto && String(e.codigoBarra) === code)
      );
      if (exact) seleccionar(exact.idProducto);
    });
  }

  function imprimir() {
    if (!_cola.length) return;
    const trabajos = [..._cola];
    const n = trabajos.length;

    // Optimista: limpiar cola y dar feedback inmediato
    _cola = [];
    _renderCola();
    vibrate(15);
    toast(`${n} membrete${n > 1 ? 's' : ''} enviados a impresora`, 'ok');
    const st = document.getElementById('memStatus');
    if (st) { st.style.display = ''; st.textContent = `Enviando ${n} membrete${n > 1 ? 's' : ''}…`; st.style.color = '#94a3b8'; }

    // Enviar en segundo plano — sin await, sin bloquear UI
    (async () => {
      let errCount = 0;
      for (const item of trabajos) {
        const res = await API.imprimirMembrete({
          idProducto: item.prod.idProducto,
          barcodes:   JSON.stringify(item.allEan),
          skuGrupal:  item.prod.skuBase || item.prod.idProducto
        }).catch(() => ({ ok: false }));
        if (!res.ok) errCount++;
      }
      if (errCount > 0) {
        toast(`${errCount} membrete${errCount > 1 ? 's' : ''} con error — verifica impresora`, 'danger');
        if (st) { st.textContent = `${errCount} con error`; st.style.color = '#f87171'; }
      } else {
        if (st) st.style.display = 'none';
      }
    })();
  }

  return { buscar, seleccionar, limpiar, remover, vaciarCola, abrirScanner, imprimir };
})();

// ════════════════════════════════════════════════
// CONFIG VIEW
// ════════════════════════════════════════════════
const ConfigView = (() => {
  // Guardar solo la URL GAS (sección Conexión en Tools)
  function guardar() {
    const gasUrl = document.getElementById('cfgGasUrl').value.trim();
    if (!gasUrl) { toast('Ingresa la URL del GAS', 'warn'); return; }
    window.WH_CONFIG.gasUrl = gasUrl;
    localStorage.setItem('wh_gas_url', gasUrl);
    const el = document.getElementById('toolsGasUrl');
    if (el) el.textContent = gasUrl;
    toast('URL guardada', 'ok');
  }

  // Guardar configuración de impresión (PrintNode + días alerta)
  async function guardarImpresion() {
    const printKey   = document.getElementById('cfgPrintKey').value.trim();
    const printId    = document.getElementById('cfgPrintId').value.trim();
    const diasAlerta = document.getElementById('cfgDiasAlerta').value;
    if (printKey)   await API.setConfig('PRINTNODE_API_KEY',   printKey);
    if (printId)    await API.setConfig('PRINTNODE_PRINTER_ID', printId);
    if (diasAlerta) await API.setConfig('DIAS_ALERTA_VENC',    diasAlerta);
    toast('Configuración de impresión guardada', 'ok');
  }

  return { guardar, guardarImpresion };
})();

// ════════════════════════════════════════════════
// Init
// ════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => App.init());
