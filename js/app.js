// warehouseMos — app.js  Lógica de la aplicación
'use strict';

// ════════════════════════════════════════════════
// [WH] Modales custom genéricos (reemplazan confirm/prompt/alert nativos).
// Paleta WH: verde esmeralda + ámbar. Vanilla DOM, funcionan en cualquier
// función async — drop-in replacement de los nativos del browser.
// Uso:
//   if (!await _whConfirm('texto', opts)) return;
//   const v = await _whPrompt('label', defaultValue, opts);
//   await _whAlert('mensaje', opts);
// ════════════════════════════════════════════════
function _whEsc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function _whConfirm(msg, opts) {
  opts = opts || {};
  return new Promise((resolve) => {
    const accent = opts.danger ? 'danger' : (opts.warning ? 'warn' : 'primary');
    const icoMap = { danger: '⚠', warn: '🟡', primary: '❓' };
    const titulo = opts.titulo || (opts.danger ? 'Confirmar acción' : '¿Continuar?');
    const backdrop = document.createElement('div');
    backdrop.className = 'wh-modal-generic-backdrop';
    backdrop.innerHTML = `
      <div class="wh-modal-generic wh-modal-generic-${accent}" role="dialog" aria-modal="true">
        <div class="wh-modal-generic-head">
          <span class="wh-modal-generic-ico">${icoMap[accent]}</span>
          <strong>${_whEsc(titulo)}</strong>
          <button class="wh-modal-generic-close" data-cancel>✕</button>
        </div>
        <div class="wh-modal-generic-body">${(msg || '').split('\n\n').map(p => '<p>' + _whEsc(p).replace(/\n/g, '<br>') + '</p>').join('')}</div>
        <div class="wh-modal-generic-foot">
          <button class="wh-modal-generic-btn wh-modal-generic-btn-ghost" data-cancel>${_whEsc(opts.cancelText || 'Cancelar')}</button>
          <button class="wh-modal-generic-btn wh-modal-generic-btn-${accent}" data-ok>${_whEsc(opts.okText || (opts.danger ? 'Sí, continuar' : 'Aceptar'))}</button>
        </div>
      </div>`;
    const cerrar = (v) => {
      backdrop.classList.remove('is-open');
      setTimeout(() => { try { backdrop.remove(); } catch(_){} }, 200);
      resolve(v);
    };
    backdrop.addEventListener('click', (ev) => {
      const t = ev.target;
      if (t === backdrop || (t.closest && t.closest('[data-cancel]'))) cerrar(false);
      else if (t.closest && t.closest('[data-ok]')) cerrar(true);
    });
    document.body.appendChild(backdrop);
    requestAnimationFrame(() => backdrop.classList.add('is-open'));
  });
}
function _whPrompt(label, defaultValue, opts) {
  opts = opts || {};
  return new Promise((resolve) => {
    const isLong = opts.textarea === true;
    const inputHtml = isLong
      ? `<textarea class="wh-modal-generic-input" rows="3" maxlength="${opts.maxlength || 240}" placeholder="${_whEsc(opts.placeholder || '')}">${_whEsc(defaultValue || '')}</textarea>`
      : `<input type="${opts.inputType || 'text'}" inputmode="${opts.inputMode || ''}" maxlength="${opts.maxlength || ''}" class="wh-modal-generic-input" value="${_whEsc(defaultValue || '')}" placeholder="${_whEsc(opts.placeholder || '')}" />`;
    const backdrop = document.createElement('div');
    backdrop.className = 'wh-modal-generic-backdrop';
    backdrop.innerHTML = `
      <div class="wh-modal-generic wh-modal-generic-primary" role="dialog" aria-modal="true">
        <div class="wh-modal-generic-head">
          <span class="wh-modal-generic-ico">✏</span>
          <strong>${_whEsc(opts.titulo || 'Ingresar dato')}</strong>
          <button class="wh-modal-generic-close" data-cancel>✕</button>
        </div>
        <div class="wh-modal-generic-body">
          <label class="wh-modal-generic-label">${_whEsc(label || '')}</label>
          ${inputHtml}
        </div>
        <div class="wh-modal-generic-foot">
          <button class="wh-modal-generic-btn wh-modal-generic-btn-ghost" data-cancel>Cancelar</button>
          <button class="wh-modal-generic-btn wh-modal-generic-btn-primary" data-ok>${_whEsc(opts.okText || 'Aceptar')}</button>
        </div>
      </div>`;
    const cerrar = (v) => {
      backdrop.classList.remove('is-open');
      setTimeout(() => { try { backdrop.remove(); } catch(_){} }, 200);
      resolve(v);
    };
    const getVal = () => {
      const i = backdrop.querySelector('.wh-modal-generic-input');
      return i ? String(i.value || '') : '';
    };
    backdrop.addEventListener('click', (ev) => {
      const t = ev.target;
      if (t === backdrop || (t.closest && t.closest('[data-cancel]'))) cerrar(null);
      else if (t.closest && t.closest('[data-ok]')) cerrar(getVal());
    });
    backdrop.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' && !isLong) cerrar(getVal());
      else if (ev.key === 'Escape') cerrar(null);
    });
    document.body.appendChild(backdrop);
    requestAnimationFrame(() => backdrop.classList.add('is-open'));
    setTimeout(() => {
      try { backdrop.querySelector('.wh-modal-generic-input')?.focus(); } catch(_){}
      try { backdrop.querySelector('.wh-modal-generic-input')?.select(); } catch(_){}
    }, 220);
  });
}
function _whAlert(msg, opts) {
  opts = opts || {};
  return new Promise((resolve) => {
    const backdrop = document.createElement('div');
    backdrop.className = 'wh-modal-generic-backdrop';
    backdrop.innerHTML = `
      <div class="wh-modal-generic wh-modal-generic-primary" role="dialog" aria-modal="true">
        <div class="wh-modal-generic-head">
          <span class="wh-modal-generic-ico">💡</span>
          <strong>${_whEsc(opts.titulo || 'Aviso')}</strong>
          <button class="wh-modal-generic-close" data-ok>✕</button>
        </div>
        <div class="wh-modal-generic-body">${(msg || '').split('\n\n').map(p => '<p>' + _whEsc(p).replace(/\n/g, '<br>') + '</p>').join('')}</div>
        <div class="wh-modal-generic-foot">
          <button class="wh-modal-generic-btn wh-modal-generic-btn-primary" data-ok>OK</button>
        </div>
      </div>`;
    const cerrar = () => {
      backdrop.classList.remove('is-open');
      setTimeout(() => { try { backdrop.remove(); } catch(_){} }, 200);
      resolve();
    };
    backdrop.addEventListener('click', (ev) => {
      if (ev.target === backdrop || ev.target.closest('[data-ok]')) cerrar();
    });
    document.body.appendChild(backdrop);
    requestAnimationFrame(() => backdrop.classList.add('is-open'));
  });
}

// ════════════════════════════════════════════════
// Vibración táctil (feedback)
// ════════════════════════════════════════════════
let _userActivated = false;
document.addEventListener('pointerdown', () => { _userActivated = true; }, { once: true });

function vibrate(ms = 10) {
  if (_userActivated && navigator.vibrate) navigator.vibrate(ms);
}

// ════════════════════════════════════════════════
// VOZ — Web Speech API para anuncios audibles en almacén
// Útil para que el operador no tenga que mirar la pantalla.
// Funciona offline (síntesis local del SO).
// ════════════════════════════════════════════════
function _vozAnunciar(texto, opts) {
  try {
    if (!('speechSynthesis' in window)) return;
    // Cancelar cualquier anuncio anterior para no apilarlos
    speechSynthesis.cancel();
    const utt = new SpeechSynthesisUtterance(String(texto || ''));
    utt.lang   = (opts && opts.lang) || 'es-PE';
    utt.rate   = (opts && opts.rate) || 1.05;
    utt.pitch  = (opts && opts.pitch) || 1.0;
    utt.volume = (opts && opts.volume) || 1.0;
    // Preferir voz en español si está disponible
    const voces = speechSynthesis.getVoices() || [];
    const vEs = voces.find(v => /^es/i.test(v.lang)) || voces.find(v => /spanish/i.test(v.name));
    if (vEs) utt.voice = vEs;
    speechSynthesis.speak(utt);
  } catch(_){}
}
function _vozCancelar() {
  try { if ('speechSynthesis' in window) speechSynthesis.cancel(); } catch(_){}
}
window._vozAnunciar = _vozAnunciar;

// ════════════════════════════════════════════════
// F5 — ALERTA FULLSCREEN PICKUP NUEVO
// Overlay + sonido fuerte repetido + vibración + voz.
// Almacén con ruido = combo agresivo para captar atención.
// ════════════════════════════════════════════════
let _apnSonidoTimer = null;
function mostrarAlertaPickupNuevo(pickup) {
  if (!pickup) return;
  const overlay = document.getElementById('alertaPickupNuevo');
  if (!overlay) return;
  // Llenar contenido
  const zonaEl = document.getElementById('apnZona');
  const metaEl = document.getElementById('apnMeta');
  const items  = Array.isArray(pickup.items) ? pickup.items : [];
  const totalUds = items.reduce((s, it) => s + (parseFloat(it.solicitado) || 0), 0);
  const zonaTxt = pickup.idZona || 'sin zona';
  if (zonaEl) zonaEl.textContent = zonaTxt;
  if (metaEl) metaEl.textContent = items.length + ' productos · ' + Math.round(totalUds) + ' uds · ' + (pickup.creadoPor || 'cajero');
  overlay.style.display = 'flex';
  // Sonido fuerte repetido (almacén con ruido) — 3 ciclos cada 1.6s
  try { SoundFX.pickupAlerta(); } catch(_){}
  if (_apnSonidoTimer) clearInterval(_apnSonidoTimer);
  let ciclos = 0;
  _apnSonidoTimer = setInterval(() => {
    if (overlay.style.display === 'none' || ciclos >= 2) {
      clearInterval(_apnSonidoTimer); _apnSonidoTimer = null; return;
    }
    try { SoundFX.pickupAlerta(); } catch(_){}
    ciclos++;
  }, 1700);
  // Vibración fuerte: 3 pulsos de 400ms con pausas
  if (_userActivated && navigator.vibrate) {
    navigator.vibrate([400, 200, 400, 200, 600]);
  }
  // Anuncio por voz — un poco después del primer beep para que se entienda
  // Ejemplo: "¡Pedido nuevo para zona ZONA-01! 23 productos."
  setTimeout(() => {
    _vozAnunciar(`¡Pedido nuevo para zona ${zonaTxt}! ${items.length} producto${items.length!==1?'s':''}.`,
                 { rate: 1.0, volume: 1.0 });
  }, 700);
  // Repetir la voz una vez más a los 4s si no la han cerrado todavía
  setTimeout(() => {
    if (overlay.style.display !== 'none') {
      _vozAnunciar(`Pedido nuevo · zona ${zonaTxt}.`, { rate: 1.05 });
    }
  }, 4200);
}
function cerrarAlertaPickupNuevo() {
  const overlay = document.getElementById('alertaPickupNuevo');
  if (overlay) overlay.style.display = 'none';
  if (_apnSonidoTimer) { clearInterval(_apnSonidoTimer); _apnSonidoTimer = null; }
  _vozCancelar();
}
// Exponer al window para que onclick en HTML lo encuentre
window.mostrarAlertaPickupNuevo = mostrarAlertaPickupNuevo;
window.cerrarAlertaPickupNuevo  = cerrarAlertaPickupNuevo;

// Test rápido desde DevTools console — valida overlay + sonido + voz sin
// esperar un pickup real. Uso: _testAlertaPickup()
window._testAlertaPickup = function(zona) {
  mostrarAlertaPickupNuevo({
    idZona: zona || 'ZONA-DEMO',
    creadoPor: 'tester',
    items: [
      { skuBase: 'TEST001', solicitado: 5 },
      { skuBase: 'TEST002', solicitado: 3 },
      { skuBase: 'TEST003', solicitado: 12 }
    ]
  });
};

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

// [v2.13.189] Scroll-lock del fondo mientras haya un sheet abierto VÍA abrirSheet.
// Se rastrea con un Set (no con query al DOM) para que los sheets que se abren por
// classList directo (navLp/actionDia/typePicker/etc.) NO contaminen el estado y
// dejen el body pegado. Backstop: nav() limpia el Set y libera el overflow.
const _openSheets = new Set();
function _syncSheetScrollLock() {
  try { document.body.style.overflow = _openSheets.size ? 'hidden' : ''; } catch(_){}
}
function _resetSheetScrollLock() {
  _openSheets.clear();
  try { document.body.style.overflow = ''; } catch(_){}
}

function abrirSheet(id) {
  const overlay = document.getElementById('overlay' + id.replace('sheet', ''));
  const sheet = document.getElementById(id);
  overlay?.classList.add('open');
  sheet?.classList.add('open');
  _openSheets.add(id);
  _syncSheetScrollLock();
}

function cerrarSheet(id) {
  const overlay = document.getElementById('overlay' + id.replace('sheet', ''));
  const sheet = document.getElementById(id);
  overlay?.classList.remove('open');
  sheet?.classList.remove('open');
  _openSheets.delete(id);
  _syncSheetScrollLock();
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

// Convierte string de fecha del sheet (yyyy-MM-dd) a Date local sin drift UTC.
// [v2.13.209] Tolerante a 3 formatos de `fecha` que conviven tras el cutover a Supabase:
//   - 'yyyy-MM-dd'                 (guías VIEJAS: sync de Sheet con celda Date → _sheetToObjects yyyy-MM-dd)
//   - 'yyyy-MM-ddTHH:mm:ss.sssZ'   (guías NUEVAS, escritura directa: timestamptz ISO-UTC con Z)
//   - 'yyyy-MM-dd HH:mm:ss+00'     (variante timestamptz con ESPACIO en vez de 'T' → Safari/iOS la rechaza)
// El espacio entre fecha y hora NO es ISO-8601: V8 lo tolera pero Safari devuelve Invalid Date
// (→ fecha "desaparecía" en iPhone/iPad). Normalizamos el primer espacio a 'T' antes de parsear.
function _parseLocalDate(s) {
  if (!s) return new Date(NaN);
  let str = String(s).trim();
  // Solo fecha (yyyy-MM-dd): mediodía local para evitar drift de TZ al mostrar día.
  if (str.length <= 10) return new Date(str + 'T12:00:00');
  // Normaliza 'yyyy-MM-dd HH:mm...' → 'yyyy-MM-ddTHH:mm...' (Safari-safe). Si ya tiene 'T', no cambia.
  if (str.indexOf('T') === -1) str = str.replace(' ', 'T');
  // Normaliza offset de timestamptz Postgres a ISO-8601 estricto (Safari/V8 lo exigen):
  //   '+00'   → '+00:00'   ·   '-05'   → '-05:00'     (offset de solo horas)
  //   '+0000' → '+00:00'   ·   '-0500' → '-05:00'     (offset sin ':')
  // 'Z' y '+HH:mm' ya válidos se dejan intactos. Sin esto, Safari devuelve Invalid Date y la
  // fecha "desaparece" (agrupación → 'Sin fecha', _fmtCorta → string crudo).
  str = str.replace(/([+-]\d{2})(\d{2})?$/, (m, hh, mm) => `${hh}:${mm || '00'}`);
  return new Date(str);
}

// [v2.13.181] Día (YYYY-MM-DD) SIEMPRE en zona horaria de Perú, sin importar
// cómo esté configurada la TZ del dispositivo. Espeja al backend, que usa
// Utilities.formatDate(..., 'America/Lima', 'yyyy-MM-dd'). Así frontend y
// servidor agrupan el MISMO preingreso en el MISMO día aunque una tablet quede
// mal configurada en otra zona (antes .getFullYear()/getDate() usaba la TZ local
// y podía mandar un preingreso de cerca de medianoche al día equivocado).
const WH_TZ = 'America/Lima';
function _diaPeru(fechaInput) {
  const d = fechaInput instanceof Date ? fechaInput : _parseLocalDate(fechaInput);
  if (!d || isNaN(d)) return '';
  try {
    return new Intl.DateTimeFormat('en-CA', { timeZone: WH_TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
  } catch (e) {
    // Fallback sin Intl (navegador muy viejo): Perú es UTC-5 todo el año (sin
    // DST), así que el instante −5h en UTC = fecha de Lima. Determinista, NO
    // depende de la TZ del dispositivo.
    return new Date(d.getTime() - 5 * 3600000).toISOString().slice(0, 10);
  }
}
function _hoyPeru()  { return _diaPeru(new Date()); }
function _ayerPeru() { return _diaPeru(new Date(Date.now() - 86400000)); }

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

// Hora desde timestamp embebido en ID, en TZ Lima.
// [v2.13.209] IDs viejos (GAS): 'G1745123456789' → 13 dígitos limpios. IDs nuevos (escritura
// directa): 'G_L1745123456789ab3xy9' → localId 'L'+Date.now()+random base36; el random PUEDE
// traer dígitos, así que un replace(/\D/g,'') global concatenaba basura (ej. 15 dígitos → año 7615).
// Extraemos SOLO la primera corrida de 13 dígitos (el ms epoch). Para PI1745... también funciona.
function _horaDesdeId(id) {
  const m = String(id || '').match(/(\d{13})/);   // primer bloque de 13 dígitos = ms epoch
  const ts = m ? parseInt(m[1]) : 0;
  if (!ts || ts < 1e12) return '';
  try {
    return new Intl.DateTimeFormat('es-PE', { timeZone: WH_TZ, hour: '2-digit', minute: '2-digit', hour12: false }).format(new Date(ts));
  } catch (e) {
    return new Date(ts - 5 * 3600000).toISOString().slice(11, 16);
  }
}

// Escapa para insertar en atributos onclick="..." (evita romper comillas)
function escAttr(s) { return String(s || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '&quot;'); }
function escHtml(s)  { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ════════════════════════════════════════════════
// [v2.13.3] Estados por carreta (LLENA/MEDIA/VACIA)
// Modelo del cargador: { id, nombre, carretas, estados[] } con
// estados.length === carretas. Si no hay `estados` (legacy), se asume todas
// LLENAS. [v2.13.180] El concepto de "tarifa de cargador" fue ELIMINADO: el
// cargador cobra en caja; el resumen del día solo muestra cargadores + estados.
// ════════════════════════════════════════════════
const ESTADOS_CARGA = ['LLENA', 'MEDIA', 'VACIA'];
const EMOJI_CARGA   = { LLENA: '🟢', MEDIA: '🟡', VACIA: '🔴' };
const LABEL_CARGA   = { LLENA: 'Llena', MEDIA: 'Media', VACIA: 'Casi vacía' };

function _normalizarCargador(c) {
  if (!c || typeof c !== 'object') return c;
  const carretas = Math.max(1, parseInt(c.carretas) || 1);
  let estados = Array.isArray(c.estados) ? c.estados.slice(0, carretas) : [];
  while (estados.length < carretas) estados.push('LLENA');
  estados = estados.map(e => ESTADOS_CARGA.indexOf(e) >= 0 ? e : 'LLENA');
  return Object.assign({}, c, { carretas, estados });
}

function _resumenEstadosCarretas(estados) {
  const r = { llenas: 0, medias: 0, vacias: 0, total: 0 };
  (estados || []).forEach(e => {
    if (e === 'LLENA') r.llenas++;
    else if (e === 'MEDIA') r.medias++;
    else if (e === 'VACIA') r.vacias++;
    r.total++;
  });
  return r;
}

// [v2.13.3] Sonidos + vibración háptica al togglear/agregar chip
let _ccAudio = null;
function _carretaSfx(estado) {
  try {
    if (!_ccAudio) _ccAudio = new (window.AudioContext || window.webkitAudioContext)();
    const ctx = _ccAudio, now = ctx.currentTime;
    const freq = estado === 'LLENA' ? 880 : estado === 'MEDIA' ? 600 : 380;
    const o = ctx.createOscillator(), g = ctx.createGain();
    o.type = 'sine'; o.frequency.setValueAtTime(freq, now);
    g.gain.setValueAtTime(0.0001, now);
    g.gain.exponentialRampToValueAtTime(0.10, now + 0.012);
    g.gain.exponentialRampToValueAtTime(0.001, now + 0.18);
    o.connect(g).connect(ctx.destination);
    o.start(now); o.stop(now + 0.20);
  } catch(_){}
  try {
    if (navigator.vibrate) navigator.vibrate(estado === 'VACIA' ? [20, 30, 20] : 15);
  } catch(_){}
}
function _carretaAddSfx() {
  try {
    if (!_ccAudio) _ccAudio = new (window.AudioContext || window.webkitAudioContext)();
    const ctx = _ccAudio, now = ctx.currentTime;
    [600, 900].forEach((f, i) => {
      const o = ctx.createOscillator(), g = ctx.createGain();
      const t0 = now + i * 0.05;
      o.type = 'triangle'; o.frequency.setValueAtTime(f, t0);
      g.gain.setValueAtTime(0.0001, t0);
      g.gain.exponentialRampToValueAtTime(0.10, t0 + 0.01);
      g.gain.exponentialRampToValueAtTime(0.001, t0 + 0.22);
      o.connect(g).connect(ctx.destination);
      o.start(t0); o.stop(t0 + 0.23);
    });
    if (navigator.vibrate) navigator.vibrate(10);
  } catch(_){}
}

// [v2.13.175] Lógica CENTRALIZADA del ciclo de carretas. Antes el avance
// LLENA→MEDIA→VACIA→eliminar (y alta/baja) estaba TRIPLICADO en los flujos de
// creación, edición y "día". Vive una sola vez acá; cada flujo aporta su array,
// render, persistencia y animación. Mutaciones in-place sobre el array dado.
const CarretaCiclo = {
  // Decide el siguiente paso de UNA carreta según su estado actual.
  // → { accion:'estado', estado } | { accion:'eliminar' }
  siguiente(estadoActual) {
    const a = ESTADOS_CARGA.indexOf(estadoActual) >= 0 ? estadoActual : 'LLENA';
    if (a === 'VACIA') return { accion: 'eliminar' };
    return { accion: 'estado', estado: ESTADOS_CARGA[(ESTADOS_CARGA.indexOf(a) + 1) % ESTADOS_CARGA.length] };
  },
  // Avanza el estado de una carreta. En 'eliminar' quita la carreta; si el
  // cargador queda en 0 lo quita del array. Devuelve descriptor o null.
  avanzar(arr, cargIdx, carretaIdx) {
    const c = arr && arr[cargIdx];
    if (!c || !Array.isArray(c.estados)) return null;
    // [v2.13.179] Guard de rango: si el chip clickeado ya no existe en el array
    // (DOM desfasado vs cache), abortar en vez de crear un "hole" en estados.
    if (carretaIdx < 0 || carretaIdx >= c.estados.length) return null;
    const r = this.siguiente(c.estados[carretaIdx] || 'LLENA');
    if (r.accion === 'eliminar') {
      c.estados.splice(carretaIdx, 1);
      c.carretas = c.estados.length;
      const cargadorEliminado = c.carretas === 0;
      if (cargadorEliminado) arr.splice(cargIdx, 1);
      return { accion: 'eliminar', cargadorEliminado: cargadorEliminado };
    }
    c.estados[carretaIdx] = r.estado;
    return { accion: 'estado', estado: r.estado };
  },
  // Agrega una carreta LLENA. Devuelve { accion:'add', carretaIdx, estado }.
  agregar(arr, cargIdx) {
    const c = arr && arr[cargIdx];
    if (!c) return null;
    if (!Array.isArray(c.estados)) c.estados = [];
    c.estados.push('LLENA');
    c.carretas = c.estados.length;
    return { accion: 'add', carretaIdx: c.estados.length - 1, estado: 'LLENA' };
  },
  // Quita la última carreta (compat con cambiarCarretas delta<0). Mínimo 1.
  quitarUltima(arr, cargIdx) {
    const c = arr && arr[cargIdx];
    if (!c) return null;
    c.carretas = Math.max(1, (c.carretas || 1) - 1);
    if (Array.isArray(c.estados)) c.estados = c.estados.slice(0, c.carretas);
    return { accion: 'trim' };
  }
};

// [v2.13.3] Resumen agregado de cargadores: solo conteo de carretas por estado.
function _resumenCargadoresDia(items) {
  let carretas = 0, llenas = 0, medias = 0, vacias = 0;
  (items || []).forEach(p => {
    let arr = [];
    try { arr = JSON.parse(p.cargadores || '[]'); } catch {}
    if (!Array.isArray(arr)) return;
    arr.forEach(c => {
      if (!c || typeof c !== 'object') return;
      const cn = _normalizarCargador(c);
      const r = _resumenEstadosCarretas(cn.estados);
      carretas += r.total;
      llenas   += r.llenas;
      medias   += r.medias;
      vacias   += r.vacias;
    });
  });
  return { carretas, llenas, medias, vacias };
}

// Filtra preingresos de un día (key=YYYY-MM-DD) usando la fecha LOCAL del
// cliente, igual que el pill del header. Centralizado para garantizar que
// pill y modal vean exactamente la misma data.
function _preingresosDeFecha(all, key) {
  return (all || []).filter(p => p.fecha && _diaPeru(p.fecha) === key);
}

// Para vistas que NO tienen los preingresos en mano (como Guías): mira
// la cache local y filtra preingresos cuya fecha cae en `key` (YYYY-MM-DD).
function _resumenCargadoresDiaPorFecha(key) {
  try {
    const all = OfflineManager.getPreingresosCache() || [];
    return _resumenCargadoresDia(_preingresosDeFecha(all, key));
  } catch { return { carretas: 0, llenas: 0, medias: 0, vacias: 0 }; }
}

// Construye el detalle agrupado por cargador (mismo shape que devolvía
// el endpoint server-side getCargadoresDelDia, pero 100% client-side
// usando la cache de preingresos. Garantiza consistencia con el pill
// del header y respuesta instantánea.
function _calcularCargadoresDelDia(key) {
  const all = OfflineManager.getPreingresosCache() || [];
  const items = _preingresosDeFecha(all, key);
  const provs = OfflineManager.getProveedoresCache() || [];
  const provMap = {};
  provs.forEach(p => { provMap[String(p.idProveedor)] = String(p.nombre || ''); });

  const byId = {};
  items.forEach(pi => {
    let arr = [];
    try { arr = JSON.parse(pi.cargadores || '[]'); } catch {}
    if (!Array.isArray(arr)) return;
    arr.forEach(c => {
      if (!c || typeof c !== 'object') return;
      const cn = _normalizarCargador(c);
      const id     = String(cn.id || cn.idPersonal || cn.nombre || '');
      const nombre = String(cn.nombre || cn.idPersonal || id || '');
      if (!id || !nombre) return;
      const carretas = cn.carretas;
      const r = _resumenEstadosCarretas(cn.estados);
      if (!byId[id]) byId[id] = {
        id, nombre,
        carretasTotal: 0, llenasTotal: 0, mediasTotal: 0, vaciasTotal: 0,
        preingresos: []
      };
      byId[id].carretasTotal += carretas;
      byId[id].llenasTotal   += r.llenas;
      byId[id].mediasTotal   += r.medias;
      byId[id].vaciasTotal   += r.vacias;
      byId[id].preingresos.push({
        idPreingreso: pi.idPreingreso,
        proveedor:    provMap[String(pi.idProveedor)] || pi.idProveedor || '',
        carretas,
        estados:      cn.estados,
        llenas: r.llenas, medias: r.medias, vacias: r.vacias,
        estado: String(pi.estado || '')
      });
    });
  });

  const cargadores = Object.values(byId).sort((a, b) => b.carretasTotal - a.carretasTotal);
  return {
    fecha:         key,
    cargadores,
    totalCarretas: cargadores.reduce((s, c) => s + c.carretasTotal, 0),
    totalLlenas:   cargadores.reduce((s, c) => s + c.llenasTotal,   0),
    totalMedias:   cargadores.reduce((s, c) => s + c.mediasTotal,   0),
    totalVacias:   cargadores.reduce((s, c) => s + c.vaciasTotal,   0),
    preingresos:   items.length
  };
}
// Normaliza un código de barras: elimina chars de control (GS1, null, etc.), trim, uppercase
function normCb(s) { return String(s || '').replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim().toUpperCase(); }

// ── Productos nuevos aprobados — últimos 3 días ─────────────
async function _cargarPNAprobados() {
  const cont = document.getElementById('dashPNAprob');
  const list = document.getElementById('pnAprobList');
  const cnt  = document.getElementById('pnAprobCount');
  if (!cont || !list) return;

  let pns = [];
  let errorMsg = '';

  // Estrategia robusta: intentar primero el endpoint nuevo (rápido, ya filtrado)
  // y caer al endpoint legacy getProductosNuevos si falla o no responde ok
  try {
    const resp = await API.getProductosNuevosRecientes({ dias: 3 });
    if (resp && resp.ok === false) throw new Error(resp.error || 'endpoint nuevo no ok');
    pns = Array.isArray(resp) ? resp : (resp && resp.data) || [];
    console.log('[PN aprobados] recientes via endpoint nuevo:', pns.length);
  } catch(e1) {
    console.warn('[PN aprobados] endpoint nuevo falló, intentando legacy:', e1 && e1.message);
    try {
      const resp = await API.getProductosNuevos({ estado: 'APROBADO' });
      if (resp && resp.ok === false) throw new Error(resp.error || 'legacy no ok');
      const todos = Array.isArray(resp) ? resp : (resp && resp.data) || [];
      // Filtrar por últimos 3 días en el frontend
      const corte = Date.now() - (3 * 86400000);
      pns = todos.filter(r => {
        if (!r.fechaAprobacion) return false;
        const t = new Date(r.fechaAprobacion).getTime();
        return t >= corte;
      }).map(r => {
        // Inferir tipoAprobacion desde observacion
        const obs = String(r.observacion || '').trim();
        const tipo = obs.indexOf('EQUIVALENTE') === 0 ? 'EQUIVALENTE' : 'NUEVO';
        return Object.assign({}, r, { tipoAprobacion: tipo });
      }).sort((a, b) => new Date(b.fechaAprobacion) - new Date(a.fechaAprobacion));
      console.log('[PN aprobados] via legacy:', pns.length);
    } catch(e2) {
      errorMsg = e2 && e2.message ? e2.message : 'Error desconocido';
      console.warn('[PN aprobados] legacy también falló:', errorMsg);
    }
  }

  // Si ambos endpoints fallaron → mostrar aviso
  if (errorMsg) {
    cont.classList.remove('hidden');
    cnt.textContent = '!';
    list.innerHTML = `<div class="text-xs text-amber-400 italic px-2">⚠ No se pudo cargar productos aprobados. Detalle: ${errorMsg}</div>`;
    return;
  }
  if (!pns.length) { cont.classList.add('hidden'); return; }

  // Toast: ¿hay aprobados desde la última visita?
  const lastSeenKey = 'wh_pnAprob_lastSeen';
  const lastSeen = parseInt(localStorage.getItem(lastSeenKey) || '0');
  const ahora = Date.now();
  const nuevosDesdeUltimaVez = pns.filter(p => {
    const t = new Date(p.fechaAprobacion).getTime();
    return t > lastSeen;
  });
  if (nuevosDesdeUltimaVez.length > 0 && lastSeen > 0 && typeof toast === 'function') {
    toast(`✅ ${nuevosDesdeUltimaVez.length} producto${nuevosDesdeUltimaVez.length !== 1 ? 's' : ''} nuevo${nuevosDesdeUltimaVez.length !== 1 ? 's' : ''} aprobado${nuevosDesdeUltimaVez.length !== 1 ? 's' : ''}`, 'ok');
  }
  localStorage.setItem(lastSeenKey, String(ahora));

  cont.classList.remove('hidden');
  cnt.textContent = pns.length;

  const usuarioActual = (window.AppSession && AppSession.getNombre && AppSession.getNombre()) || '';
  const nombreLow = String(usuarioActual).toLowerCase().trim();

  list.innerHTML = pns.map(p => {
    const tipo = String(p.tipoAprobacion || 'NUEVO').toUpperCase();
    const tipoCls = tipo === 'EQUIVALENTE' ? 'equiv' : 'nuevo';
    const tipoLabel = tipo === 'EQUIVALENTE' ? 'EQUIV' : 'NUEVO';
    const isMine = String(p.usuario || '').toLowerCase().trim().indexOf(nombreLow) >= 0 && nombreLow;
    const fechaApr = p.fechaAprobacion ? new Date(p.fechaAprobacion) : null;
    const dias = fechaApr ? Math.floor((Date.now() - fechaApr.getTime()) / 86400000) : 0;
    const whenTxt = dias === 0 ? 'Hoy' : dias === 1 ? 'Ayer' : `hace ${dias}d`;
    const recientCls = dias === 0 ? ' recent' : '';
    const fotoHtml = p.foto
      ? `<img src="${escHtml(p.foto)}" alt="">`
      : '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#475569" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="3"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21,15 16,10 5,21"/></svg>';
    return `
      <div class="pn-aprob-card${recientCls}">
        <div class="pn-aprob-foto">${fotoHtml}</div>
        <span class="pn-aprob-tipo ${tipoCls}">${tipoLabel}</span>
        <div class="pn-aprob-desc">${escHtml(p.descripcion || '—')}</div>
        <div class="pn-aprob-meta">▌${escHtml(p.codigoBarra || '—')}</div>
        <div class="pn-aprob-when">${whenTxt}</div>
        <div class="pn-aprob-by${isMine ? ' tu' : ''}">${isMine ? '✓ Tú' : escHtml(p.usuario || '—')}</div>
      </div>`;
  }).join('');
}

// Hora desde campo fecha de guía — solo si tiene componente de hora explícito
// [v2.13.209] Hora HH:mm de la guía SIEMPRE en TZ Lima (no la del dispositivo).
// Maneja los 3 formatos de `fecha` del cutover (ver _parseLocalDate): ISO-Z, ISO-con-espacio
// y yyyy-MM-dd. Antes usaba toLocaleTimeString (TZ del equipo) → mostraba 15:02 en vez de
// 10:02 fuera de Perú; y new Date(f) directo devolvía Invalid en Safari para el formato con
// espacio → caía a _horaDesdeId, que con los IDs nuevos 'G_L<ts><random>' extrae dígitos del
// random y produce hora basura. Ahora: si `fecha` trae hora, se formatea en Lima; solo si NO
// trae hora caemos al timestamp embebido en el ID (limpio, ver _horaDesdeId).
function _horaDesdeGuia(g) {
  const f = String(g.fecha || '');
  const tieneHora = f.includes('T') || (f.length > 10 && f.includes(':'));
  if (tieneHora) {
    const d = _parseLocalDate(f);   // Safari-safe (normaliza espacio→T)
    if (!isNaN(d)) {
      try {
        return new Intl.DateTimeFormat('es-PE', { timeZone: WH_TZ, hour: '2-digit', minute: '2-digit', hour12: false }).format(d);
      } catch (e) {
        // Fallback sin Intl: Perú = UTC-5 fijo. Resta 5h y lee HH:mm del instante en UTC.
        return new Date(d.getTime() - 5 * 3600000).toISOString().slice(11, 16);
      }
    }
  }
  return _horaDesdeId(g.idGuia);
}

// Parsea comentario → { comp: 'si'|'no'|null, compl: 'si'|'no'|null }
function _tagsFromComentario(comentario) {
  const s = String(comentario || '');
  const tags = { comp: null, compl: null };
  // [v2.13.173] Tolerante a "Sí"/"Si"/"sí"/"si" (con o sin acento) por si el
  // comentario viene de otra app o de edición manual. El builder sigue
  // escribiendo "Sí" para mantener consistencia visual.
  if (/comprobante:\s*s[ií]\b/i.test(s))    tags.comp  = 'si';
  else if (/comprobante:\s*no\b/i.test(s))  tags.comp  = 'no';
  if (/completo:\s*s[ií]\b/i.test(s))       tags.compl = 'si';
  else if (/completo:\s*no\b/i.test(s))     tags.compl = 'no';
  return tags;
}

// Extrae el texto libre quitando los prefijos de tags
function _textoLibreFromComentario(comentario) {
  return (comentario || '')
    .replace(/Comprobante:\s*(?:S[ií]|No)\s*\|?\s*/gi, '')
    .replace(/Completo:\s*(?:S[ií]|No)\s*\|?\s*/gi, '')
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
  // [v2.13.162 UX] Inactividad escalonada:
  //   - Primera vez de la sesión: 15 min (operador recién entra, se acomoda)
  //   - Después: 5 min (ya operando activamente, bloqueo estricto)
  // El flag _yaSeBloqueoUnaVez es de memoria — reset al reload de pestaña.
  const MIN_INACTIVIDAD         = 5;   // minutos — bloqueo subsecuente
  const MIN_INACTIVIDAD_PRIMERA = 15;  // minutos — primer bloqueo de la sesión
  let _yaSeBloqueoUnaVez = false;

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
      if (!res.ok && res.error === 'FUERA_DE_HORARIO') {
        // Operador/Envasador intentando entrar fuera del horario laboral
        try { SoundFX.buzzer(); } catch(e) {}
        Welcome.mostrarAlmacenCerrado(res.data || {});
        return;
      }
      if (!res.ok || res.offline) {
        document.getElementById('loginError').textContent = '❌ PIN incorrecto';
        try { SoundFX.buzzer(); } catch(e) {}
        setTimeout(() => { document.getElementById('loginError').textContent = ''; }, 2000);
        return;
      }
      sesionActual = res.data;
      _guardarSesion(sesionActual);
      document.getElementById('loginScreen').style.display = 'none';
      // [v2.13.162 UX] Reset del contador de bloqueos — login nuevo merece
      // los 15 min de gracia para acomodarse antes del primer bloqueo.
      _yaSeBloqueoUnaVez = false;
      _aplicarSesion();
      // Notificar a master+admin que ingresó este operador (push). Idempotente.
      try {
        const _mosUrlPush = window.WH_CONFIG?.mosGasUrl || '';
        if (_mosUrlPush) {
          const _devId = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
          const _nombreFull = ((res.data.nombre || '') + ' ' + (res.data.apellido || '')).trim();
          fetch(_mosUrlPush, {
            method: 'POST',
            body: JSON.stringify({
              action: 'notificarInicioSesionVendedor',
              nombre: _nombreFull,
              appOrigen: 'warehouseMos',
              deviceId: _devId
            })
          }).catch(() => {});
        }
      } catch(_) {}
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
    // [v2.13.162 UX] Reset del contador de bloqueos — login optimista offline
    // también merece los 15 min de gracia.
    _yaSeBloqueoUnaVez = false;
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
    if (sesionAnterior) _mostrarAvisoSesionAnterior(sesionAnterior);
    // Welcome screen estilo MosExpress (con stats + pre-carga)
    if (typeof Welcome !== 'undefined' && sesionActual) {
      Welcome.mostrar(sesionActual);
    } else {
      toast(`¡Hola ${sesionActual?.nombre || ''}! 👋`, 'ok', 2500);
    }
  }

  function _imprimirTicketBienvenida() {
    // Deshabilitado por solicitud del usuario — ahorrar papel.
    // El cuerpo se mantiene como no-op para no romper otros llamados existentes.
  }

  function _mostrarAvisoSesionAnterior(fecha) {
    // Toast de advertencia + modal breve (igual que MosExpress)
    toast(`⚠ Sesión anterior del ${fecha} no fue cerrada`, 'warn', 6000);
  }

  function _aplicarSesion() {
    window.WH_CONFIG.usuario   = sesionActual.nombre + ' ' + sesionActual.apellido;
    window.WH_CONFIG.idSesion  = sesionActual.idSesion;
    window.WH_CONFIG.idPersonal= sesionActual.idPersonal;
    window.WH_CONFIG.rol       = String(sesionActual.rol || '').toUpperCase();

    // Activar wake lock — pantalla activa mientras hay sesión
    _activarWakeLock();

    // Iniciar monitor de horario laboral (avisa 5min antes y cierra al límite)
    if (typeof Welcome !== 'undefined') Welcome.iniciarMonitorHorario(sesionActual.rol);

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

    // Mostrar acceso a Logs y Diagnóstico según rol
    const rolUp = String(sesionActual.rol || '').toUpperCase();
    const esAdmin  = (rolUp === 'MASTER' || rolUp === 'ADMINISTRADOR');
    const esMaster = (rolUp === 'MASTER');
    const sideLogs = document.getElementById('sideRowLogs');
    if (sideLogs) sideLogs.style.display = esAdmin ? '' : 'none';   // logs: admin + master
    const sideDiag = document.getElementById('sideRowDiag');
    if (sideDiag) sideDiag.style.display = esMaster ? '' : 'none';  // diagnóstico: solo MASTER
    // [v2.13.126 FIX] Tools (Calibrar impresora + Colas): visible para admin/master
    const sideTools = document.getElementById('sideRowTools');
    if (sideTools) sideTools.style.display = esAdmin ? '' : 'none';

    // [Pregúntale a tu almacén] FAB del chat IA — solo admin/master (es el dueño/gestor).
    // ADITIVO: si el módulo falla, no afecta el resto del login.
    try { if (typeof ChatAlmacen !== 'undefined') ChatAlmacen.init({ fab: esAdmin }); } catch (_) {}

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

    // Polling de bloqueo remoto desde MOS
    if (typeof BloqueoRemoto !== 'undefined') BloqueoRemoto.iniciar();

    // Caché admin (clave global + PINs admins) para validar offline
    if (typeof OfflineManager !== 'undefined' && OfflineManager.sincronizarAdminCache) {
      OfflineManager.sincronizarAdminCache();
    }

    // Push notifications — registrar token con nombre del operador
    setTimeout(_pushInitWH, 3000);

    // GPS tracking pasivo: cada 5 min mientras la app está visible
    setTimeout(() => _gpsRegistrarWH(false), 8000);
    if (_intervalGpsWH) clearInterval(_intervalGpsWH);
    _intervalGpsWH = setInterval(() => {
      if (document.visibilityState === 'visible') _gpsRegistrarWH(false);
    }, 5 * 60 * 1000);

    // [v2.13.37] Precarga de impresoras del ecosistema (admin/master).
    // Cuando el admin abra el modal de elegir impresora, ya estará cacheado
    // → modal abre INSTANT en vez de "⏳ Cargando impresoras..." por 2-3s.
    if (esAdmin && typeof API !== 'undefined' && API.getImpresorasEcosistema) {
      setTimeout(() => {
        API.getImpresorasEcosistema().then(r => {
          const arr = (r && r.ok && r.data) || [];
          try {
            window._whPrintersCache = { ts: Date.now(), data: arr };
            localStorage.setItem('wh_printers_cache', JSON.stringify(window._whPrintersCache));
          } catch(_){}
        }).catch(() => {});
      }, 1500);
    }
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
  let _timerListenersDone = false;
  function _iniciarTimerBloqueo() {
    _resetTimerBloqueo();
    if (_timerListenersDone) return;
    _timerListenersDone = true;
    ['touchstart','click','keydown','scroll'].forEach(ev =>
      document.addEventListener(ev, _resetTimerBloqueo, { passive: true })
    );
  }

  function _resetTimerBloqueo() {
    clearTimeout(lockTimer);
    // [v2.13.162 UX] Si admin sobreescribió, prioridad a su config. Sino:
    //   - Primer bloqueo de la sesión: 15 min (más relajado)
    //   - Bloqueos siguientes: 5 min (más estricto)
    const override = parseInt(localStorage.getItem('wh_min_inactividad'));
    const min = (override && override > 0)
      ? override
      : (_yaSeBloqueoUnaVez ? MIN_INACTIVIDAD : MIN_INACTIVIDAD_PRIMERA);
    lockTimer = setTimeout(() => bloquear(), min * 60 * 1000);
  }

  function bloquear() {
    if (!sesionActual) return;
    // [v2.13.162 UX] Marcar que ya se bloqueó al menos una vez — próximos
    // bloqueos usarán 5 min en lugar de 15 min.
    _yaSeBloqueoUnaVez = true;
    lockPinBuffer = '';
    _actualizarPuntos('lpin', 0);
    document.getElementById('lockError').textContent = '';

    const av = document.getElementById('lockAvatar');
    av.textContent = sesionActual.nombre[0] + sesionActual.apellido[0];
    av.style.background = sesionActual.color;
    document.getElementById('lockNombre').textContent = sesionActual.nombre + ' ' + sesionActual.apellido;

    localStorage.setItem('wh_lock_inicio', Date.now());

    // Init reloj, frase rotante y partículas en el lock screen
    _initLockEnhancements();

    document.getElementById('lockScreen').style.display = 'flex';

    clearInterval(lockInterval);
    lockInterval = setInterval(() => {
      const seg = Math.floor((Date.now() - parseInt(localStorage.getItem('wh_lock_inicio'))) / 1000);
      const m = Math.floor(seg / 60), s = seg % 60;
      document.getElementById('lockTiempo').textContent =
        `Bloqueado hace ${m > 0 ? m + 'm ' : ''}${s}s`;

      // Reloj central HH:MM:SS
      const n = new Date();
      const hm = String(n.getHours()).padStart(2, '0') + ':' + String(n.getMinutes()).padStart(2, '0');
      const ss = String(n.getSeconds()).padStart(2, '0');
      const elHM = document.getElementById('lockClockHM');
      const elSec = document.getElementById('lockClockSec');
      if (elHM) elHM.textContent = hm;
      if (elSec) elSec.textContent = ':' + ss;
    }, 1000);
  }

  // ───────────────────────────────────────────────
  //  Salvapantalla / Lock screen enhancements
  // ───────────────────────────────────────────────
  const _frasesAlmacen = [
    'El orden de hoy es la productividad de mañana.',
    'Cada caja bien etiquetada ahorra horas al equipo.',
    'Un almacén limpio refleja un equipo orgulloso.',
    'La precisión en el inventario es respeto al cliente.',
    'Trabajar bien hoy es no tener problemas mañana.',
    'Tu atención al detalle se nota en todo el negocio.',
    'Cada guía bien revisada protege al equipo entero.',
    'Lo que se cuenta bien, no se pierde nunca.',
    'El silencio de un almacén ordenado vale oro.',
    'Pequeñas decisiones precisas, grandes resultados.',
    'La constancia ordenada vence al esfuerzo desordenado.',
    'Eres parte del corazón de Inversiones MOS.',
    'Cuando todo está en su lugar, el equipo respira.',
    'La calidad empieza en el almacén, no en la venta.',
    'Tu trabajo es invisible cuando todo sale bien — eso es lo valioso.',
    'Mediste, contaste, etiquetaste: hiciste posible la venta.',
    'El cliente nunca te ve, pero vive de tu precisión.',
    'Una hora ordenada vale tres horas corriendo.',
    'Cada producto bien guardado es un cliente sin reclamos.',
    'Hoy no se trata de hacer más, sino de hacer mejor.'
  ];
  let _lockQuoteInterval = null;
  let _wakeLockSentinel = null;

  function _initLockEnhancements() {
    // Frase inicial aleatoria
    const elQ = document.getElementById('lockQuote');
    if (elQ) {
      elQ.textContent = '"' + _frasesAlmacen[Math.floor(Math.random() * _frasesAlmacen.length)] + '"';
    }
    // Rotar frase cada 8s con fade
    if (_lockQuoteInterval) clearInterval(_lockQuoteInterval);
    _lockQuoteInterval = setInterval(() => {
      if (!elQ) return;
      elQ.classList.add('fading');
      setTimeout(() => {
        elQ.textContent = '"' + _frasesAlmacen[Math.floor(Math.random() * _frasesAlmacen.length)] + '"';
        elQ.classList.remove('fading');
      }, 500);
    }, 8000);
    // Posicionar partículas con coords aleatorias
    for (let i = 0; i < 6; i++) {
      const p = document.getElementById('lockPart' + i);
      if (!p) continue;
      p.style.left           = (Math.random() * 100) + '%';
      p.style.setProperty('--p-dx', ((Math.random() - 0.5) * 100) + 'px');
      p.style.animationDelay = (Math.random() * 18) + 's';
      p.style.animationDuration = (14 + Math.random() * 8) + 's';
    }
    // Indicador wake lock
    const ind = document.getElementById('lockWakeLockInd');
    if (ind) ind.style.display = _wakeLockSentinel ? 'block' : 'none';
  }

  // ───────────────────────────────────────────────
  //  Wake Lock — pantalla activa con sesión
  // ───────────────────────────────────────────────
  // [v2.13.32] Chip wake-lock como TOAST efímero (3s). Antes era permanente
  // arriba a la derecha y tapaba iconos del topbar (usuario, cesta merma).
  // La función sigue ACTIVA en background — el chip es solo un indicador
  // visual al activarse. Mientras tengas sesión, la pantalla NO se apaga.
  function _whWakeRenderChip(on) {
    let chip = document.getElementById('whWakeChip');
    if (!on) { if (chip) chip.remove(); return; }
    // Evitar duplicar si el toast ya se está mostrando
    if (chip) return;
    chip = document.createElement('div');
    chip.id = 'whWakeChip';
    chip.style.cssText = 'position:fixed;top:14px;left:50%;transform:translateX(-50%) translateY(-12px);z-index:9998;display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:9999px;background:linear-gradient(135deg,rgba(251,191,36,0.95),rgba(245,158,11,0.95));border:1px solid rgba(251,191,36,0.6);color:#451a03;font-size:12px;font-weight:800;backdrop-filter:blur(6px);box-shadow:0 8px 24px rgba(0,0,0,0.4);pointer-events:none;font-family:system-ui,sans-serif;opacity:0;transition:opacity .35s ease,transform .35s cubic-bezier(.34,1.56,.64,1)';
    chip.innerHTML = '<span>🔆</span><span>Pantalla activa</span>';
    document.body.appendChild(chip);
    // Entrar
    requestAnimationFrame(() => {
      chip.style.opacity = '1';
      chip.style.transform = 'translateX(-50%) translateY(0)';
    });
    // Salir tras 3s
    setTimeout(() => {
      chip.style.opacity = '0';
      chip.style.transform = 'translateX(-50%) translateY(-12px)';
      setTimeout(() => { try { chip.remove(); } catch(_){} }, 380);
    }, 3000);
  }
  // [v2.13.91] Parámetros:
  //   forzar=true   → ignora preferencia 'wh_wakelock=0' (espía lo necesita sí o sí)
  //   silencioso=true → no muestra el chip "🔆 Pantalla activa" (espía no delata al target)
  async function _activarWakeLock(forzar, silencioso) {
    if (!('wakeLock' in navigator)) return;
    if (_wakeLockSentinel) return;
    if (!forzar && localStorage.getItem('wh_wakelock') === '0') return;
    try {
      _wakeLockSentinel = await navigator.wakeLock.request('screen');
      if (!silencioso) _whWakeRenderChip(true);
      _wakeLockSentinel.addEventListener('release', () => {
        _wakeLockSentinel = null;
        // _whWakeRenderChip(false) es idempotente: no-op si no hay chip
        _whWakeRenderChip(false);
      });
    } catch(e) { /* dispositivo no soporta o sin user gesture */ }
  }
  async function _liberarWakeLock() {
    if (_wakeLockSentinel) {
      try { await _wakeLockSentinel.release(); } catch(e) {}
      _wakeLockSentinel = null;
    }
    _whWakeRenderChip(false);
  }
  // Re-activar al volver al foreground (sesión normal o espía activa)
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'visible') return;
    if (sesionActual) _activarWakeLock();
    else if (window._espiaCliWH) _activarWakeLock(true, true); // espía: forzar + silencioso
  });

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
      if (_lockQuoteInterval) { clearInterval(_lockQuoteInterval); _lockQuoteInterval = null; }
      document.getElementById('lockScreen').style.display = 'none';
      _resetTimerBloqueo();
      _activarWakeLock(); // re-activar por si el SO lo libero al bloquear
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
    _liberarWakeLock();
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

  // ════════════════════════════════════════════════════════════
  // VERIFICACIÓN DE DISPOSITIVO (antes del login con PIN)
  // Igual que ME: el master debe aprobar el dispositivo antes de que pueda
  // operar. Si no está aprobado, pantalla candado + botón "📨 Solicitar acceso".
  // ════════════════════════════════════════════════════════════
  let _verifPollTimer = null;
  let _verifEstado = 'CARGANDO';

  // Verifica el estado del dispositivo en MOS y actualiza la pantalla candado.
  // Retorna 'ACTIVO' | 'INACTIVO' | 'PENDIENTE' | 'NO_REGISTRADO' | 'ERROR_RED'.
  // El init de la app espera este resultado: solo continua al login si es 'ACTIVO'.
  //
  // Optimizaciones (igual patrón que ME, ver MosExpress/index.html:3715):
  //   1) Cache local 1h en localStorage — si verificó hace <1h, autoriza al
  //      instante sin tocar el GAS. Llamada GAS solo en background "silenciosa"
  //      para refrescar el cache.
  //   2) Timeout AbortController 6s — evita quedar bloqueado en cold start del
  //      GAS (que puede tardar 5+s la primera llamada).
  //   3) Endpoint registrarSesionDispositivo (POST) en vez de
  //      consultarEstadoDispositivo: en una sola llamada registra el dispositivo
  //      como PENDIENTE_APROBACION si es nuevo Y devuelve el estado. Antes
  //      requería 2 round-trips.
  // [v2.13.164 cleanup] _AUTH_CACHE_KEY/_ID_KEY se mantienen escritos por
  // DeviceAuth.onAuth() en index.html (compat con código que pudiera leerlos).
  // No los leemos desde aquí — el cache del módulo es la fuente de verdad.
  const _AUTH_CACHE_KEY    = 'wh_device_auth_date';
  const _AUTH_CACHE_ID_KEY = 'wh_device_auth_id';

  async function _verificarDispositivoWH() {
    // [v2.13.155] DELEGAR a DeviceAuth (módulo compartido).
    // Antes esta función hacía el flow completo de verificación con su propio
    // overlay (verifDispScreen). Ahora el módulo lo hace global y uniforme.
    // Si DeviceAuth dice ACTIVO, retornamos ACTIVO. Si no, el módulo ya muestra
    // su overlay (con da-pre-block bloqueando el resto).
    //
    // [v2.13.161 BUG CRÍTICO FIX] ESPERAR a que DeviceAuth resuelva antes de
    // retornar. Antes leíamos el estado síncrono inmediato — si DeviceAuth aún
    // estaba en VERIFICANDO (fetch al server pendiente 200-500ms), retornábamos
    // VERIFICANDO y el caller hacía return temprano. Resultado: app quedaba
    // visible SIN llegar a llamar Session.init() → operador entraba como
    // "usuario fantasma" sin login, con WH_CONFIG.usuario = 'operador' default.
    // Síntoma reportado por el usuario 2026-06-04: WH permite ingresar
    // sin pedir login operador.
    if (!window.DeviceAuth) {
      _verifEstado = 'ERROR_RED';
      return 'ERROR_RED';
    }
    // Polling cada 100ms hasta que DeviceAuth tenga veredicto terminal.
    // Safety: 15s máximo (en cualquier caso ya cubre cache válido + fetch).
    var maxIntentos = 150;  // 150 × 100ms = 15s
    for (var i = 0; i < maxIntentos; i++) {
      var est = DeviceAuth.estado();
      var e = est && est.estado;
      if (e === 'ACTIVO') {
        _verifEstado = 'ACTIVO';
        _ocultarPantallaVerif();
        return 'ACTIVO';
      }
      // Estados terminales NO autorizados — el módulo ya muestra su overlay
      if (e === 'PENDIENTE_APROBACION' || e === 'INACTIVO' || e === 'SUSPENDIDO' ||
          e === 'NO_REGISTRADO' || e === 'SIN_VERIFICAR') {
        _verifEstado = e;
        return e;
      }
      // Estados intermedios (INIT, VERIFICANDO) → esperar
      await new Promise(function(r) { setTimeout(r, 100); });
    }
    // Timeout — el módulo nunca resolvió. Tratar como SIN_VERIFICAR (fail-CLOSED).
    _verifEstado = 'SIN_VERIFICAR';
    return 'SIN_VERIFICAR';
  }

  // [v2.13.164 cleanup] _verificarDispositivoWH_LEGACY y _verificarDispositivoSilencioso
  // eliminados (~170 LOC). DeviceAuth maneja todo el flow desde v2.13.155.

  function _ocultarPantallaVerif() {
    const el = document.getElementById('verifDispScreen');
    if (el) el.remove();
  }

  function _mostrarPantallaVerif(tipo, nombre) {
    let el = document.getElementById('verifDispScreen');
    if (!el) {
      el = document.createElement('div');
      el.id = 'verifDispScreen';
      el.style.cssText = 'position:fixed;inset:0;z-index:9998;background:linear-gradient(135deg,#0c1426,#1e293b);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;color:#e2e8f0;font-family:system-ui,-apple-system,sans-serif;';
      document.body.appendChild(el);
    }
    const ua = (navigator.userAgent || '').substring(0, 200);
    let html = '';
    if (tipo === 'no_registrado') {
      html = `
        <div style="text-align:center;max-width:400px;">
          <div style="font-size:64px;margin-bottom:16px;">🔒</div>
          <h2 style="font-size:22px;font-weight:800;margin-bottom:8px;color:#f8fafc;">Dispositivo no autorizado</h2>
          <p style="font-size:14px;color:#94a3b8;margin-bottom:20px;">Este dispositivo no tiene permiso para usar warehouseMos.</p>
          <button id="btnSolicitarAcceso" style="display:block;width:100%;background:linear-gradient(135deg,#f59e0b,#d97706);color:#0b1220;border:none;padding:14px 24px;border-radius:12px;font-size:14px;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(245,158,11,0.4);margin-bottom:10px;">📨 Solicitar acceso al admin (remoto)</button>
          <div style="font-size:11px;color:#64748b;margin:14px 0 10px;letter-spacing:.05em">─ o si está un admin contigo ─</div>
          <button id="btnActivarInSitu" style="display:block;width:100%;background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;padding:14px 24px;border-radius:12px;font-size:14px;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(16,185,129,.5);">🔑 Activar in-situ (admin presente)</button>
          <p style="font-size:11px;color:#64748b;margin-top:24px;font-family:'SF Mono',Menlo,monospace;">UUID: ${(typeof window._getDeviceIdWH==='function'?window._getDeviceIdWH():'').substring(0,12)}...</p>
        </div>`;
    } else if (tipo === 'pendiente') {
      // [v2.10.3] Agregar también la opción in-situ acá. Antes solo aparecía
      // en el estado 'no_registrado', pero registrarSesionDispositivo crea el
      // row como PENDIENTE en la 1era llamada → el usuario nunca veía la
      // pantalla 'no_registrado' y no tenía cómo activar in-situ.
      html = `
        <div style="text-align:center;max-width:400px;">
          <div style="font-size:64px;margin-bottom:16px;animation:wpulse 1.6s ease-in-out infinite;">⌛</div>
          <h2 style="font-size:22px;font-weight:800;margin-bottom:8px;color:#fbbf24;">Esperando aprobación</h2>
          <p style="font-size:14px;color:#94a3b8;margin-bottom:8px;">${nombre || 'Tu dispositivo'} está pendiente de aprobación por el administrador.</p>
          <p style="font-size:12px;color:#64748b;margin-bottom:16px;">Reintenta automáticamente cada 15s.</p>
          <div style="font-size:11px;color:#475569;background:rgba(255,255,255,0.05);padding:8px 12px;border-radius:8px;display:inline-block;margin-bottom:18px;">🔄 Verificando estado...</div>
          <div style="font-size:11px;color:#64748b;margin:6px 0 10px;letter-spacing:.05em">─ o si está un admin contigo ─</div>
          <button id="btnActivarInSitu" style="display:block;width:100%;background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;padding:14px 24px;border-radius:12px;font-size:14px;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(16,185,129,.5);">🔑 Activar in-situ (admin presente)</button>
          <p style="font-size:11px;color:#64748b;margin-top:18px;font-family:'SF Mono',Menlo,monospace;">UUID: ${(typeof window._getDeviceIdWH==='function'?window._getDeviceIdWH():'').substring(0,12)}...</p>
        </div>
        <style>@keyframes wpulse{0%,100%{transform:scale(1);opacity:1}50%{transform:scale(1.1);opacity:0.7}}</style>`;
    } else if (tipo === 'inactivo') {
      html = `
        <div style="text-align:center;max-width:400px;">
          <div style="font-size:64px;margin-bottom:16px;">🚫</div>
          <h2 style="font-size:22px;font-weight:800;margin-bottom:8px;color:#f87171;">Dispositivo desactivado</h2>
          <p style="font-size:14px;color:#94a3b8;margin-bottom:24px;">El administrador desactivó este dispositivo. Contactá al administrador.</p>
        </div>`;
    } else if (tipo === 'error_red') {
      html = `
        <div style="text-align:center;max-width:400px;">
          <div style="font-size:64px;margin-bottom:16px;">📡</div>
          <h2 style="font-size:22px;font-weight:800;margin-bottom:8px;color:#fbbf24;">Sin conexión con MOS</h2>
          <p style="font-size:14px;color:#94a3b8;margin-bottom:8px;">No se pudo verificar el dispositivo. Revisa tu conexión.</p>
          ${nombre ? `<p style="font-size:11px;color:#475569;margin-bottom:24px;font-family:monospace;">${String(nombre).substring(0, 100)}</p>` : '<div style="height:24px"></div>'}
          <button id="btnReintentarVerif" style="background:linear-gradient(135deg,#0ea5e9,#0284c7);color:#fff;border:none;padding:14px 24px;border-radius:12px;font-size:15px;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(14,165,233,0.4);">🔄 Reintentar</button>
        </div>`;
    } else if (tipo === 'cargando') {
      html = `
        <div style="text-align:center;max-width:400px;">
          <div style="font-size:64px;margin-bottom:16px;">🔄</div>
          <h2 style="font-size:20px;font-weight:800;margin-bottom:8px;color:#f8fafc;">Verificando dispositivo</h2>
          <p style="font-size:13px;color:#94a3b8;">Conectando con MOS...</p>
        </div>`;
    }
    el.innerHTML = html;
    const btn = document.getElementById('btnSolicitarAcceso');
    if (btn) btn.onclick = _solicitarAccesoDispositivo;
    const btnInSitu = document.getElementById('btnActivarInSitu');
    if (btnInSitu) btnInSitu.onclick = _abrirModalActivarInSitu;
    const btnRetry = document.getElementById('btnReintentarVerif');
    if (btnRetry) btnRetry.onclick = async () => {
      btnRetry.disabled = true;
      btnRetry.textContent = '⌛ Reintentando...';
      const r = await _verificarDispositivoWH();
      if (r === 'ACTIVO') {
        // Continuar con el flujo normal de la app después de salir del candado
        if (typeof Session !== 'undefined' && Session.init) Session.init();
      }
    };
  }

  // ════════════════════════════════════════════════════════════════════
  // ACTIVACIÓN IN-SITU — admin presente con clave 8 dígitos.
  // El admin escribe los 8 dig (4 globales + 4 PIN personal) y aprueba al
  // instante el dispositivo. Mismo flujo que anulaciones de venta o
  // conversión de NV→CPE en otras apps. Requiere endpoint MOS
  // aprobarDispositivoEnSitu (gas/Config.gs).
  // ════════════════════════════════════════════════════════════════════
  function _abrirModalActivarInSitu() {
    const ua = (navigator.userAgent || '').substring(0, 80);
    const esMobile = /iPhone|Android|Mobile|iPad/i.test(navigator.userAgent || '');
    const prefix   = esMobile ? 'Mobile' : 'Desktop';
    const ahora    = new Date();
    const hh       = String(ahora.getHours()).padStart(2, '0');
    const mm       = String(ahora.getMinutes()).padStart(2, '0');
    const nomDef   = prefix + ' WH ' + hh + ':' + mm;

    const modal = document.createElement('div');
    modal.id = 'modalActivarInSitu';
    modal.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(2,6,23,.85);backdrop-filter:blur(8px);display:flex;align-items:center;justify-content:center;padding:20px;';
    modal.innerHTML = `
      <div style="max-width:420px;width:100%;border-radius:20px;padding:22px;background:linear-gradient(135deg,#1e293b,#0f172a);border:1.5px solid rgba(16,185,129,.5);box-shadow:0 30px 60px -10px rgba(0,0,0,.65);">
        <div style="font-size:36px;margin-bottom:8px">🔑</div>
        <p style="font-size:1.05em;font-weight:900;color:#f1f5f9;margin-bottom:4px">Activar dispositivo</p>
        <p style="font-size:.78em;color:#94a3b8;margin-bottom:16px">Un administrador autoriza este dispositivo en este momento con su clave de 8 dígitos.</p>
        <label style="display:block;font-size:.7em;font-weight:700;color:#94a3b8;margin-bottom:4px;letter-spacing:.04em">NOMBRE DEL DISPOSITIVO</label>
        <input id="actInsNombre" type="text" value="${escAttr(nomDef)}"
               style="width:100%;padding:12px;border-radius:10px;background:rgba(15,23,42,.7);border:1.5px solid rgba(51,65,85,.6);color:#f1f5f9;font-size:.92em;outline:none;margin-bottom:14px"
               placeholder="ej: Tablet WH almacén central">
        <label style="display:block;font-size:.7em;font-weight:700;color:#94a3b8;margin-bottom:4px;letter-spacing:.04em">CLAVE ADMIN (8 DÍGITOS)</label>
        <input id="actInsClave" type="password" inputmode="numeric" maxlength="8" pattern="[0-9]{8}"
               style="width:100%;padding:14px;border-radius:10px;background:rgba(15,23,42,.7);border:1.5px solid rgba(16,185,129,.4);color:#f1f5f9;font-size:1.3em;font-weight:900;letter-spacing:.4em;text-align:center;outline:none;margin-bottom:6px"
               placeholder="••••••••" autocomplete="off">
        <p style="font-size:.62em;color:#64748b;margin-bottom:14px;letter-spacing:.04em">Clave global (4) + PIN personal del admin (4)</p>
        <p id="actInsErr" style="font-size:.75em;color:#fca5a5;margin-bottom:10px;min-height:18px;text-align:center"></p>
        <div style="display:flex;gap:8px">
          <button id="actInsCancel" style="flex:1;padding:13px;border-radius:11px;border:1px solid rgba(71,85,105,.55);background:rgba(71,85,105,.3);color:#cbd5e1;font-size:.85em;font-weight:800;cursor:pointer">Cancelar</button>
          <button id="actInsOk" style="flex:1;padding:13px;border-radius:11px;border:none;background:linear-gradient(135deg,#10b981,#059669);color:#fff;font-size:.85em;font-weight:900;cursor:pointer;box-shadow:0 6px 16px -3px rgba(16,185,129,.5)">✓ ACTIVAR</button>
        </div>
        <p style="font-size:.6em;color:#64748b;margin-top:14px;text-align:center;font-family:monospace">UUID: ${(typeof window._getDeviceIdWH==='function'?window._getDeviceIdWH():'').substring(0, 16)}...</p>
      </div>`;
    document.body.appendChild(modal);
    setTimeout(() => document.getElementById('actInsClave')?.focus(), 100);

    document.getElementById('actInsCancel').onclick = () => modal.remove();
    document.getElementById('actInsOk').onclick = async () => {
      const nombre = document.getElementById('actInsNombre').value.trim() || nomDef;
      const clave  = document.getElementById('actInsClave').value.trim();
      const errEl  = document.getElementById('actInsErr');
      const okBtn  = document.getElementById('actInsOk');
      errEl.textContent = '';
      if (!/^\d{8}$/.test(clave)) { errEl.textContent = 'La clave debe ser 8 dígitos numéricos'; return; }
      okBtn.disabled = true; okBtn.textContent = '⌛ Validando...';
      const mosUrl = window.WH_CONFIG?.mosGasUrl || '';
      const devId  = window._getDeviceIdWH ? window._getDeviceIdWH() : '';
      try {
        const res = await fetch(mosUrl, {
          method: 'POST',
          body: JSON.stringify({
            action: 'aprobarDispositivoEnSitu',
            deviceId: devId,
            nombreEquipo: nombre,
            app: 'warehouseMos',
            userAgent: ua,
            claveAdmin: clave
          })
        });
        const j = await res.json();
        if (!j.ok) { errEl.textContent = j.error || 'Error de conexión'; okBtn.disabled = false; okBtn.textContent = '✓ ACTIVAR'; return; }
        if (!j.data?.autorizado) {
          errEl.textContent = j.data?.error || 'Clave incorrecta';
          okBtn.disabled = false; okBtn.textContent = '✓ ACTIVAR';
          try { SoundFX?.warn?.(); } catch(_){}
          vibrate([60, 30, 60]);
          return;
        }
        // ✓ Aprobado
        try { SoundFX?.done?.(); } catch(_){}
        vibrate([30, 30, 60]);
        try { localStorage.setItem('wh_perms_check_pending', '1'); } catch(_){}
        try { window._pedirPersistentStorageWH && window._pedirPersistentStorageWH(); } catch(_){}
        modal.remove();
        // Quitar pantalla bloqueo y arrancar la app
        const verif = document.getElementById('verifDispScreen');
        if (verif) verif.remove();
        toast(`✓ Aprobado por ${j.data.aprobadoPor || 'admin'} · iniciando...`, 'ok', 3000);
        setTimeout(() => {
          if (typeof Session !== 'undefined' && Session.init) Session.init();
          // Disparar wizard de permisos automáticamente
          if (window.WhPerms?.auto) window.WhPerms.auto();
        }, 700);
      } catch (e) {
        errEl.textContent = 'Sin conexión con MOS';
        okBtn.disabled = false; okBtn.textContent = '✓ ACTIVAR';
      }
    };
  }

  async function _solicitarAccesoDispositivo() {
    const btn = document.getElementById('btnSolicitarAcceso');
    if (btn) { btn.disabled = true; btn.textContent = '⌛ Enviando...'; }
    const mosUrl = window.WH_CONFIG?.mosGasUrl || '';
    const devId = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
    const ua = (navigator.userAgent || '').substring(0, 200);
    try {
      await fetch(mosUrl, {
        method: 'POST',
        body: JSON.stringify({
          action: 'registrarSesionDispositivo',
          ID_Dispositivo: devId,
          app: 'warehouseMos',
          userAgent: ua
        })
      });
      // Pasar a pantalla "esperando aprobación"
      _verifEstado = 'PENDIENTE';
      _mostrarPantallaVerif('pendiente');
      if (!_verifPollTimer) _verifPollTimer = setInterval(_verificarDispositivoWH, 15 * 1000);
    } catch(e) {
      if (btn) { btn.disabled = false; btn.textContent = '📨 Solicitar acceso'; }
      _whAlert('Error: ' + (e.message || 'sin conexión'), { titulo: 'Error solicitar acceso' });
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

  async function cerrarSesionDesdeLock() {
    if (!await _whConfirm('¿Cerrar sesión?\n\nSe perderá el reporte de turno.', { warning: true, titulo: 'Cerrar sesión', okText: 'Cerrar sesión' })) return;
    clearTimeout(lockTimer);
    clearInterval(lockInterval);
    if (_lockQuoteInterval) { clearInterval(_lockQuoteInterval); _lockQuoteInterval = null; }
    _liberarWakeLock();
    _limpiarSesion();
    sesionActual = null;
    document.getElementById('lockScreen').style.display = 'none';
    _ocultarApp();
    mostrarLogin();
  }

  function _guardarSesion(ses) {
    // [v2.13.68] Manejo defensivo de QuotaExceededError. localStorage en
    // PCs muy usadas se llena con cachés viejos (wh_productos, wh_stock,
    // wh_guia_detalle, etc.). Si no podemos guardar la sesión, hacemos
    // limpieza emergency de los caches grandes y reintentamos.
    var payload = JSON.stringify({
      ...ses,
      fechaDia:      _hoy(),
      fechaGuardado: new Date().toISOString()
    });
    try {
      localStorage.setItem('wh_sesion', payload);
    } catch(e) {
      if (e.name === 'QuotaExceededError' || /quota/i.test(e.message)) {
        console.warn('[Sesion] localStorage lleno, limpiando caches grandes...');
        // Borrar los caches MÁS grandes en orden (los que el offline precarga)
        ['wh_guia_detalle', 'wh_preingresos', 'wh_productos', 'wh_stock',
         'wh_proveedores', 'wh_ajustes', 'wh_auditorias_c', 'wh_ubicaciones'].forEach(function(k){
          try { localStorage.removeItem(k); } catch(_){}
        });
        // Reintento
        try {
          localStorage.setItem('wh_sesion', payload);
          console.log('[Sesion] ✓ guardada tras cleanup emergency');
          if (typeof toast === 'function') toast('🧹 Cache limpiada · sesión OK', 'info', 4000);
        } catch(e2) {
          console.error('[Sesion] localStorage SIGUE lleno tras cleanup:', e2.message);
          if (typeof toast === 'function') toast('⚠ Storage lleno · usa DevTools→Application→Clear Storage', 'error', 10000);
          throw e2;
        }
      } else { throw e; }
    }
  }

  function _cargarSesion() {
    try { return JSON.parse(localStorage.getItem('wh_sesion')); }
    catch { return null; }
  }

  function _limpiarSesion() {
    localStorage.removeItem('wh_sesion');
    if (typeof BloqueoRemoto !== 'undefined') BloqueoRemoto.detener();
  }

  function _mostrarApp() {
    document.getElementById('topBar').style.display = '';
    document.querySelector('main').style.display = '';
    document.querySelector('nav').style.display = '';
    // Nav v3: bind interacciones + posicionar pill al activo
    setTimeout(() => {
      if (typeof App !== 'undefined') {
        if (App._bindNavV3) App._bindNavV3();
        if (App._moverNavPill) App._moverNavPill();
      }
    }, 150);
  }

  function _ocultarApp() {
    document.getElementById('topBar').style.display = 'none';
    document.querySelector('main').style.display = 'none';
    document.querySelector('nav').style.display = 'none';
  }

  function getSesion() { return sesionActual; }

  // ── DEVICE ID estable para tracking GPS y audio ──────────
  // [v2.13.160 ESTANDARIZACIÓN] Unificar deviceId con DeviceAuth (single
  // source of truth). Aunque WH y el módulo usan la MISMA clave
  // 'wh_device_id' por coincidencia, conviene delegar al módulo para evitar
  // drift futuro y race conditions (si app.js corriera antes de DeviceAuth).
  function _getDeviceIdWH() {
    // Source of truth: módulo DeviceAuth si está cargado
    if (window.DeviceAuth && typeof window.DeviceAuth.deviceId === 'function') {
      const daId = window.DeviceAuth.deviceId();
      if (daId) return daId;
    }
    // Fallback: lógica local (idéntica clave que el módulo)
    let id = localStorage.getItem('wh_device_id');
    if (!id) {
      try { id = crypto.randomUUID ? crypto.randomUUID() : ('WH' + Date.now() + Math.random().toString(36).slice(2)); }
      catch(_) { id = 'WH' + Date.now() + Math.random().toString(36).slice(2); }
      localStorage.setItem('wh_device_id', id);
    }
    return id;
  }
  window._getDeviceIdWH = _getDeviceIdWH;

  // ── AUDIO REMOTO + GPS — escucha comandos del SW y registra ubicación ──
  let _audioRecorder = null;
  let _audioStream = null;
  let _audioSesionId = null;
  let _audioChunkIdx = 0;
  let _audioAutoStopTimer = null;
  let _intervalGpsWH = null;

  const _audioBlobToBase64WH = (blob) => new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onloadend = () => {
      const dataUrl = r.result || '';
      const i = dataUrl.indexOf(',');
      resolve(i >= 0 ? dataUrl.substring(i + 1) : dataUrl);
    };
    r.onerror = reject;
    r.readAsDataURL(blob);
  });

  async function _audioRemotoIniciarWH(sesionId, duracionMaxSeg) {
    const mosUrl = window.WH_CONFIG?.mosGasUrl;
    if (!mosUrl) return;
    if (_audioRecorder) await _audioRemotoDetenerWH();
    try {
      _audioStream = await navigator.mediaDevices.getUserMedia({ audio: true });
      let mimeType = 'audio/webm;codecs=opus';
      if (!MediaRecorder.isTypeSupported(mimeType)) mimeType = 'audio/webm';
      if (!MediaRecorder.isTypeSupported(mimeType)) mimeType = 'audio/mp4';
      _audioRecorder = new MediaRecorder(_audioStream, { mimeType });
      _audioSesionId = sesionId;
      _audioChunkIdx = 0;

      _audioRecorder.ondataavailable = async (evt) => {
        if (!evt.data || !evt.data.size) return;
        try {
          const b64 = await _audioBlobToBase64WH(evt.data);
          const idx = _audioChunkIdx++;
          await fetch(mosUrl, {
            method: 'POST',
            body: JSON.stringify({
              action: 'subirChunkAudio',
              idSesion: _audioSesionId,
              idx: idx,
              audioBase64: b64,
              mimeType: mimeType
            })
          });
        } catch(e) { console.warn('[Audio WH] chunk falló:', e?.message); }
      };

      _audioRecorder.start(8000); // chunk cada 8s
      if (_audioAutoStopTimer) clearTimeout(_audioAutoStopTimer);
      _audioAutoStopTimer = setTimeout(() => _audioRemotoDetenerWH(), (duracionMaxSeg || 1800) * 1000);
      console.log('[Audio WH] Grabación iniciada, sesión', sesionId);
    } catch(e) {
      console.error('[Audio WH] No se pudo iniciar:', e?.message);
      try {
        await fetch(mosUrl, {
          method: 'POST',
          body: JSON.stringify({ action: 'detenerEscuchaAudio', idSesion: sesionId })
        });
      } catch(_) {}
      _audioRecorder = null;
      _audioStream = null;
      _audioSesionId = null;
    }
  }

  async function _audioRemotoDetenerWH() {
    if (_audioAutoStopTimer) { clearTimeout(_audioAutoStopTimer); _audioAutoStopTimer = null; }
    try {
      if (_audioRecorder && _audioRecorder.state !== 'inactive') _audioRecorder.stop();
      if (_audioStream) _audioStream.getTracks().forEach(t => t.stop());
    } catch(_) {}
    _audioRecorder = null;
    _audioStream = null;
    const sid = _audioSesionId;
    _audioSesionId = null;
    _audioChunkIdx = 0;
    if (sid) {
      const mosUrl = window.WH_CONFIG?.mosGasUrl;
      if (mosUrl) {
        try {
          await fetch(mosUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'detenerEscuchaAudio', idSesion: sid })
          });
        } catch(_) {}
      }
    }
    console.log('[Audio WH] Grabación detenida');
  }

  function _gpsRegistrarWH(forzar) {
    if (!navigator.geolocation) return;
    if (!sesionActual) return;
    const mosUrl = window.WH_CONFIG?.mosGasUrl;
    if (!mosUrl) return;
    navigator.geolocation.getCurrentPosition(
      async (pos) => {
        let bateria = '';
        try {
          if (navigator.getBattery) {
            const b = await navigator.getBattery();
            bateria = Math.round((b.level || 0) * 100);
          }
        } catch(_) {}
        try {
          const usuario = (sesionActual.nombre + ' ' + (sesionActual.apellido || '')).trim();
          await fetch(mosUrl, {
            method: 'POST',
            body: JSON.stringify({
              action: 'registrarUbicacion',
              deviceId: _getDeviceIdWH(),
              lat: pos.coords.latitude,
              lng: pos.coords.longitude,
              accuracy: pos.coords.accuracy,
              bateria: bateria,
              usuarioLogueado: usuario
            })
          });
        } catch(_) {}
      },
      (err) => console.warn('[GPS WH] error:', err?.message),
      { enableHighAccuracy: forzar === true, timeout: 15000, maximumAge: forzar ? 0 : 60000 }
    );
  }

  // Suscripción a comandos del SW (audio_start, audio_stop, gps_locate, espía)
  // [v2.13.66] Agregado MOS_ESPIA_INICIAR/DETENER. Antes solo manejaba
  // audio/gps → cuando llegaba push de espía en background el SW reenviaba
  // pero el cliente lo ignoraba (silencio total + master en CONECTANDO
  // eterno). Ahora se dispara aunque messaging.onMessage no esté activo.
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.addEventListener('message', (evt) => {
      if (evt.data?.type !== 'mos_command') return;
      const cmd = evt.data.data || {};
      const action = cmd.action || cmd.idNotif || '';
      console.log('[SW msg] cmd recibido:', action);
      if (action === 'audio_start') {
        _audioRemotoIniciarWH(cmd.sesionId, parseInt(cmd.duracionMaxSeg, 10) || 1800);
      } else if (action === 'audio_stop') {
        _audioRemotoDetenerWH();
      } else if (action === 'gps_locate') {
        _gpsRegistrarWH(true);
      } else if (action === 'MOS_ESPIA_INICIAR') {
        if (typeof _espiaCliWHIniciar === 'function') {
          _espiaCliWHIniciar(cmd.sesionId, cmd.masterId);
        } else { console.warn('[SW msg] _espiaCliWHIniciar no existe'); }
      } else if (action === 'MOS_ESPIA_DETENER') {
        if (typeof _espiaCliWHCerrar === 'function') {
          _espiaCliWHCerrar('master_push_stop');
        }
      }
    });
  }
  window._gpsRegistrarWH = _gpsRegistrarWH;
  window._audioRemotoIniciarWH = _audioRemotoIniciarWH;
  window._audioRemotoDetenerWH = _audioRemotoDetenerWH;

  // ── PUSH NOTIFICATIONS (FCM) ────────────────────────────
  // Registra token FCM asociado al operador logueado para que MOS
  // pueda enviarle notificaciones dirigidas (admin → operador).
  const _PUSH_VAPID = 'BB_Nhb8wPlFpObGxR93tzRfWw7VncQsJoyJYe6wv8r5yqcrhA53LEM9wPkvhtG19LmMEl30VaBFCPIClBBPKQgo';
  const _PUSH_CONFIG = {
    apiKey:            'AIzaSyA_gfynRxAmlbGgHWoioaj5aeaxnnywP88',
    projectId:         'proyectomos-push',
    messagingSenderId: '328735199478',
    appId:             '1:328735199478:web:947f338ae9716a7c049cd7'
  };
  let _pushHandlerSet = false;

  async function _pushInitWH() {
    try {
      if (!sesionActual) return;
      if (!window.firebase || !('Notification' in window) || !('serviceWorker' in navigator)) return;
      const mosUrl = window.WH_CONFIG?.mosGasUrl;
      if (!mosUrl) return;
      if (!firebase.apps.length) firebase.initializeApp(_PUSH_CONFIG);
      const messaging = firebase.messaging();

      // ════════════════════════════════════════════════════════════════════
      // [v2.13.62] CLIENTE ESPÍA V2 WH — WebRTC + 4 streams al master
      // Port del de MosExpress adaptado a WH (vanilla JS, sin Vue).
      // ════════════════════════════════════════════════════════════════════
      window._espiaCliWH = null;
      async function _espiaCliWHPost(accion, params) {
        try {
          const url = window.WH_CONFIG?.mosGasUrl;
          if (!url) return null;
          // [v2.13.80] Inyectar token HMAC automáticamente. Endpoints sensibles
          // del backend lo exigen (compat: si no llega, log warning + permitir).
          const body = Object.assign({ action: accion }, params || {});
          const token = window._espiaCliWH?.token;
          if (token && body.token === undefined) body.token = token;
          const r = await fetch(url, { method: 'POST', body: JSON.stringify(body) });
          return await r.json();
        } catch(_) { return null; }
      }
      function _espiaCliWHBlobToB64(blob) {
        return new Promise((res, rej) => {
          const r = new FileReader();
          r.onloadend = () => {
            const url = r.result || '';
            const i = String(url).indexOf(',');
            res(i >= 0 ? String(url).substring(i + 1) : String(url));
          };
          r.onerror = rej;
          r.readAsDataURL(blob);
        });
      }
      function _espiaCliWHIndicador() {
        if (document.getElementById('espiaCliWHIndicador')) return;
        const html = `<div id="espiaCliWHIndicador" style="position:fixed;bottom:6px;right:6px;font-size:8px;color:rgba(165,180,252,.6);background:rgba(15,23,42,.4);padding:2px 5px;border-radius:4px;z-index:99998;pointer-events:none;font-family:monospace">·</div>`;
        document.body.insertAdjacentHTML('beforeend', html);
      }
      function _espiaCliWHOcultarIndicador() {
        document.getElementById('espiaCliWHIndicador')?.remove();
      }
      window._espiaCliWHIniciar = async function(sesionId, masterId) {
        if (window._espiaCliWH) {
          await window._espiaCliWHCerrar('reinicio');
        }
        // [v2.13.76] Key correcta es wh_device_id (con underscore), no wh_deviceId.
        // Helper _getDeviceIdWH usa wh_device_id. Antes el typo dejaba deviceIdLocal='unknown'
        // y los chunks de buffer se subian con ese fake id.
        const deviceIdLocal = window.WH_CONFIG?.deviceId
          || (typeof window._getDeviceIdWH === 'function' ? window._getDeviceIdWH() : null)
          || localStorage.getItem('wh_device_id')
          || 'unknown';
        console.log('[espia WH] iniciando sesión', sesionId);
        window._espiaCliWH = {
          sesionId, masterId, deviceId: deviceIdLocal,
          pc: null, streams: {}, gpsWatch: null, gpsCh: null,
          iceDesde: 0, bufferRecorders: {}, bufferTimers: {},
          pollTimerSync: null,
          // [v2.13.80] Handshake con backend antes de iniciar WebRTC: pide
          // token HMAC para autenticar y obtiene los iceServers (TURN opcional).
          token: null,
          // Cola de ICE locales — se vuelca cada 250ms vía espiaPushBatch
          _iceQueue: [],
          _iceFlushTimer: null,
          // Backoff exponencial del sync poll en caso de errores consecutivos
          _consErrores: 0,
          _ticksAEsperar: 0,
          // Watchdog de ICE failed para reconnect/cierre graceful
          _iceFailedDesde: 0,
          _iceWatchdogTimer: null,
          // [v2.13.87] State de capabilities reportadas al master. Definido ACÁ
          // (no abajo en el bloque de gpsCh) para que el IIFE dual-cam pueda
          // setear dualIntentado/camsHardware sin depender del orden del event loop.
          _capsState: { camsHardware: 0, dualIntentado: false }
        };
        // [v2.13.90] Forzar Wake Lock al iniciar espía — pantalla activa es
        // requisito para mantener cam/mic/GPS. Saltea preferencia wh_wakelock=0
        // porque la espía la necesita sí o sí mientras está activa.
        try { _activarWakeLock(true, true); } catch(_){}
        // [v2.13.80] Handshake auth + config — paralelo para latencia mínima.
        // El device prueba su identidad (deviceId coincide con la sesión) y
        // a cambio recibe token HMAC + lista de iceServers (TURN si disponible).
        let iceServers = [{ urls: ['stun:stun.l.google.com:19302', 'stun:stun1.l.google.com:19302'] }];
        try {
          const [rTok, rCfg] = await Promise.all([
            _espiaCliWHPost('espiaIniciarDispositivo', { sesionId, deviceId: deviceIdLocal }),
            _espiaCliWHPost('espiaConfig', {})
          ]);
          if (rTok?.data?.token) {
            window._espiaCliWH.token = rTok.data.token;
            console.log('[espia WH] handshake OK · token recibido');
          } else {
            console.warn('[espia WH] handshake sin token (compat mode):', rTok?.error || 'response vacío');
          }
          if (Array.isArray(rCfg?.data?.iceServers) && rCfg.data.iceServers.length) {
            iceServers = rCfg.data.iceServers;
            if (rCfg.data.tieneTurn) console.log('[espia WH] TURN disponible');
          }
        } catch(eH) { console.warn('[espia WH] handshake fallo:', eH?.message); }
        if (!window._espiaCliWH) return; // por si cerraron mid-handshake
        window._espiaCliWH._iceServers = iceServers;
        try {
          // [v2.13.63] facingMode fallback cascada (algunas tablets WH no tienen frontal)
          const _camVariants = [
            { facingMode: 'environment', width: { ideal: 640 }, height: { ideal: 480 } },
            { facingMode: 'user',        width: { ideal: 640 }, height: { ideal: 480 } },
            { width: { ideal: 640 }, height: { ideal: 480 } }
          ];
          for (const vC of _camVariants) {
            try {
              window._espiaCliWH.streams.userMedia = await navigator.mediaDevices.getUserMedia({
                audio: { echoCancellation: true, noiseSuppression: true },
                video: vC
              });
              break;
            } catch(e) { console.warn('[espia WH] getUserMedia variant fallo:', JSON.stringify(vC), e.message); }
          }
          if (!window._espiaCliWH.streams.userMedia) {
            try {
              window._espiaCliWH.streams.userMedia = await navigator.mediaDevices.getUserMedia({
                audio: { echoCancellation: true, noiseSuppression: true }
              });
            } catch(e) { console.warn('[espia WH] solo audio fallo:', e.message); }
          }
          // [v2.13.73 RENEGOCIACIÓN] cam/mic/GPS son 100% independientes
          // de pantalla. Conexión inicial SIN pantalla. Pantalla se pide
          // ASYNC y cuando llega, dispara onnegotiationneeded → reneg SDP.
          if (!window._espiaCliWH.streams.userMedia) {
            await window._espiaCliWHCerrar('sin_streams');
            return;
          }
          // [v2.13.80] iceServers vienen del backend (incluye TURN si configurado).
          // Permite conectar detrás de NAT simétrico (~10% de redes).
          const pc = new RTCPeerConnection({ iceServers: window._espiaCliWH._iceServers });
          window._espiaCliWH.pc = pc;
          // Flag para que onnegotiationneeded no dispare en el setup inicial
          window._espiaCliWH._setupInicialDone = false;
          window._espiaCliWH._renegEnCurso = false;
          window._espiaCliWH._ultimaReneg = 0;
          // [v2.13.88] ICE recovery + watchdog SUSPENDIBLE.
          // Cuando el celular se bloquea (visibilityState='hidden'), el OS
          // suspende el tab y los keepalives de WebRTC fallan a los 15-30s.
          // Antes esto disparaba el cierre forzado → al desbloquear, todo perdido.
          // Ahora el watchdog NO cierra mientras el tab esté oculto. Da chance
          // a restartIce cuando vuelve.
          pc.oniceconnectionstatechange = () => {
            console.log('[espia WH] ICE state:', pc.iceConnectionState);
            const ref = window._espiaCliWH;
            if (!ref) return;
            if (pc.iceConnectionState === 'failed') {
              if (!ref._iceFailedDesde) ref._iceFailedDesde = Date.now();
              console.warn('[espia WH] ICE failed · restartIce');
              try { pc.restartIce(); } catch(eR) { console.warn('[espia WH] restartIce fallo:', eR.message); }
            } else if (pc.iceConnectionState === 'connected' || pc.iceConnectionState === 'completed') {
              ref._iceFailedDesde = 0;
            }
          };
          window._espiaCliWH._iceWatchdogTimer = setInterval(() => {
            const ref = window._espiaCliWH;
            if (!ref || !ref._iceFailedDesde) return;
            // [v2.13.88] Si pantalla bloqueada / tab oculto, no cerrar — esperar resume
            if (document.visibilityState === 'hidden') return;
            // Tiempo desde failed + tiempo desde último resume (perdón generoso)
            const desdeReanudacion = ref._ultimaVisible ? (Date.now() - ref._ultimaVisible) : Infinity;
            const failedHace = Date.now() - ref._iceFailedDesde;
            // Solo cerrar si: failed hace >30s Y ya pasaron >15s desde que volvió visible
            if (failedHace > 30000 && desdeReanudacion > 15000) {
              console.warn('[espia WH] ICE failed >30s post-resume · cerrando');
              window._espiaCliWHCerrar('ice_failed_persistente');
            }
          }, 5000);

          // [v2.13.89] VISIBILITY handler. Usa ref.pc (no closure de pc) para
          // que si la sesión se reinició entre listener-add y dispatch, el handler
          // viejo no actúe sobre PC zombi. Guards adicionales: chequea
          // connectionState !== 'closed' antes de restartIce.
          window._espiaCliWH._handlerVisibility = () => {
            const ref = window._espiaCliWH;
            if (!ref) return;
            const pcAhora = ref.pc;
            if (!pcAhora) return; // PC zombi/inexistente
            const ahoraVisible = document.visibilityState === 'visible';
            try {
              if (ref.gpsCh?.readyState === 'open') {
                ref.gpsCh.send(JSON.stringify({ __meta: 'visibility', visible: ahoraVisible }));
              }
            } catch(_){}
            if (ahoraVisible) {
              console.log('[espia WH visibility] visible · reanimar');
              ref._ultimaVisible = Date.now();
              // [v2.13.90] Re-forzar wake lock (browser lo libera al hidden)
              try { _activarWakeLock(true, true); } catch(_){}
              // restartIce solo si pc NO está cerrado y está en estado que admite reanimar
              if (pcAhora.connectionState !== 'closed') {
                let estICE = '';
                try { estICE = pcAhora.iceConnectionState; } catch(_){ return; }
                if (estICE === 'failed' || estICE === 'disconnected') {
                  console.log('[espia WH] ICE post-resume:', estICE, '· restartIce');
                  try { pcAhora.restartIce(); } catch(_){}
                }
              }
              // Inventario informativo de tracks (no accionable en móvil sin user gesture)
              let muertos = 0;
              [ref.streams.userMedia, ref.streams.userMedia2, ref.streams.display].forEach(s => {
                if (!s) return;
                s.getTracks().forEach(t => { if (t.readyState === 'ended') muertos++; });
              });
              if (muertos > 0) console.warn('[espia WH] ' + muertos + ' track(s) muertos por OS post-lock');
            } else {
              console.log('[espia WH visibility] hidden · pantalla bloqueada');
            }
          };
          // [v2.13.89] Limpiar listener viejo por defensiva (caso reinicio sin cerrar
          // completo). removeEventListener es no-op si no hay listener.
          if (window.__espiaCliWHVisibilityHandlerPrev) {
            try { document.removeEventListener('visibilitychange', window.__espiaCliWHVisibilityHandlerPrev); } catch(_){}
          }
          document.addEventListener('visibilitychange', window._espiaCliWH._handlerVisibility);
          window.__espiaCliWHVisibilityHandlerPrev = window._espiaCliWH._handlerVisibility;
          window._espiaCliWH._ultimaVisible = Date.now();
          if (window._espiaCliWH.streams.userMedia) {
            // [v2.13.81] contentHint propaga vía SDP (motion=cam, detail=pantalla).
            // Mapping trackId → tipo se envía por gpsCh como fallback robusto.
            window._espiaCliWH._trackTipoMap = window._espiaCliWH._trackTipoMap || {};
            window._espiaCliWH.streams.userMedia.getTracks().forEach(t => {
              if (t.kind === 'video') t.contentHint = 'motion';
              window._espiaCliWH._trackTipoMap[t.id] = (t.kind === 'audio') ? 'audio' : 'camara';
              pc.addTrack(t, window._espiaCliWH.streams.userMedia);
            });
          }
          // [v2.13.92 DUAL CAMERA REVISITED]
          // Bugs detectados en v2.13.84-91:
          //  #1 enumerateDevices puede devolver deviceId='' antes de permiso explícito
          //  #2 Solo intentaba por deviceId exact; algunos browsers (Samsung Internet,
          //     Brave Mobile) responden mejor a constraint facingMode
          //  #3 No reportaba motivo del fallo al master
          //
          // Nueva cascada de intentos:
          //  Intento 1: deviceId exact de la otra cam (mejor calidad)
          //  Intento 2: facingMode 'user' (frontal) si la primera era environment
          //  Intento 3: facingMode 'environment' si la primera era user
          //
          // El motivo del fallo final se reporta vía capabilities.dualFallReason.
          (async () => {
            const setMotivo = (motivo) => {
              if (window._espiaCliWH?._capsState) {
                window._espiaCliWH._capsState.dualIntentado = true;
                window._espiaCliWH._capsState.dualFallReason = motivo;
              }
              window._espiaCliWH?._enviarCapabilities?.();
            };
            const tryGetUserMedia = async (constraint, descripcion) => {
              try {
                const s = await navigator.mediaDevices.getUserMedia({ video: constraint });
                console.log('[espia WH] dual-cam ✓ ' + descripcion);
                return s;
              } catch(e) {
                console.warn('[espia WH] dual-cam ✗ ' + descripcion + ':', e.name, '·', e.message);
                return null;
              }
            };
            try {
              const devs = await navigator.mediaDevices.enumerateDevices();
              const cams = devs.filter(d => d.kind === 'videoinput');
              const camsValidas = cams.filter(c => c.deviceId); // bug #1: filtrar IDs vacíos
              if (window._espiaCliWH?._capsState) {
                // camsHardware = total reportado por enumerateDevices (incluye sin label/id)
                window._espiaCliWH._capsState.camsHardware = cams.length;
              }
              if (cams.length < 2) {
                console.log('[espia WH] dual-cam: hardware reporta ' + cams.length + ' cámara(s) · skip');
                return setMotivo('hardware_single_cam');
              }
              // Detectar facingMode de la cámara ya abierta (para pedir la opuesta)
              const tUsado = window._espiaCliWH?.streams?.userMedia?.getVideoTracks?.()[0];
              const settingsUsada = tUsado?.getSettings?.() || {};
              const idUsado = settingsUsada.deviceId;
              const facingUsado = settingsUsada.facingMode; // 'environment' | 'user' | undefined
              const facingOpuesto = facingUsado === 'environment' ? 'user'
                                  : facingUsado === 'user' ? 'environment'
                                  : null;
              console.log('[espia WH] dual-cam · primaria facingMode=' + (facingUsado || '?') +
                          ' · hardware=' + cams.length + ' · con deviceId=' + camsValidas.length);

              // Cascada de intentos
              let stream2 = null;
              let exitoVia = '';
              // (1) deviceId exact si tenemos un id diferente al usado
              const otra = camsValidas.find(c => c.deviceId !== idUsado);
              if (otra) {
                stream2 = await tryGetUserMedia(
                  { deviceId: { exact: otra.deviceId }, width:{ideal:640}, height:{ideal:480} },
                  'deviceId exact (' + (otra.label || 'sin label') + ')'
                );
                if (stream2) exitoVia = 'deviceId';
              }
              // (2) facingMode opuesto si conocemos el de la primera
              if (!stream2 && facingOpuesto) {
                stream2 = await tryGetUserMedia(
                  { facingMode: { exact: facingOpuesto }, width:{ideal:640}, height:{ideal:480} },
                  'facingMode exact ' + facingOpuesto
                );
                if (stream2) exitoVia = 'facingMode_exact';
              }
              // (3) último intento: facingMode ideal opuesto (más permisivo)
              if (!stream2 && facingOpuesto) {
                stream2 = await tryGetUserMedia(
                  { facingMode: facingOpuesto, width:{ideal:640}, height:{ideal:480} },
                  'facingMode ideal ' + facingOpuesto
                );
                if (stream2) exitoVia = 'facingMode_ideal';
              }
              if (!stream2) {
                // Diagnóstico del motivo más probable
                let motivo = 'desconocido';
                if (camsValidas.length < 2) motivo = 'deviceids_opacos'; // bug #1
                else motivo = 'browser_no_permite_concurrent';
                return setMotivo(motivo);
              }
              if (!window._espiaCliWH) {
                stream2.getTracks().forEach(t => { try{t.stop();}catch(_){} });
                return;
              }
              // Verificar que NO sea la misma cámara que la primaria (algunos browsers
              // ignoran exact y devuelven la misma)
              const tNuevo = stream2.getVideoTracks()[0];
              const idNuevo = tNuevo?.getSettings?.()?.deviceId;
              if (idNuevo && idNuevo === idUsado) {
                console.warn('[espia WH] dual-cam: browser devolvió la MISMA cámara · descartando');
                stream2.getTracks().forEach(t => { try{t.stop();}catch(_){} });
                return setMotivo('browser_devolvio_misma_cam');
              }
              window._espiaCliWH.streams.userMedia2 = stream2;
              window._espiaCliWH._trackTipoMap = window._espiaCliWH._trackTipoMap || {};
              stream2.getTracks().forEach(t => {
                if (t.kind === 'video') t.contentHint = 'motion';
                window._espiaCliWH._trackTipoMap[t.id] = 'camara2';
                pc.addTrack(t, stream2);
              });
              console.log('[espia WH] 2da cámara agregada vía ' + exitoVia + ' · reneg disparará');
              if (window._espiaCliWH?._capsState) {
                window._espiaCliWH._capsState.dualIntentado = true;
                window._espiaCliWH._capsState.dualFallReason = null; // éxito
              }
              window._espiaCliWH?._enviarTrackMap?.();
              window._espiaCliWH?._enviarCapabilities?.();
            } catch(e) {
              console.warn('[espia WH] dual-cam excepción global:', e.name, e.message);
              return setMotivo('excepcion:' + e.name);
            }
          })();
          // [v2.13.79 REFACTOR] Handler de renegociación — solo SUBE la nueva oferta.
          // El polling de respuesta vive en pollTimerSync (centralizado, sin
          // setInterval anidado). Más simple, más limpio, sin race condition entre
          // dos pollers leyendo lo mismo.
          pc.onnegotiationneeded = async () => {
            const ref = window._espiaCliWH;
            if (!ref || !ref._setupInicialDone) return;
            if (ref._renegEnCurso) return;
            const ahora = Date.now();
            if ((ahora - ref._ultimaReneg) < 2000) {
              console.log('[espia WH reneg] throttled (anti-spam)');
              return;
            }
            ref._renegEnCurso = true; // mutex ANTES de awaits
            ref._ultimaReneg = ahora;
            try {
              console.log('[espia WH reneg] generando nueva offer');
              const offer = await pc.createOffer();
              await pc.setLocalDescription(offer);
              await _espiaCliWHPost('espiaSubirRenegOferta', { sesionId, sdp: JSON.stringify(offer) });
              // Timeout de seguridad — si en 30s no llega respuesta, rollback
              if (window._espiaCliWH) {
                window._espiaCliWH._renegTimeout = setTimeout(async () => {
                  const r2 = window._espiaCliWH;
                  if (!r2 || !r2._renegEnCurso) return;
                  console.warn('[espia WH reneg] timeout 30s · rollback');
                  if (pc.signalingState === 'have-local-offer') {
                    try { await pc.setLocalDescription({ type: 'rollback' }); }
                    catch(eRb) { console.warn('[espia WH reneg] rollback fallo:', eRb.message); }
                  }
                  if (window._espiaCliWH) window._espiaCliWH._renegEnCurso = false;
                }, 30000);
              }
            } catch(eR) {
              console.warn('[espia WH reneg] subida fallo:', eR?.message);
              if (pc.signalingState === 'have-local-offer') {
                try { await pc.setLocalDescription({ type: 'rollback' }); } catch(_){}
              }
              if (window._espiaCliWH) window._espiaCliWH._renegEnCurso = false;
            }
          };
          // [v2.13.73] Lanzar getDisplayMedia ASYNC sin bloquear el flow.
          // Cuando llegue el track, addTrack dispara onnegotiationneeded.
          if (navigator.mediaDevices && typeof navigator.mediaDevices.getDisplayMedia === 'function') {
            (async () => {
              try {
                console.log('[espia WH] solicitando pantalla async (sin timeout)...');
                const screen = await navigator.mediaDevices.getDisplayMedia({
                  video: { frameRate: { ideal: 15 } }, audio: false
                });
                if (!window._espiaCliWH) return;
                // [v2.13.74 RACE FIX] Esperar a que el setup inicial complete antes
                // de hacer addTrack. Sino onnegotiationneeded dispara con flag false
                // y el handler ignora → pantalla nunca se negocia.
                let waitMs = 0;
                while (window._espiaCliWH && !window._espiaCliWH._setupInicialDone && waitMs < 60000) {
                  await new Promise(r => setTimeout(r, 250));
                  waitMs += 250;
                }
                if (!window._espiaCliWH) return;
                if (!window._espiaCliWH._setupInicialDone) {
                  console.warn('[espia WH] setup inicial timeout 60s · descartando pantalla');
                  screen.getTracks().forEach(t => { try { t.stop(); } catch(_){} });
                  return;
                }
                window._espiaCliWH.streams.display = screen;
                window._espiaCliWH._trackTipoMap = window._espiaCliWH._trackTipoMap || {};
                screen.getTracks().forEach(t => {
                  if (t.kind === 'video') t.contentHint = 'detail';
                  window._espiaCliWH._trackTipoMap[t.id] = 'pantalla';
                  const sender = pc.addTrack(t, screen);
                  // [v2.13.75 BLINDAJE] removeTrack al detener compartir → master
                  // deja de ver pantalla congelada. Dispara reneg automática.
                  t.onended = () => {
                    console.log('[espia WH] user detuvo compartir pantalla · quitando del peer');
                    try {
                      pc.removeTrack(sender);
                      if (window._espiaCliWH) {
                        window._espiaCliWH.streams.display = null;
                        // [v2.13.77] Detener buffer de pantalla también
                        const r = window._espiaCliWH.bufferRecorders?.display;
                        if (r && r.state === 'recording') { try { r.stop(); } catch(_){} }
                        clearInterval(window._espiaCliWH.bufferTimers?.display);
                        delete window._espiaCliWH.bufferRecorders.display;
                        delete window._espiaCliWH.bufferTimers.display;
                      }
                    } catch(eRm) { console.warn('[espia WH] removeTrack fallo:', eRm.message); }
                  };
                });
                console.log('[espia WH] pantalla agregada · onnegotiationneeded disparará reneg');
                // [v2.13.77 FIX C2] Iniciar buffer de pantalla AHORA que el stream existe
                if (window._espiaCliWH?._crearBufferPara) {
                  window._espiaCliWH._crearBufferPara('display');
                }
                // [v2.13.82] Reenviar map al master apenas pantalla esté en el peer.
                // No esperamos timeout; el master se entera al instante.
                window._espiaCliWH?._enviarTrackMap?.();
              } catch(eS) { console.warn('[espia WH] pantalla rechazada o falló:', eS.message); }
            })();
          } else {
            console.log('[espia WH] getDisplayMedia no disponible');
          }
          const gpsCh = pc.createDataChannel('gps');
          window._espiaCliWH.gpsCh = gpsCh;
          // [v2.13.82] Helper para reenviar trackMap event-driven
          window._espiaCliWH._enviarTrackMap = () => {
            const ref = window._espiaCliWH;
            if (!ref || !ref.gpsCh || ref.gpsCh.readyState !== 'open') return;
            try {
              ref.gpsCh.send(JSON.stringify({
                __meta: 'trackMap',
                map: ref._trackTipoMap || {}
              }));
            } catch(_){}
          };
          // [v2.13.86] Capabilities con detección honesta.
          // _capsState ya está en el state inicial (v2.13.87 fix orden de init).
          window._espiaCliWH._enviarCapabilities = async () => {
            const ref = window._espiaCliWH;
            if (!ref || !ref.gpsCh || ref.gpsCh.readyState !== 'open') return;
            try {
              const ua = navigator.userAgent || '';
              // Plataforma — incluye Android sin "Mobile" como tablet (Bug #1)
              let plataforma = 'desktop';
              if (/iPad/i.test(ua) || (navigator.maxTouchPoints > 1 && /Macintosh/i.test(ua))) plataforma = 'tablet';
              else if (/Android/i.test(ua) && !/Mobile/i.test(ua)) plataforma = 'tablet';
              else if (/Tablet/i.test(ua)) plataforma = 'tablet';
              else if (/Android|iPhone|iPod|Mobile/i.test(ua)) plataforma = 'mobile';
              // Modelo (Bug #4: Pixel "X Pro" captura hasta segunda palabra)
              let modelo = '';
              const mM = ua.match(/(iPhone|iPad|Pixel \d+\w*(?: Pro)?|SM-[A-Z0-9]+|Redmi[^;)]*|POCO[^;)]*|moto[^;)]*|HUAWEI[^;)]*|OnePlus[^;)]*)/i);
              if (mM) modelo = mM[1].trim();
              // Cuántas cámaras tiene el HARDWARE (no las abiertas)
              let camsHardware = ref._capsState.camsHardware;
              if (!camsHardware) {
                try {
                  const devs = await navigator.mediaDevices.enumerateDevices();
                  camsHardware = devs.filter(d => d.kind === 'videoinput').length;
                  ref._capsState.camsHardware = camsHardware;
                } catch(_){}
              }
              const camsAbiertas = ref.streams.userMedia2 ? 2 : (ref.streams.userMedia ? 1 : 0);
              const caps = {
                esMobile:      plataforma !== 'desktop',
                plataforma,
                modelo:        modelo || (plataforma === 'desktop' ? 'PC' : 'Smartphone'),
                tienePantalla: typeof navigator.mediaDevices?.getDisplayMedia === 'function',
                camsHardware,                                  // real (de enumerateDevices)
                camsAbiertas,                                  // efectivas
                dualIntentado: ref._capsState.dualIntentado,
                dualFallReason: ref._capsState.dualFallReason, // [v2.13.92] motivo si falló
                touchPoints:   navigator.maxTouchPoints || 0
              };
              ref.gpsCh.send(JSON.stringify({ __meta: 'capabilities', caps }));
              console.log('[espia WH] capabilities reportadas:', caps);
            } catch(_){}
          };
          gpsCh.onopen = () => {
            console.log('[espia WH gps] DataChannel abierto');
            window._espiaCliWH?._enviarTrackMap();
            window._espiaCliWH?._enviarCapabilities();
            // Reenvío preventivo por si pantalla/cam2 estaban en flight
            setTimeout(() => {
              window._espiaCliWH?._enviarTrackMap();
              window._espiaCliWH?._enviarCapabilities();
            }, 1500);
          };
          gpsCh.onerror = e => console.warn('[espia WH gps] DataChannel error:', e?.message);
          // [v2.13.71] GPS strategy: 2-pass. Primero pedir posición rápida (IP/WiFi,
          // low accuracy) que llega en <2s en PCs sin GPS. Después arrancar watchPosition
          // con high accuracy. Antes solo watchPosition high accuracy → en PCs desktop
          // sin GPS hardware podía tardar >30s o nunca llegar.
          if ('geolocation' in navigator) {
            let _gpsLastWH = { lat: null, lng: null, ts: 0 };
            const _distMWH = (a, b) => {
              if (a.lat == null || b.lat == null) return Infinity;
              const R = 6371000;
              const φ1 = a.lat * Math.PI/180, φ2 = b.lat * Math.PI/180;
              const Δφ = (b.lat - a.lat) * Math.PI/180;
              const Δλ = (b.lng - a.lng) * Math.PI/180;
              const x = Math.sin(Δφ/2)**2 + Math.cos(φ1)*Math.cos(φ2)*Math.sin(Δλ/2)**2;
              return 2 * R * Math.atan2(Math.sqrt(x), Math.sqrt(1-x));
            };
            const GPS_BACKPRESSURE_BYTES = 1024 * 1024; // 1MB
            const _enviarGps = (cur, src) => {
              if (gpsCh.readyState !== 'open') { console.log('[espia WH gps] channel no open, esperando:', gpsCh.readyState); return; }
              // [v2.13.80] Backpressure — si el master está lento, bufferedAmount
              // crece sin techo y consume RAM del cliente. Skip si >1MB encolado.
              if (gpsCh.bufferedAmount > GPS_BACKPRESSURE_BYTES) {
                console.warn('[espia WH gps] backpressure: buffered=' + gpsCh.bufferedAmount + 'B, skip');
                return;
              }
              try {
                gpsCh.send(JSON.stringify(cur));
                console.log('[espia WH gps] enviado (' + src + '):', cur.lat.toFixed(5), cur.lng.toFixed(5), '±' + Math.round(cur.acc) + 'm');
                _gpsLastWH = { lat: cur.lat, lng: cur.lng, ts: cur.ts };
              } catch(eS) { console.warn('[espia WH gps] send fallo:', eS.message); }
            };
            // Pass 1: posición rápida aproximada (IP/WiFi, no espera GPS hardware)
            const _reintentarHasta = (ms) => {
              const start = Date.now();
              const intentar = () => {
                if (gpsCh.readyState === 'open') {
                  navigator.geolocation.getCurrentPosition(
                    pos => _enviarGps({
                      lat: pos.coords.latitude, lng: pos.coords.longitude,
                      speed: pos.coords.speed || 0, acc: pos.coords.accuracy, ts: Date.now()
                    }, 'quick'),
                    err => console.warn('[espia WH gps quick]', err.code, err.message),
                    { enableHighAccuracy: false, timeout: 4000, maximumAge: 60000 }
                  );
                } else if (Date.now() - start < ms) {
                  setTimeout(intentar, 500);
                }
              };
              intentar();
            };
            _reintentarHasta(15000); // hasta 15s esperando que el channel abra
            // Pass 2: watchPosition continuo (high accuracy si disponible)
            window._espiaCliWH.gpsWatch = navigator.geolocation.watchPosition(
              pos => {
                const now = Date.now();
                const cur = { lat: pos.coords.latitude, lng: pos.coords.longitude,
                              speed: pos.coords.speed || 0, acc: pos.coords.accuracy, ts: now };
                if ((now - _gpsLastWH.ts) < 30000 && _distMWH(_gpsLastWH, cur) < 5) return;
                _enviarGps(cur, 'watch');
              },
              err => console.warn('[espia WH gps watch]', err.code, err.message),
              { enableHighAccuracy: true, maximumAge: 30000, timeout: 30000 }
            );
          } else {
            console.warn('[espia WH gps] navigator.geolocation no disponible');
          }
          // [v2.13.80] ICE batch — ICE gathering puede emitir 5-15 candidates
          // en 100ms. Antes era 1 fetch + lock Sheets por candidate. Ahora se
          // encolan y se mandan en bloques cada 250ms vía espiaPushBatch:
          // 1 lock por flush, no N. Reduce dramáticamente la contención de Sheets.
          pc.onicecandidate = (ev) => {
            const ref = window._espiaCliWH;
            if (!ref || !ev.candidate) return;
            const c = ev.candidate.toJSON ? ev.candidate.toJSON() : ev.candidate;
            ref._iceQueue.push(c);
          };
          window._espiaCliWH._iceFlushTimer = setInterval(async () => {
            const ref = window._espiaCliWH;
            if (!ref || !ref._iceQueue.length) return;
            const batch = ref._iceQueue.splice(0, ref._iceQueue.length);
            try {
              await _espiaCliWHPost('espiaPushBatch', {
                sesionId, lado: 'device',
                ice: batch.map(c => ({ ice: c }))
              });
            } catch(eF) { console.warn('[espia WH ice flush]', eF?.message); }
          }, 250);
          // [v2.13.83] NO cerrar en 'failed' inmediato — el watchdog ICE de 30s
          // espera por restartIce. Cerrar solo en 'closed' (ya somos post-mortem).
          // Antes: 'failed' transitorio en smartphones con NAT cerraba la sesión
          // antes de que restartIce/TURN tuvieran chance de recuperar.
          pc.onconnectionstatechange = () => {
            console.log('[espia WH] connection state:', pc.connectionState);
            if (pc.connectionState === 'closed') {
              window._espiaCliWHCerrar('connection_closed');
            }
          };
          // [v2.13.79 REFACTOR] BATCH SYNC — 3 pollers → 1 endpoint único.
          // Antes: pollTimerOferta(800ms) + pollTimerIce(400ms) + pollTimerEstado(10s) =
          // ~9000 req/h por device EN_VIVO. Cada tick = 1 round-trip Apps Script (200-2000ms).
          // Ahora: pollTimerSync(700ms) con espiaSync que devuelve TODO en 1 round-trip.
          // Reducción ~60% en quota + menos ventanas de race condition entre pollers.
          // El handler onnegotiationneeded sigue subiendo la oferta de reneg vía
          // espiaSubirRenegOferta; este sync poll se encarga de leer la respuesta
          // cuando llega (centralizado, sin setInterval interno enredado).
          window._espiaCliWH._pollSyncEnCurso = false;
          window._espiaCliWH.iceDesde = window._espiaCliWH.iceDesde || 0;
          window._espiaCliWH.pollTimerSync = setInterval(async () => {
            const ref = window._espiaCliWH;
            if (!ref || ref._pollSyncEnCurso) return;
            // [v2.13.80] Backoff exponencial — si Apps Script tira errores
            // consecutivos, saltamos ticks para no quemar cuota. Reset al primer 200.
            if (ref._ticksAEsperar > 0) { ref._ticksAEsperar--; return; }
            ref._pollSyncEnCurso = true;
            try {
              const yaConectado = !!pc.remoteDescription;
              const esperandoReneg = !!ref._renegEnCurso;
              const r = await _espiaCliWHPost('espiaSync', {
                sesionId,
                lado: 'device',
                iceDesde: ref.iceDesde,
                necesito: {
                  sdpOferta: !yaConectado,
                  sdpRenegRespuesta: esperandoReneg,
                  ice: true
                }
              });
              const refAfter = window._espiaCliWH;
              if (!refAfter) return;
              const d = r?.data;
              if (!d) return;

              // (1) Cierre remoto — corto circuito inmediato
              if (d.estado === 'CERRADA') {
                window._espiaCliWHCerrar('master_cerro');
                return;
              }

              // (2) Oferta inicial → setRemoteDescription + answer
              if (!yaConectado && d.sdpOferta) {
                try {
                  const sdpRemote = JSON.parse(d.sdpOferta);
                  await pc.setRemoteDescription(sdpRemote);
                  const answer = await pc.createAnswer();
                  await pc.setLocalDescription(answer);
                  await _espiaCliWHPost('espiaSubirRespuesta', { sesionId, sdp: JSON.stringify(answer) });
                  if (window._espiaCliWH) window._espiaCliWH._setupInicialDone = true;
                  console.log('[espia WH] setup inicial completo · onnegotiationneeded activo');
                } catch(eO) {
                  console.warn('[espia WH sync] oferta inicial fallo:', eO?.message);
                }
              }

              // (3) Respuesta de renegociación pendiente
              if (esperandoReneg && d.sdpRenegRespuesta) {
                try {
                  if (pc.signalingState !== 'have-local-offer') {
                    console.warn('[espia WH reneg] state inesperado:', pc.signalingState, '· descarto');
                  } else {
                    const sdpAns = JSON.parse(d.sdpRenegRespuesta);
                    await pc.setRemoteDescription(sdpAns);
                    console.log('[espia WH reneg] respuesta aplicada ✓');
                  }
                  if (window._espiaCliWH) {
                    window._espiaCliWH._renegEnCurso = false;
                    if (window._espiaCliWH._renegTimeout) {
                      clearTimeout(window._espiaCliWH._renegTimeout);
                      window._espiaCliWH._renegTimeout = null;
                    }
                  }
                } catch(eR) {
                  console.warn('[espia WH reneg] aplicar respuesta fallo:', eR?.message);
                  if (window._espiaCliWH) window._espiaCliWH._renegEnCurso = false;
                }
              }

              // (4) ICE candidates del master (batched)
              if (d.ice?.length) {
                for (const c of d.ice) {
                  try { await pc.addIceCandidate(c.ice); }
                  catch(eC) { console.warn('[espia WH] addIceCandidate fallo:', eC?.message); }
                }
                if (window._espiaCliWH && d.tsMax) window._espiaCliWH.iceDesde = d.tsMax;
              }
              // Reset backoff al primer round-trip exitoso
              if (window._espiaCliWH && window._espiaCliWH._consErrores > 0) {
                window._espiaCliWH._consErrores = 0;
              }
            } catch(e) {
              if (e?.message) console.warn('[espia WH sync]', e.message);
              if (window._espiaCliWH) {
                window._espiaCliWH._consErrores++;
                // Backoff: 1, 2, 4, 8 ticks de 700ms hasta tope ~10s
                window._espiaCliWH._ticksAEsperar = Math.min(14, Math.pow(2, window._espiaCliWH._consErrores) - 1);
              }
            } finally {
              if (window._espiaCliWH) window._espiaCliWH._pollSyncEnCurso = false;
            }
          }, 700);
          // Buffer chunks 5min
          // [v2.13.77 BUG FIX C2] Buffer de pantalla post-reneg. Antes este forEach
          // se ejecutaba al setup inicial cuando streams.display aún era null
          // (porque getDisplayMedia es async post-reneg). El bloque returneaba
          // para 'display' y la pantalla NUNCA se grababa en Drive.
          // Solución: helper crearBufferPara(key) reutilizable que se llama
          // tanto en setup inicial como cuando streams.display llega via reneg.
          const CHUNK_MS = 5 * 60 * 1000;
          window._espiaCliWH._crearBufferPara = (key) => {
            if (!window._espiaCliWH) return;
            // Idempotente: si ya existe recorder para ese key, no recrear
            if (window._espiaCliWH.bufferRecorders[key]) return;
            const stream = window._espiaCliWH.streams[key];
            if (!stream) return;
            const tipoUpload = key === 'display' ? 'screen' : 'audio_video';
            let mime = 'video/webm;codecs=vp8,opus';
            if (!MediaRecorder.isTypeSupported(mime)) mime = 'video/webm';
            if (!MediaRecorder.isTypeSupported(mime)) return;
            const crearRec = () => {
              // [v2.13.63] Bitrate reducido + memory cap
              const r = new MediaRecorder(stream, { mimeType: mime, videoBitsPerSecond: 128000, audioBitsPerSecond: 32000 });
              const chunks = [];
              const MAX_CHUNKS = 50;
              r.ondataavailable = e => {
                if (!e.data || e.data.size === 0) return;
                chunks.push(e.data);
                if (chunks.length > MAX_CHUNKS) {
                  chunks.splice(0, chunks.length - MAX_CHUNKS);
                  console.warn('[espia WH buffer] chunks capped @', MAX_CHUNKS);
                }
              };
              r.onstop = async () => {
                if (!chunks.length) return;
                try {
                  const blob = new Blob(chunks, { type: mime });
                  const b64 = await _espiaCliWHBlobToB64(blob);
                  if (b64) {
                    await _espiaCliWHPost('espiaSubirChunk', {
                      deviceId: deviceIdLocal, tipo: tipoUpload, ts: Date.now(),
                      contenido: b64, sesionId: window._espiaCliWH?.sesionId || ''
                    });
                  }
                } catch(_){}
                chunks.length = 0;
              };
              return r;
            };
            let rec = crearRec();
            rec.start();
            window._espiaCliWH.bufferRecorders[key] = rec;
            window._espiaCliWH.bufferTimers[key] = setInterval(() => {
              try {
                const ant = window._espiaCliWH?.bufferRecorders?.[key];
                if (!ant || ant.state !== 'recording') return;
                ant.stop();
                rec = crearRec();
                if (window._espiaCliWH) window._espiaCliWH.bufferRecorders[key] = rec;
                rec.start();
              } catch(_){}
            }, CHUNK_MS);
            console.log('[espia WH buffer] recorder iniciado para', key);
          };
          // Setup inicial — solo userMedia (display puede no estar aún)
          ['userMedia', 'display'].forEach(k => window._espiaCliWH._crearBufferPara(k));
          _espiaCliWHIndicador();
        } catch(e) {
          console.warn('[espia WH] error iniciando:', e);
          await window._espiaCliWHCerrar('error_init');
        }
      };
      window._espiaCliWHCerrar = async function(motivo) {
        if (!window._espiaCliWH) return;
        console.log('[espia WH] cerrando:', motivo);
        const ref = window._espiaCliWH;
        window._espiaCliWH = null;
        // [v2.13.80] Limpiar TODOS los timers (sync, ICE flush, watchdog, reneg, legacy)
        try { clearInterval(ref.pollTimerSync); } catch(_){}
        try { clearInterval(ref._iceFlushTimer); } catch(_){}
        try { clearInterval(ref._iceWatchdogTimer); } catch(_){}
        // [v2.13.88] Quitar listener de visibilitychange para no leakear
        if (ref._handlerVisibility) {
          try { document.removeEventListener('visibilitychange', ref._handlerVisibility); } catch(_){}
        }
        try { clearInterval(ref.pollTimerOferta); } catch(_){}
        try { clearInterval(ref.pollTimerIce); } catch(_){}
        try { clearInterval(ref.pollTimerEstado); } catch(_){}
        try { clearTimeout(ref._renegTimeout); } catch(_){}
        // [v2.13.80] Flush final ICE pendientes (best-effort)
        if (ref._iceQueue?.length && ref.sesionId && ref.token) {
          try {
            await _espiaCliWHPost.call(null, 'espiaPushBatch', {
              sesionId: ref.sesionId, lado: 'device', token: ref.token,
              ice: ref._iceQueue.map(c => ({ ice: c }))
            });
          } catch(_){}
        }
        // [v2.13.80] AWAIT MediaRecorder stops para que el último chunk se suba.
        // Antes: rec.stop() es async (dispara onstop con el blob final), pero
        // pc.close() ocurría inmediato después → onstop interrumpido → último
        // 0-5min de video perdido. Ahora esperamos a que todos terminen.
        const stopPromises = Object.entries(ref.bufferRecorders || {}).map(([_, rec]) => {
          return new Promise(resolve => {
            if (!rec || rec.state !== 'recording') return resolve();
            const onStopOnce = () => resolve();
            rec.addEventListener('stop', onStopOnce, { once: true });
            try { rec.stop(); } catch(_) { resolve(); }
            // safety timeout: si por algún motivo onstop no dispara en 6s, seguir
            setTimeout(() => resolve(), 6000);
          });
        });
        await Promise.all(stopPromises);
        Object.values(ref.bufferTimers || {}).forEach(t => { try { clearInterval(t); } catch(_){} });
        Object.values(ref.streams || {}).forEach(s => {
          try { s?.getTracks().forEach(t => t.stop()); } catch(_){}
        });
        if (ref.gpsWatch) try { navigator.geolocation.clearWatch(ref.gpsWatch); } catch(_){}
        // Cerrar gpsCh explícito (libera SCTP)
        if (ref.gpsCh) try { ref.gpsCh.close(); } catch(_){}
        try { ref.pc?.close(); } catch(_){}
        if (ref.sesionId) {
          _espiaCliWHPost.call(null, 'espiaCerrarSesion', {
            sesionId: ref.sesionId, motivo: motivo || 'manual', lado: 'device', token: ref.token
          });
        }
        _espiaCliWHOcultarIndicador();
      };
      window.addEventListener('beforeunload', () => {
        const ref = window._espiaCliWH;
        if (!ref?.sesionId) return;
        // [v2.13.80] sendBeacon con Blob application/json para que el backend
        // doPost parsee correctamente. Antes sendBeacon de string crudo enviaba
        // Content-Type text/plain → en raras ocasiones Apps Script ignoraba el
        // body → cierre nunca se registraba en RTC_SIGNALING.
        try {
          const url = window.WH_CONFIG?.mosGasUrl || '';
          const payload = JSON.stringify({
            action: 'espiaCerrarSesion',
            sesionId: ref.sesionId,
            motivo: 'page_unload', lado: 'device',
            token: ref.token
          });
          const blob = new Blob([payload], { type: 'application/json' });
          navigator.sendBeacon(url, blob);
        } catch(_){}
      });

      // Handler de primer plano (la app está abierta).
      // 1) Comandos data-only (audio_start, audio_stop, gps_locate) → ejecutar silencioso.
      //    Sin este handler los pierde cuando la app está en foreground.
      // 2) Notificaciones visibles (title/body) → toast + notif.
      if (!_pushHandlerSet) {
        _pushHandlerSet = true;
        messaging.onMessage(async payload => {
          if (payload.data && payload.data.action) {
            const cmd = payload.data;
            console.log('[Push] comando foreground WH:', cmd.action);
            if (cmd.action === 'audio_start' && typeof window._audioRemotoIniciarWH === 'function') {
              window._audioRemotoIniciarWH(cmd.sesionId, parseInt(cmd.duracionMaxSeg, 10) || 1800);
            } else if (cmd.action === 'audio_stop' && typeof window._audioRemotoDetenerWH === 'function') {
              window._audioRemotoDetenerWH();
            } else if (cmd.action === 'gps_locate' && typeof window._gpsRegistrarWH === 'function') {
              window._gpsRegistrarWH(true);
            } else if (cmd.action === 'MOS_ESPIA_INICIAR' || cmd.idNotif === 'MOS_ESPIA_INICIAR') {
              // [v2.13.62] Iniciar cliente espía v2 WebRTC
              if (typeof _espiaCliWHIniciar === 'function') {
                _espiaCliWHIniciar(cmd.sesionId, cmd.masterId);
              }
            } else if (cmd.action === 'MOS_ESPIA_DETENER' || cmd.idNotif === 'MOS_ESPIA_DETENER') {
              if (typeof _espiaCliWHCerrar === 'function') {
                _espiaCliWHCerrar('master_push_stop');
              }
            }
            return;
          }
          const t = payload.notification?.title || '';
          const b = payload.notification?.body  || '';
          if (typeof toast === 'function') toast('🔔 ' + t + (b ? ': ' + b : ''), 'info', 8000);
          try {
            const reg = await navigator.serviceWorker.ready;
            reg.showNotification(t, {
              body: b,
              icon:  'https://levo19.github.io/MOS/icons/icon-192.png',
              badge: 'https://levo19.github.io/MOS/icons/icon-192.png',
              vibrate: [200, 100, 200],
              tag: 'wh-push-fg'
            });
          } catch(_) {}
        });
      }

      const permission = Notification.permission === 'default'
        ? await Notification.requestPermission()
        : Notification.permission;
      if (permission !== 'granted') {
        console.warn('[Push WH] permiso denegado:', permission);
        return;
      }

      const swReg = await navigator.serviceWorker.ready;
      const token = await messaging.getToken({ vapidKey: _PUSH_VAPID, serviceWorkerRegistration: swReg });
      if (!token) {
        console.warn('[Push WH] getToken devolvió vacío');
        return;
      }

      // Registrar token en MOS GAS asociado al operador + deviceId (targeted exacto)
      const usuario = (sesionActual.nombre + ' ' + (sesionActual.apellido || '')).trim();
      const _devIdToken = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
      // [v2.13.64] Si deviceId vacío, no registres (token sin device = inútil para targeted push)
      if (!_devIdToken) {
        console.warn('[Push WH] deviceId vacío — no registro hasta tenerlo');
        return;
      }
      try {
        await fetch(mosUrl, {
          method: 'POST',
          body: JSON.stringify({
            action: 'registrarPushToken',
            token, usuario,
            appOrigen: 'warehouseMos',
            deviceId: _devIdToken,
            dispositivo: 'warehouseMos · ' + (navigator.userAgent || '').substring(0, 80)
          })
        });
        console.log('[Push WH] token registrado en GAS ✅');
      } catch(eReg) {
        console.warn('[Push WH] registro POST falló:', eReg?.message);
      }
      // [v2.13.64] Auto-verify: 8s después, confirmar que el token está activo
      // en PUSH_TOKENS. Si no, retry una vez. Resuelve casos donde el registro
      // se perdió (red intermitente) sin que el user lo note.
      setTimeout(async () => {
        try {
          const verifUrl = mosUrl + '?action=verificarMiTokenRegistrado&deviceId=' + encodeURIComponent(_devIdToken);
          const vr = await fetch(verifUrl);
          const vj = await vr.json();
          const registrado = vj?.data?.registrado || vj?.registrado;
          if (!registrado) {
            console.warn('[Push WH] token NO está en backend tras 8s · reintento');
            if (typeof toast === 'function') toast('⚠ Re-registrando notificaciones…', 'warn', 4000);
            fetch(mosUrl, {
              method: 'POST',
              body: JSON.stringify({
                action: 'registrarPushToken',
                token, usuario,
                appOrigen: 'warehouseMos',
                deviceId: _devIdToken,
                dispositivo: 'warehouseMos · retry · ' + (navigator.userAgent || '').substring(0, 60)
              })
            }).catch(() => {});
          } else {
            console.log('[Push WH] token verificado ✅ (' + (vj?.data?.tokens || 1) + ' activo/s)');
          }
        } catch(eV) { console.warn('[Push WH] verify token falló:', eV?.message); }
      }, 8000);
    } catch(e) {
      console.warn('[Push WH] init error:', e?.message);
    }
  }
  window._pushInitWH = _pushInitWH;

  // Verificar cierre forzado cada minuto
  setInterval(_verificarCierreForzado, 60000);

  return { init, mostrarLogin,
           pinTecla, pinAtras, lockTecla, lockAtras,
           bloquear, confirmarCierre, cerrarTurnoFinal,
           cerrarSesionDesdeLock,
           getSesion,
           // Exportadas para que App.init las pueda invocar
           verificarDispositivo: _verificarDispositivoWH };
})();

// ════════════════════════════════════════════════
// BLOQUEO REMOTO — el admin desactiva al operador en MOS
// Polling 30s a getEstadoBloqueoUsuario; muestra overlay de
// candado y pide clave de admin para acceso temporal de 15 min.
// ════════════════════════════════════════════════
const BloqueoRemoto = (() => {
  let _pollTimer = null;
  let _countdownTimer = null;
  let _state = { bloqueado: false, unlockVigente: false, unlockHasta: 0, msRestantes: 0 };
  let _warningShown = false;
  let _activo = false;

  function _mosUrl() { return window.WH_CONFIG?.mosGasUrl || ''; }

  async function _check() {
    const ses = Session.getSesion();
    if (!ses || !ses.idPersonal) return;
    if (!_mosUrl() || !navigator.onLine) return;
    try {
      // Heartbeat para MOS: deviceId + nombre permiten que registrarSesionDispositivo
      // actualice Ultima_Conexion y Ultima_Sesion en DISPOSITIVOS. Sin esto los
      // dispositivos WH siempre muestran "hace Nh" porque jamás se les hacía heartbeat.
      const devId = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
      const nombreFull = ((ses.nombre || '') + ' ' + (ses.apellido || '')).trim();
      const _ua = (navigator.userAgent || '').substring(0, 200);
      const url = _mosUrl() + '?action=getEstadoBloqueoUsuario'
                + '&idPersonal=' + encodeURIComponent(ses.idPersonal)
                + '&nombre=' + encodeURIComponent(nombreFull)
                + '&deviceId=' + encodeURIComponent(devId)
                + '&userAgent=' + encodeURIComponent(_ua)
                + '&appOrigen=warehouseMos';
      const r = await fetch(url);
      const j = await r.json();
      if (!j || !j.ok || !j.data) return;
      const prev = _state;
      _state = {
        bloqueado: !!j.data.bloqueado,
        unlockVigente: !!j.data.unlockVigente,
        unlockHasta: parseInt(j.data.unlockHasta, 10) || 0,
        msRestantes: parseInt(j.data.msRestantes, 10) || 0,
        motivo: j.data.motivo || ''
      };
      // Transición: pasó a bloqueado → mostrar overlay
      if (!prev.bloqueado && _state.bloqueado) {
        _mostrarOverlay(ses);
      }
      // Transición: estaba bloqueado y ahora unlock vigente → ocultar
      if (prev.bloqueado && _state.unlockVigente) {
        _ocultarOverlay();
        _arrancarCountdown();
      }
      // Transición: reactivado en MOS (estado=1)
      if (prev.bloqueado && !_state.bloqueado && !_state.unlockVigente) {
        _ocultarOverlay();
        _ocultarBanner();
        _warningShown = false;
        toast('Cuenta reactivada por administrador', 'ok', 4000);
      }
    } catch(e) { /* tolerar */ }
  }

  function _mostrarOverlay(ses) {
    const el = document.getElementById('lockRemotoScreen');
    if (!el) return;
    const nombre = ses ? (ses.nombre + ' ' + (ses.apellido || '')).trim() : '';
    const elNm = document.getElementById('lockRemNombre');
    if (elNm) elNm.textContent = nombre;
    const elIn = document.getElementById('lockRemPin');
    if (elIn) elIn.value = '';
    const elErr = document.getElementById('lockRemError');
    if (elErr) elErr.textContent = '';
    el.style.display = 'flex';
    setTimeout(() => { if (elIn) elIn.focus(); }, 100);
  }

  function _ocultarOverlay() {
    const el = document.getElementById('lockRemotoScreen');
    if (el) el.style.display = 'none';
  }

  function _mostrarBanner(segs) {
    const el = document.getElementById('lockRemBanner');
    const txt = document.getElementById('lockRemBannerTxt');
    if (!el || !txt) return;
    const m = Math.floor(segs / 60);
    const s = segs % 60;
    const fmtTm = m + ':' + (s < 10 ? '0' : '') + s;
    txt.textContent = 'Acceso temporal · ' + fmtTm;
    el.style.display = 'block';
    if (segs <= 60) {
      el.style.background = 'rgba(239,68,68,0.95)';
      el.style.color = '#fff';
    } else {
      el.style.background = 'rgba(245,158,11,0.95)';
      el.style.color = '#0f172a';
    }
  }

  function _ocultarBanner() {
    const el = document.getElementById('lockRemBanner');
    if (el) el.style.display = 'none';
  }

  function _arrancarCountdown() {
    if (_countdownTimer) clearInterval(_countdownTimer);
    _warningShown = false;
    const tick = () => {
      const ms = (_state.unlockHasta || 0) - Date.now();
      if (ms <= 0) {
        if (_state.unlockVigente) {
          _state.unlockVigente = false;
          _state.bloqueado = true;
          _state.msRestantes = 0;
          _mostrarOverlay(Session.getSesion());
        }
        _ocultarBanner();
        clearInterval(_countdownTimer);
        _countdownTimer = null;
        return;
      }
      const segs = Math.floor(ms / 1000);
      _mostrarBanner(segs);
      if (ms <= 60000 && !_warningShown) {
        _warningShown = true;
        toast('⚠ Acceso temporal por vencer · queda menos de 1 min', 'warn', 5000);
      }
    };
    tick();
    _countdownTimer = setInterval(tick, 1000);
  }

  async function intentarDesbloqueo() {
    const elIn = document.getElementById('lockRemPin');
    const elErr = document.getElementById('lockRemError');
    const elBtn = document.getElementById('lockRemBtn');
    const clave = String(elIn?.value || '').trim();
    if (!clave) { if (elErr) elErr.textContent = 'Ingresa la clave de admin'; return; }
    const ses = Session.getSesion();
    if (!ses) return;
    if (elBtn) { elBtn.disabled = true; elBtn.textContent = 'VALIDANDO...'; }
    try {
      const r = await fetch(_mosUrl(), {
        method: 'POST',
        body: JSON.stringify({
          action: 'desbloquearUsuarioTemporal',
          idPersonal: ses.idPersonal,
          nombre: ses.nombre,
          appOrigen: 'warehouseMos',
          claveAdmin: clave
        })
      });
      const j = await r.json();
      if (!j || !j.ok) {
        if (elErr) elErr.textContent = (j && j.error) || 'Error de conexión';
        return;
      }
      if (!j.data?.autorizado) {
        if (elErr) elErr.textContent = j.data?.error || 'Clave incorrecta';
        if (elIn) elIn.value = '';
        return;
      }
      _state = {
        bloqueado: false,
        unlockVigente: true,
        unlockHasta: parseInt(j.data.unlockHasta, 10),
        msRestantes: parseInt(j.data.msRestantes, 10) || (15 * 60 * 1000),
        motivo: ''
      };
      _ocultarOverlay();
      _arrancarCountdown();
      toast('Acceso otorgado · 15 minutos · validado por ' + (j.data.validadoPor || 'admin'), 'ok', 4000);
    } catch(e) {
      if (elErr) elErr.textContent = 'Sin conexión con MOS';
    } finally {
      if (elBtn) { elBtn.disabled = false; elBtn.textContent = 'DESBLOQUEAR 15 MIN'; }
    }
  }

  function iniciar() {
    if (_activo) return;
    _activo = true;
    setTimeout(_check, 4000);
    _pollTimer = setInterval(_check, 30 * 1000);
  }

  function detener() {
    _activo = false;
    if (_pollTimer) { clearInterval(_pollTimer); _pollTimer = null; }
    if (_countdownTimer) { clearInterval(_countdownTimer); _countdownTimer = null; }
    _ocultarOverlay();
    _ocultarBanner();
    _state = { bloqueado: false, unlockVigente: false, unlockHasta: 0, msRestantes: 0 };
    _warningShown = false;
  }

  return { iniciar, detener, intentarDesbloqueo };
})();
window.BloqueoRemoto = BloqueoRemoto;

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

  // ── Alertas de cuadre stock (auditoría diaria 22:00) ──────
  let _alertasCache = [];

  async function cargarAlertasStock() {
    try {
      const res = await API.getAlertasStock({ soloPendientes: true });
      if (!res.ok) return;
      _alertasCache = res.data || [];
      _renderAlertasStock();
    } catch(e) { /* silencioso, sin red */ }
  }

  function _renderAlertasStock() {
    const card = document.getElementById('cardAlertasStock');
    const list = document.getElementById('listAlertasStock');
    const cnt  = document.getElementById('alertasStockCount');
    const btn  = document.getElementById('btnVerTodasAlertas');
    if (!card || !list) return;
    if (!_alertasCache.length) { card.style.display = 'none'; return; }
    card.style.display = '';
    if (cnt) cnt.textContent = `${_alertasCache.length} producto${_alertasCache.length !== 1 ? 's' : ''}`;
    // Top 3 con mayor diferencia absoluta
    const top = [..._alertasCache].sort((a, b) => Math.abs(b.diferencia) - Math.abs(a.diferencia)).slice(0, 3);
    list.innerHTML = top.map(a => `
      <div class="flex items-center justify-between text-xs py-1 border-b border-red-900/30">
        <div class="flex-1 min-w-0">
          <p class="text-slate-200 font-semibold truncate">${escHtml(a.descripcion)}</p>
          <p class="text-slate-500 font-mono text-[10px]">${escHtml(a.codigoProducto)}</p>
        </div>
        <div class="text-right flex-shrink-0 ml-2">
          <span class="font-bold ${a.diferencia > 0 ? 'text-amber-400' : 'text-red-400'}">${a.diferencia > 0 ? '+' : ''}${fmt(a.diferencia, 1)}</span>
        </div>
      </div>`).join('');
    if (btn) btn.style.display = _alertasCache.length > 3 ? '' : 'none';
  }

  // Estado del modal de alertas: filtro y búsqueda
  let _alertaFiltro = 'todas';   // 'todas' | 'salidas' | 'entradas'
  let _alertaQuery  = '';

  function verTodasAlertas() {
    _alertaFiltro = 'todas';
    _alertaQuery  = '';
    const inp = document.getElementById('alertaSearchInput');
    if (inp) inp.value = '';
    const sub = document.getElementById('alertasStockSub');
    const fechaUlt = _alertasCache[0]?.fecha;
    if (sub) sub.textContent = fechaUlt ? `Última auditoría: ${new Date(fechaUlt).toLocaleString('es-PE')}` : '';
    _renderAlertasFiltradas();
    abrirSheet('sheetAlertasStock');
  }

  function setAlertaFiltro(filtro) {
    _alertaFiltro = filtro;
    _renderAlertasFiltradas();
  }

  function buscarAlerta(q) {
    _alertaQuery = String(q || '').toLowerCase().trim();
    _renderAlertasFiltradas();
  }

  function _renderAlertasFiltradas() {
    const list = document.getElementById('listAlertasStockFull');
    if (!list) return;

    // Conteos por categoría (para los chips)
    const totalSalidas  = _alertasCache.filter(a => a.diferencia < 0).length;
    const totalEntradas = _alertasCache.filter(a => a.diferencia > 0).length;
    const total         = _alertasCache.length;
    const cT  = document.getElementById('chipAlertasTodas');
    const cS  = document.getElementById('chipAlertasSalidas');
    const cE  = document.getElementById('chipAlertasEntradas');
    if (cT) cT.textContent = `Todas (${total})`;
    if (cS) cS.textContent = `Salidas fantasma (${totalSalidas})`;
    if (cE) cE.textContent = `Entradas fantasma (${totalEntradas})`;
    [cT, cS, cE].forEach(c => c && c.classList.remove('chip-active'));
    const activo = _alertaFiltro === 'salidas' ? cS : (_alertaFiltro === 'entradas' ? cE : cT);
    if (activo) activo.classList.add('chip-active');

    // Filtrar
    let lista = [..._alertasCache];
    if (_alertaFiltro === 'salidas')  lista = lista.filter(a => a.diferencia < 0);
    if (_alertaFiltro === 'entradas') lista = lista.filter(a => a.diferencia > 0);
    if (_alertaQuery) {
      lista = lista.filter(a =>
        String(a.descripcion).toLowerCase().includes(_alertaQuery) ||
        String(a.codigoProducto).toLowerCase().includes(_alertaQuery)
      );
    }
    lista.sort((a, b) => Math.abs(b.diferencia) - Math.abs(a.diferencia));

    if (!lista.length) {
      list.innerHTML = '<p class="text-xs text-slate-500 text-center py-6">Sin alertas en esta categoría</p>';
      return;
    }

    list.innerHTML = lista.map(a => {
      const safeId = escAttr(a.idAlerta);
      const safeCb = escAttr(a.codigoProducto);
      const safeNm = escAttr(a.descripcion);
      const teor   = a.stockTeorico;
      const explica = a.diferencia > 0
        ? `Stock <b>+${fmt(a.diferencia,1)}</b> mayor a lo esperado — entradas no registradas o salidas anuladas tras cerrar guía`
        : `Stock <b>${fmt(a.diferencia,1)}</b> menor a lo esperado — salidas no registradas o cierres duplicados`;
      return `
      <div class="card-sm" style="border:1px solid rgba(239,68,68,.3);padding:10px 12px">
        <div class="flex items-start justify-between gap-2 mb-1.5">
          <div class="flex-1 min-w-0">
            <p class="font-bold text-sm text-slate-100 truncate">${escHtml(a.descripcion)}</p>
            <p class="text-[10px] text-slate-500 font-mono">${escHtml(a.codigoProducto)}</p>
          </div>
          <span class="font-black text-base ${a.diferencia > 0 ? 'text-amber-400' : 'text-red-400'} flex-shrink-0">${a.diferencia > 0 ? '+' : ''}${fmt(a.diferencia,1)}</span>
        </div>
        <div class="grid grid-cols-2 gap-1 text-[11px] mb-2">
          <div><span class="text-slate-500">Stock real:</span> <b class="text-slate-200">${fmt(a.stockReal,1)}</b></div>
          <div><span class="text-slate-500">Teórico:</span> <b class="text-slate-200">${fmt(a.stockTeorico,1)}</b></div>
        </div>
        <p class="text-[11px] text-slate-400 leading-tight mb-2">${explica}</p>
        <div class="flex gap-1.5 flex-wrap">
          <button onclick="Dashboard.auditarDesdeAlerta('${safeId}','${safeCb}','${safeNm}',${teor})"
                  class="text-[11px] px-2.5 py-1 rounded font-bold"
                  style="background:rgba(14,165,233,.18);color:#38bdf8;border:1px solid rgba(14,165,233,.4)">
            📋 Auditar
          </button>
          <button onclick="Dashboard.aceptarTeorico('${safeId}')"
                  class="text-[11px] px-2.5 py-1 rounded font-bold"
                  style="background:rgba(16,185,129,.18);color:#34d399;border:1px solid rgba(16,185,129,.4)">
            ⚡ Aceptar teórico
          </button>
          <button onclick="Dashboard.marcarAlertaRevisada('${safeId}')"
                  class="text-[11px] px-2.5 py-1 rounded font-bold"
                  style="background:rgba(100,116,139,.18);color:#94a3b8;border:1px solid rgba(100,116,139,.3)">
            ✓ Solo revisar
          </button>
        </div>
      </div>`;
    }).join('');
  }

  function auditarDesdeAlerta(idAlerta, codigoBarra, nombre, stockTeorico) {
    cerrarSheet('sheetAlertasStock');
    setTimeout(() => {
      // Buscar skuBase del producto en cache
      const prods = OfflineManager.getProductosCache();
      const p = prods.find(x =>
        String(x.codigoBarra || '') === codigoBarra ||
        String(x.idProducto  || '') === codigoBarra ||
        String(x.skuBase     || '') === codigoBarra
      );
      const skuBase = p?.skuBase || p?.idProducto || codigoBarra;
      ProductosView.abrirAuditBarcode(codigoBarra, nombre, skuBase, {
        prefillFisico: stockTeorico,
        idAlerta
      });
    }, 300);
  }

  async function aceptarTeorico(idAlerta) {
    if (!await _whConfirm('¿Aplicar el stock teórico?\n\nSe creará un ajuste automático que iguala el stock real al teórico.', { warning: true, titulo: 'Aplicar stock teórico', okText: 'Aplicar' })) return;
    try {
      const res = await API.aceptarTeoricoAlerta({
        idAlerta,
        usuario: window.WH_CONFIG?.usuario || ''
      });
      if (res.ok) {
        const aj = res.data?.ajusteAplicado || 0;
        toast(`✓ Stock corregido${aj !== 0 ? ` (ajuste ${aj > 0 ? '+' : ''}${fmt(aj, 1)})` : ''}`, 'ok', 3000);
        // Quitar de la lista local y re-render
        _alertasCache = _alertasCache.filter(a => a.idAlerta !== idAlerta);
        _renderAlertasStock();
        _renderAlertasFiltradas();
      } else {
        toast('Error: ' + (res.error || 'No se pudo aplicar'), 'danger', 5000);
      }
    } catch(e) { toast('Sin conexión', 'warn', 3000); }
  }

  async function marcarAlertaRevisada(idAlerta) {
    try {
      const res = await API.marcarAlertaRevisada({ idAlerta });
      if (res.ok) {
        _alertasCache = _alertasCache.filter(a => a.idAlerta !== idAlerta);
        _renderAlertasStock();
        verTodasAlertas();
        toast('Alerta marcada como revisada', 'ok', 2000);
      } else {
        toast('Error: ' + (res.error || 'sin respuesta'), 'danger');
      }
    } catch(e) { toast('Sin conexión', 'warn'); }
  }

  return { toggle, cargarAlertasStock, verTodasAlertas, marcarAlertaRevisada,
           setAlertaFiltro, buscarAlerta, auditarDesdeAlerta, aceptarTeorico };
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

  async function init() {
    // Restaurar GAS URL si fue guardada localmente
    const gasUrl = localStorage.getItem('wh_gas_url');
    if (gasUrl) {
      console.log('[App] wh_gas_url desde localStorage:', gasUrl);
      window.WH_CONFIG.gasUrl = gasUrl;
    }
    console.log('[App] GAS URL activa:', window.WH_CONFIG.gasUrl);

    // [v2.13.118] Inicializar sistema centralizado de membretes/adhesivos.
    // WH API.post retorna {ok,data} → no desempaca. Llamadas a endpoints
    // del propio WH no llevan prefijo wh_.
    try {
      if (window.MembreteSystem && window.MembreteSystem.iniciar) {
        window.MembreteSystem.iniciar({
          apiPost:        function(action, params) {
            // [v2.13.126 FIX] API.post espera (action_string, params_object) como
            // 2 args separados. Antes pasaba 1 objeto → backend recibía
            // action=[object Object] → "Acción no reconocida".
            // [v2.13.154 DEFENSA] Validar que action no esté vacía antes de llamar.
            // Si llega vacía, log + rechazar — el endpoint vacío genera el error
            // "Acción no reconocida: " (sin nombre) en backend.
            if (!action || typeof action !== 'string') {
              console.warn('[WH apiPost] action inválida:', action, 'params:', params);
              return Promise.reject(new Error('action vacía pasada a apiPost'));
            }
            return API.post(action, params || {});
          },
          usuario:        function() { return (window.WH_CONFIG && WH_CONFIG.usuario) || ''; },
          origen:         'WH',
          unwrapData:     false,
          endpointPrefix: ''
        });
      }
    } catch(_) {}

    // [v2.13.121] Sistema centralizado de seguridad: widget + modal acceso/horario.
    // SeguridadSystem espera apiPost que devuelva data desempaquetada (unwrap=true).
    try {
      if (window.SeguridadSystem && window.SeguridadSystem.iniciar) {
        window.SeguridadSystem.iniciar({
          apiPost: function(action, params) {
            // [v2.13.126 FIX] API.post espera 2 args separados, no 1 objeto.
            return API.post(action, params || {})
              .then(function(r) { return r && r.data !== undefined ? r.data : r; });
          },
          usuario:        function() { return (window.WH_CONFIG && WH_CONFIG.usuario) || ''; },
          rol:            function() { return (window.WH_CONFIG && WH_CONFIG.rol) || ''; },
          idPersonal:     function() { return (window.WH_CONFIG && WH_CONFIG.idPersonal) || ''; },
          app:            'warehouseMos',
          unwrapData:     true,
          endpointPrefix: ''
        });
      }
    } catch(_) {}

    // ── Verificación de dispositivo BLOQUEANTE (igual que ME) ──
    // El overlay #verifDispScreen ya está visible desde el HTML inicial (z-index
    // 9998), así que la app NUNCA se ve hasta que el dispositivo esté ACTIVO.
    // Si el resultado NO es ACTIVO, no continuamos con Session.init() —
    // dejamos el candado puesto. El flujo continúa por:
    //   - botón "Solicitar acceso" → polling cada 15s
    //   - botón "Reintentar" (error de red) → reintenta y llama Session.init()
    //   - aprobación remota desde MOS → polling detecta ACTIVO y oculta candado,
    //     pero el resto de init no corrió: el operador debe recargar la pestaña.
    const verifResult = await Session.verificarDispositivo();
    if (verifResult !== 'ACTIVO') return;

    // Multi-dispositivo: al volver a foreground (cambio de pestaña / unlock),
    // refrescar datos operacionales para ver cambios de otros dispositivos
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState !== 'visible') return;
      if (typeof OfflineManager === 'undefined') return;
      // El throttle interno de 15s previene spam
      OfflineManager.precargarOperacional(false).catch(() => {});
    });

    // Indicador de cola de operaciones serializadas (para que el operador vea
    // que el sistema está procesando — evita el "no respondió, vuelvo a tocar")
    window.addEventListener('wh:opqueue', (e) => {
      const ind = document.getElementById('opQueueIndicator');
      if (!ind) return;
      const n = e.detail?.count || 0;
      if (n === 0) { ind.style.display = 'none'; return; }
      ind.style.display = 'inline-block';
      ind.textContent = `📤 ${n} en cola`;
    });

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
      // [Fix v2.9.1] Cuando un envasado offline-encolado se sincroniza,
      // recargar la lista para reemplazar el ENV_OPT_* por el id real.
      if (currentView === 'envasados' && changed.includes('envasados'))   EnvasadosView.cargar();
    });

    // Pull-to-refresh en la vista principal — también dispara OpLog.flush()
    // para que el operador pueda forzar reconciliación de ops pendientes.
    const mainContent = document.getElementById('mainContent');
    PullToRefresh.init(mainContent, () => {
      if (window.OpLog && typeof OpLog.flush === 'function') OpLog.flush();
      if (window.Mermas && typeof Mermas.refreshBadge === 'function') Mermas.refreshBadge();
      if (currentView === 'guias')        GuiasView.cargar();
      else if (currentView === 'productos')    ProductosView.cargar();
      else if (currentView === 'dashboard')    cargarDashboard();
      else if (currentView === 'preingresos')  PreingresosView.cargar?.();
    });

    // F10: modo nocturno auto vía prefers-color-scheme (default ya dark)
    // y bonus: refrescar badge mermas al cargar la app
    setTimeout(() => {
      if (window.Mermas && typeof Mermas.refreshBadge === 'function') Mermas.refreshBadge();
    }, 1500);

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

    // Pausar cámaras al salir de su vista
    if (currentView === 'despacho' && viewName !== 'despacho') {
      DespachoView.pauseCamera();
    }
    if (currentView === 'productos' && viewName !== 'productos') {
      ProductosView.cerrarProdCamara();
    }

    closeUserMenu();
    // [v2.13.189] Backstop: navegar siempre libera el scroll del fondo (un sheet
    // de la vista anterior no debe dejar el body bloqueado en la nueva vista).
    _resetSheetScrollLock();

    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const el = document.getElementById('view-' + viewName);
    if (el) { el.classList.add('active'); el.classList.add('slide-up'); }

    // Marcar botón de nav activo por data-view
    document.querySelectorAll('.nav-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.view === viewName);
    });
    // F-nav-v3: mover la píldora al botón activo
    _moverNavPill();

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
      case 'logs':        LogsView.cargar(); break;
      case 'diagnostico': DiagnosticoView.cargar(); break;
    }

    // Flotante del despacho: mostrar si hay despacho activo y NO estás en
    // la vista despacho; ocultar al entrar a despacho.
    try { DespachoView.renderFlotante(); } catch(_){}
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

    // [v2.13.121] Widget "Mi horario" — pulse glow + countdown
    try {
      if (window.SeguridadSystem && window.SeguridadSystem.arrancarWidgetMiHorario) {
        window.SeguridadSystem.arrancarWidgetMiHorario('segWidgetHorario');
      }
    } catch(_) {}

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

    // Alertas de cuadre stock (en background, no bloquea render)
    Dashboard.cargarAlertasStock();

    // Logo alert dot (topbar + sidebar)
    const totalAlertas = contadores.alertasTotal ?? 0;
    document.getElementById('logoAlertDot')?.classList.toggle('hidden', totalAlertas === 0);
    const sideLogoAlert = document.getElementById('sideLogoAlertDot');
    const sideArrow     = document.getElementById('sideAlertArrow');
    if (sideLogoAlert) sideLogoAlert.classList.toggle('visible', totalAlertas > 0);
    if (sideArrow)     sideArrow.classList.toggle('visible',     totalAlertas > 0);

    // Badges en nav: guías abiertas
    const guiasAbiertas = (alertas.guiasAbiertasTardias?.length ?? 0);
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

    // Productos nuevos aprobados (últimos 3 días) — carga en background
    _cargarPNAprobados();

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

  async function showUsuarioDialog() {
    const nombre = await _whPrompt('Nombre del operador:', window.WH_CONFIG.usuario, { titulo: 'Cambiar operador', maxlength: 60 });
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

  // ════════════════════════════════════════════════
  // Bottom nav v3 — Pill flotante + long-press menú + sonido
  // ════════════════════════════════════════════════
  function _moverNavPill() {
    const pill = document.getElementById('navPill');
    if (!pill) return;
    const bottomNav = document.getElementById('bottomNav');
    // En modo sidebar (tablet/desktop), el bottom-nav está oculto → no mover pill
    if (!bottomNav || getComputedStyle(bottomNav).display === 'none') {
      pill.classList.remove('ready');
      return;
    }
    const active = document.querySelector('#bottomNav .nav-btn-v3.active');
    if (!active) { pill.classList.remove('ready'); return; }
    const row = active.parentElement;
    const rb = row.getBoundingClientRect();
    const ab = active.getBoundingClientRect();
    pill.style.left  = (ab.left - rb.left) + 'px';
    pill.style.width = ab.width + 'px';
    pill.classList.add('ready');
  }
  window.addEventListener('resize', () => _moverNavPill());
  window.addEventListener('orientationchange', () => setTimeout(_moverNavPill, 250));

  // Long-press menú con sub-acciones por tab
  const _LP_ACCIONES = {
    dashboard: {
      titulo: '📊 Dashboard · acción rápida',
      items: [
        { ico:'⚠', tit:'Ver alertas',     sub:'Stock bajo + diferencias', act:() => { nav('dashboard'); if (typeof Dashboard !== 'undefined' && Dashboard.verTodasAlertas) Dashboard.verTodasAlertas(); } },
        { ico:'🔄', tit:'Forzar sync',     sub:'Recargar datos del server', act:() => { OfflineManager.precargarOperacional?.(true); if (typeof toast === 'function') toast('Sincronizando…', 'info', 1500); } },
        { ico:'🏠', tit:'Ir al inicio',    sub:'Cargar Dashboard',           act:() => nav('dashboard') }
      ]
    },
    guias: {
      titulo: '📋 Guías · acción rápida',
      items: [
        { ico:'📋', tit:'Nueva guía',         sub:'Ingreso o salida',          act:() => { nav('guias'); setTimeout(() => abrirTypePicker(), 180); } },
        { ico:'📥', tit:'Nuevo preingreso',   sub:'Recepción de proveedor',    act:() => { nav('preingresos'); setTimeout(() => { if (window.PreingresosView && PreingresosView.nuevo) PreingresosView.nuevo(); }, 200); } },
        { ico:'🛺', tit:'Cargadores del día', sub:'Sumar al resumen',          act:() => { if (window.Cargadores) Cargadores.abrir(); } },
        { ico:'🗑', tit:'Cesta de mermas',    sub:'Pendientes + procesar',     act:() => { if (window.Mermas) Mermas.abrirCesta(); } }
      ]
    },
    productos: {
      titulo: '📦 Productos · acción rápida',
      items: [
        { ico:'⚠', tit:'Bajo mínimo',         sub:'Productos críticos',        act:() => { nav('productos'); if (window.ProductosView && ProductosView.toggleFiltro) setTimeout(() => ProductosView.toggleFiltro('bajo'), 250); } },
        { ico:'📅', tit:'Por vencer',         sub:'Lotes próximos a expirar',  act:() => { nav('productos'); if (window.ProductosView && ProductosView.toggleFiltro) setTimeout(() => ProductosView.toggleFiltro('porVencer'), 250); } },
        { ico:'🕵', tit:'Modo auditoría',     sub:'Verificar conteo físico',   act:() => { nav('productos'); if (window.ProductosView && ProductosView.toggleAuditoriaDia) setTimeout(() => ProductosView.toggleAuditoriaDia(), 250); } }
      ]
    }
  };

  function _abrirLpNavMenu(navKey) {
    const cfg = _LP_ACCIONES[navKey];
    if (!cfg) return;
    document.getElementById('navLpTitle').textContent = cfg.titulo;
    const list = document.getElementById('navLpList');
    list.innerHTML = '';
    cfg.items.forEach((it, idx) => {
      const btn = document.createElement('button');
      btn.className = 'act-sheet-item';
      btn.innerHTML = `<span class="act-sheet-ico">${it.ico}</span>
        <span><span class="act-sheet-tit">${it.tit}</span>
          <span class="act-sheet-sub">${it.sub}</span></span>`;
      btn.onclick = () => { cerrarLpMenu(); try { it.act(); } catch(e){ console.warn(e); } };
      list.appendChild(btn);
    });
    document.getElementById('overlayNavLp').style.display = 'block';
    document.getElementById('sheetNavLp').classList.add('open');
    if (navigator.vibrate) navigator.vibrate(15);
    if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click();
  }
  function cerrarLpMenu() {
    document.getElementById('overlayNavLp').style.display = 'none';
    document.getElementById('sheetNavLp').classList.remove('open');
  }

  // Bind long-press + tap-ripple + sound on todos los .nav-btn-v3
  function _bindNavV3() {
    document.querySelectorAll('.nav-btn-v3').forEach(btn => {
      if (btn._v3Bound) return;
      btn._v3Bound = true;
      let lpTimer = null;
      let lpFired = false;
      const startLp = (x, y) => {
        lpFired = false;
        lpTimer = setTimeout(() => {
          lpFired = true;
          const key = btn.dataset.lp;
          if (key) _abrirLpNavMenu(key);
        }, 550);
        // Ripple visual
        const rip = document.createElement('span');
        rip.className = 'nav-ripple';
        const rect = btn.getBoundingClientRect();
        rip.style.left = (x - rect.left) + 'px';
        rip.style.top  = (y - rect.top) + 'px';
        rip.style.width = rip.style.height = '40px';
        btn.appendChild(rip);
        setTimeout(() => rip.remove(), 600);
      };
      const cancelLp = () => { if (lpTimer) { clearTimeout(lpTimer); lpTimer = null; } };
      btn.addEventListener('touchstart', e => {
        const t = e.touches[0];
        startLp(t.clientX, t.clientY);
      }, { passive: true });
      btn.addEventListener('touchend',    cancelLp);
      btn.addEventListener('touchcancel', cancelLp);
      btn.addEventListener('touchmove',   cancelLp);
      btn.addEventListener('mousedown',   e => startLp(e.clientX, e.clientY));
      btn.addEventListener('mouseup',     cancelLp);
      btn.addEventListener('mouseleave',  cancelLp);

      // Bloquear el click si long-press disparó
      btn.addEventListener('click', e => {
        if (lpFired) { e.preventDefault(); e.stopImmediatePropagation(); lpFired = false; return; }
        if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click();
        if (navigator.vibrate) navigator.vibrate(8);
      }, true);
    });
  }
  // Logo WH: tap → animación expandir "WareHouse" temporalmente (decorativo)
  function _whLogoEfecto() {
    const logo = document.getElementById('whLogo');
    if (!logo) return;
    logo.classList.add('wh-expanded');
    if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click();
    if (navigator.vibrate) navigator.vibrate(8);
    setTimeout(() => logo.classList.remove('wh-expanded'), 1600);
  }

  // ════════════════════════════════════════════════
  // F2-F7 — Action sheet + Type picker + Subtype picker + Entidad picker
  // ════════════════════════════════════════════════
  let _tpState = { direccion: '', subtipo: '', entidad: null };
  let _entidadPickerState = { tipo: '', lista: [], onPick: null };
  let _reabrirCtx = null;          // {idGuia} pendiente de clave admin
  const ADMIN_REMEMBER_KEY = 'wh_admin_remember_until';

  function abrirActionDia() {
    document.getElementById('overlayActionDia').style.display = 'block';
    document.getElementById('sheetActionDia').classList.add('open');
  }
  function cerrarActionDia() {
    document.getElementById('overlayActionDia').style.display = 'none';
    document.getElementById('sheetActionDia').classList.remove('open');
  }

  function abrirTypePicker() {
    _tpState = { direccion: '', subtipo: '', entidad: null };
    document.getElementById('overlayTypePicker').style.display = 'block';
    document.getElementById('sheetTypePicker').classList.add('open');
  }
  function cerrarTypePicker() {
    document.getElementById('overlayTypePicker').style.display = 'none';
    document.getElementById('sheetTypePicker').classList.remove('open');
  }
  function volverTypePicker() {
    cerrarSubtypePicker();
    abrirTypePicker();
  }

  function elegirDireccion(dir) {
    _tpState.direccion = dir;
    cerrarTypePicker();
    abrirSubtypePicker(dir);
  }

  function abrirSubtypePicker(dir) {
    const title = dir === 'INGRESO' ? '↓ INGRESO · ¿de dónde viene?' : '↑ SALIDA · ¿a dónde va?';
    const opts = dir === 'INGRESO'
      ? [
          { tipo: 'INGRESO_PROVEEDOR',  ico: '🛺', titulo: 'De PROVEEDOR',         desc: 'Compra normal · catálogo proveedores' },
          { tipo: 'INGRESO_JEFATURA',   ico: '👤', titulo: 'De JEFATURA',          desc: 'Devolución admin · admin+master MOS' },
          { tipo: 'INGRESO_DEVOLUCION_ZONA', ico: '📍', titulo: 'De ZONA (devolución)', desc: 'Material que regresa de una zona' }
        ]
      : [
          { tipo: 'SALIDA_ZONA',        ico: '📍', titulo: 'A ZONA (despacho)', desc: 'Lo más frecuente · catálogo de zonas' },
          { tipo: 'SALIDA_JEFATURA',    ico: '👤', titulo: 'A JEFATURA',        desc: 'Entrega a admin · admin+master MOS' },
          { tipo: 'SALIDA_DEVOLUCION',  ico: '🛺', titulo: 'Devol. a PROVEEDOR', desc: 'Casos especiales · catálogo proveedores' }
        ];
    document.getElementById('subtypeTitle').textContent = title;
    const html = opts.map(o => `
      <button class="act-sheet-item" onclick="App.elegirSubtipo('${o.tipo}')">
        <span class="act-sheet-ico">${o.ico}</span>
        <span><span class="act-sheet-tit">${o.titulo}</span>
          <span class="act-sheet-sub">${o.desc}</span></span>
      </button>`).join('');
    document.getElementById('subtypeOptions').innerHTML = html;
    document.getElementById('overlaySubtypePicker').style.display = 'block';
    document.getElementById('sheetSubtypePicker').classList.add('open');
  }
  function cerrarSubtypePicker() {
    document.getElementById('overlaySubtypePicker').style.display = 'none';
    document.getElementById('sheetSubtypePicker').classList.remove('open');
  }

  function elegirSubtipo(subtipo) {
    _tpState.subtipo = subtipo;
    cerrarSubtypePicker();

    // Tipos que requieren entidad
    if (/PROVEEDOR/.test(subtipo)) {
      abrirEntidadPicker('proveedor', (ent) => _completarNuevaGuia(subtipo, ent));
    } else if (/JEFATURA/.test(subtipo)) {
      abrirEntidadPicker('jefatura', (ent) => _completarNuevaGuia(subtipo, ent));
    } else if (/ZONA|DEVOLUCION/.test(subtipo)) {
      abrirEntidadPicker('zona', (ent) => _completarNuevaGuia(subtipo, ent));
    } else {
      _completarNuevaGuia(subtipo, null);
    }
  }

  function _completarNuevaGuia(subtipo, entidad) {
    if (window.GuiasView && GuiasView.crearConTipo) {
      GuiasView.crearConTipo(subtipo, entidad);
    } else {
      // Fallback: abrir el sheetGuia legacy con tipo pre-seleccionado
      const sel = document.getElementById('guiaTipo');
      if (sel) sel.value = subtipo;
      abrirSheet('sheetGuia');
    }
  }

  function abrirEntidadPicker(tipo, onPick) {
    _entidadPickerState = { tipo, lista: [], onPick };
    const titles = { proveedor: 'Elegir proveedor', jefatura: 'Elegir jefatura', zona: 'Elegir zona' };
    document.getElementById('entidadPickerTitle').textContent = titles[tipo] || 'Elegir';
    document.getElementById('entidadPickerSearch').value = '';

    let lista = [];
    if (tipo === 'proveedor') {
      lista = (OfflineManager.getProveedoresCache() || []).filter(p => {
        const n = String(p.nombre || '').trim().toUpperCase();
        const activo = String(p.estado || '') === '1';
        return activo && n.indexOf('CARGADOR') !== 0;
      });
    } else if (tipo === 'jefatura') {
      lista = (OfflineManager.getPersonalCache() || []).filter(p => {
        const rol = String(p.rol || '').toUpperCase();
        return rol === 'ADMIN' || rol === 'MASTER';
      });
    } else if (tipo === 'zona') {
      lista = OfflineManager.getZonasCache?.() || [];
    }
    _entidadPickerState.lista = lista;
    _renderEntidadPicker(lista);
    document.getElementById('overlayEntidadPicker').style.display = 'block';
    document.getElementById('sheetEntidadPicker').classList.add('open');
  }
  function cerrarEntidadPicker() {
    document.getElementById('overlayEntidadPicker').style.display = 'none';
    document.getElementById('sheetEntidadPicker').classList.remove('open');
  }
  function _renderEntidadPicker(list) {
    const cont = document.getElementById('entidadPickerList');
    if (!list.length) { cont.innerHTML = '<p style="color:#64748b;font-size:13px;text-align:center;padding:18px 0">Sin coincidencias</p>'; return; }
    const tipo = _entidadPickerState.tipo;
    const nameOf = e => tipo === 'jefatura'
      ? `${e.nombre || ''} ${e.apellido || ''}`.trim()
      : (e.nombre || e.descripcion || e.idZona || e.idProveedor || '');
    const idOf = e => e.idProveedor || e.idPersonal || e.idZona || '';
    cont.innerHTML = list.slice(0, 50).map(e => `
      <div class="ep-item" onclick="App._pickEntidad('${(idOf(e)||'').replace(/'/g,'&#39;')}')">
        <div style="flex:1">
          <div class="ep-item-name">${nameOf(e)}</div>
        </div>
      </div>`).join('');
  }
  function filtrarEntidad(q) {
    q = String(q || '').toLowerCase().trim();
    const all = _entidadPickerState.lista;
    if (!q) return _renderEntidadPicker(all);
    const out = all.filter(e => {
      const txt = (e.nombre || e.descripcion || e.apellido || '').toLowerCase();
      return txt.includes(q);
    });
    _renderEntidadPicker(out);
  }
  function _pickEntidad(id) {
    const list = _entidadPickerState.lista;
    const ent = list.find(e => (e.idProveedor || e.idPersonal || e.idZona) === id);
    cerrarEntidadPicker();
    if (typeof _entidadPickerState.onPick === 'function') {
      _entidadPickerState.onPick(ent);
    }
  }
  function dictarEntidad() {
    if (!window.Voice || !Voice.supported().stt) {
      if (typeof toast === 'function') toast('Dictado por voz no disponible', 'warn');
      return;
    }
    const inp = document.getElementById('entidadPickerSearch');
    const mic = document.getElementById('entidadPickerMic');
    if (mic) mic.classList.add('listening');
    Voice.listen({
      onResult: (txt, final) => { if (inp) { inp.value = txt; filtrarEntidad(txt); } },
      onEnd:    () => { if (mic) mic.classList.remove('listening'); }
    });
  }

  function dictarCargador() {
    if (!window.Voice || !Voice.supported().stt) return;
    const inp = document.getElementById('cargBuscarInput');
    const mic = document.getElementById('cargMic');
    if (mic) mic.classList.add('listening');
    Voice.listen({
      onResult: (txt) => { if (inp) { inp.value = txt; if (window.Cargadores) Cargadores._filtrar(txt); } },
      onEnd:    () => { if (mic) mic.classList.remove('listening'); }
    });
  }

  function dictarBuscarGuia() {
    if (!window.Voice || !Voice.supported().stt) {
      if (typeof toast === 'function') toast('Dictado por voz no disponible', 'warn');
      return;
    }
    const inp = document.getElementById('inputBuscarGuia');
    const mic = document.getElementById('micBuscarGuia');
    if (mic) mic.classList.add('listening');
    Voice.listen({
      onResult: (txt) => {
        if (inp) {
          inp.value = txt;
          if (window.GuiasView) GuiasView.buscar(txt);
        }
      },
      onEnd:    () => { if (mic) mic.classList.remove('listening'); },
      onError:  () => { if (mic) mic.classList.remove('listening'); }
    });
  }

  function dictarBuscarPre() {
    if (!window.Voice || !Voice.supported().stt) {
      if (typeof toast === 'function') toast('Dictado por voz no disponible', 'warn');
      return;
    }
    const inp = document.getElementById('inputBuscarPre');
    const mic = document.getElementById('micBuscarPre');
    if (mic) mic.classList.add('listening');
    Voice.listen({
      onResult: (txt) => {
        if (inp) {
          inp.value = txt;
          if (window.PreingresosView) PreingresosView.buscar(txt);
        }
      },
      onEnd:    () => { if (mic) mic.classList.remove('listening'); },
      onError:  () => { if (mic) mic.classList.remove('listening'); }
    });
  }

  // ── Reabrir guía con clave admin ──
  function abrirReabrirAdmin(idGuia) {
    _reabrirCtx = { idGuia };
    // Ventana de gracia 5 min: si la guía se cerró hace < 5 min, no pedir clave
    if (_dentroGraciaCierre(idGuia)) {
      if (typeof toast === 'function') toast('🕒 Reabriendo en gracia (sin clave)', 'info');
      return _ejecutarReabrir(idGuia);
    }
    // Recordar 30 min
    const remUntil = parseInt(localStorage.getItem(ADMIN_REMEMBER_KEY) || '0', 10);
    if (remUntil && Date.now() < remUntil) {
      return _ejecutarReabrir(idGuia);
    }
    document.getElementById('reabrirAdminInput').value = '';
    document.getElementById('reabrirAdminRemember').checked = false;
    document.getElementById('overlayReabrirAdmin').style.display = 'block';
    document.getElementById('sheetReabrirAdmin').classList.add('open');
    setTimeout(() => document.getElementById('reabrirAdminInput').focus(), 200);
  }
  function cerrarReabrirAdmin() {
    document.getElementById('overlayReabrirAdmin').style.display = 'none';
    document.getElementById('sheetReabrirAdmin').classList.remove('open');
    _reabrirCtx = null;
  }
  async function confirmarReabrirAdmin() {
    if (!_reabrirCtx) return;
    const clave = document.getElementById('reabrirAdminInput').value.trim();
    if (clave.length !== 8) return (typeof toast === 'function' && toast('Clave debe ser de 8 dígitos', 'warn'));
    try {
      // [v2.13.38] Auditoría enriquecida (tier + tiempo + deviceId)
      const t0 = _reabrirCtx._t0 || Date.now();
      const devId = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
      const res = await API.verificarClaveAdmin({
        clave,
        accion: 'REABRIR_GUIA',
        refDocumento: _reabrirCtx.idGuia,
        appOrigen: 'warehouseMos',
        tier: 1,
        cache_hit: 0,
        tiempo_verify_ms: Date.now() - t0,
        deviceId: devId,
        dispositivo: devId
      });
      if (!res || !res.ok) return (typeof toast === 'function' && toast('Clave incorrecta', 'warn'));
      if (document.getElementById('reabrirAdminRemember').checked) {
        localStorage.setItem(ADMIN_REMEMBER_KEY, String(Date.now() + 30 * 60 * 1000));
      }
      const idGuia = _reabrirCtx.idGuia;
      cerrarReabrirAdmin();
      await _ejecutarReabrir(idGuia);
    } catch(e) {
      if (typeof toast === 'function') toast('Error de conexión', 'warn');
    }
  }
  async function _ejecutarReabrir(idGuia) {
    try {
      const usuario = getUsuario();
      const res = await API.reabrirGuia({ idGuia, usuario });
      if (res && res.ok) {
        if (typeof toast === 'function') toast('✓ Guía reabierta', 'ok');
        if (typeof SoundFX !== 'undefined' && SoundFX.done) SoundFX.done();
        // [v2.13.186 BUG reabrir] PARCHEAR el cache ANTES del silentRefresh.
        // Antes: precargarOperacional corría async (sin await) y silentRefresh
        // leía el cache TODAVÍA VIEJO → la guía recién reabierta se re-pintaba
        // CERRADA → no se podía editar cantidad ("desbloqueé pero no edita").
        OfflineManager.patchGuiaCache?.(idGuia, { estado: 'ABIERTA' });
        OfflineManager.precargarOperacional?.(true);
        if (window.GuiasView && GuiasView.silentRefresh) GuiasView.silentRefresh();
        // Si el detalle de ESTA guía está abierto, reflejar el estado ahí también
        if (window.GuiasView && GuiasView.marcarGuiaAbierta) GuiasView.marcarGuiaAbierta(idGuia);
      } else if (typeof toast === 'function') {
        toast('No se pudo reabrir: ' + ((res && res.error) || '?'), 'warn');
      }
    } catch(e) {
      if (typeof toast === 'function') toast('Sin conexión', 'warn');
    }
  }
  function _dentroGraciaCierre(idGuia) {
    try {
      const map = JSON.parse(localStorage.getItem('wh_gracia_cierre') || '{}');
      const ts = map[idGuia];
      if (!ts) return false;
      return (Date.now() - ts) < 5 * 60 * 1000;
    } catch(_) { return false; }
  }
  function _registrarCierre(idGuia) {
    try {
      const map = JSON.parse(localStorage.getItem('wh_gracia_cierre') || '{}');
      map[idGuia] = Date.now();
      localStorage.setItem('wh_gracia_cierre', JSON.stringify(map));
    } catch(_){}
  }

  // ── Procesar mermas (clave admin) ──
  function abrirProcesarMermas() {
    document.getElementById('procesarMermasInput').value = '';
    document.getElementById('overlayProcesarMermas').style.display = 'block';
    document.getElementById('sheetProcesarMermas').classList.add('open');
    setTimeout(() => document.getElementById('procesarMermasInput').focus(), 200);
  }
  function cerrarProcesarMermas() {
    document.getElementById('overlayProcesarMermas').style.display = 'none';
    document.getElementById('sheetProcesarMermas').classList.remove('open');
  }
  async function confirmarProcesarMermas() {
    const clave = document.getElementById('procesarMermasInput').value.trim();
    if (clave.length !== 8) return (typeof toast === 'function' && toast('Clave de 8 dígitos', 'warn'));
    if (window.Mermas) {
      const res = await Mermas.procesarEliminacion(clave);
      if (res) cerrarProcesarMermas();
    }
  }

  // ── Helpers públicos ──
  function getUsuario() {
    try { return (window.WH_CONFIG && WH_CONFIG.usuario) || ''; } catch(_) { return ''; }
  }

  async function actualizarChipDia() {
    if (typeof Cargadores !== 'undefined') {
      const n = await Cargadores.refreshCountDia();
      const el = document.getElementById('chipCargadoresDia');
      if (el) el.textContent = '🛺 ' + n;
    }
  }
  async function actualizarBadgeMermas() {
    if (typeof Mermas !== 'undefined') {
      const n = await Mermas.refreshBadge();
      const btn = document.getElementById('topCestaBtn');
      if (btn) btn.classList.toggle('has-pending', n > 0);
    }
  }

  // ════════════════════════════════════════════════
  // 📜 Historial completo de guía (admin/master)
  // ════════════════════════════════════════════════
  let _historialActual = null;

  async function abrirHistorial(idGuia, titulo) {
    document.getElementById('histGuiaTitle').textContent = `${idGuia} · ${titulo || ''}`;
    document.getElementById('histLoading').style.display  = 'block';
    document.getElementById('histTimeline').style.display = 'none';
    document.getElementById('histError').style.display    = 'none';
    document.getElementById('overlayHistorialGuia').style.display = 'block';
    document.getElementById('sheetHistorialGuia').classList.add('open');
    try {
      const rol = String(window.WH_CONFIG?.rol || '').toUpperCase();
      const res = await API.getHistorialGuia({
        idGuia,
        idPersonal: window.WH_CONFIG?.idPersonal || '',
        usuario:    window.WH_CONFIG?.usuario   || '',
        rol
      });
      if (!res || !res.ok) {
        document.getElementById('histLoading').style.display = 'none';
        const err = document.getElementById('histError');
        err.textContent = (res && res.error) || 'Error al obtener historial';
        err.style.display = 'block';
        return;
      }
      _historialActual = res.data;
      _renderHistorial(res.data);
    } catch(e) {
      document.getElementById('histLoading').style.display = 'none';
      const err = document.getElementById('histError');
      err.textContent = 'Sin conexión: ' + e.message;
      err.style.display = 'block';
    }
  }

  function cerrarHistorial() {
    document.getElementById('overlayHistorialGuia').style.display = 'none';
    document.getElementById('sheetHistorialGuia').classList.remove('open');
    _historialActual = null;
  }

  function _evtCssClass(tipo) {
    const t = String(tipo || '').toUpperCase();
    if (t === 'CREACION')           return 'evt-creacion';
    if (t.indexOf('CIERRE') >= 0 || t.indexOf('CERRAR') >= 0) return 'evt-cierre';
    if (t.indexOf('STOCK') >= 0)    return 'evt-stock';
    if (t.indexOf('PN_') >= 0)      return 'evt-pn';
    if (t.indexOf('MERMA') >= 0)    return 'evt-merma';
    if (t.indexOf('OP_') === 0)     return 'evt-op';
    if (t.indexOf('LEGACY') >= 0)   return 'evt-legacy';
    return '';
  }

  function _renderHistorial(data) {
    document.getElementById('histLoading').style.display = 'none';
    const cont = document.getElementById('histTimeline');
    cont.style.display = 'block';
    const eventos = data.eventos || [];
    if (!eventos.length) {
      cont.innerHTML = '<p style="text-align:center;color:#64748b;padding:30px">Sin eventos registrados.</p>';
      return;
    }

    let html = '';
    let lastDay = '';
    eventos.forEach((e, i) => {
      const ts = e.ts ? new Date(e.ts) : null;
      const dayKey = ts ? ts.toISOString().substring(0, 10) : 'sin fecha';
      if (dayKey !== lastDay) {
        const label = ts
          ? ts.toLocaleDateString('es-PE', { weekday: 'long', day: 'numeric', month: 'long' })
          : 'Sin fecha';
        html += `<div class="hist-day-sep">${label}</div>`;
        lastDay = dayKey;
      }
      const horaStr = ts ? ts.toTimeString().substring(0, 5) : '--:--';
      const detStr = JSON.stringify(e.detalle || {}, null, 2);
      html += `
        <div class="hist-event ${_evtCssClass(e.tipo)}">
          <span class="hist-icon">${e.icono || '•'}</span>
          <div class="hist-body">
            <div class="hist-titulo">${_escHtml(e.titulo || e.tipo || '')}</div>
            <div class="hist-meta">
              <span class="hist-meta-ts">${horaStr}</span>
              ${e.usuario  ? '<span>· ' + _escHtml(e.usuario)  + '</span>' : ''}
              ${e.deviceId ? '<span>· 📱 ' + _escHtml(e.deviceId.substring(0,10)) + '</span>' : ''}
              ${e.estado   ? '<span>· ' + _escHtml(e.estado)   + '</span>' : ''}
              ${e.error    ? '<span style="color:#f87171">· ' + _escHtml(e.error) + '</span>' : ''}
            </div>
            <div class="hist-detalle" onclick="this.classList.toggle('expanded')">${_escHtml(detStr)}</div>
          </div>
        </div>`;
    });
    cont.innerHTML = html;
  }

  function _escHtml(s) {
    return String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
  }

  async function copiarHistorialJSON() {
    if (!_historialActual) return;
    const text = JSON.stringify(_historialActual, null, 2);
    try {
      await navigator.clipboard.writeText(text);
      if (typeof toast === 'function') toast('JSON copiado al portapapeles', 'ok');
    } catch(e) {
      if (typeof toast === 'function') toast('No se pudo copiar', 'warn');
    }
  }

  function descargarHistorialJSON() {
    if (!_historialActual) return;
    const text = JSON.stringify(_historialActual, null, 2);
    const blob = new Blob([text], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `historial_${_historialActual.idGuia}_${Date.now()}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  return { init, nav, abrirMas, navMas,
           toggleModoEnvasador,
           toggleUserMenu, closeUserMenu,
           toggleSideUserMenu, closeSideUserMenu,
           syncForzado, checkUpdate,
           instalarPWA: () => window._installPWA?.(),
           cargarDashboard, showUsuarioDialog,
           cargarProductosMaestro, cargarProveedoresMaestro,
           getProductosMaestro, getProveedoresMaestro,
           getView: () => currentView,
           // F2-F7 ───
           abrirActionDia, cerrarActionDia,
           abrirTypePicker, cerrarTypePicker, volverTypePicker,
           elegirDireccion, elegirSubtipo,
           abrirSubtypePicker, cerrarSubtypePicker,
           abrirEntidadPicker, cerrarEntidadPicker, filtrarEntidad, _pickEntidad,
           dictarEntidad, dictarCargador, dictarBuscarGuia, dictarBuscarPre,
           abrirReabrirAdmin, cerrarReabrirAdmin, confirmarReabrirAdmin,
           abrirProcesarMermas, cerrarProcesarMermas, confirmarProcesarMermas,
           actualizarChipDia, actualizarBadgeMermas,
           abrirHistorial, cerrarHistorial,
           copiarHistorialJSON, descargarHistorialJSON,
           cerrarLpMenu,
           _bindNavV3, _moverNavPill, _whLogoEfecto,
           getUsuario,
           _registrarCierre };
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
  // [v2.13.186 BUG reabrir] Marca una guía como ABIERTA en el estado local del
  // módulo (todas + _guiaActual) y re-renderiza. Lo llama _ejecutarReabrir (App),
  // que vive fuera de este módulo y no puede tocar _guiaActual directo.
  function marcarGuiaAbierta(idGuia) {
    const idx = todas.findIndex(g => g.idGuia === idGuia);
    if (idx >= 0) { todas[idx].estado = 'ABIERTA'; render(_filtrarYBuscar()); }
    if (_guiaActual?.idGuia === idGuia) {
      _guiaActual.estado = 'ABIERTA';
      _mostrarDetalleSheet(_guiaActual, false);
    }
  }

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
    const clT = document.getElementById('clearGuiaTabletSearch');
    if (clT) clT.style.display = _busquedaQ ? 'flex' : 'none';
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
    const inpT = document.getElementById('guiaTabletSearch');
    if (inpT) inpT.value = '';
    const clT = document.getElementById('clearGuiaTabletSearch');
    if (clT) clT.style.display = 'none';
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
    const active = filtroActual !== 'TODAS';
    // Mobile dot + dropdown
    const dot = document.getElementById('guiaFilterDot');
    if (dot) dot.style.display = active ? 'block' : 'none';
    document.querySelectorAll('.guia-fopt').forEach(b =>
      b.classList.toggle('sel', b.dataset.filtro === filtroActual));
    _cerrarFiltroMenu();
    // Tablet chips
    document.querySelectorAll('[data-gtfiltro]').forEach(b =>
      b.classList.toggle('active', b.dataset.gtfiltro === filtroActual));
    render(_filtrarYBuscar());
  }

  function _searchFocus(focused) {
    const toolbar = document.getElementById('guiasToolbar');
    if (!toolbar) return;
    if (focused) {
      toolbar.classList.add('srch-focused');
      _cerrarFiltroMenu();
    } else {
      setTimeout(() => toolbar.classList.remove('srch-focused'), 160);
    }
  }

  // ── Tablet chips filter (panel #guiaTabletToolbar) ────────────────────
  function filtrarTablet(f) {
    // Toggle: click active chip → back to TODAS
    filtrar(filtroActual === f ? 'TODAS' : f);
  }

  function _searchFocusGuiaTablet(focused) {
    const toolbar = document.getElementById('guiaTabletToolbar');
    if (!toolbar) return;
    if (focused) toolbar.classList.add('srch-focused');
    else setTimeout(() => toolbar.classList.remove('srch-focused'), 160);
  }

  function _getProvNombre(idProveedor) {
    if (!idProveedor) return '';
    const p = OfflineManager.getProveedoresCache().find(x => x.idProveedor === idProveedor);
    return p ? (p.nombre || idProveedor) : idProveedor;
  }

  function _renderGuiaCard(g) {
    const isEnvasado  = g.tipo === 'SALIDA_ENVASADO' || g.tipo === 'INGRESO_ENVASADO';
    const isMerma     = g.tipo === 'SALIDA_MERMA';
    const isIngreso   = g.tipo?.startsWith('INGRESO');
    const isAbierta   = g.estado === 'ABIERTA';
    const borderColor = isMerma ? '#dc2626' : isEnvasado ? '#475569' : isAbierta ? '#f59e0b' : isIngreso ? '#22c55e' : '#3b82f6';
    const tipoLabel   = (isMerma ? '🗑 ' : '') + (TIPO_LABELS[g.tipo] || g.tipo || '—');
    const provNombre  = _getProvNombre(g.idProveedor) || g.usuario || '—';
    const hora        = _horaDesdeGuia(g);
    const fechaCorta  = _fmtCorta(g.fecha);

    // Guías de envasado: card gris, solo lectura
    if (isEnvasado) {
      const detCache = OfflineManager.getGuiaDetalleCache();
      const numItems = detCache.filter(d => d.idGuia === g.idGuia && d.observacion !== 'ANULADO').length;
      return `
      <div class="guia-card" id="guia-${(g.idGuia||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
           style="border-left-color:#475569;opacity:.65;cursor:default"
           onclick="GuiasView.verDetalle('${escAttr(g.idGuia)}')">
        <div class="card-row-top">
          <span class="card-tipo-chip" style="background:rgba(71,85,105,.2);color:#64748b">${tipoLabel}</span>
          <span style="font-size:10px;color:#475569">🔒 sistema</span>
        </div>
        <p class="card-name" style="color:#64748b">${escHtml(g.usuario || '—')}</p>
        <div class="card-row-bottom">
          <span class="card-meta">${fechaCorta}${numItems ? ' · ' + numItems + ' prod' : ''}</span>
        </div>
      </div>`;
    }

    const chipBg  = isAbierta ? 'rgba(245,158,11,.15)' : isIngreso ? 'rgba(34,197,94,.15)' : 'rgba(59,130,246,.15)';
    const chipCol = isAbierta ? '#fbbf24' : isIngreso ? '#4ade80' : '#60a5fa';
    const estadoDot = isAbierta
      ? `<span style="width:8px;height:8px;border-radius:50%;background:#f59e0b;display:inline-block;flex-shrink:0" title="Abierta"></span>`
      : `<span style="width:8px;height:8px;border-radius:50%;background:#22c55e;display:inline-block;flex-shrink:0" title="Cerrada"></span>`;

    const fotoTag = g.foto ? `<span class="pre-qtag pre-qtag-slate">📷</span>` : '';
    const preTag  = g.idPreingreso
      ? `<span onclick="event.stopPropagation();GuiasView.irAPreingreso('${escAttr(g.idPreingreso)}')"
             class="pre-qtag pre-qtag-blue" title="Ver preingreso" style="cursor:pointer;user-select:none">📋</span>`
      : '';

    const detCache  = OfflineManager.getGuiaDetalleCache();
    const detItems  = detCache.filter(d => d.idGuia === g.idGuia && d.observacion !== 'ANULADO');
    const numItems  = detItems.length;
    const totalUds  = detItems.reduce((s, d) => s + (parseFloat(d.cantidadRecibida) || 0), 0);
    const udsStr    = totalUds % 1 === 0 ? String(totalUds) : fmt(totalUds, 1);
    const metaExtra = numItems > 0 ? ` · ${numItems} prod · ${udsStr} uds` : '';

    const pnPend  = (OfflineManager.getPNCache() || []).filter(p => p.idGuia === g.idGuia && p.estado === 'PENDIENTE').length;
    const pnBadge = pnPend ? `<span style="background:#78350f;color:#fde68a;font-size:9px;font-weight:800;
      padding:2px 6px;border-radius:4px;flex-shrink:0;letter-spacing:.04em;cursor:pointer"
      onclick="event.stopPropagation();GuiasView.abrirModalPN('',${JSON.stringify(g.idGuia)})"
      title="${pnPend} producto(s) nuevo(s) pendiente(s)">N${pnPend > 1 ? ' ' + pnPend : ''}</span>` : '';

    const waBtn = `<button onclick="event.stopPropagation();GuiasView.compartirWA('${escAttr(g.idGuia)}')"
      class="card-act card-act-wa" title="Compartir por WhatsApp">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
    </button>`;
    const printBtn = `<button onclick="event.stopPropagation();GuiasView.imprimirTicket('${escAttr(g.idGuia)}')"
      class="card-act card-act-print" title="Imprimir ticket">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M2.5 8a.5.5 0 1 0 0-1 .5.5 0 0 0 0 1z"/><path d="M5 1a2 2 0 0 0-2 2v2H2a2 2 0 0 0-2 2v3a2 2 0 0 0 2 2h1v1a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2v-1h1a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-1V3a2 2 0 0 0-2-2H5zM4 3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2H4V3zm1 5a2 2 0 0 0-2 2v1H2a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v-1a2 2 0 0 0-2-2H5zm7 2v3a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1z"/></svg>
    </button>`;

    // Foto block (izquierda) — thumb si existe, placeholder con tipo si no
    const fotoUrl = g.foto ? _normalizeDriveUrl(g.foto) : '';
    const fotoIcon = isIngreso ? '↓' : '↑';
    const fotoBlock = fotoUrl
      ? `<div class="gcard-photo" onclick="event.stopPropagation();Photos&&Photos.lightbox('${escAttr(fotoUrl)}')"><img src="${escAttr(fotoUrl)}" loading="lazy" onerror="this.style.opacity='.3'"/></div>`
      : `<div class="gcard-photo placeholder" title="${tipoLabel}">${fotoIcon}</div>`;

    return `
    <div class="guia-card" id="guia-${(g.idGuia||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
         style="border-left-color:${borderColor}"
         onclick="GuiasView.verDetalle('${escAttr(g.idGuia)}')">
      ${fotoBlock}
      <div class="gcard-body">
        <div class="card-row-top">
          <span class="card-tipo-chip" style="background:${chipBg};color:${chipCol}">${tipoLabel}</span>
          <div class="flex items-center gap-2 flex-shrink-0">${pnBadge}${preTag}${estadoDot}</div>
        </div>
        <p class="card-name" style="font-size:16px">${escHtml(provNombre)}</p>
        <div class="card-row-bottom">
          <span class="card-meta">${fechaCorta}${hora ? ' · ' + hora : ''}${metaExtra}</span>
          <div class="card-actions">${waBtn}${printBtn}</div>
        </div>
      </div>
    </div>`;
  }

  function render(list) {
    const container = document.getElementById('listGuias');
    if (!container) return;
    const optCards = Array.from(container.querySelectorAll('.card-optimistic'));
    if (!list.length) {
      // [v2.13.188] Distinguir "no hay" de "sin coincidencias" (filtro/búsqueda activos)
      const filtrando = !!(filtroActual || _busquedaQ);
      const msg = filtrando ? 'Sin coincidencias para el filtro o la búsqueda' : 'No hay guías';
      container.innerHTML = `<p class="text-slate-500 text-center py-8 text-sm">${msg}</p>`;
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

    // [v2.13.181] Día en TZ de Perú (consistente con el pill de cargadores y el backend)
    const hoyKey  = _hoyPeru();
    const ayerKey = _ayerPeru();
    const months  = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    function _gKey(g) {
      return (g.fecha && _diaPeru(g.fecha)) || '0000-00-00';
    }
    function _gLabel(key) {
      if (!key || key === '0000-00-00') return 'Sin fecha';
      if (key === hoyKey)  return 'Hoy';
      if (key === ayerKey) return 'Ayer';
      const [yyyy, mm, dd] = key.split('-'); // key = 'YYYY-MM-DD' del día de Perú
      return yyyy === hoyKey.slice(0, 4)
        ? `${parseInt(dd, 10)} ${months[parseInt(mm, 10) - 1]}`
        : `${dd}/${mm}/${yyyy}`;
    }

    const groupMap = {};
    sorted.forEach(g => {
      const k = _gKey(g);
      if (!groupMap[k]) groupMap[k] = [];
      groupMap[k].push(g);
    });

    container.innerHTML = Object.entries(groupMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([key, items]) => {
        // [v2.13.3] Pill de cargadores: carretas + advertencia si alguna no llena
        const r = _resumenCargadoresDiaPorFecha(key);
        const hdrCls = r.carretas > 0 ? 'pre-date-hdr pre-date-hdr-row' : 'pre-date-hdr';
        const mid    = r.carretas > 0 ? '<span class="pre-hdr-line"></span>' : '';
        const alerta = ((r.medias || 0) + (r.vacias || 0)) > 0 ? '<span>·</span><span>⚠</span>' : '';
        const pill = r.carretas > 0
          ? `<button onclick="PreingresosView.abrirCargadoresDia('${key}')"
                     class="carg-pill-btn inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-semibold">
                <span>🛺</span><span>${r.carretas} cart</span>${alerta}
              </button>`
          : '';
        const addBtn = `<button class="day-header-add" onclick="App.abrirActionDia()" title="Crear (guía/preingreso/cargador)">+</button>`;
        const countChip = `<span class="day-chip" title="${items.length} guías del día">📋 ${items.length}</span>`;
        return `<div class="${hdrCls}"><span>${_gLabel(key)}</span>${mid}${countChip}${pill}${addBtn}</div>
                <div class="pre-date-group">${items.map(_renderGuiaCard).join('')}</div>`;
      }).join('');

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
      <p class="text-xs text-slate-400 truncate">${escHtml(provNombre || 'Sin proveedor')}</p>`;
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
          <span class="text-xs ${esI ? 'tag-ok' : 'tag-blue'}">${escHtml(label)}</span>
          <span class="text-xs text-emerald-400 font-bold">ABIERTA</span>
        </div>
        <p class="text-sm font-bold text-slate-100 truncate mt-1">${escHtml(provNombre || 'Sin proveedor')}</p>
        <p class="text-xs text-slate-500 font-mono">${escHtml(idGuia || '')}</p>`;
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
    const equivs   = OfflineManager.getEquivalenciasCache();
    const prodMap  = {};
    // Index del nombre por idProducto/skuBase/codigoBarra del maestro.
    // skuBase solo se indexa para productos BASE (factor=1, activos): evita que una
    // presentación (factor>1) con el mismo skuBase sobreescriba el nombre del base.
    prods.forEach(p => {
      const name = p.descripcion || p.nombre || '';
      if (!name) return;
      prodMap[p.idProducto] = name;
      if (p.codigoBarra) prodMap[String(p.codigoBarra)] = name;
      const isBase = parseFloat(p.factorConversion || 1) === 1
                  && p.estado !== '0' && p.estado !== 0;
      if (isBase && p.skuBase) prodMap[String(p.skuBase).trim().toUpperCase()] = name;
    });
    // Index del nombre por codigoBarra de cada equivalente → resuelve al producto base via skuBase
    equivs.forEach(e => {
      if (!e.codigoBarra || !e.skuBase) return;
      const skuKey = String(e.skuBase).trim().toUpperCase();
      const name   = prodMap[skuKey];
      if (name) prodMap[String(e.codigoBarra)] = name;
    });

    let guia = guias.find(g => g.idGuia === idGuia);
    if (!guia) {
      // Fallback: mostrar loading y pedir a GAS
      _abrirDetalleConGAS(idGuia);
      return;
    }

    const detalle = detalles
      .filter(d => d.idGuia === idGuia)
      .map(d => ({
        ...d,
        descripcionProducto: prodMap[d.codigoProducto]   // lookup por idProducto o codigoBarra
          || d.descripcionProducto                        // preservar el nombre ya cacheado
          || d.codigoProducto                             // fallback: código crudo
      }));

    _guiaActual = { ...guia, detalle };
    _mostrarDetalleSheet(_guiaActual);

    // 2. Refrescar desde GAS en background (actualiza si hay cambios)
    if (navigator.onLine) {
      API.getGuia(idGuia).then(res => {
        if (!res.ok || res.offline) return;
        // Guard: si el usuario ya abrió otra guía, descartar esta respuesta stale
        if (_guiaActual?.idGuia !== idGuia) return;

        // Rescatar ítems locales pendientes que se agregaron mientras GAS respondía.
        // Si los borramos sobreescribiendo _guiaActual, el siguiente escaneo
        // no los encontraría y crearía una segunda fila duplicada en GAS.
        const pendingLocal = (_guiaActual.detalle || []).filter(d => d._local);

        _guiaActual = res.data;

        if (pendingLocal.length && Array.isArray(_guiaActual.detalle)) {
          // [v2.13.215][FIX 40x #5] ¿Es PN una línea? (es PN si trae el flag _esPN,
          // o si su observación arranca con 'PN_' — espeja la heurística del optimista).
          const _esLineaPN = (d) => !!d._esPN ||
            String(d.observacion || '').toUpperCase().indexOf('PN_') === 0;
          pendingLocal.forEach(p => {
            // Solo re-inyectar si el server aún no lo confirmó (agregarDetalle puede haber llegado antes)
            if (_guiaActual.detalle.some(d => d.idDetalle === p.idDetalle)) return;
            // [v2.13.215][FIX 40x #5] Dedup secundario: la card optimista (idDetalle
            // local 'DL...'/'DL_PN_...') y la fila real de Supabase ('DET_...') tienen
            // distinto idDetalle → el match por id exacto no las colapsa y la card local
            // se re-inyecta = DUPLICADO visual. Si el id local arranca con 'DL' y el
            // server ya trae una línea con el MISMO código (y misma naturaleza PN/no-PN),
            // es la misma línea ya confirmada → NO re-inyectar la local.
            const idLocal = String(p.idDetalle || '');
            if (idLocal.indexOf('DL') === 0 && p.codigoProducto) {
              const codP  = String(p.codigoProducto).toUpperCase();
              const pnP   = _esLineaPN(p);
              const yaEsta = _guiaActual.detalle.some(d =>
                String(d.codigoProducto || '').toUpperCase() === codP &&
                _esLineaPN(d) === pnP);
              if (yaEsta) return;
            }
            _guiaActual.detalle.push(p);
          });
        }

        _mostrarDetalleSheet(_guiaActual, false);
        if (Array.isArray(_guiaActual.detalle)) {
          OfflineManager.actualizarDetallesGuia(idGuia, _guiaActual.detalle);
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
    const esDiaAnterior = g.fecha && _diaPeru(g.fecha) < _hoyPeru();  // [v2.13.183] día en TZ Perú

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
        <p class="font-bold text-base leading-tight" style="color:#64748b">${escHtml(g.usuario || '—')}</p>
        <p class="text-xs mt-0.5" style="color:#475569">${fmtFecha(g.fecha)} · Solo lectura</p>`;
      // [v2.13.185] (eliminadas 3 líneas no-op: los IDs guiaDet*Panel no existen;
      // la vista de solo-lectura se arma por innerHTML arriba, no por classList.)
    } else {

    // Lock button
    const lockBtn = `
      <button onclick="GuiasView.toggleEstadoGuia()"
              style="display:flex;align-items:center;justify-content:center;
                     width:36px;height:36px;border-radius:10px;border:1.5px solid
                     ${abierta ? 'rgba(245,158,11,.6)' : 'rgba(100,116,139,.4)'};
                     background:${abierta ? 'rgba(245,158,11,.12)' : 'rgba(30,41,59,.6)'};
                     cursor:pointer;flex-shrink:0;transition:all .2s"
              title="${abierta ? 'Cerrar guía' : 'Reabrir (admin)'}">
        ${abierta
          ? `<svg width="18" height="18" viewBox="0 0 16 16" fill="${'#fbbf24'}">
               <path d="M8 1a2 2 0 0 1 2 2v2h.5A1.5 1.5 0 0 1 12 6.5V14a1.5 1.5 0 0 1-1.5 1.5h-5A1.5 1.5 0 0 1 4 14V6.5A1.5 1.5 0 0 1 5.5 5H6V3a2 2 0 0 1 2-2zm0 1a1 1 0 0 0-1 1v2h2V3a1 1 0 0 0-1-1z"/>
             </svg>`
          : `<svg width="18" height="18" viewBox="0 0 16 16" fill="${'#64748b'}">
               <path d="M8 1a2 2 0 0 1 2 2v4H6V3a2 2 0 0 1 2-2zm3 6V3a3 3 0 0 0-6 0v4a2 2 0 0 0-2 2v5a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2zM5 8h6v5H5z"/>
             </svg>`}
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
      <p class="font-black text-lg text-white leading-tight" onclick="GuiasView.deselectItem()">${escHtml(provNombreHdr)}</p>
      <p class="text-xs text-slate-500 mt-0.5" onclick="GuiasView.deselectItem()">${fmtFecha(g.fecha)} · ${g.usuario || '—'}</p>
      ${esDiaAnterior && abierta ? `<p class="text-xs text-amber-400 mt-1 font-semibold">⚠ Guía de un día anterior aún abierta</p>` : ''}
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
        <button onclick="GuiasView.editarGuia()" class="flex items-center gap-1.5 py-1.5 px-3 rounded-lg text-xs bg-slate-800 text-slate-400 transition-colors">
          <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
            <path d="M12.146.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1 0 .708l-10 10a.5.5 0 0 1-.168.11l-5 2a.5.5 0 0 1-.65-.65l2-5a.5.5 0 0 1 .11-.168l10-10zM11.207 2.5 13.5 4.793 14.793 3.5 12.5 1.207 11.207 2.5zm1.586 3L10.5 3.207 4 9.707V10h.5a.5.5 0 0 1 .5.5v.5h.5a.5.5 0 0 1 .5.5v.5h.293l6.5-6.5zm-9.761 5.175-.106.106-1.528 3.821 3.821-1.528.106-.106A.5.5 0 0 1 5 12.5V12h-.5a.5.5 0 0 1-.5-.5V11h-.5a.5.5 0 0 1-.468-.325z"/>
          </svg>
          Editar
        </button>` : ''}
        ${(['ADMIN','MASTER'].includes(String(window.WH_CONFIG?.rol || '').toUpperCase())) ? `
        <button onclick="App.abrirHistorial('${escAttr(g.idGuia)}','${escAttr(_getProvNombre(g.idProveedor) || g.usuario || '')}')"
                class="flex items-center gap-1.5 py-1.5 px-3 rounded-lg text-xs bg-slate-800 text-slate-400 transition-colors ml-auto"
                title="Historial completo (admin/master)">
          📜 Historial
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
        cEl.innerHTML = `<p class="text-xs text-slate-400 italic mb-3">${escHtml(g.comentario)}</p>`;
      } else {
        cEl.innerHTML = '';
      }
    }

    } // end else (no esEnvasado)

    const items    = (g.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const totalUds = items.reduce((s, d) => s + (parseFloat(d.cantidadRecibida) || 0), 0);
    document.getElementById('guiaDetCount').textContent = `${items.length} ítem${items.length !== 1 ? 's' : ''}`;
    // Footer totales
    const footerEl  = document.getElementById('guiaDetFooter');
    const footerVal = document.getElementById('guiaDetFooterVal');
    if (footerEl && footerVal) {
      if (items.length) {
        const udsStr = totalUds % 1 === 0 ? String(totalUds) : fmt(totalUds, 1);
        footerVal.textContent = `${items.length} prod · ${udsStr} uds`;
        footerEl.style.display = 'flex';
      } else {
        footerEl.style.display = 'none';
      }
    }

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
                  <p class="text-base font-bold text-white leading-snug">${escHtml(d.descripcionProducto || d.codigoProducto)}</p>
                  <p class="text-xs text-slate-500 font-mono mt-0.5">${escHtml(d.codigoProducto)}</p>
                </div>
                <div class="flex items-center gap-1 flex-shrink-0 mt-0.5">
                  <button onclick="GuiasView.inlineQtyDelta(-1)"
                          class="text-slate-300 text-xl leading-none w-7 h-7 flex items-center justify-center rounded-md active:bg-slate-600 select-none">−</button>
                  <span id="inlineQtyDisplay"
                        class="text-base font-black text-white w-14 text-center select-none cursor-pointer"
                        onclick="GuiasView.inlineQtyTap()">${_selQty}</span>
                  <input id="inlineQtyInput" type="number" step="any" inputmode="decimal"
                         value="${_selQty}" readonly tabindex="-1"
                         class="text-base font-black text-white bg-transparent border-b border-slate-500 text-center w-14 focus:outline-none focus:border-blue-400 hidden"
                         oninput="GuiasView.inlineQtyInput(this.value)"
                         onblur="GuiasView.inlineQtyBlurFull(this.value)"/>
                  <button onclick="GuiasView.inlineQtyDelta(1)"
                          class="text-blue-400 text-xl leading-none w-7 h-7 flex items-center justify-center rounded-md active:bg-slate-600 select-none">+</button>
                </div>
              </div>
              <div class="flex gap-2">
                ${esIngreso ? `
                <button onclick="GuiasView.inlinePickVenc()" id="inlineVencBtn"
                        class="flex-1 py-2 rounded-lg border ${vencTxt ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-xs flex items-center justify-center gap-1.5">
                  <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/></svg>
                  ${escHtml(vencTxt || 'Vencimiento')}
                </button>` : ''}
                <button onclick="GuiasView.inlineDelete(${idx})"
                        class="py-2 px-3 rounded-lg border border-red-800/60 text-red-400 text-xs flex items-center justify-center">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/><path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3h11V2h-11z"/></svg>
                </button>
              </div>
            </div>`;
          }

          // ── Tarjeta colapsada ──────────────────────────────
          // [v2.13.183] "hoy" en TZ Perú; vf normalizado a su parte de fecha
          // (la fecha de vencimiento es un día calendario, no se convierte TZ).
          const hoy       = _hoyPeru();
          const vf        = String(d.fechaVencimiento || '').slice(0, 10);
          const isVencido = vf && vf < hoy;
          const isSoon    = vf && !isVencido && vf <= _diaPeru(new Date(Date.now() + 30*86400000));
          const venc      = vf
            ? `<span style="font-size:.68rem;${isVencido ? 'color:#f87171;font-weight:700' : 'color:#fbbf24'} " class="block mt-0.5">
                ${isVencido ? '⚠ VENCIDO' : isSoon ? '⚠ Venc próx:'  : 'Venc:'} ${vf}</span>`
            : '';
          const qtyZero   = parseFloat(d.cantidadRecibida) === 0;
          const qtyColor  = qtyZero ? '#f87171' : '#fff';
          const iTag  = d._indirect
            ? `<span style="font-size:9px;font-weight:800;padding:1px 4px;border-radius:3px;
                            background:rgba(124,58,237,.18);color:#a78bfa;
                            border:1px solid rgba(124,58,237,.4);margin-right:4px;flex-shrink:0;vertical-align:middle">i</span>`
            : '';
          const rowBg = isVencido ? 'rgba(239,68,68,.07)' : '';
          // Mini-badge tipo match (✓ canónico · ↕E equivalente · ↕C completo · 🆕 nuevo)
          const matchBadge = (() => {
            const obs = String(d.observacion || '').toUpperCase();
            if (obs === 'PN_PENDIENTE') return '<span class="item-match-badge imb-nuevo" title="Producto nuevo pendiente">🆕</span>';
            if (d._indirect)            return '<span class="item-match-badge imb-equiv" title="Equivalente">↕E</span>';
            return '<span class="item-match-badge imb-canonico" title="Canónico">✓</span>';
          })();
          // Sync dot: refleja estado real (local/saving/failed/saved)
          let syncDot, syncLbl = '';
          if (d._saveFailed) {
            syncDot = '<span class="sync-dot on-failed" title="error al guardar — toca para reintentar"></span>';
            syncLbl = '<span style="color:#f87171">⚠ no guardado</span>';
          } else if (d._local || d._saving) {
            syncDot = '<span class="sync-dot on-saving" title="guardando…"></span>';
            syncLbl = '<span style="color:#fbbf24">guardando…</span>';
          } else {
            syncDot = '<span class="sync-dot on-saved" title="guardado"></span>';
          }
          return `
          <div class="flex items-center gap-3 py-3 px-3 border-b border-slate-700/50 cursor-pointer active:bg-slate-700/20 rounded-lg${pendiente}"
               style="${rowBg ? 'background:' + rowBg + ';' : 'background:rgba(30,41,59,.4);'}border-radius:10px;margin-bottom:6px"
               data-det-id="${d.idDetalle || ''}" data-det-idx="${idx}" onclick="GuiasView.selectItem(${idx})">
            <div class="flex-1 min-w-0">
              <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
                ${matchBadge}
                <p class="font-bold text-slate-100 leading-tight" style="font-size:15px;flex:1;min-width:0">${iTag}${escHtml(d.descripcionProducto || d.codigoProducto)}</p>
              </div>
              <p class="text-xs text-slate-500 font-mono mt-1" style="display:flex;align-items:center;gap:6px">
                <span>${escHtml(d.codigoProducto)}</span>
                ${syncDot}
                ${syncLbl}
              </p>
              ${venc}
            </div>
            <span class="text-lg font-black flex-shrink-0" style="color:${qtyColor}">${qtyZero ? '⚠ 0' : '×' + fmt(d.cantidadRecibida, Number.isInteger(parseFloat(d.cantidadRecibida)) ? 0 : 2)}</span>
          </div>`;
        }).join('')
      : '<p class="text-slate-500 text-sm text-center py-4">Sin ítems registrados</p>';

    const monto = parseFloat(g.montoTotal) || 0;
    document.getElementById('guiaDetMontoVal').textContent = monto > 0 ? `S/. ${fmt(monto, 2)}` : '—';
    document.getElementById('guiaDetMonto').style.display = monto > 0 ? 'block' : 'none';

    if (abierta) requestAnimationFrame(_initSwipeGuia);

    const acciones = document.getElementById('guiaDetAcciones');
    const _camItemCount = ((g.detalle || []).filter(d => d.observacion !== 'ANULADO')).length;
    const _camBadge = _camItemCount > 0
      ? `<span style="background:#7c3aed;color:#fff;border-radius:9px;padding:1px 6px;
                      font-size:.65em;font-weight:900;line-height:1.4;margin-left:2px">${_camItemCount}</span>`
      : '';
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
          CÁMARA${_camBadge}
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
  let _selDetId   = '';   // [v2.13.183] idDetalle seleccionado (fuente de verdad al commitear)
  let _selQty     = 0;
  let _selVenc    = '';
  let _selOrigQty = 0;
  let _selOrigVenc = '';
  let _selGuiaId  = '';

  // [v2.13.183] Surfacing de error para ediciones optimistas de detalle: si el
  // backend rechaza o no hay red, avisar (antes era .catch(()=>{}) silencioso →
  // mostraba "Guardado" aunque no se guardara y al refrescar revertía).
  function _avisarFalloDetalle(p, label) {
    try {
      p.then(r => { if (r && r.ok === false) { try { toast('⚠ No se guardó ' + label + ': ' + (r.error || ''), 'warn', 4000); } catch(_){} } })
       .catch(() => { try { toast('⚠ Sin conexión — ' + label + ' no se guardó', 'warn', 3500); } catch(_){} });
    } catch(_){}
  }

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
    _selDetId    = d.idDetalle;   // [v2.13.183] fuente de verdad al commitear
    _selQty      = parseFloat(d.cantidadRecibida) || 0;
    _selVenc     = d.fechaVencimiento || '';
    _selOrigQty  = _selQty;
    _selOrigVenc = _selVenc;
    _mostrarDetalleSheet(_guiaActual, false);
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

    // [v2.13.183] Resolver el ítem por idDetalle (no por índice): si un
    // silentRefresh re-renderizó y la lista filtrada cambió, el índice quedaría
    // stale y editaríamos el ítem equivocado.
    const d = (_guiaActual.detalle || []).find(x => x.idDetalle === _selDetId);
    if (!d) return;

    const qtyChanged  = _selQty !== _selOrigQty;
    const vencChanged = _selVenc !== _selOrigVenc;
    if (!qtyChanged && !vencChanged) return;

    const idDetalle = d.idDetalle;
    if (_selQty <= 0) {
      d.observacion = 'ANULADO';
      OfflineManager.addDetalleCache(d);
      _avisarFalloDetalle(API.anularDetalle({ idDetalle }), 'la anulación');
      toast('Ítem eliminado', 'warn', 1200);
      return;
    }
    d.cantidadRecibida = _selQty;
    d.fechaVencimiento = _selVenc;
    OfflineManager.addDetalleCache(d);
    if (qtyChanged)  _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle, cantidadRecibida: _selQty }), 'la cantidad');
    if (vencChanged) _avisarFalloDetalle(API.actualizarFechaVencimiento({ idDetalle, fechaVencimiento: _selVenc }), 'el vencimiento');
    toast('Guardado', 'ok', 1000);
  }

  function inlineQtyDelta(delta) {
    _selQty = Math.max(0, parseFloat(_selQty || 0) + delta);
    const inp  = document.getElementById('inlineQtyInput');
    const span = document.getElementById('inlineQtyDisplay');
    if (inp)  inp.value        = _selQty;
    if (span) span.textContent = _selQty;
  }

  // El usuario toca el número → ocultar span, mostrar input con foco
  function inlineQtyTap() {
    const inp  = document.getElementById('inlineQtyInput');
    const span = document.getElementById('inlineQtyDisplay');
    if (!inp || !span) return;
    inp.removeAttribute('readonly');
    inp.removeAttribute('tabindex');
    span.classList.add('hidden');
    inp.classList.remove('hidden');
    inp.focus();
    inp.select();
  }

  function inlineQtyInput(val) {
    const n = parseFloat(val);
    _selQty = isNaN(n) ? 0 : Math.max(0, n);
  }

  // Al salir del input: volver a mostrar el span con el valor actualizado
  function inlineQtyBlurFull(val) {
    const n = parseFloat(val);
    if (!isNaN(n)) _selQty = Math.max(0, n);
    const inp  = document.getElementById('inlineQtyInput');
    const span = document.getElementById('inlineQtyDisplay');
    if (inp) { inp.setAttribute('readonly', ''); inp.setAttribute('tabindex', '-1'); inp.classList.add('hidden'); }
    if (span) { span.textContent = _selQty; span.classList.remove('hidden'); }
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
      btn.innerHTML = `<svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5zM1 4v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V4H1z"/></svg> ${_selVenc ? escHtml(_selVenc) : 'Vencimiento'}`;
    }
  }

  function inlineDelete(idx) {
    const items = (_guiaActual?.detalle || []).filter(d => d.observacion !== 'ANULADO');
    const d = items[idx];
    if (!d) return;
    d.observacion = 'ANULADO';
    OfflineManager.addDetalleCache(d);
    _selIdx = -1;
    _mostrarDetalleSheet(_guiaActual, false);
    _avisarFalloDetalle(API.anularDetalle({ idDetalle: d.idDetalle }), 'la anulación');
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
        <p class="font-bold text-white text-base">${escHtml(d.descripcionProducto || d.codigoProducto)}</p>
        <p class="text-xs text-slate-500 font-mono">${escHtml(d.codigoProducto)}</p>
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
      OfflineManager.addDetalleCache(d);
      _mostrarDetalleSheet(_guiaActual, false);
      _avisarFalloDetalle(API.anularDetalle({ idDetalle: _editItemId }), 'la anulación');
      toast('Ítem eliminado', 'warn', 1500);
      return;
    }

    d.cantidadRecibida = qtyFinal;
    d.fechaVencimiento = _editItemVenc;
    OfflineManager.addDetalleCache(d);
    _mostrarDetalleSheet(_guiaActual, false);

    // [v2.13.98] Antes ambos catch silenciaban errores → user veía 'Ítem guardado'
    // pero el backend había rechazado. Al imprimir, salía la versión vieja.
    //
    // [v2.13.99 senior audit] Manejar 3 casos:
    //   r.offline=true   → encolado en cola offline (info, no error)
    //   r.ok=false       → backend rechazó (ERROR visible al user)
    //   exception .catch → fetch tiró antes del retry interno (raro)
    function _toastResultado(r, label, e) {
      if (e) {
        toast('⚠ Error red al guardar ' + label + ' — verificá conexión', 'error', 6000);
        console.error('[itemEdit] ' + label + ' network fail:', e?.message);
      } else if (r && r.offline) {
        // Cae a wh_queue, se sincronizará cuando vuelva la red.
        // No es error pero el user debe saber que aún no llegó al backend.
        toast('📡 Sin red — ' + label + ' guardada en cola', 'warn', 4000);
        console.log('[itemEdit] ' + label + ' encolada offline · localId=' + (r.localId || '?'));
      } else if (r && r.ok === false) {
        toast('⚠ No se pudo guardar ' + label + ': ' + (r.error || 'backend rechazó'), 'error', 6000);
        console.error('[itemEdit] ' + label + ' rechazado:', r);
      }
    }
    if (qtyChanged) {
      API.actualizarCantidadDetalle({ idDetalle: _editItemId, cantidadRecibida: qtyFinal })
        .then(r => _toastResultado(r, 'cantidad'))
        .catch(e => _toastResultado(null, 'cantidad', e));
    }
    if (vencChanged) {
      API.actualizarFechaVencimiento({ idDetalle: _editItemId, fechaVencimiento: _editItemVenc })
        .then(r => _toastResultado(r, 'vencimiento'))
        .catch(e => _toastResultado(null, 'vencimiento', e));
    }
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
      </svg> ${val ? escHtml(val) : 'Agregar vencimiento'}`;
      btn.className = `w-full py-3 rounded-xl border ${val ? 'border-amber-500 text-amber-300' : 'border-slate-600 text-slate-400'} text-sm font-bold mb-3 flex items-center justify-center gap-2`;
    }
  }

  function eliminarItemEdit() {
    const d = (_guiaActual?.detalle || []).filter(x => x.observacion !== 'ANULADO')[_editItemIdx];
    if (!d) return;
    d.observacion = 'ANULADO';
    OfflineManager.addDetalleCache(d);
    cerrarSheet('sheetEditItem');
    _mostrarDetalleSheet(_guiaActual, false);
    _avisarFalloDetalle(API.anularDetalle({ idDetalle: _editItemId }), 'la anulación');
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

  // Local dot-indicator updater for admin PIN (8 dots: 4 globales + 4 admin)
  function _updAdminDots(n) {
    for (let i = 0; i < 8; i++) {
      const el = document.getElementById('apn' + i);
      if (!el) continue;
      const filled = i < n;
      // Globales (0-3): ámbar. Admin (4-7): emerald.
      const colorVal = i < 4 ? 'bg-amber-400' : 'bg-emerald-400';
      const colorBor = i < 4 ? 'border-slate-600' : 'border-emerald-700';
      el.className = filled
        ? 'w-3.5 h-3.5 rounded-full ' + colorVal
        : 'w-3.5 h-3.5 rounded-full border-2 ' + colorBor;
    }
  }

  // Admin PIN dialog para reabrir guía — clave 8 dígitos (global+admin)
  let _pinGuiaTarget = null;
  let _adminPinBuf   = '';

  function _pedirAdminPin(idGuia) {
    _pinGuiaTarget = idGuia;
    _adminPinBuf   = '';
    _updAdminDots(0);
    document.getElementById('adminPinError').textContent = '';
    document.getElementById('adminPinModal').style.display = 'flex';
    // Refrescar caché en background para próxima vez
    if (typeof OfflineManager !== 'undefined' && OfflineManager.sincronizarAdminCache) {
      OfflineManager.sincronizarAdminCache();
    }
  }

  function adminPinTecla(d) {
    if (_adminPinBuf.length >= 8) return;
    _adminPinBuf += d;
    _updAdminDots(_adminPinBuf.length);
    if (_adminPinBuf.length === 8) setTimeout(_verificarAdminPin, 150);
  }

  function adminPinAtras() {
    _adminPinBuf = _adminPinBuf.slice(0, -1);
    _updAdminDots(_adminPinBuf.length);
  }

  function _validarLocalmenteAdmin(clave) {
    const cache = OfflineManager.getAdminCache();
    if (!cache || !cache.globalPin) return null;
    if (clave.length !== 8 || !/^\d{8}$/.test(clave)) return { ok: false, error: 'Clave debe ser de 8 dígitos' };
    const globalPart = clave.substring(0, 4);
    const userPart = clave.substring(4, 8);
    if (globalPart !== cache.globalPin) return { ok: false, error: 'Clave incorrecta' };
    const admin = (cache.adminPins || []).find(a => String(a.pin) === userPart);
    if (!admin) return { ok: false, error: 'Clave incorrecta' };
    return { ok: true, validadoPor: 'admin:' + admin.nombre + ' (offline)', idPersonal: admin.idPersonal };
  }

  async function _verificarAdminPin() {
    const clave = _adminPinBuf;
    let resultado = null;

    // Online primero — valida contra MOS y registra en auditoría
    const mosUrl = window.WH_CONFIG?.mosGasUrl || '';
    if (navigator.onLine && mosUrl) {
      try {
        const r = await fetch(mosUrl, {
          method: 'POST',
          body: JSON.stringify({
            action: 'verificarClaveAdmin',
            clave: clave,
            accion: 'REABRIR_GUIA',
            refDocumento: _pinGuiaTarget || '',
            appOrigen: 'warehouseMos',
            dispositivo: window.WH_CONFIG?.usuario || ''
          })
        });
        const j = await r.json();
        if (j?.ok && j.data) resultado = j.data;
      } catch(e) { /* fallback offline */ }
    }
    // Fallback offline
    if (!resultado) {
      const local = _validarLocalmenteAdmin(clave);
      if (local) resultado = local.ok
        ? { autorizado: true, validadoPor: local.validadoPor }
        : { autorizado: false, error: local.error };
      else resultado = { autorizado: false, error: 'Sin caché. Conecta a internet primero.' };
    }

    if (!resultado.autorizado) {
      document.getElementById('adminPinError').textContent = resultado.error || 'Clave incorrecta';
      _adminPinBuf = '';
      _updAdminDots(0);
      setTimeout(() => { document.getElementById('adminPinError').textContent = ''; }, 1800);
      return;
    }

    document.getElementById('adminPinModal').style.display = 'none';
    _adminPinBuf = '';
    _updAdminDots(0);
    const res = await API.reabrirGuia({ idGuia: _pinGuiaTarget });
    if (res.ok || res.offline) {
      // [v2.13.186 BUG reabrir] Parchear TAMBIÉN el cache wh_guias — sin esto,
      // el próximo silentRefresh (sync tick) revertía el estado a CERRADA y la
      // guía recién reabierta no dejaba editar cantidades.
      OfflineManager.patchGuiaCache?.(_pinGuiaTarget, { estado: 'ABIERTA' });
      if (_guiaActual?.idGuia === _pinGuiaTarget) {
        _guiaActual.estado = 'ABIERTA';
        _mostrarDetalleSheet(_guiaActual, false);
      }
      const idx = todas.findIndex(g => g.idGuia === _pinGuiaTarget);
      if (idx >= 0) { todas[idx].estado = 'ABIERTA'; render(_filtrarYBuscar()); }
      toast('Guía reabierta · ' + (resultado.validadoPor || 'admin'), 'ok');
    } else {
      toast('Error: ' + res.error, 'danger');
    }
  }

  // ════════════════════════════════════════════════════════════
  // AGREGAR ÍTEM — modo cámara (continua) y modo scanner HID
  // ════════════════════════════════════════════════════════════

  // Compat: abrirAgregarItem legacy → usa cámara
  function abrirAgregarItem() { abrirCamaraItem(); }

  // ── Estado sesión cámara (items escaneados en esta apertura) ─
  let _camSession      = {}; // { codigoBarra: { prod, qty } }
  let _lastScanHistory = []; // [codigoBarra, ...] — orden cronológico para undo
  let _camUnknownList  = []; // [{ code }] — códigos no encontrados en sesión actual
  let _torchOn         = false;
  let _camActive       = false; // true solo cuando GuiasView abrió el scannerModal

  function _clearCamSession() { _camSession = {}; _lastScanHistory = []; _camUnknownList = []; }

  function _addToCamList(prod) {
    const cb       = String(prod._scannedCb || prod.codigoBarra || '');
    if (!cb) return;
    const autoSum  = !!_camSession[cb];
    if (_camSession[cb]) { _camSession[cb].qty++; }
    else                 { _camSession[cb] = { prod, qty: 1 }; }
    _lastScanHistory.push(cb);
    _renderCamList();
    if (autoSum) {
      // Pulsar la fila y el número del producto ya existente
      requestAnimationFrame(() => {
        const row = document.querySelector(`[data-cam-cb="${CSS.escape(cb)}"]`);
        if (!row) return;
        row.classList.remove('cam-row-pulse');
        const qtyBtn = row.querySelector('.cam-qty-btn');
        if (qtyBtn) qtyBtn.classList.remove('cam-qty-pulse');
        void row.offsetWidth; // force reflow
        row.classList.add('cam-row-pulse');
        if (qtyBtn) qtyBtn.classList.add('cam-qty-pulse');
        setTimeout(() => {
          row.classList.remove('cam-row-pulse');
          if (qtyBtn) qtyBtn.classList.remove('cam-qty-pulse');
        }, 400);
      });
    }
  }

  function _renderCamList() {
    const list  = document.getElementById('camScannedList');
    const count = document.getElementById('scanListCount');
    if (!list) return;
    const items  = Object.values(_camSession);
    const total  = items.reduce((s, i) => s + i.qty, 0);
    if (count) count.textContent = total ? total + ' unid.' : '0 unid.';
    const hasItems = total > 0;
    const undoBtn  = document.getElementById('camUndoBtn');
    const clearBtn = document.getElementById('camClearBtn');
    if (undoBtn)  undoBtn.style.display  = _lastScanHistory.length > 0 ? 'inline-block' : 'none';
    if (clearBtn) clearBtn.style.display = hasItems ? 'inline-block' : 'none';

    // Ítems ya existentes en la guía que NO están en la sesión actual
    const detalleCompleto = _guiaActual?.detalle || [];
    const btnStyle = `width:32px;height:32px;border-radius:8px;border:1px solid #334155;
      background:#0f172a;color:#94a3b8;font-size:1.05em;font-weight:700;cursor:pointer;
      display:flex;align-items:center;justify-content:center;flex-shrink:0;
      -webkit-tap-highlight-color:transparent`;

    const guiaId = _guiaActual?.idGuia || '';
    // saving = ítem aún pendiente de GAS — solo cambia borde, no bloquea UX
    // totalQty = cantidad real en detalle (no el contador de sesión)
    let html = items.map(({ prod, qty }) => {
      const cb     = String(prod._scannedCb || prod.codigoBarra || '');
      const cbE    = escAttr(cb);
      const saving = detalleCompleto.some(d => d.codigoProducto === cb && d._local === true);
      const det    = detalleCompleto.find(d => d.codigoProducto === cb && d.observacion !== 'ANULADO');
      const totalQty = det ? (parseFloat(det.cantidadRecibida) || qty) : qty;
      return `<div data-cam-cb="${cbE}" style="display:flex;align-items:center;gap:8px;padding:9px 12px;
                border-radius:11px;background:#1e293b;
                border:1px solid ${saving ? '#475569' : '#334155'};
                margin-bottom:7px">
        <div style="flex:1;min-width:0">
          <p style="font-size:.83em;font-weight:700;color:#f1f5f9;
                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(prod.descripcion || cb)}</p>
          <p style="font-size:.67em;color:#64748b;font-family:monospace">${escHtml(cb)}</p>
        </div>
        <div style="display:flex;align-items:center;gap:4px;flex-shrink:0">
          <button onclick="GuiasView.camQtyMinus('${cbE}')"
                  style="${btnStyle}">−</button>
          <button class="cam-qty-btn" onclick="GuiasView.camQtyEdit('${cbE}')"
                  title="Toca para editar cantidad"
                  style="min-width:46px;height:32px;border-radius:8px;padding:0 10px;
                         background:#7c3aed;border:1px solid transparent;
                         color:#fff;font-size:.85em;font-weight:800;
                         cursor:pointer;-webkit-tap-highlight-color:transparent">
            ×${totalQty}
          </button>
          <button onclick="GuiasView.camQtyPlus('${cbE}')"
                  style="${btnStyle}">+</button>
        </div>
      </div>`;
    }).join('');

    // Sección "Ya en guía" — ítems confirmados no presentes en la sesión actual
    const previos = detalleCompleto.filter(d =>
      d.observacion !== 'ANULADO' &&
      !d._local &&
      !_camSession[String(d.codigoProducto || '')]
    );
    if (previos.length) {
      if (items.length) {
        html += `<p style="font-size:.62em;color:#475569;text-transform:uppercase;
                   letter-spacing:.07em;margin:10px 2px 6px;font-weight:700">Ya en guía</p>`;
      }
      html += previos.map(d => {
        const cb  = String(d.codigoProducto || '');
        const cbE = escAttr(cb);
        const nom = escHtml(d.descripcionProducto || cb);
        const qty = parseFloat(d.cantidadRecibida) || 0;
        return `<div style="display:flex;align-items:center;gap:8px;padding:8px 12px;
                  border-radius:11px;background:#0f172a;border:1px solid #1e293b;
                  margin-bottom:6px">
          <div style="flex:1;min-width:0">
            <p style="font-size:.83em;font-weight:700;color:#94a3b8;
                      white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${nom}</p>
            <p style="font-size:.67em;color:#475569;font-family:monospace">${escHtml(cb)}</p>
          </div>
          <div style="display:flex;align-items:center;gap:4px;flex-shrink:0">
            <button onclick="GuiasView.camPrevQtyMinus('${cbE}')" style="${btnStyle}">−</button>
            <span style="min-width:36px;text-align:center;font-size:.85em;
                         color:#94a3b8;font-weight:700">×${qty}</span>
            <button onclick="GuiasView.camPrevQtyPlus('${cbE}')" style="${btnStyle}">+</button>
          </div>
        </div>`;
      }).join('');
    }

    // Sección "No encontrados" — códigos escaneados sin match en catálogo
    if (_camUnknownList.length) {
      if (items.length || previos.length) {
        html += `<div style="height:1px;background:#1e293b;margin:8px 0"></div>`;
      }
      html += `<p style="font-size:.62em;color:#f87171;text-transform:uppercase;
                 letter-spacing:.07em;margin:10px 2px 6px;font-weight:700">No encontrados</p>`;
      html += _camUnknownList.map(u => {
        const cE = escAttr(u.code);
        return `<div style="display:flex;align-items:center;gap:8px;padding:9px 12px;
                  border-radius:11px;background:#2d0a0a;border:1px solid #7f1d1d;
                  margin-bottom:7px">
          <div style="flex:1;min-width:0">
            <p style="font-size:.83em;font-weight:700;color:#f87171">⚠ No registrado</p>
            <p style="font-size:.67em;color:#64748b;font-family:monospace">${escHtml(u.code)}</p>
          </div>
          <button onclick="GuiasView.abrirModalPN('${cE}','${escAttr(guiaId)}')"
                  style="flex-shrink:0;background:#fff3f3;color:#dc2626;
                         border:1px solid #fca5a5;border-radius:7px;
                         padding:4px 12px;font-size:.75em;font-weight:700;cursor:pointer">
            + Nuevo
          </button>
        </div>`;
      }).join('');
    }

    if (!html) {
      list.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;
          justify-content:center;padding:32px 20px;gap:10px;color:#334155">
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2"/>
          <rect x="7" y="7" width="10" height="10" rx="1"/>
        </svg>
        <p style="font-size:.78em;text-align:center;line-height:1.5">
          Apunta la cámara al código de barras<br>
          <span style="color:#1e293b;font-size:.9em">Los productos aparecerán aquí</span>
        </p>
      </div>`;
      return;
    }
    list.innerHTML = html;
  }

  // ── MODO CÁMARA ──────────────────────────────────────────────
  function abrirCamaraItem() {
    if (!_guiaActual) return;
    _camActive = true;
    _clearCamSession();
    _renderCamList();
    _torchOn = false;
    const tb = document.getElementById('scanTorchBtn');
    if (tb) tb.style.background = 'rgba(255,255,255,.14)';
    const picker = document.getElementById('camPicker');
    if (picker) picker.style.display = 'none';
    // Reset zoom slider
    const zoomWrap = document.getElementById('scanZoomWrap');
    const zoomRange = document.getElementById('scanZoomRange');
    if (zoomWrap) zoomWrap.style.display = 'none';
    if (zoomRange) { zoomRange.value = 1; document.getElementById('scanZoomLabel').textContent = '1×'; }
    document.getElementById('scannerModal').classList.add('open');
    _setScanStatus('ready');
    // F6: sonido scanReady para feedback "cámara lista"
    if (SoundFX.scanReady) setTimeout(() => SoundFX.scanReady(), 200);
    Scanner.start('scanVideo', _onCamResult, err => {
      toast('Error cámara: ' + err, 'danger');
      document.getElementById('scannerModal').classList.remove('open');
    }, { continuous: true, cooldown: 1500 });
    // Auto-torch: intentar encender linterna al abrir (útil en almacén)
    setTimeout(async () => {
      if (!Scanner.isActive()) return;
      // Inicializar zoom si el dispositivo lo soporta
      const zoomCaps = Scanner.getZoomCaps();
      if (zoomCaps && zoomWrap && zoomRange) {
        zoomRange.min   = zoomCaps.min;
        zoomRange.max   = zoomCaps.max;
        zoomRange.step  = zoomCaps.step || 0.1;
        zoomRange.value = zoomCaps.min;
        zoomWrap.style.display = 'flex';
      }
      // Auto-torch silencioso
      const ok = await Scanner.toggleTorch(true);
      if (ok) {
        _torchOn = true;
        if (tb) tb.style.background = 'rgba(251,191,36,.9)';
      }
    }, 900);
  }

  function cerrarCamara() {
    clearTimeout(_statusResetTimer);
    _torchOn = false;
    Scanner.stop();
    document.getElementById('scannerModal').classList.remove('open');
    const picker = document.getElementById('camPicker');
    if (picker) picker.style.display = 'none';
    // Solo volver al detalle de guía si GuiasView abrió la cámara
    if (_camActive && _guiaActual) {
      abrirSheet('sheetGuiaDetalle');
      _mostrarDetalleSheet(_guiaActual, false);
    }
    _camActive = false;
    // Toast + sonido resumen
    const sessionItems = Object.values(_camSession);
    const total = sessionItems.reduce((s, i) => s + i.qty, 0);
    if (total > 0) {
      const prods = sessionItems.length;
      SoundFX.done();
      toast(`✓ ${total} ítem${total !== 1 ? 's' : ''} · ${prods} producto${prods !== 1 ? 's' : ''} agregados`, 'ok', 2800);
    }
  }

  async function toggleTorch() {
    _torchOn = !_torchOn;
    const ok = await Scanner.toggleTorch(_torchOn);
    const btn = document.getElementById('scanTorchBtn');
    if (!ok) {
      _torchOn = false;
      toast('Linterna no disponible en este dispositivo', 'warn', 2500);
    }
    if (btn) btn.style.background = (_torchOn && ok)
      ? 'rgba(251,191,36,.9)' : 'rgba(255,255,255,.15)';
  }

  // Barra de estado entre cámara y lista — feedback del último escaneo
  let _statusResetTimer = null;

  function _setScanStatus(type, text, rawCod) {
    const bar = document.getElementById('scanStatusBar');
    if (!bar) return;
    clearTimeout(_statusResetTimer);

    // F6: aplicar clase de estado de cámara al bar (border colorido por estado)
    bar.classList.remove('cam-state-listo','cam-state-procesando','cam-state-preguntando','cam-state-descubrir','cam-state-bloqueado');
    const stateClass = {
      ready: 'cam-state-listo',
      ok: 'cam-state-listo',
      prefijo: 'cam-state-preguntando',
      no_existe: 'cam-state-descubrir',
      blocked: 'cam-state-bloqueado',
      procesando: 'cam-state-procesando'
    }[type];
    if (stateClass) bar.classList.add(stateClass);

    if (type === 'ready') {
      bar.innerHTML = '<span style="color:#334155;font-size:.75em">— listo para escanear —</span>';
      bar.style.background = '#0f172a';
      return;
    }

    const guiaId = _guiaActual?.idGuia || '';
    const cfgs = {
      ok:       { bg: '#022c22', col: '#34d399', icon: '✓', dur: 2200 },
      prefijo:  { bg: '#2c1a00', col: '#fb923c', icon: '↕', dur: 0   },
      no_existe:{ bg: '#2d0a0a', col: '#f87171', icon: '⚠', dur: 0   }
    };
    const c = cfgs[type] || cfgs.ok;
    const newBtn = type === 'no_existe'
      ? `<button onclick="GuiasView.abrirModalPN('${escAttr(rawCod||text)}','${escAttr(guiaId)}')"
           style="flex-shrink:0;background:#fff3f3;color:#dc2626;border:1px solid #fca5a5;
                  border-radius:7px;padding:4px 10px;font-size:.71em;font-weight:700;cursor:pointer">
           + Nuevo
         </button>` : '';
    bar.style.background = c.bg;
    bar.innerHTML = `
      <span style="color:${c.col};font-size:.82em;font-weight:700;flex-shrink:0">${c.icon}</span>
      <span style="color:${c.col};font-size:.76em;flex:1;white-space:nowrap;overflow:hidden;
                   text-overflow:ellipsis;margin-left:6px">${escHtml(text)}</span>
      ${newBtn}`;
    if (c.dur > 0) {
      _statusResetTimer = setTimeout(() => _setScanStatus('ready'), c.dur);
    }
  }

  // Callback del scanner continuo — procesa código sin cerrar cámara
  function _onCamResult(cod) {
    const codStr = normCb(cod);
    if (!codStr) return;
    // Si el picker está abierto, no interrumpir — el usuario está eligiendo
    const picker = document.getElementById('camPicker');
    if (picker?.style.display === 'flex') return;

    const candidatos = _buscarCandidatos(codStr);
    const esIngreso = String(_guiaActual?.tipo || '').toUpperCase().startsWith('INGRESO');

    if (!candidatos.length) {
      // Acumular en lista persistente (deduplicar por código)
      if (!_camUnknownList.find(u => u.code === codStr)) {
        _camUnknownList.push({ code: codStr });
        _renderCamList();
      }
      _setScanStatus('no_existe', codStr + ' · no registrado');
      // Caso 3: en INGRESO, código nuevo → sonido scanNuevo + vibración curiosa
      if (esIngreso && SoundFX.scanNuevo) {
        SoundFX.scanNuevo();
        vibrate([40, 30, 40]);
      } else {
        SoundFX.warn();
        vibrate([60, 30, 60]);
      }
      return;
    }
    if (candidatos[0]._exacto) {
      const autoSum = _agregarProductoDirecto(candidatos[0], false);
      _addToCamList(candidatos[0]);
      _setScanStatus('ok', candidatos[0].descripcion || (candidatos[0]._scannedCb || candidatos[0].codigoBarra));
      if (!autoSum) SoundFX.beep();
      return;
    }
    // Caso 2: prefijo → picker overlay + sonido scanIncompleto
    _setScanStatus('prefijo', 'Prefijo · ' + candidatos.length + ' productos coinciden');
    if (SoundFX.scanIncompleto) SoundFX.scanIncompleto();
    vibrate(30);
    _mostrarCamPicker(candidatos, codStr);
  }

  function _mostrarCamPicker(candidatos, codStr) {
    document.getElementById('camPickerCod').textContent = codStr;
    const esIngreso = String(_guiaActual?.tipo || '').toUpperCase().startsWith('INGRESO');
    const noneBtn = esIngreso
      ? `<button onclick="GuiasView.cerrarCamPicker();GuiasView.abrirModalPN('${escAttr(codStr)}','${escAttr(_guiaActual?.idGuia || '')}')"
                style="width:100%;text-align:left;padding:11px 13px;border-radius:11px;
                       border:1.5px dashed #7c3aed;margin-top:8px;background:rgba(124,58,237,.08);
                       display:flex;align-items:center;gap:10px;cursor:pointer;
                       -webkit-tap-highlight-color:transparent">
          <div style="flex-shrink:0;width:32px;height:32px;border-radius:8px;background:#7c3aed33;
                      display:flex;align-items:center;justify-content:center;font-size:1.1em">🆕</div>
          <div style="flex:1">
            <p style="font-size:.83em;font-weight:800;color:#c084fc">Ninguno · es otro código</p>
            <p style="font-size:.69em;color:#94a3b8;margin-top:2px">Registrar ${codStr} como nuevo</p>
          </div>
        </button>`
      : '';
    document.getElementById('camPickerList').innerHTML = candidatos.map(p => {
      const cb      = String(p._scannedCb || p.codigoBarra || '');
      const display = String(p.codigoBarra || cb);
      const cbHtml  = display.startsWith(codStr)
        ? `<strong style="color:#fbbf24">${escHtml(codStr)}</strong>${escHtml(display.slice(codStr.length))}`
        : escHtml(display);
      return `<button onclick="GuiasView.seleccionarItemCamara('${escAttr(cb)}')"
              style="width:100%;text-align:left;padding:11px 13px;border-radius:11px;
                     border:1px solid #1e293b;margin-bottom:7px;
                     background:#1e293b;display:flex;align-items:center;gap:10px;
                     cursor:pointer;-webkit-tap-highlight-color:transparent"
              ontouchstart="this.style.borderColor='#7c3aed';this.style.background='#1e1b4b'"
              ontouchend="this.style.borderColor='#1e293b';this.style.background='#1e293b'">
        <div style="flex-shrink:0;width:32px;height:32px;border-radius:8px;background:#7c3aed22;
                    display:flex;align-items:center;justify-content:center">
          <span style="font-size:.75em;color:#a78bfa;font-weight:700">CB</span>
        </div>
        <div style="flex:1;min-width:0">
          <p style="font-size:.83em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || cb))}</p>
          <p style="font-size:.69em;color:#64748b;font-family:monospace;margin-top:2px">${cbHtml}</p>
        </div>
      </button>`;
    }).join('') + noneBtn;
    const picker = document.getElementById('camPicker');
    picker.style.display = 'flex';
  }

  function cerrarCamPicker() {
    const picker = document.getElementById('camPicker');
    if (picker) picker.style.display = 'none';
    _setScanStatus('ready');
  }

  function seleccionarItemCamara(scannedCb) {
    const picker = document.getElementById('camPicker');
    if (picker) picker.style.display = 'none';
    const candidatos = _buscarCandidatos(scannedCb);
    const prod = candidatos[0];
    if (!prod) return;
    _agregarProductoDirecto(prod, true);
    _addToCamList(prod);
    _setScanStatus('ok', prod.descripcion || (prod._scannedCb || prod.codigoBarra));
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
    codStr = normCb(codStr);
    const candidatos = _buscarCandidatos(codStr);
    document.getElementById('hidPicker').style.display = 'none';

    if (!candidatos.length) {
      _updateHidDisplay('⚠ No existe: ' + codStr);
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
    // Prefijo → picker dentro del sheet
    document.getElementById('hidPickerCod').textContent = codStr;
    document.getElementById('hidPickerList').innerHTML = candidatos.map(p => {
      const cb      = String(p._scannedCb || p.codigoBarra || '');
      const display = String(p.codigoBarra || cb);
      const cbHtml  = display.startsWith(codStr)
        ? `<strong style="color:#fbbf24">${escHtml(codStr)}</strong>${escHtml(display.slice(codStr.length))}`
        : escHtml(display);
      return `<button onclick="GuiasView.seleccionarItemHid('${escAttr(cb)}')"
              style="width:100%;text-align:left;padding:8px 10px;border-radius:9px;
                     border:1px solid rgba(124,58,237,.2);margin-bottom:5px;
                     background:rgba(124,58,237,.06);display:flex;align-items:center;gap:8px;
                     cursor:pointer;-webkit-tap-highlight-color:transparent">
        <div style="flex:1;min-width:0">
          <p style="font-size:.82em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || cb))}</p>
          <p style="font-size:.69em;color:#64748b;font-family:monospace">${cbHtml}</p>
        </div>
      </button>`;
    }).join('');
    document.getElementById('hidPicker').style.display = 'block';
  }

  function seleccionarItemHid(scannedCb) {
    document.getElementById('hidPicker').style.display = 'none';
    const candidatos = _buscarCandidatos(scannedCb);
    const prod = candidatos[0];
    if (!prod) return;
    _agregarProductoDirecto(prod, true);
    _agregarItemHidList(prod);
    _updateHidDisplay('— listo para escanear —');
    setTimeout(() => _enfocarHid(), 100);
  }

  // Lista HID: auto-suma si el mismo producto se escanea de nuevo
  function _agregarItemHidList(prod) {
    const list = document.getElementById('hidProductList');
    if (!list) return;
    const cb = String(prod._scannedCb || prod.codigoBarra || prod.idProducto || '');
    const existing = list.querySelector(`[data-hid-cb="${CSS.escape(cb)}"]`);
    if (existing) {
      const countEl = existing.querySelector('.hid-count');
      const cur = parseInt(countEl?.textContent?.replace('×', '') || '1');
      if (countEl) countEl.textContent = '×' + (cur + 1);
      existing.classList.remove('item-slide-in');
      requestAnimationFrame(() => existing.classList.add('item-slide-in'));
      return;
    }
    const div = document.createElement('div');
    div.className = 'item-slide-in';
    div.dataset.hidCb = cb;
    div.style.cssText = 'display:flex;align-items:center;gap:9px;padding:9px 11px;' +
      'border-radius:11px;background:#1e293b;border:1px solid #334155;margin-bottom:6px';
    div.innerHTML = `<span style="color:#10b981;font-size:1.1em;flex-shrink:0">✓</span>
      <div style="flex:1;min-width:0">
        <p style="font-size:.82em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(prod.descripcion || cb))}</p>
        <p style="font-size:.69em;color:#64748b;font-family:monospace">${escHtml(cb)}</p>
      </div>
      <span class="hid-count" style="background:#7c3aed;border-radius:6px;padding:3px 9px;
            font-size:.8em;font-weight:700;color:#fff;flex-shrink:0">×1</span>`;
    list.insertBefore(div, list.firstChild);
  }

  // ── Búsqueda de candidatos — solo codigoBarra de canónico + equivalencias.
  // Aplica regla de oro WH (ver comentario header en _buscarDespCandidatos).
  // En guías de INGRESO el caller decide qué hacer si retorna []:
  //   - Si exacto → procesar normal
  //   - Si prefijo (length>1) → operador elige cuál
  //   - Si vacío → ofrecer "registrar producto nuevo" (registrarPN)
  // En guías de SALIDA, vacío = error "no existe en catálogo".
  // _scannedCb: código que físicamente se escaneó (puede ser equiv ≠ prod canónico).
  function _buscarCandidatos(codStr) {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    const cNorm  = normCb(codStr);
    if (!cNorm) return [];

    // 1. Exacto en PRODUCTOS_MASTER → código escaneado ES el canónico
    const exacto = prods.find(p => String(p.codigoBarra || '').trim().toUpperCase() === cNorm);
    if (exacto) return [{ ...exacto, _exacto: true }];

    // 2. Exacto en EQUIVALENCIAS → resolver al producto base (factor=1); guardar código escaneado
    const equiv = equivs.find(e => String(e.codigoBarra || '').trim().toUpperCase() === cNorm);
    if (equiv) {
      const skuB = String(equiv.skuBase || '').trim().toUpperCase();
      const prod = prods.find(p =>
        parseFloat(p.factorConversion || 1) === 1 &&
        p.estado !== '0' && p.estado !== 0 &&
        (String(p.idProducto  || '').trim().toUpperCase() === skuB ||
         String(p.skuBase     || '').trim().toUpperCase() === skuB ||
         String(p.codigoBarra || '').trim().toUpperCase() === skuB)
      );
      if (prod) return [{ ...prod, _exacto: true, _scannedCb: cNorm }];
    }

    // 3. Prefijo en PRODUCTOS_MASTER (mín. 3 chars)
    if (cNorm.length >= 3) {
      const porMaestro = prods
        .filter(p => String(p.codigoBarra || '').trim().toUpperCase().startsWith(cNorm))
        .map(p => ({ ...p })); // copias

      // Prefijo en EQUIVALENCIAS → resolver al producto base (factor=1, sin duplicar por idProducto)
      const idsYa = new Set(porMaestro.map(p => p.idProducto));
      equivs.filter(e => String(e.codigoBarra || '').trim().toUpperCase().startsWith(cNorm)).forEach(e => {
        const skuB = String(e.skuBase || '').trim().toUpperCase();
        const base  = prods.find(p =>
          parseFloat(p.factorConversion || 1) === 1 &&
          p.estado !== '0' && p.estado !== 0 &&
          (String(p.idProducto  || '').trim().toUpperCase() === skuB ||
           String(p.skuBase     || '').trim().toUpperCase() === skuB ||
           String(p.codigoBarra || '').trim().toUpperCase() === skuB)
        );
        if (base && !idsYa.has(base.idProducto)) {
          porMaestro.push({ ...base, _scannedCb: String(e.codigoBarra).trim() });
          idsYa.add(base.idProducto);
        }
      });

      if (porMaestro.length) return porMaestro.slice(0, 10);
    }

    return [];
  }

  function _agregarProductoDirecto(prod, indirecto) {
    if (!_guiaActual) return;
    // Clave = código escaneado (equiv o master) — GAS acepta ambos y stock es independiente por barcode
    const cb   = String(prod._scannedCb || prod.codigoBarra || prod.idProducto || '');
    const desc = prod.descripcion || prod.nombre || cb;

    // Auto-suma: si el mismo codigoBarra ya está en detalle → incrementar
    if (!_guiaActual.detalle) _guiaActual.detalle = [];
    const existing = _guiaActual.detalle.find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );
    if (existing) {
      // Si el ítem está en edición inline, usar _selQty como base (puede haber cambios no guardados)
      const activeItems = _guiaActual.detalle.filter(d => d.observacion !== 'ANULADO');
      const existingIdx = activeItems.indexOf(existing);
      const baseQty = (_selIdx >= 0 && existingIdx === _selIdx)
        ? (parseFloat(_selQty) || parseFloat(existing.cantidadRecibida) || 0)
        : (parseFloat(existing.cantidadRecibida) || 0);
      existing.cantidadRecibida = baseQty + 1;
      if (_selIdx >= 0 && existingIdx === _selIdx) {
        _selQty = existing.cantidadRecibida;
        _selOrigQty = existing.cantidadRecibida;
      }
      existing._saving = true;
      existing._saveFailed = false;
      _mostrarDetalleSheet(_guiaActual, false);
      vibrate(12);
      SoundFX.beepDouble();
      if (existing.idDetalle && !existing._local) {
        API.actualizarCantidadDetalle({
          idDetalle: existing.idDetalle,
          cantidadRecibida: existing.cantidadRecibida
        }).then(r => {
          existing._saving = false;
          if (r && r.ok) {
            existing._saveFailed = false;
            if (SoundFX.savedTick) SoundFX.savedTick();
          } else if (!r || !r.offline) {
            existing._saveFailed = true;
          }
          _mostrarDetalleSheet(_guiaActual, false);
        }).catch(() => {
          existing._saving = false;
          existing._saveFailed = true;
          _mostrarDetalleSheet(_guiaActual, false);
        });
      } else {
        existing._saving = false;
      }
      return true;
    }

    const localId = 'DL' + Date.now();
    const itemOptimista = {
      idDetalle: localId, idGuia: _guiaActual.idGuia,
      codigoProducto: cb,
      descripcionProducto: desc,
      cantidadEsperada: 0, cantidadRecibida: 1,
      precioUnitario: 0, fechaVencimiento: '', observacion: '',
      _local: true, _indirect: !!indirecto
    };
    _guiaActual.detalle.push(itemOptimista);
    _mostrarDetalleSheet(_guiaActual, false);
    requestAnimationFrame(() => {
      const el = document.querySelector(`[data-det-id="${localId}"]`);
      if (el) el.classList.add('item-slide-in');
    });
    toast((indirecto ? '↕ ' : '✓ ') + desc, 'ok', 1500);
    vibrate(15);

    const _idGuia = _guiaActual.idGuia;
    // [v2.13.215][FIX 40x #2 BLOQUEANTE] localId DETERMINÍSTICO atado a la card
    // optimista (`localId` = 'DL'+ts, estable por línea). Sin esto, post() inyecta
    // un _genLocalId() nuevo por invocación → el dedup del RPC (por local_id) no
    // protege contra doble-invocación lógica → doble fila en wh.guia_detalle.
    API.agregarDetalle({
      idGuia: _idGuia,
      codigoProducto: cb,
      cantidadEsperada: 0, cantidadRecibida: 1,
      precioUnitario: 0, fechaVencimiento: '',
      localId: 'DET_' + localId
    }).then(res => {
      // [v2.13.183] Guard: si el operador cambió de guía mientras el POST viajaba,
      // _guiaActual ya es otra → no tocar su estado (como hace verDetalle).
      if (!_guiaActual || _guiaActual.idGuia !== _idGuia) return;
      if (res.ok && !res.offline) {
        const idx = _guiaActual.detalle?.findIndex(d => d.idDetalle === localId);
        if (idx >= 0) {
          // Preservar cantidad local: puede haberse incrementado mientras GAS respondía
          const localQty = parseFloat(_guiaActual.detalle[idx].cantidadRecibida) || 1;
          const gasQty   = parseFloat(res.data.cantidadRecibida) || 1;
          const itemFinal = {
            ...res.data,
            idGuia: res.data.idGuia || _idGuia,
            descripcionProducto: res.data.descripcionProducto || desc,
            cantidadRecibida: localQty,
            _local: false, _indirect: !!indirecto
          };
          _guiaActual.detalle[idx] = itemFinal;
          _mostrarDetalleSheet(_guiaActual, false);
          OfflineManager.addDetalleCache(itemFinal);
          // Si el local fue incrementado mientras GAS estaba en vuelo → sincronizar
          if (localQty > gasQty) {
            _avisarFalloDetalle(API.actualizarCantidadDetalle({
              idDetalle: res.data.idDetalle,
              cantidadRecibida: localQty
            }), 'la cantidad');
          }
        }
        _renderCamList(); // quitar ⏳
      } else if (!res.offline) {
        // GAS rechazó → revertir detalle y limpiar sesión cámara
        _guiaActual.detalle = _guiaActual.detalle.filter(d => d.idDetalle !== localId);
        _mostrarDetalleSheet(_guiaActual, false);

        const isNotFound = res.error === 'PRODUCTO_NO_ENCONTRADO';
        if (isNotFound) {
          // Producto no existe en GAS: eliminar TODA la entrada de sesión
          // (auto-sums sobre un local que nunca se confirmó también son inválidos)
          delete _camSession[cb];
          // Mover a la lista de desconocidos para que el usuario pueda registrarlo
          if (!_camUnknownList.find(u => u.code === cb)) _camUnknownList.push({ code: cb });
        } else if (_camSession[cb]) {
          // Otro error (red, GAS caído, etc.) → solo descontar 1 del optimista
          _camSession[cb].qty = Math.max(0, _camSession[cb].qty - 1);
          if (_camSession[cb].qty === 0) delete _camSession[cb];
        }
        _renderCamList();
        // Feedback: barra de estado si cámara abierta, toast si no
        const camOpen = document.getElementById('scannerModal')?.classList.contains('open');
        const errMsg  = isNotFound
          ? 'No registrado en catálogo' : (res.error || res.mensaje || 'Error al guardar');
        SoundFX.error();
        vibrate([80, 40, 80]);
        if (camOpen) {
          // Pasar cb como rawCod para que el botón "+Nuevo" use el código real
          _setScanStatus('no_existe', errMsg + ' · ' + desc, cb);
        } else {
          toast(res.error === 'PRODUCTO_NO_ENCONTRADO'
            ? 'Producto no registrado en el sistema' : 'Error: ' + errMsg,
            res.error === 'PRODUCTO_NO_ENCONTRADO' ? 'warn' : 'danger', 4000);
        }
      }
    }).catch(() => {
      // [v2.13.185] Excepción/red: revertir el optimista y avisar (no silenciar).
      if (_guiaActual && _guiaActual.idGuia === _idGuia) {
        _guiaActual.detalle = (_guiaActual.detalle || []).filter(d => d.idDetalle !== localId);
        _mostrarDetalleSheet(_guiaActual, false);
      }
      if (_camSession[cb]) {
        _camSession[cb].qty = Math.max(0, (_camSession[cb].qty || 1) - 1);
        if (_camSession[cb].qty === 0) delete _camSession[cb];
      }
      _renderCamList();
      try { SoundFX.error(); } catch(_){}
      toast('⚠ Sin conexión — el escaneo no se guardó', 'warn', 4000);
    });
  }

  // ── Controles de cantidad en la lista de sesión cámara ───────

  function camQtyPlus(cb) {
    const entry = _camSession[cb];
    if (!entry) return;
    entry.qty++;
    const item = (_guiaActual?.detalle || []).find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );
    if (item) {
      item.cantidadRecibida = (parseFloat(item.cantidadRecibida) || 0) + 1;
      if (item.idDetalle && !item._local) {
        item._saving = true;
        item._saveFailed = false;
        API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida })
          .then(r => {
            item._saving = false;
            if (r && r.ok) {
              if (SoundFX.savedTick) SoundFX.savedTick();
            } else if (!r || !r.offline) {
              item._saveFailed = true;
            }
            _mostrarDetalleSheet(_guiaActual, false);
          })
          .catch(() => {
            item._saving = false;
            item._saveFailed = true;
            _mostrarDetalleSheet(_guiaActual, false);
          });
      }
      _mostrarDetalleSheet(_guiaActual, false);
    }
    _renderCamList();
    SoundFX.beepDouble();
    vibrate(10);
  }

  // Anti-anulación accidental: si qty=1 y presiona −, requiere doble-tap en 3s
  let _pendingMinusCb = null;
  let _pendingMinusTimer = null;

  function camQtyMinus(cb) {
    const entry = _camSession[cb];
    if (!entry) return;
    const item = (_guiaActual?.detalle || []).find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );

    if (entry.qty <= 1) {
      // Confirmación: necesita doble-tap en 3s para anular
      if (_pendingMinusCb !== cb) {
        _pendingMinusCb = cb;
        clearTimeout(_pendingMinusTimer);
        _pendingMinusTimer = setTimeout(() => { _pendingMinusCb = null; }, 3000);
        toast('Toca − otra vez en 3s para eliminar', 'warn', 3000);
        try { SoundFX.warn(); } catch(e) {}
        vibrate([20, 30, 20]);
        return;
      }
      // Confirmado: anular
      clearTimeout(_pendingMinusTimer);
      _pendingMinusCb = null;
      delete _camSession[cb];
      if (item) {
        if (item.idDetalle && !item._local) {
          item.observacion = 'ANULADO'; item.cantidadRecibida = 0;
          _avisarFalloDetalle(API.anularDetalle({ idDetalle: item.idDetalle }), 'la anulación');
        } else {
          _guiaActual.detalle = _guiaActual.detalle.filter(d => d !== item);
        }
        _mostrarDetalleSheet(_guiaActual, false);
      }
    } else {
      // qty > 1: decremento normal sin confirmación
      _pendingMinusCb = null;
      entry.qty--;
      if (item) {
        item.cantidadRecibida = Math.max(0, (parseFloat(item.cantidadRecibida) || 0) - 1);
        if (item.idDetalle && !item._local) {
          _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida }), 'la cantidad');
        }
        _mostrarDetalleSheet(_guiaActual, false);
      }
    }
    _renderCamList();
    vibrate(8);
  }

  async function camQtyEdit(cb) {
    const entry = _camSession[cb];
    if (!entry) return;
    // [v2.x.x] Modal custom — antes prompt() nativo bloqueaba el scanner
    const input = await _whPrompt('Nueva cantidad (decimales: usa punto):', String(entry.qty), {
      titulo: entry.prod.descripcion || cb,
      inputMode: 'decimal',
      maxlength: 10
    });
    if (input === null) return; // cancelado
    const newQty = parseFloat(input.replace(',', '.'));
    if (isNaN(newQty) || newQty < 0) { toast('Cantidad inválida', 'warn', 2000); return; }
    if (newQty === 0) { camQtyMinus(cb); return; }
    const diff = newQty - entry.qty;
    if (diff === 0) return;
    entry.qty = newQty;
    const item = (_guiaActual?.detalle || []).find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );
    if (item) {
      item.cantidadRecibida = Math.max(0, (parseFloat(item.cantidadRecibida) || 0) + diff);
      if (item.idDetalle && !item._local) {
        _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida }), 'la cantidad');
      }
      _mostrarDetalleSheet(_guiaActual, false);
    }
    _renderCamList();
    SoundFX.beep();
    vibrate(10);
  }

  // ── +/− sobre ítems "Ya en guía" (previos confirmados) ──────────
  function camPrevQtyPlus(cb) {
    const item = (_guiaActual?.detalle || []).find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );
    if (!item) return;
    item.cantidadRecibida = (parseFloat(item.cantidadRecibida) || 0) + 1;
    if (item.idDetalle && !item._local) {
      _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida }), 'la cantidad');
    }
    _mostrarDetalleSheet(_guiaActual, false);
    _renderCamList();
    SoundFX.beepDouble(); vibrate(8);
  }

  function camPrevQtyMinus(cb) {
    const item = (_guiaActual?.detalle || []).find(d =>
      d.codigoProducto === cb && d.observacion !== 'ANULADO'
    );
    if (!item) return;
    const cur = parseFloat(item.cantidadRecibida) || 0;
    if (cur <= 1) return;
    item.cantidadRecibida = cur - 1;
    if (item.idDetalle && !item._local) {
      _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida }), 'la cantidad');
    }
    _mostrarDetalleSheet(_guiaActual, false);
    _renderCamList();
    vibrate(8);
  }

  function _rescanear() {
    abrirCamaraItem();
  }

  // ── Controles adicionales cámara ──────────────────────────────

  function camUndoLast() {
    if (!_lastScanHistory.length) return;
    const cb = _lastScanHistory.pop();
    const entry = _camSession[cb];
    if (!entry) { _renderCamList(); return; }
    if (entry.qty <= 1) {
      delete _camSession[cb];
      // Revertir en detalle
      const item = (_guiaActual?.detalle || []).find(d =>
        d.codigoProducto === cb && d.observacion !== 'ANULADO'
      );
      if (item) {
        if (item.idDetalle && !item._local) {
          item.observacion = 'ANULADO'; item.cantidadRecibida = 0;
          _avisarFalloDetalle(API.anularDetalle({ idDetalle: item.idDetalle }), 'la anulación');
        } else {
          _guiaActual.detalle = _guiaActual.detalle.filter(d => d !== item);
        }
        _mostrarDetalleSheet(_guiaActual, false);
      }
    } else {
      entry.qty--;
      const item = (_guiaActual?.detalle || []).find(d =>
        d.codigoProducto === cb && d.observacion !== 'ANULADO'
      );
      if (item) {
        item.cantidadRecibida = Math.max(0, (parseFloat(item.cantidadRecibida) || 0) - 1);
        if (item.idDetalle && !item._local) {
          _avisarFalloDetalle(API.actualizarCantidadDetalle({ idDetalle: item.idDetalle, cantidadRecibida: item.cantidadRecibida }), 'la cantidad');
        }
        _mostrarDetalleSheet(_guiaActual, false);
      }
    }
    _renderCamList();
    vibrate(15);
    toast('↩ Deshecho', 'warn', 1200);
  }

  async function camLimpiarTodo() {
    if (!Object.keys(_camSession).length) return;
    const total = Object.values(_camSession).reduce((s, i) => s + i.qty, 0);
    if (!await _whConfirm(`¿Limpiar los ${total} ítems de esta sesión?\n\nLos ya guardados en GAS quedan en la guía.`, { warning: true, titulo: 'Limpiar sesión', okText: 'Limpiar' })) return;
    // Anular solo los locales (aún no confirmados por GAS)
    if (_guiaActual?.detalle) {
      const cbs = new Set(Object.keys(_camSession));
      _guiaActual.detalle = _guiaActual.detalle.map(d => {
        if (cbs.has(d.codigoProducto) && d._local) return null; // quitar local
        return d;
      }).filter(Boolean);
      _mostrarDetalleSheet(_guiaActual, false);
    }
    _clearCamSession();
    _renderCamList();
    vibrate(20);
  }

  function camSetZoom(val) {
    const v = parseFloat(val);
    Scanner.setZoom(v);
    const lbl = document.getElementById('scanZoomLabel');
    if (lbl) lbl.textContent = v.toFixed(1) + '×';
  }

  // Swipe-left en ítems del detalle de guía → anular
  function _initSwipeGuia() {
    const container = document.getElementById('guiaDetItems');
    if (!container || container._swipeInit) return;
    container._swipeInit = true;
    let sx = 0, sy = 0, el = null, moved = false;
    container.addEventListener('touchstart', e => {
      const item = e.target.closest('[data-det-idx]');
      if (!item) { el = null; return; }
      sx = e.touches[0].clientX; sy = e.touches[0].clientY;
      el = item; moved = false;
      el.style.transition = 'none';
    }, { passive: true });
    container.addEventListener('touchmove', e => {
      if (!el) return;
      const dx = e.touches[0].clientX - sx;
      const dy = e.touches[0].clientY - sy;
      if (Math.abs(dy) > Math.abs(dx) + 8) { el.style.transform = ''; el = null; return; }
      if (dx > 0) { el.style.transform = ''; el = null; return; }
      moved = true;
      const clamped = Math.max(dx, -110);
      el.style.transform = `translateX(${clamped}px)`;
      el.style.background = dx < -55 ? 'rgba(220,38,38,.22)' : '';
    }, { passive: true });
    container.addEventListener('touchend', e => {
      if (!el) return;
      const dx = e.changedTouches[0].clientX - sx;
      el.style.transition = 'transform .18s ease, background .18s';
      if (moved && dx < -80) {
        el.style.transform = 'translateX(-110%)';
        el.style.opacity = '0';
        const idx = parseInt(el.dataset.detIdx);
        vibrate([40, 20, 40]);
        setTimeout(() => { if (!isNaN(idx)) inlineDelete(idx); }, 180);
      } else {
        el.style.transform = '';
        el.style.background = '';
      }
      el = null;
    }, { passive: true });
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

    // Barcode input + toggle
    const codInput    = document.getElementById('pnCodigoBarra');
    const autoWrap    = document.getElementById('pnAutoGenWrap');
    const autoToggle  = document.getElementById('pnAutoGenToggle');
    const autoThumb   = document.getElementById('pnAutoGenThumb');
    const hasCode     = !!_pnCodigo;

    if (codInput) {
      codInput.value    = _pnCodigo || '';
      codInput.readOnly = false;
      codInput.style.opacity = '1';
    }
    // Si viene con código escaneado: ocultar toggle; si es manual: mostrar toggle
    if (autoWrap)   autoWrap.style.display  = hasCode ? 'none' : 'flex';
    // Siempre resetear toggle a OFF
    if (autoToggle) { autoToggle.dataset.on = '0'; autoToggle.style.background = '#334155'; }
    if (autoThumb)  { autoThumb.style.left = '2px'; autoThumb.style.background = '#64748b'; }

    const cant = document.getElementById('pnCantidad');
    if (cant) cant.value = '1';
    const fv = document.getElementById('pnFechaVenc');
    if (fv) fv.value = '';
    const obs = document.getElementById('pnObservaciones');
    if (obs) { obs.value = ''; obs.style.borderColor = '#334155'; obs.style.boxShadow = ''; }
    const marca = document.getElementById('pnMarca');
    if (marca) marca.value = ''; // [v2.13.48] reset campo marca
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
    // Pausar scanner continuo para que no dispare sonidos mientras el modal está abierto
    if (document.getElementById('scannerModal')?.classList.contains('open')) Scanner.stop();
    // [v2.13.188] Foco automático: si falta código, al input de código; si ya
    // viene escaneado, a la descripción (campo obligatorio).
    setTimeout(() => {
      try { (hasCode ? (obs || codInput) : (codInput || obs))?.focus(); } catch(_){}
    }, 120);
  }

  function _pnCodigoChanged() {
    const v = (document.getElementById('pnCodigoBarra')?.value || '').trim();
    _pnCodigo = v;
    // Si el usuario escribe manualmente, apagar auto-gen
    const toggle = document.getElementById('pnAutoGenToggle');
    if (toggle?.dataset.on === '1') _pnResetAutoGen();
  }

  function _pnToggleAutoGen() {
    const toggle  = document.getElementById('pnAutoGenToggle');
    const thumb   = document.getElementById('pnAutoGenThumb');
    const codInput = document.getElementById('pnCodigoBarra');
    if (!toggle) return;
    const isOn = toggle.dataset.on === '1';
    if (!isOn) {
      // Activar: generar código NLEV + timestamp
      const gen = 'NLEV' + Date.now();
      _pnCodigo = gen;
      if (codInput) { codInput.value = gen; codInput.readOnly = true; codInput.style.opacity = '.55'; }
      toggle.dataset.on  = '1';
      toggle.style.background = '#f59e0b';
      if (thumb) { thumb.style.left = '20px'; thumb.style.background = '#0f172a'; }
    } else {
      _pnResetAutoGen();
    }
  }

  function _pnResetAutoGen() {
    const toggle  = document.getElementById('pnAutoGenToggle');
    const thumb   = document.getElementById('pnAutoGenThumb');
    const codInput = document.getElementById('pnCodigoBarra');
    _pnCodigo = (codInput?.value || '').trim();
    if (codInput) { codInput.readOnly = false; codInput.style.opacity = '1'; }
    if (toggle) { toggle.dataset.on = '0'; toggle.style.background = '#334155'; }
    if (thumb)  { thumb.style.left = '2px'; thumb.style.background = '#64748b'; }
  }

  function cerrarModalPN() {
    const modal = document.getElementById('modalProductoNuevo');
    if (modal) modal.classList.remove('open');
    // Reanudar scanner si la cámara sigue abierta
    if (document.getElementById('scannerModal')?.classList.contains('open')) {
      Scanner.start('scanVideo', _onCamResult,
        err => toast('Error cámara: ' + err, 'danger'),
        { continuous: true, cooldown: 1500 }
      );
      _setScanStatus('ready');
    }
  }

  function _pnFotoSeleccionada(input) {
    // [v2.13.49] Handler más robusto + feedback claro si algo falla.
    try {
      const file = input.files && input.files[0];
      if (!file) return; // usuario canceló el picker — no es error
      // Validación tamaño (10MB max) para evitar payloads gigantes
      const sizeMB = file.size / 1024 / 1024;
      if (sizeMB > 10) {
        try { toast('⚠ Foto muy grande (' + sizeMB.toFixed(1) + 'MB) · máx 10MB', 'warn', 4000); } catch(_){}
        input.value = '';
        return;
      }
      const reader = new FileReader();
      reader.onerror = (err) => {
        console.error('[PN foto] FileReader error', err);
        try { toast('❌ No se pudo leer la foto · intenta otra', 'error', 4000); } catch(_){}
      };
      reader.onload = e => {
        try {
          const dataUrl = e.target.result;
          _pnFotoBase64 = String(dataUrl || '').split(',')[1] || '';
          _pnFotoMime   = file.type || 'image/jpeg';
          const img = document.getElementById('pnFotoImg');
          if (img) img.src = dataUrl;
          const prev = document.getElementById('pnFotoPreview');
          if (prev) prev.style.display = 'block';
          const btnTxt = document.getElementById('pnFotoBtn');
          if (btnTxt) btnTxt.textContent = '✓ Foto lista — toca para cambiar';
          try { vibrate?.(15); } catch(_){}
          try { toast('📷 Foto cargada (' + sizeMB.toFixed(1) + 'MB)', 'ok', 2000); } catch(_){}
        } catch(eIn) {
          console.error('[PN foto] onload handler error', eIn);
          try { toast('❌ Error procesando la foto', 'error', 4000); } catch(_){}
        }
      };
      reader.readAsDataURL(file);
    } catch(e) {
      console.error('[PN foto] handler outer error', e);
      try { toast('❌ Error al subir foto: ' + (e.message || e), 'error', 4000); } catch(_){}
    }
  }

  let _pnSubmitting = false;
  async function confirmarRegistrarPN() {
    if (_pnSubmitting) return;  // guard contra doble click
    _pnSubmitting = true;

    const modal    = document.getElementById('modalProductoNuevo');
    const idGuia   = modal?.dataset.pnGuia || '';
    const cantidad = parseFloat(document.getElementById('pnCantidad')?.value) || 1;
    const fechaVenc = document.getElementById('pnFechaVenc')?.value || '';
    const obs      = (document.getElementById('pnObservaciones')?.value || '').trim();
    const marca    = (document.getElementById('pnMarca')?.value || '').trim();

    // [v2.13.48] VALIDACIÓN OBLIGATORIA descripción — antes el operador
    // confirmaba con campo vacío y la fila quedaba sin descripción en la hoja.
    if (!obs) {
      const obsEl = document.getElementById('pnObservaciones');
      if (obsEl) {
        obsEl.style.borderColor = '#f87171';
        obsEl.style.boxShadow = '0 0 0 3px rgba(248,113,113,.25)';
        obsEl.focus();
        setTimeout(() => {
          obsEl.style.borderColor = '#334155';
          obsEl.style.boxShadow = '';
        }, 2500);
      }
      try { toast('⚠ Falta la descripción del producto', 'warn', 3000); } catch(_) {}
      try { vibrate?.([60, 80, 60]); } catch(_) {}
      // [BUG B] CRÍTICO: liberar el guard antes de salir. Sin esto, _pnSubmitting
      // quedaba en true tras una validación fallida y TODO click posterior a
      // "Registrar" era mudo (entraba al guard de arriba y retornaba en silencio),
      // aunque el operador ya hubiera completado la descripción. Síntoma exacto:
      // "doy click en Registrar y no pasa nada".
      _pnSubmitting = false;
      return;  // NO enviar al backend
    }

    const params = {
      codigoBarra:     _pnCodigo || '',
      idGuia,
      cantidad,
      fechaVencimiento: fechaVenc,
      descripcion:     obs,
      marca:           marca,   // [v2.13.48] el backend lo esperaba pero el frontend nunca lo mandaba
      usuario:         window.WH_CONFIG?.usuario || '',
      idSesion:        window.WH_CONFIG?.idSesion || ''
    };
    if (_pnFotoBase64) {
      params.fotoBase64 = _pnFotoBase64;
      params.mimeType   = _pnFotoMime;
    }

    // Limpiar estado del modal ANTES de la llamada para evitar reuso del codigoBarra
    const cbAEnviar = _pnCodigo;
    _pnCodigo = '';
    _pnFotoBase64 = null;
    _pnFotoMime   = '';

    cerrarModalPN();
    _ocultarPNOffer();

    // ════════════════════════════════════════════════════════════════════
    // [v2.13.213] PERCEPCIÓN INSTANTÁNEA — flujo OPTIMISTA del producto nuevo.
    // Antes: se hacía `await API.registrarPN(...)` ANTES de mostrar el ítem en
    // el detalle. Como registrarProductoNuevo NO está cableado a Supabase
    // (cae a GAS: UPSERT en hoja PRODUCTO_NUEVO + foto a Drive + sync 1:1 a
    // GUIA_DETALLE → ~2-5s), la card recién aparecía al volver el GAS y quedaba
    // pegada en "guardando…" (nacía con _local:true y nadie lo bajaba).
    //   ⚠ CROSS-DOMAIN (verificado): MOS lee los PN desde la HOJA WH
    //   (getProductosNuevosWarehouse → _abrirWhSheet('PRODUCTO_NUEVO')), NO desde
    //   wh.producto_nuevo. Por eso el ruteo del PN DEBE seguir yendo a GAS para
    //   que MOS los vea y apruebe. Solo optimizamos la PERCEPCIÓN: pintamos la
    //   card YA, guardamos el PN en background, y bajamos/elevamos el indicador
    //   según resulte. NO se cambia el ruteo ni el guardado real.
    // ════════════════════════════════════════════════════════════════════
    const codDisplay = cbAEnviar || '';
    let optIdx = -1;
    if (idGuia && _guiaActual?.idGuia === idGuia) {
      if (!_guiaActual.detalle) _guiaActual.detalle = [];
      // Idempotencia visual: reemplazar la línea PN existente del mismo código
      // (doble registro del mismo producto) en vez de duplicar la card.
      const existIdx = _guiaActual.detalle.findIndex(d =>
        String(d.codigoProducto || '').toUpperCase() === String(codDisplay || '').toUpperCase()
        && codDisplay
        && (d._esPN || String(d.observacion || '').toUpperCase().indexOf('PN_') === 0));
      const itemOpt = {
        idDetalle: existIdx >= 0
          ? _guiaActual.detalle[existIdx].idDetalle
          : ('DL_PN_' + Date.now()),
        idGuia,
        codigoProducto: codDisplay,
        descripcionProducto: (obs || codDisplay || '(nuevo)') + ' ⬡N',
        cantidadEsperada: 0,
        cantidadRecibida: cantidad,
        precioUnitario: 0,
        fechaVencimiento: fechaVenc,
        observacion: 'PN_PENDIENTE',
        idProductoNuevo: '',
        // [v2.13.213] _local:true → (1) muestra el indicador "guardando…" y
        // (2) lo PROTEGE del preserve-logic de getGuia (que re-inyecta solo los
        // _local mientras el GAS aún no confirmó la línea en GUIA_DETALLE). En
        // la reconciliación bajamos _local Y _saving (igual que el escaneo). El
        // bug viejo era que el _local NUNCA se bajaba; acá sí.
        _esPN: true, _local: true, _saving: true, _saveFailed: false
      };
      if (existIdx >= 0) { _guiaActual.detalle[existIdx] = itemOpt; optIdx = existIdx; }
      else { _guiaActual.detalle.push(itemOpt); optIdx = _guiaActual.detalle.length - 1; }
      _mostrarDetalleSheet(_guiaActual, false);
    }

    // Feedback inmediato (la card YA está en pantalla).
    toast('✓ ' + (obs || codDisplay || 'Producto nuevo') + ' · guardando…', 'ok', 2200);
    vibrate(15);

    // El guard anti-doble-submit ya no necesita cubrir la llamada de red: el
    // modal está cerrado y el ítem ya se pintó. Liberar de inmediato.
    _pnSubmitting = false;

    // Mantener foco en scanner si estaba abierto
    if (document.getElementById('sheetScanInput')?.classList.contains('open')) {
      setTimeout(() => _enfocarHid(), 250);
    }

    const _idGuiaPN = idGuia;
    // [v2.13.215][FIX 40x #4] Capturar el gate de escritura directa UNA vez al
    // inicio del flujo (junto con el resto de los datos del modal). El gate se
    // usaba DENTRO del .then de registrarPN (segundos después); si el flag global
    // cambiaba en esa ventana, registrar+detalle quedaban incoherentes. Capturado
    // acá, registrar y línea de detalle son coherentes con el mismo estado.
    const _directaPN = !!API._escrituraDirectaActiva?.();
    // Reconcilia la card optimista (por idProductoNuevo de respuesta, o por el
    // idDetalle local) con el resultado del backend. Guard de guía como en
    // _agregarProductoDirecto: si el operador cambió de guía, no tocar UI.
    const _localDetId = (optIdx >= 0) ? _guiaActual.detalle[optIdx].idDetalle : null;
    const _reubicar = () => {
      if (!_guiaActual || _guiaActual.idGuia !== _idGuiaPN) return -1;
      if (!Array.isArray(_guiaActual.detalle)) return -1;
      return _guiaActual.detalle.findIndex(d => d.idDetalle === _localDetId);
    };

    // ── Guardado REAL en background (GAS, ruteo intacto). NO bloquea la UI. ──
    API.registrarPN(params).then(res => {
      const smeQuedo = _guiaActual && _guiaActual.idGuia === _idGuiaPN;
      if (res && (res.ok || res.offline)) {
        const cod  = res.data?.codigoBarra || codDisplay || '';
        const idPN = res.data?.idProductoNuevo || ('PN_L_' + Date.now());
        const esIdempotente = !!res.data?.idempotente;

        // Persistir en cache local — solo si NO es idempotente (ya estaba)
        if (!esIdempotente) {
          const pnCache = OfflineManager.getPNCache();
          pnCache.unshift({ idProductoNuevo: idPN, idGuia: _idGuiaPN, codigoBarra: cod,
            cantidad, fechaVencimiento: fechaVenc, descripcion: obs, estado: 'PENDIENTE',
            fechaRegistro: new Date().toISOString() });
          OfflineManager.setPNCache(pnCache);
        }

        // Reconciliar la card optimista: sellar código/idPN reales devueltos por
        // el backend (el GAS pudo generar un código NLEV). Solo bajamos el
        // indicador "guardando…" si la respuesta es CONFIRMADA por el server
        // (res.ok && !res.offline); en puro offline la card sigue _local (la cola
        // la reconciliará al sincronizar) — mismo criterio que _agregarProductoDirecto.
        if (smeQuedo && _localDetId) {
          const i = _reubicar();
          if (i >= 0) {
            const it = _guiaActual.detalle[i];
            it.codigoProducto      = cod || it.codigoProducto;
            it.descripcionProducto = (obs || cod || '(nuevo)') + ' ⬡N';
            it.idProductoNuevo     = idPN;
            if (res.ok && !res.offline) {
              it._saveFailed = false; it._local = false;
              // [v2.13.215][FIX 40x #1/#3] COHERENCIA: con escritura directa ON, el PN
              // ya está en wh.producto_nuevo pero su LÍNEA en wh.guia_detalle aún se va
              // a guardar en el .then de abajo. NO bajar _saving todavía: si lo
              // bajáramos, la card parecería "guardada" y el guard de cierre (FIX #1) no
              // bloquearía mientras la línea está en vuelo → cierre sin esa línea
              // (descuadre). Se baja _saving cuando el agregarDetalle confirme (o se
              // marca _saveFailed si falla). Con escritura directa OFF, el flujo viejo
              // (sync GAS→Sheet) es la fuente → acá sí se sella.
              if (!_directaPN) {
                it._saving = false;
                if (SoundFX.savedTick) SoundFX.savedTick();
              }
              // Persistir en el cache del detalle para que un getGuia/silentRefresh
              // posterior no haga parpadear la card (mismo patrón que el escaneo).
              try { OfflineManager.addDetalleCache?.(it); } catch(_){}
            }
            _mostrarDetalleSheet(_guiaActual, false);
          }
        }

        // ════════════════════════════════════════════════════════════════
        // [v2.13.214] FIX cutover Supabase — la LÍNEA DE DETALLE del PN.
        // Causa raíz: registrarProductoNuevo (GAS) hace una "sync 1:1" del PN a
        // la hoja GUIA_DETALLE (modelo viejo). Pero las guías DIRECTAS 'G_L...'
        // viven en wh.guia_detalle (Supabase), NO en el Sheet → la línea caía en
        // un Sheet huérfano y NUNCA llegaba a Supabase. El front lee el detalle
        // de Supabase → detalle vacío (el PN existía en wh.producto_nuevo, pero
        // sin su línea en wh.guia_detalle).
        //   FIX: con escritura directa ON, tras registrar el PN en GAS (que sigue
        //   intacto, para que MOS lo apruebe leyendo la hoja), agregamos la línea
        //   directo a wh.guia_detalle vía API.agregarDetalle (action
        //   'agregarDetalleGuia', cableada al RPC agregar_detalle_guia). Esto:
        //     (a) hace aparecer la línea en la guía directa,
        //     (b) genera el lote con la fecha de vencimiento al CERRAR la guía.
        //   Usamos el código REAL devuelto por GAS (`cod` = res.data.codigoBarra:
        //   el que tipeó el operador, o un NLEV auto-generado por GAS).
        //   La guía recién creada está ABIERTA → el RPC solo inserta la línea, NO
        //   mueve stock (el stock se aplica al cerrar). Idempotente por local_id
        //   en el RPC → un reintento/doble-tap no duplica la línea. Va en
        //   background como el resto (no bloquea la UI optimista).
        //   ⚠ Solo con escritura directa ON. Con OFF, el flujo viejo (GAS sync
        //   Sheet) sigue, y agregar acá duplicaría la línea cuando el Sheet sea
        //   la fuente → por eso el gate es obligatorio.
        if (res.ok && !res.offline && _directaPN && cod && _idGuiaPN) {
          API.agregarDetalle({
            idGuia:           _idGuiaPN,
            codigoProducto:   cod,
            cantidadEsperada: 0,
            cantidadRecibida: cantidad,
            precioUnitario:   0,
            fechaVencimiento: fechaVenc || '',
            observacion:      'PN_PENDIENTE',
            // [v2.13.215][FIX 40x #2 BLOQUEANTE] localId DETERMINÍSTICO derivado del
            // id local del PN (DL_PN_...). Estable por línea → doble-tap/reintento del
            // mismo PN colapsa a la misma fila vía el dedup del RPC (por local_id).
            localId:          'PNDET_' + (_localDetId || cod)
          }).then(rDet => {
            const sigueAqui = _guiaActual && _guiaActual.idGuia === _idGuiaPN;
            if (sigueAqui && rDet && rDet.ok && rDet.data) {
              // Sellar la card optimista con el idDetalle REAL de Supabase para que
              // edición/anulación posteriores apunten a la fila real (no al DL_PN_ local)
              // y un getGuia no la duplique.
              const i = _reubicar();
              if (i >= 0) {
                const it = _guiaActual.detalle[i];
                it.idDetalle = rDet.data.idDetalle || it.idDetalle;
                if (rDet.data.idLote) it.idLote = rDet.data.idLote;
                it._local = false; it._saving = false; it._saveFailed = false;
                try { OfflineManager.addDetalleCache?.(it); } catch(_) {}
                try { if (SoundFX.savedTick) SoundFX.savedTick(); } catch(_) {}
                _mostrarDetalleSheet(_guiaActual, false);
              }
            } else if (sigueAqui && rDet && rDet.offline) {
              // Línea encolada offline (idempotente por localId). Marcar _local para
              // que el guard de cierre (FIX #1) la trate como en vuelo hasta que la
              // cola sincronice (no se cierra una guía con su línea solo encolada).
              const i = _reubicar();
              if (i >= 0) {
                _guiaActual.detalle[i]._saving = false;
                _guiaActual.detalle[i]._local = true;
                _guiaActual.detalle[i]._saveFailed = false;
                _mostrarDetalleSheet(_guiaActual, false);
              }
            } else if (sigueAqui && rDet && !rDet.ok && !rDet.offline) {
              // [v2.13.215][FIX 40x #3] El PN se registró en GAS pero su LÍNEA de
              // detalle en wh.guia_detalle NO se guardó (rechazo del backend). Marcar
              // la card en error: sin esto el operador cree que guardó y al cerrar
              // falta el stock de esa línea. Encadena con el FIX #1 (el guard de
              // cierre bloquea por _saveFailed hasta que se reintente).
              const i = _reubicar();
              if (i >= 0) {
                _guiaActual.detalle[i]._saving = false;
                _guiaActual.detalle[i]._saveFailed = true;
                _mostrarDetalleSheet(_guiaActual, false);
              }
              toast('⚠ La línea del producto nuevo no se guardó. Reintentá antes de cerrar.', 'warn', 4000);
            }
          }).catch(() => {
            // [v2.13.215][FIX 40x #3] Excepción/red al guardar la LÍNEA del PN.
            // Marcar _saveFailed para que el guard de cierre (FIX #1) lo bloquee y el
            // operador reintente; antes este catch era mudo ("la cola reintenta") y la
            // guía se cerraba sin la línea → descuadre de stock.
            const sigueAqui = _guiaActual && _guiaActual.idGuia === _idGuiaPN;
            if (sigueAqui) {
              const i = _reubicar();
              if (i >= 0) {
                _guiaActual.detalle[i]._saving = false;
                _guiaActual.detalle[i]._saveFailed = true;
                _mostrarDetalleSheet(_guiaActual, false);
              }
            }
            toast('⚠ Sin conexión — la línea del producto nuevo no se guardó. Reintentá antes de cerrar.', 'warn', 4000);
          });
        }

        // Refrescar lista de guías para actualizar badge N
        if (typeof render === 'function' && typeof _filtrar === 'function') {
          render(_filtrar(todas, filtroActual));
        }
      } else if (res && !res.offline) {
        // GAS rechazó (no es timeout/offline) → marcar la card en error. NO la
        // borramos: el indicador "⚠ no guardado" deja claro que falta reintentar
        // (coherente con el flujo de escaneo) y no se pierde lo que tipeó.
        if (smeQuedo && _localDetId) {
          const i = _reubicar();
          if (i >= 0) { _guiaActual.detalle[i]._saving = false; _guiaActual.detalle[i]._saveFailed = true; _mostrarDetalleSheet(_guiaActual, false); }
        }
        toast('⚠ No se guardó el producto nuevo: ' + (res.error || 'reintenta'), 'warn', 4000);
        try { vibrate?.([60, 80, 60]); } catch(_) {}
      }
    }).catch(() => {
      const smeQuedo = _guiaActual && _guiaActual.idGuia === _idGuiaPN;
      if (smeQuedo && _localDetId) {
        const i = _reubicar();
        if (i >= 0) { _guiaActual.detalle[i]._saving = false; _guiaActual.detalle[i]._saveFailed = true; _mostrarDetalleSheet(_guiaActual, false); }
      }
      toast('⚠ Sin conexión — el producto nuevo no se guardó', 'warn', 3500);
    });
  }

  function abrirPNSinCodigo() {
    abrirModalPN('', _guiaActual?.idGuia || null);
  }

  let _cerrandoGuia = false;
  async function confirmarCerrarGuia() {
    if (_cerrandoGuia) return;   // [v2.13.190] guard anti-doble-submit (evita doble POST de cierre)
    if (!_guiaActual) return;
    const det = (_guiaActual.detalle || []).filter(d => d.observacion !== 'ANULADO');
    if (!det.length) { toast('Agrega al menos un ítem antes de cerrar', 'warn'); return; }

    // [v2.13.215][FIX 40x #1 BLOQUEANTE] No cerrar con líneas EN VUELO.
    // El detalle (escaneo y PN) se guarda en BACKGROUND (.then sin await). Si el
    // operador cierra antes de que agregarDetalle commitee en wh.guia_detalle, el
    // orquestador cerrarGuia arma el detalle desde lo COMMITEADO → aplica stock
    // SIN esa línea (descuadre + lote huérfano) y la línea commitea después contra
    // una guía YA CERRADA (stock fuera del cierre). Bloquear hasta que termine.
    //   _saving      → POST en vuelo (aún no respondió el backend).
    //   _local       → optimista no sellado por el server (offline o sin confirmar).
    //   _saveFailed  → el backend rechazó/falló la línea (hay que reintentarla).
    const enVuelo = (_guiaActual.detalle || [])
      .filter(d => d.observacion !== 'ANULADO' && (d._saving || d._local || d._saveFailed));
    if (enVuelo.length) {
      const fallidas = enVuelo.filter(d => d._saveFailed).length;
      if (fallidas) {
        toast('⚠ ' + fallidas + ' ítem(s) no se guardaron. Reintentá esa(s) línea(s) antes de cerrar.', 'warn', 4000);
      } else {
        toast('Esperá: ' + enVuelo.length + ' ítem(s) aún guardándose. Reintentá en unos segundos.', 'warn', 3500);
      }
      try { vibrate?.([60, 80, 60]); } catch(_) {}
      return;
    }

    // Advertir si hay productos nuevos pendientes
    const pnPend = (OfflineManager.getPNCache() || []).filter(p => p.idGuia === _guiaActual.idGuia && p.estado === 'PENDIENTE').length;
    if (pnPend) toast(`⚠ ${pnPend} producto(s) nuevo(s) sin aprobar en esta guía`, 'warn', 4000);

    _cerrandoGuia = true;
    try {
    // Optimista: actualizar estado en UI inmediatamente
    _guiaActual.estado = 'CERRADA';
    // [v2.13.186] Parchear cache también (simétrico con reabrir) — sin esto un
    // silentRefresh en la ventana del POST re-pintaba la guía ABIERTA.
    OfflineManager.patchGuiaCache?.(_guiaActual.idGuia, { estado: 'CERRADA' });
    _mostrarDetalleSheet(_guiaActual, false);
    const idx = todas.findIndex(g => g.idGuia === _guiaActual.idGuia);
    if (idx >= 0) { todas[idx].estado = 'CERRADA'; render(_filtrarYBuscar()); }

    const res = await API.cerrarGuia(_guiaActual.idGuia, window.WH_CONFIG.usuario);
    if (res.ok || res.offline) {
      const monto = res.data?.montoTotal;
      toast(`Guía cerrada${monto ? ` · S/. ${fmt(monto, 2)}` : ''}`, 'ok', 3000);
      if (res.ok && !res.offline) {
        _guiaActual.montoTotal = monto || 0;
        if (idx >= 0) todas[idx].montoTotal = monto || 0;
        _mostrarDetalleSheet(_guiaActual, false);
        render(_filtrarYBuscar());
      }
      // F9: Ventana de gracia 5min + TTS post-cierre (cantidad + nombre producto)
      if (typeof App !== 'undefined' && App._registrarCierre) App._registrarCierre(_guiaActual.idGuia);
      try {
        if (window.Voice && Voice.supported().tts) {
          const items = (_guiaActual.detalle || [])
            .filter(d => d.observacion !== 'ANULADO' && (parseFloat(d.cantidadRecibida) || 0) > 0)
            .map(d => ({
              cantidad: parseFloat(d.cantidadRecibida) || 0,
              nombre:   d.descripcionProducto || d.codigoProducto
            }));
          if (items.length) setTimeout(() => Voice.leerItems(items), 800);
        }
      } catch(_){}
    } else {
      // Revertir si GAS rechazó
      _guiaActual.estado = 'ABIERTA';
      OfflineManager.patchGuiaCache?.(_guiaActual.idGuia, { estado: 'ABIERTA' });   // [v2.13.186] revert simétrico en cache
      if (idx >= 0) todas[idx].estado = 'ABIERTA';
      _mostrarDetalleSheet(_guiaActual, false);
      render(_filtrarYBuscar());
      toast('Error: ' + res.error, 'danger');
    }
    } finally { _cerrandoGuia = false; }
  }

  let _creandoGuia = false;
  async function crearGuia() {
    if (_creandoGuia) return;   // [v2.13.185] guard de doble-click (evita 2 cards/2 guías)
    const tipo        = document.getElementById('guiaTipo').value;
    const idProveedor = document.getElementById('guiaProveedor').value;
    // [v2.13.185] Validar ANTES de optimista/cerrar sheet (si falla, el form sigue abierto)
    if (!tipo) { toast('Selecciona el tipo de guía', 'warn'); return; }
    if (tipo === 'INGRESO_PROVEEDOR' && !idProveedor) { toast('Selecciona un proveedor', 'warn'); return; }

    _creandoGuia = true;
    try {
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
      if (res.ok && !res.offline && res.data?.idGuia) {
        // [v2.13.211] Escritura DIRECTA (Supabase): la guía REAL ya existe con su
        // fecha y estado ABIERTA. La card optimista quedaba "sin fecha" (no llevaba
        // el campo fecha → caía al grupo 'Sin fecha') y NO era clickeable (el <div>
        // optimista no tiene onclick verDetalle). Solución: descartar la optimista,
        // refrescar el listado desde el backend (trae la guía real con fecha y card
        // navegable) y abrir su detalle para empezar a agregar productos.
        removeOptimisticGuia(tempId);
        toast(`Guía ${res.data.idGuia} creada`, 'ok');
        // Forzar fetch inmediato (ignora throttle) → cache con la guía real.
        OfflineManager.precargarOperacional(true).then(() => {
          const fresh = OfflineManager.getGuiasCache();
          if (fresh.length) { todas = fresh; render(_filtrarYBuscar()); }
        }).catch(() => {});
        // Abrir el detalle de la guía recién creada (misma ruta que las guías
        // viejas: si aún no está en cache, _abrirDetalleConGAS la pide al backend).
        verDetalle(res.data.idGuia);
      } else if (res.ok) {
        // Camino OFFLINE (cola) o sin escritura directa: la respuesta puede tardar,
        // mantener el patrón optimista — la card temporal cubre la espera.
        finalizeOptimisticGuia(tempId, res.data?.idGuia, tipo, provNombre);
        toast('Guía guardada (offline) — sincronizando…', 'ok');
      } else {
        removeOptimisticGuia(tempId);
        toast('Error: ' + (res.error || 'no se pudo crear'), 'danger');
      }
    } finally {
      _creandoGuia = false;
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
    // Buscar fotos del preingreso en caché
    const piCache = OfflineManager.getPreingresosCache();
    const pi = piCache.find(x => x.idPreingreso === _guiaActual.idPreingreso);
    const fotos = pi?.fotos ? String(pi.fotos).split(',').map(s => s.trim()).filter(Boolean) : [];

    if (!fotos.length) { toast('El preingreso no tiene fotos', 'warn'); return; }

    FotoPicker.abrir(_guiaActual.idPreingreso, fotos, async (fileId) => {
      if (!fileId) return;  // canceló o eligió "sin foto"
      const fotoEl = document.getElementById('guiaDetFotoSection');
      if (fotoEl) fotoEl.innerHTML = '<div class="flex justify-center py-3"><div class="spinner"></div></div>';
      const res = await API.copiarFotoDePreingreso({
        idGuia:       _guiaActual.idGuia,
        idPreingreso: _guiaActual.idPreingreso,
        fileId
      }).catch(() => ({ ok: false, error: 'Sin conexión' }));
      if (res.ok && res.data?.url) {
        _guiaActual.foto = res.data.url;
        _mostrarDetalleSheet(_guiaActual, false);
        toast('Foto copiada', 'ok', 1500);
      } else {
        toast('Error: ' + (res.error || 'no se pudo copiar'), 'danger');
        _mostrarDetalleSheet(_guiaActual, false);
      }
    });
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
    // [v2.13.185] No silenciar el clear de la columna: si falla, la foto sigue en
    // el servidor y reaparecerá al refrescar → avisar para reintentar. (No revierto
    // local para no mostrar una imagen de Drive que quizá ya se borró.)
    API.actualizarGuia({ idGuia: _guiaActual.idGuia, foto: '' })
      .then(r => { if (r && r.ok === false) { try { toast('⚠ La foto no se quitó en el servidor — reintenta', 'warn', 4000); } catch(_){} } })
      .catch(() => { try { toast('⚠ Sin conexión — la foto no se quitó del servidor', 'warn', 3500); } catch(_){} });
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
      // [v2.13.190] Capturar también el RECHAZO del backend (ok:false), no solo
      // la excepción de red — antes mostraba "actualizada" aunque el server fallara.
      .then(r => { if (r && r.ok === false) { try { toast('⚠ No se guardaron los cambios: ' + (r.error || ''), 'warn', 4000); } catch(_){} } })
      .catch(() => { try { toast('⚠ Sin conexión — cambios no guardados', 'warn', 3500); } catch(_){} });

    toast('Guía actualizada', 'ok', 1500);
  }

  // Auto-guardar comentario al cerrar el sheet de detalle
  function cerrarGuiaDetalle() {
    if (_guiaActual && _guiaActual.estado === 'ABIERTA') {
      const textoExtra = document.getElementById('guiaComentarioEdit')?.value || '';
      const nuevoComentario = _buildComentario(_tagsGuia, textoExtra);
      if (nuevoComentario !== (_guiaActual.comentario || '')) {
        _guiaActual.comentario = nuevoComentario;
        const idG = _guiaActual.idGuia;
        // [v2.13.183] No silenciar: si falla, avisar para que reintente.
        API.actualizarGuia({ idGuia: idG, comentario: nuevoComentario })
          .then(r => { if (r && r.ok === false) { try { toast('⚠ No se guardó el comentario: ' + (r.error || ''), 'warn', 4000); } catch(_){} } })
          .catch(() => { try { toast('⚠ Sin conexión — comentario no guardado', 'warn', 3500); } catch(_){} });
        // Refrescar card en lista
        const idx = todas.findIndex(g => g.idGuia === idG);
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
        PreingresosView.cargar().then(() => PreingresosView.abrirDetalle(idPreingreso)).catch(() => { try { toast('No se pudo abrir el preingreso', 'warn', 3000); } catch(_){} });
      }
    }, 380);
  }

  // Abrir foto guía en carrusel (usa _guiaActual)
  function verFotoGuia() {
    if (!_guiaActual?.foto) { toast('Sin foto', 'info'); return; }
    abrirCarrusel([_normalizeDriveUrl(_guiaActual.foto)], _guiaActual.idGuia);
  }

  // ── F2-F7 — Crear guía desde type picker + entidad ──
  function crearConTipo(subtipo, entidad) {
    nueva();
    const sel = document.getElementById('guiaTipo');
    if (sel) {
      sel.value = subtipo;
      // Disparar el change handler para sincronizar UI (mostrar/ocultar zona/proveedor)
      sel.dispatchEvent(new Event('change'));
    }
    if (entidad) {
      if (entidad.idProveedor) {
        const ps = document.getElementById('guiaProveedor');
        if (ps) ps.value = entidad.idProveedor;
      } else if (entidad.idZona) {
        const zs = document.getElementById('guiaZona');
        if (zs) zs.value = entidad.idZona;
      } else if (entidad.idPersonal) {
        const com = document.getElementById('guiaComentario');
        if (com) com.value = `Jefatura: ${entidad.nombre || ''} ${entidad.apellido || ''}`.trim();
      }
    }
    abrirSheet('sheetGuia');
  }

  return {
    cargar, filtrar, toggleFiltro, _searchFocus, silentRefresh, verDetalle,
    marcarGuiaAbierta,   // [v2.13.186] reabrir desde App → refleja ABIERTA en módulo
    crearConTipo,
    filtrarTablet, _searchFocusGuiaTablet,
    buscar, buscarClear,
    abrirAgregarItem, abrirCamaraItem, abrirScannerItem,
    cerrarCamara, cerrarCamPicker, seleccionarItemCamara, toggleTorch,
    camQtyPlus, camQtyMinus, camQtyEdit, camUndoLast, camLimpiarTodo, camSetZoom,
    camPrevQtyPlus, camPrevQtyMinus,
    cerrarScannerItem, _enfocarHid,
    seleccionarItemHid,
    _procesarCodigoEscaneado: _buscarCandidatos, _rescanear,
    toggleEstadoGuia, adminPinTecla, adminPinAtras,
    confirmarCerrarGuia, crearGuia, nueva,
    toggleTagNueva,
    onFotoGuiaSeleccionada, copiarFotoDePreingreso, verFotoGuia,
    toggleTagGuia, cerrarGuiaDetalle, irAPreingreso,
    injectOptimisticGuia, finalizeOptimisticGuia, removeOptimisticGuia,
    selectItem, deselectItem,
    inlineQtyDelta, inlineQtyInput, inlineQtyBlurFull, inlineQtyTap,
    inlinePickVenc, inlineVencChanged, inlineDelete,
    toggleFotoPanel, toggleNotasPanel, editarGuia, guardarCambiosGuia,
    eliminarFotoGuia,
    filtrarChip,
    compartirWA,
    imprimirTicket,
    abrirModalPN, cerrarModalPN, _pnFotoSeleccionada, confirmarRegistrarPN,
    _pnCodigoChanged, _pnToggleAutoGen,
    abrirPNSinCodigo, _ocultarPNOffer
  };

  function imprimirTicket(idGuia) {
    const base       = location.href.split('?')[0].replace(/index\.html$/, '').replace(/\/$/, '');
    const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(idGuia)}`;
    vibrate(10);
    // [v2.13.59] fuerzaCopia:true SALTA el dedup de 60s del backend.
    // Este es el botón manual "🖨 imprimir copia" del historial — el operador
    // explícitamente quiere otra copia aunque ya se haya impreso al cerrar.
    // PrintHub: admin/master ve el selector de impresora; usuario normal
    // imprime directo en la del almacén.
    PrintHub.imprimir('imprimirTicketGuia',
                      { idGuia, reporteUrl, fuerzaCopia: true, motivo: 'copia_manual' },
                      'Guía ' + idGuia)
      .then(res => {
        if (res && res.ok === false) toast('Error al imprimir: ' + (res.error || ''), 'warn', 4000);
      })
      .catch(() => toast('Error al imprimir', 'warn', 3000));
  }

  // [v2.13.31] Cola persistente de impresiones pendientes (localStorage).
  // Se drenan al volver online o al volver a abrir la app.
  const _PRINT_QUEUE_KEY = 'wh_print_queue_v1';

  function _getPrintQueue() {
    try { return JSON.parse(localStorage.getItem(_PRINT_QUEUE_KEY) || '[]'); } catch(_) { return []; }
  }
  function _setPrintQueue(q) {
    try { localStorage.setItem(_PRINT_QUEUE_KEY, JSON.stringify(q || [])); } catch(_){}
  }
  function _encolarImpresionPendiente(idGuia, titulo) {
    if (!idGuia) return;
    const q = _getPrintQueue();
    if (q.some(x => x.idGuia === idGuia)) return; // ya en cola
    q.push({ idGuia, titulo: titulo || ('Guía ' + idGuia), ts: Date.now(), intentos: 0 });
    _setPrintQueue(q);
    try { _renderPrintQueueBadge && _renderPrintQueueBadge(); } catch(_){}
  }
  function _removerImpresionPendiente(idGuia) {
    _setPrintQueue(_getPrintQueue().filter(x => x.idGuia !== idGuia));
    try { _renderPrintQueueBadge && _renderPrintQueueBadge(); } catch(_){}
  }

  // [v2.13.59] Política revisada: NUNCA reintentar inmediato.
  // Bug histórico: backoff 1s/3s/8s producía 2-3 copias del mismo ticket porque
  // cuando el HTTP del navegador timeouteaba, PrintNode ya había aceptado el job
  // y la impresora había sacado el papel. El reintento mandaba otra copia.
  //
  // Ahora:
  //   - 1 sola llamada al backend (que tiene dedup interno de 60s)
  //   - Si excepción HTTP/red → asumimos que PrintNode probablemente sí imprimió
  //     pero no lo sabemos seguro → encolamos UNA vez para que el operador
  //     verifique visualmente y use el botón "Reintentar" del toast persistente
  //   - Si r.ok === false con error transitorio (printer offline, apikey, 5xx)
  //     → un solo reintento tras 3s; si vuelve a fallar, encolar
  //   - Si r.ya_impresa === true (backend dedupó) → toast informativo, NO encolar
  async function _imprimirConReintento(idGuia, reporteUrl, titulo) {
    let resp = null;
    let lastErr = null;
    try {
      resp = await API.imprimirTicketGuia({ idGuia, reporteUrl });
      if (resp && resp.ok !== false) {
        _removerImpresionPendiente(idGuia);
        if (resp.ya_impresa) {
          toast('🖨 Ticket ya impreso · usa "Imprimir copia" en historial si necesitas otra', 'info', 4500);
        }
        return resp;
      }
      lastErr = (resp && resp.error) || 'error PrintNode';
    } catch(e) {
      // Excepción HTTP: timeout, red caída, Apps Script colgado.
      // PrintNode SUELE imprimir igual → no reintentar inmediato. Encolar para verificación.
      lastErr = 'timeout: revisa si el ticket salió de la impresora';
      _encolarImpresionPendiente(idGuia, titulo);
      _toastImpresionFallida(idGuia, titulo, lastErr);
      return { ok: false, error: lastErr, posiblementeImpreso: true };
    }

    // Llegamos acá si el backend respondió ok:false explícito.
    // Solo reintentamos UNA vez si el error suena recuperable (no si es "guía no existe").
    const errStr = String(lastErr || '').toLowerCase();
    const esTransitorio = /printer|offline|503|502|504|apikey|red|network|timeout|fetch/i.test(errStr);
    if (esTransitorio) {
      toast('🖨 Error transitorio · reintento único en 3s…', 'info', 3500);
      await new Promise(r => setTimeout(r, 3000));
      try {
        const r2 = await API.imprimirTicketGuia({ idGuia, reporteUrl });
        if (r2 && r2.ok !== false) {
          _removerImpresionPendiente(idGuia);
          toast('🖨 Impresión OK tras 1 reintento', 'ok', 3000);
          return r2;
        }
        lastErr = (r2 && r2.error) || lastErr;
      } catch(e2) {
        lastErr = (e2 && e2.message) || 'sin red';
      }
    }

    // Falló de forma duradera — encolar + toast persistente con CTA
    _encolarImpresionPendiente(idGuia, titulo);
    _toastImpresionFallida(idGuia, titulo, lastErr);
    return { ok: false, error: lastErr };
  }

  // Toast persistente (sin auto-cierre) con botón Reintentar.
  function _toastImpresionFallida(idGuia, titulo, error) {
    try {
      const cont = document.getElementById('toastContainer') || document.body;
      const id = 't_imp_' + idGuia;
      // Si ya hay un toast persistente para esta guía, no duplicar
      if (document.getElementById(id)) return;
      const el = document.createElement('div');
      el.id = id;
      el.className = 'toast toast-warn toast-persistente';
      el.style.cssText = 'background:#7c2d12;color:#fed7aa;border:2px solid #ea580c;padding:12px 14px;border-radius:10px;display:flex;flex-direction:column;gap:8px;max-width:340px;box-shadow:0 8px 24px rgba(0,0,0,.5);margin-bottom:8px;font-size:13px;animation:cjFlyIn .3s';
      el.innerHTML = `
        <div style="display:flex;align-items:start;gap:8px">
          <span style="font-size:18px">🖨</span>
          <div style="flex:1">
            <div style="font-weight:800">No se imprimió el ticket</div>
            <div style="font-size:11px;opacity:.85;margin-top:2px">${titulo || idGuia} · ${String(error||'error').substring(0,60)}</div>
          </div>
          <button onclick="this.closest('.toast-persistente').remove()" style="background:none;border:none;color:#fed7aa;font-size:18px;font-weight:900;cursor:pointer;line-height:1">×</button>
        </div>
        <div style="display:flex;gap:6px">
          <button onclick="window._whReintentarImpresion('${idGuia.replace(/'/g,'')}')"
                  style="flex:1;padding:8px;background:#ea580c;color:#fff;border:none;border-radius:6px;font-weight:800;cursor:pointer;font-size:12px">
            🖨 Reintentar ahora
          </button>
        </div>`;
      cont.appendChild(el);
    } catch(_){
      toast('🖨 Ticket NO impreso: ' + (error||'') + ' · usa botón 🖨 en historial', 'warn', 10000);
    }
  }

  // Expuesto global para el onclick del toast persistente
  window._whReintentarImpresion = async function(idGuia) {
    const base = location.origin + location.pathname.replace(/\/[^/]*$/, '');
    const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(idGuia)}`;
    // Cerrar toast persistente
    const t = document.getElementById('t_imp_' + idGuia);
    if (t) t.remove();
    toast('🖨 Reintentando impresión…', 'info', 3000);
    const r = await _imprimirConReintento(idGuia, reporteUrl, 'Guía ' + idGuia);
    if (r && r.ok !== false) toast('🖨 Impresión OK', 'ok', 3000);
  };

  // Drenado de cola al cargar / volver online.
  // [v2.13.59] Pasar dedupSiempre:true — el backend chequea si la guía YA fue
  // impresa en algún momento (no solo últimos 60s). Si sí, devuelve ya_impresa
  // y limpiamos la entrada de la cola sin disparar otra copia.
  async function _drenarColaImpresion() {
    const q = _getPrintQueue();
    if (!q.length) return;
    for (const item of q) {
      const base = location.origin + location.pathname.replace(/\/[^/]*$/, '');
      const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(item.idGuia)}`;
      try {
        const r = await API.imprimirTicketGuia({
          idGuia:        item.idGuia,
          reporteUrl,
          dedupSiempre:  true   // backend solo imprime si NUNCA se imprimió esta guía
        });
        if (r && r.ok !== false) {
          _removerImpresionPendiente(item.idGuia);
          if (r.ya_impresa) {
            // El backend dedupó: la guía ya estaba impresa, no volvemos a sacar papel
            console.log('[print-queue] dedupado · ya impresa:', item.idGuia);
          } else {
            toast(`🖨 ${item.titulo || item.idGuia} · impresión pendiente OK`, 'ok', 4000);
          }
        }
      } catch(_){}
    }
  }
  // Reintentar cola al volver online (red recuperada)
  window.addEventListener('online', () => { setTimeout(_drenarColaImpresion, 2000); });
  // Reintentar al cargar la app si hay pendientes (3s después para no pisar carga inicial)
  setTimeout(_drenarColaImpresion, 3000);

  function _renderPrintQueueBadge() {
    // Placeholder por si querés agregar un badge en el header con la cola
    // pendiente. Por ahora solo log. El user ve el toast persistente.
    const q = _getPrintQueue();
    if (q.length) console.log('[print-queue] pendientes:', q.length);
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
// Historial de envasados agrupado por día (últimos 7 días, desc), solo del usuario actual
function _renderEnvasadosPorDia(list, container, opts) {
  // [v2.10.1] Reglas de visibilidad:
  //   - Admin / Master: ven TODOS los envasados
  //   - Operador normal: ve solo SUS envasados
  //   - Comparación tolerante (trim + toLowerCase).
  // [v2.13.116] opts permite vista "historial 14d" vs vista "lista de hoy":
  //   - opts.modo: 'hoy' (default) | 'historial14d'
  //   - opts.query: filtro de búsqueda por producto+usuario
  opts = opts || {};
  const modo  = opts.modo || 'hoy';
  const query = String(opts.query || '').trim().toLowerCase();
  const rol           = String(window.WH_CONFIG?.rol || '').toUpperCase();
  const esAdminMaster = (rol === 'MASTER' || rol === 'ADMIN');
  const usuarioActual = String(window.WH_CONFIG?.usuario || '').trim().toLowerCase();
  const hoy           = new Date().toISOString().split('T')[0];
  const desde14       = (typeof _fechaDesde14Dias === 'function') ? _fechaDesde14Dias() : hoy;

  // Mapa codigoBarra → descripcion del maestro (para legibilidad).
  // Fallback en cascada: App.getProductosMaestro (memoria) → OfflineManager
  // (cache local) → idProducto. Asegura que aunque la app esté arrancando
  // y App no haya cargado todavía, igual mostremos algo descriptivo.
  const prodMap = {};
  function _addProd(p) {
    const desc = p.descripcion || p.nombre || p.idProducto || '';
    if (!desc) return;
    if (p.codigoBarra) prodMap[String(p.codigoBarra)] = desc;
    if (p.skuBase)     prodMap[String(p.skuBase)]     = desc;
    if (p.idProducto)  prodMap[String(p.idProducto)]  = prodMap[String(p.idProducto)] || desc;
  }
  try {
    const a = (window.App && App.getProductosMaestro && App.getProductosMaestro()) || [];
    a.forEach(_addProd);
    if (window.OfflineManager?.getProductosCache) {
      OfflineManager.getProductosCache().forEach(_addProd);
    }
  } catch(_){}

  // [v2.13.116] Filtrar por rango según modo
  let visible;
  if (modo === 'historial14d') {
    visible = list.filter(e => {
      const f = String(e.fecha || '').substring(0, 10);
      return f >= desde14 && f <= hoy;
    });
  } else {
    visible = list.filter(e => String(e.fecha || '').substring(0, 10) === hoy);
  }
  if (!esAdminMaster && usuarioActual) {
    visible = visible.filter(e =>
      String(e.usuario || '').trim().toLowerCase() === usuarioActual
    );
  }
  // [v2.13.116] Buscador: producto (derivado o base) + usuario (solo admin/master)
  if (query) {
    visible = visible.filter(e => {
      const cbDer  = String(e.codigoProductoEnvasado || '');
      const cbBase = String(e.codigoProductoBase     || '');
      const descD  = String(e.descripcionProductoEnvasado || prodMap[cbDer]  || cbDer).toLowerCase();
      const descB  = String(e.descripcionProductoBase     || prodMap[cbBase] || cbBase).toLowerCase();
      const usr    = String(e.usuario || '').toLowerCase();
      return descD.indexOf(query) >= 0
          || descB.indexOf(query) >= 0
          || cbDer.toLowerCase().indexOf(query) >= 0
          || cbBase.toLowerCase().indexOf(query) >= 0
          || usr.indexOf(query) >= 0;
    });
  }

  if (!visible.length) {
    const msg = modo === 'historial14d'
      ? (query ? 'Sin resultados para "' + query + '"' :
                 (esAdminMaster ? 'Sin envasados en las últimas 2 semanas' : 'No registraste envasados en las últimas 2 semanas'))
      : (esAdminMaster ? 'Sin envasados registrados hoy' : 'No registraste envasados hoy');
    container.innerHTML = `<p class="text-slate-500 text-center py-8 text-sm">${msg}</p>`;
    return;
  }
  const grupos = {};
  visible.forEach(e => {
    const key = String(e.fecha || '').substring(0, 10);
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(e);
  });
  const dias = Object.keys(grupos).sort((a, b) => b.localeCompare(a));
  container.innerHTML = dias.map(dia => {
    const items = grupos[dia].map(e => {
      const efPct   = parseFloat(e.eficienciaPct) || 0;
      const efColor = efPct >= 95 ? 'text-emerald-400' : efPct >= 85 ? 'text-amber-400' : 'text-red-400';
      const est     = String(e.estado || '').toUpperCase();
      const anulado = est === 'ANULADO' || est === 'ANULADO_DUPLICADO';
      const optimista = String(e.idEnvasado || '').indexOf('ENV_OPT_') === 0;
      const cardCls = anulado ? ' opacity-50' : '';
      const tagAnul = anulado
        ? `<span class="text-[10px] font-bold px-1.5 py-0.5 rounded ml-1" style="background:rgba(239,68,68,.15);color:#fca5a5">✕ ${est === 'ANULADO_DUPLICADO' ? 'DUPLICADO' : 'ANULADO'}</span>`
        : '';
      const tagOpt = optimista
        ? `<span class="text-[10px] font-bold px-1.5 py-0.5 rounded ml-1" style="background:rgba(99,102,241,.15);color:#a5b4fc" title="Pendiente confirmación del servidor">⏳ sync</span>`
        : '';
      // Botones editar/anular: solo si NO está anulado y NO es optimista (esperá que confirme).
      // Ambos requieren clave admin (8 dígitos de ESTACIONES.ALMACEN.adminPin en MOS).
      // [v2.13.150] Botón "🖨 Adhesivo" — reimprime adhesivos del envasado.
      // Aparece SIEMPRE que haya codigoProductoEnvasado (incluso si anulado, ya
      // que el operador puede necesitar reimprimir un anulado para reusar el
      // batch físico). No requiere clave admin — es operación no destructiva.
      const btnAdhesivo = (e.idEnvasado && e.codigoProductoEnvasado)
        ? `<button onclick="event.stopPropagation();WhAdhesivoReprint.abrir('${escAttr(e.idEnvasado)}')"
                   class="text-[10px] px-2 py-1 rounded"
                   style="background:rgba(34,197,94,.12);color:#86efac;border:1px solid rgba(34,197,94,.35)"
                   title="Reimprimir adhesivos de este envasado">🖨 Adhesivo</button>`
        : '';
      const acciones = (!anulado && !optimista && e.idEnvasado)
        ? `<button onclick="event.stopPropagation();EnvasadosView.pedirAuthEditar('${escAttr(e.idEnvasado)}')"
                   class="text-[10px] px-2 py-1 rounded ml-auto"
                   style="background:rgba(99,102,241,.12);color:#a5b4fc;border:1px solid rgba(99,102,241,.35)"
                   title="Editar cantidad · requiere clave admin">✏ Editar</button>
           <button onclick="event.stopPropagation();EnvasadosView.pedirAuthAnular('${escAttr(e.idEnvasado)}')"
                   class="text-[10px] px-2 py-1 rounded"
                   style="background:rgba(239,68,68,.12);color:#fca5a5;border:1px solid rgba(239,68,68,.35)"
                   title="Anular este envasado · requiere clave admin">🚫 Anular</button>
           ${btnAdhesivo}`
        : btnAdhesivo;  // [v2.13.150] anulados: solo botón adhesivo
      const cbDer  = String(e.codigoProductoEnvasado || '');
      const cbBase = String(e.codigoProductoBase || '');
      // [v2.10.5] Prioridad: campo del backend (siempre resuelve) → prodMap
      // local (fallback para items optimistic recién registrados) → código.
      const descDer  = e.descripcionProductoEnvasado || prodMap[cbDer]  || cbDer  || '—';
      const descBase = e.descripcionProductoBase     || prodMap[cbBase] || cbBase || '—';
      const mermaReal = parseFloat(e.mermaReal) || 0;

      const hora = _fmtHora(e.fechaImpresion || e.fecha);
      return `<div class="card-sm${cardCls}">
        <div class="flex items-start justify-between gap-2 mb-1">
          <div class="flex-1 min-w-0">
            <p class="text-sm font-bold text-slate-100 leading-tight">
              ${hora ? `<span class="text-[10px] font-mono text-amber-400/80 mr-1">${hora}</span>` : ''}
              ${escHtml(descDer)}
              <span class="text-[10px] font-mono text-slate-500 ml-1 whitespace-nowrap" style="letter-spacing:-.05em">▌${escHtml(cbDer)}</span>
              ${tagAnul}${tagOpt}
            </p>
            <p class="text-[11px] text-slate-400 mt-0.5">
              <span class="text-slate-500">desde</span> ${escHtml(descBase)}
              <span class="font-mono text-slate-600" style="letter-spacing:-.05em">▌${escHtml(cbBase)}</span>
              <span class="text-slate-500"> · ${fmt(e.cantidadBase, 1)} ${escHtml(e.unidadBase || '')}</span>
            </p>
          </div>
          <span class="${efColor} font-bold text-sm shrink-0">${efPct}%</span>
        </div>
        <div class="flex items-center justify-between gap-2 mt-1.5">
          <div class="text-xs text-slate-300 min-w-0 truncate">
            <span class="font-bold text-slate-100">${fmt(e.unidadesProducidas)} uds</span>
            ${mermaReal > 0 ? `<span class="text-amber-400 ml-2">· Merma: ${fmt(mermaReal)}</span>` : ''}
            <span class="text-slate-400 ml-2">· 👤 ${escHtml(e.usuario || '—')}</span>
          </div>
          <div class="flex gap-1 shrink-0">${acciones}</div>
        </div>
      </div>`;
    }).join('');
    // [v2.13.116] Header de día con formato nombrado + estilo prominente para HOY/AYER
    const lbl = (typeof _fmtDiaLabel === 'function') ? _fmtDiaLabel(dia) : { texto: fmtFecha(dia), destacado: false };
    const headerCls = lbl.destacado
      ? 'text-sm font-bold text-amber-400 mt-4 mb-2 px-1 tracking-wide border-b border-amber-500/20 pb-1'
      : 'text-xs font-semibold text-slate-400 mt-3 mb-1 px-1 uppercase tracking-wide';
    return `<p class="${headerCls}">${escHtml(lbl.texto || lbl)}</p>${items}`;
  }).join('');
}

function _fechaDesde7Dias() {
  const d = new Date();
  d.setDate(d.getDate() - 6);
  return d.toISOString().split('T')[0];
}

// [v2.13.116] Historial de Modo Envasador — 14 días.
function _fechaDesde14Dias() {
  const d = new Date();
  d.setDate(d.getDate() - 13);
  return d.toISOString().split('T')[0];
}

// [v2.13.116] Convierte una fecha 'YYYY-MM-DD' a label legible para humanos.
//   0 días → "HOY · Martes 2 jun"
//   1 día  → "AYER · Lunes 1 jun"
//   2-14   → "Sábado 30 may"
// Marca destacada para HOY/AYER que en el render usa estilo más prominente.
function _fmtDiaLabel(yyyymmdd) {
  const partes = String(yyyymmdd).split('-');
  if (partes.length !== 3) return yyyymmdd;
  const d = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
  d.setHours(0, 0, 0, 0);
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const diffDias = Math.round((hoy.getTime() - d.getTime()) / (24 * 60 * 60 * 1000));
  const dias  = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const diaSem = dias[d.getDay()];
  const num    = d.getDate();
  const mes    = meses[d.getMonth()];
  if (diffDias === 0) return { destacado: true,  texto: 'HOY · ' + diaSem + ' ' + num + ' ' + mes };
  if (diffDias === 1) return { destacado: true,  texto: 'AYER · ' + diaSem + ' ' + num + ' ' + mes };
  return                    { destacado: false, texto: diaSem + ' ' + num + ' ' + mes };
}

// Hora 'HH:mm' desde un timestamp string o Date.
// [v2.13.117 AUDIT FIX] Solo mostrar hora si la fecha tiene componente
// de tiempo. Si es solo "2026-06-02", new Date() la parsea como UTC
// midnight → en zona horaria local muestra 19:00 (Perú) o 21:00 (AR)
// — hora basura. Mejor no mostrar nada.
function _fmtHora(fechaTs) {
  if (!fechaTs) return '';
  const s = String(fechaTs);
  // Solo procesar si tiene 'T' (ISO) o ' ' (date+time space-separated) o ':' (HH:MM)
  if (s.indexOf('T') < 0 && s.indexOf(' ') < 0 && s.indexOf(':') < 0) return '';
  const d = fechaTs instanceof Date ? fechaTs : new Date(fechaTs);
  if (isNaN(d.getTime())) return '';
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return hh + ':' + mm;
}

// ENVASADOS VIEW
// ════════════════════════════════════════════════
const EnvasadosView = (() => {
  let derivados = [];
  let productosMaestro = [];

  async function cargar() {
    const fd        = _fechaDesde7Dias();
    const container = document.getElementById('listEnvasados');

    // Optimistic: mostrar caché de inmediato
    const cached = OfflineManager.getEnvasadosCache()
      .filter(e => String(e.fecha || '').substring(0, 10) >= fd);
    if (cached.length) {
      _renderEnvasadosPorDia(cached, container);
    } else {
      loading('listEnvasados', true);
    }

    // Fondo: actualizar desde servidor
    const res = await API.getEnvasados({ fechaDesde: fd }).catch(() => ({ ok: false }));
    if (res.ok) {
      // [v2.10.4] Merge robusto: preservar TODO lo del cache local que no
      // venga en res.data (no solo ENV_OPT_*). Antes el optimistic se reemplazaba
      // por el real con id ENV..., y si cargar() corría antes que getEnvasados
      // reflejara el nuevo registro, el real se borraba y el usuario veía el
      // envasado aparecer y desaparecer hasta navegar y volver.
      const cacheLocal = OfflineManager.getEnvasadosCache();
      const idsBackend = new Set(res.data.map(e => String(e.idEnvasado || '')));
      const pendientes = cacheLocal.filter(e =>
        e.idEnvasado && !idsBackend.has(String(e.idEnvasado))
      );
      const fusionado = pendientes.length
        ? [...pendientes, ...res.data]
        : res.data;
      OfflineManager.guardarEnvasadosCache(fusionado);
      _renderEnvasadosPorDia(fusionado, container);
    }
  }

  let _selDerivadoId = '';   // id del derivado seleccionado actualmente

  function nuevo(preIdBase, preIdDerivado) {
    productosMaestro = App.getProductosMaestro();
    derivados = productosMaestro.filter(p => p.codigoProductoBase && p.codigoProductoBase !== '');

    // Auto-fill fecha vencimiento = hoy + 12 meses
    const fv = new Date();
    fv.setFullYear(fv.getFullYear() + 1);
    document.getElementById('envFechaVenc').value = fv.toISOString().split('T')[0];

    // Reset
    document.getElementById('envUnidades').value = 1;
    document.getElementById('envNombreProducto').textContent = '';
    document.getElementById('envasadoFactorInfo').classList.add('hidden');
    document.getElementById('envHistorialMini').classList.add('hidden');
    document.getElementById('envProductoDerivado').value = '';
    document.getElementById('envBuscarDerivado').value = '';
    // Reset checkbox imprimir — siempre arranca marcado (bug #9)
    document.getElementById('envImprimirEtiq').checked = true;
    // [Fix v2.9.1] Resetear el botón "Registrar". Si un envasado previo está
    // colgado en background (timeout largo) y el operador abre otro, el botón
    // conservaba "Registrando..." + disabled del flujo anterior.
    const _btnReg = document.getElementById('btnRegistrarEnvasado');
    if (_btnReg) {
      _btnReg.disabled = false;
      _btnReg.textContent = 'Registrar envasado';
    }
    _selDerivadoId = '';

    if (preIdDerivado) {
      // Pre-seleccionado desde EnvasadorView card → ir directo al panel 2
      seleccionarDerivado(preIdDerivado, /*lock=*/true);
    } else {
      // Mostrar panel de búsqueda
      document.getElementById('envPanelBuscar').classList.remove('hidden');
      document.getElementById('envPanelSeleccion').classList.add('hidden');
      filtrarDerivados('');
    }

    abrirSheet('sheetEnvasado');
    if (!preIdDerivado) {
      setTimeout(() => document.getElementById('envBuscarDerivado')?.focus(), 250);
    }
  }

  // Filtra la lista de derivados por nombre/codigo/marca y la renderiza
  // como cards seleccionables. Muestra "máx producibles" en cada card
  // para que el operador elija con contexto.
  function filtrarDerivados(query) {
    const cont = document.getElementById('envListaDerivados');
    if (!cont) return;
    const q = String(query || '').trim().toLowerCase();
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => {
      stockMap[String(s.codigoProducto || s.idProducto)] = s;
    });
    let list = derivados;
    if (q) {
      list = list.filter(d => {
        const txt = ((d.descripcion || '') + ' ' + (d.codigoBarra || '') + ' ' + (d.marca || '') + ' ' + (d.idProducto || '')).toLowerCase();
        return txt.indexOf(q) >= 0;
      });
    }
    if (!list.length) {
      cont.innerHTML = `<p class="text-xs text-slate-500 italic text-center py-4">Sin coincidencias${q ? ' para "' + escHtml(q) + '"' : ''}</p>`;
      return;
    }
    cont.innerHTML = list.slice(0, 60).map(d => {
      const prodBase = productosMaestro.find(p =>
        (p.skuBase && p.skuBase === d.codigoProductoBase) ||
        p.idProducto === d.codigoProductoBase
      );
      const cbBase = prodBase ? String(prodBase.codigoBarra) : '';
      const stockBase = parseFloat((stockMap[cbBase] || {}).cantidadDisponible || 0);
      const fb = parseFloat(d.factorConversionBase) || 0;
      const maxP = fb > 0 ? Math.floor(stockBase / fb) : 0;
      const maxBadge = maxP > 0
        ? `<span class="text-[10px] text-emerald-400 font-bold">${maxP} uds máx</span>`
        : `<span class="text-[10px] text-amber-400 font-bold">⚠ sin stock</span>`;
      const idAttr = escAttr(d.idProducto);
      return `<button onclick="EnvasadosView.seleccionarDerivado('${idAttr}')"
        class="w-full text-left px-3 py-2 rounded-lg border transition-all active:scale-[.98]"
        style="background:rgba(15,23,42,.6);border-color:#1e293b;">
        <div class="flex items-center justify-between gap-2">
          <div class="min-w-0 flex-1">
            <p class="text-sm font-semibold text-slate-100 truncate">${escHtml(d.descripcion || d.idProducto)}</p>
            <p class="text-[10px] text-slate-500 font-mono truncate">${escHtml(d.codigoBarra || d.idProducto)}</p>
          </div>
          ${maxBadge}
        </div>
      </button>`;
    }).join('') + (list.length > 60 ? `<p class="text-[10px] text-slate-600 italic text-center py-1">+ ${list.length - 60} más · refiná tu búsqueda</p>` : '');
  }

  // Selecciona un derivado: setea hidden input, llama onDerivadoChange,
  // cambia al panel de selección (oculta el buscador), pinta historial mini.
  function seleccionarDerivado(idDerivado, locked) {
    if (!idDerivado) return;
    const prod = derivados.find(d => d.idProducto === idDerivado);
    if (!prod) return;
    _selDerivadoId = idDerivado;
    document.getElementById('envProductoDerivado').value = idDerivado;
    document.getElementById('envSelNombre').textContent = prod.descripcion || idDerivado;
    document.getElementById('envSelMeta').textContent  = (prod.codigoBarra || idDerivado) + (prod.marca ? ' · ' + prod.marca : '');
    document.getElementById('envPanelBuscar').classList.add('hidden');
    document.getElementById('envPanelSeleccion').classList.remove('hidden');
    onDerivadoChange(idDerivado);
    _pintarHistorialMini(idDerivado);
    try { if (typeof SoundFX !== 'undefined' && SoundFX.beep) SoundFX.beep(); } catch(_){}
  }

  function cambiarDerivado() {
    _selDerivadoId = '';
    document.getElementById('envProductoDerivado').value = '';
    document.getElementById('envPanelSeleccion').classList.add('hidden');
    document.getElementById('envPanelBuscar').classList.remove('hidden');
    document.getElementById('envasadoFactorInfo').classList.add('hidden');
    document.getElementById('envHistorialMini').classList.add('hidden');
    document.getElementById('envNombreProducto').textContent = '';
    document.getElementById('envBuscarDerivado').value = '';
    filtrarDerivados('');
    setTimeout(() => document.getElementById('envBuscarDerivado')?.focus(), 100);
  }

  // Muestra los últimos 3 envasados de ESE derivado en el sheet — ayuda
  // a detectar visualmente si el operador está a punto de duplicar.
  function _pintarHistorialMini(idDerivado) {
    const cont = document.getElementById('envHistorialMini');
    if (!cont) return;
    const prod = derivados.find(d => d.idProducto === idDerivado);
    if (!prod) { cont.classList.add('hidden'); return; }
    const cbDer = String(prod.codigoBarra || '');
    const recientes = OfflineManager.getEnvasadosCache()
      .filter(e => String(e.codigoProductoEnvasado) === cbDer &&
                   String(e.estado || '').toUpperCase() !== 'ANULADO' &&
                   String(e.estado || '').toUpperCase() !== 'ANULADO_DUPLICADO')
      .sort((a, b) => String(b.fecha || '').localeCompare(String(a.fecha || '')))
      .slice(0, 3);
    if (!recientes.length) { cont.classList.add('hidden'); return; }
    cont.classList.remove('hidden');
    cont.innerHTML = `
      <div class="bg-slate-800/60 rounded-xl p-2 border border-slate-700">
        <p class="text-[10px] font-semibold text-slate-400 uppercase tracking-wide mb-1">📋 Últimos de este producto</p>
        <div class="space-y-0.5">
          ${recientes.map(e => {
            const f = String(e.fecha || '').substring(0, 10);
            const hora = (() => { try { return new Date(e.fecha).toLocaleTimeString('es-PE',{hour:'2-digit',minute:'2-digit'}); } catch(_){ return ''; } })();
            return `<div class="flex items-center justify-between text-[11px]">
              <span class="text-slate-400">${escHtml(f)} ${escHtml(hora)} · ${escHtml(e.usuario || '—')}</span>
              <span class="font-bold text-slate-200">${fmt(e.unidadesProducidas)} uds</span>
            </div>`;
          }).join('')}
        </div>
      </div>`;
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

  // [Bug #7 cleanup] calcularProyeccion era código muerto: leía envCantBase
  // que NO existe en el HTML del sheet → siempre daba 0 unidades. Se deja
  // como no-op para no romper si algo aún la llama. Las unidades se setean
  // directo en el input grande del sheet o con los presets/+- buttons.
  function calcularProyeccion() {
    const idDerivado = document.getElementById('envProductoDerivado')?.value || _selDerivadoId;
    const prod = derivados.find(p => p.idProducto === idDerivado);
    if (!prod) return;
    const factor   = parseFloat(prod.factorConversion)  || 1;
    const merma    = parseFloat(prod.mermaEsperadaPct)   || 0;
    // No-op: dejado sin efecto. Los presets / +- arman la cantidad.
    void factor; void merma;
    return;
    /* legacy */
    const cantBase = 0;
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
    const factorBase = parseFloat(prod.factorConversionBase) || 0;
    const cantBase   = producidas * factorBase;

    // ── Alertas NO bloqueantes (decisión del usuario) ──
    // Stock base no se valida como bloqueo: el operador puede envasar
    // aunque el stock base esté en 0 o quede negativo. Solo le avisamos.
    // Igual con la cantidad: puede pasar el máximo calculado, solo aviso.
    const cbBase     = prodBase ? String(prodBase.codigoBarra) : '';
    const stockEntry = OfflineManager.getStockCache().find(s => String(s.codigoProducto) === cbBase);
    const stockBase  = stockEntry ? (parseFloat(stockEntry.cantidadDisponible) || 0) : 0;
    const maxProd    = factorBase > 0 ? Math.floor(stockBase / factorBase) : 0;
    const avisos = [];
    if (stockBase <= 0)        avisos.push('⚠ Sin stock base registrado');
    else if (cantBase > stockBase) avisos.push(`⚠ Faltan ${fmt(cantBase - stockBase, 1)} ${prodBase?.unidad || 'kg'} — stock quedará negativo`);
    if (maxProd > 0 && producidas > maxProd) avisos.push(`⚠ Excedes el máximo (${maxProd} uds calculados)`);
    if (producidas >= 500)     avisos.push(`⚠ Cantidad alta: ${producidas} uds`);

    // Fecha venc suave: solo aviso si vacía o pasada (no bloquea)
    const hoyStr = new Date().toISOString().split('T')[0];
    if (!fechaVenc)              avisos.push('⚠ Sin fecha de vencimiento');
    else if (fechaVenc < hoyStr) avisos.push('⚠ Fecha de vencimiento pasada');

    if (avisos.length > 0) {
      const ok = await _whConfirm(avisos.join('\n') + '\n\n¿Continuar de todas formas?', { warning: true, titulo: 'Avisos' });
      if (!ok) return;
    }

    const btn = document.getElementById('btnRegistrarEnvasado');
    btn.disabled = true;
    btn.textContent = 'Registrando...';

    // idempotencyKey único POR CLICK — protege ante retries por timeout.
    const idempotencyKey = 'ENV-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);

    // ID temporal del envasado optimista (para poder REMOVERLO si falla el backend)
    const idEnvOptimista = 'ENV_OPT_' + Date.now();

    // Snapshot de las cantidades aplicadas optimistamente — sirve para
    // hacer rollback exacto si el backend falla.
    const cbDerivado = String(prod.codigoBarra || '');
    const cbBaseStr  = prodBase?.codigoBarra ? String(prodBase.codigoBarra) : '';

    // Optimistic: inyectar en caché y cerrar modal de inmediato.
    // [v2.10.5] Incluir descripciones legibles para que la card del historial
    // se vea bien aunque el cache local de productos no tenga los códigos.
    OfflineManager.inyectarEnvasadoCache({
      idEnvasado:             idEnvOptimista,
      codigoProductoBase:     cbBaseStr || prod.codigoProductoBase || '',
      cantidadBase:           cantBase,
      unidadBase:             prodBase?.unidad || '',
      codigoProductoEnvasado: cbDerivado,
      unidadesProducidas:     producidas,
      mermaReal:              0,
      eficienciaPct:          100,
      fecha:                  new Date().toISOString().split('T')[0],
      usuario:                window.WH_CONFIG.usuario,
      estado:                 'COMPLETADO',
      descripcionProductoEnvasado: prod.descripcion || '',
      descripcionProductoBase:     prodBase?.descripcion || ''
    });
    toast(`${producidas} uds registradas${imprimir ? ' · enviando etiquetas...' : ''}`, 'ok', 4000);
    // [v2.10.4] TTS optimista: hablar AL TOCAR el botón (no al volver del
    // backend). Antes el operador esperaba 5+ segundos hasta escuchar la voz,
    // perdiendo el efecto anti-fraude. Ahora suena inmediato.
    _decirEnVoz(`${producidas} unidades registradas de ${prod.descripcion || cbDerivado}`);
    cerrarSheet('sheetEnvasado');
    // NOTA: NO llamar cargar() acá. cargar() hace API.getEnvasados() y
    // sobrescribe el cache con la lista del backend, que aún no tiene
    // este envasado (el POST va en background). Antes esto hacía
    // "parpadear" la lista. Sólo re-renderizamos local desde cache.
    _renderDesdeCache();

    // Patch stock cache local
    if (cbBaseStr) OfflineManager.patchStockCache(cbBaseStr, -cantBase);
    OfflineManager.patchStockCache(cbDerivado, producidas);
    window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['stock'] } }));

    // Helper de ROLLBACK — revierte TODO lo optimista cuando el backend falla.
    // Antes no existía esto: el patch de stock quedaba aplicado aunque el
    // envasado no se guardara, divergiendo el cache local del real.
    function _rollbackOptimista(motivo) {
      try {
        if (cbBaseStr) OfflineManager.patchStockCache(cbBaseStr, +cantBase);
        OfflineManager.patchStockCache(cbDerivado, -producidas);
        OfflineManager.removerEnvasadoCache(idEnvOptimista);
        _renderDesdeCache();
        window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['stock'] } }));
      } catch(_) {}
      toast('⚠ Envasado revertido · ' + motivo + ' · volvé a intentar', 'danger', 7000);
    }

    // [v2.13.115 OPTIMISTIC PRO] Disparar el modal de LOTE ya mismo, antes
    // del await de registrarEnvasado. Antes el modal aparecía 15s después
    // (esperaba a que registrarEnvasado responda). Ahora aparece instantáneo.
    // El lote se crea en PARALELO al registro del envasado.
    // Si registrarEnvasado falla, cancelamos el lote en el rollback.
    if (imprimir) {
      try {
        WhLoteAdhesivo.crearYEjecutar({
          codigoBarra:     String(prod.codigoBarra),
          descripcion:     prod.descripcion || '',
          total:           producidas,
          fechaEnvasado:   fechaVenc || new Date().toISOString().split('T')[0]
        });
      } catch (e) {
        toast('No se pudo iniciar lote: ' + (e?.message || ''), 'danger', 5000);
      }
    }

    // [v2.13.108] Si imprimir=true, NO le decimos al backend que imprima
    // directo (el método viejo manda 1 sola request a PrintNode → drift).
    // GAS en segundo plano
    API.registrarEnvasado({
      codigoBarra:        prod.codigoBarra,
      unidadesProducidas: producidas,
      fechaVencimiento:   fechaVenc,
      imprimirEtiquetas:  false,    // ← el lote se encarga ahora
      usuario:            window.WH_CONFIG.usuario,
      idempotencyKey:     idempotencyKey
    }).then(res => {
      if (!res || res.ok === false) {
        _rollbackOptimista(res?.error || 'Error desconocido');
        // [v2.13.115/117] Si el envasado falló, cancelar el lote optimista
        // que disparamos en paralelo, para evitar imprimir etiquetas de
        // un envasado que no se registró. Usar cancelarSilencioso para
        // NO mostrar confirm al usuario (rollback automático).
        if (imprimir) {
          try { WhLoteAdhesivo.cancelarSilencioso('Envasado no se registró'); } catch(_) {}
        }
        return;
      }
      // Reemplazar el envasado optimista con el idEnvasado REAL del backend
      const idReal = res.data?.idEnvasado || idEnvOptimista;
      if (idReal !== idEnvOptimista) {
        OfflineManager.removerEnvasadoCache(idEnvOptimista);
        OfflineManager.inyectarEnvasadoCache({
          idEnvasado:             idReal,
          codigoProductoBase:     cbBaseStr || prod.codigoProductoBase || '',
          cantidadBase:           cantBase,
          unidadBase:             prodBase?.unidad || '',
          codigoProductoEnvasado: cbDerivado,
          descripcionProductoEnvasado: prod.descripcion || '',
          descripcionProductoBase:     prodBase?.descripcion || '',
          unidadesProducidas:     producidas,
          mermaReal:              0,
          eficienciaPct:          100,
          fecha:                  new Date().toISOString().split('T')[0],
          usuario:                window.WH_CONFIG.usuario,
          estado:                 'COMPLETADO',
          idGuiaSalida:           res.data.idGuiaSalida || '',
          idGuiaIngreso:          res.data.idGuiaIngreso || ''
        });
        _renderDesdeCache();
      }
      OfflineManager.precargarOperacional(true).catch(() => {});

      // 🎉 Celebración + banner deshacer
      _celebrarEnvasado(idReal, prod.descripcion || cbDerivado, producidas);
      // El lote de impresión ya se disparó ANTES (línea 8688 aprox) — corre
      // en paralelo y no depende del registro del envasado.
    }).catch((e) => {
      _rollbackOptimista('sin conexión');
    }).finally(() => {
      btn.disabled = false;
      btn.textContent = 'Registrar envasado';
    });
  }

  // Anular un envasado individual (error de captura, corrección manual).
  // Optimista: revierte stock cache + marca como ANULADO en cache, llama
  // backend, y hace rollback si el backend falla.
  async function anular(idEnvasado, codigoDerivado, unidades) {
    if (!idEnvasado) return;
    if (!await _whConfirm('¿Anular este envasado?\n\nRevierte el stock consumido y producido. El registro queda como ANULADO en el historial.', { danger: true, titulo: 'Anular envasado', okText: 'Anular' })) return;
    // Buscar la entrada en cache para hacer rollback exacto si falla
    const cache = OfflineManager.getEnvasadosCache();
    const env = cache.find(e => e.idEnvasado === idEnvasado);
    if (!env) { toast('Envasado no encontrado en cache local', 'warn'); return; }

    const cbBase  = String(env.codigoProductoBase || '');
    const cbDer   = String(env.codigoProductoEnvasado || codigoDerivado || '');
    const cantB   = parseFloat(env.cantidadBase) || 0;
    const uds     = parseFloat(env.unidadesProducidas || unidades) || 0;
    const estadoPrev = env.estado;

    // Optimista: revertir stock + marcar como ANULADO
    if (cbBase) OfflineManager.patchStockCache(cbBase, +cantB);
    if (cbDer)  OfflineManager.patchStockCache(cbDer,  -uds);
    env.estado = 'ANULADO';
    OfflineManager.guardarEnvasadosCache(cache);
    _renderDesdeCache();
    window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['stock'] } }));
    toast('↺ Envasado anulado · stock revertido', 'ok', 3000);

    try {
      const res = await API.anularEnvasadoManual({
        idEnvasado: idEnvasado,
        usuario:    window.WH_CONFIG?.usuario || 'manual',
        motivo:     'anulación manual desde lista'
      });
      if (!res || res.ok === false) {
        throw new Error(res?.error || 'Error backend');
      }
      OfflineManager.precargarOperacional(true).catch(() => {});
    } catch(e) {
      // Rollback: devolver el stock que revertimos y restaurar estado
      if (cbBase) OfflineManager.patchStockCache(cbBase, -cantB);
      if (cbDer)  OfflineManager.patchStockCache(cbDer,  +uds);
      env.estado = estadoPrev || 'COMPLETADO';
      OfflineManager.guardarEnvasadosCache(cache);
      _renderDesdeCache();
      window.dispatchEvent(new CustomEvent('wh:data-refresh', { detail: { changed: ['stock'] } }));
      toast('⚠ No se pudo anular: ' + (e.message || e) + ' — reintentá', 'danger', 6000);
    }
  }

  // Celebración + banner deshacer rápido tras registrar exitosamente.
  // El operador tiene 12s para deshacer si se equivocó — un click anula
  // todo (revierte stock + marca ANULADO).
  let _bannerDeshacerTimer = null;
  function _celebrarEnvasado(idReal, descripcion, uds) {
    try {
      if (typeof SoundFX !== 'undefined') {
        if (SoundFX.done) SoundFX.done();
        else if (SoundFX.savedTick) SoundFX.savedTick();
      }
    } catch(_){}
    try { vibrate([20, 30, 60]); } catch(_){}
    // Confetti inline simple: emojis flotantes
    _confettiEnvasado();
    _mostrarBannerDeshacer(idReal, descripcion || '', uds || 0);
    // [v2.10.4] TTS movido a registrar() para hablar al instante del click,
    // no después de la confirmación del backend (era ~5s tarde).
  }

  // TTS — usa SpeechSynthesis (nativo). Velocidad 0.95, español PE/ES.
  // Cancela cualquier voz previa antes de hablar (evita acumulación).
  function _decirEnVoz(texto) {
    try {
      if (!('speechSynthesis' in window)) return;
      const u = new SpeechSynthesisUtterance(String(texto || ''));
      u.lang = 'es-PE';
      u.rate = 0.95;
      u.pitch = 1;
      u.volume = 1;
      window.speechSynthesis.cancel();
      window.speechSynthesis.speak(u);
    } catch(_){}
  }

  function _confettiEnvasado() {
    const cap = document.createElement('div');
    cap.className = 'env-confetti-layer';
    cap.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:9994;overflow:hidden';
    const emojis = ['📦','✨','🎉','🟢','🟢','📦'];
    for (let i = 0; i < 12; i++) {
      const sp = document.createElement('span');
      sp.textContent = emojis[i % emojis.length];
      sp.style.cssText = `position:absolute;left:${10+Math.random()*80}%;top:-20px;font-size:${18+Math.random()*14}px;animation:envConfettiFall ${1.2+Math.random()*0.8}s cubic-bezier(.6,.04,.98,.34) forwards;animation-delay:${Math.random()*.2}s`;
      cap.appendChild(sp);
    }
    document.body.appendChild(cap);
    setTimeout(() => cap.remove(), 2200);
  }

  function _mostrarBannerDeshacer(idEnvasado, descripcion, uds) {
    if (_bannerDeshacerTimer) { clearTimeout(_bannerDeshacerTimer); _bannerDeshacerTimer = null; }
    let banner = document.getElementById('envBannerDeshacer');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'envBannerDeshacer';
      banner.className = 'env-banner-deshacer';
      document.body.appendChild(banner);
    }
    const idAttr = String(idEnvasado).replace(/'/g, '&#39;');
    const desc = String(descripcion).substring(0, 32);
    banner.innerHTML = `
      <div class="env-banner-inner">
        <span class="env-banner-ico">✅</span>
        <div class="env-banner-text">
          <p class="env-banner-title">${uds} uds · ${escHtml(desc)}</p>
          <p class="env-banner-sub" id="envBannerCd">Deshacer disponible 12s</p>
        </div>
        <button class="env-banner-btn"
                onclick="EnvasadosView._dispararDeshacer('${idAttr}')">↺ Deshacer</button>
      </div>`;
    banner.classList.add('is-open');
    let secs = 12;
    const cdEl = document.getElementById('envBannerCd');
    const tick = () => {
      secs--;
      if (cdEl) cdEl.textContent = secs > 0 ? `Deshacer disponible ${secs}s` : 'Deshacer expiró';
      if (secs <= 0) { _ocultarBannerDeshacer(); return; }
      _bannerDeshacerTimer = setTimeout(tick, 1000);
    };
    _bannerDeshacerTimer = setTimeout(tick, 1000);
  }
  function _ocultarBannerDeshacer() {
    if (_bannerDeshacerTimer) { clearTimeout(_bannerDeshacerTimer); _bannerDeshacerTimer = null; }
    const b = document.getElementById('envBannerDeshacer');
    if (b) b.classList.remove('is-open');
  }
  function _dispararDeshacer(idEnvasado) {
    _ocultarBannerDeshacer();
    anular(idEnvasado, '', 0);
  }

  // Re-render desde el cache local (sin llamar al backend). Usado tras
  // inyección/rollback optimista para que la UI refleje el cache actual.
  function _renderDesdeCache() {
    try {
      const fd  = _fechaDesde7Dias();
      const cnt = document.getElementById('listEnvasados');
      if (!cnt) return;
      const cached = OfflineManager.getEnvasadosCache()
        .filter(e => String(e.fecha || '').substring(0, 10) >= fd);
      _renderEnvasadosPorDia(cached, cnt);
    } catch(_) {}
  }

  // ────────────────────────────────────────────────────────────
  // AUTH ADMIN para editar/anular envasados — flow de 2 modales
  // ────────────────────────────────────────────────────────────
  // Patrón estilo ME (MosExpress):
  //   1) Click "Editar" o "Anular" → abre modal "Llamar al admin"
  //      (solo clave 8 dígitos). El operador NO conoce esta clave.
  //   2) Admin viene, ingresa clave → valida contra MOS.
  //   3) Si OK → abre modal de la operación (cantidad/motivo).
  //   4) Operador (o admin) completa los datos → backend ejecuta
  //      revalidando la clave (defensa en profundidad).
  //
  // Sin "recordar 30 min": cada operación pide autorización fresca,
  // garantía de que el admin esté presente para cada edición/anulación.
  let _envAuthCtx = null;   // { modo, idEnvasado, descripcion, udsActuales, cbDerivado, clave? }

  // PASO 1 — abrir modal de autorización admin
  function _abrirModalAuth(modo, idEnvasado) {
    const cache = OfflineManager.getEnvasadosCache();
    const env = cache.find(e => e.idEnvasado === idEnvasado);
    if (!env) { toast('Envasado no encontrado en caché local', 'warn'); return; }
    const descripcion = env.descripcionProductoEnvasado
                     || env.codigoProductoEnvasado
                     || idEnvasado;
    _envAuthCtx = {
      modo,
      idEnvasado,
      descripcion,
      udsActuales: parseFloat(env.unidadesProducidas) || 0,
      cbDerivado:  String(env.codigoProductoEnvasado || ''),
      clave:       null
    };
    document.getElementById('envAuthSub').textContent =
      modo === 'editar'
        ? 'Para EDITAR la cantidad, el admin debe ingresar su clave.'
        : 'Para ANULAR este envasado, el admin debe ingresar su clave.';
    document.getElementById('envAuthProd').textContent = descripcion;
    document.getElementById('envAuthCantActual').textContent =
      `Cantidad registrada: ${_envAuthCtx.udsActuales} uds`;
    document.getElementById('envAuthClave').value = '';
    document.getElementById('envAuthErr').textContent = '';
    document.getElementById('overlayEnvAuth').style.display = 'block';
    document.getElementById('sheetEnvAuth').classList.add('open');
    setTimeout(() => document.getElementById('envAuthClave').focus(), 200);

    // Enter en el input dispara la validación
    document.getElementById('envAuthClave').onkeydown = (ev) => {
      if (ev.key === 'Enter') validarAuth();
    };
  }

  // [v2.12] Cierra el sheet de auth. Por default LIMPIA el ctx (cancelación
  // del usuario). Pasar preservarCtx=true cuando el ctx debe sobrevivir para
  // el siguiente paso (ej. al transicionar al modal de acción tras validar).
  function cerrarAuth(preservarCtx) {
    document.getElementById('overlayEnvAuth').style.display = 'none';
    document.getElementById('sheetEnvAuth').classList.remove('open');
    if (!preservarCtx) _envAuthCtx = null;
  }

  function pedirAuthEditar(idEnvasado) { _abrirModalAuth('editar', idEnvasado); }
  function pedirAuthAnular(idEnvasado) { _abrirModalAuth('anular', idEnvasado); }

  // Helpers de overlay loading dentro de los bottom-sheets
  function _toggleLoading(sheetId, loadingId, on, opts) {
    const sheet = document.getElementById(sheetId);
    const ovl   = document.getElementById(loadingId);
    if (!sheet || !ovl) return;
    if (on) {
      if (opts && opts.title) {
        const t = ovl.querySelector('.sheet-loading-title');
        if (t) t.innerHTML = opts.title + '<span class="sheet-loading-dots"></span>';
      }
      if (opts && opts.sub) {
        const s = ovl.querySelector('.sheet-loading-sub');
        if (s) s.textContent = opts.sub;
      }
      ovl.classList.add('is-on');
      sheet.classList.add('is-processing');
    } else {
      ovl.classList.remove('is-on');
      sheet.classList.remove('is-processing');
    }
  }

  // [v2.13.38] Catálogo de tiers de auth en WH (alineado con MOS)
  const _WH_AUTH_TIERS = {
    'REABRIR_GUIA': 1, 'EDITAR_ENVASADO': 1,
    'ANULAR_ENVASADO': 2, 'APROBAR_DISPOSITIVO_INSITU': 2, 'PROCESAR_MERMAS': 2
  };
  // _envAuthT0 — timestamp para medir tiempo_verify_ms
  let _envAuthT0 = 0;

  // Valida la clave contra MOS (via WH backend). Si ok, pasa al PASO 2.
  async function validarAuth() {
    if (!_envAuthCtx) return;
    if (!_envAuthT0) _envAuthT0 = Date.now();
    const clave = document.getElementById('envAuthClave').value.trim();
    const errEl = document.getElementById('envAuthErr');
    if (clave.length !== 8 || !/^\d+$/.test(clave)) {
      errEl.textContent = 'La clave debe ser 8 dígitos numéricos';
      return;
    }
    errEl.textContent = '';
    _toggleLoading('sheetEnvAuth', 'envAuthLoading', true);
    try {
      const accion = _envAuthCtx.modo === 'editar' ? 'EDITAR_ENVASADO' : 'ANULAR_ENVASADO';
      const tier = _WH_AUTH_TIERS[accion] || 2;
      const tiempo_verify_ms = Date.now() - _envAuthT0;
      const devId = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
      const res = await API.verificarClaveAdmin({
        clave,
        accion,
        refDocumento: _envAuthCtx.idEnvasado,
        appOrigen: 'warehouseMos',
        // [v2.13.38] Auditoría enriquecida
        tier: tier,
        cache_hit: 0,
        tiempo_verify_ms: tiempo_verify_ms,
        deviceId: devId,
        dispositivo: devId
      });
      _toggleLoading('sheetEnvAuth', 'envAuthLoading', false);
      if (!res || !res.ok) {
        errEl.textContent = '✗ ' + (res?.error || 'Clave incorrecta');
        return;
      }
      // OK: guardar clave para revalidación en backend final + abrir PASO 2
      _envAuthCtx.clave        = clave;
      _envAuthCtx.validadoPor  = (res.data && res.data.validadoPor) || 'admin';
      _envAuthCtx.idPersonal   = (res.data && res.data.idPersonal)  || '';
      // [v2.12 FIX] Cerrar PASO 1 PRESERVANDO el ctx + abrir PASO 2
      // Antes: cerrarAuth() borraba _envAuthCtx → _abrirModalAccion salía
      // silencio (chequea if !_envAuthCtx return) → modal nunca se mostraba.
      cerrarAuth(true);
      // Pequeño feedback visual de éxito antes del siguiente modal
      try { _envBeep && _envBeep('ok'); } catch(_){}
      _abrirModalAccion();
    } catch(e) {
      _toggleLoading('sheetEnvAuth', 'envAuthLoading', false);
      errEl.textContent = '✗ Error de conexión: ' + (e.message || e);
    }
  }

  // [v2.12] WebAudio simple para feedback (ok = 2 tonos asc, success = arpegio)
  let _envAudioCtx = null;
  function _envBeep(tipo) {
    try {
      if (!_envAudioCtx) _envAudioCtx = new (window.AudioContext || window.webkitAudioContext)();
      const ctx = _envAudioCtx; const now = ctx.currentTime;
      const notas = tipo === 'ok' ? [880, 1320] : tipo === 'success' ? [659, 880, 1320] : [400];
      notas.forEach((f, i) => {
        const o = ctx.createOscillator(), g = ctx.createGain();
        const t0 = now + i * 0.08;
        o.type = 'sine'; o.frequency.setValueAtTime(f, t0);
        g.gain.setValueAtTime(0.0001, t0);
        g.gain.exponentialRampToValueAtTime(0.12, t0 + 0.015);
        g.gain.exponentialRampToValueAtTime(0.001, t0 + 0.25);
        o.connect(g).connect(ctx.destination);
        o.start(t0); o.stop(t0 + 0.27);
      });
    } catch(_){}
  }

  // PASO 2 — modal con los detalles de la operación (post-auth)
  function _abrirModalAccion() {
    if (!_envAuthCtx) return;
    const ctx = _envAuthCtx;
    document.getElementById('envAccionTitulo').textContent =
      ctx.modo === 'editar' ? '✏ Editar cantidad' : '🚫 Anular envasado';
    document.getElementById('envAccionProd').textContent = ctx.descripcion;
    document.getElementById('envAccionCantActual').textContent =
      `Cantidad actual: ${ctx.udsActuales} uds`;
    // [v2.12] Mostrar quién autorizó (banner verde sutil arriba del modal)
    const autorizadorEl = document.getElementById('envAccionAutorizador');
    if (autorizadorEl) {
      autorizadorEl.innerHTML = `🔓 Autorizado por <b>${(ctx.validadoPor || 'admin').replace(/[<>&"]/g, '')}</b>`;
      autorizadorEl.classList.remove('hidden');
    }
    document.getElementById('envAccionCantRow').classList.toggle('hidden', ctx.modo !== 'editar');
    document.getElementById('envAccionCant').value = ctx.modo === 'editar' ? ctx.udsActuales : '';
    document.getElementById('envAccionDelta').textContent = '';
    document.getElementById('envAccionMotivo').value = '';
    document.getElementById('btnEnvAccionOk').textContent =
      ctx.modo === 'editar' ? 'Aplicar cambio' : 'Confirmar anulación';
    document.getElementById('overlayEnvAccion').style.display = 'block';
    document.getElementById('sheetEnvAccion').classList.add('open');

    if (ctx.modo === 'editar') {
      const inp = document.getElementById('envAccionCant');
      inp.oninput = () => {
        const n = parseFloat(inp.value) || 0;
        const d = n - ctx.udsActuales;
        const el = document.getElementById('envAccionDelta');
        if (d === 0) { el.textContent = ''; return; }
        const signo = d > 0 ? '+' : '';
        el.textContent = `Cambio: ${signo}${d} uds`;
        el.style.color = d > 0 ? '#34d399' : '#f87171';
      };
      setTimeout(() => inp.focus(), 200);
    } else {
      setTimeout(() => document.getElementById('envAccionMotivo').focus(), 200);
    }
  }

  function cerrarAccion() {
    document.getElementById('overlayEnvAccion').style.display = 'none';
    document.getElementById('sheetEnvAccion').classList.remove('open');
    _envAuthCtx = null;
  }

  async function confirmarAccion() {
    if (!_envAuthCtx || !_envAuthCtx.clave) {
      toast('Autorización vencida, volvé a iniciar', 'warn');
      cerrarAccion();
      return;
    }
    const ctx    = _envAuthCtx;
    const motivo = document.getElementById('envAccionMotivo').value.trim();
    if (!motivo) { toast('El motivo es obligatorio', 'warn'); return; }

    // Validaciones tempranas (antes del overlay)
    let nuevasUds = null;
    if (ctx.modo === 'editar') {
      nuevasUds = parseFloat(document.getElementById('envAccionCant').value);
      if (!isFinite(nuevasUds) || nuevasUds < 0) {
        toast('Cantidad inválida', 'warn'); return;
      }
      if (nuevasUds === ctx.udsActuales) {
        toast('No hay cambio en la cantidad', 'warn'); return;
      }
    }

    // Activar overlay con mensaje contextual según la operación
    const loadingTitle = ctx.modo === 'editar' ? 'Corrigiendo cantidad' : 'Anulando envasado';
    const loadingSub   = ctx.modo === 'editar'
      ? 'Ajustando stock base, stock derivado y guías de ingreso/salida'
      : 'Revirtiendo stock y anulando detalles de guías';
    _toggleLoading('sheetEnvAccion', 'envAccionLoading', true, { title: loadingTitle, sub: loadingSub });

    try {
      let res;
      if (ctx.modo === 'editar') {
        res = await API.corregirUnidadesEnvasado({
          idEnvasado:     ctx.idEnvasado,
          nuevasUnidades: nuevasUds,
          motivo,
          usuario:        window.WH_CONFIG.usuario,
          claveAdmin:     ctx.clave
        });
      } else {
        res = await API.anularEnvasadoConClave({
          idEnvasado: ctx.idEnvasado,
          motivo,
          usuario:    window.WH_CONFIG.usuario,
          claveAdmin: ctx.clave
        });
      }

      _toggleLoading('sheetEnvAccion', 'envAccionLoading', false);

      if (!res || !res.ok) {
        toast('✗ ' + (res?.error || 'Error desconocido'), 'danger', 5000);
        return;
      }

      cerrarAccion();
      const d = res.data || {};
      // [v2.12] Sonido de éxito + toast con resumen detallado del cambio
      try { _envBeep && _envBeep('success'); } catch(_){}
      if (ctx.modo === 'editar') {
        const delta = d.udsNuevas - d.udsViejas;
        const signo = delta > 0 ? '+' : '';
        toast(`✓ Corregido · ${d.udsViejas}→${d.udsNuevas} uds (${signo}${delta}) · por ${ctx.validadoPor}`, 'ok', 6000);
        _decirEnVoz(`Corregido. ${d.udsViejas} cambiado a ${d.udsNuevas} unidades de ${d.descripcion || ctx.descripcion}`);
      } else {
        const baseRev = d.cantBaseRestit ? ` · ${d.cantBaseRestit} ${ctx.unidadBase || 'KGM'} restituidos al base` : '';
        toast(`✓ Anulado · ${d.udsAnuladas} uds revertidas${baseRev} · por ${ctx.validadoPor}`, 'ok', 6000);
        _decirEnVoz(`${d.udsAnuladas} unidades anuladas de ${d.descripcion || ctx.descripcion}`);
      }
      cargar();
      OfflineManager.precargarOperacional(true).catch(() => {});

    } catch(e) {
      _toggleLoading('sheetEnvAccion', 'envAccionLoading', false);
      toast('✗ Error de conexión: ' + (e.message || e), 'danger', 5000);
    }
  }

  return { cargar, nuevo, onDerivadoChange, calcularProyeccion, ajustarUnidades, setUnidades, registrar, anular,
           filtrarDerivados, seleccionarDerivado, cambiarDerivado,
           pedirAuthEditar, pedirAuthAnular, cerrarAuth, validarAuth, cerrarAccion, confirmarAccion,
           _dispararDeshacer };
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
    let list = _filtroUrg ? _catalog.filter(b => b.worstUrg) : _catalog;
    // [v2.13.33] Filtro búsqueda inteligente
    if (typeof _searchQuery !== 'undefined' && _searchQuery) {
      list = list.filter(b => _matchBase(b, _searchQuery));
    }

    if (!list.length) {
      if (typeof _searchQuery !== 'undefined' && _searchQuery) {
        container.innerHTML = `<div class="card text-center py-8"><p class="text-2xl mb-2">🔍</p><p class="font-semibold">Sin resultados para "${escHtml(_searchQuery)}"</p><p class="text-xs text-slate-500 mt-1">Probá con menos palabras o el código de barras</p></div>`;
      } else {
        container.innerHTML = _filtroUrg
          ? '<div class="card text-center py-8"><p class="text-2xl mb-2">✅</p><p class="font-semibold">¡Sin urgentes!</p></div>'
          : '<p class="text-slate-500 text-center py-8 text-sm">Sin productos envasables configurados</p>';
      }
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
    // [v2.13.116] NO activar filtro de urgentes por default.
    // El operador ve TODA la lista al entrar; el rayito en cards de
    // productos urgentes ya da la señal visual. Si quiere filtrar, hace
    // toggle manual del botón ⚡.
    _filtroUrg = false;
    _updateUrgBtn();
    _render();

    _timer = setInterval(() => {
      _catalog = _buildCatalog();
      _updateUrgBtn();
      _render();
    }, 120_000);
  }

  // [v2.13.116] Historial: 14 días, agrupado por día nombrado, con buscador
  // que filtra por producto + usuario.
  let _histQuery = '';
  let _histCacheLista = [];  // cache de la última fusión para re-render rápido al buscar

  async function verHistorial() {
    document.getElementById('envCatalogPanel').classList.add('hidden');
    document.getElementById('envHistorialPanel').classList.remove('hidden');
    const fd        = _fechaDesde14Dias();
    const container = document.getElementById('listEnvasadorHistorial');

    // Limpiar buscador al entrar
    const inp = document.getElementById('envHistSearch');
    if (inp) inp.value = '';
    _histQuery = '';

    // Optimistic: mostrar caché de inmediato
    const cached = OfflineManager.getEnvasadosCache()
      .filter(e => String(e.fecha || '').substring(0, 10) >= fd);
    _histCacheLista = cached;
    if (cached.length) {
      _renderEnvasadosPorDia(cached, container, { modo: 'historial14d', query: _histQuery });
    } else {
      container.innerHTML = '<div class="flex justify-center py-8"><div class="spinner"></div></div>';
    }

    // Fondo: actualizar desde servidor con rango de 14 días
    const res = await API.getEnvasados({ fechaDesde: fd }).catch(() => ({ ok: false }));
    if (res.ok) {
      const cacheLocal = OfflineManager.getEnvasadosCache();
      const idsBackend = new Set(res.data.map(e => String(e.idEnvasado || '')));
      const pendientes = cacheLocal.filter(e =>
        e.idEnvasado && !idsBackend.has(String(e.idEnvasado))
      );
      const fusionado = pendientes.length ? [...pendientes, ...res.data] : res.data;
      OfflineManager.guardarEnvasadosCache(fusionado);
      _histCacheLista = fusionado;
      _renderEnvasadosPorDia(fusionado, container, { modo: 'historial14d', query: _histQuery });
    }
  }

  // [v2.13.116] Llamado desde el oninput del buscador (HTML envHistSearch)
  function buscarHistorial(q) {
    _histQuery = String(q || '').trim();
    const container = document.getElementById('listEnvasadorHistorial');
    if (!container) return;
    _renderEnvasadosPorDia(_histCacheLista, container, { modo: 'historial14d', query: _histQuery });
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

  // [v2.13.33] Búsqueda inteligente: voz + código (canon+equiv) + descripción
  let _searchQuery = '';
  let _searchOpen = false;

  function searchToggle() {
    _searchOpen = !_searchOpen;
    const tb = document.getElementById('envToolbar');
    const inp = document.getElementById('envSearchInput');
    if (!tb) return;
    tb.classList.toggle('env-search-open', _searchOpen);
    try { SoundFX && SoundFX.ping && SoundFX.ping(); } catch(_){}
    if (_searchOpen) {
      setTimeout(() => { try { inp && inp.focus(); } catch(_){} }, 300);
    } else {
      // Limpiar al cerrar
      _searchQuery = '';
      if (inp) inp.value = '';
      try { Voice && Voice.stopListen && Voice.stopListen(); } catch(_){}
      const mic = document.getElementById('envSearchMic');
      if (mic) mic.classList.remove('recording');
      _render();
    }
  }

  function searchInput(val) {
    _searchQuery = String(val || '').trim();
    _render();
    // Si es código de barras escaneado (8+ dígitos puros), sonido + flash al match
    if (/^\d{8,}$/.test(_searchQuery)) {
      setTimeout(() => {
        const cards = document.querySelectorAll('#listEnvasadorCatalog .env-base-card');
        if (cards.length === 1) {
          cards[0].classList.add('env-search-result-flash');
          try { SoundFX && SoundFX.beep && SoundFX.beep(); } catch(_){}
          setTimeout(() => cards[0].classList.remove('env-search-result-flash'), 700);
        }
      }, 50);
    }
  }

  function searchVoice() {
    if (!window.Voice || !Voice.supported || !Voice.supported().stt) {
      _whAlert('🎤 Tu navegador no soporta reconocimiento de voz.', { titulo: 'Voz no disponible' });
      return;
    }
    const mic = document.getElementById('envSearchMic');
    const inp = document.getElementById('envSearchInput');
    if (!inp) return;
    if (mic && mic.classList.contains('recording')) {
      try { Voice.stopListen(); } catch(_){}
      mic.classList.remove('recording');
      return;
    }
    if (mic) mic.classList.add('recording');
    try { SoundFX && SoundFX.click && SoundFX.click(); } catch(_){}
    Voice.listen({
      lang: 'es-PE',
      interim: true,
      onResult: (txt, isFinal) => {
        inp.value = txt;
        searchInput(txt);
        if (isFinal) {
          if (mic) mic.classList.remove('recording');
        }
      },
      onError: () => { if (mic) mic.classList.remove('recording'); },
      onEnd:   () => { if (mic) mic.classList.remove('recording'); }
    });
  }

  // Match: prioridad código → descripción
  // Si query es numérica 8+ dígitos: solo código (canon + equiv)
  // Si no: descripción (case-insensitive, tokens, sin acentos)
  function _matchBase(base, query) {
    if (!query) return true;
    const q = query.toLowerCase().trim();
    const esCodigo = /^\d{4,}$/.test(q);

    if (esCodigo) {
      // Buscar en codigoBarra del base + sus derivados
      const codigosBase = [base.codigoBarra, base.idProducto]
        .filter(Boolean).map(c => String(c).toLowerCase());
      const codigosDeriv = (base.derivados || []).flatMap(d =>
        [d.codigoBarra, d.idProducto].filter(Boolean).map(c => String(c).toLowerCase())
      );
      // Equivalencias por skuBase
      const equivs = (typeof OfflineManager !== 'undefined' && OfflineManager.getEquivalenciasCache)
        ? OfflineManager.getEquivalenciasCache() : [];
      const skuBase = String(base.skuBase || base.idProducto || '').toUpperCase();
      const equivCodes = equivs
        .filter(e => String(e.skuBase || '').toUpperCase() === skuBase)
        .map(e => String(e.codigoBarra || '').toLowerCase());
      const todos = [...codigosBase, ...codigosDeriv, ...equivCodes];
      return todos.some(c => c.includes(q));
    }

    // Descripción — case + sin acentos, tokens
    const norm = s => String(s || '')
      .toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
    const desc = norm(base.descripcion) + ' ' +
      (base.derivados || []).map(d => norm(d.descripcion)).join(' ');
    const tokens = norm(q).split(/\s+/).filter(Boolean);
    return tokens.every(t => desc.includes(t));
  }

  return { cargar, toggleUrgFilter, verHistorial, verCatalogo, envasar, silentRefresh,
           searchToggle, searchInput, searchVoice,
           buscarHistorial };
})();

// ════════════════════════════════════════════════
// PRINT HUB — selector de impresora para admin/master
// ════════════════════════════════════════════════
// Envuelve TODA impresión de WH (guías, membretes, historial, cargadores).
//  - Usuario normal → imprime directo en la impresora del almacén.
//  - Admin/master   → modal moderno para elegir cualquier impresora activa
//    del ecosistema (puede estar físicamente en WH pero mandar a una caja).
const PrintHub = (() => {
  let _pendiente = null; // { apiMethod, params, resolve }

  function _esAdminMaster() {
    const r = String(window.WH_CONFIG?.rol || '').toUpperCase();
    return r === 'MASTER' || r === 'ADMINISTRADOR';
  }

  // [v2.13.37] Filtros del modal — persisten entre aperturas en sesión.
  // filtroZona: null = mi zona (WH_CONFIG.zona) · '__all__' = todas las zonas
  // filtroTipo: null = todos · 'TICKET' · 'ADHESIVO' · 'ZPL'
  let _filtros = { filtroZona: null, filtroTipo: null };

  // Punto de entrada único. apiMethod = nombre del método en API
  // (ej. 'imprimirTicketGuia'). titulo = texto descriptivo para el modal.
  function imprimir(apiMethod, params, titulo) {
    params = params || {};
    if (!_esAdminMaster()) {
      // Usuario operativo → directo, sin modal
      return API[apiMethod](params);
    }
    return new Promise((resolve) => {
      _pendiente = { apiMethod, params, resolve };
      const ov  = document.getElementById('modalSelImpresora');
      const lst = document.getElementById('selImpresoraLista');
      const tit = document.getElementById('selImpresoraTitulo');
      if (tit) tit.textContent = titulo || 'Imprimir';
      // [v2.13.37] Pintar INSTANT desde cache si existe
      let pintadoFromCache = false;
      try {
        const raw = localStorage.getItem('wh_printers_cache');
        if (raw) {
          const cache = JSON.parse(raw);
          if (cache && cache.data && (Date.now() - (cache.ts || 0) < 5 * 60 * 1000)) {
            window._whPrintersCache = cache;
            if (lst) _renderLista(cache.data);
            pintadoFromCache = true;
          }
        }
      } catch(_){}
      if (!pintadoFromCache && lst) {
        lst.innerHTML = '<div class="selimp-loading">⏳ Cargando impresoras...</div>';
      }
      if (ov) { ov.style.display = 'flex'; requestAnimationFrame(() => ov.classList.add('is-open')); }
      // Siempre refrescar en bg (incluso si pintamos de cache)
      API.getImpresorasEcosistema().then(r => {
        const imps = (r && r.ok && r.data) || [];
        // Actualizar cache
        try {
          const cache = { ts: Date.now(), data: imps };
          window._whPrintersCache = cache;
          localStorage.setItem('wh_printers_cache', JSON.stringify(cache));
        } catch(_){}
        _renderLista(imps);
      }).catch(() => {
        if (!pintadoFromCache && lst) {
          lst.innerHTML = '<div class="selimp-err">No se pudieron cargar las impresoras. Toca "Almacén" para usar la de siempre.</div>' +
            '<button class="selimp-card selimp-default" onclick="PrintHub.elegir(\'\')"><div class="selimp-ico">🏭</div><div class="selimp-info"><div class="selimp-nombre">Impresora del Almacén</div><div class="selimp-meta">Por defecto</div></div></button>';
        }
      });
    });
  }

  // [v2.13.37] Aplica los filtros activos a una lista de impresoras
  function _aplicarFiltros(imps) {
    if (!Array.isArray(imps)) return [];
    const miZona = String(window.WH_CONFIG?.zona || '').trim();
    return imps.filter(p => {
      // Filtro por tipo
      if (_filtros.filtroTipo && String(p.tipo || '').toUpperCase() !== _filtros.filtroTipo) return false;
      // Filtro por zona — null = mi zona (default), '__all__' = todas
      if (_filtros.filtroZona === '__all__') return true;
      const zonaProp = String(p.idZona || '').trim();
      const zonaUser = _filtros.filtroZona || miZona;
      if (!zonaUser) return true; // si no hay zona del user, mostrar todas
      return zonaProp === zonaUser || !zonaProp; // permite impresoras sin zona asignada
    });
  }

  // [v2.13.37] Cambia un filtro y re-renderiza
  function toggleZona() {
    _filtros.filtroZona = _filtros.filtroZona === '__all__' ? null : '__all__';
    const imps = (window._whPrintersCache && window._whPrintersCache.data) || [];
    _renderLista(imps);
  }
  function toggleTipo() {
    const tipos = [null, 'TICKET', 'ADHESIVO', 'ZPL'];
    const idx = tipos.indexOf(_filtros.filtroTipo);
    _filtros.filtroTipo = tipos[(idx + 1) % tipos.length];
    const imps = (window._whPrintersCache && window._whPrintersCache.data) || [];
    _renderLista(imps);
  }

  // [v2.13.37] _renderLista refactorizado — filtros + agrupación por zona
  function _renderLista(impsRaw) {
    const lst = document.getElementById('selImpresoraLista');
    if (!lst) return;
    const _stTx = {
      ONLINE: 'online', PC_OFFLINE: 'PC apagada', PRINTER_OFFLINE: 'apagada',
      SIN_PAPEL: 'sin papel', SIN_TINTA: 'sin tinta', ATASCO: 'atasco',
      TAPA_ABIERTA: 'tapa abierta', PAUSED: 'pausada', DISABLED: 'deshabilitada',
      ERROR: 'error', SIN_ID: 'sin ID', ID_INVALIDO: 'ID inválido', UNKNOWN: 'sin señal'
    };
    // Aplicar filtros activos
    const imps = _aplicarFiltros(impsRaw);
    const miZona = String(window.WH_CONFIG?.zona || '').trim();
    const tipoLbl = _filtros.filtroTipo
      ? (_filtros.filtroTipo === 'TICKET' ? '🧾 tickets'
       : _filtros.filtroTipo === 'ADHESIVO' ? '🏷 adhesivos' : '📄 ZPL')
      : '🧾 todos los tipos';
    const zonaLbl = _filtros.filtroZona === '__all__' ? '🌐 todas las zonas' : `📍 ${miZona || 'mi zona'}`;
    const totalRaw = (impsRaw || []).length;
    const totalFilt = imps.length;
    // Barra de filtros + contador
    let html = `
      <div class="selimp-filtros">
        <button class="selimp-filtro-chip ${_filtros.filtroZona === '__all__' ? 'is-all' : ''}"
                onclick="PrintHub.toggleZona()">${zonaLbl}</button>
        <button class="selimp-filtro-chip selimp-filtro-tipo"
                onclick="PrintHub.toggleTipo()">${tipoLbl}</button>
        <span class="selimp-contador">${totalFilt === totalRaw ? totalRaw : (totalFilt + '/' + totalRaw)} impresora${totalFilt !== 1 ? 's' : ''}</span>
      </div>
      <button class="selimp-card selimp-default" onclick="PrintHub.elegir('')">
        <div class="selimp-ico">🏭</div>
        <div class="selimp-info">
          <div class="selimp-nombre">Impresora del Almacén</div>
          <div class="selimp-meta">Por defecto · donde estás ahora</div>
        </div>
      </button>`;
    if (!imps.length) {
      const msgEmpty = totalRaw > 0
        ? `<div class="selimp-empty">Ninguna impresora coincide con los filtros<br><span class="text-xs">Cambia los chips arriba o "🌐 todas las zonas"</span></div>`
        : `<div class="selimp-empty">No hay otras impresoras activas registradas en el ecosistema</div>`;
      html += msgEmpty;
      lst.innerHTML = html;
      return;
    }
    // Agrupar por zona+estación (mismo patrón que MOS _liqRenderPrinters)
    const zonas = {};
    imps.forEach(p => {
      const zk = p.idZona || '_sin_zona';
      const zLbl = p.zonaNombre || p.idZona || '(sin zona)';
      if (!zonas[zk]) zonas[zk] = { label: zLbl, estaciones: {}, sinEst: [] };
      if (p.idEstacion) {
        const ek = p.idEstacion;
        if (!zonas[zk].estaciones[ek]) {
          zonas[zk].estaciones[ek] = { label: p.estacionNombre || p.idEstacion, items: [] };
        }
        zonas[zk].estaciones[ek].items.push(p);
      } else {
        zonas[zk].sinEst.push(p);
      }
    });
    const _renderCard = (i) => {
      const ico = String(i.app || i.appOrigen || '').indexOf('express') >= 0 ? '🛒' : '🏭';
      const enUso = i.enUso
        ? ` · 🟢 en uso${i.enUsoPor ? ' (' + escHtml(i.enUsoPor) + ')' : ''}`
        : '';
      const st = i.state || (i.online ? 'ONLINE' : 'UNKNOWN');
      const stTxt = _stTx[st] || 'sin señal';
      const stIcon = i.icon || (st === 'ONLINE' ? '🟢' : '❔');
      const disabled = st !== 'ONLINE';
      const cardCls = `selimp-card ${i.enUso ? 'is-enuso' : ''} ${disabled ? 'is-disabled' : ''} selimp-st-${st.toLowerCase()}`;
      const onClick = disabled
        ? `onclick="event.preventDefault();PrintHub._avisar('${escAttr(i.reason || stTxt)}')"`
        : `onclick="PrintHub.elegir('${escAttr(i.printNodeId)}')"`;
      const reasonLine = (disabled && i.reason) ? `<div class="selimp-reason">${escHtml(i.reason)}</div>` : '';
      const compInfo = i.computer
        ? ` · 💻 ${escHtml(i.computer)}${i.computerState && i.computerState !== 'connected' ? ' <span class="selimp-pc-off">(desconectada)</span>' : ''}`
        : '';
      const tipoIco = i.tipo === 'ADHESIVO' ? '🏷 ' : i.tipo === 'ZPL' ? '📄 ' : '🧾 ';
      const subMeta = (i.tipo ? tipoIco + i.tipo : '') + compInfo + enUso;
      return `
      <button class="${cardCls}" ${onClick} title="${escAttr(i.reason || stTxt)}">
        <div class="selimp-ico">${ico}</div>
        <div class="selimp-info">
          <div class="selimp-nombre">
            ${escHtml(i.nombre)}
            <span class="selimp-state-pill selimp-state-${st.toLowerCase()}">${stIcon} ${stTxt}</span>
          </div>
          <div class="selimp-meta">${subMeta}</div>
          ${reasonLine}
        </div>
      </button>`;
    };
    Object.keys(zonas)
      .sort((a, b) => zonas[a].label.localeCompare(zonas[b].label))
      .forEach(zk => {
        const z = zonas[zk];
        html += `<div class="selimp-zona-header">📍 ${escHtml(z.label)}</div>`;
        Object.keys(z.estaciones).sort((a, b) => z.estaciones[a].label.localeCompare(z.estaciones[b].label)).forEach(ek => {
          const e = z.estaciones[ek];
          html += `<div class="selimp-est-header">🏷 ${escHtml(e.label)}</div>`;
          html += e.items.map(_renderCard).join('');
        });
        if (z.sinEst.length) {
          html += z.sinEst.map(_renderCard).join('');
        }
      });
    lst.innerHTML = html;
  }

  // [v2.13.36] Aviso cuando user toca una impresora no disponible
  function _avisar(razon) {
    if (typeof toast === 'function') toast('🖨 ' + (razon || 'No disponible'), 'error', 5000);
  }

  function elegir(printerIdOverride) {
    if (!_pendiente) return;
    const { apiMethod, params, resolve } = _pendiente;
    _cerrar();
    const finalParams = printerIdOverride
      ? Object.assign({}, params, { printerIdOverride: printerIdOverride })
      : params;
    if (typeof toast === 'function') toast('Enviando a impresora...', 'ok', 2000);
    try {
      const p = API[apiMethod](finalParams);
      resolve(p);
    } catch(e) { resolve(Promise.reject(e)); }
  }

  function cancelar() {
    const pend = _pendiente;
    _cerrar();
    if (pend) pend.resolve(null);
  }

  function _cerrar() {
    const ov = document.getElementById('modalSelImpresora');
    if (ov) {
      ov.classList.remove('is-open');
      setTimeout(() => { ov.style.display = 'none'; }, 260);
    }
    _pendiente = null;
  }

  return { imprimir, elegir, cancelar, _avisar, toggleZona, toggleTipo };
})();
window.PrintHub = PrintHub;

// ════════════════════════════════════════════════════════════════════
// [v2.13.39] Keypad táctil DELEGADO para inputs de clave admin
// Un solo listener global maneja TODOS los keypads (sheetEnvAuth +
// sheetReabrirAdmin + futuros). Cada keypad declara su target con
// data-target="<inputId>". Las teclas con data-d, data-del, data-ok.
// ════════════════════════════════════════════════════════════════════
document.addEventListener('click', function(e) {
  const btn = e.target.closest('.wh-key');
  if (!btn) return;
  const keypad = btn.closest('.wh-keypad');
  if (!keypad) return;
  const targetId = keypad.dataset.target;
  const input = targetId ? document.getElementById(targetId) : null;
  if (!input) return;
  e.preventDefault();
  try { if (navigator.vibrate) navigator.vibrate(15); } catch(_){}
  if (btn.hasAttribute('data-del')) {
    input.value = String(input.value || '').slice(0, -1);
    return;
  }
  if (btn.hasAttribute('data-ok-reabrir')) {
    // Confirmación específica reabrir guía
    if (typeof App !== 'undefined' && App.confirmarReabrirAdmin) {
      App.confirmarReabrirAdmin();
    }
    return;
  }
  if (btn.hasAttribute('data-ok')) {
    // Confirmación envasado auth
    if (typeof EnvasadosView !== 'undefined' && EnvasadosView.validarAuth) {
      EnvasadosView.validarAuth();
    }
    return;
  }
  const d = btn.dataset.d;
  if (d && (input.value || '').length < (parseInt(input.maxLength, 10) || 8)) {
    input.value = (input.value || '') + String(d);
    // Auto-submit al alcanzar maxlength (consistente con MOS AdminAuthModal)
    const max = parseInt(input.maxLength, 10) || 8;
    if (input.value.length === max) {
      setTimeout(() => {
        if (targetId === 'reabrirAdminInput' && typeof App !== 'undefined' && App.confirmarReabrirAdmin) App.confirmarReabrirAdmin();
        else if (targetId === 'envAuthClave' && typeof EnvasadosView !== 'undefined' && EnvasadosView.validarAuth) EnvasadosView.validarAuth();
      }, 120);
    }
  }
});

// ════════════════════════════════════════════════
// DESPACHO RÁPIDO
// ════════════════════════════════════════════════
const DespachoView = (() => {
  const CART_KEY   = 'wh_despacho_cart';
  const ZONA_KEY   = 'wh_despacho_zona';
  const HIST_KEY   = 'wh_despacho_hist';
  const PICKUP_KEY = 'wh_despacho_pickup_activo';
  let _cart = [];
  let _tipoSalida = 'SALIDA_ZONA';

  function _saveCart() { localStorage.setItem(CART_KEY, JSON.stringify(_cart)); }
  function _loadCart() {
    try { return JSON.parse(localStorage.getItem(CART_KEY) || '[]'); } catch { return []; }
  }
  // Pickup activo persistente — sobrevive refresh del browser
  function _savePickup() {
    if (_pickupActivo) localStorage.setItem(PICKUP_KEY, JSON.stringify(_pickupActivo));
    else               localStorage.removeItem(PICKUP_KEY);
  }
  function _loadPickup() {
    try { return JSON.parse(localStorage.getItem(PICKUP_KEY) || 'null'); } catch { return null; }
  }
  function _clearPickup() { localStorage.removeItem(PICKUP_KEY); }
  function _saveZona(id) { if (id) localStorage.setItem(ZONA_KEY, id); }
  function _loadZona()   { return localStorage.getItem(ZONA_KEY) || ''; }
  function _saveHist(entry) {
    try {
      const h = JSON.parse(localStorage.getItem(HIST_KEY) || '[]');
      h.unshift(entry);
      localStorage.setItem(HIST_KEY, JSON.stringify(h.slice(0, 5)));
    } catch {}
  }
  function _loadHist() {
    try { return JSON.parse(localStorage.getItem(HIST_KEY) || '[]'); } catch { return []; }
  }
  function _fmtHistTs(ts) {
    if (!ts) return '';
    const d = new Date(ts), now = new Date();
    if (d.toDateString() === now.toDateString())
      return d.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' });
    return d.toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit' });
  }
  function _histDestino(h) {
    if (h.tipo === 'SALIDA_ZONA')       return { label: 'Zona', detail: h.nombreZona || h.idZona || '—', col: '#38bdf8' };
    if (h.tipo === 'SALIDA_JEFATURA')   return { label: 'Jefatura', detail: h.nota || 'Sin comentario', col: '#a78bfa' };
    if (h.tipo === 'SALIDA_DEVOLUCION') return { label: 'Devolución', detail: h.nota || 'Sin comentario', col: '#fbbf24' };
    return { label: h.tipo || '—', detail: h.nota || '', col: '#94a3b8' };
  }

  function _renderHist() {
    const el = document.getElementById('despHistorial');
    if (!el) return;
    const hist = _loadHist();
    if (!hist.length) { el.style.display = 'none'; return; }
    el.style.display = 'block';
    el.innerHTML = `
      <p style="font-size:.68em;font-weight:800;color:#334155;margin-bottom:5px;letter-spacing:.06em;text-transform:uppercase">Últimos despachos</p>
      ${hist.map((h, idx) => {
        const dest = _histDestino(h);
        return `<div onclick="DespachoView.verHistDetalle(${idx})"
                     style="display:flex;align-items:center;gap:10px;padding:9px 4px;
                            border-bottom:1px solid #1e293b;cursor:pointer;
                            -webkit-tap-highlight-color:transparent"
                     ontouchstart="this.style.background='#0c1e30'" ontouchend="this.style.background='transparent'">
          <div style="flex:1;min-width:0">
            <p style="font-size:.74em;font-weight:800;color:${dest.col};letter-spacing:.02em">
              ${escHtml(dest.label)} · <span style="color:#cbd5e1;font-weight:600">${escHtml(dest.detail)}</span>
            </p>
            <p style="font-size:.62em;color:#475569;margin-top:1px">${_fmtHistTs(h.ts)}${h.ok ? '' : ' · <span style="color:#f87171">error</span>'}</p>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <p style="font-size:.92em;font-weight:900;color:#f1f5f9;line-height:1">${h.n}</p>
            <p style="font-size:.55em;color:#475569;letter-spacing:.05em;text-transform:uppercase">prod</p>
          </div>
          <span style="font-size:.7em;color:${h.ok ? '#34d399' : '#f87171'};flex-shrink:0">${h.ok ? '✓' : '✗'}</span>
        </div>`;
      }).join('')}`;
  }

  function verHistDetalle(idx) {
    const hist = _loadHist();
    const h = hist[idx];
    if (!h) return;
    const dest = _histDestino(h);
    const titEl = document.getElementById('histDetalleTitulo');
    const subEl = document.getElementById('histDetalleSub');
    const lstEl = document.getElementById('histDetalleLista');
    const ftEl  = document.getElementById('histDetalleFooter');
    if (!titEl || !lstEl) return;
    titEl.innerHTML = `<span style="color:${dest.col}">${escHtml(dest.label)}</span> · <span style="color:#cbd5e1">${escHtml(dest.detail)}</span>`;
    subEl.textContent = _fmtHistTs(h.ts) + (h.ok ? '' : ' · Error') + (h.idGuia && h.idGuia !== '—' ? '' : '');
    if (Array.isArray(h.items) && h.items.length) {
      lstEl.innerHTML = h.items.map(it => {
        const qtyFmt = fmt(it.cantidad, Number.isInteger(parseFloat(it.cantidad)) ? 0 : 2);
        return `<div style="display:flex;align-items:center;gap:10px;padding:9px 4px;border-bottom:1px solid #1e293b">
          <div style="flex:1;min-width:0">
            <p style="font-size:.82em;font-weight:700;color:#f1f5f9;
                      white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(it.descripcion || it.codigoBarra)}</p>
            <p style="font-size:.66em;color:#64748b;font-family:monospace">${escHtml(it.codigoBarra)}</p>
          </div>
          <span style="font-size:.92em;font-weight:900;color:#38bdf8;flex-shrink:0">${qtyFmt}</span>
        </div>`;
      }).join('');
    } else {
      lstEl.innerHTML = `
        <div style="text-align:center;padding:30px 20px;color:#475569">
          <p style="font-size:.85em;margin-bottom:6px">📦 ${h.n} producto${h.n !== 1 ? 's' : ''} despachado${h.n !== 1 ? 's' : ''}</p>
          <p style="font-size:.72em;line-height:1.5">El detalle de productos no está disponible<br>
            <span style="color:#334155">(despacho registrado antes de v1.5.35)</span>
          </p>
        </div>`;
    }
    const totalUds = (h.items || []).reduce((s, it) => s + (parseFloat(it.cantidad) || 0), 0);
    if (ftEl) ftEl.textContent = `${h.n} producto${h.n !== 1 ? 's' : ''} · ${fmt(totalUds, 2)} unidades`;
    abrirSheet('sheetHistDetalle');
  }

  // ── Estado sesión cámara despacho ────────────────────────────
  let _despLastHistory = [];
  let _despTorchOn     = false;
  let _despStatusTimer = null;

  function cargar() {
    _cart = _loadCart();
    // Rehidratar pickup activo desde localStorage (sobrevive refresh del browser
    // y navegación entre módulos — el operador puede ir a otra vista y volver
    // sin perder el progreso del despacho).
    const pSaved = _loadPickup();
    if (pSaved) _pickupActivo = pSaved;
    // [v2.13.8] Rehidratar lista sombra si la había
    try { _lsRehidratar(); } catch(_){}
    // [v2.13.15] Sync con panel compartido — refresca y arranca polling
    try { _lsStartPanelSync(); } catch(_){}
    try { _lsStartProgresoSync(); } catch(_){}
    // [v2.13.12] Forzar ocultar flotante al entrar a despacho —
    // el usuario reportó que persistía aunque estuviera en la vista despacho
    try {
      const flot = document.getElementById('despFlotante');
      if (flot) { flot.style.display = 'none'; flot.onclick = null; }
    } catch(_){}
    _renderCart();
    _updateFooter();
    _renderHist();
    _renderPickupActivoBanner();
    _renderPickupChecklistInline();
    _renderExtrasSection();
    _updateGenerarBtn();
    badgeUpdate();
    // Sync con backend en background — detecta si fue cerrado por otro/timeout
    if (pSaved) _sincronizarPickupActivo();
  }

  // ── Render del checklist inline (en view-despacho, debajo de cám/scan) ──
  // Reemplaza el sheet modal — todo va en una sola vista.
  function _renderPickupChecklistInline() {
    const cont = document.getElementById('despPickupChecklistInline');
    if (!cont) return;
    if (!_pickupActivo) { cont.style.display = 'none'; cont.innerHTML = ''; return; }
    const items = _pickupActivo.items || [];
    if (!items.length) { cont.style.display = 'none'; return; }

    // Stock cache indexado por TODOS los códigos posibles del row
    // (codigoProducto + idProducto). Así un row guardado por EAN físico
    // también es accesible si buscamos por idProducto del catálogo.
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => {
      const codP = s.codigoProducto;
      const idP  = s.idProducto;
      if (codP) stockMap[String(codP)] = s;
      if (idP)  stockMap[String(idP)]  = s;
    });
    const productos = (App.getProductosMaestro && App.getProductosMaestro()) || [];

    // Helper: SUMAR stock de todos los códigos del item (canónico + equivalentes).
    // En tabla STOCK puede haber rows separados por codigoBarra distinto pero
    // pertenecen al mismo skuBase. El stock total disponible es la suma.
    function _buscarStock(item) {
      let total = 0;
      let count = 0;
      const visited = new Set();
      const tryKey = (k) => {
        if (!k) return;
        const key = String(k);
        if (visited.has(key)) return;
        visited.add(key);
        const r = stockMap[key];
        if (r) { total += parseFloat(r.cantidadDisponible || 0); count++; }
      };
      if (Array.isArray(item.codigosOriginales)) {
        item.codigosOriginales.forEach(c => tryKey(c));
      }
      // Fallback skuBase si los codigos no encontraron nada
      if (count === 0) tryKey(item.skuBase);
      return count > 0 ? { cantidadDisponible: total, codigosEncontrados: count } : null;
    }

    // Orden: pendientes primero, completados al final
    const sorted = items.slice().sort((a, b) => {
      const aP = (parseFloat(a.despachado)||0) < (parseFloat(a.solicitado)||0) ? 0 : 1;
      const bP = (parseFloat(b.despachado)||0) < (parseFloat(b.solicitado)||0) ? 0 : 1;
      if (aP !== bP) return aP - bP;
      return String(a.nombre).localeCompare(String(b.nombre));
    });

    cont.style.display = 'block';
    cont.innerHTML = sorted.map(it => {
      const sol  = parseFloat(it.solicitado) || 0;
      const desp = parseFloat(it.despachado) || 0;
      const pct  = sol > 0 ? Math.min(100, Math.round((desp / sol) * 100)) : 0;
      const completo  = desp >= sol && sol > 0;
      const enProg    = desp > 0 && desp < sol;
      const cls       = completo ? 'is-completo' : (enProg ? 'is-progreso' : '');
      const stockRow  = _buscarStock(it);
      const stockD    = stockRow ? parseFloat(stockRow.cantidadDisponible || 0) : 0;
      const stockKnown= !!stockRow;
      const equivCount= Array.isArray(it.codigosOriginales) ? it.codigosOriginales.length : 0;
      const equivTxt  = equivCount > 1 ? ` · ${equivCount} códigos` : '';
      const icon      = completo ? '✓' : (enProg ? '⏳' : '📦');
      const prod = productos.find(p => String(p.idProducto) === String(it.skuBase));
      const esKg = _esProductoPeso(prod);
      const unidadLbl = esKg ? String(prod.unidad || 'kg').toLowerCase() : '';
      const pendiente = sol - desp;
      let stockBadge = '';
      if (!stockKnown) {
        stockBadge = '<span style="font-size:.62em;color:#94a3b8;background:rgba(71,85,105,.25);padding:1px 6px;border-radius:6px" title="No hay stock cacheado para este producto en WH">sin info</span>';
      } else if (stockD <= 0) {
        stockBadge = '<span style="font-size:.62em;color:#fca5a5;background:rgba(220,38,38,.18);border:1px solid rgba(239,68,68,.4);padding:1px 6px;border-radius:6px;font-weight:800">⚠ stock 0</span>';
      } else if (stockD < pendiente) {
        stockBadge = `<span style="font-size:.62em;color:#fca5a5;background:rgba(220,38,38,.18);border:1px solid rgba(239,68,68,.4);padding:1px 6px;border-radius:6px;font-weight:800">⚠ stock ${fmt(stockD,1)}</span>`;
      } else {
        stockBadge = `<span style="font-size:.62em;color:#86efac;background:rgba(16,185,129,.15);padding:1px 6px;border-radius:6px;font-weight:700">stock ${fmt(stockD,1)}</span>`;
      }
      const kgBadge = esKg
        ? `<span style="font-size:.6em;color:#fbbf24;background:rgba(245,158,11,.15);padding:1px 5px;border-radius:6px;margin-left:4px;font-weight:800">⚖ ${unidadLbl}</span>`
        : '';
      return `
        <div class="pkck-card ${cls}" data-sku="${escAttr(it.skuBase)}">
          <div class="pkck-row">
            <div class="pkck-icon">${icon}</div>
            <div class="flex-1 min-w-0">
              <p class="pkck-name truncate">${escHtml(it.nombre || it.skuBase)}${kgBadge}</p>
              <p class="pkck-meta truncate">${escHtml(it.skuBase)}${equivTxt} · ${stockBadge}</p>
            </div>
            <div class="pkck-qty-wrap">
              <p><span class="pkck-qty">${esKg ? fmt(desp,3) : desp}</span><span class="pkck-qty-sol"> / ${esKg ? fmt(sol,3) : sol}${esKg ? ' '+unidadLbl : ''}</span></p>
              <p class="text-[10px] text-slate-500 mt-0.5">${pct}%</p>
            </div>
          </div>
          <div class="pkck-bar-wrap">
            <div class="pkck-bar-fill" style="width:${pct}%"></div>
            ${enProg ? '<div class="pkck-bar-shimmer"></div>' : ''}
          </div>
          <div class="pkck-check-overlay">✓</div>
        </div>`;
    }).join('');
  }

  // Render sección extras (items escaneados fuera del pickup activo)
  function _renderExtrasSection() {
    const sec = document.getElementById('despExtrasSection');
    const list = document.getElementById('despExtrasList');
    if (!sec || !list) return;
    if (!_pickupActivo) { sec.style.display = 'none'; return; }
    const extras = _cart.filter(c => c._extraPickup);
    if (!extras.length) { sec.style.display = 'none'; return; }
    sec.style.display = 'block';
    list.innerHTML = extras.map(c => `
      <div class="card-sm flex items-center gap-2" style="padding:7px 10px">
        <div class="flex-1 min-w-0">
          <p class="text-sm font-semibold truncate">${escHtml(c.descripcion)}</p>
          <p class="text-[10px] text-slate-500 font-mono">${escHtml(c.codigoBarra)}</p>
        </div>
        <span class="text-sm font-bold text-amber-300">×${c.cantidad}</span>
        <button onclick="DespachoView.quitarItem('${escAttr(c.codigoBarra)}')"
                style="padding:4px 8px;border-radius:7px;border:none;cursor:pointer;
                       background:rgba(239,68,68,.18);color:#fca5a5;font-size:.7em">✕</button>
      </div>
    `).join('');
  }

  // Actualiza el botón único "GENERAR GUÍA" según contexto
  function _updateGenerarBtn() {
    const btn  = document.getElementById('despBtnGenerarGuia');
    const txt  = document.getElementById('despBtnGenerarTxt');
    const hint = document.getElementById('despBtnGenerarHint');
    if (!btn) return;
    if (_pickupActivo) {
      const items = _pickupActivo.items || [];
      const totalDesp = items.reduce((s,it) => s + (parseFloat(it.despachado)||0), 0);
      const extras = _cart.filter(c => c._extraPickup);
      const haySomething = totalDesp > 0 || extras.length > 0;
      const completo = items.length > 0 && items.every(it =>
        (parseFloat(it.despachado)||0) >= (parseFloat(it.solicitado)||0));
      btn.disabled = !haySomething;
      if (txt) {
        if (!haySomething) txt.textContent = 'GENERAR GUÍA · escanea primero';
        else if (completo) txt.textContent = 'GENERAR GUÍA · ✓ completo';
        else {
          const falt = items.filter(it => (parseFloat(it.despachado)||0) < (parseFloat(it.solicitado)||0)).length;
          txt.textContent = `GENERAR GUÍA · faltan ${falt} item${falt!==1?'s':''}`;
        }
      }
      if (hint) hint.textContent = haySomething
        ? (completo ? '¡Todo despachado! · pulsa para emitir guía' : 'Los faltantes irán en la observación')
        : 'Escanea cada producto del pickup para registrarlo';
    } else {
      const n = _cart.length;
      btn.disabled = n === 0;
      if (txt) txt.textContent = n === 0 ? 'GENERAR GUÍA · escanea primero' : `GENERAR GUÍA · ${n} producto${n!==1?'s':''}`;
      if (hint) hint.textContent = n === 0 ? 'Abre la cámara o el scan para empezar' : 'Pulsa para emitir guía de salida';
    }
  }

  // Botón único "Generar guía" — decide ruta según haya pickup activo o no
  function generarGuia() {
    if (_pickupActivo) {
      const items = _pickupActivo.items || [];
      const completo = items.length > 0 && items.every(it =>
        (parseFloat(it.despachado)||0) >= (parseFloat(it.solicitado)||0));
      cerrarDespachoPickup(!completo);
      return;
    }
    finalizar(); // flujo legacy carrito → sheet de finalizar
  }

  function pauseCamera() { Scanner.stop(); }

  // ── Cámara TELÓN inline ───────────────────────────────────
  // Modo nuevo: el header de view-despacho se colapsa con animación,
  // un panel cámara aparece desde arriba (slide-down), y la lista del pickup
  // de abajo permanece intacta. El operador escanea sin perder de vista
  // su checklist.
  function abrirDespCamara() {
    _despLastHistory = [];
    _despTorchOn = false;
    // Cerrar SCAN inline si estaba abierto
    const scanPanel = document.getElementById('despScanInlinePanel');
    if (scanPanel && scanPanel.style.display !== 'none') {
      scanPanel.style.display = 'none';
    }
    // Telón: colapsar header con animación
    const header = document.getElementById('despHeaderCollapsible');
    if (header) header.classList.add('is-collapsed');
    // Panel cámara aparece desde arriba (slide-down)
    const camPanel = document.getElementById('despCamInlinePanel');
    if (camPanel) {
      camPanel.classList.remove('is-closing');
      camPanel.style.display = 'block';
    }
    const statusEl = document.getElementById('despCamInlineStatus');
    if (statusEl) statusEl.textContent = 'Iniciando cámara…';
    // Sonido + vibración suaves
    try { SoundFX.tick && SoundFX.tick(); SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);

    // Iniciar scanner sobre el video del panel inline
    Scanner.start('despScanVideoInline', _onDespResult, err => {
      toast('Error cámara: ' + err, 'danger');
      cerrarDespCamara();
    }, { continuous: true, cooldown: 1500 });

    setTimeout(async () => {
      if (!Scanner.isActive()) return;
      if (statusEl) statusEl.textContent = 'Apunta a un código…';
      const ok = await Scanner.toggleTorch(true);
      if (ok) { _despTorchOn = true; }
      const torchBtn = document.getElementById('despCamInlineTorch');
      if (torchBtn && _despTorchOn) torchBtn.style.background = 'rgba(251,191,36,.85)';
    }, 700);
  }

  function cerrarDespCamara() {
    clearTimeout(_despStatusTimer);
    _despTorchOn = false;
    try { Scanner.stop(); } catch(_){}
    const camPanel = document.getElementById('despCamInlinePanel');
    if (camPanel) {
      camPanel.classList.add('is-closing');
      // Esperar fin de animación antes de ocultar
      setTimeout(() => {
        camPanel.style.display = 'none';
        camPanel.classList.remove('is-closing');
      }, 350);
    }
    // Telón: subir el header de vuelta
    const header = document.getElementById('despHeaderCollapsible');
    if (header) header.classList.remove('is-collapsed');
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(10);
    _renderCart();
    _updateFooter();
    _renderExtrasSection();
    _updateGenerarBtn();
    badgeUpdate();
  }

  function cerrarDespYFinalizar() {
    cerrarDespCamara();
    setTimeout(() => finalizar(), 200);
  }

  // ── Torch y zoom ─────────────────────────────────────────────
  async function toggleDespTorch() {
    _despTorchOn = !_despTorchOn;
    const ok = await Scanner.toggleTorch(_despTorchOn);
    const btn = document.getElementById('despScanTorchBtn');
    if (!ok) { _despTorchOn = false; toast('Linterna no disponible', 'warn', 2000); }
    if (btn) btn.style.background = (_despTorchOn && ok) ? 'rgba(251,191,36,.9)' : 'rgba(255,255,255,.15)';
  }

  function despSetZoom(val) {
    const v = parseFloat(val);
    Scanner.setZoom(v);
    const lbl = document.getElementById('despScanZoomLabel');
    if (lbl) lbl.textContent = v.toFixed(1) + '×';
  }

  // ── Barra de estado ──────────────────────────────────────────
  function _setDespStatus(type, text) {
    const bar = document.getElementById('despScanStatusBar');
    if (!bar) return;
    clearTimeout(_despStatusTimer);
    if (type === 'ready') {
      bar.innerHTML = '<span style="color:#334155;font-size:.75em">— listo para escanear —</span>';
      bar.style.background = '#0f172a';
      return;
    }
    const cfgs = {
      ok:         { bg: '#022c22', col: '#34d399', icon: '✓', dur: 2500 },
      sobrestock: { bg: '#2c0d00', col: '#fb923c', icon: '⚠', dur: 0   },
      prefijo:    { bg: '#2c1a00', col: '#fb923c', icon: '↕', dur: 0   },
      no_existe:  { bg: '#2d0a0a', col: '#f87171', icon: '⚠', dur: 3000 }
    };
    const c = cfgs[type] || cfgs.ok;
    bar.style.background = c.bg;
    bar.innerHTML = `
      <span style="color:${c.col};font-size:.82em;font-weight:700;flex-shrink:0">${c.icon}</span>
      <span style="color:${c.col};font-size:.76em;flex:1;white-space:nowrap;overflow:hidden;
                   text-overflow:ellipsis;margin-left:6px">${escHtml(text)}</span>`;
    if (c.dur > 0) _despStatusTimer = setTimeout(() => _setDespStatus('ready'), c.dur);
  }

  // ── Escáner GLOBAL (Parte B) ─────────────────────────────────
  // ¿Hay un despacho/pickup en curso? Solo entonces el listener global
  // de teclado enruta los escaneos hechos desde OTROS módulos (ej. el
  // operador está en el catálogo de Productos viendo el "membrete
  // digital" y apunta el escáner a la pantalla).
  function hayDespachoActivo() {
    // [v2.13.10] Lista sombra también cuenta como "actividad" para que el
    // flotante se vea desde otras vistas y el operador pueda volver al despacho.
    return !!_pickupActivo
        || (Array.isArray(_cart) && _cart.length > 0)
        || !!(_listaSombra && _listaSombra.items && _listaSombra.items.length);
  }
  // Procesa un código escaneado desde cualquier módulo. Reusa _onDespResult
  // (busca candidato → suma a pickup o cart). Como el banner del pickup
  // vive dentro de la vista despacho y no se ve desde otros módulos,
  // damos feedback con un toast prominente + sonido.
  function procesarScanGlobal(cod) {
    if (!hayDespachoActivo()) return false;
    const codStr = normCb(cod);
    const candidatos = _buscarDespCandidatos(codStr);
    if (!candidatos.length) {
      try { SoundFX.warn(); vibrate([60, 30, 60]); } catch(_){}
      toast('⚠ ' + codStr + ' · no existe en catálogo', 'warn', 3500);
      return true;
    }
    // Reusa el flujo normal (maneja pickup, cart, sobrestock, picker prefijo)
    _onDespResult(cod);
    // Feedback visible cross-módulo: el operador está en Productos, no ve
    // el banner del despacho. El toast le confirma qué entró.
    if (candidatos[0]._exacto) {
      const prod = candidatos[0];
      const nombre = prod.descripcion || codStr;
      let detalle = '';
      try {
        if (_pickupActivo) {
          const it = _matchPickupItem(prod, String(prod._scannedCb || prod.codigoBarra || ''));
          if (it) detalle = ' · ' + (parseFloat(it.despachado) || 0) + '/' + (parseFloat(it.solicitado) || 0);
        }
      } catch(_){}
      toast('✓ ' + nombre + detalle + ' — sumado al despacho', 'ok', 3000);
    }
    _renderDespFlotante();
    return true;
  }

  // ── FLOTANTE del despacho activo (control remoto cross-módulo) ──
  // Resuelve el "producto activo": del pickup (_pickupItemActivo) o del
  // despacho rápido (_despItemActivo). Retorna lo que el flotante necesita.
  function _despProductoActivo() {
    const productos = (App.getProductosMaestro && App.getProductosMaestro()) || [];
    if (_pickupActivo && _pickupItemActivo) {
      const it = (_pickupActivo.items || []).find(x => String(x.skuBase) === String(_pickupItemActivo));
      if (it) {
        const prod = productos.find(p => String(p.idProducto) === String(it.skuBase));
        const esG = _esProductoPeso(prod);
        return {
          tipo: 'pickup', id: it.skuBase, nombre: it.nombre || it.skuBase,
          cantidad: parseFloat(it.despachado) || 0, solicitado: parseFloat(it.solicitado) || 0,
          esGranel: esG, unidad: esG ? String((prod && prod.unidad) || 'kg').toLowerCase() : ''
        };
      }
    }
    if (!_pickupActivo && _despItemActivo) {
      const ci = _cart.find(c => c.codigoBarra === _despItemActivo);
      if (ci) {
        const prod = productos.find(p =>
          String(p.codigoBarra) === String(ci.codigoBarra) ||
          String(p.idProducto) === String(ci.codigoBarra));
        const esG = _esProductoPeso(prod) || _esProductoPeso({ unidad: ci.unidad });
        return {
          tipo: 'cart', id: ci.codigoBarra, nombre: ci.descripcion || ci.codigoBarra,
          cantidad: parseFloat(ci.cantidad) || 0, solicitado: null,
          esGranel: esG, unidad: esG ? String((prod && prod.unidad) || ci.unidad || 'kg').toLowerCase() : ''
        };
      }
    }
    return null;
  }

  // Muestra/oculta y re-renderiza el flotante. Visible SOLO si hay despacho
  // activo Y NO estás en la vista despacho (ahí ya tenés el control completo).
  function _renderDespFlotante() {
    const flot = document.getElementById('despFlotante');
    if (!flot) return;
    // [v2.13.22] Chequeo robusto: tanto currentView como visibilidad real del DOM.
    // Al refrescar la página dentro de despacho, currentView sigue en 'dashboard'
    // (default) hasta que se navegue, y el flotante aparecía encima del banner.
    const viewByApp = !!(window.App && App.getView && App.getView() === 'despacho');
    const viewDom   = document.getElementById('view-despacho');
    const viewByDom = !!(viewDom && viewDom.offsetParent !== null);
    const enDespacho = viewByApp || viewByDom;
    const activo = _despProductoActivo();
    // [v2.13.10] Mostrar flotante también si hay lista sombra (aunque cart vacío)
    if (!hayDespachoActivo() || enDespacho) {
      flot.style.display = 'none';
      return;
    }
    if (!activo) {
      // Sin producto activo pero hay lista sombra → mostrar resumen de sombra
      if (_listaSombra && _listaSombra.items && _listaSombra.items.length) {
        const total = _listaSombra.items.length;
        const completos = _listaSombra.items.filter(i => (i.cantidadEscaneada || 0) >= i.cantidad).length;
        const nom   = document.getElementById('despflotNombre');
        const meta  = document.getElementById('despflotMeta');
        const ctrls = document.getElementById('despflotCtrls');
        if (nom) nom.textContent = '📋 Lista sombra activa';
        if (meta) meta.textContent = `${completos}/${total} items · toca para volver al despacho`;
        if (ctrls) ctrls.innerHTML = '<span style="font-size:1.3em;opacity:.7">→</span>';
        // [v2.13.11] TODO el flotante es clickable en modo sombra (no solo el goto)
        const goFn = () => { try { App.nav('despacho'); } catch(_){} };
        flot.onclick = goFn;
        const goto = document.getElementById('despflotGoto');
        if (goto) goto.onclick = (ev) => { ev.stopPropagation(); goFn(); };
        flot.style.display = 'block';
        flot.classList.add('is-sombra');
        return;
      }
      flot.style.display = 'none';
      return;
    }
    flot.classList.remove('is-sombra');
    flot.onclick = null;
    const nom   = document.getElementById('despflotNombre');
    const meta  = document.getElementById('despflotMeta');
    const ctrls = document.getElementById('despflotCtrls');
    if (nom) nom.textContent = activo.nombre;
    if (meta) {
      meta.textContent = activo.tipo === 'pickup'
        ? 'Pickup · ' + fmt(activo.cantidad, activo.esGranel ? 3 : 0) + '/' +
          fmt(activo.solicitado, activo.esGranel ? 3 : 0) + (activo.esGranel ? ' ' + activo.unidad : '')
        : 'Despacho rápido' + (activo.esGranel ? ' · ' + activo.unidad : '');
    }
    if (ctrls) {
      if (activo.esGranel) {
        ctrls.innerHTML =
          '<input type="number" inputmode="decimal" step="0.001" min="0" class="despflot-granel" ' +
          'value="' + fmt(activo.cantidad, 3) + '" ' +
          'onchange="DespachoView.flotSetGranel(this.value)" ' +
          'onkeydown="if(event.key===\'Enter\')this.blur()">' +
          '<span class="despflot-unit">' + activo.unidad + '</span>';
      } else {
        ctrls.innerHTML =
          '<button class="despflot-btn despflot-btn-menos" onclick="DespachoView.flotMenos()"' +
          (activo.cantidad <= 0 ? ' disabled' : '') + '>−</button>' +
          '<span class="despflot-val">' + fmt(activo.cantidad, 0) + '</span>' +
          '<button class="despflot-btn despflot-btn-mas" onclick="DespachoView.flotMas()">+</button>';
      }
    }
    const goto = document.getElementById('despflotGoto');
    if (goto) goto.onclick = () => { try { App.nav('despacho'); } catch(_){} };
    flot.style.display = 'block';
    flot.classList.remove('is-flash');
    void flot.offsetWidth; // reflow → reinicia la animación de flash
    flot.classList.add('is-flash');
  }

  function flotMas() {
    const a = _despProductoActivo();
    if (!a) return;
    if (a.tipo === 'pickup') _pkckMas(a.id);
    else _adjustQty(a.id, q => q + 1);
    _renderDespFlotante();
  }
  function flotMenos() {
    const a = _despProductoActivo();
    if (!a) return;
    if (a.tipo === 'pickup') _pkckMenos(a.id);
    else _adjustQty(a.id, q => Math.max(0, q - 1));
    _renderDespFlotante();
  }
  function flotSetGranel(val) {
    const a = _despProductoActivo();
    if (!a) return;
    const qty = parseFloat(String(val).replace(',', '.'));
    if (isNaN(qty) || qty < 0) return;
    if (a.tipo === 'pickup') {
      _pkckSetGranel(a.id, qty);
    } else {
      const idx = _cart.findIndex(c => c.codigoBarra === a.id);
      if (idx >= 0) {
        _cart[idx].cantidad = qty;
        _saveCart(); _renderCart(); _updateFooter(); badgeUpdate();
      }
    }
    _renderDespFlotante();
  }

  // ── Callback scanner ─────────────────────────────────────────
  function _onDespResult(cod) {
    const codStr = normCb(cod);
    if (!codStr) return;
    const picker = document.getElementById('despCamPicker');
    if (picker?.style.display === 'flex') return;
    const candidatos = _buscarDespCandidatos(codStr);
    if (!candidatos.length) {
      _setDespStatus('no_existe', codStr + ' · no existe en catálogo');
      SoundFX.warn(); vibrate([60, 30, 60]);
      return;
    }
    if (candidatos[0]._exacto) {
      const prod = candidatos[0];
      _agregarDespDirecto(prod); // maneja sobrestock internamente
      const cb   = String(prod._scannedCb || prod.codigoBarra || '');
      // Marcar producto activo para el flotante cross-módulo. Si hay
      // pickup, _intentarSumarAPickup ya seteó _pickupItemActivo; si es
      // despacho rápido (cart), lo seteamos acá.
      if (!_pickupActivo) _despItemActivo = cb;
      _renderDespFlotante();
      const item = _cart.find(c => c.codigoBarra === cb);
      const stockD = item?.stockDisp || 0;
      if (!item || stockD === 0 || item.cantidad <= stockD) {
        const stockTxt = stockD > 0 ? ` · Stock: ${fmt(stockD,1)}` : '';
        _setDespStatus('ok', (prod.descripcion || cb) + stockTxt);
        SoundFX.beep(); vibrate(15);
      }
      return;
    }
    _setDespStatus('prefijo', 'Prefijo · ' + candidatos.length + ' productos coinciden');
    _mostrarDespPicker(candidatos, codStr);
  }

  // ═══════════════════════════════════════════════════════════════════════
  // REGLA DE ORO WH (en piedra) — qué se acepta al ESCANEAR:
  //
  // CUALQUIER GUÍA DE SALIDA (despacho rápido, pickup, transferencias):
  //   ✓ codigoBarra del canónico (factor=1)
  //   ✓ codigoBarra de equivalencia activa (apunta al canónico)
  //   ✗ skuBase puro → ERROR "no existe en catálogo"
  //   ✗ idProducto puro → ERROR (es ID interno, no escaneable)
  //   ✗ codigoBarra de presentación (factor != 1) → ERROR
  //
  // GUÍA DE INGRESO (preingresos, recepción): además del exacto, dos casos
  // especiales:
  //   ↕ PREFIJO — escaneado es prefijo de un codigoBarra existente.
  //     Ej: catálogo tiene '12345A', operador escanea '12345' → match.
  //     Caso común: empresas que no codifican EAN bien y se les añade
  //     letra final para diferenciar variantes.
  //   + NUEVO — escaneado no es prefijo ni existe → sugerir registrar
  //     como producto nuevo (registrarPN). Solo en INGRESO, nunca en salida.
  //
  // Implementación: _buscarDespCandidatos (salida) y _buscarCandidatos
  // (ingreso/guías generales). Ambos solo buscan por codigoBarra.
  // ═══════════════════════════════════════════════════════════════════════

  // ── Búsqueda por codigoBarra — maestro + equivalencias, exacto o prefijo.
  // En salidas NUNCA se acepta nuevo: si no existe → error.
  function _buscarDespCandidatos(codStr) {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    const cNorm  = normCb(codStr);
    if (!cNorm) return [];

    // 1. Exacto en PRODUCTOS_MASTER
    const exacto = prods.find(p => String(p.codigoBarra || '').trim().toUpperCase() === cNorm);
    if (exacto) return [{ ...exacto, _exacto: true }];

    // 2. Exacto en EQUIVALENCIAS → resolver al producto base (factor=1); guardar código escaneado
    const equiv = equivs.find(e => String(e.codigoBarra || '').trim().toUpperCase() === cNorm);
    if (equiv) {
      const skuB = String(equiv.skuBase || '').trim().toUpperCase();
      const prod = prods.find(p =>
        parseFloat(p.factorConversion || 1) === 1 &&
        p.estado !== '0' && p.estado !== 0 &&
        (String(p.idProducto  || '').trim().toUpperCase() === skuB ||
         String(p.skuBase     || '').trim().toUpperCase() === skuB ||
         String(p.codigoBarra || '').trim().toUpperCase() === skuB)
      );
      if (prod) return [{ ...prod, _exacto: true, _scannedCb: cNorm }];
    }

    // 3. Prefijo en maestro + equivalencias (mín. 3 chars)
    if (cNorm.length >= 3) {
      const porMaestro = prods
        .filter(p => String(p.codigoBarra || '').trim().toUpperCase().startsWith(cNorm))
        .map(p => ({ ...p }));
      const idsYa = new Set(porMaestro.map(p => p.idProducto));
      equivs.filter(e => String(e.codigoBarra || '').trim().toUpperCase().startsWith(cNorm)).forEach(e => {
        const skuB = String(e.skuBase || '').trim().toUpperCase();
        const base = prods.find(p =>
          parseFloat(p.factorConversion || 1) === 1 &&
          p.estado !== '0' && p.estado !== 0 &&
          (String(p.idProducto  || '').trim().toUpperCase() === skuB ||
           String(p.skuBase     || '').trim().toUpperCase() === skuB ||
           String(p.codigoBarra || '').trim().toUpperCase() === skuB)
        );
        if (base && !idsYa.has(base.idProducto)) {
          porMaestro.push({ ...base, _scannedCb: String(e.codigoBarra).trim() });
          idsYa.add(base.idProducto);
        }
      });
      if (porMaestro.length) return porMaestro.slice(0, 10);
    }

    return [];
  }

  // ── Agregar al carrito (auto-suma) ───────────────────────────
  function _agregarDespDirecto(prod) {
    const cb   = String(prod._scannedCb || prod.codigoBarra || '');
    const desc = prod.descripcion || cb;
    if (!cb) return;

    // ── HOOK PICKUP: si hay pickup activo y este producto matchea con un item,
    //    sumar al despachado del pickup en vez de tratarlo como item suelto.
    if (_pickupActivo && _intentarSumarAPickup(prod, cb)) {
      _despLastHistory.push(cb);
      return; // absorbido por el pickup, no entra al carrito como extra
    }

    // ── Producto a granel SIEMPRE pide peso, esté o no dentro de pickup.
    //    Si llega aquí es porque NO matcheó con pickup activo (extra) o
    //    no hay pickup activo en absoluto. Cualquier caso → modal qty.
    if (typeof _esProductoPeso === 'function' && _esProductoPeso(prod)) {
      _abrirModalQtyGranelExtra(prod, cb);
      _despLastHistory.push(cb);
      return;
    }

    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });
    const stockD = parseFloat((stockMap[prod.idProducto] || stockMap[cb] || {}).cantidadDisponible || 0);
    const idx = _cart.findIndex(c => c.codigoBarra === cb);
    if (idx >= 0) {
      _cart[idx].cantidad = (parseFloat(_cart[idx].cantidad) || 0) + 1;
      SoundFX.beepDouble(); vibrate(12);
    } else {
      _cart.push({ codigoBarra: cb, descripcion: desc, unidad: prod.unidad || '', cantidad: 1, stockDisp: stockD, _extraPickup: !!_pickupActivo });
      // Si hay pickup activo y entró como extra → flash naranja + toast suave
      if (_pickupActivo) {
        SoundFX.warn(); vibrate([20, 30]);
        toast(`🟠 Fuera de pickup · ${desc} se agrega como extra`, 'warn', 2200);
      }
    }
    _despLastHistory.push(cb);
    _saveCart();
    _renderDespList();
    // Pulse visual en la fila recién escaneada
    requestAnimationFrame(() => {
      const row = document.querySelector(`[data-desp-cb="${CSS.escape(cb)}"]`);
      if (!row) return;
      row.classList.remove('cam-row-pulse');
      const qtyBtn = row.querySelector('.desp-qty-pulse-btn');
      if (qtyBtn) qtyBtn.classList.remove('cam-qty-pulse');
      void row.offsetWidth;
      row.classList.add('cam-row-pulse');
      if (qtyBtn) qtyBtn.classList.add('cam-qty-pulse');
      setTimeout(() => { row.classList.remove('cam-row-pulse'); if (qtyBtn) qtyBtn.classList.remove('cam-qty-pulse'); }, 400);
    });
    // Mostrar stock en barra de estado
    const item = _cart.find(c => c.codigoBarra === cb);
    const qty  = item ? parseFloat(item.cantidad) : 1;
    const stockTxt = stockD > 0 ? ` · Stock: ${fmt(stockD, 1)}` : '';
    if (stockD > 0 && qty > stockD) {
      _setDespStatus('sobrestock', desc + stockTxt + ` — pides ${fmt(qty,1)}`);
      SoundFX.warn(); vibrate([40, 20, 40]);
    }
    // (si ok o sobrestock, el status ya se muestra — el caller sigue usando _setDespStatus('ok') si quiere)
  }

  // ── Picker prefijo ───────────────────────────────────────────
  function _mostrarDespPicker(candidatos, codStr) {
    document.getElementById('despCamPickerCod').textContent = codStr;
    document.getElementById('despCamPickerList').innerHTML = candidatos.map(p => {
      const cb      = String(p._scannedCb || p.codigoBarra || '');
      const display = String(p.codigoBarra || cb);
      const cbHtml  = display.startsWith(codStr)
        ? `<strong style="color:#fbbf24">${escHtml(codStr)}</strong>${escHtml(display.slice(codStr.length))}`
        : escHtml(display);
      return `<button onclick="DespachoView.seleccionarItemDesp('${escAttr(cb)}')"
              style="width:100%;text-align:left;padding:11px 13px;border-radius:11px;
                     border:1px solid #1e293b;margin-bottom:7px;
                     background:#1e293b;display:flex;align-items:center;gap:10px;
                     cursor:pointer;-webkit-tap-highlight-color:transparent"
              ontouchstart="this.style.borderColor='#0ea5e9';this.style.background='#0c1e30'"
              ontouchend="this.style.borderColor='#1e293b';this.style.background='#1e293b'">
        <div style="flex-shrink:0;width:32px;height:32px;border-radius:8px;background:#0ea5e922;
                    display:flex;align-items:center;justify-content:center">
          <span style="font-size:.75em;color:#38bdf8;font-weight:700">CB</span>
        </div>
        <div style="flex:1;min-width:0">
          <p style="font-size:.83em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || cb))}</p>
          <p style="font-size:.69em;color:#64748b;font-family:monospace;margin-top:2px">${cbHtml}</p>
        </div>
      </button>`;
    }).join('');
    document.getElementById('despCamPicker').style.display = 'flex';
  }

  function cerrarDespPicker() {
    const picker = document.getElementById('despCamPicker');
    if (picker) picker.style.display = 'none';
    _setDespStatus('ready');
  }

  function seleccionarItemDesp(scannedCb) {
    const picker = document.getElementById('despCamPicker');
    if (picker) picker.style.display = 'none';
    const candidatos = _buscarDespCandidatos(scannedCb);
    const prod = candidatos[0];
    if (!prod) return;
    _agregarDespDirecto(prod);
    const cb     = String(prod._scannedCb || prod.codigoBarra || scannedCb);
    const item   = _cart.find(c => c.codigoBarra === cb);
    const stockD = item?.stockDisp || 0;
    if (!item || stockD === 0 || item.cantidad <= stockD) {
      const stockTxt = stockD > 0 ? ` · Stock: ${fmt(stockD,1)}` : '';
      _setDespStatus('ok', (prod.descripcion || cb) + stockTxt);
      SoundFX.beep(); vibrate(15);
    }
  }

  // ── SCAN inline TELÓN ──────────────────────────────────────
  // Mismo patrón que la cámara: header colapsa, panel scan aparece desde arriba.
  // Input READONLY — solo recibe entrada del lector de barras físico (HID)
  // capturada por listener global de keydown. Esto evita errores de tipeo manual.
  let _scanHidBuffer = '';
  let _scanHidLastTs = 0;
  let _scanHidListener = null;
  const SCAN_HID_GAP_MS = 80;       // tiempo máx entre chars de un scanner real
  const SCAN_HID_RESET_MS = 600;    // si pasa más, reinicia el buffer
  const SCAN_HID_MIN_LEN = 3;       // mínimo de chars para considerar "código válido"

  function _activarScannerHid() {
    if (_scanHidListener) return;
    _scanHidBuffer = '';
    _scanHidLastTs = 0;
    _scanHidListener = (e) => {
      // Si el target es OTRO input editable, no interceptar (evita robar foco
      // a otros campos). El input de scan es readonly, así que no entra acá.
      const tgt = e.target;
      const idTgt = tgt && tgt.id;
      const esOtroInput = tgt && (tgt.tagName === 'INPUT' || tgt.tagName === 'TEXTAREA') &&
                          idTgt !== 'despScanInlineInput' && !tgt.readOnly;
      if (esOtroInput) return;

      const now = Date.now();
      // Reset buffer si pasó mucho tiempo desde el último char (evita mezclas)
      if (now - _scanHidLastTs > SCAN_HID_RESET_MS) _scanHidBuffer = '';
      const dt = now - _scanHidLastTs;
      _scanHidLastTs = now;

      const inputBox = document.getElementById('despScanInlineInput');
      const statusEl = document.getElementById('despScanInlineStatus');

      if (e.key === 'Enter' || e.key === 'Tab') {
        if (_scanHidBuffer.length >= SCAN_HID_MIN_LEN) {
          if (inputBox) inputBox.value = _scanHidBuffer;
          submitDespScanInline();
        }
        _scanHidBuffer = '';
        e.preventDefault();
        return;
      }

      // Solo aceptar caracteres alfanuméricos y guiones (típicos de barcodes EAN/Code)
      if (!/^[a-zA-Z0-9\-_.]$/.test(e.key)) return;

      // Si el primer char y dt es grande, está bien (es el inicio de la ráfaga).
      // Si NO es el primero y dt es muy grande, probablemente es tipeo humano → ignorar.
      // Pero si _scanHidBuffer está vacío, aceptamos (inicio nuevo).
      if (_scanHidBuffer.length > 0 && dt > SCAN_HID_GAP_MS) {
        _scanHidBuffer = '';
      }
      _scanHidBuffer += e.key;
      if (inputBox) inputBox.value = _scanHidBuffer;
      if (statusEl && _scanHidBuffer.length === 1) {
        statusEl.textContent = '⚡ Capturando...';
        statusEl.style.color = '#fbbf24';
      }
    };
    document.addEventListener('keydown', _scanHidListener, true);
  }
  function _desactivarScannerHid() {
    if (_scanHidListener) {
      document.removeEventListener('keydown', _scanHidListener, true);
      _scanHidListener = null;
    }
    _scanHidBuffer = '';
  }

  function toggleDespScanInline() {
    const panel = document.getElementById('despScanInlinePanel');
    if (!panel) return;
    const visible = panel.style.display !== 'none';
    if (visible) { cerrarDespScan(); return; }

    // Cerrar cámara si estaba abierta
    const camPanel = document.getElementById('despCamInlinePanel');
    if (camPanel && camPanel.style.display !== 'none') {
      try { Scanner.stop(); } catch(_){}
      camPanel.style.display = 'none';
    }
    // Telón: colapsar header
    const header = document.getElementById('despHeaderCollapsible');
    if (header) header.classList.add('is-collapsed');
    panel.classList.remove('is-closing');
    panel.style.display = 'block';
    // Limpiar input + activar listener global de scanner HID
    const inp = document.getElementById('despScanInlineInput');
    if (inp) inp.value = '';
    _activarScannerHid();
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);
  }

  function cerrarDespScan() {
    const panel = document.getElementById('despScanInlinePanel');
    if (!panel) return;
    panel.classList.add('is-closing');
    setTimeout(() => {
      panel.style.display = 'none';
      panel.classList.remove('is-closing');
    }, 320);
    const header = document.getElementById('despHeaderCollapsible');
    if (header) header.classList.remove('is-collapsed');
    _desactivarScannerHid();
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);
    _renderCart();
    _updateFooter();
    _updateGenerarBtn();
    badgeUpdate();
  }

  // [v2.13.29] Limpiar input scan para volver a escanear (sin cerrar panel)
  function limpiarScanInput() {
    const inp = document.getElementById('despScanInlineInput');
    if (inp) { inp.value = ''; try { inp.focus(); } catch(_){} }
    const st = document.getElementById('despScanInlineStatus');
    if (st) st.textContent = '';
    _cerrarInlinePicker();
    try { SoundFX.click && SoundFX.click(); } catch(_){}
  }

  function _cerrarInlinePicker() {
    const pk = document.getElementById('despScanInlinePicker');
    if (pk) { pk.style.display = 'none'; pk.innerHTML = ''; }
  }

  function _mostrarInlinePicker(candidatos, codStr) {
    const pk = document.getElementById('despScanInlinePicker');
    if (!pk) return;
    pk.innerHTML = candidatos.map((p, idx) => {
      const cb      = String(p._scannedCb || p.codigoBarra || '');
      const display = String(p.codigoBarra || cb);
      const cbHtml  = display.startsWith(codStr)
        ? `<strong style="color:#fbbf24">${escHtml(codStr)}</strong>${escHtml(display.slice(codStr.length))}`
        : escHtml(display);
      return `<button onclick="DespachoView.seleccionarItemDespInline('${escAttr(cb)}')"
              style="width:100%;text-align:left;padding:9px 11px;border-radius:8px;
                     border:1px solid #334155;margin-bottom:5px;background:#0f172a;
                     display:flex;align-items:center;gap:9px;cursor:pointer;
                     -webkit-tap-highlight-color:transparent"
              ontouchstart="this.style.background='#0c1e30'"
              ontouchend="this.style.background='#0f172a'">
        <div style="flex-shrink:0;width:26px;height:26px;border-radius:6px;background:rgba(251,191,36,.15);
                    display:flex;align-items:center;justify-content:center">
          <span style="font-size:.65em;color:#fbbf24;font-weight:800">${idx+1}</span>
        </div>
        <div style="flex:1;min-width:0">
          <p style="font-size:.78em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(String(p.descripcion || cb))}</p>
          <p style="font-size:.65em;color:#64748b;font-family:monospace;margin-top:1px">${cbHtml}</p>
        </div>
      </button>`;
    }).join('');
    pk.style.display = 'block';
  }

  function seleccionarItemDespInline(scannedCb) {
    _cerrarInlinePicker();
    const candidatos = _buscarDespCandidatos(scannedCb);
    const prod = candidatos[0];
    if (!prod) return;
    _agregarDespDirecto(prod);
    const cb       = String(prod._scannedCb || prod.codigoBarra || scannedCb);
    const item     = _cart.find(c => c.codigoBarra === cb);
    const stockD   = item?.stockDisp || 0;
    const statusEl = document.getElementById('despScanInlineStatus');
    if (stockD > 0 && item && item.cantidad > stockD) {
      if (statusEl) { statusEl.textContent = `⚠ ${prod.descripcion || cb} · sobrestock`; statusEl.style.color = '#fb923c'; }
    } else {
      if (statusEl) { statusEl.textContent = `✓ ${prod.descripcion || cb}`; statusEl.style.color = '#34d399'; }
      SoundFX.beep(); vibrate(15);
    }
    const inp = document.getElementById('despScanInlineInput');
    if (inp) { inp.value = ''; inp.focus(); }
    _renderCart(); _updateFooter(); badgeUpdate();
    setTimeout(() => { if (statusEl) { statusEl.textContent = ''; statusEl.style.color = '#475569'; } }, 2000);
  }

  function submitDespScanInline() {
    const inp = document.getElementById('despScanInlineInput');
    const statusEl = document.getElementById('despScanInlineStatus');
    if (!inp) return;
    const val = inp.value.trim();
    if (!val) return;
    _cerrarInlinePicker();
    const candidatos = _buscarDespCandidatos(val);
    if (!candidatos.length) {
      // [v2.13.30] Diagnóstico útil: si las cachés están vacías, el error
      // no es del código sino del setup. Avisar explícito al operador.
      const nProds  = (OfflineManager.getProductosCache() || []).length;
      const nEquivs = (OfflineManager.getEquivalenciasCache() || []).length;
      let msg = '⚠ ' + val + ' · no existe en catálogo';
      if (nProds === 0 || nEquivs === 0) {
        msg = `⚠ Catálogo vacío (${nProds}p / ${nEquivs}eq) — refresca app`;
      }
      if (statusEl) { statusEl.textContent = msg; statusEl.style.color = '#f87171'; }
      SoundFX.warn(); vibrate([50, 25, 50]);
      inp.select();
      return;
    }
    // Match exacto → agregar directo. Con prefijo (varios) → mostrar picker para elegir
    if (!candidatos[0]._exacto) {
      if (statusEl) {
        statusEl.textContent = `↕ Prefijo · ${candidatos.length} coincidencias — elige una`;
        statusEl.style.color = '#fbbf24';
      }
      SoundFX.warn(); vibrate([30]);
      _mostrarInlinePicker(candidatos, val);
      inp.select();
      return;
    }
    const prod = candidatos[0];
    _agregarDespDirecto(prod);
    const cb     = String(prod._scannedCb || prod.codigoBarra || val);
    const item   = _cart.find(c => c.codigoBarra === cb);
    const stockD = item?.stockDisp || 0;
    if (stockD > 0 && item && item.cantidad > stockD) {
      if (statusEl) { statusEl.textContent = `⚠ ${prod.descripcion || cb} · sobrestock`; statusEl.style.color = '#fb923c'; }
    } else {
      if (statusEl) { statusEl.textContent = `✓ ${prod.descripcion || cb}`; statusEl.style.color = '#34d399'; }
      SoundFX.beep(); vibrate(15);
    }
    inp.value = '';
    inp.focus();
    _renderCart();
    _updateFooter();
    badgeUpdate();
    // Limpiar status tras 2s
    setTimeout(() => { if (statusEl) { statusEl.textContent = ''; statusEl.style.color = '#475569'; } }, 2000);
  }

  // ── Búsqueda manual por nombre ───────────────────────────────
  function abrirDespBusqueda() {
    const panel = document.getElementById('despSearchPanel');
    if (panel) panel.style.display = 'flex';
    setTimeout(() => {
      const inp = document.getElementById('despSearchInput');
      if (inp) { inp.value = ''; inp.focus(); }
      despBuscarInput('');
    }, 80);
  }

  function cerrarDespBusqueda() {
    const panel = document.getElementById('despSearchPanel');
    if (panel) panel.style.display = 'none';
  }

  function despBuscarInput(q) {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    const qL     = String(q || '').toLowerCase().trim();
    const list   = document.getElementById('despSearchList');
    if (!list) return;
    if (qL.length < 2) {
      list.innerHTML = '<p style="color:#475569;font-size:.78em;text-align:center;padding-top:20px">Escribe al menos 2 caracteres</p>';
      return;
    }

    // Agrupar equivs por skuBase para búsqueda y expansión
    const equivsByKey = {};
    equivs.forEach(e => {
      const key = String(e.skuBase || '').trim().toUpperCase();
      if (!key || !e.codigoBarra) return;
      if (!equivsByKey[key]) equivsByKey[key] = [];
      equivsByKey[key].push(e);
    });

    // Un entry por barcode visible (maestro + cada equiv)
    const entries = [];
    const seenCb  = new Set();

    prods.forEach(p => {
      const key = String(p.skuBase || p.idProducto || '').trim().toUpperCase();
      const eqs  = equivsByKey[key] || [];
      const haystack = [
        String(p.descripcion || '').toLowerCase(),
        String(p.codigoBarra || '').toLowerCase(),
        String(p.skuBase     || '').toLowerCase(),
        String(p.idProducto  || '').toLowerCase(),
        ...eqs.map(e => String(e.codigoBarra || '').toLowerCase() + ' ' + String(e.descripcion || '').toLowerCase())
      ].join(' ');
      if (!haystack.includes(qL)) return;

      // Barcode maestro
      const masterCb = String(p.codigoBarra || '');
      if (masterCb && !seenCb.has(masterCb)) {
        seenCb.add(masterCb);
        entries.push({ prod: p, cb: masterCb, label: p.descripcion || masterCb, tag: '' });
      }
      // Barcodes equiv
      eqs.forEach(e => {
        const equivCb = String(e.codigoBarra);
        if (!seenCb.has(equivCb)) {
          seenCb.add(equivCb);
          entries.push({ prod: p, cb: equivCb, label: e.descripcion || p.descripcion || equivCb, tag: 'EQ', _scannedCb: equivCb });
        }
      });
    });

    if (!entries.length) {
      list.innerHTML = '<p style="color:#475569;font-size:.78em;text-align:center;padding-top:20px">Sin resultados</p>';
      return;
    }
    list.innerHTML = entries.slice(0, 20).map(entry => {
      const cb    = escAttr(entry.cb);
      const tagHtml = entry.tag
        ? `<span style="font-size:.6em;font-weight:800;padding:1px 5px;border-radius:4px;
                         background:rgba(251,191,36,.15);color:#fbbf24;margin-left:5px;flex-shrink:0">${entry.tag}</span>`
        : '';
      return `<button onclick="DespachoView.seleccionarDespBusqueda('${cb}')"
              style="width:100%;text-align:left;padding:10px 12px;border-radius:10px;
                     border:1px solid #1e293b;margin-bottom:6px;background:#1e293b;
                     display:flex;align-items:center;gap:10px;cursor:pointer;
                     -webkit-tap-highlight-color:transparent"
              ontouchstart="this.style.background='#0c1e30'" ontouchend="this.style.background='#1e293b'">
        <div style="flex:1;min-width:0">
          <div style="display:flex;align-items:center">
            <p style="font-size:.83em;font-weight:700;color:#f1f5f9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(entry.label)}</p>
            ${tagHtml}
          </div>
          <p style="font-size:.67em;color:#64748b;font-family:monospace">${escHtml(entry.cb)}</p>
        </div>
      </button>`;
    }).join('');
  }

  function seleccionarDespBusqueda(codigoBarra) {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    // Buscar en maestro primero, luego en equivs
    let prod = prods.find(p => String(p.codigoBarra || '') === codigoBarra);
    if (!prod) {
      const equiv = equivs.find(e => String(e.codigoBarra || '') === codigoBarra);
      if (equiv) {
        const skuB = String(equiv.skuBase || '').trim().toUpperCase();
        const master = prods.find(p =>
          String(p.idProducto  || '').trim().toUpperCase() === skuB ||
          String(p.skuBase     || '').trim().toUpperCase() === skuB ||
          String(p.codigoBarra || '').trim().toUpperCase() === skuB
        );
        if (master) prod = { ...master, _scannedCb: codigoBarra };
      }
    }
    if (!prod) return;
    cerrarDespBusqueda();
    _agregarDespDirecto(prod);
    const item   = _cart.find(c => c.codigoBarra === codigoBarra);
    const stockD = item?.stockDisp || 0;
    if (!item || stockD === 0 || item.cantidad <= stockD) {
      const stockTxt = stockD > 0 ? ` · Stock: ${fmt(stockD,1)}` : '';
      _setDespStatus('ok', (prod.descripcion || codigoBarra) + stockTxt);
      SoundFX.beep(); vibrate(15);
    }
  }

  // ── Tipo de salida ───────────────────────────────────────────
  function selTipo(tipo) {
    _tipoSalida = tipo;
    const chips = {
      SALIDA_ZONA:       { id: 'despChipZona',     bc: '#0284c7', bg: 'rgba(14,165,233,.2)',    col: '#38bdf8'  },
      SALIDA_JEFATURA:   { id: 'despChipJefatura', bc: '#7c3aed', bg: 'rgba(124,58,237,.18)',   col: '#a78bfa'  },
      SALIDA_DEVOLUCION: { id: 'despChipDev',      bc: '#b45309', bg: 'rgba(180,83,9,.2)',       col: '#fbbf24'  }
    };
    Object.entries(chips).forEach(([t, cfg]) => {
      const el = document.getElementById(cfg.id);
      if (!el) return;
      const active = t === tipo;
      el.style.borderColor = active ? cfg.bc : '#334155';
      el.style.background  = active ? cfg.bg : 'transparent';
      el.style.color       = active ? cfg.col : '#475569';
    });
    const zonaWrap = document.getElementById('despZonaWrap');
    if (zonaWrap) zonaWrap.style.display = tipo === 'SALIDA_ZONA' ? 'block' : 'none';
  }

  // ── Render lista en modal cámara ─────────────────────────────
  function _renderDespList() {
    const list  = document.getElementById('despCamScannedList');
    const count = document.getElementById('despScanListCount');
    if (!list) return;
    const total = _cart.reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    if (count) count.textContent = total ? fmt(total, 2) + ' unid.' : '0 unid.';
    const undoBtn  = document.getElementById('despUndoBtn');
    const clearBtn = document.getElementById('despClearBtn');
    if (undoBtn)  undoBtn.style.display  = _despLastHistory.length > 0 ? 'inline-block' : 'none';
    if (clearBtn) clearBtn.style.display = _cart.length > 0 ? 'inline-block' : 'none';
    if (!_cart.length) {
      list.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;
          justify-content:center;padding:32px 20px;gap:10px;color:#334155">
        <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2"/>
          <rect x="7" y="7" width="10" height="10" rx="1"/>
        </svg>
        <p style="font-size:.78em;text-align:center;line-height:1.5;color:#334155">
          Apunta la cámara al código de barras
        </p>
      </div>`;
      return;
    }
    const btnStyle = `width:32px;height:32px;border-radius:8px;border:1px solid #334155;
      background:#0f172a;color:#94a3b8;font-size:1.05em;font-weight:700;cursor:pointer;
      display:flex;align-items:center;justify-content:center;flex-shrink:0;
      -webkit-tap-highlight-color:transparent`;
    list.innerHTML = _cart.map(item => {
      const cb    = escAttr(item.codigoBarra);
      const over  = item.stockDisp > 0 && item.cantidad > item.stockDisp;
      const overW = over ? `<span style="font-size:.63em;color:#f87171;margin-left:4px">⚠ stock:${fmt(item.stockDisp,1)}</span>` : '';
      const qtyFmt = fmt(item.cantidad, Number.isInteger(parseFloat(item.cantidad)) ? 0 : 2);
      const qtyCol = over ? '#f87171' : '#38bdf8';
      return `<div data-desp-cb="${cb}" style="display:flex;align-items:center;gap:8px;padding:9px 12px;
                border-radius:11px;background:#1e293b;border:1px solid ${over?'#7f1d1d':'#334155'};
                margin-bottom:7px">
        <div style="flex:1;min-width:0">
          <p style="font-size:.83em;font-weight:700;color:#f1f5f9;
                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${escHtml(item.descripcion)}</p>
          <p style="font-size:.67em;color:#64748b;font-family:monospace">${escHtml(item.codigoBarra)}${overW}</p>
        </div>
        <div style="display:flex;align-items:center;gap:4px;flex-shrink:0">
          <button onclick="DespachoView.despDecQty('${cb}')" style="${btnStyle}">−</button>
          <button class="desp-qty-pulse-btn" onclick="DespachoView.despEditQty('${cb}')"
                  style="min-width:46px;height:32px;border-radius:8px;padding:0 10px;
                         background:rgba(14,165,233,.15);border:1px solid rgba(14,165,233,.3);
                         color:${qtyCol};font-size:.92em;font-weight:900;cursor:pointer;
                         -webkit-tap-highlight-color:transparent">
            ${qtyFmt}
          </button>
          <button onclick="DespachoView.despIncQty('${cb}')" style="${btnStyle}">+</button>
        </div>
      </div>`;
    }).join('');
  }

  // ── Controles cantidad en modal cámara ───────────────────────
  function despIncQty(cb) {
    const item = _cart.find(c => c.codigoBarra === cb);
    if (!item) return;
    item.cantidad = (parseFloat(item.cantidad) || 0) + 1;
    _despLastHistory.push(cb);
    _saveCart(); _renderDespList();
    SoundFX.beepDouble(); vibrate(10);
  }

  function despDecQty(cb) {
    const idx = _cart.findIndex(c => c.codigoBarra === cb);
    if (idx < 0) return;
    if (_cart[idx].cantidad <= 1) {
      _cart.splice(idx, 1);
    } else {
      _cart[idx].cantidad = (parseFloat(_cart[idx].cantidad) || 0) - 1;
    }
    _saveCart(); _renderDespList(); vibrate(8);
  }

  async function despEditQty(cb) {
    const item = _cart.find(c => c.codigoBarra === cb);
    if (!item) return;
    const input = await _whPrompt('Nueva cantidad (decimales: usa punto):', String(item.cantidad), { titulo: item.descripcion || cb, inputMode: 'decimal', maxlength: 10 });
    if (input === null) return;
    const newQty = parseFloat(input.replace(',', '.'));
    if (isNaN(newQty) || newQty < 0) { toast('Cantidad inválida', 'warn', 2000); return; }
    if (newQty === 0) { despDecQty(cb); return; }
    item.cantidad = newQty;
    _saveCart(); _renderDespList();
    SoundFX.beep(); vibrate(10);
  }

  // ── Undo / Limpiar ───────────────────────────────────────────
  function despUndoLast() {
    if (!_despLastHistory.length) return;
    const cb  = _despLastHistory.pop();
    const idx = _cart.findIndex(c => c.codigoBarra === cb);
    if (idx < 0) { _renderDespList(); return; }
    if (_cart[idx].cantidad <= 1) { _cart.splice(idx, 1); }
    else                         { _cart[idx].cantidad--; }
    _saveCart(); _renderDespList(); vibrate(15);
    toast('↩ Deshecho', 'warn', 1200);
  }

  async function despLimpiarTodo() {
    if (!_cart.length) return;
    if (!await _whConfirm('¿Vaciar el carrito de despacho?', { warning: true, titulo: 'Vaciar carrito', okText: 'Vaciar' })) return;
    _cart = []; _despLastHistory = [];
    _saveCart(); _renderDespList(); vibrate(20);
  }

  // ── Render carrito en vista principal ────────────────────────
  function _renderCart() {
    // [v2.13.8] Hook lista sombra: cualquier cambio del cart actualiza la sombra
    try { _lsOnCartChange(); } catch(_){}
    const el = document.getElementById('despCartList');
    if (!el) return;
    // Si hay pickup activo, el carrito legacy se esconde (los extras tienen
    // su propia sección visual abajo del checklist del pickup).
    if (_pickupActivo) { el.style.display = 'none'; return; }
    // [v2.13.23] Si hay LISTA SOMBRA activa y el cart está vacío, esconder
    // el mensaje "Abre la cámara" porque el operador ya tiene el checklist
    // morado arriba. Si tiene items extras, sí los mostramos.
    if (_listaSombra && !_cart.length) { el.style.display = 'none'; return; }
    el.style.display = 'block';
    if (!_cart.length) {
      el.innerHTML = `
        <div class="card text-center py-10">
          <p class="text-4xl mb-2">📦</p>
          <p class="text-slate-400 text-sm">Abre la cámara para escanear productos</p>
          <p class="text-slate-600 text-xs mt-1">El carrito se guarda automáticamente</p>
        </div>`;
      return;
    }
    // [v2.13.30] Cuando hay SOMBRA activa, AGRUPAR por skuBase visualmente.
    // Estructura interna del cart se mantiene (1 entry por código escaneado),
    // pero la VISTA muestra "Producto · skuBase · N un" con suma de canónico
    // + equivalentes. El backend recibe códigos individuales sin cambios.
    // Sin sombra: vista normal (1 línea por código).
    const itemsBase = _listaSombra ? _cart.slice().reverse() : _cart;
    let itemsView = itemsBase;
    if (_listaSombra) {
      // Agrupar por skuBase
      const productosCache = OfflineManager.getProductosCache() || [];
      const equivsCache    = OfflineManager.getEquivalenciasCache() || [];
      function _resolverSku(cb) {
        const cbN = normCb(cb);
        if (!cbN) return '';
        const p = productosCache.find(x =>
          (String(x.codigoBarra || '').trim().toUpperCase() === cbN ||
           String(x.idProducto  || '').trim().toUpperCase() === cbN) &&
          parseFloat(x.factorConversion || 1) === 1);
        if (p) return { sku: String(p.skuBase || p.idProducto || p.codigoBarra || ''), desc: String(p.descripcion || ''), unidad: String(p.unidad || '') };
        const e = equivsCache.find(x => String(x.codigoBarra || '').trim().toUpperCase() === cbN);
        if (e) {
          const sk = String(e.skuBase || '');
          const can = productosCache.find(x =>
            parseFloat(x.factorConversion || 1) === 1 &&
            (String(x.skuBase || '').trim().toUpperCase() === sk.toUpperCase() ||
             String(x.idProducto || '').trim().toUpperCase() === sk.toUpperCase() ||
             String(x.codigoBarra || '').trim().toUpperCase() === sk.toUpperCase()));
          return { sku: sk, desc: String(can?.descripcion || ''), unidad: String(can?.unidad || '') };
        }
        return { sku: '', desc: '', unidad: '' };
      }
      // Agrupar (preservando orden de primera aparición)
      const grupos = [];
      const skuIdx = new Map();
      itemsBase.forEach(c => {
        const info = _resolverSku(c.codigoBarra);
        const key = info.sku || ('__' + c.codigoBarra);  // si no resuelve, key única
        if (!skuIdx.has(key)) {
          skuIdx.set(key, grupos.length);
          grupos.push({
            skuBase: info.sku,
            descripcion: info.desc || c.descripcion || c.codigoBarra,
            unidad: info.unidad || c.unidad || '',
            cantidadTotal: 0,
            codigos: []   // {codigoBarra, cantidad, stockDisp}
          });
        }
        const g = grupos[skuIdx.get(key)];
        g.cantidadTotal += parseFloat(c.cantidad) || 0;
        g.codigos.push({
          codigoBarra: c.codigoBarra,
          cantidad: parseFloat(c.cantidad) || 0,
          stockDisp: c.stockDisp || 0,
          descripcion: c.descripcion
        });
      });
      // Render agrupado
      el.innerHTML = grupos.map((g, i) => {
        const esUltimo = i === 0;
        const tieneVarios = g.codigos.length > 1;
        const skuLbl = g.skuBase || '(sin SKU)';
        const codigosHtml = g.codigos.map(c => {
          const safeId = escAttr(c.codigoBarra);
          const cb = escHtml(c.codigoBarra);
          const over = c.stockDisp > 0 && c.cantidad > c.stockDisp;
          const overWarn = over ? `<span class="text-[10px] text-red-400">⚠</span>` : '';
          return `
            <div class="desp-cart-subrow">
              <span class="desp-cart-sub-cb">${cb} ${overWarn}</span>
              <div class="desp-qty-wrap" style="transform:scale(.85)">
                <button class="desp-qty-btn" onclick="DespachoView.decQty('${safeId}')">−</button>
                <input id="despQty-${safeId}" type="number" inputmode="decimal" step="0.1" min="0"
                       value="${c.cantidad}"
                       class="desp-qty-input"
                       onblur="DespachoView.blurQty('${safeId}',this.value)"
                       onfocus="this.select()">
                <button class="desp-qty-btn" onclick="DespachoView.incQty('${safeId}')">+</button>
                <button class="desp-qty-btn" style="border-color:rgba(239,68,68,.4);color:#f87171;margin-left:2px"
                        onclick="DespachoView.quitarItem('${safeId}')">✕</button>
              </div>
            </div>`;
        }).join('');
        return `
          <div class="card-sm desp-cart-item desp-cart-grupo${esUltimo ? ' desp-cart-recien' : ''}"
               style="animation-delay:${i*.03}s">
            <div class="flex items-center gap-3">
              <div class="flex-1 min-w-0">
                <p class="font-semibold text-sm truncate">${escHtml(g.descripcion)}</p>
                <p class="text-[10px] text-slate-500 font-mono">
                  <span style="color:#a5b4fc">${escHtml(skuLbl)}</span>
                  ${tieneVarios ? `<span class="ml-1 text-[9px] text-slate-500">· ${g.codigos.length} códigos</span>` : ''}
                </p>
              </div>
              <div class="text-right shrink-0">
                <div class="text-base font-bold text-amber-300">${g.cantidadTotal}</div>
                <div class="text-[10px] text-slate-500">${g.unidad || 'un'}</div>
              </div>
            </div>
            <div class="desp-cart-codigos">${codigosHtml}</div>
          </div>`;
      }).join('');
      return;
    }
    // Sin sombra: vista normal (1 línea por código)
    el.innerHTML = itemsBase.map((item, i) => {
      const over     = item.stockDisp > 0 && item.cantidad > item.stockDisp;
      const safeId   = escAttr(item.codigoBarra);
      const overWarn = over ? `<span class="text-xs text-red-400">⚠ stock: ${fmt(item.stockDisp,1)}</span>` : '';
      const esUltimo = false;
      return `
      <div class="card-sm desp-cart-item flex items-center gap-3${esUltimo ? ' desp-cart-recien' : ''}"
           id="despRow-${safeId}" style="animation-delay:${i*.03}s">
        <div class="flex-1 min-w-0">
          <p class="font-semibold text-sm truncate">${escHtml(item.descripcion)}</p>
          <p class="text-xs text-slate-400 font-mono">${escHtml(item.codigoBarra)} ${overWarn}</p>
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

  // ── Controles de cantidad en vista principal ─────────────────
  function incQty(cod) { _adjustQty(cod, q => q + 1); }
  function decQty(cod) { _adjustQty(cod, q => Math.max(0, q - 1)); }

  function _adjustQty(cod, fn) {
    const idx = _cart.findIndex(c => c.codigoBarra === cod);
    if (idx < 0) return;
    const newQty = fn(parseFloat(_cart[idx].cantidad) || 0);
    if (newQty <= 0) { quitarItem(cod); return; }
    _cart[idx].cantidad = newQty;
    _saveCart();
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
    // El footer legacy fue reemplazado por el botón único "Generar Guía" arriba.
    // Mantengo la función por compat (algunas funciones internas la llaman).
    _updateGenerarBtn();
  }

  function finalizar() {
    if (!_cart.length) return;
    // Zona select
    const zonas = OfflineManager.getZonasCache();
    const sel   = document.getElementById('despZonaSelect');
    sel.innerHTML = '<option value="">— Seleccionar zona —</option>' +
      zonas.map(z => `<option value="${escAttr(z.idZona)}">${escHtml(z.nombre || z.idZona)}</option>`).join('');
    sel.value = _pickupActivo?.idZona || _loadZona();
    // Resumen
    const n   = _cart.length;
    const uds = _cart.reduce((s, c) => s + (parseFloat(c.cantidad) || 0), 0);
    document.getElementById('despResumenProds').textContent = `${n} producto${n !== 1 ? 's' : ''}`;
    document.getElementById('despResumenUds').textContent   = `${fmt(uds,2)} unidades`;
    // Tipo chips
    selTipo(_tipoSalida);
    // Limpiar nota
    const notaEl = document.getElementById('despNotaFinal');
    if (notaEl) notaEl.value = '';
    // Conflictos de stock
    const conflictos = _cart.filter(c => c.stockDisp > 0 && c.cantidad > c.stockDisp);
    const conflPanel = document.getElementById('despConflictoPanel');
    const conflList  = document.getElementById('despConflictoList');
    if (conflPanel && conflList) {
      if (conflictos.length) {
        conflList.innerHTML = conflictos.map(c =>
          `<div>• ${escHtml(c.descripcion)}: pides <b>${fmt(c.cantidad,2)}</b>, stock <b>${fmt(c.stockDisp,1)}</b></div>`
        ).join('');
        conflPanel.style.display = 'block';
      } else {
        conflPanel.style.display = 'none';
      }
    }
    abrirSheet('sheetDespFinalizar');
  }

  let _dspGenerarBusy = false;
  function _setDspGenerarBusy(busy) {
    _dspGenerarBusy = !!busy;
    const btn = document.getElementById('btnConfirmarDespacho');
    if (!btn) return;
    btn.disabled = !!busy;
    btn.classList.toggle('opacity-50', !!busy);
    btn.classList.toggle('pointer-events-none', !!busy);
    if (busy) {
      btn.dataset._lbl = btn.innerHTML;
      btn.innerHTML = '⏳ Generando guía...';
    } else if (btn.dataset._lbl) {
      btn.innerHTML = btn.dataset._lbl;
      delete btn.dataset._lbl;
    }
  }

  async function confirmarDespacho() {
    // ─── LOCK ANTI-DOBLE-CLICK ───────────────────────────────
    // Bug histórico (12 may + 13 may): triple click generaba 3 guías en
    // <30s. El backend ahora tiene idempotencia, pero acá bloqueamos en
    // origen para que ni siquiera intente.
    if (_dspGenerarBusy) {
      toast('Espera, ya estamos generando la guía...', 'warn');
      return;
    }
    // [v2.13.10] Aviso si la lista sombra está incompleta — el operador
    // puede continuar igual (no bloquea), solo confirma. La lista se
    // cierra al cerrar el despacho sin importar el estado.
    if (_listaSombra && _listaSombra.items) {
      const restantes = _listaSombra.items.filter(i => (i.cantidadEscaneada || 0) < i.cantidad);
      if (restantes.length > 0) {
        const lista = restantes.slice(0, 5).map(i => `  · ${i.nombre} (${(i.cantidadEscaneada||0).toFixed(1)}/${i.cantidad.toFixed(1)})`).join('\n');
        const masTxt = restantes.length > 5 ? `\n  · …y ${restantes.length - 5} más` : '';
        if (!await _whConfirm(`⚠ Lista sombra incompleta — quedan ${restantes.length} items sin atender:\n\n${lista}${masTxt}\n\n¿Despachar igual?\n\nLa lista sombra se cerrará tras el despacho.`, { warning: true, titulo: 'Lista incompleta', okText: 'Despachar igual' })) {
          return;
        }
      }
    }

    // Si hay un pickup activo, ese es el camino: cierra contra cerrarPickupConDespacho
    // (emite GUIA_SALIDA con códigos reales escaneados + faltantes en observación)
    if (_pickupActivo) {
      cerrarSheet('sheetDespFinalizar');
      const completo = _pickupTotalmenteCompleto();
      cerrarDespachoPickup(!completo);
      return;
    }

    const idZona = _tipoSalida === 'SALIDA_ZONA' ? document.getElementById('despZonaSelect').value : '';
    if (_tipoSalida === 'SALIDA_ZONA' && !idZona) { toast('Selecciona la zona de destino', 'warn'); return; }
    if (idZona) _saveZona(idZona);
    const nota = document.getElementById('despNotaFinal')?.value?.trim() || '';
    // idempotencyKey: identificador único de ESTE click. Si el cliente
    // reintenta por timeout, el backend reconoce el key y retorna la guía
    // ya creada en vez de duplicarla.
    const idempotencyKey = 'DSP-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
    const payload = {
      idZona,
      tipo:    _tipoSalida,
      nota,
      usuario: window.WH_CONFIG?.usuario || '',
      items:   _cart.map(c => ({ codigoBarra: c.codigoBarra, cantidad: c.cantidad })),
      imprimir: false,  // GAS no imprime — frontend dispara impresión separada para no bloquear
      idempotencyKey
    };
    _setDspGenerarBusy(true);

    // Resolver nombre de zona para historial
    const zonas = OfflineManager.getZonasCache();
    const zonaObj = zonas.find(z => String(z.idZona) === String(idZona));
    const nombreZona = zonaObj ? (zonaObj.nombre || zonaObj.idZona) : '';

    // Optimista: vaciar carrito y cerrar sheet inmediatamente
    const cartSnapshot = [..._cart];
    const tipoSnapshot = _tipoSalida;
    const pickupSnapshot = _pickupActivo;
    const itemsSnapshot = cartSnapshot.map(c => ({
      codigoBarra: c.codigoBarra, descripcion: c.descripcion, cantidad: c.cantidad
    }));
    const histBase = { ts: Date.now(), n: cartSnapshot.length, tipo: tipoSnapshot,
                       idZona, nombreZona, nota, items: itemsSnapshot };
    _cart = []; _tipoSalida = 'SALIDA_ZONA'; _saveCart();
    cerrarSheet('sheetDespFinalizar');
    _renderCart(); _updateFooter(); badgeUpdate();
    toast('⏳ Generando guía...', 'info', 55000);

    // GAS en segundo plano — timeout generoso 55s
    Promise.race([
      API.crearDespachoRapido(payload),
      new Promise((_, rej) => setTimeout(() => rej({ timeout: true }), 55000))
    ]).then(res => {
      if (res.ok) {
        const d = res.data;
        toast(`✅ Guía ${d.idGuia} generada · Imprimiendo...`, 'ok', 6000);
        if (d.errores?.length) toast(`⚠ ${d.errores.length} ítem(s) con error`, 'warn', 5000);
        SoundFX.done(); vibrate([30, 15, 30, 15, 60]);
        _saveHist({ ...histBase, idGuia: d.idGuia, ok: true });
        // [Fix #2+#5 v2.11.2] Re-impresión validada con QR + verificación
        // de items. Pasamos `esperadoDetalles` para que el backend espere
        // a que la hoja esté completa (retry interno con flush). Si el
        // ticket sale con menos items de los esperados, mostramos warning
        // claro al operador para que reimprima desde el historial.
        const esperados = d.esperados || d.items || cartSnapshot.length;
        try {
          const base = location.origin + location.pathname.replace(/\/[^/]*$/, '');
          const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(d.idGuia)}`;
          // [v2.13.28] Si había sombra activa al despachar, pasar snapshot
          // al backend para imprimir ticket comparativo PEDIDO vs DESPACHADO.
          const sombraSnap = (_listaSombra && Array.isArray(_listaSombra.items) && _listaSombra.items.length)
            ? JSON.stringify(_listaSombra.items)
            : undefined;
          API.imprimirTicketGuia({
            idGuia: d.idGuia,
            reporteUrl,
            esperadoDetalles: esperados,
            sombraSnapshot: sombraSnap
          }).then(r2 => {
            // [v2.13.59] Si el backend dedupó (ya_impresa) → confianza implícita en
            // que la 1ra impresión salió bien. NO mostrar warning ni intentar otra copia.
            if (r2?.ya_impresa) {
              console.log('[Despacho] dedupado · ticket ya impreso por crearDespachoRapido');
              return;
            }
            const impreso = r2?.data?.detallesImpresos || 0;
            if (r2?.ok && impreso >= esperados) {
              console.log('[Despacho] ticket OK · ' + impreso + '/' + esperados + ' items');
            } else if (r2?.ok && impreso < esperados) {
              toast(`⚠ Ticket con ${impreso}/${esperados} items — reimprimí desde el historial`, 'warn', 8000);
            }
          }).catch(() => {});
        } catch (e) { /* non-fatal */ }
        if (pickupSnapshot) {
          API.actualizarPickup({ idPickup: pickupSnapshot.idPickup, estado: 'COMPLETADO' }).catch(() => {});
          _pickupsPendientes = _pickupsPendientes.filter(p => p.idPickup !== pickupSnapshot.idPickup);
          _pickupActivo = null;
        }
        // [v2.13.8] Si había lista sombra, limpiarla al cerrar despacho con éxito
        if (_listaSombra) {
          // [v2.13.15] Marcar como COMPLETADA en backend antes de limpiar local
          if (_listaSombra.idBackend) {
            API.cerrarListaSombra({
              idLista: _listaSombra.idBackend,
              items: JSON.stringify(_listaSombra.items),
              localId: 'L' + Date.now() + Math.random().toString(36).slice(2, 8)
            }).catch(() => {});
          }
          _listaSombra = null;
          _lsSave();
          _lsRender();
          try { _lsRefrescarPanel(); } catch(_){}
        }
      } else {
        SoundFX.error(); vibrate([80, 40, 80]);
        _saveHist({ ...histBase, idGuia: '—', ok: false });
        toast('Error al generar guía: ' + (res.error || 'Sin respuesta'), 'danger', 8000);
        // Restaurar carrito para que el usuario pueda reintentar
        if (!_cart.length) { _cart = cartSnapshot; _saveCart(); _renderCart(); _updateFooter(); badgeUpdate(); }
      }
      _renderHist();
    }).catch(e => {
      const msg = e?.timeout ? 'Tiempo agotado — verifica tu conexión' : 'Sin conexión';
      SoundFX.error(); vibrate([80, 40, 80]);
      _saveHist({ ...histBase, idGuia: '—', ok: false });
      toast('Error: ' + msg, 'danger', 8000);
      if (!_cart.length) { _cart = cartSnapshot; _saveCart(); _renderCart(); _updateFooter(); badgeUpdate(); }
      _renderHist();
    }).finally(() => {
      _setDspGenerarBusy(false);
    });
  }

  async function cancelar() {
    if (!_cart.length) return;
    if (!await _whConfirm('¿Vaciar el carrito de despacho?', { warning: true, titulo: 'Vaciar carrito', okText: 'Vaciar' })) return;
    _cart = []; _saveCart(); _renderCart(); _updateFooter();
  }

  // ── Badge global (carrito + pickups pendientes) ─────────────
  const PICKUPS_SNAPSHOT_KEY = 'wh_pickups_snapshot_ids';
  let _pickupsPendientes = [];
  // Snapshot persistente — sobrevive refresh para que el detector "nuevo"
  // funcione aunque cierres la app justo cuando llega un pickup.
  let _ultimosIdsPickups = new Set(_loadIdsSnapshot());
  let _pollTimer = null;
  function _loadIdsSnapshot() {
    try { return JSON.parse(localStorage.getItem(PICKUPS_SNAPSHOT_KEY) || '[]') || []; }
    catch { return []; }
  }
  function _saveIdsSnapshot(idsArr) {
    try { localStorage.setItem(PICKUPS_SNAPSHOT_KEY, JSON.stringify(idsArr || [])); } catch(_){}
  }

  function badgeUpdate() {
    _cart = _loadCart();
    const n   = _cart.length;
    const nPickups = _pickupsPendientes.length;
    const hay = nPickups > 0;

    // Badge topbar — dinámico: muestra progreso del pickup activo si existe,
    // si no, modo carrito legacy 🛒DSP. Esconder si no hay nada.
    const despBadge = document.getElementById('despModoIndicador');
    if (despBadge) {
      if (_pickupActivo) {
        const items     = _pickupActivo.items || [];
        const totalUds  = items.reduce((s,it) => s + (parseFloat(it.solicitado) || 0), 0);
        const totalDesp = items.reduce((s,it) => s + (parseFloat(it.despachado) || 0), 0);
        const completo  = _pickupTotalmenteCompleto();
        despBadge.classList.remove('hidden');
        despBadge.textContent = `📦 ${totalDesp}/${totalUds}`;
        despBadge.style.background = completo ? '#047857' : '#4338ca';
        despBadge.title = completo ? 'Pickup completo · cerrar despacho' : 'Pickup activo · seguir despachando';
      } else {
        despBadge.textContent = '🛒DSP';
        despBadge.style.background = '#065f46';
        despBadge.title = 'Volver al Despacho Rápido';
        despBadge.classList.toggle('hidden', n === 0);
      }
    }

    // Botón FAB en vista Productos
    const fab     = document.getElementById('btnDespCartFab');
    const fabN    = document.getElementById('despCartFabN');
    const fabPick = document.getElementById('despCartFabPickN');
    const dot     = document.getElementById('despPickAlertDot');
    if (fabN)    { fabN.textContent = n; fabN.style.display = n > 0 ? 'inline-flex' : 'none'; }
    if (fabPick) { fabPick.textContent = nPickups; fabPick.style.display = nPickups > 0 ? 'inline-flex' : 'none'; }
    if (dot)     { dot.style.display = hay ? 'block' : 'none'; }
    if (fab)     { fab.classList.toggle('has-pickups', hay); }

    // Lista de pickups pendientes + [v2.13.16] listas sombra del equipo
    // en el MISMO container, ordenadas: pickups primero, listas sombra después.
    // Si hay un pickup activo, ESCONDER toda la lista — el operador atiende
    // uno a la vez. Al soltar el activo, los pendientes vuelven a aparecer.
    const listaEl = document.getElementById('despPickupsLista');
    if (listaEl) {
      // Filtrar listas sombra visibles (no la que ya estoy trabajando localmente)
      const sombras = (_lsPanelData || []).filter(l =>
        !_listaSombra || l.idLista !== _listaSombra.id
      );
      if (_pickupActivo) {
        listaEl.style.display = 'none';
        listaEl.innerHTML = '';
      } else if (_pickupsPendientes.length > 0 || sombras.length > 0) {
        listaEl.style.display = 'block';
        const pickupsHtml = _pickupsPendientes.slice(0, 8).map(p => {
          const items = Array.isArray(p.items) ? p.items : [];
          const totalUds = items.reduce((s, it) => s + (parseFloat(it.solicitado) || 0), 0);
          const totalDesp = items.reduce((s, it) => s + (parseFloat(it.despachado) || 0), 0);
          const pct = totalUds > 0 ? Math.round((totalDesp / totalUds) * 100) : 0;
          const enProceso = String(p.estado) === 'EN_PROCESO';
          const fuente = String(p.fuente || '').toLowerCase();
          const fuenteIcon = fuente.indexOf('me_cierre') >= 0 ? '🛒' : '📥';
          const fuenteLbl  = fuente.indexOf('me_cierre') >= 0 ? 'Cierre caja' : (p.fuente || 'Externo');
          let hace = '';
          try {
            const t = new Date(p.fechaCreado).getTime();
            const min = Math.floor((Date.now() - t) / 60000);
            hace = min < 1 ? 'recién' : min < 60 ? ('hace ' + min + 'm') : ('hace ' + Math.floor(min/60) + 'h');
          } catch(_) {}
          // Lock visual — atendidoPor distinto al usuario actual = bloqueado.
          // Comparación normalizada para tolerar dobles espacios / mayúsculas
          // entre devices del mismo operador.
          const usuario = window.WH_CONFIG?.usuario || '';
          const atp = String(p.atendidoPor || '').trim();
          const lockedByMe    = atp && usuario && _sameUser(atp, usuario);
          const lockedByOther = atp && usuario && !lockedByMe;
          let btnHtml = '';
          if (lockedByOther) {
            btnHtml = `<button disabled class="btn btn-sm flex-shrink-0"
                       style="background:rgba(71,85,105,.3);color:#94a3b8;border-color:rgba(71,85,105,.5);cursor:not-allowed"
                       title="Atendido por ${escAttr(atp)}">
                       🔒 ${escHtml(atp)}
                     </button>`;
          } else if (lockedByMe || enProceso) {
            btnHtml = `<button onclick="DespachoView.abrirPickup('${escAttr(p.idPickup)}')"
                       class="btn btn-sm flex-shrink-0"
                       style="background:rgba(99,102,241,.25);color:#a5b4fc;border-color:rgba(99,102,241,.4)">
                       ↻ Continuar
                     </button>`;
          } else {
            btnHtml = `<button onclick="DespachoView.abrirPickup('${escAttr(p.idPickup)}')"
                       class="btn btn-sm flex-shrink-0"
                       style="background:rgba(245,158,11,.25);color:#fbbf24;border-color:rgba(245,158,11,.4)">
                       ▶ Jalar
                     </button>`;
          }
          const cardBorder = lockedByOther ? 'rgba(71,85,105,.5)' : (enProceso ? 'rgba(99,102,241,.5)' : 'rgba(245,158,11,.5)');
          const cardBg     = lockedByOther ? 'rgba(71,85,105,.08)' : (enProceso ? 'rgba(99,102,241,.08)' : 'rgba(245,158,11,.08)');
          const animation  = lockedByOther ? 'none' : (enProceso ? 'none' : 'despPickListPulse 2.4s ease-in-out infinite');
          return `
            <div class="card flex items-center gap-3"
                 style="border:1.5px solid ${cardBorder};background:${cardBg};animation:${animation}">
              <div class="flex-shrink-0" style="font-size:24px">${fuenteIcon}</div>
              <div class="flex-1 min-w-0">
                <div class="flex items-center gap-2 flex-wrap">
                  <span class="font-bold text-sm" style="color:${lockedByOther?'#94a3b8':(enProceso ? '#a5b4fc' : '#fbbf24')}">${p.idZona || '—'}</span>
                  <span class="text-xs text-slate-400">·</span>
                  <span class="text-xs text-slate-400">${fuenteLbl}</span>
                  <span class="text-xs text-slate-500">· ${hace}</span>
                  ${atp ? `<span class="text-[10px]" style="color:${lockedByMe?'#a5b4fc':'#94a3b8'};background:rgba(15,23,42,.6);padding:1px 6px;border-radius:6px;font-weight:700">🔒 ${escHtml(atp)}${lockedByMe?' (yo)':''}</span>` : ''}
                </div>
                <p class="text-xs text-slate-300 mt-0.5">
                  ${items.length} producto${items.length !== 1 ? 's' : ''} · ${Math.round(totalUds)} uds
                  ${enProceso ? ` · <span style="color:#a5b4fc;font-weight:700">${pct}% despachado</span>` : ''}
                </p>
                ${p.creadoPor ? `<p class="text-[10px] text-slate-500">cajero: ${escHtml(p.creadoPor)}</p>` : ''}
              </div>
              ${btnHtml}
            </div>`;
        }).join('');
        // [v2.13.16] Cards de listas sombra del equipo — mismo estilo card,
        // color morado para diferenciar visualmente de pickups.
        const miUsuarioLs = window.WH_CONFIG?.usuario || '';
        const sombrasHtml = sombras.slice(0, 8).map(l => {
          const esEnUso = l.estado === 'EN_USO';
          const esMia   = esEnUso && _sameUser(l.usuarioTomada, miUsuarioLs);
          const progreso = l.total ? `${l.completos}/${l.total}` : '0/0';
          let hace = '';
          try {
            const t = new Date(l.fechaCreacion).getTime();
            const min = Math.floor((Date.now() - t) / 60000);
            hace = min < 1 ? 'recién' : min < 60 ? ('hace ' + min + 'm') : ('hace ' + Math.floor(min/60) + 'h');
          } catch(_){}
          const atp = String(l.usuarioTomada || '').trim();
          let btnH;
          if (esMia) {
            btnH = `<button onclick="DespachoView.tomarListaSombraDelPanel('${escAttr(l.idLista)}')"
                     class="btn btn-sm flex-shrink-0"
                     style="background:rgba(168,85,247,.25);color:#d8b4fe;border-color:rgba(168,85,247,.4)">
                     ↩ Continuar
                   </button>`;
          } else if (esEnUso) {
            btnH = `<button disabled class="btn btn-sm flex-shrink-0"
                     style="background:rgba(71,85,105,.3);color:#94a3b8;border-color:rgba(71,85,105,.5);cursor:not-allowed"
                     title="Jalada por ${escAttr(atp)}">
                     🔒 ${escHtml(atp)}
                   </button>`;
          } else {
            btnH = `<button onclick="DespachoView.tomarListaSombraDelPanel('${escAttr(l.idLista)}')"
                     class="btn btn-sm flex-shrink-0"
                     style="background:rgba(168,85,247,.25);color:#d8b4fe;border-color:rgba(168,85,247,.4)">
                     ▶ Jalar sombra
                   </button>`;
          }
          // [v2.13.19] Botón X para anular la lista del feed. Solo visible si NO
          // está EN_USO por otro (no le quitamos a alguien que está trabajando).
          const puedeAnular = !esEnUso || esMia;
          const btnX = puedeAnular
            ? `<button onclick="event.stopPropagation();DespachoView.anularListaSombraDelFeed('${escAttr(l.idLista)}')"
                  class="lsc-btn-x" title="Eliminar lista del feed">✕</button>`
            : '';
          const cardBorder = esMia ? 'rgba(168,85,247,.55)' : (esEnUso ? 'rgba(71,85,105,.5)' : 'rgba(168,85,247,.5)');
          const cardBg     = esMia ? 'rgba(168,85,247,.12)' : (esEnUso ? 'rgba(71,85,105,.08)' : 'rgba(168,85,247,.08)');
          const anim       = esEnUso ? 'none' : 'despPickListPulse 2.4s ease-in-out infinite';
          return `
            <div class="card flex items-center gap-3"
                 style="border:1.5px solid ${cardBorder};background:${cardBg};animation:${anim}">
              <div class="flex-shrink-0" style="font-size:24px">📋</div>
              <div class="flex-1 min-w-0">
                <div class="flex items-center gap-2 flex-wrap">
                  <span class="font-bold text-sm" style="color:#d8b4fe">Lista sombra</span>
                  <span class="text-xs text-slate-400">·</span>
                  <span class="text-xs text-slate-400">${escHtml(l.usuarioCreador)}</span>
                  <span class="text-xs text-slate-500">· ${hace}</span>
                  ${atp ? `<span class="text-[10px]" style="color:${esMia?'#d8b4fe':'#94a3b8'};background:rgba(15,23,42,.6);padding:1px 6px;border-radius:6px;font-weight:700">🔒 ${escHtml(atp)}${esMia?' (yo)':''}</span>` : ''}
                </div>
                <p class="text-xs text-slate-300 mt-0.5">
                  ${l.total} producto${l.total !== 1 ? 's' : ''}
                  ${esEnUso ? ` · <span style="color:#d8b4fe;font-weight:700">${progreso}</span>` : ''}
                </p>
              </div>
              ${btnH}
              ${btnX}
            </div>`;
        }).join('');
        listaEl.innerHTML = pickupsHtml + sombrasHtml;
      } else {
        listaEl.style.display = 'none';
      }
    }
  }

  async function _pollPickups() {
    const res = await API.getPickups({ estado: 'PENDIENTE,EN_PROCESO' }).catch(() => ({ ok: false }));
    const lista = (res && res.ok) ? (res.data || []) : [];
    const idsActuales = new Set(lista.map(p => p.idPickup));
    const huboSnapshot = localStorage.getItem(PICKUPS_SNAPSHOT_KEY) !== null;
    const nuevos = !huboSnapshot ? [] : lista.filter(p => !_ultimosIdsPickups.has(p.idPickup) && p.estado === 'PENDIENTE');
    // Filtrar el pickup activo para que NO aparezca en la lista de pendientes
    // (ya está como checklist abajo — duplicar confunde).
    const activoId = _pickupActivo ? String(_pickupActivo.idPickup) : null;
    _pickupsPendientes = activoId
      ? lista.filter(p => String(p.idPickup) !== activoId)
      : lista;
    _ultimosIdsPickups = idsActuales;
    _saveIdsSnapshot([...idsActuales]);
    badgeUpdate();
    if (nuevos.length > 0) {
      // Snooze: si ya estoy despachando un pickup, no abrir overlay fullscreen.
      // Solo toast suave + sonido corto + voz breve + count actualizado en FAB.
      if (_pickupActivo) {
        try { SoundFX.beep(); } catch(_){}
        const z = nuevos[0].idZona || 'otra zona';
        const tot = nuevos[0].items ? nuevos[0].items.reduce((s,it)=>s+(parseFloat(it.solicitado)||0),0) : 0;
        toast(`📦 Llegó otro pickup · ${nuevos[0].items?.length || 0} productos · ${Math.round(tot)} uds`, 'info', 4000);
        _vozAnunciar(`Otro pedido pendiente para zona ${z}`, { rate: 1.1 });
      } else if (typeof mostrarAlertaPickupNuevo === 'function') {
        mostrarAlertaPickupNuevo(nuevos[0]);
      }
    }
  }

  // Toggles de ordenamiento del checklist (search + sort)
  function pickupSetSearch(val) {
    _pickupSearch = String(val || '');
    _renderPickupChecklistInSheet();
  }
  function pickupSetSort(modo) {
    _pickupSort = modo || 'pendientes';
    _updatePickupSortButtons();
    _renderPickupChecklistInSheet();
  }
  function _updatePickupSortButtons() {
    ['pendientes','az','zona'].forEach(m => {
      const btn = document.getElementById('pkckSort_' + m);
      if (!btn) return;
      btn.classList.toggle('is-active', m === _pickupSort);
    });
  }

  function startPoll() {
    if (_pollTimer) return;
    _pollPickups();
    // Polling 30s — antes era 120s. Ruido de almacén = se necesita aviso rápido.
    _pollTimer = setInterval(_pollPickups, 30_000);
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
  let _pickupClosing = false; // lock anti-doble-click en cerrarDespachoPickup
  let _matchResults = []; // [{nombre, qty, producto, score, status, accepted}]
  // skuBase del ÚLTIMO item escaneado. Solo ESE item muestra los controles
  // (+/- o input granel) en el checklist — los demás los ocultan, para que
  // el operador no edite por error un producto distinto al que tiene en mano.
  let _pickupItemActivo = null;
  // codigoBarra del último producto escaneado al despacho rápido (cart).
  // Lo usa el flotante cross-módulo para mostrar sus controles +/-.
  let _despItemActivo = null;

  // Normalización de nombre de operador para comparar locks de pickup.
  // Tolera dobles espacios, mayúsculas y trims (causa real: el nombre
  // guardado en PERSONAL a veces tiene espacios extra, y al comparar
  // ===-estricto contra el usuario de otro device el operador se
  // bloquea de sí mismo).
  function _normUser(u) {
    return String(u || '').trim().toLowerCase().replace(/\s+/g, ' ');
  }
  function _sameUser(a, b) {
    var na = _normUser(a), nb = _normUser(b);
    return na && nb && na === nb;
  }

  // ════════════════════════════════════════════════════════════
  // PICKUP ACTIVO — lógica core
  // ════════════════════════════════════════════════════════════
  let _autosavePickupTimer = null;
  let _autosaveLastTs = 0;

  // Buscar item del pickup que matchee con el producto/código escaneado.
  function _matchPickupItem(prod, cb) {
    if (!_pickupActivo || !Array.isArray(_pickupActivo.items)) return null;
    const cbU  = String(cb || '').toUpperCase();
    const idP  = String(prod?.idProducto || '').toUpperCase();
    const skuP = String(prod?.skuBase    || '').toUpperCase();
    return _pickupActivo.items.find(it => {
      if (!it) return false;
      const itSku = String(it.skuBase || '').toUpperCase();
      if (itSku && (itSku === idP || (skuP && itSku === skuP))) return true;
      if (Array.isArray(it.codigosOriginales) &&
          it.codigosOriginales.some(c => String(c).toUpperCase() === cbU)) return true;
      return false;
    }) || null;
  }

  // ¿Producto por peso (granel)? Solo la unidad de medida lo determina.
  // KGM (kilogramos, estándar SUNAT) y variantes locales. Es independiente de
  // si es envasable o no — un granel puede ser canónico (ajo entero pelado)
  // o un derivado que también se vende a granel (ajo en polvo). Ambos KGM.
  // Nota: factorConversion ≠ 1 indica envasado/derivado, NO determina decimal.
  function _esProductoPeso(prod) {
    if (!prod) return false;
    const u = String(prod.unidad || '').toUpperCase().trim();
    return u === 'KGM' || u === 'KG' || u === 'KGS' || u === 'GMS' || u === 'G';
  }

  // ═══════════════════════════════════════════════════════════════════════
  // REGLA DE ORO WH (en piedra) — manejo de codigos:
  //
  // ESCANEAR (matching):  acepta skuBase O codigoBarra canónico O codigoBarra
  //                       de equivalencia activa. Cualquier match suma al MISMO
  //                       item del pickup (agrupado por skuBase del producto).
  //
  // PROGRESO (memoria):   despachadoPorCodigo guarda cuántas unidades se
  //                       escanearon por cada codigoBarra REAL. Ejemplo:
  //                         { '6959749711163': 4, 'EAN-EQUIV-001': 2 }
  //
  // GUIA_SALIDA:          al cerrar, se desglosa el despacho en filas con
  //                       codigoBarra REAL (canónico o equivalente). NUNCA
  //                       skuBase porque el stock se descuenta por codigoBarra
  //                       específico, no por agrupador. Un producto puede
  //                       tener stock distinto en su canónico vs sus
  //                       equivalentes y debe descontarse del que se escaneó.
  // ═══════════════════════════════════════════════════════════════════════

  // Intenta sumar al item del pickup que matchee. Retorna true si lo absorbió.
  // Para productos por kg, abre input modal de cantidad. Para sobrescaneado,
  // muestra modal de confirmación. Suma normal con tope = solicitado.
  function _intentarSumarAPickup(prod, cb) {
    const item = _matchPickupItem(prod, cb);
    if (!item) return false;

    const sol      = parseFloat(item.solicitado) || 0;
    const prevDesp = parseFloat(item.despachado) || 0;

    // Marcar este item como el ACTIVO + recordar el código escaneado (los
    // botones +/- y el input granel operan sobre item._ultimoCb).
    item._ultimoCb = cb;
    _pickupItemActivo = item.skuBase;

    // ── Producto a granel (KGM): focus inline al input de cantidad ───
    // Antes abría un modal. Ahora la card activa muestra un input decimal
    // con focus automático — el operador pesa y escribe directo.
    if (_esProductoPeso(prod)) {
      _savePickup();
      _renderPickupChecklistInSheet(item.skuBase);
      _renderPickupChecklistInline();
      _renderPickupActivoBanner();
      _updateGenerarBtn();
      badgeUpdate();
      try { SoundFX.beep(); } catch(_){}
      vibrate(10);
      // Foco al input granel de la card activa
      setTimeout(() => {
        try {
          const inp = document.querySelector('.pkck-card[data-sku="' + (window.CSS && CSS.escape ? CSS.escape(String(item.skuBase)) : String(item.skuBase)) + '"] .pkck-ctrl-granel');
          if (inp) { inp.focus(); inp.select(); }
        } catch(_){}
      }, 70);
      return true;
    }

    // ── Detectar sobrescaneado y advertir ─────────────────────
    if (prevDesp >= sol && sol > 0) {
      _abrirModalSobrescaneado(item, cb, prevDesp, sol, prod);
      return true; // detenemos aquí, modal decide
    }

    item.despachado = prevDesp + 1;
    item.despachadoPorCodigo = item.despachadoPorCodigo || {};
    item.despachadoPorCodigo[cb] = (parseFloat(item.despachadoPorCodigo[cb]) || 0) + 1;
    _savePickup();

    const justCompletado = prevDesp < sol && item.despachado >= sol;
    const sobreDespacho  = item.despachado > sol;

    // Re-render checklist inline + banner activo + botón generar
    _renderPickupChecklistInSheet(item.skuBase);
    _renderPickupChecklistInline();
    _renderPickupActivoBanner();
    _renderExtrasSection();
    _updateGenerarBtn();
    badgeUpdate();

    // Sonidos + animaciones según estado
    if (justCompletado) {
      try { SoundFX.pickupItemOk(); } catch(_){}
      vibrate([20, 30, 60]);
      _flashItemCompleto(item.skuBase);
      // Si TODOS los items llegaron al 100% → celebración
      if (_pickupTotalmenteCompleto()) {
        setTimeout(() => {
          try { SoundFX.pickupOk(); } catch(_){}
          _confettiCelebracion();
          _setDespStatus('ok', '🎉 ¡Pickup completo! Pulsa "Cerrar despacho"');
        }, 280);
      }
    } else if (sobreDespacho) {
      try { SoundFX.warn(); } catch(_){}
      vibrate([30, 20, 30]);
    } else {
      try { SoundFX.beep(); } catch(_){}
      vibrate(12);
    }

    _scheduleAutosavePickup();
    return true;
  }

  // ── Controles +/- e input granel del item ACTIVO del checklist ──────
  // Operan sobre item._ultimoCb (el último código escaneado de ese item).
  // Re-render + sonidos + autosave compartidos en _afterPickupChange.
  function _pkckItemPorSku(skuBase) {
    if (!_pickupActivo) return null;
    return (_pickupActivo.items || []).find(it => String(it.skuBase) === String(skuBase)) || null;
  }
  function _pkckCodigoDe(item) {
    return item._ultimoCb
        || (item.despachadoPorCodigo && Object.keys(item.despachadoPorCodigo)[0])
        || (item.codigosOriginales && item.codigosOriginales[0])
        || item.skuBase;
  }
  function _afterPickupChange(item, prevDesp, sol) {
    _savePickup();
    _renderPickupChecklistInSheet(item.skuBase);
    _renderPickupChecklistInline();
    _renderPickupActivoBanner();
    _renderExtrasSection();
    _updateGenerarBtn();
    badgeUpdate();
    const desp = parseFloat(item.despachado) || 0;
    const justCompletado = prevDesp < sol && desp >= sol && sol > 0;
    try {
      if (justCompletado) {
        SoundFX.pickupItemOk(); _flashItemCompleto(item.skuBase);
        if (_pickupTotalmenteCompleto()) {
          setTimeout(() => {
            try { SoundFX.pickupOk(); } catch(_){}
            _confettiCelebracion();
            _setDespStatus('ok', '🎉 ¡Pickup completo! Pulsa "Cerrar despacho"');
          }, 280);
        }
      } else { SoundFX.beep(); }
    } catch(_){}
    vibrate(justCompletado ? [20, 30, 60] : 12);
    _scheduleAutosavePickup();
  }
  async function _pkckMas(skuBase) {
    const item = _pkckItemPorSku(skuBase);
    if (!item) return;
    const sol      = parseFloat(item.solicitado) || 0;
    const prevDesp = parseFloat(item.despachado) || 0;
    if (prevDesp >= sol && sol > 0) {
      if (!await _whConfirm(`Ya llevas ${prevDesp}/${sol} de "${item.nombre || skuBase}".\n\n¿Sumar 1 más (sobre-despacho)?`, { warning: true, titulo: 'Sobre-despacho' })) return;
    }
    const cb = _pkckCodigoDe(item);
    item._ultimoCb = cb;
    item.despachado = prevDesp + 1;
    item.despachadoPorCodigo = item.despachadoPorCodigo || {};
    item.despachadoPorCodigo[cb] = (parseFloat(item.despachadoPorCodigo[cb]) || 0) + 1;
    _pickupItemActivo = item.skuBase;
    _afterPickupChange(item, prevDesp, sol);
  }
  function _pkckMenos(skuBase) {
    const item = _pkckItemPorSku(skuBase);
    if (!item) return;
    const prevDesp = parseFloat(item.despachado) || 0;
    if (prevDesp <= 0) return;
    const cb = _pkckCodigoDe(item);
    item.despachado = Math.max(0, prevDesp - 1);
    if (item.despachadoPorCodigo && item.despachadoPorCodigo[cb] != null) {
      const nv = (parseFloat(item.despachadoPorCodigo[cb]) || 0) - 1;
      if (nv > 0) item.despachadoPorCodigo[cb] = nv;
      else delete item.despachadoPorCodigo[cb];
    }
    _pickupItemActivo = item.skuBase; // sigue activo aunque quede en 0
    _afterPickupChange(item, prevDesp, parseFloat(item.solicitado) || 0);
  }
  function _pkckSetGranel(skuBase, valorRaw) {
    const item = _pkckItemPorSku(skuBase);
    if (!item) return;
    const qty = parseFloat(String(valorRaw).replace(',', '.'));
    if (isNaN(qty) || qty < 0) return;
    const sol      = parseFloat(item.solicitado) || 0;
    const prevDesp = parseFloat(item.despachado) || 0;
    const cb = _pkckCodigoDe(item);
    // Granel: el input SETEA el total pesado (no acumula). 1 código domina.
    item._ultimoCb = cb;
    item.despachado = qty;
    item.despachadoPorCodigo = item.despachadoPorCodigo || {};
    item.despachadoPorCodigo[cb] = qty;
    _pickupItemActivo = item.skuBase;
    _afterPickupChange(item, prevDesp, sol);
  }

  // ¿Todos los items del pickup están al 100%?
  function _pickupTotalmenteCompleto() {
    if (!_pickupActivo) return false;
    return (_pickupActivo.items || []).every(it =>
      (parseFloat(it.despachado) || 0) >= (parseFloat(it.solicitado) || 0)
    );
  }

  // Autosave a backend con debounce 4s. Llama guardarProgresoPickup con lock.
  function _scheduleAutosavePickup() {
    if (_autosavePickupTimer) clearTimeout(_autosavePickupTimer);
    const lblEl = document.getElementById('pkactAutosaveLbl');
    const wrEl  = document.getElementById('pkactAutosave');
    if (lblEl) lblEl.textContent = 'guardando...';
    if (wrEl)  wrEl.classList.add('is-saving');
    _autosavePickupTimer = setTimeout(async () => {
      if (!_pickupActivo) return;
      try {
        const r = await API.guardarProgresoPickup({
          idPickup:    _pickupActivo.idPickup,
          items:       _pickupActivo.items,
          lockUsuario: window.WH_CONFIG?.usuario || ''
        });
        if (r && r.ok === false && r.conflicto) {
          if (lblEl) lblEl.textContent = `⚠ atiende ${r.atendidoPor}`;
          toast(`🔒 ${r.atendidoPor} también está atendiendo este pickup. Tu progreso puede no guardarse.`, 'warn', 6000);
          return;
        }
        _autosaveLastTs = Date.now();
        if (lblEl) lblEl.textContent = 'guardado';
        if (wrEl)  wrEl.classList.remove('is-saving');
      } catch (e) {
        if (lblEl) lblEl.textContent = 'reintentando...';
      }
    }, 4000);
  }

  // ── Modal: producto a granel (KGM) — pedir peso decimal ──────
  // Aplica a TODOS los productos cuya unidad sea KGM (granel), sean
  // canónicos (ajo entero pelado) o derivados a granel (ajo en polvo).
  // No tiene relación con "envasable": un envasable común se cuenta en NIU.
  function _abrirModalQtyGranel(prod, item, cb, prevDesp, sol) {
    const overlay = document.getElementById('pkckModalQtyKg');
    if (!overlay) return;
    const u = String(prod.unidad || 'kg').toLowerCase();
    document.getElementById('pkckModalQtyName').textContent = prod.descripcion || item.nombre || cb;
    document.getElementById('pkckModalQtyMeta').textContent =
      `Pendiente: ${fmt(sol - prevDesp, 3)} ${u} · Ya despachado: ${fmt(prevDesp, 3)} ${u}`;
    // Cambiar el sufijo del botón confirmar a la unidad real
    const okBtn = document.getElementById('pkckModalQtyOk');
    if (okBtn) okBtn.textContent = 'Confirmar ' + u;
    const input = document.getElementById('pkckModalQtyInput');
    input.value = '';
    overlay.style.display = 'flex';
    setTimeout(() => input.focus(), 80);

    const _close = () => { overlay.style.display = 'none'; };
    document.getElementById('pkckModalQtyCancel').onclick = _close;
    okBtn.onclick = async () => {
      const qty = parseFloat(String(input.value).replace(',', '.'));
      if (!qty || qty <= 0) { toast('Ingresa un peso válido', 'warn'); return; }
      const nuevoTotal = (parseFloat(item.despachado) || 0) + qty;
      // Sobre-despacho granel: si supera lo solicitado, pedir confirmación
      if (nuevoTotal > sol && sol > 0) {
        const exceso = nuevoTotal - sol;
        const msg = `Estás pesando ${fmt(qty,3)} ${u} pero faltaba solo ${fmt(sol - prevDesp,3)} ${u}.\n\nNuevo total: ${fmt(nuevoTotal,3)} ${u} (${fmt(exceso,3)} ${u} de más)\n\n¿Confirmar igual?`;
        if (!await _whConfirm(msg, { warning: true })) {
          input.select();
          try { SoundFX.warn(); } catch(_){}
          return;
        }
      }
      item.despachado = nuevoTotal;
      item.despachadoPorCodigo = item.despachadoPorCodigo || {};
      item.despachadoPorCodigo[cb] = (parseFloat(item.despachadoPorCodigo[cb]) || 0) + qty;
      _savePickup();
      _renderPickupChecklistInSheet(item.skuBase);
      _renderPickupActivoBanner();
      badgeUpdate();
      const justCompletado = (parseFloat(prevDesp) < parseFloat(sol)) && (item.despachado >= sol);
      try {
        if (justCompletado) { SoundFX.pickupItemOk(); _flashItemCompleto(item.skuBase); _checkConfettiSiCompleto(); }
        else                { SoundFX.beep(); }
      } catch(_){}
      vibrate(justCompletado ? [20,30,60] : 12);
      _scheduleAutosavePickup();
      _close();
    };
    input.onkeydown = (e) => { if (e.key === 'Enter') okBtn.click(); };
  }

  // ── Modal: sobrescaneado — confirmar o deshacer ──
  function _abrirModalSobrescaneado(item, cb, prevDesp, sol, prod) {
    const overlay = document.getElementById('pkckModalSobre');
    if (!overlay) return;
    document.getElementById('pkckModalSobreName').textContent = item.nombre || item.skuBase;
    document.getElementById('pkckModalSobreMeta').textContent =
      `Pedías ${sol} · llevas ${prevDesp}. Si confirmas, sumarás 1 más.`;
    overlay.style.display = 'flex';
    try { SoundFX.warn(); } catch(_){}
    vibrate([30, 20, 30]);
    const _close = () => { overlay.style.display = 'none'; };
    document.getElementById('pkckModalSobreCancel').onclick = _close;
    document.getElementById('pkckModalSobreOk').onclick = () => {
      item.despachado = (parseFloat(item.despachado) || 0) + 1;
      item.despachadoPorCodigo = item.despachadoPorCodigo || {};
      item.despachadoPorCodigo[cb] = (parseFloat(item.despachadoPorCodigo[cb]) || 0) + 1;
      _savePickup();
      _renderPickupChecklistInSheet(item.skuBase);
      _renderPickupActivoBanner();
      badgeUpdate();
      try { SoundFX.beep(); } catch(_){}
      vibrate(15);
      _scheduleAutosavePickup();
      _close();
    };
  }

  // ── Modal granel para EXTRAS (fuera de pickup o sin pickup activo) ──
  // Mismo modal qty pero suma al _cart en vez del pickup activo.
  function _abrirModalQtyGranelExtra(prod, cb) {
    const overlay = document.getElementById('pkckModalQtyKg');
    if (!overlay) return;
    const u = String(prod.unidad || 'kg').toLowerCase();
    document.getElementById('pkckModalQtyName').textContent = prod.descripcion || cb;
    const enPickup = _pickupActivo ? ' · fuera de pickup, irá como extra' : '';
    document.getElementById('pkckModalQtyMeta').textContent =
      `Producto a granel${enPickup}. Pesa el producto y registra el peso real.`;
    const okBtn = document.getElementById('pkckModalQtyOk');
    if (okBtn) okBtn.textContent = 'Confirmar ' + u;
    const input = document.getElementById('pkckModalQtyInput');
    input.value = '';
    overlay.style.display = 'flex';
    setTimeout(() => input.focus(), 80);

    const _close = () => { overlay.style.display = 'none'; };
    document.getElementById('pkckModalQtyCancel').onclick = _close;
    okBtn.onclick = () => {
      const qty = parseFloat(String(input.value).replace(',', '.'));
      if (!qty || qty <= 0) { toast('Ingresa un peso válido', 'warn'); return; }
      const stockMap = {};
      OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });
      const stockD = parseFloat((stockMap[prod.idProducto] || stockMap[cb] || {}).cantidadDisponible || 0);
      const idx = _cart.findIndex(c => c.codigoBarra === cb);
      if (idx >= 0) {
        _cart[idx].cantidad = (parseFloat(_cart[idx].cantidad) || 0) + qty;
      } else {
        _cart.push({
          codigoBarra: cb, descripcion: prod.descripcion || cb,
          unidad: prod.unidad || '', cantidad: qty, stockDisp: stockD,
          _extraPickup: !!_pickupActivo, _granel: true
        });
      }
      _saveCart();
      _renderCart();
      _renderDespList && _renderDespList();
      _updateFooter();
      badgeUpdate();
      try { SoundFX.beep(); } catch(_){}
      vibrate(15);
      if (_pickupActivo) {
        toast(`🟠 ${fmt(qty,3)} ${u} de ${prod.descripcion || cb} · fuera de pickup, irá como extra`, 'warn', 3000);
      } else {
        toast(`✓ ${fmt(qty,3)} ${u} de ${prod.descripcion || cb}`, 'ok', 2000);
      }
      _close();
    };
    input.onkeydown = (e) => { if (e.key === 'Enter') okBtn.click(); };
  }

  // Helper invocado al confirmar qty kg — chequea celebración
  function _checkConfettiSiCompleto() {
    if (_pickupTotalmenteCompleto()) {
      setTimeout(() => {
        try { SoundFX.pickupOk(); } catch(_){}
        _confettiCelebracion();
        _setDespStatus && _setDespStatus('ok', '🎉 ¡Pickup completo! Pulsa "Cerrar despacho"');
      }, 280);
    }
  }

  // Render del checklist dentro del sheet — UI moderna con barras de progreso
  // Estado de filtro/orden del checklist (para #5)
  let _pickupSearch = '';
  let _pickupSort   = 'pendientes'; // 'pendientes' | 'az' | 'zona'

  function _renderPickupChecklistInSheet(flashSkuBase) {
    const cont = document.getElementById('despPickChecklist');
    if (!cont || !_pickupActivo) return;
    const items = _pickupActivo.items || [];
    // Si el item activo ya no pertenece a este pickup (cambió de pickup,
    // se cerró, etc.) → limpiar para que no queden controles huérfanos.
    if (_pickupItemActivo && !items.some(it => String(it.skuBase) === String(_pickupItemActivo))) {
      _pickupItemActivo = null;
    }
    // Stock cache indexado por codigoProducto Y idProducto (un row puede
    // ser accedido por cualquiera de los dos).
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => {
      if (s.codigoProducto) stockMap[String(s.codigoProducto)] = s;
      if (s.idProducto)     stockMap[String(s.idProducto)]     = s;
    });
    const productos = (App.getProductosMaestro && App.getProductosMaestro()) || [];

    // Helper: SUMAR stock de todos los códigos (canónico + equivalentes).
    function _buscarStockSheet(item) {
      let total = 0;
      let count = 0;
      const visited = new Set();
      const tryKey = (k) => {
        if (!k) return;
        const key = String(k);
        if (visited.has(key)) return;
        visited.add(key);
        const r = stockMap[key];
        if (r) { total += parseFloat(r.cantidadDisponible || 0); count++; }
      };
      if (Array.isArray(item.codigosOriginales)) {
        item.codigosOriginales.forEach(c => tryKey(c));
      }
      if (count === 0) tryKey(item.skuBase);
      return count > 0 ? { cantidadDisponible: total, codigosEncontrados: count } : null;
    }

    // Filtrar por búsqueda + ordenar
    const q = _pickupSearch.trim().toLowerCase();
    let filtered = items.filter(it => {
      if (!q) return true;
      return String(it.nombre || '').toLowerCase().includes(q) ||
             String(it.skuBase || '').toLowerCase().includes(q);
    });
    filtered = filtered.slice().sort((a, b) => {
      if (_pickupSort === 'az') return String(a.nombre).localeCompare(String(b.nombre));
      if (_pickupSort === 'zona') {
        const za = String((productos.find(p => String(p.idProducto) === String(a.skuBase)) || {}).zona || 'zzz');
        const zb = String((productos.find(p => String(p.idProducto) === String(b.skuBase)) || {}).zona || 'zzz');
        if (za !== zb) return za.localeCompare(zb);
        return String(a.nombre).localeCompare(String(b.nombre));
      }
      // Pendientes primero (incompletos arriba, completos abajo)
      const aPend = (parseFloat(a.despachado)||0) < (parseFloat(a.solicitado)||0) ? 0 : 1;
      const bPend = (parseFloat(b.despachado)||0) < (parseFloat(b.solicitado)||0) ? 0 : 1;
      if (aPend !== bPend) return aPend - bPend;
      return String(a.nombre).localeCompare(String(b.nombre));
    });

    cont.innerHTML = filtered.map(it => {
      const sol  = parseFloat(it.solicitado) || 0;
      const desp = parseFloat(it.despachado) || 0;
      const pct  = sol > 0 ? Math.min(100, Math.round((desp / sol) * 100)) : 0;
      const completo  = desp >= sol && sol > 0;
      const enProg    = desp > 0 && desp < sol;
      const cls       = completo ? 'is-completo' : (enProg ? 'is-progreso' : '');
      const flash     = (flashSkuBase && String(flashSkuBase) === String(it.skuBase)) ? 'is-just-completed is-flash' : '';
      const stockRow  = _buscarStockSheet(it);
      const stockD    = stockRow ? parseFloat(stockRow.cantidadDisponible || 0) : 0;
      const stockKnown= !!stockRow;
      const equivCount= Array.isArray(it.codigosOriginales) ? it.codigosOriginales.length : 0;
      const equivTxt  = equivCount > 1 ? ` · ${equivCount} códigos` : '';
      const icon      = completo ? '✓' : (enProg ? '⏳' : '📦');
      // Producto maestro para detectar si es a granel (despacho decimal)
      const prod = productos.find(p => String(p.idProducto) === String(it.skuBase));
      const esKg = _esProductoPeso(prod);
      const unidadLbl = esKg ? String(prod.unidad || 'kg').toLowerCase() : '';
      // Stock badge — rojo si stockDisp < solicitado pendiente
      let stockBadge = '';
      const pendiente = sol - desp;
      if (!stockKnown) {
        stockBadge = '<span style="font-size:.62em;color:#94a3b8;background:rgba(71,85,105,.25);padding:1px 6px;border-radius:6px" title="No hay stock cacheado para este producto en WH">sin info</span>';
      } else if (stockD <= 0) {
        stockBadge = '<span style="font-size:.62em;color:#fca5a5;background:rgba(220,38,38,.18);border:1px solid rgba(239,68,68,.4);padding:1px 6px;border-radius:6px;font-weight:800">⚠ stock 0</span>';
      } else if (stockD < pendiente) {
        stockBadge = `<span style="font-size:.62em;color:#fca5a5;background:rgba(220,38,38,.18);border:1px solid rgba(239,68,68,.4);padding:1px 6px;border-radius:6px;font-weight:800">⚠ stock ${fmt(stockD,1)} · faltarán ${fmt(pendiente - stockD,1)}</span>`;
      } else {
        stockBadge = `<span style="font-size:.62em;color:#86efac;background:rgba(16,185,129,.15);padding:1px 6px;border-radius:6px;font-weight:700">stock ${fmt(stockD,1)}</span>`;
      }
      const kgBadge = esKg
        ? `<span style="font-size:.6em;color:#fbbf24;background:rgba(245,158,11,.15);padding:1px 5px;border-radius:6px;margin-left:4px;font-weight:800;letter-spacing:.04em" title="Producto a granel · despacho por peso">⚖ GRANEL · ${unidadLbl}</span>`
        : '';
      // ── Controles: SOLO el item activo (último escaneado) los muestra ──
      // El resto los oculta para que el operador no edite por error otro
      // producto. Granel → input decimal con focus. Normal → botones +/-.
      const esActivo = _pickupItemActivo && String(_pickupItemActivo) === String(it.skuBase);
      const skuAttr  = escAttr(it.skuBase);
      let ctrlsHtml = '';
      if (esActivo) {
        if (esKg) {
          ctrlsHtml = `
          <div class="pkck-ctrls">
            <span class="pkck-ctrl-lbl">Peso despachado:</span>
            <input type="number" inputmode="decimal" step="0.001" min="0"
                   class="pkck-ctrl-granel" value="${fmt(desp,3)}"
                   onchange="DespachoView.pkckSetGranel('${skuAttr}', this.value)"
                   onkeydown="if(event.key==='Enter'){this.blur();}">
            <span class="pkck-ctrl-unit">${unidadLbl}</span>
          </div>`;
        } else {
          ctrlsHtml = `
          <div class="pkck-ctrls">
            <button class="pkck-ctrl-btn pkck-ctrl-menos" ${desp <= 0 ? 'disabled' : ''}
                    onclick="DespachoView.pkckMenos('${skuAttr}')">−</button>
            <span class="pkck-ctrl-val">${desp}</span>
            <button class="pkck-ctrl-btn pkck-ctrl-mas"
                    onclick="DespachoView.pkckMas('${skuAttr}')">+</button>
          </div>`;
        }
      }
      return `
        <div class="pkck-card ${cls} ${flash} ${esActivo ? 'is-activo' : ''}" data-sku="${skuAttr}">
          <div class="pkck-row">
            <div class="pkck-icon">${icon}</div>
            <div class="flex-1 min-w-0">
              <p class="pkck-name truncate">${escHtml(it.nombre || it.skuBase)}${kgBadge}</p>
              <p class="pkck-meta truncate">${escHtml(it.skuBase)}${equivTxt} · ${stockBadge}</p>
            </div>
            <div class="pkck-qty-wrap">
              <p><span class="pkck-qty">${esKg ? fmt(desp,3) : desp}</span><span class="pkck-qty-sol"> / ${esKg ? fmt(sol,3) : sol}${esKg ? ' '+unidadLbl : ''}</span></p>
              <p class="text-[10px] text-slate-500 mt-0.5">${pct}%</p>
            </div>
          </div>
          <div class="pkck-bar-wrap">
            <div class="pkck-bar-fill" style="width:${pct}%"></div>
            ${enProg ? '<div class="pkck-bar-shimmer"></div>' : ''}
          </div>
          ${ctrlsHtml}
          <div class="pkck-check-overlay">✓</div>
        </div>`;
    }).join('') || `<div class="text-center text-slate-500 text-sm py-6">Sin coincidencias para "${escHtml(q)}"</div>`;

    // KPIs hero
    const totalUds  = items.reduce((s,it) => s + (parseFloat(it.solicitado) || 0), 0);
    const totalDesp = items.reduce((s,it) => s + (parseFloat(it.despachado) || 0), 0);
    const pctTot    = totalUds > 0 ? Math.round((totalDesp / totalUds) * 100) : 0;
    const elIt = document.getElementById('pkckKpiItems');
    const elUd = document.getElementById('pkckKpiUds');
    const elPc = document.getElementById('pkckKpiPct');
    if (elIt) elIt.textContent = items.length;
    if (elUd) elUd.textContent = `${totalDesp}/${totalUds}`;
    if (elPc) elPc.textContent = pctTot + '%';

    // CTA dinámico
    const cta  = document.getElementById('btnConfirmarPickup');
    const hint = document.getElementById('pkckFootHint');
    if (cta) {
      const completoTotal = _pickupTotalmenteCompleto();
      const tieneAlgo     = totalDesp > 0;
      // Estado del pickup en backend
      const yaEmpezo = String(_pickupActivo.estado || '').toUpperCase() === 'EN_PROCESO';
      if (!yaEmpezo && !tieneAlgo) {
        cta.className = 'pkck-cta';
        cta.innerHTML = '<span style="font-size:1.2em">▶</span><span>EMPEZAR DESPACHO</span>';
        cta.onclick = () => DespachoView.empezarPickup();
        if (hint) hint.textContent = 'Escanea cada producto · los equivalentes cuentan automáticamente';
      } else if (completoTotal) {
        cta.className = 'pkck-cta is-cerrar';
        cta.innerHTML = '<span style="font-size:1.2em">✓</span><span>CERRAR DESPACHO COMPLETO</span>';
        cta.onclick = () => DespachoView.cerrarDespachoPickup(false);
        if (hint) hint.textContent = '¡Todos los items despachados!';
      } else if (tieneAlgo) {
        cta.className = 'pkck-cta is-parcial';
        const falt = items.filter(it => (parseFloat(it.despachado)||0) < (parseFloat(it.solicitado)||0)).length;
        cta.innerHTML = `<span style="font-size:1.2em">↗</span><span>CERRAR PARCIAL · faltan ${falt}</span>`;
        cta.onclick = () => DespachoView.cerrarDespachoPickup(true);
        if (hint) hint.textContent = 'Los faltantes irán en la observación de la guía';
      } else {
        cta.className = 'pkck-cta';
        cta.innerHTML = '<span style="font-size:1.2em">▶</span><span>EMPEZAR DESPACHO</span>';
        cta.onclick = () => DespachoView.empezarPickup();
      }
    }
  }

  // Banner del pickup activo en view-despacho (visible siempre que se está despachando)
  function _renderPickupActivoBanner() {
    const banner = document.getElementById('despPickupActivoBanner');
    if (!banner) return;
    if (!_pickupActivo) { banner.style.display = 'none'; return; }
    const items     = _pickupActivo.items || [];
    const totalUds  = items.reduce((s,it) => s + (parseFloat(it.solicitado) || 0), 0);
    const totalDesp = items.reduce((s,it) => s + (parseFloat(it.despachado) || 0), 0);
    const pct       = totalUds > 0 ? Math.round((totalDesp / totalUds) * 100) : 0;
    const completo  = _pickupTotalmenteCompleto();
    banner.style.display = 'block';
    banner.classList.toggle('is-completo', completo);
    const elZ = document.getElementById('pkactZona');
    const elP = document.getElementById('pkactProgresoLbl');
    const elB = document.getElementById('pkactBar');
    if (elZ) elZ.textContent = '📦 ' + (_pickupActivo.idZona || _pickupActivo.idPickup || 'pickup');
    if (elP) elP.textContent = `${totalDesp}/${totalUds} uds · ${pct}%`;
    if (elB) elB.style.width = pct + '%';
  }

  // Flash visual sobre la card de un item recién completado
  function _flashItemCompleto(skuBase) {
    requestAnimationFrame(() => {
      const card = document.querySelector(`.pkck-card[data-sku="${CSS.escape(String(skuBase))}"]`);
      if (!card) return;
      card.classList.add('is-just-completed', 'is-flash');
      setTimeout(() => card.classList.remove('is-just-completed', 'is-flash'), 950);
    });
  }

  // Confetti celebración cuando pickup llega al 100% total
  function _confettiCelebracion() {
    const colores = ['#10b981','#34d399','#6366f1','#818cf8','#f59e0b','#fbbf24','#ec4899'];
    const n = 36;
    for (let i = 0; i < n; i++) {
      const c = document.createElement('div');
      c.className = 'pkck-confetti';
      c.style.left  = (Math.random() * 100) + 'vw';
      c.style.background = colores[i % colores.length];
      c.style.setProperty('--tx',  ((Math.random() - 0.5) * 200) + 'px');
      c.style.setProperty('--rot', (Math.random() * 1080) + 'deg');
      c.style.setProperty('--dur', (1.8 + Math.random() * 1.6) + 's');
      c.style.animationDelay = (Math.random() * 0.3) + 's';
      document.body.appendChild(c);
      setTimeout(() => c.remove(), 4000);
    }
  }

  // Abrir el sheet del pickup activo (botón "Detalle" en banner)
  function abrirSheetPickupActivo() {
    if (!_pickupActivo) return;
    _renderPickupChecklistInSheet();
    _renderPickupActivoBanner();
    document.getElementById('despPickTitulo').textContent =
      'Pickup ' + (_pickupActivo.idPickup || '');
    document.getElementById('despPickSub').textContent =
      `${_pickupActivo.idZona || '—'} · ${_pickupActivo.creadoPor || 'cajero'}`;
    // Ocultar secciones legacy
    document.getElementById('despPickSecExactos').classList.add('hidden');
    document.getElementById('despPickSecParciales').classList.add('hidden');
    document.getElementById('despPickSecNoFound').classList.add('hidden');
    abrirSheet('sheetDespPickup');
  }

  // Empezar el despacho de un pickup: marca EN_PROCESO, persiste, cierra sheet,
  // muestra banner y deja al operador escaneando.
  async function empezarPickup() {
    if (!_pickupActivo) return;
    // Inicializar estructura si viene fresca
    _pickupActivo.items = (_pickupActivo.items || []).map(it => ({
      skuBase:           it.skuBase,
      nombre:            it.nombre || it.skuBase,
      solicitado:        parseFloat(it.solicitado) || 0,
      despachado:        parseFloat(it.despachado) || 0,
      codigosOriginales: Array.isArray(it.codigosOriginales) ? it.codigosOriginales : [],
      despachadoPorCodigo: it.despachadoPorCodigo || {}
    }));
    _pickupActivo.estado = 'EN_PROCESO';
    const usuario = window.WH_CONFIG?.usuario || '';
    _pickupActivo.atendidoPor = usuario;
    _savePickup();

    // Backend: marcar EN_PROCESO + tomar lock atendidoPor explícitamente.
    // Sin tomarLock=true el atendidoPor queda vacío y la card sigue
    // mostrándose como "▶ Jalar" para todos en el siguiente poll.
    API.actualizarPickup({
      idPickup:    _pickupActivo.idPickup,
      estado:      'EN_PROCESO',
      lockUsuario: usuario,
      tomarLock:   true
    }).catch(() => {});
    // Quitar de pendientes localmente — el polling lo traerá de vuelta como
    // EN_PROCESO con atendidoPor=yo, y la UI mostrará "↻ Continuar (yo)".
    _pickupsPendientes = _pickupsPendientes.filter(p => p.idPickup !== _pickupActivo.idPickup);

    cerrarSheet('sheetDespPickup');
    _renderPickupActivoBanner();
    _renderCart();
    _updateFooter();
    badgeUpdate();
    try { SoundFX.beepDouble(); } catch(_){}
    vibrate(20);
    toast(`▶ Despacho iniciado · escanea para registrar`, 'ok', 2800);
  }

  // Cerrar despacho del pickup activo: emite GUIA_SALIDA con códigos reales.
  // esParcial=true si el operador quiere cerrar antes de tener todo.
  //
  // Lock _pickupClosing previene cierres concurrentes: si el operador
  // hace doble-click o el botón sigue clickeable durante la espera, las
  // llamadas extra retornan sin tocar el backend. Sin esto se generan
  // múltiples GUIA_SALIDA con items duplicados (bug histórico).
  async function cerrarDespachoPickup(esParcial) {
    if (!_pickupActivo) return;
    if (_pickupClosing) {
      toast('⏳ Ya se está procesando el cierre · espera...', 'warn', 2500);
      return;
    }

    // PRE-CHECK de versión: garantiza que el código que va a procesar el
    // cierre es la última publicada. Si la PWA está atascada con versión
    // vieja, esto bloquea + fuerza reload antes de que se generen
    // duplicados u otros bugs corregidos en versiones posteriores.
    if (typeof window._preCheckVersion === 'function') {
      const ok = await window._preCheckVersion();
      if (!ok) {
        toast('🔄 Actualizando · evitando duplicados', 'info', 4000);
        return;
      }
    }

    // Si parcial, confirmar con el operador
    const completo = _pickupTotalmenteCompleto();
    if (esParcial && !completo) {
      const items = _pickupActivo.items || [];
      const falt  = items.filter(it => (parseFloat(it.despachado)||0) < (parseFloat(it.solicitado)||0)).length;
      if (!await _whConfirm(`¿Cerrar despacho con ${falt} item${falt!==1?'s':''} sin completar?\n\nLos faltantes quedarán en observación de la guía.`, { warning: true, titulo: 'Cerrar despacho parcial', okText: 'Cerrar' })) return;
    }

    // Construir despachoDetalle desde despachadoPorCodigo de cada item
    const despachoDetalle = [];
    (_pickupActivo.items || []).forEach(it => {
      const dpc = it.despachadoPorCodigo || {};
      Object.keys(dpc).forEach(cod => {
        const qty = parseFloat(dpc[cod]) || 0;
        if (qty > 0) despachoDetalle.push({ codigoBarra: String(cod), cantidad: qty });
      });
    });
    // Sumar también extras del carrito (escaneos fuera de pickup)
    _cart.filter(c => c._extraPickup).forEach(c => {
      despachoDetalle.push({ codigoBarra: c.codigoBarra, cantidad: parseFloat(c.cantidad) || 0 });
    });

    if (!despachoDetalle.length) {
      toast('No hay nada despachado todavía', 'warn', 2500);
      return;
    }

    // Tomar el lock ANTES de cualquier UI optimista y await
    _pickupClosing = true;
    _setGenerarGuiaBusy(true);

    // Snapshot para revert si falla
    const pickupSnap  = JSON.parse(JSON.stringify(_pickupActivo));
    const cartSnap    = [..._cart];

    // UI optimista: limpiar todo y mostrar feedback
    cerrarSheet('sheetDespPickup');
    _cart = [];
    _saveCart();
    _renderCart();
    _updateFooter();
    toast('⏳ Generando guía de despacho...', 'info', 55000);

    try {
      const res = await Promise.race([
        API.cerrarPickupConDespacho({
          idPickup:        _pickupActivo.idPickup,
          items:           _pickupActivo.items,
          despachoDetalle: despachoDetalle,
          usuario:         window.WH_CONFIG?.usuario || '',
          imprimir:        false
        }),
        new Promise((_, rej) => setTimeout(() => rej({ timeout: true }), 55000))
      ]);
      if (res.ok) {
        const d = res.data || {};
        try { SoundFX.pickupOk(); } catch(_){}
        vibrate([30, 20, 60, 20, 100]);
        _confettiCelebracion();
        toast(`✅ Guía ${d.idGuia || ''} · ${d.estado || 'COMPLETADO'} · ${d.despachados || 0} líneas`, 'ok', 6000);

        // [v2.13.31] PRINT con reintento auto + cola persistente offline.
        // Antes: 1 solo intento fire-and-forget → si PrintNode/red fallaba,
        // user perdía el toast warn y el ticket nunca salía. Ahora:
        //   - Hasta 3 intentos con backoff 1s/3s/8s
        //   - Si TODOS fallan, encola en localStorage para reintento posterior
        //     (drenado al volver online o al abrir la app)
        //   - Toast permanente con botón "🖨 Reintentar ahora" si falla
        if (d.idGuia) {
          try {
            const base = location.origin + location.pathname.replace(/\/[^/]*$/, '');
            const reporteUrl = `${base}/reporte.html?tipo=guia&id=${encodeURIComponent(d.idGuia)}`;
            _imprimirConReintento(d.idGuia, reporteUrl, 'Guía ' + d.idGuia);
          } catch(eImp){
            _encolarImpresionPendiente(d.idGuia, 'Guía ' + d.idGuia);
            toast('🖨 No se pudo imprimir: ' + (eImp.message || eImp) + ' · encolado para reintento', 'warn', 9000);
          }
        }
        // Historial local
        const items = (pickupSnap.items || []).filter(it => (parseFloat(it.despachado)||0) > 0).map(it => ({
          codigoBarra: it.skuBase, descripcion: it.nombre, cantidad: it.despachado
        }));
        const zonas = OfflineManager.getZonasCache();
        const zonaObj = zonas.find(z => String(z.idZona) === String(pickupSnap.idZona));
        _saveHist({
          ts: Date.now(), n: items.length, tipo: 'SALIDA_ZONA',
          idZona: pickupSnap.idZona,
          nombreZona: zonaObj ? (zonaObj.nombre || zonaObj.idZona) : (pickupSnap.idZona || ''),
          nota: 'Pickup ' + pickupSnap.idPickup,
          items, idGuia: d.idGuia || '—', ok: true
        });
        _renderHist();

        // Limpiar pickup local + renders inline
        _pickupActivo = null;
        _clearPickup();
        _renderPickupActivoBanner();
        _renderPickupChecklistInline();
        _renderExtrasSection();
        _updateGenerarBtn();
        _renderCart();
        badgeUpdate();
      } else {
        try { SoundFX.error(); } catch(_){}
        vibrate([80, 40, 80]);
        toast('Error: ' + (res.error || 'no se pudo cerrar'), 'danger', 8000);
        // Revert
        _cart = cartSnap; _saveCart(); _renderCart(); _updateFooter();
      }
    } catch (e) {
      try { SoundFX.error(); } catch(_){}
      vibrate([80, 40, 80]);
      const msg = e?.timeout ? 'Tiempo agotado' : 'Sin conexión';
      toast('Error: ' + msg, 'danger', 8000);
      _cart = cartSnap; _saveCart(); _renderCart(); _updateFooter();
    } finally {
      _pickupClosing = false;
      _setGenerarGuiaBusy(false);
    }
  }

  // Pone/quita estado "ocupado" en el botón Generar Guía y CTAs del pickup.
  // Llamado solo desde cerrarDespachoPickup para impedir doble-click visible.
  function _setGenerarGuiaBusy(busy) {
    try {
      const btn = document.getElementById('btnGenerarGuia');
      if (btn) {
        btn.disabled = !!busy;
        if (busy) btn.classList.add('is-busy'); else btn.classList.remove('is-busy');
      }
      document.querySelectorAll('.pkck-cta').forEach(el => {
        el.style.pointerEvents = busy ? 'none' : '';
        el.style.opacity       = busy ? '0.55' : '';
      });
    } catch(_){}
  }


  function abrirPickupPendiente() {
    if (!_pickupsPendientes.length) return;
    abrirPickup(_pickupsPendientes[0].idPickup);
  }

  // Abrir un pickup específico por idPickup (botón "Jalar"/"Continuar" en la lista)
  // Aplica lock optimista: si otro operador ya lo está atendiendo, no permite abrir.
  // OPTIMISTA — abrir pickup al instante (sin esperar backend).
  // Lock al backend en background. Render directo en view-despacho (sin sheet).
  // Efectos: sonido beepDouble + vibración patrón + voz + toast.
  async function abrirPickup(idPickup) {
    const p = _pickupsPendientes.find(x => String(x.idPickup) === String(idPickup));
    if (!p) return;
    const usuario = window.WH_CONFIG?.usuario || '';
    const atp = String(p.atendidoPor || '').trim();
    // Lock real: bloquear solo si OTRO operador (no yo mismo en otro device).
    if (atp && usuario && !_sameUser(atp, usuario)) {
      try { SoundFX.warn(); } catch(_){}
      vibrate([30, 20, 30]);
      toast(`🔒 Lo está atendiendo ${atp}. No se puede tomar.`, 'warn', 3500);
      _vozAnunciar('Pickup ocupado por ' + atp, { rate: 1.1 });
      return;
    }

    const tieneSku = Array.isArray(p.items) && p.items.some(it => it && it.skuBase);
    if (tieneSku) {
      // Setear pickup activo INMEDIATAMENTE (optimista) — el operador ve resultado al toque
      const yaActivo = _pickupActivo && String(_pickupActivo.idPickup) === String(p.idPickup);
      if (!yaActivo) {
        _pickupActivo = {
          ...p,
          atendidoPor: usuario,
          estado: 'EN_PROCESO',
          items: (p.items || []).map(it => ({
            skuBase:           it.skuBase,
            nombre:            it.nombre || it.skuBase,
            solicitado:        parseFloat(it.solicitado) || 0,
            despachado:        parseFloat(it.despachado) || 0,
            codigosOriginales: Array.isArray(it.codigosOriginales) ? it.codigosOriginales : [],
            despachadoPorCodigo: it.despachadoPorCodigo || {}
          }))
        };
      } else {
        // Conservar progreso local, solo refrescar metadata
        _pickupActivo.estado      = 'EN_PROCESO';
        _pickupActivo.atendidoPor = p.atendidoPor || usuario;
      }
      _savePickup();
      // Quitar de la lista de pendientes (ya está activo, no debe duplicarse)
      _pickupsPendientes = _pickupsPendientes.filter(x => String(x.idPickup) !== String(p.idPickup));

      // Efectos sonoros + visuales — feedback inmediato
      try { SoundFX.beepDouble(); } catch(_){}
      vibrate([15, 25, 15]);
      _vozAnunciar(`Pickup ${p.idZona || ''} jalado`, { rate: 1.05 });

      // Render inline (sin sheet) — todo en una sola vista
      _renderPickupActivoBanner();
      _renderPickupChecklistInline();
      _renderCart();
      _renderExtrasSection();
      _updateGenerarBtn();
      badgeUpdate();
      toast(`▶ Pickup ${p.idZona || ''} listo · escanea para registrar`, 'ok', 2400);

      // Backend en background — no bloquea la UI
      API.actualizarPickup({
        idPickup:    p.idPickup,
        estado:      'EN_PROCESO',
        lockUsuario: usuario,
        tomarLock:   true
      }).then(r => {
        if (r && r.ok === false && r.conflicto) {
          try { SoundFX.warn(); } catch(_){}
          toast(`⚠ ${r.atendidoPor} también está atendiendo este pickup`, 'warn', 5000);
        }
      }).catch(() => { /* offline ok — autosave reintentará */ });
      return;
    }
    // Fallback legacy: pickup sin skuBase → fuzzy match con sheet (caso n8n viejo)
    _pickupActivo = p;
    try { SoundFX.beep(); } catch(_){}
    _runMatching(p);
    abrirSheet('sheetDespPickup');
  }

  // Soltar pickup activo — vuelve a PENDIENTE para que otro lo tome.
  // Limpia checklist inline y sección extras. Vista vuelve a estado inicial.
  // Efectos: warn + vibración + voz + toast naranja.
  async function soltarPickupActivo() {
    if (!_pickupActivo) return;
    const idP = _pickupActivo.idPickup;
    const zona = _pickupActivo.idZona || '';
    if (!await _whConfirm('¿Soltar este pickup?\n\nEl progreso despachado se conserva en la hoja.\nOtro operador podrá tomarlo.', { warning: true, titulo: 'Soltar pickup', okText: 'Soltar' })) return;

    try { SoundFX.warn(); } catch(_){}
    vibrate([30, 20, 60]);
    _vozAnunciar(`Pickup ${zona} soltado`, { rate: 1.05 });

    // Limpiar local INMEDIATAMENTE (optimista) — el operador ve la vista limpia
    _pickupActivo = null;
    _clearPickup();
    _renderPickupActivoBanner();
    _renderPickupChecklistInline();
    _renderExtrasSection();
    _updateGenerarBtn();
    badgeUpdate();
    toast(`🔓 Pickup ${zona} soltado · ya está disponible para otro operador`, 'info', 3000);

    // Backend en background
    API.liberarPickup({ idPickup: idP }).catch(() => {});
    // Forzar refresco de la lista de pendientes (para que reaparezca pronto)
    setTimeout(() => _pollPickups(), 800);
  }

  // Sincronización al hidratar: si en backend ya está cerrado, limpiar local.
  async function _sincronizarPickupActivo() {
    if (!_pickupActivo) return;
    try {
      const r = await API.getPickup({ idPickup: _pickupActivo.idPickup });
      if (r && r.ok && r.data) {
        const estadoBE = String(r.data.estado || '').toUpperCase();
        if (['COMPLETADO','CANCELADO','PARCIAL'].indexOf(estadoBE) >= 0) {
          toast(`ℹ Este pickup ya fue ${estadoBE.toLowerCase()} · limpiando estado local`, 'info', 4000);
          _pickupActivo = null;
          _clearPickup();
          _renderPickupActivoBanner();
          badgeUpdate();
          return;
        }
        // Si otro operador (no yo mismo en otro device) lo está atendiendo, avisar
        const usuario = window.WH_CONFIG?.usuario || '';
        const atpBE = String(r.data.atendidoPor || '').trim();
        if (atpBE && usuario && !_sameUser(atpBE, usuario)) {
          toast(`⚠ Otro operador (${atpBE}) está atendiendo este pickup`, 'warn', 5000);
        }
      }
    } catch(_){}
  }

  function _runMatching(pickup) {
    const items = pickup.items || [];
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => { stockMap[s.codigoProducto || s.idProducto] = s; });

    // Items con skuBase + codigosOriginales (vienen del cierre de caja ME)
    // resuelven DIRECTO al producto del catálogo — sin fuzzy match.
    const productosMaster = App.getProductosMaestro ? App.getProductosMaestro() : [];
    function _esCanonico(p) {
      // Producto canónico = factorConversion 1 (o vacío). Las presentaciones
      // tienen factorConversion >1 (ej: pack de 24 unidades, factor=24).
      return parseFloat(p.factorConversion || 1) === 1;
    }
    function _resolverDirecto(item) {
      // 1. PREFERIR el canónico cuyo idProducto === skuBase del item
      //    (es la fila del catálogo cuyo skuBase apunta a sí mismo).
      let prod = productosMaster.find(p =>
        String(p.idProducto) === String(item.skuBase) && _esCanonico(p));
      if (prod) return prod;
      // 2. Cualquier producto con skuBase === item.skuBase + canónico
      prod = productosMaster.find(p =>
        String(p.skuBase || '') === String(item.skuBase) && _esCanonico(p));
      if (prod) return prod;
      // 3. Match por idProducto exacto (sin requerir canónico)
      prod = productosMaster.find(p => String(p.idProducto) === String(item.skuBase));
      if (prod) return prod;
      // 4. Por cualquier codigoOriginal — pero buscar primero los canónicos
      if (Array.isArray(item.codigosOriginales)) {
        for (const cod of item.codigosOriginales) {
          prod = productosMaster.find(p =>
            (String(p.codigoBarra) === String(cod) || String(p.idProducto) === String(cod)) &&
            _esCanonico(p));
          if (prod) return prod;
        }
        // 5. Fallback: cualquier match aunque no sea canónico
        for (const cod of item.codigosOriginales) {
          prod = productosMaster.find(p =>
            String(p.codigoBarra) === String(cod) || String(p.idProducto) === String(cod));
          if (prod) return prod;
        }
      }
      return null;
    }

    _matchResults = items.map(item => {
      // Caso ME (skuBase explícito): resolver directo, sin fuzzy.
      // parseFloat para preservar decimales de granel (KGM solicita 1.35 kg).
      const qty = parseFloat(item.qty || item.solicitado) || 0;
      const nombreBusq = item.nombre || item.skuBase || '';
      if (item.skuBase) {
        const prod = _resolverDirecto(item);
        if (prod) {
          const stockDisp = parseFloat((stockMap[prod.idProducto] || {}).cantidadDisponible || 0);
          return { nombre: nombreBusq, qty, producto: prod, score: 1, status: 'exacto', accepted: true, stockDisp,
                   skuBase: item.skuBase, codigosOriginales: item.codigosOriginales || [] };
        }
        // Si no resuelve directo, caer a fuzzy como fallback
      }
      // Fallback fuzzy match por nombre cuando el item llega sin skuBase resoluble
      const { producto, score } = _bestMatch(nombreBusq);
      const stockDisp = producto ? parseFloat((stockMap[producto.idProducto] || {}).cantidadDisponible || 0) : 0;
      const status = score >= 0.75 ? 'exacto' : score >= 0.35 ? 'parcial' : 'nofound';
      return { nombre: nombreBusq, qty, producto, score, status, accepted: status === 'exacto', stockDisp,
               skuBase: item.skuBase || '', codigosOriginales: item.codigosOriginales || [] };
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

  // ═══ [v2.13.8] LISTA SOMBRA — IA-powered shadow checklist ═══
  // Persistencia local (efímera por sesión). Items: [{nombre, cantidad, atendida, productos:[]}]
  const LS_KEY = 'wh_lista_sombra';
  let _listaSombra = null;

  function _lsSave() {
    if (_listaSombra) localStorage.setItem(LS_KEY, JSON.stringify(_listaSombra));
    else              localStorage.removeItem(LS_KEY);
  }
  function _lsLoad() {
    try { return JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch { return null; }
  }
  function _lsRehidratar() {
    _listaSombra = _lsLoad();
    if (_listaSombra) {
      _lsRecalcular();
      _lsRender();
    }
  }

  // Normaliza string para fuzzy match: minúsculas, sin tildes, sin extras
  function _lsNorm(s) {
    return String(s || '')
      .toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9 ]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
  // Tokens significativos (≥3 chars), excluyendo stopwords cortas
  function _lsTokens(s) {
    return _lsNorm(s).split(' ').filter(t => t.length >= 3);
  }
  // Score Jaccard sobre tokens — entre 0 y 1
  function _lsScore(a, b) {
    const ta = new Set(_lsTokens(a));
    const tb = new Set(_lsTokens(b));
    if (!ta.size || !tb.size) return 0;
    let inter = 0;
    ta.forEach(t => { if (tb.has(t)) inter++; });
    return inter / Math.max(ta.size, tb.size);
  }

  // Recorre cart y actualiza cantidadEscaneada de cada item de la sombra.
  // [v2.13.18] Match por skuBase (resolviendo el código escaneado del cart al
  // canónico). Esto soporta equivalencias: si escaneas un EAN equivalente, se
  // resuelve a su skuBase canónico y matchea contra el item de la sombra.
  function _lsRecalcular() {
    if (!_listaSombra || !_listaSombra.items) return;
    const productos = OfflineManager.getProductosCache() || [];
    const equivs    = OfflineManager.getEquivalenciasCache() || [];

    // Helper: dado un codigoBarra del cart, devuelve su skuBase canónico.
    // [v2.13.24] Usa el campo skuBase del producto, no idProducto.
    function _resolverSkuBase(cb) {
      if (!cb) return '';
      const cbN = normCb(cb);
      // 1. Match directo en productos canónicos
      const p = productos.find(x =>
        (String(x.codigoBarra || '').trim().toUpperCase() === cbN ||
         String(x.idProducto  || '').trim().toUpperCase() === cbN) &&
        parseFloat(x.factorConversion || 1) === 1
      );
      if (p) return String(p.skuBase || p.idProducto || p.codigoBarra || '');
      // 2. Equivalencia → skuBase (la equivalencia ya tiene skuBase directo)
      const e = equivs.find(x => String(x.codigoBarra || '').trim().toUpperCase() === cbN);
      if (e) return String(e.skuBase || '');
      return '';
    }

    _listaSombra.items.forEach(item => {
      let cantAcum = 0;
      const matcheados = [];
      const skuItem = String(item.skuBase || '').trim().toUpperCase();
      _cart.forEach(c => {
        let matchea = false;
        if (skuItem) {
          const skuCart = String(_resolverSkuBase(c.codigoBarra) || '').trim().toUpperCase();
          if (skuCart && skuCart === skuItem) matchea = true;
        } else {
          // Fallback fuzzy si el item no tiene skuBase identificado
          const sc = _lsScore(item.nombre, c.descripcion || '');
          if (sc >= 0.6) matchea = true;
        }
        if (matchea) {
          cantAcum += parseFloat(c.cantidad) || 0;
          matcheados.push(c.codigoBarra);
        }
      });
      item.cantidadEscaneada = Math.round(cantAcum * 1000) / 1000;
      item.productos = matcheados;
    });
  }

  // [v2.13.10] Items visibles por default (toggle solo para colapsar manualmente)
  let _lsItemsAbiertos = true;
  let _lsUltimoMarcadoIdx = -1;

  // [v2.13.20] Render banner SOMBRA ACTIVA (pkact-banner.is-sombra) + checklist
  // morado debajo. Refresca también el feed para que la card de la sombra activa
  // desaparezca de ahí (igual que pickup activo se quita del feed pendientes).
  function _lsRender() {
    // [v2.13.29] Marca body para que CSS pueda compactar el cart si hay sombra
    try { document.body.classList.toggle('has-sombra', !!_listaSombra); } catch(_){}
    const banner = document.getElementById('despSombraActivaBanner');
    if (banner) {
      if (!_listaSombra || !_listaSombra.items || !_listaSombra.items.length) {
        banner.style.display = 'none';
      } else {
        banner.style.display = 'block';
        const total      = _listaSombra.items.length;
        const completos  = _listaSombra.items.filter(i => (i.cantidadEscaneada || 0) >= i.cantidad).length;
        const titulo = document.getElementById('sombraActivaTitulo');
        const lbl    = document.getElementById('sombraActivaProgresoLbl');
        const bar    = document.getElementById('sombraActivaBar');
        const creador = _listaSombra.creador ? `Lista de ${_listaSombra.creador}` : 'Sombra activa';
        if (titulo) titulo.textContent = creador;
        if (lbl)    lbl.textContent    = `${completos}/${total}`;
        if (bar)    bar.style.width    = (completos / total * 100).toFixed(1) + '%';
      }
    }
    // Render checklist morado abajo (cards estilo pickup)
    _renderListaSombraChecklistInline();
    // Refrescar el feed (la sombra activa se filtra de ahí dentro del render del feed)
    try { badgeUpdate(); } catch(_){}
  }

  function _renderListaSombraChecklistInline() {
    const cont = document.getElementById('despListaSombraChecklist');
    if (!cont) return;
    if (!_listaSombra || !_listaSombra.items || !_listaSombra.items.length) {
      cont.style.display = 'none'; cont.innerHTML = '';
      return;
    }
    const productos = OfflineManager.getProductosCache() || [];
    // [v2.13.19] Separar items: identificados (con skuBase) vs libres (sin skuBase)
    // [v2.13.29] Mantener ORDEN ORIGINAL — no reordenar al marcar. La sombra
    // es referencia visual fija; el operador localiza por posición no por estado.
    const identificados = [];
    const libres = [];
    _listaSombra.items.forEach((it, idx) => {
      if (it.skuBase) identificados.push({ it, idx });
      else libres.push({ it, idx });
    });
    cont.style.display = 'block';
    let html = '';
    // Sección identificados (cards marcables)
    html += identificados.map(({ it, idx }) => {
      const sol  = parseFloat(it.cantidad) || 0;
      const desp = parseFloat(it.cantidadEscaneada) || 0;
      const pct  = sol > 0 ? Math.min(100, Math.round((desp / sol) * 100)) : 0;
      const completo  = desp >= sol && sol > 0;
      const enProg    = desp > 0 && desp < sol;
      const cls       = completo ? 'is-completo' : (enProg ? 'is-progreso' : '');
      const flashCls  = it._flash ? ' lsck-flash' : '';
      if (it._flash) it._flash = false;
      const skuLbl    = it.skuBase;
      const nombreShow = it.nombreMaster || it.nombre;
      const prod = productos.find(p =>
        String(p.skuBase) === String(it.skuBase) ||
        String(p.idProducto) === String(it.skuBase) ||
        String(p.codigoBarra) === String(it.skuBase)
      );
      const esKg = prod && typeof _esProductoPeso === 'function' && _esProductoPeso(prod);
      const unidadLbl = esKg ? String(prod.unidad || 'kg').toLowerCase() : '';
      const icon = completo ? '✓' : (enProg ? '⏳' : '📋');
      return `
        <div class="pkck-card lsck-card ${cls}${flashCls}" data-sku-ls="${escAttr(it.skuBase)}" data-idx="${idx}">
          <div class="pkck-row">
            <div class="pkck-icon">${icon}</div>
            <div class="flex-1 min-w-0">
              <p class="pkck-name truncate">${escHtml(nombreShow)}</p>
              <p class="pkck-meta truncate">${escHtml(skuLbl)}</p>
            </div>
            <div class="pkck-qty-wrap">
              <p><span class="pkck-qty">${esKg ? fmt(desp,3) : desp}</span><span class="pkck-qty-sol"> / ${esKg ? fmt(sol,3) : sol}${esKg ? ' '+unidadLbl : ''}</span></p>
              <p class="text-[10px] text-slate-500 mt-0.5">${pct}%</p>
            </div>
          </div>
          <div class="pkck-bar-wrap">
            <div class="pkck-bar-fill" style="width:${pct}%"></div>
            ${enProg ? '<div class="pkck-bar-shimmer"></div>' : ''}
          </div>
          <div class="pkck-check-overlay">✓</div>
        </div>`;
    }).join('');

    // [v2.13.19] Sección de items LIBRES (sin SKU identificado) — solo info,
    // no se pueden marcar. El operador puede tocar uno para buscarlo manualmente.
    if (libres.length) {
      html += `
        <div class="lsck-libres-titulo">
          <span class="lsck-libres-ico">📝</span>
          <span>Libres · ${libres.length} sin identificar (solo guía)</span>
        </div>` +
        libres.map(({ it, idx }) => `
          <div class="pkck-card lsck-card lsck-libre" data-idx="${idx}"
               onclick="DespachoView.buscarItemSombra(${idx})" style="cursor:pointer">
            <div class="pkck-row">
              <div class="pkck-icon">📝</div>
              <div class="flex-1 min-w-0">
                <p class="pkck-name truncate">${escHtml(it.nombre)}</p>
                <p class="pkck-meta truncate">↻ toca para buscar producto</p>
              </div>
              <div class="pkck-qty-wrap">
                <p><span class="pkck-qty-sol">×${it.cantidad}</span></p>
              </div>
            </div>
          </div>`).join('');
    }

    cont.innerHTML = html;
    // [v2.13.29] NO auto-scroll. La sombra es ayuda visual fija — solo el
    // operador la mueve. El flash y la rotación de "pendientes primero" ya
    // ayudan a notar qué está marcado sin necesidad de mover la vista.
    if (_lsUltimoMarcadoIdx >= 0) {
      setTimeout(() => { _lsUltimoMarcadoIdx = -1; }, 3000);
    }
  }

  // [v2.13.15] Panel de listas sombra compartidas — DISPONIBLES + EN_USO
  let _lsPanelData = [];
  let _lsPanelTimer = null;

  function _lsStartPanelSync() {
    if (_lsPanelTimer) return;
    _lsRefrescarPanel();
    _lsPanelTimer = setInterval(() => _lsRefrescarPanel(), 15000);
  }
  function _lsStopPanelSync() {
    if (_lsPanelTimer) { clearInterval(_lsPanelTimer); _lsPanelTimer = null; }
  }
  function _lsRefrescarPanel() {
    API.getListasSombra({}).then(r => {
      if (!r || !r.ok) return;
      _lsPanelData = (r.data?.listas) || [];
      _lsRenderPanel();
    }).catch(() => {});
  }
  function _lsRenderPanel() {
    // [v2.13.16] Las listas sombra ahora se renderizan dentro del feed de
    // pickups (despPickupsLista) en badgeUpdate. Solo disparamos el refresh.
    try { badgeUpdate(); } catch(_){}
  }

  // [v2.13.19] Mapeo completo de items del backend al state local
  function _lsMapItemFromBackend(it) {
    return {
      nombre:       it.nombre || it.nombreMaster || '',
      cantidad:     parseFloat(it.cantidad) || 0,
      skuBase:      String(it.skuBase || ''),
      codigoBarra:  String(it.codigoBarra || ''),
      nombreMaster: it.nombreMaster || it.nombre || '',
      unidad:       it.unidad || '',
      codigoVisto:  it.codigoVisto || '',
      cantidadEscaneada: parseFloat(it.cantidadEscaneada) || 0,
      productos:    Array.isArray(it.productos) ? it.productos : []
    };
  }

  async function tomarListaSombraDelPanel(idLista) {
    if (_listaSombra && _listaSombra.id !== idLista) {
      if (!await _whConfirm('Ya tienes una lista sombra activa.\n\n¿Reemplazarla por la que vas a tomar?', { warning: true, titulo: 'Reemplazar lista' })) return;
    }
    const usuario = window.WH_CONFIG?.usuario || '';
    if (!usuario) { toast('Sin sesión activa', 'warn'); return; }
    // [v2.13.19] OPTIMISTA: usa la data del feed (que ya viene con items full)
    // para activar el checklist INMEDIATAMENTE sin esperar al backend.
    const enFeed = (_lsPanelData || []).find(l => l.idLista === idLista);
    if (!enFeed || !Array.isArray(enFeed.items)) {
      toast('Lista no encontrada en feed — refresca', 'warn');
      _lsRefrescarPanel();
      return;
    }
    _listaSombra = {
      id: idLista,
      idBackend: idLista,
      creada: enFeed.fechaCreacion || new Date().toISOString(),
      creador: enFeed.usuarioCreador || '',
      tomadaPor: usuario,
      estado: 'EN_USO',
      items: enFeed.items.map(_lsMapItemFromBackend)
    };
    _lsItemsAbiertos = true;
    _lsRecalcular();
    _lsSave();
    _lsRender();
    try { _renderDespFlotante(); } catch(_){}
    try { SoundFX.pickupOk(); } catch(_){}
    vibrate([20, 30, 40]);
    toast(`✓ Lista jalada · ${_listaSombra.items.length} productos`, 'ok', 3500);
    // Backend en background — si falla por race condition (otro la tomó),
    // revertimos y avisamos.
    API.tomarListaSombra({ idLista, usuario,
      localId: 'L' + Date.now() + Math.random().toString(36).slice(2, 8)
    }).then(r => {
      if (!r?.ok) {
        // Rollback
        const msg = r?.error === 'EN_USO_POR_OTRO' ? (r.mensaje || 'Otro operador la tomó primero')
                  : r?.error === 'YA_COMPLETADA'   ? 'Ya fue completada'
                  : ('No se pudo tomar: ' + (r?.error || 'error'));
        toast('⚠ ' + msg + ' — la quito de tu vista', 'warn', 5000);
        _listaSombra = null;
        _lsSave();
        _lsRender();
        try { _renderDespFlotante(); } catch(_){}
      }
      _lsRefrescarPanel();
    }).catch(() => {
      // Sin conexión: dejamos local + warn
      toast('⚠ Sin conexión — trabajas local, sincroniza después', 'warn', 5000);
    });
  }

  // [v2.13.19] Anular lista sombra del feed (optimista). El creador o cualquiera
  // si no está EN_USO. Marca como ANULADA en backend.
  async function anularListaSombraDelFeed(idLista) {
    if (!await _whConfirm('¿Eliminar esta lista del feed?\n\nEsta acción no se puede deshacer.', { danger: true, titulo: 'Eliminar lista', okText: 'Eliminar' })) return;
    try { SoundFX.click(); } catch(_){}
    // Optimista: quitar del panel local
    _lsPanelData = (_lsPanelData || []).filter(l => l.idLista !== idLista);
    try { badgeUpdate(); } catch(_){}
    // Si era la que tenía activa, también la cierro local
    if (_listaSombra && _listaSombra.id === idLista) {
      _listaSombra = null;
      _lsSave();
      _lsRender();
      try { _renderDespFlotante(); } catch(_){}
    }
    API.anularListaSombra({
      idLista,
      usuario: window.WH_CONFIG?.usuario || '',
      localId: 'L' + Date.now() + Math.random().toString(36).slice(2, 8)
    }).then(r => {
      if (!r?.ok && r?.error !== 'NO_ENCONTRADA') {
        toast('No se pudo eliminar: ' + (r?.error || 'error'), 'warn', 4000);
        _lsRefrescarPanel();
      } else {
        toast('Lista eliminada del feed', 'ok', 2500);
      }
    }).catch(() => {
      toast('Sin conexión — se eliminó local, pero puede reaparecer al refrescar', 'warn');
    });
  }

  // [v2.13.15] Sync periódico de progreso al backend (cada 20s)
  let _lsProgresoTimer = null;
  function _lsStartProgresoSync() {
    if (_lsProgresoTimer) return;
    _lsProgresoTimer = setInterval(() => {
      if (_listaSombra && _listaSombra.idBackend) {
        API.actualizarProgresoListaSombra({
          idLista: _listaSombra.idBackend,
          items: JSON.stringify(_listaSombra.items)
        }).catch(() => {});
      }
    }, 20000);
  }

  // [v2.13.13] Identificar SKUs — NO toca el cart. Solo resuelve qué producto
  // del catálogo corresponde a cada item de la sombra (igual que pickup pre-resolved).
  // Cuando el operador escanea físicamente ese código, el producto entra al cart
  // y por match exacto (no fuzzy) se marca el item de la sombra.
  function jalarTodoPosible() {
    if (!_listaSombra || !_listaSombra.items) { toast('No hay lista sombra activa', 'warn'); return; }
    const pendientes = _listaSombra.items
      .map((it, i) => ({ it, i }))
      .filter(({ it }) => (it.cantidadEscaneada || 0) < it.cantidad && !it.codigoBarraSugerido);
    if (!pendientes.length) {
      toast('Todo identificado ✨ — escanea para confirmar cantidades', 'info');
      return;
    }
    const productos = OfflineManager.getProductosCache() || [];
    let identificados = 0, ambiguos = 0, sinMatch = 0;
    pendientes.forEach(({ it, i }) => {
      const scored = productos
        .map(p => ({ p, s: _lsScore(it.nombre, p.descripcion || '') }))
        .filter(x => x.s >= 0.7)
        .sort((a, b) => b.s - a.s);
      if (!scored.length) { sinMatch++; return; }
      const ganador = scored[0];
      const segundo = scored[1];
      const esUnico = !segundo || (ganador.s - segundo.s) >= 0.15;
      if (!esUnico) { ambiguos++; return; }
      const cb = String(ganador.p.codigoBarra || ganador.p.idProducto || '');
      if (!cb) { sinMatch++; return; }
      // SOLO resolvemos SKU en la sombra — no tocamos el cart
      it.codigoBarraSugerido = cb;
      it.descripcionSugerida = ganador.p.descripcion || it.nombre;
      identificados++;
    });
    if (identificados > 0) {
      _lsSave();
      _lsRender();
      try { SoundFX.rocket(); } catch(_){}
      vibrate([30, 25, 50]);
    }
    const partes = [];
    if (identificados > 0) partes.push(`🔗 ${identificados} identificados`);
    if (ambiguos > 0)      partes.push(`⚠ ${ambiguos} ambiguos`);
    if (sinMatch > 0)      partes.push(`❌ ${sinMatch} sin match`);
    const resumen = partes.join(' · ') || 'Sin cambios';
    const cola = identificados > 0 ? ' — escanea para confirmar cantidades' : '';
    toast(resumen + cola, identificados > 0 ? 'ok' : 'warn', 7000);
  }

  // [v2.13.9] Tap en item de sombra → abre buscador con el nombre pre-llenado
  // y dispara la búsqueda. El operador toca un resultado y se agrega al cart.
  function buscarItemSombra(idx) {
    if (!_listaSombra || !_listaSombra.items[idx]) return;
    const it = _listaSombra.items[idx];
    // Tomar palabras significativas para buscar (las 2-3 primeras suelen ser únicas)
    const tokens = _lsTokens(it.nombre);
    const query = tokens.slice(0, 3).join(' ') || it.nombre;
    try { SoundFX.click(); } catch(_){}
    abrirDespBusqueda();
    setTimeout(() => {
      const inp = document.getElementById('despSearchInput');
      if (inp) {
        inp.value = query;
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        // Selecciono el texto para que si quiere editarlo, sea fácil
        inp.select();
      }
      despBuscarInput(query);
    }, 120);
  }

  // Llamado tras cada cambio del carrito — actualiza sombra y dispara sfx si hay nuevo match
  function _lsOnCartChange() {
    if (!_listaSombra || !_listaSombra.items) return;
    // snapshot del estado antes (para detectar transiciones)
    const antes = _listaSombra.items.map(i => ({
      esc: i.cantidadEscaneada || 0,
      completo: (i.cantidadEscaneada || 0) >= i.cantidad
    }));
    _lsRecalcular();
    let nuevosCompletos = 0;
    let nuevosParciales = 0;
    let ultimoMarcado = -1;
    let ultimoNombre = '';
    _listaSombra.items.forEach((it, i) => {
      const escAhora = it.cantidadEscaneada || 0;
      const escAntes = antes[i].esc;
      const completoAhora = escAhora >= it.cantidad;
      if (completoAhora && !antes[i].completo) {
        it._flash = true;
        nuevosCompletos++;
        ultimoMarcado = i;
        ultimoNombre = it.nombre;
      } else if (escAhora > escAntes && !completoAhora) {
        it._flash = true;
        nuevosParciales++;
        if (ultimoMarcado < 0) { ultimoMarcado = i; ultimoNombre = it.nombre; }
      }
    });
    // [v2.13.10] Track del último para auto-scroll en render
    if (ultimoMarcado >= 0) _lsUltimoMarcadoIdx = ultimoMarcado;
    _lsSave();
    _lsRender();
    // SFX + toast según resultado
    if (nuevosCompletos > 0) {
      try { SoundFX.pickupItemOk(); } catch(_){}
      vibrate([15, 25, 35]);
      _lsToastMini(`✓ Completo: ${ultimoNombre.substring(0, 32)}`);
      // Si TODO está completo, chime grande
      const restantes = _listaSombra.items.filter(i => (i.cantidadEscaneada || 0) < i.cantidad).length;
      if (restantes === 0) {
        setTimeout(() => { try { SoundFX.pickupOk(); } catch(_){} }, 250);
        _lsToastMini('🎉 Lista sombra completa');
      }
    } else if (nuevosParciales > 0) {
      try { SoundFX.savedTick && SoundFX.savedTick(); } catch(_){}
      _lsToastMini(`+ Marcado: ${ultimoNombre.substring(0, 32)}`);
    }
  }

  // [v2.13.10] Toast pequeñito que aparece DENTRO del banner — no estorba al cart
  function _lsToastMini(msg) {
    const banner = document.getElementById('despListaSombraBanner');
    if (!banner) return;
    let toast = banner.querySelector('.ls-toast-mini');
    if (!toast) {
      toast = document.createElement('div');
      toast.className = 'ls-toast-mini';
      banner.appendChild(toast);
    }
    toast.textContent = msg;
    toast.classList.remove('is-show');
    void toast.offsetWidth; // reflow
    toast.classList.add('is-show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => toast.classList.remove('is-show'), 2200);
  }

  // [v2.13.20] toggleListaSombra ya no aplica — el banner nuevo no tiene
  // colapso/expansión, los items siempre se ven como cards abajo (igual pickup).
  function toggleListaSombra() { /* no-op por compat */ }

  async function cerrarListaSombra() {
    if (!_listaSombra) return;
    // [v2.13.15] 3 opciones: liberar (vuelve a DISPONIBLE) · cancelar · ocultar local
    const tieneBackend = !!_listaSombra.idBackend;
    let ok;
    if (tieneBackend) {
      ok = await _whConfirm('¿Liberar la lista sombra?\n\nVolverá a estar DISPONIBLE para que otro operador la tome. El carrito de despacho se mantiene.', { warning: true, titulo: 'Liberar lista', okText: 'Liberar' });
    } else {
      ok = await _whConfirm('¿Quitar la lista sombra local?\n\nEl carrito de despacho se mantiene.', { titulo: 'Quitar lista local' });
    }
    if (!ok) return;
    if (tieneBackend) {
      API.liberarListaSombra({
        idLista: _listaSombra.idBackend,
        usuario: window.WH_CONFIG?.usuario || '',
        localId: 'L' + Date.now() + Math.random().toString(36).slice(2, 8)
      }).catch(() => {});
    }
    _listaSombra = null;
    _lsSave();
    _lsRender();
    try { _renderDespFlotante(); } catch(_){}
    try { _lsRefrescarPanel(); } catch(_){}
    try { SoundFX.click(); } catch(_){}
    toast(tieneBackend ? 'Lista liberada — otro operador la puede tomar' : 'Lista sombra removida', 'info', 3000);
  }

  // ── Modal IA — pegado y análisis ──────────────────────────
  let _lsPreviewBuffer = [];

  function abrirModalLista() {
    document.getElementById('modalSubirLista').style.display = 'flex';
    _lsMostrarPaso(1);
    setTimeout(() => document.getElementById('lsTextoCrudo')?.focus(), 250);
    try { SoundFX.click(); } catch(_){}
  }

  function cerrarModalLista() {
    document.getElementById('modalSubirLista').style.display = 'none';
    document.getElementById('lsTextoCrudo').value = '';
    _lsPreviewBuffer = [];
    try { SoundFX.click(); } catch(_){}
  }

  function _lsMostrarPaso(n) {
    [1, 2, 3, 4].forEach(i => {
      const el = document.getElementById('lsModalPaso' + i);
      if (el) el.style.display = (i === n) ? 'block' : 'none';
    });
  }

  function volverPaso1() {
    _lsMostrarPaso(1);
  }

  // [v2.13.25] Chunking automático: si la lista es larga, parte en bloques
  // de ~30 productos y procesa secuencialmente con progreso visible. Esto
  // permite pegar listas de cualquier tamaño sin esperar gigante ni saturar
  // el rate-limit de Claude. Costo: ~$0.0005 por chunk, prácticamente gratis.
  async function analizarListaConIA() {
    const texto = document.getElementById('lsTextoCrudo').value.trim();
    if (!texto) { toast('Pega una lista primero', 'warn'); return; }
    if (texto.length < 10) { toast('La lista parece muy corta', 'warn'); return; }
    if (texto.length > 200000) {
      toast('Lista enorme (>200k chars). Divídela manualmente en 2 pegados.', 'warn', 6000);
      return;
    }

    _lsMostrarPaso(2);
    try { SoundFX.beep(); } catch(_){}

    // ── 1. Trocear por líneas no-vacías ────────────────────────
    const lineas = texto.split(/\r?\n/).filter(l => l.trim().length > 0);
    const CHUNK_SIZE = 30;          // ~30 productos por chunk
    const CHUNK_THRESHOLD = 45;     // umbral: hasta 45 líneas = 1 chunk

    const chunks = [];
    if (lineas.length <= CHUNK_THRESHOLD) {
      // Lista chica → 1 solo chunk con el texto original (preserva cabeceras)
      chunks.push(texto);
    } else {
      // Lista grande → N chunks por líneas
      for (let i = 0; i < lineas.length; i += CHUNK_SIZE) {
        chunks.push(lineas.slice(i, i + CHUNK_SIZE).join('\n'));
      }
    }
    const totalChunks = chunks.length;
    console.log(`[ListaSombra] ${lineas.length} líneas → ${totalChunks} chunks`);

    // ── 2. Aviso al operador si va a trocear ───────────────────
    const subEl = document.getElementById('lsLoadingSub');
    if (totalChunks > 1) {
      toast(`📦 Lista larga · IA la procesará en ${totalChunks} partes`, 'info', 4000);
    }

    // ── 3. Procesar secuencial con progreso ────────────────────
    const buffer = [];
    let chunkOk = 0;
    let chunkErr = 0;
    let primerError = null;
    const subsRot = [
      'Limpiando cabeceras y normalizando cantidades',
      'Detectando códigos y descripciones',
      'Casi listo…'
    ];
    let subRotIdx = 0;

    for (let i = 0; i < chunks.length; i++) {
      if (subEl) {
        subEl.textContent = totalChunks > 1
          ? `Procesando parte ${i + 1} de ${totalChunks} · ${buffer.length} productos detectados`
          : subsRot[subRotIdx++ % subsRot.length];
      }
      try {
        const res = await API.analizarListaSombra({ texto: chunks[i] });
        if (res?.ok && Array.isArray(res.data?.items)) {
          buffer.push.apply(buffer, res.data.items);
          chunkOk++;
        } else {
          chunkErr++;
          if (!primerError) primerError = res;
          console.warn('[Chunk', i + 1, '/', totalChunks, '] error:', res?.error);
          // Si es KEY_NOT_SET, abortar todo (no tiene sentido seguir)
          if (res?.error === 'KEY_NOT_SET') break;
        }
      } catch(e) {
        chunkErr++;
        console.warn('[Chunk', i + 1, '] excepción:', e?.message);
        if (!primerError) primerError = { error: 'NETWORK', mensaje: e?.message };
      }
    }

    // ── 4. Mostrar resultado ───────────────────────────────────
    if (!buffer.length) {
      const err = primerError?.error || 'Error desconocido';
      const msg = err === 'KEY_NOT_SET'
        ? 'La API key de Claude no está configurada. Avisa al administrador.'
        : (primerError?.mensaje || 'Inténtalo de nuevo o pega la lista en formato más limpio');
      document.getElementById('lsErrorTitulo').textContent = err === 'KEY_NOT_SET' ? 'IA no configurada' : 'No pude analizar';
      document.getElementById('lsErrorMsg').textContent = msg;
      _lsMostrarPaso(4);
      try { SoundFX.warn(); } catch(_){}
      return;
    }

    _lsPreviewBuffer = buffer;
    _renderPreview();
    _lsMostrarPaso(3);
    try { SoundFX.done(); } catch(_){}
    vibrate([20, 30, 40]);

    if (totalChunks > 1) {
      const errTxt = chunkErr > 0 ? ` (${chunkErr} parte(s) fallaron)` : '';
      toast(`✓ ${totalChunks} partes procesadas · ${buffer.length} productos${errTxt}`,
            chunkErr ? 'warn' : 'ok', 5000);
    }
  }

  function _renderPreview() {
    const cont = document.getElementById('lsPreviewList');
    document.getElementById('lsPreviewTotal').textContent = _lsPreviewBuffer.length;
    cont.innerHTML = _lsPreviewBuffer.map((it, i) => `
      <div class="ls-prev-item" style="animation-delay:${Math.min(i, 12) * 35}ms">
        <span class="ls-prev-num">${i + 1}.</span>
        <span class="ls-prev-nombre">${escHtml(it.nombre)}</span>
        <input type="number" step="0.1" min="0" value="${it.cantidad}"
               class="ls-prev-cant-input"
               oninput="DespachoView._lsPrevSetCant(${i}, this.value)">
        <button onclick="DespachoView._lsPrevDel(${i})" class="ls-prev-del" title="Quitar">✕</button>
      </div>`).join('');
  }

  function _lsPrevSetCant(idx, v) {
    if (!_lsPreviewBuffer[idx]) return;
    _lsPreviewBuffer[idx].cantidad = Math.round((parseFloat(v) || 0) * 10) / 10;
  }
  function _lsPrevDel(idx) {
    _lsPreviewBuffer.splice(idx, 1);
    _renderPreview();
    try { SoundFX.click(); } catch(_){}
  }

  // [v2.13.21] Identificar skuBase canónico priorizando código sobre descripción.
  // Pasos:
  //   1. Si codigoVisto está set → match exacto contra canónicos.codigoBarra
  //      o equivalencias.codigoBarra (resolviendo al skuBase del canónico).
  //   2. Si no encuentra por código → fuzzy match por descripción (>=0.7, gap>=0.15).
  //   3. Si ni código ni descripción → null (queda como "libre" en la UI).
  function _lsIdentificarSkuBase(nombre, productos, codigoVisto) {
    // PASO 1: match exacto por código (más confiable)
    if (codigoVisto) {
      const cb = normCb(codigoVisto);
      if (cb) {
        // Canónico directo (factor=1, activo)
        const canon = productos.find(p =>
          parseFloat(p.factorConversion || 1) === 1 &&
          String(p.estado) !== '0' && String(p.estado) !== 0 &&
          String(p.codigoBarra || '').trim().toUpperCase() === cb
        );
        if (canon) {
          return {
            skuBase:     String(canon.skuBase || canon.idProducto || canon.codigoBarra || ''),
            codigoBarra: String(canon.codigoBarra || ''),
            descripcion: String(canon.descripcion || nombre),
            unidad:      String(canon.unidad || ''),
            via:         'codigo_canonico',
            score:       1.0
          };
        }
        // Equivalencia → resolver al canónico
        const equivs = OfflineManager.getEquivalenciasCache() || [];
        const eq = equivs.find(e => String(e.codigoBarra || '').trim().toUpperCase() === cb);
        if (eq) {
          const skuB = String(eq.skuBase || '').trim().toUpperCase();
          const prod = productos.find(p =>
            parseFloat(p.factorConversion || 1) === 1 &&
            String(p.estado) !== '0' && String(p.estado) !== 0 &&
            (String(p.idProducto || '').trim().toUpperCase() === skuB ||
             String(p.codigoBarra || '').trim().toUpperCase() === skuB)
          );
          if (prod) {
            return {
              skuBase:     String(prod.skuBase || prod.idProducto || prod.codigoBarra || ''),
              codigoBarra: String(prod.codigoBarra || ''),
              descripcion: String(prod.descripcion || nombre),
              unidad:      String(prod.unidad || ''),
              via:         'codigo_equivalencia',
              score:       1.0
            };
          }
        }
      }
    }
    // PASO 2: fuzzy match por descripción (fallback)
    if (!nombre) return null;
    const candidatos = productos.filter(p => {
      const factor = parseFloat(p.factorConversion || 1);
      const activo = String(p.estado) !== '0' && String(p.estado) !== 0;
      return factor === 1 && activo;
    });
    const scored = candidatos
      .map(p => ({ p, s: _lsScore(nombre, p.descripcion || '') }))
      .filter(x => x.s >= 0.7)
      .sort((a, b) => b.s - a.s);
    if (!scored.length) return null;
    const ganador = scored[0];
    const segundo = scored[1];
    if (segundo && (ganador.s - segundo.s) < 0.15) return null;
    return {
      skuBase:     String(ganador.p.skuBase || ganador.p.idProducto || ganador.p.codigoBarra || ''),
      codigoBarra: String(ganador.p.codigoBarra || ''),
      descripcion: String(ganador.p.descripcion || nombre),
      unidad:      String(ganador.p.unidad || ''),
      via:         'descripcion',
      score:       ganador.s
    };
  }

  function activarListaSombra() {
    console.log('[ListaSombra] activarListaSombra() called · buffer:', _lsPreviewBuffer);
    try {
      // [v2.13.14] Diagnóstico claro: si por alguna razón el buffer está vacío
      // (raro pero pasó), mostrar mensaje en el modal en lugar de fallar silenciosamente.
      if (!Array.isArray(_lsPreviewBuffer) || !_lsPreviewBuffer.length) {
        document.getElementById('lsErrorTitulo').textContent = 'Preview vacío';
        document.getElementById('lsErrorMsg').textContent = 'La lista no tiene productos para activar. Vuelve atrás y analiza de nuevo.';
        _lsMostrarPaso(4);
        try { SoundFX.warn(); } catch(_){}
        return;
      }
      // Re-leer cantidades desde los inputs por si el oninput no disparó en algún caso
      // (iOS Safari a veces no dispara input cuando se hace click directo sin focus)
      const inputs = document.querySelectorAll('#lsPreviewList .ls-prev-cant-input');
      inputs.forEach((inp, i) => {
        if (!_lsPreviewBuffer[i]) return;
        const v = parseFloat(inp.value);
        if (!isNaN(v)) _lsPreviewBuffer[i].cantidad = Math.round(v * 10) / 10;
      });
      const validos = _lsPreviewBuffer.filter(i => i && i.nombre && i.cantidad > 0);
      if (!validos.length) {
        document.getElementById('lsErrorTitulo').textContent = 'Cantidades inválidas';
        document.getElementById('lsErrorMsg').textContent =
          'Todas las cantidades son 0 o no válidas. Ajusta los valores y vuelve a intentar.';
        _lsMostrarPaso(4);
        try { SoundFX.warn(); } catch(_){}
        return;
      }
      const idLista = 'LS' + Date.now();
      // [v2.13.18] Identifica skuBase de cada item ANTES de subir
      const productos = OfflineManager.getProductosCache() || [];
      let identificados = 0;
      const items = validos.map(it => {
        const m = _lsIdentificarSkuBase(it.nombre, productos);
        if (m) identificados++;
        return {
          nombre: it.nombre,
          cantidad: it.cantidad,
          codigoVisto: it.codigoVisto || '',
          skuBase: m ? m.skuBase : '',
          codigoBarra: m ? m.codigoBarra : '',
          nombreMaster: m ? m.descripcion : '',
          unidad: m ? m.unidad : '',
          cantidadEscaneada: 0,
          productos: []
        };
      });
      console.log(`[ListaSombra] identificados ${identificados}/${items.length} skuBase al subir`);
      cerrarModalLista();
      try { SoundFX.rocket(); } catch(_){}
      vibrate([30, 25, 50]);
      toast(`📋 Lista subida · ${validos.length} productos · ${identificados} con SKU identificado · Tócala en el feed para jalar`, 'ok', 5000);
      // [v2.13.19] OPTIMISTA: pinta inmediato en el feed local (sin esperar backend)
      const usuario = window.WH_CONFIG?.usuario || '';
      const optEntry = {
        idLista: idLista,
        fechaCreacion: new Date().toISOString(),
        usuarioCreador: usuario,
        estado: 'DISPONIBLE',
        usuarioTomada: '',
        items: items,
        total: items.length,
        completos: 0
      };
      _lsPanelData = [optEntry].concat(_lsPanelData || []);
      try { badgeUpdate(); } catch(_){}
      // Backend en background — si falla, removemos la card y avisamos
      API.crearListaSombra({
        idLista: idLista,
        usuario: usuario,
        items: JSON.stringify(items),
        compartir: true,
        localId: 'L' + Date.now() + Math.random().toString(36).slice(2, 8)
      }).then(r => {
        if (r?.ok) {
          console.log('[ListaSombra] backend OK:', r.data);
          _lsRefrescarPanel();  // sync con backend real
        } else {
          console.warn('[ListaSombra] backend NO creado:', r);
          // Rollback optimista
          _lsPanelData = _lsPanelData.filter(l => l.idLista !== idLista);
          try { badgeUpdate(); } catch(_){}
          toast('⚠ No se pudo subir al equipo: ' + (r?.error || ''), 'danger', 5000);
        }
      }).catch(e => {
        console.warn('[ListaSombra] crear backend falló:', e?.message);
        _lsPanelData = _lsPanelData.filter(l => l.idLista !== idLista);
        try { badgeUpdate(); } catch(_){}
        toast('Sin conexión — la lista no se subió', 'warn', 4000);
      });
    } catch(err) {
      console.error('[ListaSombra] error en activar:', err);
      document.getElementById('lsErrorTitulo').textContent = 'Error al activar';
      document.getElementById('lsErrorMsg').textContent = err?.message || String(err);
      _lsMostrarPaso(4);
      try { SoundFX.error(); } catch(_){}
    }
  }

  return { cargar, pauseCamera,
           abrirDespCamara, cerrarDespCamara, cerrarDespYFinalizar,
           toggleDespTorch, despSetZoom,
           cerrarDespPicker, seleccionarItemDesp,
           despIncQty, despDecQty, despEditQty,
           despUndoLast, despLimpiarTodo,
           toggleDespScanInline, cerrarDespScan,
           submitDespScanInline, seleccionarItemDespInline, limpiarScanInput,
           verHistDetalle,
           abrirDespBusqueda, cerrarDespBusqueda, despBuscarInput, seleccionarDespBusqueda,
           selTipo,
           incQty, decQty, blurQty, quitarItem,
           finalizar, confirmarDespacho, cancelar,
           generarGuia,
           badgeUpdate, startPoll,
           abrirPickupPendiente, abrirPickup, elegirMatchParcial, confirmarPickup,
           empezarPickup, cerrarDespachoPickup, abrirSheetPickupActivo,
           soltarPickupActivo, pickupSetSearch, pickupSetSort,
           pkckMas: _pkckMas, pkckMenos: _pkckMenos, pkckSetGranel: _pkckSetGranel,
           hayDespachoActivo, procesarScanGlobal,
           flotMas, flotMenos, flotSetGranel,
           renderFlotante: _renderDespFlotante,
           // [v2.13.8] Lista sombra
           abrirModalLista, cerrarModalLista, analizarListaConIA,
           volverPaso1, activarListaSombra, toggleListaSombra, cerrarListaSombra,
           _lsPrevSetCant, _lsPrevDel,
           // [v2.13.9] Tap-to-search en item de sombra
           buscarItemSombra,
           // [v2.13.10] Jalar todo posible con fuzzy match
           jalarTodoPosible,
           // [v2.13.15] Listas compartidas
           tomarListaSombraDelPanel,
           _lsRefrescarPanel,
           // [v2.13.19] Anular del feed
           anularListaSombraDelFeed,
           // Hook expuesto para llamar desde puntos de mutación del cart
           _lsOnCartChange, _lsRehidratar };
})();

// ════════════════════════════════════════════════
// ESCÁNER GLOBAL — captura el lector HID en CUALQUIER módulo
// ════════════════════════════════════════════════
// Solo actúa si hay un despacho/pickup activo (DespachoView.hayDespachoActivo).
// Permite que el operador esté en el catálogo de Productos viendo el
// "membrete digital" y, al apuntar el escáner a la pantalla, el código
// se enrute al despacho en curso. Ignora si el foco está en un input
// (no roba lo que el usuario escribe) y distingue scanner de tecleo
// humano por la velocidad entre pulsaciones.
(function _initEscanerGlobal() {
  let buf = '';
  let lastTs = 0;
  let timer = null;

  function _esCampoEditable(el) {
    if (!el) return false;
    const tag = (el.tagName || '').toUpperCase();
    return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || el.isContentEditable;
  }

  function _procesar() {
    const code = buf.trim();
    buf = '';
    if (code.length < 3) return;
    try {
      if (window.DespachoView && DespachoView.hayDespachoActivo && DespachoView.hayDespachoActivo()) {
        DespachoView.procesarScanGlobal(code);
      }
    } catch(_) { /* tolerar */ }
  }

  document.addEventListener('keydown', (e) => {
    // Solo cuando hay despacho activo — si no, ni escuchamos.
    try {
      if (!(window.DespachoView && DespachoView.hayDespachoActivo && DespachoView.hayDespachoActivo())) return;
    } catch(_) { return; }
    // No robar el teclado si el operador está escribiendo en un campo.
    if (_esCampoEditable(document.activeElement)) return;
    // Si el sheet de escaneo HID propio del despacho está abierto, ese
    // tiene su propio listener — no duplicar.
    const hidSheet = document.getElementById('sheetScanInput');
    if (hidSheet && hidSheet.classList.contains('open')) return;

    const now = Date.now();
    if (e.key === 'Enter') {
      if (buf) { clearTimeout(timer); _procesar(); }
      return;
    }
    if (e.key.length !== 1) return; // ignorar teclas de control
    // Velocidad: si pasó >80ms entre teclas, es tecleo humano → reiniciar
    if (buf.length > 0 && (now - lastTs) > 80) buf = '';
    lastTs = now;
    buf += e.key;
    clearTimeout(timer);
    timer = setTimeout(_procesar, 120); // fallback si el scanner no manda Enter
  });
})();

// ════════════════════════════════════════════════
// PREINGRESOS VIEW
// ════════════════════════════════════════════════
const PreingresosView = (() => {
  let _filtroEstado      = '';
  let _busquedaQ         = '';
  let _tppBusq           = '';   // búsqueda del panel tablet
  let _tppFiltro         = '';   // filtro estado del panel tablet
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

  // [v2.13.174] ¿El preingreso tiene cambios persistidos que AÚN no se
  // avisaron a cajas? Compara los 5 campos críticos actuales contra el
  // snapshotAviso (lo último impreso). Stateless: si coinciden, ya se avisó.
  // Sin snapshot (preingreso viejo / nunca impreso) → no marcamos.
  function _hayCambiosSinAvisar(p) {
    try {
      const snap = p.snapshotAviso
        ? (typeof p.snapshotAviso === 'string' ? JSON.parse(p.snapshotAviso) : p.snapshotAviso)
        : null;
      if (!snap) return false;
      const actual = _snapshotAvisoLocal(
        p.idProveedor, p.monto,
        _tagsFromComentario(p.comentario),
        _textoLibreFromComentario(p.comentario)
      );
      return _diffSnapAvisoLocal(snap, actual).length > 0;
    } catch(_) { return false; }
  }

  function _renderCard(p) {
    const tieneGuia   = !!(p.idGuia && String(p.idGuia).trim());
    const nFotos      = p.fotos ? String(p.fotos).split(',').filter(Boolean).length : 0;
    const tags        = _tagsFromComentario(p.comentario);
    const borderColor = tieneGuia ? '#22c55e' : '#f59e0b';
    const provNombre  = _getProveedorNombre(p.idProveedor);
    const hora        = _horaDesdeId(p.idPreingreso);
    const fechaCorta  = _fmtCorta(p.fecha);

    const chipBg  = tieneGuia ? 'rgba(34,197,94,.12)' : 'rgba(245,158,11,.12)';
    const chipCol = tieneGuia ? '#4ade80' : '#fbbf24';
    const chipTxt = tieneGuia ? 'Con guía' : 'Sin guía';

    let nCargadores = 0;
    try { const c = JSON.parse(p.cargadores || '[]'); nCargadores = Array.isArray(c) ? c.length : 0; } catch {}
    const tagHtml = [
      tags.compl === 'si' ? '<span class="pre-qtag pre-qtag-green">Completo</span>'   : '',
      tags.compl === 'no' ? '<span class="pre-qtag pre-qtag-amber">Incompleto</span>' : '',
      tags.comp  === 'si' ? '<span class="pre-qtag pre-qtag-blue">Comp.</span>'       : '',
      nFotos > 0          ? `<span class="pre-qtag pre-qtag-slate">📷${nFotos}</span>` : '',
      nCargadores > 0     ? `<span class="pre-qtag" style="background:#451a03;color:#fbbf24">🛺${nCargadores}</span>` : '',
      _hayCambiosSinAvisar(p) ? `<span class="pre-qtag" style="background:#7c2d12;color:#fed7aa" title="Tiene cambios que no se han avisado a cajas">⚠ sin avisar</span>` : '',
    ].filter(Boolean).join('');

    const montoStr = p.monto ? ' · S/. ' + fmt(p.monto, 2) : '';

    const waBtn = `<button onclick="event.stopPropagation();PreingresosView.compartirWA('${escAttr(p.idPreingreso)}')"
      class="card-act card-act-wa" title="Compartir por WhatsApp">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
    </button>`;

    const actionBtn = tieneGuia
      ? `<button onclick="event.stopPropagation();App.nav('guias');GuiasView.verDetalle('${escAttr(p.idGuia)}')"
           class="card-act card-act-done" title="${escAttr(p.idGuia)}">
           <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M14 4.5V14a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2h5.5L14 4.5zm-3 0A1.5 1.5 0 0 1 9.5 3V1H4a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V4.5h-2z"/></svg>
           Guía
         </button>`
      : `<button onclick="event.stopPropagation();PreingresosView.crearGuiaRapido('${escAttr(p.idPreingreso)}')"
           class="card-act card-act-guia">
           + Guía
         </button>`;

    // Foto block (izquierda): carousel si múltiples, thumb si una, placeholder si ninguna
    const urls = p.fotos ? String(p.fotos).split(',').map(s => s.trim()).filter(Boolean).map(_normalizeDriveUrl) : [];
    let fotoBlock;
    if (urls.length > 1 && window.Photos && Photos.carouselHTML) {
      fotoBlock = `<div class="gcard-photo" onclick="event.stopPropagation()">${Photos.carouselHTML(urls, {size:'sm'})}</div>`;
    } else if (urls.length === 1) {
      fotoBlock = `<div class="gcard-photo" onclick="event.stopPropagation();Photos&&Photos.lightbox('${escAttr(urls[0])}')"><img src="${escAttr(urls[0])}" loading="lazy" onerror="this.style.opacity='.3'"/></div>`;
    } else {
      fotoBlock = `<div class="gcard-photo placeholder" title="Sin foto">📷</div>`;
    }

    return `
    <div class="pre-card" id="pre-${(p.idPreingreso||'').replace(/[^a-zA-Z0-9_-]/g,'_')}"
         style="border-left-color:${borderColor}"
         onclick="PreingresosView.abrirDetalle('${escAttr(p.idPreingreso)}')">
      ${fotoBlock}
      <div class="gcard-body">
        <div class="card-row-top">
          <span class="card-tipo-chip" style="background:${chipBg};color:${chipCol}">${chipTxt}</span>
          <div class="flex items-center gap-1 flex-shrink-0">${tagHtml}</div>
        </div>
        <p class="card-name" style="font-size:16px">${escHtml(provNombre)}</p>
        <div class="card-row-bottom">
          <span class="card-meta">${fechaCorta}${hora ? ' · ' + hora : ''}${montoStr}</span>
          <div class="card-actions">${waBtn}${actionBtn}</div>
        </div>
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
    const hoyKey  = _hoyPeru();
    const ayerKey = _ayerPeru();
    const months  = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    function _dateKey(p) {
      return (p.fecha && _diaPeru(p.fecha)) || '0000-00-00';
    }
    function _dateLabel(key) {
      if (!key || key === '0000-00-00') return 'Sin fecha';
      if (key === hoyKey)  return 'Hoy';
      if (key === ayerKey) return 'Ayer';
      const [yyyy, mm, dd] = key.split('-'); // key = 'YYYY-MM-DD' del día de Perú
      return yyyy === hoyKey.slice(0, 4)
        ? `${parseInt(dd, 10)} ${months[parseInt(mm, 10) - 1]}`
        : `${dd}/${mm}/${yyyy}`;
    }

    const groupMap = {};
    sorted.forEach(p => {
      const k = _dateKey(p);
      if (!groupMap[k]) groupMap[k] = [];
      groupMap[k].push(p);
    });

    container.innerHTML = Object.entries(groupMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([key, items]) => {
        const r = _resumenCargadoresDia(items);
        const hdrCls = r.carretas > 0 ? 'pre-date-hdr pre-date-hdr-row' : 'pre-date-hdr';
        const mid    = r.carretas > 0 ? '<span class="pre-hdr-line"></span>' : '';
        const alerta = ((r.medias || 0) + (r.vacias || 0)) > 0 ? '<span>·</span><span>⚠</span>' : '';
        const pill = r.carretas > 0
          ? `<button onclick="PreingresosView.abrirCargadoresDia('${key}')"
                     class="carg-pill-btn inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[10px] font-semibold">
                <span>🛺</span><span>${r.carretas} cart</span>${alerta}
              </button>`
          : '';
        return `<div class="${hdrCls}"><span>${_dateLabel(key)}</span>${mid}${pill}</div>
                <div class="pre-date-group">${items.map(_renderCard).join('')}</div>`;
      }).join('');

    // Preservar solo cards optimistas cuyo ID aún no está en la lista real
    optCards.forEach(c => {
      const rid = c.getAttribute('data-real-id') || c.id.replace('optcard_', '');
      if (!sorted.find(p => p.idPreingreso === rid)) {
        container.insertBefore(c, container.firstChild);
      }
    });

    // Activar carruseles de fotos (autoplay 4s + swipe + dots)
    if (window.Photos && Photos.initCarousels) Photos.initCarousels(container);

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
    const menu = document.getElementById('preFilterMenu');
    if (menu) menu.style.display = 'none';
    document.querySelectorAll('.pre-fopt').forEach(b => {
      b.classList.toggle('sel', b.dataset.pfiltro === estado);
    });
    const dot = document.getElementById('preFilterDot');
    if (dot) dot.style.display = estado ? 'block' : 'none';
    cargar(estado);
  }

  function _searchFocusPre(focused) {
    const toolbar = document.getElementById('preToolbar');
    if (!toolbar) return;
    if (focused) {
      toolbar.classList.add('srch-focused');
      const menu = document.getElementById('preFilterMenu');
      if (menu) menu.style.display = 'none';
    } else {
      setTimeout(() => toolbar.classList.remove('srch-focused'), 160);
    }
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
    // [v2.13.174] Feedback + autoguardado: antes las etiquetas SOLO persistían al
    // pulsar "Guardar cambios". Si el operador cambiaba un tag y cerraba, se perdía
    // (y el snapshot del aviso no lo reflejaba). Ahora autoguarda como comentario/monto.
    try { if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click(); } catch(_){}
    try { if (navigator.vibrate) navigator.vibrate(10); } catch(_){}
    _autoguardarMeta();
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
    // [v2.13.173 BUG FIX] Si había OTRO preingreso en edición, capturar sus
    // autoguardados pendientes ANTES de cambiar _editItem. Sin esto, un
    // comentario/cargador tipeado que seguía en el debounce se perdía al
    // saltar a otro preingreso. Disparo síncrono: _ejecutar* lee _editItem y
    // el DOM (= valores del ANTERIOR) y construye sus params antes del await,
    // así el save sale con los datos correctos aunque reasignemos abajo.
    if (_editItem && _editItem.idPreingreso !== idPreingreso) {
      clearTimeout(_autoguardarMetaTimer);
      clearTimeout(_autoguardarTimer);
      try { _ejecutarAutoguardarMeta(); }       catch(_){}
      try { _ejecutarAutoguardarCargadores(); } catch(_){}
    }
    const cached = OfflineManager.getPreingresosCache();
    const p = cached.find(x => x.idPreingreso === idPreingreso);
    if (p) {
      _editItem = { ...p };
      _renderModal(p);
      abrirSheet('sheetDetallePI');
      // [v2.13.110] Enganchar autoguardado tras render del modal
      setTimeout(_engancharAutoguardarMeta, 50);
    } else {
      // Aún no está en caché (recién creado) → buscar en GAS
      toast('Cargando...', 'info', 1500);
      API.getPreingresos({ idPreingreso }).then(res => {
        const item = res.ok && res.data?.find ? res.data.find(x => x.idPreingreso === idPreingreso) : null;
        if (item) {
          _editItem = { ...item };
          _renderModal(item);
          abrirSheet('sheetDetallePI');
          setTimeout(_engancharAutoguardarMeta, 50);
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
        provs.map(pv => `<option value="${escAttr(pv.idProveedor)}"${pv.idProveedor === p.idProveedor ? ' selected' : ''}>${escHtml(pv.nombre || pv.idProveedor)}</option>`).join('');
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
    if (btnS) { btnS.disabled = false; btnS.textContent = '💾 Guardar y avisar a cajas'; }
  }

  // ── Cargadores edit modal ────────────────────────────────
  function _renderCargadoresEdit() {
    const list = document.getElementById('piCargadoresList');
    if (!list) return;
    if (!_cargadoresEdit.length) {
      list.innerHTML = '<p class="text-xs text-slate-600 italic">Sin cargadores asignados</p>';
      return;
    }
    list.innerHTML = _cargadoresEdit.map((c, i) => {
      const cn = _normalizarCargador(c);
      _cargadoresEdit[i] = cn;
      return _renderCargadorBlock(cn, i, true);
    }).join('');
  }

  // [v2.13.3] Helper compartido — render del bloque de cargador con grid
  // de chips por carreta. isEdit=true para versión edición (autoguardado),
  // false para creación. Funciones del onclick siempre en PreingresosView.
  function _renderCargadorBlock(cn, i, isEdit) {
    const sufijo = isEdit ? 'Edit' : '';
    const chips = (cn.estados || []).map((est, j) => {
      const emoji = EMOJI_CARGA[est] || '🟢';
      const clase = 'carreta-chip carreta-chip-' + String(est || 'LLENA').toLowerCase();
      return `<button onclick="PreingresosView.toggleEstadoCarreta${sufijo}(${i}, ${j})"
                      class="${clase}" title="${LABEL_CARGA[est] || 'Llena'}">
        <span class="cc-emoji">${emoji}</span>
        <span class="cc-num">${j + 1}</span>
      </button>`;
    }).join('');
    return `<div class="cargador-block">
      <div class="cargador-head">
        <span class="cargador-ico">🛺</span>
        <span class="cargador-nombre">${escHtml(cn.nombre)}</span>
        <span class="cargador-resumen">${cn.carretas} carreta${cn.carretas === 1 ? '' : 's'}</span>
        <button onclick="PreingresosView.quitarCargador${sufijo}(${i})" class="cargador-quitar" title="Quitar cargador">×</button>
      </div>
      <div class="carretas-grid">
        ${chips}
        <button onclick="PreingresosView.agregarCarreta${sufijo}(${i})" class="carreta-add" title="Agregar carreta">+</button>
      </div>
    </div>`;
  }

  // [v2.13.110 BUG FIX] Autoguardado robusto: debounce + lock + idempotencyKey.
  //
  // PROBLEMA PREVIO:
  //   Cada click en cargadores disparaba 1 API.actualizarPreingreso inmediato.
  //   Si el operario hacía 5 clicks rápidos = 5 requests en vuelo. Las
  //   respuestas llegaban fuera de orden → la última en llegar (puede ser
  //   la 1ra mandada) pisaba las posteriores. Resultado: cargadores se
  //   "borraban", aparecía el estado de hace 5 saves.
  //   .catch(()=>{}) silenciaba errores → operario nunca enterado.
  //
  // FIX:
  //   - Debounce 700ms: solo se manda 1 request después de que el operario
  //     deja de tocar cargadores por 700ms.
  //   - Lock _saving: si llega otro pedido mientras hay uno en vuelo, queda
  //     pendiente; al volver el response, se dispara el último.
  //   - marcarCargadoresPendiente: protege el `cargadores` local del merge de
  //     polling (60s) hasta que el backend confirme. [v2.13.173]
  //   - Errores SÍ se muestran con toast (.catch no silencioso).
  // NOTA: la idempotencia real la da `localId` (lo inyecta API.post porque
  // actualizarPreingreso está en _IDEMPOTENT_ACTIONS). Cada save lleva un
  // localId nuevo → no se deduplican entre sí, que es justo lo que queremos.
  let _autoguardarTimer = null;
  let _autoguardarSaving = false;
  let _autoguardarPending = false;

  function _autoguardarCargadores() {
    if (!_editItem) return;
    clearTimeout(_autoguardarTimer);
    _autoguardarTimer = setTimeout(_ejecutarAutoguardarCargadores, 700);
  }

  async function _ejecutarAutoguardarCargadores() {
    if (!_editItem) return;
    if (_autoguardarSaving) {
      _autoguardarPending = true;
      return;
    }
    _autoguardarSaving = true;
    // [v2.13.111 AUDIT FIX C] Capturar idPreingreso antes del await.
    const idPreingreso = _editItem.idPreingreso;
    try {
      const cargadores = JSON.stringify(_cargadoresEdit);
      // [v2.13.173] Patch + flag pendiente ANTES del await: el cache local es
      // la fuente de verdad mientras el backend confirma, y el merge de polling
      // no debe revertirlo.
      OfflineManager.patchPreingresosCache(idPreingreso, { cargadores });
      OfflineManager.marcarCargadoresPendiente(idPreingreso);
      const r = await API.actualizarPreingreso({ idPreingreso, cargadores });
      if (r && r.ok === false) {
        try { toast('⚠ No se guardaron cargadores: ' + (r.error || ''), 'warn', 4000); } catch(_){}
      } else {
        if (_editItem && _editItem.idPreingreso === idPreingreso) {
          _editItem.cargadores = cargadores;
        }
      }
    } catch(e) {
      try { toast('⚠ Sin conexión al guardar cargadores', 'warn', 3500); } catch(_){}
    } finally {
      _autoguardarSaving = false;
      if (_autoguardarPending) {
        _autoguardarPending = false;
        setTimeout(_ejecutarAutoguardarCargadores, 0);
      }
    }
  }

  // [v2.13.111 AUDIT FIX D] Flush REAL — espera al save en curso antes
  // de lanzar el siguiente con el estado actual.
  async function _flushAutoguardarCargadores() {
    clearTimeout(_autoguardarTimer);
    const MAX_ESPERA_MS = 8000;
    const ini = Date.now();
    while (_autoguardarSaving && (Date.now() - ini) < MAX_ESPERA_MS) {
      await new Promise(res => setTimeout(res, 60));
    }
    if (_editItem) await _ejecutarAutoguardarCargadores();
  }

  // [v2.13.110 BUG FIX] Autoguardado de comentario + tags + monto.
  //
  // PROBLEMA PREVIO:
  //   El comentario, los tags (Comprobante sí/no, Completo sí/no) y el monto
  //   solo se persistían cuando el operario apretaba "GUARDAR" en el modal.
  //   Si solo editaba cargadores (que sí tienen autoguardado) y cerraba el
  //   sheet, el comentario tipeado se PERDÍA. El operario interpretaba
  //   "el comentario se borró", cuando en realidad nunca llegó al backend.
  //
  // FIX:
  //   - Listener oninput en textarea/inputs con debounce 1200ms
  //   - Reconstruye el comentario completo (tags + texto libre) y lo manda
  //   - Mismo lock/coalescing que cargadores para evitar races
  //   - NO dispara reimpresión de aviso (eso solo en guardarEdicion explícito)
  let _autoguardarMetaTimer = null;
  let _autoguardarMetaSaving = false;
  let _autoguardarMetaPending = false;

  function _autoguardarMeta() {
    if (!_editItem) return;
    clearTimeout(_autoguardarMetaTimer);
    _autoguardarMetaTimer = setTimeout(_ejecutarAutoguardarMeta, 1200);
  }

  async function _ejecutarAutoguardarMeta() {
    if (!_editItem) return;
    if (_autoguardarMetaSaving) { _autoguardarMetaPending = true; return; }
    _autoguardarMetaSaving = true;
    // [v2.13.111 AUDIT FIX C] Capturar idPreingreso ANTES del await.
    // Después del await, _editItem puede haber sido nulleado por
    // cerrarDetalle/abrir otro preingreso → crash 'Cannot read of null'.
    const idPreingreso = _editItem.idPreingreso;
    try {
      const textoExtra = (document.getElementById('piEditComentario')?.value || '').trim();
      const partes = [];
      if (_tagsEdit.comp)  partes.push('Comprobante: ' + (_tagsEdit.comp === 'si' ? 'Sí' : 'No'));
      if (_tagsEdit.compl) partes.push('Completo: ' + (_tagsEdit.compl === 'si' ? 'Sí' : 'No'));
      if (textoExtra) partes.push(textoExtra);
      const comentario = partes.join(' | ');
      const idProveedor = document.getElementById('piEditProv')?.value || _editItem.idProveedor;

      // [v2.13.173 BUG FIX] No forzar monto=0 cuando el tag "Comprobante" está
      // SIN setear (null): eso borraba un monto existente con solo editar el
      // comentario. Solo tocamos monto si el operador definió el tag —
      //   comp==='si' → monto del input · comp==='no' → 0 explícito ·
      //   comp===null → omitir (el backend preserva el valor existente).
      const params = { idPreingreso, idProveedor, comentario };
      if (_tagsEdit.comp === 'si')      params.monto = parseFloat(document.getElementById('piEditMonto')?.value) || 0;
      else if (_tagsEdit.comp === 'no') params.monto = 0;

      // [v2.13.173] Patch + flag pendiente ANTES del await: el cache local es la
      // fuente de verdad mientras el backend confirma; el merge de polling no
      // debe revertir comentario/monto/idProveedor recién editados.
      const patch = { comentario, idProveedor };
      if (params.monto !== undefined) patch.monto = params.monto;
      OfflineManager.patchPreingresosCache(idPreingreso, patch);
      OfflineManager.marcarPreingresoPendiente(idPreingreso, Object.keys(patch));

      const r = await API.actualizarPreingreso(params);
      if (r && r.ok === false) {
        try { toast('⚠ No se guardó la edición: ' + (r.error || ''), 'warn', 4000); } catch(_){}
      } else {
        // Sincronizar _editItem solo si SIGUE siendo el mismo preingreso
        // (el operario puede haber cambiado a otro durante el await).
        if (_editItem && _editItem.idPreingreso === idPreingreso) {
          _editItem.comentario = comentario;
          if (params.monto !== undefined) _editItem.monto = params.monto;
          _editItem.idProveedor = idProveedor;
        }
      }
    } catch(e) {
      try { toast('⚠ Sin conexión al autoguardar edición', 'warn', 3500); } catch(_){}
    } finally {
      _autoguardarMetaSaving = false;
      if (_autoguardarMetaPending) {
        _autoguardarMetaPending = false;
        setTimeout(_ejecutarAutoguardarMeta, 0);
      }
    }
  }

  // [v2.13.111 AUDIT FIX D] Flush REAL — espera a que el save en curso
  // termine antes de lanzar el siguiente con el estado actual. La versión
  // anterior llamaba _ejecutarAutoguardarMeta que detectaba _saving=true
  // y volvía sin esperar → guardarEdicion creía haber flusheado pero no.
  async function _flushAutoguardarMeta() {
    clearTimeout(_autoguardarMetaTimer);
    const MAX_ESPERA_MS = 8000;
    const ini = Date.now();
    // Esperar a que termine cualquier save en curso (max 8s).
    while (_autoguardarMetaSaving && (Date.now() - ini) < MAX_ESPERA_MS) {
      await new Promise(res => setTimeout(res, 60));
    }
    // Ahora ejecutar UN último save con el estado actual.
    if (_editItem) await _ejecutarAutoguardarMeta();
  }

  // Listeners de input — se enganchan cuando abrimos el sheet de edición.
  function _engancharAutoguardarMeta() {
    const com   = document.getElementById('piEditComentario');
    const monto = document.getElementById('piEditMonto');
    const prov  = document.getElementById('piEditProv');
    if (com   && !com._whAuto)   {
      com.addEventListener('input',  _autoguardarMeta);
      com.addEventListener('blur',   () => { _flushAutoguardarMeta(); });
      com._whAuto   = 1;
    }
    if (monto && !monto._whAuto) {
      monto.addEventListener('input', _autoguardarMeta);
      monto.addEventListener('blur',  () => { _flushAutoguardarMeta(); });
      monto._whAuto = 1;
    }
    if (prov  && !prov._whAuto)  {
      prov.addEventListener('change', _autoguardarMeta);
      prov._whAuto  = 1;
    }
    // Overlay: si el operario clickea fuera del sheet, flush antes de cerrar.
    const overlay = document.getElementById('overlayDetallePI');
    if (overlay && !overlay._whAuto) {
      overlay.addEventListener('click', async () => {
        // [v2.13.177] Esperar a que terminen los autoguardados ANTES de avisar y
        // limpiar — si nulleáramos _editItem antes, un save pendiente se perdería.
        try { await _flushAutoguardarMeta(); } catch(_){}
        try { await _flushAutoguardarCargadores(); } catch(_){}
        // [v2.13.178 BUG FIX] Si durante el await el operador REABRIÓ el detalle
        // (mismo u otro preingreso), el sheet vuelve a estar 'open' → NO limpiar,
        // o nullearíamos el estado del modal recién abierto.
        const sh = document.getElementById('sheetDetallePI');
        if (sh && sh.classList.contains('open')) return;
        _avisarSiCambiosSinEnviar();   // banner "sin avisar" si cerró con cambios
        _limpiarEstadoEdicion();       // revoca blobs + resetea estado fantasma
      });
      overlay._whAuto = 1;
    }
  }

  // [v2.13.3] LEGACY compat — alias para HTML cacheado viejo
  // [v2.13.175] LEGACY compat — delta>0 agrega, delta<0 quita la última.
  function cambiarCarretasEdit(idx, delta) {
    if (!_cargadoresEdit[idx]) return;
    if (delta > 0) { agregarCarretaEdit(idx); return; }
    CarretaCiclo.quitarUltima(_cargadoresEdit, idx);
    _renderCargadoresEdit();
    _autoguardarCargadores();
  }

  // [v2.13.175] Ciclo de carretas (edición) — delega en el manejador único.
  function toggleEstadoCarretaEdit(idx, carretaIdx) {
    _toggleCarretaArray(_cargadoresEdit, idx, carretaIdx, _renderCargadoresEdit, _autoguardarCargadores, '#piCargadoresList');
  }

  function agregarCarretaEdit(idx) {
    _agregarCarretaArray(_cargadoresEdit, idx, _renderCargadoresEdit, _autoguardarCargadores, '#piCargadoresList');
  }

  // [v2.13.175] Primitivas de animación CENTRALIZADAS (antes había
  // _flashCarretaChip y _desintegrarCarreta duplicando selectores). Operan
  // sobre un elemento chip ya resuelto, así sirven igual para los flujos por
  // array (creación/edición) como para el modal del día (chips anidados por pi).
  function _chipEnLista(selector, cargIdx, carretaIdx) {
    const blocks = document.querySelectorAll(selector + ' .cargador-block');
    return (blocks[cargIdx] && blocks[cargIdx].querySelectorAll('.carreta-chip')[carretaIdx]) || null;
  }
  function _flashChipEl(chip, esNueva) {
    if (!chip) return;
    chip.classList.remove('cc-flash', 'cc-flash-nueva');
    void chip.offsetWidth; // force reflow
    chip.classList.add(esNueva ? 'cc-flash-nueva' : 'cc-flash');
  }
  // Anima la desintegración del chip + SFX/vibración, luego ejecuta onDone (que
  // muta y re-renderiza). Si no hay chip, ejecuta onDone igual (no zombie).
  function _desintegrarChipEl(chip, onDone) {
    _desintegrarSfx();
    try { if (navigator.vibrate) navigator.vibrate([15, 30, 10]); } catch(_) {}
    if (!chip) { onDone(); return; }
    chip.classList.add('cc-desintegrar');
    setTimeout(onDone, 380);
  }
  // Dispara el SFX correcto según el descriptor de cambio de CarretaCiclo.
  function _sfxCarreta(cambio) {
    if (!cambio) return;
    if (cambio.accion === 'add')         _carretaAddSfx();
    else if (cambio.accion === 'estado') _carretaSfx(cambio.estado);
    // 'eliminar' → el SFX lo pone _desintegrarChipEl
  }

  // [v2.13.175] Manejador ÚNICO de toggle para flujos basados en array
  // (creación y edición). Recibe el array, índices, render, persist (o null) y
  // el selector del contenedor para las animaciones.
  function _toggleCarretaArray(arr, cargIdx, carretaIdx, render, persist, selector) {
    const c = arr[cargIdx];
    if (!c || !Array.isArray(c.estados)) return;
    if (CarretaCiclo.siguiente(c.estados[carretaIdx] || 'LLENA').accion === 'eliminar') {
      _desintegrarChipEl(_chipEnLista(selector, cargIdx, carretaIdx), () => {
        CarretaCiclo.avanzar(arr, cargIdx, carretaIdx);
        render(); if (persist) persist();
      });
      return;
    }
    const cambio = CarretaCiclo.avanzar(arr, cargIdx, carretaIdx);
    _sfxCarreta(cambio);
    render(); if (persist) persist();
    requestAnimationFrame(() => _flashChipEl(_chipEnLista(selector, cargIdx, carretaIdx), false));
  }
  function _agregarCarretaArray(arr, cargIdx, render, persist, selector) {
    const cambio = CarretaCiclo.agregar(arr, cargIdx);
    if (!cambio) return;
    _sfxCarreta(cambio);
    render(); if (persist) persist();
    requestAnimationFrame(() => _flashChipEl(_chipEnLista(selector, cargIdx, cambio.carretaIdx), true));
  }

  // [v2.13.4] SFX descendente para la desintegración (whoosh corto)
  function _desintegrarSfx() {
    try {
      const ctx = window.AudioContext ? new AudioContext() : null;
      if (!ctx) return;
      const o = ctx.createOscillator();
      const g = ctx.createGain();
      o.type = 'sawtooth';
      o.frequency.setValueAtTime(420, ctx.currentTime);
      o.frequency.exponentialRampToValueAtTime(60, ctx.currentTime + 0.32);
      g.gain.setValueAtTime(0.08, ctx.currentTime);
      g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.34);
      o.connect(g); g.connect(ctx.destination);
      o.start(); o.stop(ctx.currentTime + 0.35);
    } catch(_) {}
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
    // [v2.13.3] Default: 1 carreta LLENA
    _cargadoresEdit.push({ id, nombre, carretas: 1, estados: ['LLENA'] });
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

  // [v2.13.7] Snapshot de los 5 campos críticos espejo del backend (_snapshotAvisoFromPI)
  function _snapshotAvisoLocal(idProveedor, monto, tagsObj, textoLibre) {
    return {
      idProveedor:     String(idProveedor || ''),
      monto:           parseFloat(monto) || 0,
      tagComp:         tagsObj.comp  || null,
      tagCompl:        tagsObj.compl || null,
      comentarioLibre: String(textoLibre || '').trim()
    };
  }
  function _diffSnapAvisoLocal(a, b) {
    if (!a || !b) return [];
    const keys = ['idProveedor','monto','tagComp','tagCompl','comentarioLibre'];
    return keys.filter(k => {
      const va = a[k], vb = b[k];
      if (k === 'monto') return (parseFloat(va) || 0) !== (parseFloat(vb) || 0);
      return String(va == null ? '' : va) !== String(vb == null ? '' : vb);
    });
  }

  // [v2.13.174] Al cerrar el detalle, si hay cambios persistidos (autoguardados)
  // que difieren de lo último avisado a cajas, recordar al operador que pulse
  // "Guardar y avisar". El chip "⚠ sin avisar" de la card es el banner persistente.
  function _avisarSiCambiosSinEnviar() {
    if (!_editItem) return;
    let snap = null;
    try {
      snap = _editItem.snapshotAviso
        ? (typeof _editItem.snapshotAviso === 'string' ? JSON.parse(_editItem.snapshotAviso) : _editItem.snapshotAviso)
        : null;
    } catch(_) {}
    if (!snap) return; // sin snapshot impreso no hay base de comparación
    const idProveedor = document.getElementById('piEditProv')?.value || _editItem.idProveedor;
    const textoExtra  = (document.getElementById('piEditComentario')?.value || '').trim();
    let monto;
    if (_tagsEdit.comp === 'si')      monto = parseFloat(document.getElementById('piEditMonto')?.value) || 0;
    else if (_tagsEdit.comp === 'no') monto = 0;
    else                              monto = parseFloat(_editItem.monto) || 0;
    const actual = _snapshotAvisoLocal(idProveedor, monto, _tagsEdit, textoExtra);
    if (_diffSnapAvisoLocal(snap, actual).length > 0) {
      toast('⚠ Guardado, pero NO avisado a cajas — abre y pulsa "Guardar y avisar"', 'warn', 5500);
      try { if (typeof SoundFX !== 'undefined' && SoundFX.warn) SoundFX.warn(); } catch(_){}
    }
  }

  // [v2.13.177] Limpieza al abandonar el detalle (cerrar por overlay): revoca
  // los blobs de fotos NO guardadas (evita fuga de memoria) y resetea el estado
  // de edición para que no quede "fantasma" entre preingresos. Se llama SOLO
  // después de flushear los autoguardados (no antes — nullear _editItem cortaría
  // un save pendiente).
  function _limpiarEstadoEdicion() {
    try { _fotosNuevas.forEach(f => f && f.objectUrl && URL.revokeObjectURL(f.objectUrl)); } catch(_){}
    _fotosNuevas    = [];
    _fotosEdit      = [];
    _cargadoresEdit = [];
    _tagsEdit       = { comp: null, compl: null };
    _editItem       = null;
  }

  // [v2.13.175] Preview óptimista del aviso a cajas — espejo en pantalla del
  // ticket comparativo ANTES→AHORA. NO bloquea ni condiciona la impresión:
  // aparece al instante con datos locales, auto-cierra y se puede descartar;
  // la impresión real viaja en paralelo. Self-contained (inyecta su CSS 1 vez).
  function _ensurePreviewAvisoCss() {
    if (document.getElementById('pavStyles')) return;
    const st = document.createElement('style');
    st.id = 'pavStyles';
    st.textContent = `
    #previewAvisoOverlay{position:fixed;inset:0;z-index:10050;display:flex;align-items:flex-end;justify-content:center;background:rgba(2,6,23,.55);backdrop-filter:blur(2px);animation:pavFade .18s ease}
    #previewAvisoOverlay.pav-closing{animation:pavFade .18s ease reverse forwards}
    .pav-card{width:100%;max-width:460px;margin:0 8px 8px;background:#0f172a;border:1px solid #1e293b;border-radius:18px 18px 14px 14px;box-shadow:0 -8px 40px rgba(0,0,0,.5);overflow:hidden;animation:pavUp .26s cubic-bezier(.2,.9,.25,1)}
    .pav-head{display:flex;align-items:center;gap:10px;padding:14px 16px 10px;border-bottom:1px solid #1e293b}
    .pav-ico{font-size:20px}
    .pav-tit{font-weight:800;font-size:14px;color:#e2e8f0;letter-spacing:.2px}
    .pav-sub{font-size:11px;color:#64748b;margin-top:1px}
    .pav-body{padding:8px 14px 4px;max-height:46vh;overflow-y:auto}
    .pav-row{padding:9px 4px;border-bottom:1px dashed #1e293b}
    .pav-row:last-child{border-bottom:0}
    .pav-campo{font-size:10px;font-weight:800;letter-spacing:.6px;color:#fbbf24;text-transform:uppercase;margin-bottom:3px}
    .pav-diff{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
    .pav-antes{font-size:13px;color:#94a3b8;text-decoration:line-through;text-decoration-color:#475569;opacity:.8;word-break:break-word}
    .pav-flecha{color:#f59e0b;font-weight:900;flex-shrink:0}
    .pav-ahora{font-size:13px;color:#34d399;font-weight:700;word-break:break-word}
    .pav-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:10px 14px 13px}
    .pav-status{font-size:11px;color:#7dd3fc;display:flex;align-items:center;gap:6px}
    .pav-dot{width:7px;height:7px;border-radius:50%;background:#38bdf8;animation:pavPulse 1s infinite}
    .pav-ok{font-size:13px;font-weight:700;color:#0f172a;background:#34d399;border:0;border-radius:10px;padding:8px 16px}
    .pav-bar{height:3px;background:#f59e0b;width:100%;transform-origin:left;animation:pavBar var(--pav-dur,6s) linear forwards}
    @keyframes pavUp{from{transform:translateY(28px);opacity:.4}to{transform:translateY(0);opacity:1}}
    @keyframes pavFade{from{opacity:0}to{opacity:1}}
    @keyframes pavPulse{0%,100%{opacity:1}50%{opacity:.25}}
    @keyframes pavBar{from{transform:scaleX(1)}to{transform:scaleX(0)}}`;
    document.head.appendChild(st);
  }

  function _cerrarPreviewAviso() {
    const ov = document.getElementById('previewAvisoOverlay');
    if (!ov) return;
    if (ov._pavTimer) { clearTimeout(ov._pavTimer); ov._pavTimer = null; }
    ov.classList.add('pav-closing');
    setTimeout(() => ov.remove(), 180);
  }

  function _mostrarPreviewAviso(idPreingreso, provNombre, cambios, snapAnt, snapAct) {
    if (!cambios || !cambios.length) return;
    _ensurePreviewAvisoCss();
    const old = document.getElementById('previewAvisoOverlay');
    if (old) old.remove();
    const lblTag = v => v === 'si' ? 'Sí' : v === 'no' ? 'No' : '—';
    const filas = cambios.map(k => {
      let tit, antes, ahora;
      if (k === 'idProveedor')         { tit = 'Proveedor';   antes = _getProveedorNombre(snapAnt.idProveedor) || '—'; ahora = provNombre || '—'; }
      else if (k === 'monto')          { tit = 'Monto';       antes = 'S/ ' + (parseFloat(snapAnt.monto) || 0).toFixed(2); ahora = 'S/ ' + (parseFloat(snapAct.monto) || 0).toFixed(2); }
      else if (k === 'tagComp')        { tit = 'Comprobante'; antes = lblTag(snapAnt.tagComp); ahora = lblTag(snapAct.tagComp); }
      else if (k === 'tagCompl')       { tit = 'Completo';    antes = lblTag(snapAnt.tagCompl); ahora = lblTag(snapAct.tagCompl); }
      else if (k === 'comentarioLibre'){ tit = 'Comentario';  antes = snapAnt.comentarioLibre || '(vacío)'; ahora = snapAct.comentarioLibre || '(vacío)'; }
      else return '';
      return `<div class="pav-row">
        <div class="pav-campo">${escHtml(tit)}</div>
        <div class="pav-diff">
          <span class="pav-antes">${escHtml(String(antes))}</span>
          <span class="pav-flecha">→</span>
          <span class="pav-ahora">${escHtml(String(ahora))}</span>
        </div>
      </div>`;
    }).filter(Boolean).join('');

    const ov = document.createElement('div');
    ov.id = 'previewAvisoOverlay';
    ov.addEventListener('click', e => { if (e.target === ov) _cerrarPreviewAviso(); });
    ov.innerHTML = `
      <div class="pav-card" onclick="event.stopPropagation()">
        <div class="pav-bar" style="--pav-dur:6s"></div>
        <div class="pav-head">
          <span class="pav-ico">🖨</span>
          <div style="flex:1;min-width:0">
            <div class="pav-tit">Avisando a cajas — cambios</div>
            <div class="pav-sub">${escHtml(provNombre || '')} · ${escHtml(idPreingreso || '')}</div>
          </div>
        </div>
        <div class="pav-body">${filas}</div>
        <div class="pav-foot">
          <span class="pav-status"><span class="pav-dot"></span>Enviando a impresora…</span>
          <button class="pav-ok" onclick="PreingresosView.cerrarPreviewAviso()">Entendido</button>
        </div>
      </div>`;
    document.body.appendChild(ov);
    try { if (typeof SoundFX !== 'undefined') (SoundFX.ping || SoundFX.click)?.(); } catch(_){}
    // Auto-cierre alineado con la barra de progreso (6s). El operador puede
    // cerrarlo antes; la impresión NO depende de esto.
    ov._pavTimer = setTimeout(_cerrarPreviewAviso, 6000);
  }

  // ── Guardar edición (optimista) ──────────────────────────
  async function guardarEdicion() {
    if (!_editItem) return;

    // [v2.13.110] Flush autoguardados pendientes ANTES de guardar.
    // Garantiza que cualquier debounce en cola se ejecute primero, evitando
    // que un save manual se contradiga con un autoguardado tardío.
    try { await _flushAutoguardarMeta(); } catch(_){}
    try { await _flushAutoguardarCargadores(); } catch(_){}

    // Capturar datos y estado antes de cerrar
    const idPreingreso  = _editItem.idPreingreso;
    const idProveedor   = document.getElementById('piEditProv')?.value || _editItem.idProveedor;
    const textoExtra    = (document.getElementById('piEditComentario')?.value || '').trim();
    const partes = [];
    if (_tagsEdit.comp)  partes.push(`Comprobante: ${_tagsEdit.comp === 'si' ? 'Sí' : 'No'}`);
    if (_tagsEdit.compl) partes.push(`Completo: ${_tagsEdit.compl === 'si' ? 'Sí' : 'No'}`);
    if (textoExtra) partes.push(textoExtra);
    const comentario    = partes.join(' | ');
    // [v2.13.173 BUG FIX] No zerar un monto existente si "Comprobante" no está
    // definido (comp===null) — solo si el operador lo puso explícito en 'no'.
    let monto;
    if (_tagsEdit.comp === 'si')      monto = parseFloat(document.getElementById('piEditMonto')?.value) || 0;
    else if (_tagsEdit.comp === 'no') monto = 0;
    else                              monto = parseFloat(_editItem.monto) || 0;
    const fotosExistentes = [..._fotosEdit];
    const fotosNuevasCaptura = [..._fotosNuevas];

    // [v2.13.7] Calcular snapshot anterior (de lo que se imprimió la última vez) vs actual
    let snapAnterior = null;
    try {
      if (_editItem.snapshotAviso) {
        snapAnterior = typeof _editItem.snapshotAviso === 'string'
          ? JSON.parse(_editItem.snapshotAviso)
          : _editItem.snapshotAviso;
      }
    } catch(_) { snapAnterior = null; }
    // Fallback: si no hay snapshot (preingreso pre-v2.13.7), reconstruir desde el estado original al abrir
    if (!snapAnterior) {
      const tagsOrig = _tagsFromComentario(_editItem.comentario);
      const libreOrig = _textoLibreFromComentario(_editItem.comentario);
      snapAnterior = _snapshotAvisoLocal(_editItem.idProveedor, _editItem.monto, tagsOrig, libreOrig);
    }
    const snapActual = _snapshotAvisoLocal(idProveedor, monto, _tagsEdit, textoExtra);
    const cambios   = _diffSnapAvisoLocal(snapAnterior, snapActual);

    // Cerrar sheet y mostrar toast inmediatamente — sin esperar red
    cerrarSheet('sheetDetallePI');
    toast('Preingreso actualizado', 'ok');
    // [v2.13.174] Tick óptimista inmediato (el aviso a cajas confirma aparte con SoundFX.done)
    try { if (typeof SoundFX !== 'undefined') (SoundFX.savedTick || SoundFX.beepDouble || SoundFX.click)?.(); } catch(_){}
    try { if (navigator.vibrate) navigator.vibrate(12); } catch(_){}
    cargar(_filtroEstado, true);

    // [v2.13.175] Preview ÓPTIMISTA del diff (ANTES→AHORA). Es solo visual: la
    // impresión sale igual en el IIFE de abajo, sin importar si el operador lo
    // lee o lo cierra. Espejo en pantalla del ticket comparativo físico.
    if (cambios.length > 0) {
      _mostrarPreviewAviso(idPreingreso, _getProveedorNombre(idProveedor), cambios, snapAnterior, snapActual);
    }

    // Limpiar estado de edición
    _fotosNuevas = [];
    _fotosEdit   = [];

    // Subir fotos nuevas + actualizar en segundo plano
    (async () => {
      try { OfflineManager.setSubiendoFotos(true); } catch(_){}
      const todasFotos = [...fotosExistentes];
      for (let i = 0; i < fotosNuevasCaptura.length; i++) {
        try {
          const { b64, mime } = await _prepararFoto(fotosNuevasCaptura[i].file);
          const up = await API.subirFotoPreingreso({ idPreingreso, fotoBase64: b64, mimeType: mime, indice: fotosExistentes.length + i + 1 });
          if (up.ok && !up.offline && up.data?.url) {
            todasFotos.push({ url: up.data.url });
            // [v2.11.3] Registrar preview local como fallback ante Drive lento
            try {
              if (window.Photos && up.data.fileId && fotosNuevasCaptura[i].objectUrl) {
                Photos.registrarPreviewLocal(up.data.fileId, fotosNuevasCaptura[i].objectUrl);
              }
            } catch(_){}
          } else {
            console.warn('[FotosEdit] Error', i + 1, up.error || (up.offline ? 'sin conexión' : 'sin URL'));
          }
        } catch(e) { console.warn('[FotosEdit]', e); }
      }
      // [v2.11.3] NO revocar blob URLs — usados como fallback por photos.js
      try { OfflineManager.setSubiendoFotos(false); } catch(_){}
      const fotos = todasFotos.map(f => f.url).join(',');
      // [v2.13.173] Re-marcar pendiente: la subida de fotos pudo tardar >15s y
      // el flag del flush ya habría expirado; sin esto el poll revierte la meta.
      OfflineManager.patchPreingresosCache(idPreingreso, { comentario, monto, idProveedor, fotos });
      OfflineManager.marcarPreingresoPendiente(idPreingreso, ['comentario', 'monto', 'idProveedor']);
      await API.actualizarPreingreso({ idPreingreso, idProveedor, monto, comentario, fotos, usuario: window.WH_CONFIG.usuario })
        .catch(e => {
          console.warn('[EditPreingreso]', e);
          // [v2.13.179] No silenciar: si falla, las fotos quedaron en Drive pero
          // sin vincular a la fila. El operador debe saberlo para reintentar.
          try { toast('⚠ No se sincronizó la edición (fotos/datos) — reintenta', 'warn', 4500); } catch(_){}
        });
      // [v2.13.7] Solo reimprimir si cambió al menos uno de los 5 campos críticos
      if (cambios.length > 0) {
        // El patch óptimista del snapshotAviso (limpia el chip "sin avisar") lo
        // hace ahora _dispararAvisoCajeros centralizadamente. [v2.13.179]
        _dispararAvisoCajeros(idPreingreso, {
          silent: false,
          modoComparativo: true,
          snapshotAnterior: snapAnterior
        });
      } else {
        console.log('[EditPreingreso] sin cambios en campos críticos — no se reimprime aviso');
      }
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
    // Antes de crear, si el preingreso tiene fotos, dejar que el operador elija una
    const cached = OfflineManager.getPreingresosCache();
    const pi     = cached.find(x => x.idPreingreso === idPreingreso);
    const fotosPI = pi?.fotos
      ? String(pi.fotos).split(',').map(s => s.trim()).filter(Boolean)
      : [];

    if (fotosPI.length > 0) {
      // Selector: el callback recibe fileId o null
      FotoPicker.abrir(idPreingreso, fotosPI, (fileId) => {
        _ejecutarCrearGuia(idPreingreso, idProveedor, fileId);
      });
    } else {
      _ejecutarCrearGuia(idPreingreso, idProveedor, null);
    }
  }

  // [v2.13.173] Guard de re-entrada: evita inyectar dos cards optimistas (y dos
  // aprobaciones en vuelo) si el operador toca "+Guía" dos veces. El backend ya
  // es idempotente, pero esto evita el duplicado visual mientras viaja la red.
  const _aprobandoPI = new Set();

  async function _ejecutarCrearGuia(idPreingreso, idProveedor, fileIdFoto) {
    if (_aprobandoPI.has(idPreingreso)) return;
    _aprobandoPI.add(idPreingreso);
    const tempId     = 'G_tmp_' + Date.now();
    const provNombre = _getProveedorNombre(idProveedor);
    GuiasView.injectOptimisticGuia({ tempId, idProveedor, provNombre });
    App.nav('guias');

    const res = await API.aprobarPreingreso({ idPreingreso, usuario: window.WH_CONFIG.usuario })
      .catch(() => ({ ok: false, error: 'Sin conexión' }));

    if (res.ok) {
      toast(`Guía ${res.data.idGuia} creada`, 'ok');
      GuiasView.finalizeOptimisticGuia(tempId, res.data.idGuia, 'INGRESO_PROVEEDOR', provNombre);

      // Si el operador eligió una foto, copiarla a la guía en background
      if (fileIdFoto) {
        API.copiarFotoDePreingreso({
          idGuia: res.data.idGuia,
          idPreingreso,
          fileId: fileIdFoto
        }).catch(() => {});
      }
    } else {
      toast('Error al crear guía: ' + res.error, 'danger');
      GuiasView.removeOptimisticGuia(tempId);
    }
    _aprobandoPI.delete(idPreingreso);
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
          <span class="text-sm font-bold text-slate-100 truncate">${escHtml(provNombre)}</span>
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
    list.innerHTML = _cargadores.map((c, i) => {
      const cn = _normalizarCargador(c);
      _cargadores[i] = cn;
      return _renderCargadorBlock(cn, i, false);
    }).join('');
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
    // [v2.13.3] Default: 1 carreta LLENA
    _cargadores.push({ id, nombre, carretas: 1, estados: ['LLENA'] });
    _renderCargadores();
    document.getElementById('sheetCargadores')?.remove();
  }

  // [v2.13.3] LEGACY compat — alias de agregarCarreta/quitarCarreta.
  // Si delta > 0: agrega 1 carreta LLENA. Si delta < 0: quita la última.
  // El export sigue exponiendo `cambiarCarretas` por compat con HTML cacheado.
  // [v2.13.175] LEGACY compat — delta>0 agrega, delta<0 quita la última.
  function cambiarCarretas(idx, delta) {
    if (!_cargadores[idx]) return;
    if (delta > 0) { agregarCarreta(idx); return; }
    CarretaCiclo.quitarUltima(_cargadores, idx);
    _renderCargadores();
  }

  // [v2.13.175] Ciclo de carretas (creación) — delega en el manejador único.
  // Sin persist: el array vive en memoria hasta crear().
  function toggleEstadoCarreta(idx, carretaIdx) {
    _toggleCarretaArray(_cargadores, idx, carretaIdx, _renderCargadores, null, '#preCargadoresList');
  }

  function agregarCarreta(idx) {
    _agregarCarretaArray(_cargadores, idx, _renderCargadores, null, '#preCargadoresList');
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
    let res;
    try {
      res = await API.crearPreingreso({ idPreingreso, idProveedor, cargadores, monto, comentario, usuario: window.WH_CONFIG.usuario });
    } catch (e) {
      res = { ok: false, error: e?.message || 'Sin conexión' };
    }

    if (!res || !res.ok) {
      document.getElementById('optcard_' + tempId)?.remove();
      toast('Error: ' + (res?.error || 'desconocido'), 'danger');
      btn.disabled = false; btn.textContent = 'Registrar preingreso';
      return;
    }

    // Garantizar feedback visual incluso si _finalizeOptimisticCard lanza por
    // algún side effect (proveedores no en cache, etc.) — el card NUNCA queda
    // zombie "Registrando…".
    const idPreingresoReal = res.data?.idPreingreso || idPreingreso;
    try {
      _updateOptimisticId(tempId, idPreingresoReal);
      _finalizeOptimisticCard(idPreingresoReal, { idProveedor, monto });
    } catch (e) {
      console.warn('[Preingreso] finalize falló, removiendo optcard:', e);
      document.getElementById('optcard_' + idPreingresoReal)?.remove();
      document.getElementById('optcard_' + tempId)?.remove();
    }
    toast(res.offline
      ? `Preingreso guardado (offline) — sincronizando…`
      : `Preingreso ${idPreingresoReal} registrado`, 'ok');
    // [v2.13.174] Confirmación óptimista de registro (el aviso a cajas suena aparte)
    try { if (typeof SoundFX !== 'undefined') (SoundFX.savedTick || SoundFX.done || SoundFX.beepDouble)?.(); } catch(_){}
    try { if (navigator.vibrate) navigator.vibrate([10, 20, 10]); } catch(_){}

    // Inyectar en caché para que abrirDetalle no tenga que ir al GAS
    OfflineManager.inyectarPreingreso({
      idPreingreso: idPreingresoReal, idProveedor, cargadores, monto,
      comentario, estado: 'PENDIENTE',
      fecha: new Date().toISOString(), fotos: '',
      usuario: window.WH_CONFIG.usuario
    });
    // [v2.13.173] Si trae cargadores, blindarlos del merge de polling hasta
    // que el backend confirme la fila recién creada.
    if (cargadores) OfflineManager.marcarCargadoresPendiente(idPreingresoReal);
    btn.disabled = false; btn.textContent = 'Registrar preingreso';

    // 3. Subir fotos en segundo plano (no bloquea UI)
    const fotosCaptura = [..._fotosSeleccionadas];
    _fotosSeleccionadas = [];
    _subirFotosEnBackground(idPreingresoReal, fotosCaptura);

    // 4. Disparar aviso a cajeros activos en MosExpress (background, no bloquea)
    if (!res.offline) _dispararAvisoCajeros(idPreingresoReal);
  }

  // Aviso a cajas abiertas de MosExpress — no bloquea, muestra toast con resultado.
  // [v2.13.7] Soporta opts.modoComparativo + opts.snapshotAnterior para imprimir
  // un ticket diff (solo campos que cambiaron) en lugar del aviso completo.
  async function _dispararAvisoCajeros(idPreingreso, opts) {
    opts = opts || {};
    const base       = location.origin + location.pathname.replace(/\/[^/]*$/, '');
    const reporteUrl = `${base}/reporte.html?tipo=preingreso&id=${encodeURIComponent(idPreingreso)}`;
    if (!opts.silent) {
      toast(opts.modoComparativo ? '📤 Reimprimiendo cambios...' : '📤 Avisando a cajas...', 'info', 3000);
    }
    try {
      const payload = { idPreingreso, reporteUrl };
      if (opts.modoComparativo) {
        payload.modoComparativo  = true;
        payload.snapshotAnterior = JSON.stringify(opts.snapshotAnterior || {});
      }
      const res = await API.imprimirAvisoCajeros(payload);
      if (res.ok) {
        const okList = (res.data?.impresiones || []).filter(r => r.ok);
        if (okList.length) {
          const detalle = okList.map(r => `${r.vendedor || '—'} (${r.zona || '—'})`).join(', ');
          toast(`✓ Aviso enviado a: ${detalle}`, 'ok', 5000);
          // [v2.13.174] Confirmación sonora/háptica: el operador no mira la pantalla
          // mientras maneja la mercadería; el "done" le confirma que la caja recibió.
          try { if (typeof SoundFX !== 'undefined' && SoundFX.done) SoundFX.done(); } catch(_){}
          try { if (navigator.vibrate) navigator.vibrate([12, 28, 12]); } catch(_){}
        }
        const errList = (res.data?.impresiones || []).filter(r => !r.ok);
        if (errList.length) {
          // [v2.13.177] Nombrar las cajas que fallaron + acción concreta. El
          // snapshot ya se persiste con éxito parcial (para no re-spamear a las
          // que sí imprimieron), así que el operador debe reimprimir manual a
          // las que fallaron.
          const cuales = errList.map(r => `${r.vendedor || '—'} (${r.zona || '—'})`).join(', ');
          toast(`⚠ NO recibieron el aviso: ${cuales}. Usá "🖨 Reimprimir aviso a cajas".`, 'warn', 7000);
          try { if (typeof SoundFX !== 'undefined' && SoundFX.warn) SoundFX.warn(); } catch(_){}
        }
      } else if (res.error === 'NO_CHANGES') {
        // Modo comparativo y no hay cambios reales — no toast molesto
        console.log('[Aviso] sin cambios para reimprimir');
      } else if (res.error === 'NO_HAY_CAJEROS_ACTIVOS') {
        toast('⚠ No hay cajas abiertas — no se imprimió aviso', 'warn', 5000);
      } else {
        toast('Error aviso cajas: ' + (res.error || 'desconocido'), 'danger', 6000);
      }
      // [v2.13.179] Sincronizar el snapshotAviso LOCAL → limpia el chip
      // "⚠ sin avisar" sin esperar el poll de 60s. Aplica si la caja recibió
      // (ok) o si NO había cajas abiertas (el operador hizo su parte; no hay a
      // quién avisar — antes el chip quedaba pegado para siempre). En NO_CHANGES
      // el snapshot ya coincide; en error REAL de impresión NO se toca (el
      // cambio sigue pendiente de verdad).
      if (res.ok || res.error === 'NO_HAY_CAJEROS_ACTIVOS') {
        try {
          const pi = (OfflineManager.getPreingresosCache() || []).find(x => x.idPreingreso === idPreingreso);
          if (pi) {
            const snap = _snapshotAvisoLocal(pi.idProveedor, pi.monto,
              _tagsFromComentario(pi.comentario), _textoLibreFromComentario(pi.comentario));
            OfflineManager.patchPreingresosCache(idPreingreso, { snapshotAviso: JSON.stringify(snap) });
            PreingresosView.silentRefresh();
          }
        } catch(_){}
      }
      return res;
    } catch (e) {
      toast('Sin conexión — aviso no enviado', 'warn', 5000);
      return { ok: false, error: 'sin conexión' };
    }
  }

  // [Fix #3+5 v2.11.1] Subida resiliente + telemetría.
  // Cambios respecto al flow anterior:
  //   - Activa flag global "subiendo fotos" → precargarOperacional() omite
  //     el refresh de preingresos mientras dura la subida (Fix #2).
  //   - Después de CADA foto subida exitosamente, persiste el array
  //     parcial en backend Y en cache local. Si la red se corta a mitad
  //     de 8 fotos, las 5 primeras ya están en la hoja, no se pierden.
  //   - Logs detallados en consola (visible con __WH_DEBUG_FOTOS = true)
  //     para diagnosticar problemas de subida en producción.
  //   - Toast más claro al final con cantidad subida vs total.
  async function _subirFotosEnBackground(idPreingreso, fotos) {
    if (!fotos.length) return;
    console.log('[Fotos] inicio subida idPre=' + idPreingreso + ' total=' + fotos.length);
    try { OfflineManager.setSubiendoFotos(true); } catch(_){}
    const urls    = [];
    const fallidas = [];
    for (let i = 0; i < fotos.length; i++) {
      const indice = i + 1;
      try {
        console.log('[Fotos] subiendo ' + indice + '/' + fotos.length + '...');
        const { b64, mime } = await _prepararFoto(fotos[i].file);
        const up = await API.subirFotoPreingreso({
          idPreingreso, fotoBase64: b64, mimeType: mime, indice
        });
        if (up.ok && !up.offline && up.data?.url) {
          urls.push(up.data.url);
          // [v2.11.3] Registrar preview local como fallback. Si Drive aún no
          // generó el thumbnail (toma 30s-5min) o algún navegador no carga
          // la URL, el sistema de retry global de photos.js usa este blob
          // como reemplazo. Bypaseamos la dependencia de Drive para el
          // operador en su propio dispositivo.
          try {
            if (window.Photos && up.data.fileId && fotos[i].objectUrl) {
              Photos.registrarPreviewLocal(up.data.fileId, fotos[i].objectUrl);
            }
          } catch(_){}
          console.log('[Fotos] ✓ ' + indice + ' subida · url=' + up.data.url.substring(0, 60));
          // Persistir incrementalmente: backend + cache local. Si el
          // operador cierra la app o la red se cae después de esta foto,
          // las URLs acumuladas ya están seguras.
          try {
            const updRes = await API.actualizarFotosPreingreso({
              idPreingreso, fotos: urls.join(',')
            });
            if (updRes?.ok) {
              try { OfflineManager.patchPreingresosCache(idPreingreso, { fotos: urls.join(',') }); } catch(_){}
              console.log('[Fotos] ✓ ' + indice + ' persistida en hoja + cache (' + urls.length + ' total)');
            } else {
              console.warn('[Fotos] ✗ actualizarFotosPreingreso falló:', updRes?.error);
            }
          } catch(eUpd) {
            console.warn('[Fotos] ✗ actualizarFotosPreingreso threw:', eUpd?.message || eUpd);
          }
        } else {
          fallidas.push(indice);
          console.warn('[Fotos] ✗ foto ' + indice + ' sin URL:', up.error || (up.offline ? 'sin conexión (encolada)' : 'desconocido'));
        }
      } catch(e) {
        fallidas.push(indice);
        console.warn('[Fotos] ✗ excepción foto ' + indice + ':', e?.message || e);
      }
    }
    // [v2.11.3] NO revocamos los blob URLs inmediatamente: los mantenemos
    // vivos como fallback para el sistema de retry de photos.js cuando Drive
    // aún no generó el thumbnail. Se liberarán al recargar la página.
    // (Antes: fotos.forEach revocaba → si el operador veía la foto y Drive
    // estaba lento, le aparecía negra hasta que Drive procesara.)
    try { OfflineManager.setSubiendoFotos(false); } catch(_){}

    console.log('[Fotos] FIN · ok=' + urls.length + '/' + fotos.length + ' · fallidas=[' + fallidas.join(',') + ']');

    // Toast final con feedback claro
    if (urls.length === fotos.length) {
      toast(`✓ ${urls.length} fotos guardadas`, 'ok', 3000);
    } else if (urls.length > 0) {
      toast(`⚠ ${urls.length}/${fotos.length} fotos guardadas. ${fallidas.length} fallaron — reintentá desde editar`, 'warn', 7000);
    } else {
      toast('✗ Ninguna foto se pudo subir — reintenta desde editar', 'danger', 7000);
    }

    // Refrescar lista para que la UI muestre las fotos en cards
    try { PreingresosView.silentRefresh(); } catch(_){}
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
    let list = _tppFiltro ? todos.filter(p => p.estado === _tppFiltro) : todos;
    if (_tppBusq) list = list.filter(p =>
      _getProveedorNombre(p.idProveedor).toLowerCase().includes(_tppBusq) ||
      (p.idPreingreso || '').toLowerCase().includes(_tppBusq) ||
      (p.idProveedor  || '').toLowerCase().includes(_tppBusq)
    );
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

    const hoyKey  = _hoyPeru();
    const ayerKey = _ayerPeru();
    const months  = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    function _dateKey(p) {
      return (p.fecha && _diaPeru(p.fecha)) || '0000-00-00';
    }
    function _dateLabel(key) {
      if (!key || key === '0000-00-00') return 'Sin fecha';
      if (key === hoyKey)  return 'Hoy';
      if (key === ayerKey) return 'Ayer';
      const [yyyy, mm, dd] = key.split('-'); // key = 'YYYY-MM-DD' del día de Perú
      return yyyy === hoyKey.slice(0, 4)
        ? `${parseInt(dd, 10)} ${months[parseInt(mm, 10) - 1]}`
        : `${dd}/${mm}/${yyyy}`;
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

  function toggleTppFiltro() {
    const menu = document.getElementById('tppFilterMenu');
    if (!menu) return;
    menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
    if (menu.style.display !== 'none') {
      setTimeout(() => { document.addEventListener('click', _closeTppFiltroOutside, { once: true }); }, 0);
    }
  }

  function _closeTppFiltroOutside(e) {
    if (!e.target.closest('#tppFilterMenu') && !e.target.closest('#tppFilterBtn')) {
      const menu = document.getElementById('tppFilterMenu');
      if (menu) menu.style.display = 'none';
    }
  }

  function filtrarTpp(estado) {
    // Toggle: click active chip → back to todos
    _tppFiltro = (_tppFiltro === estado && estado !== '') ? '' : (estado || '');
    document.querySelectorAll('[data-tppfiltro]').forEach(b =>
      b.classList.toggle('active', _tppFiltro !== '' && b.dataset.tppfiltro === _tppFiltro));
    renderTppList();
  }

  function _searchFocusTpp(focused) {
    const toolbar = document.getElementById('tppToolbar');
    if (!toolbar) return;
    if (focused) toolbar.classList.add('srch-focused');
    else setTimeout(() => toolbar.classList.remove('srch-focused'), 160);
  }

  function compartirWA(idPreingreso) {
    const pi = (OfflineManager.getPreingresosCache() || []).find(x => x.idPreingreso === idPreingreso);
    if (!pi) return;
    const prov     = (OfflineManager.getProveedoresCache() || []).find(x => x.idProveedor === pi.idProveedor);
    const provNombre = prov?.nombre || pi.idProveedor || '—';
    let cargadores = [];
    try { cargadores = JSON.parse(pi.cargadores || '[]'); } catch {}
    // [v2.13.4] Sin tarifa. Desglose explícito por estado: 🟢llenas · 🟡medias · 🔴vacías
    const cargLines = cargadores.map(c => {
      const cn = _normalizarCargador(c);
      const nombre = cn.nombre || cn.id || String(cn);
      if (!cn.carretas) return `   • ${nombre}`;
      const r = _resumenEstadosCarretas(cn.estados);
      const dets = [];
      if (r.llenas > 0) dets.push(`🟢 ${r.llenas} llena${r.llenas === 1 ? '' : 's'}`);
      if (r.medias > 0) dets.push(`🟡 ${r.medias} media${r.medias === 1 ? '' : 's'}`);
      if (r.vacias > 0) dets.push(`🔴 ${r.vacias} casi vacía${r.vacias === 1 ? '' : 's'}`);
      const det = dets.join(' · ');
      return `   • *${nombre}* — ${cn.carretas} carreta${cn.carretas === 1 ? '' : 's'} (${det})`;
    });
    const url      = `${location.href.split('?')[0].replace(/index\.html$/, '').replace(/\/$/, '')}/reporte.html?tipo=preingreso&id=${encodeURIComponent(idPreingreso)}`;
    const lineas   = [
      `*📦 PREINGRESO ${idPreingreso}*`,
      `─────────────────────`,
      `🏪 *Proveedor:* ${provNombre}`,
      `💰 *Monto:* ${pi.monto ? 'S/ ' + fmt(pi.monto, 2) : '—'}`,
      `📅 *Fecha:* ${_fmtCorta(pi.fecha)}`,
      `📊 *Estado:* ${pi.estado || '—'}`,
    ];
    if (cargadores.length) {
      lineas.push(`👥 *Cargadores:*`);
      lineas.push(...cargLines);
    } else {
      lineas.push(`👥 *Cargadores:* —`);
    }
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

  function reimprimirAviso() {
    if (!_editItem) { toast('Abre un preingreso primero', 'warn'); return; }
    _dispararAvisoCajeros(_editItem.idPreingreso);
  }

  // ── Cargadores del día — modal consolidado (preingresos + guías) ──
  let _cargDiaState = { fecha: '', data: null };

  function _fmtFechaLabel(yyyymmdd) {
    if (!yyyymmdd || yyyymmdd === '0000-00-00') return 'Sin fecha';
    if (yyyymmdd === _hoyPeru())  return 'Hoy';
    if (yyyymmdd === _ayerPeru()) return 'Ayer';
    const d = new Date(yyyymmdd + 'T12:00:00'); // key ya correcto; noon = etiqueta TZ-segura
    if (isNaN(d)) return yyyymmdd;
    return d.toLocaleDateString('es-PE', { day:'2-digit', month:'long', year:'numeric' });
  }

  function abrirCargadoresDia(fechaKey) {
    _cargDiaState.fecha = fechaKey;
    document.getElementById('cargDiaFechaLbl').textContent = _fmtFechaLabel(fechaKey);
    _cargDiaState.data = _calcularCargadoresDelDia(fechaKey);
    _renderCargDiaContent();
    abrirSheet('sheetCargadoresDia');
  }

  // [v2.13.4] Modal del día editable. Chips clickeables por carreta (ciclo
  // LLENA→MEDIA→VACIA→eliminar), mismo SFX/animaciones del preingreso.
  // Edición persiste al backend usando el idPreingreso real.
  function _renderCargDiaContent() {
    const c = document.getElementById('cargDiaContenido');
    const d = _cargDiaState.data;
    if (!d) { c.innerHTML = ''; return; }
    if (!d.cargadores || !d.cargadores.length) {
      c.innerHTML = '<div class="text-xs text-slate-500 italic text-center py-6">Sin cargadores este día</div>';
      return;
    }
    const badges = (l, m, v) => {
      const p = [];
      if (l > 0) p.push(`<span class="cd-badge cd-badge-llena" title="Llenas">🟢 ${l}</span>`);
      if (m > 0) p.push(`<span class="cd-badge cd-badge-media" title="Medias">🟡 ${m}</span>`);
      if (v > 0) p.push(`<span class="cd-badge cd-badge-vacia" title="Casi vacías">🔴 ${v} c.vacía${v === 1 ? '' : 's'}</span>`);
      return p.join(' ');
    };
    const chipsHtml = (pi) => {
      const idPre = escAttr(pi.idPreingreso);
      const chips = (pi.estados || []).map((est, j) => {
        const emoji = EMOJI_CARGA[est] || '🟢';
        const clase = 'carreta-chip carreta-chip-' + String(est || 'LLENA').toLowerCase();
        return `<button onclick="PreingresosView.toggleEstadoCarretaDia('${idPre}', ${j})"
                        class="${clase}" title="${LABEL_CARGA[est] || 'Llena'}">
          <span class="cc-emoji">${emoji}</span>
          <span class="cc-num">${j + 1}</span>
        </button>`;
      }).join('');
      return `<div class="carretas-grid mt-1">
        ${chips}
        <button onclick="PreingresosView.agregarCarretaDia('${idPre}')" class="carreta-add" title="Agregar carreta">+</button>
      </div>`;
    };
    const cards = d.cargadores.map((cg, cgIdx) => `
      <div class="cargador-block" data-cg-idx="${cgIdx}">
        <div class="cargador-head">
          <span class="cargador-ico">🛺</span>
          <span class="cargador-nombre">${escHtml(cg.nombre)}</span>
          <span class="cargador-resumen">${cg.carretasTotal} carreta${cg.carretasTotal === 1 ? '' : 's'}</span>
        </div>
        <div class="text-[11px] mb-2">${badges(cg.llenasTotal||0, cg.mediasTotal||0, cg.vaciasTotal||0)}</div>
        <div class="space-y-2">
          ${cg.preingresos.map(pi => `
            <div data-id-pre="${escAttr(pi.idPreingreso)}" data-cg-id="${escAttr(cg.id)}">
              <div class="text-[11px] text-slate-400 flex items-center gap-2">
                <span class="font-mono text-slate-500">${escHtml(pi.idPreingreso)}</span>
                <span class="flex-1 truncate">${escHtml(pi.proveedor || '—')}</span>
                <span class="text-amber-300">${pi.carretas} cart</span>
              </div>
              ${chipsHtml(pi)}
            </div>`).join('')}
        </div>
      </div>`).join('');
    c.innerHTML = `<div id="cargDiaList">${cards}</div>
      <div class="rounded-lg px-3 py-3 mt-3"
           style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.3)">
        <div class="flex items-center justify-between">
          <div>
            <div class="text-xs text-slate-400">TOTAL DEL DÍA</div>
            <div class="text-[11px] text-slate-500">${d.preingresos} preingreso${d.preingresos !== 1 ? 's' : ''}</div>
          </div>
          <div class="text-3xl font-bold text-amber-300">${d.totalCarretas} <span class="text-xs text-slate-400">carretas</span></div>
        </div>
        <div class="mt-2 text-[11px]">${badges(d.totalLlenas||0, d.totalMedias||0, d.totalVacias||0)}</div>
        <div class="mt-1 text-[10px] text-slate-500 italic">Toca una carreta para cambiar su estado · 4to click la elimina</div>
      </div>`;
  }

  // [v2.13.173] Guardado robusto de cargadores para el modal del día.
  // Reemplaza el fire-and-forget anterior (.catch(()=>{})): con clicks rápidos
  // se mandaban varias escrituras concurrentes cuyo orden de LLEGADA al backend
  // no estaba garantizado → la última en llegar podía ser un estado viejo y
  // "perdía" el cambio. Ahora: debounce 600ms por preingreso + coalescing +
  // errores visibles. El flag pendiente (marcarCargadoresPendiente) ya lo pone
  // _mutarCargadorDelPreingreso para blindar el cache del merge de polling.
  // Mismo patrón que el autoguardado del modal de edición (v2.13.110/111).
  const _cargDiaSavers = {}; // idPreingreso -> { timer, saving, pending, ultimo }

  function _guardarCargadoresDia(idPreingreso, cargadoresStr) {
    let s = _cargDiaSavers[idPreingreso];
    if (!s) s = _cargDiaSavers[idPreingreso] = { timer: null, saving: false, pending: false, ultimo: '' };
    s.ultimo = cargadoresStr;
    clearTimeout(s.timer);
    s.timer = setTimeout(() => _ejecutarGuardarCargadoresDia(idPreingreso), 600);
  }

  async function _ejecutarGuardarCargadoresDia(idPreingreso) {
    const s = _cargDiaSavers[idPreingreso];
    if (!s) return;
    if (s.saving) { s.pending = true; return; }
    s.saving = true;
    // Re-marcar pendiente para cubrir la latencia del round-trip.
    OfflineManager.marcarCargadoresPendiente(idPreingreso);
    try {
      const r = await API.actualizarPreingreso({ idPreingreso, cargadores: s.ultimo });
      if (r && r.ok === false) {
        try { toast('⚠ No se guardaron carretas: ' + (r.error || ''), 'warn', 4000); } catch(_){}
      }
    } catch (e) {
      try { toast('⚠ Sin conexión al guardar carretas', 'warn', 3500); } catch(_){}
    } finally {
      s.saving = false;
      if (s.pending) { s.pending = false; setTimeout(() => _ejecutarGuardarCargadoresDia(idPreingreso), 0); }
    }
  }

  // [v2.13.4] Edición desde el modal del día — encuentra el preingreso real,
  // muta el cargador correspondiente, persiste y refresca. La carreta vacía
  // se desintegra con animación; si quedan 0 carretas el cargador desaparece.
  function _mutarCargadorDelPreingreso(idPreingreso, cargId, fnMutar) {
    const cache = OfflineManager.getPreingresosCache() || [];
    const pi = cache.find(p => p.idPreingreso === idPreingreso);
    if (!pi) { toast('Preingreso no encontrado en caché', 'warn'); return null; }
    let arr = [];
    try { arr = JSON.parse(pi.cargadores || '[]'); } catch { arr = []; }
    if (!Array.isArray(arr)) arr = [];
    const idx = arr.findIndex(c => String(c.id || c.idPersonal || c.nombre || '') === String(cargId));
    if (idx < 0) { toast('Cargador no encontrado en el preingreso', 'warn'); return null; }
    arr[idx] = _normalizarCargador(arr[idx]);
    const res = fnMutar(arr[idx], arr, idx);
    if (res === false) return null;
    // Si el cargador quedó sin carretas → eliminarlo del array
    if (arr[idx] && arr[idx].carretas === 0) arr.splice(idx, 1);
    const cargadoresStr = JSON.stringify(arr);
    // [v2.13.173] Cache local = fuente de verdad; flag pendiente protege del
    // merge de polling; el save real va debounced/coalesced (no fire-and-forget).
    OfflineManager.patchPreingresosCache(idPreingreso, { cargadores: cargadoresStr });
    OfflineManager.marcarCargadoresPendiente(idPreingreso);
    _guardarCargadoresDia(idPreingreso, cargadoresStr);
    return arr;
  }

  // [v2.13.175] Modal del día — mismo ciclo central (CarretaCiclo), pero la
  // mutación persiste por preingreso (_mutarCargadorDelPreingreso) y los chips
  // están anidados por pi, así que la animación se resuelve dentro de `cont`.
  function _refrescarCargDia() {
    _cargDiaState.data = _calcularCargadoresDelDia(_cargDiaState.fecha);
    _renderCargDiaContent();
  }
  // Re-consulta el chip tras el re-render (el DOM se reemplazó).
  function _chipDia(idPreingreso, carretaIdx) {
    const cont = document.querySelector(`#cargDiaList [data-id-pre="${idPreingreso}"]`);
    return cont ? cont.querySelectorAll('.carreta-chip')[carretaIdx] || null : null;
  }

  function toggleEstadoCarretaDia(idPreingreso, carretaIdx) {
    const cont = document.querySelector(`#cargDiaList [data-id-pre="${idPreingreso}"]`);
    if (!cont) return;
    const cargId = cont.getAttribute('data-cg-id');
    // Estado actual desde cache (fuente de verdad, no el DOM)
    const cache = OfflineManager.getPreingresosCache() || [];
    const pi = cache.find(p => p.idPreingreso === idPreingreso);
    if (!pi) return;
    let arr = []; try { arr = JSON.parse(pi.cargadores || '[]'); } catch {}
    const c0 = arr.find(c => String(c.id || c.idPersonal || c.nombre || '') === String(cargId));
    if (!c0 || !Array.isArray(c0.estados)) return;
    const paso = CarretaCiclo.siguiente(c0.estados[carretaIdx] || 'LLENA');

    if (paso.accion === 'eliminar') {
      _desintegrarChipEl(cont.querySelectorAll('.carreta-chip')[carretaIdx], () => {
        _mutarCargadorDelPreingreso(idPreingreso, cargId, (c) => {
          c.estados.splice(carretaIdx, 1);
          c.carretas = c.estados.length;
        });
        _refrescarCargDia();
      });
      return;
    }
    _mutarCargadorDelPreingreso(idPreingreso, cargId, (c) => { c.estados[carretaIdx] = paso.estado; });
    _carretaSfx(paso.estado);
    _refrescarCargDia();
    requestAnimationFrame(() => _flashChipEl(_chipDia(idPreingreso, carretaIdx), false));
  }

  function agregarCarretaDia(idPreingreso) {
    const cont = document.querySelector(`#cargDiaList [data-id-pre="${idPreingreso}"]`);
    if (!cont) return;
    const cargId = cont.getAttribute('data-cg-id');
    let nuevoIdx = 0;
    _mutarCargadorDelPreingreso(idPreingreso, cargId, (c) => {
      if (!Array.isArray(c.estados)) c.estados = [];
      c.estados.push('LLENA');
      c.carretas = c.estados.length;
      nuevoIdx = c.estados.length - 1;
    });
    _carretaAddSfx();
    _refrescarCargDia();
    requestAnimationFrame(() => _flashChipEl(_chipDia(idPreingreso, nuevoIdx), true));
  }

  // [v2.13.4] WhatsApp consolidado — desglose explícito 🟢/🟡/🔴 por cargador y preingreso
  function compartirCargadoresDiaWA() {
    const d = _cargDiaState.data;
    if (!d || !d.cargadores || !d.cargadores.length) {
      toast('Sin datos para compartir', 'warn'); return;
    }
    // [v2.13.6] '🔴 N' siempre con texto 'casi vacía' explícito — el emoji solo
    // confunde al cajero si no acompaña la palabra completa.
    const fmtDets = (l, m, v) => {
      const dets = [];
      if (l > 0) dets.push(`🟢 ${l}`);
      if (m > 0) dets.push(`🟡 ${m}`);
      if (v > 0) dets.push(`🔴 ${v} casi vacía${v === 1 ? '' : 's'}`);
      return dets.join(' · ');
    };
    const lineas = [
      `*🛺 CARGADORES — ${_fmtFechaLabel(d.fecha)}*`,
      `─────────────────────`,
    ];
    d.cargadores.forEach(cg => {
      const det = fmtDets(cg.llenasTotal||0, cg.mediasTotal||0, cg.vaciasTotal||0);
      lineas.push(`*${cg.nombre}*  —  ${cg.carretasTotal} carreta${cg.carretasTotal === 1 ? '' : 's'}   ${det}`);
      cg.preingresos.forEach(pi => {
        const detPi = fmtDets(pi.llenas||0, pi.medias||0, pi.vacias||0);
        lineas.push(`   • ${pi.idPreingreso}  ${pi.proveedor || '—'}  —  ${pi.carretas} cart  (${detPi})`);
      });
    });
    lineas.push(``, `━━━━━━━━━━━━━━━━━━━━━━━━`);
    lineas.push(`*TOTAL:* ${d.totalCarretas} carreta${d.totalCarretas === 1 ? '' : 's'}`);
    const dets = [];
    if ((d.totalLlenas||0) > 0) dets.push((d.totalLlenas) + ' 🟢 llenas');
    if ((d.totalMedias||0) > 0) dets.push((d.totalMedias) + ' 🟡 medias');
    if ((d.totalVacias||0) > 0) dets.push((d.totalVacias) + ' 🔴 casi vacías');
    if (dets.length) lineas.push(dets.join(' · '));
    lineas.push(`_InversionMos Warehouse_`);
    window.open('https://wa.me/?text=' + encodeURIComponent(lineas.join('\n')), '_blank');
  }

  async function imprimirCargadoresDia() {
    const fecha = _cargDiaState.fecha;
    if (!fecha) { toast('Abre primero el resumen del día', 'warn'); return; }
    const res = await PrintHub.imprimir('imprimirCargadoresDia', { fecha }, 'Cargadores del día ' + fecha).catch(() => null);
    if (res && res.ok) toast('Ticket impreso', 'ok');
    else if (res)      toast('Error: ' + (res?.error || 'No se pudo imprimir'), 'danger');
  }

  return { cargar, filtrar, toggleFiltro, _searchFocusPre, silentRefresh, buscar, buscarClear, crear, nuevo,
           abrirPanel, filtrarPanel, aprobarDesdePanel,
           toggleTag, toggleTagModal,
           onFotosSeleccionadas, quitarFoto, verFotos,
           onFotosEditSeleccionadas, quitarFotoEdit,
           abrirDetalle, guardarEdicion, crearGuiaDesde, crearGuiaRapido, reimprimirAviso,
           abrirCargadoresDia, compartirCargadoresDiaWA, imprimirCargadoresDia,
           toggleEstadoCarretaDia, agregarCarretaDia,
           filtrarProveedores, seleccionarProveedor, limpiarProveedor,
           abrirPickerCargador, agregarCargador, cambiarCarretas, quitarCargador, limpiarCargador,
           toggleEstadoCarreta, agregarCarreta, toggleEstadoCarretaEdit, agregarCarretaEdit,
           abrirPickerCargadorEdit, agregarCargadorEdit, cambiarCarretasEdit, quitarCargadorEdit,
           renderTppList, buscarEnPanel, limpiarBuscarPanel,
           toggleTppFiltro, filtrarTpp, _searchFocusTpp, compartirWA,
           cerrarPreviewAviso: _cerrarPreviewAviso };
})();

// ════════════════════════════════════════════════
// MERMAS VIEW
// ════════════════════════════════════════════════
const MermasView = (() => {
  let _filtro = 'EN_PROCESO';
  let _all    = [];
  let _selMerma = null;
  let _fotoBase64 = null;
  let _fotoMime   = '';

  async function cargar() {
    loading('listMermas', true);
    const res = await API.getMermas({ limit: 200 }).catch(() => ({ ok: false }));
    _all = res.ok ? res.data : [];
    // Calcular días en proceso para flag vencidas (>3 días)
    const ahora = Date.now();
    _all.forEach(m => {
      if (!m.fechaIngreso) { m.diasEnProceso = 0; m.vencida = false; return; }
      const ms = ahora - new Date(m.fechaIngreso).getTime();
      m.diasEnProceso = Math.floor(ms / (24 * 60 * 60 * 1000));
      m.vencida = String(m.estado || '').toUpperCase() === 'EN_PROCESO' && ms > (3 * 24 * 60 * 60 * 1000);
    });
    _renderTabs();
    _render();
  }

  function setFiltro(f) {
    _filtro = f;
    document.querySelectorAll('.merma-tab').forEach(b => b.classList.remove('tab-active'));
    const map = { 'EN_PROCESO': 'tabMermasProceso', 'RESUELTA': 'tabMermasResueltas', 'DESECHADA': 'tabMermasDesechadas' };
    const btn = document.getElementById(map[f]);
    if (btn) btn.classList.add('tab-active');
    _render();
  }

  function _renderTabs() {
    const enProceso = _all.filter(m => String(m.estado || '').toUpperCase() === 'EN_PROCESO').length;
    const el = document.getElementById('tabMermasNProceso');
    if (el) el.textContent = enProceso;
  }

  function _render() {
    const container = document.getElementById('listMermas');
    if (!container) return;
    const filtradas = _all.filter(m => String(m.estado || '').toUpperCase() === _filtro);
    if (!filtradas.length) {
      container.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin mermas en esta categoría</p>';
      return;
    }
    // Resolver descripciones desde caché
    const prods = OfflineManager.getProductosCache();
    const descMap = {};
    prods.forEach(p => {
      if (p.codigoBarra) descMap[String(p.codigoBarra)] = p.descripcion;
      if (p.idProducto)  descMap[String(p.idProducto)]  = p.descripcion;
    });

    container.innerHTML = filtradas.map(m => {
      const desc = descMap[String(m.codigoProducto || '').trim()] || m.codigoProducto;
      const respLabel = String(m.responsable || m.origen || '—');
      const fechaStr = m.fechaIngreso ? new Date(m.fechaIngreso).toLocaleDateString('es-PE') : '';
      const safeId = escAttr(m.idMerma);
      let actions = '';
      if (_filtro === 'EN_PROCESO') {
        actions = `
          <div class="flex gap-1.5 mt-2">
            <button onclick="MermasView.abrirResolver('${safeId}')"
                    class="btn btn-sm btn-primary text-xs px-3 py-1.5 flex-1">Resolver</button>
            ${m.foto ? `<button onclick="MermasView.verFoto('${escAttr(m.foto)}')"
                       class="btn btn-sm btn-outline text-xs px-3 py-1.5">📷</button>` : ''}
          </div>`;
      }
      let footer = '';
      if (_filtro === 'RESUELTA' || _filtro === 'DESECHADA') {
        footer = `
          <p class="text-[10.5px] text-slate-500 mt-1">
            ${m.cantidadReparada > 0 ? `<span class="text-emerald-400">✓ ${fmt(m.cantidadReparada, 1)} reparadas</span>` : ''}
            ${m.cantidadReparada > 0 && m.cantidadDesechada > 0 ? ' · ' : ''}
            ${m.cantidadDesechada > 0 ? `<span class="text-red-400">🗑 ${fmt(m.cantidadDesechada, 1)} desechadas</span>` : ''}
            ${m.idGuiaSalida ? ` · guía ${escHtml(m.idGuiaSalida)}` : ''}
          </p>
          ${m.observacionResolucion ? `<p class="text-[10.5px] text-slate-500 mt-0.5 italic">"${escHtml(m.observacionResolucion)}"</p>` : ''}`;
      }
      return `
      <div class="card-sm" style="${m.vencida ? 'border:1px solid rgba(245,158,11,.5);background:rgba(120,53,15,.08)' : ''}">
        <div class="flex items-start justify-between gap-2 mb-1">
          <div class="flex-1 min-w-0">
            <p class="font-semibold text-sm text-slate-100 truncate">${escHtml(desc)}</p>
            <p class="text-[10.5px] text-slate-500 font-mono">${escHtml(m.codigoProducto)}</p>
          </div>
          <span class="font-black text-base text-red-400 flex-shrink-0">${fmt(m.cantidadOriginal, 1)}</span>
        </div>
        <div class="flex items-center gap-2 text-[11px] text-slate-400">
          <span class="px-2 py-0.5 rounded-full" style="background:rgba(100,116,139,.18)">${escHtml(respLabel)}</span>
          <span>${escHtml(fechaStr)}</span>
          ${m.vencida ? `<span class="text-amber-400 font-bold">⚠ ${m.diasEnProceso}d sin resolver</span>` : ''}
        </div>
        ${m.motivo ? `<p class="text-[11px] text-slate-500 mt-1">${escHtml(m.motivo)}</p>` : ''}
        ${footer}
        ${actions}
      </div>`;
    }).join('');
  }

  function nueva() {
    _fotoBase64 = null; _fotoMime = '';
    document.getElementById('mermaCodigoProd').value = '';
    document.getElementById('mermaCantidad').value   = '';
    document.getElementById('mermaMotivo').value     = '';
    document.getElementById('mermaFotoLbl').textContent = '📷 Tomar foto';
    document.getElementById('mermaFotoPrev').innerHTML = '';
    // Inyectar zonas en el dropdown desde caché
    const zonas = OfflineManager.getZonasCache();
    const sel = document.getElementById('mermaResponsable');
    sel.innerHTML = `
      <option value="ALMACEN">ALMACÉN (interno)</option>
      <option value="RECEPCION">RECEPCIÓN (proveedor)</option>
      ${zonas.map(z => `<option value="${escAttr(z.idZona)}">${escHtml(z.nombre || z.idZona)}</option>`).join('')}`;
    abrirSheet('sheetMerma');
  }

  async function onFotoSeleccionada(ev) {
    const file = ev.target.files?.[0];
    if (!file) return;
    try {
      const { b64, mime } = await _prepFotoMerma(file);
      _fotoBase64 = b64;
      _fotoMime   = mime;
      const url = URL.createObjectURL(file);
      document.getElementById('mermaFotoLbl').textContent = '✓ Foto agregada — cambiar';
      document.getElementById('mermaFotoPrev').innerHTML =
        `<img src="${url}" style="max-height:120px;border-radius:8px;border:1px solid #334155">`;
    } catch(e) { toast('Error al procesar foto', 'warn'); }
  }

  // Comprime y convierte a base64 (~700px max)
  function _prepFotoMerma(file) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        const max = 700;
        let { width: w, height: h } = img;
        if (w > max || h > max) { const r = w > h ? max/w : max/h; w = w*r|0; h = h*r|0; }
        const c = document.createElement('canvas'); c.width = w; c.height = h;
        c.getContext('2d').drawImage(img, 0, 0, w, h);
        const dataUrl = c.toDataURL('image/jpeg', 0.78);
        URL.revokeObjectURL(img.src);
        resolve({ b64: dataUrl.split(',')[1], mime: 'image/jpeg' });
      };
      img.onerror = reject;
      img.src = URL.createObjectURL(file);
    });
  }

  async function crear() {
    const codigoProducto = document.getElementById('mermaCodigoProd').value.trim();
    const responsable    = document.getElementById('mermaResponsable').value;
    const cantidad       = parseFloat(document.getElementById('mermaCantidad').value);
    const motivo         = document.getElementById('mermaMotivo').value.trim();
    if (!codigoProducto || !cantidad || cantidad <= 0) {
      toast('Completa producto y cantidad', 'warn'); return;
    }
    if (!_fotoBase64) {
      toast('Foto obligatoria al registrar merma', 'warn'); return;
    }
    const btn = document.getElementById('btnRegistrarMerma');
    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const res = await API.registrarMerma({
        codigoProducto, responsable, motivo,
        cantidadOriginal: cantidad,
        fotoBase64: _fotoBase64, mimeType: _fotoMime,
        usuario: window.WH_CONFIG?.usuario || ''
      });
      if (res.ok) {
        toast('✓ Merma registrada en proceso', 'ok');
        cerrarSheet('sheetMerma');
        cargar();
      } else {
        toast('Error: ' + (res.error || 'desconocido'), 'danger', 5000);
      }
    } catch(e) { toast('Sin conexión', 'warn'); }
    finally { btn.disabled = false; btn.textContent = 'Registrar merma'; }
  }

  function abrirResolver(idMerma) {
    const m = _all.find(x => x.idMerma === idMerma);
    if (!m) return;
    _selMerma = m;
    const prods = OfflineManager.getProductosCache();
    const desc = prods.find(p => p.codigoBarra === m.codigoProducto || p.idProducto === m.codigoProducto)?.descripcion || m.codigoProducto;
    document.getElementById('resolverMermaInfo').textContent =
      `${desc} · ${fmt(m.cantidadOriginal, 1)} unidades · ${m.responsable || m.origen || ''}`;
    document.getElementById('resolverMermaReparar').value  = m.cantidadOriginal;
    document.getElementById('resolverMermaDesechar').value = 0;
    document.getElementById('resolverMermaObs').value      = '';
    document.getElementById('resolverMermaWarn').textContent = '';
    abrirSheet('sheetResolverMerma');
  }

  function balancearResolucion(modificado) {
    if (!_selMerma) return;
    const total = parseFloat(_selMerma.cantidadOriginal) || 0;
    const repInp = document.getElementById('resolverMermaReparar');
    const desInp = document.getElementById('resolverMermaDesechar');
    const warn   = document.getElementById('resolverMermaWarn');
    let rep = parseFloat(repInp.value) || 0;
    let des = parseFloat(desInp.value) || 0;
    // Auto-balance: al editar uno, ajustar el otro al complemento
    if (modificado === 'rep') des = Math.max(0, total - rep);
    if (modificado === 'des') rep = Math.max(0, total - des);
    repInp.value = rep;
    desInp.value = des;
    if (Math.abs((rep + des) - total) > 0.001) {
      warn.textContent = `⚠ Reparar + desechar debe sumar ${fmt(total, 1)} (actual: ${fmt(rep + des, 1)})`;
    } else { warn.textContent = ''; }
  }

  async function confirmarResolver() {
    if (!_selMerma) return;
    const rep = parseFloat(document.getElementById('resolverMermaReparar').value)  || 0;
    const des = parseFloat(document.getElementById('resolverMermaDesechar').value) || 0;
    const obs = document.getElementById('resolverMermaObs').value.trim();
    const total = parseFloat(_selMerma.cantidadOriginal) || 0;
    if (Math.abs((rep + des) - total) > 0.001) {
      toast(`Reparar + desechar debe sumar ${fmt(total, 1)}`, 'warn'); return;
    }
    if (rep === 0 && des === 0) { toast('Indica al menos una cantidad', 'warn'); return; }

    const btn = document.getElementById('btnConfirmarResolver');
    btn.disabled = true; btn.textContent = 'Aplicando...';
    try {
      const res = await API.resolverMerma({
        idMerma: _selMerma.idMerma,
        cantidadReparada:  rep,
        cantidadDesechada: des,
        observacionResolucion: obs,
        usuario: window.WH_CONFIG?.usuario || ''
      });
      if (res.ok) {
        const msgDesecho = des > 0 ? ` · ${fmt(des, 1)} a guía semanal` : '';
        toast(`✓ Resuelto${msgDesecho}`, 'ok', 4000);
        cerrarSheet('sheetResolverMerma');
        _selMerma = null;
        cargar();
      } else {
        toast('Error: ' + (res.error || 'desconocido'), 'danger', 5000);
      }
    } catch(e) { toast('Sin conexión', 'warn'); }
    finally { btn.disabled = false; btn.textContent = 'Aplicar resolución'; }
  }

  function verFoto(url) {
    if (!url) return;
    window.open(url, '_blank');
  }

  return { cargar, crear, nueva, setFiltro, onFotoSeleccionada,
           abrirResolver, balancearResolucion, confirmarResolver, verFoto };
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
  let _renderGen    = 0;     // cancela chunks de render anteriores

  // ── Estado de ajuste manual ───────────────────────────────
  let _ajusteTarget = null; // { codigoBarra, nombre }

  // ── Estado de auditoría diaria ────────────────────────────
  const _AUDIT_KEY  = 'wh_audit_dia';
  let _auditDia     = null;  // { fecha, skus:[...30], auditados:{sku:[cods]} }
  let _auditModo    = false; // modo filtro activo
  let _auditTarget  = null;  // barcode actualmente en auditoría

  // ── helpers ────────────────────────────────────
  function _s(id)          { return _stockMap[id] || { cantidadDisponible: 0, stockMinimo: 0, stockMaximo: 0 }; }
  function _buildMap(list) { (list || []).forEach(s => { _stockMap[s.codigoProducto || s.idProducto] = s; }); }

  // [Fix cutover Supabase] El stock que muestra esta vista DEBE venir del stock
  // EN VIVO (API.getStock → wh.stock_enriquecido vía lectura directa Supabase, o
  // getStockFlip por GAS), NO del cache wh_stock que puebla descargarOperacional
  // (que históricamente leía el Sheet congelado → ajustes directos no se veían y
  // ALSOL salía 0 / SIN STOCK). Trae stock fresco, reconstruye _stockMap y
  // re-renderiza. Si falla (offline/error), conserva lo que ya hay en el map
  // (cache) — nunca deja la vista peor que antes.
  async function _refrescarStockVivo() {
    try {
      const res = await API.getStock({ soloAlertas: false });
      const data = res && res.ok && Array.isArray(res.data) ? res.data : null;
      if (!data || !data.length) return false;
      _buildMap(data);  // overwrite por codigoProducto con valores vivos
      return true;
    } catch (_) { return false; }
  }

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

    // ── Pre-cálculo único de caches usadas por flags _dormido / _porVencer ──
    // Antes cada tap de chip filtro recalculaba O(n×m). Ahora O(m+n) total
    // al construir grupos: una sola pasada sobre detalles, guías y lotes.
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const lotes    = (OfflineManager.getLotesCache?.() || OfflineManager.getLotesVencimientoCache?.() || []);
    const gMap = {};
    guias.forEach(gg => { gMap[gg.idGuia] = gg; });
    const hace30   = Date.now() - 30 * 86400000;
    const en7dias  = Date.now() + 7 * 86400000;
    // codigoBarra → fecha último mov (epoch ms)
    const ultMovPorCod = {};
    detalles.forEach(d => {
      const cb = d.codigoProducto;
      if (!cb) return;
      const f = gMap[d.idGuia]?.fecha;
      if (!f) return;
      const t = new Date(f).getTime();
      if (!t) return;
      if (!ultMovPorCod[cb] || ultMovPorCod[cb] < t) ultMovPorCod[cb] = t;
    });
    // codigoBarra → tiene lote por vencer (<7d)
    const porVencerPorCod = {};
    lotes.forEach(l => {
      const cb = l.codigoProducto || l.codigoBarra;
      if (!cb || !l.fechaVencimiento) return;
      const t = new Date(l.fechaVencimiento).getTime();
      if (t > 0 && t < en7dias) porVencerPorCod[cb] = true;
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

      // Stock total + bajoMin + flags pre-calculadas (una sola pasada)
      let stockTotal = 0, bajoMin = false;
      let ultimoMov = 0, porVencer = false;
      children.forEach(c => {
        const s  = _s(c.codigoBarra);
        const st = s.cantidadDisponible || 0;
        const mn = parseFloat(s.stockMinimo || base.stockMinimo || 0);
        stockTotal += st;
        if (mn > 0 && st <= mn) bajoMin = true;
        const t = ultMovPorCod[c.codigoBarra] || 0;
        if (t > ultimoMov) ultimoMov = t;
        if (porVencerPorCod[c.codigoBarra]) porVencer = true;
      });
      // _dormido = sin movimientos en últimos 30 días
      const _dormido   = ultimoMov === 0 || ultimoMov < hace30;
      const _porVencer = porVencer;

      return { skuBase: g.skuBase, base, children, stockTotal, bajoMin, _dormido, _porVencer };
    }).sort((a, b) => String(a.base.descripcion || '').localeCompare(String(b.base.descripcion || ''), 'es'));
  }

  // ── Render lista de grupos ──────────────────────
  // Precomputa rotación y último movimiento UNA vez para todos los grupos (O(m+n), no O(m×n))
  function _render(grupos) {
    const el = document.getElementById('listProductos');
    if (!grupos.length) { el.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin productos</p>'; return; }

    // Invalidar cualquier render chunked anterior
    const gen = ++_renderGen;

    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const hace30   = Date.now() - 30 * 86400000;
    const gMap = {};
    guias.forEach(g => { gMap[g.idGuia] = g; });
    const cbCount = {}; // codigoBarra → nº movimientos últimos 30 días
    const cbFecha = {}; // codigoBarra → fecha último movimiento
    detalles.forEach(d => {
      const cb   = d.codigoProducto;
      if (!cb) return;
      const guia = gMap[d.idGuia];
      if (!guia) return;
      const fecha = guia.fecha;
      if (fecha && (!cbFecha[cb] || fecha > cbFecha[cb])) cbFecha[cb] = fecha;
      if (fecha && new Date(fecha) >= hace30) cbCount[cb] = (cbCount[cb] || 0) + 1;
    });

    const CHUNK = 40;
    // Primer bloque: render síncrono — el usuario ve contenido de inmediato
    el.innerHTML = grupos.slice(0, CHUNK).map(g => _cardGrupo(g, cbCount, cbFecha)).join('');

    if (grupos.length <= CHUNK) return;

    // Bloques restantes: se agregan progresivamente sin bloquear el hilo principal
    let i = CHUNK;
    const renderNext = () => {
      if (gen !== _renderGen) return; // render más reciente en curso, cancelar
      if (i >= grupos.length) return;
      el.insertAdjacentHTML('beforeend', grupos.slice(i, i + CHUNK).map(g => _cardGrupo(g, cbCount, cbFecha)).join(''));
      i += CHUNK;
      if (i < grupos.length) setTimeout(renderNext, 0);
    };
    setTimeout(renderNext, 0);
  }

  // Helper: escapa regex y resalta el término dentro de un texto
  function _highlightTerm(texto, query) {
    if (!query || !texto) return escHtml(texto || '');
    const tokens = query.toLowerCase().split(/\s+/).filter(t => t.length >= 2);
    if (!tokens.length) return escHtml(texto);
    const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const safe = escHtml(texto);
    let out = safe;
    tokens.forEach(t => {
      const re = new RegExp('(' + escapeRe(t) + ')', 'gi');
      out = out.replace(re, '<span class="prod-search-hl">$1</span>');
    });
    return out;
  }

  function _cardGrupo(g, cbCount, cbFecha) {
    const codigos = g.children.map(c => c.codigoBarra).filter(Boolean);
    // Rotación: suma de movimientos de todos los barcodes del grupo en los últimos 30 días
    const nRot = codigos.reduce((s, cb) => s + (cbCount[cb] || 0), 0);
    const rot  = nRot >= 10 ? { nivel: 'ALTA',  color: 'text-emerald-400', dot: 'bg-emerald-400' }
               : nRot >= 4  ? { nivel: 'MEDIA', color: 'text-amber-400',   dot: 'bg-amber-400'   }
               :               { nivel: 'BAJA',  color: 'text-slate-500',   dot: 'bg-slate-600'   };
    // Último movimiento: la fecha más reciente entre todos los barcodes del grupo
    const ulti = codigos.map(cb => cbFecha[cb]).filter(Boolean)
                        .sort((a, b) => b.localeCompare(a))[0] || null;
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

    // Estados visuales modernos — flags pre-calculadas en _agrupar (O(1))
    const dormidoCls  = g._dormido   ? 'is-dormido'    : '';
    const vencerCls   = g._porVencer ? 'is-por-vencer' : '';

    // Highlight del término buscado en la descripción
    const descRender  = _highlightTerm(g.base.descripcion || g.skuBase, _queryActual);

    return `
    <div class="prod-card ca ${accentCls} ${auditCls} ${dormidoCls} ${vencerCls}" id="grp-${sid}" style="position:relative">
      <!-- Cabecera -->
      <div class="flex items-start gap-2">
        <div class="flex-1 min-w-0">
          <div class="flex items-center gap-2 flex-wrap">
            <p class="font-bold text-sm leading-snug">${descRender}</p>
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
        <button id="memb-${sid}" onclick="event.stopPropagation();ProductosView.imprimirMembrete('${escAttr(g.skuBase)}','${escAttr(g.base.idProducto || '')}','${sid}')"
                class="btn btn-outline text-xs py-1 px-2 flex items-center gap-1"
                title="Imprimir membrete (canónico + equivalentes)">
          🖨 Membrete
        </button>
        <button onclick="event.stopPropagation();ProductosView.verCodigos('${escAttr(g.skuBase)}')"
                class="btn btn-outline text-xs py-1 px-2 flex items-center gap-1 prod-btn-codigos"
                title="Ver códigos de barra en pantalla para escanear directo">
          <span class="wh-bar-ico">▐│▌║▏</span> Códigos
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

    // ── Señales visuales + sonoras post-render ──────────────────────
    requestAnimationFrame(() => {
      if (exactGrupo) {
        const sid  = exactGrupo.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
        const card = document.getElementById('grp-' + sid);
        if (card) {
          // ✓ Match exacto = animación + sonido + scroll a la card
          card.classList.add('is-match-exact');
          setTimeout(() => card.classList.remove('is-match-exact'), 1300);
          const panel = document.getElementById('eqs-' + sid);
          const chev  = document.getElementById('chev-' + sid);
          if (panel && panel.classList.contains('hidden')) {
            panel.classList.remove('hidden');
            if (chev) chev.style.transform = 'rotate(180deg)';
          }
          card.scrollIntoView({ behavior: 'smooth', block: 'center' });
          try { SoundFX.beepDouble(); } catch(_){}
          vibrate([10, 20, 30]);
        }
        _filtrados.forEach(g => {
          if (g !== exactGrupo) {
            const s = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
            const el = document.getElementById('grp-' + s);
            if (el) el.classList.add('card-dim');
          }
        });
      } else if (_filtrados.length > 0) {
        // ↕ Match parcial (prefijo / contenido) = borde ámbar pulsante
        _filtrados.forEach(g => {
          const sid = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
          const el  = document.getElementById('grp-' + sid);
          if (el) el.classList.add('is-match-prefix');
        });
        try { SoundFX.beep && SoundFX.beep(); } catch(_){}
      } else {
        // ✗ Sin coincidencias
        try { SoundFX.warn && SoundFX.warn(); } catch(_){}
        vibrate([40, 20, 40]);
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

  // Filtro de tipo en historial moderno (Todos / Entradas / Salidas / Ajustes)
  let _histFiltroTipo = 'all';
  let _histMovsCache  = []; // movimientos cacheados para re-render con filtro
  let _histStockCache = 0;

  async function verHistorial(codigosStr, nombre) {
    // Acepta barcode único o pipe-separated para grupos multi-barcode
    const codigos = String(codigosStr).split('|').map(s => s.trim()).filter(Boolean);
    _histTarget = { codigos, codigo: codigos[0], nombre };
    _histFiltroTipo = 'all';
    document.querySelectorAll('.hist-tipo-chip').forEach(c => {
      c.classList.toggle('is-active', c.dataset.tipo === 'all');
    });

    document.getElementById('histNombre').textContent = nombre;
    document.getElementById('histCodigo').textContent = codigos.length > 1
      ? `${codigos[0]} · ${codigos.length} códigos`
      : codigos[0];
    document.getElementById('histList').innerHTML =
      '<div class="flex justify-center py-8"><div class="spinner"></div></div>';
    // [v2.13.54] Limpiar footer mientras carga
    const _ftPrev = document.getElementById('histFooter');
    if (_ftPrev) _ftPrev.innerHTML = '';
    abrirSheet('sheetHistorial');
    try { SoundFX.beep && SoundFX.beep(); } catch(_){}
    vibrate(12);

    // KPIs hero
    const stockTotal = codigos.reduce((sum, c) => sum + (_s(c).cantidadDisponible || 0), 0);
    const stockMin   = _s(codigos[0]).stockMinimo || 0;
    _histStockCache  = stockTotal;
    // Movs 30d (de la cache local — rápido)
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const gMap = {}; guias.forEach(g => { gMap[g.idGuia] = g; });
    const hace30 = Date.now() - 30 * 86400000;
    const codSet = new Set(codigos);
    const movs30 = detalles.filter(d => {
      if (!codSet.has(d.codigoProducto)) return false;
      const f = gMap[d.idGuia]?.fecha;
      return f && new Date(f).getTime() > hace30;
    }).length;
    const setKpi = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.textContent = v;
    };
    setKpi('histStockActual', fmt(stockTotal));
    setKpi('histStockMin',    fmt(stockMin));
    setKpi('histKpiMovs',     movs30);

    const local = _movimientosLocal(codigos);
    if (local.length) {
      _histMovsCache = local;
      _renderHistorial(local, true, stockTotal);
    }

    const res = await API.getHistorialStock(codigos.join(',')).catch(() => ({ ok: false }));
    if (res.ok) {
      const gasData = res.data || [];
      const gasIds  = new Set(gasData.map(m => m.idGuia).filter(Boolean));
      const extras  = local.filter(m => m.idGuia && !gasIds.has(m.idGuia));
      const merged  = [...gasData, ...extras].sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
      _histMovsCache = merged;
      if (merged.length) {
        _renderHistorial(merged, false, stockTotal);
      } else if (!local.length) {
        document.getElementById('histList').innerHTML = `
          <div class="hist-empty">
            <div class="hist-empty-icon">📭</div>
            <p class="text-sm">Sin movimientos registrados</p>
            <p class="text-xs mt-1 text-slate-600">El producto aún no tiene entradas ni salidas</p>
          </div>`;
      }
    }
  }

  // Filtro chip Todos/Entradas/Salidas/Ajustes — re-render desde cache local
  function histFiltrarTipo(tipo) {
    _histFiltroTipo = tipo || 'all';
    document.querySelectorAll('.hist-tipo-chip').forEach(c => {
      c.classList.toggle('is-active', c.dataset.tipo === _histFiltroTipo);
    });
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);
    _renderHistorial(_histMovsCache, false, _histStockCache);
  }

  // Auditar desde el modal historial — usa el primer codigoBarra del grupo
  function histAuditar() {
    if (!_histTarget?.codigos?.length) return;
    cerrarSheet('sheetHistorial');
    const cb = _histTarget.codigos[0];
    setTimeout(() => abrirAuditBarcode(cb, _histTarget.nombre, ''), 200);
  }

  function _renderHistorial(movs, esLocal, stockActual) {
    if (!movs.length && esLocal) return;
    const stock = stockActual ?? 0;

    // Balance corriente hacia atrás desde stock actual
    let bal = stock;
    const conBal = movs.map(m => {
      const row = { ...m, bal };
      bal = m.esIngreso ? bal - m.cantidad : bal + m.cantidad;
      return row;
    });

    const fBadge = document.getElementById('histFuenteBadge');
    if (fBadge) fBadge.textContent = esLocal ? '📦 caché local' : '☁ sincronizado';

    // Categoriza tipo de movimiento (para clases CSS y filtro)
    const _categoria = (m) => {
      const t = (m.tipo || '').toUpperCase();
      if (t.includes('INI'))     return 'ini';
      if (m.fuente === 'ajuste' || t.includes('AJUSTE')) return m.esIngreso ? 'ajuste-inc' : 'ajuste-dec';
      return m.esIngreso ? 'ingreso' : 'salida';
    };
    const _label = {
      ingreso:    'INGRESO',
      salida:     'SALIDA',
      'ajuste-inc': 'AJUSTE ▲',
      'ajuste-dec': 'AJUSTE ▼',
      ini:        'INICIAL'
    };
    const _signo = (cat) => (cat === 'ingreso' || cat === 'ajuste-inc' || cat === 'ini') ? '+' : '−';

    // Aplicar filtro chip
    const filtrados = conBal.filter(m => {
      if (_histFiltroTipo === 'all') return true;
      const cat = _categoria(m);
      if (_histFiltroTipo === 'ingreso') return cat === 'ingreso' || cat === 'ini';
      if (_histFiltroTipo === 'salida')  return cat === 'salida';
      if (_histFiltroTipo === 'ajuste')  return cat === 'ajuste-inc' || cat === 'ajuste-dec';
      return true;
    });

    if (!filtrados.length) {
      document.getElementById('histList').innerHTML = `
        <div class="hist-empty">
          <div class="hist-empty-icon">${_histFiltroTipo === 'all' ? '📭' : '🔎'}</div>
          <p class="text-sm">${_histFiltroTipo === 'all' ? 'Sin movimientos registrados' : 'Sin movimientos de este tipo'}</p>
        </div>`;
      return;
    }

    // [v2.13.54] Helper: card de lote — colores semánticos según días para vencer
    const _diasParaVencer = (fechaIso) => {
      if (!fechaIso) return null;
      const fv = new Date(fechaIso);
      if (isNaN(fv.getTime())) return null;
      return Math.ceil((fv.getTime() - Date.now()) / 86400000);
    };
    const _claseLote = (dias) => {
      if (dias === null) return 'lote-sin';
      if (dias < 0)      return 'lote-vencido';
      if (dias <= 7)     return 'lote-vencido';
      if (dias <= 30)    return 'lote-proxim';
      return 'lote-ok';
    };
    const _fmtFv = (fechaIso) => {
      if (!fechaIso) return '—';
      const d = new Date(fechaIso); if (isNaN(d.getTime())) return String(fechaIso).slice(0, 10);
      return ('0'+d.getDate()).slice(-2) + '/' + ('0'+(d.getMonth()+1)).slice(-2) + '/' + d.getFullYear();
    };
    const _renderLoteChip = (lote, cantidadConsumida) => {
      if (!lote) return '';
      const dias = _diasParaVencer(lote.fechaVencimiento);
      const cls  = _claseLote(dias);
      let badge;
      if (dias === null) badge = 'sin fecha';
      else if (dias < 0)  badge = `🚨 vencido ${Math.abs(dias)}d`;
      else if (dias <= 7) badge = `⚠ ${dias}d`;
      else if (dias <= 30) badge = `⏳ ${dias}d`;
      else                 badge = `✓ ${dias}d`;
      const cantTxt = (cantidadConsumida != null) ? ` <b style="color:#fbbf24">−${fmt(cantidadConsumida)}u</b>` : '';
      return `
        <div class="hist-lote-chip ${cls}">
          <span class="hist-lote-id">🏷 ${escHtml(lote.idLote)}</span>${cantTxt}
          <span class="hist-lote-fv">📅 ${_fmtFv(lote.fechaVencimiento)}</span>
          <span class="hist-lote-badge">${badge}</span>
        </div>`;
    };

    document.getElementById('histList').innerHTML = filtrados.map(m => {
      const cat = _categoria(m);
      const lbl = _label[cat] || cat;
      const sign = _signo(cat);
      const tipoChip = `<span class="hist-mov-tipo" style="background:rgba(15,23,42,.6);color:#94a3b8">${lbl}</span>`;
      // [v2.13.54] Render de lotes anidados según tipo de movimiento
      let lotesHtml = '';
      if (m.esIngreso && m.lote) {
        lotesHtml = `<div class="hist-lotes-wrap">${_renderLoteChip(m.lote, null)}</div>`;
      } else if (!m.esIngreso && Array.isArray(m.lotesConsumidos) && m.lotesConsumidos.length > 0) {
        lotesHtml = `<div class="hist-lotes-wrap is-salida">${m.lotesConsumidos.map(l => _renderLoteChip(l, l.cantidad)).join('')}</div>`;
      }
      return `
        <div class="hist-timeline-row is-${cat}">
          <div class="flex-1 min-w-0">
            <div class="flex items-baseline justify-between gap-2">
              <span class="hist-mov-amount is-${cat}">${sign}${fmt(m.cantidad)}</span>
              <span class="hist-mov-fecha">${fmtFecha(m.fecha)}</span>
            </div>
            <p class="hist-mov-meta">
              ${tipoChip}<span style="color:#94a3b8">${escHtml(m.tipo || '')}</span>${m.idGuia ? ` · <span class="font-mono text-[10px]">${escHtml(m.idGuia)}</span>` : ''}
            </p>
            ${m.origen ? `<p class="text-[10px] text-slate-600 truncate mt-0.5">${escHtml(m.origen)}</p>` : ''}
            ${lotesHtml}
            <div class="flex items-center gap-2 mt-1.5 flex-wrap">
              <span class="hist-mov-saldo">Saldo <strong>${fmt(m.bal)}</strong></span>
              ${m.usuario ? `<span class="text-[10px] text-slate-500">por ${escHtml(m.usuario)}</span>` : ''}
            </div>
          </div>
        </div>`;
    }).join('');

    // [v2.13.54] Resumen FOOTER de lotes activos del producto
    try {
      const codigos = (_histTarget && _histTarget.codigos) || [];
      if (codigos.length && API && API.getLotesFIFO) {
        const allLotes = [];
        Promise.all(codigos.map(cb => API.getLotesFIFO({ codigoProducto: cb }).then(r => (r && r.ok ? (r.data || []) : [])).catch(() => [])))
          .then(arrs => {
            arrs.forEach(a => allLotes.push(...a));
            const cont = document.getElementById('histFooter');
            if (!cont) return;
            if (!allLotes.length) {
              cont.innerHTML = `<div class="hist-footer-empty">Sin lotes activos · stock sin trazabilidad de vencimiento</div>`;
              return;
            }
            cont.innerHTML = `
              <div class="hist-footer-ttl">📦 Lotes activos · ${allLotes.length}</div>
              <div class="hist-footer-grid">
                ${allLotes.map(l => _renderLoteChip({ idLote: l.idLote, fechaVencimiento: l.fechaVencimiento }, l.cantidadActual)).join('')}
              </div>`;
            // [v2.13.54] Sonido suave cuando aparecen lotes activos
            try { SoundFX.done && SoundFX.done(); } catch(_){}
          });
      }
    } catch(_) {}
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

    const res = await PrintHub.imprimir('imprimirHistorialStock', {
      texto,
      codigoProducto: _histTarget.codigo
    }, 'Historial ' + (_histTarget.codigo || '')).catch(() => ({ ok: false }));
    if (res) toast(res.ok ? 'Impreso ✓' : 'No se pudo imprimir — revisa config GAS', res.ok ? 'ok' : 'warn');
  }

  // ── Aplicar query activa ─────────────────────────
  function _aplicarQuery() {
    if (_auditModo) {
      _filtrados = _grupos.filter(g => _auditDia?.skus.includes(g.skuBase));
      _render(_aplicarFiltroChip(_filtrados));
    } else if (_queryActual) {
      buscar(_queryActual);
    } else {
      _filtrados = [..._grupos];
      _render(_aplicarFiltroChip(_filtrados));
    }
    _renderMetrics();
  }

  // ─── Filtros chip (Todos / Stock bajo / Críticos / Vencer / Dormidos) ───
  let _filtroChip = 'all';

  // Helpers ahora son O(1) por item — usan flags pre-calculadas en _agrupar.
  // Esto evita recalcular cache de detalles/guías/lotes cada tap de chip filtro
  // (antes era O(n×m) = lento con 2k+ productos).
  function _grupoTieneVencimientoProximo(g) { return !!g._porVencer; }
  function _grupoEstaDormido(g)              { return !!g._dormido; }

  function _aplicarFiltroChip(grupos) {
    if (_filtroChip === 'all') return grupos;
    if (_filtroChip === 'bajo')    return grupos.filter(g => g.bajoMin && g.stockTotal > 0);
    if (_filtroChip === 'critico') return grupos.filter(g => g.stockTotal === 0);
    if (_filtroChip === 'vencer')  return grupos.filter(g => g._porVencer);
    if (_filtroChip === 'dormido') return grupos.filter(g => g._dormido);
    return grupos;
  }

  function toggleFiltro(filter) {
    _filtroChip = (_filtroChip === filter && filter !== 'all') ? 'all' : filter;
    document.querySelectorAll('#prodChipsRow .prod-chip').forEach(b => {
      b.classList.toggle('is-active', b.dataset.filter === _filtroChip);
    });
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);
    _aplicarQuery();
  }

  function _renderMetrics() {
    if (!_grupos.length) return;
    // Una sola pasada por _grupos en lugar de 5 .filter() — más rápido en 2k items
    let total = 0, bajo = 0, critico = 0, vencer = 0, dormido = 0;
    for (const g of _grupos) {
      total++;
      if (g.bajoMin && g.stockTotal > 0) bajo++;
      if (g.stockTotal === 0) critico++;
      if (g._porVencer) vencer++;
      if (g._dormido)   dormido++;
    }
    const set = (id, n) => {
      const el = document.getElementById(id);
      if (el) el.textContent = n;
    };
    set('chipNumAll',     total);
    set('chipNumBajo',    bajo);
    set('chipNumCritico', critico);
    set('chipNumVencer',  vencer);
    set('chipNumDormido', dormido);
  }

  // ─── Búsqueda por voz (Web Speech API) ───
  let _vozRecognition = null;
  let _vozActivo = false;
  function toggleVozBusqueda() {
    if (_vozActivo) { _detenerVoz(); return; }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { toast('Tu navegador no soporta búsqueda por voz', 'warn'); return; }
    _vozRecognition = new SR();
    _vozRecognition.lang = 'es-PE';
    _vozRecognition.continuous = false;
    _vozRecognition.interimResults = true;
    const btn = document.getElementById('prodMicBtn');
    if (btn) btn.classList.add('is-listening');
    _vozActivo = true;
    toast('🎤 Habla ahora · ej: "vinagre"', 'info', 2200);
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(15);
    _vozRecognition.onresult = (e) => {
      const last = e.results[e.results.length - 1];
      const text = last[0].transcript.trim();
      const inp = document.getElementById('inputBuscarProd');
      if (inp) inp.value = text;
      if (last.isFinal) {
        buscar(text);
        _detenerVoz();
      }
    };
    _vozRecognition.onerror = () => _detenerVoz();
    _vozRecognition.onend = () => _detenerVoz();
    try { _vozRecognition.start(); } catch(_){ _detenerVoz(); }
  }
  function _detenerVoz() {
    try { _vozRecognition && _vozRecognition.stop(); } catch(_){}
    _vozRecognition = null;
    _vozActivo = false;
    const btn = document.getElementById('prodMicBtn');
    if (btn) btn.classList.remove('is-listening');
  }

  // ═══════════════════════════════════════════════════════════════════════
  // PRECARGA + AUTO-REFRESH 60s SIN PARPADEO
  //
  // El sistema central (OfflineManager.iniciarRefreshOperacional) ya hace:
  //   1. Precarga al iniciar la app (PERSONAL, PRODUCTOS_MASTER, EQUIVALENCIAS,
  //      STOCK_ZONAS, GUIAS_CABECERA, GUIAS_DETALLE, PREINGRESOS, AJUSTES,
  //      AUDITORIAS, PROVEEDORES — todas las tablas maestras y operacionales).
  //   2. setInterval 60s para refresh background.
  //   3. dispatch wh:data-refresh cuando hay cambios → llama silentRefresh.
  //
  // Lo que faltaba: silentRefresh usaba _render(innerHTML) que causaba
  // parpadeo con 2k+ cards. Ahora usa _renderDiff que solo actualiza
  // las cards que cambiaron (basado en snapshot por skuBase).
  // ═══════════════════════════════════════════════════════════════════════
  let _cardSnapshots = new Map(); // skuBase → snapshot (string) para diff

  function _snapshotGrupo(g) {
    // Campos que pueden cambiar entre refreshes y deben triggerear update
    return [
      g.stockTotal,
      g.bajoMin ? 1 : 0,
      g._dormido ? 1 : 0,
      g._porVencer ? 1 : 0,
      g.children.length,
      g.base.descripcion || '',
      g.base.stockMinimo || 0,
      g.base.stockMaximo || 0
    ].join('·');
  }

  // Render con diff: si la card existe y su snapshot no cambió, no la toca.
  // Si cambió, reemplaza el HTML in-place + flash sutil.
  // Si es nuevo, append. Si desapareció, remove.
  function _renderDiff(grupos) {
    const list = document.getElementById('listProductos');
    if (!list) return;
    if (!grupos.length) {
      list.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin productos</p>';
      _cardSnapshots.clear();
      return;
    }

    // Pre-cómputo de cbCount/cbFecha (usa la misma lógica que _render)
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias    = OfflineManager.getGuiasCache();
    const hace30   = Date.now() - 30 * 86400000;
    const gMap = {}; guias.forEach(g => { gMap[g.idGuia] = g; });
    const cbCount = {}, cbFecha = {};
    detalles.forEach(d => {
      const cb = d.codigoProducto;
      if (!cb) return;
      const f = gMap[d.idGuia]?.fecha;
      if (!f) return;
      if (!cbFecha[cb] || f > cbFecha[cb]) cbFecha[cb] = f;
      if (new Date(f) >= hace30) cbCount[cb] = (cbCount[cb] || 0) + 1;
    });

    const newKeys = new Set();
    let huboCambios = 0;
    grupos.forEach((g, idx) => {
      const sid = g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_');
      newKeys.add(sid);
      const newSnap = _snapshotGrupo(g);
      const oldSnap = _cardSnapshots.get(sid);
      const cardEl  = document.getElementById('grp-' + sid);

      if (!cardEl) {
        // Card nueva → insertar
        const html = _cardGrupo(g, cbCount, cbFecha);
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        const newCard = tmp.firstElementChild;
        list.appendChild(newCard);
        _cardSnapshots.set(sid, newSnap);
        return;
      }
      if (oldSnap === newSnap) {
        // No cambió — no tocar el DOM (cero parpadeo)
        return;
      }
      // Cambió → reemplazar in-place con flash sutil
      const html = _cardGrupo(g, cbCount, cbFecha);
      const tmp = document.createElement('div');
      tmp.innerHTML = html;
      const newCard = tmp.firstElementChild;
      newCard.classList.add('is-updated');
      cardEl.replaceWith(newCard);
      _cardSnapshots.set(sid, newSnap);
      huboCambios++;
    });

    // Cards eliminados (productos que ya no aparecen)
    [...list.querySelectorAll('.prod-card')].forEach(c => {
      const id = c.id?.replace(/^grp-/, '');
      if (id && !newKeys.has(id)) {
        c.remove();
        _cardSnapshots.delete(id);
      }
    });
  }

  // Helper para filtrar por query sin recargar todo el flow de buscar()
  function _matchQuery(g, q) {
    const qL = String(q).toLowerCase();
    const tokens = qL.split(/\s+/).filter(Boolean);
    const hay = [
      g.base.descripcion || '', g.skuBase || '', g.base.idProducto || '',
      ...g.children.map(c => (c.descripcion || '') + ' ' + (c.codigoBarra || ''))
    ].join(' ').toLowerCase();
    return tokens.every(t => hay.includes(t));
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
    const pend = _pendientesCount();
    const total = _auditDia.skus.length;
    if (pend === 0) {
      btn.style.display = 'none'; // ocultar cuando todas completadas
    } else {
      btn.style.display = '';
      btn.innerHTML = `<svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M5.5 2A3.5 3.5 0 0 0 2 5.5v5A3.5 3.5 0 0 0 5.5 14h5a3.5 3.5 0 0 0 3.5-3.5V8a.5.5 0 0 1 1 0v2.5a4.5 4.5 0 0 1-4.5 4.5h-5A4.5 4.5 0 0 1 1 10.5v-5A4.5 4.5 0 0 1 5.5 1H8a.5.5 0 0 1 0 1H5.5z"/><path d="M16 3a3 3 0 1 1-6 0 3 3 0 0 1 6 0z"/></svg> ${pend}&nbsp;<span style="opacity:.6;font-weight:400">/ ${total}</span>`;
      btn.className = 'audit-badge audit-badge-pending' + (_auditModo ? ' active' : '');
    }
    const dot = document.getElementById('navDotAudit');
    if (dot) dot.style.display = pend > 0 ? '' : 'none';
    const sideDot = document.getElementById('sideNavDotAudit');
    if (sideDot) sideDot.style.display = pend > 0 ? '' : 'none';
  }

  function _searchFocusProd(focused) {
    const toolbar = document.getElementById('prodToolbar');
    if (!toolbar) return;
    if (focused) toolbar.classList.add('srch-focused');
    else setTimeout(() => toolbar.classList.remove('srch-focused'), 160);
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

  function abrirAuditBarcode(codigoBarra, nombre, skuBase, opts) {
    opts = opts || {};
    const cod = String(codigoBarra);
    const s   = _s(cod);
    _auditTarget = { codigoBarra: cod, nombre, skuBase, idAlerta: opts.idAlerta || null };
    document.getElementById('auditNombre').textContent   = nombre;
    document.getElementById('auditCodigo').textContent   = cod;
    document.getElementById('auditStockSis').textContent = fmt(s.cantidadDisponible || 0);
    // Pre-llenar conteo con teórico (si viene de una alerta) — el usuario puede editarlo
    document.getElementById('auditConteo').value =
      (opts.prefillFisico !== undefined && opts.prefillFisico !== null)
        ? String(opts.prefillFisico) : '';
    document.getElementById('auditObs').value =
      opts.idAlerta ? 'Auditoría desde alerta de cuadre' : '';
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
      // Si la auditoría vino de una alerta de cuadre, marcarla como revisada
      else if (target.idAlerta && typeof Dashboard !== 'undefined') {
        Dashboard.marcarAlertaRevisada(target.idAlerta);
      }
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
    // setTimeout(0): cede al browser DESPUÉS del paint, así el tab y el skeleton se ven antes del trabajo pesado
    // (requestAnimationFrame dispara ANTES del paint — incorrecto para este caso)
    await new Promise(r => setTimeout(r, 0));

    // Si ya tenemos grupos de la sesión actual, los re-usamos — _agrupar cuesta tiempo con muchos productos
    if (!_grupos.length) {
      const prods  = OfflineManager.getProductosCache();
      const equivs = OfflineManager.getEquivalenciasCache();
      _buildMap(OfflineManager.getStockCache());
      _grupos = _agrupar(prods, equivs);
    }
    _initAuditDia();
    _aplicarQuery();

    // [Fix cutover Supabase] Stock EN VIVO inmediato: pisa el cache (que puede
    // traer 0 del Sheet congelado) con el stock real de Supabase y re-renderiza.
    _refrescarStockVivo().then(ok => {
      if (!ok) return;
      _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
      _actualizarBadge();
      _aplicarQuery();
    }).catch(() => {});

    // Refrescar datos en background — actualiza _grupos y re-renderiza si hubo cambios
    OfflineManager.precargarOperacional().then(async () => {
      _buildMap(OfflineManager.getStockCache());
      // Stock vivo gana sobre el cache operacional (que puede estar congelado).
      await _refrescarStockVivo();
      _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
      _actualizarBadge();
      _aplicarQuery();
    }).catch(() => {});
  }

  // Re-render desde caché sin spinner ni API call (llamado por wh:data-refresh).
  // Usa render DIFF: solo las cards cuyo snapshot cambió se actualizan +
  // flash sutil; el resto del DOM no se toca → cero parpadeo.
  function silentRefresh() {
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();
    _buildMap(OfflineManager.getStockCache());
    _grupos = _agrupar(prods, equivs);
    // Indicador sutil: chip "todos" parpadea brevemente
    const chipAll = document.querySelector('#prodChipsRow [data-filter="all"]');
    if (chipAll) {
      chipAll.classList.add('is-refreshing');
      setTimeout(() => chipAll.classList.remove('is-refreshing'), 800);
    }
    // Render desde el estado actual (diff si la vista está visible).
    const _repintar = () => {
      const list = document.getElementById('listProductos');
      const yaRenderizado = list && list.querySelector('.prod-card');
      // [v2.13.68] Guard contra currentView undefined (cuando se llama desde
      // offline/precarga antes que la app inicialice la variable global).
      const _vistaProd = (typeof currentView !== 'undefined' && currentView === 'productos');
      if (yaRenderizado && _vistaProd) {
        const visibles = _aplicarFiltroChip(
          _queryActual ? _grupos.filter(g => _matchQuery(g, _queryActual)) : _grupos
        );
        _renderDiff(visibles);
        _renderMetrics();
      } else {
        _aplicarQuery();
      }
    };
    _repintar();
    // [Fix cutover Supabase] El cache wh_stock (descargarOperacional) puede estar
    // congelado (Sheet). El stock vivo de Supabase debe GANAR: tras pintar desde
    // cache, traemos el stock en vivo y, si difiere, reagrupamos y repintamos.
    // Esto es lo que evita el "F5 vuelve a 0": el live pisa el 0 del cache.
    _refrescarStockVivo().then(ok => {
      if (!ok) return;
      _grupos = _agrupar(OfflineManager.getProductosCache(), OfflineManager.getEquivalenciasCache());
      _repintar();
    }).catch(() => {});
  }

  // ── Cámara inline de búsqueda ────────────────────────────────
  let _prodCamTimer = null;

  function _setProdCamStatus(type, text) {
    const el = document.getElementById('prodCamStatus');
    if (!el) return;
    clearTimeout(_prodCamTimer);
    if (type === 'ready') {
      el.textContent  = '— apunta al código de barras —';
      el.style.color  = '#475569';
      el.style.background = '#0f172a';
      return;
    }
    const cfgs = {
      ok:        { col: '#34d399', bg: '#022c22', dur: 1200 },
      no_existe: { col: '#f87171', bg: '#2d0a0a', dur: 3500 }
    };
    const c = cfgs[type] || cfgs.ok;
    el.style.color  = c.col;
    el.style.background = c.bg;
    el.textContent  = (type === 'ok' ? '✓ ' : '⚠ ') + text;
    if (c.dur) _prodCamTimer = setTimeout(() => _setProdCamStatus('ready'), c.dur);
  }

  function _onProdCamResult(cod) {
    const cNorm  = normCb(cod);
    if (!cNorm) return;
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();

    // 1. Exacto en PRODUCTOS_MASTER
    let found = prods.find(p => String(p.codigoBarra || '').trim().toUpperCase() === cNorm);

    // 2. Exacto en EQUIVALENCIAS → resolver al producto maestro
    if (!found) {
      const equiv = equivs.find(e => String(e.codigoBarra || '').trim().toUpperCase() === cNorm);
      if (equiv) {
        const skuB = String(equiv.skuBase || '').trim().toUpperCase();
        found = prods.find(p =>
          String(p.idProducto  || '').trim().toUpperCase() === skuB ||
          String(p.skuBase     || '').trim().toUpperCase() === skuB ||
          String(p.codigoBarra || '').trim().toUpperCase() === skuB
        );
      }
    }

    if (!found) {
      _setProdCamStatus('no_existe', cNorm + ' · no existe en catálogo');
      SoundFX.warn(); vibrate([50, 25, 50]);
      _vozAnunciar && _vozAnunciar('No encontrado', { rate: 1.1 });
      return;
    }

    // ✓ Match perfecto al escanear: efecto fuerte + sonido + voz + flash card
    SoundFX.beepDouble(); vibrate([15, 30, 60]);
    _setProdCamStatus('ok', found.descripcion || cNorm);
    _vozAnunciar && _vozAnunciar(found.descripcion || 'Producto encontrado', { rate: 1.05 });
    const inp = document.getElementById('inputBuscarProd');
    const searchCode = found.codigoBarra || cNorm;
    if (inp) { inp.value = searchCode; }
    buscar(searchCode);
    setTimeout(() => cerrarProdCamara(), 900);
  }

  function abrirProdCamara() {
    const strip = document.getElementById('prodCamStrip');
    if (!strip) return;
    // Telón: colapsar header (chips + toolbar) con animación
    const header = document.getElementById('prodHeaderCollapsible');
    if (header) header.classList.add('is-collapsed');
    strip.classList.remove('is-closing');
    strip.style.display = 'block';
    _setProdCamStatus('ready');
    const btn = document.getElementById('prodCamBtn');
    if (btn) btn.style.background = 'rgba(14,165,233,.25)';
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(10);
    Scanner.start('prodCamVideo', _onProdCamResult, err => {
      toast('Error cámara: ' + err, 'danger');
      cerrarProdCamara();
    }, { continuous: true, cooldown: 1200 });
  }

  function cerrarProdCamara() {
    clearTimeout(_prodCamTimer);
    try { Scanner.stop(); } catch(_){}
    const strip = document.getElementById('prodCamStrip');
    if (strip) {
      strip.classList.add('is-closing');
      setTimeout(() => {
        strip.style.display = 'none';
        strip.classList.remove('is-closing');
      }, 320);
    }
    // Restaurar header
    const header = document.getElementById('prodHeaderCollapsible');
    if (header) header.classList.remove('is-collapsed');
    const btn = document.getElementById('prodCamBtn');
    if (btn) btn.style.background = '';
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(8);
  }

  function toggleProdCamara() {
    const strip = document.getElementById('prodCamStrip');
    if (strip && strip.style.display !== 'none') cerrarProdCamara();
    else abrirProdCamara();
  }

  // ═══════════════════════════════════════════════════════════════════════
  // SHEET DETALLE PRODUCTO — Modal moderno con tabs (Stock/Movs/Lotes/Códigos)
  // Tap en card del producto → abre este sheet con info completa.
  // Acciones: Despachar (envía a despacho rápido) · Auditar · Historial.
  // NUNCA precio ni edición — WH solo mueve mercadería en cantidades.
  // ═══════════════════════════════════════════════════════════════════════
  let _detSkuActivo = null;
  let _detTabActivo = 'stock';

  function abrirSheetDetalleProducto(skuBase) {
    if (!skuBase) return;
    const grupo = _grupos.find(g => String(g.skuBase) === String(skuBase));
    if (!grupo) return;
    _detSkuActivo = skuBase;
    _detTabActivo = 'stock';
    // Hero
    const titEl = document.getElementById('prodDetTitulo');
    const skuEl = document.getElementById('prodDetSku');
    if (titEl) titEl.textContent = grupo.base.descripcion || skuBase;
    if (skuEl) skuEl.textContent = `${skuBase} · ${grupo.base.unidad || ''}`;
    // KPIs
    const codigos = grupo.children.map(c => c.codigoBarra).filter(Boolean);
    const detalles = OfflineManager.getGuiaDetalleCache();
    const guias = OfflineManager.getGuiasCache();
    const gMap = {};
    guias.forEach(g => { gMap[g.idGuia] = g; });
    const hace30 = Date.now() - 30 * 86400000;
    const setCodigos = new Set(codigos);
    const movs30 = detalles.filter(d => {
      if (!setCodigos.has(d.codigoProducto)) return false;
      const f = gMap[d.idGuia]?.fecha;
      return f && new Date(f).getTime() > hace30;
    }).length;
    const setKpi = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };
    setKpi('prodDetKpiStock', fmt(grupo.stockTotal));
    setKpi('prodDetKpiRot',   movs30);
    setKpi('prodDetKpiCodes', codigos.length);
    // Tabs
    document.querySelectorAll('.prod-detail-tab').forEach(t => {
      t.classList.toggle('is-active', t.dataset.tab === 'stock');
    });
    detSetTab('stock');
    abrirSheet('sheetProdDetalle');
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(10);
  }

  function detSetTab(tab) {
    if (!_detSkuActivo) return;
    _detTabActivo = tab;
    document.querySelectorAll('.prod-detail-tab').forEach(t => {
      t.classList.toggle('is-active', t.dataset.tab === tab);
    });
    const cont = document.getElementById('prodDetTabContent');
    if (!cont) return;
    cont.classList.remove('prod-detail-tab-content');
    void cont.offsetWidth;
    cont.classList.add('prod-detail-tab-content');
    const grupo = _grupos.find(g => String(g.skuBase) === String(_detSkuActivo));
    if (!grupo) { cont.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Producto no encontrado</p>'; return; }

    if (tab === 'stock') {
      // Breakdown stock por cada codigoBarra del grupo (canónico + equivalentes)
      cont.innerHTML = grupo.children.map(c => {
        const s = _s(c.codigoBarra);
        const cant = parseFloat(s.cantidadDisponible) || 0;
        const tag = c.origen === 'equiv' ? 'is-equiv' : 'is-canonico';
        const tagTxt = c.origen === 'equiv' ? 'EQUIV' : 'CANÓN';
        return `
          <div class="prod-stock-row ${tag}">
            <span class="prod-stock-tag ${tag}">${tagTxt}</span>
            <div class="flex-1 min-w-0">
              <p class="text-xs font-mono truncate">${escHtml(c.codigoBarra)}</p>
              <p class="text-[10px] text-slate-500 truncate">${escHtml(c.descripcion || '')}</p>
            </div>
            <span class="font-black text-sm ${cant > 0 ? 'text-emerald-400' : 'text-slate-500'}">${fmt(cant)}</span>
          </div>`;
      }).join('') +
      `<p class="text-[10px] text-slate-500 mt-3 text-center">Total grupo: <span class="text-emerald-400 font-bold">${fmt(grupo.stockTotal)}</span> unidades · stock se descuenta por código real al despachar</p>`;
    } else if (tab === 'movs') {
      // Timeline últimos 30 movimientos
      const codSet = new Set(grupo.children.map(c => c.codigoBarra));
      const movs = detalles
        .filter(d => codSet.has(d.codigoProducto))
        .map(d => {
          const g = gMap[d.idGuia] || {};
          return {
            fecha: g.fecha,
            tipo: g.tipo || '',
            cant: parseFloat(d.cantidad) || 0,
            cb: d.codigoProducto,
            idGuia: d.idGuia
          };
        })
        .filter(m => m.fecha)
        .sort((a,b) => String(b.fecha).localeCompare(String(a.fecha)))
        .slice(0, 30);
      if (!movs.length) {
        cont.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin movimientos registrados</p>';
        return;
      }
      cont.innerHTML = movs.map(m => {
        const esEntrada = String(m.tipo).indexOf('ENTRADA') === 0 || String(m.tipo).indexOf('INGRESO') >= 0;
        const cls = esEntrada ? 'is-entrada' : 'is-salida';
        const tipoIcon = esEntrada ? '↓' : '↑';
        return `
          <div class="prod-timeline-item ${cls}">
            <div class="flex-1 min-w-0">
              <p class="prod-timeline-tipo ${cls}">${tipoIcon} ${escHtml(m.tipo)} · ${fmt(m.cant)}</p>
              <p class="prod-timeline-fecha">${fmtFecha(m.fecha)} · <span class="font-mono">${escHtml(m.cb)}</span></p>
              <p class="prod-timeline-meta">Guía ${escHtml(m.idGuia)}</p>
            </div>
          </div>`;
      }).join('');
    } else if (tab === 'lotes') {
      const lotes = (OfflineManager.getLotesCache?.() || OfflineManager.getLotesVencimientoCache?.() || [])
        .filter(l => grupo.children.some(c => String(c.codigoBarra) === String(l.codigoProducto || l.codigoBarra)));
      if (!lotes.length) {
        cont.innerHTML = '<p class="text-slate-500 text-center py-8 text-sm">Sin lotes con vencimiento registrados</p>';
        return;
      }
      const hoy = Date.now();
      const en7d = hoy + 7 * 86400000;
      cont.innerHTML = lotes.sort((a,b) =>
        new Date(a.fechaVencimiento || 0) - new Date(b.fechaVencimiento || 0)
      ).map(l => {
        const t = new Date(l.fechaVencimiento || 0).getTime();
        const critico = t > 0 && t < en7d;
        const vencido = t > 0 && t < hoy;
        const color = vencido ? '#f87171' : critico ? '#fbbf24' : '#94a3b8';
        return `
          <div class="card-sm flex items-center gap-3" style="border-color:${color}3a">
            <div style="font-size:18px;flex-shrink:0">${vencido ? '⛔' : critico ? '⚠' : '📅'}</div>
            <div class="flex-1 min-w-0">
              <p class="text-xs font-bold" style="color:${color}">${l.fechaVencimiento || 'Sin fecha'}</p>
              <p class="text-[10px] text-slate-500 font-mono truncate">${escHtml(l.codigoProducto || l.codigoBarra || '')}</p>
            </div>
            <span class="font-bold text-sm" style="color:${color}">${fmt(l.cantidadDisponible || l.cantidad || 0)}</span>
          </div>`;
      }).join('');
    } else if (tab === 'codes') {
      cont.innerHTML = grupo.children.map(c => {
        const tag = c.origen === 'equiv' ? 'is-equiv' : 'is-canonico';
        const tagTxt = c.origen === 'equiv' ? 'EQUIV' : 'CANÓNICO';
        return `
          <div class="prod-stock-row ${tag}">
            <span class="prod-stock-tag ${tag}">${tagTxt}</span>
            <div class="flex-1 min-w-0">
              <p class="text-xs font-mono truncate">${escHtml(c.codigoBarra)}</p>
              <p class="text-[10px] text-slate-500 truncate">${escHtml(c.descripcion || '')}</p>
            </div>
          </div>`;
      }).join('') +
      `<p class="text-[10px] text-slate-500 mt-3 text-center">${grupo.children.length} código${grupo.children.length!==1?'s':''} aceptado${grupo.children.length!==1?'s':''} al escanear · canónico es lo nominal, equivalentes son aliases válidos</p>`;
    }
  }

  function detDespacharActual() {
    if (!_detSkuActivo) return;
    const grupo = _grupos.find(g => String(g.skuBase) === String(_detSkuActivo));
    if (!grupo) return;
    cerrarSheet('sheetProdDetalle');
    try { SoundFX.beepDouble && SoundFX.beepDouble(); } catch(_){}
    vibrate(15);
    toast(`📦 Yendo a despacho rápido para ${grupo.base.descripcion || _detSkuActivo}`, 'info', 2500);
    setTimeout(() => App.nav('despacho'), 250);
  }

  function detAuditarActual() {
    if (!_detSkuActivo) return;
    const grupo = _grupos.find(g => String(g.skuBase) === String(_detSkuActivo));
    if (!grupo || !grupo.children.length) return;
    cerrarSheet('sheetProdDetalle');
    const c0 = grupo.children[0];
    setTimeout(() => abrirAuditBarcode(c0.codigoBarra, c0.descripcion || grupo.base.descripcion, _detSkuActivo), 200);
  }

  function detHistorialActual() {
    if (!_detSkuActivo) return;
    const grupo = _grupos.find(g => String(g.skuBase) === String(_detSkuActivo));
    if (!grupo || !grupo.children.length) return;
    cerrarSheet('sheetProdDetalle');
    const codigos = grupo.children.map(c => c.codigoBarra).join('|');
    setTimeout(() => verHistorial(codigos, grupo.base.descripcion || _detSkuActivo), 200);
  }

  // ═══════════════════════════════════════════════════════════════════════
  // SWIPE GESTURES + LONG-PRESS en cards de productos (mobile-friendly).
  // Swipe izquierda → Auditar  ·  Swipe derecha → Historial
  // Long-press (500ms) → menú contextual
  // ═══════════════════════════════════════════════════════════════════════
  let _swipeStartX = null;
  let _swipeStartY = null;
  let _swipeCardEl = null;
  let _lpTimer = null;

  function _attachGestures() {
    const list = document.getElementById('listProductos');
    if (!list || list._gestAttached) return;
    list._gestAttached = true;

    // Click en card body → abre sheet detalle.
    // Si el click viene de un botón/input interno, NO abrir (deja seguir su acción).
    list.addEventListener('click', (e) => {
      if (e.target.closest('button, input, select, textarea, a')) return;
      const card = e.target.closest('.prod-card');
      if (!card || !card.id?.startsWith('grp-')) return;
      const cardId = card.id.replace(/^grp-/, '');
      const grupo = _grupos.find(g => g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_') === cardId);
      if (!grupo) return;
      // Solo abrir si no estamos en medio de un swipe (evita falso positivo)
      if (_swipeCardEl) return;
      abrirSheetDetalleProducto(grupo.skuBase);
    });

    list.addEventListener('touchstart', (e) => {
      const card = e.target.closest?.('.prod-card');
      if (!card) return;
      _swipeStartX = e.touches[0].clientX;
      _swipeStartY = e.touches[0].clientY;
      _swipeCardEl = card;
      card.classList.add('prod-card-swipeable');
      // Long-press timer
      clearTimeout(_lpTimer);
      _lpTimer = setTimeout(() => {
        if (_swipeCardEl !== card) return;
        _abrirLpMenu(card, _swipeStartX, _swipeStartY);
        _lpTimer = null;
      }, 500);
    }, { passive: true });

    list.addEventListener('touchmove', (e) => {
      if (!_swipeCardEl || _swipeStartX === null) return;
      const dx = e.touches[0].clientX - _swipeStartX;
      const dy = Math.abs(e.touches[0].clientY - _swipeStartY);
      if (dy > 12) { _resetSwipe(); return; }
      if (Math.abs(dx) > 14) {
        clearTimeout(_lpTimer);
        _lpTimer = null;
        _swipeCardEl.classList.toggle('swipe-left',  dx < -30);
        _swipeCardEl.classList.toggle('swipe-right', dx >  30);
        _swipeCardEl.style.transform = `translateX(${Math.max(-80, Math.min(80, dx * 0.4))}px)`;
      }
    }, { passive: true });

    list.addEventListener('touchend', (e) => {
      clearTimeout(_lpTimer); _lpTimer = null;
      if (!_swipeCardEl || _swipeStartX === null) { _resetSwipe(); return; }
      const dx = (e.changedTouches[0].clientX - _swipeStartX);
      const card = _swipeCardEl;
      const sku  = card.id.replace(/^grp-/, '').replace(/_/g, '');
      // Buscar el grupo real (ojo: el id tiene caracteres reemplazados)
      const grupo = _grupos.find(g => g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_') === card.id.replace(/^grp-/, ''));
      if (!grupo) { _resetSwipe(); return; }
      if (dx < -60) {
        // Swipe izquierda → Auditar
        try { SoundFX.beepDouble && SoundFX.beepDouble(); } catch(_){}
        vibrate(15);
        const c0 = grupo.children[0];
        if (c0) abrirAuditBarcode(c0.codigoBarra, c0.descripcion || grupo.base.descripcion, grupo.skuBase);
      } else if (dx > 60) {
        // Swipe derecha → Historial
        try { SoundFX.beepDouble && SoundFX.beepDouble(); } catch(_){}
        vibrate(15);
        const codigos = grupo.children.map(c => c.codigoBarra).join('|');
        verHistorial(codigos, grupo.base.descripcion || grupo.skuBase);
      }
      _resetSwipe();
    });

    list.addEventListener('touchcancel', () => { clearTimeout(_lpTimer); _resetSwipe(); });
  }

  function _resetSwipe() {
    if (_swipeCardEl) {
      _swipeCardEl.style.transform = '';
      _swipeCardEl.classList.remove('swipe-left', 'swipe-right');
    }
    _swipeStartX = null;
    _swipeStartY = null;
    _swipeCardEl = null;
  }

  function _abrirLpMenu(card, clickX, clickY) {
    if (!card) return;
    const sku = card.id.replace(/^grp-/, '');
    const grupo = _grupos.find(g => g.skuBase.replace(/[^a-zA-Z0-9_-]/g, '_') === sku);
    if (!grupo) return;
    try { SoundFX.click && SoundFX.click(); } catch(_){}
    vibrate(20);
    // Ripple visual
    const rip = document.createElement('div');
    rip.className = 'prod-card-lp-ripple';
    const rect = card.getBoundingClientRect();
    rip.style.left = (clickX - rect.left) + 'px';
    rip.style.top  = (clickY - rect.top) + 'px';
    card.style.position = 'relative';
    card.appendChild(rip);
    setTimeout(() => rip.remove(), 700);
    // Menú
    const menu = document.createElement('div');
    menu.className = 'prod-lp-menu';
    menu.innerHTML = `
      <div class="prod-lp-menu-item" data-act="detalle"><span class="prod-lp-menu-icon">📦</span>Ver detalle</div>
      <div class="prod-lp-menu-item" data-act="despachar"><span class="prod-lp-menu-icon">🚚</span>Despachar este producto</div>
      <div class="prod-lp-menu-item" data-act="auditar"><span class="prod-lp-menu-icon">🕵</span>Auditar</div>
      <div class="prod-lp-menu-item" data-act="historial"><span class="prod-lp-menu-icon">📊</span>Historial</div>
    `;
    document.body.appendChild(menu);
    const mw = menu.offsetWidth, mh = menu.offsetHeight;
    let mx = clickX - mw/2, my = clickY - mh - 10;
    if (mx < 8) mx = 8;
    if (mx + mw > window.innerWidth - 8) mx = window.innerWidth - mw - 8;
    if (my < 60) my = clickY + 20;
    menu.style.left = mx + 'px';
    menu.style.top  = my + 'px';
    const close = () => { menu.remove(); document.removeEventListener('click', close); };
    menu.addEventListener('click', (e) => {
      const it = e.target.closest?.('.prod-lp-menu-item');
      if (!it) return;
      const act = it.dataset.act;
      close();
      const c0 = grupo.children[0];
      if (act === 'detalle')   abrirSheetDetalleProducto(grupo.skuBase);
      if (act === 'despachar') {
        try { SoundFX.beepDouble && SoundFX.beepDouble(); } catch(_){}
        toast(`📦 Yendo a despacho rápido para ${grupo.base.descripcion}`, 'info', 2500);
        setTimeout(() => App.nav('despacho'), 250);
      }
      if (act === 'auditar' && c0) abrirAuditBarcode(c0.codigoBarra, c0.descripcion || grupo.base.descripcion, grupo.skuBase);
      if (act === 'historial') verHistorial(grupo.children.map(c => c.codigoBarra).join('|'), grupo.base.descripcion);
    });
    setTimeout(() => document.addEventListener('click', close, { once: true }), 50);
  }

  // Llamar attachGestures cada vez que se renderice la lista
  const _origRender = _render;
  _render = function(grupos) {
    _origRender(grupos);
    setTimeout(_attachGestures, 50);
  };

  // ── Imprimir membrete con lockout 2s anti-doble-click ──
  const _membLock = new Set();
  function imprimirMembrete(skuBase, idProducto, sid) {
    const key = String(skuBase || idProducto);
    if (_membLock.has(key)) {
      if (typeof toast === 'function') toast('Espera 2s para reimprimir', 'info', 1200);
      return;
    }

    // [v2.13.120] Si el módulo MembreteSystem está cargado, abrir menú
    // ME|WH para que el operario elija tipo. Multi-código auto-detectado
    // desde el grupo.children. Sistema nuevo con polling + fire-and-forget.
    const grupoNew = _grupos.find(gr => gr.skuBase === skuBase);
    if (window.MembreteSystem && MembreteSystem.abrirMenuProductoCard) {
      const childrenCodes = grupoNew
        ? grupoNew.children.map(c => String(c.codigoBarra || '')).filter(Boolean)
        : [];
      MembreteSystem.abrirMenuProductoCard({
        idProducto:  idProducto || skuBase,
        codigoBarra: childrenCodes[0] || idProducto || skuBase,
        descripcion: (grupoNew && grupoNew.base && grupoNew.base.descripcion) || skuBase || '',
        skuBase:     skuBase || '',
        codigos:     childrenCodes
      });
      return;
    }

    // ── FALLBACK al sistema viejo (PrintHub) si MembreteSystem no cargó ──
    _membLock.add(key);

    // Feedback inmediato visual: deshabilitar botón 2s
    const btn = document.getElementById('memb-' + sid);
    if (btn) {
      btn.disabled = true;
      btn.style.opacity = '.55';
      btn.style.pointerEvents = 'none';
    }
    setTimeout(() => {
      _membLock.delete(key);
      if (btn) {
        btn.disabled = false;
        btn.style.opacity = '';
        btn.style.pointerEvents = '';
      }
    }, 2000);

    // Mandar todos los codigoBarra al GAS (canónico + equivalentes) si están disponibles
    const grupo = _grupos.find(gr => gr.skuBase === skuBase);
    const barcodes = grupo
      ? grupo.children.map(c => String(c.codigoBarra || '')).filter(Boolean)
      : [];

    // PrintHub: admin/master ve el selector de impresora; usuario normal directo.
    PrintHub.imprimir('imprimirMembrete', {
      idProducto: idProducto || skuBase,
      barcodes: barcodes.length ? JSON.stringify(barcodes) : ''
    }, 'Membrete ' + (skuBase || idProducto || '')).then(res => {
      if (res === null) return; // admin canceló el modal
      if (res && res.ok) {
        if (typeof toast === 'function') toast('✓ Membrete enviado a impresora', 'ok', 1800);
        if (typeof SoundFX !== 'undefined' && SoundFX.savedTick) SoundFX.savedTick();
        if (navigator.vibrate) navigator.vibrate(10);
      } else {
        const msg = (res && res.error) || 'Error desconocido';
        if (typeof toast === 'function') toast('Impresora: ' + msg, 'warn', 3500);
        if (typeof SoundFX !== 'undefined' && SoundFX.warn) SoundFX.warn();
      }
    }).catch(e => {
      if (typeof toast === 'function') toast('Sin conexión a impresora', 'warn', 3500);
    });
  }

  // ── MEMBRETE DIGITAL — overlay con códigos de barra escaneables ──
  // El operador apunta el escáner físico a la PANTALLA para "leer" el
  // producto sin que el membrete físico esté pegado en el andamio.
  // Solo enruta el escaneo al despacho si hay un despacho/pickup activo
  // (Parte B — listener global de DespachoView).
  function verCodigos(skuBase) {
    const g = (_grupos || []).find(gr => String(gr.skuBase) === String(skuBase));
    if (!g) { toast('Producto no encontrado', 'warn'); return; }
    const codigos = (g.children || [])
      .map(c => String(c.codigoBarra || '').trim())
      .filter(Boolean);
    if (!codigos.length) { toast('Este producto no tiene códigos de barra', 'warn'); return; }

    const cont = document.getElementById('codigosOverlayBody');
    const titEl = document.getElementById('codigosOverlayTitulo');
    if (titEl) titEl.textContent = g.base.descripcion || skuBase;
    if (cont) {
      // Un bloque por código: SVG del barcode + el número debajo. El
      // primero es el canónico, el resto equivalentes.
      cont.innerHTML = codigos.map((cb, idx) => `
        <div class="cod-barcode-card">
          <div class="cod-barcode-tag">${idx === 0 ? '★ CANÓNICO' : 'EQUIVALENTE ' + idx}</div>
          <svg class="cod-barcode-svg" id="codbc-${idx}"></svg>
          <div class="cod-barcode-num">${escHtml(cb)}</div>
        </div>`).join('');
    }
    const ov = document.getElementById('codigosOverlay');
    if (ov) { ov.style.display = 'flex'; requestAnimationFrame(() => ov.classList.add('is-open')); }
    try { SoundFX.beep && SoundFX.beep(); } catch(_){}
    try { vibrate && vibrate(12); } catch(_){}

    // Renderizar los barcodes con JsBarcode (tras el reflow del overlay).
    // Code128 codifica cualquier texto (sirve para EAN reales y códigos
    // internos WH... por igual). Fondo blanco + quiet zone amplio para
    // que el escáner lo lea bien desde la pantalla.
    setTimeout(() => {
      codigos.forEach((cb, idx) => {
        const svg = document.getElementById('codbc-' + idx);
        if (!svg || typeof JsBarcode === 'undefined') return;
        try {
          JsBarcode(svg, cb, {
            format: 'CODE128', width: 2.4, height: 90,
            displayValue: false, margin: 14, background: '#ffffff', lineColor: '#000000'
          });
        } catch(e) { svg.outerHTML = '<div class="cod-barcode-err">⚠ No se pudo generar el código</div>'; }
      });
    }, 60);
  }
  function cerrarCodigos() {
    const ov = document.getElementById('codigosOverlay');
    if (ov) { ov.classList.remove('is-open'); setTimeout(() => { ov.style.display = 'none'; }, 250); }
  }

  return { cargar, silentRefresh, buscar, buscarClear, _searchFocusProd, toggleGrupo, toggleAuditoriaDia,
           abrirAuditBarcode, confirmarAuditoria,
           abrirAjuste, abrirAjusteDesdeHistorial, previewAjuste, confirmarAjuste,
           verHistorial, imprimirHistorial, imprimirMembrete, verCodigos, cerrarCodigos,
           histFiltrarTipo, histAuditar,
           abrirProdCamara, cerrarProdCamara, toggleProdCamara,
           toggleFiltro, toggleVozBusqueda,
           abrirSheetDetalleProducto, detSetTab,
           detDespacharActual, detAuditarActual, detHistorialActual };
})();


// ════════════════════════════════════════════════
// MEMBRETE VIEW — cola de impresión de membretes
// ════════════════════════════════════════════════
const MembreteView = (() => {
  let _cola = []; // [{ prod, allEan }]

  function _esActivoEquiv(v) {
    if (v === true || v === 1) return true;
    const s = String(v || '').trim().toUpperCase();
    return s === '1' || s === 'TRUE' || s === 'YES' || s === 'SI' || s === 'S';
  }

  function _buildEan(prod) {
    const equivs = OfflineManager.getEquivalenciasCache();
    const skuKey = String(prod.skuBase || prod.idProducto || '').trim().toUpperCase();
    const altCodes = equivs
      .filter(e => String(e.skuBase || '').trim().toUpperCase() === skuKey
                && _esActivoEquiv(e.activo)
                && e.codigoBarra)
      .map(e => String(e.codigoBarra).trim());
    const allEan = [];
    if (prod.codigoBarra) allEan.push(String(prod.codigoBarra).trim());
    altCodes.forEach(c => { if (c && !allEan.includes(c)) allEan.push(c); });
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

    const ql     = val.toLowerCase();
    const qU     = val.toUpperCase();
    const prods  = OfflineManager.getProductosCache();
    const equivs = OfflineManager.getEquivalenciasCache();

    // Solo productos base de almacén: factor=1 y activos
    const baseProds = prods.filter(p => {
      if (p.estado === '0' || p.estado === 0) return false;
      return parseFloat(p.factorConversion || 1) === 1;
    });

    // Agrupar equivs por skuBase (normalizado)
    const equivsByKey = {};
    equivs.forEach(e => {
      if (!_esActivoEquiv(e.activo)) return;
      const k = String(e.skuBase || '').trim().toUpperCase();
      if (!k || !e.codigoBarra) return;
      if (!equivsByKey[k]) equivsByKey[k] = [];
      equivsByKey[k].push(String(e.codigoBarra).trim());
    });

    // Dedup por skuBase: 1 entrada por producto base
    const seenSku = new Set();
    const matches = [];
    for (const p of baseProds) {
      const skuKey = String(p.skuBase || p.idProducto || '').trim().toUpperCase();
      if (seenSku.has(skuKey)) continue;
      const eqCbs = equivsByKey[skuKey] || [];
      const haystack = [
        String(p.descripcion || '').toLowerCase(),
        String(p.idProducto  || '').toLowerCase(),
        String(p.skuBase     || '').toLowerCase(),
        String(p.codigoBarra || '').toLowerCase(),
        ...eqCbs.map(c => c.toLowerCase())
      ].join(' ');
      if (haystack.includes(ql)) {
        seenSku.add(skuKey);
        const allCbs = [];
        if (p.codigoBarra) allCbs.push(String(p.codigoBarra).trim());
        eqCbs.forEach(c => { if (c && !allCbs.includes(c)) allCbs.push(c); });
        matches.push({ prod: p, allCbs });
      }
      if (matches.length >= 12) break;
    }

    if (!matches.length) {
      sEl.innerHTML = `<div style="padding:10px 12px;font-size:12px;color:#64748b;">Sin resultados</div>`;
    } else {
      sEl.innerHTML = matches.map(({ prod: p, allCbs }) => {
        const yaEnCola = _cola.some(i => i.prod.idProducto === p.idProducto);
        const sku   = String(p.skuBase || p.idProducto || '');
        const nCb   = allCbs.length;
        const preview = allCbs.length
          ? (allCbs.slice(0, 2).join(' · ') + (allCbs.length > 2 ? ` · +${allCbs.length - 2} más` : ''))
          : 'sin códigos';
        return `<button onclick="MembreteView.seleccionar('${escAttr(p.idProducto)}')"
                style="display:flex;align-items:flex-start;gap:8px;width:100%;text-align:left;
                       padding:10px 12px;background:none;border:none;cursor:pointer;
                       border-bottom:1px solid #1e293b;transition:background .1s;"
                onmouseenter="this.style.background='#1e293b'"
                onmouseleave="this.style.background='none'">
          <div style="flex:1;min-width:0">
            <div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">
              <span style="flex:1;font-size:13px;font-weight:700;color:${yaEnCola ? '#475569' : '#e2e8f0'};
                           white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
                ${escHtml(p.descripcion || sku)}
              </span>
              <span style="flex-shrink:0;font-size:10px;font-weight:700;padding:1px 6px;
                           border-radius:99px;background:${nCb > 1 ? 'rgba(59,130,246,.18)' : 'rgba(71,85,105,.2)'};
                           color:${nCb > 1 ? '#60a5fa' : '#64748b'}">
                ${nCb} cb${nCb !== 1 ? 's' : ''}
              </span>
              ${yaEnCola ? '<span style="font-size:10px;color:#22c55e;flex-shrink:0">✓</span>' : ''}
            </div>
            <p style="font-size:10.5px;color:#64748b;font-family:monospace;
                      white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
              ${escHtml(sku)} · ${escHtml(preview)}
            </p>
          </div>
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
      const cN     = String(code).trim().toUpperCase();

      // Solo productos base
      const baseProds = prods.filter(p => {
        if (p.estado === '0' || p.estado === 0) return false;
        return parseFloat(p.factorConversion || 1) === 1;
      });

      // 1. Match en maestro
      let exact = baseProds.find(p =>
        String(p.idProducto  || '').trim().toUpperCase() === cN ||
        String(p.codigoBarra || '').trim().toUpperCase() === cN ||
        String(p.skuBase     || '').trim().toUpperCase() === cN
      );

      // 2. Match en EQUIVALENCIAS → resolver al producto base
      if (!exact) {
        const equiv = equivs.find(e =>
          String(e.codigoBarra || '').trim().toUpperCase() === cN && _esActivoEquiv(e.activo)
        );
        if (equiv) {
          const skuB = String(equiv.skuBase || '').trim().toUpperCase();
          exact = baseProds.find(p =>
            String(p.skuBase     || '').trim().toUpperCase() === skuB ||
            String(p.idProducto  || '').trim().toUpperCase() === skuB ||
            String(p.codigoBarra || '').trim().toUpperCase() === skuB
          );
        }
      }
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
    // [fix ALTO-4] validar formato de deployment de Apps Script antes de persistir.
    // Sin esto, una URL errada/malformada se fijaba para SIEMPRE (wh_gas_url está en PRESERVAR)
    // → la app quedaba pegada a un /exec equivocado (la preocupación #1 materializada).
    if (!/^https:\/\/script\.google\.com\/macros\/s\/AKfyc[\w-]{20,}\/exec$/.test(gasUrl)) {
      toast('URL inválida — debe ser https://script.google.com/macros/s/AKfyc.../exec', 'warn');
      return;
    }
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
// LOGS VIEW — historial de movimientos de stock
// Solo accesible para MASTER / ADMINISTRADOR
// ════════════════════════════════════════════════
// ════════════════════════════════════════════════
// FOTO PICKER — selector reutilizable de foto desde preingreso.
// Uso: FotoPicker.abrir(idPreingreso, fotos, onSelect)
//   fotos: array de URLs (col 'fotos' del preingreso)
//   onSelect: (fileId | null) => void   (null = sin foto)
// ════════════════════════════════════════════════
const FotoPicker = (() => {
  let _onSelect = null;

  function _extractFileId(url) {
    const m = String(url || '').match(/[?&]id=([^&]+)/);
    return m ? m[1] : '';
  }

  async function abrir(idPreingreso, fotos, onSelect) {
    _onSelect = onSelect || (() => {});
    const arr = Array.isArray(fotos) ? fotos.filter(Boolean)
      : String(fotos || '').split(',').map(s => s.trim()).filter(Boolean);
    if (!arr.length) {
      // Sin fotos → callback null directo
      _onSelect(null);
      _onSelect = null;
      return;
    }
    if (arr.length === 1) {
      // Una sola foto: confirmar rápido sin abrir modal
      const fid = _extractFileId(arr[0]);
      if (await _whConfirm('El preingreso tiene 1 foto adjunta.\n\n¿Usarla en la guía?', { titulo: 'Foto preingreso' })) {
        _onSelect(fid || null);
      } else {
        _onSelect(null);
      }
      _onSelect = null;
      return;
    }
    // Múltiples fotos: abrir selector
    const sub = document.getElementById('selectorFotoPISub');
    if (sub) sub.textContent = `${arr.length} fotos disponibles — toca una para usarla en la guía`;
    const grid = document.getElementById('selectorFotoPIGrid');
    if (grid) {
      grid.innerHTML = arr.map((url, i) => {
        const fid = _extractFileId(url);
        return `
        <button onclick="FotoPicker.elegir('${escAttr(fid)}')"
                style="aspect-ratio:1;border:2px solid #334155;border-radius:10px;
                       background:#0f172a;cursor:pointer;overflow:hidden;
                       padding:0;position:relative"
                ontouchstart="this.style.borderColor='#0ea5e9'"
                ontouchend="this.style.borderColor='#334155'">
          <img src="${escAttr(url)}" alt="Foto ${i+1}"
               style="width:100%;height:100%;object-fit:cover"
               onerror="this.style.display='none';this.nextElementSibling.style.display='flex'"/>
          <div style="display:none;width:100%;height:100%;align-items:center;justify-content:center;color:#475569;font-size:.7em">${i+1}</div>
          <span style="position:absolute;bottom:6px;right:6px;background:rgba(0,0,0,.7);
                       color:#fff;font-size:.65em;padding:2px 7px;border-radius:99px">${i+1}</span>
        </button>`;
      }).join('');
    }
    abrirSheet('sheetSelectorFotoPI');
  }

  function elegir(fileId) {
    cerrarSheet('sheetSelectorFotoPI');
    if (_onSelect) _onSelect(fileId || null);
    _onSelect = null;
  }

  function confirmarSinFoto() {
    cerrarSheet('sheetSelectorFotoPI');
    if (_onSelect) _onSelect(null);
    _onSelect = null;
  }

  return { abrir, elegir, confirmarSinFoto };
})();

// ════════════════════════════════════════════════
// WELCOME — pantalla post-login con stats + pre-carga inteligente
// ════════════════════════════════════════════════
const Welcome = (() => {
  const FRASES = [
    'Un almacén ordenado es trabajo bien hecho',
    'Cada producto bien acomodado es un cliente feliz',
    'El detalle de hoy es la calidad de mañana',
    'Pequeñas tareas hechas bien construyen grandes resultados',
    'Trabaja con calma, los errores cuestan más que el tiempo',
    'Quien controla el stock, controla el negocio',
    'El almacén es la columna del negocio',
    'Hoy es un día para hacer el inventario sin dudas',
    'Un envase bien hecho es un cliente que vuelve',
    'La precisión hoy ahorra problemas mañana',
    'Cada guía cerrada bien es una operación segura',
    'Tu trabajo silencioso sostiene toda la empresa',
    'Sin el almacén, no hay venta',
    'Contar bien es más rápido que contar dos veces',
    'El orden es la base del control',
    'Un buen almacenero ve lo que otros no ven',
    'Hoy decides cómo termina el día',
    'La constancia vence al talento',
    'Cada día es una nueva oportunidad de hacerlo perfecto',
    'El almacén bien cuidado vende solo'
  ];

  let _saludoOpened = false;
  let _typewriterTimer = null;

  function _saludo() {
    const h = new Date().getHours();
    if (h >= 5  && h < 12) return 'Buenos días';
    if (h >= 12 && h < 19) return 'Buenas tardes';
    return 'Buenas noches';
  }

  async function mostrar(sesion) {
    if (_saludoOpened) return;
    _saludoOpened = true;

    const overlay = document.getElementById('welcomeOverlay');
    if (!overlay) return;

    // Llenar datos básicos disponibles
    const nombre = sesion.nombre || '';
    const apellido = sesion.apellido || '';
    const iniciales = ((nombre[0] || '?') + (apellido[0] || '')).toUpperCase();
    const av = document.getElementById('welAvatar');
    if (av) {
      av.textContent = iniciales;
      if (sesion.color) av.style.background = sesion.color;
    }
    document.getElementById('welSaludo').textContent = `¡${_saludo()}, ${nombre}! ⚡`;
    document.getElementById('welBrand').textContent  = `InversionMos · ${(sesion.rol || '').toUpperCase()}`;
    document.getElementById('welHoraInicio').textContent = '⏰ ' + (sesion.horaInicio || '—').substring(0, 5);

    // Frase aleatoria con typewriter
    const frase = FRASES[Math.floor(Math.random() * FRASES.length)];
    _typewriter(frase);

    // Mostrar overlay + sonido
    overlay.style.display = 'flex';
    try { SoundFX.welcome(); } catch(e) {}

    // Pre-carga + datos welcome en paralelo
    _ejecutarPrecarga(sesion);
  }

  function _typewriter(texto) {
    const el = document.getElementById('welQuote');
    if (!el) return;
    if (_typewriterTimer) clearInterval(_typewriterTimer);
    el.textContent = '';
    let i = 0;
    _typewriterTimer = setInterval(() => {
      el.textContent = texto.substring(0, ++i);
      if (i >= texto.length) { clearInterval(_typewriterTimer); _typewriterTimer = null; }
    }, 32);
  }

  function _setProgreso(pct, label) {
    const bar = document.getElementById('welProgBar');
    const lbl = document.getElementById('welProgLabel');
    if (bar) bar.style.width = Math.min(100, pct) + '%';
    if (lbl) lbl.textContent = label || '';
  }

  async function _ejecutarPrecarga(sesion) {
    const jobs = [
      { name: 'maestros',    fn: () => OfflineManager.precargar(true) },
      { name: 'operacional', fn: () => OfflineManager.precargarOperacional(true) },
      { name: 'welcome',     fn: () => API.getWelcomeData({ idPersonal: sesion.idPersonal }) }
    ];
    let done = 0;
    let welcomeData = null;

    _setProgreso(5, 'Cargando datos...');

    for (const job of jobs) {
      try {
        const res = await job.fn();
        if (job.name === 'welcome' && res?.ok) welcomeData = res.data;
      } catch(e) { /* silencioso */ }
      done++;
      _setProgreso((done / jobs.length) * 100, `${done}/${jobs.length} · ${job.name}`);
    }

    _setProgreso(100, 'Listo');

    if (welcomeData) _renderEstadisticas(welcomeData);

    // Auto-cerrar tras 5s si el usuario no entra
    setTimeout(() => { if (_saludoOpened) cerrar(); }, 5000);
  }

  function _renderEstadisticas(d) {
    document.getElementById('welRacha').textContent    = '🔥 ' + (d.racha || 0);
    document.getElementById('welSesiones').textContent = '📅 ' + (d.totalSesiones || 0);

    // Pendientes
    const pend = [];
    if (d.pendientes) {
      if (d.pendientes.mermasVencidas > 0)  pend.push(`⚠ ${d.pendientes.mermasVencidas} merma(s) vencida(s) (>3 días)`);
      if (d.pendientes.envasesUrgentes > 0) pend.push(`📦 ${d.pendientes.envasesUrgentes} envase(s) urgente(s)`);
      if (d.pendientes.auditoriasHoy > 0)   pend.push(`📅 ${d.pendientes.auditoriasHoy} auditoría(s) para hoy`);
    }
    if (d.sinResolverAyer > 0) pend.push(`🔻 ${d.sinResolverAyer} merma(s) sin resolver`);
    if (pend.length) {
      document.getElementById('welPendList').innerHTML = pend.map(p => `<div>${escHtml(p)}</div>`).join('');
      document.getElementById('welPendientes').style.display = '';
      try { SoundFX.ping(); } catch(e) {}
    }

    // Ayer
    if (d.ayer) {
      const partes = [];
      if (d.ayer.guiasCerradas > 0)    partes.push(`${d.ayer.guiasCerradas} guías cerradas`);
      if (d.ayer.unidadesEnvasadas > 0) partes.push(`${fmt(d.ayer.unidadesEnvasadas, 0)} unid. envasadas`);
      if (d.ayer.mermasResueltas > 0)  partes.push(`${d.ayer.mermasResueltas} mermas resueltas`);
      if (partes.length) {
        document.getElementById('welAyerTexto').textContent = 'Ayer: ' + partes.join(' · ');
        document.getElementById('welAyer').style.display = '';
      }
    }
  }

  function cerrar() {
    if (!_saludoOpened) return;
    _saludoOpened = false;
    if (_typewriterTimer) clearInterval(_typewriterTimer);
    const overlay = document.getElementById('welcomeOverlay');
    if (overlay) overlay.style.display = 'none';
  }

  // ── Almacén cerrado por horario ─────────────────────
  function mostrarAlmacenCerrado(info) {
    // [v2.13.122] Usar modal moderno SeguridadSystem si está cargado
    // (incluye botones "solicitar extensión" + "notifícame cuando abra")
    try {
      if (window.SeguridadSystem && window.SeguridadSystem.abrirModalFueraHorario) {
        // [v2.13.125 FIX] Soportar números decimales (7.5 → "07:30")
        var _numAHHMM = function(n) {
          var h = Math.floor(n);
          var m = Math.round((n - h) * 60);
          if (m === 60) { h++; m = 0; }
          return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
        };
        var _segAp = info && info.apertura != null ? String(info.apertura) : '07:00';
        var _segCi = info && info.cierre   != null ? String(info.cierre)   : '19:00';
        if (typeof info.apertura === 'number') _segAp = _numAHHMM(info.apertura);
        if (typeof info.cierre   === 'number') _segCi = _numAHHMM(info.cierre);
        window.SeguridadSystem.abrirModalFueraHorario(info && info.motivo || 'fuera', _segAp, _segCi);
        try { SoundFX.buzzer(); } catch(_) {}
        return;
      }
    } catch(_) {}
    // Fallback: overlay viejo
    const overlay = document.getElementById('almacenCerradoOverlay');
    if (!overlay) return;
    const dia = info.dia || 1;
    const horario = (dia === 7)
      ? 'Domingo: 07:00 - 16:00'
      : 'Lun-Sáb: 07:00 - 19:00';
    document.getElementById('acHorario').textContent = horario;
    // Calcular falta hasta apertura
    const ahora = new Date();
    const apertura = new Date(ahora);
    if (info.motivo === 'despues_cierre') apertura.setDate(apertura.getDate() + 1);
    apertura.setHours(info.apertura || 7, 0, 0, 0);
    const diff = apertura - ahora;
    const horas = Math.floor(diff / 3600000);
    const mins  = Math.floor((diff % 3600000) / 60000);
    document.getElementById('acFaltan').textContent =
      `Faltan ${horas}h ${mins}m para abrir`;
    overlay.style.display = 'flex';
    try { SoundFX.buzzer(); } catch(e) {}
  }

  function cerrarAlmacenCerrado() {
    const overlay = document.getElementById('almacenCerradoOverlay');
    if (overlay) overlay.style.display = 'none';
  }

  function intentarAccesoAdmin() {
    cerrarAlmacenCerrado();
    // Volver a mostrar pantalla de login para que admin/master ingrese su PIN
    const loginScr = document.getElementById('loginScreen');
    if (loginScr) loginScr.style.display = 'flex';
  }

  // ── Extensión de horario in-situ (admin/master con clave 8 dig) ────
  // [v2.13.163] DELEGADO al módulo compartido ExtensorHorario. Antes esta
  // sección manejaba su propio modal y flow con 1h FIJA. Ahora se unifica
  // con ME — mismo modal, mismas opciones (20m/1h/2h), mismo backend.
  // Auditoría tier 2 vía verificarClaveAdmin('EXTENDER_HORARIO_DISPOSITIVO').
  const _WH_DESBLOQUEO_KEY_LEGACY = 'wh_desbloqueo_hasta';   // legacy (1h fija)

  function _desbloqueoVigente() {
    // Source of truth: módulo compartido (ext_horario_hasta).
    // Compat con clave legacy si operador hizo extensión antes de v2.13.163.
    try {
      if (window.ExtensorHorario) {
        const ms = ExtensorHorario.vigente();
        if (ms > 0) return Date.now() + ms;
      }
      const legacy = parseInt(localStorage.getItem(_WH_DESBLOQUEO_KEY_LEGACY) || '0', 10);
      return legacy && Date.now() < legacy ? legacy : 0;
    } catch(_) { return 0; }
  }

  function abrirDesbloqueoTemporal() {
    if (!window.ExtensorHorario) {
      // Fallback fail-soft: módulo no cargó (CDN caído)
      toast('Módulo de extensión no disponible. Recarga la app.', 'error', 5000);
      return;
    }
    const mosUrl = window.WH_CONFIG?.mosGasUrl || '';
    const devId  = (typeof window._getDeviceIdWH === 'function') ? window._getDeviceIdWH() : '';
    if (!mosUrl || !devId) {
      toast('Configuración incompleta — recarga la app.', 'error', 4000);
      return;
    }
    ExtensorHorario.abrir({
      mosGasUrl: mosUrl,
      app:       'warehouseMos',
      deviceId:  devId,
      onSuccess: function(d) {
        // Extensión activada → cerrar overlay de almacén cerrado + volver al login
        cerrarAlmacenCerrado();
        toast(`🔓 Extensión activa · +${d.minutos} min · ${d.aprobadoPor || 'admin'}`, 'ok', 5500);
        const loginScr = document.getElementById('loginScreen');
        if (loginScr) loginScr.style.display = 'flex';
      }
    });
  }

  // [v2.13.164] Stubs cerrarDesbloqueoTemporal/confirmarDesbloqueoTemporal
  // removidos junto con whDesbloqueoModal en index.html. Ningún call site
  // queda activo en el HTML actual ni en el resto del JS.

  // ── Aviso 5 min antes del cierre + verificación periódica ──
  let _avisoMostrado = false;
  let _cierreInterval = null;

  function iniciarMonitorHorario(rol) {
    const rolUp = String(rol || '').toUpperCase();
    if (rolUp === 'MASTER' || rolUp === 'ADMIN' || rolUp === 'ADMINISTRADOR') return;  // sin restricción
    if (_cierreInterval) clearInterval(_cierreInterval);
    _cierreInterval = setInterval(() => _checkCierreInminente(), 60 * 1000); // cada minuto
    _checkCierreInminente();
  }

  function _checkCierreInminente() {
    const ahora = new Date();
    const dia   = ahora.getDay() === 0 ? 7 : ahora.getDay(); // 1=lun, 7=dom
    const cierreH = (dia === 7) ? 16 : 19;
    const aperturaH = 7;
    const horaActual = ahora.getHours() + ahora.getMinutes() / 60;

    // Si hay desbloqueo temporal vigente → ignorar bloqueo
    const desHasta = _desbloqueoVigente();
    if (desHasta) {
      // Actualizar el contador visible en overlay (si está abierto)
      const minRest = Math.ceil((desHasta - Date.now()) / 60000);
      const elInfo = document.getElementById('acDesbloqueoInfo');
      const elMin  = document.getElementById('acDesbloqueoMin');
      if (elInfo && elMin) { elMin.textContent = String(minRest); elInfo.style.display = 'block'; }
      return;
    }

    // Fuera de horario → forzar logout y pantalla de cierre
    if (horaActual >= cierreH || horaActual < aperturaH) {
      Welcome.mostrarAlmacenCerrado({ apertura: aperturaH, cierre: cierreH, dia,
                                        motivo: horaActual >= cierreH ? 'despues_cierre' : 'antes_apertura' });
      try { SoundFX.closeAlarm(); } catch(e) {}
      // Cerrar sesión local
      try { localStorage.removeItem('wh_sesion'); } catch(e) {}
      if (_cierreInterval) { clearInterval(_cierreInterval); _cierreInterval = null; }
      return;
    }

    // 5 min antes del cierre → aviso una sola vez
    const minutosAlCierre = (cierreH - horaActual) * 60;
    if (minutosAlCierre > 0 && minutosAlCierre <= 5 && !_avisoMostrado) {
      _avisoMostrado = true;
      try { SoundFX.bell(); } catch(e) {}
      toast(`🔔 El almacén cierra en ${Math.ceil(minutosAlCierre)} minutos · guarda tu trabajo`, 'warn', 8000);
    }
  }

  return { mostrar, cerrar, mostrarAlmacenCerrado, cerrarAlmacenCerrado, intentarAccesoAdmin, iniciarMonitorHorario,
           abrirDesbloqueoTemporal };
})();

const LogsView = (() => {
  let _movimientos = [];
  let _cargando    = false;

  const TIPO_LABEL = {
    'CIERRE_GUIA':       { txt: 'Cierre guía',       color: '#3b82f6', bg: 'rgba(59,130,246,.15)' },
    'ANULACION_DETALLE': { txt: 'Anulación',         color: '#f87171', bg: 'rgba(248,113,113,.15)' },
    'EDICION_CANTIDAD':  { txt: 'Edición cantidad',  color: '#fbbf24', bg: 'rgba(251,191,36,.15)' },
    'AUTO_SUMA_DETALLE': { txt: 'Auto-suma',         color: '#a78bfa', bg: 'rgba(167,139,250,.15)' },
    'REABRIR_REVERSO':   { txt: 'Reabrir guía',      color: '#fb923c', bg: 'rgba(251,146,60,.15)' },
    'AJUSTE_MANUAL':     { txt: 'Ajuste manual',     color: '#34d399', bg: 'rgba(52,211,153,.15)' },
    'AUDITORIA':         { txt: 'Auditoría',         color: '#60a5fa', bg: 'rgba(96,165,250,.15)' },
    'ENVASADO_BASE':     { txt: 'Envasado base',     color: '#c084fc', bg: 'rgba(192,132,252,.15)' },
    'ENVASADO_DERIVADO': { txt: 'Envasado derivado', color: '#c084fc', bg: 'rgba(192,132,252,.15)' },
    'APROBACION_PN':     { txt: 'Aprobación PN',     color: '#34d399', bg: 'rgba(52,211,153,.15)' },
    'INDEFINIDO':        { txt: 'Indefinido',        color: '#94a3b8', bg: 'rgba(148,163,184,.15)' }
  };

  async function cargar() {
    if (_cargando) return;
    _cargando = true;
    const list = document.getElementById('logsList');
    if (list) list.innerHTML = '<p class="text-xs text-slate-500 text-center py-6">Cargando movimientos...</p>';
    try {
      const res = await API.getStockMovimientos({ limit: 500 });
      _movimientos = res.ok ? (res.data || []) : [];
      filtrar();
    } catch(e) {
      if (list) list.innerHTML = '<p class="text-xs text-red-400 text-center py-6">Sin conexión</p>';
    } finally {
      _cargando = false;
    }
  }

  function filtrar() {
    const q     = (document.getElementById('logsFiltroProducto')?.value || '').toLowerCase().trim();
    const tipo  = document.getElementById('logsFiltroTipo')?.value || '';
    let lista = [..._movimientos];
    if (q) {
      // Buscar por código + descripción (resolvemos descripción desde caché)
      const prods = OfflineManager.getProductosCache();
      const descMap = {};
      prods.forEach(p => {
        if (p.codigoBarra) descMap[String(p.codigoBarra).trim()] = String(p.descripcion || '').toLowerCase();
        if (p.idProducto)  descMap[String(p.idProducto)] = String(p.descripcion || '').toLowerCase();
      });
      lista = lista.filter(m => {
        const cb = String(m.codigoProducto || '').toLowerCase();
        const desc = descMap[String(m.codigoProducto || '').trim()] || '';
        return cb.includes(q) || desc.includes(q);
      });
    }
    if (tipo) lista = lista.filter(m => m.tipoOperacion === tipo);
    _render(lista);
  }

  function _render(lista) {
    const list = document.getElementById('logsList');
    const tot  = document.getElementById('logsTotal');
    if (!list) return;
    if (tot) tot.textContent = lista.length + ' / ' + _movimientos.length + ' movimientos';
    if (!lista.length) {
      list.innerHTML = '<p class="text-xs text-slate-500 text-center py-6">Sin movimientos</p>';
      return;
    }
    // Resolver descripciones
    const prods = OfflineManager.getProductosCache();
    const descMap = {};
    prods.forEach(p => {
      const name = p.descripcion || '';
      if (!name) return;
      if (p.codigoBarra) descMap[String(p.codigoBarra).trim()] = name;
      if (p.idProducto)  descMap[String(p.idProducto)] = name;
    });

    list.innerHTML = lista.map(m => {
      const cfg = TIPO_LABEL[m.tipoOperacion] || TIPO_LABEL.INDEFINIDO;
      const desc = descMap[String(m.codigoProducto || '').trim()] || m.codigoProducto;
      const delta = parseFloat(m.delta) || 0;
      const fechaStr = m.fecha ? new Date(m.fecha).toLocaleString('es-PE') : '';
      return `
      <div class="card-sm" style="padding:9px 11px">
        <div class="flex items-start justify-between gap-2 mb-1">
          <div class="flex-1 min-w-0">
            <p class="font-semibold text-sm text-slate-100 truncate">${escHtml(desc)}</p>
            <p class="text-[10px] text-slate-500 font-mono">${escHtml(m.codigoProducto)}</p>
          </div>
          <span class="font-black text-sm flex-shrink-0"
                style="color:${delta > 0 ? '#34d399' : '#f87171'}">${delta > 0 ? '+' : ''}${fmt(delta, 2)}</span>
        </div>
        <div class="flex items-center justify-between gap-2 text-[11px]">
          <span style="display:inline-block;padding:1px 8px;border-radius:99px;
                       background:${cfg.bg};color:${cfg.color};font-weight:700;font-size:10px">
            ${cfg.txt}
          </span>
          <span class="text-slate-500">${escHtml(fechaStr)}</span>
        </div>
        <div class="text-[10px] text-slate-500 mt-1 flex gap-3">
          <span>antes: <b class="text-slate-300">${fmt(m.stockAntes, 1)}</b></span>
          <span>después: <b class="text-slate-300">${fmt(m.stockDespues, 1)}</b></span>
          ${m.usuario ? `<span class="ml-auto">${escHtml(m.usuario)}</span>` : ''}
        </div>
        ${m.origen ? `<p class="text-[10px] text-slate-600 font-mono mt-0.5">origen: ${escHtml(m.origen)}</p>` : ''}
      </div>`;
    }).join('');
  }

  return { cargar, filtrar };
})();

// ════════════════════════════════════════════════
// DIAGNOSTICO VIEW — auto-tests guiados de bugs de cuadre
// ════════════════════════════════════════════════
const DiagnosticoView = (() => {
  const TESTS = [
    {
      id: 'TEST_1',
      titulo: 'Scan rápido mismo producto (10 escaneos)',
      objetivo: 'Detecta race condition al escanear el mismo producto muy rápido',
      requiereProducto: true,
      pasos: [
        'Abre la sección Guías y crea una nueva guía SALIDA_ZONA',
        'Abre la cámara y escanea el producto seleccionado 10 VECES SEGUIDAS, lo más rápido que puedas (uno cada ~0.3s)',
        'Espera 5 segundos para que GAS termine',
        'Cierra la cámara (NO cierres la guía aún)'
      ],
      esperado: '1 fila en el detalle con qty=10'
    },
    {
      id: 'TEST_2',
      titulo: 'Scan + edición concurrente (out-of-order)',
      objetivo: 'Detecta si actualizarCantidad llega fuera de orden',
      requiereProducto: true,
      pasos: [
        'Crea una guía SALIDA_ZONA nueva',
        'Escanea el producto seleccionado UNA vez (qty=1)',
        'Toca el ítem, edita la cantidad a 15 (escribe rápido en el input)',
        'Antes de cerrar la edición, ESCANEA otra vez el mismo producto',
        'Confirma la edición y espera 5 segundos'
      ],
      esperado: 'qty final >=15 (idealmente 16)'
    },
    {
      id: 'TEST_3',
      titulo: 'Cerrar guía inmediatamente tras escaneos',
      objetivo: 'Detecta si el cierre llega antes que las escrituras pendientes',
      requiereProducto: true,
      pasos: [
        'Crea una guía SALIDA_ZONA nueva',
        'Escanea el producto seleccionado 5 veces rápido',
        'INMEDIATAMENTE cierra la cámara',
        'INMEDIATAMENTE pulsa CERRAR GUÍA',
        'Espera 30 segundos'
      ],
      esperado: 'Guía CERRADA, qty=5, stock bajó exactamente 5 unidades'
    },
    {
      id: 'TEST_4',
      titulo: 'Despacho rápido doble-click',
      objetivo: 'Detecta si el doble-click crea guías duplicadas',
      requiereProducto: false,
      pasos: [
        'Ve al módulo Despacho rápido',
        'Escanea 3 productos diferentes',
        'Pulsa "Generar guía" Y INMEDIATAMENTE pulsa otra vez (doble click rápido)',
        'Espera 15 segundos'
      ],
      esperado: 'Solo 1 guía creada (NO 2 duplicadas)'
    },
    {
      id: 'TEST_5',
      titulo: 'Botones +/- rápidos en cámara',
      objetivo: 'Detecta inconsistencia entre cliente y GAS al manipular qty',
      requiereProducto: true,
      pasos: [
        'Crea una guía SALIDA_ZONA nueva',
        'Escanea el producto UNA vez (qty=1)',
        'Espera 8 segundos para que GAS confirme',
        'Pulsa el botón + 5 VECES SEGUIDAS rápido',
        'Pulsa el botón − 2 VECES rápido',
        'Cierra la cámara'
      ],
      esperado: 'qty final = 4 (1 + 5 − 2)'
    },
    {
      id: 'TEST_6',
      titulo: 'Reabrir y editar cantidad',
      objetivo: 'Detecta si reabrir/recerrar duplica descuento de stock',
      requiereProducto: true,
      pasos: [
        'Crea una guía SALIDA_ZONA con el producto, qty=5',
        'CIERRA la guía (stock baja 5)',
        'REABRE la guía (admin PIN)',
        'Edita la cantidad del ítem a 3 (en lugar de 5)',
        'Cierra la guía nuevamente',
        'Espera 10 segundos'
      ],
      esperado: 'Stock final = stock original − 3 (no − 5)'
    }
  ];

  let _resultados = {};  // idTest → { estado, fecha, mensaje, idEjecucion }
  let _ejecucionActiva = null;
  let _testActivo = null;

  async function cargar() {
    try {
      const res = await API.getResultadosDiagnostico();
      if (res.ok) {
        // Indexar por idTest, tomar el más reciente
        _resultados = {};
        (res.data || []).forEach(r => {
          if (!_resultados[r.idTest] || new Date(r.fechaInicio) > new Date(_resultados[r.idTest].fechaInicio)) {
            _resultados[r.idTest] = r;
          }
        });
      }
    } catch(e) {}
    _render();
  }

  function _render() {
    const list = document.getElementById('diagListTests');
    if (!list) return;
    list.innerHTML = TESTS.map((t, i) => {
      const ult = _resultados[t.id];
      let badge = '<span class="text-xs text-slate-500">⏸ Pendiente</span>';
      if (ult) {
        if (ult.estado === 'PASS') badge = '<span class="text-xs text-emerald-400 font-bold">✓ PASS</span>';
        else if (ult.estado === 'FAIL') badge = '<span class="text-xs text-red-400 font-bold">✗ FAIL</span>';
        else if (ult.estado === 'RUNNING') badge = '<span class="text-xs text-amber-400 font-bold">⚙ Ejecutando</span>';
      }
      return `
      <div class="card-sm" style="padding:11px 13px;cursor:pointer"
           onclick="DiagnosticoView.iniciarTest('${t.id}')">
        <div class="flex items-start justify-between gap-2">
          <div class="flex-1 min-w-0">
            <p class="font-bold text-sm text-slate-100">TEST ${i + 1} — ${escHtml(t.titulo)}</p>
            <p class="text-[11px] text-slate-400 mt-0.5">${escHtml(t.objetivo)}</p>
            ${ult && ult.mensaje ? `<p class="text-[11px] mt-1 ${ult.estado === 'PASS' ? 'text-emerald-300' : 'text-red-300'}">${escHtml(ult.mensaje)}</p>` : ''}
          </div>
          <div class="text-right flex-shrink-0">
            ${badge}
            ${ult && ult.fechaInicio ? `<p class="text-[10px] text-slate-500 mt-1">${new Date(ult.fechaInicio).toLocaleString('es-PE')}</p>` : ''}
          </div>
        </div>
      </div>`;
    }).join('');
  }

  async function iniciarTest(idTest) {
    const t = TESTS.find(x => x.id === idTest);
    if (!t) return;
    _testActivo = t;
    _ejecucionActiva = null;

    const tit = document.getElementById('testDiagTitulo');
    const sub = document.getElementById('testDiagSub');
    const cnt = document.getElementById('testDiagContent');
    if (tit) tit.textContent = t.titulo;
    if (sub) sub.textContent = t.objetivo;

    // Pantalla de SETUP
    let setupHTML = '';
    if (t.requiereProducto) {
      const prods = OfflineManager.getProductosCache();
      const opts = prods
        .filter(p => p.codigoBarra && p.estado === '1')
        .slice(0, 200)
        .map(p => `<option value="${escAttr(p.codigoBarra)}">${escHtml(p.descripcion)} — ${escHtml(p.codigoBarra)}</option>`)
        .join('');
      setupHTML = `
        <div class="space-y-3">
          <div>
            <label class="text-xs text-slate-400">Selecciona un producto físico que tengas disponible:</label>
            <select id="diagProdSel" class="input mt-1">
              <option value="">— Elegir producto —</option>${opts}
            </select>
            <p class="text-[11px] text-slate-500 mt-1">⚠ Los tests modifican datos reales (stock, guías). Úsalos solo en datos de prueba.</p>
          </div>
          <div class="card-sm" style="border:1px solid rgba(14,165,233,.3);background:rgba(14,165,233,.05);padding:10px">
            <p class="text-xs font-bold text-blue-300 mb-2">📋 Pasos a seguir (lee bien antes):</p>
            <ol class="text-xs text-slate-300 space-y-1.5" style="padding-left:1.2em;list-style-type:decimal">
              ${t.pasos.map(p => `<li>${escHtml(p)}</li>`).join('')}
            </ol>
            <p class="text-[11px] text-emerald-400 mt-3"><b>Esperado:</b> ${escHtml(t.esperado)}</p>
          </div>
          <button onclick="DiagnosticoView.confirmarSetup()" id="btnTestSetup"
                  class="btn btn-primary w-full">▶ Iniciar test</button>
        </div>`;
    } else {
      setupHTML = `
        <div class="space-y-3">
          <div class="card-sm" style="border:1px solid rgba(14,165,233,.3);background:rgba(14,165,233,.05);padding:10px">
            <p class="text-xs font-bold text-blue-300 mb-2">📋 Pasos a seguir (lee bien antes):</p>
            <ol class="text-xs text-slate-300 space-y-1.5" style="padding-left:1.2em;list-style-type:decimal">
              ${t.pasos.map(p => `<li>${escHtml(p)}</li>`).join('')}
            </ol>
            <p class="text-[11px] text-emerald-400 mt-3"><b>Esperado:</b> ${escHtml(t.esperado)}</p>
          </div>
          <button onclick="DiagnosticoView.confirmarSetup()" id="btnTestSetup"
                  class="btn btn-primary w-full">▶ Iniciar test</button>
        </div>`;
    }
    if (cnt) cnt.innerHTML = setupHTML;
    abrirSheet('sheetTestDiag');
  }

  async function confirmarSetup() {
    if (!_testActivo) return;
    let codigoProducto = '';
    if (_testActivo.requiereProducto) {
      codigoProducto = document.getElementById('diagProdSel')?.value || '';
      if (!codigoProducto) { toast('Selecciona un producto', 'warn'); return; }
    }
    const btn = document.getElementById('btnTestSetup');
    if (btn) { btn.disabled = true; btn.textContent = 'Iniciando...'; }
    try {
      const res = await API.iniciarTestDiagnostico({
        idTest: _testActivo.id,
        codigoProducto,
        usuario: window.WH_CONFIG?.usuario || ''
      });
      if (!res.ok) { toast('Error: ' + res.error, 'danger'); return; }
      _ejecucionActiva = res.data.idEjecucion;

      // Pantalla de ejecución (recordatorio + botón finalizar)
      const cnt = document.getElementById('testDiagContent');
      if (cnt) cnt.innerHTML = `
        <div class="space-y-3">
          <div class="card-sm" style="border:1.5px solid rgba(245,158,11,.5);background:rgba(120,53,15,.15);padding:12px">
            <p class="text-sm font-bold text-amber-300 mb-2">⚙ Test en ejecución</p>
            <p class="text-xs text-slate-300 mb-2">Realiza los pasos físicos del test:</p>
            <ol class="text-xs text-slate-300 space-y-1.5" style="padding-left:1.2em;list-style-type:decimal">
              ${_testActivo.pasos.map(p => `<li>${escHtml(p)}</li>`).join('')}
            </ol>
            ${codigoProducto ? `<p class="text-[11px] text-amber-200 mt-2 font-mono">Producto: ${escHtml(codigoProducto)}</p>` : ''}
          </div>
          <p class="text-xs text-slate-400 text-center">Cuando hayas completado TODOS los pasos:</p>
          <button onclick="DiagnosticoView.finalizar()" id="btnTestFin"
                  class="btn btn-primary w-full">✓ Ya completé todos los pasos · Validar</button>
          <button onclick="DiagnosticoView.cancelar()" class="btn btn-outline w-full text-xs">Cancelar test</button>
        </div>`;
    } catch(e) { toast('Sin conexión', 'warn'); }
    finally { if (btn) { btn.disabled = false; btn.textContent = '▶ Iniciar test'; } }
  }

  async function finalizar() {
    if (!_ejecucionActiva) return;
    const btn = document.getElementById('btnTestFin');
    if (btn) { btn.disabled = true; btn.textContent = 'Validando...'; }
    try {
      const res = await API.finalizarTestDiagnostico({ idEjecucion: _ejecucionActiva });
      if (!res.ok) { toast('Error validando: ' + res.error, 'danger'); return; }
      _renderResultado(res.data);
    } catch(e) { toast('Sin conexión', 'warn'); }
    finally { if (btn) { btn.disabled = false; btn.textContent = '✓ Ya completé · Validar'; } }
  }

  function _renderResultado(r) {
    const cnt = document.getElementById('testDiagContent');
    if (!cnt) return;
    const passed = r.pass === true;
    cnt.innerHTML = `
      <div class="space-y-3">
        <div class="card-sm" style="border:2px solid ${passed ? 'rgba(52,211,153,.6)' : 'rgba(248,113,113,.6)'};
             background:${passed ? 'rgba(6,78,59,.2)' : 'rgba(127,29,29,.2)'};padding:14px;text-align:center">
          <p class="text-3xl mb-1">${passed ? '✓' : '✗'}</p>
          <p class="text-base font-bold ${passed ? 'text-emerald-300' : 'text-red-300'}">${passed ? 'TEST PASS' : 'TEST FAIL'}</p>
          <p class="text-xs ${passed ? 'text-emerald-200' : 'text-red-200'} mt-2">${escHtml(r.mensaje || '')}</p>
        </div>
        ${r.esperado || r.obtenido ? `
        <div class="card-sm" style="padding:10px;background:#0f172a">
          ${r.esperado ? `<p class="text-[11px] text-slate-400">Esperado: <b class="text-slate-200">${escHtml(String(r.esperado))}</b></p>` : ''}
          ${r.obtenido ? `<p class="text-[11px] text-slate-400 mt-1">Obtenido: <b class="text-slate-200">${escHtml(String(r.obtenido))}</b></p>` : ''}
          ${r.idGuia ? `<p class="text-[10px] text-slate-500 mt-1 font-mono">Guía: ${escHtml(r.idGuia)}</p>` : ''}
        </div>` : ''}
        ${passed ? `<button onclick="DiagnosticoView.cerrarYsiguiente()" class="btn btn-primary w-full">Continuar al siguiente</button>`
                 : `<button onclick="DiagnosticoView.cerrar()" class="btn btn-outline w-full">Cerrar — revisar bug</button>`}
      </div>`;
    try { passed ? SoundFX.done() : SoundFX.error(); } catch(e) {}
    cargar();  // refresca lista
  }

  function cancelar() {
    _ejecucionActiva = null;
    _testActivo = null;
    cerrarSheet('sheetTestDiag');
  }

  function cerrar() {
    _ejecucionActiva = null;
    _testActivo = null;
    cerrarSheet('sheetTestDiag');
    cargar();
  }

  function cerrarYsiguiente() {
    const idx = TESTS.findIndex(t => t.id === _testActivo?.id);
    cerrar();
    if (idx >= 0 && idx < TESTS.length - 1) {
      setTimeout(() => iniciarTest(TESTS[idx + 1].id), 400);
    }
  }

  // ── Test interno automático (sin acción manual) ──────────
  function _llenarSelectorProductos() {
    const sel = document.getElementById('testInternoProd');
    if (!sel) return;
    const prods = OfflineManager.getProductosCache();
    const stockMap = {};
    OfflineManager.getStockCache().forEach(s => {
      stockMap[String(s.codigoProducto || '').trim()] = parseFloat(s.cantidadDisponible) || 0;
    });
    // Solo productos activos con codigoBarra y stock >= 50 (para no causar negativos)
    const opts = prods
      .filter(p => p.codigoBarra && p.estado === '1')
      .filter(p => (stockMap[String(p.codigoBarra).trim()] || 0) >= 50)
      .slice(0, 100)
      .map(p => `<option value="${escAttr(p.codigoBarra)}">${escHtml(p.descripcion)} — stock ${stockMap[String(p.codigoBarra).trim()]}</option>`)
      .join('');
    sel.innerHTML = '<option value="">— Producto con stock ≥ 50 —</option>' + opts;
  }

  async function ejecutarTestInterno() {
    const sel = document.getElementById('testInternoProd');
    const cb = sel?.value;
    if (!cb) { toast('Selecciona un producto con stock ≥ 50', 'warn'); return; }
    const btn = document.getElementById('btnTestInterno');
    const resEl = document.getElementById('testInternoResultado');
    if (btn) { btn.disabled = true; btn.textContent = '⚙ Ejecutando 9 tests...'; }
    if (resEl) resEl.innerHTML = '<p class="text-xs text-slate-400 text-center py-3">Procesando...</p>';
    try {
      const res = await API.runInternalTests({ codigoBarra: cb });
      if (!res.ok) {
        if (resEl) resEl.innerHTML = `<p class="text-xs text-red-400">Error: ${escHtml(res.error || '')}</p>`;
        return;
      }
      const d = res.data;
      const passColor = c => c ? 'text-emerald-400' : 'text-red-400';
      const passIcon  = c => c ? '✓' : '✗';
      let html = `
        <div class="card-sm" style="padding:10px;background:#0f172a;border:1px solid #334155">
          <p class="text-sm font-bold mb-2">
            <span class="text-emerald-400">${d.pass} pass</span> ·
            <span class="text-red-400">${d.fail} fail</span> ·
            ${d.total} total
          </p>
          <p class="text-[10px] text-slate-500 mb-2">
            Stock original: ${d.stockOriginal} → final: ${d.stockFinal}
            ${Math.abs(d.stockFinal - d.stockOriginal) <= 0.01 ? ' ✓' : ' ⚠ DIFF'}
            ${d.idGuiaBorrada ? ' · guía test borrada: ' + escHtml(d.idGuiaBorrada) : ''}
          </p>
          <div class="space-y-1">
            ${d.resultados.map(r => `
              <div class="text-[11px] border-l-2 pl-2 ${r.pass ? 'border-emerald-500' : 'border-red-500'}">
                <span class="${passColor(r.pass)} font-bold">${passIcon(r.pass)}</span>
                <span class="text-slate-200">${escHtml(r.test)}</span>
                <span class="text-slate-500 block">${escHtml(r.detalle || '')}</span>
                ${!r.pass && r.esperado !== undefined ? `<span class="text-slate-500 block">esperado: <b>${escHtml(String(r.esperado))}</b> · obtenido: <b>${escHtml(String(r.obtenido))}</b></span>` : ''}
              </div>
            `).join('')}
          </div>
        </div>`;
      if (resEl) resEl.innerHTML = html;
      try { (d.fail === 0 ? SoundFX.done : SoundFX.error)(); } catch(e) {}
      // Refrescar histórico
      _cargarHistorialInterno();
    } catch(e) {
      if (resEl) resEl.innerHTML = `<p class="text-xs text-red-400">Sin conexión</p>`;
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = '▶ Ejecutar tests internos'; }
    }
  }

  function _renderHistorialInterno() {
    const cont = document.getElementById('testInternoHistorial');
    const list = document.getElementById('testInternoHistList');
    if (!cont || !list) return;
    const histInternos = Object.values(_resultados)
      .filter(r => r.idTest === 'INTERNO_AUTO');
    // _resultados está indexado por idTest, así que solo hay 1 INTERNO_AUTO (el más reciente)
    // Usamos getResultadosDiagnostico para traer los últimos 10 internos completos
    if (!histInternos.length) { cont.style.display = 'none'; return; }
  }

  async function _cargarHistorialInterno() {
    try {
      const res = await API.getResultadosDiagnostico();
      if (!res.ok) return;
      const internos = (res.data || []).filter(r => r.idTest === 'INTERNO_AUTO').slice(0, 8);
      const cont = document.getElementById('testInternoHistorial');
      const list = document.getElementById('testInternoHistList');
      if (!cont || !list) return;
      if (!internos.length) { cont.style.display = 'none'; return; }
      cont.style.display = '';
      list.innerHTML = internos.map(h => {
        const ok = h.estado === 'PASS';
        return `
        <div class="flex items-center justify-between text-[11px] py-1 border-b border-slate-700/50">
          <span class="${ok ? 'text-emerald-400' : 'text-red-400'} font-bold">${ok ? '✓' : '✗'}</span>
          <span class="flex-1 text-slate-300 ml-2 truncate">${escHtml(h.mensaje || '')}</span>
          <span class="text-slate-500 text-[10px] ml-2">${h.fechaInicio ? new Date(h.fechaInicio).toLocaleString('es-PE') : ''}</span>
        </div>`;
      }).join('');
    } catch(e) {}
  }

  // Llenar selector cuando se carga la vista
  const _origCargar = cargar;
  async function cargarConSelector() {
    await _origCargar();
    _llenarSelectorProductos();
    _cargarHistorialInterno();
  }

  return { cargar: cargarConSelector, iniciarTest, confirmarSetup, finalizar, cancelar, cerrar, cerrarYsiguiente, ejecutarTestInterno };
})();

// ════════════════════════════════════════════════
// [v2.13.53] MODAL LOTES — FIFO + Historial trazable
// ════════════════════════════════════════════════
let _lotCtxCodigo = '';
let _lotCtxNombre = '';
let _lotFifoCache = [];
let _lotHistCache = [];
let _lotTabActual = 'fifo';

async function abrirModalLotes(codigoProducto, nombreOpcional) {
  if (!codigoProducto) { toast('Falta código producto', 'warn'); return; }
  _lotCtxCodigo = String(codigoProducto).trim();
  _lotCtxNombre = String(nombreOpcional || '').trim();
  _lotTabActual = 'fifo';
  document.getElementById('lotHdrCodigo').textContent = _lotCtxNombre || _lotCtxCodigo;
  document.getElementById('lotHdrSub').textContent = 'FIFO · vence primero, sale primero';
  document.getElementById('lotTabFifo').classList.add('active');
  document.getElementById('lotTabHist').classList.remove('active');
  document.getElementById('lotContenidoFifo').style.display = '';
  document.getElementById('lotContenidoHist').style.display = 'none';
  document.getElementById('lotContenidoFifo').innerHTML =
    '<div class="lot-empty">⏳ Cargando lotes…</div>';
  abrirModal('modalLotes');
  try { SoundFX.beep?.(); } catch(_){}
  vibrate(15);
  try {
    const r = await API.getLotesFIFO({ codigoProducto: _lotCtxCodigo });
    _lotFifoCache = (r && r.ok) ? (r.data || []) : [];
    _renderLotesFifo();
    if (_lotFifoCache.length > 0) { try { SoundFX.done?.(); } catch(_){} }
  } catch(e) {
    document.getElementById('lotContenidoFifo').innerHTML =
      '<div class="lot-empty" style="color:#f87171">⚠ Error: ' + (e.message || e) + '</div>';
    try { SoundFX.warn?.(); } catch(_){}
  }
}

function cerrarModalLotes() {
  cerrarModal('modalLotes');
  _lotCtxCodigo = ''; _lotFifoCache = []; _lotHistCache = [];
}

async function cambiarTabLote(tab) {
  _lotTabActual = tab;
  document.getElementById('lotTabFifo').classList.toggle('active', tab === 'fifo');
  document.getElementById('lotTabHist').classList.toggle('active', tab === 'hist');
  document.getElementById('lotContenidoFifo').style.display = tab === 'fifo' ? '' : 'none';
  document.getElementById('lotContenidoHist').style.display = tab === 'hist' ? '' : 'none';
  try { SoundFX.beep?.(); } catch(_){}
  vibrate(10);
  if (tab === 'hist' && !_lotHistCache.length) {
    document.getElementById('lotContenidoHist').innerHTML =
      '<div class="lot-empty">⏳ Cargando historial…</div>';
    try {
      const idLote = (_lotFifoCache[0] && _lotFifoCache[0].idLote) || '';
      if (!idLote) {
        document.getElementById('lotContenidoHist').innerHTML =
          '<div class="lot-empty">Sin lotes activos para mostrar historial</div>';
        return;
      }
      // Cargar historial de TODOS los lotes activos
      const all = await Promise.all(_lotFifoCache.map(l =>
        API.getHistorialLote({ idLote: l.idLote }).then(r => (r && r.ok ? r.data : [])).catch(() => [])
      ));
      _lotHistCache = [].concat.apply([], all);
      _lotHistCache.sort((a, b) => String(b.ts).localeCompare(String(a.ts)));
      _renderLotesHistorial();
    } catch(e) {
      document.getElementById('lotContenidoHist').innerHTML =
        '<div class="lot-empty" style="color:#f87171">⚠ Error: ' + (e.message || e) + '</div>';
    }
  }
}

function _renderLotesFifo() {
  const cont = document.getElementById('lotContenidoFifo');
  if (!_lotFifoCache.length) {
    cont.innerHTML = `
      <div class="lot-empty">
        <div style="font-size:2.5rem;margin-bottom:8px">📭</div>
        <div>Sin lotes activos para este producto</div>
        <div style="font-size:.65rem;margin-top:4px;opacity:.7">
          Los lotes se crean al ingresar mercadería con fecha de vencimiento
        </div>
      </div>`;
    return;
  }
  const html = _lotFifoCache.map((l, idx) => {
    const dias = l.diasRestantes;
    let cardCls = 'ok', badgeTxt = '✓ vigente', badgeCls = 'lot-bd-ok';
    if (dias !== null) {
      if (dias < 0)      { cardCls = 'vencido'; badgeTxt = '🚨 VENCIDO ' + Math.abs(dias) + 'd'; badgeCls = 'lot-bd-vencido'; }
      else if (dias <= 7)  { cardCls = 'vencido'; badgeTxt = '⚠ vence en ' + dias + 'd'; badgeCls = 'lot-bd-vencido'; }
      else if (dias <= 30) { cardCls = 'proxim';  badgeTxt = '⏳ ' + dias + 'd'; badgeCls = 'lot-bd-proxim'; }
      else                 { badgeTxt = '✓ ' + dias + 'd'; }
    } else {
      badgeTxt = 'sin fecha'; badgeCls = '';
    }
    const fifoBadge = idx === 0 ? '<span class="lot-badge lot-bd-fifo">🎯 FIFO</span>' : '';
    const fechaFmt = l.fechaVencimiento ? new Date(l.fechaVencimiento).toLocaleDateString('es-PE') : '—';
    return `
      <div class="lot-card ${cardCls}">
        <div class="lot-top">
          <div>
            <span class="lot-cant">${l.cantidadActual} u</span>
            <span class="lot-id" style="margin-left:8px">${l.idLote}</span>
          </div>
          <div style="display:flex;gap:5px">
            ${fifoBadge}
            <span class="lot-badge ${badgeCls}">${badgeTxt}</span>
          </div>
        </div>
        <div class="lot-meta">
          <span>📅 Vence: <b style="color:#f8fafc">${fechaFmt}</b></span>
          <span>📋 Guía: <b style="color:#cbd5e1;font-family:ui-monospace,monospace">${l.idGuia || '—'}</b></span>
        </div>
      </div>`;
  }).join('');
  cont.innerHTML = html;
}

function _renderLotesHistorial() {
  const cont = document.getElementById('lotContenidoHist');
  if (!_lotHistCache.length) {
    cont.innerHTML = '<div class="lot-empty">Sin movimientos registrados</div>';
    return;
  }
  const html = _lotHistCache.map(m => {
    const tsStr = m.ts ? new Date(m.ts).toLocaleString('es-PE') : '';
    return `
      <div class="lot-hist-item">
        <span class="lot-hist-acc ${m.accion || ''}">${m.accion || '?'}</span>
        <div style="flex:1">
          <div style="color:#cbd5e1"><b style="color:#f8fafc">${m.cantidad || 0}u</b> · ${m.motivo || ''}</div>
          <div style="color:#94a3b8;font-size:.6rem;margin-top:2px">
            ${tsStr} · ${m.usuario || '—'} · guía ${m.idGuia || '—'}
          </div>
        </div>
      </div>`;
  }).join('');
  cont.innerHTML = html;
}

// ════════════════════════════════════════════════════════════════════
// [v2.13.108] WhLoteAdhesivo — orquestación de impresión por lotes
// ════════════════════════════════════════════════════════════════════
//
// Cliente del sistema de lotes del backend (Envasados.gs). Maneja:
//   - Creación del lote
//   - Sub-jobs uno a uno con polling
//   - GAPDETECT condicional (al inicio + post OUT_OF_PAPER)
//   - Modal de progreso con barra, velocidad, ETA
//   - Detección de fin de rollo + reanudación
//
// API pública:
//   WhLoteAdhesivo.crearYEjecutar({ codigoBarra, total, descripcion?, vto?, fechaEnvasado? })
//   WhLoteAdhesivo.continuar()
//   WhLoteAdhesivo.cancelar()
//   WhLoteAdhesivo.cerrar()
//
// Es un módulo self-contained — no depende de MOS. Reutiliza el mismo
// backend `crearLoteAdhesivo`/`imprimirSubLoteAdhesivo` que MOS.
const WhLoteAdhesivo = (() => {
  'use strict';
  let _state = null;
  // _state = { idLote, total, completadas, subJobSize, status, ultimoError,
  //            descripcion, codigoBarra, vto, tInicio, orquestando }

  async function crearYEjecutar(opts) {
    if (_state) {
      try { toast('Ya hay un lote en curso. Termínalo primero.', 'warn'); } catch(_) {}
      return;
    }
    const cb = String(opts.codigoBarra || '').trim();
    const total = parseInt(opts.total) || 0;
    if (!cb || total <= 0) {
      try { toast('Datos de lote inválidos', 'error'); } catch(_) {}
      return;
    }

    // [v2.13.113 OPTIMISTIC] Abrir modal de progreso INMEDIATAMENTE con
    // estado "Iniciando…" placeholder. El operario ve feedback instantáneo
    // mientras el backend crea el lote (que puede tardar 1-3s con red lenta).
    _abrirModalProgreso({
      idLote:      '',  // se completa cuando responde el backend
      total:       total,
      completadas: 0,
      subJobSize:  10,  // estimación inicial (backend la confirma)
      descripcion: opts.descripcion || '',
      codigoBarra: cb,
      vto:         opts.vto || ''
    });
    _setStatus('CREADO');

    // 1. Crear lote en backend (en background, ya con modal abierto)
    let r;
    try {
      r = await API.post('crearLoteAdhesivo', {
        codigoBarra:      cb,
        descripcion:      opts.descripcion || '',
        total:            total,
        usuario:          window.WH_CONFIG?.usuario || '',
        origen:           'WH',
        vto:              opts.vto || '',
        fechaEnvasado:    opts.fechaEnvasado || '',
        idempotencyKey:   'wh_lote_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)
      });
    } catch (e) {
      _setStatus('PAUSADO_ERROR', 'Sin conexión al crear el lote: ' + (e?.message || ''));
      return;
    }
    if (r && r.ok === false) {
      _setStatus('PAUSADO_ERROR', 'Backend rechazó: ' + (r.error || 'desconocido'));
      return;
    }
    // WH API.post retorna {ok, data}. NO desempaca como MOS.
    const d = r.data || r;
    // Re-check: usuario puede haber cancelado el modal mientras esperaba.
    if (!_state) return;
    // 2. Completar metadata del state con la respuesta real
    _state.idLote      = d.idLote;
    _state.total       = d.total || total;
    _state.subJobSize  = d.subJobSize || 10;
    _state.descripcion = d.descripcion || _state.descripcion;
    _state.vto         = d.vto || _state.vto;
    _render();
    // 3. Arrancar orquestación
    _orquestar(d.idLote);
  }

  async function _orquestar(idLote, runOpts) {
    if (!_state || _state.idLote !== idLote) return;
    if (_state.orquestando) return;
    _state.orquestando = true;
    // [v2.13.109 AUDIT FIX #2] requireGapDetect debe consumirse SOLO en el
    // primer ciclo. Antes era const → todos los sub-jobs subsiguientes
    // hacían GAPDETECT también, perdiendo ~1 etq por cada 10.
    // [v2.13.143] QUITADO el `_state.completadas === 0` que forzaba GAPDETECT
    // automático en el primer sub-job de CADA lote. Eso consumía 3 etiquetas
    // gratis por lote, contradiciendo la estrategia 1-calibración-por-rollo.
    // Ahora SOLO se hace GAPDETECT si el caller lo pidió explícitamente
    // (botón "🔧 Calibrar rollo nuevo" del modal calibrador).
    let pendingGapDetect = !!(runOpts && runOpts.requireGapDetect);
    try {
      while (_state && _state.idLote === idLote) {
        if (['CANCELADO','COMPLETADO','PAUSADO_USUARIO','PAUSADO_OUT_PAPER','PAUSADO_ERROR'].indexOf(_state.status) >= 0) break;
        const necesitaCal = pendingGapDetect;
        if (necesitaCal) _setStatus('CALIBRANDO');
        let r;
        try {
          r = await API.post('imprimirSubLoteAdhesivo', {
            idLote:           idLote,
            requireGapDetect: necesitaCal
          });
          if (necesitaCal) pendingGapDetect = false;
        } catch (e) {
          _setStatus('PAUSADO_ERROR', 'Sin conexión: ' + (e?.message || ''));
          break;
        }
        if (!_state || _state.idLote !== idLote) break;
        if (r && r.ok === false) {
          _setStatus(r.status || 'PAUSADO_ERROR', r.error || 'Error desconocido');
          if ((r.status || '') === 'PAUSADO_OUT_PAPER') _mostrarRolloAgotado();
          break;
        }
        // [v2.13.112 BUG FIX] WH API.post NO desempaca data — leer de r.data.
        const subRes = r.data || r;
        _state.completadas = subRes.completadas || _state.completadas;
        _state.status      = subRes.status || 'IMPRIMIENDO';
        _render();
        if (_state.status === 'COMPLETADO') {
          _celebrar();
          break;
        }
        await new Promise(res => setTimeout(res, 250));
      }
    } finally {
      if (_state) _state.orquestando = false;
    }
  }

  function _setStatus(status, errMsg) {
    if (!_state) return;
    _state.status = status;
    if (errMsg) _state.ultimoError = errMsg;
    _render();
  }

  function _abrirModalProgreso(meta) {
    _state = {
      idLote:      meta.idLote,
      total:       meta.total,
      completadas: 0,
      subJobSize:  meta.subJobSize || 10,
      descripcion: meta.descripcion,
      codigoBarra: meta.codigoBarra,
      vto:         meta.vto,
      status:      'CREADO',
      ultimoError: '',
      orquestando: false,
      tInicio:     Date.now()
    };
    const html = `
      <div class="wh-lote-overlay" id="whLoteOverlay">
        <div class="wh-lote-modal">
          <div class="wh-lote-head">
            <div class="wh-lote-emoji">🏭</div>
            <div class="wh-lote-text">
              <div class="wh-lote-title">LOTE DE IMPRESIÓN</div>
              <div class="wh-lote-sub">${_esc(_state.descripcion || _state.codigoBarra)}</div>
            </div>
          </div>
          <div class="wh-lote-body">
            <div class="wh-lote-stat">
              <div class="wh-lote-chip" id="whLoteChip">⏳ creado</div>
              <div class="wh-lote-counter" id="whLoteCounter">0 / ${_state.total}</div>
            </div>
            <div class="wh-lote-bar"><div class="wh-lote-fill" id="whLoteFill" style="width:0%"></div></div>
            <div class="wh-lote-info">
              <span id="whLoteVel">— etq/min</span>
              <span id="whLoteEta">estimado: —</span>
            </div>
            <div class="wh-lote-err" id="whLoteErr" style="display:none"></div>
            <div class="wh-lote-actions" id="whLoteActions">
              <button class="wh-lote-btn-warn" onclick="WhLoteAdhesivo.cancelar()">⊘ Cancelar lote</button>
            </div>
          </div>
        </div>
      </div>`;
    document.body.insertAdjacentHTML('beforeend', html);
    _render();
  }

  function _render() {
    if (!_state) return;
    const pct = _state.total > 0 ? (_state.completadas / _state.total * 100) : 0;
    const fill = document.getElementById('whLoteFill');
    if (fill) fill.style.width = pct.toFixed(1) + '%';
    const counter = document.getElementById('whLoteCounter');
    if (counter) counter.textContent = _state.completadas + ' / ' + _state.total;
    const chip = document.getElementById('whLoteChip');
    if (chip) {
      const map = {
        CREADO:            { cls: 'info',  txt: '<span class="wh-lote-spin">◐</span> preparando lote…' },
        CALIBRANDO:        { cls: 'warn',  txt: '<span class="wh-lote-spin">◐</span> calibrando rollo…' },
        IMPRIMIENDO:       { cls: 'ok',    txt: '🖨 imprimiendo' },
        PAUSADO_USUARIO:   { cls: 'warn',  txt: '⏸ pausado' },
        PAUSADO_OUT_PAPER: { cls: 'error', txt: '🛑 rollo agotado' },
        PAUSADO_ERROR:     { cls: 'error', txt: '❌ error' },
        COMPLETADO:        { cls: 'ok',    txt: '✅ completado' },
        CANCELADO:         { cls: 'warn',  txt: '⊘ cancelado' }
      };
      const m = map[_state.status] || { cls: 'info', txt: _state.status };
      chip.className = 'wh-lote-chip wh-lote-chip-' + m.cls;
      chip.innerHTML = m.txt;  // soporta spinner HTML
    }
    const elapsedSec = (Date.now() - _state.tInicio) / 1000;
    if (_state.completadas > 0 && elapsedSec > 1) {
      const velMin = (_state.completadas / elapsedSec * 60).toFixed(0);
      const restante = _state.total - _state.completadas;
      const segRestantes = restante / (_state.completadas / elapsedSec);
      const elVel = document.getElementById('whLoteVel');
      const elEta = document.getElementById('whLoteEta');
      if (elVel) elVel.textContent = velMin + ' etq/min';
      if (elEta) elEta.textContent = 'estimado: ' + Math.ceil(segRestantes) + ' seg';
    }
    const elErr = document.getElementById('whLoteErr');
    if (elErr) {
      if (_state.ultimoError) {
        elErr.style.display = 'block';
        elErr.textContent = '⚠ ' + _state.ultimoError;
      } else {
        elErr.style.display = 'none';
      }
    }
  }

  function _mostrarRolloAgotado() {
    const actions = document.getElementById('whLoteActions');
    if (!actions || !_state) return;
    const restante = _state.total - _state.completadas;
    actions.innerHTML = `
      <div class="wh-lote-alert">
        <div class="wh-lote-alert-title">🛑 Rollo agotado</div>
        <div class="wh-lote-alert-msg">
          Se imprimieron ~${_state.completadas} de ${_state.total}.<br>
          Faltan ${restante}. Cambiá el rollo y dale a Continuar.<br>
          <span class="wh-lote-alert-warn">⚠ Hasta ${_state.subJobSize} pueden duplicarse del rollo viejo.</span>
        </div>
        <div class="wh-lote-alert-btns">
          <button class="wh-lote-btn-primary" onclick="WhLoteAdhesivo.continuar()">✓ Continuar (rollo nuevo)</button>
          <button class="wh-lote-btn-warn" onclick="WhLoteAdhesivo.cancelar()">⊘ Cancelar</button>
        </div>
      </div>`;
  }

  function _celebrar() {
    try { toast('✅ Lote completado: ' + _state.total + ' adhesivos', 'ok', 6000); } catch(_) {}
    try { vibrate && vibrate([100, 50, 100]); } catch(_) {}
    setTimeout(cerrar, 2500);
  }

  function continuar() {
    if (!_state) return;
    _setStatus('CALIBRANDO');
    const actions = document.getElementById('whLoteActions');
    if (actions) actions.innerHTML = '<button class="wh-lote-btn-warn" onclick="WhLoteAdhesivo.cancelar()">⊘ Cancelar lote</button>';
    _orquestar(_state.idLote, { requireGapDetect: true });
  }

  // [Lote4 · M4-WH] confirm() nativo -> confirmación INLINE en el panel del lote
  // (mismo patrón que _mostrarRolloAgotado; la casa prohíbe prompt/confirm/alert nativos).
  function cancelar() {
    if (!_state) return;
    const actions = document.getElementById('whLoteActions');
    if (!actions) { _cancelarConfirmado(); return; }   // sin panel → cancelar directo
    actions.innerHTML = `
      <div class="wh-lote-alert">
        <div class="wh-lote-alert-title">¿Cancelar el lote?</div>
        <div class="wh-lote-alert-msg">
          ${_state.completadas} de ${_state.total} ya impresas <b>no se borran</b> del rollo.
        </div>
        <div class="wh-lote-alert-btns">
          <button class="wh-lote-btn-warn" onclick="WhLoteAdhesivo._cancelarConfirmado()">⊘ Sí, cancelar</button>
          <button class="wh-lote-btn-primary" onclick="WhLoteAdhesivo._abortarCancelacion()">← Volver</button>
        </div>
      </div>`;
    try { vibrate && vibrate([40]); } catch(_) {}
  }
  async function _cancelarConfirmado() {
    if (!_state) return;
    try { await API.post('cancelarLoteAdhesivo', { idLote: _state.idLote }); } catch(_) {}
    _setStatus('CANCELADO');
    setTimeout(cerrar, 800);
  }
  // Volver del prompt de cancelación: restaurar el botón Cancelar normal.
  function _abortarCancelacion() {
    const actions = document.getElementById('whLoteActions');
    if (actions) actions.innerHTML = '<button class="wh-lote-btn-warn" onclick="WhLoteAdhesivo.cancelar()">⊘ Cancelar lote</button>';
  }

  // [v2.13.115] Cancelar sin pedir confirm — usado por rollback automático
  // cuando el envasado paralelo falla y el lote ya estaba en vuelo.
  async function cancelarSilencioso(motivo) {
    if (!_state) return;
    if (_state.idLote) {
      try { await API.post('cancelarLoteAdhesivo', { idLote: _state.idLote }); } catch(_) {}
    }
    _setStatus('CANCELADO', motivo || '');
    setTimeout(cerrar, 1000);
  }

  function cerrar() {
    const ov = document.getElementById('whLoteOverlay');
    if (ov) ov.remove();
    _state = null;
  }

  function _esc(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  return { crearYEjecutar, continuar, cancelar, _cancelarConfirmado, _abortarCancelacion, cancelarSilencioso, cerrar };
})();

// ════════════════════════════════════════════════════════════════════
// WhAdhesivoReprint — [v2.13.150]
// ════════════════════════════════════════════════════════════════════
// Mini-modal de "reimprimir adhesivos" para el historial del Envasador.
// Pedido del usuario 2026-06-04: que el envasador pueda re-imprimir
// adhesivos de un envasado anterior sin pasar por MOS.
//
// Flujo:
//   1. Click "🖨 Adhesivo" en una card del historial
//   2. Modal pide cantidad (default = unidadesProducidas del envasado)
//   3. Botón "Imprimir" → llama WhLoteAdhesivo.crearYEjecutar(...) que
//      abre el modal de progreso ya existente y orquesta los sub-jobs.
//
// Reutiliza la infra de WhLoteAdhesivo (modal progreso, polling, etc.).
// ════════════════════════════════════════════════════════════════════
const WhAdhesivoReprint = (() => {
  'use strict';

  function abrir(idEnvasado) {
    // 1. Buscar el envasado en el cache offline
    let env = null;
    try {
      const cache = OfflineManager.getEnvasadosCache ? OfflineManager.getEnvasadosCache() : [];
      env = cache.find(x => String(x.idEnvasado) === String(idEnvasado));
    } catch (e) {}
    if (!env) {
      try { toast('Envasado no encontrado en cache', 'error'); } catch(_) {}
      return;
    }
    const codigoBarra = String(env.codigoProductoEnvasado || '').trim();
    if (!codigoBarra) {
      try { toast('Sin código envasado para imprimir adhesivos', 'warn'); } catch(_) {}
      return;
    }
    // descripcion: priorizar campo backend, fallback al mapa local
    const descripcion = env.descripcionProductoEnvasado
                     || (window._prodMap && window._prodMap[codigoBarra])
                     || codigoBarra;
    const defaultCantidad = Math.max(1, parseInt(env.unidadesProducidas) || 1);
    const fechaEnvasado = env.fecha || env.fechaImpresion || '';

    // [v2.13.165 PREVIEW] Construir lista de siblings (otros productos envasables)
    // para que AdhesivoPreview detecte highlights diferenciales. Sin esto, solo
    // el último token (peso) queda destacado — funcional pero menos contextual.
    let siblings = [];
    try {
      const stockArr = (window.todosProductos || []) || [];
      const norm = window.AdhesivoPreview ? AdhesivoPreview.normalize : (s => String(s||''));
      stockArr.forEach(p => {
        const c = String(p.codigoBarra || p.idProducto || '').toUpperCase();
        if (c.indexOf('WH') !== 0) return;
        const desc = p.descripcion || '';
        if (!desc) return;
        siblings.push(norm(desc).toUpperCase().split(/\s+/).filter(Boolean));
      });
    } catch(_) {}

    _abrirModal({
      idEnvasado:   idEnvasado,
      codigoBarra:  codigoBarra,
      descripcion:  descripcion,
      defaultCant:  defaultCantidad,
      fechaEnvasado: fechaEnvasado,
      siblings:     siblings
    });
  }

  function _abrirModal(datos) {
    if (document.getElementById('whAdhReprintOverlay')) return;

    // [v2.13.165 PREVIEW] Procesar datos con AdhesivoPreview para render
    // pixel-perfect del adhesivo. Si el módulo no está cargado (CDN caído),
    // omitir el preview y mostrar solo el modal de cantidad como fallback.
    let previewHtml = '';
    let svgIdPreview = null;
    if (window.AdhesivoPreview) {
      AdhesivoPreview.inyectarCss();
      const proc = AdhesivoPreview.procesar({
        codigoBarra:   datos.codigoBarra,
        descripcion:   datos.descripcion,
        fechaEnvasado: datos.fechaEnvasado,
        siblings:      datos.siblings || []
      });
      if (proc) {
        svgIdPreview = 'whAdhReprintSvg_' + Date.now();
        const cardHtml = AdhesivoPreview.renderHtml(proc, {
          cantidad: datos.defaultCant,
          svgId:    svgIdPreview
        });
        // Scale 0.5 para que el adhesivo de 600px quepa cómodo en el modal
        previewHtml = '<div style="display:flex;justify-content:center;padding:14px 0;overflow:hidden">'
                    +   '<div style="transform:scale(.55);transform-origin:top center;height:165px">'
                    +     cardHtml
                    +   '</div>'
                    + '</div>';
        // Guardamos los datos procesados para el dibujo del barcode tras insertar
        window._whAdhReprintProc = proc;
        window._whAdhReprintSvgId = svgIdPreview;
      }
    }

    const html = `
      <div id="whAdhReprintOverlay"
           class="fixed inset-0 z-[99996] flex items-center justify-center p-4"
           style="background:rgba(2,6,23,.78);backdrop-filter:blur(12px)"
           onclick="if(event.target===this) WhAdhesivoReprint.cerrar()">
        <div class="w-full max-w-md rounded-2xl border overflow-hidden shadow-2xl"
             style="background:linear-gradient(180deg,#0a1424,#070d18);border-color:rgba(34,197,94,.4);max-height:92vh;display:flex;flex-direction:column"
             onclick="event.stopPropagation()">
          <header class="p-5 border-b border-slate-800 flex items-center gap-3"
                  style="background:linear-gradient(135deg,rgba(34,197,94,.10),rgba(22,163,74,.04))">
            <div class="text-4xl">🏷</div>
            <div class="flex-1 min-w-0">
              <div class="text-sm font-black text-emerald-400 tracking-wider">REIMPRIMIR ADHESIVO</div>
              <div class="text-xs text-slate-400 mt-1 truncate">${_esc(datos.descripcion)}</div>
              <div class="text-[10px] font-mono text-slate-500 mt-0.5">▌${_esc(datos.codigoBarra)}</div>
            </div>
            <button onclick="WhAdhesivoReprint.cerrar()"
                    class="text-slate-400 hover:text-slate-200 text-xl px-2"
                    title="Cerrar">✕</button>
          </header>

          ${previewHtml}

          <section class="p-5 space-y-4" style="overflow-y:auto">
            <div>
              <label class="text-xs font-bold text-slate-300 uppercase tracking-wider">Cantidad a imprimir</label>
              <div class="flex items-center gap-2 mt-2">
                <button onclick="WhAdhesivoReprint.delta(-1)"
                        class="w-12 h-12 rounded-lg bg-slate-800 hover:bg-slate-700 text-emerald-400 text-2xl font-bold">−</button>
                <input id="whAdhReprintCant" type="number"
                       value="${datos.defaultCant}" min="1" max="999"
                       oninput="WhAdhesivoReprint.setCant(this.value)"
                       class="flex-1 h-12 px-4 text-center text-2xl font-bold rounded-lg bg-slate-900 border border-slate-700 text-slate-100">
                <button onclick="WhAdhesivoReprint.delta(1)"
                        class="w-12 h-12 rounded-lg bg-slate-800 hover:bg-slate-700 text-emerald-400 text-2xl font-bold">+</button>
              </div>
              <div class="text-[11px] text-slate-500 mt-2">
                Por defecto = ${datos.defaultCant} (unidades envasadas)
              </div>
            </div>

            <div class="text-[11px] text-slate-400 bg-slate-900/60 rounded-lg p-3 border border-slate-800">
              <div class="flex items-start gap-2">
                <span class="text-amber-400 mt-0.5">💡</span>
                <span>El adhesivo lleva la fecha de envasado original (vto +1 año). Reimprimir NO altera stock ni crea nuevo envasado.</span>
              </div>
            </div>
          </section>

          <footer class="p-4 border-t border-slate-800 flex gap-2">
            <button onclick="WhAdhesivoReprint.cerrar()"
                    class="flex-1 h-12 rounded-lg border border-slate-700 text-slate-300 hover:bg-slate-800 font-bold">
              Cancelar
            </button>
            <button id="whAdhReprintBtnOk"
                    onclick="WhAdhesivoReprint.imprimir()"
                    class="flex-1 h-12 rounded-lg bg-emerald-500 hover:bg-emerald-400 text-slate-900 font-black">
              🖨 Imprimir
            </button>
          </footer>
        </div>
      </div>`;
    document.body.insertAdjacentHTML('beforeend', html);
    window._whAdhReprintDatos = datos;
    document.addEventListener('keydown', _keydown);
    setTimeout(() => {
      // [v2.13.165] Dibujar el barcode SVG tras insertar al DOM
      try {
        if (window.AdhesivoPreview && window._whAdhReprintProc && window._whAdhReprintSvgId) {
          const svgEl = document.getElementById(window._whAdhReprintSvgId);
          if (svgEl) AdhesivoPreview.dibujarBarcode(svgEl, window._whAdhReprintProc.codigoBarra);
        }
      } catch(eBC) { console.warn('[WhAdhReprint] preview barcode fail:', eBC && eBC.message); }
      const inp = document.getElementById('whAdhReprintCant');
      if (inp) { inp.focus(); inp.select(); }
    }, 100);
  }

  function _keydown(e) {
    if (e.key === 'Escape') cerrar();
    if (e.key === 'Enter' && document.getElementById('whAdhReprintOverlay')) {
      e.preventDefault();
      imprimir();
    }
  }

  function _actualizarCantTag(n) {
    // [v2.13.165] Sincronizar la "×cantidad" del preview con el input
    const tag = document.querySelector('#whAdhReprintOverlay .adhesivo-cantidad-tag');
    if (tag) tag.textContent = '×' + n;
  }

  function delta(d) {
    const inp = document.getElementById('whAdhReprintCant');
    if (!inp) return;
    const nuevo = Math.max(1, Math.min(999, (parseInt(inp.value) || 1) + d));
    inp.value = nuevo;
    _actualizarCantTag(nuevo);
  }

  function setCant(v) {
    // sanitiza pero no bloquea — el max se aplica al imprimir
    let n = parseInt(v) || 1;
    const inp = document.getElementById('whAdhReprintCant');
    if (n < 1)   { n = 1;   if (inp) inp.value = 1; }
    if (n > 999) { n = 999; if (inp) inp.value = 999; }
    _actualizarCantTag(n);
  }

  function imprimir() {
    const datos = window._whAdhReprintDatos;
    if (!datos) return;
    const inp = document.getElementById('whAdhReprintCant');
    const cant = Math.max(1, Math.min(999, parseInt(inp && inp.value) || datos.defaultCant));

    cerrar();
    // [v2.13.152 FIX] WhLoteAdhesivo se declara con `const` a nivel de archivo.
    // `const` NO se expone a window automáticamente (solo `var` lo hace). El
    // check `window.WhLoteAdhesivo` siempre era false → toast 'no disponible'.
    // Fix: referencia directa al const del scope. Como WhAdhesivoReprint se
    // declara DESPUÉS de WhLoteAdhesivo, este último ya está definido.
    if (typeof WhLoteAdhesivo !== 'undefined' && WhLoteAdhesivo.crearYEjecutar) {
      WhLoteAdhesivo.crearYEjecutar({
        codigoBarra:    datos.codigoBarra,
        descripcion:    datos.descripcion,
        total:          cant,
        vto:            '',  // backend lo recalcula desde fechaEnvasado
        fechaEnvasado:  datos.fechaEnvasado
      });
    } else {
      try { toast('Sistema de lotes no disponible', 'error'); } catch(_) {}
    }
  }

  function cerrar() {
    document.removeEventListener('keydown', _keydown);
    const ov = document.getElementById('whAdhReprintOverlay');
    if (ov) ov.remove();
    // [v2.13.165] Limpieza completa de globals del preview
    try { delete window._whAdhReprintDatos; }   catch(_) { window._whAdhReprintDatos = null; }
    try { delete window._whAdhReprintProc; }    catch(_) { window._whAdhReprintProc = null; }
    try { delete window._whAdhReprintSvgId; }   catch(_) { window._whAdhReprintSvgId = null; }
  }

  function _esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])
    );
  }

  return { abrir, delta, setCant, imprimir, cerrar };
})();
// [v2.13.152] Expongo a window para que los onclick="WhAdhesivoReprint.xxx()"
// del HTML inline funcionen. `const` NO crea propiedad en window por sí solo.
window.WhAdhesivoReprint = WhAdhesivoReprint;

// ════════════════════════════════════════════════
// Init
// ════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => App.init());
