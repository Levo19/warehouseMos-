// ============================================================
// warehouseMos — clienteInbox.js
// Polling para alertar pedidos nuevos del Portal Cliente (pedido.html)
//
// Cada 20 segundos consulta `clienteInboxPolling` y, si hay pedidos
// confirmados desde el último check, muestra un banner pulsante y
// suena una "campanada" distinta a la alerta de pickup ME, para que
// el almacenero distinga "pedido de cliente externo" de los pickups.
// ============================================================

(function() {
  'use strict';

  var POLL_MS    = 20000;     // 20s
  var LS_KEY_TS  = 'wh_cliInbox_lastTs';
  var pollTimer  = null;

  // ── API helpers ────────────────────────────────────────────
  function gasUrl() {
    return (window.cfg && window.cfg.gasUrl) || '';
  }
  async function consultar(desde) {
    var url = gasUrl(); if (!url) return null;
    try {
      var qs = '?action=clienteInboxPolling&desde=' + encodeURIComponent(desde || 0);
      var r  = await fetch(url + qs, { method: 'GET' });
      var d  = await r.json();
      return (d && d.ok && d.data) ? d.data : null;
    } catch(e) {
      return null;
    }
  }

  // ── Sonido distintivo (Web Audio API, 3 tonos ascendentes) ─
  var _ac;
  function ac() { return _ac ||= new (window.AudioContext || window.webkitAudioContext)(); }
  function dingDong() {
    try {
      [880, 1175, 1568].forEach(function(f, i) {
        setTimeout(function() {
          var a = ac(), o = a.createOscillator(), g = a.createGain();
          o.type = 'sine'; o.frequency.value = f; o.connect(g); g.connect(a.destination);
          g.gain.setValueAtTime(0, a.currentTime);
          g.gain.linearRampToValueAtTime(0.28, a.currentTime + 0.02);
          g.gain.exponentialRampToValueAtTime(0.0001, a.currentTime + 0.4);
          o.start(); o.stop(a.currentTime + 0.4);
        }, i * 150);
      });
    } catch(e) {}
  }

  // ── Inyectar CSS una sola vez ──────────────────────────────
  function ensureStyles() {
    if (document.getElementById('cliInboxStyles')) return;
    var s = document.createElement('style');
    s.id = 'cliInboxStyles';
    s.textContent = `
      .cliInboxBanner {
        position: fixed; top: 12px; left: 50%; transform: translateX(-50%);
        max-width: 92vw; min-width: 280px; z-index: 9998;
        background: linear-gradient(135deg, #0ea5e9, #10b981);
        color: #03121a; padding: 14px 16px; border-radius: 14px;
        box-shadow: 0 16px 40px -10px rgba(14,165,233,.55), 0 0 0 6px rgba(14,165,233,.08);
        font-family: system-ui, -apple-system, sans-serif;
        animation: cliInPop .35s ease both, cliInPulse 1.4s ease-in-out 0.35s 4;
      }
      @keyframes cliInPop { from { transform: translate(-50%, -120%); opacity: 0; } to { transform: translate(-50%, 0); opacity: 1; } }
      @keyframes cliInPulse { 0%,100% { box-shadow: 0 16px 40px -10px rgba(14,165,233,.55), 0 0 0 6px rgba(14,165,233,.08); } 50% { box-shadow: 0 16px 40px -10px rgba(14,165,233,.55), 0 0 0 14px rgba(14,165,233,.18); } }
      .cliInboxBanner .titulo { font-weight: 800; font-size: 14px; display: flex; align-items: center; gap: 8px; }
      .cliInboxBanner .sub { font-size: 12.5px; margin-top: 4px; opacity: .92; }
      .cliInboxBanner .actions { display: flex; gap: 8px; margin-top: 10px; }
      .cliInboxBanner button { border: 0; border-radius: 9px; padding: 7px 12px; font-weight: 700; cursor: pointer; font-size: 12.5px; }
      .cliInboxBanner .btn-primary { background: #03121a; color: #fff; }
      .cliInboxBanner .btn-ghost { background: rgba(255,255,255,.5); color: #03121a; }
      .cliInboxBanner .badge { background: #03121a; color: #fff; border-radius: 999px; padding: 1px 8px; font-size: 11px; font-weight: 800; }
    `;
    document.head.appendChild(s);
  }

  // ── Mostrar banner ─────────────────────────────────────────
  function mostrarBanner(items) {
    ensureStyles();
    // Limpia banners previos
    document.querySelectorAll('.cliInboxBanner').forEach(function(b) { b.remove(); });
    var div = document.createElement('div');
    div.className = 'cliInboxBanner';
    var primero = items[0];
    var resto   = items.length - 1;
    div.innerHTML = `
      <div class="titulo">🔔 Nuevo pedido cliente <span class="badge">${items.length}</span></div>
      <div class="sub">
        <b>${escapeHtml(primero.cliente || primero.token || 'Cliente')}</b> · ${primero.items} ítems
        ${resto > 0 ? ' · y ' + resto + ' más' : ''}
      </div>
      <div class="actions">
        <button class="btn-primary" data-act="ver">▶ Ver lista sombra</button>
        <button class="btn-ghost" data-act="cerrar">Cerrar</button>
      </div>
    `;
    div.querySelector('[data-act="cerrar"]').onclick = function() { div.remove(); };
    div.querySelector('[data-act="ver"]').onclick = function() {
      div.remove();
      // Intentar abrir el módulo de listas sombra / despacho rápido si está disponible
      if (window.MOS && typeof window.MOS.irADespachoRapido === 'function') {
        try { window.MOS.irADespachoRapido(); return; } catch(e) {}
      }
      // Fallback: si hay hash para listas sombra
      try { location.hash = '#listas'; } catch(e) {}
    };
    document.body.appendChild(div);
    // Vibración cross-browser
    try { if (navigator.vibrate) navigator.vibrate([60, 80, 60, 80, 120]); } catch(e) {}
    dingDong();
    // Auto-dismiss tras 12s si no interactúa
    setTimeout(function() { if (div.parentNode) div.remove(); }, 12000);
  }

  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, function(c) {
      return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
    });
  }

  // ── Loop de polling ────────────────────────────────────────
  async function tick() {
    var lastTs = parseFloat(localStorage.getItem(LS_KEY_TS) || '0');
    // Primera vez: marcar el "ahora" como baseline para no notificar pedidos viejos
    if (!lastTs) {
      var ini = await consultar(0);
      if (ini && ini.ahora) localStorage.setItem(LS_KEY_TS, String(ini.ahora));
      return;
    }
    var d = await consultar(lastTs);
    if (!d) return;
    if (d.nuevos && d.nuevos.length > 0) mostrarBanner(d.nuevos);
    if (d.ahora) localStorage.setItem(LS_KEY_TS, String(d.ahora));
  }

  function start() {
    if (pollTimer) return;
    tick();
    pollTimer = setInterval(tick, POLL_MS);
  }
  function stop() { if (pollTimer) { clearInterval(pollTimer); pollTimer = null; } }

  // "Despertar" audio context al primer tap (requerido por iOS)
  function wake() { try { if (_ac && _ac.state === 'suspended') _ac.resume(); } catch(e) {} }
  document.addEventListener('click', wake, { once: true });
  document.addEventListener('touchstart', wake, { once: true });

  // Arranca cuando el DOM esté listo (después de api.js y app.js)
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    setTimeout(start, 500);  // pequeño delay para que window.cfg esté listo
  }

  // Pausar polling si la pestaña está oculta (ahorra batería)
  document.addEventListener('visibilitychange', function() {
    if (document.hidden) stop(); else start();
  });

  // Exponer por si app.js quiere disparar manual
  window.cliInbox = { start: start, stop: stop, tick: tick, _ding: dingDong };
})();
