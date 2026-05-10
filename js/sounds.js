// warehouseMos — sounds.js
// Síntesis de audio con Web Audio API (sin archivos externos)
const SoundFX = (() => {
  let _ctx = null;
  let _comp = null;

  function _getCtx() {
    if (!_ctx) {
      try { _ctx = new (window.AudioContext || window.webkitAudioContext)(); } catch(e) {}
    }
    if (!_ctx) return null;
    if (_ctx.state === 'suspended') _ctx.resume().catch(() => {});
    if (!_comp) {
      _comp = _ctx.createDynamicsCompressor();
      _comp.threshold.value = -6;
      _comp.knee.value      = 2;
      _comp.ratio.value     = 12;
      _comp.attack.value    = 0.001;
      _comp.release.value   = 0.08;
      _comp.connect(_ctx.destination);
    }
    return _ctx;
  }

  function _tone(freq, dur, type = 'square', vol = 0.8, delay = 0) {
    try {
      const ctx = _getCtx();
      if (!ctx) return;
      const t    = ctx.currentTime + delay;
      const osc  = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(_comp);
      osc.type = type;
      osc.frequency.setValueAtTime(freq, t);
      gain.gain.setValueAtTime(0, t);
      gain.gain.linearRampToValueAtTime(vol, t + 0.005);
      gain.gain.exponentialRampToValueAtTime(0.001, t + dur);
      osc.start(t);
      osc.stop(t + dur + 0.02);
    } catch(e) {}
  }

  return {
    // Scan exacto — beep corto y agudo tipo caja registradora
    beep:       () => _tone(1900, 0.13, 'square', 0.9),

    // Mismo producto +1 — doble beep ascendente
    beepDouble: () => { _tone(1500, 0.09, 'square', 0.8, 0); _tone(2100, 0.09, 'square', 0.85, 0.11); },

    // No encontrado — tono descendente duro (square para cortar el ruido)
    warn:       () => { _tone(600, 0.15, 'square', 0.85, 0); _tone(360, 0.22, 'square', 0.8, 0.16); },

    // Error GAS — buzz grave y largo, sawtooth
    error:      () => _tone(180, 0.35, 'sawtooth', 0.9),

    // Listo / cerrar cámara con ítems — chime ascendente 2 notas
    done:       () => { _tone(880, 0.18, 'sine', 0.85, 0); _tone(1320, 0.25, 'sine', 0.9, 0.20); },

    // Login OK — woosh metálico + chime industrial (3 tonos descendente-ascendente)
    welcome: () => {
      _tone(220, 0.08, 'sawtooth', 0.5, 0);     // woosh corto grave
      _tone(440, 0.10, 'triangle', 0.7, 0.06);
      _tone(660, 0.12, 'triangle', 0.8, 0.13);
      _tone(880, 0.18, 'sine',     0.9, 0.22);  // chime metálico final
    },

    // Buzzer corto — acceso denegado / error UX
    buzzer: () => {
      _tone(140, 0.12, 'square', 0.95, 0);
      _tone(120, 0.18, 'square', 0.9, 0.13);
    },

    // Factory bell — aviso 5 min antes del cierre (3 toques)
    bell: () => {
      _tone(1200, 0.13, 'triangle', 0.85, 0);
      _tone(1200, 0.13, 'triangle', 0.85, 0.25);
      _tone(1200, 0.18, 'triangle', 0.9,  0.50);
    },

    // Cierre forzado — sirena descendente larga
    closeAlarm: () => {
      _tone(900, 0.30, 'sawtooth', 0.85, 0);
      _tone(700, 0.30, 'sawtooth', 0.85, 0.30);
      _tone(500, 0.40, 'sawtooth', 0.9,  0.60);
    },

    // Ping notificación — para nuevos pendientes
    ping: () => _tone(2200, 0.10, 'sine', 0.7),

    // Click suave — para ajustes manuales +/- (no satura al repetirse)
    click: () => _tone(1100, 0.04, 'sine', 0.4),

    // Pickup nuevo — alerta fuerte y repetida (almacén con ruido).
    // 3 ciclos de tonos urgentes para que se escuche aunque haya ruido.
    pickupAlerta: () => {
      _tone(1800, 0.20, 'square', 1.0, 0.00);
      _tone(2400, 0.20, 'square', 1.0, 0.22);
      _tone(1800, 0.20, 'square', 1.0, 0.55);
      _tone(2400, 0.20, 'square', 1.0, 0.77);
      _tone(1800, 0.30, 'square', 1.0, 1.10);
      _tone(2400, 0.40, 'square', 1.0, 1.35);
    },

    // Pickup completado — chime grande de éxito (4 notas ascendentes).
    pickupOk: () => {
      _tone(523,  0.12, 'triangle', 0.85, 0.00);  // C5
      _tone(659,  0.12, 'triangle', 0.90, 0.13);  // E5
      _tone(784,  0.14, 'triangle', 0.95, 0.26);  // G5
      _tone(1047, 0.30, 'sine',     1.00, 0.40);  // C6 sustained
    },

    // Pickup item completo — beep agudo + chime corto (sentido de "ok marcado")
    pickupItemOk: () => {
      _tone(2200, 0.08, 'sine', 0.7, 0.00);
      _tone(2800, 0.18, 'sine', 0.8, 0.08);
    },

    // ════ Sonidos del flujo de update ════

    // Llegó nueva versión — chime descendente-ascendente moderno (notificación)
    updateArrived: () => {
      _tone(523, 0.10, 'sine',     0.7, 0.00);   // C5
      _tone(784, 0.12, 'triangle', 0.85, 0.10);  // G5
      _tone(1047, 0.18, 'sine',    0.85, 0.22);  // C6
    },

    // Update lista para instalar — chime tipo "ta-da" exitoso
    updateReady: () => {
      _tone(659,  0.10, 'triangle', 0.8, 0.00);  // E5
      _tone(880,  0.10, 'triangle', 0.85, 0.10); // A5
      _tone(1175, 0.12, 'triangle', 0.9, 0.20);  // D6
      _tone(1568, 0.30, 'sine',     1.0, 0.32);  // G6 sustained
    },

    // Cohete despegando — sweep ascendente largo + sustain agudo
    rocket: () => {
      // Rumble grave (sawtooth) que sube — efecto cohete
      try {
        const ctx = _getCtx(); if (!ctx) return;
        const t = ctx.currentTime;
        const osc1 = ctx.createOscillator();
        const g1   = ctx.createGain();
        osc1.connect(g1); g1.connect(_comp);
        osc1.type = 'sawtooth';
        osc1.frequency.setValueAtTime(120, t);
        osc1.frequency.exponentialRampToValueAtTime(2400, t + 0.85);
        g1.gain.setValueAtTime(0, t);
        g1.gain.linearRampToValueAtTime(0.6, t + 0.05);
        g1.gain.exponentialRampToValueAtTime(0.001, t + 0.95);
        osc1.start(t); osc1.stop(t + 1.0);
      } catch(_){}
      // Beeps agudos al final (señal lift-off completo)
      _tone(1568, 0.10, 'square', 0.7, 0.85);
      _tone(2093, 0.18, 'square', 0.9, 0.95);
    },

    // Tick urgente — countdown final
    tickUrgent: () => _tone(1500, 0.04, 'square', 0.6),
  };
})();
