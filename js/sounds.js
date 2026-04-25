// warehouseMos — sounds.js
// Síntesis de audio con Web Audio API (sin archivos externos)
const SoundFX = (() => {
  let _ctx = null;

  function _getCtx() {
    if (!_ctx) {
      try { _ctx = new (window.AudioContext || window.webkitAudioContext)(); } catch(e) {}
    }
    if (_ctx?.state === 'suspended') _ctx.resume().catch(() => {});
    return _ctx;
  }

  function _tone(freq, dur, type = 'square', vol = 0.22, delay = 0) {
    try {
      const ctx = _getCtx();
      if (!ctx) return;
      const t    = ctx.currentTime + delay;
      const osc  = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
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
    // Scan exacto — beep corto tipo caja registradora
    beep:       () => _tone(1900, 0.09, 'square', 0.22),

    // Mismo producto +1 — doble beep ascendente (distinguible del primero)
    beepDouble: () => { _tone(1500, 0.06, 'square', 0.18, 0); _tone(2100, 0.06, 'square', 0.18, 0.09); },

    // No encontrado — tono descendente de advertencia
    warn:       () => { _tone(520, 0.12, 'sine', 0.2, 0); _tone(340, 0.16, 'sine', 0.18, 0.13); },

    // Error GAS — buzz corto
    error:      () => _tone(160, 0.22, 'sawtooth', 0.18),

    // Listo / cerrar cámara con ítems — chime ascendente 2 notas
    done:       () => { _tone(880, 0.12, 'sine', 0.2, 0); _tone(1320, 0.18, 'sine', 0.2, 0.15); },
  };
})();
