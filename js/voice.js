// ============================================================
// warehouseMos — js/voice.js
// Wrapper Web Speech API (STT + TTS).
//
// API:
//   Voice.listen({lang, onResult, onError})
//   Voice.stopListen()
//   Voice.speak(text, {rate, voice, lang})
//   Voice.cancel()
//   Voice.supported() → { stt, tts }
// ============================================================

(function() {
  'use strict';

  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  let _rec = null;
  let _listening = false;

  function supported() {
    return {
      stt: !!SR,
      tts: !!(window.speechSynthesis && window.SpeechSynthesisUtterance)
    };
  }

  function listen(opts) {
    opts = opts || {};
    if (!SR) {
      if (opts.onError) opts.onError(new Error('STT no soportado en este navegador'));
      return;
    }
    if (_listening) { stopListen(); }
    _rec = new SR();
    _rec.lang           = opts.lang || 'es-PE';
    _rec.continuous     = !!opts.continuous;
    _rec.interimResults = opts.interim !== false;
    _rec.maxAlternatives = 1;

    _rec.onresult = e => {
      const last = e.results[e.results.length - 1];
      const txt  = last[0].transcript.trim();
      const isFinal = last.isFinal;
      if (typeof opts.onResult === 'function') opts.onResult(txt, isFinal);
    };
    _rec.onerror = e => { if (typeof opts.onError === 'function') opts.onError(e); };
    _rec.onend   = () => { _listening = false; if (typeof opts.onEnd === 'function') opts.onEnd(); };

    try { _rec.start(); _listening = true; }
    catch(e) { if (typeof opts.onError === 'function') opts.onError(e); }
  }

  function stopListen() {
    if (_rec && _listening) {
      try { _rec.stop(); } catch(_){}
    }
    _listening = false;
  }

  // ── TTS ──
  function speak(text, opts) {
    opts = opts || {};
    if (!window.speechSynthesis) return Promise.resolve();
    return new Promise(res => {
      const u = new SpeechSynthesisUtterance(text);
      u.lang  = opts.lang || 'es-PE';
      u.rate  = opts.rate || 1;
      u.pitch = opts.pitch || 1;
      u.volume= opts.volume != null ? opts.volume : 1;
      if (opts.voice) {
        const v = window.speechSynthesis.getVoices().find(x => x.name === opts.voice);
        if (v) u.voice = v;
      }
      u.onend = res;
      u.onerror = res;
      window.speechSynthesis.speak(u);
    });
  }

  function cancel() {
    if (window.speechSynthesis) window.speechSynthesis.cancel();
  }

  // ── Lectura de items de guía (formato simple: "cantidad nombreProducto") ──
  async function leerItems(items) {
    cancel();
    if (!items || !items.length) return;
    const partes = items
      .filter(i => (parseFloat(i.cantidad) || 0) > 0)
      .map(i => {
        const c = parseFloat(i.cantidad) || 0;
        const n = String(i.nombre || i.descripcion || i.codigoProducto || 'producto').trim();
        return c + ' ' + n;
      });
    if (!partes.length) return;
    const texto = partes.join(', ');
    await speak(texto, { rate: 0.95 });
  }

  window.Voice = {
    listen, stopListen, speak, cancel,
    leerItems, supported,
    isListening: () => _listening
  };
})();
