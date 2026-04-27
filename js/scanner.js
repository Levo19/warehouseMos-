// warehouseMos — scanner.js
// Estrategia: BarcodeDetector nativo (Chrome/Android) → ZXing fallback
// Mejoras v2: constraints 720p + focusMode continuous, double-confirm anti-falsos, torch
const Scanner = (() => {
  let _active      = false;
  let _stream      = null;
  let _videoEl     = null;
  let _raf         = null;
  let _zxReader    = null;
  let _detector    = null;
  let _continuous  = false;
  let _cooldown    = 1200;
  let _lastCode    = null;
  let _lastCodeTs  = 0;
  // Anti-doble-conteo: el código tiene que "desaparecer" del frame para volver a contarse.
  // _emptyFrames cuenta frames sin detección consecutivos; al superar el umbral, _lastCode se libera.
  let _emptyFrames = 0;
  const _EMPTY_FRAMES_RESET = 5;
  // Double-confirm: el mismo código debe leerse 2 veces en _CONFIRM_WINDOW ms
  let _pendingCode = null;
  let _pendingTs   = 0;
  const _CONFIRM_WINDOW = 1500;

  // Longitudes esperadas por formato — descarta lecturas parciales
  const _FORMAT_LEN = {
    ean_13: 13, ean_8: 8, upc_a: 12, upc_e: 8,
    itf: null, code_128: null, code_39: null, codabar: null,
    qr_code: null, data_matrix: null, aztec: null
  };
  function _formatoValido(code, format) {
    if (!format) return code.length >= 6; // sin formato → mínimo 6 chars
    const expected = _FORMAT_LEN[format];
    if (expected == null) return code.length >= 4; // formatos variables: solo descarta súper cortos
    return code.length === expected;
  }

  // Constraints optimizadas para lectura de códigos de barra
  // 1080p ideal: detecta barras finas y códigos pequeños mejor que 720p.
  // ideal (no exact) — dispositivos viejos siguen entregando lo que pueden.
  const _CONSTRAINTS = {
    video: {
      facingMode: { ideal: 'environment' },
      width:      { ideal: 1920 },
      height:     { ideal: 1080 },
      frameRate:  { ideal: 15, max: 30 },
      advanced:   [{ focusMode: 'continuous' }]
    },
    audio: false
  };

  function _loadZXing() {
    return new Promise((res, rej) => {
      if (window.ZXing) return res();
      const s = document.createElement('script');
      s.src = 'https://unpkg.com/@zxing/library@0.20.0/umd/index.min.js';
      s.onload = res;
      s.onerror = () => {
        const s2 = document.createElement('script');
        s2.src = 'https://unpkg.com/@zxing/library@latest/umd/index.min.js';
        s2.onload = res; s2.onerror = rej;
        document.head.appendChild(s2);
      };
      document.head.appendChild(s);
    });
  }

  function _hasBarcodeDetector() { return typeof BarcodeDetector !== 'undefined'; }

  async function _buildDetector() {
    try {
      const supported = await BarcodeDetector.getSupportedFormats().catch(() => []);
      const want = ['ean_13','ean_8','code_128','code_39','itf','qr_code',
                    'upc_a','upc_e','data_matrix','codabar','aztec'];
      const formats = want.filter(f => supported.length === 0 || supported.includes(f));
      return new BarcodeDetector({ formats: formats.length ? formats : ['ean_13','code_128'] });
    } catch (e) { return null; }
  }

  // Devuelve true solo cuando el mismo código se lee por 2ª vez dentro de _CONFIRM_WINDOW
  function _confirm(code) {
    const now = Date.now();
    if (code === _pendingCode && now - _pendingTs <= _CONFIRM_WINDOW) {
      _pendingCode = null; _pendingTs = 0;
      return true;
    }
    _pendingCode = code; _pendingTs = now;
    return false;
  }

  function _rafLoop(onResult) {
    if (!_active || !_videoEl) return;
    if (_videoEl.readyState < 2 || _videoEl.paused) {
      _raf = requestAnimationFrame(() => _rafLoop(onResult));
      return;
    }
    _detector.detect(_videoEl).then(codes => {
      if (!_active) return;
      if (codes.length > 0) {
        _emptyFrames = 0; // reset porque hay detección
        const code   = codes[0].rawValue;
        const format = codes[0].format;
        // Filtrar lecturas inválidas/parciales por longitud según formato
        if (!_formatoValido(code, format)) {
          _raf = requestAnimationFrame(() => _rafLoop(onResult));
          return;
        }
        // Doble-confirm: el mismo código debe leerse 2 veces seguidas para evitar lecturas parciales
        if (!_confirm(code)) {
          setTimeout(() => { if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult)); }, 60);
          return;
        }
        if (_continuous) {
          // Solo aceptar si es código distinto al último confirmado.
          // Para repetir el mismo código, debe primero desaparecer del frame (_emptyFrames).
          if (code !== _lastCode) {
            _lastCode = code; _lastCodeTs = Date.now();
            onResult(code);
          }
          setTimeout(() => { if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult)); }, 300);
        } else {
          stop(); onResult(code);
        }
      } else {
        // Sin detección este frame: contar; si llega al umbral, liberar _lastCode
        _emptyFrames++;
        if (_emptyFrames >= _EMPTY_FRAMES_RESET) {
          _lastCode = null;
        }
        _raf = requestAnimationFrame(() => _rafLoop(onResult));
      }
    }).catch(() => {
      if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult));
    });
  }

  async function _startNative(onResult, onError) {
    try {
      _stream = await navigator.mediaDevices.getUserMedia(_CONSTRAINTS);
      _videoEl.srcObject = _stream;
      await _videoEl.play().catch(() => {});
      _rafLoop(onResult);
    } catch (err) {
      _active = false;
      if (onError) onError(err.message || 'Error cámara');
    }
  }

  async function _startZXing(onResult, onError) {
    await _loadZXing();

    const hints = new Map();
    hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [
      ZXing.BarcodeFormat.EAN_13, ZXing.BarcodeFormat.EAN_8,
      ZXing.BarcodeFormat.CODE_128, ZXing.BarcodeFormat.CODE_39,
      ZXing.BarcodeFormat.ITF, ZXing.BarcodeFormat.QR_CODE,
      ZXing.BarcodeFormat.DATA_MATRIX, ZXing.BarcodeFormat.UPC_A,
      ZXing.BarcodeFormat.UPC_E, ZXing.BarcodeFormat.CODABAR,
    ]);
    hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
    _zxReader = new ZXing.BrowserMultiFormatReader(hints, 500);

    const _zxCb = (result, err) => {
      if (result && _active) {
        _emptyFrames = 0;
        const text = result.getText();
        if (!_confirm(text)) return;
        if (_continuous) {
          // Mismo anti-doble-conteo: solo si es código distinto
          if (text !== _lastCode) {
            _lastCode = text; _lastCodeTs = Date.now();
            onResult(text);
          }
        } else { stop(); onResult(text); }
      } else if (_active && _continuous) {
        // ZXing reporta NotFoundException cuando no detecta — usar para liberar _lastCode
        _emptyFrames++;
        if (_emptyFrames >= _EMPTY_FRAMES_RESET) _lastCode = null;
      }
      if (err && err.name && !err.name.includes('NotFoundException')) {
        console.warn('[Scanner/ZXing]', err.name);
      }
    };

    try {
      // Pre-adquirir stream con constraints mejoradas, pasarlo a ZXing
      _stream = await navigator.mediaDevices.getUserMedia(_CONSTRAINTS);
      _videoEl.srcObject = _stream;
      if (typeof _zxReader.decodeFromStream === 'function') {
        await _zxReader.decodeFromStream(_stream, _videoEl, _zxCb);
      } else {
        // ZXing < 0.19: fallback sin stream personalizado
        throw new Error('decodeFromStream not available');
      }
    } catch (_) {
      // Último recurso: ZXing gestiona su propia cámara
      try {
        const devices = await _zxReader.listVideoInputDevices();
        const backCam = devices.find(d => /back|rear|trasera|environment/i.test(d.label))
          || devices[devices.length - 1];
        await _zxReader.decodeFromVideoDevice(backCam?.deviceId ?? undefined, _videoEl.id, _zxCb);
        _stream = _videoEl.srcObject || null;
      } catch (err2) {
        _active = false;
        if (onError) onError(err2.message || 'Error cámara');
      }
    }
  }

  async function start(videoElementId, onResult, onError, options) {
    if (_active) stop();
    _continuous  = options?.continuous || false;
    _cooldown    = options?.cooldown   || 1200;
    _lastCode    = null; _lastCodeTs  = 0;
    _pendingCode = null; _pendingTs   = 0;
    _emptyFrames = 0;

    _videoEl = document.getElementById(videoElementId);
    if (!_videoEl) { if (onError) onError('Video no encontrado'); return; }
    _active = true;

    // Tap-to-focus: registrar listener (el video puede ser distinto en cada modal)
    _videoEl.addEventListener('click',     _onVideoTap);
    _videoEl.addEventListener('touchstart', _onVideoTap, { passive: true });
    _videoEl.style.cursor = 'crosshair';

    if (_hasBarcodeDetector()) {
      _detector = await _buildDetector();
      if (_detector) { await _startNative(onResult, onError); return; }
    }
    await _startZXing(onResult, onError);
  }

  function stop() {
    _active = false;
    if (_raf)      { cancelAnimationFrame(_raf); _raf = null; }
    if (_zxReader) { try { _zxReader.reset(); } catch (e) {} _zxReader = null; }
    if (_stream)   { _stream.getTracks().forEach(t => t.stop()); _stream = null; }
    if (_videoEl) {
      _videoEl.removeEventListener('click', _onVideoTap);
      _videoEl.removeEventListener('touchstart', _onVideoTap);
      _videoEl.style.cursor = '';
    }
    if (_videoEl)  { _videoEl.srcObject = null; _videoEl = null; }
    _pendingCode = null; _pendingTs = 0;
  }

  // ── Tap-to-focus: el usuario toca el video → forzar re-enfoque
  // Útil en sensores tipo Samsung A56 cuyo autofocus continuous no enfoca cerca
  let _tapBusy = false;
  async function _doTapFocus(x, y) {
    if (!_stream || _tapBusy) return false;
    const track = _stream.getVideoTracks()[0];
    if (!track) return false;
    const caps = track.getCapabilities?.() || {};
    _tapBusy = true;
    try {
      // Si soporta pointsOfInterest + focusMode, usar tap-to-focus real
      const adv = {};
      if (caps.pointsOfInterest && x != null && y != null) {
        adv.pointsOfInterest = [{ x, y }];
      }
      if (Array.isArray(caps.focusMode) && caps.focusMode.includes('single-shot')) {
        adv.focusMode = 'single-shot';
      }
      if (Object.keys(adv).length) {
        await track.applyConstraints({ advanced: [adv] });
        // Volver a continuous tras enfocar
        setTimeout(() => {
          _tapBusy = false;
          if (Array.isArray(caps.focusMode) && caps.focusMode.includes('continuous')) {
            track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
          }
        }, 800);
        return true;
      }
    } catch (e) { /* dispositivo no soporta — ignorar */ }
    _tapBusy = false;
    return false;
  }

  function _onVideoTap(e) {
    if (!_videoEl) return;
    const rect = _videoEl.getBoundingClientRect();
    const cx = e.touches ? e.touches[0].clientX : e.clientX;
    const cy = e.touches ? e.touches[0].clientY : e.clientY;
    if (cx == null || cy == null) { _doTapFocus(0.5, 0.5); return; }
    const x = (cx - rect.left) / rect.width;
    const y = (cy - rect.top)  / rect.height;
    _doTapFocus(Math.max(0, Math.min(1, x)), Math.max(0, Math.min(1, y)));
  }

  async function toggleTorch(on) {
    if (!_stream) return false;
    const track = _stream.getVideoTracks()[0];
    if (!track) return false;
    // Verificar capability primero — algunos Samsung rechazan el constraint si no se chequea
    try {
      const caps = track.getCapabilities?.();
      if (caps && 'torch' in caps && caps.torch === false) return false;
    } catch (e) { /* ignorar — algunos navegadores no exponen getCapabilities */ }
    try {
      await track.applyConstraints({ advanced: [{ torch: !!on }] });
      return true;
    } catch (e) { return false; }
  }

  async function setZoom(value) {
    if (!_stream) return false;
    const track = _stream.getVideoTracks()[0];
    if (!track) return false;
    try {
      await track.applyConstraints({ advanced: [{ zoom: value }] });
      return true;
    } catch (e) { return false; }
  }

  function getZoomCaps() {
    if (!_stream) return null;
    const track = _stream.getVideoTracks()[0];
    if (!track) return null;
    const caps = track.getCapabilities?.();
    return caps?.zoom || null; // { min, max, step } o null si no soportado
  }

  function isActive() { return _active; }

  return { start, stop, isActive, toggleTorch, setZoom, getZoomCaps };
})();
