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
  // Double-confirm: el mismo código debe leerse 2 veces en _CONFIRM_WINDOW ms
  let _pendingCode = null;
  let _pendingTs   = 0;
  const _CONFIRM_WINDOW = 1000;

  // Constraints optimizadas para lectura de códigos de barra
  const _CONSTRAINTS = {
    video: {
      facingMode: { ideal: 'environment' },
      width:      { ideal: 1280 },
      height:     { ideal: 720 },
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
        const code = codes[0].rawValue;
        if (!_confirm(code)) {
          // Primera lectura — esperar confirmación rápida
          setTimeout(() => { if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult)); }, 80);
          return;
        }
        if (_continuous) {
          const now = Date.now();
          if (code !== _lastCode || now - _lastCodeTs > _cooldown) {
            _lastCode = code; _lastCodeTs = now;
            onResult(code);
          }
          setTimeout(() => { if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult)); }, 300);
        } else {
          stop(); onResult(code);
        }
      } else {
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
        const text = result.getText();
        if (!_confirm(text)) return;
        if (_continuous) {
          const now = Date.now();
          if (text !== _lastCode || now - _lastCodeTs > _cooldown) {
            _lastCode = text; _lastCodeTs = now;
            onResult(text);
          }
        } else { stop(); onResult(text); }
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

    _videoEl = document.getElementById(videoElementId);
    if (!_videoEl) { if (onError) onError('Video no encontrado'); return; }
    _active = true;

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
    if (_videoEl)  { _videoEl.srcObject = null; _videoEl = null; }
    _pendingCode = null; _pendingTs = 0;
  }

  async function toggleTorch(on) {
    if (!_stream) return false;
    const track = _stream.getVideoTracks()[0];
    if (!track) return false;
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
