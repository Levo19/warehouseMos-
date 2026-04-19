// warehouseMos — scanner.js
// Escáner de códigos de barra (cámara)
// Estrategia: BarcodeDetector nativo (Chrome/Android rápido) → ZXing fallback
const Scanner = (() => {
  let _active   = false;
  let _stream   = null;
  let _videoEl  = null;
  let _raf      = null;
  let _zxReader = null;
  let _detector = null;

  // ── Carga ZXing solo si lo necesitamos ─────────────────
  function _loadZXing() {
    return new Promise((res, rej) => {
      if (window.ZXing) return res();
      const s = document.createElement('script');
      // Versión pinada: @latest puede resolver a versiones con bugs en móvil
      s.src = 'https://unpkg.com/@zxing/library@0.20.0/umd/index.min.js';
      s.onload = res;
      s.onerror = () => {
        // fallback al CDN original si unpkg falla
        const s2 = document.createElement('script');
        s2.src = 'https://unpkg.com/@zxing/library@latest/umd/index.min.js';
        s2.onload = res; s2.onerror = rej;
        document.head.appendChild(s2);
      };
      document.head.appendChild(s);
    });
  }

  // ── BarcodeDetector nativo (Chrome 83+, Android Chrome rápido) ─
  function _hasBarcodeDetector() {
    return typeof BarcodeDetector !== 'undefined';
  }

  async function _buildDetector() {
    try {
      // Detectar formatos soportados por el browser/dispositivo
      const supported = await BarcodeDetector.getSupportedFormats().catch(() => []);
      const want = ['ean_13','ean_8','code_128','code_39','itf','qr_code',
                    'upc_a','upc_e','data_matrix','codabar','aztec'];
      const formats = want.filter(f => supported.length === 0 || supported.includes(f));
      return new BarcodeDetector({ formats: formats.length ? formats : ['ean_13','code_128'] });
    } catch (e) {
      return null;
    }
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
        stop();
        onResult(code);
      } else {
        _raf = requestAnimationFrame(() => _rafLoop(onResult));
      }
    }).catch(() => {
      if (_active) _raf = requestAnimationFrame(() => _rafLoop(onResult));
    });
  }

  // ── Path A: BarcodeDetector (getusermedia manual + rAF loop) ──
  async function _startNative(onResult, onError) {
    try {
      _stream = await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: 'environment' },
          width:  { ideal: 1920 },
          height: { ideal: 1080 },
        },
        audio: false,
      });
      _videoEl.srcObject = _stream;
      await _videoEl.play().catch(() => {});
      _rafLoop(onResult);
    } catch (err) {
      _active = false;
      if (onError) onError(err.message || 'Error cámara');
    }
  }

  // ── Path B: ZXing (maneja su propia cámara internamente) ──────
  async function _startZXing(onResult, onError) {
    await _loadZXing();
    const hints = new Map();
    hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [
      ZXing.BarcodeFormat.EAN_13,
      ZXing.BarcodeFormat.EAN_8,
      ZXing.BarcodeFormat.CODE_128,
      ZXing.BarcodeFormat.CODE_39,
      ZXing.BarcodeFormat.ITF,
      ZXing.BarcodeFormat.QR_CODE,
      ZXing.BarcodeFormat.DATA_MATRIX,
      ZXing.BarcodeFormat.UPC_A,
      ZXing.BarcodeFormat.UPC_E,
      ZXing.BarcodeFormat.CODABAR,
    ]);
    hints.set(ZXing.DecodeHintType.TRY_HARDER, true);

    _zxReader = new ZXing.BrowserMultiFormatReader(hints, 500);

    try {
      const devices = await _zxReader.listVideoInputDevices();
      const backCam = devices.find(d =>
        /back|rear|trasera|environment/i.test(d.label)
      ) || devices[devices.length - 1]; // el último suele ser la trasera
      const deviceId = backCam?.deviceId ?? undefined;

      await _zxReader.decodeFromVideoDevice(deviceId, _videoEl.id, (result, err) => {
        if (result && _active) {
          const text = result.getText();
          stop();
          onResult(text);
        }
        // NotFoundException es normal (cada frame sin código), ignorar
        if (err && err.name && !err.name.includes('NotFoundException')) {
          console.warn('[Scanner/ZXing]', err.name);
        }
      });
    } catch (err) {
      _active = false;
      if (onError) onError(err.message || 'Error cámara');
    }
  }

  // ── API pública ───────────────────────────────────────────────
  async function start(videoElementId, onResult, onError) {
    // Si ya hay uno activo, detenerlo primero
    if (_active) stop();

    _videoEl = document.getElementById(videoElementId);
    if (!_videoEl) { if (onError) onError('Elemento video no encontrado'); return; }
    _active = true;

    if (_hasBarcodeDetector()) {
      _detector = await _buildDetector();
      if (_detector) {
        await _startNative(onResult, onError);
        return;
      }
    }
    // Fallback ZXing
    await _startZXing(onResult, onError);
  }

  function stop() {
    _active = false;
    if (_raf) { cancelAnimationFrame(_raf); _raf = null; }
    if (_zxReader) { try { _zxReader.reset(); } catch (e) {} _zxReader = null; }
    if (_stream)   { _stream.getTracks().forEach(t => t.stop()); _stream = null; }
    if (_videoEl)  { _videoEl.srcObject = null; _videoEl = null; }
  }

  function isActive() { return _active; }

  return { start, stop, isActive };
})();
