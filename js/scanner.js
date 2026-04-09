// warehouseMos — scanner.js  Escáner de códigos de barra (cámara)
const Scanner = (() => {
  let codeReader = null;
  let active = false;
  let videoEl = null;

  function _loadZXing() {
    return new Promise((res, rej) => {
      if (window.ZXing) return res();
      const s = document.createElement('script');
      s.src = 'https://unpkg.com/@zxing/library@latest/umd/index.min.js';
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  async function start(videoElementId, onResult, onError) {
    await _loadZXing();
    videoEl = document.getElementById(videoElementId);
    if (!videoEl) return;

    try {
      codeReader = new ZXing.BrowserMultiFormatReader();
      active = true;

      const devices = await codeReader.listVideoInputDevices();
      // Preferir cámara trasera
      const backCam = devices.find(d =>
        d.label.toLowerCase().includes('back') ||
        d.label.toLowerCase().includes('rear') ||
        d.label.toLowerCase().includes('trasera')
      );
      const deviceId = backCam ? backCam.deviceId : (devices[0]?.deviceId);

      await codeReader.decodeFromVideoDevice(deviceId, videoElementId, (result, err) => {
        if (result && active) {
          const text = result.getText();
          stop();
          onResult(text);
        }
        if (err && !(err instanceof ZXing.NotFoundException)) {
          console.warn('[Scanner]', err);
        }
      });
    } catch (err) {
      active = false;
      if (onError) onError(err.message);
    }
  }

  function stop() {
    if (codeReader) {
      codeReader.reset();
      codeReader = null;
    }
    if (videoEl) {
      const stream = videoEl.srcObject;
      if (stream) stream.getTracks().forEach(t => t.stop());
      videoEl.srcObject = null;
    }
    active = false;
  }

  function isActive() { return active; }

  return { start, stop, isActive };
})();
