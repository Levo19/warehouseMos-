// ============================================================
// warehouseMos — js/sharesheet.js
// Hoja de compartir con VISTA PREVIA. Resuelve el "a veces solo texto":
//  • Móvil: la Web Share API con archivo exige activación de usuario reciente; si
//    el fetch+render demoran, esa ventana expira y el share del archivo falla.
//    Aquí la imagen YA está lista y el share se dispara en el toque de "Enviar"
//    (activación FRESCA) → confiable, adjunta la imagen.
//  • PC / WhatsApp Web: NO se puede adjuntar un archivo por web. Lo mejor: COPIAR
//    la imagen al portapapeles (para pegarla con Ctrl+V en el chat) + descargarla
//    + abrir WhatsApp Web con el texto.
// window.ShareSheet(blob, filename, texto)
// ============================================================
(function() {
  'use strict';
  function _descargar(url, filename) {
    try { const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); } catch(_) {}
  }
  function _toast(msg, tipo, ms) { try { if (typeof window.toast === 'function') window.toast(msg, tipo || 'info', ms || 5000); } catch(_) {} }

  window.ShareSheet = function(blob, filename, texto) {
    if (!blob) { window.open('https://wa.me/?text=' + encodeURIComponent(texto || ''), '_blank'); return; }
    const url = URL.createObjectURL(blob);
    const file = (function(){ try { return new File([blob], filename, { type: 'image/png' }); } catch(_) { return null; } })();
    const puedeArchivo = !!(file && navigator.canShare && navigator.canShare({ files: [file] }));
    const sendLabel = puedeArchivo ? '📲 Enviar por WhatsApp' : '📋 Copiar imagen + abrir WhatsApp';
    const nota = puedeArchivo ? '' :
      'En PC WhatsApp Web no adjunta archivos por web. Al tocar el botón: se <b>copia la imagen</b> (pégala con <b>Ctrl+V</b> en el chat), se <b>descarga</b> (por si prefieres arrastrarla) y se abre WhatsApp con el texto.';

    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;z-index:13000;background:rgba(3,6,12,.92);display:flex;align-items:center;justify-content:center;padding:16px;-webkit-tap-highlight-color:transparent';
    ov.innerHTML =
      '<div style="max-width:430px;width:100%;background:#0e1626;border:1px solid #24324a;border-radius:18px;overflow:hidden;display:flex;flex-direction:column;max-height:90vh">' +
        '<div style="padding:12px 15px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #1a2740">' +
          '<span style="font-weight:800;color:#e7eef8;font-size:14px">Vista previa</span>' +
          '<button data-x style="background:#131d30;border:1px solid #24324a;color:#93a4c2;border-radius:8px;width:30px;height:30px;font-size:15px;cursor:pointer">✕</button>' +
        '</div>' +
        '<div style="overflow:auto;background:#05080e;flex:1;display:flex;justify-content:center;padding:10px">' +
          '<img src="' + url + '" alt="" style="max-width:100%;height:auto;border-radius:8px;align-self:flex-start">' +
        '</div>' +
        (nota ? '<div style="padding:10px 15px 0;color:#7f8ba3;font-size:11px;line-height:1.45">' + nota + '</div>' : '') +
        '<div style="padding:12px 15px;display:flex;gap:9px">' +
          '<button data-send style="flex:1;background:#25d366;color:#04310f;border:none;border-radius:12px;padding:13px;font-weight:800;font-size:14px;cursor:pointer">' + sendLabel + '</button>' +
          '<button data-dl title="Descargar imagen" style="background:#131d30;border:1px solid #334155;color:#cbd5e1;border-radius:12px;padding:0 15px;font-weight:800;font-size:16px;cursor:pointer">⬇️</button>' +
        '</div>' +
      '</div>';

    // difiere el revoke del objectURL: revocarlo en el mismo tick que a.click() puede ABORTAR la descarga.
    const close = () => { const u = url; ov.remove(); setTimeout(() => { try { URL.revokeObjectURL(u); } catch(_) {} }, 1800); };
    ov.addEventListener('click', e => { if (e.target === ov) close(); });
    ov.querySelector('[data-x]').onclick = close;
    ov.querySelector('[data-dl]').onclick = () => { _descargar(url, filename); _toast('Imagen descargada', 'ok', 2500); };
    ov.querySelector('[data-send]').onclick = async () => {
      // Este toque es una activación de usuario FRESCA → móvil: share confiable.
      if (puedeArchivo) {
        try { await navigator.share({ files: [file], text: texto }); close(); return; }
        catch(e) { if (e && e.name === 'AbortError') return; /* canceló → seguir en la vista previa */ }
      }
      // PC / WhatsApp Web: la imagen no se adjunta por web. Copiar al portapapeles (pegar con Ctrl+V) + descargar
      // + abrir WhatsApp. ORDEN CLAVE: INICIAR clipboard.write (necesita FOCO + activación) ANTES de window.open
      // (que roba el foco); luego abrir/descargar y recién AWAIT la copia. Así ni la copia ni el open se bloquean.
      let clipP = null;
      try { if (navigator.clipboard && window.ClipboardItem) clipP = navigator.clipboard.write([ new ClipboardItem({ [blob.type || 'image/png']: blob }) ]); } catch(_) { clipP = null; }
      window.open('https://wa.me/?text=' + encodeURIComponent(texto || ''), '_blank');
      _descargar(url, filename);
      let copiado = false;
      if (clipP) { try { await clipP; copiado = true; } catch(_) { copiado = false; } }
      close();
      _toast(copiado
        ? '✅ Imagen copiada — pégala con Ctrl+V en el chat de WhatsApp (también se descargó)'
        : '⬇️ Imagen descargada — arrástrala al chat de WhatsApp', 'info', 6500);
    };
    document.body.appendChild(ov);
  };
})();
