// ============================================================
// warehouseMos — js/sharesheet.js
// Compartir una imagen por WhatsApp DIRECTO (sin modal de preview — el modal era un paso extra
// que hacía lento el proceso):
//  • Móvil: hoja nativa de compartir con la imagen adjunta (navigator.share con archivo).
//  • PC / WhatsApp Web (no adjunta archivos por web): copia la imagen al portapapeles (pegar con
//    Ctrl+V), la descarga y abre WhatsApp con el texto.
// window.ShareSheet(blob, filename, texto)  — se llama desde voucher.js y cargadores.js.
// ============================================================
(function() {
  'use strict';
  function _descargar(url, filename) {
    try { const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); } catch(_) {}
  }
  function _toast(msg, tipo, ms) { try { if (typeof window.toast === 'function') window.toast(msg, tipo || 'info', ms || 5000); } catch(_) {} }

  window.ShareSheet = async function(blob, filename, texto) {
    if (!blob) { window.open('https://wa.me/?text=' + encodeURIComponent(texto || ''), '_blank'); return; }
    let file = null; try { file = new File([blob], filename, { type: 'image/png' }); } catch(_) {}
    // MÓVIL: compartir la imagen directamente (hoja nativa → el usuario elige WhatsApp).
    if (file && navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], text: texto }); return; }
      catch(e) { if (e && e.name === 'AbortError') return; /* canceló; si falló por otra razón → plan PC */ }
    }
    // PC / WhatsApp Web: iniciar la copia (necesita FOCO) ANTES de abrir WhatsApp (que roba el foco);
    // luego descargar y recién await la copia.
    let clipP = null;
    try { if (navigator.clipboard && window.ClipboardItem) clipP = navigator.clipboard.write([ new ClipboardItem({ [blob.type || 'image/png']: blob }) ]); } catch(_) { clipP = null; }
    window.open('https://wa.me/?text=' + encodeURIComponent(texto || ''), '_blank');
    const url = URL.createObjectURL(blob);
    _descargar(url, filename);
    setTimeout(() => { try { URL.revokeObjectURL(url); } catch(_) {} }, 1800);
    let copiado = false;
    if (clipP) { try { await clipP; copiado = true; } catch(_) { copiado = false; } }
    _toast(copiado
      ? '✅ Imagen copiada — pégala con Ctrl+V en el chat de WhatsApp (también se descargó)'
      : '⬇️ Imagen descargada — adjúntala en WhatsApp', 'info', 5500);
  };
})();
