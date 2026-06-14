// impresion-directa.js — [PASO 5 · B5] Armado de tickets/etiquetas en el NAVEGADOR (réplica fiel del GAS).
// Cada `armarXxx(params)` devuelve { printerHint, title, base64 } SIN resolver el printerId real ni imprimir:
// el caller resuelve el printerId (por printerHint 'TICKET'/'ADHESIVO') y manda con API.imprimirDirecto(id, base64, title).
// ⚠️ VALIDACIÓN: el armado es determinístico pero SOLO se valida de verdad IMPRIMIENDO en la impresora física.
// Antes de confiar en cualquier tipo, imprimir 1 de prueba y comparar contra la versión GAS.
//
// Estado de portación:
//   ✅ Bienvenida turno (ESC/POS) — Code.gs:786-853
//   ⏳ Etiqueta Caserito/envasado (TSPL2) — Envasados.gs (bitmap logo + highlight + wrap + drift) — PENDIENTE
//   ⏳ Aviso a cajeros (ESC/POS + QR) — Reporte.gs — PENDIENTE
//   ⏳ Membrete estándar (ESC/POS) — Productos.gs — PENDIENTE
//   ⏳ Membrete ME/WH (TSPL2) — Membretes.gs — PENDIENTE
const ImpresionDirecta = (() => {

  // base64 de los BYTES UTF-8 de un string — idéntico a Utilities.base64Encode(string) de GAS.
  // Maneja bytes de control (\x1b \x1d, <128) y multibyte (acentos) igual que el backend.
  function _b64Utf8(str) {
    return btoa(unescape(encodeURIComponent(String(str))));
  }

  // Fecha/hora en America/Lima (espeja Session.getScriptTimeZone() del GAS, que corre en TZ Perú).
  function _fmtLima(withDate) {
    // new Date() en el navegador del usuario (no es el sandbox de build) — TZ del dispositivo, reformateada a Lima.
    const opts = { timeZone: 'America/Lima', hour12: false,
      year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
    const p = new Intl.DateTimeFormat('es-PE', opts).formatToParts(new Date())
      .reduce((a, x) => (a[x.type] = x.value, a), {});
    return withDate ? `${p.day}/${p.month}/${p.year} ${p.hour}:${p.minute}:${p.second}`
                    : `${p.hour}:${p.minute}:${p.second}`;
  }

  // ── BIENVENIDA DE TURNO — réplica EXACTA de imprimirBienvenida (Code.gs:786-853) ──
  // params: { nombre, apellido, rol, horaInicio?, empresa? }. empresa default 'InversionMos' (= _getConfigValue del GAS).
  function armarBienvenida(params) {
    params = params || {};
    const empresa  = String(params.empresa || 'InversionMos');
    const nombre   = String(params.nombre   || '');
    const apellido = String(params.apellido || '');
    const rol      = String(params.rol      || '');
    const hora     = String(params.horaInicio || _fmtLima(false));
    const ahora    = _fmtLima(true);

    let t = '';
    t += '\x1b\x40';                          // Init impresora
    t += '\x1b\x61\x01';                      // Centrar
    t += '\x1b\x21\x30';                      // Doble alto + ancho
    t += empresa + '\n';
    t += '\x1b\x21\x00';                      // Normal
    t += 'warehouseMos — Almácen\n';// (texto literal del GAS, se conserva igual)
    t += '================================\n\n';
    t += '\x1b\x21\x10';                      // Doble alto
    t += 'INICIO DE TURNO\n';
    t += '\x1b\x21\x00\n';
    t += '\x1b\x61\x01';
    t += '\x1b\x21\x20';                      // Doble ancho
    t += nombre + ' ' + apellido + '\n';
    t += '\x1b\x21\x00';
    t += rol + '\n\n';
    t += '================================\n';
    t += '\x1b\x61\x00';                      // Izquierda
    t += 'Hora inicio : ' + hora + '\n';
    t += 'Impreso     : ' + ahora + '\n';
    t += '\n';
    t += '\x1b\x61\x01';
    t += 'Bienvenido al turno. ¡Mucho éxito!\n';
    t += '\n\n\n\n\n';
    t += '\x1d\x56\x00';                      // Corte

    return { printerHint: 'TICKET', title: 'Bienvenida ' + nombre + ' ' + apellido, base64: _b64Utf8(t) };
  }

  return {
    _b64Utf8,        // expuesto para los builders que arman ARRAY de bytes (TSPL2) cuando se porten
    armarBienvenida,
  };
})();
