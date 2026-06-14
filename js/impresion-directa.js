// impresion-directa.js — [PASO 5 · B5] Armado de tickets/etiquetas en el NAVEGADOR (réplica fiel del GAS).
// Cada `armarXxx(params)` devuelve { printerHint, title, base64 } SIN resolver el printerId real ni imprimir:
// el caller resuelve el printerId (por printerHint 'TICKET'/'ADHESIVO') y manda con API.imprimirDirecto(id, base64, title).
// ⚠️ VALIDACIÓN: el armado es determinístico pero SOLO se valida de verdad IMPRIMIENDO en la impresora física.
// Antes de confiar en cualquier tipo, imprimir 1 de prueba y comparar contra la versión GAS.
//
// Estado de portación:
//   ✅ Bienvenida turno (ESC/POS) — Code.gs:786-853
//   ✅ Etiqueta Caserito/envasado (TSPL2) — Envasados.gs (bitmap logo + highlight + wrap; SIN drift acumulado, offset base)
//   ✅ Aviso a cajeros (ESC/POS + QR + modo comparativo) — Reporte.gs:1316-1505 (normal), 1190-1312 (comparativo)
//   ✅ Membrete estándar (ESC/POS) — Productos.gs:2214-2395
//   ✅ Membrete ME góndola (TSPL2) — Membretes.gs:120-240 (precio MEGA Font 5; offset base)
//   ✅ Membrete WH almacén (TSPL2) — Membretes.gs:262-368 (logo + highlight; offset base)
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

  // base64 de un ARRAY de bytes (0-255) — idéntico a Utilities.base64Encode(blob.getBytes()) de GAS.
  // Los builders ESC/POS del aviso/membrete arman un array `B` de enteros byte por byte (b1/bStr/bLn)
  // y en GAS lo pasan por Utilities.newBlob(B).getBytes() → base64. Acá replicamos esa conversión:
  // array de bytes → binary string (cada char = 1 byte) → btoa. NO usar _b64Utf8 (re-encodearía
  // los bytes >127 como UTF-8 multibyte y rompería el barcode/QR).
  function _bytesB64(arr) {
    let bin = '';
    for (let i = 0; i < arr.length; i++) bin += String.fromCharCode(arr[i] & 0xff);
    return btoa(bin);
  }

  // ── Helpers de armado (espejo EXACTO de Reporte.gs / Productos.gs) ──

  // Word-wrap inteligente que corta palabras largas como último recurso — _wrapPalabras (Reporte.gs:1510-1539).
  function _wrapPalabras(texto, anchoPrimero, anchoResto) {
    texto = String(texto || '').trim();
    if (!texto) return [''];
    if (anchoResto == null) anchoResto = anchoPrimero;

    var palabras = texto.split(/\s+/);
    var lineas   = [];
    var cur      = '';
    var ancho    = anchoPrimero;

    for (var i = 0; i < palabras.length; i++) {
      var p = palabras[i];
      while (p.length > ancho) {
        if (cur) { lineas.push(cur); cur = ''; ancho = anchoResto; }
        lineas.push(p.substring(0, ancho));
        p = p.substring(ancho);
      }
      var sep = cur ? ' ' : '';
      if ((cur + sep + p).length <= ancho) {
        cur = cur + sep + p;
      } else {
        lineas.push(cur);
        cur = p;
        ancho = anchoResto;
      }
    }
    if (cur) lineas.push(cur);
    return lineas;
  }

  // Word-wrap SIN cortar palabras largas — _membWrap (Productos.gs:2378-2395). Distinto de _wrapPalabras.
  function _membWrap(text, maxLen) {
    var words = String(text || '').trim().split(/\s+/);
    var lines = [];
    var cur   = '';
    for (var i = 0; i < words.length; i++) {
      var w = words[i];
      if (!cur) {
        cur = w;
      } else if ((cur + ' ' + w).length <= maxLen) {
        cur += ' ' + w;
      } else {
        lines.push(cur);
        cur = w;
      }
    }
    if (cur) lines.push(cur);
    return lines;
  }

  // Padding izquierda↔derecha a 48 chars — _padLine48 (Reporte.gs:181-187).
  function _padLine48(left, right) {
    var l = String(left || '');
    var r = String(right || '');
    var pad = 48 - l.length - r.length;
    if (pad < 1) pad = 1;
    return l + Array(pad + 1).join(' ') + r;
  }

  // Etiquetas legibles para tags — _lblTag (Reporte.gs:1182-1187).
  function _lblTag(g, v) {
    if (!v) return '(no marcado)';
    if (g === 'comp')  return v === 'si' ? 'Con comprobante' : 'Sin comprobante';
    if (g === 'compl') return v === 'si' ? 'Pedido completo' : 'Pedido INCOMPLETO';
    return String(v);
  }

  // Parse del comentario libre + tags comprobante/completo — _parseComentarioPI (Reporte.gs:1140-1156).
  // Espejo EXACTO del backend (que a su vez espeja _tagsFromComentario del frontend).
  function _parseComentarioPI(c) {
    var s = String(c || '');
    var tags = { comp: null, compl: null };
    if (/comprobante:\s*s[ií]\b/i.test(s))    tags.comp  = 'si';
    else if (/comprobante:\s*no\b/i.test(s))  tags.comp  = 'no';
    if (/completo:\s*s[ií]\b/i.test(s))       tags.compl = 'si';
    else if (/completo:\s*no\b/i.test(s))     tags.compl = 'no';
    var libre = s
      .replace(/Comprobante:\s*(?:S[ií]|No)\s*\|?\s*/gi, '')
      .replace(/Completo:\s*(?:S[ií]|No)\s*\|?\s*/gi, '')
      .replace(/^\s*\|\s*/, '').replace(/\s*\|\s*$/, '')
      .trim();
    return { tagComp: tags.comp, tagCompl: tags.compl, comentarioLibre: libre };
  }

  // Snapshot de los 5 campos críticos desde un objeto preingreso — _snapshotAvisoFromPI (Reporte.gs:1159-1168).
  function _snapshotAvisoFromPI(pi) {
    var p = _parseComentarioPI(pi.comentario);
    return {
      idProveedor:     String(pi.idProveedor || ''),
      monto:           parseFloat(pi.monto) || 0,
      tagComp:         p.tagComp,
      tagCompl:        p.tagCompl,
      comentarioLibre: p.comentarioLibre
    };
  }

  // Claves que difieren entre dos snapshots — _diffSnapshotsAviso (Reporte.gs:1171-1179).
  function _diffSnapshotsAviso(a, b) {
    if (!a || !b) return [];
    var keys = ['idProveedor','monto','tagComp','tagCompl','comentarioLibre'];
    return keys.filter(function(k) {
      var va = a[k]; var vb = b[k];
      if (k === 'monto') return (parseFloat(va) || 0) !== (parseFloat(vb) || 0);
      return String(va == null ? '' : va) !== String(vb == null ? '' : vb);
    });
  }

  // Fecha estilo "2 mayo 1:11pm" en TZ Lima — espeja el bloque Utilities.formatDate
  // de _construirAvisoIngresoBytes (Reporte.gs:1353-1371). Devuelve '' si la fecha es inválida.
  function _fechaAvisoLima(rawFecha) {
    try {
      var d = (rawFecha instanceof Date) ? rawFecha : new Date(rawFecha || new Date());
      if (isNaN(d.getTime())) return '';
      var meses = ['enero','febrero','marzo','abril','mayo','junio',
                   'julio','agosto','septiembre','octubre','noviembre','diciembre'];
      var opts = { timeZone: 'America/Lima', hour12: false,
        month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' };
      var p = new Intl.DateTimeFormat('es-PE', opts).formatToParts(d)
        .reduce(function(a, x) { a[x.type] = x.value; return a; }, {});
      var dia    = parseInt(p.day, 10);
      var mesIdx = parseInt(p.month, 10) - 1;
      var hh     = parseInt(p.hour, 10);
      var mm     = parseInt(p.minute, 10);
      var mes    = meses[mesIdx] || '';
      var ampm   = hh >= 12 ? 'pm' : 'am';
      var hh12   = hh % 12; if (hh12 === 0) hh12 = 12;
      var horaTxt = mm === 0 ? (hh12 + ampm) : (hh12 + ':' + (mm < 10 ? '0' : '') + mm + ampm);
      return dia + ' ' + mes + ' ' + horaTxt;
    } catch(e) { return ''; }
  }

  // ── AVISO A CAJEROS (preingreso) — réplica de imprimirAvisoCajeros (Reporte.gs:967) ──
  // El GAS resuelve el preingreso/proveedor/cajas desde Sheets; ACÁ todo llega resuelto en params.
  // params (modo normal):
  //   { idPreingreso, proveedorNombre, monto, cargadores, comentario, fotos, fecha, reporteUrl }
  //   - monto: número o string (se imprime tal cual con prefijo 'S/. ', igual que el GAS usa pi.monto).
  //   - cargadores: array de objetos { nombre|idPersonal, carretas, estados:[] } o array de strings,
  //                 O un string JSON (se parsea con JSON.parse igual que el GAS hace con pi.cargadores).
  //   - comentario: string (incluye tags "Comprobante:.. | Completo:.. | <libre>"). Se imprime COMPLETO.
  //   - fotos: string CSV de ids (solo se cuenta cuántas hay, como el GAS).
  //   - fecha: Date o string parseable (se formatea a "2 mayo 1:11pm" en TZ Lima).
  //   - reporteUrl: string para el QR (vacío => sin QR).
  // params (modo comparativo): además
  //   { modoComparativo:true, snapshotAnterior:{ idProveedor, monto, tagComp, tagCompl, comentarioLibre },
  //     proveedorNombreAnterior } y el snapshot ACTUAL se deriva de proveedorNombre+monto+comentario.
  // Devuelve { printerHint:'TICKET', title, base64 }.
  function armarAvisoCajeros(params) {
    params = params || {};
    var provName   = String(params.proveedorNombre || '');
    var reporteUrl = String(params.reporteUrl || '');
    // `pi` equivalente al row de hoja que consume el GAS.
    var pi = {
      idPreingreso: String(params.idPreingreso || ''),
      idProveedor:  String(params.idProveedor || ''),
      monto:        params.monto,
      comentario:   params.comentario,
      cargadores:   params.cargadores,
      fotos:        params.fotos,
      fecha:        params.fecha,
      fechaCreacion: params.fechaCreacion
    };

    var modoComparativo = !!params.modoComparativo;
    var snapAnt = null;
    if (modoComparativo) {
      try {
        snapAnt = (typeof params.snapshotAnterior === 'string')
          ? JSON.parse(params.snapshotAnterior || '{}')
          : (params.snapshotAnterior || {});
      } catch(e) { snapAnt = null; }
      if (!snapAnt || !Object.keys(snapAnt).length) modoComparativo = false;
    }
    var snapAct = _snapshotAvisoFromPI(pi);
    if (modoComparativo) {
      var difsChk = _diffSnapshotsAviso(snapAnt, snapAct);
      if (!difsChk.length) {
        // El GAS retorna { ok:false, error:'NO_CHANGES' }; acá conservamos esa señal sin imprimir.
        return { ok: false, error: 'NO_CHANGES', mensaje: 'No hay cambios en campos críticos' };
      }
    }

    var bytes;
    if (modoComparativo) {
      var provNameAnt = '';
      if (snapAnt && snapAnt.idProveedor && snapAnt.idProveedor !== pi.idProveedor) {
        provNameAnt = String(params.proveedorNombreAnterior || '');
      }
      bytes = _construirAvisoComparativoBytes(pi, provName, provNameAnt, reporteUrl, snapAnt, snapAct);
    } else {
      bytes = _construirAvisoIngresoBytes(pi, provName, reporteUrl);
    }

    return {
      printerHint: 'TICKET',
      title: 'Aviso preingreso ' + (pi.idPreingreso || ''),
      base64: _bytesB64(bytes)
    };
  }

  // Bytes ESC/POS del ticket de aviso de ingreso — _construirAvisoIngresoBytes (Reporte.gs:1316-1505).
  function _construirAvisoIngresoBytes(pi, provName, reporteUrl) {
    var B = [];
    function b1(v)   { B.push(v & 0xff); }
    function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
    function bLn(s)  { bStr(s); b1(0x0a); }

    var SEP  = '================================================';
    var SEP2 = '------------------------------------------------';

    b1(0x1b); b1(0x40);

    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x21); b1(0x38);
    bLn('WAREHOUSE');
    bLn('MOS');
    b1(0x1b); b1(0x21); b1(0x00);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('AVISO INGRESO');
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);

    if (provName) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
      var nameLines = _wrapPalabras(String(provName).toUpperCase(), 24);
      for (var nl = 0; nl < nameLines.length && nl < 2; nl++) bLn(nameLines[nl]);
      b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x61); b1(0x00);
    }

    var fechaPI = _fechaAvisoLima(pi.fecha || pi.fechaCreacion);
    if (fechaPI) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn(fechaPI);
      b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x61); b1(0x00);
    }

    bLn(SEP);

    if (pi.monto) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('PREPARAR PARA PAGAR');
      b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x21); b1(0x38); b1(0x1b); b1(0x45); b1(0x01);
      bLn('S/. ' + pi.monto);
      b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x61); b1(0x00);
      bLn(SEP);
    }

    var cargs = [];
    if (Array.isArray(pi.cargadores)) cargs = pi.cargadores;
    else { try { cargs = JSON.parse(pi.cargadores || '[]'); } catch(e) {} }
    if (cargs.length) {
      bLn('');
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('Cargadores:');
      b1(0x1b); b1(0x45); b1(0x00);
      var totLlenas = 0, totMedias = 0, totVacias = 0, totCarretas = 0;
      cargs.forEach(function(c) {
        var nombre = (typeof c === 'object') ? (c.nombre || c.idPersonal || '') : String(c);
        if (!nombre) return;
        var carretas = (typeof c === 'object' && c.carretas) ? parseInt(c.carretas) || 0 : 0;
        var estadosArr = [];
        if (typeof c === 'object' && Array.isArray(c.estados)) {
          estadosArr = c.estados.slice(0, carretas).map(function(e) {
            return (e === 'MEDIA' || e === 'VACIA') ? e : 'LLENA';
          });
        }
        while (estadosArr.length < carretas) estadosArr.push('LLENA');
        var ll = 0, md = 0, vc = 0;
        estadosArr.forEach(function(e) {
          if (e === 'LLENA') ll++; else if (e === 'MEDIA') md++; else if (e === 'VACIA') vc++;
        });
        totLlenas += ll; totMedias += md; totVacias += vc; totCarretas += carretas;
        if (carretas > 0) {
          b1(0x1b); b1(0x45); b1(0x01);
          bLn(_padLine48('  - ' + nombre, carretas + ' carreta' + (carretas === 1 ? '' : 's')));
          b1(0x1b); b1(0x45); b1(0x00);
          var dets = [];
          if (ll > 0) dets.push(ll + ' LLENA' + (ll === 1 ? '' : 'S'));
          if (md > 0) dets.push(md + ' MEDIA' + (md === 1 ? '' : 'S'));
          if (vc > 0) dets.push(vc + ' CASI VACIA' + (vc === 1 ? '' : 'S'));
          if (dets.length) bLn('      ' + dets.join(' / '));
        } else {
          bLn('  - ' + nombre);
        }
      });
      if (cargs.length > 1 && (totMedias + totVacias > 0)) {
        bLn(SEP2);
        b1(0x1b); b1(0x45); b1(0x01);
        bLn(_padLine48('  TOTAL', totCarretas + ' carreta' + (totCarretas === 1 ? '' : 's')));
        var tdets = [];
        if (totLlenas > 0) tdets.push(totLlenas + ' L');
        if (totMedias > 0) tdets.push(totMedias + ' M');
        if (totVacias > 0) tdets.push(totVacias + ' CV');
        if (tdets.length) bLn('      ' + tdets.join(' / '));
        b1(0x1b); b1(0x45); b1(0x00);
      }
    }

    if (pi.comentario) {
      bLn(SEP2);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('Comentario:');
      b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
      var comLines = _wrapPalabras(String(pi.comentario), 24);
      for (var ci = 0; ci < comLines.length; ci++) bLn(comLines[ci]);
      b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
    }

    bLn(SEP);

    var nFotos = pi.fotos ? String(pi.fotos).split(',').filter(Boolean).length : 0;
    if (nFotos > 0) {
      b1(0x1b); b1(0x61); b1(0x01);
      bLn(nFotos + ' imagen' + (nFotos !== 1 ? 'es' : '') + ' adjunta' + (nFotos !== 1 ? 's' : ''));
      bLn('ver en el reporte digital');
      b1(0x1b); b1(0x61); b1(0x00);
      bLn(SEP);
    }

    if (reporteUrl) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('VER PREINGRESO COMPLETO');
      b1(0x1b); b1(0x45); b1(0x00);

      var qrLen = reporteUrl.length + 3;
      var qrpL  = qrLen & 0xff;
      var qrpH  = (qrLen >> 8) & 0xff;
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x04); b1(0x00); b1(0x31); b1(0x41); b1(0x32); b1(0x00);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x43); b1(0x05);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x45); b1(0x31);
      b1(0x1d); b1(0x28); b1(0x6b); b1(qrpL); b1(qrpH); b1(0x31); b1(0x50); b1(0x30);
      bStr(reporteUrl);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x51); b1(0x30);

      bLn('Escanea para fotos y detalles');
      b1(0x1b); b1(0x61); b1(0x00);
      bLn(SEP);
    }

    b1(0x1b); b1(0x4a); b1(160);
    b1(0x1d); b1(0x56); b1(0x00);

    return B;
  }

  // Bytes ESC/POS del ticket COMPARATIVO — _construirAvisoComparativoBytes (Reporte.gs:1190-1312).
  function _construirAvisoComparativoBytes(pi, provNameActual, provNameAnterior, reporteUrl, snapAnt, snapAct) {
    var B = [];
    function b1(v)   { B.push(v & 0xff); }
    function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
    function bLn(s)  { bStr(s); b1(0x0a); }

    var SEP  = '================================================';
    var SEP2 = '------------------------------------------------';

    b1(0x1b); b1(0x40);

    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x21); b1(0x38);
    bLn('WAREHOUSE');
    bLn('MOS');
    b1(0x1b); b1(0x21); b1(0x00);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('*** AVISO ACTUALIZADO ***');
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);

    if (provNameActual) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x21); b1(0x10); b1(0x1b); b1(0x45); b1(0x01);
      var nL = _wrapPalabras(String(provNameActual).toUpperCase(), 24);
      for (var nl = 0; nl < nL.length && nl < 2; nl++) bLn(nL[nl]);
      b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x61); b1(0x00);
    }
    b1(0x1b); b1(0x61); b1(0x01);
    bLn('Preingreso ' + (pi.idPreingreso || ''));
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);

    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x45); b1(0x01);
    bLn('-- DATOS CORREGIDOS --');
    b1(0x1b); b1(0x45); b1(0x00);
    b1(0x1b); b1(0x61); b1(0x00);
    bLn(SEP);

    var difs = _diffSnapshotsAviso(snapAnt, snapAct);

    function _seccion(titulo, antes, ahora) {
      b1(0x1b); b1(0x45); b1(0x01);
      bLn(' ' + titulo);
      b1(0x1b); b1(0x45); b1(0x00);
      bLn('   ANTES:  ' + (antes == null || antes === '' ? '(vacio)' : String(antes)));
      bLn('   AHORA:  ' + (ahora == null || ahora === '' ? '(vacio)' : String(ahora)));
      bLn(SEP2);
    }

    difs.forEach(function(k) {
      if (k === 'idProveedor') {
        _seccion('PROVEEDOR',
          provNameAnterior || snapAnt.idProveedor || '(sin proveedor)',
          provNameActual   || snapAct.idProveedor || '(sin proveedor)');
      } else if (k === 'monto') {
        _seccion('MONTO',
          'S/. ' + (parseFloat(snapAnt.monto) || 0).toFixed(2),
          'S/. ' + (parseFloat(snapAct.monto) || 0).toFixed(2));
      } else if (k === 'tagComp') {
        _seccion('COMPROBANTE', _lblTag('comp', snapAnt.tagComp), _lblTag('comp', snapAct.tagComp));
      } else if (k === 'tagCompl') {
        _seccion('COMPLETO', _lblTag('compl', snapAnt.tagCompl), _lblTag('compl', snapAct.tagCompl));
      } else if (k === 'comentarioLibre') {
        b1(0x1b); b1(0x45); b1(0x01);
        bLn(' COMENTARIO');
        b1(0x1b); b1(0x45); b1(0x00);
        var antL = _wrapPalabras(snapAnt.comentarioLibre || '(vacio)', 40);
        var actL = _wrapPalabras(snapAct.comentarioLibre || '(vacio)', 40);
        bLn('   ANTES:');
        antL.forEach(function(l) { bLn('     ' + l); });
        bLn('   AHORA:');
        actL.forEach(function(l) { bLn('     ' + l); });
        bLn(SEP2);
      }
    });

    bLn(SEP);

    if (difs.indexOf('monto') >= 0 && (parseFloat(snapAct.monto) || 0) > 0) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('PREPARAR PARA PAGAR');
      b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x21); b1(0x38); b1(0x1b); b1(0x45); b1(0x01);
      bLn('S/. ' + (parseFloat(snapAct.monto) || 0).toFixed(2));
      b1(0x1b); b1(0x21); b1(0x00); b1(0x1b); b1(0x45); b1(0x00);
      b1(0x1b); b1(0x61); b1(0x00);
      bLn(SEP);
    }

    if (reporteUrl) {
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x45); b1(0x01);
      bLn('VER PREINGRESO COMPLETO');
      b1(0x1b); b1(0x45); b1(0x00);
      var qrLen = reporteUrl.length + 3;
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x04); b1(0x00); b1(0x31); b1(0x41); b1(0x32); b1(0x00);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x43); b1(0x05);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x45); b1(0x31);
      b1(0x1d); b1(0x28); b1(0x6b); b1(qrLen & 0xff); b1((qrLen >> 8) & 0xff); b1(0x31); b1(0x50); b1(0x30);
      bStr(reporteUrl);
      b1(0x1d); b1(0x28); b1(0x6b); b1(0x03); b1(0x00); b1(0x31); b1(0x51); b1(0x30);
      bLn('Escanea para ver fotos y detalles');
      b1(0x1b); b1(0x61); b1(0x00);
      bLn(SEP);
    }

    b1(0x1b); b1(0x4a); b1(160);
    b1(0x1d); b1(0x56); b1(0x00);

    return B;
  }

  // ── MEMBRETE ESTÁNDAR (SKU + EAN) — réplica de imprimirMembrete (Productos.gs:2214-2375) ──
  // El GAS resuelve producto/EAN desde Sheets; ACÁ todo llega resuelto en params.
  // params: { nombre, sku, barcodes }
  //   - nombre: descripción del producto (fallback a sku si vacío, igual que prod.descripcion || sku).
  //   - sku: skuBase.
  //   - barcodes: array de strings de EAN ya resueltos (o string JSON). Si vacío => usa [sku] (mismo
  //               fallback final del GAS: `if (!allEan.length) allEan.push(sku)`).
  // Devuelve { printerHint:'TICKET', title, base64 }.
  function armarMembreteStd(params) {
    params = params || {};
    var sku = String(params.sku || '');

    var allEan = [];
    if (Array.isArray(params.barcodes)) {
      allEan = params.barcodes.map(String).filter(Boolean);
    } else if (params.barcodes) {
      try {
        var parsed = JSON.parse(String(params.barcodes));
        if (Array.isArray(parsed)) allEan = parsed.map(String).filter(Boolean);
      } catch(e) {}
    }
    if (!allEan.length) allEan.push(sku);

    var numEan      = allEan.length;
    var hasMultiple = numEan > 1;

    var desc      = String(params.nombre || sku).toUpperCase();
    var nameLines = _membWrap(desc, 20).slice(0, 2);

    var B = [];
    function b1(v)   { B.push(v & 0xff); }
    function bStr(s) { for (var k = 0; k < s.length; k++) B.push(s.charCodeAt(k) & 0xff); }
    function bLn(s)  { bStr(s); b1(0x0a); }

    var SEPEQ  = '================================================';
    var SEPDA  = '------------------------------------------------';

    b1(0x1b); b1(0x40);

    b1(0x1b); b1(0x61); b1(0x01);
    b1(0x1b); b1(0x21); b1(0x38);
    for (var ni = 0; ni < nameLines.length; ni++) { bLn(nameLines[ni]); }
    b1(0x1b); b1(0x21); b1(0x00);

    b1(0x1b); b1(0x21); b1(0x08);
    bLn('SKU: ' + sku);
    b1(0x1b); b1(0x21); b1(0x00);

    b1(0x1b); b1(0x61); b1(0x01);
    bLn(SEPEQ);

    for (var ei = 0; ei < numEan; ei++) {
      var bd = '{B' + allEan[ei];
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1d); b1(0x68); b1(80);
      b1(0x1d); b1(0x77); b1(0x02);
      b1(0x1d); b1(0x48); b1(0x02);
      b1(0x1d); b1(0x66); b1(0x00);
      b1(0x1d); b1(0x6b); b1(0x49);
      b1(bd.length & 0xff);
      bStr(bd);
      b1(0x0a);
      b1(0x1b); b1(0x61); b1(0x01);
      var isLast = (ei === numEan - 1);
      bLn(isLast ? SEPEQ : SEPDA);
    }

    if (hasMultiple) {
      var skuBd = '{B' + sku;
      b1(0x1b); b1(0x61); b1(0x01);
      b1(0x1b); b1(0x21); b1(0x08);
      bLn('SKU: ' + sku);
      b1(0x1b); b1(0x21); b1(0x00);
      b1(0x1d); b1(0x68); b1(60);
      b1(0x1d); b1(0x77); b1(0x02);
      b1(0x1d); b1(0x48); b1(0x02);
      b1(0x1d); b1(0x66); b1(0x00);
      b1(0x1d); b1(0x6b); b1(0x49);
      b1(skuBd.length & 0xff);
      bStr(skuBd);
      b1(0x0a);
      b1(0x1b); b1(0x61); b1(0x01);
      bLn(SEPEQ);
    }

    b1(0x1b); b1(0x4a); b1(160);
    b1(0x1d); b1(0x56); b1(0x00);

    return {
      printerHint: 'TICKET',
      title: 'Membrete ' + sku,
      base64: _bytesB64(B)
    };
  }

  // ════════════════════════════════════════════════════════════════════
  // TSPL2 (impresora TSC TTP-244CE 203 DPI, adhesivos 50×25mm = 400×200 dots)
  // Réplica FIEL del GAS (Envasados.gs + Membretes.gs). Comparten helpers de
  // texto/barcode/logo. Cada builder devuelve { printerHint:'ADHESIVO', title, base64 }.
  //
  // ENCODING: el GAS arma un ARRAY de bytes (0-255) — _strToBytesEtq concatena
  // `charCodeAt(i) & 0xFF` de los comandos TSPL ASCII; _hexToBytesEtq vuelca los
  // bytes CRUDOS del bitmap del logo (hex→byte) en medio del array. Luego GAS hace
  // Utilities.base64Encode(bytes) sobre ese array de enteros. Acá replicamos byte a
  // byte con _bytesB64(arr) (array de bytes → binary string → btoa). NUNCA _b64Utf8:
  // re-encodearía como UTF-8 multibyte los bytes >127 del bitmap → logo corrupto.
  //
  // DRIFT: el GAS compensa drift acumulado por-print leyendo Script Properties del
  // rollo (ADHESIVO_OFFSET_Y, ADHESIVO_DRIFT_DOTS_POR_PRINT, ADHESIVO_PRINTS_DESDE_CAL).
  // Esos Properties NO existen en el navegador → acá usamos offsetY = OFFSET BASE = 0,
  // que es exactamente lo que daría el GAS con drift=0 (offsetBase 0 + comp 0).
  // TODO(drift): si en el futuro se quiere compensar drift en cliente, habría que
  // exponer esos 3 valores como params; por ahora se usa el base (drift=0).
  // ════════════════════════════════════════════════════════════════════

  // Defaults del GAS para los Script Properties del adhesivo (PropertiesService).
  // Envasados.gs:493-495 / Membretes.gs:128-130 — || 2 / || 8 / || 4.
  var ADHESIVO_GAP_MM   = 2;   // ADHESIVO_GAP_MM   default
  var ADHESIVO_DENSITY  = 8;   // ADHESIVO_DENSITY  default
  var ADHESIVO_SPEED    = 4;   // ADHESIVO_SPEED    default
  // offsetY base (drift=0): GAS = offsetBase(0) + Math.round(driftDots(0) * prints) = 0.
  var ADHESIVO_OFFSET_Y_BASE = 0;

  // ── Logo bitmap (184x36 dots = 23 bytes/row × 36) — Envasados.gs:313-337 ──
  // LITERAL, copiado byte-a-byte (NO regenerado). Fuente: logo-tonys.svg / gen.py.
  var LOGO_W_BYTES = 23;
  var LOGO_H = 36;
  var LOGO_TSPL_HEX =
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE3FFFFC000' +
'1FE03FE0780C078083F80FFFFFFFFFFFFFFFC1FFFFC0001F800FE0380C078083C001FFFFFFFF' +
'FFFFFF007FFFC0001E0007E0380C0701838000FFFFFFFFFFFFFE003FFFC0001E0003E0380E03' +
'018380007FFFFFFFFFFFFC001FFFC0001C0001E0180E03018300007FFFFFFFFFFFF8000FFFC0' +
'001C0201E0180E03018300C07FFFFFFFFFFFE00003FFFE01FC0701E0180F03038301C07FFFFF' +
'FFFFFFC00001FFFE01F80701E0080F03038301C07FFFFFFFFFFF800000FFFE01F80701E0080F' +
'0203C301C07FFFFFFFFFFF0000007FFE01F80701E0080F8007FF00FFFFFFFFFFFFFC0000001F' +
'FE01F80701E0080F8007FF007FFFFFFFFFFFF000000007FE01F80701E0000F8007FF001FFFFF' +
'FFFFFFF000000007FE01F80701E0000F800FFF800FFFFFFFFFFFF000000007FE01F80701E000' +
'0FC00FFFC003FFFFFFFFFFF000000007FE01F80701E0000FC00FFFE001FFFFFFFFFFF0000000' +
'07FE01F80701E0000FC00FFFF000FFFFFFFFFFFE0000003FFE01F80701E0000FE01FFFFC007F' +
'FFFFFFFFFE0000003FFE01F80701E0000FE01FFFFE007FFFFFFFFFFE3E003E3FFE01F80701E0' +
'400FE01FFFFF807FFFFFFFFFFE3E003E3FFE01F80701E0400FE01FFF01C03FFFFFFFFFFE3E00' +
'3E3FFE01F80701E0400FE01FFF01C03FFFFFFFFFFE3EFFBE3FFE01F80701E0600FE01FFF01C0' +
'3FFFFFFFFFFE00FF803FFE01F80701E0600FE01FFF01C03FFFFFFFFFFE00FF803FFE01FC0701' +
'E0600FE01FFF01C03FFFFFFFFFFE00FF803FFE01FC0601E0700FE01FFF00C03FFFFFFFFFFE00' +
'FF803FFE01FC0001E0700FE01FFF80007FFFFFFFFFFE00FF803FFE01FE0003E0700FE01FFF80' +
'007FFFFFFFFFFE00FF803FFE01FE0007E0700FE01FFFC000FFFFFFFFFFFE00FF803FFE01FF80' +
'0FE0780FE01FFFE001FFFFFFFFFFFE00FF803FFE01FFE03FE0780FE01FFFF807FFFFFFFFFFFE' +
'00FF803FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE00FF803FFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFE00FF803FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';

  // ── Logos de membrete (Membretes.gs:32-77) — LITERALES. _getLogoHexParaTipo
  // hace fallback al logo Tony's si falta. El builder WH usa MEMBRETE_WH. ──
  var LOGO_TIENDA_ME_HEX =
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC1FFFFC000' +
'1C07800381E03C001FF800FFFFFFFFFFF000000007C0001C07800380E03C0007F800FFFFFFFF' +
'FFF000000007C0001C07800380E03C0003F000FFFFFFFFFFF000000007C0001C07800380E03C' +
'0001F0007FFFFFFFFFF000000007C0001C07800380603C0001F0007FFFFFFFFFF000000007C0' +
'001C07800380603C0601F0007FFFFFFFFFF000000007FE01FC0780FF80603C0700F0007FFFFF' +
'FFFFF000000007FE01FC0780FF80203C0700F0007FFFFFFFFFFC0000001FFE01FC0780FF8020' +
'3C0700F0207FFFFFFFFFFC0000001FFE01FC0780FF80203C0700E0207FFFFFFFFFFC0000001F' +
'FE01FC0780FF80203C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFF' +
'FFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780' +
'003C0700E0203FFFFFFFFFFC7FE3FF1FFE01FC07800780003C0700E0203FFFFFFFFFFC7FE3FF' +
'1FFE01FC07800780003C0700E0203FFFFFFFFFFC0000001FFE01FC07800780003C0700E0301F' +
'FFFFFFFFFC0000001FFE01FC0780FF80003C0700C0701FFFFFFFFFFC0000001FFE01FC0780FF' +
'81003C0700C0701FFFFFFFFFFC7FE3FF1FFE01FC0780FF81003C0700C0701FFFFFFFFFFC7FE3' +
'FF1FFE01FC0780FF81003C0700C0001FFFFFFFFFFC7FE3FF1FFE01FC0780FF81803C0700C000' +
'1FFFFFFFFFFC7FE3FF1FFE01FC0780FF81803C0700C0001FFFFFFFFFFC7FE3FF1FFE01FC0780' +
'FF81803C0700C0000FFFFFFFFFFC0000001FFE01FC07800381C03C0600C0000FFFFFFFFFFC00' +
'7F001FFE01FC07800381C03C000180700FFFFFFFFFFC007F001FFE01FC07800381C03C000180' +
'700FFFFFFFFFFC007F001FFE01FC07800381C03C000180700FFFFFFFFFFC007F001FFE01FC07' +
'800381E03C000380700FFFFFFFFFFC007F001FFE01FC07800381E03C000F80700FFFFFFFFFFC' +
'007F001FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC007F001FFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFC007F001FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';
  var LOGO_ALMACEN_WH_HEX =
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC00' +
'7E03F801C007C007FF01FF000703C07FFFE0020007FC007E03F801C007C007FC007F000701C0' +
'7FFFA0020007F8007E03F801C0078007F8001F000701C07FFF20020007F8003E03F801C00780' +
'03F0000F000701C07FFE20020007F8003E03F801C0078003E0000F000700C07FFC20020007F8' +
'003E03F800C0078003E0180F000700C07FF820020007F8003E03F800C0078003E0380701FF00' +
'C07FF0203E0007F8003E03F800C0078003C0380701FF00407FF0203E0007F8103E03F8008007' +
'8103C0380701FF00407FF0203E0007F0103E03F80080070103C0380701FF00407FF0203E0007' +
'F0101E03F80080070101C0380701FF00407FF0203E0007F0101E03F80080070101C03807000F' +
'00007FF0203E0007F0101E03F80000070101C03807000F00007FF0203E0007F0101E03F81004' +
'070101C03FFF000F00007FF0203E0007F0101E03F81004070101C03FFF000F00007FF0203E00' +
'07F0101E03F81004070101C03FFF000F00007FF0203E0007F0180E03F81004070180C03FFF00' +
'0F00007FFFFFFFFFFFE0380E03F81004060380C0380701FF00007FFFFFFFFFFFE0380E03F818' +
'04060380C0380701FF02007FFFFFFFFFFFE0380E03F81804060380C0380701FF02007FFFFFFF' +
'FFFFE0000E03F8180C060000C0380701FF02007FFFFFFFFFFFE0000E03F8180C060000C03807' +
'01FF03007FF0203E0007E0000E03F8180C060000C0380701FF03007FF0203E0007E0000603F8' +
'180C06000060380F01FF03007FF0203E0007E000060018180C06000060180F000703807FF020' +
'3E0007C0380600181C0C04038060000F000703807FF0203E0007C0380600181C0C0403807000' +
'1F000703807FF0203E0007C0380600181C1C04038070003F000703807FF0203E0007C0380600' +
'181C1C0403807C007F000703C07FF0203E0007C0380600181C1C0403807F01FF000703C07FF0' +
'203E0007FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0203E0007FFFFFFFFFFFFFFFFFFFFFF' +
'FFFFFFFFFFFFFFF0203E0007FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF';

  // _getLogoHexParaTipo — Membretes.gs:79-84.
  function _getLogoHexParaTipo(tipo) {
    if (tipo === 'MEMBRETE_ME' && LOGO_TIENDA_ME_HEX)  return LOGO_TIENDA_ME_HEX;
    if (tipo === 'MEMBRETE_WH' && LOGO_ALMACEN_WH_HEX) return LOGO_ALMACEN_WH_HEX;
    return LOGO_TSPL_HEX;
  }

  // Tabla manual de meses ES — Envasados.gs:341 (no depende de locale).
  var MESES_ES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

  // _calcBarcodeAdaptativo — Envasados.gs:22-51. Single source of truth del barcode.
  function _calcBarcodeAdaptativo(codigo) {
    var bc = String(codigo || '').replace(/"/g, '');
    var bcLen = bc.length;
    var modules = 11 * bcLen + 35;
    var narrowBc, quietZoneMin;
    if (modules * 3 <= 340) {
      narrowBc = 3; quietZoneMin = 30;  // bcLen <= 7  ÓPTIMO
    } else if (modules * 2 <= 360) {
      narrowBc = 2; quietZoneMin = 20;  // bcLen <= 13 estándar legible
    } else if (modules * 2 <= 376) {
      narrowBc = 2; quietZoneMin = 12;  // bcLen = 14  apretado pero legible
    } else {
      narrowBc = 2; quietZoneMin = 8;   // bcLen > 14  al límite físico
    }
    var barcodeWidth = modules * narrowBc;
    var barcodeHeight = 48;             // 48 dots (6.0mm)
    var barcodeX = Math.max(quietZoneMin, Math.floor((400 - barcodeWidth) / 2));
    return {
      bc:            bc,
      bcLen:         bcLen,
      narrowBc:      narrowBc,
      barcodeWidth:  barcodeWidth,
      barcodeHeight: barcodeHeight,
      quietZoneMin:  quietZoneMin,
      barcodeX:      barcodeX
    };
  }

  // _normalizeEtq — Envasados.gs:344-348. NFD + strip diacríticos.
  function _normalizeEtq(s) {
    if (s === null || s === undefined) return '';
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  // _calcVencimientoEtq — Envasados.gs:350-357. fechaEnvasado + 1 año exacto → "MES/yyyy".
  function _calcVencimientoEtq(fechaEnvasado) {
    var d = fechaEnvasado ? new Date(fechaEnvasado) : new Date();
    d.setFullYear(d.getFullYear() + 1);
    return MESES_ES[d.getMonth()] + '/' + d.getFullYear();
  }

  // _detectHighlightsEtq — Envasados.gs:361-384. Tokens diferenciadores (+ último=peso).
  function _detectHighlightsEtq(targetTokens, allTokenized) {
    var hl = {};
    hl[targetTokens.length - 1] = true;
    for (var i = 0; i < allTokenized.length; i++) {
      var s = allTokenized[i];
      if (s === targetTokens) continue;
      if (s[0] !== targetTokens[0]) continue;
      for (var pos = 1; pos < targetTokens.length; pos++) {
        if (hl[pos]) continue;
        var prior = true;
        for (var k = 0; k < pos; k++) {
          if (s[k] !== targetTokens[k]) { prior = false; break; }
        }
        if (prior && s[pos] !== undefined && s[pos] !== targetTokens[pos]) {
          hl[pos] = true;
          break;
        }
      }
    }
    var out = [];
    for (var key in hl) if (hl[key]) out.push(parseInt(key));
    out.sort(function(a,b){ return a - b; });
    return out;
  }

  // _fontWidthEtq — Envasados.gs:387. Font 3 normal (16), Font 4 highlight (24).
  function _fontWidthEtq(isHighlight) { return isHighlight ? 24 : 16; }

  // _wrapTokensEtq — Envasados.gs:391-437. Word-wrap a max 2 líneas, smart-highlight.
  function _wrapTokensEtq(tokens, highlights) {
    var MAX_W = 370, SPACE = 8;
    var widths = tokens.map(function(t, i) {
      return t.length * _fontWidthEtq(highlights.indexOf(i) >= 0);
    });
    var isHl = function(i) { return highlights.indexOf(i) >= 0; };

    var total = 0;
    for (var i = 0; i < widths.length; i++) total += widths[i] + (i > 0 ? SPACE : 0);
    if (total <= MAX_W) {
      return [tokens.map(function(t, i) { return { tok: t, hl: isHl(i), w: widths[i] }; })];
    }

    var firstHl = highlights.length > 0 ? highlights[0] : tokens.length;
    if (firstHl > 0 && firstHl < tokens.length) {
      var l1 = [], w1 = 0;
      for (var a = 0; a < firstHl; a++) {
        w1 += widths[a] + (l1.length > 0 ? SPACE : 0);
        l1.push({ tok: tokens[a], hl: false, w: widths[a] });
      }
      var l2 = [], w2 = 0;
      for (var b = firstHl; b < tokens.length; b++) {
        w2 += widths[b] + (l2.length > 0 ? SPACE : 0);
        l2.push({ tok: tokens[b], hl: isHl(b), w: widths[b] });
      }
      if (w1 <= MAX_W && w2 <= MAX_W) return [l1, l2];
    }

    var lines = [[]], curW = 0;
    for (var c = 0; c < tokens.length; c++) {
      var sep = lines[lines.length - 1].length === 0 ? 0 : SPACE;
      if (curW + sep + widths[c] <= MAX_W) {
        lines[lines.length - 1].push({ tok: tokens[c], hl: isHl(c), w: widths[c] });
        curW += sep + widths[c];
      } else if (lines.length === 1) {
        lines.push([{ tok: tokens[c], hl: isHl(c), w: widths[c] }]);
        curW = widths[c];
      } else {
        lines[1].push({ tok: tokens[c], hl: isHl(c), w: widths[c] });
        curW += sep + widths[c];
      }
    }
    return lines;
  }

  // _hexToBytesEtq — Envasados.gs:439-443. Hex string → array de bytes (logo crudo).
  function _hexToBytesEtq(hex) {
    var arr = [];
    for (var i = 0; i < hex.length; i += 2) arr.push(parseInt(hex.substr(i, 2), 16));
    return arr;
  }

  // _strToBytesEtq — Envasados.gs:445-449. ASCII string → array de bytes (& 0xFF).
  function _strToBytesEtq(s) {
    var arr = [];
    for (var i = 0; i < s.length; i++) arr.push(s.charCodeAt(i) & 0xFF);
    return arr;
  }

  // Tokeniza el catálogo de envasables para smart-highlight. En el GAS sale de
  // Sheets (_getAllEnvasablesTokens, Envasados.gs:452-468); acá llega por param
  // `allEnvasables` (array de { descripcion } o { tokens }). Si no llega → [].
  function _normalizeEnvasables(allEnvasables) {
    if (!Array.isArray(allEnvasables)) return [];
    return allEnvasables.map(function(p) {
      if (p && Array.isArray(p.tokens)) return { tokens: p.tokens };
      return { tokens: _normalizeEtq((p && p.descripcion) || '').split(/\s+/) };
    });
  }

  // ── ETIQUETA CASERITO / ENVASADO (TSPL2) — _buildTSPLEtq (Envasados.gs:484-608) ──
  // params: { codigoBarra, descripcion, codigoDerivado?, cantidad?, unidades?,
  //           fechaVencimiento?, idLote?, fechaEnvasado?, allEnvasables? }
  //   - codigoBarra: código del derivado (va al barcode Code128 y al texto).
  //   - descripcion: nombre del producto (se normaliza, tokeniza, wrap + highlight).
  //   - fechaVencimiento: si viene, se respeta (legacy override, igual que el GAS
  //                       _imprimirEtiquetasEnvasado:639-651 — vto = fecha tal cual,
  //                       formateada MES/yyyy SIN sumar año). Si NO viene, se calcula
  //                       fechaEnvasado (o hoy) + 1 año vía _calcVencimientoEtq.
  //   - unidades: cantidad de etiquetas a emitir (PRINT 1,1 por cada una). Default 1.
  //   - codigoDerivado/cantidad/idLote: aceptados por compat con el caller del GAS;
  //                       NO se imprimen en la etiqueta (el GAS tampoco los usa en TSPL).
  //   - allEnvasables: catálogo tokenizado para el smart-highlight (opcional).
  // Devuelve { printerHint:'ADHESIVO', title, base64 }.
  //
  // DRIFT: offsetY = ADHESIVO_OFFSET_Y_BASE (0). El header (SIZE/GAP/DIRECTION/
  // DENSITY/SPEED) se emite UNA vez; luego N bloques (CLS + contenido + PRINT 1,1).
  function armarEtiquetaCaserito(params) {
    params = params || {};
    var producto = {
      codigoBarra: String(params.codigoBarra || ''),
      descripcion: String(params.descripcion || params.codigoBarra || '')
    };
    var unidades = parseInt(params.unidades) || 1;
    var allEnv = _normalizeEnvasables(params.allEnvasables);

    // Vto: réplica de _imprimirEtiquetasEnvasado (Envasados.gs:639-653).
    // Si vino fechaVencimiento explícita → MES/yyyy de ESA fecha (sin +1 año).
    // Si no → fechaEnvasado (o hoy) +1 año vía _calcVencimientoEtq.
    var vto;
    if (params.fechaVencimiento) {
      var dOverride = new Date(params.fechaVencimiento);
      vto = MESES_ES[dOverride.getMonth()] + '/' + dOverride.getFullYear();
    } else {
      var fechaEnvasado = params.fechaEnvasado || new Date();
      vto = _calcVencimientoEtq(fechaEnvasado);
    }

    var bytes = _buildBytesEtqCaserito(producto, vto, unidades, allEnv);

    return {
      printerHint: 'ADHESIVO',
      title: 'Etiqueta ' + _normalizeEtq(producto.descripcion),
      base64: _bytesB64(bytes)
    };
  }

  // Bytes TSPL2 de la etiqueta de envasado — espejo EXACTO de _buildTSPLEtq, con
  // offsetY=base (drift desactivado) y vto ya resuelto por el caller.
  function _buildBytesEtqCaserito(producto, vto, unidades, allEnvasables) {
    var descNorm = _normalizeEtq(producto.descripcion);
    var tokens = descNorm.split(/\s+/);
    var allTok = allEnvasables.map(function(p) { return p.tokens; });
    var highlights = _detectHighlightsEtq(tokens, allTok);
    var lines = _wrapTokensEtq(tokens, highlights);

    var gapMm   = ADHESIVO_GAP_MM;
    var density = ADHESIVO_DENSITY;
    var speed   = ADHESIVO_SPEED;

    // ── Header GLOBAL del job (una sola vez, antes del loop) — Envasados.gs:519-527 ──
    var headerGlobal = [
      'SIZE 50 mm,25 mm',
      'GAP ' + gapMm + ' mm,0 mm',
      'DIRECTION 1',
      'DENSITY ' + density,
      'SPEED ' + speed,
      ''
    ].join('\r\n');
    var bytes = _strToBytesEtq(headerGlobal);

    var _bcCalc = _calcBarcodeAdaptativo(producto.codigoBarra);
    var bc = _bcCalc.bc;
    var narrowBc = _bcCalc.narrowBc;
    var barcodeHeight = _bcCalc.barcodeHeight;
    var barcodeX = _bcCalc.barcodeX;
    var frameX1 = 10, frameX2 = 389;
    var cmL = 12;
    var codigoFontW = 8;
    var codigoWidth = bc.length * codigoFontW;
    var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));

    var N = unidades || 1;
    for (var iEtq = 0; iEtq < N; iEtq++) {
      // TODO(drift): el GAS aquí calcula offsetY_i = clamp(offsetBase + driftDots ×
      // (printsBase + iEtq)). Sin Script Properties del rollo en cliente, usamos el
      // OFFSET BASE constante (= drift desactivado). Coincide con el GAS si drift=0.
      var offsetY = ADHESIVO_OFFSET_Y_BASE;

      // ── CLS + LOGO ──
      bytes = bytes.concat(_strToBytesEtq(
        'CLS\r\n' +
        'BITMAP 5,' + (2 + offsetY) + ',' + LOGO_W_BYTES + ',' + LOGO_H + ',0,'
      ));
      bytes = bytes.concat(_hexToBytesEtq(LOGO_TSPL_HEX));
      bytes = bytes.concat(_strToBytesEtq('\r\n'));

      // ── Vto + separador ──
      bytes = bytes.concat(_strToBytesEtq('TEXT 232,' + (12 + offsetY) + ',"2",0,1,1,"Vto ' + vto + '"\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR 5,' + (42 + offsetY) + ',390,1\r\n'));

      // ── Descripción con centrado vertical ──
      var DESC_AREA_Y0 = 46, DESC_AREA_H = 72;
      var LINE_H = 38, SPACE = 8;
      var startY;
      if (lines.length === 1) {
        var lineHasHl = lines[0].some(function(t) { return t.hl; });
        var lineHeight = lineHasHl ? 32 : 24;
        var baselineOffset = lineHasHl ? 0 : 4;
        startY = DESC_AREA_Y0 + Math.floor((DESC_AREA_H - lineHeight) / 2) - baselineOffset + offsetY;
      } else {
        startY = DESC_AREA_Y0 + offsetY;
      }
      for (var li = 0; li < lines.length; li++) {
        var line = lines[li];
        var totalW = 0;
        for (var ti = 0; ti < line.length; ti++) totalW += line[ti].w + (ti > 0 ? SPACE : 0);
        var x = Math.max(5, Math.round((400 - totalW) / 2));
        var y = startY + li * LINE_H;
        for (var tj = 0; tj < line.length; tj++) {
          var o = line[tj];
          var font = o.hl ? '4' : '3';
          var yAdj = o.hl ? y : y + 4;
          var safe = String(o.tok).replace(/"/g, "'");
          bytes = bytes.concat(_strToBytesEtq('TEXT ' + x + ',' + yAdj + ',"' + font + '",0,1,1,"' + safe + '"\r\n'));
          x += o.w + SPACE;
        }
      }

      // ── Frame corner marks + barcode + código ──
      var barcodeY = 124 + offsetY;
      var frameY1 = 118 + offsetY, frameY2 = 196 + offsetY;
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
      bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + bc + '"\r\n'));
      var codigoY = barcodeY + barcodeHeight + 8;
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + bc + '"\r\n'));

      // ── PRINT esta etiqueta ──
      bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
    }

    return bytes;
  }

  // ── MEMBRETE ME góndola (TSPL2) — _buildTSPLMembreteMe (Membretes.gs:120-240) ──
  // Una etiqueta por item (el caller concatena para lote). Precio MEGA Font 5.
  // params: { codigoBarra, descripcion, precio, skuBase?, esSkuBase?, idProducto?,
  //           allEnvasables? }
  //   - esSkuBase=true → usa skuBase como código principal (barcode + texto).
  //   - precio: número (o string parseable); se imprime 'S/ ' + toFixed(2).
  // Devuelve { printerHint:'ADHESIVO', title, base64 }. offsetY = base (drift off).
  function armarMembreteMe(params) {
    params = params || {};
    var producto = {
      codigoBarra: String(params.codigoBarra || ''),
      descripcion: String(params.descripcion || ''),
      precio:      params.precio,
      skuBase:     String(params.skuBase || ''),
      esSkuBase:   !!params.esSkuBase,
      idProducto:  String(params.idProducto || '')
    };
    var allEnv = _normalizeEnvasables(params.allEnvasables);
    var bytes = _buildBytesMembreteMe(producto, allEnv);
    var codigoTitle = String(
      (producto.esSkuBase ? producto.skuBase : producto.codigoBarra)
      || producto.codigoBarra || producto.skuBase || producto.idProducto || ''
    );
    return {
      printerHint: 'ADHESIVO',
      title: 'Membrete ME ' + (codigoTitle || _normalizeEtq(producto.descripcion)),
      base64: _bytesB64(bytes)
    };
  }

  function _buildBytesMembreteMe(producto, allEnvasables) {
    var descNorm = _normalizeEtq(producto.descripcion || '');
    var tokens = descNorm.split(/\s+/);
    var allTok = (allEnvasables || []).map(function(p) { return p.tokens; });
    var highlights = _detectHighlightsEtq(tokens, allTok);
    var lines = _wrapTokensEtq(tokens, highlights);

    var gapMm   = ADHESIVO_GAP_MM;
    var density = ADHESIVO_DENSITY;
    var speed   = ADHESIVO_SPEED;
    // TODO(drift): GAS = _calcularOffsetEfectivoParaPrint() (drift acumulado de
    // Properties). Sin Properties en cliente → offset base (drift desactivado).
    var offsetY = ADHESIVO_OFFSET_Y_BASE;

    var header = [
      'SIZE 50 mm,25 mm',
      'GAP ' + gapMm + ' mm,0 mm',
      'DIRECTION 1',
      'DENSITY ' + density,
      'SPEED ' + speed,
      'CLS'
    ].join('\r\n') + '\r\n';
    var bytes = _strToBytesEtq(header);

    // ─── 1) DESCRIPCIÓN Font 3, 1 LÍNEA centrada, área Y=2-26 ───
    var SPACE = 8;
    var primeraLinea = (lines[0] || []).map(function(t) {
      return { tok: t.tok, w: t.w, hl: t.hl };
    });
    if (lines.length > 1 && primeraLinea.length > 0) {
      var ultimo = primeraLinea[primeraLinea.length - 1];
      ultimo.tok = ultimo.tok.replace(/[\.,]+$/, '') + '..';
      ultimo.w = ultimo.tok.length * 16;
    }
    var lineW = 0;
    for (var ti = 0; ti < primeraLinea.length; ti++) {
      lineW += primeraLinea[ti].w + (ti > 0 ? SPACE : 0);
    }
    while (lineW > 380 && primeraLinea.length > 1) {
      var quitado = primeraLinea.pop();
      lineW -= quitado.w + SPACE;
      var nuevoUlt = primeraLinea[primeraLinea.length - 1];
      if (nuevoUlt.tok.slice(-2) !== '..') {
        lineW -= nuevoUlt.w;
        nuevoUlt.tok = nuevoUlt.tok.replace(/[\.,]+$/, '') + '..';
        nuevoUlt.w = nuevoUlt.tok.length * 16;
        lineW += nuevoUlt.w;
      }
    }
    var descX = Math.max(5, Math.round((400 - lineW) / 2));
    var descY = 2 + offsetY;
    for (var tj = 0; tj < primeraLinea.length; tj++) {
      var o = primeraLinea[tj];
      var safe = String(o.tok).replace(/"/g, "'");
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + descX + ',' + descY + ',"3",0,1,1,"' + safe + '"\r\n'));
      descX += o.w + SPACE;
    }

    // ─── 2) PRECIO MEGA Font 5 (32×48) CENTRADO en Y=30 ───
    var precioStr = 'S/ ' + (parseFloat(producto.precio) || 0).toFixed(2);
    var precioFontW = 32;
    var precioWidth = precioStr.length * precioFontW;
    var precioX = Math.max(5, Math.round((400 - precioWidth) / 2));
    var precioY = 30 + offsetY;
    bytes = bytes.concat(_strToBytesEtq('TEXT ' + precioX + ',' + precioY + ',"5",0,1,1,"' + precioStr + '"\r\n'));

    // ─── 3) Línea decorativa gruesa bajo precio Y=82 ───
    var lineaDecoY = 82 + offsetY;
    var lineaDecoX = Math.max(30, precioX - 8);
    var lineaDecoW = Math.min(340, precioWidth + 16);
    bytes = bytes.concat(_strToBytesEtq('BAR ' + lineaDecoX + ',' + lineaDecoY + ',' + lineaDecoW + ',3\r\n'));

    // ─── 4) BARCODE altura 48 con frame + corner marks Y=88-148 ───
    var codigo = String(
      (producto.esSkuBase ? producto.skuBase : producto.codigoBarra)
      || producto.codigoBarra
      || producto.skuBase
      || producto.idProducto
      || ''
    ).replace(/"/g, '');
    if (!codigo) codigo = 'SIN-CODIGO';
    var _bcCalc = _calcBarcodeAdaptativo(codigo);
    var narrowBc = _bcCalc.narrowBc;
    var barcodeHeight = _bcCalc.barcodeHeight;
    var barcodeX = _bcCalc.barcodeX;
    var barcodeY = 94 + offsetY;
    var frameX1 = 10, frameX2 = 389;
    var frameY1 = 88 + offsetY, frameY2 = 148 + offsetY, cmL = 12;
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + codigo + '"\r\n'));

    // ─── 5) Texto código Font 1 centrado Y=152 ───
    var codigoFontW = 8;
    var codigoWidth = codigo.length * codigoFontW;
    var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));
    var codigoY = 152 + offsetY;
    bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + codigo + '"\r\n'));

    bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
    return bytes;
  }

  // ── MEMBRETE WH almacén (TSPL2) — _buildTSPLMembreteWh (Membretes.gs:262-368) ──
  // Una etiqueta por item (el caller concatena para lote). Logo + highlight.
  // params: { descripcion, codigo|codigoBarra, skuBase?, idProducto?,
  //           esCabecera?, indice?, total?, allEnvasables? }
  //   - codigo: código principal (barcode + texto). Fallback codigoBarra/skuBase/idProducto.
  //   - esCabecera=true → tag "CAB" en esquina sup-der (solo si total>1).
  //   - indice/total: posición en la serie (tag "i/total" si total>1). Default 1/1.
  // Devuelve { printerHint:'ADHESIVO', title, base64 }. offsetY = base (drift off).
  function armarMembreteWh(params) {
    params = params || {};
    var producto = {
      descripcion: String(params.descripcion || ''),
      codigo:      String(params.codigo || params.codigoBarra || ''),
      codigoBarra: String(params.codigoBarra || ''),
      skuBase:     String(params.skuBase || ''),
      idProducto:  String(params.idProducto || '')
    };
    var esCabecera = !!params.esCabecera;
    var indice = parseInt(params.indice) || 1;
    var total  = parseInt(params.total)  || 1;
    var allEnv = _normalizeEnvasables(params.allEnvasables);

    var bytes = _buildBytesMembreteWh(producto, esCabecera, indice, total, allEnv);
    var codigoTitle = String(producto.codigo || producto.codigoBarra || producto.skuBase || producto.idProducto || '');
    return {
      printerHint: 'ADHESIVO',
      title: 'Membrete WH ' + (codigoTitle || _normalizeEtq(producto.descripcion)),
      base64: _bytesB64(bytes)
    };
  }

  function _buildBytesMembreteWh(producto, esCabecera, indice, total, allEnvasables) {
    var descNorm = _normalizeEtq(producto.descripcion || '');
    var tokens = descNorm.split(/\s+/);
    var allTok = (allEnvasables || []).map(function(p) { return p.tokens; });
    var highlights = _detectHighlightsEtq(tokens, allTok);
    var lines = _wrapTokensEtq(tokens, highlights);

    var gapMm   = ADHESIVO_GAP_MM;
    var density = ADHESIVO_DENSITY;
    var speed   = ADHESIVO_SPEED;
    // TODO(drift): GAS = _calcularOffsetEfectivoParaPrint(); acá offset base.
    var offsetY = ADHESIVO_OFFSET_Y_BASE;

    var logoHex = _getLogoHexParaTipo('MEMBRETE_WH');

    var header = [
      'SIZE 50 mm,25 mm',
      'GAP ' + gapMm + ' mm,0 mm',
      'DIRECTION 1',
      'DENSITY ' + density,
      'SPEED ' + speed,
      'CLS',
      'BITMAP 5,' + (2 + offsetY) + ',' + LOGO_W_BYTES + ',' + LOGO_H + ',0,'
    ].join('\r\n');

    var bytes = _strToBytesEtq(header);
    bytes = bytes.concat(_hexToBytesEtq(logoHex));
    bytes = bytes.concat(_strToBytesEtq('\r\n'));

    // Indicador esquina sup-der SOLO si serie multi-código.
    if (total > 1) {
      var tagTexto = esCabecera ? 'CAB' : (indice + '/' + total);
      var tagX = 400 - tagTexto.length * 8 - 8;
      bytes = bytes.concat(_strToBytesEtq('TEXT ' + tagX + ',' + (4 + offsetY) + ',"2",0,1,1,"' + tagTexto + '"\r\n'));
    }

    // Descripción área Y=46-118 (centrada igual que envasado).
    var DESC_AREA_Y0 = 46, DESC_AREA_H = 72;
    var LINE_H = 38, SPACE = 8;
    var startY;
    if (lines.length === 1) {
      var lineHasHl = lines[0].some(function(t) { return t.hl; });
      var lineHeight = lineHasHl ? 32 : 24;
      var baselineOffset = lineHasHl ? 0 : 4;
      startY = DESC_AREA_Y0 + Math.floor((DESC_AREA_H - lineHeight) / 2) - baselineOffset + offsetY;
    } else {
      startY = DESC_AREA_Y0 + offsetY;
    }
    for (var li = 0; li < Math.min(lines.length, 2); li++) {
      var line = lines[li];
      var totalW = 0;
      for (var ti = 0; ti < line.length; ti++) totalW += line[ti].w + (ti > 0 ? SPACE : 0);
      var x = Math.max(5, Math.round((400 - totalW) / 2));
      var y = startY + li * LINE_H;
      for (var tj = 0; tj < line.length; tj++) {
        var o = line[tj];
        var font = o.hl ? '4' : '3';
        var yAdj = o.hl ? y : y + 4;
        var safe = String(o.tok).replace(/"/g, "'");
        bytes = bytes.concat(_strToBytesEtq('TEXT ' + x + ',' + yAdj + ',"' + font + '",0,1,1,"' + safe + '"\r\n'));
        x += o.w + SPACE;
      }
    }

    // Barcode adaptativo + frame corner marks (igual que envasado).
    var codigo = String(producto.codigo || producto.codigoBarra || producto.skuBase || producto.idProducto || '').replace(/"/g, '');
    if (!codigo) codigo = 'SIN-CODIGO';
    var _bcCalc = _calcBarcodeAdaptativo(codigo);
    var narrowBc = _bcCalc.narrowBc;
    var barcodeHeight = _bcCalc.barcodeHeight;
    var barcodeX = _bcCalc.barcodeX;
    var barcodeY = 124 + offsetY;
    var frameX1 = 10, frameX2 = 389;
    var frameY1 = 118 + offsetY, frameY2 = 196 + offsetY, cmL = 12;
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY1 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + frameY1 + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + frameY2 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX1 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + (frameX2 - cmL + 1) + ',' + frameY2 + ',' + cmL + ',1\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BAR ' + frameX2 + ',' + (frameY2 - cmL + 1) + ',1,' + cmL + '\r\n'));
    bytes = bytes.concat(_strToBytesEtq('BARCODE ' + barcodeX + ',' + barcodeY + ',"128",' + barcodeHeight + ',0,0,' + narrowBc + ',' + narrowBc + ',"' + codigo + '"\r\n'));

    var codigoFontW = 8;
    var codigoWidth = codigo.length * codigoFontW;
    var codigoX = Math.max(frameX1 + 4, Math.floor((400 - codigoWidth) / 2));
    var codigoY = barcodeY + barcodeHeight + 8;
    bytes = bytes.concat(_strToBytesEtq('TEXT ' + codigoX + ',' + codigoY + ',"1",0,1,1,"' + codigo + '"\r\n'));

    bytes = bytes.concat(_strToBytesEtq('PRINT 1,1\r\n'));
    return bytes;
  }

  return {
    _b64Utf8,        // expuesto para los builders que arman ARRAY de bytes (TSPL2) cuando se porten
    _bytesB64,       // expuesto: base64 de array de bytes (ESC/POS raw)
    armarBienvenida,
    armarAvisoCajeros,
    armarMembreteStd,
    armarEtiquetaCaserito,
    armarMembreteMe,
    armarMembreteWh,
  };
})();

// Export para entorno Node (validación/tests). En navegador es no-op.
if (typeof module !== 'undefined' && module.exports) { module.exports = ImpresionDirecta; }
