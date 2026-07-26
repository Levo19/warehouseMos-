// ============================================================
// warehouseMos — js/voucher.js
// Voucher-imagen compartible de una GUÍA o PREINGRESO (+ sus anexos), estilo
// comprobante con corte de ticket, QR al reporte público y aclaración.
// Datos: wh.voucher_data (SQL 566). QR: qrcode-generator. Comparte imagen (Web
// Share con archivo en móvil; en PC descarga PNG + abre WhatsApp con el link).
// ============================================================

(function() {
  'use strict';

  const BASE = 'https://levo19.github.io/warehouseMos-/reporte.html';
  const TIPO = {
    INGRESO_PROVEEDOR:'📥 Ingreso de proveedor', INGRESO_JEFATURA:'📥 Ingreso de jefatura',
    INGRESO_DEVOLUCION_ZONA:'📥 Devolución de zona', SALIDA_ZONA:'📤 Salida a zona',
    SALIDA_DEVOLUCION:'📤 Devolución', SALIDA_JEFATURA:'📤 Salida jefatura',
    SALIDA_ENVASADO:'📤 Envasado', SALIDA_MERMA:'📤 Merma', TRANSFORMACION:'♻️ Transformación'
  };

  // [fix] no usar \b tras "Sí": la í es no-ASCII y \b de JS no la reconoce como límite → el tag no parseaba.
  // Se ancla con lookahead a espacio / pipe / fin.
  function _tags(c){ const s = String(c||''); const B = '(?=\\s|\\||$)';
    const has = (k,v) => new RegExp(k + ':\\s*' + v + B, 'i').test(s);
    return {
      comp:  has('comprobante','s[ií]') ? 'si' : has('comprobante','no') ? 'no' : null,
      compl: has('completo','s[ií]')    ? 'si' : has('completo','no')    ? 'no' : null }; }
  function _libre(c){ return String(c||'')
    .replace(/Comprobante:\s*(?:S[ií]|No)\s*\|?\s*/gi,'').replace(/Completo:\s*(?:S[ií]|No)\s*\|?\s*/gi,'')
    .replace(/^\s*\|\s*/,'').replace(/\s*\|\s*$/,'').trim(); }
  function _money(m){ const v = parseFloat(m); return isNaN(v) ? null : 'S/ ' + v.toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2}); }
  function _tipoLabel(t){ return TIPO[t] || (t ? String(t).replace(/_/g,' ').toLowerCase() : '—'); }
  function _link(tipo, id){ return BASE + '?tipo=' + tipo + '&id=' + encodeURIComponent(id); }
  function _fmtFecha(f){ const raw = String(f||'').slice(0,10); if(!raw) return ''; const p=raw.split('-'); if(p.length<3) return raw;
    const ms=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']; return parseInt(p[2],10)+' '+ms[parseInt(p[1],10)-1]+' '+p[0]; }
  function _accent(t){
    if (t.kind === 'preingreso') return { main:'#5b3fb8', bg:'#efe9fb', ribbon:'PREINGRESO', rb:'#5b3fb8' };
    const g = String(t.tipoGuia||'');
    if (g.indexOf('INGRESO') === 0) return { main:'#15803d', bg:'#e7f4ec', ribbon:'GUÍA', rb:'#15803d' };
    if (g.indexOf('SALIDA')  === 0) return { main:'#1d4ed8', bg:'#e7edfb', ribbon:'GUÍA', rb:'#1d4ed8' };
    return { main:'#b45309', bg:'#fbf0dd', ribbon:'GUÍA', rb:'#b45309' };
  }
  function _estadoPill(est){
    const s = String(est||'').toUpperCase();
    if (s.indexOf('CERR') === 0 || s.indexOf('PROC') === 0) return { txt:'✓ '+(est||''), bg:'#e7f4ec', fg:'#15803d' };
    if (s.indexOf('ANUL') === 0) return { txt:est, bg:'#fbe4e4', fg:'#b91c1c' };
    if (s.indexOf('ABIER') === 0 || s.indexOf('PEND') === 0) return { txt:'● '+(est||''), bg:'#fbf0dd', fg:'#b45309' };
    return { txt: est||'—', bg:'#efe9fb', fg:'#5b3fb8' };
  }

  // ── layout de un ticket (mide si draw=false, dibuja si draw=true). Devuelve el nuevo y. ──
  function _ticket(x, t, y0, W, PAD, draw, rr, wrap, trunc) {
    let y = y0;
    const isG = t.kind === 'guia';
    const ac = _accent(t);
    const headerH = isG ? 98 : 82;
    if (draw) {
      x.fillStyle = ac.bg; x.fillRect(PAD, y, W - 2*PAD, headerH);
      x.fillStyle = 'rgba(0,0,0,.06)'; x.fillRect(PAD, y+headerH-1, W-2*PAD, 1);
      // brand + ribbon
      x.fillStyle = '#8b93a1'; x.font = '800 10px -apple-system,Segoe UI,sans-serif';
      x.fillText(isG ? 'InversionMos · Almacén' : 'Preingreso', PAD+16, y+18);
      const rw = x.measureText(ac.ribbon).width + 18;
      x.fillStyle = ac.rb; rr(W-PAD-16-rw, y+9, rw, 16, 8); x.fill();
      x.fillStyle = '#fff'; x.font = '800 9.5px ui-monospace,monospace'; x.textAlign='center';
      x.fillText(ac.ribbon, W-PAD-16-rw/2, y+20); x.textAlign='left';
      // proveedor (principal)
      x.fillStyle = '#1b1e26'; x.font = '900 20px -apple-system,Segoe UI,sans-serif';
      x.fillText(trunc(t.proveedor || '—', W-2*PAD-32, '900 20px -apple-system,Segoe UI,sans-serif'), PAD+16, y+43);
      // subtítulo (tipo guía) / nada en preingreso
      let subY = y+62;
      if (isG) { x.fillStyle = '#565d6b'; x.font = '700 13px -apple-system,Segoe UI,sans-serif'; x.fillText(_tipoLabel(t.tipoGuia), PAD+16, y+62); subY = y+62; }
      // id + estado pill
      const id = isG ? t.idGuia : t.idPreingreso;
      x.fillStyle = '#1b1e26'; x.font = '700 13px ui-monospace,monospace'; x.fillText(id, PAD+16, y+ (isG?86:70));
      const pi = _estadoPill(t.estado); x.font = '800 11px -apple-system,sans-serif';
      const pw = x.measureText(pi.txt).width + 22, py = y + (isG?74:58);
      x.fillStyle = pi.bg; rr(W-PAD-16-pw, py, pw, 22, 11); x.fill();
      x.fillStyle = pi.fg; x.textAlign='center'; x.fillText(pi.txt, W-PAD-16-pw/2, py+15); x.textAlign='left';
    }
    y += headerH;

    if (isG) {
      // marcas: completo / comprobante
      const tg = _tags(t.comentario);
      const marks = [];
      if (tg.compl === 'si') marks.push(['✅ Completo','#e7f4ec','#15803d']);
      else if (tg.compl === 'no') marks.push(['⚠️ Incompleto','#fbf0dd','#b45309']);
      if (tg.comp === 'si') marks.push(['🧾 Con comprobante','#e7edfb','#1d4ed8']);
      else if (tg.comp === 'no') marks.push(['Sin comprobante','#f4f1e8','#8b93a1']);
      if (marks.length) {
        if (draw) { let mx = PAD+16; x.font='800 12px -apple-system,sans-serif';
          marks.forEach(m => { const w = x.measureText(m[0]).width + 20; x.fillStyle=m[1]; rr(mx,y+2,w,26,9); x.fill(); x.fillStyle=m[2]; x.textAlign='left'; x.fillText(m[0],mx+10,y+19); mx+=w+7; }); }
        y += 34;
      }
      // post-it comentario libre
      const libre = _libre(t.comentario);
      if (libre) y = _postit(x, y, W, PAD, '✍️ Comentario del operador', libre, draw, rr, wrap);
      // productos
      const det = t.detalle || [];
      if (draw) { x.fillStyle='#8b93a1'; x.font='800 10.5px -apple-system,sans-serif'; x.fillText('PRODUCTOS · ' + det.length, PAD+16, y+16); }
      y += 24;
      det.forEach((d, i) => {
        if (draw) {
          if (i) { x.strokeStyle='#efe9dd'; x.beginPath(); x.moveTo(PAD+16,y); x.lineTo(W-PAD-16,y); x.stroke(); }
          const col = d.esProductoNuevo ? '#1d4ed8' : d.esSinIdentificar ? '#b45309' : '#15803d';
          x.fillStyle = col; x.beginPath(); x.arc(PAD+20, y+14, 4, 0, 7); x.fill();
          x.fillStyle = '#1b1e26'; x.font='600 13.5px -apple-system,sans-serif';
          let nm = d.descripcion || d.codigoProducto || '—';
          const flag = d.esProductoNuevo ? '  🆕 nuevo' : d.esSinIdentificar ? '  ⚠️ sin identificar' : '';
          x.fillText(trunc(nm, W-2*PAD-150, '600 13.5px -apple-system,sans-serif') + flag, PAD+32, y+19);
          x.fillStyle = '#565d6b'; x.font='700 13px ui-monospace,monospace'; x.textAlign='right';
          x.fillText(String(d.cantidad != null ? d.cantidad : ''), W-PAD-16, y+19); x.textAlign='left';
        }
        y += 27;
      });
      y += 10;
    } else {
      // preingreso: monto + tags + post-it
      const mon = _money(t.monto);
      if (draw && mon) {
        x.fillStyle='#15803d'; x.font='900 23px -apple-system,Segoe UI,sans-serif';
        x.fillText(mon, PAD+16, y+22);
        const mw = x.measureText(mon).width;   // medir con la fuente del monto (evita solape del label)
        x.fillStyle='#8b93a1'; x.font='700 11px -apple-system,sans-serif';
        x.fillText('monto declarado', PAD+16+mw+12, y+21);
      }
      y += 34;
      const tags = [];
      const carg = t.cargadores || [];
      if (carg.length) tags.push('🛺 ' + carg.length + ' cargador' + (carg.length===1?'':'es') + (carg.length? ' · ' + carg.slice(0,3).join(', ') : ''));
      const nf = (t.fotos||[]).length; if (nf) tags.push('📷 ' + nf + ' foto' + (nf===1?'':'s'));
      if (t.fecha) tags.push('📅 ' + _fmtFecha(t.fecha));
      if (draw && tags.length) { let mx=PAD+16; x.font='700 12px -apple-system,sans-serif';
        tags.forEach(tg => { const w=x.measureText(tg).width+20; x.fillStyle='#f4f1e8'; rr(mx,y,w,26,9); x.fill(); x.fillStyle='#565d6b'; x.fillText(tg,mx+10,y+17); mx+=w+7; if(mx>W-2*PAD-80){/*wrap*/} }); }
      if (tags.length) y += 34;
      const libre = _libre(t.comentario);
      if (libre) y = _postit(x, y, W, PAD, '✍️ Nota del preingreso', libre, draw, rr, wrap);
      y += 8;
    }
    return y;
  }

  function _postit(x, y, W, PAD, label, texto, draw, rr, wrap) {
    const innerW = W - 2*PAD - 26;
    const lines = wrap(texto, innerW, 'italic 600 13px -apple-system,sans-serif');
    const boxH = 22 + lines.length * 18 + 12;
    if (draw) {
      x.fillStyle = '#fff6c9'; rr(PAD+16, y, W-2*PAD-32, boxH, 5); x.fill();
      x.strokeStyle = '#f2e39a'; x.lineWidth = 1; rr(PAD+16, y, W-2*PAD-32, boxH, 5); x.stroke();
      x.fillStyle = '#a9820c'; x.font = '800 9.5px -apple-system,sans-serif'; x.textAlign='left';
      x.fillText(label.toUpperCase(), PAD+29, y+16);
      x.fillStyle = '#4a3d10'; x.font = 'italic 600 13px -apple-system,sans-serif';
      lines.forEach((ln, i) => x.fillText(ln, PAD+29, y+34 + i*18));
    }
    return y + boxH + 12;
  }

  // ── dibuja toda la pila y devuelve el canvas ──
  function _dibujar(data) {
    const W = 720, PAD = 14, dpr = Math.min(3, window.devicePixelRatio || 2);
    const tickets = data.tickets || [];
    const link = _link(data.tipo, data.idPrincipal);
    // ctx de medición
    const mc = document.createElement('canvas'), mx = mc.getContext('2d');
    const rrNo = () => {}, truncM = (s,w,f) => _trunc(mx, s, w, f), wrapM = (s,w,f) => _wrap(mx, s, w, f);
    let H = PAD;
    tickets.forEach((t, i) => { H = _ticket(mx, t, H, W, PAD, false, rrNo, wrapM, truncM); if (i < tickets.length-1) H += 26; });
    // footer: solo un QR pequeño + una línea corta (el link y las condiciones van en el TEXTO, no en la imagen)
    const footH = 16 + 100 + 22 + 8;
    H += footH + PAD;

    const cv = document.createElement('canvas'); cv.width = W*dpr; cv.height = H*dpr;
    const x = cv.getContext('2d'); x.scale(dpr, dpr); x.textBaseline = 'alphabetic';
    const rr = (px,py,w,h,r) => { r=Math.min(r,h/2,w/2); x.beginPath(); x.moveTo(px+r,py); x.arcTo(px+w,py,px+w,py+h,r); x.arcTo(px+w,py+h,px,py+h,r); x.arcTo(px,py+h,px,py,r); x.arcTo(px,py,px+w,py,r); x.closePath(); };
    const trunc = (s,w,f) => _trunc(x, s, w, f), wrap = (s,w,f) => _wrap(x, s, w, f);
    // fondo stage + tarjeta papel
    x.fillStyle = '#e5e2da'; x.fillRect(0,0,W,H);
    x.fillStyle = '#fdfcf8'; rr(PAD-4, PAD-4, W-2*(PAD-4), H-2*(PAD-4), 18); x.fill();
    x.strokeStyle = '#e7e2d6'; x.lineWidth=1; rr(PAD-4, PAD-4, W-2*(PAD-4), H-2*(PAD-4), 18); x.stroke();

    let y = PAD;
    tickets.forEach((t, i) => {
      y = _ticket(x, t, y, W, PAD, true, rr, wrap, trunc);
      if (i < tickets.length-1) {
        // corte de ticket
        const cy = y + 13;
        x.fillStyle = '#e5e2da';
        x.beginPath(); x.arc(PAD-4, cy, 11, 0, 7); x.fill();
        x.beginPath(); x.arc(W-PAD+4, cy, 11, 0, 7); x.fill();
        x.strokeStyle = '#d9d3c4'; x.lineWidth = 2; x.setLineDash([6,5]);
        x.beginPath(); x.moveTo(PAD+10, cy); x.lineTo(W-PAD-10, cy); x.stroke(); x.setLineDash([]);
        y += 26;
      }
    });

    // footer: QR PEQUEÑO centrado + una línea corta. Nada más (link/condiciones van en el texto).
    y += 6;
    x.strokeStyle = '#e7e2d6'; x.beginPath(); x.moveTo(PAD+16,y); x.lineTo(W-PAD-16,y); x.stroke(); y += 16;
    let qrOk = false;
    try {
      if (typeof qrcode === 'function') {
        const qr = qrcode(0,'M'); qr.addData(link); qr.make();
        const count = qr.getModuleCount(), quiet = 3, total = count + quiet*2;
        const cell = Math.max(2, Math.floor(90/total)), size = cell*total, bx = (W-size)/2, by = y;
        x.fillStyle = '#fff'; rr(bx-6, by-6, size+12, size+12, 9); x.fill();
        x.strokeStyle = '#e7e2d6'; rr(bx-6, by-6, size+12, size+12, 9); x.stroke();
        x.fillStyle = '#14151b';
        for (let r=0;r<count;r++) for (let c=0;c<count;c++) if (qr.isDark(r,c)) x.fillRect(bx+(c+quiet)*cell, by+(r+quiet)*cell, cell, cell);
        y += size + 10;
        x.fillStyle = '#565d6b'; x.font = '700 12.5px -apple-system,Segoe UI,sans-serif'; x.textAlign = 'center';
        x.fillText('📲 Escanéame — reporte en vivo con fotos', W/2, y+4); x.textAlign = 'left';
        y += 18; qrOk = true;
      }
    } catch(_) {}
    if (!qrOk) { x.fillStyle = '#1d4ed8'; x.font = '700 12px ui-monospace,monospace'; x.textAlign='center'; x.fillText(link.replace('https://',''), W/2, y+14); x.textAlign='left'; y += 24; }
    return cv;
  }

  function _trunc(ctx, s, maxW, font) { ctx.font = font; s = String(s||''); if (ctx.measureText(s).width <= maxW) return s;
    while (s.length > 1 && ctx.measureText(s + '…').width > maxW) s = s.slice(0,-1); return s + '…'; }
  function _wrap(ctx, text, maxW, font) { if (font) ctx.font = font; const words = String(text||'').split(' '), lines=[]; let line='';
    words.forEach(w => { const t = line ? line+' '+w : w; if (ctx.measureText(t).width > maxW && line) { lines.push(line); line=w; } else line=t; });
    if (line) lines.push(line); return lines.length ? lines : ['']; }

  function _textoWA(data) {
    const t = (data.tickets||[])[0]; if (!t) return '';
    const isG = t.kind === 'guia';
    const L = ['*🧾 ' + (t.proveedor || '—') + '*'];
    L.push((isG ? _tipoLabel(t.tipoGuia) : '📦 Preingreso') + '  ·  ' + (isG ? t.idGuia : t.idPreingreso));
    L.push('Estado: ' + (t.estado || '—'));
    if (isG) {
      const tg = _tags(t.comentario), m = [];
      if (tg.compl==='si') m.push('✅ Completo'); else if (tg.compl==='no') m.push('⚠️ Incompleto');
      if (tg.comp==='si') m.push('🧾 Con comprobante'); else if (tg.comp==='no') m.push('Sin comprobante');
      if (m.length) L.push(m.join('  ·  '));
      const libre = _libre(t.comentario); if (libre) L.push('📝 ' + libre);
      const det = t.detalle || []; if (det.length) { L.push('', '*Productos (' + det.length + '):*');
        det.forEach(d => L.push('• ' + (d.descripcion || d.codigoProducto) + ' — ' + (d.cantidad!=null?d.cantidad:'') + (d.esProductoNuevo?' 🆕':d.esSinIdentificar?' ⚠️':''))); }
    } else {
      const mon = _money(t.monto); if (mon) L.push('💵 ' + mon);
      const carg = t.cargadores||[]; if (carg.length) L.push('🛺 ' + carg.join(', '));
      const libre = _libre(t.comentario); if (libre) L.push('📝 ' + libre);
    }
    // anexos (conteo)
    const anexos = (data.tickets||[]).slice(1);
    if (anexos.length) { L.push('', '*Anexado:*'); anexos.forEach(a => L.push('🔗 ' + (a.kind==='guia'?a.idGuia:a.idPreingreso) + (a.kind==='preingreso'&&_money(a.monto)?(' · '+_money(a.monto)):''))); }
    // bloque profesional: leer el QR / abrir el link + condiciones (texto plano, acompaña a la imagen)
    L.push('',
      '━━━━━━━━━━━━━━━━━━',
      '📲 *Escanea el código QR* de la imagen, o abre este enlace para ver el reporte:',
      _link(data.tipo, data.idPrincipal),
      '',
      'ℹ️ *Importante:*',
      '• Es un reporte *en tiempo real*: se actualiza solo.',
      '• La información es *un resumen al momento* y puede *cambiar durante el día* (más productos, correcciones o fotos).',
      '• Abre el enlace para ver el estado *actualizado* y todas las *fotos*.',
      (data.generado ? '🕒 Emitido: ' + data.generado : ''),
      '',
      '_InversionMos · warehouseMos_');
    return L.filter(function(l){ return l !== null && l !== undefined; }).join('\n');
  }

  async function compartir(tipo, id) {
    if (typeof toast === 'function') toast('Generando voucher…', 'info', 1500);
    let data;
    try { const r = await API.get('voucherData', { tipo, id }); data = (r && r.ok && r.data) ? r.data : null; }
    catch(_) { data = null; }
    if (!data || !(data.tickets||[]).length) { if (typeof toast === 'function') toast('No se pudo generar el voucher', 'warn'); return; }
    const texto = _textoWA(data);
    let canvas;
    try { canvas = _dibujar(data); } catch(e) { window.open('https://wa.me/?text=' + encodeURIComponent(texto), '_blank'); return; }
    const fname = (data.idPrincipal || 'voucher') + '.png';
    canvas.toBlob(async (blob) => {
      if (!blob) { window.open('https://wa.me/?text=' + encodeURIComponent(texto), '_blank'); return; }
      const file = new File([blob], fname, { type: 'image/png' });
      try { if (navigator.canShare && navigator.canShare({ files:[file] })) { await navigator.share({ files:[file], text: texto }); return; } }
      catch(_) {}
      try { const u = URL.createObjectURL(blob); const a = document.createElement('a'); a.href=u; a.download=fname; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(u),5000); } catch(_){}
      window.open('https://wa.me/?text=' + encodeURIComponent(texto), '_blank');
      if (typeof toast === 'function') toast('Imagen descargada — adjúntala en WhatsApp (el link ya va en el texto)', 'info', 4200);
    }, 'image/png');
  }

  window.Voucher = { compartir, _dibujar, _textoWA };
})();
