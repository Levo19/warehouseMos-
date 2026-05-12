// ============================================================
// warehouseMos — js/mermas.js
// Cesta de mermas: V2 con secciones (pendientes/descartado/solucionado),
// agregar, solucionar (sliders recuperar/descartar), procesar eliminación.
// ============================================================

(function() {
  'use strict';

  let _data = { pendientes: [], descartado: [], solucionado: [] };
  let _producto = null;

  function abrirCesta() {
    // Optimista: abre el sheet INMEDIATO con lo cacheado, refresh en background
    document.getElementById('overlayMermasCesta').style.display = 'block';
    document.getElementById('sheetMermasCesta').classList.add('open');
    // Render con lo que haya (puede estar vacío la primera vez)
    if (_data && (_data.pendientes || _data.descartado || _data.solucionado)) {
      _render();
    } else {
      const cont = document.getElementById('mermaSecPendientes');
      if (cont) cont.innerHTML = '<p class="mc-empty">Cargando…</p>';
    }
    // Fetch en background — al volver, _render() actualiza la UI
    refresh().catch(() => {});
  }

  function cerrarCesta() {
    document.getElementById('overlayMermasCesta').style.display = 'none';
    document.getElementById('sheetMermasCesta').classList.remove('open');
    if (typeof App !== 'undefined' && App.actualizarBadgeMermas) App.actualizarBadgeMermas();
  }

  async function refresh() {
    try {
      const res = await API.get('getMermasCesta');
      if (res && res.ok && res.data) {
        _data = res.data;
        _render();
      }
    } catch(e) { console.warn('cesta mermas:', e.message); }
  }

  function _render() {
    const total = _data.totalPendientes + _data.totalDescartado;
    const totalEl = document.getElementById('mermaCestaTotal');
    if (totalEl) totalEl.textContent = total;

    const secP = document.getElementById('mermaSecPendientes');
    const secD = document.getElementById('mermaSecDescartado');
    const secS = document.getElementById('mermaSecSolucionado');
    if (secP) secP.innerHTML = _data.pendientes.map(_rowHtml.bind(null, 'pendiente')).join('') ||
      '<p class="mc-empty">Sin mermas pendientes.</p>';
    if (secD) secD.innerHTML = _data.descartado.map(_rowHtml.bind(null, 'descartado')).join('') ||
      '<p class="mc-empty">Nada descartado esperando proceso.</p>';
    if (secS) secS.innerHTML = _data.solucionado.map(_rowHtml.bind(null, 'solucionado')).join('') ||
      '<p class="mc-empty">Sin mermas solucionadas (últimos 7 días).</p>';
  }

  function _rowHtml(seccion, m) {
    const cb = String(m.codigoProducto || '');
    const motivo = String(m.motivo || '');
    const zona   = String(m.responsable || m.origen || '');
    const orig = parseFloat(m.cantidadOriginal) || 0;
    const rep  = parseFloat(m.cantidadReparada) || 0;
    const des  = parseFloat(m.cantidadDesechada) || 0;
    const pend = parseFloat(m.cantidadPendiente) || 0;
    const foto = String(m.foto || '');
    const fotoEl = foto ? `<img class="mc-thumb" src="${foto}" onclick="Photos.lightbox('${foto}')"/>` : '';

    if (seccion === 'pendiente') {
      return `<div class="mc-row mc-row-pend">
        ${fotoEl}
        <div class="mc-info">
          <div class="mc-name">${_escHtml(cb)}</div>
          <div class="mc-meta">${pend} pendiente · ${zona ? '📍 '+_escHtml(zona)+' · ' : ''}${_escHtml(motivo)}</div>
        </div>
        <button class="mc-act" onclick="Mermas.abrirSolucionar('${_escAttr(m.idMerma)}')">Solucionar</button>
      </div>`;
    }
    if (seccion === 'descartado') {
      return `<div class="mc-row mc-row-desc">
        ${fotoEl}
        <div class="mc-info">
          <div class="mc-name">${_escHtml(cb)}</div>
          <div class="mc-meta">×${des} para eliminar · ${zona ? '📍 '+_escHtml(zona) : ''}</div>
        </div>
      </div>`;
    }
    // solucionado
    return `<div class="mc-row mc-row-ok">
      ${fotoEl}
      <div class="mc-info">
        <div class="mc-name">✓ ${_escHtml(cb)}</div>
        <div class="mc-meta">${rep} recuperado · ${des} descartado</div>
      </div>
    </div>`;
  }

  // ── Agregar ──
  function abrirAgregar(productoPreseleccionado) {
    _producto = productoPreseleccionado || null;
    document.getElementById('overlayMermaAdd').style.display = 'block';
    document.getElementById('sheetMermaAdd').classList.add('open');
    document.getElementById('mermaAddInfo').textContent = _producto
      ? `${_producto.nombre || _producto.codigoBarra}`
      : 'Escanea o ingresa el código del producto.';
    document.getElementById('mermaAddCantidad').value = '1';
    document.getElementById('mermaAddMotivo').value = '';
    const zonaInput = document.getElementById('mermaAddZona');
    if (zonaInput) zonaInput.value = '';
  }

  function cerrarAgregar() {
    document.getElementById('overlayMermaAdd').style.display = 'none';
    document.getElementById('sheetMermaAdd').classList.remove('open');
    _producto = null;
  }

  async function confirmarAgregar() {
    const cantidad = parseFloat(document.getElementById('mermaAddCantidad').value) || 0;
    const motivo   = document.getElementById('mermaAddMotivo').value.trim();
    const zona     = (document.getElementById('mermaAddZona') || {}).value || '';
    const codigo   = _producto && _producto.codigoBarra ? String(_producto.codigoBarra) : '';
    if (!codigo)   return (typeof toast === 'function' && toast('Selecciona producto', 'warn'));
    if (cantidad <= 0) return (typeof toast === 'function' && toast('Cantidad inválida', 'warn'));
    if (!motivo)   return (typeof toast === 'function' && toast('Motivo requerido', 'warn'));

    // Encolar como op idempotente
    const usuario = (window.App && App.getUsuario && App.getUsuario()) || '';
    if (window.OpLog) {
      OpLog.enqueue({
        tipo: 'MERMA_AGREGAR',
        payload: {
          codigoProducto:   codigo,
          cantidadOriginal: cantidad,
          motivo, zonaResponsable: zona, usuario
        }
      });
      if (typeof SoundFX !== 'undefined' && SoundFX.beep) SoundFX.beep();
      if (typeof toast === 'function') toast('Merma agregada', 'ok');
      cerrarAgregar();
      setTimeout(refresh, 600);
    } else {
      try {
        const res = await API.post('agregarAMermas', {
          codigoProducto:   codigo,
          cantidadOriginal: cantidad,
          motivo, zonaResponsable: zona, usuario
        });
        if (res && res.ok) { cerrarAgregar(); refresh(); }
        else if (typeof toast === 'function') toast('Error: ' + ((res&&res.error)||'?'), 'warn');
      } catch(e) {
        if (typeof toast === 'function') toast('Sin conexión', 'warn');
      }
    }
  }

  // ── Solucionar (sliders) ──
  let _solucionandoMerma = null;

  function abrirSolucionar(idMerma) {
    const m = _data.pendientes.find(x => x.idMerma === idMerma);
    if (!m) return;
    _solucionandoMerma = m;
    const pend = parseFloat(m.cantidadPendiente) || 0;
    document.getElementById('mermaSolNombre').textContent = String(m.codigoProducto || '');
    document.getElementById('mermaSolPendTotal').textContent = pend;
    document.getElementById('mermaSolRecup').value = '0';
    document.getElementById('mermaSolDesc').value  = '0';
    document.getElementById('mermaSolRecupMax').textContent = pend;
    document.getElementById('mermaSolDescMax').textContent  = pend;
    _actualizarSumaSol();
    document.getElementById('overlayMermaSol').style.display = 'block';
    document.getElementById('sheetMermaSol').classList.add('open');
  }

  function cerrarSolucionar() {
    document.getElementById('overlayMermaSol').style.display = 'none';
    document.getElementById('sheetMermaSol').classList.remove('open');
    _solucionandoMerma = null;
  }

  function _actualizarSumaSol() {
    const r = parseFloat(document.getElementById('mermaSolRecup').value) || 0;
    const d = parseFloat(document.getElementById('mermaSolDesc').value)  || 0;
    const orig = parseFloat((_solucionandoMerma && _solucionandoMerma.cantidadPendiente) || 0);
    document.getElementById('mermaSolSumaActual').textContent = r + d;
    document.getElementById('mermaSolSumaMax').textContent = orig;
    const ok = (r + d) > 0 && (r + d) <= orig;
    document.getElementById('mermaSolOk').disabled = !ok;
  }

  async function confirmarSolucionar() {
    if (!_solucionandoMerma) return;
    const r = parseFloat(document.getElementById('mermaSolRecup').value) || 0;
    const d = parseFloat(document.getElementById('mermaSolDesc').value)  || 0;
    const obs = (document.getElementById('mermaSolObs') || {}).value || '';
    const usuario = (window.App && App.getUsuario && App.getUsuario()) || '';
    const idMerma = _solucionandoMerma.idMerma;
    if (window.OpLog) {
      OpLog.enqueue({
        tipo: 'MERMA_SOLUCIONAR',
        payload: { idMerma, deltaRecuperado: r, deltaDescartado: d, observacion: obs, usuario }
      });
      if (typeof SoundFX !== 'undefined' && SoundFX.beep) SoundFX.beep();
      cerrarSolucionar();
      setTimeout(refresh, 600);
    } else {
      try {
        const res = await API.post('solucionarMerma', {
          idMerma, deltaRecuperado: r, deltaDescartado: d, observacion: obs, usuario
        });
        if (res && res.ok) { cerrarSolucionar(); refresh(); }
      } catch(e) {}
    }
  }

  // ── Procesar eliminación (requiere clave admin) ──
  async function procesarEliminacion(claveAdmin) {
    const usuario = (window.App && App.getUsuario && App.getUsuario()) || '';
    try {
      const res = await API.post('procesarEliminacionMermas', { claveAdmin, usuario });
      if (res && res.ok) {
        if (typeof toast === 'function') toast(`Guía ${res.data.idGuiaSalida} generada · ${res.data.procesados} procesados`, 'ok');
        if (typeof SoundFX !== 'undefined' && SoundFX.done) SoundFX.done();
        refresh();
        return res.data;
      } else {
        if (typeof toast === 'function') toast('Error: ' + ((res && res.error) || '?'), 'warn');
      }
    } catch(e) {
      if (typeof toast === 'function') toast('Sin conexión', 'warn');
    }
  }

  async function refreshBadge() {
    try {
      const res = await API.get('contadorMermasPendientes');
      if (res && res.ok) {
        const n = (res.data && res.data.count) || 0;
        const el = document.getElementById('topMermasBadge');
        if (el) {
          el.textContent = n;
          el.style.display = n > 0 ? 'inline-block' : 'none';
        }
        return n;
      }
    } catch(e) {}
    return 0;
  }

  function _bindInputs() {
    ['mermaSolRecup', 'mermaSolDesc'].forEach(id => {
      const el = document.getElementById(id);
      if (el && !el._bound) {
        el._bound = true;
        el.addEventListener('input', _actualizarSumaSol);
      }
    });
  }

  function _escHtml(s) { return String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }
  function _escAttr(s) { return String(s||'').replace(/'/g, '&#39;'); }

  // Polling badge cada 90s
  setInterval(() => { refreshBadge(); }, 90000);
  // Bind inputs cuando el DOM esté listo
  if (document.readyState !== 'loading') setTimeout(_bindInputs, 500);
  else document.addEventListener('DOMContentLoaded', () => setTimeout(_bindInputs, 500));

  window.Mermas = {
    abrirCesta, cerrarCesta, refresh,
    abrirAgregar, cerrarAgregar, confirmarAgregar,
    abrirSolucionar, cerrarSolucionar, confirmarSolucionar,
    procesarEliminacion, refreshBadge
  };
})();
