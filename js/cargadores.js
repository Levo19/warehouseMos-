// ============================================================
// warehouseMos — js/cargadores.js
// Sistema cargadores independiente de preingresos.
// Modal moderno: buscador + lista de coincidencias + resumen
// abajo con contadores. Cada tap agrega +1; cada [-] resta 1.
// ============================================================

(function() {
  'use strict';

  let _master = [];     // catálogo cargadores filtrado prefijo CARGADOR
  let _resumen = [];    // resumen del día actual
  let _fechaActual = '';
  let _polling = null;

  function _hoyStr() {
    const d = new Date();
    const z = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${z(d.getMonth()+1)}-${z(d.getDate())}`;
  }

  async function _cargarMaster() {
    if (_master.length) return _master;
    try {
      const res = await API.get('listarCargadoresMaster');
      _master = (res && res.ok && res.data) ? res.data : [];
    } catch(e) {
      console.warn('cargadores master:', e.message);
      _master = [];
    }
    return _master;
  }

  async function _cargarResumen(fecha) {
    fecha = fecha || _hoyStr();
    try {
      const res = await API.get('getResumenCargadoresDia', { fecha });
      if (res && res.ok && res.data) {
        _resumen = res.data.cargadores || [];
        _fechaActual = res.data.fecha || fecha;
        return res.data;
      }
    } catch(e) { console.warn('resumen cargadores:', e.message); }
    _resumen = [];
    _fechaActual = fecha;
    return { fecha, total: 0, cargadores: [] };
  }

  function abrir(fecha) {
    fecha = fecha || _hoyStr();
    // Optimista: abre el modal INMEDIATO con lo cacheado
    document.getElementById('overlayCargadores').style.display = 'block';
    document.getElementById('modalCargadores').classList.add('open');
    const inp = document.getElementById('cargBuscarInput');
    if (inp) { inp.value = ''; setTimeout(() => inp.focus(), 100); }
    // Render con cache si hay; placeholder si no
    if (_master.length || _resumen.length) {
      _render();
      _filtrar('');
    } else {
      const listEl = document.getElementById('cargCoincidencias');
      if (listEl) listEl.innerHTML = '<p style="color:#64748b;font-size:13px;padding:14px 0;text-align:center">Cargando cargadores…</p>';
    }
    // Fetch en background, refresca UI cuando llega
    Promise.all([_cargarMaster(), _cargarResumen(fecha)]).then(() => {
      _render();
      _filtrar(inp ? inp.value : '');
    }).catch(() => {});
  }

  function cerrar() {
    document.getElementById('overlayCargadores').style.display = 'none';
    document.getElementById('modalCargadores').classList.remove('open');
    if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
  }

  function _render() {
    const total = _resumen.reduce((s, c) => s + c.count, 0);
    const totalEl = document.getElementById('cargTotal');
    if (totalEl) totalEl.textContent = total;

    const resEl = document.getElementById('cargResumen');
    if (resEl) {
      if (!_resumen.length) {
        resEl.innerHTML = '<p style="color:#64748b;font-size:13px;text-align:center;padding:18px 0">Sin cargadores aún hoy.</p>';
      } else {
        resEl.innerHTML = _resumen.map(c => `
          <div class="carg-row">
            <span class="carg-name">${_escHtml(c.nombre || c.idCargador)}</span>
            <span class="carg-count">×${c.count}</span>
            <button class="carg-minus" onclick="Cargadores.quitar('${_escAttr(c.idCargador)}')" title="Quitar uno">−</button>
          </div>
        `).join('');
      }
    }
  }

  function _filtrar(query) {
    query = String(query || '').trim().toLowerCase();
    const listEl = document.getElementById('cargCoincidencias');
    if (!listEl) return;
    let candidatos = _master;
    if (query) {
      candidatos = _master.filter(c =>
        String(c.nombre || '').toLowerCase().includes(query) ||
        String(c.idCargador || '').toLowerCase().includes(query)
      );
    }
    if (!candidatos.length) {
      listEl.innerHTML = '<p style="color:#64748b;font-size:13px;padding:14px 0;text-align:center">Sin coincidencias.</p>';
      return;
    }
    listEl.innerHTML = candidatos.slice(0, 12).map(c => `
      <button class="carg-match" onclick="Cargadores.agregar('${_escAttr(c.idCargador)}', '${_escAttr(c.nombre)}')">
        <span class="carg-match-name">${_escHtml(c.nombre)}</span>
        <span class="carg-match-add">+1</span>
      </button>
    `).join('');
  }

  async function agregar(idCargador, nombre) {
    const usuario = (window.App && App.getUsuario && App.getUsuario()) || '';
    try {
      const res = await API.post('addCargadorDia', {
        idCargador, nombre, fecha: _fechaActual || _hoyStr(),
        usuario, deviceId: (window.OpLog && OpLog.deviceId()) || ''
      });
      if (res && res.ok) {
        if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click();
        if (navigator.vibrate) navigator.vibrate(8);
        await _cargarResumen(_fechaActual || _hoyStr());
        _render();
      } else {
        if (typeof toast === 'function') toast('No se pudo agregar: ' + ((res && res.error) || 'error'), 'warn');
      }
    } catch(e) {
      if (typeof toast === 'function') toast('Sin conexión · cargador no guardado', 'warn');
    }
  }

  async function quitar(idCargador) {
    try {
      const res = await API.post('removeCargadorDia', {
        idCargador, fecha: _fechaActual || _hoyStr()
      });
      if (res && res.ok) {
        if (typeof SoundFX !== 'undefined' && SoundFX.click) SoundFX.click();
        if (navigator.vibrate) navigator.vibrate(8);
        await _cargarResumen(_fechaActual || _hoyStr());
        _render();
      }
    } catch(e) {}
  }

  function startPolling() {
    if (_polling) return;
    _polling = setInterval(() => {
      _cargarResumen(_hoyStr()).then(() => {
        if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
      });
    }, 60000);
  }

  function getCountDia(fecha) {
    fecha = fecha || _hoyStr();
    if (_fechaActual === fecha) {
      return _resumen.reduce((s, c) => s + c.count, 0);
    }
    return 0;
  }

  async function refreshCountDia(fecha) {
    const r = await _cargarResumen(fecha || _hoyStr());
    return r.total || 0;
  }

  function _escHtml(s) { return String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }
  function _escAttr(s) { return String(s||'').replace(/'/g, '&#39;'); }

  window.Cargadores = {
    abrir, cerrar, agregar, quitar,
    startPolling, getCountDia, refreshCountDia,
    _filtrar,
    _hoyStr
  };
})();
