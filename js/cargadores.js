// ============================================================
// warehouseMos — js/cargadores.js  ·  v2 (termómetro + fotos)
// Cargadores INDEPENDIENTES del preingreso. Cada cargador del día tiene un
// NIVEL de carga (termómetro táctil 0-100%, color por tramo + sonido) y sus
// FOTOS (a Supabase Storage, con carrusel). Modelo: 1 fila por (fecha,cargador)
// en wh.cargadores_log (RPCs cargador_dia_* — SQL 557).
// ============================================================

(function() {
  'use strict';

  let _master = [];      // catálogo de cargadores (proveedores prefijo CARGADOR)
  let _dia = [];         // [{idCargador, nombre, nivel, fotos:[], count}]
  let _fechaActual = '';
  let _polling = null;
  let _cssDone = false;
  let _audio = null;

  const MIN_GUARDAR = 10;   // % mínimo para considerar la carga "registrada"

  function _hoyStr() {
    try { return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date()); }
    catch (e) { return new Date(Date.now() - 5 * 3600000).toISOString().slice(0, 10); }
  }
  function _escHtml(s) { return String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }
  function _escAttr(s) { return String(s||'').replace(/'/g, '&#39;').replace(/"/g, '&quot;'); }
  function _usuario(){ return (window.App && App.getUsuario && App.getUsuario()) || ''; }
  function _deviceId(){ return localStorage.getItem('wh_device_id') || ''; }

  // ── escala del termómetro ──
  function _thColor(p){ return p >= 75 ? '#10b981' : p >= 50 ? '#fbbf24' : p >= 25 ? '#f97316' : '#ef4444'; }
  function _thColorT(p){ return p >= 75 ? '#6ee7b7' : p >= 50 ? '#fcd34d' : p >= 25 ? '#fdba74' : '#fca5a5'; }
  function _thWord(p){ return p >= 100 ? '🔥 Recontra lleno' : p >= 75 ? 'Lleno' : p >= 50 ? 'Media carga' : p >= 25 ? 'Poca carga' : 'Casi vacío'; }
  function _tramo(p){ return p >= 100 ? 4 : p >= 75 ? 3 : p >= 50 ? 2 : p >= 25 ? 1 : 0; }

  function _tono(tramo){
    try {
      if (!_audio) _audio = new (window.AudioContext || window.webkitAudioContext)();
      const ctx = _audio, now = ctx.currentTime;
      const freq = [360, 480, 620, 820, 1040][Math.max(0, Math.min(4, tramo))];
      const o = ctx.createOscillator(), g = ctx.createGain();
      o.type = 'sine'; o.frequency.setValueAtTime(freq, now);
      g.gain.setValueAtTime(0.0001, now);
      g.gain.exponentialRampToValueAtTime(0.09, now + 0.012);
      g.gain.exponentialRampToValueAtTime(0.001, now + 0.16);
      o.connect(g).connect(ctx.destination); o.start(now); o.stop(now + 0.18);
    } catch(_){}
    try { if (navigator.vibrate) navigator.vibrate(tramo >= 4 ? [15,30,15] : 12); } catch(_){}
  }

  function _injectCSS(){
    if (_cssDone) return; _cssDone = true;
    const st = document.createElement('style'); st.id = 'cargV2CSS';
    st.textContent = `
      .cgv-card{background:#1a1505;border:1px solid #854d0e;border-radius:14px;padding:12px 13px;margin-bottom:11px;animation:cgvIn .4s cubic-bezier(.34,1.4,.5,1)}
      @keyframes cgvIn{from{opacity:0;transform:translateY(-6px) scale(.98)}to{opacity:1;transform:none}}
      .cgv-head{display:flex;align-items:center;gap:8px;margin-bottom:11px}
      .cgv-ico{font-size:16px} .cgv-name{font-size:13.5px;font-weight:800;color:#fcd34d;flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .cgv-x{color:#5f7192;font-size:18px;background:transparent;border:none;cursor:pointer;width:24px;height:24px}
      .cgv-x:hover{color:#f87171}
      .cgv-th-top{display:flex;align-items:baseline;gap:8px;margin-bottom:7px}
      .cgv-pct{font-size:24px;font-weight:900;font-family:ui-monospace,monospace;line-height:1;letter-spacing:-.02em}
      .cgv-word{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}
      .cgv-track{position:relative;height:30px;border-radius:999px;background:#0a0f18;border:1px solid #854d0e;cursor:ew-resize;touch-action:none;box-shadow:inset 0 2px 6px rgba(0,0,0,.5);user-select:none}
      .cgv-fill{position:absolute;left:0;top:0;bottom:0;border-radius:999px;background:linear-gradient(90deg,#ef4444,#f97316,#fbbf24,#10b981);background-size:600% 100%;transition:width .12s linear}
      .cgv-fill::after{content:'';position:absolute;inset:0;border-radius:999px;background:linear-gradient(180deg,rgba(255,255,255,.26),transparent 45%)}
      .cgv-ticks i{position:absolute;top:6px;bottom:6px;width:1.5px;background:rgba(255,255,255,.14);pointer-events:none}
      .cgv-min{position:absolute;top:-4px;bottom:-4px;left:10%;width:2px;pointer-events:none;background:repeating-linear-gradient(180deg,rgba(239,68,68,.7) 0 4px,transparent 4px 7px)}
      .cgv-handle{position:absolute;top:50%;width:26px;height:26px;border-radius:50%;background:#fff;box-shadow:0 3px 10px rgba(0,0,0,.55),0 0 0 3px rgba(0,0,0,.22);transform:translate(-50%,-50%);pointer-events:none;display:grid;place-items:center}
      .cgv-handle::before{content:'⋮⋮';font-size:11px;color:#334155;letter-spacing:-2px;font-weight:900}
      .cgv-scale{display:flex;justify-content:space-between;font-size:9px;color:#5f7192;font-family:ui-monospace,monospace;margin-top:5px;padding:0 2px}
      .cgv-warn{font-size:10px;color:#fca5a5;margin-top:5px;font-weight:600}
      .cgv-fotos{margin-top:12px;border-top:1px dashed #854d0e;padding-top:10px}
      .cgv-fotos-lbl{font-size:9px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#fbbf24;margin-bottom:7px;display:flex;align-items:center;gap:6px}
      .cgv-fotos-lbl .n{background:rgba(245,158,11,.14);border:1px solid rgba(245,158,11,.3);border-radius:999px;padding:0 6px;font-size:9px}
      .cgv-fotos-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
      .cgv-th-img{width:56px;height:56px;border-radius:11px;overflow:hidden;border:1px solid #854d0e;cursor:pointer;background:#0a0f18;position:relative}
      .cgv-th-img img{width:100%;height:100%;object-fit:cover;display:block}
      .cgv-th-more{width:56px;height:56px;border-radius:11px;border:1px solid #854d0e;background:rgba(252,211,77,.08);color:#fcd34d;font-size:13px;font-weight:800;display:flex;align-items:center;justify-content:center;cursor:pointer}
      .cgv-foto-add{width:56px;height:56px;border-radius:11px;border:2px dashed #854d0e;background:rgba(252,211,77,.04);color:#fbbf24;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1px;cursor:pointer}
      .cgv-foto-add:hover{border-color:#fcd34d;background:rgba(252,211,77,.08)}
      .cgv-foto-add .cam{font-size:20px;line-height:1}.cgv-foto-add small{font-size:8px;font-weight:800}
      .cgv-foto-add.busy{opacity:.6;pointer-events:none}
      .cgv-foto-none{font-size:10px;color:#5f7192;font-style:italic;flex:1}
      /* carrusel */
      .cgv-carousel{position:fixed;inset:0;z-index:12000;background:rgba(3,6,12,.92);display:flex;flex-direction:column;justify-content:center;align-items:center;padding:14px}
      .cgv-car-box{width:100%;max-width:460px;background:#05080e;border:1px solid #243247;border-radius:18px;overflow:hidden}
      .cgv-car-top{padding:11px 14px;display:flex;align-items:center;gap:9px;border-bottom:1px solid #1a2740}
      .cgv-car-top .t{font-weight:800;font-size:13px;color:#fcd34d;flex:1}
      .cgv-car-top .x{width:30px;height:30px;border-radius:50%;background:#131d30;border:1px solid #243247;color:#93a4c2;font-size:14px;cursor:pointer}
      .cgv-car-stage{position:relative;aspect-ratio:4/3;background:#0a0f18}
      .cgv-car-stage img{width:100%;height:100%;object-fit:contain}
      .cgv-car-nav{position:absolute;top:50%;transform:translateY(-50%);width:36px;height:36px;border-radius:50%;background:rgba(10,15,24,.8);border:1px solid #243247;color:#fff;font-size:18px;cursor:pointer}
      .cgv-car-nav.l{left:10px}.cgv-car-nav.r{right:10px}
      .cgv-car-count{position:absolute;top:10px;right:12px;font-size:10px;font-weight:800;font-family:ui-monospace,monospace;color:#fff;background:rgba(10,15,24,.7);border-radius:999px;padding:3px 9px}
      .cgv-car-dots{display:flex;gap:6px;justify-content:center;padding:11px}
      .cgv-car-dots i{width:7px;height:7px;border-radius:50%;background:#243247;transition:.2s}.cgv-car-dots i.on{background:#fcd34d;width:20px;border-radius:99px}
      @media(prefers-reduced-motion:reduce){.cgv-card,.cgv-fill{animation:none;transition:none}}
    `;
    document.head.appendChild(st);
  }

  async function _cargarMaster() {
    if (_master.length) return _master;
    try { const res = await API.get('listarCargadoresMaster'); _master = (res && res.ok && res.data) ? res.data : []; }
    catch(e){ _master = []; }
    return _master;
  }
  async function _cargarResumen(fecha) {
    fecha = fecha || _hoyStr();
    try {
      const res = await API.get('getResumenCargadoresDia', { fecha });
      if (res && res.ok && res.data) { _dia = res.data.cargadores || []; _fechaActual = res.data.fecha || fecha; return res.data; }
    } catch(e){}
    _dia = []; _fechaActual = fecha;
    return { fecha, total: 0, cargadores: [] };
  }

  function abrir(fecha) {
    _injectCSS();
    fecha = fecha || _hoyStr();
    document.getElementById('overlayCargadores').style.display = 'block';
    document.getElementById('modalCargadores').classList.add('open');
    const inp = document.getElementById('cargBuscarInput');
    if (inp) { inp.value = ''; setTimeout(() => inp.focus(), 100); }
    if (_master.length || _dia.length) { _render(); _filtrar(''); }
    else { const l = document.getElementById('cargCoincidencias'); if (l) l.innerHTML = '<p style="color:#64748b;font-size:13px;padding:14px 0;text-align:center">Cargando…</p>'; }
    Promise.all([_cargarMaster(), _cargarResumen(fecha)]).then(() => { _render(); _filtrar(inp ? inp.value : ''); }).catch(()=>{});
  }
  function cerrar() {
    document.getElementById('overlayCargadores').style.display = 'none';
    document.getElementById('modalCargadores').classList.remove('open');
    if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
  }

  // ── render de las tarjetas (termómetro + fotos) ──
  function _cardHTML(c) {
    const id = _escAttr(c.idCargador);
    const p = Math.max(0, Math.min(100, parseInt(c.nivel) || 0));
    const fotos = Array.isArray(c.fotos) ? c.fotos : [];
    const col = _thColor(p), colT = _thColorT(p);
    const thumbs = fotos.slice(0, 2).map((u, i) => `<div class="cgv-th-img" onclick="Cargadores._carousel('${id}',${i})"><img src="${_escAttr(u)}" loading="lazy" onerror="this.parentNode.style.display='none'"></div>`).join('');
    const more = fotos.length > 2 ? `<div class="cgv-th-more" onclick="Cargadores._carousel('${id}',2)">+${fotos.length - 2}</div>` : '';
    return `<div class="cgv-card" id="cgvCard_${id}" data-id="${id}">
      <div class="cgv-head"><span class="cgv-ico">🛺</span><span class="cgv-name">${_escHtml(c.nombre || c.idCargador)}</span>
        <button class="cgv-x" onclick="Cargadores.quitar('${id}')" title="Quitar cargador">✕</button></div>
      <div class="cgv-th">
        <div class="cgv-th-top"><span class="cgv-pct" id="cgvPct_${id}" style="color:${col}">${p}%</span><span class="cgv-word" id="cgvWord_${id}" style="color:${colT}">${_thWord(p)}</span></div>
        <div class="cgv-track" id="cgvTrack_${id}" data-id="${id}">
          <div class="cgv-fill" id="cgvFill_${id}" style="width:${p}%;background-position:${p ? (100 - (100/Math.max(p,1))*100) : 0}% 0"></div>
          <div class="cgv-ticks"><i style="left:25%"></i><i style="left:50%"></i><i style="left:75%"></i></div>
          <div class="cgv-min"></div>
          <div class="cgv-handle" id="cgvHandle_${id}" style="left:${p}%"></div>
        </div>
        <div class="cgv-scale"><span>0</span><span>25</span><span>50</span><span>75</span><span>100</span></div>
        <div class="cgv-warn" id="cgvWarn_${id}" style="${p >= MIN_GUARDAR ? 'display:none' : ''}">⚠ Jala la barra — mínimo ${MIN_GUARDAR}% para registrar la carga.</div>
      </div>
      <div class="cgv-fotos">
        <div class="cgv-fotos-lbl">📷 Fotos <span class="n">${fotos.length}</span></div>
        <div class="cgv-fotos-row">
          ${thumbs}${more}
          <div class="cgv-foto-add" id="cgvAdd_${id}" onclick="Cargadores._pickFoto('${id}')"><span class="cam">+</span><small>foto</small></div>
          ${fotos.length ? '' : '<span class="cgv-foto-none">Sin fotos · agrega evidencia</span>'}
        </div>
      </div>
    </div>`;
  }

  function _render() {
    const totalEl = document.getElementById('cargTotal');
    if (totalEl) totalEl.textContent = _dia.length;
    const resEl = document.getElementById('cargResumen');
    if (!resEl) return;
    if (!_dia.length) { resEl.innerHTML = '<p style="color:#64748b;font-size:13px;text-align:center;padding:18px 0">Sin cargadores hoy. Busca uno arriba y agrégalo.</p>'; return; }
    resEl.innerHTML = _dia.map(_cardHTML).join('');
    _bindDrag();
  }

  function _filtrar(query) {
    query = String(query || '').trim().toLowerCase();
    const listEl = document.getElementById('cargCoincidencias');
    if (!listEl) return;
    let cand = _master;
    if (query) cand = _master.filter(c => String(c.nombre||'').toLowerCase().includes(query) || String(c.idCargador||'').toLowerCase().includes(query));
    const yaHoy = new Set(_dia.map(c => String(c.idCargador)));
    if (!cand.length) { listEl.innerHTML = '<p style="color:#64748b;font-size:13px;padding:14px 0;text-align:center">Sin coincidencias.</p>'; return; }
    listEl.innerHTML = cand.slice(0, 12).map(c => {
      const ya = yaHoy.has(String(c.idCargador));
      return `<button class="carg-match" ${ya ? 'style="opacity:.5"' : ''} onclick="Cargadores.agregar('${_escAttr(c.idCargador)}','${_escAttr(c.nombre)}')">
        <span class="carg-match-name">${_escHtml(c.nombre)}</span>
        <span class="carg-match-add">${ya ? '✓ hoy' : '+ agregar'}</span></button>`;
    }).join('');
  }

  // ── agregar cargador (upsert nivel 0) ──
  async function agregar(idCargador, nombre) {
    if (_dia.some(c => String(c.idCargador) === String(idCargador))) {   // ya está hoy → scroll a su card
      const el = document.getElementById('cgvCard_' + _escAttr(idCargador));
      if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); el.style.boxShadow = '0 0 0 3px rgba(252,211,77,.4)'; setTimeout(() => el.style.boxShadow = '', 900); }
      return;
    }
    // optimista
    _dia.unshift({ idCargador, nombre, nivel: 0, fotos: [], count: 1 });
    _render(); _filtrar(document.getElementById('cargBuscarInput')?.value || '');
    _tono(0);
    try {
      const res = await API.post('cargadorDiaUpsert', { idCargador, nombre, fecha: _fechaActual || _hoyStr(), usuario: _usuario(), deviceId: _deviceId() });
      if (!res || res.ok === false) { if (typeof toast === 'function') toast('No se pudo agregar: ' + ((res && res.error) || 'error'), 'warn'); }
    } catch(e) { if (typeof toast === 'function') toast('Sin conexión · reintenta', 'warn'); }
  }

  async function quitar(idCargador) {
    const nombre = (_dia.find(c => String(c.idCargador) === String(idCargador)) || {}).nombre || 'cargador';
    if (window.App && App.confirmar) { const ok = await App.confirmar('¿Quitar a ' + nombre + ' de hoy?'); if (!ok) return; }
    _dia = _dia.filter(c => String(c.idCargador) !== String(idCargador));
    _render(); _filtrar(document.getElementById('cargBuscarInput')?.value || '');
    if (navigator.vibrate) navigator.vibrate([10, 30, 10]);
    try { await API.post('cargadorDiaQuitar', { idCargador, fecha: _fechaActual || _hoyStr() }); }
    catch(e){ if (typeof toast === 'function') toast('No se pudo quitar — reintenta', 'warn'); }
  }

  // ── termómetro: drag por puntero ──
  let _drag = null;   // { id, track }
  function _pctFromEvent(track, clientX) {
    const r = track.getBoundingClientRect();
    return Math.max(0, Math.min(100, Math.round(((clientX - r.left) / Math.max(1, r.width)) * 100)));
  }
  function _paintNivel(id, p) {
    p = Math.max(0, Math.min(100, Math.round(p)));
    const fill = document.getElementById('cgvFill_' + id), handle = document.getElementById('cgvHandle_' + id),
          pct = document.getElementById('cgvPct_' + id), word = document.getElementById('cgvWord_' + id),
          warn = document.getElementById('cgvWarn_' + id);
    const col = _thColor(p), colT = _thColorT(p);
    if (fill) { fill.style.width = p + '%'; fill.style.backgroundPosition = (p ? (100 - (100 / Math.max(p, 1)) * 100) : 0) + '% 0'; if (p >= 75) fill.style.boxShadow = '0 0 14px -2px rgba(16,185,129,.55)'; else fill.style.boxShadow = ''; }
    if (handle) handle.style.left = p + '%';
    if (pct) { pct.textContent = p + '%'; pct.style.color = col; }
    if (word) { word.textContent = _thWord(p); word.style.color = colT; }
    if (warn) warn.style.display = p >= MIN_GUARDAR ? 'none' : '';
  }
  function _bindDrag() {
    document.querySelectorAll('.cgv-track').forEach(track => {
      if (track._bound) return; track._bound = true;
      track.addEventListener('pointerdown', ev => {
        ev.preventDefault();
        const id = track.getAttribute('data-id');
        _drag = { id, track, lastTramo: _tramo(parseInt(document.getElementById('cgvPct_' + id)?.textContent) || 0) };
        try { track.setPointerCapture(ev.pointerId); } catch(_){}
        const p = _pctFromEvent(track, ev.clientX);
        _onDragMove(p);
      });
      track.addEventListener('pointermove', ev => { if (_drag && _drag.track === track) _onDragMove(_pctFromEvent(track, ev.clientX)); });
      const end = () => { if (_drag && _drag.track === track) _endDrag(); };
      track.addEventListener('pointerup', end);
      track.addEventListener('pointercancel', end);
    });
  }
  function _onDragMove(p) {
    if (!_drag) return;
    _paintNivel(_drag.id, p);
    const t = _tramo(p);
    if (t !== _drag.lastTramo) { _tono(t); _drag.lastTramo = t; }
  }
  let _saveTimers = {};
  function _endDrag() {
    if (!_drag) return;
    const id = _drag.id, p = Math.max(0, Math.min(100, Math.round(parseInt(document.getElementById('cgvPct_' + id)?.textContent) || 0)));
    _drag = null;
    const c = _dia.find(x => String(x.idCargador) === String(id)); if (c) c.nivel = p;
    // persistir (debounce por cargador)
    clearTimeout(_saveTimers[id]);
    _saveTimers[id] = setTimeout(async () => {
      try { const res = await API.post('cargadorDiaSetNivel', { idCargador: id, nivel: p, fecha: _fechaActual || _hoyStr(), nombre: (c && c.nombre) || '', usuario: _usuario(), deviceId: _deviceId() });
        if (!res || res.ok === false) { if (typeof toast === 'function') toast('No se guardó el nivel — reintenta', 'warn'); }
      } catch(e){ if (typeof toast === 'function') toast('Sin conexión · nivel no guardado', 'warn'); }
    }, 350);
  }

  // ── fotos ──
  function _pickFoto(idCargador) {
    let inp = document.getElementById('cgvFileInput');
    if (!inp) { inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/*'; inp.capture = 'environment'; inp.id = 'cgvFileInput'; inp.style.display = 'none'; document.body.appendChild(inp); }
    inp.value = '';
    inp.onchange = (ev) => { const f = ev.target.files && ev.target.files[0]; if (f) _subirFoto(idCargador, f); };
    inp.click();
  }
  async function _subirFoto(idCargador, file) {
    if (file.size > 6 * 1024 * 1024) { if (typeof toast === 'function') toast('Foto muy pesada (máx 6MB)', 'warn'); return; }
    const add = document.getElementById('cgvAdd_' + _escAttr(idCargador)); if (add) add.classList.add('busy');
    const b64 = await new Promise(res => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.readAsDataURL(file); });
    try {
      const res = await API.post('cargadorDiaAddFoto', { idCargador, fecha: _fechaActual || _hoyStr(), fotoBase64: b64, mimeType: file.type || 'image/jpeg', nombre: (_dia.find(c => String(c.idCargador) === String(idCargador)) || {}).nombre || '', usuario: _usuario(), deviceId: _deviceId() });
      if (res && res.ok && res.data) {
        const c = _dia.find(x => String(x.idCargador) === String(idCargador));
        if (c) { c.fotos = res.data.fotos || c.fotos; _actualizarCardFotos(idCargador); }
        if (navigator.vibrate) navigator.vibrate(12);
      } else { if (typeof toast === 'function') toast('No se subió la foto: ' + ((res && res.error) || '?'), 'warn'); }
    } catch(e){ if (typeof toast === 'function') toast('No se subió la foto — reintenta', 'warn'); }
    finally { if (add) add.classList.remove('busy'); }
  }
  function _actualizarCardFotos(idCargador) {
    const card = document.getElementById('cgvCard_' + _escAttr(idCargador)); if (!card) return;
    const c = _dia.find(x => String(x.idCargador) === String(idCargador)); if (!c) return;
    const wrap = document.createElement('div'); wrap.innerHTML = _cardHTML(c);
    const nuevo = wrap.querySelector('.cgv-fotos');
    const viejo = card.querySelector('.cgv-fotos');
    if (nuevo && viejo) viejo.replaceWith(nuevo);
  }

  // ── carrusel ──
  function _carousel(idCargador, startIdx) {
    const c = _dia.find(x => String(x.idCargador) === String(idCargador)); if (!c) return;
    const fotos = Array.isArray(c.fotos) ? c.fotos : []; if (!fotos.length) return;
    let idx = Math.max(0, Math.min(fotos.length - 1, parseInt(startIdx) || 0));
    let ov = document.getElementById('cgvCarousel');
    if (ov) ov.remove();
    ov = document.createElement('div'); ov.id = 'cgvCarousel'; ov.className = 'cgv-carousel';
    ov.addEventListener('click', e => { if (e.target === ov) ov.remove(); });
    const paint = () => {
      ov.innerHTML = `<div class="cgv-car-box">
        <div class="cgv-car-top"><span style="font-size:16px">🛺</span><span class="t">${_escHtml(c.nombre || c.idCargador)}</span><button class="cgv-car-x x" onclick="document.getElementById('cgvCarousel').remove()">✕</button></div>
        <div class="cgv-car-stage">
          <img src="${_escAttr(fotos[idx])}" alt="">
          <span class="cgv-car-count">${idx + 1} / ${fotos.length}</span>
          ${fotos.length > 1 ? `<button class="cgv-car-nav l" data-nav="-1">‹</button><button class="cgv-car-nav r" data-nav="1">›</button>` : ''}
        </div>
        ${fotos.length > 1 ? `<div class="cgv-car-dots">${fotos.map((_, i) => `<i class="${i === idx ? 'on' : ''}"></i>`).join('')}</div>` : '<div style="height:10px"></div>'}
      </div>`;
      ov.querySelectorAll('[data-nav]').forEach(b => b.onclick = () => { idx = (idx + parseInt(b.getAttribute('data-nav')) + fotos.length) % fotos.length; paint(); });
    };
    paint();
    document.body.appendChild(ov);
  }

  function startPolling() {
    if (_polling) return;
    _polling = setInterval(() => {
      if (document.hidden) return;
      _cargarResumen(_hoyStr()).then(() => { if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia(); });
    }, 60000);
  }
  function getCountDia(fecha) { fecha = fecha || _hoyStr(); return (_fechaActual === fecha) ? _dia.length : 0; }
  async function refreshCountDia(fecha) { const r = await _cargarResumen(fecha || _hoyStr()); return (r.cargadores || []).length; }

  window.Cargadores = {
    abrir, cerrar, agregar, quitar,
    startPolling, getCountDia, refreshCountDia,
    _filtrar, _hoyStr, _pickFoto, _carousel
  };
})();
