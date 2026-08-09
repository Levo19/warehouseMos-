// ============================================================
// warehouseMos — js/cargadores.js  ·  v3 (CARGAS = log de eventos)
// Un cargador puede hacer VARIAS cargas al día; cada carga tiene su propio
// termómetro (nivel %), sus fotos (Storage) y su hora. El PROTAGONISTA es el
// resumen del día agrupado por cargador. El "agregar" se colapsa a un botón
// "+ Carga nueva" (top-3 los que más traen + buscador). Modelo: wh.cargadores_log
// (1 fila = 1 carga, id_log = idCarga único) — RPCs cargador_carga_* (SQL 562/563).
// ============================================================

(function() {
  'use strict';

  let _master = [];      // catálogo [{idCargador, nombre, veces}]
  let _dia = [];         // cargas: [{idCarga, idCargador, nombre, nivel, fotos:[], hora, ts, _prov}]
  let _fechaActual = '';
  let _cssDone = false;
  let _audio = null;
  let _addOpen = false;  // panel "agregar" desplegado
  let _searchVal = '';
  let _generado = '';    // hora de emisión del último resumen (server)

  const MIN_GUARDAR = 10;   // % mínimo para considerar la carga "registrada"

  function _hoyStr() {
    try { return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Lima', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date()); }
    catch (e) { return new Date(Date.now() - 5 * 3600000).toISOString().slice(0, 10); }
  }
  function _horaAhora() {
    try { return new Intl.DateTimeFormat('es-PE', { timeZone: 'America/Lima', hour: '2-digit', minute: '2-digit', hour12: false }).format(new Date()); }
    catch (e) { return ''; }
  }
  function _nuevaIdCarga() { return 'CRG_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8); }
  function _escHtml(s) { return String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }
  function _escAttr(s) { return String(s||'').replace(/'/g, '&#39;').replace(/"/g, '&quot;'); }
  // Argumento seguro para un string JS de comilla simple DENTRO de onclick="...". Escapa backslash y comilla
  // para JS, y &/" como entidad para el atributo (el parser las decodifica antes de que JS lea el string).
  // Sin esto, un nombre con apóstrofo (D'Angelo) rompe el handler o abre XSS.
  function _escJs(s) { return String(s==null?'':s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/&/g,'&amp;').replace(/"/g,'&quot;'); }
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
    const st = document.createElement('style'); st.id = 'cargV3CSS';
    st.textContent = `
      /* ── modal responsive: bottom-sheet en mobile, panel centrado en PC/tablet ── */
      @media (min-width:760px){
        .modal-cargadores{left:50%;right:auto;transform:translateX(-50%) translateY(100%);
          width:min(620px,94vw);max-height:90vh;border-radius:22px;border:1px solid #1e293b}
        .modal-cargadores.open{transform:translateX(-50%) translateY(0)}
      }
      .cgv-stats{display:flex;gap:8px;margin:2px 0 12px}
      .cgv-stat{flex:1;background:rgba(245,158,11,.06);border:1px solid rgba(245,158,11,.18);border-radius:12px;padding:9px 12px}
      .cgv-stat .k{font-size:9px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#94794b}
      .cgv-stat .v{font-size:22px;font-weight:900;font-family:ui-monospace,monospace;color:#fcd34d;line-height:1.05}
      .cgv-stat.b .v{color:#6ee7b7}
      /* ── grupo por cargador ── */
      .cgv-group{background:#140f04;border:1px solid #5c3908;border-radius:16px;padding:11px 11px 4px;margin-bottom:13px;animation:cgvIn .35s cubic-bezier(.34,1.4,.5,1)}
      @keyframes cgvIn{from{opacity:0;transform:translateY(-6px) scale(.985)}to{opacity:1;transform:none}}
      .cgv-ghead{display:flex;align-items:center;gap:9px;padding:2px 3px 10px}
      .cgv-ghead .ico{font-size:19px;filter:drop-shadow(0 2px 6px rgba(252,211,77,.4))}
      .cgv-ghead .nm{font-size:15.5px;font-weight:850;color:#fcd34d;flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-.01em}
      .cgv-ghead .cnt{font-size:10.5px;font-weight:800;color:#fbbf24;background:rgba(245,158,11,.14);border:1px solid rgba(245,158,11,.3);border-radius:999px;padding:2px 9px;white-space:nowrap}
      .cgv-ghead .addm{width:30px;height:30px;border-radius:9px;border:1.5px dashed #7a5210;background:rgba(252,211,77,.05);color:#fcd34d;font-size:17px;font-weight:800;cursor:pointer;display:grid;place-items:center;line-height:1}
      .cgv-ghead .addm:active{transform:scale(.92)}
      /* ── carga (evento) ── */
      .cgv-carga{background:#0d1420;border:1px solid #24324a;border-radius:13px;padding:10px 11px;margin-bottom:8px;animation:cgvIn .3s ease}
      .cgv-carga.prov{border-color:#7a5210;border-style:dashed;background:rgba(124,82,16,.08)}
      .cgv-ctop{display:flex;align-items:center;gap:8px;margin-bottom:9px}
      .cgv-hora{font-size:11px;font-weight:800;font-family:ui-monospace,monospace;color:#cbd5e1;background:#1a2740;border:1px solid #2b3c5a;border-radius:7px;padding:2px 8px;display:inline-flex;align-items:center;gap:4px}
      .cgv-edit{font-size:10px;font-weight:700;font-family:ui-monospace,monospace;color:#93a4c2;background:rgba(148,164,194,.1);border-radius:6px;padding:2px 6px}
      .cgv-pct{font-size:22px;font-weight:900;font-family:ui-monospace,monospace;line-height:1;letter-spacing:-.02em;margin-left:auto}
      .cgv-word{font-size:10.5px;font-weight:800;text-transform:uppercase;letter-spacing:.03em}
      .cgv-x{color:#5f7192;font-size:17px;background:transparent;border:none;cursor:pointer;width:24px;height:24px;line-height:1}
      .cgv-x:active{color:#f87171}
      .cgv-track{position:relative;height:30px;border-radius:999px;background:#060b13;border:1px solid #2b3c5a;cursor:ew-resize;touch-action:none;box-shadow:inset 0 2px 6px rgba(0,0,0,.5);user-select:none}
      .cgv-fill{position:absolute;left:0;top:0;bottom:0;border-radius:999px;background:#ef4444;transition:width .12s linear,background .2s}
      .cgv-fill::after{content:'';position:absolute;inset:0;border-radius:999px;background:linear-gradient(180deg,rgba(255,255,255,.26),transparent 45%)}
      .cgv-ticks i{position:absolute;top:6px;bottom:6px;width:1.5px;background:rgba(255,255,255,.14);pointer-events:none}
      .cgv-min{position:absolute;top:-4px;bottom:-4px;left:10%;width:2px;pointer-events:none;background:repeating-linear-gradient(180deg,rgba(239,68,68,.7) 0 4px,transparent 4px 7px)}
      .cgv-handle{position:absolute;top:50%;width:26px;height:26px;border-radius:50%;background:#fff;box-shadow:0 3px 10px rgba(0,0,0,.55),0 0 0 3px rgba(0,0,0,.22);transform:translate(-50%,-50%);pointer-events:none;display:grid;place-items:center}
      .cgv-handle::before{content:'⋮⋮';font-size:11px;color:#334155;letter-spacing:-2px;font-weight:900}
      .cgv-scale{display:flex;justify-content:space-between;font-size:9px;color:#5f7192;font-family:ui-monospace,monospace;margin-top:4px;padding:0 2px}
      .cgv-warn{font-size:10px;color:#fca5a5;margin-top:5px;font-weight:600}
      /* ── fotos grandes en fila con carrusel ── */
      .cgv-fotos{margin-top:11px;border-top:1px dashed #24324a;padding-top:9px}
      .cgv-fotos-row{display:flex;gap:9px;align-items:center;overflow-x:auto;padding-bottom:3px;scrollbar-width:thin;-webkit-overflow-scrolling:touch}
      .cgv-fotos-row::-webkit-scrollbar{height:5px}.cgv-fotos-row::-webkit-scrollbar-thumb{background:#334155;border-radius:9px}
      .cgv-th-img{width:88px;height:88px;flex:0 0 auto;border-radius:13px;overflow:hidden;border:1px solid #3a4d6e;cursor:zoom-in;background:#060b13;position:relative}
      .cgv-th-img img{width:100%;height:100%;object-fit:cover;display:block;transition:transform .25s}
      /* thumb en subida: preview local + spinner (certeza de que está guardando) */
      .cgv-th-loading{cursor:default;border-color:#7a5210}
      .cgv-th-loading img{opacity:.4;filter:blur(1px)}
      .cgv-spin{position:absolute;inset:0;margin:auto;width:26px;height:26px;border:3px solid rgba(252,211,77,.28);border-top-color:#fcd34d;border-radius:50%;animation:cgvSpin .7s linear infinite}
      .cgv-th-loading::after{content:'guardando…';position:absolute;left:0;right:0;bottom:0;font-size:8px;font-weight:800;text-align:center;color:#fcd34d;background:rgba(5,8,14,.72);padding:2px 0;letter-spacing:.02em}
      @keyframes cgvSpin{to{transform:rotate(360deg)}}
      .cgv-th-img:active img{transform:scale(1.06)}
      .cgv-th-img .z{position:absolute;right:4px;bottom:4px;background:rgba(5,8,14,.72);border-radius:6px;font-size:10px;padding:1px 4px;color:#e7eef8}
      .cgv-foto-add{width:88px;height:88px;flex:0 0 auto;border-radius:13px;border:2px dashed #5c3908;background:rgba(252,211,77,.04);color:#fbbf24;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;cursor:pointer}
      .cgv-foto-add:active{border-color:#fcd34d;background:rgba(252,211,77,.09)}
      .cgv-foto-add .cam{font-size:26px;line-height:1}.cgv-foto-add small{font-size:9px;font-weight:800}
      .cgv-foto-add.busy{opacity:.6;pointer-events:none}
      /* ── vacío ── */
      .cgv-empty{text-align:center;padding:26px 14px 22px}
      .cgv-empty .moto{width:74px;height:74px;margin:0 auto 12px;border-radius:20px;border:2px dashed #7a5210;background:rgba(252,211,77,.05);display:grid;place-items:center;font-size:34px;position:relative}
      .cgv-empty .moto .plus{position:absolute;right:-6px;bottom:-6px;width:26px;height:26px;border-radius:50%;background:#fcd34d;color:#3a2606;font-size:18px;font-weight:900;display:grid;place-items:center;box-shadow:0 3px 9px rgba(0,0,0,.4)}
      .cgv-empty p{color:#94a3b8;font-size:13px;margin:0 0 4px}.cgv-empty small{color:#64748b;font-size:11px}
      /* ── agregar (colapsable) ── */
      .cgv-addbtn{width:100%;display:flex;align-items:center;justify-content:center;gap:9px;padding:14px;border-radius:14px;border:2px dashed #5c3908;background:rgba(252,211,77,.05);color:#fcd34d;font-size:14.5px;font-weight:850;cursor:pointer;margin-top:4px;transition:.15s}
      .cgv-addbtn:active{transform:scale(.99);background:rgba(252,211,77,.1)}
      .cgv-addbtn .b{font-size:20px}
      .cgv-add{background:#0e1626;border:1px solid #24324a;border-radius:16px;padding:12px;margin-top:4px;animation:cgvIn .28s ease}
      .cgv-add-hd{display:flex;align-items:center;gap:8px;margin-bottom:10px}
      .cgv-add-hd .t{font-size:12px;font-weight:850;color:#fcd34d;letter-spacing:.02em;text-transform:uppercase;flex:1}
      .cgv-add-hd .x{background:#131d30;border:1px solid #24324a;color:#93a4c2;border-radius:8px;width:28px;height:28px;font-size:14px;cursor:pointer}
      .cgv-qlbl{font-size:10px;font-weight:800;color:#7f8ba3;text-transform:uppercase;letter-spacing:.06em;margin:2px 0 7px}
      .cgv-quick{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:11px}
      .cgv-qcard{background:#131d30;border:1px solid #2b3c5a;border-radius:12px;padding:11px 8px;cursor:pointer;text-align:center;transition:.12s;min-width:0}
      .cgv-qcard:active{transform:scale(.96);border-color:#fcd34d}
      .cgv-qcard .qi{font-size:22px;line-height:1}
      .cgv-qcard .qn{font-size:12px;font-weight:800;color:#e7eef8;margin-top:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .cgv-qcard .qv{font-size:9px;color:#7f8ba3;margin-top:2px}
      .cgv-srch{display:flex;align-items:center;gap:8px;padding:11px 12px;border-radius:11px;background:#131d30;border:1px solid #2b3c5a}
      .cgv-srch input{flex:1;background:transparent;border:none;outline:none;color:#f1f5f9;font-size:14px}
      .cgv-drop{margin-top:7px;max-height:230px;overflow-y:auto}
      .cgv-match{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;padding:11px 13px;border-radius:10px;background:rgba(30,41,59,.45);border:1px solid transparent;color:#f1f5f9;cursor:pointer;margin-bottom:5px;text-align:left}
      .cgv-match:active{background:#1e293b;transform:scale(.99)}
      .cgv-match .mn{font-size:14px;font-weight:700}
      .cgv-match .mv{font-size:11px;color:#7f8ba3}
      .cgv-match .ma{font-size:12px;font-weight:850;color:#4ade80;white-space:nowrap}
      .cgv-nores{color:#64748b;font-size:12.5px;padding:14px 6px;text-align:center;line-height:1.5}
      /* carrusel */
      .cgv-carousel{position:fixed;inset:0;z-index:12000;background:rgba(3,6,12,.93);display:flex;flex-direction:column;justify-content:center;align-items:center;padding:14px}
      .cgv-car-box{width:100%;max-width:520px;background:#05080e;border:1px solid #243247;border-radius:18px;overflow:hidden}
      .cgv-car-top{padding:11px 14px;display:flex;align-items:center;gap:9px;border-bottom:1px solid #1a2740}
      .cgv-car-top .t{font-weight:800;font-size:13px;color:#fcd34d;flex:1}
      .cgv-car-top .x{width:30px;height:30px;border-radius:50%;background:#131d30;border:1px solid #243247;color:#93a4c2;font-size:14px;cursor:pointer}
      .cgv-car-stage{position:relative;aspect-ratio:4/3;background:#0a0f18}
      .cgv-car-stage img{width:100%;height:100%;object-fit:contain}
      .cgv-car-nav{position:absolute;top:50%;transform:translateY(-50%);width:40px;height:40px;border-radius:50%;background:rgba(10,15,24,.8);border:1px solid #243247;color:#fff;font-size:20px;cursor:pointer}
      .cgv-car-nav.l{left:10px}.cgv-car-nav.r{right:10px}
      .cgv-car-count{position:absolute;top:10px;right:12px;font-size:10px;font-weight:800;font-family:ui-monospace,monospace;color:#fff;background:rgba(10,15,24,.7);border-radius:999px;padding:3px 9px}
      .cgv-car-dots{display:flex;gap:6px;justify-content:center;padding:11px}
      .cgv-car-dots i{width:7px;height:7px;border-radius:50%;background:#243247;transition:.2s}.cgv-car-dots i.on{background:#fcd34d;width:20px;border-radius:99px}
      @media(prefers-reduced-motion:reduce){.cgv-group,.cgv-carga,.cgv-fill,.cgv-add{animation:none;transition:none}}
    `;
    document.head.appendChild(st);
  }

  // [100% Supabase] master = proveedores prefijo "CARGADOR" (cache local instantáneo) + RPC directo
  // (aparece al instante lo que el admin crea en MOS; el operador solo registra).
  function _leerCargProv() {
    const all = (window.OfflineManager && OfflineManager.getProveedoresCache) ? (OfflineManager.getProveedoresCache() || []) : [];
    return all
      .filter(p => String(p.nombre || '').trim().toUpperCase().indexOf('CARGADOR') === 0 && String(p.estado || '') === '1')
      .map(p => ({ idCargador: p.idProveedor, nombre: (String(p.nombre || '').replace(/^CARGADOR\s*/i, '').trim() || p.nombre), veces: 0 }));
  }
  let _masterLoading = false;
  async function _cargarMaster() {
    try {
      let prov = _leerCargProv();
      if (prov.length) _master = prov;
      let rpcDefinitivo = false;
      if (!_masterLoading) {
        _masterLoading = true;
        try {
          const res = await API.get('listarCargadoresMaster');
          if (res && res.ok && Array.isArray(res.data)) { prov = res.data; rpcDefinitivo = true; }   // autoritativo (incluye [])
        } catch(_) {}
        _masterLoading = false;
      }
      if (!rpcDefinitivo && !prov.length && window.OfflineManager && OfflineManager.precargar) {
        try { await OfflineManager.precargar('manual'); } catch(_){}
        prov = _leerCargProv();
      }
      _master = prov;
      if (_addOpen) _renderAdd();
    } catch(e) { _masterLoading = false; }
    return _master;
  }

  // Resumen del día → aplana a lista de cargas. FUSIONA preservando cargas provisionales
  // locales (aún no persistidas) para que el poller/refresh del chip no las pise.
  async function _cargarResumen(fecha) {
    fecha = fecha || _hoyStr();
    try {
      const res = await API.get('resumenCargasDia', { fecha });
      if (res && res.ok && res.data) {
        const nuevaFecha = res.data.fecha || fecha;
        const flat = [];
        (res.data.cargadores || []).forEach(cg => {
          (cg.cargas || []).forEach(k => flat.push({
            idCarga: k.idCarga, idCargador: cg.idCargador, nombre: cg.nombre,
            nivel: parseInt(k.nivel) || 0, fotos: Array.isArray(k.fotos) ? k.fotos : [],
            hora: k.hora || '', editado: k.editado || '', edito: !!k.edito, ts: k.ts || '', _prov: false
          }));
        });
        if (_fechaActual === nuevaFecha) {
          const idsServer = new Set(flat.map(c => String(c.idCarga)));
          const provLocales = _dia.filter(c => c && c._prov && !idsServer.has(String(c.idCarga)));
          _dia = provLocales.concat(flat);
        } else {
          _dia = flat;
        }
        _fechaActual = nuevaFecha;
        _generado = res.data.generado || _generado;
        return res.data;
      }
    } catch(e){}
    _fechaActual = _fechaActual || fecha;
    return { fecha: _fechaActual, totalCargas: _dia.length, cargadores: [] };
  }

  function abrir(fecha) {
    _injectCSS();
    fecha = fecha || _hoyStr();
    document.getElementById('overlayCargadores').style.display = 'block';
    document.getElementById('modalCargadores').classList.add('open');
    _addOpen = false; _searchVal = '';
    // [rev] Si abro un día DISTINTO al que ya tengo en memoria, NO mostrar los datos del día anterior
    // (era el "primero muestra hoy y al rato el día que clickié"). Vaciar + mostrar "Cargando…".
    if (_fechaActual === fecha) {
      _dia = _dia.filter(c => c && !c._prov);   // mismo día → cache instantáneo (descarta provisionales)
      _render();
    } else {
      _dia = []; _fechaActual = fecha;
      const res = document.getElementById('cargResumen');
      if (res) res.innerHTML = '<div style="text-align:center;padding:30px 14px;color:#64748b;font-size:13px"><div class="cgv-spin" style="position:static;margin:0 auto 12px"></div>Cargando cargas del día…</div>';
      const add = document.getElementById('cargAdd'); if (add) add.innerHTML = '';
    }
    Promise.all([_cargarMaster(), _cargarResumen(fecha)]).then(() => _render()).catch(()=>{});
  }
  function cerrar() {
    document.getElementById('overlayCargadores').style.display = 'none';
    document.getElementById('modalCargadores').classList.remove('open');
    _dia = _dia.filter(c => c && !c._prov);   // provisionales abandonadas (<10%, sin foto) no cuentan
    if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
    // refrescar el mapa {fecha→nº cargas} → la moto de cada día de guías/preingresos refleja lo recién registrado
    if (typeof App !== 'undefined' && App.precargarConteoCargas) App.precargarConteoCargas();
  }

  // ── agrupar cargas por cargador (orden de primera aparición en _dia) ──
  function _grupos() {
    const orden = [], map = {};
    _dia.forEach(c => {
      const k = String(c.idCargador);
      if (!map[k]) { map[k] = { idCargador: c.idCargador, nombre: c.nombre, cargas: [] }; orden.push(k); }
      if (c.nombre && !map[k].nombre) map[k].nombre = c.nombre;
      map[k].cargas.push(c);
    });
    return orden.map(k => map[k]);
  }

  // ── HTML de una carga (termómetro + fotos) ──
  function _cargaHTML(c) {
    const id = _escAttr(c.idCarga);
    const p = Math.max(0, Math.min(100, parseInt(c.nivel) || 0));
    const fotos = Array.isArray(c.fotos) ? c.fotos : [];
    const col = _thColor(p), colT = _thColorT(p);
    // [540] la miniatura de 88px se pide a /render/image (los PNG crudos pesan ~5MB c/u);
    // el carrusel sigue abriendo el ORIGINAL en máxima calidad.
    const _th = (u) => ((window.Photos && Photos.thumb) ? Photos.thumb(u, 240) : u);
    const thumbs = fotos.map((u, i) => `<div class="cgv-th-img" onclick="Cargadores._carousel('${id}',${i})"><img src="${_escAttr(_th(u))}" loading="lazy" decoding="async" onerror="this.parentNode.style.display='none'"><span class="z">🔍</span></div>`).join('');
    return `<div class="cgv-carga ${c._prov ? 'prov' : ''}" id="cgvCarga_${id}" data-id="${id}">
      <div class="cgv-ctop">
        <span class="cgv-hora">⏱ ${_escHtml(c.hora || _horaAhora())}</span>
        ${(c.edito && c.editado && c.editado !== c.hora) ? `<span class="cgv-edit" title="Última edición">✎ ${_escHtml(c.editado)}</span>` : ''}
        <span class="cgv-word" id="cgvWord_${id}" style="color:${colT}">${_thWord(p)}</span>
        <span class="cgv-pct" id="cgvPct_${id}" style="color:${col}">${p}%</span>
        <button class="cgv-x" onclick="Cargadores.quitar('${id}')" title="Quitar esta carga">✕</button>
      </div>
      <div class="cgv-track" id="cgvTrack_${id}" data-id="${id}">
        <div class="cgv-fill" id="cgvFill_${id}" style="width:${p}%;background:${col}"></div>
        <div class="cgv-ticks"><i style="left:25%"></i><i style="left:50%"></i><i style="left:75%"></i></div>
        <div class="cgv-min"></div>
        <div class="cgv-handle" id="cgvHandle_${id}" style="left:${p}%"></div>
      </div>
      <div class="cgv-scale"><span>0</span><span>25</span><span>50</span><span>75</span><span>100</span></div>
      <div class="cgv-warn" id="cgvWarn_${id}" style="${p >= MIN_GUARDAR ? 'display:none' : ''}">⚠ Jala la barra — mínimo ${MIN_GUARDAR}% para registrar esta carga.</div>
      <div class="cgv-fotos">
        <div class="cgv-fotos-row" id="cgvFotos_${id}">
          ${thumbs}
          <div class="cgv-foto-add" id="cgvAdd_${id}" onclick="Cargadores._pickFoto('${id}')"><span class="cam">＋</span><small>foto</small></div>
        </div>
      </div>
    </div>`;
  }

  function _grupoHTML(g) {
    const gid = _escJs(g.idCargador), gnom = _escJs(g.nombre || g.idCargador);
    return `<div class="cgv-group">
      <div class="cgv-ghead">
        <span class="ico">🛺</span>
        <span class="nm">${_escHtml(g.nombre || g.idCargador)}</span>
        <span class="cnt">${g.cargas.length} carga${g.cargas.length === 1 ? '' : 's'}</span>
        <button class="addm" onclick="Cargadores.agregar('${gid}','${gnom}')" title="Otra carga de ${_escHtml(g.nombre)}">＋</button>
      </div>
      ${g.cargas.map(_cargaHTML).join('')}
    </div>`;
  }

  function _render() {
    const res = document.getElementById('cargResumen');
    if (!res) return;
    const grupos = _grupos();
    const totCargas = _dia.length, totCarg = grupos.length;
    const totalEl = document.getElementById('cargTotal'); if (totalEl) totalEl.textContent = totCargas;
    if (!grupos.length) {
      res.innerHTML = `<div class="cgv-empty">
        <div class="moto">🛺<span class="plus">＋</span></div>
        <p>Sin cargas registradas hoy.</p>
        <small>Toca “+ Carga nueva” para registrar la primera.</small>
      </div>`;
    } else {
      res.innerHTML = `<div class="cgv-stats">
          <div class="cgv-stat"><div class="k">Cargas</div><div class="v">${totCargas}</div></div>
          <div class="cgv-stat b"><div class="k">Cargadores</div><div class="v">${totCarg}</div></div>
        </div>` + grupos.map(_grupoHTML).join('');
    }
    _bindDrag();
    _renderAdd();
  }

  // ── panel agregar (colapsable) ──
  function _renderAdd() {
    const box = document.getElementById('cargAdd'); if (!box) return;
    if (!_addOpen) {
      box.innerHTML = `<button class="cgv-addbtn" onclick="Cargadores.toggleAdd(true)"><span class="b">＋</span> Carga nueva</button>`;
      return;
    }
    const top3 = _master.slice(0, 3);
    const quick = top3.length ? `<div class="cgv-qlbl">Los que más traen · toca para registrar</div>
      <div class="cgv-quick">${top3.map(c => `<div class="cgv-qcard" onclick="Cargadores.agregar('${_escJs(c.idCargador)}','${_escJs(c.nombre)}')">
        <div class="qi">🛺</div><div class="qn">${_escHtml(c.nombre)}</div><div class="qv">${c.veces ? c.veces + ' días' : 'nuevo'}</div></div>`).join('')}</div>` : '';
    box.innerHTML = `<div class="cgv-add">
      <div class="cgv-add-hd"><span class="t">Agregar carga</span><button class="x" onclick="Cargadores.toggleAdd(false)">✕</button></div>
      ${quick}
      <div class="cgv-srch"><span style="color:#64748b">🔎</span>
        <input id="cargBuscarInput" type="text" placeholder="Buscar cargador…" value="${_escAttr(_searchVal)}" oninput="Cargadores._filtrar(this.value)"/>
        <button class="carg-mic" id="cargMic" onclick="App.dictarCargador && App.dictarCargador()" title="Dictado">🎤</button>
      </div>
      <div class="cgv-drop" id="cargDrop"></div>
    </div>`;
    _filtrar(_searchVal);
    const inp = document.getElementById('cargBuscarInput'); if (inp) setTimeout(() => inp.focus(), 60);
  }
  function toggleAdd(open) { _addOpen = !!open; if (!open) _searchVal = ''; _renderAdd(); }

  function _filtrar(query) {
    _searchVal = String(query || '');
    const drop = document.getElementById('cargDrop'); if (!drop) return;
    const ql = _searchVal.trim().toLowerCase();
    let cand = _master;
    if (ql) cand = _master.filter(c => String(c.nombre || '').toLowerCase().includes(ql) || String(c.idCargador || '').toLowerCase().includes(ql));
    if (!cand.length) {
      drop.innerHTML = ql
        ? `<p class="cgv-nores">No existe “${_escHtml(_searchVal.trim())}”.<br>Los cargadores los crea el admin en <b style="color:#fcd34d">MOS → 🚚 Cargadores</b>; apenas lo cree aparecerá aquí.</p>`
        : `<p class="cgv-nores">Escribe el nombre del cargador o toca uno de arriba.</p>`;
      return;
    }
    drop.innerHTML = cand.slice(0, 20).map(c => `<button class="cgv-match" onclick="Cargadores.agregar('${_escJs(c.idCargador)}','${_escJs(c.nombre)}')">
      <span><span class="mn">${_escHtml(c.nombre)}</span> ${c.veces ? '<span class="mv">· ' + c.veces + ' días</span>' : ''}</span>
      <span class="ma">+ carga</span></button>`).join('');
  }

  // ── agregar una CARGA nueva (provisional hasta ≥10% o foto) ──
  async function agregar(idCargador, nombre) {
    const idCarga = _nuevaIdCarga();
    _dia.unshift({ idCarga, idCargador, nombre, nivel: 0, fotos: [], hora: _horaAhora(), ts: '', _prov: true });
    _addOpen = false; _searchVal = '';
    _render();
    _tono(0);
    if (typeof toast === 'function') toast('Jala la barra a ≥' + MIN_GUARDAR + '% para registrar la carga', 'info', 2600);
    const el = document.getElementById('cgvCarga_' + _escAttr(idCarga));
    if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
  }

  async function quitar(idCarga) {
    const c = _dia.find(x => String(x.idCarga) === String(idCarga));
    const nombre = (c || {}).nombre || 'cargador';
    // [rev senior] cancelar cualquier setNivel pendiente de ESTA carga: si el operador jala la barra y
    // dentro de 350ms toca ✕, el upsert diferido podía RESUCITAR la carga borrada. Cancelarlo primero.
    if (_saveTimers[idCarga]) { clearTimeout(_saveTimers[idCarga]); delete _saveTimers[idCarga]; }
    // Confirmación SIEMPRE (evita borrar una carga de casualidad). Modal moderno _whConfirm.
    if (typeof window._whConfirm === 'function') {
      const p = (c && parseInt(c.nivel)) || 0;
      const nf = (c && Array.isArray(c.fotos)) ? c.fotos.length : 0;
      const det = [p ? p + '%' : null, nf ? nf + (nf === 1 ? ' foto' : ' fotos') : null].filter(Boolean).join(' · ');
      const ok = await window._whConfirm(
        `¿Quitar esta carga de ${nombre}${c && c.hora ? ' (⏱ ' + c.hora + ')' : ''}?` + (det ? `\n\nSe elimina: ${det}.` : ''),
        { danger: true, titulo: '🛺 Quitar carga', okText: 'Quitar' });
      if (!ok) return;
    }
    _dia = _dia.filter(x => String(x.idCarga) !== String(idCarga));
    _render();
    if (navigator.vibrate) navigator.vibrate([10, 30, 10]);
    if (c && c._prov) return;   // nunca se persistió
    try { const r = await API.post('cargadorCargaQuitar', { idCarga });
      if (r && r.ok && typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
    } catch(e){ if (typeof toast === 'function') toast('No se pudo quitar — reintenta', 'warn'); }
  }

  // ── termómetro: drag por puntero (id = idCarga) ──
  let _drag = null;
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
    if (fill) { fill.style.width = p + '%'; fill.style.background = col; fill.style.boxShadow = p >= 75 ? '0 0 14px -2px rgba(16,185,129,.55)' : ''; }
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
        _onDragMove(_pctFromEvent(track, ev.clientX));
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
    const c = _dia.find(x => String(x.idCarga) === String(id)); if (c) c.nivel = p;
    if (c && c._prov && p < MIN_GUARDAR) return;   // provisional <10% → no persistir
    clearTimeout(_saveTimers[id]);
    _saveTimers[id] = setTimeout(async () => {
      delete _saveTimers[id];   // el timer ya disparó: no dejar handles rancios acumulados en el mapa
      try { const res = await API.post('cargadorCargaSetNivel', { idCarga: id, idCargador: c && c.idCargador, nivel: p, fecha: _fechaActual || _hoyStr(), nombre: (c && c.nombre) || '', usuario: _usuario(), deviceId: _deviceId() });
        if (res && res.ok) {
          if (c) { c._prov = false; const card = document.getElementById('cgvCarga_' + _escAttr(id)); if (card) card.classList.remove('prov'); }
          if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
        } else if (typeof toast === 'function') toast('No se guardó el nivel — reintenta', 'warn');
      } catch(e){ if (typeof toast === 'function') toast('Sin conexión · nivel no guardado', 'warn'); }
    }, 350);
  }

  // ── fotos (por carga) ──
  // [fix móvil] Antes el input tenía capture='environment' → en el celular abría DIRECTO la cámara,
  // sin dejar subir fotos ya existentes. Ahora, en táctil, mostramos un selector Cámara/Galería;
  // en escritorio (sin cámara relevante) abre directo el diálogo de archivos.
  let _fotoChooserId = null;
  function _pickFoto(idCarga) {
    _fotoChooserId = idCarga;
    const esTactil = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);
    if (!esTactil) { _fotoDesde(false); return; }   // desktop → galería/archivos directo
    let ov = document.getElementById('cgvFotoChooser'); if (ov) ov.remove();
    if (!document.getElementById('cgvChooserKf')) {
      const st = document.createElement('style'); st.id = 'cgvChooserKf';
      st.textContent = '@keyframes cgvChIn{from{opacity:0}to{opacity:1}}@keyframes cgvChUp{from{transform:translateY(14px)}to{transform:translateY(0)}}';
      document.head.appendChild(st);
    }
    ov = document.createElement('div');
    ov.id = 'cgvFotoChooser';
    ov.style.cssText = 'position:fixed;inset:0;z-index:2147483000;background:rgba(2,6,23,.72);backdrop-filter:blur(6px);display:flex;align-items:flex-end;justify-content:center;padding:16px;padding-bottom:calc(16px + env(safe-area-inset-bottom,0px));animation:cgvChIn .18s ease';
    ov.onclick = (e) => { if (e.target === ov) _cerrarFotoChooser(); };
    ov.innerHTML = `
      <div style="width:100%;max-width:420px;background:#0f172a;border:1px solid #1e293b;border-radius:18px;padding:14px;box-shadow:0 -12px 44px rgba(0,0,0,.55);animation:cgvChUp .22s cubic-bezier(.34,1.56,.64,1)">
        <div style="text-align:center;font-size:11px;font-weight:800;color:#94a3b8;letter-spacing:.6px;text-transform:uppercase;margin-bottom:12px">Agregar foto del cargador</div>
        <div style="display:flex;gap:10px">
          <button onclick="Cargadores._fotoDesde(true)" style="flex:1;display:flex;flex-direction:column;align-items:center;gap:7px;background:linear-gradient(135deg,rgba(99,102,241,.16),rgba(99,102,241,.05));border:1px solid rgba(99,102,241,.4);color:#c7d2fe;border-radius:14px;padding:18px 12px;font-size:13px;font-weight:700;cursor:pointer">
            <span style="font-size:30px;line-height:1">📷</span>Cámara</button>
          <button onclick="Cargadores._fotoDesde(false)" style="flex:1;display:flex;flex-direction:column;align-items:center;gap:7px;background:linear-gradient(135deg,rgba(16,185,129,.16),rgba(16,185,129,.05));border:1px solid rgba(16,185,129,.4);color:#a7f3d0;border-radius:14px;padding:18px 12px;font-size:13px;font-weight:700;cursor:pointer">
            <span style="font-size:30px;line-height:1">🖼</span>Galería</button>
        </div>
        <button onclick="Cargadores._cerrarFotoChooser()" style="width:100%;margin-top:10px;background:transparent;border:0;color:#64748b;font-size:13px;font-weight:600;padding:10px;cursor:pointer">Cancelar</button>
      </div>`;
    document.body.appendChild(ov);
  }
  function _cerrarFotoChooser() { const ov = document.getElementById('cgvFotoChooser'); if (ov) ov.remove(); }
  function _fotoDesde(usarCamara) {
    _cerrarFotoChooser();
    const idCarga = _fotoChooserId; if (!idCarga) return;
    let inp = document.getElementById('cgvFileInput');
    if (!inp) { inp = document.createElement('input'); inp.type = 'file'; inp.id = 'cgvFileInput'; inp.style.display = 'none'; document.body.appendChild(inp); }
    inp.accept = 'image/*';
    if (usarCamara) inp.setAttribute('capture', 'environment'); else inp.removeAttribute('capture');
    inp.value = '';
    inp.onchange = (ev) => { const f = ev.target.files && ev.target.files[0]; if (f) _subirFoto(idCarga, f); };
    inp.click();
  }
  // Comprime la foto en el navegador (cámara de celular ~5MB → ~0.3-0.6MB) → sube MUCHO más rápido.
  // Máx lado mayor 1600px, JPEG calidad .82. No toca gif/svg. Devuelve {blob, mime} o null (usar original).
  async function _comprimirImagen(file, maxDim, q) {
    try {
      if (!/^image\//.test(file.type || '') || /gif|svg/.test(file.type || '')) return null;
      const dataUrl = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.onerror = () => r(null); fr.readAsDataURL(file); });
      if (!dataUrl) return null;
      const img = await new Promise(r => { const im = new Image(); im.onload = () => r(im); im.onerror = () => r(null); im.src = dataUrl; });
      if (!img || !img.width) return null;
      let w = img.width, h = img.height;
      const scale = Math.min(1, maxDim / Math.max(w, h));
      w = Math.max(1, Math.round(w * scale)); h = Math.max(1, Math.round(h * scale));
      const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
      cv.getContext('2d').drawImage(img, 0, 0, w, h);
      const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', q || 0.82));
      return (blob && blob.size) ? { blob, mime: 'image/jpeg' } : null;
    } catch(_) { return null; }
  }
  async function _subirFoto(idCarga, file) {
    if (file.size > 20 * 1024 * 1024) { if (typeof toast === 'function') toast('Foto muy pesada (máx 20MB)', 'warn'); return; }
    const c = _dia.find(x => String(x.idCarga) === String(idCarga));
    const add = document.getElementById('cgvAdd_' + _escAttr(idCarga)); if (add) add.classList.add('busy');
    // [v2.13.496] preview OPTIMISTA con spinner: aparece la miniatura al instante mientras sube → certeza.
    const localUrl = (() => { try { return URL.createObjectURL(file); } catch(_) { return ''; } })();
    let loadEl = null;
    const row = document.getElementById('cgvFotos_' + _escAttr(idCarga));
    if (row) {
      loadEl = document.createElement('div'); loadEl.className = 'cgv-th-img cgv-th-loading';
      loadEl.innerHTML = `<img src="${localUrl}" alt=""><span class="cgv-spin"></span>`;
      const ab = row.querySelector('.cgv-foto-add'); if (ab) row.insertBefore(loadEl, ab); else row.appendChild(loadEl);
    }
    // comprimir antes de subir (acelera la subida y aligera el reporte en vivo)
    let blob = file, mime = file.type || 'image/jpeg';
    try { const rc = await _comprimirImagen(file, 1600, 0.82); if (rc && rc.blob) { blob = rc.blob; mime = rc.mime; } } catch(_){}
    const b64 = await new Promise(res => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = () => res(''); r.readAsDataURL(blob); });
    try {
      const res = await API.post('cargadorCargaAddFoto', { idCarga, idCargador: c && c.idCargador, fecha: _fechaActual || _hoyStr(), fotoBase64: b64, mimeType: mime, nombre: (c && c.nombre) || '', usuario: _usuario(), deviceId: _deviceId() });
      if (res && res.ok && res.data) {
        if (c) { c.fotos = res.data.fotos || c.fotos; c._prov = false; }
        if (navigator.vibrate) navigator.vibrate(12);
        _actualizarFotos(idCarga);   // re-render: quita el thumb de carga y pinta la foto real
        if (typeof App !== 'undefined' && App.actualizarChipDia) App.actualizarChipDia();
      } else {
        if (loadEl) loadEl.remove();
        if (typeof toast === 'function') toast('No se subió la foto: ' + ((res && res.error) || '?'), 'warn');
      }
    } catch(e){ if (loadEl) loadEl.remove(); if (typeof toast === 'function') toast('No se subió la foto — reintenta', 'warn'); }
    finally { if (add) add.classList.remove('busy'); if (localUrl) { try { URL.revokeObjectURL(localUrl); } catch(_){} } }
  }
  function _actualizarFotos(idCarga) {
    const card = document.getElementById('cgvCarga_' + _escAttr(idCarga)); if (!card) return;
    const c = _dia.find(x => String(x.idCarga) === String(idCarga)); if (!c) return;
    const wrap = document.createElement('div'); wrap.innerHTML = _cargaHTML(c);
    const nuevo = wrap.querySelector('.cgv-fotos'), viejo = card.querySelector('.cgv-fotos');
    if (nuevo && viejo) viejo.replaceWith(nuevo);
    card.classList.toggle('prov', !!c._prov);
  }

  // ── carrusel (fotos de UNA carga) ──
  function _carousel(idCarga, startIdx) {
    const c = _dia.find(x => String(x.idCarga) === String(idCarga)); if (!c) return;
    const fotos = Array.isArray(c.fotos) ? c.fotos : []; if (!fotos.length) return;
    let idx = Math.max(0, Math.min(fotos.length - 1, parseInt(startIdx) || 0));
    let ov = document.getElementById('cgvCarousel'); if (ov) ov.remove();
    ov = document.createElement('div'); ov.id = 'cgvCarousel'; ov.className = 'cgv-carousel';
    ov.addEventListener('click', e => { if (e.target === ov) ov.remove(); });
    const paint = () => {
      ov.innerHTML = `<div class="cgv-car-box">
        <div class="cgv-car-top"><span style="font-size:16px">🛺</span><span class="t">${_escHtml(c.nombre || c.idCargador)} · ${_escHtml(c.hora || '')}</span><button class="x" onclick="document.getElementById('cgvCarousel').remove()">✕</button></div>
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

  // ── compartir por WhatsApp (con link en vivo) + imprimir ──
  // agrupa las cargas PERSISTIDAS por cargador (orden por actividad reciente); usado por compartir/imagen.
  function _gruposPersistidos() {
    const persist = _dia.filter(c => c && !c._prov);
    const map = {}, orden = [];
    persist.forEach(c => { const k = String(c.idCargador); if (!map[k]) { map[k] = { nombre: c.nombre, cargas: [] }; orden.push(k); } map[k].cargas.push(c); });
    orden.forEach(k => map[k].cargas.sort((a, b) => String(a.hora).localeCompare(String(b.hora))));
    return { persist, orden, map, prom: persist.length ? Math.round(persist.reduce((s, c) => s + (parseInt(c.nivel) || 0), 0) / persist.length) : 0 };
  }
  function _linkReporte(fecha) { return 'https://levo19.github.io/warehouseMos-/cargadores.html?fecha=' + fecha; }
  function _textoWA(fecha) {
    const g = _gruposPersistidos();
    const L = ['*🛺 CARGAS DEL DÍA*', fecha, '─────────────────────'];
    g.orden.forEach(k => {
      const gr = g.map[k];
      L.push(`*${gr.nombre}* — ${gr.cargas.length} carga${gr.cargas.length === 1 ? '' : 's'}`);
      gr.cargas.forEach((c, i) => {
        const p = parseInt(c.nivel) || 0;
        const ico = p >= 75 ? '🟢' : p >= 50 ? '🟡' : p >= 25 ? '🟠' : '🔴';
        const nf = (c.fotos || []).length;
        const ed = (c.edito && c.editado && c.editado !== c.hora) ? ' ✎' + c.editado : '';
        L.push(`   ${i + 1}) ⏱${c.hora || ''}  ${ico} ${p}%${nf ? ' 📷' + nf : ''}${ed}`);
      });
    });
    L.push('', '━━━━━━━━━━━━━━━━━━', `*${g.persist.length}* carga(s) · *${g.orden.length}* cargador(es) · promedio *${g.prom}%*`, '',
      '📲 *Escanea el código QR* de la imagen, o abre este enlace para ver el reporte:',
      _linkReporte(fecha), '',
      'ℹ️ *Importante:*',
      '• Es un reporte *en tiempo real*: se actualiza solo cada 60s.',
      '• La información es *un resumen al momento* y puede *cambiar durante el día* (más cargas o fotos).',
      '• Abre el enlace para ver el estado *actualizado* y las *fotos*.');
    if (_generado) L.push('🕒 Emitido: ' + _generado);
    L.push('', '_InversionMos · warehouseMos_');
    return L.join('\n');
  }
  // Compartir TEXTO (fallback / opción rápida)
  function compartir() {
    const fecha = _fechaActual || _hoyStr();
    if (!_dia.filter(c => !c._prov).length) { if (typeof toast === 'function') toast('Sin cargas para compartir', 'warn'); return; }
    window.open('https://wa.me/?text=' + encodeURIComponent(_textoWA(fecha)), '_blank');
  }
  // Compartir IMAGEN: dibuja el reporte en canvas y lo comparte (Web Share con archivo en móvil;
  // en PC descarga el PNG y abre WhatsApp con el link). El link al reporte en vivo va SIEMPRE en el texto.
  async function compartirImagen() {
    const fecha = _fechaActual || _hoyStr();
    const g = _gruposPersistidos();
    if (!g.persist.length) { if (typeof toast === 'function') toast('Sin cargas para compartir', 'warn'); return; }
    let canvas;
    try { canvas = _dibujarReporte(fecha, g); } catch(e) { compartir(); return; }
    const texto = _textoWA(fecha);
    canvas.toBlob((blob) => {
      if (!blob) { compartir(); return; }
      // Vista previa + botón Enviar (activación FRESCA → share confiable en móvil; en PC copia+descarga)
      if (window.ShareSheet) { window.ShareSheet(blob, 'cargas-' + fecha + '.png', texto); return; }
      compartir();
    }, 'image/png');
  }
  // Dibuja el reporte del día en un canvas (tema oscuro dorado WH). QR grande en la cabecera (derecha),
  // mini-dashboard a la izquierda. El link y las condiciones van en el TEXTO de WhatsApp. Devuelve el canvas.
  function _dibujarReporte(fecha, g) {
    const W = 720, PAD = 30, dpr = Math.min(2, window.devicePixelRatio || 2);   // dpr 2: imagen más liviana
    const link = _linkReporte(fecha);
    let qr = null; try { if (typeof qrcode === 'function') { qr = qrcode(0, 'M'); qr.addData(link); qr.make(); } } catch(_) { qr = null; }
    let qrSize = 0;
    if (qr) { const cnt = qr.getModuleCount(), tot = cnt + 6, cell = Math.max(2, Math.floor(104 / tot)); qrSize = cell * tot; }
    const kpiY = PAD + 62;
    const headerBottom = Math.max(kpiY + 66, qr ? (PAD + qrSize + 26) : 0) + 16;
    // medir alto
    let H = headerBottom;
    g.orden.forEach(k => {
      H += 40;
      g.map[k].cargas.forEach(c => { H += 46; if (c.edito && c.editado && c.editado !== c.hora) H += 16; });
      H += 12;
    });
    H += PAD;   // sin footer: el QR ya está en la cabecera
    const cv = document.createElement('canvas');
    cv.width = W * dpr; cv.height = H * dpr;
    const x = cv.getContext('2d'); x.scale(dpr, dpr); x.textBaseline = 'alphabetic';
    const col = (p) => p >= 75 ? '#10b981' : p >= 50 ? '#fbbf24' : p >= 25 ? '#f97316' : '#ef4444';
    const rr = (px, py, w, h, r) => { x.beginPath(); x.moveTo(px + r, py); x.arcTo(px + w, py, px + w, py + h, r); x.arcTo(px + w, py + h, px, py + h, r); x.arcTo(px, py + h, px, py, r); x.arcTo(px, py, px + w, py, r); x.closePath(); };
    x.fillStyle = '#0b0f17'; x.fillRect(0, 0, W, H);
    // QR en la cabecera (derecha) + chip "Escanéame"
    let leftRight = W - PAD;
    if (qr) {
      const cnt = qr.getModuleCount(), quiet = 3, total = cnt + quiet * 2, cell = Math.max(2, Math.floor(104 / total)), size = cell * total;
      const bx = W - PAD - size, by = PAD;
      x.fillStyle = '#fff'; rr(bx-4, by-4, size+8, size+8, 9); x.fill();
      x.fillStyle = '#000';
      for (let r=0;r<cnt;r++) for (let c=0;c<cnt;c++) if (qr.isDark(r,c)) x.fillRect(bx+(c+quiet)*cell, by+(r+quiet)*cell, cell, cell);
      const chip = '📲 Escanéame'; x.font = '800 11px -apple-system,Segoe UI,sans-serif';
      const cw = x.measureText(chip).width + 16, chy = by + size + 7;
      x.fillStyle = '#fcd34d'; rr(bx + size/2 - cw/2, chy, cw, 18, 9); x.fill();
      x.fillStyle = '#3a2606'; x.textAlign = 'center'; x.fillText(chip, bx + size/2, chy + 13); x.textAlign = 'left';
      leftRight = bx - 16;
    }
    // título + fecha (izquierda)
    x.fillStyle = '#fcd34d'; x.font = '800 30px -apple-system,Segoe UI,Roboto,sans-serif';
    x.fillText('🛺 Cargas del día', PAD, PAD + 30);
    x.fillStyle = '#93a4c2'; x.font = '600 14px ui-monospace,monospace';
    x.fillText(fecha + (_generado ? '   ·   emitido ' + _generado : ''), PAD, PAD + 52);
    // mini-dashboard (KPIs) a la izquierda, sin invadir el QR
    const kpis = [['Cargas', g.persist.length, '#6ee7b7'], ['Cargadores', g.orden.length, '#93c5fd'], ['Promedio', g.prom + '%', '#fcd34d']];
    const gap = 10, kw = (leftRight - PAD - 2 * gap) / 3;
    kpis.forEach((k, i) => {
      const kx = PAD + i * (kw + gap);
      x.fillStyle = '#12100a'; rr(kx, kpiY, kw, 66, 12); x.fill();
      x.strokeStyle = 'rgba(245,158,11,.22)'; x.lineWidth = 1; rr(kx, kpiY, kw, 66, 12); x.stroke();
      x.fillStyle = '#7f8ba3'; x.font = '800 10px -apple-system,sans-serif';
      x.fillText(String(k[0]).toUpperCase(), kx + 12, kpiY + 22);
      x.fillStyle = k[2]; x.font = '900 25px ui-monospace,monospace';
      x.fillText(String(k[1]), kx + 12, kpiY + 52);
    });
    let y = headerBottom;
    // grupos
    g.orden.forEach(k => {
      const gr = g.map[k];
      x.fillStyle = '#fcd34d'; x.font = '800 19px -apple-system,sans-serif';
      x.fillText('🛺 ' + gr.nombre, PAD, y + 20);
      x.fillStyle = '#fbbf24'; x.font = '800 12px ui-monospace,monospace'; x.textAlign = 'right';
      x.fillText(gr.cargas.length + ' carga' + (gr.cargas.length === 1 ? '' : 's'), W - PAD, y + 19);
      x.textAlign = 'left'; y += 34;
      gr.cargas.forEach((c, i) => {
        const p = Math.max(0, Math.min(100, parseInt(c.nivel) || 0));
        x.fillStyle = '#cbd5e1'; x.font = '800 13px ui-monospace,monospace';
        x.fillText((i + 1) + ')  ⏱ ' + (c.hora || ''), PAD, y + 15);
        const nf = (c.fotos || []).length;
        let extra = (nf ? '  📷' + nf : '') + ((c.edito && c.editado && c.editado !== c.hora) ? '   ✎ ' + c.editado : '');
        if (extra) { x.fillStyle = '#7f8ba3'; x.font = '700 11px ui-monospace,monospace'; x.fillText(extra, PAD + 132, y + 15); }
        // barra
        const bx = PAD, bw = W - 2 * PAD - 62, by = y + 24, bh = 16;
        x.fillStyle = '#060b13'; rr(bx, by, bw, bh, 8); x.fill();
        x.fillStyle = col(p); rr(bx, by, Math.max(bh, bw * p / 100), bh, 8); x.fill();
        x.fillStyle = col(p); x.font = '900 15px ui-monospace,monospace'; x.textAlign = 'right';
        x.fillText(p + '%', W - PAD, by + 14); x.textAlign = 'left';
        y += 46; if (c.edito && c.editado && c.editado !== c.hora) y += 16;
      });
      y += 12;
    });
    return cv;
  }
  async function imprimir() {
    const fecha = _fechaActual || _hoyStr();
    if (!_dia.filter(c => !c._prov).length) { if (typeof toast === 'function') toast('Sin cargas para imprimir', 'warn'); return; }
    if (!window.PrintHub || !PrintHub.imprimir) { if (typeof toast === 'function') toast('Impresión no disponible aquí', 'warn'); return; }
    try {
      const res = await PrintHub.imprimir('imprimirCargadoresDia', { fecha }, 'Cargadores del día ' + fecha);
      if (res && res.ok) { if (typeof toast === 'function') toast('Ticket impreso ✓', 'ok'); }
      else if (typeof toast === 'function') toast('No se pudo imprimir: ' + ((res && res.error) || '?'), 'warn');
    } catch(e) { if (typeof toast === 'function') toast('No se pudo imprimir', 'warn'); }
  }

  // [rev senior] contar solo cargas PERSISTIDAS (no provisionales 0% sin registrar) → el chip no infla.
  function getCountDia(fecha) { fecha = fecha || _hoyStr(); return (_fechaActual === fecha) ? _dia.filter(c => c && !c._prov).length : 0; }
  async function refreshCountDia(fecha) {
    fecha = fecha || _hoyStr();
    const m = document.getElementById('modalCargadores');
    if (m && m.classList.contains('open') && _fechaActual === fecha) return _dia.filter(c => c && !c._prov).length;   // fuente viva (solo persistidas)
    const r = await _cargarResumen(fecha);
    return (r.totalCargas != null) ? r.totalCargas : _dia.filter(c => c && !c._prov).length;
  }

  window.Cargadores = {
    abrir, cerrar, agregar, quitar, compartir, compartirImagen, imprimir, toggleAdd,
    getCountDia, refreshCountDia,
    _filtrar, _hoyStr, _pickFoto, _fotoDesde, _cerrarFotoChooser, _carousel
  };
})();
