// ============================================================
// warehouseMos — js/photos.js
// Sistema fotos: lightbox + miniruleta + picker fuente +
// compresión client-side antes de subir.
// ============================================================

(function() {
  'use strict';

  const MAX_DIM = 1600;
  const JPEG_Q  = 0.85;

  // ── Lightbox fullscreen ──
  function lightbox(urls, idx) {
    if (typeof urls === 'string') urls = [urls];
    if (!urls || !urls.length) return;
    idx = idx || 0;
    let cur = idx;

    const overlay = document.createElement('div');
    overlay.className = 'photo-lightbox';
    overlay.innerHTML = `
      <div class="photo-lightbox-bg" onclick="this.parentNode.remove()"></div>
      <button class="photo-lb-close" onclick="this.parentNode.remove()">✕</button>
      <img class="photo-lb-img" src="${urls[cur]}" />
      ${urls.length > 1 ? `
        <button class="photo-lb-prev">‹</button>
        <button class="photo-lb-next">›</button>
        <div class="photo-lb-dots">${urls.map((_,i)=>`<span class="${i===cur?'on':''}"></span>`).join('')}</div>
      ` : ''}
    `;
    document.body.appendChild(overlay);

    if (urls.length > 1) {
      const prev = overlay.querySelector('.photo-lb-prev');
      const next = overlay.querySelector('.photo-lb-next');
      const img  = overlay.querySelector('.photo-lb-img');
      const dots = overlay.querySelectorAll('.photo-lb-dots span');
      const goto = i => {
        cur = (i + urls.length) % urls.length;
        img.src = urls[cur];
        dots.forEach((d, j) => d.classList.toggle('on', j === cur));
      };
      prev.onclick = () => goto(cur - 1);
      next.onclick = () => goto(cur + 1);
      let startX = 0;
      img.addEventListener('touchstart', e => { startX = e.touches[0].clientX; });
      img.addEventListener('touchend',   e => {
        const dx = e.changedTouches[0].clientX - startX;
        if (Math.abs(dx) > 50) goto(cur + (dx < 0 ? 1 : -1));
      });
    }
  }

  // ── Miniruleta (carousel autoplay) ──
  // Renderiza HTML del carousel. El caller debe meterlo en el DOM e invocar
  // initCarousels() para activarlos.
  function carouselHTML(urls, opts) {
    opts = opts || {};
    if (!urls || !urls.length) return '';
    const dataAttr = 'data-pcarousel="' + urls.length + '"';
    const dataUrls = 'data-urls="' + urls.map(u => u.replace(/"/g,'&quot;')).join('|') + '"';
    const size = opts.size || 'md';
    return `
      <div class="pcarousel pcar-${size}" ${dataAttr} ${dataUrls}>
        <img class="pcar-img" src="${urls[0]}" loading="lazy"/>
        ${urls.length > 1 ? `<div class="pcar-dots">${urls.map((_,i)=>`<span class="${i===0?'on':''}"></span>`).join('')}</div>` : ''}
      </div>
    `;
  }

  function initCarousels(root) {
    root = root || document;
    root.querySelectorAll('.pcarousel[data-pcarousel]').forEach(el => {
      if (el._inited) return;
      el._inited = true;
      const urls = (el.getAttribute('data-urls') || '').split('|').filter(Boolean);
      if (urls.length < 2) return;
      let idx = 0;
      const img  = el.querySelector('.pcar-img');
      const dots = el.querySelectorAll('.pcar-dots span');
      const tick = () => {
        idx = (idx + 1) % urls.length;
        img.src = urls[idx];
        dots.forEach((d, j) => d.classList.toggle('on', j === idx));
      };
      let timer = setInterval(tick, 4000);
      el.addEventListener('mouseenter', () => clearInterval(timer));
      el.addEventListener('mouseleave', () => { timer = setInterval(tick, 4000); });
      el.addEventListener('click', e => {
        e.stopPropagation();
        lightbox(urls, idx);
      });
      let startX = 0;
      el.addEventListener('touchstart', e => { startX = e.touches[0].clientX; clearInterval(timer); });
      el.addEventListener('touchend',   e => {
        const dx = e.changedTouches[0].clientX - startX;
        if (Math.abs(dx) > 30) {
          idx = (idx + (dx < 0 ? 1 : -1) + urls.length) % urls.length;
          img.src = urls[idx];
          dots.forEach((d, j) => d.classList.toggle('on', j === idx));
        }
        timer = setInterval(tick, 4000);
      });
    });
  }

  // ── Compresión client-side antes de subir ──
  function comprimir(file) {
    return new Promise((res, rej) => {
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          let w = img.width, h = img.height;
          if (w > MAX_DIM || h > MAX_DIM) {
            if (w > h) { h = Math.round(h * MAX_DIM / w); w = MAX_DIM; }
            else        { w = Math.round(w * MAX_DIM / h); h = MAX_DIM; }
          }
          const canvas = document.createElement('canvas');
          canvas.width = w; canvas.height = h;
          canvas.getContext('2d').drawImage(img, 0, 0, w, h);
          const dataUrl = canvas.toDataURL('image/jpeg', JPEG_Q);
          const b64 = dataUrl.split(',')[1];
          res({ base64: b64, mime: 'image/jpeg', size: b64.length });
        };
        img.onerror = rej;
        img.src = reader.result;
      };
      reader.onerror = rej;
      reader.readAsDataURL(file);
    });
  }

  // ── Subir a Drive vía GAS endpoint genérico ──
  async function subir(entidad, idEntidad, file) {
    const c = await comprimir(file);
    const res = await API.post('subirFotoEntidad', {
      entidad, idEntidad,
      base64: c.base64, mimeType: c.mime
    });
    if (!res || !res.ok) throw new Error((res && res.error) || 'upload falló');
    return res.data;
  }

  // ── Picker de fuente al crear guía desde preingreso ──
  // opts: { fotosPreingreso: [url,...], onChoose: ({source, url|file}) => void }
  function abrirPickerFuente(opts) {
    const overlay = document.getElementById('overlayPhotoSource');
    const modal   = document.getElementById('modalPhotoSource');
    if (!overlay || !modal) return;
    const lista = document.getElementById('photoSourceFromPre');
    if (lista) {
      const urls = opts.fotosPreingreso || [];
      lista.innerHTML = urls.length
        ? urls.map((u, i) => `<button class="ps-thumb" data-url="${u}"><img src="${u}"/></button>`).join('')
        : '<p style="color:#64748b;font-size:12px;text-align:center;width:100%">Sin fotos en el preingreso.</p>';
      lista.querySelectorAll('.ps-thumb').forEach(b => {
        b.onclick = () => {
          lista.querySelectorAll('.ps-thumb').forEach(x => x.classList.remove('sel'));
          b.classList.add('sel');
          modal._picked = { source: 'pre', url: b.dataset.url };
        };
      });
    }
    const camBtn = document.getElementById('photoSourceCam');
    const galBtn = document.getElementById('photoSourceGal');
    const noBtn  = document.getElementById('photoSourceNone');
    const okBtn  = document.getElementById('photoSourceOk');
    const fileInput = document.getElementById('photoSourceFile');
    if (camBtn) camBtn.onclick = () => { fileInput.accept = 'image/*'; fileInput.capture = 'environment'; fileInput.click(); };
    if (galBtn) galBtn.onclick = () => { fileInput.accept = 'image/*'; fileInput.removeAttribute('capture'); fileInput.click(); };
    if (fileInput) fileInput.onchange = e => {
      const f = e.target.files && e.target.files[0];
      if (f) modal._picked = { source: 'file', file: f };
    };
    if (noBtn) noBtn.onclick = () => { modal._picked = { source: 'none' }; };
    if (okBtn) okBtn.onclick = () => {
      const picked = modal._picked || { source: 'none' };
      overlay.style.display = 'none';
      modal.classList.remove('open');
      if (typeof opts.onChoose === 'function') opts.onChoose(picked);
    };
    overlay.style.display = 'block';
    modal.classList.add('open');
    modal._picked = null;
  }

  function cerrarPickerFuente() {
    const overlay = document.getElementById('overlayPhotoSource');
    const modal   = document.getElementById('modalPhotoSource');
    if (overlay) overlay.style.display = 'none';
    if (modal) modal.classList.remove('open');
  }

  window.Photos = {
    lightbox, carouselHTML, initCarousels,
    comprimir, subir,
    abrirPickerFuente, cerrarPickerFuente
  };
})();
