// ============================================================
// warehouseMos — Service Worker
// Cambia VERSION en cada deploy para invalidar caché
// ============================================================
const VERSION = '1.0.54';
const CACHE   = 'warehouse-v' + VERSION;
const ASSETS  = [
  './',
  './index.html',
  './manifest.json',
  './version.json',
  './js/app.js',
  './js/api.js',
  './js/offline.js',
  './js/scanner.js',
  './assets/icon.png',
  'https://cdn.tailwindcss.com',
  'https://unpkg.com/@zxing/library@0.20.0/umd/index.min.js'
];

// ── Instalar: cachear assets con no-store (ignora CDN) ──────
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(
        ASSETS.map(url => new Request(url, { cache: 'no-store' }))
      ))
      .then(() => self.skipWaiting())
  );
});

// ── Activar: borrar cachés viejos y reclamar clientes ───────
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// ── Fetch: caché primero, red como fallback ──────────────────
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // No interceptar GAS
  if (url.hostname.includes('script.google.com')) return;

  // version.json: siempre desde red para detectar cambios
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (!res || res.status !== 200) return res;
        if (res.type !== 'basic' && res.type !== 'cors') return res;
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return res;
      }).catch(() => {
        if (e.request.destination === 'document') return caches.match('./index.html');
        return Response.error();
      });
    })
  );
});

// ── Mensaje SKIP_WAITING desde la app ───────────────────────
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});
