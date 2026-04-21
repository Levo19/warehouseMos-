// ============================================================
// warehouseMos — Service Worker
// Cambia VERSION en cada deploy para invalidar caché
// ============================================================
const VERSION = '1.0.85';
const CACHE   = 'warehouse-v' + VERSION;

// Solo assets locales — CDN se cachea en el fetch handler al primer uso
const LOCAL_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './version.json',
  './js/app.js',
  './js/api.js',
  './js/offline.js',
  './js/scanner.js',
];

// ── Instalar: cachear cada asset individualmente (un fallo no mata el install)
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache =>
      Promise.allSettled(
        LOCAL_ASSETS.map(url =>
          cache.add(new Request(url, { cache: 'no-store' }))
            .catch(err => console.warn('[SW] No se pudo cachear:', url, err))
        )
      )
    ).then(() => self.skipWaiting())
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

  // No interceptar GAS (tanto script.google.com como script.googleusercontent.com)
  if (url.hostname.includes('google.com') || url.hostname.includes('googleusercontent.com')) return;

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
        // Solo cachear respuestas same-origin o CORS (no opaque)
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
