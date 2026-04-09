// warehouseMos — Service Worker v1.0
const CACHE = 'warehouse-v1';
const STATIC = [
  './',
  './index.html',
  './js/app.js',
  './js/api.js',
  './js/scanner.js',
  'https://cdn.tailwindcss.com',
  'https://unpkg.com/@zxing/library@latest/umd/index.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(STATIC.map(u => {
      try { return new Request(u, { mode: 'no-cors' }); }
      catch { return u; }
    }))).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // No interceptar peticiones al GAS (siempre necesitan red)
  if (e.request.url.includes('script.google.com')) return;

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(response => {
        if (response && response.status === 200 && response.type !== 'opaque') {
          const clone = response.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return response;
      }).catch(() => {
        if (e.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
