// ============================================================
// warehouseMos — Service Worker
// Cambia VERSION en cada deploy para invalidar caché
// ============================================================

// ── Firebase Cloud Messaging (background push) ─────────────
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey:            'AIzaSyA_gfynRxAmlbGgHWoioaj5aeaxnnywP88',
  projectId:         'proyectomos-push',
  messagingSenderId: '328735199478',
  appId:             '1:328735199478:web:947f338ae9716a7c049cd7'
});

const _fcmMsg = firebase.messaging();
_fcmMsg.onBackgroundMessage(payload => {
  // Comandos data-only (audio_start, audio_stop, gps_locate) → reenviar al cliente sin notificación
  if (payload.data && payload.data.action) {
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clients => {
      clients.forEach(c => c.postMessage({ type: 'mos_command', data: payload.data }));
    });
    return;
  }
  const title = payload.notification?.title || 'warehouseMos';
  const body  = payload.notification?.body  || '';
  self.registration.showNotification(title, {
    body,
    icon:    'https://levo19.github.io/MOS/icon-192.png',
    badge:   'https://levo19.github.io/MOS/icon-192.png',
    tag:     'wh-push',
    vibrate: [200, 100, 200]
  });
});

const VERSION = '2.13.6';
const CACHE   = 'warehouse-v' + VERSION;

// Solo assets locales — CDN se cachea en el fetch handler al primer uso
const LOCAL_ASSETS = [
  './',
  './index.html',
  './reporte.html',
  './manifest.json',
  './version.json',
  './js/app.js',
  './js/api.js',
  './js/offline.js',
  './js/scanner.js',
  './js/sounds.js',
];

// ── Instalar: cachear secuencial con reporte de progreso ──
// Cada asset cacheado dispara postMessage al cliente para mostrar
// barra de progreso real en el banner de update. Un fallo individual
// no aborta el install (Promise.allSettled mental).
self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    const total = LOCAL_ASSETS.length;
    let done = 0;
    async function _broadcast(payload) {
      const cs = await self.clients.matchAll({ includeUncontrolled: true, type: 'window' });
      cs.forEach(c => { try { c.postMessage(payload); } catch(_){} });
    }
    await _broadcast({ type: 'sw-install-progress', done: 0, total, version: VERSION });
    for (const url of LOCAL_ASSETS) {
      try {
        await cache.add(new Request(url, { cache: 'no-store' }));
      } catch (err) {
        console.warn('[SW] No se pudo cachear:', url, err);
      }
      done++;
      await _broadcast({ type: 'sw-install-progress', done, total, version: VERSION });
    }
    await _broadcast({ type: 'sw-install-done', total, version: VERSION });
  })());
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

  // version.json: siempre desde red. Si falla red, devolver caché o respuesta
  // de error explícita (NUNCA undefined — eso rompe respondWith).
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(
      fetch(e.request).catch(async () => {
        const cached = await caches.match(e.request);
        return cached || new Response('{"version":"offline"}', {
          status: 503, headers: { 'Content-Type': 'application/json' }
        });
      })
    );
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        // Si la red devuelve algo inválido, fallback a respuesta de error explícita
        if (!res) return Response.error();
        if (res.status !== 200) return res;
        // Solo cachear respuestas same-origin o CORS (no opaque)
        if (res.type !== 'basic' && res.type !== 'cors') return res;
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return res;
      }).catch(async () => {
        if (e.request.destination === 'document') {
          const indexCached = await caches.match('./index.html');
          return indexCached || Response.error();
        }
        return Response.error();
      });
    })
  );
});

// ── Mensaje SKIP_WAITING desde la app ───────────────────────
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});
