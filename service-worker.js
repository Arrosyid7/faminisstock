// ============================================================
// SERVICE WORKER — Sistem Stok Faminis Barokah
// Strategy: Cache First untuk asset lokal, Network Only untuk API
// ============================================================

const CACHE_NAME = 'faminis-stok-v2';
const OFFLINE_URL = './offline.html';

// Semua asset statis yang di-pre-cache saat install
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './Index.html',
  './manifest.json',
  './offline.html',
  './icons/icon-48.png',
  './icons/icon-72.png',
  './icons/icon-96.png',
  './icons/icon-128.png',
  './icons/icon-144.png',
  './icons/icon-168.png',
  './icons/icon-192.png',
  './icons/icon-384.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/screenshot-mobile.png'
];

// =====================
// EVENT: INSTALL
// =====================
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[SW] Pre-caching assets...');
      return Promise.allSettled(
        ASSETS_TO_CACHE.map((url) =>
          cache.add(url).catch((err) => {
            console.warn('[SW] Gagal cache:', url, err);
          })
        )
      );
    }).then(() => {
      console.log('[SW] Install selesai, langsung aktif.');
      return self.skipWaiting();
    })
  );
});

// =====================
// EVENT: ACTIVATE
// =====================
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter((name) => name !== CACHE_NAME)
          .map((name) => {
            console.log('[SW] Hapus cache lama:', name);
            return caches.delete(name);
          })
      );
    }).then(() => {
      console.log('[SW] Aktif, mengambil alih semua klien.');
      return self.clients.claim();
    })
  );
});

// =====================
// EVENT: FETCH
// =====================
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  if (event.request.method !== 'GET') {
    return;
  }

  if (!url.protocol.startsWith('http')) return;

  const isCdn = (
    url.hostname === 'cdn.tailwindcss.com' ||
    url.hostname === 'unpkg.com' ||
    url.hostname === 'cdn.jsdelivr.net'
  );

  if (isCdn) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response.ok) {
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
          }
          return response;
        })
        .catch(() => {
          return caches.match(event.request).then((cached) => {
            if (cached) return cached;
            return new Response('', { status: 503, statusText: 'CDN tidak tersedia (offline)' });
          });
        })
    );
    return;
  }

  if (url.hostname === 'lh3.googleusercontent.com') {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response.ok) {
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
          }
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        fetch(event.request).then((response) => {
          if (response && response.ok) {
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, response));
          }
        }).catch(() => {});
        return cachedResponse;
      }

      return fetch(event.request).then((response) => {
        if (response.ok) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
        }
        return response;
      }).catch(() => {
        if (event.request.mode === 'navigate') {
          return caches.match(OFFLINE_URL).then((offlinePage) => {
            return offlinePage || new Response(
              '<h1>Offline</h1><p>Buka ulang saat ada koneksi.</p>',
              { headers: { 'Content-Type': 'text/html' } }
            );
          });
        }
        return new Response('', { status: 503, statusText: 'Offline' });
      });
    })
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
