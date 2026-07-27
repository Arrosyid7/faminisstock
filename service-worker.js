// ============================================================
// SERVICE WORKER — Sistem Stok Faminis Barokah
// Strategy: Cache First untuk asset lokal, Network Only untuk API
// ============================================================

const CACHE_NAME = 'faminis-stok-v1';
const OFFLINE_URL = './offline.html';

// Semua asset statis yang di-pre-cache saat install
const ASSETS_TO_CACHE = [
  './',
  './Index.html',
  './manifest.json',
  './offline.html',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png'
];

// =====================
// EVENT: INSTALL
// =====================
// Pre-cache semua asset statis. Kalau salah satu gagal di-cache,
// service worker tetap ter-install (tidak gagal total).
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[SW] Pre-caching assets...');
      // Gunakan addAll dengan error handling per-item supaya 1 file yang
      // gagal tidak membatalkan seluruh instalasi.
      return Promise.allSettled(
        ASSETS_TO_CACHE.map((url) =>
          cache.add(url).catch((err) => {
            console.warn('[SW] Gagal cache:', url, err);
          })
        )
      );
    }).then(() => {
      console.log('[SW] Install selesai, langsung aktif.');
      // skipWaiting: service worker baru langsung aktif tanpa tunggu tab lama ditutup.
      return self.skipWaiting();
    })
  );
});

// =====================
// EVENT: ACTIVATE
// =====================
// Hapus cache versi lama saat service worker baru aktif.
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
      // Langsung kontrol semua tab yang sudah terbuka tanpa perlu reload.
      return self.clients.claim();
    })
  );
});

// =====================
// EVENT: FETCH
// =====================
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // ─── Abaikan permintaan non-GET (POST, dll) ──────────────
  // Semua panggilan API ke Apps Script pakai POST — biarkan browser
  // menanganinya langsung. Kalau POST gagal (offline), browser sendiri
  // yang akan throw network error, dan app akan menampilkan pesan error.
  if (event.request.method !== 'GET') {
    return; // tidak perlu event.respondWith — browser tangani sendiri
  }

  // ─── Abaikan URL chrome-extension:// atau non-http ───────
  if (!url.protocol.startsWith('http')) return;

  // ─── CDN pihak ketiga (Tailwind, Chart.js, html5-qrcode, JsBarcode) ─
  // Strategy: Network First, fallback ke cache kalau offline.
  const isCdn = (
    url.hostname === 'cdn.tailwindcss.com' ||
    url.hostname === 'unpkg.com' ||
    url.hostname === 'cdn.jsdelivr.net'
  );

  if (isCdn) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          // Simpan ke cache untuk offline
          if (response.ok) {
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
          }
          return response;
        })
        .catch(() => {
          // Offline — coba dari cache
          return caches.match(event.request).then((cached) => {
            if (cached) return cached;
            // CDN tidak tersedia dan tidak di-cache — biarkan error alami
            return new Response('', { status: 503, statusText: 'CDN tidak tersedia (offline)' });
          });
        })
    );
    return;
  }

  // ─── Gambar Google Drive (logo) ────────────────────────────
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

  // ─── Asset lokal (HTML, manifest, icon) ────────────────────
  // Strategy: Cache First → Network fallback → Offline page
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Tetap update cache di background (stale-while-revalidate)
        fetch(event.request).then((response) => {
          if (response && response.ok) {
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, response));
          }
        }).catch(() => {});
        return cachedResponse;
      }

      // Tidak di cache — ambil dari network
      return fetch(event.request).then((response) => {
        if (response.ok) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
        }
        return response;
      }).catch(() => {
        // Benar-benar offline dan tidak ada cache — tampilkan offline page
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

// =====================
// EVENT: MESSAGE
// =====================
// Terima perintah dari halaman, misal untuk skip waiting paksa.
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
