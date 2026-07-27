// Service Worker Sistem Stok Faminis Barokah
// Tujuan utamanya CUMA supaya app ini bisa di-install ke home screen & kebuka cepat pas dibuka
// ulang -- BUKAN mode transaksi offline penuh, karena data stok/transaksi tetap wajib realtime
// dari Google Apps Script (lihat pengecualian origin di bawah).

const CACHE_NAME = 'faminis-stok-shell-v1';
const APP_SHELL = [
  './',
  './Index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch((e) => console.log('Gagal cache app shell (diabaikan):', e))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // PENTING: JANGAN ikut campur permintaan ke Apps Script Web App (URL_API) atau ke CDN pihak
  // ketiga (Tailwind, html5-qrcode, JsBarcode, Chart.js) -- biarkan lewat langsung ke jaringan
  // seperti biasa, supaya data stok/transaksi & script selalu versi terbaru, tidak ke-cache basi.
  if (url.origin !== self.location.origin) {
    return;
  }

  // Untuk file app shell milik sendiri: coba jaringan dulu (biar selalu dapat versi terbaru
  // kalau online), baru jatuh ke cache kalau offline/gagal -- jadi app tetap kebuka biar sinyal
  // lagi jelek, walau data di dalamnya baru bisa dimuat begitu online lagi.
  event.respondWith(
    fetch(req)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(req))
  );
});
