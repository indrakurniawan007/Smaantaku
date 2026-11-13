// Service worker Absen Sholat: simpan halaman + pustaka scanner agar aplikasi tetap bisa dibuka saat offline.
// Letakkan SEFOLDER dengan index.html (harus lewat HTTPS). Naikkan VER bila isi file ini diubah.
const VER = 'absen-v1';
const CDN = [
  'https://cdn.tailwindcss.com/3.4.17',
  'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

self.addEventListener('install', function (e) {
  e.waitUntil((async function () {
    const c = await caches.open(VER);
    await Promise.all(CDN.map(function (u) {
      return fetch(new Request(u, { mode: 'no-cors' })).then(function (r) { return c.put(u, r); }).catch(function () {});
    }));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', function (e) {
  e.waitUntil((async function () {
    const ks = await caches.keys();
    await Promise.all(ks.filter(function (k) { return k !== VER; }).map(function (k) { return caches.delete(k); }));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', function (e) {
  const req = e.request;
  if (req.method !== 'GET') return;               // POST ke Apps Script tidak disentuh
  const url = new URL(req.url);

  // pustaka CDN: ambil dari simpanan dulu
  if (CDN.indexOf(req.url) >= 0) {
    e.respondWith(caches.match(req.url).then(function (hit) {
      return hit || fetch(req).then(function (r) {
        const salinan = r.clone();
        caches.open(VER).then(function (c) { c.put(req.url, salinan); });
        return r;
      });
    }));
    return;
  }

  // halaman aplikasi: utamakan versi terbaru dari internet, cadangan dari simpanan saat offline
  if (url.origin === location.origin && (req.mode === 'navigate' || url.pathname.endsWith('.html'))) {
    e.respondWith(fetch(req).then(function (r) {
      const salinan = r.clone();
      caches.open(VER).then(function (c) { c.put(req, salinan); });
      return r;
    }).catch(function () {
      return caches.match(req, { ignoreSearch: true }).then(function (hit) {
        return hit || caches.match(self.registration.scope);
      });
    }));
  }
});
