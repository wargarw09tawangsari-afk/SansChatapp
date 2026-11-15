// ============================================================================
FILE: service-worker.js (Gabungan untuk PWA SansChat - GitHub Pages)
============================================================================
const CACHE_NAME = 'sanschat-v6';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './utama.html',
  './login.html',
  './manifest.json',
  'https://cdn-icons-png.flaticon.com/512/3135/3135715.png'
];

// 1. EVENT INSTALL: Menyimpan aset penting ke dalam cache browser
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('Membuka cache dan menyimpan aset statis...');
        return cache.addAll(ASSETS_TO_CACHE);
      })
      .catch((err) => console.error('Gagal melakukan caching aset:', err))
  );
});

// 2. EVENT ACTIVATE: Membersihkan cache versi lama agar memori tetap optimal
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('Menghapus cache versi lama:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 3. EVENT FETCH: Mengelola permintaan jaringan & fallback offline
self.addEventListener('fetch', (event) => {
  // Abaikan request non-GET atau request eksternal ke Google Apps Script agar data dinamis tetap real-time
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        // Jika jaringan aktif, simpan salinan terbaru ke cache
        if (networkResponse && networkResponse.status === 200) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }
        return networkResponse;
      })
      .catch(() => {
        // Jika sedang offline, ambil dari cache browser
        return caches.match(event.request).then((cachedResponse) => {
          if (cachedResponse) {
            return cachedResponse;
          }
          // Fallback darurat jika halaman utama diminta dalam kondisi offline total
          if (event.request.mode === 'navigate') {
            return caches.match('./index.html');
          }
        });
      })
  );
});

// 4. EVENT NOTIFICATION CLICK: Menangani interaksi notifikasi PWA
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ('focus' in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow('./index.html');
    })
  );
});
