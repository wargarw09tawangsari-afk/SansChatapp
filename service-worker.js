const CACHE_NAME = 'sanschat-v38';
const USER_META_CACHE = 'sanschat-user-meta-v1';
const S3_API_BASE = "https://script.google.com/macros/s/AKfycbyHBqAVQ044nhszVhbKIeP2qi6Wwafu5KgfU648L6RNQSBwYPd_DXWm-TOqUwO7KVg/exec";
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './login.html',
  './utama.html',
  './single-frame.css',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE)));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME && k !== USER_META_CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SET_ACTIVE_USER') {
    event.waitUntil(
      caches.open(USER_META_CACHE).then((c) => c.put('/active-user-meta', new Response(JSON.stringify({
        username: String(event.data.username || '').toUpperCase()
      }))))
    );
  }
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) return;
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const networkFetch = fetch(event.request).then((res) => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return res;
      }).catch(() => cachedResponse);
      return cachedResponse || networkFetch;
    })
  );
});
