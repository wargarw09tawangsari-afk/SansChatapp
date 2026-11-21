const CACHE_NAME = 'sanschat-v36';
const USER_META_CACHE = 'sanschat-user-meta-v1';
const S3_API_BASE = "https://script.google.com/macros/s/AKfycbyHBqAVQ044nhszVhbKIeP2qi6Wwafu5KgfU648L6RNQSBwYPd_DXWm-TOqUwO7KVg/exec";
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './login.html',
  './utama.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

let swActiveUsername = "";
let swRingtone = "standard";
let lastNotifiedCallId = "";
let lastNotifiedChatTime = 0;

async function persistSwUserMeta(username, ringtone) {
  if (!username) return;
  swActiveUsername = String(username).trim().toUpperCase();
  swRingtone = ringtone || "standard";
  try {
    const cache = await caches.open(USER_META_CACHE);
    await cache.put('/active-user-meta', new Response(JSON.stringify({
      username: swActiveUsername,
      ringtone: swRingtone
    })));
  } catch (e) {}
}

async function restoreSwUserMeta() {
  if (swActiveUsername) return swActiveUsername;
  try {
    const cache = await caches.open(USER_META_CACHE);
    const resp = await cache.match('/active-user-meta');
    if (resp) {
      const data = await resp.json();
      swActiveUsername = (data.username || "").trim().toUpperCase();
      swRingtone = data.ringtone || "standard";
    }
  } catch (e) {}
  return swActiveUsername;
}

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME && k !== USER_META_CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SET_ACTIVE_USER') {
    event.waitUntil(persistSwUserMeta(event.data.username, event.data.ringtone));
  }
  if (event.data && event.data.type === 'CLEAR_ACTIVE_USER') {
    swActiveUsername = "";
    event.waitUntil(caches.delete(USER_META_CACHE));
  }
  if (event.data && event.data.type === 'TRIGGER_BG_CHECK') {
    event.waitUntil((async () => {
      const u = event.data.username || await restoreSwUserMeta();
      if (u) {
        await persistSwUserMeta(u, event.data.ringtone || swRingtone);
        await checkBackgroundAlertsForUser(u);
      }
    })());
  }
});

self.addEventListener('periodicsync', (event) => {
  if (event.tag === 'sanschat-bg-watcher') {
    event.waitUntil((async () => {
      const u = await restoreSwUserMeta();
      if (u) await checkBackgroundAlertsForUser(u);
    })());
  }
});

self.addEventListener('sync', (event) => {
  if (event.tag === 'sanschat-bg-sync') {
    event.waitUntil((async () => {
      const u = await restoreSwUserMeta();
      if (u) await checkBackgroundAlertsForUser(u);
    })());
  }
});

async function checkBackgroundAlertsForUser(username) {
  if (!username) return;
  try {
    const callResp = await fetch(S3_API_BASE, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'checkIncomingCall', username: username })
    });
    const callJson = await callResp.json();
    const incoming = callJson && callJson.data;
    if (incoming && incoming.callerName) {
      const callKey = (incoming.roomUrl || incoming.callerName) + "_" + (incoming.type || "voice");
      if (lastNotifiedCallId !== callKey) {
        lastNotifiedCallId = callKey;
        const isVideo = incoming.type === 'video';
        await self.registration.showNotification(
          isVideo ? `📹 VIDEO CALL MASUK: ${incoming.callerName}` : `📞 TELPON MASUK: ${incoming.callerName}`,
          {
            body: `${incoming.callerName} memanggil ${username}. Ketuk TERIMA untuk menjawab!`,
            icon: './icon-192.png',
            badge: './icon-192.png',
            tag: 'sanschat-incoming-call',
            renotify: true,
            requireInteraction: true,
            vibrate: [600, 200, 600, 200, 600, 200, 600],
            data: { actionType: 'incoming_call', username: username, callerName: incoming.callerName, callType: incoming.type || 'voice', roomUrl: incoming.roomUrl || '' },
            actions: [{ action: 'accept_call', title: '✅ TERIMA' }, { action: 'reject_call', title: '❌ TOLAK' }]
          }
        );
      }
    }
    const chatResp = await fetch(S3_API_BASE, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'hasUnreadPrivateMessages', username: username })
    });
    const chatJson = await chatResp.json();
    if (chatJson && chatJson.data === true) {
      const now = Date.now();
      if (now - lastNotifiedChatTime > 12000) {
        lastNotifiedChatTime = now;
        await self.registration.showNotification(
          `💬 CHAT MASUK UNTUK ${username}`,
          {
            body: 'Ketuk TERIMA CHAT untuk membuka dan membalas pesan.',
            icon: './icon-192.png',
            badge: './icon-192.png',
            tag: 'sanschat-incoming-chat',
            renotify: true,
            requireInteraction: true,
            vibrate: [500, 200, 500, 200, 500],
            data: { actionType: 'incoming_chat', username: username },
            actions: [{ action: 'accept_chat', title: '💬 TERIMA CHAT' }, { action: 'dismiss_chat', title: '🔕 ABAIKAN' }]
          }
        );
      }
    }
  } catch (e) {}
}

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const data = event.notification.data || {};
  const username = data.username || swActiveUsername || "";
  event.waitUntil((async () => {
    if (event.action === 'reject_call' && username) {
      try {
        await fetch(S3_API_BASE, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'updateCallStatus', username: username, status: 'REJECTED' })
        });
      } catch (e) {}
      return;
    }
    if (event.action === 'dismiss_chat') return;
    const allClients = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const client of allClients) {
      if ('focus' in client) {
        await client.focus();
        client.postMessage({ type: 'OPEN_FROM_NOTIFICATION', actionType: data.actionType || 'incoming_chat', username: username });
        return;
      }
    }
    const targetUrl = username ? `./utama.html?user=${encodeURIComponent(username)}` : './index.html';
    if (clients.openWindow) await clients.openWindow(targetUrl);
  })());
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const reqUrl = new URL(event.request.url);
  if (reqUrl.origin !== self.location.origin) return;
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((resp) => {
        if (resp && resp.status === 200 && resp.type === 'basic') {
          const clone = resp.clone();
          caches.open(CACHE_NAME).then((c) => c.put(event.request, clone));
        }
        return resp;
      }).catch(() => cached);
    })
  );
});
