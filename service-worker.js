const CACHE_NAME = 'sanschat-v29';
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

self.addEventListener('install', (event) => {
 self.skipWaiting();
 event.waitUntil(
 caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
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

self.addEventListener('message', (event) => {
 if (event.data && event.data.type === 'SET_ACTIVE_USER') {
 swActiveUsername = String(event.data.username || "").trim().toUpperCase();
 swRingtone = event.data.ringtone || "standard";
 }
 if (event.data && event.data.type === 'TRIGGER_BG_CHECK') {
 swActiveUsername = String(event.data.username || swActiveUsername || "").trim().toUpperCase();
 event.waitUntil(checkBackgroundAlertsForUser(swActiveUsername));
 }
});

self.addEventListener('periodicsync', (event) => {
 if (event.tag === 'sanschat-bg-watcher' && swActiveUsername) {
 event.waitUntil(checkBackgroundAlertsForUser(swActiveUsername));
 }
});

self.addEventListener('sync', (event) => {
 if (event.tag === 'sanschat-bg-sync' && swActiveUsername) {
 event.waitUntil(checkBackgroundAlertsForUser(swActiveUsername));
 }
});

async function checkBackgroundAlertsForUser(username) {
 if (!username) return;
 try {
 // 1. Cek Panggilan Suara / Video Call Masuk
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
 isVideo ? `📹 Video Call Masuk: ${incoming.callerName}` : `📞 Telpon Masuk: ${incoming.callerName}`,
 {
 body: `${incoming.callerName} memanggil ${username}. Ketuk Terima untuk menjawab!`,
 icon: './icon-192.png',
 badge: './icon-192.png',
 tag: 'sanschat-incoming-call',
 renotify: true,
 requireInteraction: true,
 vibrate: [500, 200, 500, 200, 500, 200, 500],
 data: {
 actionType: 'incoming_call',
 username: username,
 callerName: incoming.callerName,
 callType: incoming.type || 'voice',
 roomUrl: incoming.roomUrl || ''
 },
 actions: [
 { action: 'accept_call', title: '✅ TERIMA' },
 { action: 'reject_call', title: '❌ TOLAK' }
 ]
 }
 );
 }
 }

 // 2. Cek Chat Privat Masuk
 const chatResp = await fetch(S3_API_BASE, {
 method: 'POST',
 headers: { 'Content-Type': 'text/plain;charset=utf-8' },
 body: JSON.stringify({ action: 'hasUnreadPrivateMessages', username: username })
 });
 const chatJson = await chatResp.json();
 if (chatJson && chatJson.data === true) {
 await self.registration.showNotification(
 `💬 Pesan Privat Baru untuk ${username}`,
 {
 body: 'Ketuk untuk membuka dan membalas pesan di Sc-SansChat.',
 icon: './icon-192.png',
 badge: './icon-192.png',
 tag: 'sanschat-unread-private',
 renotify: true,
 vibrate: [300, 150, 300],
 data: {
 actionType: 'open_private_chat',
 username: username
 },
 actions: [
 { action: 'open_chat', title: '💬 BUKA CHAT' }
 ]
 }
 );
 }
 } catch (e) {}
}

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

self.addEventListener('notificationclick', (event) => {
 const notifData = event.notification.data || {};
 const clickedAction = event.action || 'default';
 event.notification.close();

 const targetUser = encodeURIComponent(notifData.username || swActiveUsername || '');
 let targetUrl = `./utama.html?user=${targetUser}`;

 if (notifData.actionType === 'incoming_call') {
 if (clickedAction === 'reject_call') {
 event.waitUntil(
 fetch(S3_API_BASE, {
 method: 'POST',
 headers: { 'Content-Type': 'text/plain;charset=utf-8' },
 body: JSON.stringify({ action: 'updateCallStatus', username: notifData.username || swActiveUsername, status: 'ENDED' })
 }).catch(() => {})
 );
 return;
 }
 const cType = encodeURIComponent(notifData.callType || 'voice');
 const cName = encodeURIComponent(notifData.callerName || 'Warga');
 const cUrl = encodeURIComponent(notifData.roomUrl || '');
 const autoAns = clickedAction === 'accept_call' ? '1' : '0';
 targetUrl = `./utama.html?user=${targetUser}&incomingCall=1&autoAnswer=${autoAns}&callType=${cType}&callerName=${cName}&roomUrl=${cUrl}`;
 } else if (notifData.actionType === 'open_private_chat') {
 targetUrl = `./utama.html?user=${targetUser}&openPrivate=1`;
 }

 event.waitUntil(
 clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
 for (const client of clientList) {
 if ('navigate' in client) {
 client.navigate(targetUrl);
 return client.focus();
 }
 }
 if (clients.openWindow) return clients.openWindow(targetUrl);
 })
 );
});
