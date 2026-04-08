// Service Worker - 勤怠記録 PWA
const CACHE = 'kintai-v3';
const ASSETS = [
  './index.html',
  './manifest.json',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // POSTリクエストはキャッシュしない（GAS送信・Firebase通信）
  if (e.request.method !== 'GET') return;

  // 外部API通信はキャッシュしない
  if (url.includes('script.google.com')) return;
  if (url.includes('firestore.googleapis.com')) return;
  if (url.includes('firebase')) return;
  if (url.includes('googleapis.com')) return;
  if (url.includes('gstatic.com')) return;
  if (url.includes('fonts.googleapis.com')) return;

  // GETリクエストのみキャッシュ
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (!res || res.status !== 200 || res.type === 'opaque') return res;
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return res;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
