const CACHE_NAME = 'aw109-ems-cache-v9.4'; // ←ここをv9.4に上げました
const urlsToCache = [
  './',
  './index.html',
  './search.html',
  './report.html',
  './library.html',
  './admin.html',
  './portal.html'
];

self.addEventListener('install', event => {
  self.skipWaiting(); // 新しいバージョンを即座にアクティブにする
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          // 古いバージョンのキャッシュを完全に削除
          if (cache !== CACHE_NAME) return caches.delete(cache);
        })
      );
    }).then(() => self.clients.claim()) // 開いている画面のコントロールを即座に奪う
  );
});
