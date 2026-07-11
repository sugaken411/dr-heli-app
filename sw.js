const CACHE_NAME = 'aw109-cache-v12.2'; // バージョンを上げて古いキャッシュを破壊
const urlsToCache = [
  './',
  './index.html',
  './search.html',
  './report.html',
  './debriefing.html',
  './debrief_search.html',
  './library.html',
  './summary.html',
  './admin.html',
  './portal.html',
  './manifest.json',
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/chart.js',
  'https://cdnjs.cloudflare.com/ajax/libs/suncalc/1.9.0/suncalc.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(urlsToCache))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName); // 古いキャッシュをすべて削除
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // POSTリクエストやAPIへの通信はキャッシュせずスルー（常に最新を取得）
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return;
  }
  
  event.respondWith(
    fetch(event.request)
      .then(response => {
        // ネットワークから取得できたらキャッシュを更新
        if (response && response.status === 200 && response.type === 'basic') {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseToCache));
        }
        return response;
      })
      .catch(() => caches.match(event.request)) // オフライン時はキャッシュを返す
  );
});
