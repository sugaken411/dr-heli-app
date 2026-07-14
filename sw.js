const CACHE_NAME = 'aw109-cache-v13.0'; // バージョンを一斉引き上げ古いキャッシュを完全にパージ・破壊
const urlsToCache = [
  './',
  './portal.html',
  './index.html',
  './search.html',
  './management.html', // 🌟 追加：PWAオフライン環境でのカレンダー画面白化バグを完全に抑止
  './checklist.html',
  './report.html',
  './debriefing.html',
  './debrief_search.html',
  './library.html',
  './summary.html',
  './admin.html',
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
            return caches.delete(cacheName); // 古いバージョンv12.x系のキャッシュをすべて論理削除
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // POSTリクエストやGAS側APIへの通信はキャッシュせず常にスルー（動的インデックスマッピング通信の最新性を保護）
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return;
  }
  
  event.respondWith(
    fetch(event.request)
      .then(response => {
        // ネットワークから正常取得できたらキャッシュを動的更新
        if (response && response.status === 200 && response.type === 'basic') {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseToCache));
        }
        return response;
      })
      .catch(() => caches.match(event.request)) // 電波瞬断・オフライン時はローカルキャッシュを返却
  );
});
