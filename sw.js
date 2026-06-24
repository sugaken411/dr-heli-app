const CACHE_NAME = 'aw109-ems-cache-v9.2';
const urlsToCache = [
  './',
  './index.html',
  './search.html',
  './report.html',
  './library.html',
  './admin.html',
  './portal.html'
];

// インストール時に指定したファイルをキャッシュ
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
});

// 通信発生時の処理（ネットワーク優先、圏外ならキャッシュを返す）
self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});

// 古いキャッシュの削除
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          if (cache !== CACHE_NAME) return caches.delete(cache);
        })
      );
    })
  );
});
