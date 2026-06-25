// バージョンを更新（古いキャッシュを強制破棄させます）
const CACHE_NAME = 'aw109-ems-cache-v9.7';
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

// 🌟 【最重要修正】通信のコントロール
self.addEventListener('fetch', event => {
  // GETリクエスト以外（GASへのPOST送信など）、または GAS(script.google.com) への通信の場合は
  // Service Workerを完全にスルーして、ブラウザ本来の直接通信を行わせる
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return; 
  }

  // HTMLやCSSなどの表示用ファイルのみ、キャッシュを確認する
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
