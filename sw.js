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
      .then(cache => Promise.all(
        // 🌟 addAllは1件でも失敗すると全体が失敗する（外部CDNが一時的に落ちているだけで
        // オフライン機能が全滅しかねない）ため、1件ずつ個別に試し、失敗しても続行する
        urlsToCache.map(url => cache.add(url).catch(err => console.warn('[SW] キャッシュ失敗（続行）:', url, err)))
      ))
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

  // 🌟 事案検索コンソール→症例登録（編集）への画面遷移で、この fetch が電波不良により
  // レスポンスもエラーも返らないまま「ハング」し、location.href による画面遷移そのものが
  // 止まって「通信タイムアウト」表示のまま固まる不具合が報告された（2026-09-08）。
  // 明示的なタイムアウトで打ち切り、キャッシュへフォールバックすることで画面遷移自体は必ず進むようにする。
  const NAV_TIMEOUT_MS = 6000;
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), NAV_TIMEOUT_MS);

  event.respondWith(
    fetch(event.request, { signal: controller.signal })
      .then(response => {
        clearTimeout(timeoutId);
        // ネットワークから正常取得できたらキャッシュを動的更新
        if (response && response.status === 200 && response.type === 'basic') {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseToCache));
        }
        return response;
      })
      .catch(() => {
        clearTimeout(timeoutId);
        return caches.match(event.request); // 電波瞬断・オフライン・タイムアウト時はローカルキャッシュを返却
      })
  );
});
