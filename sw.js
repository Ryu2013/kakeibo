const CACHE_NAME = 'kakeibo-v2';
const STATIC_ASSETS = ['./manifest.json'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return; // GASなど外部への通信は素通し

  // HTML(アプリ本体)は開発中なので常に最新を優先し、オフライン時だけキャッシュにフォールバック
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          caches.open(CACHE_NAME).then((c) => c.put(e.request, res.clone()));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // その他の静的アセットはキャッシュ優先
  e.respondWith(caches.match(e.request).then((cached) => cached || fetch(e.request)));
});
