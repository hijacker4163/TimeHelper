// Service Worker (sw.js) - 用於快取離線資源與在行動裝置/背景顯示推播通知
const CACHE_NAME = 'time-calc-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/manifest.json'
];

// 安裝 Service Worker 並快取基礎外殼
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE).catch((err) => {
        console.warn('部分靜態資源快取略過:', err);
      });
    })
  );
});

// 啟動階段清理舊快取並立刻接管頁面
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 攔截網路請求以支援離線存取基礎頁面
self.addEventListener('fetch', (event) => {
  // 只攔截 GET 請求，API 請求直接放行
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com') || event.request.url.includes('/api/')) {
    return;
  }
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});

// 監聽來自網頁端的訊息以觸發推播
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SHOW_NOTIFICATION') {
    const { title, options } = event.data;
    event.waitUntil(
      self.registration.showNotification(title, options)
    );
  }
});

// 點擊通知時自動喚醒或切換至該應用分頁
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ('focus' in client) {
          return client.focus();
        }
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow('/');
      }
    })
  );
});
