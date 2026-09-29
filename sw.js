const CACHE_NAME = 'h2i-nilai-v2-cache-2026-09';
const FILES_TO_CACHE = [
  './index.html',
  './manifest.json',
  './icon-192.svg',
  './icon-512.svg'
];

// 1. Install Service Worker & Cache Files
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(FILES_TO_CACHE);
    })
  );
  self.skipWaiting(); // Force activation
});

// 2. Activate Service Worker & Clean Old Cache
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keyList) => {
      return Promise.all(keyList.map((key) => {
        if (key !== CACHE_NAME) {
          console.log('[ServiceWorker] Removing old cache', key);
          return caches.delete(key);
        }
      }));
    })
  );
  self.clients.claim();
});

// 3. Fetch Event (Offline First Strategy)
self.addEventListener('fetch', (event) => {
  // Hanya cache request GET (Abaikan POST ke API Google Script)
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((response) => {
      const path = new URL(event.request.url).pathname;
      const alwaysFresh = /\/(index\.html|manifest\.json|sw\.js|icon-192\.svg|icon-512\.svg)$/.test(path);
      if (alwaysFresh) {
        return fetch(event.request).then((networkResponse) => {
          if (networkResponse && networkResponse.ok) {
            const copy = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return networkResponse;
        }).catch(() => response || Promise.reject(new Error('Offline')));
      }
      if (response) return response;
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.ok && event.request.url.startsWith(self.location.origin)) {
          const copy = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return networkResponse;
      });
    })
  );
});
