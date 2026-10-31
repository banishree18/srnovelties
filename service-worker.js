const CACHE_NAME = 'srnovelties-v1';
const OFFLINE_PAGE = '/offline.html';

// Save the offline page when the app is installed
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.add(OFFLINE_PAGE))
  );
  self.skipWaiting();
});

// Remove old saved versions
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Network first: always try the internet, use saved copy only when offline
self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only handle normal page/file loads from our own website.
  // Orders and live products (Google Apps Script) are never touched.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    fetch(req)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        return response;
      })
      .catch(() =>
        caches.match(req).then((saved) => saved || caches.match(OFFLINE_PAGE))
      )
  );
});