const CACHE_NAME = 'classroom-helper-offline-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.png',
  './small-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // Add each file on its own so one missing file cannot break the whole install
      Promise.all(ASSETS.map((asset) => cache.add(asset).catch(() => {})))
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Network-first: when online you always get the newest files (no version bump needed
// while testing). When offline, the cached copy is used.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request).then((response) => {
      if (response && response.status === 200 && response.type === 'basic') {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
      }
      return response;
    }).catch(() =>
      caches.match(event.request).then((cached) => cached || caches.match('./index.html'))
    )
  );
});
