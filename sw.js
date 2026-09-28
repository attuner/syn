const CACHE_NAME = 'kutumb-radio-v3';
const STATIC_ASSETS = [
  './index.html',
  './admin.html',
  './manifest.json',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

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

self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // Real-world radio streams, Google Apps Script API calls, and Google Drive audio proxy 
  // must always stream live directly from the network without stale caching
  if (
    url.includes('script.google.com') ||
    url.includes('drive.google.com') ||
    url.includes('googleusercontent.com') ||
    event.request.method !== 'GET'
  ) {
    return;
  }

  // Network-first strategy for dynamic HTML pages to ensure schedule updates appear instantly
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Cache fallback for static styling and fonts
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      return cachedResponse || fetch(event.request);
    })
  );
});