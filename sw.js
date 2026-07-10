// Service Worker — Automatizaciones Madero
// v1.7.2 — bumped to force cache invalidation and reload new index.html

const CACHE_NAME = 'madero-app-v1.7.2';
const URLS_TO_CACHE = [
  '/',
  '/index.html',
  '/app.html',
  '/offline.html',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png',
];

// Install: cache core files
self.addEventListener('install', function(event) {
  self.skipWaiting(); // activate immediately, don't wait
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(URLS_TO_CACHE);
    })
  );
});

// Activate: delete ALL old caches
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames
          .filter(function(name) { return name !== CACHE_NAME; })
          .map(function(name) {
            console.log('[SW] Eliminando cache viejo:', name);
            return caches.delete(name);
          })
      );
    }).then(function() {
      return self.clients.claim(); // take control of all open tabs
    })
  );
});

// Fetch: network first for HTML, cache first for assets
self.addEventListener('fetch', function(event) {
  var url = new URL(event.request.url);

  // Always fetch HTML fresh from network (never serve stale index.html)
  if (event.request.destination === 'document' ||
      url.pathname === '/' ||
      url.pathname.endsWith('.html')) {
    event.respondWith(
      fetch(event.request)
        .then(function(response) {
          // Update cache with fresh version
          if (response.ok) {
            var clone = response.clone();
            caches.open(CACHE_NAME).then(function(cache) {
              cache.put(event.request, clone);
            });
          }
          return response;
        })
        .catch(function() {
          // Offline fallback
          return caches.match(event.request) ||
                 caches.match('/offline.html');
        })
    );
    return;
  }

  // Cache first for everything else (icons, fonts, etc.)
  event.respondWith(
    caches.match(event.request).then(function(cached) {
      return cached || fetch(event.request).then(function(response) {
        if (response.ok) {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, clone);
          });
        }
        return response;
      });
    })
  );
});
