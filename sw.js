// FabTracker Service Worker v1
// Cache básico para funcionar offline (pantalla de carga)

var CACHE_NAME = 'fabtracker-v1';
var urlsToCache = [
  './app.html',
  './icon-192.png',
  './icon-512.png',
  './manifest.json'
];

// Instalar SW y cachear archivos base
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      console.log('SW: Cacheando archivos base');
      return cache.addAll(urlsToCache);
    })
  );
  self.skipWaiting();
});

// Activar y limpiar caches viejos
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.filter(function(name) {
          return name !== CACHE_NAME;
        }).map(function(name) {
          console.log('SW: Eliminando cache viejo:', name);
          return caches.delete(name);
        })
      );
    })
  );
  self.clients.claim();
});

// Estrategia: Network First (siempre intenta red, fallback a cache)
// Así siempre tienes la versión más reciente cuando hay internet
self.addEventListener('fetch', function(event) {
  // Solo cachear requests del mismo origen
  if (!event.request.url.startsWith(self.location.origin)) return;
  // No cachear requests a Microsoft/Graph API
  if (event.request.url.includes('microsoftonline') ||
      event.request.url.includes('graph.microsoft') ||
      event.request.url.includes('sharepoint')) return;

  event.respondWith(
    fetch(event.request)
      .then(function(response) {
        // Guardar copia en cache si es válida
        if (response && response.status === 200) {
          var responseClone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, responseClone);
          });
        }
        return response;
      })
      .catch(function() {
        // Sin red — usar cache
        return caches.match(event.request).then(function(cached) {
          if (cached) return cached;
          // Fallback a app.html para navegación
          if (event.request.mode === 'navigate') {
            return caches.match('./app.html');
          }
        });
      })
  );
});
