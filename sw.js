// Service Worker — Automatizaciones Madero
// ⬆️ SUBE ESTE NÚMERO en cada publicación para forzar limpieza de caché vieja.
//    (El HTML ya es "network-first", así que app.html/index.html se actualizan solos;
//     bumpear la versión solo limpia cachés de recursos antiguos.)
const CACHE_NAME = 'madero-app-v1.7.4';

// Rutas RELATIVAS al scope del SW — funcionan igual en la raíz o en subcarpeta
// (/madero-app/ y /madero-app-dev/), a diferencia de las rutas absolutas '/app.html'.
const URLS_TO_CACHE = [
  './',
  'index.html',
  'app.html',
  'offline.html',
  'manifest.json',
  'icon-192.png',
  'icon-512.png',
];

// Install: precachear núcleo (resistente: si algún archivo falta, no aborta la instalación)
self.addEventListener('install', function(event) {
  self.skipWaiting(); // activar de inmediato, no esperar
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return Promise.all(URLS_TO_CACHE.map(function(u) {
        return cache.add(u).catch(function(err) {
          console.warn('[SW] No se pudo precachear', u, '—', err && err.message);
        });
      }));
    })
  );
});

// Permitir que la página fuerce la activación de una versión nueva en espera
self.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

// Activate: borrar TODAS las cachés viejas
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
      return self.clients.claim(); // tomar control de las pestañas abiertas
    })
  );
});

// Fetch: network-first para HTML (nunca servir HTML viejo), cache-first para recursos
self.addEventListener('fetch', function(event) {
  var req = event.request;

  // IMPORTANTE: solo interceptar GET del MISMO ORIGEN.
  // Las llamadas a los flujos de Power Automate (POST), a Microsoft Graph y a CDNs
  // deben pasar directo a la red — Cache.put NO soporta POST y las rompería.
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // HTML siempre fresco desde la red (con respaldo a caché si no hay conexión)
  if (event.request.destination === 'document' ||
      url.pathname === '/' ||
      url.pathname.endsWith('.html')) {
    event.respondWith(
      fetch(event.request)
        .then(function(response) {
          if (response && response.ok) {
            var clone = response.clone();
            caches.open(CACHE_NAME).then(function(cache) { cache.put(event.request, clone); });
          }
          return response;
        })
        .catch(function() {
          return caches.match(event.request).then(function(hit) {
            return hit || caches.match('offline.html') || caches.match('./');
          });
        })
    );
    return;
  }

  // Cache-first para lo demás (iconos, fuentes, etc.)
  event.respondWith(
    caches.match(event.request).then(function(cached) {
      return cached || fetch(event.request).then(function(response) {
        if (response && response.ok) {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) { cache.put(event.request, clone); });
        }
        return response;
      });
    })
  );
});
