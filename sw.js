const CACHE_NAME = 'viandas-shell-v1';
const APP_FILES = [
  './index.html',
  './menu.html',
  './productos.html',
  './resumen.html',
  './cambios.html',
  './historial.html',
  './reportes.html',
  './usuario.html',
  './styles.css',
  './app-common.js',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_FILES))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys
        .filter(key => key.startsWith('viandas-shell-') && key !== CACHE_NAME)
        .map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== 'GET' || url.origin !== self.location.origin) return;

  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then(response => {
          if (response.ok) {
            const responseToCache = response.clone();
            event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.put(request, responseToCache)));
          }
          return response;
        })
        .catch(async () => (await caches.match(request)) || caches.match('./index.html'))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then(cached => {
      if (cached) return cached;
      return fetch(request).then(response => {
        if (response.ok) {
          const responseToCache = response.clone();
          event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.put(request, responseToCache)));
        }
        return response;
      });
    })
  );
});
