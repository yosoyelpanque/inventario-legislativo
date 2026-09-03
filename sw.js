// Refresco de publicación: conserva una única versión de caché para la corrección Oracle.
const CACHE = 'inventario-legislativo-v13';
const CORE = [
  './', './index.html', './styles.css?v=13', './app.js?v=13', './db.js', './parser.js', './reports.js',
  './manifest.json', './vendor/jszip.min.js', './assets/camara-logo.png',
  './assets/icon-192.png', './assets/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE).map((key) => caches.delete(key)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request).then((response) => {
    if (response.ok && new URL(event.request.url).origin === location.origin) {
      const copy = response.clone();
      caches.open(CACHE).then((cache) => cache.put(event.request, copy));
    }
    return response;
  }).catch(() => caches.match('./index.html'))));
});
