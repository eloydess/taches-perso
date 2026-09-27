/* Tâches perso — service worker minimal.
   Réseau d'abord pour les fichiers de l'application (toujours la dernière version),
   copie locale seulement en secours. Les appels au serveur Apps Script ne sont jamais mis en cache. */
var CACHE = 'taches-perso-v1';
var FICHIERS = ['./', 'index.html', 'manifest.webmanifest', 'apple-touch-icon.png', 'icone-192.png', 'icone-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FICHIERS); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (cles) {
    return Promise.all(cles.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req).then(function (rep) {
      var copie = rep.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copie); });
      return rep;
    }).catch(function () {
      return caches.match(req).then(function (r) { return r || caches.match('index.html'); });
    })
  );
});
