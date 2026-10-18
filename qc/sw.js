/* Network first, cache as fallback, so the page still opens when the kitchen wifi drops. */
var V = 'eqc-v1';
self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(caches.open(V).then(function (c) {
    return c.addAll(['./', './index.html', './config.js', './manifest.webmanifest']);
  }).catch(function () {}));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== V; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(function (r) {
    var copy = r.clone();
    caches.open(V).then(function (c) { c.put(e.request, copy); });
    return r;
  }).catch(function () {
    return caches.match(e.request).then(function (m) { return m || caches.match('./index.html'); });
  }));
});
