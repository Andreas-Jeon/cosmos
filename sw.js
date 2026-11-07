/* Cosmos Mobile service worker: app files only (index, scripts, icons). OneDrive data is never cached here. */
var V = "cosmos-shell-v1";
var SHELL = ["./", "index.html", "loader.js", "msal-browser.min.js", "manifest.webmanifest", "icon-192.png", "icon-512.png"];
self.addEventListener("install", function (e) { e.waitUntil(caches.open(V).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); })); });
self.addEventListener("activate", function (e) { e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== V; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); })); });
self.addEventListener("fetch", function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(function (r) {
    if (r.ok) { var cp = r.clone(); caches.open(V).then(function (c) { c.put(e.request, cp); }); }
    return r;
  }).catch(function () { return caches.match(e.request, {ignoreSearch: true}).then(function (r) { return r || caches.match("index.html"); }); }));
});
