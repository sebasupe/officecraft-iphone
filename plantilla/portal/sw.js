// Service Worker ligero del portal OfficeCraft (/).
// Nunca toca las rutas /word/, /deck/ ni /grid/ porque cada app tiene su propio SW independiente.
const VERSION = "__VERSION__";
const CACHE = "officecraft-portal-" + VERSION;
const PRECACHE = __PRECACHE__;

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(PRECACHE.map((u) => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith("officecraft-portal-") && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (/\/(word|deck|grid)(\/|$)/.test(url.pathname)) return;

  if (req.mode === "navigate") {
    event.respondWith(caches.open(CACHE).then((c) => c.match("./index.html").then((hit) => hit || fetch(req))));
    return;
  }
  event.respondWith(
    caches.open(CACHE).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req)))
  );
});
