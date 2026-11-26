// __APP_NAME__ para iPhone: guarda la app completa en el dispositivo para uso sin conexión.
// Cuando se publica una versión nueva en GitHub, este archivo cambia, se descarga solo lo nuevo
// (reutilizando los archivos .wasm/.js con huella que no cambiaron) y se limpia solo el caché viejo
// de esta misma app sin tocar las demás apps de OfficeCraft.
const VERSION = "__VERSION__";
const ENGINE_VER = "__ENGINE_VER__";
const PREFIX = "__CACHE_PREFIX__";
const CACHE = PREFIX + VERSION;
const PRECACHE = __PRECACHE__;
const WASM = "./__WASM__";
const WASM_BYTES = __WASM_BYTES__;

async function broadcast(message) {
  const windows = await self.clients.matchAll({ includeUncontrolled: true, type: "window" });
  windows.forEach((client) => client.postMessage(message));
}

async function fetchWasm() {
  const response = await fetch(new Request(WASM, { cache: "reload" }));
  if (!response.ok) throw new Error("No se pudo bajar __APP_NAME__ (" + response.status + ")");
  if (!response.body) return response;
  const reader = response.body.getReader();
  const chunks = [];
  let loaded = 0;
  let reported = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    loaded += value.byteLength;
    if (loaded - reported > 384 * 1024) {
      reported = loaded;
      broadcast({ type: "progress", loaded, total: WASM_BYTES });
    }
  }
  broadcast({ type: "progress", loaded, total: WASM_BYTES });
  return new Response(new Blob(chunks, { type: "application/wasm" }), {
    headers: { "Content-Type": "application/wasm" },
  });
}

const HASHED = /-[0-9a-f]{16}(_bg|_loader)?\.(wasm|js|ico)$/;
async function reuse(cache, url) {
  if (!HASHED.test(url)) return false;
  const keys = (await caches.keys()).filter((k) => k.startsWith(PREFIX) && k !== CACHE);
  for (const key of keys) {
    const oldCache = await caches.open(key);
    const old = await oldCache.match(url);
    if (old) {
      await cache.put(url, old);
      return true;
    }
  }
  return false;
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE);
      const rest = [];
      for (const url of PRECACHE) {
        if (url !== WASM && !(await reuse(cache, url))) rest.push(url);
      }
      if (rest.length) {
        await cache.addAll(rest.map((url) => new Request(url, { cache: "reload" })));
      }
      if (!(await reuse(cache, WASM))) {
        await cache.put(WASM, await fetchWasm());
      }
      await self.skipWaiting();
    })().catch((error) => {
      broadcast({
        type: "install-failed",
        message: String((error && error.message) || error),
      });
      throw error;
    })
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const older = (await caches.keys()).filter((key) => key.startsWith(PREFIX) && key !== CACHE);
      await Promise.all(older.map((key) => caches.delete(key)));
      await self.clients.claim();
      broadcast({ type: "activated", version: ENGINE_VER });
    })()
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      (async () => {
        const cache = await caches.open(CACHE);
        return (await cache.match("./index.html")) || fetch(request);
      })()
    );
    return;
  }

  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE);
      const hit =
        (await cache.match(request, { ignoreSearch: true })) ||
        (await caches.match(request, { ignoreSearch: true }));
      if (hit) return hit;
      const response = await fetch(request);
      if (response.ok && response.type === "basic") {
        cache.put(request, response.clone()).catch(() => {});
      }
      return response;
    })()
  );
});
