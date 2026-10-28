/*
 * Service worker de cobro.html (alcance: solo esa página).
 * Guarda la página, el ícono y config.json para que la app instalada abra
 * aunque no haya internet (muestra "Sin conexión" hasta que vuelva).
 * Página y config: primero la red (así toma cambios al subirlos), y si no
 * hay conexión, la copia. Las consultas al backend NO pasan por acá
 * (son de otro dominio) y nunca se guardan.
 */
const CACHE_NAME = "cobro-qr-v2";
const ARCHIVOS = ["cobro.html", "manifest-cobro.json", "config.json", "icon-192.png", "icon-512.png", "admin/vendor/qrcode.min.js"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((c) => c.addAll(ARCHIVOS)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((nombres) => Promise.all(
      nombres.filter((n) => n.startsWith("cobro-qr-") && n !== CACHE_NAME).map((n) => caches.delete(n))
    ))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(event.request)
      .then((resp) => {
        if (resp && resp.ok) { const copia = resp.clone(); caches.open(CACHE_NAME).then((c) => c.put(event.request, copia)); }
        return resp;
      })
      .catch(() => caches.match(event.request, { ignoreSearch: true }))
  );
});
