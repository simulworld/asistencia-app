// ════════════════════════════════════════════════════════════════
// SERVICE WORKER — Sistema de Asistencia v3.0
// Estrategia: Cache-first para assets, Network-first para la API
// ════════════════════════════════════════════════════════════════

const CACHE_NAME    = "asistencia-v3";
const ASSETS_CACHE  = ["./", "./index.html", "./manifest.json", "./sw.js"];

// ── Instalación: pre-cachea los archivos de la app ──────────────
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(ASSETS_CACHE))
      .then(() => self.skipWaiting()) // activa inmediatamente sin esperar cierre de tabs
  );
});

// ── Activación: limpia caches viejas ────────────────────────────
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: estrategia diferenciada ──────────────────────────────
self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);

  // Llamadas a la API de Apps Script → siempre red (nunca cachear datos reales)
  if (url.hostname.includes("script.google.com")) {
    event.respondWith(fetch(event.request));
    return;
  }

  // index.html → network-first para que las nuevas versiones se reflejen
  // pronto; si no hay conexión, usa la copia cacheada.
  if (url.pathname.endsWith("/index.html") || url.pathname.endsWith("/")) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response && response.status === 200 && response.type === "basic") {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // Otros assets → cache-first con fallback a red.
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response && response.status === 200 && response.type === "basic") {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      });
    })
  );
});
