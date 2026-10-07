// AHQS Healthcare Service Worker — offline cache for core pages and resources
const CACHE_NAME = "ahqs-healthcare-v2";
const OFFLINE_URL = "/offline.html";

const CORE_ASSETS = [
  "/",
  "/index.html",
  "/about.html",
  "/services.html",
  "/products.html",
  "/resources.html",
  "/contact.html",
  "/faq.html",
  "/manifest.json",
  "/favicon.svg",
  "/ahqs-official-logo-hd.png",
  "/downloads/free-resources/ahqs-audit-tool-template.xlsx",
  "/downloads/free-resources/ahqs-risk-register-template.xlsx",
  "/downloads/free-resources/ahqs-capa-tracking-template.xlsx",
  "/downloads/free-resources/ahqs-quality-indicator-dashboard.xlsx",
  "/downloads/free-resources/ahqs-document-control-register.xlsx",
  "/downloads/free-resources/ahqs-competency-assessment-form.xlsx",
  "/assets/ahqs-conversion.css",
  "/assets/ahqs-conversion.js",
  "/offline.html"
];

// Install — cache core assets
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(CORE_ASSETS).catch(() => {});
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      );
    })
  );
  self.clients.claim();
});

// Fetch — cache-first for core assets, network-first for everything else
self.addEventListener("fetch", (event) => {
  const request = event.request;

  // Only handle GET requests
  if (request.method !== "GET") return;

  // Skip cross-origin requests
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  // Skip Formspree and analytics
  if (url.hostname.includes("formspree.io") || url.hostname.includes("google")) return;

  event.respondWith(
    caches.match(request).then((cached) => {
      const fetchPromise = fetch(request)
        .then((response) => {
          // Cache successful responses
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, clone).catch(() => {}));
          }
          return response;
        })
        .catch(() => cached);

      return cached || fetchPromise;
    })
  );
});
