const CACHE = "budgie-v10"; // bump to force clients to refetch
const SHELL = ["./", "./index.html", "./app.js", "./manifest.json", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
// Network-first for the app itself (so updates show up), cache fallback offline.
// Never touch sync API calls (POST) or cross-origin requests other than the Excel reader.
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;
  if (!sameOrigin && url.hostname !== "cdnjs.cloudflare.com") return;
  e.respondWith(
    // same-origin files: always revalidate so an update isn't hidden behind the browser's HTTP cache
    fetch(req, sameOrigin ? { cache: "no-cache" } : undefined).then((res) => {
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then((r) => r || caches.match("./index.html")))
  );
});
