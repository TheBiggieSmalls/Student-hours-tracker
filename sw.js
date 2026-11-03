const CACHE = "werkuren-v7";
const FILES = ["./", "index.html", "excel.html", "excel-sync.js", "exceljs.min.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png"];
// Alles wordt bij de installatie in het geheugen van het toestel gezet, zodat de app volledig offline werkt.
self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(u => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Eerst het geheugen (werkt offline), op de achtergrond verversen zodat updates bij de volgende keer openen komen.
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(hit => {
      const net = fetch(e.request).then(r => {
        if (r && r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
        return r;
      }).catch(() => hit || caches.match("index.html"));
      return hit || net;
    })
  );
});
