/* Offline support: the app itself is cached; forms and lists are cached by data.js */
const SHELL = "eagle-shell-v4";
const FILES = ["./", "index.html", "css/app.css", "manifest.json", "icons/icon-192.png",
  "js/app.js", "js/engine.js", "js/data.js", "js/backend.js", "js/config.js",
  "lib/jszip.min.js", "lib/qrcode.js", "lib/xlsx.mini.min.js",
  "lib/firebase-app.js", "lib/firebase-auth.js", "lib/firebase-firestore.js",
  "lib/pdf-lib.min.js", "lib/fontkit.umd.min.js", "js/pdfdraw.js", "js/arabic.js",
  "fonts/NotoNaskhArabic-400.ttf", "fonts/NotoNaskhArabic-700.ttf"];
self.addEventListener("install", e => { e.waitUntil(caches.open(SHELL).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("eagle-shell") && k !== SHELL).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;     // Firebase & GitHub: network
  // network first (always newest app), cached copy when offline
  e.respondWith(fetch(e.request, { cache: "no-cache" }).then(r => { const c = r.clone(); caches.open(SHELL).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request).then(r => r || caches.match("index.html"))));
});
