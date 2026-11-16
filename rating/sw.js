// Рейтинг влучань — офлайн-режим: сторінка звіту завжди спершу з мережі (щоб бачити свіжий звіт),
// без інтернету — остання збережена копія.
const CACHE = 'rv-v1';
const CORE = ['./', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).catch(() => {}));
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const sameOrigin = url.origin === self.location.origin;
  // 1) сторінка звіту: мережа -> кеш
  if (req.mode === 'navigate' || (sameOrigin && url.pathname.endsWith('/index.html'))) {
    e.respondWith((async () => {
      try {
        const fresh = await fetch(req, { cache: 'no-store' });
        const c = await caches.open(CACHE); c.put('./', fresh.clone());
        return fresh;
      } catch (err) {
        return (await caches.match('./')) || (await caches.match(req)) || Response.error();
      }
    })());
    return;
  }
  // 2) бібліотеки (Leaflet, контури областей) і свої файли: кеш -> мережа
  if (sameOrigin || url.hostname === 'unpkg.com' || url.hostname === 'cdn.jsdelivr.net') {
    e.respondWith((async () => {
      const hit = await caches.match(req);
      if (hit) return hit;
      try {
        const res = await fetch(req);
        if (res && (res.ok || res.type === 'opaque')) (await caches.open(CACHE)).put(req, res.clone());
        return res;
      } catch (err) { return hit || Response.error(); }
    })());
    return;
  }
  // 3) плитки карти: лише мережа (без накопичення в пам'яті телефона)
});
