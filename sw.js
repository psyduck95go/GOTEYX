/* 可愛反派狠餃色 Service Worker
   - 網頁本身：網路優先（有新版就拿新版，離線時用快取）
   - 圖片、字型、CDN 程式庫：先用快取、背景更新
   更新網站內容後，請把 VERSION 改掉，讓舊快取被清除。 */
const VERSION = 'v2';
const CACHE = `goteyx-${VERSION}`;
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
  './images/dumplings/d1-cabbage-pork.svg',
  './images/dumplings/d2-shrimp-scallop.svg',
  './images/dumplings/d3-chili-oil-wonton.svg',
  './images/dumplings/d4-shiitake-veg.svg'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('goteyx-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (!/^https?:$/.test(url.protocol)) return;

  // 頁面導覽：網路優先
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./index.html', copy)); return res; })
        .catch(() => caches.match('./index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // 其他資源：快取優先，背景更新
  event.respondWith(
    caches.open(CACHE).then(cache => cache.match(req).then(cached => {
      const network = fetch(req).then(res => {
        if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
        return res;
      }).catch(() => cached);
      return cached || network;
    }))
  );
});
