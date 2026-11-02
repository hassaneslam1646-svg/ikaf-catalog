// تطبيق مخزون إكاف: الصفحة بتتجاب من النت الأول (عشان المخزون يبقى أحدث حاجة)
// ولو النت فاصل بتفتح آخر نسخة اتحمّلت. الصور بتتحفظ وتتحدث في الخلفية.
const V = 'ikaf-v1';
const PAGE = V + '-page', IMG = V + '-img', LIB = V + '-lib';
const CORE = ['./', 'logo.png', 'manifest.webmanifest', 'icons/icon-192.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(PAGE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => !k.startsWith(V)).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

async function networkFirst(req) {
  const c = await caches.open(PAGE);
  try {
    const r = await fetch(req, {cache: 'no-store'});
    if (r.ok) c.put('./', r.clone());
    return r;
  } catch (err) {
    return (await c.match('./')) || Response.error();
  }
}
async function staleWhileRevalidate(req, name) {
  const c = await caches.open(name);
  const hit = await c.match(req);
  const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => hit);
  return hit || net;
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (req.mode === 'navigate') { e.respondWith(networkFirst(req)); return; }
  if (u.origin === location.origin) {
    if (/\.(webp|jpe?g|png)$/i.test(u.pathname)) { e.respondWith(staleWhileRevalidate(req, IMG)); return; }
    if (/\/(vendor\/|manifest)/.test(u.pathname)) { e.respondWith(staleWhileRevalidate(req, LIB)); return; }
    return;
  }
  if (/fonts\.(googleapis|gstatic)\.com$/.test(u.hostname)) e.respondWith(staleWhileRevalidate(req, LIB));
});
