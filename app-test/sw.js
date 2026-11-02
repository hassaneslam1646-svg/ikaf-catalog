// تطبيق مخزون إكاف: الصفحة بتتجاب من النت الأول (عشان المخزون يبقى أحدث حاجة)
// ولو النت فاصل بتفتح آخر نسخة اتحمّلت. الصور بتتحفظ وتتحدث في الخلفية.
const V = 'ikaf-v3';
const PAGE = V + '-page', IMG = V + '-img', LIB = V + '-lib';
const CORE = ['./', 'logo.png', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/splash.webp'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(PAGE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => !k.startsWith(V)).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// النت لو بطيء جدًا (واي فاي من غير إنترنت مثلًا) منستناش أكتر من 5 ثواني: نفتح آخر نسخة،
// والنسخة الجديدة بتكمل تحميل في الخلفية وتظهر في الفتحة الجاية
async function networkFirst(req, event) {
  const c = await caches.open(PAGE);
  const net = fetch(req, {cache: 'no-store'}).then(r => { if (r.ok) c.put('./', r.clone()); return r; });
  if (event) event.waitUntil(net.catch(() => {}));
  const cached = await c.match('./');
  if (!cached) return net.catch(() => Response.error());
  const timeout = new Promise(res => setTimeout(() => res(null), 5000));
  try {
    const r = await Promise.race([net, timeout]);
    return (r && r.ok) ? r : cached;
  } catch (err) {
    return cached;
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
  if (req.mode === 'navigate') { e.respondWith(networkFirst(req, e)); return; }
  if (u.origin === location.origin) {
    if (/\.(webp|jpe?g|png)$/i.test(u.pathname)) { e.respondWith(staleWhileRevalidate(req, IMG)); return; }
    if (/\/(vendor\/|manifest)/.test(u.pathname)) { e.respondWith(staleWhileRevalidate(req, LIB)); return; }
    return;
  }
  if (/fonts\.(googleapis|gstatic)\.com$/.test(u.hostname)) e.respondWith(staleWhileRevalidate(req, LIB));
});
