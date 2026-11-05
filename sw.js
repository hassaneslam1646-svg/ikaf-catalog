// تطبيق مخزون إكاف — نسخة مبسطة ومتوافقة مع آيفون (Safari)
// - الصفحة: من النت الأول، ولو مفيش نت (أو النت واقف أكتر من 5 ثواني) آخر نسخة محفوظة
// - الصور: من النت عادي زي اللينك بالظبط، ولو مفيش نت بس بنعرض المحفوظ
const V = 'ikaf-v4';
const PAGE = V + '-page', IMG = V + '-img';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(PAGE)
    .then(c => Promise.all(['./', 'logo.png', 'icons/splash.webp', 'icons/icon-192.png']
      .map(u => fetch(u, {cache: 'reload'}).then(r => r.ok && c.put(u, r)).catch(() => {}))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => !k.startsWith(V)).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function timeout(ms) { return new Promise(res => setTimeout(() => res(null), ms)); }

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (u.origin !== location.origin) return;              // أي حاجة بره الموقع: المتصفح يتعامل معاها لوحده

  if (req.mode === 'navigate') {
    const net = fetch(req).then(r => {
      if (r && r.ok) { const cl = r.clone(); caches.open(PAGE).then(c => c.put('./', cl)).catch(() => {}); }
      return r;
    });
    e.respondWith(
      Promise.race([net.catch(() => null), timeout(5000)]).then(r => {
        if (r && r.ok) return r;
        return caches.open(PAGE).then(c => c.match('./')).then(h => h || net);
      })
    );
    return;
  }

  // الصور: النت الأول دايمًا (مفيش أي تأخير)، والمحفوظ بس لو النت فشل
  if (/\.(webp|jpe?g|png)$/i.test(u.pathname)) {
    e.respondWith(
      fetch(req).then(r => {
        if (r && r.ok && /\/(thumb|icons)\/|logo\.png$/.test(u.pathname)) {
          const cl = r.clone();
          e.waitUntil(caches.open(IMG).then(c => c.put(req, cl)).catch(() => {}));
        }
        return r;
      }).catch(() => caches.match(req).then(h => h || Response.error()))
    );
  }
});
