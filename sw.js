/* Service worker de "Mi día": guarda la app en el móvil para que abra sin conexión.
   Las tareas del calendario NO pasan por aquí: la app guarda las últimas que cargó.
   Cuando cambies index.html, sube también este archivo cambiando el número de VERSION. */
const VERSION = 'v1';
const CACHE = 'midia-' + VERSION;
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-180.png', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('midia-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* Pide a la red con un tiempo máximo; si no hay conexión, usa la copia guardada */
function networkFirst(req){
  return new Promise(resolve => {
    let done = false;
    const fallback = () => caches.match(req, {ignoreSearch: true})
      .then(r => r || caches.match('./index.html'))
      .then(r => { if (!done){ done = true; resolve(r || Response.error()); } });
    const timer = setTimeout(fallback, 4000);
    fetch(req).then(res => {
      clearTimeout(timer);
      if (res && res.ok){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      if (!done){ done = true; resolve(res); }
    }).catch(() => { clearTimeout(timer); fallback(); });
  });
}

function cacheFirst(req){
  return caches.match(req).then(hit => hit || fetch(req).then(res => {
    if (res && (res.ok || res.type === 'opaque')){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  }));
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin){
    e.respondWith(req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/') ? networkFirst(req) : cacheFirst(req));
  } else if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com'){
    e.respondWith(cacheFirst(req));
  }
  /* El resto (el calendario de Google) va directo a internet, sin guardar nada aquí */
});
