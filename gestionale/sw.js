const CACHE = 'asglab-v16';
const ASSETS = [
  '/gestionale/',
  '/gestionale/index.html',
  '/gestionale/xlsx.min.js'
];

self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){ return c.addAll(ASSETS); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){return k!==CACHE;}).map(function(k){return caches.delete(k);}));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e){
  // Solo GET, solo stesso dominio
  if(e.request.method !== 'GET') return;
  if(!e.request.url.startsWith(self.location.origin)) return;
  // Nextcloud: sempre network, mai cache
  if(e.request.url.includes('remote.php') || e.request.url.includes('dash.asglab.it')) return;

  e.respondWith(
    caches.match(e.request).then(function(cached){
      var net = fetch(e.request).then(function(res){
        if(res && res.status === 200){
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      }).catch(function(){ return cached; });
      // Serve cache subito se disponibile, aggiorna in background
      return cached || net;
    })
  );
});
