// ASG LAB Gestionale - Service Worker v18
// Strategia: stale-while-revalidate per asset stessa-origine,
// network-only per WebDAV e cloud.asglab.it (no cache dati sensibili).
const CACHE = 'asglab-v18';
const ASSETS = [
  '/gestionale/',
  '/gestionale/index.html',
  '/gestionale/manifest.json',
  '/gestionale/icon-192.png',
  '/gestionale/xlsx.min.js'
];

self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){
      // addAll fallisce se anche un solo asset manca: usa add singolarmente per resilienza
      return Promise.allSettled(ASSETS.map(function(a){return c.add(a).catch(function(){});}));
    })
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

// URL/host che NON devono mai essere cachati (dati sensibili o auth-bound)
function shouldBypassCache(url){
  return url.includes('remote.php') ||      // WebDAV Nextcloud
         url.includes('/ocs/') ||            // Nextcloud OCS API
         url.includes('cloud.asglab.it') ||  // Tutto cloud NC
         url.includes('dash.asglab.it');     // Proxy dash
}

self.addEventListener('fetch', function(e){
  // Solo GET
  if(e.request.method !== 'GET') return;

  var url = e.request.url;

  // Bypass per dati sensibili (non passa nemmeno dalla cache)
  if(shouldBypassCache(url)) return;

  // Solo stesso dominio per cache
  if(!url.startsWith(self.location.origin)) return;

  e.respondWith(
    caches.match(e.request).then(function(cached){
      var net = fetch(e.request).then(function(res){
        // Cache solo risposte 200 OK basic (no opaque, no redirect, no error)
        if(res && res.status === 200 && res.type === 'basic'){
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      }).catch(function(){
        // Offline: ritorna cache se disponibile, altrimenti errore
        return cached || new Response('Offline', {status: 503, statusText: 'Offline'});
      });
      // Stale-while-revalidate: serve cache subito, aggiorna in background
      return cached || net;
    })
  );
});

// Permette al client di forzare skipWaiting (per il banner "nuova versione")
self.addEventListener('message', function(e){
  if(e.data && e.data.type === 'SKIP_WAITING'){
    self.skipWaiting();
  }
});
