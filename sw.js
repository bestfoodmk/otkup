/* Service worker: ја чува целата апликација на уредот за да се отвора и без интернет.
   Податоците (списоци, внесови) НЕ минуваат тука — тие одат директно до SharePoint кога има мрежа. */
var VERZIJA = "otkup-2026-10-08a";
var DATOTEKI = [
  "./", "./index.html", "./app.js", "./zaednicko.js", "./config.js", "./manifest.webmanifest",
  "./lib/jsQR.js", "./lib/msal-browser.min.js",
  "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"
];
self.addEventListener("install", function(e){
  e.waitUntil(caches.open(VERZIJA).then(function(c){ return c.addAll(DATOTEKI); }).then(function(){ return self.skipWaiting(); }));
});
self.addEventListener("activate", function(e){
  e.waitUntil(caches.keys().then(function(ks){
    return Promise.all(ks.filter(function(k){ return k !== VERZIJA; }).map(function(k){ return caches.delete(k); }));
  }).then(function(){ return self.clients.claim(); }));
});
self.addEventListener("fetch", function(e){
  var url = new URL(e.request.url);
  // Само сопствените датотеки од кешот; Graph/login барањата одат директно на мрежа.
  if(url.origin !== self.location.origin || e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request, {ignoreSearch:true}).then(function(odg){
      var mreza = fetch(e.request).then(function(r){
        if(r && r.ok){ var kopija = r.clone(); caches.open(VERZIJA).then(function(c){ c.put(e.request, kopija); }); }
        return r;
      }).catch(function(){ return odg; });
      return odg || mreza;
    })
  );
});
