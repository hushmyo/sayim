const AD="sayim-v4";
const DOSYALAR=["./","index.html","manifest.json"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(AD).then(c=>c.addAll(DOSYALAR)).then(()=>self.skipWaiting()));});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==AD).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=="GET"||u.origin!==location.origin) return;   // Microsoft/Graph istekleri aga gider
  e.respondWith(fetch(e.request).then(r=>{const k=r.clone();caches.open(AD).then(c=>c.put(e.request,k));return r;}).catch(()=>caches.match(e.request)));
});
