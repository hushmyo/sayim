const V='sayim-secfix-v1';
const PREFIX='sayim-';
const BASE=new URL(self.registration.scope);
const DOSYA=['./','index.html','app.js','manifest.json','icon-192.png','icon-512.png'];
const ASSETS=new Set(DOSYA.map(p=>new URL(p,BASE).href));
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(DOSYA)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(n=>(n==='sayim'||n.startsWith(PREFIX))&&n!==V).map(n=>caches.delete(n)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin||!ASSETS.has(u.href))return;
  e.respondWith((async()=>{
    const c=await caches.open(V);
    try{const r=await fetch(e.request,{cache:'no-cache'});if(r.ok)await c.put(e.request,r.clone());return r}
    catch(_){const m=await c.match(e.request);if(m)return m;if(e.request.mode==='navigate')return(await c.match(new URL('index.html',BASE)))||Response.error();return Response.error()}
  })())
});
