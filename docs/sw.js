const C="pvh-shell-v3";
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png"])).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  if(e.request.method!=="GET")return;
  const u=new URL(e.request.url);
  /* Only the app shell is cached. Data fetches (OneDrive and friends) always go
     to the network, so a new build is actually seen and no record data is left
     behind in the cache. */
  if(u.origin!==self.location.origin)return;
  if(/\.json(\?|$)/i.test(u.pathname+u.search))return;
  e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).then(res=>{const cl=res.clone();caches.open(C).then(c=>c.put(e.request,cl));return res;}).catch(()=>caches.match("index.html"))));
});