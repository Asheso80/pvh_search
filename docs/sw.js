const C="pvh-shell-203bdc4a";
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png"])).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener("fetch",e=>{
  if(e.request.method!=="GET")return;
  const u=new URL(e.request.url);
  /* Only the app shell is cached. Data fetches (Dropbox and friends) always go
     to the network, so a new build is actually seen and no record data is left
     behind in the cache. */
  if(u.origin!==self.location.origin)return;
  if(/\.json(\?|$)/i.test(u.pathname+u.search))return;
  /* The page itself is network-first: a redeployed app is picked up on the next
     launch that has a signal, instead of the cached copy being served forever.
     The cache is the fallback, so offline still works, and a 3s cap means a
     flaky connection in the field falls back rather than hanging. */
  const isPage=e.request.mode==="navigate"||/(^|\/)(index\.html)?$/.test(u.pathname);
  if(isPage){
    e.respondWith(Promise.race([
      fetch(e.request).then(res=>{
        const cl=res.clone();
        caches.open(C).then(c=>c.put("index.html",cl));
        return res;
      }),
      new Promise(r=>setTimeout(()=>r(null),3000))
    ]).then(res=>res||caches.match("index.html",{ignoreSearch:true}).then(c=>c||fetch(e.request)))
      .catch(()=>caches.match("index.html",{ignoreSearch:true})));
    return;
  }
  e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).then(res=>{const cl=res.clone();caches.open(C).then(c=>c.put(e.request,cl));return res;}).catch(()=>caches.match("index.html"))));
});