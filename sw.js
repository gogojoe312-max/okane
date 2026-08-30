const C='okane-2026-08-30-a';
self.addEventListener('install',e=>{self.skipWaiting();});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.map(k=>k!==C&&caches.delete(k)))));self.clients.claim();});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET'||!e.request.url.startsWith(self.location.origin))return;
  e.respondWith(
    fetch(e.request,{cache:'no-store'}).then(r=>{const cl=r.clone();caches.open(C).then(c=>c.put(e.request,cl));return r;})
    .catch(()=>caches.match(e.request))
  );
});
