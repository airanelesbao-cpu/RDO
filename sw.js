// Service worker: o app abre na hora e funciona sem internet.
// Os arquivos do app ficam guardados no aparelho; a versão nova é baixada em segundo plano.
// Os dados (atividades, fotos) NÃO passam por aqui: vão ao Supabase e ficam guardados no aparelho pelo próprio app.
const VERSAO = 'rdo-v5', FONTES = 'rdo-fontes';
const ARQUIVOS = ['./', 'index.html', 'config.js', 'native.js', 'privacidade.html', 'termos.html', 'manifest.webmanifest', 'lib/pdf-lib.min.js', 'lib/jszip.min.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-180.png'];
// Chave única por arquivo (sem ?parâmetros), para a cópia nova sempre substituir a antiga.
const chave = u => { const x = new URL(u.url || u, location.href); x.search = ''; x.hash = ''; if (x.pathname.endsWith('/')) x.pathname += 'index.html'; return x.href; };
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => Promise.all(ARQUIVOS.map(async u => {
    const r = await fetch(new Request(u, { cache: 'reload' }));
    if (!r.ok) { if (u === 'index.html') throw new Error(u); return; }
    await c.put(chave(u), r);
  }))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO && k !== FONTES).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
async function avisar() {
  const cs = await self.clients.matchAll({ type: 'window' });
  cs.forEach(c => c.postMessage({ tipo: 'nova-versao' }));
}
async function fonte(req) {
  const c = await caches.open(FONTES), hit = await c.match(req);
  if (hit) return hit;
  try { const r = await fetch(req); if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }
  catch (e) { return Response.error(); }
}
async function app(e) {
  const req = e.request;
  const key = chave(req);
  const hit = await caches.match(key);
  const rede = fetch(req, req.mode === 'navigate' ? { cache: 'no-cache' } : undefined).then(async r => {
    if (r && r.ok) {
      const c = await caches.open(VERSAO);
      if (hit && (req.mode === 'navigate' || /index\.html$/.test(req.url))) {
        const a = await hit.clone().text(), b = await r.clone().text();
        if (a !== b) avisar();
      }
      if (!r.redirected) await c.put(key, r.clone());
    }
    return r;
  }).catch(() => null);
  if (hit && req.mode === 'navigate') {            // página principal: tenta a rede primeiro (até 4 s), assim a versão nova aparece logo
    const r = await Promise.race([rede, new Promise(res => setTimeout(() => res(null), 4000))]);
    if (r && r.ok && !r.redirected) return r;
    e.waitUntil(rede); return hit;
  }
  if (hit) { e.waitUntil(rede); return hit; }   // abre na hora, mesmo com sinal fraco
  const r = await rede;
  if (r) return r;
  return (req.mode === 'navigate' && await caches.match(chave('index.html'))) || Response.error();
}
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET') return;
  if (/(^|\.)fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) { e.respondWith(fonte(req)); return; }
  if (url.origin !== location.origin) return;                         // Supabase: direto na rede
  if (url.searchParams.has('chk')) return;                            // checagem de versão: sempre na rede
  if (/^\/(rest|auth|storage|functions)\//.test(url.pathname)) return;
  e.respondWith(app(e));
});
