/* Orçamento de Chapas Prime — service worker
   Para forçar atualização nos aparelhos: mude o número da versão e faça o commit. */
const VERSAO = 'orcamento-chapas-v8';
const ARQUIVOS = ['.', 'index.html', 'calc.js', 'excel.js', 'modelo-orcamento.xlsx', 'manifest.json', 'icon.svg', 'icon-maskable.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('n8n.mcamargo.uk')) return;   // conferência com a Programação: sempre ao vivo
  e.respondWith(
    fetch(e.request).then(r => {
      const copia = r.clone();
      caches.open(VERSAO).then(c => c.put(e.request, copia)).catch(() => {});
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
