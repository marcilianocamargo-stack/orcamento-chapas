/* Orçamento de Chapas Prime — service worker
   Para forçar atualização nos aparelhos: mude o número da versão (aqui e o ?v= dos scripts no index.html) e faça o commit. */
const VERSAO = 'orcamento-chapas-v13';
const ARQUIVOS = ['.', 'index.html', 'calc.js?v=13', 'excel.js?v=13', 'modelo-orcamento.xlsx', 'manifest.json', 'icon.svg', 'icon-maskable.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ARQUIVOS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
/* Sempre busca na rede (o GitHub Pages deixa o navegador guardar os arquivos por 10 min — sem o
   "no-cache" a página nova podia abrir com o calc.js velho). Sem internet, usa a cópia guardada. */
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('n8n.mcamargo.uk')) return;   // conferência com a Programação: sempre ao vivo
  const mesmaOrigem = new URL(e.request.url).origin === self.location.origin;
  const rede = mesmaOrigem ? fetch(e.request.url, { cache: 'no-cache', credentials: 'same-origin' }) : fetch(e.request);
  e.respondWith(
    rede.then(r => {
      if (r.ok) { const copia = r.clone(); caches.open(VERSAO).then(c => c.put(e.request, copia)).catch(() => {}); }
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
