/*
 * Service Worker do APP Manutenção Eusébio (PWA).
 *
 * Objetivo: apenas tornar o app instalável na tela inicial do celular,
 * com uma rede de segurança simples para abrir mesmo sem internet no
 * momento do acesso.
 *
 * IMPORTANTE — estratégia "network-first" (rede primeiro), de propósito:
 * os dados do app (ordens de serviço, orçamento, etc.) são atualizados a
 * cada publicação via atualizar_app.bat. Um cache agressivo ("cache
 * first") prenderia o usuário numa versão antiga do app/dados até o
 * cache expirar. Por isso, sempre que há conexão, o Service Worker busca
 * a versão mais nova na rede; o cache só entra em ação como fallback
 * quando a rede falha (ex.: sem sinal).
 */

const CACHE_NAME = 'app-manutencao-eusebio-v1';
const PRECACHE_URLS = [
  './',
  './manifest.json',
  './img/logo-app.png',
  './img/logo-app-192.png',
];

self.addEventListener('install', (event) => {
  // Ativa a nova versão do Service Worker assim que instalada, sem
  // esperar todas as abas antigas fecharem.
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .catch(() => { /* pré-cache é best-effort; nunca bloqueia a instalação */ })
  );
});

self.addEventListener('activate', (event) => {
  // Remove caches de versões antigas deste Service Worker e assume o
  // controle das páginas já abertas imediatamente.
  event.waitUntil(
    caches.keys()
      .then((names) => Promise.all(
        names.filter((name) => name !== CACHE_NAME).map((name) => caches.delete(name))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  // Só intercepta requisições GET (a única coisa que este app faz).
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Sucesso na rede: serve a resposta mais recente e atualiza o
        // cache em segundo plano, para o fallback offline nunca ficar
        // muito desatualizado.
        const responseCopy = response.clone();
        caches.open(CACHE_NAME)
          .then((cache) => cache.put(event.request, responseCopy))
          .catch(() => {});
        return response;
      })
      .catch(() => {
        // Sem rede: tenta servir do cache; se também não tiver, deixa a
        // falha normal do navegador acontecer (não mascara o erro).
        return caches.match(event.request);
      })
  );
});
