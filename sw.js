/* ═══════════════════════════════════════════════════════════════════════════
   APR REMO ENGENHARIA — Service Worker
   Estratégia:
     • same-origin  → stale-while-revalidate (cache imediato + atualização bg)
     • CDN externo  → cache-first (fontes Barlow + html2pdf.js)
     • Drive / sync → sem interceptação (requer rede por design)
   ═══════════════════════════════════════════════════════════════════════════ */

'use strict';

const CACHE_NAME = 'apr-remo-v1';

// Assets pré-cacheados no install (shell mínimo do app)
const PRE_CACHE = [
  './',
  './index.html',
];

// Hosts CDN cacheados de forma lazy (cache-first)
const CDN_HOSTS = [
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'cdn.jsdelivr.net',
  'unpkg.com',
];

// ── Install: pré-cachear o shell ─────────────────────────────────────────────
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) { return cache.addAll(PRE_CACHE); })
      .then(function() { return self.skipWaiting(); })
  );
});

// ── Activate: limpar caches de versões anteriores ────────────────────────────
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys()
      .then(function(keys) {
        return Promise.all(
          keys
            .filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
        );
      })
      .then(function() { return self.clients.claim(); })
  );
});

// ── Fetch: rotear por estratégia ─────────────────────────────────────────────
self.addEventListener('fetch', function(event) {
  var req = event.request;

  // Ignorar requisições não-GET e protocolos não-http
  if (req.method !== 'GET') return;
  var url;
  try { url = new URL(req.url); } catch(e) { return; }
  if (!url.protocol.startsWith('http')) return;

  // Same-origin → stale-while-revalidate
  if (url.origin === self.location.origin) {
    event.respondWith(swrStrategy(req));
    return;
  }

  // CDN externo → cache-first
  var isCdn = CDN_HOSTS.some(function(h) { return url.hostname.indexOf(h) !== -1; });
  if (isCdn) {
    event.respondWith(cacheFirstStrategy(req));
    return;
  }

  // Demais (Drive sync, etc.) → sem interceptação
});

// ── Stale-while-revalidate ───────────────────────────────────────────────────
// Entrega o cache imediatamente; atualiza o cache em background via rede.
// Na próxima visita o conteúdo já estará atualizado.
function swrStrategy(req) {
  return caches.open(CACHE_NAME).then(function(cache) {
    return cache.match(req).then(function(cached) {
      var networkUpdate = fetch(req).then(function(res) {
        if (res && res.status === 200) cache.put(req, res.clone());
        return res;
      }).catch(function() { return null; });

      return cached || networkUpdate;
    });
  });
}

// ── Cache-first ──────────────────────────────────────────────────────────────
// Retorna do cache; se ausente, busca na rede, armazena e retorna.
// Ideal para fontes e scripts CDN que raramente mudam.
function cacheFirstStrategy(req) {
  return caches.open(CACHE_NAME).then(function(cache) {
    return cache.match(req).then(function(cached) {
      if (cached) return cached;

      return fetch(req).then(function(res) {
        // Não cachear respostas opacas (possíveis erros silenciosos)
        if (res && res.status === 200 && res.type !== 'opaque') {
          cache.put(req, res.clone());
        }
        return res;
      }).catch(function() {
        // Recurso CDN indisponível offline — retorna erro de rede
        return new Response('', { status: 503, statusText: 'Offline' });
      });
    });
  });
}

// ── Mensagem SKIP_WAITING ────────────────────────────────────────────────────
// Disparada pelo banner "Nova versão disponível" no index.html
self.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
