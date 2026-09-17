/// <reference lib="webworker" />
import { build, files, version } from '$service-worker';

const worker = globalThis.self as unknown as ServiceWorkerGlobalScope;
const cacheName = `noffice-assets-${version}`;
const root = new URL('./', worker.registration.scope).href;
const assets = [...build, ...files].map((path) => new URL(path, root).href);
worker.addEventListener('install', (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(cacheName);
      await cache.addAll([...assets, root]);
    })(),
  );
});
worker.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      for (const key of await caches.keys())
        if (key.startsWith('noffice-assets-') && key !== cacheName) await caches.delete(key);
      await worker.clients.claim();
    })(),
  );
});
worker.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET' || new URL(request.url).origin !== worker.location.origin) return;
  if (!assets.includes(request.url) && request.mode !== 'navigate') return;
  event.respondWith(
    (async () => {
      const cache = await caches.open(cacheName);
      // Keep each open client on a consistent cached shell and asset version.
      const cached = await cache.match(request.mode === 'navigate' ? root : request.url);
      return cached || fetch(request);
    })(),
  );
});
