// ==========================================
// SOLLO ERP - OFFLINE SERVICE WORKER (v5.2 Enterprise)
// ==========================================

// --- NEW CODE: Bump the version to 5.5 to clear the cache! ---
const CACHE_NAME = 'sollo-erp-v5.5-enterprise';
// --- END OF NEW CODE ---

const TIMEOUT_MS = 3000; // 3-second timeout to defeat Lie-Fi

// Critical assets to cache immediately upon installation
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './style.css',
    './app.js',
    './ui.js',
    './db.js',
    './utils.js',
    './cloud.js',
    './icon-192.png', // Add this!
    './icon-512.png', // Add this!
    './manifest.json',
    'https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@24,400,0,0',
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/flatpickr/4.6.13/flatpickr.min.css',
    'https://cdnjs.cloudflare.com/ajax/libs/flatpickr/4.6.13/flatpickr.min.js'
];

// 1. Install Event
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[Service Worker] Caching v5.2 enterprise assets');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
    self.skipWaiting(); 
});

// 2. Activate Event
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('[Service Worker] Clearing old cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
    self.clients.claim(); 
});

// 3. Fetch Event (Split Strategy for Ultimate Speed)
self.addEventListener('fetch', (event) => {
    // Strictly ignore blob and data URLs
    if (event.request.url.startsWith('blob:') || event.request.url.startsWith('data:')) return;
    if (event.request.method !== 'GET') return;

    const url = new URL(event.request.url);
    const isStaticAsset = url.pathname.endsWith('.css') || url.pathname.endsWith('.js') || url.pathname.endsWith('.png') || url.hostname.includes('fonts') || url.hostname.includes('cdnjs');

    if (isStaticAsset) {
        // STRATEGY 1: CACHE-FIRST (For CSS, JS, Fonts, Images)
        // Instantly loads from hardware. Updates cache silently in the background.
        event.respondWith(
            caches.match(event.request).then((cachedResponse) => {
                const fetchPromise = fetch(event.request).then((networkResponse) => {
                    if (networkResponse && networkResponse.status === 200) {
                        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, networkResponse.clone()));
                    }
                    return networkResponse;
                }).catch(() => { console.log('[Service Worker] Offline: Using static cached asset.'); });

                return cachedResponse || fetchPromise;
            })
        );
    } else {
        // STRATEGY 2: NETWORK-FIRST with Timeout Fallback (For HTML / App Shell)
        // Ensures you always get the latest UI updates if online.
        event.respondWith(
            new Promise((resolve) => {
                let isResolved = false;

                const timeoutId = setTimeout(() => {
                    if (!isResolved) {
                        caches.match(event.request).then((cachedResponse) => {
                            if (cachedResponse) {
                                isResolved = true;
                                resolve(cachedResponse);
                            }
                        });
                    }
                }, TIMEOUT_MS);

                fetch(event.request).then((networkResponse) => {
                    clearTimeout(timeoutId);
                    if (!isResolved) {
                        if (networkResponse && networkResponse.status === 200 && !event.request.url.startsWith('chrome-extension')) {
                            const responseToCache = networkResponse.clone();
                            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseToCache));
                        }
                        isResolved = true;
                        resolve(networkResponse);
                    }
                }).catch(() => {
                    clearTimeout(timeoutId);
                    if (!isResolved) {
                        caches.match(event.request).then((cachedResponse) => {
                            if (cachedResponse) {
                                resolve(cachedResponse);
                            } else {
                                if (event.request.mode === 'navigate' || event.request.headers.get('accept').includes('text/html')) {
                                    caches.match('./index.html', { ignoreSearch: true }).then((cachedHtml) => {
                                        resolve(cachedHtml || new Response('<div style="text-align:center; padding:50px; font-family:sans-serif;"><h2>App Offline</h2><p>Please check your connection and reload.</p></div>', { status: 503, headers: { 'Content-Type': 'text/html' } }));
                                    });
                                } else {
                                    resolve(new Response('', { status: 503, statusText: 'Service Unavailable' }));
                                }
                            }
                        });
                    }
                });
            })
        );
    }
});
