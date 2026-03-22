const CACHE_NAME = 'courier-printer-v3'; // Bump version to force update

// Files to save for offline use
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './style.css',
    './script.js',
    './manifest.json',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js', 
    'https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap',
    'https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,1,0'
];

// Install event - Cache all assets
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
        .then(cache => {
            console.log('Opened cache');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
});

// Fetch event - Network First with Dynamic Caching for offline fonts & CDNs
self.addEventListener('fetch', event => {
    // Skip requests from browser extensions or unsupported protocols
    if (!event.request.url.startsWith('http')) return;

    event.respondWith(
        fetch(event.request)
        .then(networkResponse => {
            // Dynamically cache successful network responses (like hidden font files)
            if (networkResponse && networkResponse.status === 200) {
                const responseClone = networkResponse.clone();
                caches.open(CACHE_NAME).then(cache => {
                    cache.put(event.request, responseClone);
                });
            }
            return networkResponse;
        })
        .catch(() => {
            // If offline, fall back to the locally cached files
            return caches.match(event.request);
        })
    );
});

// Activate event - Clean up old caches if we update the version
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cache => {
                    if (cache !== CACHE_NAME) {
                        return caches.delete(cache);
                    }
                })
            );
        })
    );
});
