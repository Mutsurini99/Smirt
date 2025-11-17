const CACHE_NAME = 'smirt-pwa-v4'; // ⬆️ Incrementato per correzioni errori
const urlsToCache = [
  './',
  './index.html',
  './manifest.json',
  './SMIRT_Icon_Final.svg',
  './js/user-manager.js',
  './users.json'
];

// Install event - cache resources with better error handling
self.addEventListener('install', (event) => {
  console.log('[SW] Install evento');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[SW] Cache aperta');
        // Cache resources individually to handle failures gracefully
        return Promise.allSettled(
          urlsToCache.map(url => {
            return cache.add(url).catch(error => {
              console.warn(`[SW] Impossibile cachare ${url}:`, error);
              return null; // Continue with other resources
            });
          })
        );
      })
      .then(() => {
        console.log('[SW] Risorse cachate (con possibili fallimenti gestiti)');
        self.skipWaiting();
      })
      .catch((error) => {
        console.error('[SW] Errore durante il caching:', error);
      })
  );
});

// Activate event - clean up old caches
self.addEventListener('activate', (event) => {
  console.log('[SW] Activate evento');
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Eliminando cache obsoleta:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
      console.log('[SW] Service worker attivato');
      self.clients.claim();
    })
  );
});

// Fetch event - serve from cache when offline with improved error handling
self.addEventListener('fetch', (event) => {
  // Skip non-GET requests
  if (event.request.method !== 'GET') {
    return;
  }

  // Skip external APIs (Google Sheets, etc.)
  if (event.request.url.includes('script.google.com') || 
      event.request.url.includes('googleapis.com') ||
      event.request.url.includes('chrome-extension') ||
      event.request.url.includes('tailwindcss.com')) {
    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        // Return cached version or fetch from network
        if (response) {
          console.log('[SW] Serving from cache:', event.request.url);
          return response;
        }
        
        console.log('[SW] Fetching from network:', event.request.url);
        return fetch(event.request).then((response) => {
          // Don't cache non-successful responses or non-basic types
          if (!response || response.status !== 200 || response.type !== 'basic') {
            return response;
          }

          // Clone the response
          const responseToCache = response.clone();

          // Add to cache with error handling
          caches.open(CACHE_NAME)
            .then((cache) => {
              return cache.put(event.request, responseToCache);
            })
            .catch((error) => {
              console.warn('[SW] Errore aggiornamento cache:', error);
            });

          return response;
        });
      })
      .catch((error) => {
        console.log('[SW] Fetch fallback for:', event.request.url);
        
        // Fallback for navigation requests
        if (event.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
        
        // For other requests, return a generic offline response
        return new Response('Contenuto non disponibile offline', {
          status: 200,
          headers: { 'Content-Type': 'text/plain' }
        });
      })
  );
});

// Background sync for offline form submissions
self.addEventListener('sync', (event) => {
  console.log('[SW] Background sync:', event.tag);
  
  if (event.tag === 'background-sync-reports') {
    event.waitUntil(
      syncOfflineReports()
    );
  }
});

// Sync offline reports when connection is restored
async function syncOfflineReports() {
  try {
    console.log('[SW] Sincronizzazione rapporti offline...');
    
    // Get offline reports from IndexedDB
    const reports = await getOfflineReports();
    
    if (reports && reports.length > 0) {
      for (const report of reports) {
        try {
          await submitReport(report);
          await removeOfflineReport(report.id);
          console.log('[SW] Rapporto sincronizzato:', report.id);
        } catch (error) {
          console.error('[SW] Errore sincronizzazione rapporto:', error);
        }
      }
    }
  } catch (error) {
    console.error('[SW] Errore generale sincronizzazione:', error);
  }
}

// IndexedDB helpers (simplified for this example)
async function getOfflineReports() {
  // In a real implementation, this would fetch from IndexedDB
  return [];
}

async function removeOfflineReport(id) {
  // In a real implementation, this would remove from IndexedDB
  console.log('[SW] Report rimosso:', id);
}

async function submitReport(report) {
  // In a real implementation, this would submit to Google Sheets
  console.log('[SW] Submitting report:', report);
}

// Push notifications (for future enhancements)
self.addEventListener('push', (event) => {
  console.log('[SW] Push message received');
  
  const options = {
    body: event.data ? event.data.text() : 'Nuovo aggiornamento disponibile',
    icon: './SMIRT_Icon_Final.svg', // ✅ Aggiornato a SVG
    badge: './SMIRT_Icon_Final.svg', // ✅ Aggiornato a SVG
    vibrate: [200, 100, 200],
    data: {
      dateOfArrival: Date.now(),
      primaryKey: 1
    },
    actions: [
      {
        action: 'explore',
        title: 'Visualizza',
        icon: './SMIRT_Icon_Final.svg' // ✅ Aggiornato a SVG
      },
      {
        action: 'close',
        title: 'Chiudi',
        icon: './SMIRT_Icon_Final.svg' // ✅ Aggiornato a SVG
      }
    ]
  };

  event.waitUntil(
    self.registration.showNotification('SMIRT - Rapporti Intervento', options)
  );
});

// Notification click handler
self.addEventListener('notificationclick', (event) => {
  console.log('[SW] Notification click received:', event.action);

  event.notification.close();

  if (event.action === 'explore') {
    event.waitUntil(
      clients.openWindow('./index.html')
    );
  }
});

// Message handling from main thread
self.addEventListener('message', (event) => {
  console.log('[SW] Message received:', event.data);
  
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  
  if (event.data && event.data.type === 'GET_VERSION') {
    event.ports[0].postMessage({ version: CACHE_NAME });
  }
});

console.log('[SW] Service Worker caricato - Versione con modifiche date e buoni lavoro');
