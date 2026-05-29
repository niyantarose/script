import "./core/mangaUpdatesClient.js";
import "./core/titleAnalysis.js";
import "./backgrounds/taiwan.js";
import "./backgrounds/aladin.js";

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
