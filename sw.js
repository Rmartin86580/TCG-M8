/* Service worker minimal : il ne garde RIEN en mémoire et n'active aucun mode hors connexion.
   Il existe seulement parce que certains navigateurs Android exigent un service worker pour
   proposer « Installer l'application » (icône + plein écran, voir manifest.webmanifest).
   Chaque requête passe directement par le réseau, comme sans service worker. */
self.addEventListener("install", () => self.skipWaiting());

self.addEventListener("activate", e => {
  // nettoie d'éventuels caches laissés par une version précédente
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith("m8-")).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Gestionnaire volontairement vide : le navigateur traite la requête normalement.
self.addEventListener("fetch", () => {});
