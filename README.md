# Ma collection — Gentle Mates TCG

Application web pour suivre sa collection de cartes Gentle Mates TCG (extension « Tout Commence Ici »).
Tout fonctionne dans le navigateur : aucune donnée n'est envoyée sur un serveur. Chaque personne qui
ouvre la page garde sa propre collection dans son propre navigateur.

## Ce que le site utilise

- `index.html` — l'application (un seul fichier)
- `images-cartes/` — les visuels des cartes
- `logo/` — logos et symboles d'essence

Les autres fichiers (`build_html.py`, `build4.py`, le classeur Excel) servent seulement à régénérer
`index.html` quand on change les données. Voir `REPRISE-CLAUDE-CODE.md`.

## Publier avec GitHub Pages

1. Crée un compte sur https://github.com, puis un nouveau dépôt **public** (bouton « New »).
2. Envoie-y les fichiers du dossier. Le plus simple : installe **GitHub Desktop**
   (https://desktop.github.com), puis « File » → « Add local repository » → choisis ce dossier →
   « Publish repository ». Le fichier `.gitignore` empêche d'envoyer ta sauvegarde perso.
   (Le glisser-déposer sur le site de GitHub est limité à 100 fichiers, et il y en a plus.)
3. Sur la page du dépôt : « Settings » → « Pages » → « Build and deployment » →
   Source « Deploy from a branch », branche `main`, dossier `/ (root)` → « Save ».
4. Attends 1 à 2 minutes. L'adresse du site apparaît en haut de la page « Pages » :
   `https://TON-PSEUDO.github.io/NOM-DU-DEPOT/`

Envoie ce lien à tes amis. Pour mettre à jour le site plus tard : régénère `index.html`
(`python build_html.py`), puis dans GitHub Desktop fais « Commit » et « Push origin ».

## Retrouver sa collection sur tous ses appareils

Les données sont stockées dans le navigateur de chaque appareil. Pour les synchroniser :

1. Dans l'appli, ouvre une extension, descends en bas de page et clique sur **Synchroniser**.
2. Clique sur le lien « Crée un jeton sur GitHub » (il faut être connecté à ton compte), garde
   l'option « gist » cochée et rien d'autre, choisis « No expiration », puis « Generate token ».
3. Copie le jeton (il commence par `ghp_`), colle-le dans l'appli et clique sur **Activer**.
4. Sur chaque autre appareil : ouvre le site, « Synchroniser », colle le même jeton.

Ta collection est alors enregistrée dans un fichier secret (« gist ») de ton compte GitHub, mis à
jour tout seul quand tu coches une carte. Si les deux appareils ont changé chacun de leur côté,
l'appli te demande quelle version garder. Les anciennes versions restent consultables dans
l'historique du gist.

Ne partage jamais ton jeton. Si tu le perds ou s'il fuit, supprime-le sur
https://github.com/settings/tokens et recommence.

Sans jeton, l'appli marche comme avant (données locales). Pour une sauvegarde à la main, utilise
« Exporter mes données » et « Importer un fichier ».
