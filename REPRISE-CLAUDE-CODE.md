# Projet : suivi de ma collection — TCG Gentle Mates

Fichier de reprise. Contexte, état actuel et décisions prises au fil d'une session
précédente. À lire avant toute modification.

---

## 1. Le besoin

Suivre une collection du jeu de cartes Gentle Mates (extension « Tout Commence Ici »),
consultable depuis un PC Windows, un iPhone et un iPad. L'utilisateur n'est pas
développeur : il faut privilégier ce qui marche sans maintenance.

Ce qu'il veut voir pour chaque carte : le visuel, le nom, la rareté, le type, l'essence,
s'il la possède, s'il en a des doublons, et si une version reverse existe.

---

## 2. Les fichiers

```
ma-collection.html      l'application, un seul fichier autonome
index.html              copie identique, écrite aussi par build_html.py — c'est ce que GitHub Pages sert
README.md, .gitignore   mode d'emploi de publication GitHub Pages (27/10/2026) ; .gitignore exclut la
                        sauvegarde perso (ma-collection-m8*.json) et dos-cartes/ (45 Mo, inutile au site)
build_html.py           le générateur qui la produit — NE PAS ÉDITER LE HTML DIRECTEMENT
build4.py               génère le classeur Excel
Suivi_collection_cartes.xlsx   le classeur, source des données cartes
images-cartes/          001.png … 127.png, P01.png … (visuels des cartes)
logo/                   habillage fourni par l'utilisateur (voir §5)
jeu-de-test/            données fictives pour tester l'onglet Comparer (3 amis + une collection « moi »,
                        mode d'emploi dans LISEZ-MOI.txt) ; ignoré par git, usage local seulement
dos-cartes/             133 dos de cartes personnalisés (PNG, 1003×1421) — voir ci-dessous
dos-cartes-planche.pdf  ces dos en planches A4, 9 par page à 63 mm de large (15 pages)
checklist-collection.pdf   feuille de suivi imprimable pour le classeur
```

**Dos de cartes personnalisés (17/09/2026).** Reproduction du visuel créé par l'utilisateur
(capture d'un dos Essence Orange 002/122) pour toutes les cartes : liseré violet, rangées de
glyphes haut/bas (extraites de sa capture), logo Gentle Mates noir (`logo/logo2.png` recoloré),
numéro, nom dans une pastille, symbole d'essence en noir (`logo/essence-*.png` recolorés ;
essence « Neutre » pour les cartes sans essence). Numéro affiché `NNN/122` (y compris 123–127,
comme sur les vraies cartes) et `P01`… pour les promos. Le texte de la pastille rétrécit
automatiquement pour les noms longs. Les scripts générateurs étaient dans le dossier temporaire
de la session (non conservés) ; le dossier ne contient que les sorties.

**Dossier épuré le 15/09/2026.** À la demande de l'utilisateur, tout ce qui ne sert pas au
fonctionnement a été retiré : `constructeur.html` (l'outil lui-même, voir §6 — `build_html.py`
ne l'écrit plus), `ma-collection-nomade.html` (sortie de cet outil, 19 Mo), `projet-collection/`
(dossier vide) et `serveur/` (le projet ASP.NET, code source, non régénérable — voir §7).
Si le besoin de synchronisation entre appareils revient, ce projet serveur est à refaire de
zéro, cette fois avec les chemins relatifs corrigés dès le départ.

**Point important sur la chaîne de génération.** `build4.py` écrit le classeur Excel.
`build_html.py` lit ce classeur et en injecte les données dans le HTML. Donc :

- pour modifier les **données** (cartes, raretés, promos) → éditer `build4.py`,
  puis relancer `build4.py` **et** `build_html.py`
- pour modifier l'**interface** → éditer le gabarit HTML dans `build_html.py`,
  puis relancer `build_html.py` seul

Les chemins étaient auparavant codés en dur vers `/home/claude/collection/...` (résidu d'une
session précédente dans un autre environnement). Corrigé le 14/09/2026 : `build_html.py`
utilise désormais le dossier du script (relatif), donc `python build_html.py` tourne
directement depuis ce dossier sur le PC Windows de l'utilisateur.

Après `build4.py`, vérifier que le classeur s'ouvre sans erreur de formule (pas d'outil de
recalcul automatique installé sur ce PC — une relecture rapide dans Excel suffit).

Vérifier la syntaxe JS après chaque génération : extraire le contenu de `<script>` et
passer `node --check`. Le HTML fait ~1000 lignes de JS, une erreur casse tout en silence.

---

## 3. Les données des cartes

**133 entrées** au total :

- **001 à 122** — la cardlist. Numéro, nom, rareté, type et essence, relevés sur
  gentlemates.com/tcg-cards le 14/09/2026 (les caractéristiques étaient encodées dans les
  noms de fichiers des visuels). Fiables.
- **123 à 127** — les cartes secrètes, numérotées 123/122 à 127/122. Noms et numérotation
  confirmés pour les 5, y compris la 123 (« Le Quatrième du Nom »), par la checklist
  officielle du jeu (photo fournie par l'utilisateur le 14/09/2026 — voir plus bas). Ni
  type ni essence au départ : le site ne les publiait pas. **06/10/2026 : la 123 (Le
  Quatrième du Nom — Joueur, Orange) et la 125 (6W26L — Joueur, Mauve) sont dévoilées** :
  visuels officiels récupérés dans `images-cartes/` (ils remplacent les images floues de
  teasing), type et essence renseignés via `SECRETES_REVELEES` dans `build4.py`, dos de cartes
  `dos-cartes/123.png` et `125.png` mis à jour avec leur symbole d'essence. Restent cachées :
  124, 126, 127 — pour les dévoiler, ajouter leurs entrées dans `SECRETES_REVELEES`, remplacer
  leur image dans `images-cartes/`, relancer `build4.py` puis `build_html.py`. Les visuels
  viennent de la page gentlemates.com/tcg-cards (images servies par Supabase, le nom de
  fichier encode `numéro_rareté_type_essence_nom`).
- **P01 à P06** — les promos, hors boosters. P01 et P02 sont les deux cartes vendues au
  Major de Paris, P05 « Tout Commence Ici » et P06 « Le Marionnettiste » viennent des
  packs. P03 (Squeezie le Toxique) et P04 (Brawks Déter) confirmées par la même checklist
  officielle — ce n'était jusque-là qu'une déduction.

**Checklist officielle (14/09/2026).** L'utilisateur a fourni une photo de la checklist
papier officielle du jeu (numéro, nom, rareté, une ou deux cases à cocher par carte). Elle
a confirmé la numérotation 123-127 et P03/P04 ci-dessus, et révélé un point qui n'était pas
documenté ailleurs : **une version reverse n'existe que pour les raretés Essence, Commune,
Peu commune et Rare** — les cartes Épique, Légendaire, Secrète et Promo n'ont qu'une seule
case sur la checklist, donc pas de reverse. Aucune exception constatée sur les 133 cartes.
Cette règle est maintenant codée dans `build_html.py` (`NOREV_RARETE`, fonction `autoRev`)
et remplace la question manuelle « Une version reverse existe ? » posée avant pour chaque
carte : elle est déduite automatiquement de la rareté, avec une option pour la corriger
carte par carte si jamais la règle se révèle fausse dans un cas précis (menu déroulant du
panneau de détail, toujours présent).

Répartition : 33 communes, 24 peu communes, 28 rares, 18 épiques, 10 légendaires,
5 secrètes, 6 promos, 9 essences.

Les 10 essences : Rose, Orange, Jade, Violette, Jaune, Mauve, Verte, Bleue, Azur, Neutre.

Trois noms apparaissent plusieurs fois dans la cardlist (CHAMPIONS DU MAJOR ×3,
Ravitaillement ×2, Leakof ×2) : ce sont des cartes distinctes, signalées en note.

**Ne pas inventer de données manquantes.** Tout ce qui est déduit plutôt que constaté est
signalé en note dans le classeur, et doit le rester.

**Plusieurs extensions à l'avenir (décidé le 14/09/2026, en attente).** Le champ `ext`
existe déjà par carte (actuellement toujours "Tout Commence Ici") mais le numéro seul
(`001`, `P03`…) sert encore de clé unique partout : stockage localStorage, nom de fichier
image, comparaison entre amis. Une 2ᵉ extension qui recommence sa numérotation à 001
entrerait en collision. Décision : **ne rien changer tant qu'il n'y a pas de vraies
données de 2ᵉ extension** — plutôt que de deviner un format à l'avance. Le jour venu,
prévoir : une clé composée du genre `TCI-001` / `EXT2-001` (migration des données
existantes à faire sans perte), des visuels rangés par sous-dossier
(`images-cartes/TCI/001.png`), et un filtre « Extension » dans la barre de filtres.

---

## 4. L'application — état actuel

Un fichier HTML autonome, français, thème sombre. Données dans le localStorage du
navigateur, export et import JSON pour passer d'un appareil à l'autre.

**Trois onglets** : Collection (par défaut), Statistiques, Comparer (voir plus bas). Les
statistiques ont été déplacées dans un onglet à la demande de l'utilisateur, qui ne veut
pas les voir à chaque ouverture.

**En-tête** : logo à gauche (76 px, 58 px sur mobile — quasiment toute la hauteur du
bandeau), compteur à droite, bande de progression carte par carte colorée par essence
dessous. Reste toujours à cette taille, ne grossit plus au chargement et ne change plus
au scroll — une version précédente agrandissait le logo puis le rétrécissait au défilement,
ça glitchait avec le flou de l'en-tête (`backdrop-filter`) et l'utilisateur a demandé à le
garder petit en permanence.

**Barre d'outils** : recherche, tri, bascule galerie/tableau, taille des vignettes
(340/230/150 px, mémorisée), bouton Filtres, bouton Ajouter une carte.

**Filtres** repliés derrière un bouton, groupés sous trois titres (Statut, Rareté,
Essence), avec une pastille indiquant le nombre de filtres actifs. Ils étaient auparavant
étalés en trois rangées permanentes : l'utilisateur a jugé ça inutilisable.

**Vignette** : visuel (grisé si la carte n'est pas possédée), badge de statut, icône
d'essence en haut à droite du bloc info (58 px — l'utilisateur les voulait très grandes),
numéro, nom, rareté, puis la saisie.

**Saisie — modèle important.** Une case à cocher marque l'exemplaire rangé dans le
classeur. Une fois cochée, une ligne « Doublons » apparaît avec des boutons − et +. En
interne, le total est stocké dans `qn` (normale) et `qr` (reverse) : cocher met à 1, les
doublons ajoutent au-delà. La ligne Reverse n'apparaît que si une version reverse existe
pour cette carte — déduit automatiquement de la rareté (voir §3, checklist officielle),
modifiable au cas par cas dans le panneau de détail. Pas de clic droit : tout doit rester
utilisable au doigt.

**Statut calculé** : Complète (vert), Manque reverse (orange), Manquante (rouge).

**Panneau de détail** au clic : mêmes cases et doublons, existence d'une reverse, état,
prix payé, valeur estimée, date, emplacement, notes. Les cartes ajoutées par l'utilisateur
peuvent y être supprimées.

**Passe « profondeur » (14/09/2026).** L'utilisateur trouvait le design trop plat / « appli
basique », surtout l'en-tête et les couleurs. Ajouté sans changer le thème violet sombre :
ombres portées sur cartes/panneaux/boutons (relief au survol), dégradés doux sur les
surfaces (stats, cartes, boutons) au lieu d'aplats, halo derrière le logo et liseré dégradé
rose→violet→azur sous l'en-tête, lueur sur les badges de statut et les barres de
progression. Tout dans le `<style>` de `build_html.py` — rien touché dans le HTML généré
directement.

**Mode rapide (14/09/2026).** Bouton dans la barre d'outils : une fois activé, un tap sur
la vignette (au lieu d'ouvrir le panneau de détail) coche directement « Normale ». Pensé
pour cocher plusieurs cartes de suite juste après avoir ouvert des boosters. Les cases
fines (reverse, doublons) restent cliquables normalement, indépendamment du mode.

**Copier les manquantes / Copier les doublons (14/09/2026).** Deux boutons en pied de
page qui mettent dans le presse-papier une liste texte (numéro + nom) — les cartes qui
manquent, ou celles possédées en plusieurs exemplaires — pour la coller ailleurs
(Discord, SMS…).

**Onglet Comparer (14/09/2026).** Permet de charger le fichier JSON exporté par un ou
plusieurs amis (bouton « Importer des amis », sélection multiple) sans toucher à sa
propre collection — chargé en mémoire seulement, pas persistant après rechargement de la
page. Affiche un tableau croisé (une colonne par personne, un point de statut + le nombre
de doublons par carte) et, pour chaque ami, un bouton « Résumé d'échange » qui copie dans
le presse-papier ce qu'on peut lui donner (nos doublons qu'il n'a pas) et lui demander
(ses doublons qu'on n'a pas). Aucune synchronisation en temps réel : chacun exporte son
fichier et se l'échange à la main, cohérent avec le reste (pas de serveur nécessaire).

**Page d'accueil : choix de l'extension (15/09/2026).** À l'ouverture, un écran affiche
une tuile par extension présente dans `CARDS` (déduites automatiquement du champ `ext` —
donc une nouvelle extension apparaît toute seule le jour où ses cartes sont ajoutées,
sans toucher au code), avec la progression de chacune. Cliquer dessus entre dans la
collection filtrée à cette extension ; un bouton « ← Extensions » dans l'en-tête permet
de revenir choisir. Le dernier choix est mémorisé (`localStorage`) pour ne pas repasser
par cet écran à chaque ouverture. En marge de ce chantier, correction d'un import : les
6 promos avaient `ext:"Promo"` au lieu de `"Tout Commence Ici"` dans `build4.py` (elles
faisaient donc croire à une extension séparée) — corrigé, et le même défaut par défaut
dans le formulaire « Ajouter une carte » (`build_html.py`).

**Onglet Administration (15/09/2026).** Un quatrième onglet, à côté de Comparer, pour
gérer une extension sans passer par moi ni par le pipeline Python :

- **Créer une extension** — bouton sur la page d'accueil (« + Créer une extension »).
  Le nom est stocké dans `localStorage` (`EXT_EXTRA`) même sans aucune carte dedans, pour
  que la tuile existe tout de suite sur l'accueil.
- **Ajouter des cartes** — même formulaire que l'ancien bouton « Ajouter une carte »,
  mais pré-rempli avec l'extension en cours, et avec un bouton « Ajouter et continuer »
  pour en saisir plusieurs à la suite sans rouvrir le formulaire à chaque fois.
- **Modifier / supprimer n'importe quelle carte**, y compris celles du set officiel — pas
  seulement celles ajoutées à la main comme avant. Techniquement : un système de
  surcharges (`OVERRIDES`, par numéro de carte) et une liste de numéros masqués
  (`DELETED`), tous les deux en `localStorage`. Le numéro d'une carte officielle ne peut
  pas être changé (ça casserait le lien avec son visuel dans `images-cartes/`), le reste
  oui. `CARDS` est maintenant reconstruit par une fonction unique, `rebuildCards()`, à
  chaque changement de `EXTRA`/`OVERRIDES`/`DELETED`.

Tout ça reste local à l'appareil (comme le reste des saisies) : ça ne modifie ni le
classeur Excel ni le fichier HTML généré. Si une extension créée ainsi doit devenir
« officielle » (avec ses propres visuels, dans le classeur), il faudra la reporter à la
main dans `build4.py` le moment venu.

**Info et logo d'extension (15/09/2026).** Le bouton « Créer une extension » ouvre
maintenant un vrai formulaire (nom, info en texte libre, logo) plutôt qu'une simple boîte
de saisie. Le logo est choisi comme un fichier depuis l'appareil, réduit dans le
navigateur (canvas, 700 px max) et stocké en base64 dans `localStorage` (`EXT_META`) —
pas besoin de déposer un fichier dans `logo/`. Un bouton « Modifier » dans l'onglet
Administration permet de changer l'info et le logo après coup (le nom, lui, ne change
plus une fois l'extension créée). `extLogo()` regarde d'abord `EXT_META`, et ne retombe
sur la convention `logo/<slug>.png` que si rien n'a été mis en ligne.

**Bug corrigé au passage.** Le panneau de détail (`#panel`) et son fond (`#scrim`)
s'étaient retrouvés à l'intérieur du conteneur `#app`, cliquant sur « Créer une
extension » n'affichait donc rien tant qu'aucune extension n'était encore choisie
(`#app` est masqué sur l'écran d'accueil). Déplacés en dehors de `#app` dans
`build_html.py`.

**Administration cachée par défaut (15/09/2026).** L'onglet Administration et le bouton
« + Créer une extension » sont masqués (`hidden`) tant que le mode admin n'est pas
déverrouillé. Pour le déverrouiller : **taper 5 fois de suite sur le logo** (celui de
l'accueil ou celui de l'en-tête, les deux marchent) en moins de 2 secondes — un toast
confirme, et l'état est retenu sur cet appareil (`localStorage`, clé `m8tcg.admin`). Les
mêmes 5 tapotements le reverrouillent. Important : comme tout le reste dans cette appli,
c'est du **local, pas de la vraie sécurité** — n'importe qui inspectant le code source
verrait comment ça marche. Ça sert à ne pas exposer ces boutons par accident à un ami qui
utilise juste « Comparer », pas à protéger des données sensibles.

**Export/import complets (15/09/2026).** Bug trouvé en cherchant la prochaine
amélioration : « Exporter mes données » ne sauvegardait que `store` et `EXTRA` — tout le
travail d'Administration (extensions créées, cartes modifiées ou supprimées) était donc
perdu au réimport sur un autre appareil. Corrigé : le JSON exporté (désormais `v:3`)
inclut aussi `overrides`, `deleted`, `extsExtra` et `extMeta`, et l'import les restaure
tous. Testé par un aller-retour complet (export puis import sur un état vidé).

**Suppression d'une extension (15/09/2026).** Bouton « Supprimer « … » » dans l'onglet
Administration (absent sur « Tout Commence Ici », qui ne peut pas être supprimée — c'est
l'extension d'origine, intégrée aux données, pas une extension créée à la main). Supprime
l'extension elle-même (`EXT_EXTRA`, `EXT_META`), les cartes que tu y avais ajoutées à la
main (`EXTRA`, avec leurs saisies dans `store`), et réaffecte à leur extension d'origine
les cartes officielles que tu y avais éventuellement déplacées via une surcharge — sans
perdre le reste de ce que tu avais modifié dessus. Si l'extension supprimée était celle
affichée à l'écran, retour automatique à la page d'accueil.

---

## 5. Habillage (dossier `logo/`)

Chargé au démarrage, chaque élément facultatif, la page reste correcte sans :

- `logo.png` (ou .svg, .webp) — remplace le titre texte
- `banniere.jpg` — fond de l'en-tête, assombri par un dégradé
- `icone.png` — 512 px, icône d'écran d'accueil iOS
- `dos-carte.png` — visuel de repli pour les cartes sans image
- `essence-rose.png`, `essence-azur.png`, etc. — une par essence, minuscules sans accent

---

## 6. La question non résolue : ouvrir la page sur iPhone

**Le problème.** Un dossier HTML + images fonctionne sur PC. Sur iOS, ouvrir le HTML
depuis l'app Fichiers ne charge pas les fichiers voisins.

**Tentative 1** — `constructeur.html` : un outil local qui lisait les dossiers
`images-cartes` et `logo`, compressait les visuels et produisait un fichier unique avec
tout en base64. **Testé : échec sur iPhone.** La visionneuse d'iOS affiche la page mais
n'exécute pas le JavaScript, donc aucune carte ne s'affiche. Le fichier fonctionnait
parfaitement sur PC. Retiré du projet le 15/09/2026 (dossier épuré, voir §2) — piste
abandonnée, inutile de la retenter telle quelle.

**Conclusion** : il faut une vraie adresse web pour que Safari exécute les scripts. Pistes
évoquées, aucune retenue à ce jour :

- Netlify Drop — déposer le fichier unique, obtenir une URL, ajouter à l'écran d'accueil.
  La plus simple. C'était la dernière proposition sur la table.
- Serveur ASP.NET sur son PC (§7) — écarté comme trop compliqué.
- Tailscale ou Cloudflare Tunnel — écartés pour la même raison.

L'utilisateur a un port RDP ouvert sur sa box et ne veut pas en ouvrir d'autre.

---

## 7. La synchronisation serveur (abandonnée)

Il existait un projet ASP.NET Core 8 complet (dossier `serveur/`), jamais déployé : il
servait les fichiers statiques et exposait `GET`/`PUT /api/collection` pour stocker la
collection en JSON côté serveur, avec sauvegardes horodatées et authentification Basic.
**Supprimé le 15/09/2026** dans le nettoyage du dossier (code source, non régénérable —
voir §2). Si le besoin de synchronisation entre appareils redevient prioritaire, ce projet
est à refaire de zéro.

**Appli installable, SANS mode hors connexion (27/10/2026).** Demande explicite de l'utilisateur :
icône + plein écran oui, hors connexion non. Fichiers statiques à la racine (non générés par
`build_html.py`) : `manifest.webmanifest` (nom, icônes, plein écran, couleurs), `sw.js`, dossier
`icons/` (icônes dessinées depuis `logo/logo.png` sur le fond de l'appli, dont une « maskable » et
`apple-touch-icon.png`). `sw.js` est un service worker **volontairement minimal** : gestionnaire
`fetch` vide, aucun cache, il purge les caches `m8-*` d'une éventuelle version précédente ; il ne
sert qu'à satisfaire les navigateurs Android qui exigent un service worker pour proposer
« Installer l'application ». `build_html.py` ajoute les balises dans `<head>` (manifest,
apple-touch-icon, favicon — `logo/icone.png` n'a jamais existé), enregistre le service worker
(seulement en http/https, donc pas en `file://`) et ajoute la section « Appli sur ton téléphone »
dans la roue crantée (bouton Installer si le navigateur le propose, sinon mode d'emploi
iPhone/Android). Une première version avec cache des images (mode hors connexion testé et
fonctionnel, 218 Mo) a été écrite puis retirée : si l'envie revient, c'est un cache d'abord pour
`images-cartes/` et `logo/` + réseau d'abord pour la page, avec une constante de version de cache
à changer quand on remplace des images. Pas testé sur un vrai téléphone.

**Reverse sans normale (27/10/2026).** On peut posséder la reverse d'une carte sans en avoir la
normale. `statutOf` la classait à tort « complète » ; elle est maintenant « partielle » (pastille
orange) tant que les deux versions n'y sont pas, pour les cartes qui ont une reverse. Nouvelle
fonction `versionState(data,c)` (Complète / Sans reverse / Sans normale / Absente) utilisée par la
colonne « Statut » du tableau de la collection (plus de symboles ●◐○) et par Comparer. Le filtre
de statut « Reverse manquante » s'appelle « Une version manque ». Le Full set compte une carte dès
qu'une version est possédée (reverse seule incluse) ; le Master set compte chaque version à part.
Le classeur Excel (`build4.py`) était déjà correct.

**Écran Comparer refait (27/10/2026).** Un ami à la fois (puces en haut, `cmpFriend`) au lieu
d'un tableau avec tous les amis et des symboles ●◐○. Contenu : 3 chiffres (cartes de l'ami, en
commun, échanges possibles donner ⇄ recevoir), deux listes « Tu peux donner à X » / « X peut te
donner » (8 lignes + « Voir N autres », bouton Copier), puis le tableau de toutes les cartes en
deux colonnes (toi / l'ami) avec étiquettes en mots (Complète, Sans reverse, Sans normale, Absente
+ « N en double ») et trois filtres (toutes / chez l'ami pas chez toi / chez toi pas chez l'ami).
Règle d'échange (`tradeList`) : on propose un **doublon de version** (normale en double si l'autre
n'a aucune normale ; reverse en double si l'autre n'a aucune reverse et que la reverse existe),
plus fine que l'ancienne règle « carte entière manquante ». Noms d'amis et de cartes échappés
(`cmpEsc`). Le conteneur `.cmpwrap` a `grid-template-columns:minmax(0,1fr)` : sans ça le tableau
élargit toute la page sur mobile. Jeu de test : dossier `jeu-de-test/`.

**Écran Statistiques refait (27/10/2026).** Fonction `stats()` : deux anneaux SVG (Full set =
cartes différentes, Master set = normales + reverses, 227 au total), trois chiffres (cartes
ouvertes, **reverses trouvées sur le total** (ex. 3 / 94 : cartes ayant une reverse cochée sur
toutes les cartes qui en ont une), **doublons** = somme de `dupOf`), barres « Par rareté » (couleurs
`RCOLOR`) et « Par essence » (couleurs d'essence), et un bloc « Argent » (investi / valeur / écart)
**affiché seulement si au moins un prix ou une valeur est saisi**. La barre « reverses par
rareté » a été supprimée (le Master set la résume). Attention : `setTab("stats")` ne recalcule
pas, `stats()` est appelée par `render()` — après un changement de données hors interface, appeler
`render()`.

**Roue crantée « Réglages » (27/10/2026).** Le pied de page a disparu : tout est dans un panneau
ouvert par la roue (`.gear`, `#gearHome` sur l'accueil, `#gearApp` dans l'en-tête, fonction
`openSettings()`) : synchronisation, export JSON/CSV, import, copier manquantes/doublons, « Tout
effacer », et — seulement si le mode administration est déverrouillé (5 taps secrets, inchangé) —
gérer les cartes de l'extension, créer une extension, quitter l'administration. Les anciens
handlers de boutons sont devenus des fonctions (`exportJson`, `exportCsv`, `copyMissing`,
`copyDup`, `importPick`, `resetAll`) ; `#file` (import) reste dans le corps de la page. L'onglet
« Administration » n'apparaît plus que lorsqu'on est dedans. `#newExtBtn` et `<footer>` supprimés.
Un point rose sur la roue signale un souci de synchro (jeton refusé, choix de version à faire).

**Barre d'outils refaite (27/10/2026).** Onglets Collection / Statistiques / Comparer en
navigation soulignée (`.tabnav`) ; rangée 1 = recherche en pilule + bouton Filtres ; rangée 2 =
tri, vue galerie/tableau en icônes SVG intégrées (aucune dépendance externe), curseur de taille
`#sizeRange` (150–340, remplace les 3 boutons Grandes/Moyennes/Petites ; même clé `m8tcg.ui`),
« Mode rapide » et « Ajouter » (seul bouton rose, classe `.btn.primary`). « Importer des amis »
n'est plus à côté des onglets : il est dans l'onglet Comparer (`#addFriends` recréé dans
`renderCompare`). `.tools` est maintenant une grille de deux `.toolrow` (`display:grid` dans
`setTab`) ; sur mobile la 2e rangée (`.second`) passe en grille 2 colonnes.

**Mise à jour 27/10/2026 — la synchronisation existe, autrement.** L'appli est publiée via
GitHub Pages (voir `README.md`) et synchronise la collection par un **gist secret** du compte
GitHub de l'utilisateur (roue crantée en haut à droite → « Synchroniser mes appareils »). Aucun serveur. Code :
bloc « synchronisation » de `build_html.py` (remplace l'ancien `pull()/push()` vers
`api/collection`, retiré).

- Jeton GitHub « classic » avec la seule portée `gist`, collé une fois par appareil, stocké en
  `localStorage` (`m8tcg.sync` = `{token, id, cloudAt, last}`). Fichier en ligne :
  `m8tcg-collection.json`, même contenu que l'export JSON (v3).
- `m8tcg.syncDirty` = « changements locaux pas encore envoyés ». Chaque `save*()` appelle
  `changed()` → envoi différé de 1,5 s (lecture du gist puis PATCH).
- Détection de conflit : le fichier en ligne contient `updatedAt` ; s'il diffère de
  `SYNC.cloudAt` (dernière version vue) **et** que l'appareil a des changements non envoyés,
  l'appli ouvre un panneau « Quelle version garder ? » — jamais d'écrasement silencieux.
  Si seul le cloud a changé, il est appliqué sans question. Resynchro au retour sur l'onglet
  (`visibilitychange`) et à la reconnexion.
- Premier branchement d'un appareil : gist existant retrouvé par son nom de fichier ; appareil
  vide → récupération directe, appareil déjà rempli → même panneau de choix.
- Testé le 27/10/2026 avec un faux serveur GitHub dans le navigateur (création, envoi auto,
  deuxième appareil, jeton refusé, conflit dans les deux sens, désactivation). Confirmé
  fonctionnel par l'utilisateur avec le vrai GitHub le 27/10/2026.
- Publication : dépôt https://github.com/Rmartin86580/TCG-M8 (branche `main`), site sur
  https://rmartin86580.github.io/TCG-M8/ (GitHub Pages, racine). Le dépôt local est
  initialisé dans ce dossier (identité git locale `Rmartin86580`, e-mail noreply). Pour
  publier une modification : `python build_html.py` (régénère `index.html`), puis
  `git add -A`, `git commit`, `git push`.
- Historique : chaque envoi crée une révision du gist (consultable sur gist.github.com), ce
  qui sert de filet de sécurité si une version est écrasée par erreur.

`ma-collection.html` garde encore le code qui essaie d'interroger `api/collection` au
chargement (fonctions `pull()`/`push()`) : sans serveur en face, ça échoue silencieusement
et l'appli bascule en mode local, comme prévu — rien à corriger dans l'immédiat, mais ce
code est aujourd'hui mort tant qu'aucun serveur n'existe.

---

## 8. Ton de travail attendu

L'utilisateur va droit au but, écrit vite, et n'aime pas les détours. Il apprécie qu'on
lui propose plusieurs options quand un choix d'interface se présente, et qu'on lui dise
franchement quand une piste ne marche pas plutôt que d'insister. Réponses en français,
tutoiement.

Il fournit lui-même ses visuels et ses fichiers. Ne pas supposer qu'un dossier est en
place : vérifier.
