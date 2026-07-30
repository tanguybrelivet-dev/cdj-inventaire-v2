# Inventaire Terrain — mettre le prototype sur votre téléphone

> **Mise à jour** : l'export « vers le modèle Word » génère désormais un
> vrai fichier `.docx` (au lieu d'un HTML déguisé qui ne s'ouvrait pas dans
> l'app Word sur iPhone). Sur mobile, il utilise le partage natif iOS
> (Partager → Enregistrer dans Fichiers / Ouvrir dans Word) plutôt qu'un
> téléchargement classique. Pensez à remplacer les fichiers de votre dépôt
> GitHub par ceux de ce zip pour que Netlify redéploie la correction.

Ce dossier est un vrai petit projet web (React + Vite). Une fois déployé en
ligne, vous pourrez l'ouvrir sur votre téléphone et l'**ajouter à l'écran
d'accueil** : il se comporte alors comme une app (icône, plein écran, pas de
barre d'adresse), sans passer par l'App Store / Google Play.

## Option A — la plus simple (aucune ligne de commande), ~5 minutes

1. Allez sur https://vercel.com (ou https://app.netlify.com/drop) et créez un
   compte gratuit si besoin.
2. **Netlify Drop** : faites simplement glisser tout ce dossier
   `inventaire-app` sur https://app.netlify.com/drop → un lien public est
   généré immédiatement (ex. `https://xxxx.netlify.app`).
   *(Netlify Drop build automatiquement un projet Vite/React, aucune
   configuration nécessaire.)*
3. Ouvrez ce lien **depuis votre téléphone** (Safari sur iPhone, Chrome sur
   Android).
4. Sur iPhone : bouton **Partager** → **Sur l'écran d'accueil**.
   Sur Android : menu **⋮** → **Ajouter à l'écran d'accueil** / **Installer
   l'application**.
5. Une icône bleu marine avec le badge orange « HJ » apparaît sur votre
   écran d'accueil — cliquez dessus, l'app s'ouvre en plein écran.

## Option B — via GitHub + Vercel (si vous voulez un lien stable et des mises
à jour faciles)

1. Créez un dépôt sur https://github.com et déposez-y ce dossier.
2. Sur https://vercel.com, cliquez **Add New → Project**, importez ce
   dépôt GitHub.
3. Vercel détecte automatiquement Vite ; laissez les réglages par défaut et
   cliquez **Deploy**.
4. Vous obtenez un lien du type `https://inventaire-terrain.vercel.app` —
   ouvrez-le sur votre téléphone et suivez l'étape 4 de l'option A.

## Option C — tester en local avant de déployer (si vous êtes à l'aise avec
le terminal)

```bash
npm install
npm run dev
```

Puis ouvrez l'adresse affichée (`http://localhost:5173`) sur un ordinateur,
ou `npm run build && npm run preview` pour tester la version de production.

## Ce que ça donne concrètement une fois installé

- Icône sur l'écran d'accueil, ouverture en plein écran (pas de barre Safari/
  Chrome visible)
- Fonctionne sur iPhone et Android, sans validation App Store
- Appareil photo, micro (dictée) et stockage fonctionnent comme dans un vrai
  navigateur mobile — **testez bien ces trois points une fois installé**,
  ils étaient bridés dans l'aperçu Claude par le bac à sable

## Tableau de bord multi-appareils (nouveau)

L'app peut maintenant enregistrer les dossiers dans le cloud (Supabase,
gratuit) pour les retrouver depuis n'importe quel appareil — téléphone
terrain **et** ordinateur du bureau.

### 1. Créer le projet Supabase (5 minutes)

1. Allez sur https://supabase.com, créez un compte gratuit
2. **New project** → donnez-lui un nom (ex. `cdj-inventaire`), choisissez un
   mot de passe de base de données (à conserver, pas besoin de le retenir
   ensuite), région **Europe (Paris ou Frankfort)**, cliquez **Create**
3. Une fois le projet prêt, allez dans **SQL Editor** (menu de gauche) →
   **New query**, collez le contenu du fichier `supabase_setup.sql` fourni
   dans ce dossier, cliquez **Run**
4. Allez dans **Project Settings → API** : copiez la **Project URL** et la
   clé **anon public**

### 2. Connecter l'app à Supabase

Ouvrez `src/supabaseClient.js` et remplacez les deux valeurs :

```js
const SUPABASE_URL = "https://VOTRE-PROJET.supabase.co";
const SUPABASE_ANON_KEY = "VOTRE-CLE-ANON-PUBLIC";
```

Uploadez ce fichier modifié sur GitHub (avec le reste si besoin) — Netlify
redéploiera automatiquement. Un bouton **« Enregistrer le dossier »** (vert)
apparaît alors dans l'onglet Inventaire, et le **Tableau de bord** (accessible
depuis le menu d'accueil) affiche tous les dossiers enregistrés.

### 3. Intégrer le tableau de bord dans huissier-concarneau.fr

Comme c'est un site différent (WordPress), la façon la plus simple est
d'intégrer le tableau de bord **dans une iframe**, sans toucher au code du
site :

1. Trouvez l'adresse de votre app Netlify (ex.
   `https://fluffy-starship-675d0f.netlify.app` ou votre domaine personnalisé)
2. Dans WordPress, éditez la page où vous voulez afficher le tableau de bord
3. Ajoutez un bloc **HTML personnalisé** (ou un widget "Texte/HTML" selon
   votre thème) avec ce code, en remplaçant l'URL par la vôtre :

```html
<iframe
  src="https://VOTRE-SITE.netlify.app/?view=dashboard"
  style="width:100%; height:800px; border:0; border-radius:12px;"
  title="Tableau de bord des inventaires"
></iframe>
```

Le paramètre `?view=dashboard` fait que l'app s'ouvre **directement** sur le
tableau de bord (sans passer par le menu), et masque le bouton "Retour"
puisqu'elle est intégrée dans une autre page.

### Limites actuelles à connaître

- Les **photos** sont pour l'instant stockées uniquement sur l'appareil qui
  les a prises (pas encore synchronisées dans le cloud) — l'enregistrement
  cloud sauvegarde le texte des lots (zone, catégorie, désignation, valeur),
  pas encore les images. À ajouter si besoin (nécessite Supabase Storage).
- Pas de compte utilisateur : toute personne ayant le lien de l'app peut lire
  et modifier tous les dossiers (politique Supabase volontairement ouverte
  pour rester simple à ce stade — voir le commentaire dans
  `supabase_setup.sql` pour la restreindre plus tard).



## Pour aller plus loin (production réelle)

Ce prototype a maintenant une vraie base de données (Supabase) pour les
dossiers/lots, mais il reste encore quelques briques pour une mise en
production complète :
- synchroniser aussi les **photos** dans le cloud (Supabase Storage)
- des **comptes utilisateurs** (au lieu d'un accès ouvert à toute personne
  ayant le lien)
- les vraies API Sirene / Pappers via un petit serveur relais (clé API non
  exposée dans l'app)
- éventuellement un vrai wrapper natif (Capacitor) pour publier sur les
  stores et fiabiliser caméra/micro/mode hors-ligne

Dites-le si vous voulez qu'on avance sur l'une de ces briques.
