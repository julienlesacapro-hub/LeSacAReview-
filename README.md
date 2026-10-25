# Revues SIGAPS

Application web (installable sur téléphone et ordinateur) pour interroger la liste des revues SIGAPS : **rang (A+ → NC) et IF**, **titre**, **soumission gratuite / payante**, **bouton vers la revue**. Résultats triés par rang puis IF.

- Recherche FR / EN par domaine ou titre (ex. « trachée » → ORL, laryngologie…), case « Termes exacts », filtres (année, rang, IF, soumission, MEDLINE, JCR, domaine, thématique, case reports, vérification, colonnes ajoutées…).
- Les données sont **intégrées au dépôt** (`data.enc`) : aucune recherche de fichier, la requête est automatique.

## ⚠ Confidentialité

Le classeur indique que les Impact Factors (Clarivate) **ne peuvent être ni diffusés ni rendus publics**. Ce dépôt étant public, **le classeur Excel ne doit jamais y être déposé** (`.gitignore` bloque les `.xlsx`). Seul `data.enc`, **chiffré (AES‑256‑GCM)**, est publié ; l'appli le déchiffre dans votre navigateur avec un code d'accès que vous saisissez une fois par appareil. Choisissez une phrase d'au moins 12 caractères : le fichier chiffré étant public, un code faible pourrait être deviné. Plus sûr encore : passez le dépôt en privé (Settings → General → Danger Zone ; GitHub Pages sur dépôt privé exige un plan payant).

## Installation (une seule fois)

1. Fusionner la branche dans `main`, puis **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`**. L'adresse est `https://julienlesacapro-hub.github.io/LeSacAReview-/`.
2. Ouvrir l'adresse, cliquer sur l'icône ⬆ (« Mettre à jour les données ») et suivre la procédure ci‑dessous pour produire le premier `data.enc`.
3. Installer l'appli : iPhone/iPad → Safari › Partager › *Sur l'écran d'accueil* ; Android/Chrome → menu › *Installer l'application* ; ordinateur → icône d'installation dans la barre d'adresse.

## Mettre à jour le tableau Excel

1. Modifier le classeur comme d'habitude (nouvelles revues, corrections, nouvelle année : ajouter les colonnes `20XX` **IF** + **Cat.** à la suite, même format que 2021‑2025). Garder l'onglet **Revues**, ses deux lignes d'en‑têtes et les noms de colonnes.
2. Ouvrir l'appli → icône ⬆ → **Choisir le classeur Excel** → saisir le code d'accès (vide = garder le code actuel ; 12 caractères minimum pour un nouveau code) → **Générer data.enc**. La lecture se fait dans le navigateur, rien n'est envoyé. Le résumé indique les revues ajoutées/retirées, les années détectées et les anomalies (rang inconnu, doublon…).
3. **Télécharger data.enc**, puis sur GitHub : dépôt → **Add file → Upload files** → glisser `data.enc` à la racine (il remplace l'ancien) → **Commit changes** (sur `main`).
4. Après 1 à 2 minutes (déploiement Pages) l'appli affiche les nouvelles données sur tous les appareils (fermer/rouvrir l'appli si besoin). Le bouton « aperçu » permet de contrôler le résultat avant de publier.

Changer de code d'accès : générer `data.enc` avec le nouveau code ; chaque appareil le redemandera une fois.

## Fonctionnement des données

- Colonnes reconnues par leur **nom** (pas par leur lettre) ; une **colonne ajoutée** devient automatiquement un filtre et apparaît dans la fiche.
- Rang/IF affichés = dernière année disposant d'un rang (année en petit ; en orange si ce n'est pas la dernière année du fichier). Le filtre « Année de référence » force une année précise. Sans rang : NC.
- Soumission : « Gratuite », « Payante », « Gratuite · OA payant » (mixte) ou « À vérifier », déduit du texte de la colonne *Soumission / publication* (texte intégral au survol et dans la fiche).
- Bouton « Site de la revue » = colonne *Lien vers la revue* ; s'il est vide, bouton « Fiche NLM » (catalogue NLM, d'après le NLMid).
- Recherche approchée : racines de mots, fautes de frappe d'un caractère, et `lexique.js` (correspondances FR↔EN organe/maladie → domaines, rédigées à la main : à compléter au besoin). Les titres du classeur sont des abréviations NLM.

## Développement

`node tests/test.mjs chemin/classeur.xlsx` : lecture, chiffrement/déchiffrement et recherche sur le vrai fichier. Fichiers : `index.html`, `style.css`, `app.js`, `core.js`, `lexique.js`, `sw.js` (hors ligne), `vendor/` (SheetJS 0.18.5, licence Apache‑2.0).
