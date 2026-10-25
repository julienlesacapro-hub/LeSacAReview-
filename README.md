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

## Comptes utilisateurs (Supabase) — accès individuel par e‑mail

Par défaut l'appli utilise un code d'accès partagé (`data.enc`). Pour donner un **compte personnel** à d'autres personnes (e‑mail + mot de passe initial à changer + « mot de passe oublié » par e‑mail), l'appli s'appuie sur Supabase (gratuit à cette échelle). Les données ne sont alors **plus dans le dépôt** : elles sont dans un stockage privé, lisible uniquement par les comptes actifs.

> ⚠ Vérifier avant d'inviter quiconque : le classeur indique que les IF (Clarivate) « ne peuvent être ni diffusés ni rendus publics » hors du cadre SIGAPS. À confirmer avec votre DRCI / établissement.

### Mise en place (une seule fois)

1. Créer un compte sur supabase.com → **New project** (région UE, noter le mot de passe de la base).
2. **Authentication → Sign In / Providers → Email** : désactiver **Allow new users to sign up** (indispensable : seuls les comptes que vous créez peuvent entrer) ; activer *Confirm email* si proposé ; régler la longueur minimale du mot de passe à 12.
3. **Authentication → URL Configuration** : *Site URL* = `https://julienlesacapro-hub.github.io/LeSacAReview-/` et la même adresse dans *Redirect URLs* (sinon le lien « mot de passe oublié » ne reviendra pas sur l'appli).
4. **SQL Editor → New query** : coller le contenu de `supabase/setup.sql` → **Run**.
5. **Authentication → Users → Add user → Create new user** : créer **votre** compte (e‑mail + mot de passe, *Auto Confirm User* coché). Puis exécuter dans le SQL Editor (en remplaçant l'adresse) la requête « compte administrateur » en commentaire en bas de `supabase/setup.sql`.
6. **Project Settings → API** : copier *Project URL* et la clé publique *anon / publishable* dans `config.js` (jamais la clé `service_role`). Commit sur `main`. L'appli passe automatiquement en mode comptes.
7. Ouvrir l'appli, se connecter avec votre compte, icône ⬆ → choisir le classeur Excel → **Générer et publier**.
8. Supprimer `data.enc` du dépôt (devenu inutile ; c'est une copie chiffrée publique).

### Créer un compte pour une autre personne

**Authentication → Users → Add user → Create new user** : e‑mail + mot de passe initial (cocher *Auto Confirm User*), puis lui transmettre ces identifiants. À sa première connexion, l'appli lui impose de choisir un mot de passe personnel (12 caractères min.) ; **tant qu'il n'est pas changé, le serveur refuse l'accès aux données** (règle SQL, pas seulement l'écran).
Retirer l'accès : même page → menu du compte → *Delete user* (ou *Ban*). Forcer à nouveau un changement de mot de passe : requête en commentaire dans `supabase/setup.sql`.

### Mot de passe oublié

Écran de connexion → **Mot de passe oublié ?** → saisir l'e‑mail → l'utilisateur reçoit un lien, choisit un nouveau mot de passe, et est connecté. Le service d'e‑mail intégré de Supabase est limité (quelques e‑mails par heure) et son expéditeur est générique : pour un usage régulier, configurer un SMTP personnalisé (Authentication → SMTP Settings). Un lien expiré ou déjà utilisé (certaines messageries le « consomment » en prévisualisant) se remplace en refaisant la demande.

### Mettre à jour les données (administrateur)

Se connecter avec le compte administrateur → icône ⬆ → classeur Excel → **Générer et publier** : la mise à jour est immédiate pour tous, sans passer par GitHub.

### Limites

- Mot de passe oublié, comptes et mots de passe sont gérés par Supabase ; le code de l'appli ne les voit jamais.
- Les utilisateurs gardent sur leur appareil une copie du dernier fichier reçu (pour le mode hors ligne) : supprimer un compte n'efface pas cette copie.
- Le comportement face à un vrai projet Supabase n'a été testé que contre un serveur simulé (connexion, première connexion, oubli, publication) ; la règle SQL et l'envoi réel d'e‑mail sont à vérifier à la mise en place (créer un compte de test).

## Mettre à jour le tableau Excel (mode code d'accès partagé)

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

`node tests/test.mjs chemin/classeur.xlsx` : lecture, chiffrement/déchiffrement et recherche sur le vrai fichier. Fichiers : `index.html`, `style.css`, `app.js`, `core.js`, `lexique.js`, `config.js` + `supabase/setup.sql` (comptes), `sw.js` (hors ligne), `vendor/` (SheetJS 0.18.5 et supabase-js 2.116.0, licences incluses).
