# Journal de développement

## 2026-09-16 — Rôles, personnel et groupes

- Reprise de la première maquette de rôles afin de la compléter côté serveur.
- Ajout de l'API versionnée `backend/main.py` : sessions signées, migrations MariaDB additives, gestion des groupes, responsable direct, type salarié/stagiaire et contrôle des ressources par rôle.
- Ajout des pages `Administration` et `Personnel` : un administrateur peut créer des groupes et attribuer les personnes; un responsable ne voit que son périmètre autorisé.
- La configuration des fiches reste conservée dans `Presence`; aucune migration ne supprime les comptes ou les fiches existantes.
- Suppression des artefacts Python et traduction de la documentation fonctionnelle en français.
- Ajustement du modèle : l'appartenance à un Groupe devient facultative pour un Employé; un Admin peut être choisi comme responsable direct.
- Ajout d'une structure hiérarchique Graphviz dans Administration; les Groupes sont conservés comme informations de nœud afin de garder un arbre lisible.
- Remplacement du rôle unique par des étiquettes combinables (`Admin`, `Responsable`, `Employe`); une personne Admin peut donc aussi être rattachée comme Employé à un responsable et à un Groupe.

### Vérifications locales

- Compilation des modules Python modifiés : OK.
- `python -m unittest test_access_control.py` : 6 tests réussis.

## 2026-09-14 — Réinitialisation de mot de passe

- Ajout d'une vue de réinitialisation dédiée : un lien contenant le paramètre `token` affiche les champs de nouveau mot de passe et de confirmation au lieu de revenir au formulaire de connexion.
- Validation locale de la confirmation et de la longueur minimale, puis appel de l'API `POST /reset-password` avec gestion explicite des liens invalides ou expirés.
- Après succès, le token est retiré de l'URL avant le retour à la connexion afin de ne pas pouvoir être réutilisé dans le navigateur.
- Migration des identifiants de l'API vers un fichier d'environnement protégé sur le VPS, chargé par `presence-app.service`. Les données MariaDB n'ont pas été modifiées.

## 2026-09-11 — Publication du calendrier de présence

- Publication du calendrier mensuel visuel et du mode de modification groupée.
- Correction de la compatibilité de connexion pour les comptes historiques dont `role` est vide mais `is_admin` est défini.
- Alignement de l'environnement de production sur Streamlit `1.60.0` et correction de la politique CSP nécessaire au composant de calendrier.

## 2026-09-10 — Calendrier mensuel visuel

- Ajout d'un calendrier par semaines avec deux demi-journées par date.
- Ajout des statuts Travail, Congé payé, Absence, Arrêt maladie, Férié et Autre.
- Ajout du mode `Modifier`, de la sélection multiple par glissement, de cartes à hauteur constante et de barres continues pour les statuts adjacents.
- Synchronisation des exceptions du calendrier avec les exports Excel afin que les congés et absences modifiés soient présents dans les fiches générées.
