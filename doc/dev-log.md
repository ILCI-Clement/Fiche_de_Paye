# Journal de développement

## 2026-09-08 — Cadrage et lancement de la refonte

### Décisions fonctionnelles enregistrées

- Le produit reste un outil interne de gestion des présences et des indemnités de stage.
- Les responsables travaillent sur leurs propres données. Les administrateurs conservent une vue globale. Les employés gardent le mode de connexion actuel et gèrent uniquement leur profil.
- Une bibliothèque de personnes, une file d’export, un historique, une corbeille et une page d’accueil centrée sur la création de fiches seront ajoutés.
- Les fiches utilisent le planning par défaut de chaque personne et un calendrier mensuel séparant matin et après-midi pour les exceptions.
- Les exports Word, PDF et Excel sont demandés pour salariés et stagiaires, individuellement ou par lot.
- L’archivage, la restauration depuis la corbeille et les règles de conservation de 60 jours sont définis dans `dev-requirements.md`.

### Travaux lancés

- Harmonisation de la langue des fichiers projet : français pour l’interface et la documentation, anglais pour les identifiants de code.
- Préparation d’un modèle de données indépendant de l’ancienne structure JSON imbriquée.
- Refonte progressive de l’interface de gestion des fiches, sans modifier les modèles Excel et Word existants.

### Première livraison technique

- Ajout de `app_core.py`, qui centralise les types de personne, les statuts de fiche, le planning hebdomadaire par défaut, la création de fiches indépendantes, la détection des doublons et la validation métier de base.
- Ajout de `pages/Home.py` et intégration de la page d’accueil pour les administrateurs et responsables. Cette page prépare le futur point d’entrée de création et les compteurs de suivi.
- Conversion de `overview.md` et `dev-requirements.md` en français pour respecter la règle de langue du projet.

### Deuxième livraison technique

- Remplacement de l’ancienne page unique `Fiches.py` par une interface organisée en cinq espaces : personnes, création, file d’export, historique et corbeille.
- Ajout de la création de personnes, de la création depuis une personne ou une fiche vierge, de la détection de doublons et du calendrier mensuel par demi-journée.
- Ajout des actions de cycle de vie : terminer, archiver, désarchiver, envoyer à la corbeille et restaurer.
- Conservation de l’API de configuration actuelle avec des requêtes à délai limité. La couche d’export sera reconnectée à la nouvelle file dans la prochaine étape.

### Troisième livraison technique

- Ajout de `export_service.py` pour produire des exports Word, PDF et Excel à partir des nouvelles fiches mensuelles.
- Ajout des exports fusionnés : une page par personne dans Word/PDF et une feuille par personne dans Excel.
- Ajout de l’export séparé par lot : un ZIP contient un dossier identifié par personne et ses fichiers sélectionnés.
- Ajout de la dépendance `reportlab` pour la génération PDF.
- Test mémoire réussi pour les trois formats et le ZIP, sans écrire de document métier sur le disque.

### Limites connues avant la suite

- Les règles de saisie quotidienne des employés et d’approbation restent en attente de confirmation métier.
- Les brouillons non enregistrés doivent être persistés localement dans le navigateur. Streamlit ne fournit pas ce stockage durable nativement ; cette intégration sera isolée afin de ne pas mélanger les brouillons locaux avec les données serveur.
- L’API Presence et son schéma ne sont pas présents dans ce dépôt. Les nouveaux objets persistés devront rester compatibles avec l’endpoint de configuration existant ou nécessiter une évolution coordonnée de l’API.

## 2026-09-08 — Vérification fonctionnelle et correction des régressions

### Clarification du périmètre

- Pas de salaire ni de calcul de paie pour les salariés. Les stagiaires conservent les indemnités, le taux horaire et le transport.
- Les entrées précédentes décrivaient surtout une ossature : les exports contenaient un résumé d’exceptions, sans calendrier complet ni montants de stage. La présente itération corrige ces omissions ; elle ne constitue pas une validation de production de toutes les exigences.

### Changements réalisés

- Calcul partagé des demi-journées et heures mensuelles à partir du planning figé de la fiche. Prise en compte des jours fériés français, repos, week-ends travaillés et remplacements manuels. Les horaires inversés sont rejetés.
- Ajout d’une vue mensuelle en grille et d’un éditeur matin/après-midi pour le jour sélectionné.
- Rétablissement des champs de contrat, de stage et d’indemnité. Une modification de personne ne réécrit pas les fiches existantes. Le montant du stage conserve la base manuelle précédente ; aucune nouvelle règle de rémunération des congés n’est inventée.
- Réutilisation du modèle Excel salarié et du modèle Word stagiaire sans modifier les fichiers sources. Export du mois complet pour les salariés, et de l’indemnité pour les stagiaires. Les exports PDF enveloppent le texte au lieu de tronquer silencieusement les dernières lignes.
- Validation avant export ; préparation atomique des fichiers. Un échec ne marque plus la fiche comme exportée. Le statut change au clic sur le téléchargement (le navigateur ne confirme pas l’écriture effective sur disque).
- Les téléchargements préparés sont isolés par compte et invalidés lorsque les données changent. Les noms identiques ne produisent plus de chemins ZIP en conflit. Les textes utilisateur ne sont pas interprétés comme formules Excel.
- Import explicite et répétable des anciennes configurations, sans effacer les sources. Un chargement API en échec bloque l’écriture au lieu de préparer une configuration vide.
- Restauration de l’état précédent depuis la corbeille, avec contrôle des doublons. Édition manuelle et réexport des archives, désarchivage, archivage par sélection.
- Recherche par nom/type et filtrage par période dans la file et les archives. Suppression définitive avec une boîte de dialogue de confirmation et une action d’annulation.
- Boutons opérationnels sur l’accueil ; les compteurs ne montrent pas de faux zéros avant chargement des données.
- Correction des versions Streamlit contradictoires dans les dépendances, conversion de `requirements.txt` en UTF-8 et ajout des instructions de démarrage.

### Vérification

- 22 tests automatisés passent : calculs, snapshots, jours fériés, migration, doublons, corbeille, contenu des exports, préparation en échec et parcours Streamlit avec API simulée, dont correction d’archive, confirmation de suppression et conservation des descriptions PDF longues.
- Vérification des véritables cellules Excel, textes Word et pages PDF, au-delà de la simple taille des fichiers générés.
- Aucun appel au serveur de production, aucun déploiement, aucun modèle source modifié.
- La fidélité visuelle complète Word/Excel/PDF et le parcours navigateur réel restent à vérifier. Les tests Streamlit ne prouvent pas la durabilité des brouillons dans le navigateur.

### Fonctionnalités encore incomplètes

- Brouillons locaux persistants hors session, sauvegarde serveur par fiche et accès global administrateur.
- Historique des fiches sauvegardées non archivées et ouverture directe des doublons.
- Expiration à 60 jours exécutée côté serveur.
- Mise en page fidèle entre tous les formats, notamment exports mixtes. Une description très longue peut nécessiter une page de continuation plutôt que perdre du contenu.

## 2026-09-08 — Mise à disposition du test local

- Ajout du point d’entrée isolé `local_preview.py` et du lanceur PowerShell `start-local.ps1`, avec écoute uniquement sur `127.0.0.1:8501`. Le parcours de connexion de production reste inchangé.
- Le premier prototype comportait deux personnes fictives ; elles ont ensuite été retirées par la migration documentée ci-dessous. Le test local actuel démarre sans données de démonstration.
- Sauvegarde explicite des données de test dans une base SQLite locale ignorée par Git. Les brouillons incomplets peuvent être sauvegardés localement ; la validation reste obligatoire avant export. Aucun appel à l’API Presence dans ce parcours.
- Le bouton de création de l’accueil ouvre désormais directement l’onglet de création. Correction d’une normalisation de case à cocher qui marquait à tort une fiche avec congé payé comme brouillon lors de son affichage.
- Vérification de 24 tests automatisés, dont le parcours sans secrets ni réseau et la relecture après sauvegarde locale. `pip check` ne signale aucun conflit.
- Vérification dans le navigateur du chargement, de la navigation, des fiches fictives terminées et de la préparation d’un PDF fusionné avec bouton de téléchargement. Le téléchargement final et la fidélité visuelle de tous les formats restent à valider par l’utilisateur.
- Instructions de lancement et limites documentées dans le README. Le service de test est lancé et la page locale ouverte ; aucun déploiement public ou accès aux données de production.
- Ce mode local n’implémente pas la persistance automatique dans le navigateur : les modifications non sauvegardées restent liées à la session. Les autres limites de production de l’entrée précédente restent ouvertes.

## 2026-09-08 — Interaction directe avec le calendrier

- Suppression des personnes et fiches de démonstration livrées dans le test local. Une migration ciblée retire aussi ces anciens exemples de la base locale sans supprimer les données créées par l’utilisateur.
- Remplacement de la liste déroulante des dates par des boutons dans la grille mensuelle. La date active est mise en évidence et un clic ouvre directement son édition.
- Déplacement du panneau matin/après-midi au-dessus de la grille. Les statuts sont des choix visibles et directs ; la précision, la raison et les heures du statut `Autre` s’affichent dans ce panneau.
- Vérification du parcours réel dans le navigateur : sélection du 15 septembre, changement du matin vers `Autre` et apparition du champ de raison avant les boutons du calendrier.
- Confirmation technique et automatisée du parcours de production : connexion obligatoire via l’API Presence et trois rôles `Admin`, `Responsable`, `Employe`. Le test local reste volontairement sans connexion.
- Les 28 tests automatisés passent, dont les nouveaux contrôles de connexion anonyme, de routage d’un employé vers son profil et des trois choix de rôle dans l’administration.
- Les contrôles de pages existent toujours, mais l’accès administrateur global aux fiches de tous les responsables et l’autorisation effective côté API ne peuvent pas être garantis par ce seul dépôt.

## 2026-09-08 — Nommage des fichiers exportés

- Les exports d’une seule fiche utilisent désormais le préfixe `NOM_Prénom_MM-AAAA`, suivi du type de document, dans les trois formats.
- Les fichiers contenus dans un ZIP individuel reprennent exactement ce même préfixe. Le ZIP d’une seule personne est lui aussi nommé avec la personne et la période.
- Un export fusionné de plusieurs personnes du même mois commence par `MM-AAAA`. Une sélection couvrant plusieurs mois utilise le préfixe `multi-periodes` afin de ne pas afficher une période trompeuse.

## 2026-09-08 — Tableau mensuel des stagiaires

- Correction d’une omission : les exports des stagiaires ne contenaient que les informations d’indemnité et de transport.
- L’Excel stagiaire conserve le récapitulatif d’indemnité en haut et ajoute un tableau mensuel avec date, matin, après-midi, heures journalières et total mensuel.
- Le PDF stagiaire contient désormais le même tableau après son récapitulatif.
- Le Word stagiaire conserve le modèle d’indemnité existant et ajoute le tableau mensuel. Des bordures sont appliquées directement lorsque le modèle ne contient pas le style Word `Table Grid`.
- Les contrôles automatisés vérifient le premier et le dernier jour du mois ainsi que le total d’heures dans les trois formats.

## 2026-09-08 — Préparation de la connexion au stockage VPS

- Confirmation dans l’historique du dépôt que le frontal accédait déjà aux données MariaDB du VPS par l’intermédiaire de l’API Presence, et non par une connexion SQL directe.
- Ajout de `presence_storage.py` pour centraliser `GET /get-config/{username}` et `POST /save-config/{username}`, l’authentification Bearer, les délais d’attente, l’encodage du nom de compte et les erreurs de transport ou de format.
- Conservation du contrat JSON historique et des clés anciennes. La sérialisation crée une copie et ne modifie plus les dates présentes dans l’espace de travail en mémoire.
- Ajout de tests isolés du réseau et du serveur. Aucun appel et aucune écriture de production ne sont réalisés.
- Documentation du contrat connu et des contrôles encore nécessaires dans `doc/storage-contract.md`. Le code du serveur Presence, le schéma MariaDB et les secrets de production ne sont pas présents dans ce dépôt.
- La sauvegarde individuelle et la persistance locale des brouillons restent volontairement désactivées : l’endpoint historique remplace une configuration complète et doit d’abord être protégé contre les écrasements concurrents.
