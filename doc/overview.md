# Fiches de présence

> 历史记录：文中的 Streamlit 路径已于 2026-10-23 分离至[源码备案](../archive/streamlit/README.md)。当前程序定位以[模块索引](modules.md)为准。

## Objectif

Cette application Streamlit permet de préparer des fiches mensuelles de présence pour des salariés et des stagiaires. Elle propose un calendrier par demi-journée, l'export des documents et une administration des comptes. L'interface est en français; le code et les identifiants techniques sont en anglais.

## Composants

Voir [l'index des modules](modules.md) pour les responsabilités et les dépendances actuelles.

- `online_test.py` : point d'entrée et navigation selon le rôle.
- `pages/Login.py` : connexion, demande et confirmation de réinitialisation du mot de passe.
- `remember_me.py` : synchronisation contrôlée du jeton de connexion persistant dans le navigateur.
- `pages/Fiches.py` : création des fiches, calendrier, sauvegarde et exports.
- `fiche_exports.py`, `fiche_serialization.py` et `planning.py` : validation, génération, dates et planning par défaut partagés.
- `pages/TransportReceipts.py`, `document_ui.py` et `document_downloads.py` : dépôt, cartes, filtres, archives et ZIP des Documents.
- `pages/Profile.py` : modification du profil connecté.
- `pages/Admin.py` : administration centrée sur la fiche utilisateur, avec sélection des comptes, sections repliables, départements, étiquettes et affectation hiérarchique directe.
- `pages/People.py` : vue opérationnelle du personnel accessible aux responsables autorisés.
- `backend/main.py` : API FastAPI versionnée qui applique l'authentification et les autorisations.
- `backend/documents.py` et `backend/esign.py` : routes Documents, conversion PDF et signature.
- `backend/contract_history.py` et `contract_periods.py` : historique des contrats et rattachement mensuel des fiches.
- `ExcelGen.py` et `DocxGen.py` : génération des exports Excel et Word.
- `calendar_view.py` : calendrier mensuel et sélection groupée de demi-journées.

## Fonctionnement métier

Une fiche est créée pour un mois et une année, à partir d'une personne existante ou d'une fiche vide. Les horaires par défaut de chaque personne remplissent le calendrier; les samedis et dimanches sont des jours de repos par défaut pour les nouveaux plannings. Chaque matin et après-midi peut être défini comme travail, congé payé, absence, arrêt maladie, férié ou autre. Une modification ne change que la fiche du mois concerné, jamais le planning par défaut de la personne.

Le calendrier possède un mode consultation et un mode `Modifier`. Dans ce dernier, plusieurs demi-journées peuvent être sélectionnées par clic ou glissement de souris, puis recevoir un statut en une action. Les séquences adjacentes de même statut sont affichées sous forme de barres continues à extrémités arrondies.

Les fiches sont enregistrées dans la configuration de leur propriétaire. Les salariés disposent d'un export Excel et les stagiaires d'un export Word; les originaux peuvent être convertis en PDF. L'export groupé produit un ZIP contenant un fichier par fiche. La suppression d'une fiche est réservée à l'administrateur; aucune corbeille de 60 jours n'est actuellement implémentée.

Le module Documents conserve les justificatifs par mois de rattachement, avec les types Transport ou Autre. La liste distingue les documents à traiter et les archives; les cartes peuvent être sélectionnées pour un téléchargement ZIP. L'archivage conserve le fichier. Le rétablissement d'un document archivé n'est pas encore proposé.

Chaque salarié et responsable utilise son propre compte pour déposer ses factures.
Le propriétaire est automatiquement lié à la session et l'API refuse une attribution
à autrui. Un Admin peut déposer pour une autre personne. Le responsable peut déposer
ses propres documents et consulter ceux de son équipe; ses droits d'archivage,
de correction du mois et de suppression ne sont pas étendus.

La signature ClawShow suit l'ordre salarié puis responsable direct, avec deux
zones de signature distinctes. Le parcours salarié en production a été confirmé
par l'utilisateur; le parcours complet stagiaire reste à vérifier.

## Utilisateurs et accès

L'API délivre à la connexion un jeton de session signé et limité dans le temps. L'option « Rester connecté sur cet appareil » ajoute un jeton persistant opaque, stocké dans un cookie `Secure` (en HTTPS) et `SameSite=Strict`, qui expire après 30 jours. Il ne contient pas le mot de passe et est invalidé à la déconnexion ou lors d'un changement de mot de passe. Les appels sensibles utilisent le jeton de session; le contrôle d'accès est appliqué par l'API, pas uniquement par le menu Streamlit.

| Rôle | Accès |
| --- | --- |
| `Admin` | Tous les comptes, groupes, fiches et paramètres d'organisation. |
| `Responsable` | Consultation du périmètre autorisé et saisie des présences de son équipe; pas de modification des comptes ou contrats ni de suppression des fiches. |
| `Employe` | Son profil et sa propre consultation; il ne peut ni gérer des personnes ni enregistrer une fiche pour autrui. |

Les groupes sont créés et activés par un administrateur. Un employé peut appartenir à plusieurs groupes, sans que cela soit obligatoire; son responsable direct peut être un `Responsable` ou un `Admin`. L'historique des contrats accepte les types `salarie`, `stagiaire` et `alternance` et refuse les périodes qui se chevauchent. Un changement de contrat en cours de mois peut produire deux fiches liées à leurs périodes respectives.

La page Administration affiche une structure en arbre : `Organisation` au sommet, suivie des responsables directs et de leurs employés. Les Groupes sont affichés dans les nœuds des employés, sans devenir des parents supplémentaires dans l'arbre.

## Stockage et déploiement

MariaDB conserve les comptes dans `users` et les configurations de fiches dans `Presence`. Les migrations d'organisation sont additives : elles ajoutent des colonnes facultatives à `users` et créent `organization_groups` ainsi que `user_group_memberships`. Elles ne suppriment ni comptes ni fiches existantes.

En production, l'API tourne sur `127.0.0.1:8001` sous `presence-app.service`, et l'interface Streamlit sous `presence-interface.service`. Les secrets sont lus depuis le fichier d'environnement protégé du serveur et ne doivent jamais être ajoutés au dépôt.

La suppression explicite d'un compte supprime actuellement ses fiches, contrats
et justificatifs associés, y compris les documents qu'il a déposés ou archivés
pour autrui. Elle ne constitue pas un archivage et ne permet pas de restauration.
Le [plan de maintenance](code-audit.md) décrit les limites de cache, les erreurs
de chargement et le problème de nom de fonction dans les actions groupées.

## Vérification avant publication

1. Compiler les modules Python modifiés; exécuter les tests demandés et distinguer leurs résultats des vérifications de service.
2. Sauvegarder le code de l'API et exporter la base MariaDB avant toute migration.
3. Déployer l'API et l'interface dans la même fenêtre de maintenance courte, puis redémarrer les deux services.
4. Vérifier l'état des services, la page publique, la connexion d'un administrateur et l'absence de perte de comptes ou de fiches.
