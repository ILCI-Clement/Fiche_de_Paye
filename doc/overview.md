# Fiches de présence

## Objectif

Cette application Streamlit permet de préparer des fiches mensuelles de présence pour des salariés et des stagiaires. Elle propose un calendrier par demi-journée, l'export des documents et une administration des comptes. L'interface est en français; le code et les identifiants techniques sont en anglais.

## Composants

- `online_test.py` : point d'entrée et navigation selon le rôle.
- `pages/Login.py` : connexion, demande et confirmation de réinitialisation du mot de passe.
- `pages/Fiches.py` : création des fiches, calendrier, sauvegarde et exports.
- `pages/Profile.py` : modification du profil connecté.
- `pages/Admin.py` : administration des comptes et des groupes.
- `pages/People.py` : vue opérationnelle du personnel accessible aux responsables autorisés.
- `backend/main.py` : API FastAPI versionnée qui applique l'authentification et les autorisations.
- `ExcelGen.py` et `DocxGen.py` : génération des exports Excel et Word.
- `calendar_view.py` : calendrier mensuel et sélection groupée de demi-journées.

## Fonctionnement métier

Une fiche est créée pour un mois et une année, à partir d'une personne existante ou d'une fiche vide. Les horaires par défaut de chaque personne remplissent le calendrier; les samedis et dimanches sont des jours de repos par défaut pour les nouveaux plannings. Chaque matin et après-midi peut être défini comme travail, congé payé, absence, arrêt maladie, férié ou autre. Une modification ne change que la fiche du mois concerné, jamais le planning par défaut de la personne.

Le calendrier possède un mode consultation et un mode `Modifier`. Dans ce dernier, plusieurs demi-journées peuvent être sélectionnées par clic ou glissement de souris, puis recevoir un statut en une action. Les séquences adjacentes de même statut sont affichées sous forme de barres continues à extrémités arrondies.

Les fiches peuvent être conservées dans une file, archivées volontairement, restaurées, mises à la corbeille et récupérées pendant 60 jours. Les exports Word, PDF et Excel sont proposés séparément. Un export groupé peut produire une fiche par page ou par fichier, avec des noms incluant la personne, le mois et l'année. Les fiches archivées restent figées, sauf modification volontaire suivie d'un retour dans la file.

## Utilisateurs et accès

L'API délivre à la connexion un jeton de session signé et limité dans le temps. Les appels sensibles utilisent ce jeton; le contrôle d'accès est appliqué par l'API, pas uniquement par le menu Streamlit.

| Rôle | Accès |
| --- | --- |
| `Admin` | Tous les comptes, groupes, fiches et paramètres d'organisation. |
| `Responsable` | Sa propre fiche et les employés qui lui sont rattachés directement ou qui appartiennent à au moins un groupe qu'il gère. |
| `Employe` | Son profil et sa propre consultation; il ne peut ni gérer des personnes ni enregistrer une fiche pour autrui. |

Les groupes sont créés et activés par un administrateur. Un employé peut appartenir à plusieurs groupes; un responsable peut gérer plusieurs groupes. Ces relations, le responsable direct et le type `salarie` ou `stagiaire` sont enregistrés dans MariaDB.

## Stockage et déploiement

MariaDB conserve les comptes dans `users` et les configurations de fiches dans `Presence`. Les migrations d'organisation sont additives : elles ajoutent des colonnes facultatives à `users` et créent `organization_groups` ainsi que `user_group_memberships`. Elles ne suppriment ni comptes ni fiches existantes.

En production, l'API tourne sur `127.0.0.1:8001` sous `presence-app.service`, et l'interface Streamlit sous `presence-interface.service`. Les secrets sont lus depuis le fichier d'environnement protégé du serveur et ne doivent jamais être ajoutés au dépôt.

## Vérification avant publication

1. Compiler les modules Python modifiés et exécuter les tests unitaires disponibles.
2. Sauvegarder le code de l'API et exporter la base MariaDB avant toute migration.
3. Déployer l'API et l'interface dans la même fenêtre de maintenance courte, puis redémarrer les deux services.
4. Vérifier l'état des services, la page publique, la connexion d'un administrateur et l'absence de perte de comptes ou de fiches.
