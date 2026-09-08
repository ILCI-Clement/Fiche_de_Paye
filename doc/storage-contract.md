# Contrat de stockage Presence

## Architecture constatée

L’application Streamlit ne se connecte pas directement à MariaDB. Les versions historiques du dépôt chargent les données du VPS au moyen de l’API Presence. MariaDB se trouve donc derrière cette API et ses identifiants ne doivent pas être copiés dans le frontal.

Le dépôt actuel ne contient ni le code du serveur Presence, ni le schéma MariaDB, ni une configuration de production locale. Une écriture réelle ne doit pas être tentée tant que ces éléments et une sauvegarde de la base n’ont pas été vérifiés.

## Interface existante

Les opérations de configuration observées sont les suivantes :

- `GET /get-config/{username}` avec un jeton Bearer ;
- `POST /save-config/{username}` avec un jeton Bearer et un objet JSON dans le corps ;
- la réponse de lecture est directement l’objet de configuration, sans enveloppe supplémentaire dans le code historique.

La nouvelle couche `presence_storage.py` conserve ce contrat, encode le nom du compte dans l’URL, impose un délai d’attente, valide la forme générale de la réponse et transforme les dates sans modifier les données en mémoire.

## Structure fonctionnelle actuelle

Le document de travail utilise `schema_version: 2` et conserve au minimum :

- `people` : bibliothèque des salariés et stagiaires ;
- `forms` : fiches actives, exportées ou archivées ;
- `trash` : fiches placées dans la corbeille.

Les anciennes clés restent présentes lors du chargement et de la sauvegarde afin de permettre un import contrôlé des anciennes données.

## Points à confirmer côté serveur

Avant la connexion de production, il faut vérifier dans le dépôt ou sur le VPS :

- la table et la colonne qui contiennent la configuration JSON ;
- le type MariaDB de cette colonne et sa taille maximale ;
- le comportement lorsqu’un utilisateur ne possède encore aucune configuration ;
- la limite de taille du corps HTTP et la gestion des écritures simultanées ;
- le contrôle d’autorisation qui empêche un responsable de lire ou modifier les données d’un autre compte ;
- la sauvegarde et la restauration de la base avant toute migration ;
- le mécanisme serveur nécessaire à la suppression définitive après 60 jours.

## Sauvegarde individuelle

Le besoin fonctionnel impose que le responsable choisisse, fiche par fiche, si une donnée doit être envoyée au serveur. L’endpoint historique remplace une configuration JSON complète. Une stratégie de fusion et de versionnement doit donc être définie côté API avant d’activer la sauvegarde individuelle, afin qu’un onglet ancien ne puisse pas écraser des modifications plus récentes.

Les brouillons non envoyés au serveur devront rester dans un stockage local durable propre au navigateur. La base SQLite du mode de test n’est pas ce stockage de production.
