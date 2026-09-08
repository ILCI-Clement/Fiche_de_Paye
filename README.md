# Gestion des présences

Application interne Streamlit, interface française. Python 3.12 ou supérieur.

## Installation

Depuis la racine du dépôt, sous PowerShell :

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run online_test.py
```

Configurer `.streamlit/secrets.toml` localement (ne pas le versionner) :

```toml
URL_PRESENCE = "https://adresse-de-votre-api"
PRESENCE_TOKEN = "votre-jeton"
```

Le serveur Presence et la base de données ne sont pas inclus dans le dépôt. Le point d’entrée `online_test.py` conserve la connexion réelle.

Le frontal ne doit pas recevoir d’identifiants MariaDB : il utilise l’API Presence comme passerelle de stockage. Le contrat actuellement connu et les vérifications nécessaires avant une migration sont décrits dans `doc/storage-contract.md`.

## Test local isolé

Après installation des dépendances, démarrer depuis PowerShell :

```powershell
.\start-local.ps1
```

Puis ouvrir [le test local](http://127.0.0.1:8501). Si la politique PowerShell bloque le script, exécuter directement :

```powershell
.\.venv\Scripts\python.exe -B -m streamlit run local_preview.py --server.address 127.0.0.1 --server.port 8501 --server.headless true --browser.gatherUsageStats false
```

Ce point d’entrée distinct n’exige ni compte ni configuration API et ne contacte pas le serveur Presence. Il écoute uniquement sur la boucle locale. Ne pas le publier sur Internet.

Le test local démarre avec une bibliothèque et une file vides. Créer les personnes et fiches de test nécessaires, uniquement avec des données fictives, puis tester les demi-journées du calendrier, les exports Word/PDF/Excel, l’archivage et la restauration.

Le bouton « Sauvegarder les données de test sur cet ordinateur » conserve explicitement l’espace de test dans `.local-test/preview.sqlite3`, exclu de Git. Il ne s’agit pas du stockage dans le navigateur prévu pour la version finale. Les changements non sauvegardés peuvent être perdus à la fermeture ou au rechargement de la session. Utiliser uniquement des données fictives. Arrêter le serveur avec `Ctrl+C` dans son terminal.

## Vérifications

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```

Les tests utilisent des personnes fictives, génèrent les exports en mémoire et simulent les échanges API. Ils ne modifient pas le serveur.

## État de la refonte

Les salariés disposent d’un suivi de présence sans salaire. Les stagiaires conservent le calcul d’indemnité existant. L’Excel salarié utilise le modèle du dépôt ; le Word stagiaire utilise le modèle Word original. Les autres formats disposent actuellement d’une présentation fonctionnelle dont la fidélité visuelle reste à valider.

Les anciennes configurations peuvent être importées explicitement depuis la page de gestion ; leurs données sources sont conservées dans la configuration.

Le stockage durable des brouillons dans le navigateur, la sauvegarde serveur par fiche, la vue administrateur globale et le nettoyage serveur automatique de la corbeille restent à intégrer. En l’état, une fiche non sauvegardée sur le serveur dépend de la session Streamlit. Ne pas utiliser cette version pour des brouillons réels que l’on souhaite conserver uniquement en local.

Les décisions métier se trouvent dans `doc/dev-requirements.md`, et l’avancement vérifié dans `doc/dev-log.md`.
