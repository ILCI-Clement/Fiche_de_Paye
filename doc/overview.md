# Fiche de Paye / Fiches de présence

## État actuel du projet

L’application est un outil Streamlit interne de présence et d’indemnités de stage en cours de refonte. La page principale propose une bibliothèque de personnes, des fiches mensuelles avec planning figé, un calendrier matin/après-midi, une file d’export, des archives et une corbeille. Les exports Word, PDF et Excel sont disponibles. Aucun salaire n’est calculé pour les salariés ; les indemnités des stagiaires sont conservées.

Les tests automatisés vérifient les calculs et le contenu des fichiers. La persistance des brouillons dans le navigateur, la sauvegarde individuelle, l’accès global administrateur et le nettoyage serveur après 60 jours ne sont pas encore implémentés. Les détails vérifiés sont consignés dans `doc/dev-log.md`.

Ce n’est pas encore un système complet de gestion des présences : il n’existe pas de pointage quotidien, de demande de congé, de circuit de validation, d’historique métier structuré, de tableau de bord ou de notification automatique.

## Structure du dépôt

- `online_test.py` : point d’entrée Streamlit et navigation selon le rôle.
- `pages/Login.py` : connexion et demande de réinitialisation du mot de passe.
- `pages/Profile.py` : modification du profil connecté.
- `pages/Fiches.py` : saisie des fiches, synchronisation de configuration et exports.
- `pages/Admin.py` : création, liste et suppression des comptes.
- `app_core.py` : création, validation, import des anciennes configurations et transitions de fiches.
- `attendance.py` : calcul des demi-journées, heures mensuelles et indemnités.
- `export_service.py` : génération, validation et préparation des exports.
- `ExcelGen.py` et `DocxGen.py` : anciens générateurs, conservés comme référence, non appelés par la nouvelle interface.
- `Fiche_Exemple.xlsx` et `template_stagiaire.docx` : modèles d’export existants.

## Rôles et persistance actuelle

La navigation affiche :

| Rôle | Pages accessibles |
| --- | --- |
| `Admin` | Profil, création de fiches, administration |
| `Responsable` | Profil, création de fiches |
| `Employe` | Profil |

Les données de fiches sont actuellement enregistrées comme une seule configuration JSON par utilisateur connecté via l’API Presence. Le dépôt ne contient ni la base de données, ni le code de cette API. Les contrôles de rôle dans Streamlit ne suffisent donc pas à prouver la sécurité effective des routes API.

Les routes utilisées côté client sont : `POST /login`, `POST /forgot-password`, `PUT /update_profile/{username}`, `GET /get-config/{username}`, `POST /save-config/{username}`, `POST /create-user`, `GET /list-users` et `DELETE /delete-user/{username}`.

## Fonctionnement des anciens générateurs (référence)

Pour un salarié, l’interface collecte identité, responsable, contrat, planning hebdomadaire, congés payés, absences et arrêts maladie. Pour un stagiaire, elle collecte identité, dates de stage, jours, heures, taux horaire, transport et remboursement.

Le générateur Excel écrit toutes les valeurs et tous les calculs dans un classeur sans formules. Il utilise les jours fériés français, mais force actuellement les week-ends en `WEEK-END`, même si un planning les rend travaillés. Le générateur Word calcule l’indemnité de stage et le remboursement transport, puis écrit le résultat dans le modèle existant.

## Limites initiales (avant la refonte)

1. Validation incomplète des dates, horaires, taux et chevauchements avant génération.
2. Conversion incorrecte des créneaux passant minuit et libellé incorrect pour certaines absences ou certains arrêts l’après-midi.
3. Chemins de modèles dépendants du répertoire de lancement.
4. Requêtes HTTP sans politique commune de délai, d’erreur ou de validation des réponses.
5. Absence de tests automatisés, de documentation de démarrage et d’exemple de secrets Streamlit.
6. Aucune relation structurée entre responsable, personne, fiche et historique.
7. Risque d’état résiduel dans les widgets Streamlit dynamiques.
8. L’e-mail de rappel de fin de contrat annoncé dans l’interface n’est pas implémenté.

Plusieurs de ces limites sont corrigées dans le nouveau parcours : validation des heures, chemins de modèles absolus, délais réseau, snapshots et tests. Les anciens générateurs restent inchangés ; les limites de la version actuelle figurent dans le journal de développement.

## Direction de la refonte

Un point d’entrée séparé, `local_preview.py`, permet maintenant de tester la refonte sans API ni connexion de production, avec des personnes fictives et une sauvegarde explicite SQLite sur l’ordinateur. Il est réservé à un lancement en boucle locale et ne remplace pas le futur stockage des brouillons dans le navigateur. Voir le README pour le démarrage.

La cible fonctionnelle est décrite dans `doc/dev-requirements.md`. La refonte introduit une bibliothèque de personnes, des fiches mensuelles indépendantes, une file d’export, un historique, une corbeille et un calendrier mensuel matin/après-midi, tout en préservant les modèles de sortie existants.
