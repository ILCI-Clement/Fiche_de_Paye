# Exigences de développement

## Positionnement confirmé

- Application interne pour les responsables et leurs collègues. Aucun accès public, aucune inscription publique et aucune fonctionnalité commerciale ne sont prévus.
- Le nombre d’utilisateurs attendu est compris entre 10 et 50.
- Les rôles existants `Admin`, `Responsable` et `Employe` sont conservés. Un administrateur gère toutes les données. Un responsable utilise la bibliothèque de personnes, la file et les exports. Un employé gère uniquement son profil.
- Le mode de connexion actuel est conservé. Un compte n’est pas obligatoirement associé à une fiche de personne.
- Les responsables ne voient et ne gèrent que les personnes, fiches, historiques et éléments de corbeille qu’ils ont créés. Les administrateurs ont accès à toutes les données.
- Cette version ne crée pas de relation d’affectation unique entre une personne et un responsable.

## Personnes et fiches mensuelles

- Le système couvre salariés et stagiaires.
- Une bibliothèque de personnes est distincte des fiches mensuelles. La création d’une fiche depuis la bibliothèque reprend les informations réutilisables : identité, contrat ou stage, responsable et planning.
- Une personne peut passer du type salarié au type stagiaire, et inversement.
- Chaque personne possède un planning hebdomadaire par défaut, avec des horaires distincts selon les jours. Les horaires peuvent donc varier d’une personne à l’autre.
- Toute nouvelle fiche mensuelle reprend le planning par défaut. Une modification ultérieure du planning d’une personne n’affecte aucune fiche déjà créée ou archivée.
- La première version conserve les informations existantes : contrat, planning, congés payés (`CP`), absences (`ABS`) et arrêts maladie (`AM`) pour les salariés ; dates de stage, taux horaire, jours, heures quotidiennes, transport et taux de remboursement pour les stagiaires.
- Salariés et stagiaires utilisent un calendrier mensuel partagé, découpé en matin et après-midi, pour déclarer les exceptions.
- Les exceptions sont saisies dans une vue calendrier. Un clic sur une date la sélectionne directement, sans liste déroulante ; le panneau de modification et les champs de raison sont affichés au-dessus du calendrier. Chaque demi-journée n’accepte qu’un seul statut : travail normal, CP, absence, arrêt maladie, jour férié/fermeture, ou `Autre`. Le statut `Autre` comporte un libellé et un nombre d’heures saisis manuellement.
- Les jours fériés français sont préremplis automatiquement. Le responsable peut ajouter ou modifier des jours fériés spéciaux et des fermetures d’entreprise.
- Il n’y a pas de fonction distincte pour modifier ponctuellement les horaires d’une journée normale, ni de modification de planning en masse sur des fiches existantes.
- Un responsable ne peut pas créer deux fiches pour la même personne et le même mois. Pour une fiche vierge, la détection utilise le nom, le type de personne et le mois. En cas de doublon, l’application affiche un message et permet d’ouvrir la fiche existante.
- Une fiche vierge ne crée jamais une personne dans la bibliothèque et ne peut pas modifier une personne existante. La création de personne reste une action séparée.
- Toute sauvegarde, export ou archivage applique une validation commune. Les champs obligatoires manquants et les conflits du calendrier bloquent l’action et sont signalés précisément.

## Accueil, file, historique et corbeille

- La page d’accueil donne la priorité à la création d’une nouvelle fiche : depuis une personne existante ou depuis une fiche vierge.
- Une file de fiches permet de créer plusieurs fiches, de les compléter, puis de les traiter individuellement ou par lot.
- Les statuts visibles dans la file sont au minimum `Brouillon`, `Terminé` et `Exporté`.
- Après export, une fiche reste dans la file. Le système ne sauvegarde, n’archive, ne déplace et ne supprime jamais automatiquement une fiche exportée.
- Le responsable peut archiver une ou plusieurs fiches exportées. Une fiche non archivée reste dans la file d’export.
- Une fiche sauvegardée sur le serveur est disponible dans l’historique. L’historique permet de rechercher par nom, type et mois, puis de modifier, réexporter ou supprimer une fiche.
- Une fiche archivée est figée : elle ne suit pas automatiquement les modifications de la personne. Le responsable peut toutefois la modifier manuellement. Après une correction, elle reste archivée par défaut, mais peut être désarchivée manuellement et retourne alors dans la file.
- La suppression déplace d’abord la fiche dans la corbeille. La suppression définitive exige une seconde confirmation. Une fiche peut être restaurée ; les éléments de corbeille sont supprimés définitivement après 60 jours.
- Bibliothèque de personnes, file et historique proposent une recherche et des filtres au minimum par nom, type et mois.

## Sauvegarde des brouillons

- La personne qui remplit une fiche choisit explicitement de la sauvegarder ou non sur le serveur.
- Une fiche sauvegardée devient une donnée persistante qui peut être rouverte et réexportée lors d’une connexion ultérieure.
- Une fiche non sauvegardée reste un brouillon local sur le navigateur et l’appareil qui l’ont créée. Elle survit au rechargement, à la déconnexion et à la fermeture du navigateur jusqu’à sa suppression manuelle.
- Les brouillons locaux ne sont ni synchronisés entre navigateurs, ni synchronisés entre appareils.

## Exports

- Les salariés n’ont aucun calcul ni champ de salaire dans cette application. Les stagiaires conservent leur indemnité de stage, le taux horaire et le remboursement du transport (clarification du 8 septembre 2026).
- Le calcul de l’indemnité de stage conserve provisoirement la base existante saisie manuellement : jours indemnisés × heures par jour × taux horaire. Une exception de calendrier ne change pas automatiquement cette base ; la règle métier de rapprochement reste à confirmer.
- Salariés et stagiaires peuvent chacun être exportés séparément en Word, PDF ou Excel.
- Les trois formats contiennent, pour les salariés comme pour les stagiaires, un tableau mensuel avec la date, le statut du matin, le statut de l’après-midi et les heures. Pour les stagiaires, ce tableau complète les informations d’indemnité et de transport sans les remplacer.
- Les modèles Excel des salariés et Word des stagiaires sont conservés sans modification de mise en page ou de champs.
- Les combinaisons qui n’ont pas encore de modèle natif, par exemple Word pour un salarié ou Excel pour un stagiaire, reprennent la structure de contenu et le style visuel des modèles existants, sans nouvelle direction artistique.
- L’utilisateur peut exporter une sélection ou toute la file.
- En export fusionné, chaque personne commence sur une page distincte dans Word et PDF ; dans Excel, chaque personne occupe sa propre feuille.
- En export séparé par lot, le téléchargement est un ZIP. Chaque personne dispose d’un dossier nommé à partir de son nom, contenant ses fichiers exportés.
