# Exigences — rôles, personnel et Groupes

État aligné sur le code publié le 14 octobre 2026. Les limites et les corrections
à prévoir sont détaillées dans [le plan de maintenance](code-audit.md).

## Objectif

Le système Presence doit limiter les données visibles et modifiables au périmètre
organisationnel de l'utilisateur connecté, tout en conservant les comptes, les
fiches et les rôles existants.

## Rôles et types de personnel

- `Admin` : étiquette de permission système; elle peut être combinée à
  `Responsable` et/ou `Employe`.
- `Responsable` : gère uniquement les Employés dont il est le responsable direct
  ou qui appartiennent à au moins un Groupe qu'il gère.
- `Employe` : consulte uniquement son profil et ses résultats autorisés ; il ne
  peut pas modifier les présences, Groupes ou autres comptes.
- Un `Employe` porte le type `salarie` ou `stagiaire`. Ce type ne constitue pas
  un rôle supplémentaire.

Les étiquettes historiques `Admin`, `Responsable` et `Employe` sont conservées.
Un compte peut en porter plusieurs. Un ancien compte sans type de personnel est
interprété comme `salarie`.

## Modèle organisationnel

- Un Groupe possède un nom unique et un état actif/inactif.
- Un Employé possède un seul responsable direct, qui peut avoir le rôle
  `Responsable` ou `Admin`. Son appartenance à un ou plusieurs Groupes est
  facultative.
- Un Responsable gère un ou plusieurs Groupes.
- Un Groupe inactif ne supprime ni utilisateurs ni fiches historiques. Il ne peut
  plus être attribué à un nouveau compte.
- Les associations sont conservées dans des tables relationnelles afin qu'un
  utilisateur puisse appartenir à plusieurs Groupes.

## Autorisations

### Admin

- Crée, modifie et supprime les comptes.
- Crée, renomme, active ou désactive les Groupes.
- Assigne les Groupes gérés aux Responsables, puis les Groupes facultatifs et le
  responsable direct (Responsable ou Admin) aux Employés.
- Accède à toutes les fiches de présence.

### Responsable

- Consulte les personnes de son périmètre et prépare leurs présences.
- Enregistre uniquement les champs de présence autorisés par
  `backend/attendance_permissions.py`.
- Ne crée, ne modifie et ne supprime aucun compte, contrat ou Groupe.
- Ne supprime pas les fiches; dépose ses propres Documents et consulte ceux de
  son équipe sans les modifier, archiver ou supprimer.

### Employe

- Peut se connecter et consulter les informations qui lui sont attribuées.
- Ne peut pas modifier les présences, l'organisation ni les autres comptes.
- Doit utiliser son compte personnel pour déposer ses factures et justificatifs;
  l'API refuse une attribution à un autre utilisateur. Un Admin peut déposer pour autrui.

## API et persistance

- L'API émet après connexion un jeton de session signé et limité dans le temps.
  Les routes protégées ne doivent plus accepter un rôle transmis par le client.
- L'API vérifie chaque lecture, création, modification et suppression ; masquer
  une page Streamlit ne constitue pas une sécurité suffisante.
- La migration MariaDB est additive : ajout de `employee_type` et
  `manager_username` dans `users`, puis création de `organization_groups` et
  `user_group_memberships`. Elle ne modifie pas les fiches existantes.
- Actuellement, la suppression d'un compte est définitive : son contenu Presence,
  les fiches associées dans les autres configurations, ses contrats et les
  justificatifs associés sont supprimés. Les fichiers déposés ou archivés par ce
  compte pour d'autres personnes sont aussi concernés. Les associations de
  Groupes sont retirées et les liens de responsable direct sont neutralisés.
  Une politique de conservation doit être définie avant de modifier ce comportement.

## Interface

- La page Administration permet à un Admin de gérer les Groupes, créer des
  comptes et modifier les affectations.
- La page `Mon équipe` présente à un Responsable les personnes de son périmètre
  en consultation. La saisie des présences se fait dans `Création de Fiches`.
- Les choix de Groupes affichent uniquement les Groupes actifs et autorisés.
- Les erreurs d'autorisation et les données organisationnelles incomplètes sont
  affichées explicitement en français.

## Critères d'acceptation

1. Les comptes existants peuvent toujours se connecter.
2. Un Responsable ne peut ni consulter ni appeler l'API pour une ressource hors
   de son périmètre.
3. Un Employé ne peut pas enregistrer une fiche de présence, même par appel API
   direct.
4. Un Admin peut gérer plusieurs Groupes et attribuer plusieurs Groupes à un
   Employé ou à un Responsable.
5. Les migrations additives préservent les fiches MariaDB existantes. La suppression
   explicite d'un compte suit actuellement le comportement définitif décrit ci-dessus.
