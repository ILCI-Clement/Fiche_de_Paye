# Exigences — rôles, personnel et Groupes

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

- Crée uniquement des comptes `Employe` ; il devient automatiquement leur
  Responsable direct. Les Groupes attribués restent limités à ses propres Groupes.
- Voit, modifie et supprime uniquement les Employés de son périmètre.
- Ne peut ni créer ni modifier un Admin, un autre Responsable ou un Groupe.
- Ne peut pas attribuer un Groupe qu'il ne gère pas.

### Employe

- Peut se connecter et consulter les informations qui lui sont attribuées.
- Ne peut pas modifier les présences, l'organisation ni les autres comptes.

## API et persistance

- L'API émet après connexion un jeton de session signé et limité dans le temps.
  Les routes protégées ne doivent plus accepter un rôle transmis par le client.
- L'API vérifie chaque lecture, création, modification et suppression ; masquer
  une page Streamlit ne constitue pas une sécurité suffisante.
- La migration MariaDB est additive : ajout de `employee_type` et
  `manager_username` dans `users`, puis création de `organization_groups` et
  `user_group_memberships`. Elle ne modifie pas les fiches existantes.
- La suppression d'un compte conserve son contenu Presence historique. Les
  associations de Groupes sont retirées et les liens de responsable direct sont
  neutralisés.

## Interface

- La page Administration permet à un Admin de gérer les Groupes, créer des
  comptes et modifier les affectations.
- La page Personnel présente à un Responsable uniquement ses Employés et lui
  permet d'en créer dans son périmètre.
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
5. Les fiches MariaDB existantes restent intactes après migration et après la
   suppression d'un compte.
