export const workspacePages = {
  attendance: {
    title: 'Création de fiches',
    description: 'Retrouvez vos fiches par mois et préparez les exports ou la signature. Dans le calendrier, cliquez ou glissez sur les demi-journées pour sélectionner ; partez d’une case sélectionnée pour désélectionner, puis choisissez un statut et appliquez. Le bouton en bas à droite enregistre vos modifications.',
  },
  documents: {
    title: 'Documents et justificatifs',
    description: 'Déposez et retrouvez vos documents par mois. Recherchez par nom, numéro ou employé, filtrez par type et consultez les archives. Une sélection peut être téléchargée en ZIP.',
  },
  leave: {
    title: 'Congés payés',
    description: 'Envoyez une demande de congé à votre responsable direct et suivez son traitement. Les responsables et administrateurs peuvent examiner les demandes de leur périmètre. Chaque décision conserve son auteur et sa date ; les demandes retirées sont masquées.',
  },
  team: {
    title: 'Mon équipe',
    description: 'Consultez la structure de l’entreprise et la liste de votre équipe selon votre périmètre autorisé. Les profils mensuels regroupent les présences, documents, signatures et congés accessibles. Les comptes et contrats sont gérés par un administrateur.',
  },
  admin: {
    title: 'Comptes et départements',
    description: 'Recherchez un compte et gérez ses rôles, départements, responsable et contrats. La création d’un compte de connexion est distincte de l’inscription au registre du personnel.',
  },
  dashboard: {
    title: 'Tableau de bord',
    description: 'Consultez la répartition des effectifs, les échéances contractuelles et les entretiens annuels. Les actions disponibles dépendent de vos droits.',
  },
  profile: {
    title: 'Paramètres du compte',
    description: 'Modifiez votre nom d’utilisateur, votre adresse e-mail ou votre mot de passe avec votre mot de passe actuel. Enregistrez d’abord vos autres modifications : la mise à jour renouvelle la session et recharge votre espace.',
  },
  personnel: {
    title: 'Registre du personnel',
    description: 'Enregistrez et recherchez les salariés, stagiaires et alternants. Une ligne représente une personne ; le bouton « … » ouvre ses informations, périodes d’emploi et modifications. Le compte de connexion est géré séparément.',
  },
} as const

export type WorkspacePage = keyof typeof workspacePages
