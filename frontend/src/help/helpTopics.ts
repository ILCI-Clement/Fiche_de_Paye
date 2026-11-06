import { canManageTeam, isAdmin, isSuperAdmin } from '../rolePolicy.ts'
import type { WorkspacePage } from '../workspacePages'

export interface HelpTopic {
  id: WorkspacePage | 'start'
  title: string
  intro: string
  steps: readonly string[]
  note?: string
}

const topics: readonly HelpTopic[] = [
  {
    id: 'start', title: 'Premiers pas',
    intro: 'Retrouvez ici les étapes essentielles pour utiliser votre espace RH.',
    steps: [
      'Lors de la création de votre compte, un e-mail vous communique vos informations de connexion et votre mot de passe initial.',
      'Connectez-vous avec votre adresse e-mail ou votre identifiant. Si vous utilisez un mot de passe temporaire, modifiez-le dans Paramètres.',
      'Utilisez le menu à gauche pour accéder aux fonctionnalités. Sur mobile, ouvrez le menu avec le bouton à trois traits.',
      'Le bouton « ? » en bas du menu permet de revenir à ce guide sans quitter votre travail en cours.',
    ],
    note: 'Sur un appareil partagé, ne cochez pas « Se souvenir de moi » et déconnectez-vous après utilisation.',
  },
  {
    id: 'leave', title: 'Congés payés',
    intro: 'Déposez une demande de congé et suivez sa décision.',
    steps: [
      'Cliquez sur « Nouvelle demande ». Choisissez les dates et les périodes : journée entière, matin ou après-midi, puis envoyez la demande.',
      'Retrouvez son état dans les onglets En attente, Approuvées, Refusées ou Annulées. Les compteurs sont cliquables pour filtrer la liste.',
      'Cliquez sur « Détails » pour consulter la demande et son historique.',
      'Une demande encore en attente peut être retirée directement, sans e-mail. Elle ne compte plus parmi les demandes en attente.',
      'Pour un congé déjà approuvé, demandez son annulation : celle-ci doit être approuvée par le responsable ou un administrateur.',
    ],
    note: 'Le responsable reçoit un e-mail pour une nouvelle demande ou une demande d’annulation. L’employé reçoit le résultat par e-mail. Sans responsable disponible, la demande reste accessible aux administrateurs.',
  },
  {
    id: 'documents', title: 'Documents',
    intro: 'Déposez vos factures et justificatifs, puis retrouvez-les par période.',
    steps: [
      'Choisissez le mois et l’année concernés en haut de la page.',
      'Cliquez sur « Déposer un document », sélectionnez sa catégorie et votre fichier, puis confirmez l’envoi. Formats acceptés : PDF, JPG et PNG, jusqu’à 10 Mo par fichier.',
      'Utilisez la recherche et le filtre de type pour retrouver un document. Basculez entre « À traiter » et « Archives » si nécessaire.',
      'Sur une ligne, les icônes permettent de prévisualiser ou télécharger le document. Les autres actions disponibles sont dans « … ».',
      'Pour télécharger plusieurs fichiers, cochez-les puis cliquez sur « Télécharger le ZIP ». La sélection est limitée à 100 fichiers et 100 Mo.',
    ],
  },
  {
    id: 'medical', title: 'Visites médicales',
    intro: 'Retrouvez vos dates de visite et les attestations correspondantes.',
    steps: [
      'Choisissez une année et, si nécessaire, un mois. Téléchargez les attestations filtrées avec l’icône du haut, ou toutes les attestations d’un salarié avec l’icône de sa ligne. Le ZIP contient les PDF originaux classés par année et mois ; chaque fichier porte le nom du salarié et la date de visite. Dans l’historique, cochez les visites pour exporter une sélection (100 fichiers / 100 Mo maximum).',
      'Cliquez sur Enregistrer une visite, saisissez la date réelle et ajoutez une attestation PDF de 10 Mo maximum si elle est disponible.',
      'Vous pouvez enregistrer la date seule et ajouter le PDF plus tard depuis l’historique.',
      'Cliquez sur votre nom ou l’icône Historique pour consulter les visites précédentes et ouvrir leurs attestations.',
      'Les responsables consultent les visites et PDF de leurs collaborateurs selon les droits actuels de Mon équipe, y compris les anciennes visites. Ils peuvent modifier uniquement leur propre dossier ; les administrateurs peuvent corriger tous les dossiers accessibles.',
      'Utilisez le crayon pour corriger la date ou remplacer le PDF. L’icône Historique conserve les anciennes versions. Les Admin peuvent archiver et restaurer ; seul le Super Admin peut supprimer définitivement une visite déjà archivée et ses attestations.',
    ],
  },
  {
    id: 'team', title: 'Mon équipe',
    intro: 'Consultez votre équipe et l’organisation de l’entreprise selon vos droits.',
    steps: [
      'L’onglet de liste présente les collègues de vos départements ou les personnes que vous gérez, selon votre profil.',
      'Utilisez la recherche pour retrouver une personne et l’onglet « Organigramme de l’entreprise » pour consulter la structure.',
      'Cliquez sur « Mon profil » pour voir votre suivi mensuel. Choisissez une année et un mois pour consulter les présences, documents, signatures et congés accessibles.',
      'Le profil affiche le total des jours de congé approuvés pour l’année choisie, dates futures incluses, et le détail du mois. Une demi-journée vaut 0,5 jour. Les anciennes durées non vérifiables sont signalées.',
    ],
    note: 'Les profils des autres personnes sont accessibles uniquement si vos droits le permettent. Pour corriger votre affectation, contactez un administrateur.',
  },
  {
    id: 'profile', title: 'Paramètres',
    intro: 'Gérez vos informations de connexion et consultez votre dossier associé.',
    steps: [
      'Dans « Paramètres du compte », modifiez votre identifiant ou votre adresse e-mail si nécessaire.',
      'Pour changer de mot de passe, saisissez le nouveau mot de passe dans les deux champs correspondants. Sinon, laissez ces champs vides.',
      'Saisissez votre mot de passe actuel, puis cliquez sur « Enregistrer » en haut à droite.',
      'L’onglet « Données du personnel » présente votre dossier associé. Les corrections de ces données sont gérées par les administrateurs.',
    ],
    note: 'L’identifiant de connexion n’est pas le nom affiché dans l’application. La mise à jour du compte peut renouveler la session et recharger votre espace.',
  },
  {
    id: 'dashboard', title: 'Tableau de bord',
    intro: 'Identifiez les tâches à traiter dans votre périmètre.',
    steps: [
      'Dans « À traiter », cliquez sur Congés à valider, Signatures à suivre, Archives à vérifier ou Dossiers à vérifier.',
      'Utilisez la recherche et la période pour retrouver une tâche. Cliquez sur la ligne pour ouvrir la demande ou le dossier correspondant.',
      'Les indications de dossier incomplet permettent de rejoindre la rubrique à compléter.',
      'Consultez la répartition des effectifs et développez les rubriques de contrats ou d’entretiens pour voir les informations accessibles.',
    ],
  },
  {
    id: 'attendance', title: 'Fiches de temps',
    intro: 'Préparez une fiche mensuelle, enregistrez les présences, puis lancez la signature.',
    steps: [
      'Choisissez le mois, l’année et la personne, puis cliquez sur « Ouvrir / créer sa fiche ». Si le contrat manque, le formulaire vous propose de le compléter selon vos droits.',
      'Vérifiez « Informations personnelles et contrat » et « Planning hebdomadaire » avant de renseigner le calendrier.',
      'Dans « Calendrier mensuel », cliquez ou glissez sur les demi-journées à modifier : M pour le matin, A pour l’après-midi.',
      'Choisissez le statut souhaité, puis cliquez sur « Appliquer ». Cliquez ensuite sur « Enregistrer » en bas à droite.',
      'Cliquez sur « Préparer la signature », vérifiez les informations et destinataires, puis confirmez l’envoi. Le salarié signe en premier, puis son responsable.',
      'Ouvrez « Signé » pour suivre les envois et consulter ou télécharger les PDF disponibles, regroupés par année et par mois.',
    ],
    note: 'Une fiche signée, en cours de signature ou archivée est protégée : son calendrier est en consultation uniquement. Télécharger une fiche ne déclenche pas son envoi pour signature.',
  },
  {
    id: 'personnel', title: 'Utilisateurs',
    intro: 'Enregistrez les personnes et gérez leurs dossiers.',
    steps: [
      'Cliquez sur « Enregistrer un salarié » et complétez les champs obligatoires, dont l’identité et la date d’entrée. Les situations particulières affichent leurs champs complémentaires.',
      'Pour créer aussi un accès à l’application, cochez la création du compte, renseignez l’e-mail et choisissez la fonction, les départements et le responsable éventuel. Les droits Admin sont séparés de la fonction Employé ou Responsable.',
      'Pour une personne extérieure à l’entreprise, utilisez l’option correspondante : seul le compte de connexion est créé, sans inscription au registre des salariés.',
      'Recherchez une personne dans le tableau et ouvrez son dossier pour modifier les informations ou les périodes d’emploi. Les corrections conservent un historique.',
      'Utilisez le filtre de dossier pour retrouver les archives. Les icônes en haut à droite du tableau donnent accès aux exports et à l’historique des instantanés.',
    ],
    note: 'La sortie effective et la fin prévue du contrat sont deux dates distinctes. Le numéro de sécurité sociale est chiffré et masqué par défaut ; son affichage est temporaire et contrôlé.',
  },
  {
    id: 'admin', title: 'Administration',
    intro: 'Gérez l’organisation, les droits et l’état des comptes.',
    steps: [
      'Choisissez Normal, Désactivés, Archivés, Connexion supprimée ou Tous, puis recherchez et ouvrez le compte concerné.',
      'Dans « Organisation », choisissez la fonction Employé ou Responsable, les départements et le responsable direct éventuel. Les droits système se règlent avec les cases Admin et, si autorisé, Super Admin.',
      'Pour ajuster les subordonnés, ouvrez « Personnes sous responsabilité », cliquez sur « Modifier la liste », cochez ou décochez les personnes, puis vérifiez et confirmez les affectations.',
      'Une personne retirée peut rester temporairement sans responsable. Les droits de gestion d’un département restent indépendants de ce lien direct.',
      'Enregistrez les modifications d’organisation avec le bouton en bas à droite. Le bouton près du nom permet de désactiver ou réactiver le compte, après confirmation.',
      'Utilisez « Gérer les départements » pour maintenir les départements. Le panneau de stockage indique la capacité utilisée sur le serveur.',
    ],
    note: 'La désactivation conserve l’historique. Les actions de suppression sont réservées au Super Admin et restent soumises aux contrôles du serveur.',
  },
]

export function helpTopicsFor(subject: { role_tags: readonly string[] }): HelpTopic[] {
  return topics.filter(topic => {
    if (topic.id === 'admin' || topic.id === 'personnel') return isAdmin(subject)
    if (topic.id === 'dashboard' || topic.id === 'attendance') return canManageTeam(subject)
    return true
  }).map(topic => topic.id === 'personnel' ? {
    ...topic, note: `${topic.note} ${isSuperAdmin(subject) ? 'Le Super Admin peut exporter en PDF ou en Excel.' : 'L’export Admin est disponible en PDF.'}`,
  } : topic)
}
