# Journal de développement

## 2026-09-11 — Publication du calendrier de présence

- Publication du calendrier mensuel visuel et du mode de modification groupée sur la branche `main` du dépôt GitHub.
- Déploiement du commit `11f4ef6` sur l'interface Presence de production, dans `/var/www/presence-interface`.
- Création d'une sauvegarde du code applicatif sur le VPS avant la mise à jour, puis redémarrage contrôlé de `presence-interface.service`.
- Vérification du service actif et de la réponse HTTP `200` de l'URL publique après redémarrage.
- Aucune commande MariaDB, suppression ou migration de données n'a été exécutée : les comptes et les fiches existantes restent inchangés.
- Correction de compatibilité de connexion : l'API historique peut renvoyer `role: null` avec `is_admin`. L'interface déduit désormais le rôle `Admin` ou `Responsable` dans ce cas, sans modifier le compte concerné.
- Correction du panneau d'administration : la liste des utilisateurs contenait une f-string avec des guillemets incompatibles, qui ne se manifestait qu'une fois la page Admin redevenue accessible.
- Alignement de l'environnement de production sur Streamlit `1.60.0`, version utilisée lors de la validation locale du composant calendrier. La version `1.58.0` affichait le composant de sélection sous forme d'espace réservé vide.

## 2026-09-10 — Calendrier mensuel visuel

- Ajout de `calendar_view.py`, un composant Streamlit de calendrier mensuel.
- Affichage par semaines avec deux demi-journées par date (matin / après-midi).
- Modification directe par clic après sélection du type de saisie au-dessus du calendrier.
- Ajout des statuts `Travail`, `Congé payé`, `Absence`, `Arrêt maladie`, `Férié`, `Autre` et `Réinitialiser`.
- Les journées consécutives du même statut sont rendues sous forme de barres visuellement continues avec extrémités arrondies.
- Les employés et les stagiaires utilisent le même calendrier.
- Les anciennes données sont complétées uniquement avec les clés manquantes (`calendar_overrides` et planning incomplet) afin de préserver la compatibilité JSON/MariaDB.

### Vérifications

- Compilation Python de l'application et des générateurs : OK.
- Tests des fonctions de calendrier (statuts, demi-journées, réinitialisation) : OK.
- Test Streamlit de la page employé : calendrier et contrôles chargés sans exception.
- Test Streamlit de la page stagiaire : calendrier et champs stagiaire chargés sans exception.
- Prévisualisation locale démarrée sur `http://127.0.0.1:8502` avec une API de test isolée.
- Ajout de `test_support/local_presence_api.py` pour redémarrer cette API de test locale sans accès au VPS; les données restent uniquement en mémoire et repartent vides après son arrêt.
- Le flux de connexion conserve la navigation Streamlit existante (`st.rerun`) après authentification; le lien vers `Fiches` reste disponible dans le menu.
- Correction d’une tentative de navigation programmatique incompatible avec `st.navigation`; aucune modification de l’authentification n’est conservée.
- Test navigateur local : sélection d'un congé payé sur une demi-journée, saisie d'un motif `Autre` et sauvegarde simulée : OK.

## 2026-09-10 — Éditeur de calendrier intégré

- Suppression de l'aperçu clair et de la liste de boutons située sous le calendrier.
- Le calendrier utilise désormais un fond sombre cohérent avec l'application et des couleurs de statut à faible saturation.
- Chaque matin et après-midi est un contrôle direct dans sa cellule. Un clic ouvre un menu contextuel avec les statuts disponibles.
- Le statut `Autre` demande un motif dans ce même menu avant application.
- Le statut `Travail` utilise le planning par défaut de la personne; les horaires d'une demi-journée ne sont pas modifiables dans ce menu.
- Sur les fenêtres étroites, la grille conserve une largeur lisible et défile horizontalement dans sa propre zone plutôt que de casser les libellés en colonnes verticales.

### Vérifications

- Analyse syntaxique Python : OK.
- Test Streamlit du composant : chargement sans exception.
- Test navigateur local : ouverture d'un menu de demi-journée et passage de `Travail` à `Congé payé` : OK.

## 2026-09-10 — Synchronisation calendrier et export

- Correction d'un conflit entre le calendrier et les anciens compteurs manuels de congés, absences et arrêts maladies.
- Après une modification dans le calendrier, les anciens compteurs conservaient parfois leur ancienne valeur et retiraient la nouvelle exception au redémarrage Streamlit.
- Les clés des compteurs sont désormais synchronisées avec le nombre courant d'exceptions; une absence, un congé payé ou un arrêt ajouté depuis le calendrier reste présent jusqu'à modification explicite.

### Vérifications

- Test Streamlit : un congé payé ajouté depuis une demi-journée du calendrier survit au redémarrage déclenché par le contrôle manuel : OK.
- Test Excel : congé payé, absence et arrêt maladie ajoutés depuis le calendrier sont affichés dans la fiche exportée : OK.

## 2026-09-10 — Modification groupée du calendrier

- Ajout de deux modes : consultation par défaut et modification explicite via le bouton `Modifier`.
- Le mode de consultation affiche uniquement les barres de statut continues, avec des extrémités arrondies pour les séquences consécutives.
- Le mode de modification propose une case indépendante pour chaque matin et après-midi.
- Une sélection de plusieurs demi-journées ouvre un seul menu d'action pour appliquer un statut commun, puis la sélection est automatiquement effacée.
- Ajout de `Terminer` pour quitter le mode de modification et de `Annuler la sélection` pour annuler un lot en attente.

## 2026-09-10 — Sélection par glisser-déposer

- Remplacement des cases Streamlit du mode de modification par un petit composant local `calendar_selector_component`.
- Les demi-journées conservent une taille fixe. Un clic les sélectionne ou les désélectionne; un glisser-déposer ajoute ou retire une série de demi-journées selon l'état de la première cellule.
- Les libellés longs sont tronqués dans le mode de modification au lieu d'augmenter la hauteur d'une carte; la vue de consultation conserve les horaires complets.
- Les sélections restent désormais dans le composant navigateur. Le calendrier ne transmet une action à Streamlit qu'au moment de l'application d'un statut, ce qui évite le rechargement visuel après chaque clic ou glisser-déposer.
- Augmentation uniforme de la hauteur des cartes de demi-journée afin d'afficher les horaires sur plusieurs lignes sans modifier la largeur des colonnes.
- Les nouveaux plannings définissent samedi et dimanche comme jours de repos par défaut; les plannings existants restent inchangés.
- La sélection groupée reste séparée du statut : elle n'altère aucune présence avant validation dans `Modifier les X demi-journées sélectionnées`.

### Vérifications

- Test navigateur local : le composant de sélection, les cellules à taille fixe, les contrôles de fin de modification et le compteur de sélection se chargent sans erreur.

### Vérifications

- Test Streamlit : sélection de deux demi-journées, application groupée d'un congé payé et effacement automatique de la sélection : OK.
