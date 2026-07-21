import streamlit as st
from datetime import date, datetime
import requests
import time
from DocxGen import generer_docx_stagiaire
from ExcelGen import remplir_fiche_paie
import zipfile
import io

st.set_page_config(page_title="Générateur automatique de fiche de présence", layout="wide")

# Secrets de streamlit
TOKEN = st.secrets["PRESENCE_TOKEN"]
API_URL = st.secrets["URL_PRESENCE"]
HORAIRES = [f"{h:02d}:{m:02d}" for h in range(7, 21) for m in (0, 30)]
HORAIRES.insert(0, "") # Option vide pour les jours non travaillés

# Configuration du header pour les requêtes
headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

# Les dates sont transformées en chaînes de caractères (ISO format).
def serialize_dates(data):
    """Convertir date en string"""
    if isinstance(data, dict):
        return {k: serialize_dates(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [serialize_dates(v) for v in data]
    elif isinstance(data, date):
        return data.isoformat()
    return data

# Les chaînes sont retransformées en objets datetime.date pour être compatibles avec les widgets Streamlit.
def deserialize_dates(data):
    """Convertir string en date"""
    if isinstance(data, dict):
        return {k: deserialize_dates(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [deserialize_dates(v) for v in data]
    elif isinstance(data, str):
        try:
            return date.fromisoformat(data)
        except ValueError:
            return data
    return data

# INITIALISATION SESSION
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

# LOGOUT
def logout():
    st.session_state.logged_in = False
    st.session_state.username = None
    st.rerun()

# LOGIN
def login_page():
    # Récupération des paramètres de l'URL
    params = st.query_params

    if "token" in params:
        # --- MODE RÉINITIALISATION (Si l'URL contient un jeton) ---
        st.title(" Nouveau mot de passe")
        token = params["token"]
        new_pass = st.text_input("Saisissez votre nouveau mot de passe", type="password", width=400)
        confirm_pass = st.text_input("Confirmez votre nouveau mot de passe", type="password", width=400)
        if st.button("Valider le nouveau mot de passe", type="primary"):
            if new_pass == confirm_pass:
                res = requests.post(f"{API_URL}/reset-password", json={"token": token, "new_password": new_pass})
                if res.status_code == 200:
                    st.success("Votre mot de passe a bien été modifié ! Vous pouvez maintenant vous connecter.")
                    # On vide les paramètres URL pour revenir à l'écran de connexion normal
                    st.query_params.clear()
                    st.button("Retour à la connexion")
                else:
                    st.error("Le lien est invalide ou a expiré.")
            else:
                st.warning("Les mots de passe ne correspondent pas.")
    else:
        st.title("Connexion")

        username = st.text_input("Nom d'utilisateur", width=400)
        password = st.text_input("Mot de passe", type="password", width=400)

        if st.button("Se connecter", type="primary"):
            try:
                # On demande à l'API de vérifier
                res = requests.post(f"{API_URL}/login", json={"username": username, "password": password})
                if res.status_code == 200:
                    data = res.json()
                    st.session_state.logged_in = True
                    st.session_state.username = data["username"]
                    st.session_state.is_admin = data["is_admin"] # On stocke s'il est admin
                    st.success("Connexion réussie")
                    st.rerun()
                else:
                    st.error("Identifiants incorrects")
            except Exception as e:
                st.error(f"Erreur de connexion à l'API : {e}")
        
        # Toggle pour afficher le formulaire de mot de passe oublié
        forgot_tab = st.checkbox("Mot de passe oublié ?")

        if forgot_tab:
            st.subheader("Récupération de compte")
            email_recup = st.text_input("Entrez votre e-mail professionnel", width=400)
            if st.button("Recevoir le lien de récupération"):
                if email_recup:
                    res = requests.post(f"{API_URL}/forgot-password", json={"email": email_recup})
                    st.info("Si l'adresse existe, un lien vient de vous être envoyé par e-mail.")
                else:
                    st.warning("Veuillez entrer une adresse e-mail.")
            
############### Interface Streamlit #################

# Vérification de l'accès
if not st.session_state.logged_in:
    login_page()
    st.stop()  # bloque le reste de l'app

st.button("Déconnexion", on_click=logout)
username = st.session_state.username

# CHARGEMENT DES DONNEES DU VPS (MariaDB)
if "data_loaded" not in st.session_state:
    try:
        # Appel GET à l'API pour récupérer le JSON stocké
        response = requests.get(f"{API_URL}/get-config/{username}", headers=headers)
        if response.status_code == 200 and response.json():
            # On récupère les données et on convertit les strings en dates
            raw_data = response.json()
            st.session_state.user_data = {username: deserialize_dates(raw_data)}
        else:
            st.session_state.user_data = {username: {}}
        st.session_state.data_loaded = True
    except Exception as e:
        st.error(f"Erreur de connexion au serveur : {e}")
        st.session_state.user_data = {username: {}}

# Raccourci vers les données de l'utilisateur actuel
user_store = st.session_state.user_data[username]

# FORMULAIRE PRINCIPAL 
st.title("Générateur automatique de fiche de présence")
st.write(f"Bienvenue {username} !")

if st.session_state.is_admin:
    with st.sidebar.expander("PANNEAU ADMINISTRATEUR", expanded=False):
        st.subheader("Créer un nouvel utilisateur")
        new_user = st.text_input("Nom d'utilisateur", key="admin_new_user")
        new_pass = st.text_input("Mot de passe", type="password", key="admin_new_pass")
        new_is_admin = st.checkbox("Administrateur ?", value=False, help="Cochez pour que l'utilisateur soit administrateur")
            
        if st.button("Créer l'utilisateur", type="primary"):
            if new_user and new_pass:
                payload = {"new_username": new_user, "new_password": new_pass, "is_admin": new_is_admin}
                res = requests.post(f"{API_URL}/create-user", headers=headers, json=payload)
                    
                if res.status_code == 200:
                    st.success(f"Compte '{new_user}' créé avec succès !")
                else:
                    st.error(f"Erreur : {res.json().get('detail')}")
            else:
                st.warning("Veuillez remplir tous les champs.")

        st.write("---") # Ligne de séparation

        st.subheader("Liste des utilisateurs")
        
        try:
            # Appel à l'API pour récupérer la liste
            res_list = requests.get(f"{API_URL}/list-users", headers=headers)
            
            if res_list.status_code == 200:
                donnees_users = res_list.json().get("users", [])
                
                if donnees_users:
                    # Pour un affichage ultra propre, on peut utiliser un petit tableau
                    # ou boucler pour afficher des lignes avec des icônes
                    for u in donnees_users:
                        role = "Admin" if u["is_admin"] else "Utilisateur"
                        # Affichage d'une ligne stylisée pour chaque utilisateur
                        st.write(f"**{u['username']}** — *{role}*")
                else:
                    st.info("Aucun utilisateur trouvé en base.")
            else:
                st.error("Impossible de récupérer la liste des utilisateurs.")
        except Exception as e:
            st.error(f"Erreur de connexion à l'API : {e}")

        st.write("---") # Ligne de séparation
        
        st.subheader("Supprimer un utilisateur")
        user_to_del = st.text_input("Nom de l'utilisateur à supprimer", key="admin_del_user")
        
        # Sécurité Streamlit : On demande de cocher une case pour confirmer avant de cliquer
        confirmer_suppression = st.checkbox("Je confirme vouloir supprimer définitivement cet utilisateur", key="confirm_del")
        
        if st.button("Supprimer l'utilisateur", type="secondary", disabled=not confirmer_suppression):
            if user_to_del:
                # Sécurité : Éviter que l'admin connecté se supprime lui-même
                if user_to_del == st.session_state.username:
                    st.error("Vous ne pouvez pas supprimer le compte avec lequel vous êtes actuellement connecté !")
                else:
                    try:
                        # Appel à l'API avec la méthode DELETE
                        # Note : On passe le nom dans l'URL directement comme défini dans FastAPI
                        res = requests.delete(
                            f"{API_URL}/delete-user/{user_to_del}", 
                            headers=headers
                        )
                        
                        if res.status_code == 200:
                            st.success(f"Le compte '{user_to_del}' a été supprimé.")
                            # Petit rerun pour rafraîchir l'interface si nécessaire
                            st.rerun()
                        else:
                            st.error(f"Erreur : {res.json().get('detail', 'Impossible de supprimer cet utilisateur')}")
                    except Exception as e:
                        st.error(f"Erreur de communication avec l'API : {e}")
            else:
                st.warning("Veuillez saisir un nom d'utilisateur.")

# Initialisation des sous-structures si vides
if "user_data" not in st.session_state:
    st.session_state.user_data = {}
if username not in st.session_state.user_data:
    st.session_state.user_data[username] = {}
user_store = st.session_state.user_data[username]

# Sélection du mois et de l'année
now = datetime.now()
col1, col2 = st.columns(2)
with col1:
    user_store["mois"] = st.number_input("Mois", min_value=1, max_value=12, value=int(now.strftime("%m")), key="mois", help="Saisissez le numéro du mois")
with col2:
    user_store["annee"] = st.number_input("Année", min_value=2000, max_value=2100, value=int(now.strftime("%Y")), key="annee")

if "employes_data" not in user_store:
    user_store["employes_data"] = []

# Bouton pour ajouter un employé à la fin de la liste
if st.button("Ajouter un employé / stagiaire", use_container_width=True):
    user_store["employes_data"].append({
        "id": int(time.time() * 1000),
        "type": "Salarié",
        "nom": "", "responsable": "", "email_responsable": "", "ddc": None, "fdc": None, "cdi": False,
        "vacances": [], "absences": [], "arret": [],
        "planning_detail": {j: {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": True} for j in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]}
    })
    st.rerun() # On force Streamlit à recréer les onglets immédiatement

# Si la liste est vide, on affiche un message d'aide
if not user_store["employes_data"]:
    st.info("Aucun employé ou stagiaire configuré. Cliquez sur le bouton ci-dessus pour commencer.")

if user_store["employes_data"]:
    # Génération dynamique des titres des onglets (affiche le nom de l'employé s'il existe)
    labels_onglets = [
        f"Employé {idx+1}" for idx in range(len(user_store["employes_data"]))
    ]
    
    # Création des onglets pour chaque employé
    tabs = st.tabs(labels_onglets)

    for h, tab in enumerate(tabs):
        with tab:
            emp = user_store["employes_data"][h]
            
            # Si un employé n'a pas d'ID, on lui en donne un
            if "id" not in emp:
                emp["id"] = int(time.time() * 1000) + h

            emp_id = emp["id"]

            # --- BOUTON DE SUPPRESSION DE CE TAB PRÉCIS ---
            c_space, c_gen, c_del = st.columns([4, 1, 1])
            with c_space:
                st.subheader(f"Fiche de {emp['nom']}" if emp["nom"] else f"Fiche d'employé")
            with c_del:
                # Un bouton rouge aligné à droite pour supprimer l'employé courant
                if st.button("Supprimer cette fiche", key=f"del_btn_{emp_id}", type="secondary", help="Supprime définitivement cet employé de la liste"):
                    user_store["employes_data"].pop(h) # Supprime précisément l'index h
                    st.success("Fiche supprimée ! Sauvegardez pour appliquer les changements sur le serveur.")
                    st.rerun() # Recharge l'interface sans l'onglet supprimé
            with c_gen:
                if st.button("Générer cette fiche", key=f"gen_solo_btn_{emp_id}", type="primary", help="Charge uniquement la fiche de cet employé"):
                    erreur_type_solo = None
                    nom_propre = emp.get("nom", f"Fiche_{h+1}").replace(" ", "_")
                    
                    if emp.get("type") == "Salarié":
                        # Validation Salarié
                        if not emp.get("fdc"): erreur_type_solo = "du fin de contrat"
                        if not emp.get("ddc"): erreur_type_solo = "du début de contrat"
                        if emp.get("responsable") == "": erreur_type_solo = "du responsable"
                        if emp.get("nom") == "": erreur_type_solo = "du nom"
                        
                        if erreur_type_solo:
                            st.error(f"Impossible de générer : il manque l'information {erreur_type_solo} !")
                        else:
                            # Génération de l'Excel unique
                            excel_buffer = remplir_fiche_paie(user_store["mois"], user_store["annee"], emp)
                            
                            # On propose le téléchargement immédiat de cet Excel
                            st.download_button(
                                label="Télécharger l'Excel",
                                data=excel_buffer,
                                file_name=f"fiche_paie_{nom_propre}_{user_store['mois']}_{user_store['annee']}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_solo_excel_{emp_id}"
                            )
                            
                    else:
                        # Validation Stagiaire
                        if not emp.get("fds"): erreur_type_solo = "de la fin de stage"
                        if not emp.get("dds"): erreur_type_solo = "du début de stage"
                        if emp.get("nom_stagiaire") == "": erreur_type_solo = "du nom"
                        
                        if erreur_type_solo:
                            st.error(f"Impossible de générer : il manque l'information {erreur_type_solo} !")
                        else:
                            # Génération du Word unique
                            docx_buffer = generer_docx_stagiaire(emp, user_store['mois'], user_store['annee'])
                            
                            # On propose le téléchargement immédiat de ce Word
                            st.download_button(
                                label="Télécharger le Word",
                                data=docx_buffer,
                                file_name=f"Fiche_Stage_{nom_propre}_{user_store['mois']}_{user_store['annee']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_solo_docx_{emp_id}"
                            )

            # Sélection du type de contrat
            type_contrat = st.radio(
                "Type de contrat",
                ["Salarié", "Stagiaire"],
                key=f"type_contrat_{emp_id}",
                index=0 if emp.get("type", "Salarié") == "Salarié" else 1,
                horizontal=True
            )
            emp["type"] = type_contrat

            st.divider()

            # CHAMPS SALARIÉS
            if type_contrat == "Salarié":
                st.subheader("Informations Employé")
                emp["nom"] = st.text_input("NOM Prénom (Employé)", key=f"{username}_employe_nom_{emp_id}", value=emp["nom"])
                emp["responsable"] = st.text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{emp_id}", value=emp["responsable"])
                emp["email_responsable"] = st.text_input("Email du responsable", placeholder="responsable@univ-ilci.fr", key=f"{username}_resp_mail_{emp_id}", value=emp.get("email_responsable", ""))
                c1, c2 = st.columns(2)
                with c1:
                    emp["ddc"] = st.date_input("Début de contrat", key=f"ddc_{emp_id}", value=emp.get("ddc"), format="DD/MM/YYYY")
                    emp["cdi"] = st.checkbox("Contrat CDI ?", value=emp.get("cdi", False), key=f"cdi_{emp_id}")
                with c2:
                    if not emp["cdi"]:
                        emp["fdc"] = st.date_input("Fin de contrat", key=f"fdc_{emp_id}", value=emp.get("fdc") if emp.get("fdc") != "Pas de fin" else None, format="DD/MM/YYYY")
                    else:
                        emp["fdc"] = "Pas de fin"
                        st.write("Fin de contrat : N/A")
                
                # SECTION PLANNINGS ET CONGES
                # Note : Le code utilise des boucles 'while' pour synchroniser le nombre de jours saisis avec le contenu du dictionnaire 'user_store'.

                # Section Planning pour les temps partiels
                with st.expander("Temps partiel / Planning hebdomadaire"):
                    st.write("Indiquez les horaires pour chaque jour (décochez si non travaillé) :")
                    
                    # On définit les jours de la semaine
                    jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
                    
                    # On initialise la structure si besoin
                    if "planning_detail" not in emp:
                        emp["planning_detail"] = {j: {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": True} for j in jours}

                    for jour in jours:
                        st.write(f"**{jour}**")
                        c1, c2, c3, c4, c5 = st.columns([1, 2, 2, 2, 2])
                        
                        with c1:
                            emp["planning_detail"][jour]["actif"] = st.checkbox("Jour de travail", value=emp["planning_detail"][jour]["actif"], key=f"check_{emp_id}_{jour}")
                        
                        if emp["planning_detail"][jour]["actif"]:
                            with c2:
                                emp["planning_detail"][jour]["m1"] = st.selectbox("Matin de", HORAIRES, index=HORAIRES.index(emp["planning_detail"][jour]["m1"]), key=f"m1_{emp_id}_{jour}")
                            with c3:
                                emp["planning_detail"][jour]["m2"] = st.selectbox("à", HORAIRES, index=HORAIRES.index(emp["planning_detail"][jour]["m2"]), key=f"m2_{emp_id}_{jour}")
                            with c4:
                                emp["planning_detail"][jour]["a1"] = st.selectbox("Après-midi de", HORAIRES, index=HORAIRES.index(emp["planning_detail"][jour]["a1"]), key=f"a1_{emp_id}_{jour}")
                            with c5:
                                emp["planning_detail"][jour]["a2"] = st.selectbox("à", HORAIRES, index=HORAIRES.index(emp["planning_detail"][jour]["a2"]), key=f"a2_{emp_id}_{jour}")

                # Section Congés
                with st.expander("Congés payés"):
                    st.subheader("Saisir les jours de congés payés")
                    nb_jours_vac = st.number_input("Nombre de jours :", min_value=0, max_value=31, value=len(emp["vacances"]), key=f"{username}_nb_jours_vac_{emp_id}")

                    while len(emp["vacances"]) < nb_jours_vac:
                        emp["vacances"].append({
                        "date": None,
                        "matin": False,
                        "aprem": False,
                        "examen_alt": False
                    })

                    while len(emp["vacances"]) > nb_jours_vac:
                        emp["vacances"].pop()

                    for i, vac in enumerate(emp["vacances"]):
                        st.markdown(f"### Jour de CP #{i+1}")
                        col1, col2, col3, col4 = st.columns(4)

                        with col1:
                            vac["date"] = st.date_input(f"Date", key=f"{username}_date_cp_{emp_id}_{i}", format="MM/DD/YYYY", value=vac["date"])
                        with col2:
                            vac["matin"] = st.checkbox(f"Matin", value=vac["matin"], key=f"{username}_matin_{emp_id}_{i}")
                        with col3:
                            vac["aprem"] = st.checkbox(f"Après-midi", value=vac["aprem"], key=f"{username}_aprem_{emp_id}_{i}")
                        with col4:
                            vac["examen_alt"] = st.checkbox(f"Examen alternance", value=vac["examen_alt"], key=f"{username}_examen_alt_{emp_id}_{i}", help="Cochez la case si c'est un alternant qui pose des jours de congés pour les révisions de ses examens")

                # Section Absences
                with st.expander("Absences"):
                    st.subheader("Saisir les jours d'absences")
                    nb_jours_abs = st.number_input("Nombre de jours :", min_value=0, max_value=31, value=len(emp["absences"]), key=f"{username}_nb_jours_abs_{emp_id}")

                    while len(emp["absences"]) < nb_jours_abs:
                        emp["absences"].append({
                        "date": None,
                        "matin": False,
                        "aprem": False
                    })

                    while len(emp["absences"]) > nb_jours_abs:
                        emp["absences"].pop()

                    for i, abs in enumerate(emp["absences"]):
                        st.markdown(f"### Jour d'ABS #{i+1}")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            abs["date"] = st.date_input(f"Date", key=f"{username}_date_abs_{emp_id}_{i}", format="MM/DD/YYYY", value=abs["date"])
                        with col2:
                            abs["matin"] = st.checkbox(f"Matin", value=abs["matin"], key=f"{username}_matin_abs_{emp_id}_{i}")
                        with col3:
                            abs["aprem"] = st.checkbox(f"Après-midi", value=abs["aprem"], key=f"{username}_aprem_abs_{emp_id}_{i}")

                # Section Arrêts
                with st.expander("Arrêts maladies"):
                    st.subheader("Saisir les jours d'arrêts maladies")
                    nb_jours_am = st.number_input("Nombre de jours", min_value=0, max_value=31, value=len(emp["arret"]), key=f"{username}_nb_jours_am_{emp_id}")

                    while len(emp["arret"]) < nb_jours_am:
                        emp["arret"].append({
                        "date": None,
                        "matin": False,
                        "aprem": False
                    })

                    while len(emp["arret"]) > nb_jours_am:
                        emp["arret"].pop()

                    for i, am in enumerate(emp["arret"]):
                        st.markdown(f"### Jour d'AM #{i+1}")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            am["date"] = st.date_input(f"Date", key=f"{username}_date_am_{emp_id}_{i}", format="MM/DD/YYYY", value=am["date"])
                        with col2:
                            am["matin"] = st.checkbox(f"Matin", value=am["matin"], key=f"{username}_matin_am_{emp_id}_{i}")
                        with col3:
                            am["aprem"] = st.checkbox(f"Après-midi", value=am["aprem"], key=f"{username}_aprem_am_{emp_id}_{i}")
            else:
                # CHAMPS STAGIAIRES
                st.subheader("Information Stagiaire")
            
                c1, c2 = st.columns(2)
                with c1:
                    emp["nom_stagiaire"] = st.text_input("Nom du stagiaire", key=f"st_nom_{emp_id}", value=emp.get("nom_stagiaire", ""))
                    emp["responsable"] = st.text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{emp_id}", value=emp["responsable"])
                    emp["dds"] = st.date_input("Début de stage", key=f"dds_{emp_id}", value=emp.get("dds"), format="DD/MM/YYYY")
                    emp["nb_jours"] = st.number_input("Nombre de jours", key=f"st_nj_{emp_id}", value=emp.get("nb_jours", 0))
                    emp["taux_horaire"] = st.number_input("Taux horaire (€)", key=f"st_th_{emp_id}", value=emp.get("taux_horaire", 0.0))
                    emp["facture_mensuelle"] = st.number_input("Facture mensuelle (€)", key=f"st_fm_{emp_id}", value=emp.get("facture_mensuelle", 0.0))
                with c2:
                    emp["prenom_stagiaire"] = st.text_input("Prénom du stagiaire", key=f"st_pre_{emp_id}", value=emp.get("prenom_stagiaire", ""))
                    emp["email_responsable"] = st.text_input("Email du responsable", placeholder="responsable@univ-ilci.fr", key=f"{username}_resp_mail_{emp_id}", value=emp.get("email_responsable", ""))
                    emp["fds"] = st.date_input("Fin de stage", key=f"fds_{emp_id}", value=emp.get("fds"), format="DD/MM/YYYY")
                    emp["nb_heures_jour"] = st.number_input("Nombre d'heures/jour", key=f"st_nhj_{emp_id}", value=emp.get("nb_heures_jour", 0.0))
                    emp["transport"] = st.text_input("Transport", key=f"st_tr_{emp_id}", value=emp.get("transport", ""))
                    emp["taux"] = st.number_input("Taux (%)", key=f"st_tx_{emp_id}", value=emp.get("taux", 0.0))

# BOUTON DE SAUVEGARDE SUR LE VPS
st.divider()
if st.button("Sauvegarder", use_container_width=True):
    try:
        # On prépare les données (conversion des dates en texte)
        data_to_send = serialize_dates(user_store)
        
        response = requests.post(
            f"{API_URL}/save-config/{username}",
            headers=headers, 
            json=data_to_send
        )
        
        if response.status_code == 200:
            st.success("Données synchronisées avec succès !")
        else:
            st.error(f"Erreur lors de la sauvegarde: {response.status_code}")
    except Exception as e:
        st.error(f"Impossible de joindre le serveur : {e}")

# GÉNÉRATION EXCEL ET DOCX
if st.button("Générer toutes les fiches", type="primary"): 
    # On sépare les deux types de contrat
    salaries = [e for e in user_store["employes_data"] if e.get("type") == "Salarié"]
    stagiaires = [e for e in user_store["employes_data"] if e.get("type") == "Stagiaire"]

    bloquer_generation = False

    if salaries:
        # Logique de validation des champs obligatoires
        erreur_type = None
        erreur_employe = None
        categories = {
            "vacances": "le congé payé",
            "absences": "l'absence",
            "arret": "l'arrêt maladie"
        }

        for idx, employe in enumerate(salaries, start=1):
            nom_emp = employe.get("nom", "Employé sans nom")

            for key_cat, label in categories.items():
                for jour in employe[key_cat]:
                    if not jour["matin"] and not jour["aprem"]:
                        erreur_type = label
                        erreur_employe = nom_emp
                        break
                
                if erreur_type:
                    break
            
            if (not employe["fdc"]):
                erreur_type = "du fin de contrat"
                erreur_employe = f"{nom_emp} (Employé {idx})"

            if (not employe["ddc"]):
                erreur_type = "du début de contrat"
                erreur_employe = f"{nom_emp} (Employé {idx})"

            if (employe["responsable"] == ""):
                erreur_type = "du responsable"
                erreur_employe = f"{nom_emp} (Employé {idx})"

            if (employe["nom"] == ""):
                erreur_type = "du nom"
                erreur_employe = f"Employé {idx}"

            if erreur_type:
                bloquer_generation = True
                break

        if (erreur_type == "le congé payé" or erreur_type == "l'absence" or erreur_type == "l'arrêt maladie"):
            st.error(
                f"Une des deux cases 'Matin' ou 'Après-midi' pour {erreur_type} de **{erreur_employe}** n'a pas été cochée !"
            )
        elif erreur_type:
            st.error(
                f"Il manque l'information {erreur_type} pour **{erreur_employe}** !"
            )

    if stagiaires and not bloquer_generation:
        # Logique de validation des champs obligatoires
        erreur_type_stage = None
        erreur_stagiaire = None

        for idx, stagiaire in enumerate(stagiaires, start=1):
            nom_emp = stagiaire.get("nom", "Stagiaire sans nom")

            if (not stagiaire["fds"]):
                erreur_type_stage = "du fin de contrat"
                erreur_stagiaire = f"{nom_emp} (Stagiaire {idx})"

            if (not stagiaire["dds"]):
                erreur_type_stage = "du début de contrat"
                erreur_stagiaire = f"{nom_emp} (Stagiaire {idx})"

            if (stagiaire["nom_stagiaire"] == ""):
                erreur_type_stage = "du nom"
                erreur_stagiaire = f"Stagiaire {idx}"

            if (stagiaire["prenom_stagiaire"] == ""):
                erreur_type_stage = "du prénom"
                erreur_stagiaire = f"Stagiaire {idx}"

            if erreur_type_stage:
                bloquer_generation = True
                st.error(
                    f"Il manque l'information {erreur_type_stage} pour **{erreur_stagiaire}** !"
                )
    
    if not bloquer_generation and (salaries or stagiaires):
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

            if salaries:
                for idx, salarie in enumerate(salaries):
                    nom_propre = salarie.get("nom", f"Employe_{idx+1}").replace(" ", "_")
                    file_name = f"Fiche_paie_{nom_propre}_{user_store['mois']}_{user_store['annee']}.xlsx"

                    excel_buffer = remplir_fiche_paie(user_store["mois"], user_store["annee"],salarie)

                    zip_file.writestr(file_name, excel_buffer.getvalue())
            
            if stagiaires:
                for idx, stagiaire in enumerate(stagiaires):
                    nom_propre = stagiaire.get("nom_stagiaire", f"Stagiaire_{idx+1}").replace(" ", "_")
                    file_name = f"Fiche_stage_{nom_propre}_{user_store['mois']}_{user_store['annee']}.docx"

                    docx_buffer = generer_docx_stagiaire(stagiaire, user_store['mois'], user_store['annee'])

                    zip_file.writestr(file_name, docx_buffer.getvalue())

        zip_buffer.seek(0)

        st.success("Toutes les fiches individuelles ont été générées avec succès !")

        st.download_button(
            label="Télécharger toutes les fiches (Dossier ZIP)",
            data=zip_buffer,
            file_name=f"fiches_presence_{user_store['mois']}_{user_store['annee']}.zip",
            mime="application/zip",
            use_container_width=True
        )