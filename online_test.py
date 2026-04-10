import streamlit as st
from datetime import date
import requests
from DocxGen import generer_docx_stagiaire
from ExcelGen import remplir_fiche_paie

# Secrets de streamlit
TOKEN = st.secrets["PRESENCE_TOKEN"]
API_URL = st.secrets["URL_PRESENCE"]
USERS = st.secrets["USERS"]

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

# LOGOUT
def logout():
    st.session_state.logged_in = False
    st.session_state.username = None
    st.rerun()

# LOGIN
def login_page():
    st.title("Connexion")

    username = st.text_input("Nom d'utilisateur")
    password = st.text_input("Mot de passe", type="password")

    if st.button("Se connecter"):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success("Connexion réussie")
            st.rerun()
        else:
            st.error("Identifiants incorrects")
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

# Initialisation des sous-structures si vides
if "user_data" not in st.session_state:
    st.session_state.user_data = {}
if username not in st.session_state.user_data:
    st.session_state.user_data[username] = {}
user_store = st.session_state.user_data[username]

# Sélection du mois et de l'année
col1, col2 = st.columns(2)
with col1:
    user_store["mois"] = st.number_input("Mois", min_value=1, max_value=12, value=user_store.get("mois", 1), key="mois")
with col2:
    user_store["annee"] = st.number_input("Année", min_value=2000, max_value=2100, value=user_store.get("annee", 2025), key="annee")

# Gestion dynamique de la liste des employés
user_store["nb_employe"] = st.number_input("Nombre d'employés :", min_value=1, max_value=30, step=1, value=user_store.get("nb_employe", 1), key="nb_employe")
employe = [f"Employé {j+1}" for j in range(user_store["nb_employe"])]

if "employes_data" not in user_store:
    user_store["employes_data"] = []

# Ajouter des employés si on augmente le nombre
while len(user_store["employes_data"]) < user_store["nb_employe"]:
    user_store["employes_data"].append({
        "nom": "",
        "responsable": "",
        "ddc": None,
        "fdc": None,
        "cdi": False,
        "vacances": [],
        "absences": [],
        "arret": [],
        "planning": {d: True for d in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]}
    })
# Supprimer des employés si on diminue le nombre
while len(user_store["employes_data"]) > user_store["nb_employe"]:
    user_store["employes_data"].pop()

# Création des onglets pour chaque employé
tabs = st.tabs(employe)

for h, tab in enumerate(tabs):
    with tab:
        emp = user_store["employes_data"][h]
        
        # Sélection du type de contrat
        type_contrat = st.radio(
            "Type de contrat",
            ["Salarié", "Stagiaire"],
            key=f"type_contrat_{h}",
            index=0 if emp.get("type", "Salarié") == "Salarié" else 1,
            horizontal=True
        )
        emp["type"] = type_contrat

        st.divider()

        # CHAMPS SALARIÉS
        if type_contrat == "Salarié":
            st.subheader("Information Employé")
            emp["nom"] = st.text_input("NOM Prénom (Employé)", key=f"{username}_employe_nom_{h}", value=emp["nom"])
            emp["responsable"] = st.text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{h}", value=emp["responsable"])
            c1, c2 = st.columns(2)
            with c1:
                emp["ddc"] = st.date_input("Début de contrat", key=f"ddc_{h}", value=emp.get("ddc"))
                emp["cdi"] = st.checkbox("Contrat CDI ?", value=emp.get("cdi", False), key=f"cdi_{h}")
            with c2:
                if not emp["cdi"]:
                    emp["fdc"] = st.date_input("Fin de contrat", key=f"fdc_{h}", value=emp.get("fdc") if emp.get("fdc") != "Pas de fin" else None)
                else:
                    emp["fdc"] = "Pas de fin"
                    st.write("Fin de contrat : N/A")
            
            # SECTION PLANNINGS ET CONGES
            # Note : Le code utilise des boucles 'while' pour synchroniser le nombre de jours saisis avec le contenu du dictionnaire 'user_store'.

            # Section Planning pour les temps partiels
            with st.expander("Temps partiel / Planning hebdomadaire"):
                st.write("Cochez les jours travaillés :")
                cols_days = st.columns(7)
                jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
                for i, jour in enumerate(jours):
                    emp["planning"][jour] = cols_days[i].checkbox(jour[:3], value=emp["planning"].get(jour, True), key=f"plan_{h}_{jour}")

            # Section Congés
            with st.expander("Congés payés"):
                st.subheader("Saisir les jours de congés payés")
                nb_jours_vac = st.number_input("Nombre de jours :", min_value=0, max_value=31, value=len(emp["vacances"]), key=f"{username}_nb_jours_vac_{h}")

                while len(emp["vacances"]) < nb_jours_vac:
                    emp["vacances"].append({
                    "date": None,
                    "matin": False,
                    "aprem": False
                })

                while len(emp["vacances"]) > nb_jours_vac:
                    emp["vacances"].pop()

                for i, vac in enumerate(emp["vacances"]):
                    st.markdown(f"### Jour de CP #{i+1}")
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        vac["date"] = st.date_input(f"Date", key=f"{username}_date_cp_{h}_{i}", format="MM/DD/YYYY", value=vac["date"])
                    with col2:
                        vac["matin"] = st.checkbox(f"Matin", value=vac["matin"], key=f"{username}_matin_{h}_{i}")
                    with col3:
                        vac["aprem"] = st.checkbox(f"Après-midi", value=vac["aprem"], key=f"{username}_aprem_{h}_{i}")

            # Section Absences
            with st.expander("Absences"):
                st.subheader("Saisir les jours d'absences")
                nb_jours_abs = st.number_input("Nombre de jours :", min_value=0, max_value=31, value=len(emp["absences"]), key=f"{username}_nb_jours_abs_{h}")

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
                        abs["date"] = st.date_input(f"Date", key=f"{username}_date_abs_{h}_{i}", format="MM/DD/YYYY", value=abs["date"])
                    with col2:
                        abs["matin"] = st.checkbox(f"Matin", value=abs["matin"], key=f"{username}_matin_abs_{h}_{i}")
                    with col3:
                        abs["aprem"] = st.checkbox(f"Après-midi", value=abs["aprem"], key=f"{username}_aprem_abs_{h}_{i}")

            # Section Arrêts
            with st.expander("Arrêts maladies"):
                st.subheader("Saisir les jours d'arrêts maladies")
                nb_jours_am = st.number_input("Nombre de jours", min_value=0, max_value=31, value=len(emp["arret"]), key=f"{username}_nb_jours_am_{h}")

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
                        am["date"] = st.date_input(f"Date", key=f"{username}_date_am_{h}_{i}", format="MM/DD/YYYY", value=am["date"])
                    with col2:
                        am["matin"] = st.checkbox(f"Matin", value=am["matin"], key=f"{username}_matin_am_{h}_{i}")
                    with col3:
                        am["aprem"] = st.checkbox(f"Après-midi", value=am["aprem"], key=f"{username}_aprem_am_{h}_{i}")
        else:
            # CHAMPS STAGIAIRES
            st.subheader("Information Stagiaire")
        
            c1, c2 = st.columns(2)
            with c1:
                emp["nom_stagiaire"] = st.text_input("Nom du stagiaire", key=f"st_nom_{h}", value=emp.get("nom_stagiaire", ""))
                emp["dds"] = st.date_input("Début de stage", key=f"dds_{h}", value=emp.get("dds"))
                emp["nb_jours"] = st.number_input("Nombre de jours", key=f"st_nj_{h}", value=emp.get("nb_jours", 0))
                emp["taux_horaire"] = st.number_input("Taux horaire (€)", key=f"st_th_{h}", value=emp.get("taux_horaire", 0.0))
                emp["facture_mensuelle"] = st.number_input("Facture mensuelle (€)", key=f"st_fm_{h}", value=emp.get("facture_mensuelle", 0.0))
            with c2:
                emp["prenom_stagiaire"] = st.text_input("Prénom du stagiaire", key=f"st_pre_{h}", value=emp.get("prenom_stagiaire", ""))
                emp["fds"] = st.date_input("Fin de stage", key=f"fds_{h}", value=emp.get("dds"))
                emp["nb_heures_jour"] = st.number_input("Nombre d'heures/jour", key=f"st_nhj_{h}", value=emp.get("nb_heures_jour", 0.0))
                emp["transport"] = st.text_input("Transport", key=f"st_tr_{h}", value=emp.get("transport", ""))
                emp["taux"] = st.number_input("Taux (%)", key=f"st_tx_{h}", value=emp.get("taux", 0.0))

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
if st.button("Générer la fiche", type="primary"): 
    # On sépare les deux types de contrat
    salaries = [e for e in user_store["employes_data"] if e.get("type") == "Salarié"]
    stagiaires = [e for e in user_store["employes_data"] if e.get("type") == "Stagiaire"]

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
                break

        if (erreur_type == "le congé payé" or erreur_type == "l'absence" or erreur_type == "l'arrêt maladie"):
            st.error(
                f"Une des deux cases 'Matin' ou 'Après-midi' pour {erreur_type} de **{erreur_employe}** n'a pas été cochée !"
            )
        elif erreur_type:
            st.error(
                f"Il manque l'information {erreur_type} pour **{erreur_employe}** !"
            )
        else:
            # On génère un document pour tous les salariés
            buffer = remplir_fiche_paie(user_store["mois"], user_store["annee"], salaries)

            st.success("Fiche salariés générée avec succès !")

            st.download_button(
                "Télécharger la fiche remplie",
                data=buffer,
                file_name=f"fiche_paie_{user_store["mois"]}_{user_store["annee"]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    if stagiaires:
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
                erreur_type_stage = "du nom"
                erreur_stagiaire = f"Stagiaire {idx}"

            if erreur_type_stage:
                st.error(
                    f"Il manque l'information {erreur_type_stage} pour **{erreur_stagiaire}** !"
                )
            else:
                # On génère un document par stagiaire            
                buffer_docx = generer_docx_stagiaire(stagiaire, user_store['mois'], user_store['annee'])

                st.success("Fiches stagiaires générées avec succès !")
                
                st.download_button(
                    label=f"Télécharger la fiche d'indemnité de stage de {stagiaire.get('nom_stagiaire')}",
                    data=buffer_docx,
                    file_name=f"Fiche_Stage_{stagiaire.get('nom_stagiaire')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"btn_st_{idx}"
                )