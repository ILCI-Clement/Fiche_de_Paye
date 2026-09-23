import streamlit as st
from datetime import date, datetime
import requests
import time
import base64
from DocxGen import generer_docx_stagiaire
from ExcelGen import remplir_fiche_paie
from calendar_view import render_monthly_calendar
import zipfile
import io
from api_client import api_url, authenticated_headers

# Secrets de streamlit
API_URL = api_url()
HORAIRES = [f"{h:02d}:{m:02d}" for h in range(7, 21) for m in (0, 30)]
HORAIRES.insert(0, "") # Option vide pour les jours non travaillés

# Configuration du header pour les requêtes
headers = authenticated_headers()

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

if "user" not in st.session_state or st.session_state["user"] is None:
    st.warning("Veuillez vous connecter d'abord.")
    st.stop()

username = st.session_state['user']['name']
role = st.session_state['user']['role']

if "user_data" not in st.session_state:
    st.session_state.user_data = {}

# CHARGEMENT DES DONNEES DU VPS (MariaDB)
if username not in st.session_state.user_data:
    try:
        # Appel GET à l'API pour récupérer le JSON stocké
        response = requests.get(f"{API_URL}/get-config/{username}", headers=headers)
        if response.status_code == 200 and response.json():
            # On récupère les données et on convertit les strings en dates
            raw_data = response.json()
            st.session_state.user_data[username] = deserialize_dates(raw_data)
        else:
            st.session_state.user_data[username] = {}
    except Exception as e:
        st.error(f"Erreur de connexion au serveur : {e}")
        st.session_state.user_data[username] = {}

# Raccourci vers les données de l'utilisateur actuel
user_store = st.session_state.user_data[username]

# FORMULAIRE PRINCIPAL 
st.title("Générateur de fiche de présence")

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
if st.button("Ajouter un employé / stagiaire", width="stretch"):
    user_store["employes_data"].append({
        "id": int(time.time() * 1000),
        "type": "Salarié",
        "nom": "", "email_employe": "", "responsable": "", "email_responsable": "", "ddc": None, "fdc": None, "cdi": False,
        "vacances": [], "absences": [], "arret": [],
        "calendar_overrides": {},
        "planning_detail": {j: {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": j not in ("Samedi", "Dimanche")} for j in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]}
    })
    st.rerun() # On force Streamlit à recréer les onglets immédiatement

# Si la liste est vide, on affiche un message d'aide
if not user_store["employes_data"]:
    st.info("Aucun employé ou stagiaire configuré. Cliquez sur le bouton ci-dessus pour commencer.")

if user_store["employes_data"]:
    def fiche_tab_label(employee: dict, index: int) -> str:
        """Return the clearest available label for an attendance sheet tab."""
        if employee.get("type") == "Stagiaire":
            name = " ".join(
                part.strip()
                for part in (str(employee.get("prenom_stagiaire") or ""), str(employee.get("nom_stagiaire") or ""))
                if part.strip()
            )
        else:
            name = str(employee.get("prenom") or employee.get("nom") or "").strip()
        return name or f"Employé {index + 1}"

    labels_onglets = [
        fiche_tab_label(employee, index) for index, employee in enumerate(user_store["employes_data"])
    ]
    
    # Création des onglets pour chaque employé
    # Les onglets suivis conservent l'employé actif après le rafraîchissement
    # déclenché par un champ, notamment les sélecteurs de date.
    active_tab = st.session_state.get("fiche_active_tab")
    if active_tab not in labels_onglets:
        active_tab = labels_onglets[0]
    tabs = st.tabs(labels_onglets, default=active_tab, key="fiche_active_tab", on_change="rerun")

    for h, tab in enumerate(tabs):
        if not tab.open:
            continue
        with tab:
            emp = user_store["employes_data"][h]
            
            # Si un employé n'a pas d'ID, on lui en donne un
            if "id" not in emp:
                emp["id"] = int(time.time() * 1000) + h

            emp_id = emp["id"]

            generated_file_key = f"generated_fiche_{emp_id}"

            c_space, c_gen, c_del = st.columns([4, 1, 1])
            with c_space:
                st.subheader(f"Fiche de {emp['nom']}" if emp["nom"] else f"Fiche d'employé")
            with c_del:
                if st.button("Supprimer cette fiche", key=f"del_btn_{emp_id}", type="secondary", help="Supprime définitivement cet employé de la liste"):
                    user_store["employes_data"].pop(h)
                    st.session_state.pop(generated_file_key, None)
                    st.success("Fiche supprimée ! Sauvegardez pour appliquer les changements sur le serveur.")
                    st.rerun()
            with c_gen:
                if st.button("Générer cette fiche", key=f"gen_solo_btn_{emp_id}", type="primary", help="Charge uniquement la fiche de cet employé"):
                    erreur_type_solo = None
                    nom_employe_text = emp.get("nom") or emp.get("nom_stagiaire") or f"Employé {h + 1}"
                    nom_propre = nom_employe_text.replace(" ", "_")

                    if emp.get("type") == "Salarié":
                        if not emp.get("fdc"): erreur_type_solo = "du fin de contrat"
                        if not emp.get("ddc"): erreur_type_solo = "du début de contrat"
                        if emp.get("responsable") == "": erreur_type_solo = "du responsable"
                        if emp.get("nom") == "": erreur_type_solo = "du nom"

                        if erreur_type_solo:
                            st.error(f"Impossible de générer : il manque l'information {erreur_type_solo} !")
                        else:
                            excel_buffer = remplir_fiche_paie(user_store["mois"], user_store["annee"], emp)
                            st.session_state[generated_file_key] = {
                                "data": excel_buffer.getvalue(),
                                "filename": f"fiche_paie_{nom_propre}_{user_store['mois']}_{user_store['annee']}.xlsx",
                                "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "download_label": "Télécharger l'Excel",
                                "employee_name": nom_employe_text,
                            }
                            st.rerun()
                    else:
                        if not emp.get("fds"): erreur_type_solo = "de la fin de stage"
                        if not emp.get("dds"): erreur_type_solo = "du début de stage"
                        if emp.get("nom_stagiaire") == "": erreur_type_solo = "du nom"

                        if erreur_type_solo:
                            st.error(f"Impossible de générer : il manque l'information {erreur_type_solo} !")
                        else:
                            docx_buffer = generer_docx_stagiaire(emp, user_store['mois'], user_store['annee'])
                            st.session_state[generated_file_key] = {
                                "data": docx_buffer.getvalue(),
                                "filename": f"Fiche_Stage_{nom_propre}_{user_store['mois']}_{user_store['annee']}.docx",
                                "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                "download_label": "Télécharger le Word",
                                "employee_name": nom_employe_text,
                            }
                            st.rerun()

            generated_file = st.session_state.get(generated_file_key)
            c_download, c_send, _ = st.columns([1, 1, 4])
            if generated_file:
                st.caption("Fichier généré. Générez à nouveau la fiche après toute modification avant de l'envoyer.")
                with c_download:
                    st.download_button(
                        label=generated_file["download_label"],
                        data=generated_file["data"],
                        file_name=generated_file["filename"],
                        mime=generated_file["mime"],
                        key=f"dl_solo_{emp_id}",
                    )
                with c_send:
                    if st.button("Envoyer pour signature", key=f"send_solo_btn_{emp_id}", type="primary"):
                        target_email = emp.get("email_employe", "").strip()
                        if not target_email:
                            st.error("Veuillez renseigner l'adresse e-mail de l'employé dans le formulaire.")
                        else:
                            payload = {
                                "recipient_email": target_email,
                                "employee_name": generated_file["employee_name"],
                                "month": int(user_store["mois"]),
                                "year": int(user_store["annee"]),
                                "filename": generated_file["filename"],
                                "file_b64": base64.b64encode(generated_file["data"]).decode("utf-8"),
                            }
                            try:
                                response = requests.post(
                                    f"{API_URL}/send-fiche",
                                    headers=headers,
                                    json=payload,
                                    timeout=30,
                                )
                                if response.status_code == 200:
                                    st.success(f"Demande de signature envoyée à {target_email} !")
                                else:
                                    detail = response.json().get("detail", response.text)
                                    st.error(f"Erreur lors de l'envoi : {detail}")
                            except requests.RequestException as error:
                                st.error(f"Erreur de communication avec le serveur : {error}")
            else:
                with c_download:
                    st.caption("Générez la fiche pour activer l'envoi.")
                with c_send:
                    st.button(
                        "Envoyer pour signature",
                        key=f"send_solo_btn_{emp_id}",
                        disabled=True,
                        help="Générez d'abord la fiche au format PDF à signer.",
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

            # Visual monthly calendar shared by employees and interns.
            render_monthly_calendar(
                emp,
                int(user_store["mois"]),
                int(user_store["annee"]),
                f"{username}_{emp_id}",
            )

            st.divider()

            # CHAMPS SALARIÉS
            if type_contrat == "Salarié":
                st.subheader("Informations Employé")
                emp["nom"] = st.text_input("NOM Prénom (Employé)", key=f"{username}_employe_nom_{emp_id}", value=emp["nom"])
                emp["email_employe"] = st.text_input("E-mail de l'employé", placeholder="employe@univ-ilci.fr", key=f"{username}_emp_mail_{emp_id}", value=emp.get("email_employe", ""), help="L'adresse e-mail à laquelle la fiche de présence sera envoyée.")
                emp["responsable"] = st.text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{emp_id}", value=emp["responsable"])
                emp["email_responsable"] = st.text_input("Email du responsable", placeholder="responsable@univ-ilci.fr", key=f"{username}_resp_mail_{emp_id}", value=emp.get("email_responsable", ""), help="Un mail sera envoyé au responsable un mois avant la fin du contrat de l'employé.")
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
                with st.expander("Temps partiel / Planning hebdomadaire"):
                    st.write("Indiquez les horaires pour chaque jour (décochez si non travaillé) :")
                    
                    jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
                    
                    if "planning_detail" not in emp:
                        emp["planning_detail"] = {j: {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": j not in ("Samedi", "Dimanche")} for j in jours}

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
                    nb_jours_vac = st.number_input(
                        "Nombre de jours :",
                        min_value=0,
                        max_value=31,
                        value=len(emp["vacances"]),
                        key=f"{username}_nb_jours_vac_{emp_id}_{len(emp['vacances'])}",
                    )

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
                            vac["examen_alt"] = st.checkbox(f"Examen alternance", value=vac["examen_alt"] if "examen_alt" in vac else False, key=f"{username}_examen_alt_{emp_id}_{i}", help="Cochez la case si c'est un alternant qui pose des jours de congés pour les révisions de ses examens")

                # Section Absences
                with st.expander("Absences"):
                    st.subheader("Saisir les jours d'absences")
                    nb_jours_abs = st.number_input(
                        "Nombre de jours :",
                        min_value=0,
                        max_value=31,
                        value=len(emp["absences"]),
                        key=f"{username}_nb_jours_abs_{emp_id}_{len(emp['absences'])}",
                    )

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
                    nb_jours_am = st.number_input(
                        "Nombre de jours",
                        min_value=0,
                        max_value=31,
                        value=len(emp["arret"]),
                        key=f"{username}_nb_jours_am_{emp_id}_{len(emp['arret'])}",
                    )

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
                    emp["email_employe"] = st.text_input("E-mail du stagiaire", placeholder="stagiaire@univ-ilci.fr", key=f"{username}_emp_mail_{emp_id}", value=emp.get("email_employe", ""), help="L'adresse e-mail à laquelle la fiche de présence sera envoyée.")
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
if st.button("Sauvegarder", width="stretch"):
    try:
        # On prépare les données (conversion des dates en texte)
        data_to_send = serialize_dates(user_store)
        
        response = requests.post(
            f"{API_URL}/save-config/{username}",
            headers=headers, 
            json=data_to_send
        )
        
        if response.status_code == 200:
            st.success("Données synchronisées avec succès ! L'e-mail de l'employé est mémorisé pour les prochains envois.")
        else:
            st.error(f"Erreur lors de la sauvegarde: {response.status_code}")
    except Exception as e:
        st.error(f"Impossible de joindre le serveur : {e}")

# ENVOI GLOBAL POUR SIGNATURE
if st.button(
    "Envoyer toutes les fiches pour signature",
    type="primary",
    help="Génère chaque fiche puis envoie une demande de signature à l'adresse e-mail enregistrée.",
):
    envois_reussis = []
    erreurs_envoi = []
    fiches_a_envoyer = user_store["employes_data"]

    if not fiches_a_envoyer:
        st.warning("Aucune fiche n'est disponible à envoyer.")
    else:
        progress = st.progress(0, text="Préparation des demandes de signature…")
        for index, employe in enumerate(fiches_a_envoyer, start=1):
            est_stagiaire = employe.get("type") == "Stagiaire"
            emp_id = employe.get("id")
            nom = (
                " ".join(
                    part.strip()
                    for part in (str(employe.get("prenom_stagiaire") or ""), str(employe.get("nom_stagiaire") or ""))
                    if part.strip()
                )
                if est_stagiaire
                else str(employe.get("nom") or "").strip()
            )
            # Lire la valeur du widget en priorité : elle peut être plus récente
            # que la copie de la fiche en mémoire au moment du clic groupé.
            email_widget_key = f"{username}_emp_mail_{emp_id}"
            email = str(st.session_state.get(email_widget_key, employe.get("email_employe") or "")).strip()
            employe["email_employe"] = email
            informations_manquantes = []

            if not nom:
                informations_manquantes.append("nom")
            if not email:
                informations_manquantes.append("e-mail")
            if est_stagiaire:
                if not employe.get("dds"):
                    informations_manquantes.append("début de stage")
                if not employe.get("fds"):
                    informations_manquantes.append("fin de stage")
            else:
                if not employe.get("ddc"):
                    informations_manquantes.append("début de contrat")
                if not employe.get("fdc"):
                    informations_manquantes.append("fin de contrat")
                if not str(employe.get("responsable") or "").strip():
                    informations_manquantes.append("responsable")

            if informations_manquantes:
                erreurs_envoi.append(
                    f"{nom or f'Fiche {index}'} : information manquante ({', '.join(informations_manquantes)})."
                )
            else:
                try:
                    nom_fichier = nom.replace(" ", "_")
                    if est_stagiaire:
                        fichier = generer_docx_stagiaire(employe, user_store["mois"], user_store["annee"])
                        filename = f"Fiche_stage_{nom_fichier}_{user_store['mois']}_{user_store['annee']}.docx"
                    else:
                        fichier = remplir_fiche_paie(user_store["mois"], user_store["annee"], employe)
                        filename = f"fiche_paie_{nom_fichier}_{user_store['mois']}_{user_store['annee']}.xlsx"

                    payload = {
                        "recipient_email": email,
                        "employee_name": nom,
                        "month": int(user_store["mois"]),
                        "year": int(user_store["annee"]),
                        "filename": filename,
                        "file_b64": base64.b64encode(fichier.getvalue()).decode("utf-8"),
                    }
                    response = requests.post(
                        f"{API_URL}/send-fiche",
                        headers=headers,
                        json=payload,
                        timeout=90,
                    )
                    if response.status_code == 200:
                        envois_reussis.append(f"{nom} ({email})")
                    else:
                        try:
                            detail = response.json().get("detail", response.text)
                        except ValueError:
                            detail = response.text
                        erreurs_envoi.append(f"{nom} : {detail}")
                except requests.RequestException as error:
                    erreurs_envoi.append(f"{nom} : impossible de joindre le serveur ({error}).")
                except Exception as error:
                    erreurs_envoi.append(f"{nom} : impossible de préparer la fiche ({error}).")

            progress.progress(index / len(fiches_a_envoyer), text=f"Traitement de {index}/{len(fiches_a_envoyer)}…")

        progress.empty()
        if envois_reussis:
            st.success(f"{len(envois_reussis)} demande(s) de signature envoyée(s) : {', '.join(envois_reussis)}.")
        if erreurs_envoi:
            st.error("Certaines fiches n'ont pas été envoyées :\n\n" + "\n\n".join(f"- {erreur}" for erreur in erreurs_envoi))

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
            width="stretch"
        )
