from access_control import role_tags
from contract_periods import matching_contracts, bind_contract, period_sheets
import streamlit as st
from datetime import date, datetime
import requests
import time
import base64
import os
from fiche_exports import fiche_name, validate_fiche, generate_fiche_file
from fiche_batch_ui import render_batch_selection
from fiche_serialization import serialize_dates, deserialize_dates
from planning import default_planning
from calendar_view import render_monthly_calendar
from period_selector import select_period
import zipfile
import hashlib
import json
from pdf_export import convert_fiche_pdf
import io
from api_client import api_url, authenticated_headers, safe_api_request

# Secrets de streamlit
API_URL = api_url()
HORAIRES = [f"{h:02d}:{m:02d}" for h in range(7, 21) for m in (0, 30)]
HORAIRES.insert(0, "") # Option vide pour les jours non travaillés

# Configuration du header pour les requêtes
headers = authenticated_headers()


if "user" not in st.session_state or st.session_state["user"] is None:
    st.warning("Veuillez vous connecter d'abord.")
    st.stop()

username = st.session_state['user']['name']
role = st.session_state['user']['role']
USER_TAGS = role_tags(st.session_state['user'])
IS_MANAGER = "Responsable" in USER_TAGS and "Admin" not in USER_TAGS
if not ({"Admin", "Responsable"} & USER_TAGS):
    st.error("La saisie des présences est réservée aux administrateurs et responsables.")
    st.stop()

def person_text_input(*args, **kwargs):
    kwargs["disabled"] = IS_MANAGER
    return st.text_input(*args, **kwargs)


if "user_data" not in st.session_state:
    st.session_state.user_data = {}

# CHARGEMENT DES DONNEES DU VPS (MariaDB)
if username not in st.session_state.user_data:
    try:
        # Appel GET à l'API pour récupérer le JSON stocké
        response = safe_api_request("GET", f"{API_URL}/get-config/{username}", headers=headers)
        if response is not None and response.status_code == 200 and response.json():
            # On récupère les données et on convertit les strings en dates
            raw_data = response.json()
            st.session_state.user_data[username] = deserialize_dates(raw_data)
        elif response is not None:
            st.error(f"Impossible de charger la configuration ({response.status_code}).")
            st.session_state.user_data[username] = {}
        else:
            st.stop()
    except (ValueError, TypeError) as error:
        st.error(f"La configuration enregistrée est invalide : {error}")
        st.session_state.user_data[username] = {}

# Raccourci vers les données de l'utilisateur actuel
user_store = st.session_state.user_data[username]
contract_sync_key = f"fiche_contract_sync_pending_{username}"


def save_current_configuration(show_success: bool = False) -> bool:
    """Persist the current fiche configuration and make contract dates available to the dashboard."""
    try:
        response = requests.post(
            f"{API_URL}/save-config/{username}",
            headers=headers,
            json=serialize_dates(user_store),
            timeout=15,
        )
    except requests.RequestException as error:
        st.error(f"Impossible de joindre le serveur : {error}")
        return False
    if response.status_code != 200:
        st.error(f"Erreur lors de la sauvegarde: {response.status_code}")
        return False
    if show_success:
        st.success("Données synchronisées avec succès ! L'e-mail de l'employé est mémorisé pour les prochains envois.")
    return True


def mark_contract_sync_pending() -> None:
    """Defer the save until every widget value has been copied into the fiche record."""
    st.session_state[contract_sync_key] = True

# FORMULAIRE PRINCIPAL 
st.title("Générateur de fiche de présence")

# Sélection du mois et de l'année
now = datetime.now()
user_store["mois"], user_store["annee"] = select_period(
    f"fiches_{username}", "top", user_store.get("mois", now.month), user_store.get("annee", now.year)
)

if "employes_data" not in user_store:
    user_store["employes_data"] = []


def create_fiche_from_account(account: dict, manager_by_username: dict[str, dict]) -> dict:
    """Create a new monthly fiche prefilled with the selected account's known data."""
    employee_type = "Stagiaire" if account.get("employee_type") == "stagiaire" else "Salarié"
    manager_username = str(account.get("manager_id") or "")
    manager = manager_by_username.get(manager_username.casefold(), {})
    start_date = account.get("contract_start_date")
    end_date = account.get("contract_end_date")
    if isinstance(start_date, str) and start_date:
        try:
            start_date = date.fromisoformat(start_date)
        except ValueError:
            start_date = None
    if isinstance(end_date, str) and end_date:
        try:
            end_date = date.fromisoformat(end_date)
        except ValueError:
            end_date = None

    employee = {
        "id": int(time.time() * 1000),
        "type": employee_type,
        "nom": " ".join(str(account.get(field) or "").strip() for field in ("prenom", "nom")).strip(),
        "nom_stagiaire": str(account.get("last_name") or ""),
        "prenom_stagiaire": str(account.get("first_name") or ""),
        "email_employe": str(account.get("email") or ""),
        "account_username": str(account["username"]),
        "responsable": "",
        "email_responsable": str(manager.get("email") or ""),
        "ddc": start_date if employee_type == "Salarié" else None,
        "fdc": ("Pas de fin" if account.get("is_cdi") else end_date) if employee_type == "Salarié" else None,
        "cdi": bool(account.get("is_cdi")) if employee_type == "Salarié" else False,
        "dds": start_date if employee_type == "Stagiaire" else None,
        "fds": end_date if employee_type == "Stagiaire" else None,
        "nb_jours": 0,
        "nb_heures_jour": 0.0,
        "taux_horaire": 0.0,
        "facture_mensuelle": 0.0,
        "transport": "",
        "taux": 0.0,
        "vacances": [],
        "absences": [],
        "arret": [],
        "calendar_overrides": {},
        "planning_detail": default_planning(),
    }
    return employee


def account_already_in_fiches(employee: dict, account: dict) -> bool:
    """Match stable account links first, with email/name fallbacks for legacy fiches."""
    username_key = str(account.get("username") or "").strip().casefold()
    email_key = str(account.get("email") or "").strip().casefold()
    linked_username = str(employee.get("account_username") or "").strip().casefold()
    if username_key and linked_username == username_key:
        return True
    employee_email = str(employee.get("email_employe") or "").strip().casefold()
    if email_key and employee_email and email_key == employee_email:
        return True
    legacy_name = str(employee.get("nom") or "").strip().casefold()
    return bool(username_key and legacy_name == username_key)


def display_person_label(value: str) -> str:
    """Show the mailbox name without creating an automatic email link."""
    return str(value or "").strip().split("@", 1)[0]


def account_display_label(account: dict) -> str:
    """Use an available full name, falling back to the account email."""
    full_name = " ".join(
        str(account.get(field) or "").strip()
        for field in ("last_name", "first_name")
        if str(account.get(field) or "").strip()
    )
    return display_person_label(full_name or str(account.get("email") or account.get("username") or "Compte employé").strip())


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
    email = str(employee.get("email_employe") or "").strip()
    label = display_person_label(name or email or str(employee.get("account_username") or "").strip() or f"Employé {index + 1}")
    if employee.get("contract_id"):
        label += f" · {str(employee['sheet_period_start'])[8:10]}–{str(employee['sheet_period_end'])[8:10]} · {employee['contract_snapshot']['employee_type']}"
    return label


accounts_response = safe_api_request("GET", f"{API_URL}/list-users", headers=headers)
accounts = []
if accounts_response is not None and accounts_response.status_code == 200:
    accounts = list(accounts_response.json().get("users", []))
elif accounts_response is not None:
    try:
        st.warning(str(accounts_response.json().get("detail", "Les comptes existants ne sont pas disponibles.")))
    except ValueError:
        st.warning("Les comptes existants ne sont pas disponibles.")

manager_by_username = {
    str(account.get("username") or "").casefold(): account
    for account in accounts
}


# Bouton pour ajouter un employé à la fin de la liste
if "Admin" in USER_TAGS and st.button("Ajouter un employé / stagiaire", width="stretch"):
    user_store["employes_data"].append({
        "id": int(time.time() * 1000),
        "type": "Salarié",
        "nom": "", "email_employe": "", "account_username": "", "account_username": "", "responsable": "", "email_responsable": "", "ddc": None, "fdc": None, "cdi": False,
        "vacances": [], "absences": [], "arret": [],
        "calendar_overrides": {},
        "planning_detail": {j: {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": j not in ("Samedi", "Dimanche")} for j in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]}
    })
    st.rerun() # On force Streamlit à recréer les onglets immédiatement

# A searchable selector keeps existing accounts available without a button grid.
if accounts:
    accounts_by_username = {str(account["username"]): account for account in accounts}
    with st.form("select_existing_employee"):
        selected_username = st.selectbox(
            "Sélectionner l'employé existant",
            options=list(accounts_by_username),
            index=None,
            placeholder="Rechercher par nom de compte ou e-mail…",
            format_func=lambda key: f"{account_display_label(accounts_by_username[key])} ({display_person_label(key)})",
        )
        open_existing = st.form_submit_button("Ouvrir / créer sa fiche")
    if open_existing:
        if selected_username is None:
            st.warning("Sélectionnez une personne dans la liste.")
        else:
            account = accounts_by_username[selected_username]
            contract_response = requests.get(f"{API_URL}/users/{selected_username}/contracts", headers=headers, timeout=15)
            if contract_response.status_code != 200:
                st.error("Impossible de charger les contrats de cette personne.")
                st.stop()
            contracts = matching_contracts(contract_response.json()["contracts"], int(user_store["mois"]), int(user_store["annee"]))
            if not contracts:
                st.warning("Aucun contrat pour ce mois. Un administrateur doit renseigner le contrat dans Administration.")
                st.stop()
            for contract in contracts:
                existing_fiche = next(((index, fiche) for index, fiche in enumerate(user_store["employes_data"])
                    if fiche.get("account_username") == selected_username and fiche.get("contract_id") == contract["id"]
                    and fiche.get("sheet_month") == user_store["mois"] and fiche.get("sheet_year") == user_store["annee"]), None)
                if existing_fiche:
                    index, fiche = existing_fiche
                else:
                    fiche = bind_contract(create_fiche_from_account(account, manager_by_username), contract, int(user_store["mois"]), int(user_store["annee"]))
                    prior = next((e for e in reversed(user_store["employes_data"]) if e.get("account_username") == selected_username), None)
                    if prior:
                        import copy
                        for field in ("nom", "nom_stagiaire", "prenom_stagiaire", "responsable", "email_responsable", "planning_detail"):
                            if prior.get(field):
                                fiche[field] = copy.deepcopy(prior[field])
                        if prior.get("type") == fiche.get("type"):
                            for field in ("taux_horaire", "nb_heures_jour", "transport", "taux"):
                                if field in prior:
                                    fiche[field] = prior[field]
                    user_store["employes_data"].append(fiche)
                    index = len(user_store["employes_data"]) - 1
                st.session_state["fiche_active_tab"] = fiche_tab_label(fiche, index)
                st.session_state[f"fiche_active_id_{username}_{user_store['annee']}_{user_store['mois']}"] = str(fiche["id"])
            save_current_configuration()
            st.rerun()

# Si la liste est vide, on affiche un message d'aide
if not user_store["employes_data"]:
    st.info("Aucun employé ou stagiaire configuré. Cliquez sur le bouton ci-dessus pour commencer.")

if user_store["employes_data"]:
    visible_indices = [i for i, employee in enumerate(user_store["employes_data"])
        if not employee.get("contract_id") or (employee.get("sheet_month") == user_store["mois"] and employee.get("sheet_year") == user_store["annee"])]
    if not visible_indices:
        st.info("Sélectionnez une personne pour créer sa fiche de ce mois.")
        st.stop()
    # Keep a separate persistent sheet ID: labels can change while editing.
    from uuid import uuid4
    for i in visible_indices:
        user_store["employes_data"][i].setdefault("id", uuid4().hex)
    labels_onglets = [fiche_tab_label(user_store["employes_data"][i], i) for i in visible_indices]
    for position, label in enumerate(list(labels_onglets)):
        if labels_onglets.count(label) > 1:
            for duplicate in [j for j, value in enumerate(labels_onglets) if value == label]:
                labels_onglets[duplicate] = f"{label} · fiche {duplicate + 1}"

    active_id_key = f"fiche_active_id_{username}_{user_store['annee']}_{user_store['mois']}"
    active_id = st.session_state.get(active_id_key)
    active_position = next((j for j, i in enumerate(visible_indices)
        if str(user_store["employes_data"][i]["id"]) == active_id), 0)
    st.session_state[active_id_key] = str(user_store["employes_data"][visible_indices[active_position]]["id"])
    st.session_state["fiche_active_tab"] = labels_onglets[active_position]

    def remember_fiche_tab():
        label = st.session_state.get("fiche_active_tab")
        if label in labels_onglets:
            selected_index = visible_indices[labels_onglets.index(label)]
            st.session_state[active_id_key] = str(user_store["employes_data"][selected_index]["id"])

    tabs = st.tabs(labels_onglets, default=labels_onglets[0], key="fiche_active_tab", on_change=remember_fiche_tab)

    for tab_index, tab in enumerate(tabs):
        h = visible_indices[tab_index]
        if not tab.open:
            continue
        with tab:
            emp = user_store["employes_data"][h]
            
            # Si un employé n'a pas d'ID, on lui en donne un
            if "id" not in emp:
                emp["id"] = int(time.time() * 1000) + h

            emp_id = emp["id"]

            generated_file_key = f"generated_fiche_{emp_id}"
            generated_pdf_key = f"{generated_file_key}_pdf"

            c_space, c_gen, c_del = st.columns([4, 1, 1])
            with c_space:
                fiche_label = labels_onglets[tab_index]
                st.subheader(f"Fiche de {fiche_label}" if fiche_label else "Fiche d'employé")
            with c_del:
                if "Admin" in USER_TAGS and st.button("Supprimer cette fiche", key=f"del_btn_{emp_id}", type="secondary", help="Supprime définitivement cet employé de la liste"):
                    user_store["employes_data"].pop(h)
                    st.session_state.pop(generated_file_key, None)
                    st.session_state.pop(generated_pdf_key, None)
                    st.success("Fiche supprimée ! Sauvegardez pour appliquer les changements sur le serveur.")
                    st.rerun()
            with c_gen:
                if st.button("Générer cette fiche", key=f"gen_solo_btn_{emp_id}", type="primary", help="Charge uniquement la fiche de cet employé"):
                    try:
                        validate_fiche(emp)
                        generated = generate_fiche_file(emp, user_store["mois"], user_store["annee"])
                    except Exception as error:
                        st.error(f"Impossible de générer la fiche : {error}")
                    else:
                        st.session_state.pop(generated_pdf_key, None)
                        st.session_state[generated_file_key] = generated
                        st.rerun()

            generated_file = st.session_state.get(generated_file_key)
            c_download, c_pdf, c_send, _ = st.columns([1, 1, 1, 3])
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
                with c_pdf:
                    generated_pdf = st.session_state.get(generated_pdf_key)
                    if generated_pdf:
                        st.download_button(
                            label="Télécharger le PDF",
                            data=generated_pdf["data"],
                            file_name=generated_pdf["filename"],
                            mime="application/pdf",
                            key=f"dl_pdf_solo_{emp_id}",
                            on_click="ignore",
                        )
                    elif st.button("Préparer le PDF", key=f"pdf_solo_btn_{emp_id}"):
                        try:
                            with st.spinner("Préparation du PDF…"):
                                pdf_data = convert_fiche_pdf(
                                    API_URL, headers, generated_file["filename"], generated_file["data"]
                                )
                            stem = os.path.splitext(generated_file["filename"])[0]
                            st.session_state[generated_pdf_key] = {
                                "data": pdf_data,
                                "filename": f"{stem}.pdf",
                            }
                            st.rerun()
                        except ValueError as error:
                            st.error(f"Impossible de préparer le PDF : {error}")
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
                                    st.success(response.json().get("message", f"Demande de signature envoyée à {target_email} !"))
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
                horizontal=True,
                disabled=IS_MANAGER or bool(emp.get("contract_id"))
            )
            emp["type"] = type_contrat

            st.divider()

            # Visual monthly calendar shared by employees and interns.
            with st.expander("Calendrier mensuel", expanded=False):
                select_period(f"fiches_{username}", f"calendar_{emp_id}", user_store["mois"], user_store["annee"])
                render_monthly_calendar(
                    emp,
                    int(user_store["mois"]),
                    int(user_store["annee"]),
                    f"{username}_{emp_id}",
                )

            st.divider()

            # CHAMPS SALARIÉS
            if type_contrat == "Salarié":
                with st.expander("Informations personnelles et contrat", expanded=False):
                    st.subheader("Informations Employé")
                    emp["nom"] = person_text_input("NOM Prénom (Employé)", key=f"{username}_employe_nom_{emp_id}", value=emp["nom"])
                    emp["email_employe"] = person_text_input("E-mail de l'employé", placeholder="employe@univ-ilci.fr", key=f"{username}_emp_mail_{emp_id}", value=emp.get("email_employe", ""), help="L'adresse e-mail à laquelle la fiche de présence sera envoyée.")
                    emp["responsable"] = person_text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{emp_id}", value=emp["responsable"])
                    emp["email_responsable"] = person_text_input("Email du responsable", placeholder="responsable@univ-ilci.fr", key=f"{username}_resp_mail_{emp_id}", value=emp.get("email_responsable", ""), help="Un mail sera envoyé au responsable un mois avant la fin du contrat de l'employé.")
                    c1, c2 = st.columns(2)
                    with c1:
                        emp["ddc"] = st.date_input("Début de contrat", key=f"ddc_{emp_id}", value=emp.get("ddc"), format="DD/MM/YYYY", disabled=IS_MANAGER or bool(emp.get("contract_id")))
                        emp["cdi"] = st.checkbox("Contrat CDI ?", value=emp.get("cdi", False), key=f"cdi_{emp_id}", disabled=IS_MANAGER or bool(emp.get("contract_id")))
                    with c2:
                        if not emp["cdi"]:
                            emp["fdc"] = st.date_input("Fin de contrat", key=f"fdc_{emp_id}", value=emp.get("fdc") if emp.get("fdc") != "Pas de fin" else None, format="DD/MM/YYYY", disabled=IS_MANAGER or bool(emp.get("contract_id")))
                        else:
                            emp["fdc"] = "Pas de fin"
                            st.write("Fin de contrat : N/A")
                
                # SECTION PLANNINGS ET CONGES
                with st.expander("Temps partiel / Planning hebdomadaire"):
                    st.write("Indiquez les horaires pour chaque jour (décochez si non travaillé) :")
                    
                    jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
                    
                    if "planning_detail" not in emp:
                        emp["planning_detail"] = default_planning()

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

            else:
                # CHAMPS STAGIAIRES
                with st.expander("Informations du stage et rémunération", expanded=False):
                    st.subheader("Information Stagiaire")
            
                    c1, c2 = st.columns(2)
                    with c1:
                        emp["nom_stagiaire"] = person_text_input("Nom du stagiaire", key=f"st_nom_{emp_id}", value=emp.get("nom_stagiaire", ""))
                        emp["email_employe"] = person_text_input("E-mail du stagiaire", placeholder="stagiaire@univ-ilci.fr", key=f"{username}_emp_mail_{emp_id}", value=emp.get("email_employe", ""), help="L'adresse e-mail à laquelle la fiche de présence sera envoyée.")
                        emp["responsable"] = person_text_input("NOM Prénom (Responsable)", key=f"{username}_resp_nom_{emp_id}", value=emp["responsable"])
                        emp["dds"] = st.date_input("Début de stage", key=f"dds_{emp_id}", value=emp.get("dds"), format="DD/MM/YYYY", disabled=IS_MANAGER or bool(emp.get("contract_id")))
                        emp["nb_jours"] = st.number_input("Nombre de jours", key=f"st_nj_{emp_id}", value=emp.get("nb_jours", 0))
                        emp["taux_horaire"] = st.number_input("Taux horaire (€)", key=f"st_th_{emp_id}", value=emp.get("taux_horaire", 0.0), disabled=IS_MANAGER)
                        emp["facture_mensuelle"] = st.number_input("Facture mensuelle (€)", key=f"st_fm_{emp_id}", value=emp.get("facture_mensuelle", 0.0), disabled=IS_MANAGER)
                    with c2:
                        emp["prenom_stagiaire"] = person_text_input("Prénom du stagiaire", key=f"st_pre_{emp_id}", value=emp.get("prenom_stagiaire", ""))
                        emp["email_responsable"] = person_text_input("Email du responsable", placeholder="responsable@univ-ilci.fr", key=f"{username}_resp_mail_{emp_id}", value=emp.get("email_responsable", ""))
                        emp["fds"] = st.date_input("Fin de stage", key=f"fds_{emp_id}", value=emp.get("fds"), format="DD/MM/YYYY", disabled=IS_MANAGER or bool(emp.get("contract_id")))
                        emp["nb_heures_jour"] = st.number_input("Nombre d'heures/jour", key=f"st_nhj_{emp_id}", value=emp.get("nb_heures_jour", 0.0))
                        emp["transport"] = person_text_input("Transport", key=f"st_tr_{emp_id}", value=emp.get("transport", ""))
                        emp["taux"] = st.number_input("Taux (%)", key=f"st_tx_{emp_id}", value=emp.get("taux", 0.0), disabled=IS_MANAGER)

# BOUTON DE SAUVEGARDE SUR LE VPS
st.divider()
if st.session_state.pop(contract_sync_key, False) and save_current_configuration():
    st.toast("Dates de contrat synchronisées avec le Tableau de bord.")

if st.button("Sauvegarder", width="stretch"):
    save_current_configuration(show_success=True)

# Selection is shared by download and signing, without changing signer resolution.
selected_batch_sheets = render_batch_selection(
    period_sheets(user_store["employes_data"], user_store["mois"], user_store["annee"]),
    username=username, month=user_store["mois"], year=user_store["annee"], label=fiche_tab_label,
)
if selected_batch_sheets is not None:
    st.caption("Signature : salarié d’abord, puis son responsable direct enregistré. "
               "Le téléchargement ZIP ne déclenche aucun e-mail.")
    # ENVOI GLOBAL POUR SIGNATURE
    if st.button(
        "Envoyer la sélection pour signature",
        disabled=not selected_batch_sheets,
        type="primary",
        help="Génère uniquement les fiches cochées puis envoie les demandes une par une.",
    ):
        envois_reussis = []
        erreurs_envoi = []
        fiches_a_envoyer = selected_batch_sheets

        if not fiches_a_envoyer:
            st.warning("Aucune fiche n'est disponible à envoyer.")
        else:
            progress = st.progress(0, text="Préparation des demandes de signature…")
            for index, employe in enumerate(fiches_a_envoyer, start=1):
                nom = fiche_name(employe) or f"Fiche {index}"
                emp_id = employe.get("id")
                email_widget_key = f"{username}_emp_mail_{emp_id}"
                email = str(st.session_state.get(email_widget_key, employe.get("email_employe") or "")).strip()
                employe["email_employe"] = email
                try:
                    validate_fiche(employe, require_email=True)
                except ValueError as error:
                    erreurs_envoi.append(f"{nom} : {error}")
                else:
                    try:
                        generated = generate_fiche_file(employe, user_store["mois"], user_store["annee"])
                        payload = {
                            "recipient_email": email,
                            "employee_name": nom,
                            "month": int(user_store["mois"]),
                            "year": int(user_store["annee"]),
                            "filename": generated["filename"],
                            "file_b64": base64.b64encode(generated["data"]).decode("utf-8"),
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

    # Export des fiches du mois sélectionné.
    batch_sheets = selected_batch_sheets
    batch_format = st.segmented_control(
        "Format de l'export groupé", ["PDF", "Excel / Word"], default="PDF", key="fiche_batch_format"
    ) or "PDF"
    batch_fingerprint = hashlib.sha256(json.dumps(
        {"username": username, "month": user_store["mois"], "year": user_store["annee"],
         "format": batch_format, "sheets": batch_sheets},
        sort_keys=True, default=str, ensure_ascii=False,
    ).encode("utf-8")).hexdigest()
    if st.session_state.get("fiche_batch_export", {}).get("fingerprint") != batch_fingerprint:
        st.session_state.pop("fiche_batch_export", None)

    if st.button("Préparer la sélection en PDF" if batch_format == "PDF" else "Générer la sélection",
                 type="primary", disabled=not batch_sheets, key="prepare_batch_export"):
        st.session_state.pop("fiche_batch_export", None)
        export_errors = []
        for employee in batch_sheets:
            try:
                validate_fiche(employee, require_first_name=True, check_exceptions=True)
            except ValueError as error:
                export_errors.append(f"{fiche_name(employee) or 'Personne'} : {error}")
        if export_errors:
            st.error("Corrigez les fiches avant l'export :\n\n" + "\n\n".join(export_errors))
        else:
            zip_buffer = io.BytesIO()
            progress = st.progress(0, text="Préparation des fiches…")
            export_errors = []
            sheets_to_export = batch_sheets
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, employee in enumerate(sheets_to_export):
                    name = fiche_name(employee) or f"Personne_{idx+1}"
                    try:
                        generated = generate_fiche_file(employee, user_store["mois"], user_store["annee"], ordinal=idx+1)
                        file_name, data = generated["filename"], generated["data"]
                        if batch_format == "PDF":
                            data = convert_fiche_pdf(API_URL, headers, file_name, data)
                            file_name = os.path.splitext(file_name)[0] + ".pdf"
                        zip_file.writestr(file_name, data)
                    except Exception as error:
                        export_errors.append(f"{name} : {error}")
                    progress.progress((idx+1)/len(sheets_to_export), text=f"Fiche {idx+1}/{len(sheets_to_export)}")
            progress.empty()
            if export_errors:
                st.error("L'export n'a pas pu être terminé. Corrigez les erreurs puis réessayez :\n\n"
                         + "\n\n".join(f"- {error}" for error in export_errors))
            else:
                suffix = "pdf" if batch_format == "PDF" else "excel_word"
                st.session_state["fiche_batch_export"] = {
                    "fingerprint": batch_fingerprint,
                    "data": zip_buffer.getvalue(),
                    "filename": f"fiches_presence_{user_store['annee']}_{int(user_store['mois']):02d}_{suffix}.zip",
                    "count": len(sheets_to_export),
                }

    batch_export = st.session_state.get("fiche_batch_export")
    if batch_export:
        st.success(f"{batch_export['count']} fiche(s) prête(s) à télécharger.")
        st.download_button(
            label="Télécharger les PDF sélectionnés (ZIP)" if batch_format == "PDF" else "Télécharger les fiches sélectionnées (ZIP)",
            data=batch_export["data"], file_name=batch_export["filename"], mime="application/zip",
            key="download_batch_export", on_click="ignore", width="stretch",
        )
