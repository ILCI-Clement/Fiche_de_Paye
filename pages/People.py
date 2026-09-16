"""Personnel view for managers and administrators."""

from __future__ import annotations

import requests
import streamlit as st

from access_control import role_tags
from api_client import api_url, authenticated_headers


API_URL = api_url()
HEADERS = authenticated_headers()
USER = st.session_state.get("user")

USER_TAGS = role_tags(USER) if USER else set()

if not USER or not ({"Admin", "Responsable"} & USER_TAGS):
    st.error("Cette page est réservée aux administrateurs et responsables.")
    st.stop()


def detail(response: requests.Response) -> str:
    try:
        return str(response.json().get("detail", "Une erreur est survenue."))
    except ValueError:
        return "Une erreur est survenue."


users_response = requests.get(f"{API_URL}/list-users", headers=HEADERS, timeout=10)
groups_response = requests.get(f"{API_URL}/groups", headers=HEADERS, timeout=10)
if users_response.status_code != 200 or groups_response.status_code != 200:
    st.error(detail(users_response if users_response.status_code != 200 else groups_response))
    st.stop()

users = list(users_response.json().get("users", []))
groups = list(groups_response.json().get("groups", []))
group_names = {int(group["id"]): str(group["name"]) for group in groups if group.get("is_active")}
group_ids = list(group_names)

st.title("Personnel")
st.caption("Les utilisateurs affichés appartiennent à votre périmètre de responsabilité.")

if "Admin" in USER_TAGS:
    st.info("La gestion complète des comptes et des Groupes est disponible dans Administration.")

st.subheader("Employés gérés")
if users:
    st.dataframe(
        [
            {
                "Utilisateur": person["username"],
                "Type": person.get("employee_type") or "salarie",
                "Responsable": person.get("manager_id") or "",
                "Groupes": ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in person.get("group_ids", [])),
                "E-mail": person["email"],
            }
            for person in users
            if "Employe" in set(person.get("role_tags", [person.get("role")]))
        ],
        width="stretch",
        hide_index=True,
    )
else:
    st.info("Aucun Employé ne fait actuellement partie de votre périmètre.")

if "Responsable" in USER_TAGS and "Admin" not in USER_TAGS:
    st.divider()
    st.subheader("Créer un Employé")
    if not group_ids:
        st.caption("Aucun Groupe actif ne vous est attribué. Vous pouvez créer l'Employé sans Groupe; vous serez son Responsable direct.")
    with st.form("manager_create_employee", clear_on_submit=True):
        username = st.text_input("Nom d'utilisateur")
        email = st.text_input("E-mail professionnel")
        password = st.text_input("Mot de passe", type="password")
        employee_type = st.selectbox("Type de personnel", ["salarie", "stagiaire"])
        selected_groups = st.multiselect("Groupes d'appartenance (facultatif)", group_ids, format_func=lambda group_id: group_names[group_id])
        if st.form_submit_button("Créer l'Employé", type="primary"):
            response = requests.post(
                f"{API_URL}/create-user",
                headers=HEADERS,
                json={
                    "new_username": username,
                    "new_mail": email,
                    "new_password": password,
                    "new_role": "Employe",
                    "employee_type": employee_type,
                    "group_ids": selected_groups,
                },
                timeout=10,
            )
            if response.status_code == 200:
                st.success("Employé créé.")
                st.rerun()
            else:
                st.error(detail(response))
