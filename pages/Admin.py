"""Administration of organization groups and user assignments."""

from __future__ import annotations

import requests
import streamlit as st

from api_client import api_url, authenticated_headers


API_URL = api_url()
HEADERS = authenticated_headers()
CURRENT_USER = st.session_state.get("user")

if not CURRENT_USER or CURRENT_USER.get("role") != "Admin":
    st.error("Cette page est réservée aux administrateurs.")
    st.stop()


def api_error(response: requests.Response) -> str:
    try:
        return str(response.json().get("detail", "Une erreur est survenue."))
    except ValueError:
        return "Une erreur est survenue."


def fetch_groups() -> list[dict]:
    response = requests.get(f"{API_URL}/groups", headers=HEADERS, timeout=10)
    if response.status_code != 200:
        st.error(api_error(response))
        return []
    return list(response.json().get("groups", []))


def fetch_users() -> list[dict]:
    response = requests.get(f"{API_URL}/list-users", headers=HEADERS, timeout=10)
    if response.status_code != 200:
        st.error(api_error(response))
        return []
    return list(response.json().get("users", []))


st.title("Administration")
st.caption("Gérez les Groupes, les comptes et les périmètres de responsabilité.")

groups = fetch_groups()
users = fetch_users()
group_names = {int(group["id"]): str(group["name"]) for group in groups}
active_group_ids = [int(group["id"]) for group in groups if group.get("is_active")]

st.subheader("Groupes")
with st.form("create_group_form", clear_on_submit=True):
    new_group_name = st.text_input("Nom du nouveau Groupe")
    submitted_group = st.form_submit_button("Créer le Groupe", type="primary")
    if submitted_group:
        response = requests.post(f"{API_URL}/groups", headers=HEADERS, json={"name": new_group_name}, timeout=10)
        if response.status_code == 200:
            st.success("Groupe créé.")
            st.rerun()
        else:
            st.error(api_error(response))

if groups:
    selected_group_id = st.selectbox(
        "Modifier un Groupe",
        options=[int(group["id"]) for group in groups],
        format_func=lambda group_id: group_names[group_id],
        key="admin_selected_group",
    )
    selected_group = next(group for group in groups if int(group["id"]) == selected_group_id)
    with st.form("update_group_form"):
        edited_group_name = st.text_input("Nom", value=str(selected_group["name"]))
        edited_group_active = st.checkbox("Groupe actif", value=bool(selected_group["is_active"]))
        if st.form_submit_button("Enregistrer le Groupe"):
            response = requests.put(
                f"{API_URL}/groups/{selected_group_id}",
                headers=HEADERS,
                json={"name": edited_group_name, "is_active": edited_group_active},
                timeout=10,
            )
            if response.status_code == 200:
                st.success("Groupe mis à jour.")
                st.rerun()
            else:
                st.error(api_error(response))
else:
    st.info("Créez au moins un Groupe avant d'attribuer des Responsables ou des Employés.")

st.divider()
st.subheader("Créer un utilisateur")
new_role = st.selectbox("Rôle", ["Employe", "Responsable", "Admin"], key="admin_new_role")
with st.form("create_user_form", clear_on_submit=True):
    new_username = st.text_input("Nom d'utilisateur")
    new_email = st.text_input("E-mail professionnel")
    new_password = st.text_input("Mot de passe", type="password")
    employee_type = "salarie"
    group_ids: list[int] = []
    managed_group_ids: list[int] = []
    manager_id: str | None = None
    if new_role == "Employe":
        employee_type = st.selectbox("Type de personnel", ["salarie", "stagiaire"])
        group_ids = st.multiselect(
            "Groupes d'appartenance", active_group_ids, format_func=lambda group_id: group_names[group_id]
        )
        managers = [user["username"] for user in users if user.get("role") == "Responsable"]
        manager_id = st.selectbox("Responsable direct", managers, index=None, placeholder="Sélectionnez un Responsable")
    elif new_role == "Responsable":
        managed_group_ids = st.multiselect(
            "Groupes gérés", active_group_ids, format_func=lambda group_id: group_names[group_id]
        )
    if st.form_submit_button("Créer l'utilisateur", type="primary"):
        payload = {
            "new_username": new_username,
            "new_mail": new_email,
            "new_password": new_password,
            "new_role": new_role,
            "employee_type": employee_type,
            "group_ids": group_ids,
            "managed_group_ids": managed_group_ids,
            "manager_id": manager_id,
        }
        response = requests.post(f"{API_URL}/create-user", headers=HEADERS, json=payload, timeout=10)
        if response.status_code == 200:
            st.success("Utilisateur créé.")
            st.rerun()
        else:
            st.error(api_error(response))

st.divider()
st.subheader("Utilisateurs")
if users:
    rows = []
    for user in users:
        rows.append(
            {
                "Utilisateur": user["username"],
                "Rôle": user["role"],
                "Type": user.get("employee_type") if user["role"] == "Employe" else "",
                "Responsable": user.get("manager_id") or "",
                "Groupes": ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in user.get("group_ids", [])),
                "Groupes gérés": ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in user.get("managed_group_ids", [])),
                "E-mail": user["email"],
            }
        )
    st.dataframe(rows, width="stretch", hide_index=True)

    selected_username = st.selectbox("Modifier l'organisation d'un utilisateur", [user["username"] for user in users])
    selected_user = next(user for user in users if user["username"] == selected_username)
    updated_role = st.selectbox(
        "Rôle attribué",
        ["Employe", "Responsable", "Admin"],
        index=["Employe", "Responsable", "Admin"].index(selected_user["role"]),
        key="admin_updated_role",
    )
    with st.form("update_user_organization_form"):
        updated_employee_type = selected_user.get("employee_type") or "salarie"
        updated_group_ids = selected_user.get("group_ids", [])
        updated_managed_group_ids = selected_user.get("managed_group_ids", [])
        updated_manager_id = selected_user.get("manager_id")
        if updated_role == "Employe":
            updated_employee_type = st.selectbox(
                "Type de personnel", ["salarie", "stagiaire"], index=["salarie", "stagiaire"].index(updated_employee_type)
            )
            updated_group_ids = st.multiselect(
                "Groupes d'appartenance", active_group_ids, default=updated_group_ids, format_func=lambda group_id: group_names[group_id]
            )
            manager_options = [user["username"] for user in users if user.get("role") == "Responsable"]
            manager_index = manager_options.index(updated_manager_id) if updated_manager_id in manager_options else None
            updated_manager_id = st.selectbox("Responsable direct", manager_options, index=manager_index, placeholder="Sélectionnez un Responsable")
        elif updated_role == "Responsable":
            updated_managed_group_ids = st.multiselect(
                "Groupes gérés", active_group_ids, default=updated_managed_group_ids, format_func=lambda group_id: group_names[group_id]
            )
        if st.form_submit_button("Enregistrer l'organisation"):
            response = requests.put(
                f"{API_URL}/users/{selected_username}/organization",
                headers=HEADERS,
                json={
                    "role": updated_role,
                    "employee_type": updated_employee_type,
                    "group_ids": updated_group_ids,
                    "managed_group_ids": updated_managed_group_ids,
                    "manager_id": updated_manager_id,
                },
                timeout=10,
            )
            if response.status_code == 200:
                st.success("Organisation mise à jour.")
                st.rerun()
            else:
                st.error(api_error(response))

    st.subheader("Supprimer un utilisateur")
    deletable_users = [user["username"] for user in users if user["username"] != CURRENT_USER["name"]]
    if deletable_users:
        delete_username = st.selectbox("Utilisateur à supprimer", deletable_users)
        confirm_delete = st.checkbox("Je confirme la suppression du compte. Les fiches existantes seront conservées.")
        if st.button("Supprimer l'utilisateur", type="secondary", disabled=not confirm_delete):
            response = requests.delete(f"{API_URL}/delete-user/{delete_username}", headers=HEADERS, timeout=10)
            if response.status_code == 200:
                st.success(response.json().get("message", "Compte supprimé."))
                st.rerun()
            else:
                st.error(api_error(response))
