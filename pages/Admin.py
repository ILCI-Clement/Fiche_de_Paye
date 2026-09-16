"""Administration of organization groups and user assignments."""

from __future__ import annotations

import requests
import streamlit as st

from access_control import role_tags
from api_client import api_url, authenticated_headers


API_URL = api_url()
HEADERS = authenticated_headers()
CURRENT_USER = st.session_state.get("user")

if not CURRENT_USER or "Admin" not in role_tags(CURRENT_USER):
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
active_group_id_set = set(active_group_ids)


def user_tags(member: dict) -> set[str]:
    return set(member.get("role_tags", [member.get("role")]))


def organization_function(member: dict) -> str:
    tags = user_tags(member)
    if "Responsable" in tags:
        return "Responsable"
    if "Employe" in tags:
        return f"Employé · {member.get('employee_type') or 'salarie'}"
    if "Admin" in tags:
        return "Administrateur"
    return "—"


def organization_rows(members: list[dict]) -> list[dict[str, str]]:
    return [
        {
            "Personne": member["username"],
            "Étiquettes": ", ".join(member.get("role_tags", [member["role"]])),
            "Responsable direct": member.get("manager_id") or "—",
            "Fonction": organization_function(member),
        }
        for member in members
    ]


def is_general_management_group(group: dict) -> bool:
    name = str(group.get("name", "")).casefold()
    return "direction générale" in name or "direction generale" in name


def users_assigned_to_group(group_id: int) -> list[dict]:
    return [
        user
        for user in users
        if group_id in {
            *{int(member_group_id) for member_group_id in (user.get("group_ids") or [])},
            *{int(managed_group_id) for managed_group_id in (user.get("managed_group_ids") or [])},
        }
    ]


with st.expander("Structure des équipes", expanded=True):
    st.caption("Affichage complet par département, responsable direct et responsabilité de département.")
    general_groups = [group for group in groups if is_general_management_group(group)]
    general_ids = {int(group["id"]) for group in general_groups}
    st.subheader("Direction générale")
    general_member_names = {
        user["username"]
        for group_id in general_ids
        for user in users_assigned_to_group(group_id)
    }
    general_members = [user for user in users if user["username"] in general_member_names]
    if general_members:
        st.dataframe(organization_rows(general_members), width="stretch", hide_index=True)
    else:
        st.info("Attribuez un responsable ou un membre au Groupe Direction générale pour l'afficher ici.")

    st.subheader("Départements")
    department_groups = [group for group in groups if int(group["id"]) not in general_ids]
    for group in department_groups:
        group_id = int(group["id"])
        members = users_assigned_to_group(group_id)
        with st.expander(f"{group['name']} · {len(members)} personne(s)", expanded=False):
            if members:
                st.dataframe(organization_rows(members), width="stretch", hide_index=True)
            else:
                st.caption("Aucune personne n'est encore attribuée à ce département.")

    unassigned_members = [user for user in users if not user.get("group_ids")]
    if unassigned_members:
        with st.expander(f"Sans département · {len(unassigned_members)} personne(s)", expanded=False):
            st.dataframe(organization_rows(unassigned_members), width="stretch", hide_index=True)

with st.expander("Groupes", expanded=False):
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
            "Modifier un Groupe", options=[int(group["id"]) for group in groups],
            format_func=lambda group_id: group_names[group_id], key="admin_selected_group",
        )
        selected_group = next(group for group in groups if int(group["id"]) == selected_group_id)
        with st.form("update_group_form"):
            edited_group_name = st.text_input("Nom", value=str(selected_group["name"]))
            edited_group_active = st.checkbox("Groupe actif", value=bool(selected_group["is_active"]))
            if st.form_submit_button("Enregistrer le Groupe"):
                response = requests.put(f"{API_URL}/groups/{selected_group_id}", headers=HEADERS, json={"name": edited_group_name, "is_active": edited_group_active}, timeout=10)
                if response.status_code == 200:
                    st.success("Groupe mis à jour.")
                    st.rerun()
                else:
                    st.error(api_error(response))
    else:
        st.info("Créez un Groupe lorsque vous souhaitez structurer une équipe.")

st.divider()
st.subheader("Créer un utilisateur")
new_role_tags = st.multiselect("Étiquettes d'identité", ["Admin", "Responsable", "Employe"], default=["Employe"], key="admin_new_role_tags")
with st.form("create_user_form", clear_on_submit=True):
    new_username = st.text_input("Nom d'utilisateur")
    new_email = st.text_input("E-mail professionnel")
    new_password = st.text_input("Mot de passe", type="password")
    employee_type = "salarie"
    group_ids: list[int] = []
    managed_group_ids: list[int] = []
    manager_id: str | None = None
    if "Employe" in new_role_tags:
        employee_type = st.selectbox("Type de personnel", ["salarie", "stagiaire"])
        group_ids = st.multiselect(
            "Groupes d'appartenance (facultatif)", active_group_ids, format_func=lambda group_id: group_names[group_id]
        )
        if not active_group_ids:
            st.caption("Aucun Groupe actif n'est encore créé. L'attribution à un Groupe est facultative.")
        managers = [user["username"] for user in users if user.get("role") in {"Responsable", "Admin"}]
        manager_id = st.selectbox(
            "Responsable direct",
            managers,
            index=None,
            placeholder="Sélectionnez un Responsable ou un Administrateur",
        )
    if "Responsable" in new_role_tags:
        managed_group_ids = st.multiselect(
            "Groupes gérés", active_group_ids, format_func=lambda group_id: group_names[group_id]
        )
    if st.form_submit_button("Créer l'utilisateur", type="primary"):
        payload = {
            "new_username": new_username,
            "new_mail": new_email,
            "new_password": new_password,
            "role_tags": new_role_tags,
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
                "Étiquettes": ", ".join(user.get("role_tags", [user["role"]])),
                "Fonction": organization_function(user),
                "Responsable": user.get("manager_id") or "",
                "Groupes": ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in user.get("group_ids", [])),
                "Groupes gérés": ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in user.get("managed_group_ids", [])),
                "E-mail": user["email"],
            }
        )
    st.dataframe(rows, width="stretch", hide_index=True)

    selected_username = st.selectbox("Modifier l'organisation d'un utilisateur", [user["username"] for user in users])
    selected_user = next(user for user in users if user["username"] == selected_username)
    updated_role_tags = st.multiselect(
        "Étiquettes attribuées",
        ["Employe", "Responsable", "Admin"],
        default=selected_user.get("role_tags", [selected_user["role"]]),
        key="admin_updated_role",
    )
    with st.form("update_user_organization_form"):
        updated_employee_type = selected_user.get("employee_type") or "salarie"
        updated_group_ids = [group_id for group_id in selected_user.get("group_ids", []) if group_id in active_group_id_set]
        updated_managed_group_ids = [group_id for group_id in selected_user.get("managed_group_ids", []) if group_id in active_group_id_set]
        updated_manager_id = selected_user.get("manager_id")
        if "Employe" in updated_role_tags:
            updated_employee_type = st.selectbox(
                "Type de personnel", ["salarie", "stagiaire"], index=["salarie", "stagiaire"].index(updated_employee_type)
            )
            updated_group_ids = st.multiselect(
                "Groupes d'appartenance (facultatif)", active_group_ids, default=updated_group_ids, format_func=lambda group_id: group_names[group_id]
            )
            if not active_group_ids:
                st.caption("Aucun Groupe actif n'est encore créé. L'attribution à un Groupe est facultative.")
            manager_options = [user["username"] for user in users if user.get("role") in {"Responsable", "Admin"}]
            manager_index = manager_options.index(updated_manager_id) if updated_manager_id in manager_options else None
            updated_manager_id = st.selectbox(
                "Responsable direct",
                manager_options,
                index=manager_index,
                placeholder="Sélectionnez un Responsable ou un Administrateur",
            )
        if "Responsable" in updated_role_tags:
            updated_managed_group_ids = st.multiselect(
                "Groupes gérés", active_group_ids, default=updated_managed_group_ids, format_func=lambda group_id: group_names[group_id]
            )
        if st.form_submit_button("Enregistrer l'organisation"):
            response = requests.put(
                f"{API_URL}/users/{selected_username}/organization",
                headers=HEADERS,
                json={
                    "role_tags": updated_role_tags,
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
