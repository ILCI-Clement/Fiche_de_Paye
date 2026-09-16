"""Administration of accounts, departments, and reporting relationships."""

from __future__ import annotations

import requests
import streamlit as st

from access_control import role_tags
from api_client import api_url, authenticated_headers


API_URL = api_url()
HEADERS = authenticated_headers()
CURRENT_USER = st.session_state.get("user")
IDENTITY_TAGS = ["Employe", "Responsable", "Admin"]

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


def user_tags(user: dict) -> set[str]:
    return set(user.get("role_tags", [user.get("role")]))


def organization_function(user: dict) -> str:
    tags = user_tags(user)
    if "Responsable" in tags:
        return "Responsable"
    if "Employe" in tags:
        return f"Employé · {user.get('employee_type') or 'salarie'}"
    if "Admin" in tags:
        return "Administrateur"
    return "—"


st.title("Administration")
st.caption("Consultez un compte, puis gérez ses informations et ses relations organisationnelles.")

groups = fetch_groups()
users = fetch_users()
users_by_name = {user["username"]: user for user in users}
group_names = {int(group["id"]): str(group["name"]) for group in groups}
active_group_ids = [int(group["id"]) for group in groups if group.get("is_active")]


def department_names(user: dict, relation: str) -> str:
    key = "group_ids" if relation == "member" else "managed_group_ids"
    return ", ".join(group_names.get(int(group_id), str(group_id)) for group_id in (user.get(key) or [])) or "—"


def manager_options(exclude_username: str | None = None) -> list[str]:
    return [
        user["username"]
        for user in users
        if user["username"] != exclude_username and {"Admin", "Responsable"} & user_tags(user)
    ]


def active_memberships(user: dict, relation: str) -> list[int]:
    key = "group_ids" if relation == "member" else "managed_group_ids"
    return [int(group_id) for group_id in (user.get(key) or []) if int(group_id) in active_group_ids]


def organization_payload(
    user: dict,
    *,
    tags: list[str] | None = None,
    employee_type: str | None = None,
    group_ids: list[int] | None = None,
    managed_group_ids: list[int] | None = None,
    manager_id: str | None = None,
) -> dict:
    selected_tags = tags if tags is not None else list(user.get("role_tags", [user["role"]]))
    return {
        "role_tags": selected_tags,
        "employee_type": employee_type or user.get("employee_type") or "salarie",
        "group_ids": active_memberships(user, "member") if group_ids is None else group_ids,
        "managed_group_ids": active_memberships(user, "manager") if managed_group_ids is None else managed_group_ids,
        "manager_id": user.get("manager_id") if manager_id is None else manager_id,
    }


def save_organization(username: str, payload: dict) -> None:
    response = requests.put(
        f"{API_URL}/users/{username}/organization",
        headers=HEADERS,
        json=payload,
        timeout=10,
    )
    if response.status_code == 200:
        st.success("Informations enregistrées.")
        st.rerun()
    st.error(api_error(response))


with st.expander("Utilisateurs enregistrés", expanded=True):
    with st.container(horizontal=True, horizontal_alignment="distribute", vertical_alignment="center"):
        st.caption(f"{len(users)} utilisateur(s) enregistré(s). Cliquez sur une ligne pour ouvrir sa fiche.")
        with st.popover("Ajouter un utilisateur", icon=":material/person_add:"):
            with st.form("create_user_form", clear_on_submit=True):
                new_username = st.text_input("Nom d'utilisateur")
                new_email = st.text_input("E-mail professionnel")
                new_password = st.text_input("Mot de passe", type="password")
                new_tags = st.multiselect("Étiquettes d'identité", IDENTITY_TAGS, default=["Employe"])
                new_employee_type = st.selectbox("Type de personnel", ["salarie", "stagiaire"])
                new_group_ids = st.multiselect(
                    "Départements de travail", active_group_ids, format_func=lambda group_id: group_names[group_id]
                )
                new_managed_group_ids = st.multiselect(
                    "Départements gérés", active_group_ids, format_func=lambda group_id: group_names[group_id]
                )
                new_manager_id = st.selectbox(
                    "Responsable direct", manager_options(), index=None,
                    placeholder="Sélectionnez un responsable si nécessaire",
                )
                if st.form_submit_button("Créer l'utilisateur", type="primary"):
                    response = requests.post(
                        f"{API_URL}/create-user",
                        headers=HEADERS,
                        json={
                            "new_username": new_username,
                            "new_mail": new_email,
                            "new_password": new_password,
                            "role_tags": new_tags,
                            "employee_type": new_employee_type,
                            "group_ids": new_group_ids,
                            "managed_group_ids": new_managed_group_ids,
                            "manager_id": new_manager_id,
                        },
                        timeout=10,
                    )
                    if response.status_code == 200:
                        st.success("Utilisateur créé.")
                        st.rerun()
                    else:
                        st.error(api_error(response))

    overview_rows = [
        {
            "Utilisateur": user["username"],
            "Fonction": organization_function(user),
            "Département": department_names(user, "member"),
            "Responsable direct": user.get("manager_id") or "—",
        }
        for user in users
    ]
    selection = st.dataframe(
        overview_rows,
        width="stretch",
        hide_index=True,
        key="admin_user_overview",
        on_select="rerun",
        selection_mode="single-row",
    )
    if selection.selection.rows:
        st.session_state["admin_selected_username"] = overview_rows[selection.selection.rows[0]]["Utilisateur"]

selected_username = st.session_state.get("admin_selected_username")
if selected_username not in users_by_name:
    selected_username = users[0]["username"] if users else None
    st.session_state["admin_selected_username"] = selected_username

if selected_username:
    selected_user = users_by_name[selected_username]
    st.subheader(f"Fiche de {selected_username}")

    with st.expander("Informations de base", expanded=True):
        st.table(
            {
                "Nom d'utilisateur": selected_user["username"],
                "E-mail professionnel": selected_user["email"],
                "Type de personnel": selected_user.get("employee_type") or "—",
            },
            border="horizontal",
            width="content",
        )

    with st.expander("Fonction et étiquettes", expanded=False):
        with st.form(f"identity_{selected_username}"):
            edited_tags = st.multiselect(
                "Étiquettes d'identité", IDENTITY_TAGS,
                default=list(selected_user.get("role_tags", [selected_user["role"]])),
            )
            edited_employee_type = st.selectbox(
                "Type de personnel", ["salarie", "stagiaire"],
                index=["salarie", "stagiaire"].index(selected_user.get("employee_type") or "salarie"),
            )
            if st.form_submit_button("Enregistrer la fonction", type="primary"):
                save_organization(
                    selected_username,
                    organization_payload(selected_user, tags=edited_tags, employee_type=edited_employee_type),
                )

    with st.expander("Départements", expanded=False):
        with st.form(f"departments_{selected_username}"):
            edited_group_ids = st.multiselect(
                "Départements de travail", active_group_ids,
                default=active_memberships(selected_user, "member"),
                format_func=lambda group_id: group_names[group_id],
            )
            edited_managed_group_ids = st.multiselect(
                "Départements gérés", active_group_ids,
                default=active_memberships(selected_user, "manager"),
                format_func=lambda group_id: group_names[group_id],
            )
            if st.form_submit_button("Enregistrer les départements", type="primary"):
                save_organization(
                    selected_username,
                    organization_payload(
                        selected_user,
                        group_ids=edited_group_ids,
                        managed_group_ids=edited_managed_group_ids,
                    ),
                )

    with st.expander("Organisation", expanded=False):
        direct_manager = selected_user.get("manager_id") or "Sans responsable direct"
        st.markdown(f"**Responsable direct :** {direct_manager}")
        direct_reports = [user for user in users if user.get("manager_id") == selected_username]
        if direct_reports:
            st.markdown("**Employés sous responsabilité :**")
            st.dataframe(
                [
                    {
                        "Utilisateur": user["username"],
                        "Fonction": organization_function(user),
                        "Département": department_names(user, "member"),
                    }
                    for user in direct_reports
                ],
                width="stretch",
                hide_index=True,
            )
        else:
            st.caption("Aucun Employé n'est actuellement rattaché à cette personne.")

        if {"Admin", "Responsable"} & user_tags(selected_user):
            with st.popover("Ajouter un employé", icon=":material/person_add:"):
                employees = [
                    user for user in users
                    if user["username"] != selected_username and "Employe" in user_tags(user)
                ]
                if employees:
                    employee_names = [
                        user["username"] for user in sorted(employees, key=lambda user: user["username"].casefold())
                    ]
                    employee_name = st.selectbox("Employé existant", employee_names)
                    current_manager = users_by_name[employee_name].get("manager_id") or "Aucun responsable"
                    st.caption(f"Responsable actuel : {current_manager}")
                    if st.button("Affecter", type="primary", key=f"assign_{selected_username}"):
                        response = requests.patch(
                            f"{API_URL}/users/{employee_name}/direct-manager",
                            headers=HEADERS,
                            json={"manager_id": selected_username},
                            timeout=10,
                        )
                        if response.status_code == 200:
                            st.success("Employé affecté au responsable.")
                            st.rerun()
                        else:
                            st.error(api_error(response))
                else:
                    st.caption("Aucun Employé existant ne peut être affecté.")

    with st.expander("Actions du compte", expanded=False):
        if selected_username == CURRENT_USER["name"]:
            st.caption("Votre propre compte ne peut pas être supprimé ici.")
        else:
            confirmation = st.checkbox(
                "Je confirme la suppression du compte. Les fiches existantes sont conservées.",
                key=f"delete_confirmation_{selected_username}",
            )
            if st.button("Supprimer le compte", disabled=not confirmation, key=f"delete_{selected_username}"):
                response = requests.delete(
                    f"{API_URL}/delete-user/{selected_username}", headers=HEADERS, timeout=10
                )
                if response.status_code == 200:
                    st.session_state["admin_selected_username"] = None
                    st.success("Compte supprimé.")
                    st.rerun()
                else:
                    st.error(api_error(response))

with st.expander("Gérer les départements", expanded=False):
    with st.container(horizontal=True, horizontal_alignment="distribute", vertical_alignment="center"):
        st.caption("Les départements servent à identifier le périmètre de travail; ils ne déterminent pas la hiérarchie.")
        with st.popover("Ajouter un département", icon=":material/add:"):
            with st.form("create_group_form", clear_on_submit=True):
                new_group_name = st.text_input("Nom du département")
                if st.form_submit_button("Créer le département", type="primary"):
                    response = requests.post(
                        f"{API_URL}/groups", headers=HEADERS, json={"name": new_group_name}, timeout=10
                    )
                    if response.status_code == 200:
                        st.success("Département créé.")
                        st.rerun()
                    else:
                        st.error(api_error(response))

    if groups:
        selected_group_id = st.selectbox(
            "Département à modifier", [int(group["id"]) for group in groups],
            format_func=lambda group_id: group_names[group_id],
        )
        selected_group = next(group for group in groups if int(group["id"]) == selected_group_id)
        with st.form("update_group_form"):
            edited_group_name = st.text_input("Nom du département", value=str(selected_group["name"]))
            edited_group_active = st.checkbox("Département actif", value=bool(selected_group["is_active"]))
            if st.form_submit_button("Enregistrer le département", type="primary"):
                response = requests.put(
                    f"{API_URL}/groups/{selected_group_id}", headers=HEADERS,
                    json={"name": edited_group_name, "is_active": edited_group_active}, timeout=10,
                )
                if response.status_code == 200:
                    st.success("Département mis à jour.")
                    st.rerun()
                else:
                    st.error(api_error(response))
    else:
        st.info("Créez un département lorsqu'il sera nécessaire.")
