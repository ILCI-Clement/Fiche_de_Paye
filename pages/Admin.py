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
DEPARTMENT_BADGE_COLORS = ("blue", "green", "violet", "orange", "yellow", "red")


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


def department_ids(member: dict) -> list[int]:
    return sorted(
        {
            *(int(group_id) for group_id in (member.get("group_ids") or [])),
            *(int(group_id) for group_id in (member.get("managed_group_ids") or [])),
        }
    )


def department_labels(member: dict) -> list[tuple[int, str]]:
    return [(group_id, group_names[group_id]) for group_id in department_ids(member) if group_id in group_names]


def can_supervise(member: dict) -> bool:
    return bool({"Admin", "Responsable"} & user_tags(member))


users_by_name = {user["username"]: user for user in users}
reports_by_manager: dict[str, list[dict]] = {}
for candidate in users:
    manager_name = candidate.get("manager_id")
    if manager_name:
        reports_by_manager.setdefault(str(manager_name), []).append(candidate)


def available_employees(manager: dict) -> list[dict]:
    return sorted(
        [
            employee
            for employee in users
            if employee["username"] != manager["username"] and "Employe" in user_tags(employee)
        ],
        key=lambda employee: employee["username"].casefold(),
    )


def render_person_card(member: dict, ancestry: set[str]) -> None:
    member_name = member["username"]
    if member_name in ancestry:
        st.error(f"Boucle hiérarchique détectée pour {member_name}.")
        return

    with st.container(border=True):
        details, actions = st.columns([4, 1], vertical_alignment="center")
        with details:
            st.markdown(f"**{member_name}**")
            st.caption(f"{organization_function(member)} · {', '.join(member.get('role_tags', [member['role']]))}")
            with st.container(horizontal=True, gap="small"):
                for group_id, label in department_labels(member):
                    st.badge(label, color=DEPARTMENT_BADGE_COLORS[group_id % len(DEPARTMENT_BADGE_COLORS)])
        if can_supervise(member):
            with actions:
                with st.popover("Ajouter un employé", icon=":material/person_add:", key=f"assign_{member_name}"):
                    candidates = available_employees(member)
                    if not candidates:
                        st.caption("Aucun Employé existant ne peut être affecté.")
                    else:
                        candidate_names = [candidate["username"] for candidate in candidates]
                        selected_name = st.selectbox(
                            "Employé existant",
                            candidate_names,
                            key=f"assign_employee_{member_name}",
                        )
                        selected_employee = users_by_name[selected_name]
                        current_manager = selected_employee.get("manager_id") or "Aucun responsable"
                        st.caption(f"Responsable actuel : {current_manager}")
                        if st.button("Affecter", type="primary", key=f"confirm_assign_{member_name}"):
                            response = requests.patch(
                                f"{API_URL}/users/{selected_name}/direct-manager",
                                headers=HEADERS,
                                json={"manager_id": member_name},
                                timeout=10,
                            )
                            if response.status_code == 200:
                                st.success("Employé affecté au responsable.")
                                st.rerun()
                            else:
                                st.error(api_error(response))

        reports = reports_by_manager.get(member_name, [])
        for report in sorted(reports, key=lambda person: person["username"].casefold()):
            render_person_card(report, ancestry | {member_name})


with st.expander("Structure des équipes", expanded=True):
    st.caption("La structure suit le responsable direct. Les Groupes sont affichés comme départements de travail.")
    with st.container(border=True):
        st.subheader("Direction générale")
        root_managers = [
            user
            for user in users
            if can_supervise(user) and not user.get("manager_id")
        ]
        if root_managers:
            for manager in sorted(root_managers, key=lambda person: person["username"].casefold()):
                render_person_card(manager, set())
        else:
            st.info("Attribuez au moins un Responsable ou Administrateur pour afficher la structure.")

    employees_without_manager = [
        user
        for user in users
        if "Employe" in user_tags(user) and not user.get("manager_id")
    ]
    if employees_without_manager:
        with st.expander(f"Sans responsable direct · {len(employees_without_manager)} personne(s)", expanded=False):
            for employee in sorted(employees_without_manager, key=lambda person: person["username"].casefold()):
                render_person_card(employee, set())

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
