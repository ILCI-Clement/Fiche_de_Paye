"""Operational dashboard for contract alerts and annual interview to-dos."""

from __future__ import annotations

from datetime import date

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


def api_error(response: requests.Response) -> str:
    try:
        return str(response.json().get("detail", "Une erreur est survenue."))
    except ValueError:
        return "Une erreur est survenue."


def as_date(value: object) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


response = requests.get(f"{API_URL}/dashboard", headers=HEADERS, timeout=15)
if response.status_code != 200:
    st.error(api_error(response))
    st.stop()

dashboard = response.json()
contracts = list(dashboard.get("contracts", []))
interviews = list(dashboard.get("interviews", []))
open_interviews = [item for item in interviews if not item.get("completed")]
is_admin = "Admin" in USER_TAGS
auto_send_enabled = bool(dashboard.get("contract_reminders_auto_send"))
contract_signature = "|".join(
    f"{item.get('employee_name')}:{item.get('contract_end_date')}" for item in contracts
)
contract_prompt_key = "contract_reminder_prompt_acknowledged"


def send_due_contract_reminders() -> requests.Response:
    return requests.post(f"{API_URL}/contract-reminders/send", headers=HEADERS, timeout=30)


@st.dialog("Envoyer les rappels de fin de contrat ?")
def confirm_contract_reminders() -> None:
    st.write(f"{len(contracts)} contrat(s) arrivent à échéance dans les 10 prochains jours.")
    st.caption("Les e-mails seront envoyés une seule fois aux administrateurs et aux responsables directs concernés.")
    send_column, cancel_column = st.columns(2)
    with send_column:
        if st.button("Envoyer maintenant", type="primary", width="stretch"):
            send_response = send_due_contract_reminders()
            if send_response.status_code == 200:
                result = send_response.json()
                st.session_state[contract_prompt_key] = contract_signature
                st.success(f"{result['sent']} e-mail(s) de rappel envoyé(s).")
                st.rerun()
            st.error(api_error(send_response))
    with cancel_column:
        if st.button("Pas maintenant", width="stretch"):
            st.session_state[contract_prompt_key] = contract_signature
            st.rerun()

st.title("Tableau de bord")
st.caption("Suivez les échéances contractuelles et les entretiens annuels à réaliser.")

metric_contracts, metric_interviews, metric_completed = st.columns(3)
metric_contracts.metric("Contrats à échéance (10 jours)", len(contracts))
metric_interviews.metric(f"Entretiens annuels à faire ({dashboard['year']})", len(open_interviews))
metric_completed.metric("Entretiens terminés", len(interviews) - len(open_interviews))

with st.expander("Contrats arrivant à échéance", expanded=True):
    description_column, settings_column = st.columns([3, 2])
    with description_column:
        st.caption("Les rappels sont destinés aux administrateurs et aux responsables directs, jamais à l'employé concerné.")
    with settings_column:
        if is_admin:
            desired_auto_send = st.toggle(
                "Envoi automatique des rappels",
                value=auto_send_enabled,
                key=f"contract_reminders_auto_send_{int(auto_send_enabled)}",
                help="Si cette option est désactivée, une confirmation est demandée dans ce tableau de bord avant tout envoi.",
            )
            if desired_auto_send != auto_send_enabled:
                setting_response = requests.patch(
                    f"{API_URL}/settings/contract-reminders",
                    headers=HEADERS,
                    json={"auto_send": desired_auto_send},
                    timeout=15,
                )
                if setting_response.status_code == 200:
                    st.rerun()
                st.error(api_error(setting_response))
            if st.button(
                "Envoyer les rappels maintenant",
                disabled=not contracts,
                width="stretch",
                help="Ouvre une confirmation avant l'envoi des rappels actuellement dus.",
            ):
                confirm_contract_reminders()
        else:
            st.caption("L'envoi est décidé par un administrateur.")
    if contracts:
        st.dataframe(
            [
                {
                    "Employé": item.get("employee_name", item["employee_username"]),
                    "Fin de contrat": as_date(item["contract_end_date"]).strftime("%d/%m/%Y"),
                    "Jours restants": item["days_remaining"],
                    "Responsable direct": item.get("manager_name") or item.get("manager_username") or "—",
                }
                for item in contracts
            ],
            hide_index=True,
            width="stretch",
        )
    else:
        st.success("Aucun contrat ne se termine dans les 10 prochains jours.")

if contracts and is_admin and not auto_send_enabled and st.session_state.get(contract_prompt_key) != contract_signature:
    confirm_contract_reminders()

with st.expander(f"Entretiens annuels {dashboard['year']}", expanded=True):
    st.caption("Une tâche est créée automatiquement chaque année pour chaque personne enregistrée. Mettez-la à jour après l'entretien.")
    if not interviews:
        st.info("Aucun entretien annuel ne relève de votre périmètre.")
    interview_columns = st.columns(3)
    for index, interview in enumerate(interviews):
        username = str(interview["employee_username"])
        completed = bool(interview.get("completed"))
        status = "Terminé" if completed else "À faire"
        due_date_label = as_date(interview["due_date"]).strftime("%d/%m/%Y")
        with interview_columns[index % len(interview_columns)]:
            with st.expander(f"{username} · {status}", expanded=False):
                st.caption(f"Date cible : {due_date_label}")
                with st.form(f"annual_interview_{username}_{dashboard['year']}"):
                    due_date = st.date_input(
                        "Date cible",
                        value=as_date(interview["due_date"]),
                        key=f"interview_due_{username}_{dashboard['year']}",
                    )
                    notes = st.text_area(
                        "Remarques",
                        value=str(interview.get("notes") or ""),
                        key=f"interview_notes_{username}_{dashboard['year']}",
                        placeholder="Ex. date proposée, points à aborder, compte rendu…",
                    )
                    completed_value = st.checkbox(
                        "Entretien réalisé",
                        value=completed,
                        key=f"interview_completed_{username}_{dashboard['year']}",
                    )
                    if st.form_submit_button("Enregistrer", type="primary"):
                        update_response = requests.patch(
                            f"{API_URL}/annual-interviews/{username}/{dashboard['year']}",
                            headers=HEADERS,
                            json={
                                "due_date": due_date.isoformat(),
                                "notes": notes,
                                "completed": completed_value,
                            },
                            timeout=15,
                        )
                        if update_response.status_code == 200:
                            st.rerun()
                        st.error(api_error(update_response))
