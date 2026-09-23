"""Transport receipt uploads and archive management."""

from __future__ import annotations

import base64
from datetime import datetime

import requests
import streamlit as st

from access_control import role_tags
from api_client import api_url, authenticated_headers


API_URL = api_url()
HEADERS = authenticated_headers()
USER = st.session_state.get("user")
USER_TAGS = role_tags(USER) if USER else set()

if not USER:
    st.warning("Veuillez vous connecter d'abord.")
    st.stop()


def response_detail(response: requests.Response) -> str:
    try:
        return str(response.json().get("detail", "Une erreur est survenue."))
    except ValueError:
        return "Une erreur est survenue."


def load_receipts() -> list[dict]:
    response = requests.get(
        f"{API_URL}/transport-receipts",
        headers=HEADERS,
        params={"include_archived": "true"},
        timeout=15,
    )
    if response.status_code != 200:
        st.error(response_detail(response))
        st.stop()
    return list(response.json().get("receipts", []))


def format_created_at(value: object) -> str:
    try:
        return datetime.fromisoformat(str(value)).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(value)


def created_this_month(receipt: dict) -> bool:
    try:
        created_at = datetime.fromisoformat(str(receipt.get("created_at")))
        now = datetime.now()
        return created_at.year == now.year and created_at.month == now.month
    except ValueError:
        return False


st.title("Justificatifs de transport")
st.caption("Déposez un justificatif de transport pour un utilisateur enregistré. Les fichiers acceptés sont PDF, JPG et PNG, jusqu'à 10 Mo.")

assignees_response = requests.get(f"{API_URL}/transport-receipts/assignees", headers=HEADERS, timeout=15)
if assignees_response.status_code != 200:
    st.error(response_detail(assignees_response))
    st.stop()
assignees = list(assignees_response.json().get("users", []))

with st.expander("Ajouter un justificatif", expanded=True):
    with st.form("transport_receipt_upload", clear_on_submit=True):
        employee_username = st.selectbox("Employé concerné", assignees)
        uploaded_file = st.file_uploader("Fichier", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=False)
        submitted = st.form_submit_button("Envoyer le justificatif", type="primary")

        if submitted:
            if uploaded_file is None:
                st.error("Sélectionnez un fichier PDF, JPG ou PNG.")
            else:
                payload = {
                    "employee_username": employee_username,
                    "original_filename": uploaded_file.name,
                    "file_b64": base64.b64encode(uploaded_file.getvalue()).decode("utf-8"),
                }
                try:
                    response = requests.post(f"{API_URL}/transport-receipts", headers=HEADERS, json=payload, timeout=30)
                except requests.RequestException as error:
                    st.error(f"Impossible de joindre le serveur : {error}")
                else:
                    if response.status_code == 200:
                        st.success("Justificatif enregistré.")
                        st.rerun()
                    else:
                        st.error(response_detail(response))

receipts = load_receipts()
current_month_count = sum(1 for receipt in receipts if created_this_month(receipt))
metric_label = "Justificatifs déposés ce mois"
st.metric(metric_label, current_month_count)

show_archived = st.checkbox("Afficher les justificatifs archivés", value=False)
visible_receipts = [receipt for receipt in receipts if show_archived or not receipt.get("archived")]

if "Admin" in USER_TAGS:
    st.subheader("Tous les justificatifs")
    st.caption("Vous pouvez consulter, archiver ou supprimer chaque justificatif.")
else:
    st.subheader("Mes justificatifs")
    st.caption("Vous pouvez consulter les justificatifs qui vous sont attribués et supprimer ceux que vous avez déposés.")

if not visible_receipts:
    st.info("Aucun justificatif à afficher.")

for receipt in visible_receipts:
    receipt_id = int(receipt["id"])
    archived = bool(receipt.get("archived"))
    status = "Archivé" if archived else "À traiter"
    with st.container(border=True):
        info_column, download_column, archive_column, delete_column = st.columns([5, 1.4, 1.2, 1.2])
        with info_column:
            st.markdown(f"**{receipt['original_filename']}**")
            st.caption(
                f"Employé concerné : {receipt['employee_username']} · "
                f"Déposé par : {receipt['uploaded_by']} · "
                f"{format_created_at(receipt.get('created_at'))} · {status}"
            )
        with download_column:
            try:
                file_response = requests.get(
                    f"{API_URL}/transport-receipts/{receipt_id}/file",
                    headers=HEADERS,
                    timeout=30,
                )
                if file_response.status_code == 200:
                    st.download_button(
                        "Télécharger",
                        data=file_response.content,
                        file_name=receipt["original_filename"],
                        mime=receipt["mime_type"],
                        key=f"download_receipt_{receipt_id}",
                    )
                else:
                    st.caption("Fichier indisponible")
            except requests.RequestException:
                st.caption("Fichier indisponible")
        if "Admin" in USER_TAGS:
            with archive_column:
                if not archived and st.button("Archiver", key=f"archive_receipt_{receipt_id}"):
                    response = requests.patch(
                        f"{API_URL}/transport-receipts/{receipt_id}/archive",
                        headers=HEADERS,
                        timeout=15,
                    )
                    if response.status_code == 200:
                        st.rerun()
                    st.error(response_detail(response))
        with delete_column:
            may_delete = "Admin" in USER_TAGS or receipt["uploaded_by"] == USER["name"]
            if may_delete and st.button("Supprimer", key=f"delete_receipt_{receipt_id}", type="secondary"):
                response = requests.delete(
                    f"{API_URL}/transport-receipts/{receipt_id}",
                    headers=HEADERS,
                    timeout=15,
                )
                if response.status_code == 200:
                    st.rerun()
                st.error(response_detail(response))
