"""Transport receipt uploads and archive management."""

from __future__ import annotations

import base64
from datetime import datetime

import requests
import streamlit as st

from access_control import role_tags
from api_client import api_url, authenticated_headers, safe_api_request


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
    response = safe_api_request(
        "GET",
        f"{API_URL}/transport-receipts",
        headers=HEADERS,
        params={"include_archived": "true"},
    )
    if response is None:
        st.stop()
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

assignees_response = safe_api_request("GET", f"{API_URL}/transport-receipts/assignees", headers=HEADERS)
if assignees_response is None:
    st.stop()
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
                response = safe_api_request(
                    "POST", f"{API_URL}/transport-receipts", headers=HEADERS, json=payload, timeout=30
                )
                if response is not None and response.status_code == 200:
                    st.success("Justificatif enregistré.")
                    st.rerun()
                elif response is not None:
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
            download_key = f"receipt_download_{receipt_id}"
            if download_key in st.session_state:
                st.download_button(
                    "Télécharger",
                    data=st.session_state[download_key],
                    file_name=receipt["original_filename"],
                    mime=receipt["mime_type"],
                    key=f"download_receipt_{receipt_id}",
                )
            elif st.button("Préparer", key=f"prepare_receipt_{receipt_id}"):
                file_response = safe_api_request(
                    "GET", f"{API_URL}/transport-receipts/{receipt_id}/file", headers=HEADERS, timeout=30
                )
                if file_response is not None and file_response.status_code == 200:
                    st.session_state[download_key] = file_response.content
                    st.rerun()
                elif file_response is not None:
                    st.error(response_detail(file_response))
        if "Admin" in USER_TAGS:
            with archive_column:
                if not archived and st.button("Archiver", key=f"archive_receipt_{receipt_id}"):
                    response = safe_api_request(
                        "PATCH", f"{API_URL}/transport-receipts/{receipt_id}/archive", headers=HEADERS
                    )
                    if response is not None and response.status_code == 200:
                        st.rerun()
                    if response is not None:
                        st.error(response_detail(response))
        with delete_column:
            may_delete = "Admin" in USER_TAGS or receipt["uploaded_by"] == USER["name"]
            confirm_delete_key = f"confirm_delete_receipt_{receipt_id}"
            if may_delete and st.button("Supprimer", key=f"delete_receipt_{receipt_id}", type="secondary"):
                st.session_state[confirm_delete_key] = True
                st.rerun()

        if may_delete and st.session_state.get(confirm_delete_key):
            st.warning(
                "Cette action supprimera définitivement ce justificatif. "
                "Voulez-vous continuer ?"
            )
            confirm_column, cancel_column, _ = st.columns([1.5, 1.2, 5])
            with confirm_column:
                confirm_delete = st.button(
                    "Confirmer la suppression",
                    key=f"confirm_delete_button_{receipt_id}",
                    type="primary",
                )
            with cancel_column:
                cancel_delete = st.button("Annuler", key=f"cancel_delete_{receipt_id}")

            if cancel_delete:
                st.session_state.pop(confirm_delete_key, None)
                st.rerun()

            if confirm_delete:
                response = safe_api_request(
                    "DELETE", f"{API_URL}/transport-receipts/{receipt_id}", headers=HEADERS
                )
                st.session_state.pop(confirm_delete_key, None)
                if response is not None and response.status_code == 200:
                    st.session_state.pop(f"receipt_download_{receipt_id}", None)
                    st.rerun()
                if response is not None:
                    st.error(response_detail(response))
