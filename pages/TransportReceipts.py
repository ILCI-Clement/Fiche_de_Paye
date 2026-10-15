"""Transport receipt uploads and archive management."""

from __future__ import annotations

import base64
from datetime import datetime

import requests
import streamlit as st
from period_selector import select_period, MONTH_NAMES
from document_ui import receipt_period, render_receipt_card, render_selection_controls, render_bulk_download

from access_control import role_tags
from api_client import api_url, authenticated_headers, safe_api_request
from api_client import response_detail


API_URL = api_url()
HEADERS = authenticated_headers()
USER = st.session_state.get("user")
USER_TAGS = role_tags(USER) if USER else set()

if not USER:
    st.warning("Veuillez vous connecter d'abord.")
    st.stop()


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


def belongs_to_selected_period(receipt: dict) -> bool:
    try:
        return receipt_period(receipt) == (selected_month, selected_year)
    except (ValueError, TypeError):
        return False


st.title("Documents")
selected_month, selected_year = select_period("transport", "top", datetime.now().month, datetime.now().year)
st.caption("Filtre selon le mois de rattachement du document, indépendamment de sa date de dépôt.")
st.caption("Les fichiers acceptés sont PDF, JPG et PNG, jusqu'à 10 Mo.")

if st.session_state.get("receipt_flash"):
    st.success(st.session_state.pop("receipt_flash"))

IS_ADMIN = "Admin" in USER_TAGS
assignees = [USER["name"]]
if IS_ADMIN:
    assignees_response = safe_api_request("GET", f"{API_URL}/transport-receipts/assignees", headers=HEADERS)
    if assignees_response is None:
        st.stop()
    if assignees_response.status_code != 200:
        st.error(response_detail(assignees_response))
        st.stop()
    assignees = list(assignees_response.json().get("users", []))

if assignees:
    with st.expander("Ajouter un justificatif", expanded=True):
        document_type = st.segmented_control(
            "Type de document", ["Transport", "Autre"], default="Transport", key="receipt_document_type"
        ) or "Transport"
        page_period = (selected_month, selected_year)
        reset_upload_period = st.session_state.pop("receipt_reset_upload_period", False)
        if st.session_state.get("receipt_upload_period_context") != page_period or reset_upload_period:
            st.session_state["receipt_upload_period_context"] = page_period
            st.session_state["receipt_custom_period"] = False
            st.session_state["receipt_upload_month"] = selected_month
            st.session_state["receipt_upload_year"] = selected_year
        upload_month, upload_year = selected_month, selected_year
        if st.checkbox("Modifier le mois de rattachement", key="receipt_custom_period"):
            month_column, year_column = st.columns(2)
            with month_column:
                upload_month = st.selectbox(
                    "Mois du document", range(1, 13), format_func=lambda month: MONTH_NAMES[month - 1],
                    key="receipt_upload_month",
                )
            with year_column:
                upload_year = st.number_input(
                    "Année du document", min_value=2000, max_value=2100, key="receipt_upload_year"
                )
        upload_period_label = f"{MONTH_NAMES[upload_month - 1]} {upload_year}"
        st.info(f"Ce document sera classé dans : {upload_period_label}.")
        with st.form("transport_receipt_upload", clear_on_submit=True):
            if IS_ADMIN:
                employee_username = st.selectbox("Employé concerné", assignees)
            else:
                employee_username = USER["name"]
                st.caption(f"Document personnel · {employee_username}")
            other_type = ""
            if document_type == "Autre":
                other_type = st.text_input(
                    "Précisez le type de document", max_chars=120, placeholder="Ex. : repas, matériel…"
                )
            uploaded_file = st.file_uploader("Fichier", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=False)
            submitted = st.form_submit_button(f"Envoyer pour {upload_period_label}", type="primary")

            if submitted:
                if uploaded_file is None:
                    st.error("Sélectionnez un fichier PDF, JPG ou PNG.")
                elif document_type == "Autre" and not other_type.strip():
                    st.error("Précisez le type de document avant de l'envoyer.")
                else:
                    payload = {
                        "employee_username": employee_username,
                        "receipt_month": int(upload_month),
                        "receipt_year": int(upload_year),
                        "document_type": document_type.lower(),
                        "other_type": other_type.strip() if document_type == "Autre" else None,
                        "original_filename": uploaded_file.name,
                        "file_b64": base64.b64encode(uploaded_file.getvalue()).decode("utf-8"),
                    }
                    try:
                        response = requests.post(
                            f"{API_URL}/transport-receipts", headers=HEADERS, json=payload, timeout=30
                        )
                    except requests.RequestException:
                        response = None
                        st.error("Le serveur est indisponible. Réessayez l'envoi du document.")
                    if response is not None and response.status_code == 200:
                        st.session_state["receipt_flash"] = f"Document enregistré dans {upload_period_label}."
                        st.session_state["receipt_reset_upload_period"] = True
                        st.rerun()
                    elif response is not None:
                        st.error(response_detail(response))

receipts = load_receipts()
current_month_count = sum(1 for receipt in receipts if belongs_to_selected_period(receipt))
select_period("transport", "bottom", selected_month, selected_year)
summary_column, folder_column, type_column = st.columns([1.6, 1.2, 1.2], gap="small")
with folder_column:
    document_folder = st.segmented_control(
        "Dossier", ["À traiter", "Archives"], default="À traiter", key="receipt_folder"
    ) or "À traiter"
show_archived = document_folder == "Archives"
with type_column:
    type_filter = st.selectbox(
        "Filtrer par type", ["Tous", "Transport", "Autre", "Non renseigné"], key="receipt_type_filter",
    )
visible_receipts = [
    receipt for receipt in receipts
    if belongs_to_selected_period(receipt)
    and bool(receipt.get("archived")) == show_archived
    and (type_filter == "Tous"
         or (type_filter == "Transport" and receipt.get("document_type") == "transport")
         or (type_filter == "Autre" and receipt.get("document_type") == "autre")
         or (type_filter == "Non renseigné" and receipt.get("document_type") not in {"transport", "autre"}))
]
with summary_column:
    list_title = "Archives" if show_archived else "Tous les justificatifs" if "Admin" in USER_TAGS else "Mes justificatifs"
    st.markdown(f"#### {list_title}")
    st.caption(f"Documents rattachés à {MONTH_NAMES[selected_month - 1]} {selected_year} · {current_month_count}")
    st.caption(f"{len(visible_receipts)} document(s) correspondant aux filtres.")

if not visible_receipts:
    st.info("Aucun justificatif à afficher.")

selection_context = (USER["name"], selected_year, selected_month, document_folder, type_filter)
bulk_download_slot = render_selection_controls(visible_receipts, selection_context)

for row_start in range(0, len(visible_receipts), 3):
    card_columns = st.columns(3, gap="small")
    for card_column, receipt in zip(card_columns, visible_receipts[row_start:row_start + 3]):
        with card_column:
            render_receipt_card(receipt, api_base=API_URL, headers=HEADERS, user=USER, user_tags=USER_TAGS)

render_bulk_download(
    bulk_download_slot, visible_receipts, selection_context,
    month=selected_month, year=selected_year, archived=show_archived,
    api_base=API_URL, headers=HEADERS,
)
