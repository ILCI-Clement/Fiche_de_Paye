"""Compact document cards and selection/download controls; page owns filters and authentication."""
from datetime import datetime
import json
import requests
import streamlit as st
from api_client import response_detail, safe_api_request
from period_selector import MONTH_NAMES
from document_downloads import build_document_archive


def format_created_at(value: object) -> str:
    try:
        return datetime.fromisoformat(str(value)).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(value)


def receipt_period(receipt: dict) -> tuple[int, int]:
    if receipt.get("receipt_month") and receipt.get("receipt_year"):
        return int(receipt["receipt_month"]), int(receipt["receipt_year"])
    created_at = datetime.fromisoformat(str(receipt.get("created_at")))
    return created_at.month, created_at.year


def render_receipt_card(receipt: dict, *, api_base, headers, user, user_tags) -> None:
    API_URL, HEADERS, USER, USER_TAGS = api_base, headers, user, user_tags
    receipt_id = int(receipt["id"])
    receipt_month, receipt_year = receipt_period(receipt)
    archived = bool(receipt.get("archived"))
    status = "Archivé" if archived else "À traiter"
    category = receipt.get("document_type")
    category_label = ("Transport" if category == "transport" else
                      f"Autre — {receipt.get('other_type') or 'Non précisé'}" if category == "autre" else
                      "Non renseigné")
    with st.container(border=True):
        st.checkbox("Sélectionner", key=f"receipt_selected_{receipt_id}")
        st.markdown(f"**{receipt['original_filename']} · {MONTH_NAMES[receipt_month - 1]} {receipt_year}**")
        st.badge(category_label, color="blue" if category == "transport" else "orange" if category == "autre" else "gray")
        st.caption(f"{receipt['employee_username']}")
        with st.container():
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
        confirm_delete_key = f"confirm_delete_receipt_{receipt_id}"
        with st.expander("Détails et actions", expanded=bool(st.session_state.get(confirm_delete_key))):
            st.caption(
                f"Déposé par : {receipt['uploaded_by']} · "
                f"{format_created_at(receipt.get('created_at'))} · {status}"
            )
            archive_column = st.container()
            delete_column = st.container()
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
                may_delete = "Admin" in USER_TAGS or ("Responsable" not in USER_TAGS and receipt["uploaded_by"] == USER["name"])
                confirm_delete_key = f"confirm_delete_receipt_{receipt_id}"
                if may_delete and st.button("Supprimer", key=f"delete_receipt_{receipt_id}", type="secondary"):
                    st.session_state[confirm_delete_key] = True
                    st.rerun()

            if "Admin" in USER_TAGS:
                with st.popover("Corriger le mois", use_container_width=True):
                    with st.form(f"receipt_period_form_{receipt_id}"):
                        corrected_month = st.selectbox(
                            "Mois", range(1, 13), index=receipt_month - 1,
                            format_func=lambda month: MONTH_NAMES[month - 1], key=f"receipt_edit_month_{receipt_id}",
                        )
                        corrected_year = st.number_input(
                            "Année", min_value=2000, max_value=2100, value=receipt_year,
                            key=f"receipt_edit_year_{receipt_id}",
                        )
                        if st.form_submit_button("Enregistrer le mois"):
                            try:
                                response = requests.patch(
                                    f"{API_URL}/transport-receipts/{receipt_id}/period", headers=HEADERS,
                                    json={"receipt_month": int(corrected_month), "receipt_year": int(corrected_year)}, timeout=15,
                                )
                            except requests.RequestException:
                                response = None
                                st.error("Le serveur est indisponible. La modification n'a pas été enregistrée.")
                            if response is not None and response.status_code == 200:
                                st.session_state["receipt_flash"] = (
                                    f"Mois corrigé : {MONTH_NAMES[corrected_month - 1]} {corrected_year}."
                                )
                                st.rerun()
                            elif response is not None:
                                st.error(response_detail(response))

            if may_delete and st.session_state.get(confirm_delete_key):
                st.warning(
                    "Cette action supprimera définitivement ce justificatif. "
                    "Voulez-vous continuer ?"
                )
                confirm_column, cancel_column = st.columns(2)
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


def render_selection_controls(visible_receipts, selection_context):
    if st.session_state.get("receipt_selection_context") != selection_context:
        for state_key in list(st.session_state):
            if state_key.startswith("receipt_selected_"):
                del st.session_state[state_key]
        st.session_state.pop("receipt_bulk_zip", None)
        st.session_state["receipt_selection_context"] = selection_context

    if visible_receipts:
        with st.container(horizontal=True):
            if st.button("Tout sélectionner", key="receipt_select_all"):
                for receipt in visible_receipts:
                    st.session_state[f"receipt_selected_{int(receipt['id'])}"] = True
            if st.button("Tout désélectionner", key="receipt_select_none"):
                for receipt in visible_receipts:
                    st.session_state[f"receipt_selected_{int(receipt['id'])}"] = False
        return st.container()


def render_bulk_download(bulk_download_slot, visible_receipts, selection_context, *, month, year, archived, api_base, headers):
    API_URL, HEADERS = api_base, headers

    if visible_receipts:
        selected_receipts = [
            receipt for receipt in visible_receipts
            if st.session_state.get(f"receipt_selected_{int(receipt['id'])}", False)
        ]
        archive_signature = json.dumps([selection_context, selected_receipts], sort_keys=True, default=str)
        prepared_archive = st.session_state.get("receipt_bulk_zip")
        if prepared_archive and prepared_archive["signature"] != archive_signature:
            st.session_state.pop("receipt_bulk_zip", None)
            prepared_archive = None
        with bulk_download_slot:
            st.caption(f"{len(selected_receipts)} document(s) sélectionné(s). Le ZIP contient uniquement cette sélection (100 Mo maximum).")
            if st.button("Préparer le ZIP", disabled=not selected_receipts, key="receipt_prepare_zip", type="primary"):
                st.session_state.pop("receipt_bulk_zip", None)
                prepared_archive = None
                try:
                    with st.spinner("Préparation des documents…"):
                        archive_data = build_document_archive(selected_receipts, API_URL, HEADERS)
                    prepared_archive = {"signature": archive_signature, "data": archive_data}
                    st.session_state["receipt_bulk_zip"] = prepared_archive
                except ValueError as exc:
                    st.error(str(exc))
            if prepared_archive:
                st.download_button(
                    "Télécharger le ZIP", data=prepared_archive["data"],
                    file_name=f"Documents_{year}-{month:02d}_{'archives' if archived else 'a-traiter'}.zip",
                    mime="application/zip", key="receipt_download_zip", on_click="ignore",
                )
