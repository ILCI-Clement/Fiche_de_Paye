"""Month-scoped attendance selection for explicit download and signing actions."""
import streamlit as st


def render_batch_selection(sheets, *, username, month, year, label):
    context = (username, int(year), int(month))
    prefix = "fiche_batch_selected_"
    if st.session_state.get("fiche_batch_selection_context") != context:
        for key in list(st.session_state):
            if key.startswith(prefix):
                del st.session_state[key]
        st.session_state.pop("fiche_batch_export", None)
        st.session_state["fiche_batch_selection_context"] = context
        st.session_state["fiche_batch_panel_open"] = False

    keys = [f"{prefix}{username}_{year}_{month}_{sheet['id']}" for sheet in sheets]
    for key in list(st.session_state):
        if key.startswith(prefix) and key not in keys:
            del st.session_state[key]

    if not st.session_state.get("fiche_batch_panel_open"):
        if not st.button("Export / signature groupés", key="open_fiche_batch", disabled=not sheets, width="stretch"):
            return None
        st.session_state["fiche_batch_panel_open"] = True

    st.subheader("Sélection des fiches")
    st.caption("Cochez les fiches du mois à télécharger ou à envoyer pour signature. Aucun envoi automatique.")
    all_column, none_column, close_column = st.columns(3)
    with all_column:
        if st.button("Tout sélectionner", key="fiche_batch_select_all"):
            for key in keys:
                st.session_state[key] = True
    with none_column:
        if st.button("Tout désélectionner", key="fiche_batch_select_none"):
            for key in keys:
                st.session_state[key] = False
    with close_column:
        if st.button("Fermer", key="close_fiche_batch"):
            st.session_state["fiche_batch_panel_open"] = False
            st.session_state.pop("fiche_batch_export", None)
            return None

    columns = st.columns(2)
    selected = []
    for index, (sheet, key) in enumerate(zip(sheets, keys)):
        with columns[index % 2]:
            if st.checkbox(label(sheet, index), key=key):
                selected.append(sheet)
    st.caption(f"{len(selected)} fiche(s) sélectionnée(s) sur {len(sheets)}.")
    return selected
