import streamlit as st


if "user" not in st.session_state or st.session_state["user"] is None:
    st.warning("Veuillez vous connecter pour accéder à l’accueil.")
    st.stop()


st.title("Gestion des fiches")
st.caption("Attendance and internship allowance management")

st.subheader("Créer une nouvelle fiche")
st.write(
    "Créez une fiche mensuelle depuis une personne existante ou commencez une fiche vierge."
)

first_column, second_column = st.columns(2)
with first_column:
    st.info("Depuis une personne existante\n\nRéutilisez son identité et son planning par défaut.")
    if st.button('Choisir une personne', type='primary'):
        st.session_state['creation_source'] = 'Personne existante'
        st.session_state['management_initial_tab'] = 'Créer une fiche'
        st.switch_page('pages/Fiches.py')
with second_column:
    st.info("Fiche vierge\n\nSaisissez uniquement les données nécessaires pour ce mois.")
    if st.button('Créer une fiche vierge'):
        st.session_state['creation_source'] = 'Fiche vierge'
        st.session_state['management_initial_tab'] = 'Créer une fiche'
        st.switch_page('pages/Fiches.py')

st.subheader("Suivi")
workspace = st.session_state.get(f"workspace_{st.session_state['user']['name']}")
forms = workspace.get('forms', []) if workspace is not None else []
draft_count = sum(form.get('state') == 'draft' for form in forms) if workspace is not None else '—'
exported_count = sum(form.get('state') == 'exported' and not form.get('archived') for form in forms) if workspace is not None else '—'
archived_count = sum(bool(form.get('archived')) for form in forms) if workspace is not None else '—'
first_metric, second_metric, third_metric = st.columns(3)
first_metric.metric("Brouillons", draft_count)
second_metric.metric("Exportées non archivées", exported_count)
third_metric.metric("Archivées", archived_count)

if workspace is None:
    st.caption('Ouvrez la gestion des fiches pour charger vos données et les compteurs.')
