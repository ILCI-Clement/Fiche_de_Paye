"""Explicit local testing entry point; never imports the production login pages."""

import streamlit as st
from preview_storage import load_preview_workspace

st.set_page_config(page_title='Présences — test local', page_icon='📅', layout='wide')
st.session_state['_local_preview'] = True
st.session_state['user'] = {'name': 'local-preview', 'email': 'demo@example.invalid', 'role': 'Responsable'}
if 'workspace_local-preview' not in st.session_state:
    st.session_state['workspace_local-preview'] = load_preview_workspace()

st.sidebar.title('Présences')
st.sidebar.info('TEST LOCAL\n\nDonnées fictives. Aucune connexion à l’API de production.')
st.sidebar.caption('Les sauvegardes explicites sont conservées sur cet ordinateur. Les brouillons non sauvegardés restent liés à la session.')
navigation = st.navigation([
    st.Page('pages/Home.py', title='Accueil', icon=':material/home:', default=True),
    st.Page('pages/Fiches.py', title='Gestion des fiches', icon=':material/calendar_month:'),
])
navigation.run()
