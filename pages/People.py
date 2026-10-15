"""Mon équipe entry: authentication, scope selection, then delegated rendering."""
import streamlit as st
from access_control import role_tags
from organization_ui import render_organization
from team_api import fetch_directory
from team_ui import department_rows, personnel_rows, render_team_table

user = st.session_state.get('user')
if not user:
    st.error('Veuillez vous connecter.')
    st.stop()

tags = role_tags(user)
st.title('Mon équipe')
organization = fetch_directory('organization-directory')
if organization is not None:
    render_organization(organization['users'], organization['groups'])
st.divider()

if not {'Admin', 'Responsable'}.intersection(tags):
    directory = fetch_directory('my-team')
    if directory is not None:
        render_team_table(department_rows(directory), title='Mon département',
                          empty_message='Vous n’êtes affecté à aucun département actif.')
else:
    users = fetch_directory('list-users')
    groups = fetch_directory('groups')
    if users is not None and groups is not None:
        st.caption('Les utilisateurs affichés appartiennent à votre périmètre de responsabilité.')
        if 'Admin' in tags:
            st.info('La gestion complète des comptes et des Groupes est disponible dans Administration.')
        render_team_table(personnel_rows(users.get('users', []), groups.get('groups', [])),
                          title='Employés gérés',
                          empty_message='Aucun Employé ne fait actuellement partie de votre périmètre.')
    if 'Responsable' in tags and 'Admin' not in tags:
        st.caption('Les comptes et contrats sont gérés par un administrateur.')
        st.page_link('pages/Fiches.py', label='Créer / remplir une fiche de présence',
                     icon=':material/calendar_month:')
