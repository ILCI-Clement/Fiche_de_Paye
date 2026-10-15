"""Team table projections and rendering, independent of endpoint selection."""
import streamlit as st
from access_control import role_tags


def department_rows(directory):
    names = {int(group['id']): group['name'] for group in directory['groups']}
    return [{
        'Nom': member['name'],
        'Département': ', '.join(names[i] for i in dict.fromkeys(
            member['group_ids'] + member['managed_group_ids']) if i in names),
        'Responsable de': ', '.join(names[i] for i in member['managed_group_ids'] if i in names),
    } for member in directory['members']]


def personnel_rows(users, groups):
    names = {int(group['id']): group['name'] for group in groups if group.get('is_active')}
    return [{
        'Utilisateur': person['username'],
        'Type': person.get('employee_type') or 'salarie',
        'Responsable': person.get('manager_id') or '',
        'Groupes': ', '.join(names.get(int(gid), str(gid)) for gid in person.get('group_ids', [])),
        'E-mail': person['email'],
    } for person in users if 'Employe' in role_tags(person)]


def render_team_table(rows, *, title, empty_message):
    st.subheader(title)
    if rows:
        st.dataframe(rows, hide_index=True, width='stretch')
    else:
        st.info(empty_message)
