"""Contract history editor used by the administration page."""
from datetime import date
import requests
import streamlit as st

TYPES = {'salarie': 'Salarié', 'stagiaire': 'Stagiaire', 'alternance': 'Alternance'}


def contract_editor(api, headers, username):
    response = requests.get(f'{api}/users/{username}/contracts', headers=headers, timeout=15)
    if response.status_code != 200:
        st.error('Impossible de charger les contrats.')
        return
    contracts = response.json()['contracts']
    options = [None] + [c['id'] for c in contracts]
    by_id = {c['id']: c for c in contracts}
    selected = st.selectbox('Contrat à gérer', options,
        format_func=lambda key: 'Nouveau contrat' if key is None else f"{TYPES.get(by_id[key]['employee_type'], by_id[key]['employee_type'])} · {by_id[key]['start_date']} → {by_id[key]['end_date'] or 'CDI'}",
        key=f'contract_choice_{username}')
    current = by_id.get(selected, {})
    suffix = f'{username}_{selected}'
    st.caption('Les contrats historiques sont conservés. Les périodes ne doivent pas se chevaucher.')
    with st.form(f'contract_history_{suffix}'):
        kind = st.selectbox('Type de contrat', list(TYPES),
            index=list(TYPES).index(current.get('employee_type', 'salarie')),
            format_func=TYPES.get)
        start = st.date_input('Début de contrat', value=date.fromisoformat(current['start_date']) if current else None)
        cdi = st.checkbox('CDI / sans date de fin', value=bool(current.get('is_cdi')))
        end = st.date_input('Fin de contrat (laisser vide pour un CDI)', value=date.fromisoformat(current['end_date']) if current.get('end_date') else None)
        submitted = st.form_submit_button('Enregistrer le contrat', type='primary')
    if submitted:
        if start is None:
            st.error('Indiquez une date de début.')
            return
        payload = {'employee_type':kind,'start_date':start.isoformat(), 'end_date':end.isoformat() if end else None,'is_cdi':cdi}
        url = f'{api}/users/{username}/contracts' + (f'/{selected}' if selected is not None else '')
        result = requests.request('PATCH' if selected is not None else 'POST', url, headers=headers,json=payload,timeout=15)
        if result.status_code == 200:
            st.success('Contrat enregistré.')
            st.rerun()
        else:
            st.error(result.json().get('detail','Impossible de sauvegarder.'))
