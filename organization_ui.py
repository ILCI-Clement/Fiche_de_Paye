"""Production organization rendering; no API calls or preview fixtures."""
import streamlit as st
from organization_cards import organization_html
from organization_model import build_organization


def render_organization(users, groups):
    model = build_organization(users, groups)
    with st.expander('Organigramme de l’entreprise', expanded=False):
        with st.container(horizontal=True):
            st.badge(f"{len(model['departments'])} départements actifs", color='blue')
            st.badge(f"{model['people_count']} personnes", color='gray')
        st.html(organization_html(model))
        if model['unassigned']:
            st.subheader('Personnes sans département actif')
            st.dataframe(model['unassigned'], hide_index=True, width='stretch')
        if model['issues']:
            with st.expander('Affectations à vérifier', expanded=False):
                for issue in model['issues']:
                    st.warning(issue)
        return model
