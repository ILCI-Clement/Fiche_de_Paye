import streamlit as st
# from datetime import date, datetime
# import requests
# import time
# from DocxGen import generer_docx_stagiaire
# from ExcelGen import remplir_fiche_paie
# import zipfile
# import io

st.set_page_config(page_title="Fiches de présences", layout="wide")

if "user" not in st.session_state:
    st.session_state.user = None

login_page = st.Page("pages/Login.py", title="Connexion")
fiches_page = st.Page("pages/Fiches.py", title="Création de Fiches")
profile_page = st.Page("pages/Profile.py", title="Infos Personnelles")
admin_page = st.Page("pages/Admin.py", title="Administration")

def get_pages_for_user():
    user = st.session_state["user"]

    if not user:
        return [login_page]
    
    role = user.get("role")
    
    if role == "Admin":
        return [profile_page, fiches_page, admin_page]

    elif role == "Responsable":
        return [profile_page, fiches_page]
    
    elif role == "Employe":
        return [profile_page]
    
    return [login_page]

pages = get_pages_for_user()
pg = st.navigation(pages)

if st.session_state["user"]:
    st.sidebar.write(f"Connecté en tant que : **{st.session_state['user']['name']}** ({st.session_state['user']['role']})")
    if st.sidebar.button("Se déconnecter"):
        st.session_state["user"] = None
        st.rerun()
    
pg.run()
