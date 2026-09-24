import requests
import streamlit as st
from access_control import normalize_role, role_tags
from api_client import api_url
from remember_me import remembered_token
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

clear_remembered_login = bool(st.session_state.get("clear_remembered_login"))
remember_token = remembered_token(
    token_to_store=(st.session_state["user"] or {}).get("remember_token"),
    clear=clear_remembered_login,
)
st.session_state["clear_remembered_login"] = False

if not clear_remembered_login and not st.session_state["user"] and remember_token:
    try:
        restored = requests.post(
            f"{api_url()}/restore-session",
            json={"remember_token": remember_token},
            timeout=10,
        )
        if restored.status_code == 200:
            data = restored.json()
            role = normalize_role(data)
            st.session_state["user"] = {
                "name": data["username"],
                "email": data["email"],
                "role": role,
                "role_tags": data.get("role_tags", [role]),
                "is_admin": bool(data.get("is_admin")),
                "id": data.get("id", data.get("user_id")),
                "managed_group_ids": data.get("managed_group_ids", data.get("groups", [])),
                "group_ids": data.get("group_ids", []),
                "auth_token": data.get("auth_token"),
                "remember_token": remember_token,
                "data": {},
            }
            st.rerun()
    except requests.RequestException:
        pass

if not st.session_state["user"]:
    st.markdown("<style>[data-testid='stSidebar']{display:none;}</style>", unsafe_allow_html=True)

login_page = st.Page("pages/Login.py", title="Connexion")
fiches_page = st.Page("pages/Fiches.py", title="Création de Fiches")
profile_page = st.Page("pages/Profile.py", title="Infos Personnelles")
admin_page = st.Page("pages/Admin.py", title="Administration")
people_page = st.Page("pages/People.py", title="Personnel")
transport_receipts_page = st.Page("pages/TransportReceipts.py", title="Justificatifs de transport")
dashboard_page = st.Page("pages/Dashboard.py", title="Tableau de bord")

def get_pages_for_user():
    user = st.session_state["user"]

    if not user:
        return [login_page]
    
    role = normalize_role(user)
    tags = role_tags(user)
    if role is None:
        return [login_page]
    user["role"] = role
    
    if "Admin" in tags:
        return [dashboard_page, profile_page, fiches_page, transport_receipts_page, people_page, admin_page]

    if "Responsable" in tags:
        return [dashboard_page, profile_page, fiches_page, transport_receipts_page, people_page]
    
    if "Employe" in tags:
        return [profile_page, transport_receipts_page]
    
    return [login_page]

pages = get_pages_for_user()
pg = st.navigation(pages)

if st.session_state["user"]:
    st.sidebar.write(f"Connecté en tant que : **{st.session_state['user']['name']}** ({st.session_state['user']['role']})")
    if st.sidebar.button("Se déconnecter"):
        logout_remember_token = remember_token or st.session_state["user"].get("remember_token")
        if logout_remember_token:
            try:
                requests.post(
                    f"{api_url()}/forget-session",
                    json={"remember_token": logout_remember_token},
                    timeout=10,
                )
            except requests.RequestException:
                pass
        st.session_state["clear_remembered_login"] = True
        st.session_state["user"] = None
        st.rerun()
    
pg.run()
