import streamlit as st
import requests

# Secrets de streamlit
TOKEN = st.secrets["PRESENCE_TOKEN"]
API_URL = st.secrets["URL_PRESENCE"]

st.title("Page de Connexion")

username = st.text_input("Nom d'utilisateur", width=400)
password = st.text_input("Mot de passe", type="password", width=400)

if st.button("Se connecter"):
    try:
        # On demande à l'API de vérifier
        res = requests.post(f"{API_URL}/login", json={"username": username, "password": password})
        if res.status_code == 200:
            data = res.json()

            st.session_state["user"] = {
            "name": data["username"],
            "email": data["email"],
            "role": data["role"],
            "data": {}
            }

            st.success("Connexion réussie")
            st.rerun()
        else:
            st.error("Identifiants incorrects")
    except Exception as e:
        st.error(f"Erreur de connexion à l'API : {e}")

# Toggle pour afficher le formulaire de mot de passe oublié
forgot_tab = st.checkbox("Mot de passe oublié ?")

if forgot_tab:
    st.subheader("Récupération de compte")
    email_recup = st.text_input("Entrez votre e-mail professionnel", width=400)
    if st.button("Recevoir le lien de récupération"):
        if email_recup:
            res = requests.post(f"{API_URL}/forgot-password", json={"email": email_recup})
            st.info("Si l'adresse est associée à un compte, un lien vient de vous être envoyé par e-mail.")
        else:
            st.warning("Veuillez entrer une adresse e-mail.")