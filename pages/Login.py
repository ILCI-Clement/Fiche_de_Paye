import streamlit as st
import requests
from access_control import normalize_role
from api_client import api_url

# Secrets de streamlit
API_URL = api_url()

st.title("Page de Connexion")

reset_token = st.query_params.get("token")

if reset_token:
    st.subheader("Réinitialisation du mot de passe")
    st.caption("Choisissez un nouveau mot de passe pour votre compte.")
    new_password = st.text_input("Nouveau mot de passe", type="password", width=400)
    confirm_password = st.text_input("Confirmer le nouveau mot de passe", type="password", width=400)

    if st.button("Enregistrer le nouveau mot de passe", type="primary"):
        if not new_password or not confirm_password:
            st.warning("Veuillez renseigner et confirmer votre nouveau mot de passe.")
        elif new_password != confirm_password:
            st.error("Les deux mots de passe ne correspondent pas.")
        elif len(new_password) < 8:
            st.warning("Le mot de passe doit comporter au moins 8 caractères.")
        else:
            try:
                response = requests.post(
                    f"{API_URL}/reset-password",
                    json={"token": reset_token, "new_password": new_password},
                    timeout=10,
                )
                if response.status_code == 200:
                    st.query_params.clear()
                    st.success("Votre mot de passe a été réinitialisé. Vous pouvez maintenant vous connecter.")
                    st.rerun()
                else:
                    st.error("Ce lien est invalide ou a expiré. Veuillez demander un nouveau lien.")
            except requests.RequestException:
                st.error("La réinitialisation est momentanément indisponible. Veuillez réessayer plus tard.")

    st.info("Le lien de récupération est valable pendant 15 minutes et ne peut être utilisé qu'une seule fois.")
    st.stop()

username = st.text_input("Nom d'utilisateur", width=400)
password = st.text_input("Mot de passe", type="password", width=400)
remember_me = st.checkbox("Rester connecté sur cet appareil", value=True)

if st.button("Se connecter"):
    try:
        # Request account authentication from the Presence API.
        res = requests.post(
            f"{API_URL}/login",
            json={"username": username, "password": password, "remember_me": remember_me},
            timeout=10,
        )
        if res.status_code == 200:
            data = res.json()
            role = normalize_role(data)
            if role is None:
                st.error("Le rôle de ce compte n'est pas reconnu.")
                st.stop()

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
                "remember_token": data.get("remember_token"),
                "data": {},
            }
            if not remember_me:
                st.session_state["clear_remembered_login"] = True

            st.success("Connexion réussie")
            st.rerun()
        else:
            st.error("Identifiants incorrects")
    except requests.RequestException:
        st.error("Erreur de connexion à l'API. Veuillez réessayer plus tard.")

# Toggle pour afficher le formulaire de mot de passe oublié
forgot_tab = st.checkbox("Mot de passe oublié ?")

if forgot_tab:
    st.subheader("Récupération de compte")
    email_recup = st.text_input("Entrez votre e-mail professionnel", width=400)
    if st.button("Recevoir le lien de récupération"):
        if email_recup:
            try:
                requests.post(f"{API_URL}/forgot-password", json={"email": email_recup}, timeout=10)
                st.info("Si l'adresse est associée à un compte, un lien vient de vous être envoyé par e-mail.")
            except requests.RequestException:
                st.error("La demande est momentanément indisponible. Veuillez réessayer plus tard.")
        else:
            st.warning("Veuillez entrer une adresse e-mail.")
