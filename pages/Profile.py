import streamlit as st
import requests

headers = {
    "Authorization": f"Bearer {st.secrets['PRESENCE_TOKEN']}"
}

st.title("Mes Infos Personnelles")

current_user = st.session_state["user"]["name"]

with st.form("form_profile"):
    st.subheader("Identifiants")
    new_username = st.text_input("Nom d'utilisateur", value=current_user)
    new_email = st.text_input("Adresse e-mail", value=st.session_state["user"]["email"])

    st.subheader("Changer de mot de passe (optionnel)")
    new_pass = st.text_input("Nouveau mot de passe", type="password")
    confirm_pass = st.text_input("Confirmer le nouveau mot de passe", type="password")

    st.divider()
    st.subheader("Confirmation obligatoire")
    current_pass = st.text_input("Mot de passe actuel", type="password")

    submitted = st.form_submit_button("Enregistrer les modifications", type="primary")

    if submitted:
        if not current_pass:
            st.error("Veuillez renseigner votre mot de passe actuel pour valider.")
        elif new_pass and new_pass != confirm_pass:
            st.error("Les deux nouveaux mots de passe ne correspondent pas.")
        else:
            payload = {
                "current_password": current_pass,
                "new_username": new_username if new_username != current_user else None,
                "new_email": new_email if new_email else None,
                "new_password": new_pass if new_pass else None,
                "confirm_password": confirm_pass if confirm_pass else None
            }
            
            try:
                res = requests.put(f"{st.secrets['URL_PRESENCE']}/update_profile/{current_user}", json=payload, headers=headers)
                if res.status_code == 200:
                    data = res.json()
                    st.success(data["message"])
                    # Mise à jour de la session locale avec le nouveau nom
                    st.session_state["user"]["name"] = data["username"]
                    st.rerun()
                else:
                    st.error(res.json().get("detail", "Une erreur est survenue."))
            except Exception as e:
                st.error(f"Erreur de communication avec l'API : {e}")