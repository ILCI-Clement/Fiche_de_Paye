import streamlit as st
import requests

TOKEN = st.secrets["PRESENCE_TOKEN"]
API_URL = st.secrets["URL_PRESENCE"]

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

st.title("Panneau Administratif")

st.subheader("Créer un nouvel utilisateur")

new_user = st.text_input("Nom d'utilisateur", key="admin_new_user")
new_mail = st.text_input("E-mail professionnel (exemple@univ-ilci.fr)", key="admin_new_mail")
new_pass = st.text_input("Mot de passe", type="password", key="admin_new_pass")
new_is_admin = st.checkbox("Administrateur ?", value=False, help="Cochez pour que l'utilisateur soit administrateur")
new_role = st.radio("Role de l'utilisateur", ["Admin", "Responsable", "Employe"], help="Cochez le rôle que vous voulez attribuer")

if st.button("Créer l'utilisateur", type="primary"):
    if new_user and new_mail and new_pass:
        payload = {"new_username": new_user, "new_mail": new_mail, "new_password": new_pass, "is_admin": new_is_admin, "new_role": new_role}
        res = requests.post(f"{API_URL}/create-user", headers=headers, json=payload)
                    
        if res.status_code == 200:
            st.success(f"Compte '{new_user}' créé avec succès !")
        else:
            st.error(f"Erreur : {res.json().get('detail')}")
    else:
        st.warning("Veuillez remplir tous les champs.")

st.write("---") # Ligne de séparation

st.subheader("Liste des utilisateurs")
        
try:
    # Appel à l'API pour récupérer la liste
    res_list = requests.get(f"{API_URL}/list-users", headers=headers)
            
    if res_list.status_code == 200:
        donnees_users = res_list.json().get("users", [])
                
        if donnees_users:
            for u in donnees_users:
                # Affichage d'une ligne stylisée pour chaque utilisateur
                st.write(f"**{u['username']}** — *{u["role"]}* — {u['email']}")
        else:
            st.error("Impossible de récupérer la liste des utilisateurs.")
except Exception as e:
    st.error(f"Erreur de connexion à l'API : {e}")

st.write("---") # Ligne de séparation
        
st.subheader("Supprimer un utilisateur")

user_to_del = st.text_input("Nom de l'utilisateur à supprimer", key="admin_del_user")
        
# Sécurité Streamlit : On demande de cocher une case pour confirmer avant de cliquer
confirmer_suppression = st.checkbox("Je confirme vouloir supprimer définitivement cet utilisateur", key="confirm_del")
        
if st.button("Supprimer l'utilisateur", type="secondary", disabled=not confirmer_suppression):
    if user_to_del:
        # Sécurité : Éviter que l'admin connecté se supprime lui-même
        if user_to_del == st.session_state['user']['name']:
            st.error("Vous ne pouvez pas supprimer le compte avec lequel vous êtes actuellement connecté !")
        else:
            try:
                # Appel à l'API avec la méthode DELETE
                # Note : On passe le nom dans l'URL directement comme défini dans FastAPI
                res = requests.delete(f"{API_URL}/delete-user/{user_to_del}", headers=headers)
                        
                if res.status_code == 200:
                    st.success(f"Le compte '{user_to_del}' a été supprimé.")
                    # Petit rerun pour rafraîchir l'interface si nécessaire
                    st.rerun()
                else:
                    st.error(f"Erreur : {res.json().get('detail', 'Impossible de supprimer cet utilisateur')}")
            except Exception as e:
                st.error(f"Erreur de communication avec l'API : {e}")
            else:
                st.warning("Veuillez saisir un nom d'utilisateur.")
