from docxtpl import DocxTemplate
import io
from num2words import num2words

def generer_docx_stagiaire(data_stagiaire, mois, annee):
    # 1. Chargement du modèle
    doc = DocxTemplate("template_stagiaire.docx")
    
    # 2. Calculs automatiques
    th = data_stagiaire.get("taux_horaire", 0)
    nj = data_stagiaire.get("nb_jours", 0)
    nhj = data_stagiaire.get("nb_heures_jour", 0)
    
    total_stage = round(th * nj * nhj, 2)
    
    facture = data_stagiaire.get("facture_mensuelle", 0)
    taux_tr = data_stagiaire.get("taux", 0) / 100
    total_transport = round(facture * taux_tr, 2)
    
    total_general = round(total_stage + total_transport, 2)

    mois_string = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    
    # Conversion du total en lettres (Français)
    total_lettres = num2words(total_general, lang='fr', to='currency')

    # 3. Dictionnaire de correspondance avec les balises du Word
    context = {
        "nom": data_stagiaire.get("nom_stagiaire").upper(),
        "prenom": data_stagiaire.get("prenom_stagiaire").capitalize(),
        "mois": mois_string[mois-1],
        "annee": annee,
        "dds":data_stagiaire.get("dds"),
        "fds":data_stagiaire.get("fds"),
        "taux_horaire": f"{th:.2f}",
        "nb_jours": nj,
        "nb_heures_jour": nhj,
        "total_stage": f"{total_stage:.2f}",
        "transport": data_stagiaire.get("transport", ""),
        "facture_mensuelle": f"{facture:.2f}",
        "taux": f"{data_stagiaire.get('taux')}%",
        "total_transport": f"{total_transport:.2f}",
        "total": f"{total_general:.2f}",
        "total_lettres": total_lettres.capitalize()
    }

    # 4. Remplissage du document
    doc.render(context)
    
    # 5. Sauvegarde dans un buffer mémoire pour Streamlit
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer