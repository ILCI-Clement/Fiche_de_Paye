import streamlit as st
import calendar
from datetime import date, datetime, time
import holidays
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
import io

def heures_vers_texte(nb_heures):
    heures = int(nb_heures)
    minutes = int(round((nb_heures - heures) * 60))
    return f"{heures:02d}:{minutes:02d}"

def calculer_heures(ws, ligne_debut, ligne_fin, col_heure_debut, col_heure_fin, col_resultat):
    fmt = "%H:%M"
    for row in range(ligne_debut, ligne_fin + 1):
        debut = ws.cell(row=row, column=col_heure_debut).value
        fin = ws.cell(row=row, column=col_heure_fin).value

        if isinstance(debut, str) and isinstance(fin, str):
            try:
                h_debut = datetime.strptime(debut, fmt)
                h_fin = datetime.strptime(fin, fmt)
                diff = (h_fin - h_debut).seconds / 3600
                ws.cell(row=row, column=col_resultat, value=heures_vers_texte(diff))
            except ValueError:
                ws.cell(row=row, column=col_resultat, value="")
        else:
            ws.cell(row=row, column=col_resultat, value="")

def somme(ws, lignes, col_source, col_resultat, ligne_total, somme):
    total_minutes = 0

    if somme == "semaine":
        for row in lignes:
            val = ws.cell(row=row, column=col_source).value
            if isinstance(val, str):
                try:
                    heures, minutes = map(int, val.split(":"))
                    total_minutes += heures * 60 + minutes
                except:
                    pass

    elif somme == "total":
        for col in col_source:
            val = ws.cell(row=lignes, column=col).value
            if isinstance(val, str):
                try:
                    heures, minutes = map(int, val.split(":"))
                    total_minutes += heures * 60 + minutes
                except:
                    pass

    total_heures = total_minutes // 60
    reste_minutes = total_minutes % 60
    ws.cell(row=ligne_total, column=col_resultat, value=f"{total_heures:02d}:{reste_minutes:02d}")

def remplir_calendrier(ws, mois, annee, vacances, nom, responsable):
    mois_string = ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"]

    ws.cell(row=4, column=28, value=nom)
    ws.cell(row=8, column=28, value=responsable)
    ws.cell(row=4, column=14, value=mois_string[mois-1])
    ws.cell(row=7, column=14, value=annee)
    
    clnn = 3

    for i in range(5):
        lgn = 15
        for j in range(7):
            ws.cell(row=lgn, column=clnn, value="")
            for k in range(4):
                for l in range(2):
                    ws.cell(row=lgn+l, column=clnn+k+1, value="")
            lgn += 2
        clnn += 5

    colonnes = [15, 17, 19, 21, 23, 25, 27]
    jours_feries = holidays.France(years=annee)

    nb_jours = calendar.monthrange(annee, mois)[1]

    ligne = 3
    jour = 1
    premier_jour = date(annee, mois, 1)
    decalage = premier_jour.weekday()

    while jour <= nb_jours:
        for i, col in enumerate(colonnes):
            if i < decalage and jour == 1:
                continue
            if jour > nb_jours:
                break

            cell = ws.cell(row=col, column=ligne, value=jour)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)
            d = date(annee, mois, jour)

            if d in jours_feries:
                ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                cell = ws.cell(row=col, column=ligne+1, value="FERIE")
                cell.alignment = Alignment(horizontal="center", vertical="center")

            elif d in vacances:
                if vacances[d] == (True, True):
                    ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                    cell = ws.cell(row=col, column=ligne+1, value=f"CP\n09:00 à 17:00")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif vacances[d] == (True, False):
                    cell = ws.cell(row=col+1, column=ligne+4, value="04:00")
                    ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                    cell = ws.cell(row=col, column=ligne+1, value=f"CP\n09:00 à 12:00")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif vacances[d] == (False, True):
                    cell = ws.cell(row=col, column=ligne+4, value="03:00")
                    ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                    cell = ws.cell(row=col, column=ligne+1, value=f"CP\n13:00 à 17:00")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


            elif d.weekday() in [5, 6]:
                ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                cell = ws.cell(row=col, column=ligne+1, value="WEEK-END")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            else:
                cell = ws.cell(row=col, column=ligne+1, value="09:00")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell = ws.cell(row=col, column=ligne+2, value="à")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell = ws.cell(row=col, column=ligne+3, value="12:00")
                cell.alignment = Alignment(horizontal="center", vertical="center")

                cell = ws.cell(row=col+1, column=ligne+1, value="13:00")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell = ws.cell(row=col+1, column=ligne+2, value="à")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell = ws.cell(row=col+1, column=ligne+3, value="17:00")
                cell.alignment = Alignment(horizontal="center", vertical="center")

                calculer_heures(ws, col, col, ligne+1, ligne+3, ligne+4)
                calculer_heures(ws, col+1, col+1, ligne+1, ligne+3, ligne+4)

            jour += 1

        decalage = 0
        ligne += 5

    somme(ws, lignes=[15,16,17,18,19,20,21,22,23,24], col_source=7, col_resultat=3, ligne_total=29, somme="semaine")
    somme(ws, lignes=[15,16,17,18,19,20,21,22,23,24], col_source=12, col_resultat=8, ligne_total=29, somme="semaine")
    somme(ws, lignes=[15,16,17,18,19,20,21,22,23,24], col_source=17, col_resultat=13, ligne_total=29, somme="semaine")
    somme(ws, lignes=[15,16,17,18,19,20,21,22,23,24], col_source=22, col_resultat=18, ligne_total=29, somme="semaine")
    somme(ws, lignes=[15,16,17,18,19,20,21,22,23,24], col_source=27, col_resultat=23, ligne_total=29, somme="semaine")

    somme(ws, lignes=29, col_source=[3,8,13,18,23], col_resultat=28, ligne_total=29, somme="total")

def remplir_fiche_paie(fichier_entree, mois, annee, employes_data):
    wb = load_workbook(fichier_entree)
    modele = wb.active

    for employe in employes_data:
        ws = wb.copy_worksheet(modele)
        if employe["nom"]:
            ws.title = employe["nom"]
        else:
            ws.title = "Sans nom"
    
        vacances = employe["vacances"]

        remplir_calendrier(ws, mois, annee, vacances, employe["nom"], employe["responsable"])

    wb.remove(modele)
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# Interface Streamlit

st.title("Générateur automatique de fiche de présence (Excel)")

uploaded_excel = "Fiche_Exemple.xlsx"

col1, col2 = st.columns(2)
with col1:
    mois = st.number_input("Mois", min_value=1, max_value=12, value=1)
with col2:
    annee = st.number_input("Année", min_value=2000, max_value=2100, value=2025)



nb_employe = st.number_input("Nombre d'employés :", min_value=1, max_value=10, step=1)

employe = [f"Employé {j+1}" for j in range(nb_employe)]

employes_data = []
tabs = st.tabs(employe)

for h, tab in enumerate(tabs):
    with tab:
        st.subheader("Information Employé")
        nom = st.text_input("NOM Prénom de l'employé", key=f"employe_nom_{h}")
        responsable = st.text_input("NOM Prénom du Responsable", key=f"resp_nom_{h}")
        st.subheader("Saisir les jours de vacances")

        nb_jours_vac = st.number_input("Nombre de jours de vacances :", min_value=0, max_value=31, value=0, key=f"nb_jours_vac_{h}")

        vacances = {}

        for i in range(nb_jours_vac):
            st.markdown(f"### Jour de vacances #{i+1}")
            vac_col1, vac_col2, vac_col3 = st.columns(3)
            with vac_col1:
                d = st.date_input(f"Date du jour {i+1}", key=f"date_{h}_{i}")
            with vac_col2:
                t1 = st.checkbox(f"Matin du jour {i+1}", value=False, key=f"matin_{h}_{i}")
            with vac_col3:
                t2 = st.checkbox(f"Après-Midi du jour {i+1}", value=False, key=f"aprem_{h}_{i}")

            vacances[d] = (t1, t2)

        employes_data.append({"nom": nom, "responsable": responsable, "vacances": vacances})



if st.button("Générer la fiche"): 
    check_erreur = False
    for a in employes_data:
        b = a["vacances"]
        for c in b:
            if b[c] == (False, False):
                check_erreur = True
    if(check_erreur == True):
        st.error(f"Il faut cocher au moins une des deux cases pour le jour de congé de {a["nom"]} !")
    else:
        buffer = remplir_fiche_paie(uploaded_excel, mois, annee, employes_data)

        st.success("Fiche générée avec succès !")

        st.download_button(
            "Télécharger la fiche remplie",
            data=buffer,
            file_name=f"fiche_paie_{mois}_{annee}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )