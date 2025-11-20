import streamlit as st
import calendar
from datetime import date, datetime, time
import holidays
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
import io

def lire_vacances(file):
    vacances = {}
    for ligne in file.getvalue().decode("utf-8").splitlines():
        parts = ligne.strip().split()
        if len(parts) == 3:
            jour = datetime.strptime(parts[0], "%d-%m-%Y").date()
            debut = parts[1]
            fin = parts[2]
            vacances[jour] = (debut, fin)
    return vacances

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

def remplir_fiche_paie(fichier_entree, mois, annee, fichier_vacances):
    wb = load_workbook(fichier_entree)
    ws = wb.active

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
    vacances = lire_vacances(fichier_vacances)

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
                debut, fin = vacances[d]
                ws.merge_cells(start_row=col, start_column=ligne+1, end_row=col+1, end_column=ligne+3)
                cell = ws.cell(row=col, column=ligne+1, value=f"VACANCES\n{debut} - {fin}")
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

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# -------------------------
# Interface Streamlit
# -------------------------

st.title("Générateur automatique de fiche de paie (Excel)")

uploaded_excel = st.file_uploader("Importer le fichier Excel modèle", type=["xlsx"])
uploaded_vacances = st.file_uploader("Importer le fichier vacances.txt", type=["txt"])

col1, col2 = st.columns(2)
with col1:
    mois = st.number_input("Mois", min_value=1, max_value=12, value=1)
with col2:
    annee = st.number_input("Année", min_value=2000, max_value=2100, value=2025)

if st.button("Générer la fiche"):

    if not uploaded_excel or not uploaded_vacances:
        st.error("Merci d'importer l'Excel modèle **et** le fichier vacances.txt.")
    else:
        buffer = remplir_fiche_paie(uploaded_excel, mois, annee, uploaded_vacances)

        st.success("Fiche générée avec succès !")

        st.download_button(
            "Télécharger la fiche remplie",
            data=buffer,
            file_name=f"fiche_paie_{mois}_{annee}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
