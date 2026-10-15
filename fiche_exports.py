"""Shared fiche validation and original-file generation for all export flows."""

from contract_periods import contract_filename
from DocxGen import generer_docx_stagiaire
from ExcelGen import remplir_fiche_paie


def fiche_name(employee: dict) -> str:
    if employee.get("type") == "Stagiaire":
        return " ".join(str(employee.get(key) or "").strip() for key in ("prenom_stagiaire", "nom_stagiaire")).strip()
    return str(employee.get("nom") or "").strip()


def validate_fiche(employee: dict, *, require_email=False, require_first_name=False, check_exceptions=False) -> None:
    """Keep each flow's validation options explicit while sharing their rules."""
    intern = employee.get("type") == "Stagiaire"
    fields = [("nom_stagiaire", "nom"), ("dds", "début de stage"), ("fds", "fin de stage")] if intern else [
        ("nom", "nom"), ("ddc", "début de contrat"), ("fdc", "fin de contrat"), ("responsable", "responsable")
    ]
    if intern and require_first_name:
        fields.append(("prenom_stagiaire", "prénom"))
    if require_email:
        fields.append(("email_employe", "e-mail"))
    missing = [label for key, label in fields if not str(employee.get(key) or "").strip()]
    if missing:
        raise ValueError(f"Information manquante : {', '.join(missing)}.")
    if check_exceptions and not intern:
        for key, label in (("vacances", "congé payé"), ("absences", "absence"), ("arret", "arrêt maladie")):
            if any(not day.get("matin") and not day.get("aprem") for day in employee.get(key, [])):
                raise ValueError(f"Cochez Matin ou Après-midi pour chaque {label}.")


def generate_fiche_file(employee: dict, month: int, year: int, *, ordinal: int | None = None) -> dict:
    """Generate an Excel/Word original with a safe, contract-aware filename."""
    intern = employee.get("type") == "Stagiaire"
    name = fiche_name(employee) or "Personne"
    safe_name = "".join(char if char.isalnum() or char in "_-" else "_" for char in name)
    safe_name = contract_filename(safe_name, employee)
    prefix, extension = ("fiche_presence_stage", "docx") if intern else ("fiche_presence", "xlsx")
    suffix = f"_{ordinal}" if ordinal is not None else ""
    generated = generer_docx_stagiaire(employee, month, year) if intern else remplir_fiche_paie(month, year, employee)
    return {
        "data": generated.getvalue(),
        "filename": f"{prefix}_{safe_name}_{month}_{year}{suffix}.{extension}",
        "mime": "application/vnd.openxmlformats-officedocument." + ("wordprocessingml.document" if intern else "spreadsheetml.sheet"),
        "download_label": "Télécharger le Word" if intern else "Télécharger l'Excel",
        "employee_name": name,
    }
