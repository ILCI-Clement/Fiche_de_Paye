"""Domain helpers for the attendance application."""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime
from typing import Any
from uuid import uuid4
import math
from attendance import duration_minutes


PERSON_TYPES = {"employee", "intern"}
FORM_STATES = {"draft", "complete", "exported", "archived", "trash"}
EXCEPTION_TYPES = {
    "work",
    "rest",
    "paid_leave",
    "absence",
    "sick_leave",
    "public_holiday",
    "other",
}
WEEK_DAYS = (
    "Lundi",
    "Mardi",
    "Mercredi",
    "Jeudi",
    "Vendredi",
    "Samedi",
    "Dimanche",
)


def current_timestamp() -> str:
    """Return a timezone-aware ISO timestamp."""
    return datetime.now().astimezone().isoformat(timespec="seconds")


def default_schedule() -> dict[str, dict[str, Any]]:
    """Build the standard weekly schedule used by new people."""
    return {
        day: {
            "active": day not in {"Samedi", "Dimanche"},
            "morning_start": "09:00",
            "morning_end": "12:00",
            "afternoon_start": "13:00",
            "afternoon_end": "17:00",
        }
        for day in WEEK_DAYS
    }


def person_display_name(person: dict[str, Any]) -> str:
    """Return a stable, readable display name for a person record."""
    first_name = person.get("first_name", "").strip()
    last_name = person.get("last_name", "").strip()
    return " ".join(part for part in (last_name.upper(), first_name) if part) or "Sans nom"


def create_person(person_type: str, **values: Any) -> dict[str, Any]:
    """Create a person record with reusable fields and a private identifier."""
    if person_type not in PERSON_TYPES:
        raise ValueError("Unsupported person type")

    record = {
        "id": str(uuid4()),
        "type": person_type,
        "first_name": "",
        "last_name": "",
        "supervisor_name": "",
        "supervisor_email": "",
        "start_date": None,
        "end_date": None,
        "permanent_contract": False,
        "hourly_rate": 0.0,
        "day_count": 0.0,
        "daily_hours": 7.0,
        "transport": "",
        "transport_cost": 0.0,
        "transport_rate": 0.0,
        "default_schedule": default_schedule(),
        "created_at": current_timestamp(),
        "updated_at": current_timestamp(),
        "archived": False,
    }
    record.update(values)
    return record


def create_form(
    year: int,
    month: int,
    person: dict[str, Any] | None = None,
    *,
    person_type: str | None = None,
    first_name: str = "",
    last_name: str = "",
) -> dict[str, Any]:
    """Create an independent monthly form from a person snapshot or blank data."""
    if not 1 <= month <= 12:
        raise ValueError("Month must be between 1 and 12")

    if person is not None:
        person_type = person["type"]
        first_name = person.get("first_name", "")
        last_name = person.get("last_name", "")
        schedule = deepcopy(person.get("default_schedule", default_schedule()))
        person_id = person["id"]
        snapshot = deepcopy(person)
    else:
        if person_type not in PERSON_TYPES:
            raise ValueError("A blank form requires a person type")
        schedule = default_schedule()
        person_id = None
        snapshot = {}

    return {
        "id": str(uuid4()),
        "person_id": person_id,
        "person_type": person_type,
        "first_name": first_name,
        "last_name": last_name,
        "person_snapshot": snapshot,
        "year": year,
        "month": month,
        "schedule": schedule,
        "exceptions": {},
        "state": "draft",
        "archived": False,
        "exported_formats": [],
        "created_at": current_timestamp(),
        "updated_at": current_timestamp(),
    }


def form_identity(form: dict[str, Any]) -> tuple[str, str, int, int]:
    """Return the fields used to prevent a duplicate monthly form."""
    person_key = form.get("person_id") or "|".join(
        (
            form.get("last_name", "").strip().casefold(),
            form.get("first_name", "").strip().casefold(),
        )
    )
    return (person_key, form["person_type"], form["year"], form["month"])


def find_duplicate_form(forms: list[dict[str, Any]], candidate: dict[str, Any]) -> dict[str, Any] | None:
    """Find a non-deleted form with the same person and target month."""
    identity = form_identity(candidate)
    return next(
        (
            form
            for form in forms
            if form.get("id") != candidate.get("id")
            and form.get("state") != "trash"
            and (form_identity(form) == identity or (
                form.get('person_id') and form.get('person_id') == candidate.get('person_id')
                and (form['year'], form['month']) == (candidate['year'], candidate['month'])
            ) or (
                (form['year'], form['month'], form['person_type']) ==
                (candidate['year'], candidate['month'], candidate['person_type'])
                and ' '.join((form.get('last_name', '') + ' ' + form.get('first_name', '')).casefold().split()) ==
                ' '.join((candidate.get('last_name', '') + ' ' + candidate.get('first_name', '')).casefold().split())
            ))
        ),
        None,
    )


def validate_form(form: dict[str, Any]) -> list[str]:
    """Return French validation messages for a monthly form."""
    errors: list[str] = []
    if not form.get("last_name", "").strip():
        errors.append("Le nom de la personne est obligatoire.")
    if form.get("person_type") not in PERSON_TYPES:
        errors.append("Le type de personne est invalide.")
    if not isinstance(form.get("month"), int) or not 1 <= form["month"] <= 12:
        errors.append("Le mois est invalide.")
    if not isinstance(form.get("year"), int) or not 2000 <= form["year"] <= 2100:
        errors.append("L’année est invalide.")
    details = form.get('person_snapshot', {})
    if not details.get('supervisor_name', '').strip():
        errors.append('Le responsable est obligatoire.')
    parsed_dates = {}
    for field in ('start_date', 'end_date'):
        if field == 'end_date' and details.get('permanent_contract') and form.get('person_type') == 'employee':
            continue
        try:
            parsed_dates[field] = date.fromisoformat(str(details.get(field)))
        except (TypeError, ValueError):
            errors.append('La date de début est obligatoire.' if field == 'start_date' else 'La date de fin est obligatoire.')
    if len(parsed_dates) == 2 and parsed_dates['end_date'] < parsed_dates['start_date']:
        errors.append('La date de fin doit suivre la date de début.')
    for day, schedule in form.get('schedule', {}).items():
        if not schedule.get('active'):
            continue
        try:
            for half in ('morning', 'afternoon'):
                duration_minutes(schedule.get(f'{half}_start'), schedule.get(f'{half}_end'))
            if schedule.get('morning_end') and schedule.get('afternoon_start') and schedule['morning_end'] > schedule['afternoon_start']:
                raise ValueError('Les créneaux du matin et de l’après-midi se chevauchent.')
        except ValueError as error:
            errors.append(f'{day} : {error}')
    if form.get('person_type') == 'intern':
        if not form.get('first_name', '').strip():
            errors.append('Le prénom du stagiaire est obligatoire.')
        for field, label, maximum in [('hourly_rate', 'taux horaire', None), ('day_count', 'nombre de jours', 31), ('daily_hours', 'heures par jour', 24), ('transport_cost', 'facture transport', None), ('transport_rate', 'taux de remboursement', 100)]:
            value = details.get(field, 0)
            if not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0 or (maximum is not None and value > maximum):
                errors.append(f'Valeur invalide : {label}.')
    for day_key, half_days in form.get("exceptions", {}).items():
        try:
            parsed_day = date.fromisoformat(day_key)
            if (parsed_day.year, parsed_day.month) != (form.get('year'), form.get('month')):
                errors.append(f'{day_key} : date hors du mois sélectionné.')
        except (TypeError, ValueError):
            errors.append("Une date du calendrier est invalide.")
            continue
        for half_day, item in half_days.items():
            if half_day not in {"morning", "afternoon"}:
                errors.append("Une période du calendrier est invalide.")
            elif item.get("type") not in EXCEPTION_TYPES:
                errors.append("Un statut du calendrier est invalide.")
            elif item.get("type") == "other" and item.get("hours") is None:
                errors.append("Le statut Autre exige un nombre d’heures.")
            elif item.get('type') == 'other':
                value = item['hours']
                if not isinstance(value, (int, float)) or not math.isfinite(value) or not 0 <= value <= 24 or not item.get('label', '').strip():
                    errors.append(f'{day_key} : Autre exige une description et des heures entre 0 et 24.')
    return errors


def ensure_workspace(data: dict[str, Any] | None) -> dict[str, Any]:
    """Ensure that a remote configuration contains the new workspace keys."""
    if data is not None and not isinstance(data, dict):
        raise ValueError('Format de configuration serveur invalide.')
    workspace = data if data is not None else {}
    workspace.setdefault("schema_version", 2)
    workspace.setdefault("people", [])
    workspace.setdefault("forms", [])
    workspace.setdefault("trash", [])
    if any(not isinstance(workspace[key], list) for key in ('people', 'forms', 'trash')):
        raise ValueError('Structure de configuration serveur invalide.')
    return workspace


def import_legacy_forms(workspace):
    """Import old records explicitly, preserving the original source payload."""
    if workspace.get('legacy_imported'):
        return 0
    year, month = int(workspace.get('annee', date.today().year)), int(workspace.get('mois', date.today().month))
    imported = []
    for index, source in enumerate(workspace.get('employes_data', [])):
        person_type = 'intern' if source.get('type') == 'Stagiaire' else 'employee'
        form = create_form(year, month, person_type=person_type, first_name=source.get('prenom_stagiaire', '') if person_type == 'intern' else '', last_name=source.get('nom_stagiaire', '') if person_type == 'intern' else source.get('nom', ''))
        field_map = {'supervisor_name': 'responsable', 'supervisor_email': 'email_responsable', 'hourly_rate': 'taux_horaire', 'day_count': 'nb_jours', 'daily_hours': 'nb_heures_jour', 'transport': 'transport', 'transport_cost': 'facture_mensuelle', 'transport_rate': 'taux', 'permanent_contract': 'cdi'}
        form['person_snapshot'] = {target: deepcopy(source[key]) for target, key in field_map.items() if key in source}
        form['person_snapshot']['start_date'] = source.get('dds' if person_type == 'intern' else 'ddc')
        form['person_snapshot']['end_date'] = source.get('fds' if person_type == 'intern' else 'fdc')
        for day, shift in source.get('planning_detail', {}).items():
            form['schedule'][day] = {'active': shift.get('actif', False), 'morning_start': shift.get('m1', ''), 'morning_end': shift.get('m2', ''), 'afternoon_start': shift.get('a1', ''), 'afternoon_end': shift.get('a2', '')}
        for key, status in [('arret', 'sick_leave'), ('absences', 'absence'), ('vacances', 'paid_leave')]:
            for event in source.get(key, []):
                raw_date = str(event.get('date'))
                for half, legacy_half in [('morning', 'matin'), ('afternoon', 'aprem')]:
                    if event.get(legacy_half):
                        form['exceptions'].setdefault(raw_date, {})[half] = {'type': status, 'exam_leave': bool(event.get('examen_alt'))}
        form['legacy_index'] = index
        if not find_duplicate_form(workspace['forms'] + imported, form):
            imported.append(form)
    workspace['forms'].extend(imported)
    workspace['legacy_imported'] = True
    return len(imported)


def move_to_trash(workspace: dict[str, Any], form_id: str) -> bool:
    """Move a form to the recycle bin without deleting its data."""
    forms = workspace["forms"]
    form = next((item for item in forms if item["id"] == form_id), None)
    if form is None:
        return False
    forms.remove(form)
    form['previous_state'] = form.get('state', 'draft')
    form["state"] = "trash"
    form["deleted_at"] = current_timestamp()
    workspace["trash"].append(form)
    return True


def restore_from_trash(workspace: dict[str, Any], form_id: str) -> bool:
    """Restore a recycled form to the export queue."""
    trash = workspace["trash"]
    form = next((item for item in trash if item["id"] == form_id), None)
    if form is None:
        return False
    if find_duplicate_form(workspace['forms'], form):
        raise ValueError('Une fiche existe déjà pour cette personne et ce mois. La restauration est bloquée.')
    trash.remove(form)
    form['state'] = form.pop('previous_state', 'archived' if form.get('archived') else 'draft')
    form['archived'] = form['state'] == 'archived'
    form.pop("deleted_at", None)
    workspace["forms"].append(form)
    return True
