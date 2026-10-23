"""Shared half-day attendance rules used by the React API and file generation."""

from __future__ import annotations
from contract_periods import outside_contract

from datetime import date
from typing import Any

try:
    import holidays
except ImportError:  # pragma: no cover - the production requirements include it
    holidays = None

from planning import WEEKDAYS, DEFAULT_SCHEDULE, DEFAULT_WORKING_DAYS

STATUS_CLASS = {
    "Travail": "work",
    "Congé payé": "paid-leave",
    "Absence": "absence",
    "Arrêt maladie": "sick-leave",
    "Férié": "holiday",
    "Autre": "other",
    "Repos": "rest",
}



def ensure_calendar_data(employee: dict[str, Any]) -> None:
    """Add only missing calendar-compatible keys to a legacy employee record."""

    employee.setdefault("vacances", [])
    employee.setdefault("absences", [])
    employee.setdefault("arret", [])
    planning = employee.setdefault("planning_detail", {})
    if not isinstance(planning, dict):
        planning = {}
        employee["planning_detail"] = planning
    for day in WEEKDAYS:
        schedule = planning.setdefault(day, {})
        if not isinstance(schedule, dict):
            schedule = {}
            planning[day] = schedule
        schedule_defaults = {**DEFAULT_SCHEDULE, "actif": day in DEFAULT_WORKING_DAYS}
        for key, default in schedule_defaults.items():
            schedule.setdefault(key, default)
    employee.setdefault("calendar_overrides", {})


def _date_key(value: Any) -> str:
    if isinstance(value, date):
        return value.isoformat()
    return str(value or "")[:10]


def _matching_entry(entries: list[dict[str, Any]], day_key: str) -> dict[str, Any] | None:
    for entry in entries:
        if isinstance(entry, dict) and _date_key(entry.get("date")) == day_key:
            return entry
    return None


def _remove_part(entries: list[dict[str, Any]], day_key: str, period: str) -> None:
    """Remove one half-day from a legacy exception list and clean empty rows."""

    kept: list[dict[str, Any]] = []
    for entry in entries:
        if not isinstance(entry, dict) or _date_key(entry.get("date")) != day_key:
            kept.append(entry)
            continue
        entry[period] = False
        if entry.get("matin") or entry.get("aprem"):
            kept.append(entry)
    entries[:] = kept


def _clear_exception(employee: dict[str, Any], day_key: str, period: str) -> None:
    for collection in ("vacances", "absences", "arret"):
        _remove_part(employee.setdefault(collection, []), day_key, period)


def _set_exception(
    employee: dict[str, Any],
    collection: str,
    day_key: str,
    period: str,
) -> None:
    entries = employee.setdefault(collection, [])
    entry = _matching_entry(entries, day_key)
    if entry is None:
        entry = {
            "date": day_key,
            "matin": False,
            "aprem": False,
        }
        if collection == "vacances":
            entry["examen_alt"] = False
        entries.append(entry)
    entry[period] = True


def apply_calendar_status(
    employee: dict[str, Any],
    day: date,
    period: str,
    status: str,
    reason: str = "",
) -> None:
    """Apply a clicked half-day while preserving the existing data model."""

    if outside_contract(employee, day):
        return
    ensure_calendar_data(employee)
    day_key = day.isoformat()
    overrides = employee["calendar_overrides"]

    _clear_exception(employee, day_key, period)

    if status == "Réinitialiser":
        if day_key in overrides:
            overrides[day_key].pop(period, None)
            if not overrides[day_key]:
                overrides.pop(day_key, None)
        return

    exception_map = {
        "Congé payé": "vacances",
        "Absence": "absences",
        "Arrêt maladie": "arret",
    }
    if status in exception_map:
        overrides.setdefault(day_key, {}).pop(period, None)
        _set_exception(employee, exception_map[status], day_key, period)
        if not overrides.get(day_key):
            overrides.pop(day_key, None)
        return

    overrides.setdefault(day_key, {})[period] = {
        "status": status,
        "reason": reason.strip() if status == "Autre" else "",
    }


def _holiday_name(day: date) -> str:
    if holidays is None:
        return ""
    try:
        return str(holidays.France(years=day.year).get(day, ""))
    except Exception:
        return ""


def resolve_cell(
    employee: dict[str, Any],
    day: date,
    period: str,
) -> tuple[str, str, str]:
    """Return ``(status, detail, css_class)`` for one calendar half-day."""

    if outside_contract(employee, day):
        return "Repos", "Hors contrat", "rest"
    ensure_calendar_data(employee)
    day_key = day.isoformat()
    override = employee["calendar_overrides"].get(day_key, {}).get(period)
    if isinstance(override, dict) and override.get("status"):
        status = str(override["status"])
        return status, str(override.get("reason", "")), STATUS_CLASS.get(status, "other")

    for collection, status in (
        ("vacances", "Congé payé"),
        ("absences", "Absence"),
        ("arret", "Arrêt maladie"),
    ):
        entry = _matching_entry(employee.get(collection, []), day_key)
        if entry and entry.get(period):
            detail = "Examen alternance" if entry.get("examen_alt") else ""
            return status, detail, STATUS_CLASS[status]

    holiday_name = _holiday_name(day)
    if holiday_name:
        return "Férié", holiday_name, STATUS_CLASS["Férié"]

    planning = employee.get("planning_detail", {})
    schedule = planning.get(WEEKDAYS[day.weekday()], {})
    if schedule.get("actif", True):
        start_key, end_key = ("m1", "m2") if period == "matin" else ("a1", "a2")
        start = schedule.get(start_key, "")
        end = schedule.get(end_key, "")
        detail = f"{start}–{end}" if start and end else ""
        return "Travail", detail, STATUS_CLASS["Travail"]
    return "Repos", "", STATUS_CLASS["Repos"]
