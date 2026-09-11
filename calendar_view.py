"""Interactive monthly calendar for presence forms.

The calendar uses a small local Streamlit component for fixed-size, drag-selectable
half-day cells. It edits the current JSON-compatible employee structure in place.
"""

from __future__ import annotations

import calendar
from datetime import date
from html import escape
from pathlib import Path
import re
from typing import Any

try:
    import holidays
except ImportError:  # pragma: no cover - the production requirements include it
    holidays = None

WEEKDAYS = (
    "Lundi",
    "Mardi",
    "Mercredi",
    "Jeudi",
    "Vendredi",
    "Samedi",
    "Dimanche",
)

STATUS_CLASS = {
    "Travail": "work",
    "Congé payé": "paid-leave",
    "Absence": "absence",
    "Arrêt maladie": "sick-leave",
    "Férié": "holiday",
    "Autre": "other",
    "Repos": "rest",
}

DEFAULT_SCHEDULE = {
    "m1": "09:00",
    "m2": "12:00",
    "a1": "13:00",
    "a2": "17:00",
    "actif": True,
}
DEFAULT_WORKING_DAYS = frozenset(WEEKDAYS[:5])

_COMPONENT_DIRECTORY = Path(__file__).with_name("calendar_selector_component")


def _default_planning() -> dict[str, dict[str, Any]]:
    return {
        day: {**DEFAULT_SCHEDULE, "actif": day in DEFAULT_WORKING_DAYS}
        for day in WEEKDAYS
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


def _short_cell_text(status: str, detail: str) -> str:
    labels = {
        "Travail": "Travail",
        "Congé payé": "CP",
        "Absence": "ABS",
        "Arrêt maladie": "AM",
        "Férié": "Férié",
        "Autre": "Autre",
        "Repos": "Repos",
    }
    label = labels.get(status, status)
    return f"{label} · {detail}" if detail else label


def _short_button_text(period: str, status: str) -> str:
    labels = {
        "Travail": "Travail",
        "Congé payé": "CP",
        "Absence": "Abs.",
        "Arrêt maladie": "Maladie",
        "Férié": "Férié",
        "Autre": "Autre",
        "Repos": "Repos",
    }
    return f"{'AM' if period == 'matin' else 'PM'} · {labels.get(status, status)}"


def _safe_key(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]", "-", value)


def _render_styles() -> None:
    import streamlit as st

    st.markdown(
        """
        <style>
        .presence-calendar-heading {
            color: #e7edf5; font-size: 0.87rem; font-weight: 700;
            text-align: center; padding: 0.25rem 0.1rem 0.45rem;
        }
        .presence-calendar-date {
            color: #d9e2ef; font-weight: 700; font-size: 0.82rem;
            padding: 0.1rem 0.35rem 0.4rem;
        }
        .presence-calendar-empty {
            color: #65748a; font-size: 0.78rem; min-height: 5.2rem;
        }
        .presence-preview { width: 100%; overflow-x: auto; font-size: 0.78rem; }
        .presence-preview-header, .presence-preview-week {
            display: grid; grid-template-columns: repeat(7, minmax(6.35rem, 1fr)); min-width: 44.5rem;
        }
        .presence-preview-header div {
            color: #9aa9bd; font-weight: 700; text-align: center; padding: 0.35rem 0.2rem;
        }
        .presence-preview-week { border-left: 1px solid #293341; }
        .presence-preview-day {
            min-height: 5.9rem; background: #0d1219; border-top: 1px solid #293341;
            border-right: 1px solid #293341; padding: 0.35rem 0;
        }
        .presence-preview-day.empty { background: #0a0f15; }
        .presence-preview-date { color: #d9e2ef; font-weight: 700; padding: 0 0.35rem 0.32rem; }
        .presence-preview-segment {
            height: 3.25rem; margin: 0.12rem 0; padding: 0.28rem 0.34rem;
            overflow: hidden; color: #eef4fb;
            border: 1px solid transparent;
        }
        .presence-preview-segment span { color: #c6d2e0; font-size: 0.66rem; margin-right: 0.2rem; }
        .presence-preview-segment.work { background: #1e4269; border-color: #4d85bd; }
        .presence-preview-segment.paid-leave { background: #1f513c; border-color: #4c9b72; }
        .presence-preview-segment.absence { background: #6b3b24; border-color: #b87549; }
        .presence-preview-segment.sick-leave { background: #6b3038; border-color: #ba6974; }
        .presence-preview-segment.holiday { background: #4a3e76; border-color: #887ac2; }
        .presence-preview-segment.other { background: #66521b; border-color: #b89a3c; }
        .presence-preview-segment.rest { background: #202833; border-color: #465364; color: #b7c2d0; }
        .presence-preview-segment.round-left { border-radius: 0.75rem 0 0 0.75rem; }
        .presence-preview-segment.round-right { border-radius: 0 0.75rem 0.75rem 0; }
        .presence-preview-segment.round-left.round-right { border-radius: 0.75rem; }
        div[data-baseweb="popover"] {
            background: #151b24; border: 1px solid #394455;
            color: #eef3f8;
        }
        div[data-baseweb="popover"] [data-testid="stMarkdownContainer"] p {
            color: #d9e2ef;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


STATUS_PALETTE = {
    "work": ("#1e4269", "#4d85bd"),
    "paid-leave": ("#1f513c", "#4c9b72"),
    "absence": ("#6b3b24", "#b87549"),
    "sick-leave": ("#6b3038", "#ba6974"),
    "holiday": ("#4a3e76", "#887ac2"),
    "other": ("#66521b", "#b89a3c"),
    "rest": ("#202833", "#465364"),
}


def _preview_segment(
    status: str,
    detail: str,
    css_class: str,
    left_edge: bool,
    right_edge: bool,
    period: str,
) -> str:
    classes = ["presence-preview-segment", css_class]
    if left_edge:
        classes.append("round-left")
    if right_edge:
        classes.append("round-right")
    return (
        f"<div class='{' '.join(classes)}'><span>{'AM' if period == 'matin' else 'PM'}</span>"
        f"{escape(_short_cell_text(status, detail))}</div>"
    )


def _render_calendar_preview(employee: dict[str, Any], month: int, year: int) -> None:
    """Render the read-only dark calendar with connected status bars."""

    import streamlit as st

    html = ["<div class='presence-preview'>", "<div class='presence-preview-header'>"]
    html.extend(f"<div>{escape(day[:3])}</div>" for day in WEEKDAYS)
    html.append("</div>")
    for week in calendar.monthcalendar(year, month):
        html.append("<div class='presence-preview-week'>")
        states: dict[tuple[int, str], tuple[str, str, str]] = {}
        for day_number in week:
            if day_number:
                current_day = date(year, month, day_number)
                for period in ("matin", "aprem"):
                    states[(day_number, period)] = resolve_cell(employee, current_day, period)
        for index, day_number in enumerate(week):
            if not day_number:
                html.append("<div class='presence-preview-day empty'></div>")
                continue
            segments: list[str] = []
            for period in ("matin", "aprem"):
                status, detail, css_class = states[(day_number, period)]
                previous_day = week[index - 1] if index else 0
                next_day = week[index + 1] if index < 6 else 0
                previous_same = bool(
                    previous_day
                    and states[(previous_day, period)][0] == status
                    and states[(previous_day, period)][2] == css_class
                )
                next_same = bool(
                    next_day
                    and states[(next_day, period)][0] == status
                    and states[(next_day, period)][2] == css_class
                )
                segments.append(
                    _preview_segment(
                        status,
                        detail,
                        css_class,
                        not previous_same,
                        not next_same,
                        period,
                    )
                )
            html.append(
                f"<div class='presence-preview-day'><div class='presence-preview-date'>{day_number}</div>"
                + "".join(segments)
                + "</div>"
            )
        html.append("</div>")
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def _selection_key(key_prefix: str, day: date, period: str) -> str:
    return f"{key_prefix}_calendar_selection_{day.isoformat()}_{period}"


def _render_edit_grid(
    employee: dict[str, Any],
    month: int,
    year: int,
    key_prefix: str,
) -> None:
    """Render a fixed-size, drag-selectable half-day calendar component."""

    import streamlit as st
    import streamlit.components.v1 as components

    calendar_component = components.declare_component(
        "presence_calendar_selector",
        path=str(_COMPONENT_DIRECTORY),
    )
    cells: list[dict[str, str | int]] = []
    for week in calendar.monthcalendar(year, month):
        for day_number in week:
            if not day_number:
                continue
            current_day = date(year, month, day_number)
            for period in ("matin", "aprem"):
                status, detail, css_class = resolve_cell(employee, current_day, period)
                cells.append(
                    {
                        "key": _selection_key(key_prefix, current_day, period),
                        "date": current_day.isoformat(),
                        "day": day_number,
                        "weekday": current_day.weekday(),
                        "period": "AM" if period == "matin" else "PM",
                        "label": _short_cell_text(status, detail),
                        "style": css_class,
                    }
                )

    value = calendar_component(
        cells=cells,
        weekdays=[day[:3] for day in WEEKDAYS],
        key=f"{key_prefix}_calendar_drag_selector_{year}_{month}",
        default=None,
    )
    if not isinstance(value, dict) or value.get("action") != "apply":
        return

    status = str(value.get("status", ""))
    if status not in {"Travail", "Congé payé", "Absence", "Arrêt maladie", "Férié", "Autre", "Réinitialiser"}:
        return
    reason = str(value.get("reason", ""))
    if status == "Autre" and not reason.strip():
        return

    selected_keys = {str(item) for item in value.get("keys", [])}
    selected_cells = {
        str(cell["key"]): (date.fromisoformat(str(cell["date"])), "matin" if cell["period"] == "AM" else "aprem")
        for cell in cells
    }
    for selected_key in selected_keys:
        selected_day = selected_cells.get(selected_key)
        if selected_day:
            apply_calendar_status(employee, selected_day[0], selected_day[1], status, reason)
    if selected_keys:
        st.rerun()


def render_monthly_calendar(
    employee: dict[str, Any],
    month: int,
    year: int,
    key_prefix: str,
) -> None:
    """Render a read-only calendar with an explicit batch-edit mode."""

    import streamlit as st

    ensure_calendar_data(employee)
    _render_styles()
    st.subheader("Calendrier mensuel")
    edit_key = f"{key_prefix}_calendar_edit_mode"
    is_editing = st.session_state.get(edit_key, False)

    if not is_editing:
        st.caption("Consultez les présences du mois ou ouvrez le mode de modification groupée.")
        if st.button("Modifier", key=f"{key_prefix}_calendar_start_edit", type="primary"):
            st.session_state[edit_key] = True
            st.rerun()
        _render_calendar_preview(employee, month, year)
        return

    st.caption(
        "Cliquez ou faites glisser pour sélectionner les matinées et après-midis à modifier, "
        "puis appliquez un statut à la sélection."
    )
    if st.button("Terminer", key=f"{key_prefix}_calendar_finish_edit"):
        st.session_state[edit_key] = False
        st.rerun()
    _render_edit_grid(employee, month, year, key_prefix)
