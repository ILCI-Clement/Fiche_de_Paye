"""Shared defaults for new attendance records; never overwrite saved schedules."""

WEEKDAYS = ("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche")
DEFAULT_SCHEDULE = {"m1": "09:00", "m2": "12:00", "a1": "13:00", "a2": "17:00", "actif": True}
DEFAULT_WORKING_DAYS = frozenset(WEEKDAYS[:5])


def default_planning() -> dict:
    return {
        day: {**DEFAULT_SCHEDULE, "actif": day in DEFAULT_WORKING_DAYS}
        for day in WEEKDAYS
    }
