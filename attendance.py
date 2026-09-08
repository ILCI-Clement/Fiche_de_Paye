"""Shared attendance calculations; exported totals never depend on widget state."""

import calendar
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import holidays

WEEKDAYS = ('Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche')
STATUS_LABELS = {
    'work': 'Travail', 'rest': 'Repos', 'paid_leave': 'CP',
    'absence': 'ABS', 'sick_leave': 'AM', 'public_holiday': 'Férié / fermeture',
    'other': 'Autre',
}


def duration_minutes(start, end):
    """Reject incomplete or reversed same-day shifts instead of wrapping 24 hours."""
    if not start and not end:
        return 0
    try:
        difference = datetime.strptime(end, '%H:%M') - datetime.strptime(start, '%H:%M')
    except (TypeError, ValueError) as error:
        raise ValueError('Horaire invalide : utiliser HH:MM.') from error
    if difference.total_seconds() <= 0:
        raise ValueError('La fin du créneau doit suivre son début.')
    return int(difference.total_seconds() // 60)


def format_minutes(minutes):
    return f'{minutes // 60:02d}:{minutes % 60:02d}'


def monthly_attendance(form):
    """Calculate each half-day from a snapshot; manual statuses override holidays."""
    year, month = form['year'], form['month']
    public_holidays = holidays.France(years=year, language='fr')
    days = []
    for number in range(1, calendar.monthrange(year, month)[1] + 1):
        day = date(year, month, number)
        schedule = form['schedule'].get(WEEKDAYS[day.weekday()], {})
        row = {'date': day, 'weekday': WEEKDAYS[day.weekday()]}
        for half in ('morning', 'afternoon'):
            start, end = schedule.get(f'{half}_start', ''), schedule.get(f'{half}_end', '')
            default_type = 'work' if schedule.get('active') and start and end else 'rest'
            if day in public_holidays:
                default_type = 'public_holiday'
            override = form.get('exceptions', {}).get(day.isoformat(), {}).get(half, {})
            status = override.get('type', default_type)
            minutes = duration_minutes(start, end) if status == 'work' else 0
            if status == 'other':
                minutes = int(Decimal(str(override.get('hours', 0))) * 60)
            label = STATUS_LABELS[status]
            if status == 'work':
                label = f'{start}–{end}'
            elif status == 'other':
                label = override.get('label', 'Autre')
            elif status == 'paid_leave' and override.get('exam_leave'):
                label = 'CP examen alternance'
            row[half] = {'type': status, 'label': label, 'minutes': minutes, 'start': start, 'end': end}
        row['minutes'] = row['morning']['minutes'] + row['afternoon']['minutes']
        days.append(row)
    return days


def internship_amounts(form):
    """Keep the legacy explicit day-count basis until payment rules are confirmed."""
    details = form.get('person_snapshot', {})
    def money(value):
        return value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    stage = money(Decimal(str(details.get('hourly_rate', 0))) * Decimal(str(details.get('day_count', 0))) * Decimal(str(details.get('daily_hours', 0))))
    transport = money(Decimal(str(details.get('transport_cost', 0))) * Decimal(str(details.get('transport_rate', 0))) / 100)
    return {'stage': stage, 'transport': transport, 'total': stage + transport}
