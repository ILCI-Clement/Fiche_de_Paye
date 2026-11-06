"""Read-only approved leave totals from historical attendance synchronization cells."""
import json
from datetime import date

from .leave_attendance import PAID, leave_cells


def annual_leave_totals(cursor, username, year, requests):
    approved = {row['id']: row for row in requests if row['status'] == 'approved'}
    monthly = [set() for _ in range(12)]
    unknown = set()
    if approved:
        # Include archived requests: archiving is presentation, not cancellation.
        args = (username, date(year, 12, 31), date(year, 1, 1))
        scope = "r.employee_username=%s AND r.status='approved' AND r.start_date<=%s AND r.end_date>=%s"
        cursor.execute(f'''SELECT e.request_id FROM leave_events e JOIN leave_requests r ON r.id=e.request_id
            WHERE {scope} AND e.action='attendance_synced' ''', args)
        synced = {row['request_id'] for row in cursor.fetchall()}
        cursor.execute(f'''SELECT c.request_id,c.cell_date,c.cell_period,c.after_state
            FROM leave_attendance_cells c JOIN leave_requests r ON r.id=c.request_id
            WHERE {scope} AND c.cell_date>=%s AND c.cell_date<=%s''',
                       (*args, date(year, 1, 1), date(year, 12, 31)))
        cells = cursor.fetchall()
        expected = {}
        for identifier, row in approved.items():
            start, end = (date.fromisoformat(str(row[key])[:10]) for key in ('start_date', 'end_date'))
            clipped = {**row, 'start_date': max(start, date(year, 1, 1)),
                       'end_date': min(end, date(year, 12, 31)),
                       'start_slot': row['start_slot'] if start.year == year else 'full',
                       'end_slot': row['end_slot'] if end.year == year else 'full'}
            expected[identifier] = set(leave_cells(clipped))
        recorded = set()
        for row in cells:
            identifier = row['request_id']
            if identifier not in approved:
                continue
            recorded.add(identifier)
            try:
                state = json.loads(row['after_state']) if isinstance(row['after_state'], str) else row['after_state']
                current = date.fromisoformat(str(row['cell_date'])[:10])
                period = row['cell_period']
                if (current, period) not in expected[identifier] or state['resolved'][0] != PAID:
                    raise ValueError('Invalid leave cell')
                monthly[current.month - 1].add((current, period))
            except (ValueError, TypeError, KeyError, IndexError):
                unknown.add(identifier)
        # Old approvals without synchronization history have no reliable schedule snapshot.
        # Do not guess weekend/holiday/part-time rules from today's schedule.
        unknown |= set(approved) - synced - recorded
    unknown_months = [False] * 12
    for identifier in unknown:
        row = approved[identifier]
        for month in range(1, 13):
            if str(row['start_date'])[:7] <= f'{year}-{month:02d}' <= str(row['end_date'])[:7]:
                unknown_months[month - 1] = True
    return {'approved_days': sum(len(cells) for cells in monthly) / 2,
            'approved_half_days': sum(len(cells) for cells in monthly),
            'incomplete': bool(unknown), 'unverified_requests': len(unknown),
            'months': [{'approved_days': len(cells) / 2, 'incomplete': unknown_months[i]}
                       for i, cells in enumerate(monthly)]}
