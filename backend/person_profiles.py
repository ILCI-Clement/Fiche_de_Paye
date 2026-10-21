"""Read-only yearly overview; authorization precedes every employee data query."""
import json
from calendar import monthrange
from datetime import date, datetime
from fastapi import Depends, HTTPException, Query


def register_profile_routes(app, *, get_db_connection, get_current_user, load_user, attempts):
    @app.get('/person-profiles/{username}')
    def profile(username: str, year: int = Query(ge=2000, le=2100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                # Re-read actor and target: neither role nor relationship comes from the client.
                current = load_user(cursor, actor['username'])
                target = load_user(cursor, username)
                if not current or not current.get('is_active', True) or not target or not (
                    current['username'] == target['username']
                    or 'Admin' in current.get('role_tags', [])
                    or current['username'] == target.get('manager_username')
                ):
                    raise HTTPException(403, 'Ce profil est réservé à la personne, son responsable direct et un administrateur.')
                cursor.execute('''SELECT g.name FROM organization_groups g
                    JOIN user_group_memberships m ON m.group_id=g.id
                    WHERE m.username=%s AND m.relation='member' AND g.is_active=1 ORDER BY g.name''', (target['username'],))
                groups = [row['name'] for row in cursor.fetchall()]
                months = [{'month': m, 'attendance_sheets': 0, 'documents': 0, 'archived_documents': 0,
                           'leaves': [], 'signature_requests': []} for m in range(1, 13)]
                cursor.execute('''SELECT receipt_month,receipt_year,created_at,archived FROM transport_receipts
                    WHERE employee_username=%s''', (target['username'],))
                for row in cursor.fetchall():
                    raw = row['created_at']
                    if isinstance(raw, (datetime, date)):
                        uploaded = raw
                    else:
                        uploaded = datetime.fromisoformat(str(raw).replace('Z', '+00:00'))
                    y, m = row.get('receipt_year') or uploaded.year, row.get('receipt_month') or uploaded.month
                    if int(y) == year and 1 <= int(m) <= 12:
                        months[int(m) - 1]['documents'] += 1
                        months[int(m) - 1]['archived_documents'] += int(bool(row['archived']))
                cursor.execute('''SELECT id,start_date,end_date,start_slot,end_slot,status FROM leave_requests
                    WHERE employee_username=%s AND start_date<=%s AND end_date>=%s ORDER BY start_date,id''',
                    (target['username'], date(year, 12, 31), date(year, 1, 1)))
                for row in cursor.fetchall():
                    start, end = date.fromisoformat(str(row['start_date'])[:10]), date.fromisoformat(str(row['end_date'])[:10])
                    item = {**row, 'start_date': start.isoformat(), 'end_date': end.isoformat()}
                    for month in months:
                        m = month['month']
                        if start <= date(year, m, monthrange(year, m)[1]) and end >= date(year, m, 1):
                            month['leaves'].append(item)
                # Older sheets may live in a manager's workspace; never return that workspace.
                seen, legacy, invalid = set(), set(), 0
                email = str(target.get('email') or '').strip().casefold()
                cursor.execute('SELECT username FROM users WHERE LOWER(email)=LOWER(%s)', (email,))
                email_unique = len(cursor.fetchall()) == 1
                cursor.execute('SELECT user_id,form_content FROM Presence')
                for stored in cursor.fetchall():
                    try:
                        config = stored['form_content']
                        if isinstance(config, str): config = json.loads(config)
                        records = config.get('employes_data', [])
                        if not isinstance(records, list): raise ValueError()
                    except (TypeError, ValueError, AttributeError):
                        invalid += 1
                        continue
                    for index, sheet in enumerate(records):
                        if not isinstance(sheet, dict): continue
                        linked = sheet.get('account_username')
                        if linked != target['username'] and not (
                            not linked and email_unique and email
                            and str(sheet.get('email_employe') or '').strip().casefold() == email
                        ): continue
                        key = str(sheet.get('id') or f"{stored['user_id']}:{index}")
                        m, y = sheet.get('sheet_month'), sheet.get('sheet_year')
                        if not m or not y:
                            legacy.add(key)
                        elif y == year and isinstance(m, int) and 1 <= m <= 12 and (m, key) not in seen:
                            seen.add((m, key)); months[m - 1]['attendance_sheets'] += 1
                signatures, signatures_available = attempts.profile_receipts(target['username'], year)
                for item in signatures:
                    months[item['month'] - 1]['signature_requests'].append(item)
                manager = load_user(cursor, target['manager_username']) if target.get('manager_username') else None
                return {'username': target['username'], 'name': str(target.get('email') or target['username']).split('@')[0],
                        'groups': groups, 'manager': manager['username'] if manager else None, 'year': year,
                        'months': months, 'legacy_unbound_sheets': len(legacy), 'attendance_incomplete': bool(invalid),
                        'signatures_available': signatures_available}
        finally:
            db.close()
