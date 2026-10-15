"""Authenticated employee department directory, separate from personnel records."""
from fastapi import Depends
from organization_model import build_organization


def directory_people(rows, memberships):
    """Project public directory fields in linear time; discard private user fields."""
    links = {}
    for membership in memberships:
        links.setdefault(membership['username'], []).append(membership)
    people = []
    for row in rows:
        email = str(row.get('email') or row['username'])
        person = {'username': row['username'], 'display_name': email.split('@', 1)[0],
                  'is_ceo': email.casefold() == 'william.tang@univ-ilci.fr',
                  'group_ids': [], 'managed_group_ids': []}
        for membership in links.get(row['username'], []):
            field = {'member': 'group_ids', 'manager': 'managed_group_ids'}.get(membership['relation'])
            if field and int(membership['group_id']) not in person[field]:
                person[field].append(int(membership['group_id']))
        people.append(person)
    return people


def team_members(rows):
    """Deduplicate already authorized membership rows and hide complete email addresses."""
    unique_rows = {row['username']: row for row in rows}
    people = directory_people(unique_rows.values(), rows)
    return [{'name': person['display_name'], 'group_ids': person['group_ids'],
             'managed_group_ids': person['managed_group_ids']} for person in people]


def register_team_directory(app, *, get_db_connection, get_current_user):
    @app.get('/organization-directory')
    def organization_directory(actor=Depends(get_current_user)):
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                cursor.execute('SELECT id,name,is_active FROM organization_groups ORDER BY name')
                groups = cursor.fetchall()
                cursor.execute('SELECT username,email FROM users ORDER BY username')
                rows = cursor.fetchall()
                cursor.execute('SELECT username,group_id,relation FROM user_group_memberships')
                memberships = cursor.fetchall()
                people = directory_people(rows, memberships)
                return {'groups': groups, 'users': people,
                        'organization': build_organization(people, groups)}
        finally:
            connection.close()

    @app.get('/my-team')
    def my_team(actor=Depends(get_current_user)):
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                # Memberships are read from the server, not accepted from the client.
                cursor.execute('''SELECT g.id, g.name FROM organization_groups g
                    JOIN user_group_memberships m ON m.group_id=g.id
                    WHERE m.username=%s AND m.relation='member' AND g.is_active=1
                    ORDER BY g.name''', (actor['username'],))
                groups = cursor.fetchall()
                if not groups:
                    return {'groups': [], 'members': []}
                ids = [int(group['id']) for group in groups]
                placeholders = ','.join(['%s'] * len(ids))
                cursor.execute(f'''SELECT u.username, u.email, m.group_id, m.relation
                    FROM user_group_memberships m JOIN users u ON u.username=m.username
                    WHERE m.group_id IN ({placeholders}) ORDER BY u.username''', tuple(ids))
                rows = cursor.fetchall()
                return {'groups': groups, 'members': team_members(rows)}
        finally:
            connection.close()
