"""Medical evidence uses the existing live team relation for read access only."""
from .person_visibility import account_list_visible, daily_person
from .role_policy import is_admin


def may_read_employee(cursor, actor, employee, load_user, can_manage):
    username = employee.get('linked_username')
    if is_admin(actor) or username == actor['username']:
        return True
    if not username or employee.get('is_archived') or not can_manage:
        return False
    target = load_user(cursor, username)
    return bool(daily_person(target) and account_list_visible(target, actor)
                and can_manage(cursor, actor, target))


def may_edit_employee(actor, employee):
    return bool(not employee.get('is_archived')
                and (is_admin(actor) or employee.get('linked_username') == actor['username']))


def readable_usernames(cursor, actor, load_user, can_manage):
    names = [actor['username']]
    if can_manage:
        cursor.execute('SELECT username FROM users ORDER BY username')
        for row in cursor.fetchall():
            if row['username'] != actor['username'] and may_read_employee(
                    cursor, actor, {'linked_username': row['username']}, load_user, can_manage):
                names.append(row['username'])
    return names


def record_archive_filter(archived, alias='v'):
    # Personnel archives remain reachable to Admin; visit archives are independent.
    return (f' AND ({alias}.is_archived=1 OR e.is_archived=1)' if archived
            else f' AND {alias}.is_archived=0')
