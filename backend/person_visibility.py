"""Daily person catalogs exclude inactive logins and archived linked employees.

This is a visibility rule, not an authentication or historical-access policy.
"""
import sqlite3
import pymysql
from .role_policy import is_super_admin


def account_list_visible(user, actor):
    """System Super Admin identities appear only in their own account catalogs."""
    return bool(user and (not is_super_admin(user) or user['username'] == actor['username']))


def hidden_account_usernames(cursor, actor):
    cursor.execute('SELECT username,role_tags FROM users')
    return [row['username'] for row in cursor.fetchall() if not account_list_visible(row, actor)]


def archived_usernames(cursor):
    try:
        cursor.execute('SELECT linked_username FROM personnel_register WHERE is_archived=1 AND linked_username IS NOT NULL')
        return {row['linked_username'] for row in cursor.fetchall()}
    except pymysql.err.ProgrammingError as error:
        if error.args[0] != 1146:
            raise
    except sqlite3.OperationalError as error:
        if str(error) != 'no such table: personnel_register':
            raise
    return set()  # Older installations may not yet have an employee register.


def annotate_visibility(cursor, users):
    archived = archived_usernames(cursor)
    return [{**user, 'personnel_archived': user['username'] in archived} for user in users if user]


def daily_person(user):
    return bool(user and user.get('is_active', True) and not user.get('login_deleted_at')
                and not user.get('personnel_archived'))
