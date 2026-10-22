"""Minimal permanent-deletion ledger, independent from the deleted business rows."""
import json


RESOURCE_TYPES = frozenset({'account', 'personnel', 'document', 'contract', 'group',
                            'leave_request', 'annual_interview', 'attendance_sheet'})
COUNT_FIELDS = frozenset({'deleted_records', 'deleted_files', 'period_count', 'history_count',
                         'contract_count', 'document_count', 'sheet_count', 'membership_count',
                         'request_count', 'interview_count', 'affected_people', 'confirmed',
                         'previous_is_archived', 'was_active'})


def ensure_destructive_audit_schema(cursor):
    cursor.execute('''CREATE TABLE IF NOT EXISTS permanent_deletion_events (
        id BIGINT AUTO_INCREMENT PRIMARY KEY,
        actor_username VARCHAR(255) NOT NULL,
        resource_type VARCHAR(32) NOT NULL, resource_id VARCHAR(255) NOT NULL,
        metadata TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY deletion_resource (resource_type,resource_id))''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS resource_archive_events (
        id BIGINT AUTO_INCREMENT PRIMARY KEY, actor_username VARCHAR(255) NOT NULL,
        resource_type VARCHAR(32) NOT NULL, resource_id VARCHAR(255) NOT NULL,
        is_archived TINYINT(1) NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)''')


def append_archive_event(cursor, actor, resource_type, resource_id, is_archived):
    if resource_type not in RESOURCE_TYPES or type(is_archived) is not bool:
        raise ValueError('Invalid archive audit metadata.')
    username = actor['username'] if isinstance(actor, dict) else actor
    cursor.execute('''INSERT INTO resource_archive_events
        (actor_username,resource_type,resource_id,is_archived) VALUES (%s,%s,%s,%s)''',
        (username, resource_type, str(resource_id), is_archived))


def append_deletion_event(cursor, actor, resource_type, resource_id, metadata=None):
    """Write in the caller's transaction; never accept a free-form business snapshot."""
    values = metadata or {}
    if (resource_type not in RESOURCE_TYPES or len(str(resource_id)) > 255
            or not isinstance(values, dict) or not set(values).issubset(COUNT_FIELDS)
            or any(type(value) not in (int, bool) or int(value) < 0 for value in values.values())):
        raise ValueError('Invalid deletion audit metadata.')
    username = actor['username'] if isinstance(actor, dict) else actor
    cursor.execute('''INSERT INTO permanent_deletion_events
        (actor_username,resource_type,resource_id,metadata) VALUES (%s,%s,%s,%s)''',
        (username, resource_type, str(resource_id), json.dumps(values, sort_keys=True)))
