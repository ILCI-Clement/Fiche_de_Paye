"""Permanent document references, sequenced per Paris upload day in MariaDB."""
from datetime import datetime
from zoneinfo import ZoneInfo

PARIS = ZoneInfo("Europe/Paris")


def allocate_document_number(cursor, uploaded_at: datetime) -> str:
    day = uploaded_at.astimezone(PARIS).date()
    # This row lock is held until the receipt insert commits or rolls back.
    cursor.execute("INSERT IGNORE INTO document_daily_sequences (upload_day, last_sequence) VALUES (%s, 0)", (day,))
    cursor.execute("SELECT last_sequence FROM document_daily_sequences WHERE upload_day = %s FOR UPDATE", (day,))
    sequence = int(cursor.fetchone()["last_sequence"]) + 1
    cursor.execute("UPDATE document_daily_sequences SET last_sequence = %s WHERE upload_day = %s", (sequence, day))
    return f"{day:%Y%m%d}-{sequence:04d}"


def ensure_document_number_schema(get_db_connection):
    """Add schema and backfill once, preserving already issued numbers and files."""
    connection = get_db_connection()
    locked = False
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT GET_LOCK(CONCAT(DATABASE(), ':document_numbers'), 60) AS acquired")
            locked = bool(cursor.fetchone()["acquired"])
            if not locked:
                raise RuntimeError("Document numbering migration lock unavailable")
            cursor.execute("SHOW COLUMNS FROM transport_receipts")
            if "document_number" not in {row["Field"] for row in cursor.fetchall()}:
                cursor.execute("ALTER TABLE transport_receipts ADD COLUMN document_number VARCHAR(24) NULL")
            cursor.execute("""CREATE TABLE IF NOT EXISTS document_daily_sequences (
                upload_day DATE NOT NULL PRIMARY KEY,
                last_sequence INT UNSIGNED NOT NULL DEFAULT 0
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
            cursor.execute("SHOW INDEX FROM transport_receipts WHERE Key_name = 'document_number_unique'")
            if not cursor.fetchone():
                cursor.execute("ALTER TABLE transport_receipts ADD UNIQUE KEY document_number_unique (document_number)")
            # UNIX_TIMESTAMP avoids assuming the DB session's display timezone.
            cursor.execute("""SELECT id, UNIX_TIMESTAMP(created_at) AS uploaded_epoch
                FROM transport_receipts WHERE document_number IS NULL ORDER BY created_at, id FOR UPDATE""")
            rows = cursor.fetchall()
            for row in rows:
                uploaded_at = datetime.fromtimestamp(float(row["uploaded_epoch"]), PARIS)
                number = allocate_document_number(cursor, uploaded_at)
                cursor.execute("UPDATE transport_receipts SET document_number = %s WHERE id = %s AND document_number IS NULL", (number, row["id"]))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        if locked:
            with connection.cursor() as cursor:
                cursor.execute("SELECT RELEASE_LOCK(CONCAT(DATABASE(), ':document_numbers'))")
        connection.close()
