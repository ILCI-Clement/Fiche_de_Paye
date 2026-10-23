"""Presence API with organization-aware authorization.

This module is deployed as ``/var/www/presence-app/main.py``. It owns the
MariaDB schema migration for groups and enforces access server-side.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import smtplib
import threading
import time
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from pathlib import Path
from typing import Any

import bcrypt
import pymysql
from fastapi import Depends, FastAPI, Header, HTTPException
from pydantic import BaseModel, ConfigDict, EmailStr, Field

if __package__:
    from .login_guard import (COOLDOWN_SECONDS, check_cooldown, cooldown_error,
                              ensure_login_guard_schema, lock_login_attempt, record_failure, reset_failures)
    from .person_names import (annotate_person_names, load_person_accounts,
                               load_person_names, person_display_name)
    from .role_policy import (VALID_ROLES, ROLE_PRIORITY, active_super_admin_count, can_assign_roles,
                              can_manage_account, canonical_role_tags, checked_resource_actor,
                              ensure_role_audit_schema, is_admin, is_super_admin, record_role_change,
                              reserve_deleted_account_name, username_is_reserved)
else:
    from login_guard import (COOLDOWN_SECONDS, check_cooldown, cooldown_error,
                             ensure_login_guard_schema, lock_login_attempt, record_failure, reset_failures)
    from person_names import (annotate_person_names, load_person_accounts,
                              load_person_names, person_display_name)
    from role_policy import (VALID_ROLES, ROLE_PRIORITY, active_super_admin_count, can_assign_roles,
                             can_manage_account, canonical_role_tags, checked_resource_actor,
                             ensure_role_audit_schema, is_admin, is_super_admin, record_role_change,
                             reserve_deleted_account_name, username_is_reserved)


app = FastAPI(title="ILCI · Espace RH API")
VALID_EMPLOYEE_TYPES = {"salarie", "stagiaire"}
SESSION_DURATION_SECONDS = 8 * 60 * 60
REMEMBER_SESSION_DURATION_DAYS = 30
PASSWORD_RESET_DURATION_MINUTES = 15
MAX_FICHE_ATTACHMENT_BYTES = 10 * 1024 * 1024
MAX_FICHE_CONFIG_BYTES = 5 * 1024 * 1024
MAX_FICHE_EMPLOYEES = 100
MAX_TRANSPORT_RECEIPT_BYTES = 10 * 1024 * 1024
ESIGN_DOCUMENT_DURATION_HOURS = 24
REMINDER_CHECK_INTERVAL_SECONDS = 6 * 60 * 60
_reminder_worker_started = False


def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


API_TOKEN = required_env("PRESENCE_API_TOKEN")
DB_CONFIG = {
    "host": required_env("PRESENCE_DB_HOST"),
    "user": required_env("PRESENCE_DB_USER"),
    "password": required_env("PRESENCE_DB_PASSWORD"),
    "database": required_env("PRESENCE_DB_DATABASE"),
    "cursorclass": pymysql.cursors.DictCursor,
    "autocommit": False,
}
SMTP_SERVER = required_env("PRESENCE_SMTP_SERVER")
SMTP_PORT = int(required_env("PRESENCE_SMTP_PORT"))
SMTP_USER = required_env("PRESENCE_SMTP_USER")
SMTP_PASSWORD = required_env("PRESENCE_SMTP_PASSWORD")
STREAMLIT_APP_URL = required_env("PRESENCE_APP_URL")
ESIGN_DOCUMENTS_DIR = Path(
    os.getenv("PRESENCE_ESIGN_DOCUMENTS_DIR", "/var/lib/presence-app/esign-documents")
)
TRANSPORT_RECEIPTS_DIR = Path(
    os.getenv("PRESENCE_TRANSPORT_RECEIPTS_DIR", "/var/lib/presence-app/transport-receipts")
)


class ProfileUpdateRequest(BaseModel):
    current_password: str
    new_username: str | None = None
    new_email: EmailStr | None = None
    new_password: str | None = None
    confirm_password: str | None = None


class GroupRequest(BaseModel):
    name: str
    is_active: bool = True


class DirectManagerRequest(BaseModel):
    manager_id: str


class AccountPurgeRequest(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True, hide_input_in_errors=True)
    confirmation: str = Field(min_length=1, max_length=255)
    reason: str = Field(min_length=1, max_length=500)


class RememberSessionRequest(BaseModel):
    remember_token: str


class AnnualInterviewUpdateRequest(BaseModel):
    due_date: date | None = None
    notes: str | None = None
    completed: bool | None = None


class ContractUpdateRequest(BaseModel):
    contract_start_date: date | None = None
    contract_end_date: date | None = None
    is_cdi: bool = False


class ContractReminderSettingsRequest(BaseModel):
    auto_send: bool


class ContractReminderResendRequest(BaseModel):
    reminder_key: str
    contract_end_date: date


def get_db_connection() -> pymysql.Connection:
    return pymysql.connect(**DB_CONFIG)


def normalize_role(user: dict[str, Any]) -> str:
    role = user.get("role")
    if role in VALID_ROLES:
        return str(role)
    return "Admin" if user.get("is_admin") else "Responsable"


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def hash_remember_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def create_remember_token(cursor: pymysql.cursors.Cursor, username: str) -> str:
    token = secrets.token_urlsafe(48)
    expires_at = datetime.now() + timedelta(days=REMEMBER_SESSION_DURATION_DAYS)
    cursor.execute(
        "UPDATE users SET remember_token_hash = %s, remember_token_expires = %s WHERE username = %s",
        (hash_remember_token(token), expires_at, username),
    )
    return token


def verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except (TypeError, ValueError):
        return False


def encode_segment(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).rstrip(b"=").decode("ascii")


def decode_segment(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def create_session_token(username: str, session_version: int = 0) -> str:
    payload = json.dumps(
        {"username": username, "expires_at": int(time.time()) + SESSION_DURATION_SECONDS, "session_version": session_version},
        separators=(",", ":"),
    ).encode("utf-8")
    payload_part = encode_segment(payload)
    signature = hmac.new(API_TOKEN.encode("utf-8"), payload_part.encode("ascii"), hashlib.sha256).digest()
    return f"{payload_part}.{encode_segment(signature)}"


def read_session_token(token: str, *, with_version: bool = False):
    try:
        payload_part, signature_part = token.split(".", 1)
        expected = hmac.new(API_TOKEN.encode("utf-8"), payload_part.encode("ascii"), hashlib.sha256).digest()
        if not hmac.compare_digest(expected, decode_segment(signature_part)):
            raise ValueError("Invalid signature")
        payload = json.loads(decode_segment(payload_part))
        if int(payload["expires_at"]) < int(time.time()):
            raise ValueError("Expired token")
        return {"username": str(payload["username"]), "session_version": int(payload.get("session_version", 0))} if with_version else str(payload["username"])
    except (KeyError, TypeError, ValueError, json.JSONDecodeError):
        raise HTTPException(status_code=401, detail="Session invalide ou expirée.") from None


def ensure_organization_schema() -> None:
    """Apply additive, idempotent schema changes without touching existing data."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            if __package__:
                from .account_lifecycle import ensure_account_lifecycle_schema
            else:
                from account_lifecycle import ensure_account_lifecycle_schema
            ensure_account_lifecycle_schema(cursor)
            ensure_role_audit_schema(cursor)
            cursor.execute("SHOW COLUMNS FROM users")
            columns = {item["Field"] for item in cursor.fetchall()}
            ensure_login_guard_schema(cursor, columns)
            if "employee_type" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN employee_type VARCHAR(20) NULL")
            if "manager_username" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN manager_username VARCHAR(50) NULL")
            if "role_tags" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN role_tags TEXT NULL")
            if "remember_token_hash" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN remember_token_hash CHAR(64) NULL")
            if "remember_token_expires" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN remember_token_expires DATETIME NULL")
            if "contract_start_date" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN contract_start_date DATE NULL")
            if "contract_end_date" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN contract_end_date DATE NULL")
            if "is_cdi" not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN is_cdi TINYINT(1) NOT NULL DEFAULT 0")
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS organization_groups (
                    id INT NOT NULL AUTO_INCREMENT,
                    name VARCHAR(100) NOT NULL,
                    is_active TINYINT(1) NOT NULL DEFAULT 1,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (id),
                    UNIQUE KEY organization_groups_name_unique (name)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS user_group_memberships (
                    username VARCHAR(50) NOT NULL,
                    group_id INT NOT NULL,
                    relation VARCHAR(20) NOT NULL,
                    PRIMARY KEY (username, group_id, relation),
                    KEY user_group_memberships_group_index (group_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS transport_receipts (
                    id INT NOT NULL AUTO_INCREMENT,
                    employee_username VARCHAR(50) NOT NULL,
                    uploaded_by VARCHAR(50) NOT NULL,
                    original_filename VARCHAR(255) NOT NULL,
                    stored_filename VARCHAR(255) NOT NULL,
                    mime_type VARCHAR(100) NOT NULL,
                    size_bytes INT NOT NULL,
                    document_type VARCHAR(20) NULL,
                    other_type VARCHAR(120) NULL,
                    receipt_month INT NULL,
                    receipt_year INT NULL,
                    archived TINYINT(1) NOT NULL DEFAULT 0,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    archived_at DATETIME NULL,
                    archived_by VARCHAR(50) NULL,
                    PRIMARY KEY (id),
                    KEY transport_receipts_employee_index (employee_username),
                    KEY transport_receipts_uploaded_index (uploaded_by),
                    KEY transport_receipts_archived_index (archived, created_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            # Keep old receipts unclassified; only new uploads receive a category.
            cursor.execute("SHOW COLUMNS FROM transport_receipts")
            receipt_columns = {column["Field"] for column in cursor.fetchall()}
            if "document_type" not in receipt_columns:
                cursor.execute("ALTER TABLE transport_receipts ADD COLUMN document_type VARCHAR(20) NULL")
            if "other_type" not in receipt_columns:
                cursor.execute("ALTER TABLE transport_receipts ADD COLUMN other_type VARCHAR(120) NULL")
            if "receipt_month" not in receipt_columns:
                cursor.execute("ALTER TABLE transport_receipts ADD COLUMN receipt_month INT NULL")
            if "receipt_year" not in receipt_columns:
                cursor.execute("ALTER TABLE transport_receipts ADD COLUMN receipt_year INT NULL")
            cursor.execute(
                "UPDATE transport_receipts SET receipt_month = MONTH(created_at), receipt_year = YEAR(created_at) "
                "WHERE receipt_month IS NULL OR receipt_year IS NULL"
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS annual_interviews (
                    id INT NOT NULL AUTO_INCREMENT,
                    employee_username VARCHAR(50) NOT NULL,
                    interview_year INT NOT NULL,
                    due_date DATE NOT NULL,
                    notes TEXT NULL,
                    completed TINYINT(1) NOT NULL DEFAULT 0,
                    completed_at DATETIME NULL,
                    completed_by VARCHAR(50) NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (id),
                    UNIQUE KEY annual_interviews_employee_year_unique (employee_username, interview_year),
                    KEY annual_interviews_due_index (completed, due_date),
                    KEY annual_interviews_employee_index (employee_username)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS contract_end_reminders (
                    id INT NOT NULL AUTO_INCREMENT,
                    employee_username VARCHAR(50) NOT NULL,
                    contract_end_date DATE NOT NULL,
                    recipient_email VARCHAR(255) NOT NULL,
                    sent_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (id),
                    UNIQUE KEY contract_end_reminder_unique (employee_username, contract_end_date, recipient_email),
                    KEY contract_end_reminder_employee_index (employee_username)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    setting_key VARCHAR(100) NOT NULL,
                    setting_value VARCHAR(255) NOT NULL,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    PRIMARY KEY (setting_key)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            # Never enable unattended e-mails merely because this feature is deployed.
            cursor.execute(
                "INSERT IGNORE INTO app_settings (setting_key, setting_value) VALUES ('contract_reminders_auto_send', '0')"
            )
        connection.commit()
    finally:
        connection.close()


@app.on_event("startup")
def migrate_database() -> None:
    ensure_organization_schema()
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            sync_all_contract_dates(cursor)
        connection.commit()
    finally:
        connection.close()
    start_contract_reminder_worker()


def load_user(cursor: pymysql.cursors.Cursor, username: str, *, person_names=None, account_names=None) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT username, email, password_hash, is_admin, role, role_tags, employee_type, manager_username,
               contract_start_date, contract_end_date, is_cdi, created_at, is_active, session_version, login_deleted_at
        FROM users WHERE username = %s
        """,
        (username,),
    )
    user = cursor.fetchone()
    if not user:
        return None
    raw_tags = user.get("role_tags")
    try:
        tags = {str(tag) for tag in json.loads(raw_tags)} if raw_tags else set()
    except (TypeError, json.JSONDecodeError):
        tags = set()
    tags = tags & VALID_ROLES or {normalize_role(user)}
    user["role_tags"] = canonical_role_tags(tags)
    user["role"] = user["role_tags"][0]
    user["employee_type"] = user.get("employee_type") or "salarie"
    cursor.execute(
        "SELECT group_id FROM user_group_memberships WHERE username = %s AND relation = 'member' ORDER BY group_id",
        (username,),
    )
    user["group_ids"] = [row["group_id"] for row in cursor.fetchall()]
    cursor.execute(
        "SELECT group_id FROM user_group_memberships WHERE username = %s AND relation = 'manager' ORDER BY group_id",
        (username,),
    )
    user["managed_group_ids"] = [row["group_id"] for row in cursor.fetchall()]
    cursor.execute("SHOW TABLES LIKE 'personnel_contracts'")
    if cursor.fetchone():
        cursor.execute("SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date", (username,))
        history = cursor.fetchall()
        if history:
            current = next((c for c in history if not c.get('is_archived') and c["start_date"] <= date.today() and (c["end_date"] is None or c["end_date"] >= date.today())), None)
            user["contract_start_date"] = current["start_date"] if current else None
            user["contract_end_date"] = current["end_date"] if current else None
            user["is_cdi"] = current["is_cdi"] if current else False
            if current:
                user["employee_type"] = "salarie" if current["employee_type"] == "alternance" else current["employee_type"]
    names = person_names
    if names is None:
        names = getattr(cursor, '_presence_person_names', None)
        if names is None:
            names = load_person_names(cursor)
            try:
                cursor._presence_person_names = names
            except (AttributeError, TypeError):
                pass  # Some cursor adapters cannot carry request-scoped attributes.
    annotate_person_names([user], names)
    manager_username = user.get('manager_username')
    if manager_username:
        manager_names = account_names
        if manager_names is None:
            manager_names = getattr(cursor, '_presence_account_display_names', None)
            if manager_names is None:
                manager_names = {key: account['display_name'] for key, account in
                                 load_person_accounts(cursor, person_names=names).items()}
                try:
                    cursor._presence_account_display_names = manager_names
                except (AttributeError, TypeError):
                    pass
        user['manager_display_name'] = manager_names.get(manager_username, 'Ancien compte')
    else:
        user['manager_display_name'] = None
    return user


def public_user(user: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": user["username"],
        "username": user["username"],
        "email": user["email"],
        "full_name": user.get('full_name'),
        "name": person_display_name(user),
        "display_name": person_display_name(user),
        "role": user["role"],
        "role_tags": user["role_tags"],
        "is_admin": is_admin(user),
        "is_active": bool(user.get("is_active", True)),
        "login_deleted_at": user.get("login_deleted_at"),
        "employee_type": user.get("employee_type") or "salarie",
        "manager_id": user.get("manager_username"),
        "manager_display_name": user.get('manager_display_name'),
        "group_ids": user.get("group_ids", []),
        "managed_group_ids": user.get("managed_group_ids", []),
        "contract_start_date": user.get("contract_start_date"),
        "contract_end_date": user.get("contract_end_date"),
        "is_cdi": bool(user.get("is_cdi")),
        "created_at": user.get("created_at"),
    }


def parse_optional_date(value: Any) -> date | None:
    """Return an ISO date when an existing fiche contains a usable date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None
    return None


def fiche_display_name(employee: dict[str, Any]) -> str:
    """Return the employee-facing name from a fiche record."""
    if employee.get("type") == "Stagiaire":
        return " ".join(
            value.strip()
            for value in (str(employee.get("prenom_stagiaire") or ""), str(employee.get("nom_stagiaire") or ""))
            if value.strip()
        )
    return str(employee.get("nom") or "").strip()


def fiche_contract_values(employee: dict[str, Any]) -> tuple[date | None, date | None, bool]:
    if employee.get("type") == "Stagiaire":
        return (
            parse_optional_date(employee.get("dds")),
            parse_optional_date(employee.get("fds")),
            False,
        )
    is_cdi = bool(employee.get("cdi"))
    return (
        parse_optional_date(employee.get("ddc")),
        None if is_cdi else parse_optional_date(employee.get("fdc")),
        is_cdi,
    )


def resolve_fiche_account_username(cursor: pymysql.cursors.Cursor, employee: dict[str, Any]) -> str | None:
    """Resolve a fiche to an account; the recorded employee name wins over an e-mail fallback."""
    employee_name = fiche_display_name(employee)
    if employee_name:
        cursor.execute("SELECT username FROM users WHERE LOWER(username) = LOWER(%s)", (employee_name,))
        row = cursor.fetchone()
        if row:
            return str(row["username"])
    account_username = str(employee.get("account_username") or "").strip()
    if account_username:
        cursor.execute("SELECT username FROM users WHERE LOWER(username) = LOWER(%s)", (account_username,))
        row = cursor.fetchone()
        if row:
            return str(row["username"])
    email = str(employee.get("email_employe") or "").strip()
    if email:
        cursor.execute("SELECT username FROM users WHERE email = %s", (email,))
        row = cursor.fetchone()
        if row:
            return str(row["username"])
    return None


def clear_mislinked_contract(cursor: pymysql.cursors.Cursor, username: str, employee: dict[str, Any]) -> None:
    """Remove a contract only when it exactly matches the fiche that was previously linked by mistake."""
    start_date, end_date, is_cdi = fiche_contract_values(employee)
    cursor.execute(
        """
        UPDATE users
        SET contract_start_date = NULL, contract_end_date = NULL, is_cdi = 0
        WHERE username = %s
          AND contract_start_date <=> %s
          AND contract_end_date <=> %s
          AND is_cdi = %s
        """,
        (username, start_date, end_date, is_cdi),
    )


def sync_contract_dates_from_config(cursor, config):
    """Contract dates are maintained through contract history only."""
    return


def link_fiche_records_to_users(cursor: pymysql.cursors.Cursor, config: dict[str, Any]) -> bool:
    """Persist a stable account link for fiche records that originate from an account."""
    changed = False
    for employee in config.get("employes_data", []) if isinstance(config, dict) else []:
        if not isinstance(employee, dict):
            continue
        resolved_username = resolve_fiche_account_username(cursor, employee)
        previous_username = str(employee.get("account_username") or "").strip()
        if resolved_username and previous_username != resolved_username:
            if previous_username:
                clear_mislinked_contract(cursor, previous_username, employee)
            employee["account_username"] = resolved_username
            changed = True
    return changed


def sync_all_contract_dates(cursor: pymysql.cursors.Cursor) -> None:
    """Backfill contract fields from pre-existing saved attendance configurations."""
    cursor.execute("SELECT user_id, form_content FROM Presence")
    for row in cursor.fetchall():
        content = row.get("form_content")
        try:
            config = content if isinstance(content, dict) else json.loads(content)
        except (TypeError, json.JSONDecodeError):
            continue
        linked = link_fiche_records_to_users(cursor, config)
        sync_contract_dates_from_config(cursor, config)
        if linked:
            cursor.execute(
                "UPDATE Presence SET form_content = %s WHERE user_id = %s",
                (json.dumps(config, default=str), row["user_id"]),
            )


def contract_deadlines_from_fiches(cursor: pymysql.cursors.Cursor, today: date) -> list[dict[str, Any]]:
    """Return all fiche contracts ending within ten days, including people without an app account."""
    records: list[dict[str, Any]] = []
    represented_accounts: set[str] = set()
    seen_references: set[str] = set()
    users = all_users(cursor)
    accounts = {user['username']: user for user in users}
    cursor.execute("SELECT user_id, form_content FROM Presence")
    for row in cursor.fetchall():
        content = row.get("form_content")
        try:
            config = content if isinstance(content, dict) else json.loads(content)
        except (TypeError, json.JSONDecodeError):
            continue
        for employee in config.get("employes_data", []) if isinstance(config, dict) else []:
            if not isinstance(employee, dict) or employee.get('is_archived'):
                continue
            account_username = resolve_fiche_account_username(cursor, employee)
            cursor.execute("SHOW TABLES LIKE 'personnel_contracts'")
            if account_username and cursor.fetchone():
                cursor.execute("SELECT id FROM personnel_contracts WHERE username=%s LIMIT 1", (account_username,))
                if cursor.fetchone():
                    continue
            account = accounts.get(account_username)
            display_name = person_display_name(account or {'email': employee.get('email_employe')})
            _, end_date, is_cdi = fiche_contract_values(employee)
            if not isinstance(end_date, date) or is_cdi:
                continue
            days_remaining = (end_date - today).days
            if not 0 <= days_remaining <= 10:
                continue
            identity_name = fiche_display_name(employee) or account_username or display_name
            record_id = f"{row['user_id']}:{employee.get('id') or identity_name.casefold()}"
            reminder_key = account_username or f"fiche:{hashlib.sha256(record_id.encode()).hexdigest()[:32]}"
            if reminder_key in seen_references:
                continue
            seen_references.add(reminder_key)
            if account_username:
                represented_accounts.add(account_username)
            manager_username = account.get("manager_username") if account else None
            manager = accounts.get(manager_username)
            manager_email = str(employee.get("email_responsable") or "").strip()
            if not manager_email and manager:
                manager_email = str(manager["email"])
            records.append(
                {
                    "reminder_key": reminder_key,
                    "employee_username": account_username,
                    "employee_name": display_name,
                    "contract_end_date": end_date,
                    "days_remaining": days_remaining,
                    "manager_username": manager_username,
                    "manager_name": person_display_name(manager or {'email': manager_email}),
                    "manager_email": manager_email,
                }
            )

    # Keep contracts maintained directly on an account when that account has no fiche yet.
    for user in users:
        if user["username"] in represented_accounts:
            continue
        end_date = user.get("contract_end_date")
        if not isinstance(end_date, date) or bool(user.get("is_cdi")):
            continue
        days_remaining = (end_date - today).days
        if not 0 <= days_remaining <= 10:
            continue
        manager_username = user.get("manager_username")
        manager = accounts.get(manager_username)
        records.append(
            {
                "reminder_key": user["username"],
                "employee_username": user["username"],
                "employee_name": person_display_name(user),
                "contract_end_date": end_date,
                "days_remaining": days_remaining,
                "manager_username": manager_username,
                "manager_name": person_display_name(manager),
                "manager_email": str(manager["email"]) if manager else "",
            }
        )
    return records


def eligible_for_annual_interview(user: dict[str, Any]) -> bool:
    # The Admin label is an application permission, not an employment status:
    # administrators also need their yearly interview to be tracked.
    return bool(user.get('is_active', True) and not user.get('login_deleted_at')
                and {"Admin", "Employe", "Responsable"} & set(user["role_tags"]))


def ensure_annual_interviews(cursor: pymysql.cursors.Cursor, year: int) -> None:
    """Create one annual-interview task per active personnel account and year."""
    for user in all_users(cursor):
        if user and eligible_for_annual_interview(user):
            cursor.execute(
                """
                INSERT IGNORE INTO annual_interviews (employee_username, interview_year, due_date)
                VALUES (%s, %s, %s)
                """,
                (user["username"], year, date(year, 12, 31)),
            )


def all_users(cursor: pymysql.cursors.Cursor) -> list[dict[str, Any]]:
    cursor.execute("SELECT username,email FROM users ORDER BY username")
    accounts = cursor.fetchall()
    names = load_person_names(cursor, [account['username'] for account in accounts])
    labels = {account['username']: person_display_name(account, names) for account in accounts}
    return [user for account in accounts if (user := load_user(cursor, str(account['username']),
                                                             person_names=names, account_names=labels))]


def users_visible_to_actor(cursor: pymysql.cursors.Cursor, actor: dict[str, Any], users: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if "Admin" in actor["role_tags"]:
        return users
    return [
        user
        for user in users
        if user["username"] == actor["username"] or user_can_manage(cursor, actor, user)
    ]


def get_current_user(authorization: str | None = Header(default=None)) -> dict[str, Any]:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Authentification requise.")
    session = read_session_token(authorization.removeprefix("Bearer "), with_version=True)
    username = session["username"]
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            user = load_user(cursor, username)
        if not user or not user.get("is_active", True) or user.get("login_deleted_at") or user.get("session_version", 0) != session["session_version"]:
            raise HTTPException(status_code=401, detail="Session invalide ou compte désactivé.")
        return user
    finally:
        connection.close()


def require_roles(*roles: str):
    def validate(user: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
        if not (set(roles) & set(user["role_tags"])):
            raise HTTPException(status_code=403, detail="Autorisation insuffisante.")
        return user

    return validate


def user_can_manage(cursor: pymysql.cursors.Cursor, actor: dict[str, Any], target: dict[str, Any]) -> bool:
    if "Admin" in actor["role_tags"]:
        return True
    if "Responsable" not in actor["role_tags"] or "Employe" not in target["role_tags"]:
        return False
    if target.get("manager_username") == actor["username"]:
        return True
    return bool(set(actor["managed_group_ids"]) & set(target["group_ids"]))


def is_valid_direct_manager(user: dict[str, Any] | None) -> bool:
    """An administrator can also be assigned as an employee's direct manager."""
    return bool(user and user.get("is_active", True) and ({"Admin", "Responsable"} & set(user["role_tags"])))


def would_create_management_cycle(
    cursor: pymysql.cursors.Cursor,
    employee_username: str,
    manager_username: str,
) -> bool:
    """Return whether assigning the manager would create a direct-manager cycle."""
    current_username: str | None = manager_username
    visited: set[str] = set()
    while current_username and current_username not in visited:
        if current_username == employee_username:
            return True
        visited.add(current_username)
        current_user = load_user(cursor, current_username)
        current_username = current_user.get("manager_username") if current_user else None
    return False


def requested_role_tags(payload: dict[str, Any], fallback: list[str] | None = None) -> list[str]:
    raw_tags = payload.get("role_tags", fallback or [payload.get("new_role", "Employe")])
    if not isinstance(raw_tags, (list, tuple, set)):
        raise HTTPException(status_code=400, detail="Les étiquettes d'identité sont invalides.")
    tags = {str(tag) for tag in (raw_tags or [])}
    if not tags or not tags.issubset(VALID_ROLES):
        raise HTTPException(status_code=400, detail="Les étiquettes d'identité sont invalides.")
    return canonical_role_tags(tags)


def validate_group_ids(cursor: pymysql.cursors.Cursor, group_ids: list[Any], *, active_only: bool = True) -> list[int]:
    normalized = sorted({int(group_id) for group_id in group_ids})
    if not normalized:
        return []
    placeholders = ", ".join(["%s"] * len(normalized))
    predicate = "AND is_active = 1" if active_only else ""
    cursor.execute(f"SELECT id FROM organization_groups WHERE id IN ({placeholders}) {predicate}", tuple(normalized))
    found = {row["id"] for row in cursor.fetchall()}
    if found != set(normalized):
        raise HTTPException(status_code=400, detail="Un ou plusieurs Groupes sont invalides ou inactifs.")
    return normalized


def replace_memberships(cursor: pymysql.cursors.Cursor, username: str, relation: str, group_ids: list[int]) -> None:
    cursor.execute("DELETE FROM user_group_memberships WHERE username = %s AND relation = %s", (username, relation))
    if group_ids:
        cursor.executemany(
            "INSERT INTO user_group_memberships (username, group_id, relation) VALUES (%s, %s, %s)",
            [(username, group_id, relation) for group_id in group_ids],
        )


def require_target_access(cursor: pymysql.cursors.Cursor, actor: dict[str, Any], target_username: str, *, edit: bool) -> dict[str, Any]:
    target = load_user(cursor, target_username)
    if not target:
        raise HTTPException(status_code=404, detail="Utilisateur introuvable.")
    if target_username == actor["username"]:
        if edit and not ({"Admin", "Responsable"} & set(actor["role_tags"])):
            raise HTTPException(status_code=403, detail="Un employé ne peut pas modifier sa fiche de présence.")
        return target
    if user_can_manage(cursor, actor, target):
        return target
    raise HTTPException(status_code=403, detail="Cet utilisateur ne fait pas partie de votre périmètre.")


def send_password_reset_email(email: str, token: str) -> bool:
    reset_url = f"{STREAMLIT_APP_URL}/?token={token}"
    message = MIMEMultipart()
    message["From"] = SMTP_USER
    message["To"] = email
    message["Subject"] = "Réinitialisation de votre mot de passe"
    message.attach(
        MIMEText(
            f"""<h3>Demande de réinitialisation de mot de passe</h3>
            <p>Bonjour,</p>
            <p>Une demande de réinitialisation a été effectuée pour votre compte ILCI · Espace RH.</p>
            <p><a href=\"{reset_url}\">Réinitialiser mon mot de passe</a></p>
            <p>Ce lien est valable pendant 15 minutes.</p>""",
            "html",
        )
    )
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, email, message.as_string())
        return True
    except smtplib.SMTPException:
        return False


def contract_reminders_auto_enabled(cursor: pymysql.cursors.Cursor) -> bool:
    cursor.execute(
        "SELECT setting_value FROM app_settings WHERE setting_key = 'contract_reminders_auto_send'"
    )
    setting = cursor.fetchone()
    return bool(setting and str(setting["setting_value"]).strip() == "1")


def send_contract_end_reminder_email(recipient_email: str, contract: dict[str, Any]) -> bool:
    """Send a single, actionable contract-end reminder to an admin or direct manager."""
    end_date = contract.get("contract_end_date")
    formatted_end_date = end_date.strftime("%d/%m/%Y") if isinstance(end_date, date) else str(end_date)
    message = MIMEMultipart()
    message["From"] = SMTP_USER
    message["To"] = recipient_email
    message["Subject"] = f"Rappel : fin de contrat de {contract['employee_name']}"
    employee_label = escape(str(contract['employee_name']))
    message.attach(
        MIMEText(
            f"""<h3>Fin de contrat à anticiper</h3>
            <p>Bonjour,</p>
            <p>Le contrat de <strong>{employee_label}</strong> se termine le
            <strong>{formatted_end_date}</strong> ({contract['days_remaining']} jour(s) restant(s)).</p>
            <p>Merci de vérifier la situation dans le Tableau de bord de Fiches de présence.</p>""",
            "html",
        )
    )
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, recipient_email, message.as_string())
        return True
    except smtplib.SMTPException:
        return False


def contract_reminder_recipients(cursor: pymysql.cursors.Cursor, contract: dict[str, Any]) -> set[str]:
    """Return the administrators and direct manager who must receive a reminder."""
    users = all_users(cursor)
    inactive_emails = {str(user.get("email") or "").strip().casefold() for user in users if not user.get("is_active", True)}
    recipients = {
        str(user["email"]).strip()
        for user in users
        if user.get("is_active", True) and "Admin" in user["role_tags"] and str(user.get("email") or "").strip()
    }
    manager_email = str(contract.get("manager_email") or "").strip()
    if manager_email and manager_email.casefold() not in inactive_emails:
        recipients.add(manager_email)
    return recipients


def claim_contract_end_reminder(
    cursor: pymysql.cursors.Cursor,
    contract: dict[str, Any],
    recipient_email: str,
) -> bool:
    """Reserve a normal reminder before sending so concurrent workers cannot duplicate it."""
    cursor.execute(
        """
        INSERT IGNORE INTO contract_end_reminders (employee_username, contract_end_date, recipient_email)
        VALUES (%s, %s, %s)
        """,
        (contract["reminder_key"], contract["contract_end_date"], recipient_email),
    )
    return cursor.rowcount == 1


def release_contract_end_reminder_claim(
    cursor: pymysql.cursors.Cursor,
    contract: dict[str, Any],
    recipient_email: str,
) -> None:
    """Allow a later attempt only when SMTP explicitly reports a failed delivery."""
    cursor.execute(
        """
        DELETE FROM contract_end_reminders
        WHERE employee_username = %s AND contract_end_date = %s AND recipient_email = %s
        """,
        (contract["reminder_key"], contract["contract_end_date"], recipient_email),
    )


def run_contract_end_reminders(*, force: bool = False) -> dict[str, int | bool]:
    """Send only on explicit request unless the administrator enabled automatic delivery."""
    today = date.today()
    sent_count = 0
    contract_count = 0
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            if not force and not contract_reminders_auto_enabled(cursor):
                return {"sent": 0, "contracts": 0, "automatic": False}
            for contract in contract_deadlines_from_fiches(cursor, today):
                contract_count += 1
                for recipient_email in contract_reminder_recipients(cursor, contract):
                    if not claim_contract_end_reminder(cursor, contract, recipient_email):
                        continue
                    # Commit the reservation before SMTP. If this process stops after delivery,
                    # a later worker will not risk sending the same HR reminder again.
                    connection.commit()
                    if send_contract_end_reminder_email(recipient_email, contract):
                        sent_count += 1
                    else:
                        release_contract_end_reminder_claim(cursor, contract, recipient_email)
                        connection.commit()
        connection.commit()
        return {"sent": sent_count, "contracts": contract_count, "automatic": not force}
    finally:
        connection.close()


def contract_reminder_worker() -> None:
    while True:
        try:
            run_contract_end_reminders()
        except Exception:
            # A transient mail or database outage must not stop the API service.
            pass
        time.sleep(REMINDER_CHECK_INTERVAL_SECONDS)


def start_contract_reminder_worker() -> None:
    global _reminder_worker_started
    if _reminder_worker_started:
        return
    _reminder_worker_started = True
    threading.Thread(target=contract_reminder_worker, name="contract-reminders", daemon=True).start()


@app.post("/login")
def login(payload: dict[str, Any]) -> dict[str, Any]:
    identifier = str(payload.get("username", "")).strip()
    password = str(payload.get("password", ""))
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            # Resolve to the canonical account before issuing tokens. Ambiguous
            # username/e-mail matches must never silently select an identity.
            cursor.execute(
                "SELECT username FROM users WHERE LOWER(TRIM(username)) = LOWER(%s) "
                "OR LOWER(TRIM(email)) = LOWER(%s) LIMIT 2",
                (identifier, identifier),
            )
            matches = cursor.fetchall() if identifier else []
            if len(matches) != 1:
                raise HTTPException(status_code=401, detail="Identifiants incorrects.")
            username = matches[0]["username"]
        # Discard the resolution snapshot before the locking read. load_user below
        # must see current account status/roles, not a snapshot from before the lock.
        connection.commit()
        with connection.cursor() as cursor:
            state = lock_login_attempt(cursor, username, identifier)
            if not state:
                raise HTTPException(status_code=401, detail="Identifiants incorrects.")
            user = load_user(cursor, username)
            if not user or not user.get("is_active", True) or user.get("login_deleted_at"):
                raise HTTPException(status_code=401, detail="Identifiants incorrects.")
            check_cooldown(state)
            if not verify_password(password, user["password_hash"]):
                locked = record_failure(cursor, username, state)
                connection.commit()  # Persist the failure before returning an HTTP error.
                if locked:
                    raise cooldown_error(COOLDOWN_SECONDS)
                raise HTTPException(status_code=401, detail="Identifiants incorrects.")
            reset_failures(cursor, username)
        response = public_user(user)
        response["status"] = "success"
        response["auth_token"] = create_session_token(username, user.get("session_version", 0))
        if bool(payload.get("remember_me")):
            with connection.cursor() as cursor:
                response["remember_token"] = create_remember_token(cursor, username)
        connection.commit()
        return response
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


@app.post("/restore-session")
def restore_session(payload: RememberSessionRequest) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT username FROM users WHERE remember_token_hash = %s AND remember_token_expires > %s",
                (hash_remember_token(payload.remember_token), datetime.now()),
            )
            row = cursor.fetchone()
            user = load_user(cursor, row["username"]) if row else None
        if not user or not user.get("is_active", True) or user.get("login_deleted_at"):
            raise HTTPException(status_code=401, detail="Session persistante expirée.")
        response = public_user(user)
        response["status"] = "success"
        response["auth_token"] = create_session_token(user["username"], user.get("session_version", 0))
        return response
    finally:
        connection.close()


@app.post("/forget-session")
def forget_session(payload: RememberSessionRequest) -> dict[str, str]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                "UPDATE users SET remember_token_hash = NULL, remember_token_expires = NULL "
                "WHERE remember_token_hash = %s",
                (hash_remember_token(payload.remember_token),),
            )
        connection.commit()
    finally:
        connection.close()
    return {"status": "success"}


@app.get("/groups")
def list_groups(actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable"))) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            if "Admin" in actor["role_tags"]:
                cursor.execute("SELECT id, name, is_active FROM organization_groups ORDER BY name")
            else:
                groups = actor["managed_group_ids"]
                if not groups:
                    return {"groups": []}
                placeholders = ", ".join(["%s"] * len(groups))
                cursor.execute(
                    f"SELECT id, name, is_active FROM organization_groups WHERE id IN ({placeholders}) AND is_active = 1 ORDER BY name",
                    tuple(groups),
                )
            return {"groups": cursor.fetchall()}
    finally:
        connection.close()


@app.post("/groups")
def create_group(payload: GroupRequest, _: dict[str, Any] = Depends(require_roles("Admin"))) -> dict[str, Any]:
    name = payload.name.strip()
    if not name:
        raise HTTPException(status_code=400, detail="Le nom du Groupe est obligatoire.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("INSERT INTO organization_groups (name, is_active) VALUES (%s, %s)", (name, payload.is_active))
            group_id = cursor.lastrowid
        connection.commit()
        return {"id": group_id, "name": name, "is_active": payload.is_active}
    except pymysql.err.IntegrityError:
        raise HTTPException(status_code=400, detail="Ce nom de Groupe existe déjà.") from None
    finally:
        connection.close()


@app.put("/groups/{group_id}")
def update_group(group_id: int, payload: GroupRequest, _: dict[str, Any] = Depends(require_roles("Admin"))) -> dict[str, Any]:
    name = payload.name.strip()
    if not name:
        raise HTTPException(status_code=400, detail="Le nom du Groupe est obligatoire.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("UPDATE organization_groups SET name = %s, is_active = %s WHERE id = %s", (name, payload.is_active, group_id))
            if cursor.rowcount == 0:
                raise HTTPException(status_code=404, detail="Groupe introuvable.")
        connection.commit()
        return {"id": group_id, "name": name, "is_active": payload.is_active}
    except pymysql.err.IntegrityError:
        raise HTTPException(status_code=400, detail="Ce nom de Groupe existe déjà.") from None
    finally:
        connection.close()


@app.post("/create-user")
def create_user(payload: dict[str, Any], actor: dict[str, Any] = Depends(require_roles("Admin"))) -> dict[str, Any]:
    username = str(payload.get("new_username", "")).strip()
    email = str(payload.get("new_mail", "")).strip()
    password = str(payload.get("new_password", ""))
    tags = requested_role_tags(payload)
    employee_type = str(payload.get("employee_type") or "salarie")
    if not username or not email or len(password) < 8:
        raise HTTPException(status_code=400, detail="Nom, e-mail et mot de passe de 8 caractères minimum sont obligatoires.")
    if employee_type not in VALID_EMPLOYEE_TYPES:
        raise HTTPException(status_code=400, detail="Type de personnel invalide.")
    if "Responsable" in actor["role_tags"] and "Admin" not in actor["role_tags"] and tags != ["Employe"]:
        raise HTTPException(status_code=403, detail="Un Responsable peut uniquement créer un Employé.")

    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            actor = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not can_assign_roles(actor, tags):
                raise HTTPException(status_code=403, detail="Vous ne pouvez pas attribuer ces droits.")
            if username_is_reserved(cursor, username):
                raise HTTPException(status_code=409, detail="Ce nom de compte est réservé et ne peut pas être réutilisé.")
            group_ids = validate_group_ids(cursor, list(payload.get("group_ids") or []))
            managed_group_ids = validate_group_ids(cursor, list(payload.get("managed_group_ids") or []))
            manager_username = payload.get("manager_id")
            if "Responsable" in actor["role_tags"] and "Admin" not in actor["role_tags"]:
                allowed = set(actor["managed_group_ids"])
                if not set(group_ids).issubset(allowed):
                    raise HTTPException(status_code=403, detail="Les Groupes choisis doivent être gérés par ce Responsable.")
                manager_username = actor["username"]
                managed_group_ids = []
            elif manager_username:
                manager = load_user(cursor, str(manager_username))
                if not is_valid_direct_manager(manager):
                    raise HTTPException(status_code=400, detail="Le Responsable direct ou l'Administrateur est invalide.")
            if "Employe" in tags and not manager_username:
                raise HTTPException(status_code=400, detail="Un Employé doit avoir un Responsable direct ou un Administrateur.")
            if "Responsable" in tags and "Admin" not in tags and not managed_group_ids:
                raise HTTPException(status_code=400, detail="Un Responsable doit gérer au moins un Groupe.")

            cursor.execute(
                """
                INSERT INTO users (username, email, password_hash, is_admin, role, employee_type, manager_username)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (username, email, hash_password(password), "Admin" in tags, tags[0], employee_type, manager_username),
            )
            cursor.execute("UPDATE users SET role_tags = %s WHERE username = %s", (json.dumps(tags), username))
            replace_memberships(cursor, username, "member", group_ids)
            replace_memberships(cursor, username, "manager", managed_group_ids if "Responsable" in tags else [])
            record_role_change(cursor, actor['username'], username, [], tags, action='created')
        connection.commit()
        return {"status": "success", "message": "Compte créé."}
    except pymysql.err.IntegrityError:
        connection.rollback()
        raise HTTPException(status_code=400, detail="Ce nom ou cet e-mail est déjà utilisé.") from None
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def admin_job_titles(cursor: pymysql.cursors.Cursor, actor: dict[str, Any],
                     users: list[dict[str, Any]]) -> dict[str, str]:
    """Read only explicitly linked job titles for the administration selector.

    This is intentionally separate from public_user and the public name directory:
    a manager's account scope does not grant access to the employee register.
    """
    if not is_admin(actor):
        return {}
    usernames = sorted({user['username'] for user in users if user})
    if not usernames:
        return {}
    cursor.execute("SHOW TABLES LIKE 'personnel_register'")
    if not cursor.fetchone():
        return {}
    titles: dict[str, str] = {}
    try:
        for offset in range(0, len(usernames), 500):
            batch = usernames[offset:offset + 500]
            placeholders = ','.join(['%s'] * len(batch))
            cursor.execute(f'''SELECT linked_username,job_title FROM personnel_register
                WHERE linked_username IN ({placeholders})''', tuple(batch))
            for row in cursor.fetchall():
                title = row.get('job_title')
                username = row.get('linked_username')
                if username in batch and isinstance(title, str) and title.strip():
                    titles[username] = title.strip()
    except pymysql.err.ProgrammingError as exc:
        if exc.args and exc.args[0] == 1146:
            return {}
        raise
    return titles


@app.get("/list-users")
def list_users(actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            candidates = all_users(cursor)
            if is_admin(actor):
                visible = candidates
            elif "Responsable" in actor["role_tags"]:
                visible = [user for user in candidates if user and user_can_manage(cursor, actor, user)]
            else:
                visible = [user for user in candidates if user and user["username"] == actor["username"]]
            users = [public_user(user) for user in visible if user]
            if is_admin(actor):
                titles = admin_job_titles(cursor, actor, visible)
                for user in users:
                    user['job_title'] = titles.get(user['username'])
            return {"status": "success", "users": users}
    finally:
        connection.close()


@app.get("/dashboard")
def get_dashboard(actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable"))) -> dict[str, Any]:
    """Return the operational to-do view for administrators and managers."""
    current_year = date.today().year
    today = date.today()
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            ensure_annual_interviews(cursor, current_year)
            auto_send = contract_reminders_auto_enabled(cursor)
            users = all_users(cursor)
            visible_users = [user for user in users_visible_to_actor(cursor, actor, users)
                             if user.get('is_active', True) and not user.get('login_deleted_at')]
            visible_usernames = {user["username"] for user in visible_users}
            candidates = contract_deadlines_from_fiches(cursor, today)
            if "Admin" in actor["role_tags"]:
                contracts = candidates
            else:
                contracts = [
                    contract
                    for contract in candidates
                    if contract.get("employee_username") == actor["username"]
                    or contract.get("manager_username") == actor["username"]
                    or str(contract.get("manager_email") or "").casefold() == str(actor["email"]).casefold()
                ]
            if visible_usernames:
                placeholders = ", ".join(["%s"] * len(visible_usernames))
                cursor.execute(
                    f"""
                    SELECT employee_username, interview_year, due_date, notes, completed, completed_at, completed_by
                    FROM annual_interviews
                    WHERE interview_year = %s AND employee_username IN ({placeholders}) AND is_archived=0
                    ORDER BY completed ASC, due_date ASC, employee_username ASC
                    """,
                    (current_year, *sorted(visible_usernames)),
                )
                interviews = cursor.fetchall()
            else:
                interviews = []
            by_username = {user['username']: user for user in users}
            for interview in interviews:
                employee = by_username.get(interview['employee_username'])
                interview['employee_name'] = person_display_name(employee)
                interview['manager_name'] = employee.get('manager_display_name') if employee else None
                completed_by = interview.get('completed_by')
                interview['completed_by_name'] = (
                    person_display_name(by_username[completed_by]) if completed_by in by_username
                    else 'Ancien compte' if completed_by else None
                )
        connection.commit()
        return {
            "year": current_year,
            "contract_reminders_auto_send": auto_send,
            "contracts": sorted(contracts, key=lambda item: (item["days_remaining"], item["employee_name"].casefold())),
            "interviews": interviews,
        }
    finally:
        connection.close()


@app.patch("/settings/contract-reminders")
def update_contract_reminder_settings(
    payload: ContractReminderSettingsRequest,
    _: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO app_settings (setting_key, setting_value)
                VALUES ('contract_reminders_auto_send', %s)
                ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)
                """,
                ("1" if payload.auto_send else "0",),
            )
        connection.commit()
        return {"status": "success", "auto_send": payload.auto_send}
    finally:
        connection.close()


@app.post("/contract-reminders/send")
def send_contract_reminders_now(
    _: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, int | bool]:
    """Explicitly dispatch the currently due contract reminders after UI confirmation."""
    return run_contract_end_reminders(force=True)


@app.post("/contract-reminders/resend")
def resend_contract_reminder(
    payload: ContractReminderResendRequest,
    _: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    """Explicitly resend one reminder after an administrator confirms the action."""
    reminder_key = payload.reminder_key.strip()
    if not reminder_key or len(reminder_key) > 100:
        raise HTTPException(status_code=400, detail="Référence de rappel invalide.")

    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            contract = next(
                (
                    item
                    for item in contract_deadlines_from_fiches(cursor, date.today())
                    if item["reminder_key"] == reminder_key and item["contract_end_date"] == payload.contract_end_date
                ),
                None,
            )
            if not contract:
                raise HTTPException(status_code=404, detail="Ce contrat n'est plus dans la période de rappel.")
            recipients = contract_reminder_recipients(cursor, contract)

        sent_count = sum(
            1 for recipient_email in recipients if send_contract_end_reminder_email(recipient_email, contract)
        )
        if not sent_count:
            raise HTTPException(status_code=502, detail="Aucun e-mail de rappel n'a pu être envoyé.")
        return {
            "status": "success",
            "sent": sent_count,
            "employee_name": contract["employee_name"],
        }
    finally:
        connection.close()


@app.patch("/annual-interviews/{employee_username}/{interview_year}")
def update_annual_interview(
    employee_username: str,
    interview_year: int,
    payload: AnnualInterviewUpdateRequest,
    actor: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    if interview_year < 2000 or interview_year > 2100:
        raise HTTPException(status_code=400, detail="Année d'entretien invalide.")
    if payload.notes is not None and len(payload.notes) > 5000:
        raise HTTPException(status_code=400, detail="La remarque est trop longue.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            require_target_access(cursor, actor, employee_username, edit=True)
            ensure_annual_interviews(cursor, interview_year)
            cursor.execute(
                "SELECT * FROM annual_interviews WHERE employee_username = %s AND interview_year = %s",
                (employee_username, interview_year),
            )
            interview = cursor.fetchone()
            if not interview:
                raise HTTPException(status_code=404, detail="Entretien annuel introuvable.")
            due_date = payload.due_date if payload.due_date is not None else interview["due_date"]
            notes = payload.notes if payload.notes is not None else interview.get("notes")
            completed = bool(interview["completed"]) if payload.completed is None else payload.completed
            completed_at = datetime.now() if completed else None
            completed_by = actor["username"] if completed else None
            cursor.execute(
                """
                UPDATE annual_interviews
                SET due_date = %s, notes = %s, completed = %s, completed_at = %s, completed_by = %s
                WHERE employee_username = %s AND interview_year = %s
                """,
                (due_date, notes, completed, completed_at, completed_by, employee_username, interview_year),
            )
        connection.commit()
        return {"status": "success"}
    finally:
        connection.close()


@app.put("/users/{target_username}/organization")
def update_user_organization(
    target_username: str,
    payload: dict[str, Any],
    actor: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    """Update organization metadata while enforcing the caller's management scope."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            # Serialize role edits and re-read the actor after acquiring the lock.
            # Two administrators must not concurrently demote each other.
            current_actor = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not is_admin(current_actor):
                raise HTTPException(status_code=403, detail="Votre rôle Admin a changé. Reconnectez-vous avant de continuer.")
            actor = current_actor
            target = require_target_access(cursor, actor, target_username, edit=True)
            tags = requested_role_tags(payload, target["role_tags"])
            if not can_assign_roles(actor, tags, target):
                raise HTTPException(status_code=403, detail="Seul un Super-administrateur peut gérer ces droits.")
            if is_super_admin(target) and 'SuperAdmin' not in tags:
                if active_super_admin_count(cursor, exclude_username=target['username']) == 0:
                    raise HTTPException(status_code=409, detail="Le dernier Super-administrateur actif doit être conservé.")
                if target['username'] == actor['username']:
                    raise HTTPException(status_code=400, detail="Vous ne pouvez pas retirer vos propres droits de Super-administrateur.")
            if target_username == actor["username"] and "Admin" not in tags:
                raise HTTPException(status_code=400, detail="Vous ne pouvez pas retirer votre propre rôle Admin. Demandez à un autre administrateur.")
            employee_type = str(payload.get("employee_type", target["employee_type"]))
            if employee_type not in VALID_EMPLOYEE_TYPES:
                raise HTTPException(status_code=400, detail="Type de personnel invalide.")
            if "Responsable" in actor["role_tags"] and "Admin" not in actor["role_tags"]:
                if "Employe" not in target["role_tags"] or tags != ["Employe"]:
                    raise HTTPException(status_code=403, detail="Un Responsable peut uniquement modifier un Employé de son périmètre.")
                group_ids = validate_group_ids(cursor, list(payload.get("group_ids", target["group_ids"])))
                if not set(group_ids).issubset(set(actor["managed_group_ids"])):
                    raise HTTPException(status_code=403, detail="Les Groupes choisis doivent être gérés par ce Responsable.")
                manager_username = actor["username"]
                managed_group_ids: list[int] = []
            else:
                group_ids = validate_group_ids(cursor, list(payload.get("group_ids", target["group_ids"])))
                managed_group_ids = validate_group_ids(
                    cursor,
                    list(payload.get("managed_group_ids", target["managed_group_ids"])),
                )
                manager_username = payload.get("manager_id", target.get("manager_username"))
                if manager_username:
                    manager = load_user(cursor, str(manager_username))
                    if not is_valid_direct_manager(manager):
                        raise HTTPException(status_code=400, detail="Le Responsable direct ou l'Administrateur est invalide.")
                    if target_username == manager_username or would_create_management_cycle(cursor, target_username, manager_username):
                        raise HTTPException(status_code=400, detail="Cette affectation créerait une boucle hiérarchique.")
                if "Employe" in tags and not manager_username:
                    raise HTTPException(status_code=400, detail="Un Employé doit avoir un Responsable direct ou un Administrateur.")
                if "Responsable" in tags and "Admin" not in tags and not managed_group_ids:
                    raise HTTPException(status_code=400, detail="Un Responsable doit gérer au moins un Groupe.")
                if tags == ["Admin"]:
                    managed_group_ids = []

            cursor.execute(
                """
                UPDATE users SET role = %s, role_tags = %s, is_admin = %s, employee_type = %s, manager_username = %s
                WHERE username = %s
                """,
                (
                    tags[0],
                    json.dumps(tags),
                    "Admin" in tags,
                    employee_type,
                    manager_username,
                    target_username,
                ),
            )
            replace_memberships(cursor, target_username, "member", group_ids)
            replace_memberships(cursor, target_username, "manager", managed_group_ids if "Responsable" in tags else [])
            if set(tags) != set(target['role_tags']):
                cursor.execute('''UPDATE users SET session_version=session_version+1,
                    remember_token_hash=NULL,remember_token_expires=NULL,reset_token=NULL,
                    reset_token_expires=CURRENT_TIMESTAMP WHERE username=%s''', (target['username'],))
                record_role_change(cursor, actor['username'], target['username'], target['role_tags'], tags)
            updated = load_user(cursor, target_username)
        connection.commit()
        if set(tags) != set(target["role_tags"]):
            import logging
            logging.getLogger("uvicorn.error").info("Organization role change %s", json.dumps({
                "actor": actor["username"], "target": target_username,
                "before": target["role_tags"], "after": tags,
            }, ensure_ascii=True))
        return {"status": "success", "user": public_user(updated)}
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


@app.patch("/users/{target_username}/contract")
def update_user_contract(
    target_username: str,
    payload: ContractUpdateRequest,
    actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
) -> dict[str, Any]:
    raise HTTPException(status_code=409, detail="Utilisez la gestion de l'historique des contrats dans Administration.")


@app.patch("/users/{target_username}/direct-manager")
def assign_direct_manager(
    target_username: str,
    payload: DirectManagerRequest,
    actor: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    """Assign any existing person to a direct manager without changing other metadata."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            actor = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not is_admin(actor):
                raise HTTPException(status_code=403, detail="Autorisation insuffisante.")
            target = load_user(cursor, target_username)
            manager = load_user(cursor, payload.manager_id)
            if not target:
                raise HTTPException(status_code=404, detail="Utilisateur introuvable.")
            if not can_manage_account(actor, target):
                raise HTTPException(status_code=403, detail="Seul un Super-administrateur peut modifier ce compte.")
            if not is_valid_direct_manager(manager):
                raise HTTPException(status_code=400, detail="Le responsable sélectionné est invalide.")
            if target_username == payload.manager_id or would_create_management_cycle(cursor, target_username, payload.manager_id):
                raise HTTPException(status_code=400, detail="Cette affectation créerait une boucle hiérarchique.")
            cursor.execute(
                "UPDATE users SET manager_username = %s WHERE username = %s",
                (payload.manager_id, target_username),
            )
            updated = load_user(cursor, target_username)
        connection.commit()
        return {"status": "success", "user": public_user(updated)}
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def fiche_record_belongs_to_user(record: dict[str, Any], username: str, email: str) -> bool:
    account_username = str(
        record.get("account_username") or record.get("employee_username") or record.get("user_username") or ""
    ).strip()
    record_email = str(record.get("email_employe") or "").strip()
    return account_username == username or (bool(record_email) and record_email.casefold() == email.casefold())


def serialize_fiche_config(data: dict[str, Any]) -> str:
    """Validate the flexible legacy fiche payload while putting bounded limits around it."""
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="La configuration de fiche est invalide.")
    employees = data.get("employes_data", [])
    if employees is not None:
        if not isinstance(employees, list) or len(employees) > MAX_FICHE_EMPLOYEES:
            raise HTTPException(
                status_code=400,
                detail=f"Une fiche peut contenir au plus {MAX_FICHE_EMPLOYEES} employés.",
            )
        if any(not isinstance(employee, dict) for employee in employees):
            raise HTTPException(status_code=400, detail="Un employé de la fiche est invalide.")
    try:
        serialized = json.dumps(data, default=str)
    except (TypeError, ValueError) as error:
        raise HTTPException(status_code=400, detail="La configuration de fiche est invalide.") from error
    if len(serialized.encode("utf-8")) > MAX_FICHE_CONFIG_BYTES:
        raise HTTPException(status_code=413, detail="La configuration de fiche dépasse la limite de 5 Mo.")
    return serialized


def migrate_username_in_saved_fiches(
    cursor: pymysql.cursors.Cursor,
    old_username: str,
    new_username: str,
    old_email: str,
    new_email: str,
) -> None:
    """Update stable account links and manager details in every stored fiche configuration."""
    cursor.execute("SELECT user_id, form_content FROM Presence")
    username_fields = {
        "account_username",
        "employee_username",
        "user_username",
        "manager_username",
        "responsable_id",
    }
    for row in cursor.fetchall():
        content = row.get("form_content")
        try:
            config = content if isinstance(content, dict) else json.loads(content)
        except (TypeError, json.JSONDecodeError):
            # Leave historical corrupt content untouched; get_config reports it safely to the owner.
            continue
        employees = config.get("employes_data") if isinstance(config, dict) else None
        if not isinstance(employees, list):
            continue
        changed = False
        for employee in employees:
            if not isinstance(employee, dict):
                continue
            for field in username_fields:
                if str(employee.get(field) or "") == old_username:
                    employee[field] = new_username
                    changed = True
            # These two fields are displayed to users, but equal usernames are unambiguous links.
            if str(employee.get("responsable") or "") == old_username:
                employee["responsable"] = new_username
                changed = True
            if old_email != new_email:
                for field in ("email_employe", "email_responsable"):
                    if str(employee.get(field) or "").casefold() == old_email.casefold():
                        employee[field] = new_email
                        changed = True
        if changed:
            cursor.execute(
                "UPDATE Presence SET form_content = %s WHERE user_id = %s",
                (serialize_fiche_config(config), row["user_id"]),
            )


@app.delete("/delete-user/{username_to_delete}")
def delete_user(username_to_delete: str, payload: AccountPurgeRequest,
                actor: dict[str, Any] = Depends(require_roles("SuperAdmin"))) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            current = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not is_super_admin(current):
                raise HTTPException(403, 'La suppression définitive est réservée aux Super-administrateurs.')
            target = load_user(cursor, username_to_delete)
            if not target:
                raise HTTPException(404, 'Compte introuvable.')
            confirmations = {target['username']}
            target_email = str(target.get('email') or '').strip()
            if target_email:
                confirmations.add(target_email)
            if payload.confirmation not in confirmations:
                raise HTTPException(400, 'Saisissez l’adresse e-mail exacte du compte pour confirmer.')
            if target['username'] == current['username']:
                raise HTTPException(400, 'Vous ne pouvez pas supprimer votre propre compte.')
            if target.get('is_active', True):
                raise HTTPException(409, 'Désactivez le compte avant sa suppression définitive.')
            if is_super_admin(target) and active_super_admin_count(cursor, exclude_username=target['username']) == 0:
                raise HTTPException(409, 'Le dernier Super-administrateur actif doit être conservé.')
            statistics = perform_account_purge(cursor, current, target, reason=payload.reason)
        connection.commit()
        return {'status': 'success', 'statistics': statistics}
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def perform_account_purge(cursor, actor, target, *, reason):
    if __package__:
        from .account_purge import purge_account
    else:
        from account_purge import purge_account
    return purge_account(cursor, actor, target, reason=reason)


@app.put("/update_profile/{username}")
def update_profile(username: str, payload: ProfileUpdateRequest, actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
    if username != actor["username"]:
        raise HTTPException(status_code=403, detail="Vous pouvez uniquement modifier votre propre profil.")
    if bool(payload.new_password) != bool(payload.confirm_password) or payload.new_password != payload.confirm_password:
        raise HTTPException(status_code=400, detail="Les nouveaux mots de passe ne correspondent pas.")
    if payload.new_password and len(payload.new_password) < 8:
        raise HTTPException(status_code=400, detail="Le mot de passe doit contenir au moins 8 caractères.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            # Use the same users -> employee lock order as registry/account writes.
            cursor.execute("SELECT username FROM users ORDER BY username FOR UPDATE")
            cursor.fetchall()
            user = load_user(cursor, username)
            if (not user or not user.get("is_active", True) or user.get("login_deleted_at")
                or user.get("session_version", 0) != actor.get("session_version", 0)):
                raise HTTPException(status_code=401, detail="Votre session n’est plus valide. Reconnectez-vous.")
            if not user or not verify_password(payload.current_password, user["password_hash"]):
                raise HTTPException(status_code=403, detail="Mot de passe actuel incorrect.")
            updates: list[str] = []
            params: list[Any] = []
            target_username = username
            target_email = str(user["email"])
            if payload.new_username and payload.new_username != username:
                existing_name = load_user(cursor, payload.new_username)
                same_identity = bool(existing_name and existing_name['username'] == user['username'])
                if not same_identity and username_is_reserved(cursor, payload.new_username):
                    raise HTTPException(409, 'Ce nom de compte est réservé et ne peut pas être réutilisé.')
                updates.append("username = %s")
                params.append(payload.new_username)
                target_username = payload.new_username
            if payload.new_email and payload.new_email != user["email"]:
                updates.append("email = %s")
                target_email = str(payload.new_email)
                params.append(target_email)
            if payload.new_password:
                updates.append("password_hash = %s")
                params.append(hash_password(payload.new_password))
            if updates:
                if target_username != username:
                    reserve_deleted_account_name(cursor, user)
                cursor.execute(f"UPDATE users SET {', '.join(updates)} WHERE username = %s", (*params, username))
                if payload.new_password:
                    cursor.execute(
                        "UPDATE users SET remember_token_hash = NULL, remember_token_expires = NULL WHERE username = %s",
                        (target_username,),
                    )
                if target_username != username:
                    cursor.execute("UPDATE Presence SET user_id = %s WHERE user_id = %s", (target_username, username))
                    cursor.execute("UPDATE user_group_memberships SET username = %s WHERE username = %s", (target_username, username))
                    cursor.execute("UPDATE account_status_events SET target_username = %s WHERE target_username = %s", (target_username, username))
                    cursor.execute("UPDATE users SET manager_username = %s WHERE manager_username = %s", (target_username, username))
                    cursor.execute(
                        "UPDATE transport_receipts SET employee_username = %s WHERE employee_username = %s",
                        (target_username, username),
                    )
                    cursor.execute(
                        "UPDATE transport_receipts SET uploaded_by = %s WHERE uploaded_by = %s",
                        (target_username, username),
                    )
                    cursor.execute(
                        "UPDATE transport_receipts SET archived_by = %s WHERE archived_by = %s",
                        (target_username, username),
                    )
                    cursor.execute(
                        "UPDATE annual_interviews SET employee_username = %s WHERE employee_username = %s",
                        (target_username, username),
                    )
                    cursor.execute(
                        "UPDATE annual_interviews SET completed_by = %s WHERE completed_by = %s",
                        (target_username, username),
                    )
                    cursor.execute(
                        "UPDATE contract_end_reminders SET employee_username = %s WHERE employee_username = %s",
                        (target_username, username),
                    )
                if target_username != username:
                    cursor.execute("UPDATE personnel_contracts SET username=%s WHERE username=%s", (target_username, username))
                    rename_leave_owner(cursor, username, target_username)
                    rename_personnel_account_link(cursor, username, target_username, target_username)
                    rename_signed_fiche_owner(cursor, username, target_username)
                if target_username != username or target_email != str(user["email"]):
                    migrate_username_in_saved_fiches(
                        cursor,
                        username,
                        target_username,
                        str(user["email"]),
                        target_email,
                    )
        connection.commit()
        return {
            "status": "success",
            "message": "Profil mis à jour.",
            "username": target_username,
            "auth_token": create_session_token(target_username, user.get("session_version", 0)),
        }
    except pymysql.err.IntegrityError:
        raise HTTPException(status_code=400, detail="Ce nom ou cet e-mail est déjà utilisé.") from None
    finally:
        connection.close()


@app.post("/forgot-password")
def forgot_password(payload: dict[str, Any]) -> dict[str, str]:
    email = str(payload.get("email", ""))
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT username FROM users WHERE email = %s AND is_active=1 AND login_deleted_at IS NULL", (email,))
            user = cursor.fetchone()
            if not user:
                return {"status": "success", "message": "Si l'adresse est valide, un e-mail a été envoyé."}
            token = secrets.token_urlsafe(32)
            cursor.execute(
                "UPDATE users SET reset_token = %s, reset_token_expires = %s WHERE email = %s",
                (token, datetime.now() + timedelta(minutes=PASSWORD_RESET_DURATION_MINUTES), email),
            )
        connection.commit()
        send_password_reset_email(email, token)
        return {"status": "success", "message": "Si l'adresse est valide, un e-mail a été envoyé."}
    finally:
        connection.close()


@app.post("/reset-password")
def reset_password(payload: dict[str, Any]) -> dict[str, str]:
    token = str(payload.get("token", ""))
    new_password = str(payload.get("new_password", ""))
    if len(new_password) < 8:
        raise HTTPException(status_code=400, detail="Le mot de passe doit contenir au moins 8 caractères.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT username FROM users WHERE reset_token = %s AND reset_token_expires > %s AND is_active=1 AND login_deleted_at IS NULL",
                (token, datetime.now()),
            )
            user = cursor.fetchone()
            if not user:
                raise HTTPException(status_code=400, detail="Token invalide ou expiré.")
            cursor.execute(
                "UPDATE users SET password_hash = %s, reset_token = NULL, reset_token_expires = %s, "
                "remember_token_hash = NULL, remember_token_expires = NULL WHERE username = %s",
                (hash_password(new_password), datetime.now() - timedelta(seconds=1), user["username"]),
            )
        connection.commit()
        return {"status": "success", "message": "Mot de passe réinitialisé."}
    finally:
        connection.close()


@app.get("/get-config/{user_id}")
def get_config(user_id: str, actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            require_target_access(cursor, actor, user_id, edit=False)
            cursor.execute("SELECT form_content FROM Presence WHERE user_id = %s", (user_id,))
            result = cursor.fetchone()
            if not result:
                return {}
            content = result["form_content"]
            try:
                return content if isinstance(content, dict) else json.loads(content)
            except (TypeError, json.JSONDecodeError):
                raise HTTPException(
                    status_code=422,
                    detail="La configuration enregistrée est invalide. Contactez un administrateur pour la restaurer.",
                ) from None
    finally:
        connection.close()


@app.post("/save-config/{user_id}")
def save_config(user_id: str, data: dict[str, Any], actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, str]:
    save_config_checked(user_id, data, actor)
    return {"status": "success"}


def save_config_checked(user_id, data, actor, expected_revision=None, *, confirmed_deletions=None):
    if __package__:
        from .attendance_revision import config_revision, require_revision
    else:
        from attendance_revision import config_revision, require_revision
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            actor = checked_resource_actor(cursor, actor, load_user, lock=True)
            require_target_access(cursor, actor, user_id, edit=True)
            cursor.execute("SELECT form_content FROM Presence WHERE user_id=%s FOR UPDATE", (user_id,))
            old_row = cursor.fetchone()
            old_config = old_row["form_content"] if old_row else {}
            if isinstance(old_config, str):
                old_config = json.loads(old_config)
            require_revision(old_config, expected_revision)
            validate_manager_save(cursor, actor, old_config, data, load_user, user_can_manage,
                                  workspace_owner=user_id, confirmed_deletions=confirmed_deletions)
            serialize_fiche_config(data)
            link_fiche_records_to_users(cursor, data)
            sync_contract_dates_from_config(cursor, data)
            serialized = serialize_fiche_config(data)
            cursor.execute(
                """
                INSERT INTO Presence (user_id, form_content) VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE form_content = %s
                """,
                (user_id, serialized, serialized),
            )
        connection.commit()
        canonical = json.loads(serialized)
        return {"status": "success", "config": canonical, "revision": config_revision(canonical)}
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


# Support both package imports locally and `uvicorn main:app` on the server.
if __package__:
    from .account_lifecycle import register_account_lifecycle
    from .ui_capabilities import register_ui_capabilities
    from .leave_requests import register_leave_routes
    from .leave_store import rename_leave_owner
    from .personnel_store import rename_personnel_account_link
    from .team_directory import register_team_directory
    from .contract_history import register_contract_history
    from .attendance_permissions import validate_manager_save
    from .documents import register_document_routes
    from .esign import register_esign_routes, convert_fiche_to_pdf, submit_to_clawshow
else:
    from account_lifecycle import register_account_lifecycle
    from ui_capabilities import register_ui_capabilities
    from leave_requests import register_leave_routes
    from leave_store import rename_leave_owner
    from personnel_store import rename_personnel_account_link
    from contract_history import register_contract_history
    from team_directory import register_team_directory
    from attendance_permissions import validate_manager_save
    from documents import register_document_routes
    from esign import register_esign_routes, convert_fiche_to_pdf, submit_to_clawshow

register_account_lifecycle(app, get_db_connection=get_db_connection, get_current_user=get_current_user, load_user=load_user, public_user=public_user, hash_password=hash_password)
register_ui_capabilities(app, mail_configured=lambda: bool(SMTP_SERVER and SMTP_USER and SMTP_PASSWORD))
register_contract_history(app, get_db_connection, require_roles, get_current_user, require_target_access, load_user=load_user)
register_team_directory(app, get_db_connection=get_db_connection, get_current_user=get_current_user)
register_leave_routes(app, get_db_connection=get_db_connection, get_current_user=get_current_user)
register_document_routes(
    app, get_db_connection=get_db_connection, get_current_user=get_current_user,
    require_roles=require_roles, load_user=load_user, user_can_manage=user_can_manage,
    all_users=all_users, storage_dir=TRANSPORT_RECEIPTS_DIR, max_file_bytes=MAX_TRANSPORT_RECEIPT_BYTES,
)
if __package__:
    from .signed_fiches import register_signed_fiche_routes
    from .signed_fiche_store import rename_signed_fiche_owner
else:
    from signed_fiches import register_signed_fiche_routes
    from signed_fiche_store import rename_signed_fiche_owner
register_signed_fiche_routes(
    app, get_db_connection=get_db_connection, get_current_user=get_current_user, load_user=load_user,
    root=os.getenv('PRESENCE_SIGNED_FICHES_DIR', '/var/lib/presence-app/signed-fiches'),
)
send_fiche_for_signature = register_esign_routes(
    app, get_db_connection=get_db_connection, require_roles=require_roles,
    load_user=load_user, require_target_access=require_target_access,
    storage_dir=ESIGN_DOCUMENTS_DIR, document_duration_hours=ESIGN_DOCUMENT_DURATION_HOURS,
    max_file_bytes=MAX_FICHE_ATTACHMENT_BYTES,
)
