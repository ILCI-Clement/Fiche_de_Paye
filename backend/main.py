"""Presence API with organization-aware authorization.

This module is deployed as ``/var/www/presence-app/main.py``. It owns the
MariaDB schema migration for groups and enforces access server-side.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import json
import os
import secrets
import smtplib
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

import bcrypt
import pymysql
from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, EmailStr


app = FastAPI(title="Presence API")
VALID_ROLES = {"Admin", "Responsable", "Employe"}
ROLE_PRIORITY = ("Admin", "Responsable", "Employe")
VALID_EMPLOYEE_TYPES = {"salarie", "stagiaire"}
SESSION_DURATION_SECONDS = 8 * 60 * 60
REMEMBER_SESSION_DURATION_DAYS = 30
PASSWORD_RESET_DURATION_MINUTES = 15
MAX_FICHE_ATTACHMENT_BYTES = 10 * 1024 * 1024
MAX_TRANSPORT_RECEIPT_BYTES = 10 * 1024 * 1024
ESIGN_DOCUMENT_DURATION_HOURS = 24
CLAWSHOW_CREATE_URL = "https://esign.clawshow.ai/esign/create"
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


class RememberSessionRequest(BaseModel):
    remember_token: str


class ESignFicheRequest(BaseModel):
    recipient_email: EmailStr
    employee_name: str
    month: int
    year: int
    filename: str
    file_b64: str


class TransportReceiptUploadRequest(BaseModel):
    employee_username: str
    original_filename: str
    file_b64: str


class AnnualInterviewUpdateRequest(BaseModel):
    due_date: date | None = None
    notes: str | None = None
    completed: bool | None = None


class ContractUpdateRequest(BaseModel):
    contract_start_date: date | None = None
    contract_end_date: date | None = None
    is_cdi: bool = False


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


def create_session_token(username: str) -> str:
    payload = json.dumps(
        {"username": username, "expires_at": int(time.time()) + SESSION_DURATION_SECONDS},
        separators=(",", ":"),
    ).encode("utf-8")
    payload_part = encode_segment(payload)
    signature = hmac.new(API_TOKEN.encode("utf-8"), payload_part.encode("ascii"), hashlib.sha256).digest()
    return f"{payload_part}.{encode_segment(signature)}"


def read_session_token(token: str) -> str:
    try:
        payload_part, signature_part = token.split(".", 1)
        expected = hmac.new(API_TOKEN.encode("utf-8"), payload_part.encode("ascii"), hashlib.sha256).digest()
        if not hmac.compare_digest(expected, decode_segment(signature_part)):
            raise ValueError("Invalid signature")
        payload = json.loads(decode_segment(payload_part))
        if int(payload["expires_at"]) < int(time.time()):
            raise ValueError("Expired token")
        return str(payload["username"])
    except (KeyError, TypeError, ValueError, json.JSONDecodeError):
        raise HTTPException(status_code=401, detail="Session invalide ou expirée.") from None


def ensure_organization_schema() -> None:
    """Apply additive, idempotent schema changes without touching existing data."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SHOW COLUMNS FROM users")
            columns = {item["Field"] for item in cursor.fetchall()}
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


def load_user(cursor: pymysql.cursors.Cursor, username: str) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT username, email, password_hash, is_admin, role, role_tags, employee_type, manager_username,
               contract_start_date, contract_end_date, is_cdi, created_at
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
    user["role_tags"] = [role for role in ROLE_PRIORITY if role in tags]
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
    return user


def public_user(user: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": user["username"],
        "username": user["username"],
        "email": user["email"],
        "role": user["role"],
        "role_tags": user["role_tags"],
        "is_admin": bool(user.get("is_admin")),
        "employee_type": user.get("employee_type") or "salarie",
        "manager_id": user.get("manager_username"),
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


def sync_contract_dates_from_config(cursor: pymysql.cursors.Cursor, config: dict[str, Any]) -> None:
    """Copy fiche contract dates only to the resolved employee account, never to an unrelated e-mail owner."""
    for employee in config.get("employes_data", []) if isinstance(config, dict) else []:
        if not isinstance(employee, dict):
            continue
        account_username = resolve_fiche_account_username(cursor, employee)
        start_date, end_date, is_cdi = fiche_contract_values(employee)
        if not account_username or (not start_date and not end_date and not is_cdi):
            continue
        cursor.execute(
            """
            UPDATE users
            SET contract_start_date = %s, contract_end_date = %s, is_cdi = %s
            WHERE username = %s
            """,
            (start_date, end_date, is_cdi, account_username),
        )


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


def fiche_display_names_by_account(cursor: pymysql.cursors.Cursor) -> dict[str, str]:
    """Use the name shown on the fiche when presenting an account's contract deadline."""
    names: dict[str, str] = {}
    cursor.execute("SELECT form_content FROM Presence")
    for row in cursor.fetchall():
        content = row.get("form_content")
        try:
            config = content if isinstance(content, dict) else json.loads(content)
        except (TypeError, json.JSONDecodeError):
            continue
        for employee in config.get("employes_data", []) if isinstance(config, dict) else []:
            if not isinstance(employee, dict):
                continue
            account_username = resolve_fiche_account_username(cursor, employee)
            display_name = fiche_display_name(employee)
            if account_username and display_name and account_username not in names:
                names[account_username] = display_name
    return names


def eligible_for_annual_interview(user: dict[str, Any]) -> bool:
    # The Admin label is an application permission, not an employment status:
    # administrators also need their yearly interview to be tracked.
    return bool({"Admin", "Employe", "Responsable"} & set(user["role_tags"]))


def ensure_annual_interviews(cursor: pymysql.cursors.Cursor, year: int) -> None:
    """Create one annual-interview task per active personnel account and year."""
    cursor.execute("SELECT username FROM users ORDER BY username")
    for row in cursor.fetchall():
        user = load_user(cursor, str(row["username"]))
        if user and eligible_for_annual_interview(user):
            cursor.execute(
                """
                INSERT IGNORE INTO annual_interviews (employee_username, interview_year, due_date)
                VALUES (%s, %s, %s)
                """,
                (user["username"], year, date(year, 12, 31)),
            )


def all_users(cursor: pymysql.cursors.Cursor) -> list[dict[str, Any]]:
    cursor.execute("SELECT username FROM users ORDER BY username")
    return [user for row in cursor.fetchall() if (user := load_user(cursor, str(row["username"])))]


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
    username = read_session_token(authorization.removeprefix("Bearer "))
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            user = load_user(cursor, username)
        if not user:
            raise HTTPException(status_code=401, detail="Session invalide.")
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
    return bool(user and ({"Admin", "Responsable"} & set(user["role_tags"])))


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
    tags = {str(tag) for tag in (raw_tags or [])}
    if not tags or not tags.issubset(VALID_ROLES):
        raise HTTPException(status_code=400, detail="Les étiquettes d'identité sont invalides.")
    return [role for role in ROLE_PRIORITY if role in tags]


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
            <p>Une demande de réinitialisation a été effectuée pour votre compte Presence.</p>
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


def send_contract_end_reminder_email(recipient_email: str, employee: dict[str, Any], days_remaining: int) -> bool:
    """Send a single, actionable contract-end reminder to an admin or direct manager."""
    end_date = employee.get("contract_end_date")
    formatted_end_date = end_date.strftime("%d/%m/%Y") if isinstance(end_date, date) else str(end_date)
    message = MIMEMultipart()
    message["From"] = SMTP_USER
    message["To"] = recipient_email
    message["Subject"] = f"Rappel : fin de contrat de {employee['username']}"
    message.attach(
        MIMEText(
            f"""<h3>Fin de contrat à anticiper</h3>
            <p>Bonjour,</p>
            <p>Le contrat de <strong>{employee['username']}</strong> se termine le
            <strong>{formatted_end_date}</strong> ({days_remaining} jour(s) restant(s)).</p>
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


def run_contract_end_reminders() -> None:
    """Notify admins and direct managers once per recipient for contracts ending within ten days."""
    today = date.today()
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            users = all_users(cursor)
            admins = [user for user in users if "Admin" in user["role_tags"]]
            for employee in users:
                end_date = employee.get("contract_end_date")
                if not isinstance(end_date, date) or bool(employee.get("is_cdi")):
                    continue
                days_remaining = (end_date - today).days
                if not 0 <= days_remaining <= 10:
                    continue
                recipients = {admin["email"] for admin in admins}
                manager_username = employee.get("manager_username")
                manager = next((user for user in users if user["username"] == manager_username), None)
                if manager:
                    recipients.add(manager["email"])
                for recipient_email in recipients:
                    cursor.execute(
                        """
                        SELECT 1 FROM contract_end_reminders
                        WHERE employee_username = %s AND contract_end_date = %s AND recipient_email = %s
                        """,
                        (employee["username"], end_date, recipient_email),
                    )
                    if cursor.fetchone():
                        continue
                    if send_contract_end_reminder_email(recipient_email, employee, days_remaining):
                        cursor.execute(
                            """
                            INSERT INTO contract_end_reminders (employee_username, contract_end_date, recipient_email)
                            VALUES (%s, %s, %s)
                            """,
                            (employee["username"], end_date, recipient_email),
                        )
        connection.commit()
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


def e_sign_public_base_url() -> str:
    """Return the public HTTPS origin used by ClawShow to retrieve a PDF."""
    public_url = os.getenv("PRESENCE_ESIGN_PUBLIC_URL", "").rstrip("/")
    if not public_url.startswith("https://"):
        raise HTTPException(
            status_code=503,
            detail="L'URL publique de signature électronique n'est pas configurée sur le serveur.",
        )
    return public_url


def clean_expired_esign_documents() -> None:
    """Remove expired private PDFs and their metadata without touching other paths."""
    if not ESIGN_DOCUMENTS_DIR.is_dir():
        return
    now = time.time()
    for metadata_path in ESIGN_DOCUMENTS_DIR.glob("*.json"):
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            if float(metadata.get("expires_at", 0)) > now:
                continue
            pdf_name = str(metadata.get("pdf_name", ""))
            if pdf_name.endswith(".pdf") and "/" not in pdf_name and "\\" not in pdf_name:
                (ESIGN_DOCUMENTS_DIR / pdf_name).unlink(missing_ok=True)
            metadata_path.unlink(missing_ok=True)
        except (OSError, TypeError, ValueError, json.JSONDecodeError):
            # A malformed metadata file must not stop a signature request.
            continue


def convert_fiche_to_pdf(file_bytes: bytes, filename: str) -> bytes:
    """Convert a generated Excel or Word attendance sheet with LibreOffice."""
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xlsx", ".docx"}:
        raise HTTPException(status_code=400, detail="Seuls les fichiers Excel et Word peuvent être signés.")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise HTTPException(
            status_code=503,
            detail="La génération PDF pour la signature n'est pas encore disponible sur le serveur.",
        )

    with tempfile.TemporaryDirectory(prefix="presence-esign-") as temporary_dir:
        source_path = Path(temporary_dir) / f"fiche{suffix}"
        source_path.write_bytes(file_bytes)
        profile_path = Path(temporary_dir) / "libreoffice-profile"
        profile_path.mkdir()
        try:
            result = subprocess.run(
                [
                    soffice,
                    f"-env:UserInstallation={profile_path.as_uri()}",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temporary_dir,
                    str(source_path),
                ],
                check=False,
                capture_output=True,
                text=True,
                timeout=90,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            raise HTTPException(status_code=502, detail="La conversion de la fiche en PDF a échoué.") from error
        pdf_path = source_path.with_suffix(".pdf")
        if result.returncode != 0 or not pdf_path.is_file():
            raise HTTPException(status_code=502, detail="La conversion de la fiche en PDF a échoué.")
        pdf_bytes = pdf_path.read_bytes()
        if not pdf_bytes:
            raise HTTPException(status_code=502, detail="Le PDF généré est vide.")
        return pdf_bytes


def store_esign_pdf(pdf_bytes: bytes, original_filename: str) -> tuple[str, str]:
    """Store a short-lived PDF behind an unguessable public token."""
    ESIGN_DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)
    document_token = secrets.token_urlsafe(32)
    pdf_name = f"{document_token}.pdf"
    expires_at = time.time() + ESIGN_DOCUMENT_DURATION_HOURS * 3600
    (ESIGN_DOCUMENTS_DIR / pdf_name).write_bytes(pdf_bytes)
    (ESIGN_DOCUMENTS_DIR / f"{document_token}.json").write_text(
        json.dumps({"pdf_name": pdf_name, "expires_at": expires_at, "filename": original_filename}),
        encoding="utf-8",
    )
    return document_token, pdf_name


def delete_esign_pdf(document_token: str, pdf_name: str) -> None:
    (ESIGN_DOCUMENTS_DIR / pdf_name).unlink(missing_ok=True)
    (ESIGN_DOCUMENTS_DIR / f"{document_token}.json").unlink(missing_ok=True)


def sanitize_transport_receipt_filename(filename: str) -> str:
    """Keep only a safe display name for an uploaded transport receipt."""
    safe_name = str(filename or "").replace("\\", "/").rsplit("/", 1)[-1].strip()
    if not safe_name or len(safe_name) > 255:
        raise HTTPException(status_code=400, detail="Le nom du fichier est invalide.")
    return safe_name


def validate_transport_receipt_file(file_bytes: bytes, filename: str) -> tuple[str, str]:
    """Allow only a small, verifiable set of receipt formats."""
    extension = Path(filename).suffix.lower()
    valid_files = {
        ".pdf": ("application/pdf", b"%PDF-"),
        ".png": ("image/png", b"\x89PNG\r\n\x1a\n"),
        ".jpg": ("image/jpeg", b"\xff\xd8\xff"),
        ".jpeg": ("image/jpeg", b"\xff\xd8\xff"),
    }
    expected = valid_files.get(extension)
    if not expected or not file_bytes.startswith(expected[1]):
        raise HTTPException(status_code=400, detail="Seuls les fichiers PDF, JPG et PNG valides sont acceptés.")
    return extension, expected[0]


def transport_receipt_response(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": int(row["id"]),
        "employee_username": str(row["employee_username"]),
        "uploaded_by": str(row["uploaded_by"]),
        "original_filename": str(row["original_filename"]),
        "mime_type": str(row["mime_type"]),
        "size_bytes": int(row["size_bytes"]),
        "archived": bool(row["archived"]),
        "created_at": row["created_at"],
        "archived_at": row.get("archived_at"),
        "archived_by": row.get("archived_by"),
    }


def load_transport_receipt(cursor: pymysql.cursors.Cursor, receipt_id: int) -> dict[str, Any]:
    cursor.execute("SELECT * FROM transport_receipts WHERE id = %s", (receipt_id,))
    receipt = cursor.fetchone()
    if not receipt:
        raise HTTPException(status_code=404, detail="Justificatif introuvable.")
    return receipt


def can_view_transport_receipt(actor: dict[str, Any], receipt: dict[str, Any]) -> bool:
    return "Admin" in actor["role_tags"] or receipt["employee_username"] == actor["username"]


@app.get("/transport-receipts/assignees")
def list_transport_receipt_assignees(
    _: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    """Return every registered person so an uploader can assign a receipt."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT username FROM users ORDER BY username")
            return {"users": [str(row["username"]) for row in cursor.fetchall()]}
    finally:
        connection.close()


@app.post("/transport-receipts")
def upload_transport_receipt(
    payload: TransportReceiptUploadRequest,
    actor: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    try:
        employee_username = str(payload.employee_username).strip()
        if not employee_username:
            raise HTTPException(status_code=400, detail="Sélectionnez l'employé concerné.")
        original_filename = sanitize_transport_receipt_filename(payload.original_filename)
        file_bytes = base64.b64decode(payload.file_b64, validate=True)
        if not file_bytes or len(file_bytes) > MAX_TRANSPORT_RECEIPT_BYTES:
            raise HTTPException(status_code=400, detail="Le fichier est vide ou dépasse 10 Mo.")
        extension, mime_type = validate_transport_receipt_file(file_bytes, original_filename)

        connection = get_db_connection()
        stored_path: Path | None = None
        try:
            with connection.cursor() as cursor:
                cursor.execute("SELECT 1 FROM users WHERE username = %s", (employee_username,))
                if not cursor.fetchone():
                    raise HTTPException(status_code=400, detail="L'employé sélectionné n'existe pas.")

                TRANSPORT_RECEIPTS_DIR.mkdir(parents=True, exist_ok=True)
                stored_filename = f"{secrets.token_urlsafe(32)}{extension}"
                stored_path = TRANSPORT_RECEIPTS_DIR / stored_filename
                stored_path.write_bytes(file_bytes)
                cursor.execute(
                    """
                    INSERT INTO transport_receipts
                        (employee_username, uploaded_by, original_filename, stored_filename, mime_type, size_bytes)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    """,
                    (employee_username, actor["username"], original_filename, stored_filename, mime_type, len(file_bytes)),
                )
                receipt_id = int(cursor.lastrowid)
                receipt = load_transport_receipt(cursor, receipt_id)
            connection.commit()
            return {"status": "success", "receipt": transport_receipt_response(receipt)}
        except Exception:
            connection.rollback()
            if stored_path:
                stored_path.unlink(missing_ok=True)
            raise
        finally:
            connection.close()
    except HTTPException:
        raise
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="Le fichier envoyé est invalide.") from None


@app.get("/transport-receipts")
def list_transport_receipts(
    include_archived: bool = True,
    actor: dict[str, Any] = Depends(get_current_user),
) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            conditions: list[str] = []
            params: list[Any] = []
            if "Admin" not in actor["role_tags"]:
                conditions.append("employee_username = %s")
                params.append(actor["username"])
            if not include_archived:
                conditions.append("archived = 0")
            where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
            cursor.execute(f"SELECT * FROM transport_receipts {where} ORDER BY created_at DESC, id DESC", tuple(params))
            return {"receipts": [transport_receipt_response(row) for row in cursor.fetchall()]}
    finally:
        connection.close()


@app.get("/transport-receipts/{receipt_id}/file")
def download_transport_receipt(
    receipt_id: int,
    actor: dict[str, Any] = Depends(get_current_user),
) -> FileResponse:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            receipt = load_transport_receipt(cursor, receipt_id)
        if not can_view_transport_receipt(actor, receipt):
            raise HTTPException(status_code=403, detail="Vous ne pouvez pas consulter ce justificatif.")
    finally:
        connection.close()

    stored_name = str(receipt["stored_filename"])
    if Path(stored_name).name != stored_name:
        raise HTTPException(status_code=404, detail="Fichier introuvable.")
    file_path = TRANSPORT_RECEIPTS_DIR / stored_name
    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="Fichier introuvable.")
    return FileResponse(file_path, media_type=str(receipt["mime_type"]), filename=str(receipt["original_filename"]))


@app.patch("/transport-receipts/{receipt_id}/archive")
def archive_transport_receipt(
    receipt_id: int,
    actor: dict[str, Any] = Depends(require_roles("Admin")),
) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            receipt = load_transport_receipt(cursor, receipt_id)
            if not receipt["archived"]:
                cursor.execute(
                    "UPDATE transport_receipts SET archived = 1, archived_at = %s, archived_by = %s WHERE id = %s",
                    (datetime.now(), actor["username"], receipt_id),
                )
                receipt = load_transport_receipt(cursor, receipt_id)
        connection.commit()
        return {"status": "success", "receipt": transport_receipt_response(receipt)}
    finally:
        connection.close()


@app.delete("/transport-receipts/{receipt_id}")
def delete_transport_receipt(
    receipt_id: int,
    actor: dict[str, Any] = Depends(get_current_user),
) -> dict[str, str]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            receipt = load_transport_receipt(cursor, receipt_id)
            can_delete = "Admin" in actor["role_tags"] or receipt["uploaded_by"] == actor["username"]
            if not can_delete:
                raise HTTPException(status_code=403, detail="Vous pouvez uniquement supprimer vos propres justificatifs.")
            cursor.execute("DELETE FROM transport_receipts WHERE id = %s", (receipt_id,))
        connection.commit()
    finally:
        connection.close()

    stored_name = str(receipt["stored_filename"])
    if Path(stored_name).name == stored_name:
        (TRANSPORT_RECEIPTS_DIR / stored_name).unlink(missing_ok=True)
    return {"status": "success", "message": "Justificatif supprimé."}


def submit_to_clawshow(
    *,
    document_url: str,
    employee_name: str,
    employee_email: str,
    month: int,
    year: int,
) -> dict[str, Any]:
    api_key = os.getenv("CLAWSHOW_ESIGN_API_KEY")
    if not api_key:
        raise HTTPException(status_code=503, detail="La signature électronique n'est pas configurée sur le serveur.")
    body = {
        "namespace": os.getenv("CLAWSHOW_ESIGN_NAMESPACE", "ilci"),
        "file_url": document_url,
        "signers": [{"name": employee_name, "email": employee_email, "order": 1, "role": "student"}],
        "reference_id": f"presence-{year}-{month:02d}-{secrets.token_hex(8)}",
    }
    request = urllib.request.Request(
        CLAWSHOW_CREATE_URL,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        try:
            provider_message = json.loads(error.read().decode("utf-8")).get("detail")
        except (UnicodeDecodeError, json.JSONDecodeError):
            provider_message = None
        raise HTTPException(
            status_code=502,
            detail=provider_message or "ClawShow n'a pas accepté la demande de signature.",
        ) from None
    except (urllib.error.URLError, TimeoutError, UnicodeDecodeError, json.JSONDecodeError):
        raise HTTPException(status_code=502, detail="ClawShow est indisponible pour le moment.") from None
    if not result.get("success"):
        raise HTTPException(status_code=502, detail="ClawShow n'a pas créé la demande de signature.")
    return result


@app.get("/esign/documents/{document_token}")
def get_esign_document(document_token: str) -> FileResponse:
    if not document_token or "/" in document_token or "\\" in document_token:
        raise HTTPException(status_code=404, detail="Document introuvable.")
    clean_expired_esign_documents()
    metadata_path = ESIGN_DOCUMENTS_DIR / f"{document_token}.json"
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        pdf_name = str(metadata["pdf_name"])
    except (OSError, KeyError, TypeError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="Document introuvable.") from None
    if not pdf_name.endswith(".pdf") or "/" in pdf_name or "\\" in pdf_name:
        raise HTTPException(status_code=404, detail="Document introuvable.")
    pdf_path = ESIGN_DOCUMENTS_DIR / pdf_name
    if not pdf_path.is_file():
        raise HTTPException(status_code=404, detail="Document introuvable.")
    return FileResponse(pdf_path, media_type="application/pdf", filename=str(metadata.get("filename", "fiche.pdf")))


@app.post("/send-fiche")
def send_fiche_for_signature(
    payload: ESignFicheRequest,
    _: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
) -> dict[str, str]:
    """Convert a generated fiche to PDF and create a ClawShow signature request."""
    try:
        if not payload.filename or payload.filename != payload.filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]:
            raise HTTPException(status_code=400, detail="Le nom du fichier est invalide.")
        if not 1 <= payload.month <= 12 or not 2000 <= payload.year <= 2100:
            raise HTTPException(status_code=400, detail="La période de la fiche est invalide.")
        file_bytes = base64.b64decode(payload.file_b64, validate=True)
        if not file_bytes or len(file_bytes) > MAX_FICHE_ATTACHMENT_BYTES:
            raise HTTPException(status_code=400, detail="Le fichier est vide ou dépasse 10 Mo.")

        clean_expired_esign_documents()
        pdf_bytes = convert_fiche_to_pdf(file_bytes, payload.filename)
        pdf_filename = f"{Path(payload.filename).stem}.pdf"
        document_token, pdf_name = store_esign_pdf(pdf_bytes, pdf_filename)
        try:
            result = submit_to_clawshow(
                document_url=f"{e_sign_public_base_url()}/esign/documents/{document_token}",
                employee_name=payload.employee_name.strip(),
                employee_email=str(payload.recipient_email),
                month=payload.month,
                year=payload.year,
            )
        except Exception:
            delete_esign_pdf(document_token, pdf_name)
            raise

        return {
            "status": "success",
            "message": f"Demande de signature envoyée à {payload.recipient_email}.",
            "document_id": str(result.get("document_id", "")),
        }
    except HTTPException:
        raise
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="Le fichier est invalide.") from None
    except Exception:
        raise HTTPException(status_code=500, detail="La demande de signature a échoué.") from None


@app.post("/login")
def login(payload: dict[str, Any]) -> dict[str, Any]:
    username = str(payload.get("username", ""))
    password = str(payload.get("password", ""))
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            user = load_user(cursor, username)
        if not user or not verify_password(password, user["password_hash"]):
            raise HTTPException(status_code=401, detail="Identifiants incorrects.")
        response = public_user(user)
        response["status"] = "success"
        response["auth_token"] = create_session_token(username)
        if bool(payload.get("remember_me")):
            with connection.cursor() as cursor:
                response["remember_token"] = create_remember_token(cursor, username)
            connection.commit()
        return response
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
        if not user:
            raise HTTPException(status_code=401, detail="Session persistante expirée.")
        response = public_user(user)
        response["status"] = "success"
        response["auth_token"] = create_session_token(user["username"])
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
def create_user(payload: dict[str, Any], actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable"))) -> dict[str, Any]:
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
        connection.commit()
        return {"status": "success", "message": f"Utilisateur {username} créé."}
    except pymysql.err.IntegrityError:
        raise HTTPException(status_code=400, detail="Ce nom ou cet e-mail est déjà utilisé.") from None
    finally:
        connection.close()


@app.get("/list-users")
def list_users(actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, Any]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT username FROM users ORDER BY username")
            candidates = [load_user(cursor, row["username"]) for row in cursor.fetchall()]
            if "Admin" in actor["role_tags"]:
                visible = candidates
            elif "Responsable" in actor["role_tags"]:
                visible = [user for user in candidates if user and user_can_manage(cursor, actor, user)]
            else:
                visible = [user for user in candidates if user and user["username"] == actor["username"]]
            return {"status": "success", "users": [public_user(user) for user in visible if user]}
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
            users = all_users(cursor)
            visible_users = users_visible_to_actor(cursor, actor, users)
            display_names = fiche_display_names_by_account(cursor)
            visible_usernames = {user["username"] for user in visible_users}
            contracts = []
            for user in visible_users:
                end_date = user.get("contract_end_date")
                if not isinstance(end_date, date) or bool(user.get("is_cdi")):
                    continue
                days_remaining = (end_date - today).days
                if 0 <= days_remaining <= 10:
                    contracts.append(
                        {
                            "employee_username": user["username"],
                            "employee_name": display_names.get(user["username"], user["username"]),
                            "contract_end_date": end_date,
                            "days_remaining": days_remaining,
                            "manager_username": user.get("manager_username"),
                        }
                    )
            if visible_usernames:
                placeholders = ", ".join(["%s"] * len(visible_usernames))
                cursor.execute(
                    f"""
                    SELECT employee_username, interview_year, due_date, notes, completed, completed_at, completed_by
                    FROM annual_interviews
                    WHERE interview_year = %s AND employee_username IN ({placeholders})
                    ORDER BY completed ASC, due_date ASC, employee_username ASC
                    """,
                    (current_year, *sorted(visible_usernames)),
                )
                interviews = cursor.fetchall()
            else:
                interviews = []
        connection.commit()
        return {
            "year": current_year,
            "contracts": sorted(contracts, key=lambda item: (item["days_remaining"], item["employee_username"].casefold())),
            "interviews": interviews,
        }
    finally:
        connection.close()


@app.patch("/annual-interviews/{employee_username}/{interview_year}")
def update_annual_interview(
    employee_username: str,
    interview_year: int,
    payload: AnnualInterviewUpdateRequest,
    actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
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
    actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
) -> dict[str, Any]:
    """Update organization metadata while enforcing the caller's management scope."""
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            target = require_target_access(cursor, actor, target_username, edit=True)
            tags = requested_role_tags(payload, target["role_tags"])
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
            updated = load_user(cursor, target_username)
        connection.commit()
        return {"status": "success", "user": public_user(updated)}
    finally:
        connection.close()


@app.patch("/users/{target_username}/contract")
def update_user_contract(
    target_username: str,
    payload: ContractUpdateRequest,
    actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
) -> dict[str, Any]:
    if payload.contract_start_date and payload.contract_end_date and payload.contract_end_date < payload.contract_start_date:
        raise HTTPException(status_code=400, detail="La fin de contrat ne peut pas précéder le début.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            require_target_access(cursor, actor, target_username, edit=True)
            cursor.execute(
                """
                UPDATE users
                SET contract_start_date = %s, contract_end_date = %s, is_cdi = %s
                WHERE username = %s
                """,
                (
                    payload.contract_start_date,
                    None if payload.is_cdi else payload.contract_end_date,
                    payload.is_cdi,
                    target_username,
                ),
            )
            updated = load_user(cursor, target_username)
        connection.commit()
        return {"status": "success", "user": public_user(updated)}
    finally:
        connection.close()


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
            target = load_user(cursor, target_username)
            manager = load_user(cursor, payload.manager_id)
            if not target:
                raise HTTPException(status_code=404, detail="Utilisateur introuvable.")
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
    finally:
        connection.close()


def fiche_record_belongs_to_user(record: dict[str, Any], username: str, email: str) -> bool:
    account_username = str(
        record.get("account_username") or record.get("employee_username") or record.get("user_username") or ""
    ).strip()
    record_email = str(record.get("email_employe") or "").strip()
    return account_username == username or (bool(record_email) and record_email.casefold() == email.casefold())


def purge_user_from_saved_fiches(cursor: pymysql.cursors.Cursor, username: str, email: str) -> None:
    """Remove a deleted person from every saved fiche configuration, not only their own account."""
    cursor.execute("SELECT user_id, form_content FROM Presence")
    for row in cursor.fetchall():
        owner = str(row["user_id"])
        if owner == username:
            cursor.execute("DELETE FROM Presence WHERE user_id = %s", (username,))
            continue
        content = row.get("form_content")
        try:
            config = content if isinstance(content, dict) else json.loads(content)
        except (TypeError, json.JSONDecodeError):
            continue
        employees = config.get("employes_data") if isinstance(config, dict) else None
        if not isinstance(employees, list):
            continue
        changed = False
        retained: list[dict[str, Any]] = []
        for employee in employees:
            if not isinstance(employee, dict):
                retained.append(employee)
                continue
            if fiche_record_belongs_to_user(employee, username, email):
                changed = True
                continue
            manager_email = str(employee.get("email_responsable") or "").strip()
            if manager_email and manager_email.casefold() == email.casefold():
                employee["responsable"] = ""
                employee["email_responsable"] = ""
                changed = True
            retained.append(employee)
        if changed:
            config["employes_data"] = retained
            cursor.execute(
                "UPDATE Presence SET form_content = %s WHERE user_id = %s",
                (json.dumps(config, default=str), owner),
            )


@app.delete("/delete-user/{username_to_delete}")
def delete_user(username_to_delete: str, actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable"))) -> dict[str, Any]:
    if username_to_delete == actor["username"]:
        raise HTTPException(status_code=400, detail="Vous ne pouvez pas supprimer votre propre compte.")
    connection = get_db_connection()
    receipt_files: list[str] = []
    try:
        with connection.cursor() as cursor:
            target = require_target_access(cursor, actor, username_to_delete, edit=True)
            if "Responsable" in actor["role_tags"] and "Admin" not in actor["role_tags"] and "Employe" not in target["role_tags"]:
                raise HTTPException(status_code=403, detail="Un Responsable ne peut supprimer qu'un Employé.")
            cursor.execute(
                """
                SELECT stored_filename FROM transport_receipts
                WHERE employee_username = %s OR uploaded_by = %s OR archived_by = %s
                """,
                (username_to_delete, username_to_delete, username_to_delete),
            )
            receipt_files = [str(row["stored_filename"]) for row in cursor.fetchall()]
            purge_user_from_saved_fiches(cursor, username_to_delete, str(target["email"]))
            cursor.execute("UPDATE users SET manager_username = NULL WHERE manager_username = %s", (username_to_delete,))
            cursor.execute(
                "DELETE FROM transport_receipts WHERE employee_username = %s OR uploaded_by = %s OR archived_by = %s",
                (username_to_delete, username_to_delete, username_to_delete),
            )
            cursor.execute("DELETE FROM annual_interviews WHERE employee_username = %s", (username_to_delete,))
            cursor.execute("UPDATE annual_interviews SET completed_by = NULL WHERE completed_by = %s", (username_to_delete,))
            cursor.execute(
                "DELETE FROM contract_end_reminders WHERE employee_username = %s OR recipient_email = %s",
                (username_to_delete, target["email"]),
            )
            cursor.execute("DELETE FROM user_group_memberships WHERE username = %s", (username_to_delete,))
            cursor.execute("DELETE FROM users WHERE username = %s", (username_to_delete,))
        connection.commit()
    finally:
        connection.close()
    for stored_name in receipt_files:
        if Path(stored_name).name == stored_name:
            (TRANSPORT_RECEIPTS_DIR / stored_name).unlink(missing_ok=True)
    return {"status": "success", "message": "Compte et données associées supprimés définitivement."}


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
            user = load_user(cursor, username)
            if not user or not verify_password(payload.current_password, user["password_hash"]):
                raise HTTPException(status_code=403, detail="Mot de passe actuel incorrect.")
            updates: list[str] = []
            params: list[Any] = []
            target_username = username
            if payload.new_username and payload.new_username != username:
                updates.append("username = %s")
                params.append(payload.new_username)
                target_username = payload.new_username
            if payload.new_email and payload.new_email != user["email"]:
                updates.append("email = %s")
                params.append(str(payload.new_email))
            if payload.new_password:
                updates.append("password_hash = %s")
                params.append(hash_password(payload.new_password))
            if updates:
                cursor.execute(f"UPDATE users SET {', '.join(updates)} WHERE username = %s", (*params, username))
                if payload.new_password:
                    cursor.execute(
                        "UPDATE users SET remember_token_hash = NULL, remember_token_expires = NULL WHERE username = %s",
                        (target_username,),
                    )
                if target_username != username:
                    cursor.execute("UPDATE Presence SET user_id = %s WHERE user_id = %s", (target_username, username))
                    cursor.execute("UPDATE user_group_memberships SET username = %s WHERE username = %s", (target_username, username))
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
        connection.commit()
        return {
            "status": "success",
            "message": "Profil mis à jour.",
            "username": target_username,
            "auth_token": create_session_token(target_username),
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
            cursor.execute("SELECT username FROM users WHERE email = %s", (email,))
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
                "SELECT username FROM users WHERE reset_token = %s AND reset_token_expires > %s",
                (token, datetime.now()),
            )
            user = cursor.fetchone()
            if not user:
                raise HTTPException(status_code=400, detail="Token invalide ou expiré.")
            cursor.execute(
                "UPDATE users SET password_hash = %s, reset_token = NULL, reset_token_expires = NULL, "
                "remember_token_hash = NULL, remember_token_expires = NULL WHERE username = %s",
                (hash_password(new_password), user["username"]),
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
            return content if isinstance(content, dict) else json.loads(content)
    finally:
        connection.close()


@app.post("/save-config/{user_id}")
def save_config(user_id: str, data: dict[str, Any], actor: dict[str, Any] = Depends(get_current_user)) -> dict[str, str]:
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            require_target_access(cursor, actor, user_id, edit=True)
            link_fiche_records_to_users(cursor, data)
            sync_contract_dates_from_config(cursor, data)
            serialized = json.dumps(data, default=str)
            cursor.execute(
                """
                INSERT INTO Presence (user_id, form_content) VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE form_content = %s
                """,
                (user_id, serialized, serialized),
            )
        connection.commit()
        return {"status": "success"}
    finally:
        connection.close()
