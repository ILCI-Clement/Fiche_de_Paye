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
import time
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any

import bcrypt
import pymysql
from fastapi import Depends, FastAPI, Header, HTTPException
from pydantic import BaseModel, EmailStr


app = FastAPI(title="Presence API")
VALID_ROLES = {"Admin", "Responsable", "Employe"}
ROLE_PRIORITY = ("Admin", "Responsable", "Employe")
VALID_EMPLOYEE_TYPES = {"salarie", "stagiaire"}
SESSION_DURATION_SECONDS = 8 * 60 * 60
REMEMBER_SESSION_DURATION_DAYS = 30
PASSWORD_RESET_DURATION_MINUTES = 15
MAX_FICHE_ATTACHMENT_BYTES = 10 * 1024 * 1024


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


class SendFicheRequest(BaseModel):
    recipient_email: EmailStr
    employee_name: str
    month: int
    year: int
    filename: str
    file_b64: str


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
        connection.commit()
    finally:
        connection.close()


@app.on_event("startup")
def migrate_database() -> None:
    ensure_organization_schema()


def load_user(cursor: pymysql.cursors.Cursor, username: str) -> dict[str, Any] | None:
    cursor.execute(
        """
        SELECT username, email, password_hash, is_admin, role, role_tags, employee_type, manager_username, created_at
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
        "created_at": user.get("created_at"),
    }


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


@app.post("/send-fiche")
def send_fiche(
    payload: SendFicheRequest,
    _: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
) -> dict[str, str]:
    """Send a generated attendance sheet to the selected employee."""
    try:
        if not payload.filename or payload.filename != payload.filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]:
            raise HTTPException(status_code=400, detail="Le nom du fichier joint est invalide.")
        file_bytes = base64.b64decode(payload.file_b64, validate=True)
        if not file_bytes or len(file_bytes) > MAX_FICHE_ATTACHMENT_BYTES:
            raise HTTPException(status_code=400, detail="Le fichier joint est vide ou dépasse 10 Mo.")

        message = MIMEMultipart()
        message["From"] = SMTP_USER
        message["To"] = payload.recipient_email
        message["Subject"] = (
            f"Fiche de présence - {payload.employee_name} "
            f"({payload.month:02d}/{payload.year})"
        )

        body_text = f"""Bonjour {payload.employee_name},

Veuillez trouver ci-joint votre fiche de présence pour le mois {payload.month:02d}/{payload.year}.

Cordialement,
L'équipe RH / Administration"""

        message.attach(MIMEText(body_text, "plain"))

        attachment = MIMEApplication(file_bytes, Name=payload.filename)
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=payload.filename,
        )
        message.attach(attachment)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, payload.recipient_email, message.as_string())

        return {
            "status": "success",
            "message": f"Fiche envoyée avec succès à {payload.recipient_email}.",
        }
    except HTTPException:
        raise
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="Le fichier joint est invalide.") from None
    except smtplib.SMTPException:
        raise HTTPException(
            status_code=502,
            detail="Le serveur de messagerie n'a pas pu envoyer la fiche.",
        ) from None
    except Exception:
        raise HTTPException(status_code=500, detail="L'envoi de la fiche a échoué.") from None


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


@app.delete("/delete-user/{username_to_delete}")
def delete_user(username_to_delete: str, actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable"))) -> dict[str, Any]:
    if username_to_delete == actor["username"]:
        raise HTTPException(status_code=400, detail="Vous ne pouvez pas supprimer votre propre compte.")
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            target = require_target_access(cursor, actor, username_to_delete, edit=True)
            if "Responsable" in actor["role_tags"] and "Admin" not in actor["role_tags"] and "Employe" not in target["role_tags"]:
                raise HTTPException(status_code=403, detail="Un Responsable ne peut supprimer qu'un Employé.")
            cursor.execute("UPDATE users SET manager_username = NULL WHERE manager_username = %s", (username_to_delete,))
            cursor.execute("DELETE FROM user_group_memberships WHERE username = %s", (username_to_delete,))
            cursor.execute("DELETE FROM users WHERE username = %s", (username_to_delete,))
        connection.commit()
        return {"status": "success", "message": "Compte supprimé. Les fiches de présence existantes sont conservées."}
    finally:
        connection.close()


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
