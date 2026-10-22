"""Document upload, scope, monthly metadata, archive and download routes."""
from typing import Any
from pathlib import Path
import base64, binascii, secrets
from datetime import datetime, timezone
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from document_downloads import safe_archive_component, MAX_ARCHIVE_BYTES
import pymysql
from fastapi import Depends, HTTPException
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel, ConfigDict, Field, StrictBool
try:
    from .document_numbers import allocate_document_number, ensure_document_number_schema
    from .destructive_audit import append_deletion_event
    from .role_policy import checked_resource_actor, is_admin, is_super_admin
    from .person_names import load_person_names, load_display_names, person_display_name
except ImportError:
    from document_numbers import allocate_document_number, ensure_document_number_schema
    from destructive_audit import append_deletion_event
    from role_policy import checked_resource_actor, is_admin, is_super_admin
    from person_names import load_person_names, load_display_names, person_display_name


class TransportReceiptUploadRequest(BaseModel):
    employee_username: str
    original_filename: str
    file_b64: str
    document_type: str = "transport"
    other_type: str | None = None
    receipt_month: int | None = None
    receipt_year: int | None = None


class ReceiptPeriodRequest(BaseModel):
    receipt_month: int
    receipt_year: int


class DocumentExportRequest(BaseModel):
    ids: list[int]


class ReceiptArchiveRequest(BaseModel):
    model_config = ConfigDict(extra='forbid')
    is_archived: StrictBool


class ReceiptDeleteRequest(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True)
    confirmation: str = Field(strict=True, min_length=1, max_length=32)


def register_document_routes(app, *, get_db_connection, get_current_user, require_roles,
                             load_user, user_can_manage, all_users, storage_dir, max_file_bytes):
    """Register routes using the same authentication dependencies as the main API."""
    TRANSPORT_RECEIPTS_DIR = storage_dir
    MAX_TRANSPORT_RECEIPT_BYTES = max_file_bytes
    app.router.add_event_handler("startup", lambda: ensure_document_number_schema(get_db_connection))
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


    def transport_receipt_response(row: dict[str, Any], display_names=None) -> dict[str, Any]:
        labels = display_names or {}
        return {
            "id": int(row["id"]),
            "document_number": row.get("document_number"),
            "employee_username": str(row["employee_username"]),
            "employee_display_name": labels.get(row["employee_username"], 'Ancien compte'),
            "uploaded_by": str(row["uploaded_by"]),
            "uploaded_by_display_name": labels.get(row["uploaded_by"], 'Ancien compte'),
            "original_filename": str(row["original_filename"]),
            "mime_type": str(row["mime_type"]),
            "size_bytes": int(row["size_bytes"]),
            "document_type": row.get("document_type"),
            "other_type": row.get("other_type"),
            "receipt_month": row.get("receipt_month"),
            "receipt_year": row.get("receipt_year"),
            "archived": bool(row["archived"]),
            "created_at": row["created_at"],
            "archived_at": row.get("archived_at"),
            "archived_by": row.get("archived_by"),
            "archived_by_display_name": (labels.get(row['archived_by'], 'Ancien compte')
                                         if row.get('archived_by') else None),
        }


    def load_transport_receipt(cursor: pymysql.cursors.Cursor, receipt_id: int, *, lock=False) -> dict[str, Any]:
        cursor.execute("SELECT * FROM transport_receipts WHERE id = %s" + (' FOR UPDATE' if lock else ''), (receipt_id,))
        receipt = cursor.fetchone()
        if not receipt:
            raise HTTPException(status_code=404, detail="Justificatif introuvable.")
        return receipt


    def can_view_transport_receipt(actor, receipt):
        if is_admin(actor) or receipt["employee_username"] == actor["username"]:
            return True
        if "Responsable" not in actor["role_tags"]:
            return False
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                target = load_user(cursor, receipt["employee_username"])
                return bool(target and user_can_manage(cursor, actor, target))
        finally:
            connection.close()


    @app.get("/transport-receipts/assignees")
    def list_transport_receipt_assignees(
        actor: dict[str, Any] = Depends(get_current_user),
    ) -> dict[str, Any]:
        """Only administrators may select a different upload owner."""
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                if is_admin(actor):
                    cursor.execute("SELECT username,email FROM users ORDER BY username")
                else:
                    cursor.execute("SELECT username,email FROM users WHERE username=%s", (actor['username'],))
                rows = cursor.fetchall()
                names = load_person_names(cursor, [row['username'] for row in rows])
                accounts = [{**row, 'full_name': names.get(row['username']),
                             'display_name': person_display_name(row, names)} for row in rows]
                return {"users": [str(row['username']) for row in rows], 'accounts': accounts}
        finally:
            connection.close()


    @app.post("/transport-receipts")
    def upload_transport_receipt(
        payload: TransportReceiptUploadRequest,
        actor: dict[str, Any] = Depends(get_current_user),
    ) -> dict[str, Any]:
        employee_username = str(payload.employee_username).strip()
        if not employee_username:
            raise HTTPException(status_code=400, detail="Sélectionnez l'employé concerné.")
        if not is_admin(actor) and employee_username != actor["username"]:
            raise HTTPException(status_code=403, detail="Vous pouvez déposer des documents uniquement pour votre propre compte.")
        receipt_month, receipt_year = payload.receipt_month, payload.receipt_year
        if receipt_month is None and receipt_year is None:
            # Compatibility with an already-open older upload page.
            now = datetime.now()
            receipt_month, receipt_year = now.month, now.year
        if receipt_month is None or receipt_year is None or not 1 <= receipt_month <= 12 or not 2000 <= receipt_year <= 2100:
            raise HTTPException(status_code=400, detail="Choisissez un mois et une année valides pour le document.")
        document_type = payload.document_type.strip().lower()
        if document_type not in {"transport", "autre"}:
            raise HTTPException(status_code=400, detail="Choisissez Transport ou Autre comme type de document.")
        other_type = str(payload.other_type or "").strip() if document_type == "autre" else None
        if document_type == "autre" and (not other_type or len(other_type) > 120):
            raise HTTPException(status_code=400, detail="Précisez le type de document (1 à 120 caractères).")
        try:
            original_filename = sanitize_transport_receipt_filename(payload.original_filename)
            file_bytes = base64.b64decode(payload.file_b64, validate=True)
            if not file_bytes or len(file_bytes) > MAX_TRANSPORT_RECEIPT_BYTES:
                raise HTTPException(status_code=400, detail="Le fichier est vide ou dépasse 10 Mo.")
            extension, mime_type = validate_transport_receipt_file(file_bytes, original_filename)

            connection = get_db_connection()
            stored_path: Path | None = None
            try:
                with connection.cursor() as cursor:
                    current = checked_resource_actor(cursor, actor, load_user, lock=True)
                    if not is_admin(current) and employee_username != current['username']:
                        raise HTTPException(403, 'Vous pouvez déposer uniquement pour votre propre compte.')
                    cursor.execute("SELECT 1 FROM users WHERE username = %s", (employee_username,))
                    if not cursor.fetchone():
                        raise HTTPException(status_code=400, detail="L'employé sélectionné n'existe pas.")

                    TRANSPORT_RECEIPTS_DIR.mkdir(parents=True, exist_ok=True)
                    stored_filename = f"{secrets.token_urlsafe(32)}{extension}"
                    stored_path = TRANSPORT_RECEIPTS_DIR / stored_filename
                    stored_path.write_bytes(file_bytes)
                    uploaded_at = datetime.now(timezone.utc)
                    document_number = allocate_document_number(cursor, uploaded_at)
                    cursor.execute(
                        """
                        INSERT INTO transport_receipts
                            (employee_username, uploaded_by, original_filename, stored_filename, mime_type, size_bytes, document_type, other_type, receipt_month, receipt_year, document_number, created_at)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, FROM_UNIXTIME(%s))
                        """,
                        (employee_username, current["username"], original_filename, stored_filename, mime_type, len(file_bytes), document_type, other_type, receipt_month, receipt_year, document_number, uploaded_at.timestamp()),
                    )
                    receipt_id = int(cursor.lastrowid)
                    receipt = load_transport_receipt(cursor, receipt_id)
                    labels = load_display_names(cursor, [receipt['employee_username'], receipt['uploaded_by']])
                connection.commit()
                return {"status": "success", "receipt": transport_receipt_response(receipt, labels)}
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
                if not is_admin(actor):
                    permitted = [actor["username"]]
                    if "Responsable" in actor["role_tags"]:
                        permitted += [u["username"] for u in all_users(cursor) if user_can_manage(cursor, actor, u)]
                    conditions.append("employee_username IN (" + ",".join(["%s"] * len(permitted)) + ")")
                    params.extend(permitted)
                if not include_archived:
                    conditions.append("archived = 0")
                where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
                cursor.execute(f"SELECT * FROM transport_receipts {where} ORDER BY created_at DESC, id DESC", tuple(params))
                rows = cursor.fetchall()
                usernames = {row.get(key) for row in rows
                             for key in ('employee_username', 'uploaded_by', 'archived_by') if row.get(key)}
                labels = load_display_names(cursor, usernames)
                return {"receipts": [transport_receipt_response(row, labels) for row in rows]}
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
        filename = str(receipt["original_filename"])
        if receipt.get("document_number"):
            filename = f"{receipt['document_number']}_{filename}"
        return FileResponse(file_path, media_type=str(receipt["mime_type"]), filename=filename)


    @app.post("/transport-receipts/export")
    def export_documents(payload: DocumentExportRequest,
                         actor: dict[str, Any] = Depends(get_current_user)):
        ids = list(dict.fromkeys(payload.ids))
        if not ids or len(ids) > 100:
            raise HTTPException(status_code=400, detail="Sélectionnez entre 1 et 100 documents.")
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                placeholders = ','.join(['%s'] * len(ids))
                cursor.execute(f'SELECT id,employee_username FROM transport_receipts WHERE id IN ({placeholders})', tuple(ids))
                labels = load_display_names(cursor, [row['employee_username'] for row in cursor.fetchall()])
        finally:
            connection.close()
        output = BytesIO()
        total = 0
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
            for receipt_id in ids:
                # Reuse the individual download's current authorization and path checks.
                response = download_transport_receipt(receipt_id, actor)
                path = Path(response.path)
                if total + path.stat().st_size > MAX_ARCHIVE_BYTES:
                    raise HTTPException(status_code=400, detail="La sélection dépasse 100 Mo.")
                with path.open("rb") as source:
                    data = source.read(MAX_ARCHIVE_BYTES - total + 1)
                total += len(data)
                if total > MAX_ARCHIVE_BYTES:
                    raise HTTPException(status_code=400, detail="La sélection dépasse 100 Mo.")
                connection = get_db_connection()
                try:
                    with connection.cursor() as cursor:
                        receipt = load_transport_receipt(cursor, receipt_id)
                    category = receipt.get("other_type") if receipt.get("document_type") == "autre" else receipt.get("document_type")
                    name = "_".join(safe_archive_component(value) for value in (
                        receipt.get("document_number") or receipt_id,
                        labels.get(receipt['employee_username'], 'Ancien compte'),
                        category or "Non renseigné", receipt["original_filename"]))
                    archive.writestr(name, data)
                finally:
                    connection.close()
        return Response(output.getvalue(), media_type="application/zip",
                        headers={"Content-Disposition": 'attachment; filename="documents.zip"'})

    @app.patch("/transport-receipts/{receipt_id}/period")
    def update_receipt_period(
        receipt_id: int,
        payload: ReceiptPeriodRequest,
        actor: dict[str, Any] = Depends(get_current_user),
    ) -> dict[str, Any]:
        if not 1 <= payload.receipt_month <= 12 or not 2000 <= payload.receipt_year <= 2100:
            raise HTTPException(status_code=400, detail="Choisissez un mois et une année valides pour le document.")
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user, lock=True)
                if not is_admin(current):
                    raise HTTPException(403, 'La modification du mois est réservée aux administrateurs.')
                load_transport_receipt(cursor, receipt_id, lock=True)
                cursor.execute(
                    "UPDATE transport_receipts SET receipt_month = %s, receipt_year = %s WHERE id = %s",
                    (payload.receipt_month, payload.receipt_year, receipt_id),
                )
                receipt = load_transport_receipt(cursor, receipt_id)
                labels = load_display_names(cursor, [receipt.get(key) for key in
                    ('employee_username', 'uploaded_by', 'archived_by') if receipt.get(key)])
            connection.commit()
            return {"status": "success", "receipt": transport_receipt_response(receipt, labels)}
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()


    @app.patch("/transport-receipts/{receipt_id}/archive")
    def archive_transport_receipt(
        receipt_id: int,
        payload: ReceiptArchiveRequest | None = None,
        actor: dict[str, Any] = Depends(get_current_user),
    ) -> dict[str, Any]:
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user, lock=True)
                if 'Responsable' in current.get('role_tags', []) and not is_admin(current):
                    raise HTTPException(403, 'Un responsable dispose ici d’un accès en lecture seule.')
                receipt = load_transport_receipt(cursor, receipt_id, lock=True)
                if not is_admin(current) and current['username'] not in {receipt['employee_username'], receipt['uploaded_by']}:
                    raise HTTPException(403, 'Vous pouvez archiver uniquement vos propres justificatifs.')
                archived = payload.is_archived if payload is not None else True
                if bool(receipt['archived']) != archived:
                    cursor.execute(
                        "UPDATE transport_receipts SET archived = %s, archived_at = %s, archived_by = %s WHERE id = %s",
                        (archived, datetime.now() if archived else None, current['username'] if archived else None, receipt_id),
                    )
                    receipt = load_transport_receipt(cursor, receipt_id)
                labels = load_display_names(cursor, [receipt.get(key) for key in
                    ('employee_username', 'uploaded_by', 'archived_by') if receipt.get(key)])
            connection.commit()
            return {"status": "success", "receipt": transport_receipt_response(receipt, labels)}
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()


    @app.delete("/transport-receipts/{receipt_id}")
    def delete_transport_receipt(
        receipt_id: int,
        payload: ReceiptDeleteRequest | None = None,
        actor: dict[str, Any] = Depends(get_current_user),
    ) -> dict[str, Any]:
        connection = get_db_connection()
        try:
            with connection.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user, lock=True)
                if not is_super_admin(current):
                    raise HTTPException(403, 'La suppression définitive est réservée au SuperAdmin.')
                receipt = load_transport_receipt(cursor, receipt_id, lock=True)
                if not receipt['archived']:
                    raise HTTPException(409, 'Archivez ce document avant sa suppression définitive.')
                if payload is None or payload.confirmation != str(receipt_id):
                    raise HTTPException(400, 'Confirmez le numéro exact du document.')
                stored_name = str(receipt['stored_filename'])
                if not stored_name or Path(stored_name).name != stored_name:
                    raise HTTPException(409, 'Le document doit être vérifié avant suppression.')
                cursor.execute("DELETE FROM transport_receipts WHERE id = %s", (receipt_id,))
                append_deletion_event(cursor, current, 'document', receipt_id,
                                      {'deleted_records': 1, 'document_count': 1,
                                       'confirmed': True, 'previous_is_archived': True})
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

        try:
            (TRANSPORT_RECEIPTS_DIR / stored_name).unlink(missing_ok=True)
        except OSError:
            return {'status': 'cleanup_pending', 'cleanup_pending': True,
                    'message': 'Document retiré du registre ; le nettoyage du fichier reste à terminer.'}
        return {"status": "success", "cleanup_pending": False, "message": "Justificatif supprimé."}


