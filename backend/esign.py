"""LibreOffice PDF conversion and the ClawShow attendance-signature workflow."""
from typing import Any
from pathlib import Path
import base64, binascii, secrets
from fastapi import Depends, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
import json, os, time, shutil, subprocess, tempfile
import urllib.error, urllib.request
from urllib.parse import urlsplit
from fastapi.responses import Response
from pydantic import EmailStr
if __package__:
    from .signature_flow import resolve_manager, signing_people, attendance_signature_fields
    from .signature_placement import place_signature_fields
else:
    from signature_flow import resolve_manager, signing_people, attendance_signature_fields
    from signature_placement import place_signature_fields

CLAWSHOW_CREATE_URL = "https://esign.clawshow.ai/esign/create"


def sandbox_signing_links(result):
    """Expose signing pages only; PDF downloads and field setup are not signer links."""
    links = []
    document_id = str(result.get('document_id') or '')
    def visit(value):
        if isinstance(value, dict):
            for key, child in value.items():
                if isinstance(child, (dict, list)) or 'url' in key.lower() or 'link' in key.lower():
                    visit(child)
        elif isinstance(value, list):
            for child in value: visit(child)
        elif isinstance(value, str):
            try:
                parsed = urlsplit(value)
                host = parsed.hostname or ''
                signing_page = (parsed.path.rstrip('/') == f'/esign/{document_id}' if document_id
                                else parsed.path.startswith('/sign/'))
                if signing_page and parsed.scheme == 'https' and (host == 'clawshow.ai' or host.endswith('.clawshow.ai')) and not parsed.username and value not in links:
                    links.append(value)
            except ValueError:
                pass
    visit(result)
    return links


def sandbox_signature_actions(result, employee_email, manager_email):
    """Use signer metadata when supplied; never assign a manager from URL order alone."""
    links = sandbox_signing_links(result)
    labels = {}
    def visit(value, inherited=None):
        if isinstance(value, dict):
            email = str(value.get('email') or '').strip().casefold()
            party = ('employee' if email == employee_email.strip().casefold() else
                     'manager' if email == manager_email.strip().casefold() else inherited)
            for key, child in value.items():
                if isinstance(child, str) and child in links and party:
                    labels[child] = party
                elif isinstance(child, (dict, list)):
                    visit(child, party)
        elif isinstance(value, list):
            for child in value: visit(child, inherited)
    visit(result)
    return [{'url': link, 'party': labels.get(link, 'unknown')} for link in links]

class ESignFicheRequest(BaseModel):
    recipient_email: EmailStr
    employee_name: str
    month: int
    year: int
    filename: str
    file_b64: str


class FichePdfRequest(BaseModel):
    filename: str
    file_b64: str


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


def submit_to_clawshow(
    *,
    document_url: str,
    employee_name: str,
    employee_email: str,
    manager_name: str,
    manager_email: str,
    month: int,
    year: int,
    intern: bool = False,
    pdf_bytes: bytes | None = None,
) -> dict[str, Any]:
    api_key = os.getenv("CLAWSHOW_ESIGN_API_KEY")
    if not api_key:
        raise HTTPException(status_code=503, detail="La signature électronique n'est pas configurée sur le serveur.")
    if not pdf_bytes:
        raise HTTPException(400, "Régénérez le PDF avant de l'envoyer pour signature.")
    body = {
        "namespace": os.getenv("CLAWSHOW_ESIGN_NAMESPACE", "ilci"),
        "file_url": document_url,
        "signature_level": "aes",
        "title": f"Feuille de présence — {month:02d}/{year}",
        "signers": signing_people(employee_name, employee_email, manager_name, manager_email),
        "signature_fields": place_signature_fields(pdf_bytes, attendance_signature_fields(intern=intern)),
        "schoolInfo": {
            "name": os.getenv("PRESENCE_ESIGN_ORGANIZATION_NAME", "ILCI"),
            "fullName": os.getenv("PRESENCE_ESIGN_ORGANIZATION_FULL_NAME", "Institut de Langues et de Commerce International"),
            "representative": {
                "name": os.getenv("PRESENCE_ESIGN_REPRESENTATIVE_NAME") or manager_name,
                "title": os.getenv("PRESENCE_ESIGN_REPRESENTATIVE_TITLE", "Responsable"),
            },
        },
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
            provider_error = json.loads(error.read().decode("utf-8"))
            provider_message = provider_error.get("detail") or provider_error.get("error") or provider_error.get("message")
        except (UnicodeDecodeError, json.JSONDecodeError):
            provider_message = None
        raise HTTPException(
            status_code=502,
            detail=str(provider_message) if provider_message else "ClawShow n'a pas accepté la demande de signature.",
        ) from None
    except (urllib.error.URLError, TimeoutError, UnicodeDecodeError, json.JSONDecodeError):
        raise HTTPException(status_code=502, detail="ClawShow est indisponible pour le moment.") from None
    if not result.get("success"):
        raise HTTPException(status_code=502, detail="ClawShow n'a pas créé la demande de signature.")
    return result


def register_esign_routes(app, *, get_db_connection, require_roles, load_user,
                          require_target_access, storage_dir, document_duration_hours, max_file_bytes):
    ESIGN_DOCUMENTS_DIR = storage_dir
    ESIGN_DOCUMENT_DURATION_HOURS = document_duration_hours
    MAX_FICHE_ATTACHMENT_BYTES = max_file_bytes
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


    @app.post("/fiches/pdf")
    def generate_fiche_pdf(
        payload: FichePdfRequest,
        _: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
    ) -> Response:
        """Convert one generated attendance sheet to a downloadable PDF."""
        try:
            filename = payload.filename.strip()
            if not filename or filename != filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]:
                raise HTTPException(status_code=400, detail="Le nom du fichier est invalide.")
            file_bytes = base64.b64decode(payload.file_b64, validate=True)
            if not file_bytes or len(file_bytes) > MAX_FICHE_ATTACHMENT_BYTES:
                raise HTTPException(status_code=400, detail="Le fichier est vide ou dépasse 10 Mo.")

            pdf_bytes = convert_fiche_to_pdf(file_bytes, filename)
            return Response(
                content=pdf_bytes,
                media_type="application/pdf",
            )
        except HTTPException:
            raise
        except (binascii.Error, ValueError):
            raise HTTPException(status_code=400, detail="Le fichier est invalide.") from None


    @app.post("/send-fiche")
    def send_fiche_for_signature(
        payload: ESignFicheRequest,
        actor: dict[str, Any] = Depends(require_roles("Admin", "Responsable")),
    ) -> dict[str, Any]:
        """Convert a generated fiche to PDF and create a ClawShow signature request."""
        manager = resolve_manager(payload.recipient_email, actor, get_db_connection, load_user, require_target_access)
        recipient_policy = getattr(app.state, 'esign_recipient_policy', None)
        if recipient_policy:
            recipient_policy(str(payload.recipient_email), manager['email'])
        try:
            if not payload.filename or payload.filename != payload.filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]:
                raise HTTPException(status_code=400, detail="Le nom du fichier est invalide.")
            if not 1 <= payload.month <= 12 or not 2000 <= payload.year <= 2100:
                raise HTTPException(status_code=400, detail="La période de la fiche est invalide.")
            file_bytes = base64.b64decode(payload.file_b64, validate=True)
            if not file_bytes or len(file_bytes) > MAX_FICHE_ATTACHMENT_BYTES:
                raise HTTPException(status_code=400, detail="Le fichier est vide ou dépasse 10 Mo.")

            clean_expired_esign_documents()
            conversion_started = time.monotonic()
            pdf_bytes = convert_fiche_to_pdf(file_bytes, payload.filename)
            conversion_seconds = time.monotonic() - conversion_started
            pdf_filename = f"{Path(payload.filename).stem}.pdf"
            document_token, pdf_name = store_esign_pdf(pdf_bytes, pdf_filename)
            try:
                if recipient_policy:
                    recipient_policy(str(payload.recipient_email), manager['email'])
                provider_started = time.monotonic()
                result = submit_to_clawshow(
                    document_url=f"{e_sign_public_base_url()}/esign/documents/{document_token}",
                    employee_name=manager['employee_name'],
                    employee_email=str(payload.recipient_email),
                    manager_name=manager["name"],
                    manager_email=manager["email"],
                    month=payload.month,
                    year=payload.year,
                    intern=Path(payload.filename).suffix.lower() == ".docx",
                    pdf_bytes=pdf_bytes,
                )
            except Exception:
                delete_esign_pdf(document_token, pdf_name)
                raise

            return {
                "status": "success",
                "message": f"Signature du salarié : {payload.recipient_email}, puis du responsable : {manager['email']}.",
                "document_id": str(result.get("document_id", "")),
                "signing_links": sandbox_signing_links(result) if os.getenv('CLAWSHOW_ESIGN_API_KEY', '').startswith('sk_test_') else [],
                "signing_actions": sandbox_signature_actions(result, str(payload.recipient_email), manager['email']) if os.getenv('CLAWSHOW_ESIGN_API_KEY', '').startswith('sk_test_') else [],
                "timings": {"pdf_conversion_seconds": round(conversion_seconds, 2),
                            "provider_seconds": round(time.monotonic() - provider_started, 2)},
            }
        except HTTPException:
            raise
        except (binascii.Error, ValueError):
            raise HTTPException(status_code=400, detail="Le fichier est invalide.") from None
        except Exception:
            raise HTTPException(status_code=500, detail="La demande de signature a échoué.") from None

    # Both frontends use the same conversion, manager resolution and provider request.
    return send_fiche_for_signature


