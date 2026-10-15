"""Build document archives from authorized individual downloads."""

from io import BytesIO
import re
from zipfile import ZIP_DEFLATED, ZipFile

import requests


MAX_ARCHIVE_BYTES = 100 * 1024 * 1024


def safe_archive_component(value: object) -> str:
    text = re.sub(r'[\x00-\x1f<>:"/\\|?*]', "_", str(value))
    return text.strip(" .")[:100] or "document"


def build_document_archive(receipts: list[dict], api_url: str, headers: dict) -> bytes:
    """Fetch each file with the caller's permissions; never return a partial ZIP."""
    output = BytesIO()
    total_bytes = 0
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for receipt in receipts:
            try:
                response = requests.get(
                    f"{api_url}/transport-receipts/{int(receipt['id'])}/file",
                    headers=headers, timeout=30, stream=True,
                )
                with response:
                    if response.status_code != 200:
                        raise ValueError(
                            f"Impossible de télécharger {receipt['original_filename']}. "
                            "Actualisez la liste et réessayez. Aucun ZIP partiel n'a été créé."
                        )
                    data = BytesIO()
                    for chunk in response.iter_content(chunk_size=64 * 1024):
                        total_bytes += len(chunk)
                        if total_bytes > MAX_ARCHIVE_BYTES:
                            raise ValueError("La sélection dépasse 100 Mo. Téléchargez les documents en plusieurs lots.")
                        data.write(chunk)
            except requests.RequestException as exc:
                raise ValueError("Le serveur est indisponible. Réessayez la préparation du ZIP.") from exc
            category = receipt.get("other_type") if receipt.get("document_type") == "autre" else receipt.get("document_type")
            name = (
                f"{safe_archive_component(receipt['employee_username'])}_"
                f"{safe_archive_component(category or 'Non renseigne')}_"
                f"{int(receipt['id'])}_"
                f"{safe_archive_component(receipt['original_filename'])}"
            )
            archive.writestr(name, data.getvalue())
    return output.getvalue()
