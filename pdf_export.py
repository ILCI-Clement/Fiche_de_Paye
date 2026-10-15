"""Convert generated attendance files through the authenticated PDF endpoint."""
import base64
import requests


def convert_fiche_pdf(api_url, headers, filename, data):
    try:
        response = requests.post(
            f"{api_url}/fiches/pdf",
            headers=headers,
            json={"filename": filename, "file_b64": base64.b64encode(data).decode("ascii")},
            timeout=100,
        )
    except requests.RequestException as error:
        raise ValueError("Le serveur de conversion PDF est indisponible. Réessayez.") from error
    if response.status_code != 200:
        try:
            detail = response.json().get("detail", "La conversion a échoué.")
        except (ValueError, AttributeError):
            detail = "La conversion a échoué. Réessayez."
        raise ValueError(str(detail))
    if not response.content.startswith(b"%PDF-"):
        raise ValueError("Le serveur n'a pas renvoyé un fichier PDF valide.")
    return response.content
