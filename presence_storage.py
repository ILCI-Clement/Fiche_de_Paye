"""Client for the existing Presence API storage gateway."""

from __future__ import annotations

from copy import deepcopy
from datetime import date
from typing import Any
from urllib.parse import quote, urlsplit

import requests

from app_core import ensure_workspace


class PresenceStorageError(RuntimeError):
    """Raised when a workspace cannot be read from or written to the API."""


def serialize_dates(value: Any) -> Any:
    """Return a JSON-compatible copy without mutating the workspace."""
    if isinstance(value, dict):
        return {key: serialize_dates(item) for key, item in value.items()}
    if isinstance(value, list):
        return [serialize_dates(item) for item in value]
    if isinstance(value, date):
        return value.isoformat()
    return value


class PresenceApiStorage:
    """Read and write a user's JSON workspace through the Presence API."""

    def __init__(
        self,
        base_url: str,
        token: str,
        *,
        timeout: float = 10,
        http_client: Any = requests,
    ) -> None:
        normalized_url = str(base_url or "").strip().rstrip("/")
        parsed_url = urlsplit(normalized_url)
        if parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc:
            raise ValueError("L’adresse de l’API Presence est invalide.")
        if not str(token or "").strip():
            raise ValueError("Le jeton de l’API Presence est manquant.")
        if timeout <= 0:
            raise ValueError("Le délai d’attente de l’API doit être positif.")

        self.base_url = normalized_url
        self.timeout = timeout
        self.http_client = http_client
        self.headers = {
            "Authorization": f"Bearer {token.strip()}",
            "Content-Type": "application/json",
        }

    def _workspace_url(self, action: str, username: str) -> str:
        normalized_username = str(username or "").strip()
        if not normalized_username:
            raise PresenceStorageError("Le nom du compte est manquant.")
        return f"{self.base_url}/{action}/{quote(normalized_username, safe='')}"

    def load_workspace(self, username: str) -> dict[str, Any]:
        """Load and validate the raw JSON configuration returned by the API."""
        try:
            response = self.http_client.get(
                self._workspace_url("get-config", username),
                headers=self.headers,
                timeout=self.timeout,
            )
            response.raise_for_status()
            payload = response.json()
            return ensure_workspace(deepcopy(payload))
        except (requests.RequestException, ValueError, TypeError) as error:
            raise PresenceStorageError(
                f"Impossible de charger les données depuis l’API Presence : {error}"
            ) from error

    def save_workspace(self, username: str, workspace: dict[str, Any]) -> None:
        """Save a JSON-compatible copy of a validated workspace."""
        try:
            response = self.http_client.post(
                self._workspace_url("save-config", username),
                headers=self.headers,
                json=serialize_dates(deepcopy(workspace)),
                timeout=self.timeout,
            )
            response.raise_for_status()
        except (requests.RequestException, ValueError, TypeError) as error:
            raise PresenceStorageError(
                f"Impossible de sauvegarder les données via l’API Presence : {error}"
            ) from error
