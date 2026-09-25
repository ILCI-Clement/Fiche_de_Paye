"""Small helpers for authenticated calls from Streamlit pages to Presence API."""

from __future__ import annotations

from typing import Any

import requests
import streamlit as st


def api_url() -> str:
    return str(st.secrets["URL_PRESENCE"])


def authenticated_headers() -> dict[str, str]:
    user: dict[str, Any] | None = st.session_state.get("user")
    token = user.get("auth_token") if user else None
    if not token:
        return {"Content-Type": "application/json"}
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def safe_api_request(
    method: str,
    url: str,
    *,
    timeout: int = 15,
    **kwargs: Any,
) -> requests.Response | None:
    """Make an API request without allowing a temporary outage to crash a Streamlit page."""
    try:
        return requests.request(method, url, timeout=timeout, **kwargs)
    except requests.RequestException as error:
        st.error(f"Impossible de joindre le serveur : {error}")
        return None
