"""Small helpers for authenticated calls from Streamlit pages to Presence API."""

from __future__ import annotations

from typing import Any

import streamlit as st


def api_url() -> str:
    return str(st.secrets["URL_PRESENCE"])


def authenticated_headers() -> dict[str, str]:
    user: dict[str, Any] | None = st.session_state.get("user")
    token = user.get("auth_token") if user else None
    if not token:
        return {"Content-Type": "application/json"}
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
