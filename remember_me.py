"""Persistent browser token cookie for the optional remember-me login flow."""

from __future__ import annotations

import streamlit as st


COOKIE_NAME = "presence_remember_token"
COOKIE_MAX_AGE_SECONDS = 30 * 24 * 60 * 60

_COOKIE_COMPONENT = st.components.v2.component(
    "presence_remember_cookie",
    html="""
    <span aria-hidden="true"></span>
    """,
    js="""
    export default function (component) {
      const { data, setStateValue } = component
      const name = data?.name
      if (!name) return

      const getCookie = (cookieName) => {
        const prefix = `${cookieName}=`
        return document.cookie.split(';').map((entry) => entry.trim())
          .find((entry) => entry.startsWith(prefix))?.slice(prefix.length) ?? ''
      }

      const secureAttribute = window.location.protocol === 'https:' ? '; Secure' : ''
      if (data?.clear) {
        document.cookie = `${name}=; Path=/; Max-Age=0; SameSite=Strict${secureAttribute}`
      } else if (data?.token) {
        const encodedToken = encodeURIComponent(data.token)
        document.cookie = `${name}=${encodedToken}; Path=/; Max-Age=${data.maxAge}; SameSite=Strict${secureAttribute}`
      }

      let rememberedToken = ''
      try {
        rememberedToken = decodeURIComponent(getCookie(name))
      } catch (_) {
        document.cookie = `${name}=; Path=/; Max-Age=0; SameSite=Strict${secureAttribute}`
      }
      setStateValue('token', data?.clear ? '' : rememberedToken)
    }
    """,
)


def remembered_token(*, token_to_store: str | None = None, clear: bool = False) -> str:
    """Synchronize the browser cookie and return its currently reported token."""
    state = st.session_state.get("presence_remember_cookie", {})
    current_token = str(state.get("token") or "") if isinstance(state, dict) else ""
    result = _COOKIE_COMPONENT(
        key="presence_remember_cookie",
        data={
            "name": COOKIE_NAME,
            "token": token_to_store or "",
            "clear": clear,
            "maxAge": COOKIE_MAX_AGE_SECONDS,
        },
        on_token_change=lambda: None,
    )
    return str(getattr(result, "token", current_token) or "")
