"""Authenticated read requests for the Mon équipe page."""
import streamlit as st
from api_client import api_url, authenticated_headers, response_detail, safe_api_request


def fetch_directory(path):
    response = safe_api_request('GET', f'{api_url()}/{path}', headers=authenticated_headers())
    if response is None:
        return None
    if response.status_code != 200:
        st.error(response_detail(response))
        return None
    return response.json()
