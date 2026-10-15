"""Convert saved fiche dates between JSON and Streamlit widgets."""
from datetime import date


def serialize_dates(data):
    """Convertir date en string"""
    if isinstance(data, dict):
        return {k: serialize_dates(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [serialize_dates(v) for v in data]
    elif isinstance(data, date):
        return data.isoformat()
    return data


def deserialize_dates(data):
    """Convertir string en date"""
    if isinstance(data, dict):
        return {k: deserialize_dates(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [deserialize_dates(v) for v in data]
    elif isinstance(data, str):
        try:
            return date.fromisoformat(data)
        except ValueError:
            return data
    return data

