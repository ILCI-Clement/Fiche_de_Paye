"""Public UI feature switches; authorization remains in the business endpoints."""
from collections.abc import Callable
from fastapi import FastAPI


def register_ui_capabilities(app: FastAPI, *, mail_configured: Callable[[], bool]):
    @app.get('/ui-capabilities')
    def capabilities():
        local = bool(getattr(app.state, 'local_snapshot', False))
        return {
            'environment': 'local-snapshot' if local else 'server',
            'account_creation': not local,
            'mail_enabled': not local and bool(mail_configured()),
        }
