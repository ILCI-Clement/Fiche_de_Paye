# API Presence

This directory contains the version-controlled FastAPI service used by the Presence interface.

## Deployment

The production service reads its credentials from `/etc/presence-app/presence-app.env`, via the systemd drop-in for `presence-app.service`. Do not commit that file or any password, SMTP credential, database credential, or signing secret.

## Signature électronique

The attendance-sheet action creates a PDF and sends a one-signer invitation through
ClawShow eSign. The production environment must provide `CLAWSHOW_ESIGN_API_KEY`;
it must only exist in `/etc/presence-app/presence-app.env`. LibreOffice is required
on the API server for the Excel/Word-to-PDF conversion. Generated PDFs are kept in
`PRESENCE_ESIGN_DOCUMENTS_DIR` (default: `/var/lib/presence-app/esign-documents`)
behind an unguessable URL for 24 hours, so ClawShow can retrieve the document.
`PRESENCE_ESIGN_PUBLIC_URL` must be the public FastAPI origin (not the Streamlit
interface URL), for example `https://presence.example.org`.

Before replacing the deployed API, take both a source backup and a MariaDB dump. At startup, the API performs only additive schema migrations: it adds optional organization columns to `users` and creates the group tables if they do not already exist. It does not delete existing user accounts or `Presence` records.

The expected command is:

```text
uvicorn main:app --host 127.0.0.1 --port 8001
```

The deployed service currently uses `/var/www/presence-app/main.py`; deployment copies this tracked source file to that path after validation.
