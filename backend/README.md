# API Presence

This directory contains the version-controlled FastAPI service used by the Presence interface.

## Deployment

The production service reads its credentials from `/etc/presence-app/presence-app.env`, via the systemd drop-in for `presence-app.service`. Do not commit that file or any password, SMTP credential, database credential, or signing secret.

Before replacing the deployed API, take both a source backup and a MariaDB dump. At startup, the API performs only additive schema migrations: it adds optional organization columns to `users` and creates the group tables if they do not already exist. It does not delete existing user accounts or `Presence` records.

The expected command is:

```text
uvicorn main:app --host 127.0.0.1 --port 8001
```

The deployed service currently uses `/var/www/presence-app/main.py`; deployment copies this tracked source file to that path after validation.
