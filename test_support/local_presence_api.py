"""Minimal local API used only by the isolated Streamlit test environment.

It deliberately keeps data in memory and never contacts the VPS or MariaDB.
"""

from __future__ import annotations

import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse


USERS = {
    "LocalTester": {
        "password": "local-test",
        "email": "local.tester@example.test",
        "role": "Responsable",
    },
}
CONFIGS: dict[str, dict] = {}


class LocalPresenceHandler(BaseHTTPRequestHandler):
    """Serve only the routes exercised by the local Streamlit preview."""

    def log_message(self, _format: str, *_args: object) -> None:
        return

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length) if length else b"{}"
        try:
            value = json.loads(raw_body.decode("utf-8"))
        except json.JSONDecodeError:
            return {}
        return value if isinstance(value, dict) else {}

    def _write_json(self, status: HTTPStatus, payload: object) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path.startswith("/get-config/"):
            username = unquote(path.removeprefix("/get-config/"))
            self._write_json(HTTPStatus.OK, CONFIGS.get(username, {}))
            return
        if path == "/list-users":
            self._write_json(
                HTTPStatus.OK,
                {"users": [{"username": name, "email": user["email"], "role": user["role"]} for name, user in USERS.items()]},
            )
            return
        self._write_json(HTTPStatus.NOT_FOUND, {"detail": "Unknown local test route"})

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        payload = self._read_json()
        if path == "/login":
            username = str(payload.get("username", ""))
            user = USERS.get(username)
            if user and payload.get("password") == user["password"]:
                self._write_json(HTTPStatus.OK, {"username": username, "email": user["email"], "role": user["role"]})
            else:
                self._write_json(HTTPStatus.UNAUTHORIZED, {"detail": "Invalid local test credentials"})
            return
        if path.startswith("/save-config/"):
            username = unquote(path.removeprefix("/save-config/"))
            CONFIGS[username] = payload
            self._write_json(HTTPStatus.OK, {"message": "Local test data saved"})
            return
        if path == "/forgot-password":
            self._write_json(HTTPStatus.OK, {"message": "Local test request accepted"})
            return
        self._write_json(HTTPStatus.NOT_FOUND, {"detail": "Unknown local test route"})

    def do_PUT(self) -> None:  # noqa: N802
        self._write_json(HTTPStatus.OK, {"message": "Local test profile updated", "username": "LocalTester"})


if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", 8003), LocalPresenceHandler)
    print("Local Presence API listening on http://127.0.0.1:8003", flush=True)
    server.serve_forever()
