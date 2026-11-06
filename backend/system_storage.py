"""Read-only filesystem capacity, restricted to freshly authorized administrators."""
from datetime import datetime, timezone
from pathlib import Path
import os
import shutil

from fastapi import Depends, HTTPException, Response

from .role_policy import checked_resource_actor, is_admin


def storage_status(paths):
    volumes, unavailable = {}, []
    for label, raw in paths:
        try:
            path = Path(raw).resolve(strict=True)
            device = path.stat().st_dev
            if device in volumes:
                volumes[device]['labels'].append(label)
                continue
            usage = shutil.disk_usage(path)
            percent = round(usage.used / usage.total * 100, 1)
            level = ('critical' if usage.used / usage.total >= .9 or usage.free < 5 * 1024**3
                     else 'warning' if usage.used / usage.total >= .8 else 'normal')
            volumes[device] = dict(labels=[label], total_bytes=usage.total,
                                   used_bytes=usage.used, free_bytes=usage.free,
                                   used_percent=percent, level=level)
        except (OSError, ValueError, ZeroDivisionError):
            unavailable.append(label)
    return dict(volumes=list(volumes.values()), unavailable=unavailable,
                checked_at=datetime.now(timezone.utc).isoformat())


def register_storage_routes(app, *, get_db_connection, get_current_user, load_user, paths):
    @app.get('/system/storage')
    def storage(response: Response, actor=Depends(get_current_user)):
        response.headers['Cache-Control'] = 'no-store'
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user)
                if not is_admin(current):
                    raise HTTPException(403, 'Accès réservé aux administrateurs.')
        finally:
            db.close()
        return storage_status(paths)


def application_storage_paths(main):
    from .medical_visits import medical_storage_dir
    return [
        ('Visites médicales', medical_storage_dir()),
        ('Application', Path(__file__).resolve().parent.parent),
        ('Documents', main.TRANSPORT_RECEIPTS_DIR),
        ('PDF à signer', main.ESIGN_DOCUMENTS_DIR),
        ('PDF signés', os.getenv('PRESENCE_SIGNED_FICHES_DIR', '/var/lib/presence-app/signed-fiches')),
    ]
