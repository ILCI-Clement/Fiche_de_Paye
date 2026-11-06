"""Calendar-aligned 03:00 Europe/Paris schedule, stored as UTC instants."""
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

PARIS = ZoneInfo('Europe/Paris')


def utc_now():
    return datetime.now(timezone.utc).replace(tzinfo=None, microsecond=0)


def next_due(now, interval):
    if interval not in (1, 3, 6, 12):
        raise ValueError('invalid_interval')
    instant = now.replace(tzinfo=timezone.utc) if now.tzinfo is None else now
    local = instant.astimezone(PARIS)
    for year in range(local.year, local.year + 2):
        for month in range(1, 13, interval):
            candidate = datetime(year, month, 1, 3, tzinfo=PARIS)
            if candidate > instant:
                return candidate.astimezone(timezone.utc).replace(tzinfo=None)
    raise ValueError('invalid_schedule')
