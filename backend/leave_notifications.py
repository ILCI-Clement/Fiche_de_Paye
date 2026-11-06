"""Transactional leave-mail outbox. No approval tokens or sensitive notes in mail."""
import logging
import os
import smtplib
import ssl
import threading
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from html import escape
from urllib.parse import urlsplit, urlunsplit

from pydantic import EmailStr, TypeAdapter, ValidationError
from .leave_store import LeaveStore, timestamp
from .leave_rules import can_review
from .person_visibility import archived_usernames
from .person_names import person_display_name

logger = logging.getLogger(__name__)
RESULT_TITLES = {
    'approved': 'Votre demande de congé est approuvée',
    'rejected': 'Votre demande de congé est refusée',
    'cancel_approved': 'Votre congé est annulé',
    'cancel_rejected': 'Annulation refusée : votre congé reste approuvé',
}


def ensure_notification_schema(connect):
    with LeaveStore(connect, write=True) as store:
        indexes = '' if store.sqlite else ', KEY notification_due (status,next_attempt_at,event_id)'
        engine = '' if store.sqlite else ' ENGINE=InnoDB DEFAULT CHARSET=utf8mb4'
        store.execute('''CREATE TABLE IF NOT EXISTS leave_notification_outbox (
            event_id BIGINT PRIMARY KEY,request_id BIGINT NOT NULL,kind VARCHAR(24) NOT NULL,
            status VARCHAR(24) NOT NULL DEFAULT 'pending',attempts INTEGER NOT NULL DEFAULT 0,
            next_attempt_at VARCHAR(40) NOT NULL,recipient_username VARCHAR(255) NULL,
            created_at VARCHAR(40) NOT NULL,updated_at VARCHAR(40) NOT NULL''' + indexes + ')' + engine)


def review_url(app_url, request_id):
    url = urlsplit(app_url)
    if url.scheme != 'https' or not url.hostname or url.username or url.password:
        raise ValueError('Public HTTPS application URL required')
    return urlunsplit((url.scheme, url.netloc, url.path or '/', f'leave={int(request_id)}', ''))


def mail_message(row, name, app_url, sender, recipient):
    link = review_url(app_url, row['id'])
    result = row['kind'] in RESULT_TITLES
    title = RESULT_TITLES[row['kind']] if result else ('Annulation de congé à valider'
        if row['kind'] == 'cancel_requested' else 'Demande de congé à valider')
    slots = {'full': 'Journée entière', 'morning': 'Matin', 'afternoon': 'Après-midi'}
    def endpoint(day, slot):
        value = day.isoformat() if hasattr(day, 'isoformat') else str(day)
        return f"{'/'.join(reversed(value[:10].split('-')))} · {slots[slot]}"
    period = endpoint(row['start_date'], row['start_slot'])
    if str(row['start_date']) != str(row['end_date']):
        period += ' → ' + endpoint(row['end_date'], row['end_slot'])
    message = EmailMessage()
    message['Subject'], message['From'], message['To'] = title, sender, recipient
    comment = str(row.get('decision_comment') or '').strip() if result else ''
    decided = str(row.get('decision_at') or '') if result else ''
    details = (f'\nDécision du {decided}' if decided else '') + (f'\nCommentaire : {comment}' if comment else '')
    footer = 'Connexion requise.' if result else 'Connexion requise. La décision se prend dans RH.'
    button = 'Voir les détails' if result else 'Voir et traiter la demande'
    message.set_content(f'{title}\n\n{name}\n{period}{details}\n\nVoir la demande : {link}\n{footer}')
    message.add_alternative(f'<p>{escape(title)}</p><p><strong>{escape(name)}</strong><br>{escape(period)}</p>'
        + (f'<p>Décision du {escape(decided)}</p>' if decided else '')
        + (f'<p>Commentaire : {escape(comment).replace(chr(10), "<br>")}</p>' if comment else '')
        + f'<p><a href="{escape(link, quote=True)}" style="display:inline-block;padding:10px 16px;background:#285ec0;color:white;border-radius:6px;text-decoration:none">{button}</a></p>'
        + f'<p>{footer}</p>', subtype='html')
    return message


def smtp_deliver(message):
    """Retry connection failures, but never automatically repeat uncertain DATA sends."""
    sending = accepted = False
    try:
        server, port = os.environ['PRESENCE_SMTP_SERVER'], int(os.environ['PRESENCE_SMTP_PORT'])
        context = ssl.create_default_context()
        transport = (smtplib.SMTP_SSL(server, port, timeout=10, context=context) if port == 465
                     else smtplib.SMTP(server, port, timeout=10))
        with transport as smtp:
            if port != 465:
                smtp.ehlo(); smtp.starttls(context=context); smtp.ehlo()
            smtp.login(os.environ['PRESENCE_SMTP_USER'], os.environ['PRESENCE_SMTP_PASSWORD'])
            sending = True
            refused = smtp.send_message(message)
            accepted = not bool(refused)
            # A QUIT failure after successful acceptance must not trigger another mail.
            sending = False
            return 'failed' if refused else 'sent'
    except smtplib.SMTPRecipientsRefused:
        return 'failed'
    except (OSError, smtplib.SMTPException, ValueError, KeyError):
        if accepted: return 'sent'
        return 'uncertain' if sending else 'retry'


class LeaveNotifications:
    def __init__(self, connect, *, deliver=smtp_deliver):
        self.connect, self.deliver = connect, deliver

    def enqueue(self, store, request_id, kind):
        if kind not in {'submitted','cancel_requested',*RESULT_TITLES}:
            raise ValueError('Unsupported leave notification event')
        # Called immediately after inserting the audit event, in the same transaction.
        event_id = store.cursor.lastrowid
        now = timestamp()
        store.execute('''INSERT INTO leave_notification_outbox
            (event_id,request_id,kind,next_attempt_at,created_at,updated_at) VALUES(%s,%s,%s,%s,%s,%s)''',
            (event_id, request_id, kind, now, now, now))

    def dispatch_one(self):
        now = timestamp()
        with LeaveStore(self.connect, write=True) as store:
            # A process crash after claiming is deliberately never automatically resent.
            stale = (datetime.now(timezone.utc) - timedelta(minutes=10)).isoformat(timespec='microseconds')
            store.execute("UPDATE leave_notification_outbox SET status='uncertain',updated_at=%s WHERE status='sending' AND updated_at<%s", (now, stale))
            suffix = '' if store.sqlite else ' FOR UPDATE'
            store.execute("SELECT * FROM leave_notification_outbox WHERE status IN ('pending','retry','no_recipient') AND next_attempt_at<=%s ORDER BY event_id LIMIT 1" + suffix, (now,))
            job = store.one()
            if not job: return False
            row = store.request(job['request_id'])
            result = job['kind'] in RESULT_TITLES
            event = None
            if result:
                store.execute('SELECT action,new_status,comment,created_at FROM leave_events WHERE id=%s AND request_id=%s', (job['event_id'],job['request_id']))
                event = store.one()
            valid_result = event and ((event['action'] == 'decided' and event['new_status'] == job['kind'])
                or event['action'] == job['kind'] and job['kind'] in {'cancel_approved','cancel_rejected'})
            actionable = row and (valid_result if result else not row.get('is_archived') and (
                row['status'] == 'pending' if job['kind'] == 'submitted' else
                row['status'] == 'approved' and row.get('cancellation_status') == 'pending'))
            if not actionable:
                store.execute("UPDATE leave_notification_outbox SET status='obsolete',updated_at=%s WHERE event_id=%s", (now,job['event_id']))
                return True
            recipient_id = row['employee_username'] if result else row.get('current_manager')
            store.execute('SELECT username,email,role_tags,is_active,login_deleted_at FROM users WHERE username=%s', (recipient_id,))
            recipient_user = store.one()
            import json
            if recipient_user: recipient_user['role_tags'] = json.loads(recipient_user['role_tags'] or '[]')
            eligible = (recipient_user and recipient_user['is_active'] and not recipient_user['login_deleted_at']
                and recipient_user['username'] not in archived_usernames(store.cursor)
                and (result or recipient_user['username'] != row['employee_username']
                    and can_review(recipient_user,row['employee_username'],row.get('current_manager'))))
            try: recipient = str(TypeAdapter(EmailStr).validate_python(recipient_user['email'])) if eligible else None
            except ValidationError: recipient = None
            if not recipient:
                later = (datetime.now(timezone.utc) + timedelta(hours=1)).isoformat(timespec='microseconds')
                store.execute("UPDATE leave_notification_outbox SET status='no_recipient',next_attempt_at=%s,updated_at=%s WHERE event_id=%s", (later,now,job['event_id']))
                return True
            try:
                message = mail_message({**row,'kind':job['kind'],
                    'decision_comment':event['comment'] if event else '', 'decision_at':event['created_at'] if event else ''}, person_display_name(
                    {'username':row['employee_username'],'email':row.get('current_email')},store.person_names()),
                    os.environ['PRESENCE_APP_URL'],os.environ['PRESENCE_SMTP_USER'],recipient)
            except (ValueError, KeyError):
                store.execute("UPDATE leave_notification_outbox SET status='failed',updated_at=%s WHERE event_id=%s", (now,job['event_id']))
                return True
            store.execute("UPDATE leave_notification_outbox SET status='sending',recipient_username=%s,attempts=attempts+1,updated_at=%s WHERE event_id=%s", (recipient_user['username'],now,job['event_id']))
        outcome = self.deliver(message)
        if outcome not in ('sent','retry','uncertain','failed'): outcome = 'uncertain'
        if outcome == 'retry' and job['attempts'] >= 4: outcome = 'failed'
        later = (datetime.now(timezone.utc) + timedelta(minutes=2 ** min(job['attempts'],4))).isoformat(timespec='microseconds')
        with LeaveStore(self.connect, write=True) as store:
            store.execute('UPDATE leave_notification_outbox SET status=%s,next_attempt_at=%s,updated_at=%s WHERE event_id=%s', (outcome,later,timestamp(),job['event_id']))
        return True

    def register(self, app):
        stop = threading.Event()
        @app.on_event('startup')
        def start():
            ensure_notification_schema(self.connect)
            if os.getenv('PRESENCE_LEAVE_NOTIFICATION_WORKER', '1') != '1': return
            def work():
                while not stop.is_set():
                    try:
                        if self.dispatch_one(): continue
                    except Exception:
                        logger.warning('Leave notification worker failed; inspect outbox status.')
                    stop.wait(30)
            thread = threading.Thread(target=work, daemon=True, name='leave-notifications')
            thread.start()
        @app.on_event('shutdown')
        def shutdown(): stop.set()
