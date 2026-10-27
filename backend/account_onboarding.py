"""Initial account credentials and one welcome-mail attempt; no plaintext persistence."""
from dataclasses import dataclass, field
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
import secrets
import smtplib
import ssl
import string
from urllib.parse import urlsplit

from pydantic import EmailStr, SecretStr, TypeAdapter, ValidationError

_EMAIL = TypeAdapter(EmailStr)


@dataclass(frozen=True)
class AccountCredentials:
    email: str
    password: SecretStr = field(repr=False)
    generated: bool


def credentials_for_account(payload):
    raw_email = payload.get('new_mail')
    if not isinstance(raw_email, str) or any(ord(c) < 32 or c in '<>' for c in raw_email):
        raise ValueError('Une adresse e-mail valide est obligatoire.')
    try:
        email = str(_EMAIL.validate_python(raw_email.strip()))
    except ValidationError:
        raise ValueError('Une adresse e-mail valide est obligatoire.') from None
    generated = payload.get('generate_password', False)
    if type(generated) is not bool:
        raise ValueError('Le mode de création du mot de passe est invalide.')
    if generated:
        if payload.get('new_password') not in (None, ''):
            raise ValueError('Choisissez un seul mode de création du mot de passe.')
        alphabet = string.ascii_uppercase + string.digits
        while True:
            password = ''.join(secrets.choice(alphabet) for _ in range(8))
            if any(c in string.ascii_uppercase for c in password) and any(c in string.digits for c in password):
                break
    else:
        password = payload.get('new_password')
        if not isinstance(password, str) or len(password) < 8:
            raise ValueError('Le mot de passe doit contenir au moins 8 caractères et au plus 72 octets.')
        try:
            encoded_length = len(password.encode('utf-8'))
        except UnicodeEncodeError:
            raise ValueError('Le mot de passe contient un caractère invalide.') from None
        if encoded_length > 72:
            raise ValueError('Le mot de passe doit contenir au moins 8 caractères et au plus 72 octets.')
    return AccountCredentials(email, SecretStr(password), generated)


def send_welcome_email(*, email, name, password, generated, app_url, server, port, sender, smtp_password):
    """Return SMTP acceptance, not inbox delivery. Never retry or log credentials."""
    try:
        url = urlsplit(app_url)
        if url.scheme != 'https' or not url.hostname or url.username or url.password:
            return False
        message = MIMEMultipart('alternative')
        message['From'], message['To'] = sender, email
        message['Subject'] = 'Votre compte ILCI · Espace RH'
        reminder = ('Ce mot de passe temporaire est à modifier dans Paramètres après votre connexion.'
                    if generated else 'Vous pouvez modifier ce mot de passe dans Paramètres.')
        message.attach(MIMEText(
            f'Bonjour {name},\n\nVotre compte Espace RH a été créé.\n'
            f'Connexion : {app_url}\nAdresse e-mail : {email}\nMot de passe initial : {password}\n\n'
            f'{reminder}\nNe partagez pas ce message.\n', 'plain', 'utf-8'))
        message.attach(MIMEText(
            f'<p>Bonjour {escape(name)},</p><p>Votre compte Espace RH a été créé.</p>'
            f'<p><a href="{escape(app_url, quote=True)}">Se connecter à Espace RH</a></p>'
            f'<p>Adresse e-mail : {escape(email)}<br>Mot de passe initial : <strong>{escape(password)}</strong></p>'
            f'<p>{escape(reminder)}</p><p>Ne partagez pas ce message.</p>', 'html', 'utf-8'))
        context = ssl.create_default_context()
        if port == 465:
            transport = smtplib.SMTP_SSL(server, port, timeout=10, context=context)
        else:
            transport = smtplib.SMTP(server, port, timeout=10)
        with transport as smtp:
            if port != 465:
                smtp.ehlo()
                smtp.starttls(context=context)
                smtp.ehlo()
            smtp.login(sender, smtp_password)
            return not bool(smtp.sendmail(sender, [email], message.as_string()))
    except (OSError, smtplib.SMTPException, ValueError):
        return False
