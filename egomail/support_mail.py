from __future__ import annotations

import smtplib
import ssl
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable


@dataclass
class SmtpSettings:
    host: str
    port: int
    security: str  # SSL, STARTTLS, NONE
    username: str
    password: str


def infer_smtp_from_imap(imap_host: str, imap_port: int = 993, imap_ssl: bool = True) -> tuple[str, int, str]:
    """Propone una configurazione SMTP partendo dall'host IMAP.

    La proposta viene sempre mostrata e confermata dall'utente prima dell'invio.
    Per provider non riconosciuti usa una trasformazione prudente imap.* -> smtp.*.
    """
    host = (imap_host or "").strip().lower()
    presets = {
        "imap.gmail.com": ("smtp.gmail.com", 465, "SSL"),
        "imap.googlemail.com": ("smtp.gmail.com", 465, "SSL"),
        "outlook.office365.com": ("smtp.office365.com", 587, "STARTTLS"),
        "imap-mail.outlook.com": ("smtp-mail.outlook.com", 587, "STARTTLS"),
        "imap.mail.yahoo.com": ("smtp.mail.yahoo.com", 465, "SSL"),
        "imap.aol.com": ("smtp.aol.com", 465, "SSL"),
        "imap.icloud.com": ("smtp.mail.me.com", 587, "STARTTLS"),
        "imap.libero.it": ("smtp.libero.it", 465, "SSL"),
        "imap.virgilio.it": ("smtp.virgilio.it", 465, "SSL"),
    }
    if host in presets:
        return presets[host]
    if host.startswith("imap."):
        return "smtp." + host[5:], 465, "SSL"
    if "imap" in host:
        return host.replace("imap", "smtp", 1), 465, "SSL"
    return host, 587, "STARTTLS"


def send_support_email(settings: SmtpSettings, recipient: str, subject: str, body: str, attachments: Iterable[str | Path] | None = None) -> None:
    recipient = recipient.strip()
    if not recipient:
        raise ValueError("Indirizzo email dell'assistenza mancante.")
    if not settings.host.strip():
        raise ValueError("Server SMTP mancante.")
    if not settings.username.strip():
        raise ValueError("Utente/mittente SMTP mancante.")

    msg = EmailMessage()
    msg["From"] = settings.username.strip()
    msg["To"] = recipient
    msg["Subject"] = subject.strip() or "Richiesta assistenza EgoMailExtractor Community"
    msg.set_content(body)

    for attachment in attachments or []:
        path = Path(attachment)
        if path.exists() and path.is_file():
            msg.add_attachment(path.read_bytes(), maintype="application", subtype="zip" if path.suffix.lower() == ".zip" else "octet-stream", filename=path.name)

    security = settings.security.upper().strip()
    context = ssl.create_default_context()
    if security == "SSL":
        with smtplib.SMTP_SSL(settings.host, settings.port, timeout=30, context=context) as server:
            server.login(settings.username, settings.password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(settings.host, settings.port, timeout=30) as server:
            server.ehlo()
            if security == "STARTTLS":
                server.starttls(context=context)
                server.ehlo()
            server.login(settings.username, settings.password)
            server.send_message(msg)
