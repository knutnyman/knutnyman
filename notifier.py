"""Notification backends for successful bookings."""

import logging
import os
import smtplib
from email.message import EmailMessage

log = logging.getLogger(__name__)


def notify(subject: str, body: str) -> None:
    """Send a notification via all configured backends."""
    log.info("[NOTIFY] %s — %s", subject, body)
    print(f"\n{'='*60}\n{subject}\n{body}\n{'='*60}\n")
    _try_email(subject, body)


def _try_email(subject: str, body: str) -> None:
    to = os.getenv("NOTIFY_EMAIL_TO")
    host = os.getenv("NOTIFY_SMTP_HOST")
    user = os.getenv("NOTIFY_SMTP_USER")
    password = os.getenv("NOTIFY_SMTP_PASS")
    port = int(os.getenv("NOTIFY_SMTP_PORT", "587"))

    if not all([to, host, user, password]):
        return  # email not configured — skip silently

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = to
    msg.set_content(body)

    try:
        with smtplib.SMTP(host, port) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(msg)
        log.info("Email notification sent to %s", to)
    except Exception as exc:
        log.warning("Email notification failed: %s", exc)
