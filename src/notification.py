import json
import smtplib
import socket
import urllib.error
import urllib.request
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from src.config import CONFIG, log


def send_teams(window_title: str, pid: int, runtime_min: float, reason: str) -> bool:
    url = CONFIG["teams_webhook_url"]
    if not url:
        log.warning("Teams webhook URL not configured — set TEAMS_WEBHOOK_URL in .env")
        return False

    hostname = socket.gethostname()
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type":    "AdaptiveCard",
                    "version": "1.4",
                    "body": [
                        {
                            "type":   "TextBlock",
                            "text":   "🚨  Excel Watchdog — Stuck Process Detected",
                            "weight": "Bolder",
                            "size":   "Medium",
                            "color":  "Attention",
                            "wrap":   True,
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "Dialog / Window", "value": window_title},
                                {"title": "Excel PID",       "value": str(pid)},
                                {"title": "Runtime",         "value": f"{runtime_min:.1f} minutes"},
                                {"title": "Reason",          "value": reason},
                                {"title": "Server",          "value": hostname},
                                {"title": "Detected at",     "value": now_str},
                            ],
                        },
                        {
                            "type":  "TextBlock",
                            "text":  "Excel is still running. Please connect to the server "
                                     "and close the error dialog manually.",
                            "wrap":  True,
                            "color": "Warning",
                            "size":  "Small",
                        },
                    ],
                },
            }
        ],
    }

    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                log.info("✅ Teams notification sent — PID %d | '%s'", pid, window_title)
                return True
            log.error("Teams returned HTTP %d", resp.status)
    except urllib.error.URLError as e:
        log.error("Teams request failed: %s", e)
    return False


def send_email(window_title: str, pid: int, runtime_min: float, reason: str) -> bool:
    host     = CONFIG["smtp_host"]
    user     = CONFIG["smtp_user"]
    password = CONFIG["smtp_password"]
    from_    = CONFIG["email_from"] or user
    to_raw   = CONFIG["email_to"]

    if not host or not to_raw:
        log.warning("Email not configured — set SMTP_HOST and EMAIL_TO in .env")
        return False

    recipients = [a.strip() for a in to_raw.split(",") if a.strip()]
    hostname   = socket.gethostname()
    now_str    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject    = f"[Excel Watchdog] Stuck process on {hostname} — {window_title}"

    body = (
        f"Excel Watchdog detected a stuck process.\n\n"
        f"Dialog / Window : {window_title}\n"
        f"Excel PID       : {pid}\n"
        f"Runtime         : {runtime_min:.1f} minutes\n"
        f"Reason          : {reason}\n"
        f"Server          : {hostname}\n"
        f"Detected at     : {now_str}\n\n"
        f"Please connect to the server and close the error dialog manually."
    )

    msg = MIMEMultipart()
    msg["From"]    = from_
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(host, CONFIG["smtp_port"], timeout=15) as smtp:
            smtp.ehlo()
            smtp.starttls()
            if user and password:
                smtp.login(user, password)
            smtp.sendmail(from_, recipients, msg.as_string())
        log.info("Email notification sent — PID %d | '%s'", pid, window_title)
        return True
    except Exception as e:
        log.error("Email send failed: %s", e)
        return False
