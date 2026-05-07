import smtplib
import urllib.error
from unittest.mock import MagicMock, patch


# ---------------------------------------------------------------------------
# Helpers — build mock context managers used by both channels
# ---------------------------------------------------------------------------

def _http_response(status=200):
    mock = MagicMock()
    mock.__enter__ = MagicMock(return_value=mock)
    mock.__exit__ = MagicMock(return_value=False)
    mock.status = status
    return mock


def _smtp_connection():
    mock = MagicMock()
    mock.__enter__ = MagicMock(return_value=mock)
    mock.__exit__ = MagicMock(return_value=False)
    return mock


# ---------------------------------------------------------------------------
# send_teams
# ---------------------------------------------------------------------------

def test_teams_skips_when_url_not_set(mods, caplog):
    mods.config.CONFIG["teams_webhook_url"] = ""
    result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False
    assert "not configured" in caplog.text


def test_teams_posts_when_url_set(mods):
    mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
    with patch("urllib.request.urlopen", return_value=_http_response(200)):
        result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is True


def test_teams_returns_false_on_url_error(mods):
    mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
    with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("timeout")):
        result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False


def test_teams_returns_false_on_non_200(mods):
    mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
    with patch("urllib.request.urlopen", return_value=_http_response(500)):
        result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False


# ---------------------------------------------------------------------------
# send_email
# ---------------------------------------------------------------------------

def test_email_skips_when_not_configured(mods, caplog):
    mods.config.CONFIG["smtp_host"] = ""
    mods.config.CONFIG["email_to"]  = ""
    result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False
    assert "not configured" in caplog.text


def test_email_skips_when_host_missing(mods):
    mods.config.CONFIG["smtp_host"] = ""
    mods.config.CONFIG["email_to"]  = "ops@example.com"
    result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False


def test_email_skips_when_to_missing(mods):
    mods.config.CONFIG["smtp_host"] = "smtp.example.com"
    mods.config.CONFIG["email_to"]  = ""
    result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False


def test_email_sends_when_configured(mods):
    mods.config.CONFIG.update({
        "smtp_host":     "smtp.example.com",
        "smtp_port":     587,
        "smtp_user":     "user@example.com",
        "smtp_password": "secret",
        "email_from":    "watchdog@example.com",
        "email_to":      "ops@example.com",
    })
    mock = _smtp_connection()
    with patch("smtplib.SMTP", return_value=mock):
        result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is True
    mock.sendmail.assert_called_once()


def test_email_multiple_recipients(mods):
    mods.config.CONFIG.update({
        "smtp_host":     "smtp.example.com",
        "smtp_port":     587,
        "smtp_user":     "",
        "smtp_password": "",
        "email_from":    "watchdog@example.com",
        "email_to":      "a@example.com, b@example.com",
    })
    mock = _smtp_connection()
    with patch("smtplib.SMTP", return_value=mock):
        result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is True
    _, recipients, _ = mock.sendmail.call_args[0]
    assert recipients == ["a@example.com", "b@example.com"]


def test_email_returns_false_on_smtp_error(mods):
    mods.config.CONFIG["smtp_host"] = "smtp.example.com"
    mods.config.CONFIG["email_to"]  = "ops@example.com"
    with patch("smtplib.SMTP", side_effect=smtplib.SMTPException("connection refused")):
        result = mods.notification.send_email("Dialog", pid=1, runtime_min=5.0, reason="test")
    assert result is False
