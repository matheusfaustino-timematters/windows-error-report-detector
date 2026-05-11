"""
Tests for the excel watchdog (src/ package).

Win32 APIs, psutil, and time.sleep are mocked throughout — these tests run on
any OS without a real Excel process.
"""

import smtplib
import sys
import types
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

import pytest


# ---------------------------------------------------------------------------
# Suppress win32 modules before any src import so HAS_WIN32 stays False
# ---------------------------------------------------------------------------
@pytest.fixture(autouse=True)
def no_win32(monkeypatch):
    monkeypatch.setitem(sys.modules, "win32gui", None)
    monkeypatch.setitem(sys.modules, "win32process", None)


def _clear_src(sys_modules):
    for key in [k for k in sys_modules if k.startswith("src") or k == "main"]:
        sys_modules.pop(key)


@pytest.fixture()
def mods(tmp_path, monkeypatch):
    """
    Re-import the entire src package with DATA_DIR redirected to tmp_path.
    Clears cached modules first so env vars are picked up fresh each test.
    """
    monkeypatch.setenv("DATA_DIR", str(tmp_path))
    monkeypatch.setenv("TEAMS_WEBHOOK_URL", "")
    monkeypatch.setenv("MAX_RUNTIME_MINUTES", "30")
    monkeypatch.setenv("CHECK_INTERVAL_SECONDS", "20")

    _clear_src(sys.modules)

    import src.config
    import src.database
    import src.notification
    import src.windows
    import src.monitor

    return types.SimpleNamespace(
        config=src.config,
        database=src.database,
        notification=src.notification,
        windows=src.windows,
        monitor=src.monitor,
    )


# ===========================================================================
# src.windows — is_hung_window, get_windows_for_pid
# ===========================================================================
class TestWindowsHelpers:
    def test_is_hung_window_returns_false_when_no_user32(self, mods):
        with patch.object(mods.windows, "user32", None):
            assert mods.windows.is_hung_window(42) is False

    def test_is_hung_window_calls_api_and_returns_true(self, mods):
        mock_user32 = MagicMock()
        mock_user32.IsHungAppWindow.return_value = 1
        with patch.object(mods.windows, "user32", mock_user32):
            result = mods.windows.is_hung_window(42)
        assert result is True
        mock_user32.IsHungAppWindow.assert_called_once_with(42)

    def test_is_hung_window_calls_api_and_returns_false(self, mods):
        mock_user32 = MagicMock()
        mock_user32.IsHungAppWindow.return_value = 0
        with patch.object(mods.windows, "user32", mock_user32):
            assert mods.windows.is_hung_window(42) is False

    def test_get_windows_for_pid_returns_empty_without_win32(self, mods):
        # HAS_WIN32 is False in all tests (win32gui suppressed by no_win32 fixture)
        assert mods.windows.get_windows_for_pid(1234) == []


# ===========================================================================
# src.windows — is_error_dialog
# ===========================================================================
class TestIsErrorDialog:
    def test_normal_workbook_not_flagged(self, mods):
        assert mods.windows.is_error_dialog("Report.xlsx - Microsoft Excel") is False

    def test_plain_excel_title_not_flagged(self, mods):
        assert mods.windows.is_error_dialog("Microsoft Excel") is False
        assert mods.windows.is_error_dialog("excel") is False

    def test_vba_runtime_error_flagged(self, mods):
        assert mods.windows.is_error_dialog("Run-time error '1004'") is True

    def test_compile_error_flagged(self, mods):
        assert mods.windows.is_error_dialog("Compile Error in Hidden Module") is True

    def test_microsoft_visual_basic_flagged(self, mods):
        assert mods.windows.is_error_dialog("Microsoft Visual Basic") is True

    def test_overflow_flagged(self, mods):
        assert mods.windows.is_error_dialog("Overflow") is True

    def test_case_insensitive(self, mods):
        assert mods.windows.is_error_dialog("AUTOMATION ERROR") is True

    def test_xls_in_title_not_flagged(self, mods):
        assert mods.windows.is_error_dialog("data.xls - Overflow") is False


# ===========================================================================
# src.database — SQLite helpers
# ===========================================================================
class TestDatabase:
    def test_init_db_creates_table(self, mods):
        session = mods.database.init_db()
        # If the table doesn't exist this query raises; success proves CREATE TABLE ran
        session.query(mods.database.NotificationLog).count()
        session.close()

    def test_not_notified_initially(self, mods):
        session = mods.database.init_db()
        assert mods.database.already_notified_today(session, "Some Dialog") is False
        session.close()

    def test_record_then_detected(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(
            session, "Some Dialog", pid=1234, runtime_minutes=5.0, reason="test"
        )
        assert mods.database.already_notified_today(session, "Some Dialog") is True
        session.close()

    def test_duplicate_insert_ignored(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(
            session, "Dialog A", pid=1, runtime_minutes=1.0, reason="r"
        )
        mods.database.record_notification(
            session, "Dialog A", pid=2, runtime_minutes=2.0, reason="r2"
        )
        count = (
            session.query(mods.database.NotificationLog)
            .filter_by(window_title="Dialog A")
            .count()
        )
        assert count == 1
        session.close()

    def test_history_returns_recent_rows(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(
            session, "Old Dialog", pid=99, runtime_minutes=0.0, reason="old"
        )
        rows = mods.database.get_notification_history(session, days=7)
        assert len(rows) == 1
        assert rows[0]["window_title"] == "Old Dialog"
        session.close()

    def test_history_excludes_old_rows(self, mods):
        session = mods.database.init_db()
        old_date = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
        session.add(
            mods.database.NotificationLog(
                window_title="Ancient Dialog",
                excel_pid=0,
                notified_date=old_date,
                notified_at=old_date + "T00:00:00",
                runtime_minutes=0.0,
                stuck_reason="old",
            )
        )
        session.commit()
        rows = mods.database.get_notification_history(session, days=7)
        assert all(r["window_title"] != "Ancient Dialog" for r in rows)
        session.close()


# ===========================================================================
# src.notification — Teams webhook
# ===========================================================================
class TestSendTeams:
    def test_skips_when_url_not_set(self, mods, caplog):
        mods.config.CONFIG["teams_webhook_url"] = ""
        result = mods.notification.send_teams(
            "Dialog", pid=1, runtime_min=5.0, reason="test"
        )
        assert result is False
        assert "not configured" in caplog.text

    def test_posts_when_url_set(self, mods):
        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 200

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = mods.notification.send_teams(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is True

    def test_returns_false_on_url_error(self, mods):
        import urllib.error

        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"

        with patch(
            "urllib.request.urlopen", side_effect=urllib.error.URLError("timeout")
        ):
            result = mods.notification.send_teams(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is False

    def test_returns_false_on_non_200(self, mods):
        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 500

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = mods.notification.send_teams(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is False


# ===========================================================================
# src.notification — send_email
# ===========================================================================
class TestSendEmail:
    def test_skips_when_not_configured(self, mods, caplog):
        mods.config.CONFIG["smtp_host"] = ""
        mods.config.CONFIG["email_to"] = ""
        result = mods.notification.send_email(
            "Dialog", pid=1, runtime_min=5.0, reason="test"
        )
        assert result is False
        assert "not configured" in caplog.text

    def test_skips_when_host_missing(self, mods, caplog):
        mods.config.CONFIG["smtp_host"] = ""
        mods.config.CONFIG["email_to"] = "ops@example.com"
        result = mods.notification.send_email(
            "Dialog", pid=1, runtime_min=5.0, reason="test"
        )
        assert result is False

    def test_skips_when_to_missing(self, mods, caplog):
        mods.config.CONFIG["smtp_host"] = "smtp.example.com"
        mods.config.CONFIG["email_to"] = ""
        result = mods.notification.send_email(
            "Dialog", pid=1, runtime_min=5.0, reason="test"
        )
        assert result is False

    def test_sends_when_configured(self, mods):
        mods.config.CONFIG["smtp_host"] = "smtp.example.com"
        mods.config.CONFIG["smtp_port"] = 587
        mods.config.CONFIG["smtp_user"] = "user@example.com"
        mods.config.CONFIG["smtp_password"] = "secret"
        mods.config.CONFIG["email_from"] = "watchdog@example.com"
        mods.config.CONFIG["email_to"] = "ops@example.com"

        mock_smtp = MagicMock()
        mock_smtp.__enter__ = MagicMock(return_value=mock_smtp)
        mock_smtp.__exit__ = MagicMock(return_value=False)

        with patch("smtplib.SMTP", return_value=mock_smtp):
            result = mods.notification.send_email(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is True
        mock_smtp.sendmail.assert_called_once()

    def test_multiple_recipients(self, mods):
        mods.config.CONFIG["smtp_host"] = "smtp.example.com"
        mods.config.CONFIG["smtp_port"] = 587
        mods.config.CONFIG["smtp_user"] = ""
        mods.config.CONFIG["smtp_password"] = ""
        mods.config.CONFIG["email_from"] = "watchdog@example.com"
        mods.config.CONFIG["email_to"] = "a@example.com, b@example.com"

        mock_smtp = MagicMock()
        mock_smtp.__enter__ = MagicMock(return_value=mock_smtp)
        mock_smtp.__exit__ = MagicMock(return_value=False)

        with patch("smtplib.SMTP", return_value=mock_smtp):
            result = mods.notification.send_email(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is True
        _, recipients, _ = mock_smtp.sendmail.call_args[0]
        assert recipients == ["a@example.com", "b@example.com"]

    def test_returns_false_on_smtp_error(self, mods):
        mods.config.CONFIG["smtp_host"] = "smtp.example.com"
        mods.config.CONFIG["email_to"] = "ops@example.com"

        with patch(
            "smtplib.SMTP", side_effect=smtplib.SMTPException("connection refused")
        ):
            result = mods.notification.send_email(
                "Dialog", pid=1, runtime_min=5.0, reason="test"
            )

        assert result is False


# ===========================================================================
# src.monitor — analyze_process
# ===========================================================================
class TestAnalyzeProcess:
    def _make_proc(self, pid=100, cpu=0.0, runtime_minutes=5):
        proc = MagicMock()
        proc.pid = pid
        create_time = (datetime.now() - timedelta(minutes=runtime_minutes)).timestamp()
        proc.create_time.return_value = create_time
        proc.cpu_percent.return_value = cpu
        return proc

    def test_ok_process_not_stuck(self, mods):
        proc = self._make_proc(cpu=10.0, runtime_minutes=5)
        with (
            patch("src.monitor.get_windows_for_pid", return_value=[]),
            patch("src.monitor.time.sleep"),
        ):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is False

    def test_error_dialog_detected(self, mods):
        proc = self._make_proc(cpu=0.0, runtime_minutes=5)
        windows = [(1, "Run-time error '1004'")]
        with (
            patch("src.monitor.get_windows_for_pid", return_value=windows),
            patch("src.monitor.time.sleep"),
        ):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Error dialog" in result["reason"]

    def test_hung_window_detected(self, mods):
        proc = self._make_proc(cpu=0.0, runtime_minutes=5)
        windows = [(42, "Microsoft Excel")]
        with (
            patch("src.monitor.get_windows_for_pid", return_value=windows),
            patch("src.monitor.is_hung_window", return_value=True),
            patch("src.monitor.time.sleep"),
        ):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Not Responding" in result["reason"]

    def test_timeout_with_idle_cpu(self, mods):
        mods.config.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=0.0, runtime_minutes=60)
        with (
            patch("src.monitor.get_windows_for_pid", return_value=[]),
            patch("src.monitor.time.sleep"),
        ):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "timeout" in result["dialog_title"]

    def test_long_running_but_active_cpu_not_stuck(self, mods):
        mods.config.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=50.0, runtime_minutes=60)
        with (
            patch("src.monitor.get_windows_for_pid", return_value=[]),
            patch("src.monitor.time.sleep"),
        ):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is False

    def test_access_denied_returns_not_stuck(self, mods):
        import psutil

        proc = MagicMock()
        proc.pid = 999
        proc.cpu_percent.side_effect = psutil.AccessDenied(999)
        with patch("src.monitor.time.sleep"):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is False


# ===========================================================================
# src.monitor — run()
# ===========================================================================
class TestMonitorRun:
    def _excel_proc(self, pid=100):
        proc = MagicMock()
        proc.pid = pid
        proc.info = {"name": "EXCEL.EXE", "pid": pid}
        return proc

    def _stuck(self, title="Run-time error '1004'"):
        return {
            "is_stuck": True,
            "dialog_title": title,
            "reason": "Error dialog",
            "runtime_min": 5.0,
        }

    def _ok(self):
        return {"is_stuck": False, "dialog_title": "", "reason": "", "runtime_min": 5.0}

    def test_run_once_no_excel_processes(self, mods):
        with (
            patch("src.monitor.psutil.process_iter", return_value=[]),
            patch("src.monitor.time.sleep"),
        ):
            mods.monitor.run(once=True)  # must not raise

    def test_run_once_ok_process_no_notification(self, mods):
        proc = self._excel_proc()
        with (
            patch("src.monitor.psutil.process_iter", return_value=[proc]),
            patch("src.monitor.analyze_process", return_value=self._ok()),
            patch("src.monitor.send_teams") as mock_teams,
            patch("src.monitor.send_email") as mock_email,
            patch("src.monitor.time.sleep"),
        ):
            mods.monitor.run(once=True)
        mock_teams.assert_not_called()
        mock_email.assert_not_called()

    def test_run_once_stuck_sends_and_records(self, mods):
        proc = self._excel_proc()
        with (
            patch("src.monitor.psutil.process_iter", return_value=[proc]),
            patch("src.monitor.analyze_process", return_value=self._stuck()),
            patch("src.monitor.send_teams", return_value=True),
            patch("src.monitor.send_email", return_value=False),
            patch("src.monitor.time.sleep"),
        ):
            mods.monitor.run(once=True)
        session = mods.database.init_db()
        assert (
            mods.database.already_notified_today(session, "Run-time error '1004'")
            is True
        )
        session.close()

    def test_run_once_already_notified_skips_send(self, mods):
        proc = self._excel_proc()
        title = "Run-time error '1004'"
        session = mods.database.init_db()
        mods.database.record_notification(
            session, title, pid=99, runtime_minutes=1.0, reason="pre"
        )
        session.close()

        with (
            patch("src.monitor.psutil.process_iter", return_value=[proc]),
            patch("src.monitor.analyze_process", return_value=self._stuck(title)),
            patch("src.monitor.send_teams") as mock_teams,
            patch("src.monitor.send_email") as mock_email,
            patch("src.monitor.time.sleep"),
        ):
            mods.monitor.run(once=True)
        mock_teams.assert_not_called()
        mock_email.assert_not_called()

    def test_run_once_both_channels_fail_skips_record(self, mods):
        proc = self._excel_proc()
        title = "Overflow"
        with (
            patch("src.monitor.psutil.process_iter", return_value=[proc]),
            patch("src.monitor.analyze_process", return_value=self._stuck(title)),
            patch("src.monitor.send_teams", return_value=False),
            patch("src.monitor.send_email", return_value=False),
            patch("src.monitor.time.sleep"),
        ):
            mods.monitor.run(once=True)
        session = mods.database.init_db()
        assert mods.database.already_notified_today(session, title) is False
        session.close()
