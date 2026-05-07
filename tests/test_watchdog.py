"""
Tests for the excel watchdog (src/ package).

Win32 APIs, psutil, and time.sleep are mocked throughout — these tests run on
any OS without a real Excel process.
"""

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
        tables = session.execute(
            mods.database.NotificationLog.__table__.metadata.sorted_tables[0].select().limit(0)
        )
        assert mods.database.NotificationLog.__tablename__ == "notification_log"
        session.close()

    def test_not_notified_initially(self, mods):
        session = mods.database.init_db()
        assert mods.database.already_notified_today(session, "Some Dialog") is False
        session.close()

    def test_record_then_detected(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(session, "Some Dialog", pid=1234, runtime_minutes=5.0, reason="test")
        assert mods.database.already_notified_today(session, "Some Dialog") is True
        session.close()

    def test_duplicate_insert_ignored(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(session, "Dialog A", pid=1, runtime_minutes=1.0, reason="r")
        mods.database.record_notification(session, "Dialog A", pid=2, runtime_minutes=2.0, reason="r2")
        count = (
            session.query(mods.database.NotificationLog)
            .filter_by(window_title="Dialog A")
            .count()
        )
        assert count == 1
        session.close()

    def test_history_returns_recent_rows(self, mods):
        session = mods.database.init_db()
        mods.database.record_notification(session, "Old Dialog", pid=99, runtime_minutes=0.0, reason="old")
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
        result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
        assert result is False
        assert "not configured" in caplog.text

    def test_posts_when_url_set(self, mods):
        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 200

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

        assert result is True

    def test_returns_false_on_url_error(self, mods):
        import urllib.error
        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"

        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("timeout")):
            result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

        assert result is False

    def test_returns_false_on_non_200(self, mods):
        mods.config.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 500

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = mods.notification.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

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
        with patch("src.monitor.get_windows_for_pid", return_value=[]), \
             patch("src.monitor.time.sleep"):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is False

    def test_error_dialog_detected(self, mods):
        proc = self._make_proc(cpu=0.0, runtime_minutes=5)
        windows = [(1, "Run-time error '1004'")]
        with patch("src.monitor.get_windows_for_pid", return_value=windows), \
             patch("src.monitor.time.sleep"):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Error dialog" in result["reason"]

    def test_hung_window_detected(self, mods):
        proc = self._make_proc(cpu=0.0, runtime_minutes=5)
        windows = [(42, "Microsoft Excel")]
        with patch("src.monitor.get_windows_for_pid", return_value=windows), \
             patch("src.monitor.is_hung_window", return_value=True), \
             patch("src.monitor.time.sleep"):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Not Responding" in result["reason"]

    def test_timeout_with_idle_cpu(self, mods):
        mods.config.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=0.0, runtime_minutes=60)
        with patch("src.monitor.get_windows_for_pid", return_value=[]), \
             patch("src.monitor.time.sleep"):
            result = mods.monitor.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "timeout" in result["dialog_title"]

    def test_long_running_but_active_cpu_not_stuck(self, mods):
        mods.config.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=50.0, runtime_minutes=60)
        with patch("src.monitor.get_windows_for_pid", return_value=[]), \
             patch("src.monitor.time.sleep"):
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
