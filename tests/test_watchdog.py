"""
Tests for excel_watchdog (main.py).

Win32 APIs and psutil are mocked throughout — these tests run on any OS
without a real Excel process.
"""

import importlib
import sqlite3
import sys
import types
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


# ---------------------------------------------------------------------------
# Patch win32 modules before importing main so HAS_WIN32 stays False on CI
# ---------------------------------------------------------------------------
@pytest.fixture(autouse=True)
def no_win32(monkeypatch):
    """Ensure win32gui/win32process are absent so HAS_WIN32 == False."""
    monkeypatch.setitem(sys.modules, "win32gui", None)
    monkeypatch.setitem(sys.modules, "win32process", None)


@pytest.fixture()
def main_mod(tmp_path, monkeypatch):
    """
    Import (or re-import) main with DATA_DIR pointing to a temp folder.
    Re-importing ensures CONFIG picks up the monkeypatched env vars.
    """
    monkeypatch.setenv("DATA_DIR", str(tmp_path))
    monkeypatch.setenv("TEAMS_WEBHOOK_URL", "")
    monkeypatch.setenv("MAX_RUNTIME_MINUTES", "30")
    monkeypatch.setenv("CHECK_INTERVAL_SECONDS", "20")

    # Remove cached module so imports are fresh
    sys.modules.pop("main", None)
    import main as m
    return m


# ===========================================================================
# is_error_dialog
# ===========================================================================
class TestIsErrorDialog:
    def test_normal_workbook_not_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Report.xlsx - Microsoft Excel") is False

    def test_plain_excel_title_not_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Microsoft Excel") is False
        assert main_mod.is_error_dialog("excel") is False

    def test_vba_runtime_error_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Run-time error '1004'") is True

    def test_compile_error_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Compile Error in Hidden Module") is True

    def test_microsoft_visual_basic_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Microsoft Visual Basic") is True

    def test_overflow_flagged(self, main_mod):
        assert main_mod.is_error_dialog("Overflow") is True

    def test_case_insensitive(self, main_mod):
        assert main_mod.is_error_dialog("AUTOMATION ERROR") is True

    def test_xls_extension_not_flagged(self, main_mod):
        # A window whose title includes an .xls filename should not be flagged
        assert main_mod.is_error_dialog("data.xls - Overflow") is False


# ===========================================================================
# SQLite helpers
# ===========================================================================
class TestDatabase:
    def test_init_db_creates_table(self, main_mod):
        conn = main_mod.init_db()
        tables = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
        assert any("notification_log" in t for t in tables)
        conn.close()

    def test_not_notified_initially(self, main_mod):
        conn = main_mod.init_db()
        assert main_mod.already_notified_today(conn, "Some Dialog") is False
        conn.close()

    def test_record_then_detected(self, main_mod):
        conn = main_mod.init_db()
        main_mod.record_notification(conn, "Some Dialog", pid=1234, runtime_minutes=5.0, reason="test")
        assert main_mod.already_notified_today(conn, "Some Dialog") is True
        conn.close()

    def test_duplicate_insert_ignored(self, main_mod):
        conn = main_mod.init_db()
        main_mod.record_notification(conn, "Dialog A", pid=1, runtime_minutes=1.0, reason="r")
        main_mod.record_notification(conn, "Dialog A", pid=2, runtime_minutes=2.0, reason="r2")
        rows = conn.execute("SELECT COUNT(*) FROM notification_log WHERE window_title='Dialog A'").fetchone()
        assert rows[0] == 1  # INSERT OR IGNORE — second insert is a no-op
        conn.close()

    def test_history_returns_recent_rows(self, main_mod):
        conn = main_mod.init_db()
        main_mod.record_notification(conn, "Old Dialog", pid=99, runtime_minutes=0.0, reason="old")
        rows = main_mod.get_notification_history(conn, days=7)
        assert len(rows) == 1
        assert rows[0]["window_title"] == "Old Dialog"
        conn.close()

    def test_history_excludes_old_rows(self, main_mod):
        conn = main_mod.init_db()
        # Insert a row with a date 10 days ago
        old_date = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
        conn.execute(
            "INSERT INTO notification_log (window_title, excel_pid, notified_date, notified_at, runtime_minutes, stuck_reason) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("Ancient Dialog", 0, old_date, old_date + "T00:00:00", 0.0, "old"),
        )
        conn.commit()
        rows = main_mod.get_notification_history(conn, days=7)
        assert all(r["window_title"] != "Ancient Dialog" for r in rows)
        conn.close()


# ===========================================================================
# Teams notification
# ===========================================================================
class TestSendTeams:
    def test_skips_when_url_not_set(self, main_mod, caplog):
        main_mod.CONFIG["teams_webhook_url"] = ""
        result = main_mod.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")
        assert result is False
        assert "not configured" in caplog.text

    def test_posts_when_url_set(self, main_mod):
        main_mod.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 200

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = main_mod.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

        assert result is True

    def test_returns_false_on_url_error(self, main_mod):
        import urllib.error
        main_mod.CONFIG["teams_webhook_url"] = "https://example.com/webhook"

        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("timeout")):
            result = main_mod.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

        assert result is False

    def test_returns_false_on_non_200(self, main_mod):
        main_mod.CONFIG["teams_webhook_url"] = "https://example.com/webhook"
        mock_resp = MagicMock()
        mock_resp.__enter__ = MagicMock(return_value=mock_resp)
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_resp.status = 500

        with patch("urllib.request.urlopen", return_value=mock_resp):
            result = main_mod.send_teams("Dialog", pid=1, runtime_min=5.0, reason="test")

        assert result is False


# ===========================================================================
# analyze_process
# ===========================================================================
class TestAnalyzeProcess:
    def _make_proc(self, pid=100, cpu=0.0, create_offset_minutes=-5):
        """Build a mock psutil.Process."""
        proc = MagicMock()
        proc.pid = pid
        create_time = (datetime.now() - timedelta(minutes=abs(create_offset_minutes))).timestamp()
        proc.create_time.return_value = create_time
        proc.cpu_percent.return_value = cpu
        return proc

    def test_ok_process_not_stuck(self, main_mod):
        proc = self._make_proc(cpu=10.0, create_offset_minutes=-5)
        with patch.object(main_mod, "get_windows_for_pid", return_value=[]):
            result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is False

    def test_error_dialog_detected(self, main_mod):
        proc = self._make_proc(cpu=0.0, create_offset_minutes=-5)
        windows = [(1, "Run-time error '1004'")]
        with patch.object(main_mod, "get_windows_for_pid", return_value=windows):
            result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Error dialog" in result["reason"]

    def test_hung_window_detected(self, main_mod):
        proc = self._make_proc(cpu=0.0, create_offset_minutes=-5)
        windows = [(42, "Microsoft Excel")]
        with patch.object(main_mod, "get_windows_for_pid", return_value=windows), \
             patch.object(main_mod, "is_hung_window", return_value=True):
            result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "Not Responding" in result["reason"]

    def test_timeout_with_idle_cpu(self, main_mod):
        main_mod.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=0.0, create_offset_minutes=-60)  # 60 min old
        with patch.object(main_mod, "get_windows_for_pid", return_value=[]):
            result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is True
        assert "timeout" in result["dialog_title"]

    def test_long_running_but_active_cpu_not_stuck(self, main_mod):
        main_mod.CONFIG["max_runtime_minutes"] = 30
        proc = self._make_proc(cpu=50.0, create_offset_minutes=-60)
        with patch.object(main_mod, "get_windows_for_pid", return_value=[]):
            result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is False

    def test_access_denied_returns_not_stuck(self, main_mod):
        import psutil
        proc = MagicMock()
        proc.pid = 999
        proc.cpu_percent.side_effect = psutil.AccessDenied(999)
        result = main_mod.analyze_process(proc)
        assert result["is_stuck"] is False
