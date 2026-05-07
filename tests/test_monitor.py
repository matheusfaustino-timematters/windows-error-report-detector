from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

import psutil
import pytest


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def make_proc():
    """Factory fixture — returns a psutil.Process mock with configurable cpu/runtime."""
    def _factory(pid=100, cpu=0.0, runtime_minutes=5):
        proc = MagicMock()
        proc.pid = pid
        proc.create_time.return_value = (
            datetime.now() - timedelta(minutes=runtime_minutes)
        ).timestamp()
        proc.cpu_percent.return_value = cpu
        return proc
    return _factory


@pytest.fixture
def excel_proc():
    """A single mock EXCEL.EXE process as seen by psutil.process_iter."""
    proc = MagicMock()
    proc.pid = 100
    proc.info = {"name": "EXCEL.EXE", "pid": 100}
    return proc


def _stuck(title="Run-time error '1004'"):
    return {"is_stuck": True, "dialog_title": title, "reason": "Error dialog", "runtime_min": 5.0}


def _ok():
    return {"is_stuck": False, "dialog_title": "", "reason": "", "runtime_min": 5.0}


# ---------------------------------------------------------------------------
# analyze_process
# ---------------------------------------------------------------------------

def test_ok_process_not_stuck(mods, make_proc):
    proc = make_proc(cpu=10.0, runtime_minutes=5)
    with patch("src.monitor.get_windows_for_pid", return_value=[]), \
         patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is False


def test_error_dialog_detected(mods, make_proc):
    proc = make_proc(cpu=0.0, runtime_minutes=5)
    with patch("src.monitor.get_windows_for_pid", return_value=[(1, "Run-time error '1004'")]), \
         patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is True
    assert "Error dialog" in result["reason"]


def test_hung_window_detected(mods, make_proc):
    proc = make_proc(cpu=0.0, runtime_minutes=5)
    with patch("src.monitor.get_windows_for_pid", return_value=[(42, "Microsoft Excel")]), \
         patch("src.monitor.is_hung_window", return_value=True), \
         patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is True
    assert "Not Responding" in result["reason"]


def test_timeout_with_idle_cpu(mods, make_proc):
    mods.config.CONFIG["max_runtime_minutes"] = 30
    proc = make_proc(cpu=0.0, runtime_minutes=60)
    with patch("src.monitor.get_windows_for_pid", return_value=[]), \
         patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is True
    assert "timeout" in result["dialog_title"]


def test_long_running_active_cpu_not_stuck(mods, make_proc):
    mods.config.CONFIG["max_runtime_minutes"] = 30
    proc = make_proc(cpu=50.0, runtime_minutes=60)
    with patch("src.monitor.get_windows_for_pid", return_value=[]), \
         patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is False


def test_access_denied_returns_not_stuck(mods):
    proc = MagicMock()
    proc.pid = 999
    proc.cpu_percent.side_effect = psutil.AccessDenied(999)
    with patch("src.monitor.time.sleep"):
        result = mods.monitor.analyze_process(proc)
    assert result["is_stuck"] is False


# ---------------------------------------------------------------------------
# run()
# ---------------------------------------------------------------------------

def test_run_once_no_excel_processes(mods):
    with patch("src.monitor.psutil.process_iter", return_value=[]), \
         patch("src.monitor.time.sleep"):
        mods.monitor.run(once=True)


def test_run_once_ok_process_no_notification(mods, excel_proc):
    with patch("src.monitor.psutil.process_iter", return_value=[excel_proc]), \
         patch("src.monitor.analyze_process", return_value=_ok()), \
         patch("src.monitor.send_teams") as mock_teams, \
         patch("src.monitor.send_email") as mock_email, \
         patch("src.monitor.time.sleep"):
        mods.monitor.run(once=True)
    mock_teams.assert_not_called()
    mock_email.assert_not_called()


def test_run_once_stuck_sends_and_records(mods, excel_proc):
    with patch("src.monitor.psutil.process_iter", return_value=[excel_proc]), \
         patch("src.monitor.analyze_process", return_value=_stuck()), \
         patch("src.monitor.send_teams", return_value=True), \
         patch("src.monitor.send_email", return_value=False), \
         patch("src.monitor.time.sleep"):
        mods.monitor.run(once=True)
    session = mods.database.init_db()
    assert mods.database.already_notified_today(session, "Run-time error '1004'") is True
    session.close()


def test_run_once_already_notified_skips_send(mods, excel_proc):
    title = "Run-time error '1004'"
    session = mods.database.init_db()
    mods.database.record_notification(session, title, pid=99, runtime_minutes=1.0, reason="pre")
    session.close()

    with patch("src.monitor.psutil.process_iter", return_value=[excel_proc]), \
         patch("src.monitor.analyze_process", return_value=_stuck(title)), \
         patch("src.monitor.send_teams") as mock_teams, \
         patch("src.monitor.send_email") as mock_email, \
         patch("src.monitor.time.sleep"):
        mods.monitor.run(once=True)
    mock_teams.assert_not_called()
    mock_email.assert_not_called()


def test_run_once_both_channels_fail_skips_record(mods, excel_proc):
    title = "Overflow"
    with patch("src.monitor.psutil.process_iter", return_value=[excel_proc]), \
         patch("src.monitor.analyze_process", return_value=_stuck(title)), \
         patch("src.monitor.send_teams", return_value=False), \
         patch("src.monitor.send_email", return_value=False), \
         patch("src.monitor.time.sleep"):
        mods.monitor.run(once=True)
    session = mods.database.init_db()
    assert mods.database.already_notified_today(session, title) is False
    session.close()
