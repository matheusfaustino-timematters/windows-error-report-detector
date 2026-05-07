from unittest.mock import MagicMock, patch

import pytest


# ---------------------------------------------------------------------------
# is_error_dialog
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("title,expected", [
    ("Report.xlsx - Microsoft Excel", False),
    ("Microsoft Excel",               False),
    ("excel",                         False),
    ("data.xls - Overflow",           False),
    ("Run-time error '1004'",         True),
    ("Compile Error in Hidden Module",True),
    ("Microsoft Visual Basic",        True),
    ("Overflow",                      True),
    ("AUTOMATION ERROR",              True),   # case-insensitive
])
def test_is_error_dialog(mods, title, expected):
    assert mods.windows.is_error_dialog(title) is expected


# ---------------------------------------------------------------------------
# is_hung_window
# ---------------------------------------------------------------------------

def test_is_hung_window_returns_false_when_no_user32(mods):
    with patch.object(mods.windows, "user32", None):
        assert mods.windows.is_hung_window(42) is False


def test_is_hung_window_returns_true_from_api(mods):
    mock_user32 = MagicMock()
    mock_user32.IsHungAppWindow.return_value = 1
    with patch.object(mods.windows, "user32", mock_user32):
        assert mods.windows.is_hung_window(42) is True
    mock_user32.IsHungAppWindow.assert_called_once_with(42)


def test_is_hung_window_returns_false_from_api(mods):
    mock_user32 = MagicMock()
    mock_user32.IsHungAppWindow.return_value = 0
    with patch.object(mods.windows, "user32", mock_user32):
        assert mods.windows.is_hung_window(42) is False


# ---------------------------------------------------------------------------
# get_windows_for_pid
# ---------------------------------------------------------------------------

def test_get_windows_for_pid_returns_empty_without_win32(mods):
    # HAS_WIN32 is False in all tests (win32gui suppressed by no_win32 fixture)
    assert mods.windows.get_windows_for_pid(1234) == []
