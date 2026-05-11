import ctypes

from src.config import ERROR_DIALOG_PATTERNS

try:
    import win32gui
    import win32process

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    user32 = ctypes.windll.user32
except AttributeError:
    user32 = None


def is_hung_window(hwnd: int) -> bool:
    if user32 is None:
        return False
    return bool(user32.IsHungAppWindow(hwnd))


def get_windows_for_pid(pid: int) -> list:
    results = []
    if not HAS_WIN32:
        return results

    def _cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid == pid:
                title = win32gui.GetWindowText(hwnd)
                if title:
                    results.append((hwnd, title))
        return True

    win32gui.EnumWindows(_cb, None)
    return results


def is_error_dialog(title: str) -> bool:
    lower = title.lower()
    if ".xls" in lower:
        return False
    if lower.strip() in ("microsoft excel", "excel"):
        return False
    return any(p in lower for p in ERROR_DIALOG_PATTERNS)
