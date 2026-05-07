import sys
import types

import pytest


@pytest.fixture(autouse=True)
def no_win32(monkeypatch):
    monkeypatch.setitem(sys.modules, "win32gui", None)
    monkeypatch.setitem(sys.modules, "win32process", None)


def _clear_src(sys_modules):
    for key in [k for k in sys_modules if k.startswith("src") or k == "main"]:
        sys_modules.pop(key)


@pytest.fixture()
def mods(tmp_path, monkeypatch):
    """Re-import the entire src package with DATA_DIR redirected to tmp_path."""
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
