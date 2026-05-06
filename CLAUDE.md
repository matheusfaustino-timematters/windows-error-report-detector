# Excel Watchdog

## What this project does

Monitors Excel processes on a Windows Server machine that run as part of an
automated pipeline: BAT scripts → VBS scripts → Excel macros (data refresh +
export). When Excel gets stuck on an error dialog or stops responding, the
watchdog detects it and sends a Microsoft Teams notification. It does **not**
kill the process — a human is expected to investigate and close the dialog.

SQLite is used to track notifications so the same issue on the same day only
fires one alert, but re-alerts the next day if still present.

---

## Current file structure

```
excel_watchdog.py     Main script — runs as a background process on the server
SETUP_GUIDE.md        Step-by-step: Teams webhook setup + Task Scheduler config
watchdog.log          Created automatically on first run (not in repo)
watchdog.db           SQLite database, created automatically on first run (not in repo)
```

---

## Tech stack

- **Python 3.10+** (no virtual environment set up yet)
- **psutil** — process enumeration and CPU inspection
- **pywin32** (`win32gui`, `win32process`) — Win32 window enumeration and
  `IsHungAppWindow()` detection
- **sqlite3** — built-in, used for notification dedup state
- **urllib** — built-in, used for Teams webhook HTTP POST (no requests library)
- **Target OS:** Windows Server (2016 / 2019 / 2022)

Install dependencies:
```
pip install psutil pywin32
```

---

## How to run

```bash
# Run the watchdog continuously (normal production mode)
python excel_watchdog.py

# Single scan — good for testing
python excel_watchdog.py --once

# Print notification history from the database
python excel_watchdog.py --history
```

---

## Key configuration (top of excel_watchdog.py)

```python
CONFIG = {
    "max_runtime_minutes":    30,    # timeout before a CPU-idle Excel is flagged
    "check_interval_seconds": 20,    # how often to poll
    "data_dir":  r"C:\ExcelWatchdog",
    "teams_webhook_url": "https://...",   # must be set before use
}

ERROR_DIALOG_PATTERNS = [...]  # lowercase title fragments that flag a dialog
```

---

## Detection logic (analyze_process in excel_watchdog.py)

Three checks run in priority order for each Excel process:

1. **IsHungAppWindow()** — Windows OS declares the window not responding
2. **Error dialog title match** — enumerates all visible windows for the PID,
   flags those matching `ERROR_DIALOG_PATTERNS` that aren't normal workbook windows
3. **Timeout + idle CPU** — running longer than `max_runtime_minutes` with <0.5% CPU

---

## SQLite dedup schema

```sql
notification_log (
    id              INTEGER PRIMARY KEY,
    window_title    TEXT,     -- the dialog/window title that triggered
    excel_pid       INTEGER,
    notified_date   TEXT,     -- 'YYYY-MM-DD'
    notified_at     TEXT,     -- full ISO datetime
    runtime_minutes REAL,
    stuck_reason    TEXT
)

UNIQUE INDEX on (window_title, notified_date)
```

Same `(window_title, date)` pair → `INSERT OR IGNORE` → no duplicate Teams post.
Different day → new row → new notification.

---

## Teams notification format

Sends an **Adaptive Card** (not a legacy MessageCard) via POST to the
Incoming Webhook URL. Card contains: dialog title, PID, runtime, reason,
server hostname, timestamp. See `send_teams()` in the script.

---

## What isn't built yet — possible next steps

- [ ] **Dashboard / web UI** — view notification history, live process status
- [ ] **Multiple server support** — run watchdog on several machines, aggregate
      alerts into one Teams channel
- [ ] **Config file** (`watchdog.toml` or `watchdog.ini`) so non-developers
      can change settings without editing Python
- [ ] **Auto-close known safe dialogs** — optionally auto-dismiss specific
      dialog titles (e.g. a "Save?" prompt) without notifying
- [ ] **Teams actionable card** — add a button to the Teams card that marks
      the alert as acknowledged in the DB
- [ ] **Retry / backoff** on Teams webhook failures
- [ ] **Unit tests** — mock `win32gui`, `psutil`, and `sqlite3` to test
      detection logic and dedup without a real Excel process
- [ ] **Installer / setup script** — automate pip install + Task Scheduler
      registration in one `setup.bat`
- [ ] **Log rotation** — `watchdog.log` grows forever today

---

## Decisions already made (don't revisit without good reason)

- **No process killing** — the watchdog only notifies; humans close dialogs
- **No email or Slack** — Teams only
- **No external HTTP library** — `urllib` keeps the dependency list short
- **SQLite over a plain text file** — easier to query, handles concurrent
  writes if the script ever runs multiple threads
- **Dedup key is `(window_title, date)`** — not PID, because PIDs recycle;
  not timestamp, because the same stuck dialog would spam every poll cycle

---

## Notes for Claude Code

- The script is designed to run on Windows Server — Win32 APIs won't be
  available in a Linux/macOS dev environment. Use `--once` with mocked data
  or a Windows VM for end-to-end testing.
- `HAS_WIN32` guards all `win32gui`/`win32process` calls — the script won't
  crash on non-Windows but dialog detection will be disabled.
- When adding features, keep all state in `watchdog.db`. Do not introduce
  additional state files.
- The `ERROR_DIALOG_PATTERNS` list uses **lowercase** strings — comparison is
  always done against `title.lower()`.
