# Excel Watchdog

Monitors Excel processes on a Windows Server for stuck error dialogs or hung windows and sends a Microsoft Teams notification when one is detected. The script does **not** kill any process — a human investigates and closes the dialog.

---

## How it works

Excel runs as part of an automated pipeline (BAT → VBS → Excel macros). When a macro triggers an error dialog, Excel silently waits for a human to click OK — blocking the whole pipeline. The watchdog polls every N seconds and fires a Teams alert the first time it detects a problem. SQLite deduplication ensures the same dialog only alerts once per day.

Three checks run in priority order for each `EXCEL.EXE` process:

1. **`IsHungAppWindow()`** — Windows OS marks the window "Not Responding"
2. **Error dialog title match** — window title matches patterns like `run-time error`, `compile error`, `automation error`, etc.
3. **Timeout + idle CPU** — process has been running longer than `MAX_RUNTIME_MINUTES` with < 0.5% CPU

---

## Requirements

- Windows Server 2016 / 2019 / 2022
- Python 3.10+
- [uv](https://docs.astral.sh/uv/) (package manager)

---

## Installation

```powershell
# Clone and enter the project directory
git clone <repo-url>
cd windows-error-report-detection

# Install dependencies
uv sync
```

---

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```powershell
Copy-Item .env.example .env
```

| Variable | Required | Default | Description |
|---|---|---|---|
| `TEAMS_WEBHOOK_URL` | **yes** | — | Teams Incoming Webhook URL |
| `MAX_RUNTIME_MINUTES` | no | `30` | Minutes before an idle Excel is flagged |
| `CHECK_INTERVAL_SECONDS` | no | `20` | How often to poll |
| `DATA_DIR` | no | `C:\ExcelWatchdog` | Folder for `watchdog.log` and `watchdog.db` |

### Getting a Teams webhook URL

**Option A — Incoming Webhook connector (classic)**

1. Teams channel → **...** → **Connectors** → search **Incoming Webhook** → **Configure**
2. Name it, click **Create**, copy the generated URL

**Option B — Power Automate Workflow (newer tenants)**

1. Teams channel → **...** → **Workflows**
2. Search **"Post to a channel when a webhook request is received"** → **Add workflow**
3. Copy the generated URL

Both options produce a URL that works with the same payload format.

---

## Usage

```powershell
# Run continuously (production mode)
uv run python main.py

# Single scan — good for testing your webhook
uv run python main.py --once

# Print the last 7 days of notifications
uv run python main.py --history
```

---

## Running tests

```powershell
uv run pytest
```

Tests mock `win32gui`, `win32process`, `psutil`, and `urllib` — no real Excel process or Teams webhook needed. All tests run on any OS.

---

## Deploying on Windows Server (Task Scheduler)

See [SETUP_GUIDE.md](SETUP_GUIDE.md) for step-by-step instructions including Task Scheduler configuration (GUI and CLI), folder layout, and tuning tips.

Quick command-line registration (run as Administrator):

```bat
schtasks /create ^
  /tn "Excel Watchdog" ^
  /tr "\"C:\Python312\python.exe\" \"C:\ExcelWatchdog\main.py\"" ^
  /sc ONSTART /delay 0002:00 ^
  /ru SYSTEM /rl HIGHEST /f
```

---

## File structure

```
main.py              Watchdog script
.env.example         Environment variable template
SETUP_GUIDE.md       Teams webhook setup + Task Scheduler walkthrough
tests/
  test_watchdog.py   Unit tests (all dependencies mocked)
pyproject.toml       Project metadata and dependencies (uv)
uv.lock              Locked dependency versions
```

Runtime files created automatically (not in repo):

```
watchdog.log         Rotating log output
watchdog.db          SQLite notification dedup state
```

---

## Customising error dialog detection

Add lowercase title fragments to `ERROR_DIALOG_PATTERNS` in `main.py`:

```python
ERROR_DIALOG_PATTERNS = [
    "microsoft visual basic",
    "run-time error",
    "data refresh failed",   # example custom pattern
    ...
]
```

Any visible window whose title matches a pattern (and doesn't contain `.xls`) is treated as a blocking dialog.
