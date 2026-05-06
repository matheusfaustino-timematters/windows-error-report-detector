# =============================================================================
# excel_watchdog.py
#
# Monitors Excel processes for stuck error dialogs or hung windows.
# Sends a Microsoft Teams notification when one is detected.
# Uses SQLite to avoid duplicate notifications — same issue on the same day
# will only notify once, but WILL notify again the following day.
#
# REQUIREMENTS:
#   pip install psutil pywin32
#   (sqlite3 is built into Python — no install needed)
#
# USAGE:
#   python excel_watchdog.py            # runs forever, polling every N seconds
#   python excel_watchdog.py --once     # single scan then exit (good for testing)
#   python excel_watchdog.py --history  # show last 7 days of notifications
# =============================================================================

import argparse
import ctypes
import json
import logging
import os
import socket
import sqlite3
import sys
import time
import urllib.request
import urllib.error
from datetime import date, datetime
from pathlib import Path

import psutil
from dotenv import load_dotenv

load_dotenv()

try:
    import win32gui
    import win32process
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("WARNING: pywin32 not installed. Run: pip install pywin32")

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
CONFIG = {
    "max_runtime_minutes":    int(os.getenv("MAX_RUNTIME_MINUTES", "30")),
    "check_interval_seconds": int(os.getenv("CHECK_INTERVAL_SECONDS", "20")),
    "data_dir":               os.getenv("DATA_DIR", r"C:\ExcelWatchdog"),
    "teams_webhook_url":      os.getenv("TEAMS_WEBHOOK_URL", ""),
}

# Window title fragments that indicate an Excel error dialog is blocking
ERROR_DIALOG_PATTERNS = [
    "microsoft visual basic",
    "run-time error",
    "compile error",
    "cannot open",
    "cannot find",
    "unable to",
    "object required",
    "subscript out of range",
    "type mismatch",
    "overflow",
    "automation error",
    "method",         # e.g. "Method 'Range' of object '_Worksheet' failed"
]

# =============================================================================

# ── PATHS ─────────────────────────────────────────────────────────────────────
DATA_DIR = Path(CONFIG["data_dir"])
DATA_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = DATA_DIR / "watchdog.log"
DB_PATH  = DATA_DIR / "watchdog.db"

# ── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("ExcelWatchdog")


# ── SQLITE STATE ──────────────────────────────────────────────────────────────
def init_db() -> sqlite3.Connection:
    """Create the database and table if they don't exist yet."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS notification_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            window_title    TEXT    NOT NULL,
            excel_pid       INTEGER NOT NULL,
            notified_date   TEXT    NOT NULL,   -- 'YYYY-MM-DD'
            notified_at     TEXT    NOT NULL,   -- full ISO datetime
            runtime_minutes REAL,
            stuck_reason    TEXT
        )
    """)
    # Unique constraint: one notification per (title, day)
    # So the same dialog on a different day will notify again
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS uq_title_date
        ON notification_log (window_title, notified_date)
    """)
    conn.commit()
    return conn


def already_notified_today(conn: sqlite3.Connection, window_title: str) -> bool:
    today = date.today().isoformat()
    row = conn.execute(
        "SELECT id FROM notification_log WHERE window_title = ? AND notified_date = ?",
        (window_title, today),
    ).fetchone()
    return row is not None


def record_notification(
    conn: sqlite3.Connection,
    window_title: str,
    pid: int,
    runtime_minutes: float,
    reason: str,
) -> None:
    today = date.today().isoformat()
    now   = datetime.now().isoformat(timespec="seconds")
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO notification_log
                (window_title, excel_pid, notified_date, notified_at, runtime_minutes, stuck_reason)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (window_title, pid, today, now, runtime_minutes, reason),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed: %s", e)


def get_notification_history(conn: sqlite3.Connection, days: int = 7) -> list:
    rows = conn.execute(
        """
        SELECT window_title, excel_pid, notified_date, notified_at, runtime_minutes, stuck_reason
        FROM notification_log
        WHERE notified_date >= date('now', ?)
        ORDER BY notified_at DESC
        """,
        (f"-{days} days",),
    ).fetchall()
    return [
        {
            "window_title":    r[0],
            "excel_pid":       r[1],
            "notified_date":   r[2],
            "notified_at":     r[3],
            "runtime_minutes": r[4],
            "stuck_reason":    r[5],
        }
        for r in rows
    ]


# ── TEAMS NOTIFICATION ────────────────────────────────────────────────────────
def send_teams(window_title: str, pid: int, runtime_min: float, reason: str) -> bool:
    """
    POST an Adaptive Card to a Teams channel via Incoming Webhook.
    Returns True if the request succeeded.
    """
    url = CONFIG["teams_webhook_url"]
    if not url:
        log.warning("Teams webhook URL not configured — set TEAMS_WEBHOOK_URL in .env")
        return False

    hostname = socket.gethostname()
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type":    "AdaptiveCard",
                    "version": "1.4",
                    "body": [
                        {
                            "type":   "TextBlock",
                            "text":   "🚨  Excel Watchdog — Stuck Process Detected",
                            "weight": "Bolder",
                            "size":   "Medium",
                            "color":  "Attention",
                            "wrap":   True,
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "Dialog / Window", "value": window_title},
                                {"title": "Excel PID",       "value": str(pid)},
                                {"title": "Runtime",         "value": f"{runtime_min:.1f} minutes"},
                                {"title": "Reason",          "value": reason},
                                {"title": "Server",          "value": hostname},
                                {"title": "Detected at",     "value": now_str},
                            ],
                        },
                        {
                            "type":  "TextBlock",
                            "text":  "Excel is still running. Please connect to the server "
                                     "and close the error dialog manually.",
                            "wrap":  True,
                            "color": "Warning",
                            "size":  "Small",
                        },
                    ],
                },
            }
        ],
    }

    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                log.info("✅ Teams notification sent — PID %d | '%s'", pid, window_title)
                return True
            log.error("Teams returned HTTP %d", resp.status)
    except urllib.error.URLError as e:
        log.error("Teams request failed: %s", e)
    return False


# ── WIN32 WINDOW INSPECTION ───────────────────────────────────────────────────
user32 = ctypes.windll.user32


def is_hung_window(hwnd: int) -> bool:
    return bool(user32.IsHungAppWindow(hwnd))


def get_windows_for_pid(pid: int) -> list:
    """Return [(hwnd, title), ...] for all visible windows owned by the given PID."""
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
    """
    Return True if a window title looks like an Excel error dialog
    rather than a normal workbook window.
    """
    lower = title.lower()
    # Normal Excel workbook windows contain the filename or say just "Microsoft Excel"
    if ".xls" in lower:
        return False
    if lower.strip() in ("microsoft excel", "excel"):
        return False
    return any(p in lower for p in ERROR_DIALOG_PATTERNS)


# ── PROCESS ANALYSIS ──────────────────────────────────────────────────────────
def analyze_process(proc: psutil.Process) -> dict:
    """
    Inspect a single Excel process and return a dict describing its state:
        {
            "is_stuck":     bool,
            "reason":       str,
            "dialog_title": str,   # the specific window title that triggered
            "runtime_min":  float,
        }
    """
    result = {
        "is_stuck":     False,
        "reason":       "",
        "dialog_title": "",
        "runtime_min":  0.0,
    }

    try:
        proc.cpu_percent(interval=None)                        # prime counter
        create_dt  = datetime.fromtimestamp(proc.create_time())
        runtime_min = (datetime.now() - create_dt).total_seconds() / 60
        result["runtime_min"] = round(runtime_min, 1)

        time.sleep(3)                                          # short window
        cpu      = proc.cpu_percent(interval=None)
        cpu_idle = cpu < 0.5

        windows = get_windows_for_pid(proc.pid)

        # ── Check 1: IsHungAppWindow ──────────────────────────────────────────
        for hwnd, title in windows:
            if is_hung_window(hwnd):
                result["is_stuck"]     = True
                result["reason"]       = "Window marked 'Not Responding' by Windows"
                result["dialog_title"] = title or f"PID {proc.pid} main window"
                return result

        # ── Check 2: Error dialog title match ─────────────────────────────────
        for _hwnd, title in windows:
            if is_error_dialog(title):
                result["is_stuck"]     = True
                result["reason"]       = f"Error dialog open: '{title}'"
                result["dialog_title"] = title
                return result

        # ── Check 3: Timeout with no CPU activity ─────────────────────────────
        if runtime_min > CONFIG["max_runtime_minutes"] and cpu_idle:
            result["is_stuck"]     = True
            result["reason"]       = (
                f"Running {runtime_min:.1f} min with ~0% CPU — "
                "likely waiting on a hidden dialog"
            )
            result["dialog_title"] = f"Excel PID {proc.pid} (timeout)"

    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
        log.debug("Could not inspect PID %d: %s", proc.pid, e)

    return result


# ── MAIN MONITOR LOOP ─────────────────────────────────────────────────────────
def run(once: bool = False) -> None:
    conn = init_db()
    log.info(
        "=== Excel Watchdog started | interval=%ds | timeout=%dmin | db=%s ===",
        CONFIG["check_interval_seconds"],
        CONFIG["max_runtime_minutes"],
        DB_PATH,
    )

    while True:
        try:
            excel_procs = [
                p for p in psutil.process_iter(["name", "pid"])
                if p.info["name"] and p.info["name"].upper() == "EXCEL.EXE"
            ]

            if excel_procs:
                log.info("Found %d Excel process(es)", len(excel_procs))
                for proc in excel_procs:
                    status = analyze_process(proc)

                    if not status["is_stuck"]:
                        log.info(
                            "Excel PID %d — OK (%.1f min runtime)",
                            proc.pid, status["runtime_min"],
                        )
                        continue

                    log.warning("STUCK: PID %d | %s", proc.pid, status["reason"])

                    # ── Dedup check: only notify once per dialog title per day ─
                    if already_notified_today(conn, status["dialog_title"]):
                        log.info(
                            "Already notified today for '%s' — will notify again tomorrow if still present",
                            status["dialog_title"],
                        )
                        continue

                    # ── Send Teams alert ──────────────────────────────────────
                    sent = send_teams(
                        window_title=status["dialog_title"],
                        pid=proc.pid,
                        runtime_min=status["runtime_min"],
                        reason=status["reason"],
                    )

                    if sent:
                        record_notification(
                            conn,
                            window_title=status["dialog_title"],
                            pid=proc.pid,
                            runtime_minutes=status["runtime_min"],
                            reason=status["reason"],
                        )
            else:
                log.debug("No Excel processes running")

        except Exception as e:
            log.error("Watchdog error: %s", e, exc_info=True)

        if once:
            break

        time.sleep(CONFIG["check_interval_seconds"])

    conn.close()


# ── CLI ───────────────────────────────────────────────────────────────────────
def main() -> None:
    _cli()


def _cli() -> None:
    parser = argparse.ArgumentParser(
        description="Excel Watchdog — detects stuck Excel processes and notifies Teams"
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run a single scan then exit (useful for testing your Teams webhook)",
    )
    parser.add_argument(
        "--history",
        action="store_true",
        help="Print the last 7 days of notifications stored in the database",
    )
    args = parser.parse_args()

    if args.history:
        conn = init_db()
        rows = get_notification_history(conn)
        if not rows:
            print("No notifications in the last 7 days.")
        else:
            header = f"{'Date':<12} {'Time':<20} {'PID':<7} {'Window Title':<45} Reason"
            print(header)
            print("-" * len(header))
            for r in rows:
                print(
                    f"{r['notified_date']:<12} "
                    f"{r['notified_at']:<20} "
                    f"{r['excel_pid']:<7} "
                    f"{str(r['window_title'])[:43]:<45} "
                    f"{r['stuck_reason']}"
                )
        conn.close()
        sys.exit(0)

    run(once=args.once)


if __name__ == "__main__":
    main()
