import logging
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

CONFIG = {
    "max_runtime_minutes":    int(os.getenv("MAX_RUNTIME_MINUTES", "30")),
    "check_interval_seconds": int(os.getenv("CHECK_INTERVAL_SECONDS", "20")),
    "data_dir":               os.getenv("DATA_DIR", r"C:\ExcelWatchdog"),
    "teams_webhook_url":      os.getenv("TEAMS_WEBHOOK_URL", ""),
    # Email (SMTP) — all empty by default; set in .env to enable
    "smtp_host":     os.getenv("SMTP_HOST", ""),
    "smtp_port":     int(os.getenv("SMTP_PORT", "587")),
    "smtp_user":     os.getenv("SMTP_USER", ""),
    "smtp_password": os.getenv("SMTP_PASSWORD", ""),
    "email_from":    os.getenv("EMAIL_FROM", ""),
    "email_to":      os.getenv("EMAIL_TO", ""),   # comma-separated for multiple recipients
}

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
    "method",  # e.g. "Method 'Range' of object '_Worksheet' failed"
]

DATA_DIR = Path(CONFIG["data_dir"])
DATA_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = DATA_DIR / "watchdog.log"
DB_PATH  = DATA_DIR / "watchdog.db"

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
