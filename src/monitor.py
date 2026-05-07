import time
from datetime import datetime

import psutil

from src.config import CONFIG, DB_PATH, log
from src.database import already_notified_today, init_db, record_notification
from src.notification import send_teams
from src.windows import get_windows_for_pid, is_error_dialog, is_hung_window


def analyze_process(proc: psutil.Process) -> dict:
    result = {
        "is_stuck":     False,
        "reason":       "",
        "dialog_title": "",
        "runtime_min":  0.0,
    }

    try:
        proc.cpu_percent(interval=None)
        create_dt   = datetime.fromtimestamp(proc.create_time())
        runtime_min = (datetime.now() - create_dt).total_seconds() / 60
        result["runtime_min"] = round(runtime_min, 1)

        time.sleep(3)
        cpu      = proc.cpu_percent(interval=None)
        cpu_idle = cpu < 0.5

        windows = get_windows_for_pid(proc.pid)

        for hwnd, title in windows:
            if is_hung_window(hwnd):
                result["is_stuck"]     = True
                result["reason"]       = "Window marked 'Not Responding' by Windows"
                result["dialog_title"] = title or f"PID {proc.pid} main window"
                return result

        for _hwnd, title in windows:
            if is_error_dialog(title):
                result["is_stuck"]     = True
                result["reason"]       = f"Error dialog open: '{title}'"
                result["dialog_title"] = title
                return result

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


def run(once: bool = False) -> None:
    session = init_db()
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

                    if already_notified_today(session, status["dialog_title"]):
                        log.info(
                            "Already notified today for '%s' — will notify again tomorrow if still present",
                            status["dialog_title"],
                        )
                        continue

                    sent = send_teams(
                        window_title=status["dialog_title"],
                        pid=proc.pid,
                        runtime_min=status["runtime_min"],
                        reason=status["reason"],
                    )

                    if sent:
                        record_notification(
                            session,
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

    session.close()
