import argparse
import sys

from src.database import get_notification_history, init_db
from src.monitor import run


def main() -> None:
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
        session = init_db()
        rows = get_notification_history(session)
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
        session.close()
        sys.exit(0)

    run(once=args.once)


if __name__ == "__main__":
    main()
