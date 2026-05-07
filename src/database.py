from datetime import date, datetime, timedelta

from sqlalchemy import Column, Float, Integer, String, UniqueConstraint, create_engine
from sqlalchemy.orm import DeclarativeBase, Session

from src.config import DB_PATH, log


class Base(DeclarativeBase):
    pass


class NotificationLog(Base):
    __tablename__ = "notification_log"
    __table_args__ = (UniqueConstraint("window_title", "notified_date", name="uq_title_date"),)

    id              = Column(Integer, primary_key=True, autoincrement=True)
    window_title    = Column(String, nullable=False)
    excel_pid       = Column(Integer, nullable=False)
    notified_date   = Column(String, nullable=False)
    notified_at     = Column(String, nullable=False)
    runtime_minutes = Column(Float)
    stuck_reason    = Column(String)


def init_db() -> Session:
    engine = create_engine(f"sqlite:///{DB_PATH}")
    Base.metadata.create_all(engine)
    return Session(engine)


def already_notified_today(session: Session, window_title: str) -> bool:
    today = date.today().isoformat()
    return (
        session.query(NotificationLog)
        .filter_by(window_title=window_title, notified_date=today)
        .first()
    ) is not None


def record_notification(
    session: Session,
    window_title: str,
    pid: int,
    runtime_minutes: float,
    reason: str,
) -> None:
    today = date.today().isoformat()
    existing = (
        session.query(NotificationLog)
        .filter_by(window_title=window_title, notified_date=today)
        .first()
    )
    if existing:
        return
    try:
        session.add(
            NotificationLog(
                window_title=window_title,
                excel_pid=pid,
                notified_date=today,
                notified_at=datetime.now().isoformat(timespec="seconds"),
                runtime_minutes=runtime_minutes,
                stuck_reason=reason,
            )
        )
        session.commit()
    except Exception as e:
        log.error("DB write failed: %s", e)
        session.rollback()


def get_notification_history(session: Session, days: int = 7) -> list:
    cutoff = (date.today() - timedelta(days=days)).isoformat()
    rows = (
        session.query(NotificationLog)
        .filter(NotificationLog.notified_date >= cutoff)
        .order_by(NotificationLog.notified_at.desc())
        .all()
    )
    return [
        {
            "window_title":    r.window_title,
            "excel_pid":       r.excel_pid,
            "notified_date":   r.notified_date,
            "notified_at":     r.notified_at,
            "runtime_minutes": r.runtime_minutes,
            "stuck_reason":    r.stuck_reason,
        }
        for r in rows
    ]
