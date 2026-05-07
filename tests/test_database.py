from datetime import datetime, timedelta


def test_init_db_creates_table(mods):
    session = mods.database.init_db()
    # Raises if the table doesn't exist — success proves CREATE TABLE ran
    session.query(mods.database.NotificationLog).count()
    session.close()


def test_not_notified_initially(mods):
    session = mods.database.init_db()
    assert mods.database.already_notified_today(session, "Some Dialog") is False
    session.close()


def test_record_then_detected(mods):
    session = mods.database.init_db()
    mods.database.record_notification(session, "Some Dialog", pid=1234, runtime_minutes=5.0, reason="test")
    assert mods.database.already_notified_today(session, "Some Dialog") is True
    session.close()


def test_duplicate_insert_ignored(mods):
    session = mods.database.init_db()
    mods.database.record_notification(session, "Dialog A", pid=1, runtime_minutes=1.0, reason="r")
    mods.database.record_notification(session, "Dialog A", pid=2, runtime_minutes=2.0, reason="r2")
    count = session.query(mods.database.NotificationLog).filter_by(window_title="Dialog A").count()
    assert count == 1
    session.close()


def test_history_returns_recent_rows(mods):
    session = mods.database.init_db()
    mods.database.record_notification(session, "Old Dialog", pid=99, runtime_minutes=0.0, reason="old")
    rows = mods.database.get_notification_history(session, days=7)
    assert len(rows) == 1
    assert rows[0]["window_title"] == "Old Dialog"
    session.close()


def test_history_excludes_old_rows(mods):
    session = mods.database.init_db()
    old_date = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
    session.add(
        mods.database.NotificationLog(
            window_title="Ancient Dialog",
            excel_pid=0,
            notified_date=old_date,
            notified_at=old_date + "T00:00:00",
            runtime_minutes=0.0,
            stuck_reason="old",
        )
    )
    session.commit()
    rows = mods.database.get_notification_history(session, days=7)
    assert all(r["window_title"] != "Ancient Dialog" for r in rows)
    session.close()
