import os
from pathlib import Path
from sqlalchemy import text
from sqlalchemy.exc import OperationalError
from sqlmodel import SQLModel, Session, create_engine

BASE_PATH = Path(__file__).resolve().parent.parent
DB_PATH = Path(
    os.getenv("DATABASE_PATH")
    or os.getenv("DB_PATH")
    or ("/app/data/hr.db" if Path("/app/data").exists() else str(BASE_PATH / "data" / "hr.db"))
)
DATABASE_URL = f"sqlite:///{DB_PATH.as_posix()}"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    try:
        SQLModel.metadata.create_all(engine)
    except OperationalError as exc:
        if "already exists" not in str(exc).lower():
            raise
    # lightweight migrations: add missing columns and relax retirement_date to allow NULL
    with engine.begin() as conn:
        cols = conn.execute(text("PRAGMA table_info(employee);")).fetchall()
        names = {c[1] for c in cols}
        if "seniority_rank" not in names:
            conn.execute(text("ALTER TABLE employee ADD COLUMN seniority_rank INTEGER;"))
        for col, ddl in [
            ("category", "TEXT"),
            ("pf_no", "TEXT"),
            ("hrms", "TEXT"),
            ("crew_id", "TEXT"),
            ("dob", "DATE"),
            ("doa", "DATE"),
            ("do_report", "DATE"),
            ("status", "TEXT"),
            ("working_at", "TEXT"),
            ("gradation", "TEXT"),
            ("grading_due", "DATE"),
            ("cli", "TEXT"),
            ("cli_id", "TEXT"),
            ("pme_due", "DATE"),
            ("technical_due", "DATE"),
            ("transportation_due", "DATE"),
        ]:
            if col not in names:
                conn.execute(text(f"ALTER TABLE employee ADD COLUMN {col} {ddl};"))
        conn.execute(text(
            """
            CREATE TABLE IF NOT EXISTS cli_bio_reference (
                cli_id TEXT PRIMARY KEY,
                cli_name TEXT NOT NULL,
                mobile_no TEXT,
                gradation TEXT NOT NULL DEFAULT '0',
                source_file TEXT,
                updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        ))
        bio_cols = conn.execute(text("PRAGMA table_info(cli_bio_reference);")).fetchall()
        bio_names = {c[1] for c in bio_cols}
        if "mobile_no" not in bio_names:
            conn.execute(text("ALTER TABLE cli_bio_reference ADD COLUMN mobile_no TEXT;"))
        conn.execute(text(
            """
            UPDATE employee
            SET role = 'LPS/SHT'
            WHERE role IN ('LPS', 'SHT', 'SHUNTER', 'LPS(SHUNTER)');
            """
        ))
        conn.execute(text(
            """
            UPDATE employee
            SET promotion_role = 'LPS/SHT'
            WHERE promotion_role IN ('LPS', 'SHT', 'SHUNTER', 'LPS(SHUNTER)');
            """
        ))
        req_total = conn.execute(
            text(
                """
                SELECT COALESCE(SUM(needed), 0)
                FROM requirement
                WHERE role IN ('LPS/SHT', 'LPS', 'SHT', 'SHUNTER', 'LPS(SHUNTER)');
                """
            )
        ).scalar()
        if req_total:
            conn.execute(
                text(
                    """
                    DELETE FROM requirement
                    WHERE role IN ('LPS/SHT', 'LPS', 'SHT', 'SHUNTER', 'LPS(SHUNTER)');
                    """
                )
            )
            conn.execute(
                text("INSERT INTO requirement (role, needed) VALUES ('LPS/SHT', :needed);"),
                {"needed": int(req_total)},
            )
        conn.execute(text("UPDATE employee SET status = 'ACTIVE' WHERE status IS NULL OR TRIM(status) = '';"))
        bio_table_exists = conn.execute(
            text("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cli_bio_reference';")
        ).fetchone()
        if bio_table_exists:
            conn.execute(
                text(
                    """
                    DELETE FROM cli_bio_reference
                    WHERE UPPER(TRIM(cli_name)) = 'J S BASAK'
                       OR UPPER(TRIM(cli_id)) = 'SDAH0049';
                    """
                )
            )
        # If retirement_date is NOT NULL, rebuild table to allow NULL and normalize placeholder date
        retirement_col = next((c for c in cols if c[1] == "retirement_date"), None)
        retirement_notnull = retirement_col and retirement_col[3] == 1
        if retirement_notnull:
            conn.execute(text("DROP TABLE IF EXISTS employee_old;"))
            conn.execute(text("ALTER TABLE employee RENAME TO employee_old;"))
            conn.execute(text(
                """
                CREATE TABLE employee (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    role TEXT NOT NULL,
                    hire_date DATE NOT NULL,
                    retirement_date DATE,
                    promotion_role TEXT,
                    promotion_ready_date DATE,
                    seniority_rank INTEGER,
                    category TEXT,
                    pf_no TEXT,
                    hrms TEXT,
                    crew_id TEXT,
                    dob DATE,
                    doa DATE,
                    do_report DATE,
                    status TEXT,
                    working_at TEXT,
                    gradation TEXT,
                    grading_due DATE,
                    cli TEXT,
                    cli_id TEXT,
                    pme_due DATE,
                    technical_due DATE,
                    transportation_due DATE
                );
                """
            ))
            conn.execute(text(
                """
                INSERT INTO employee (
                    id, name, role, hire_date, retirement_date, promotion_role, promotion_ready_date,
                    seniority_rank, category, pf_no, hrms, crew_id, dob, doa, do_report, status, working_at,
                    gradation, grading_due, cli, cli_id, pme_due, technical_due, transportation_due
                )
                SELECT
                    id, name, role, hire_date,
                    CASE WHEN retirement_date = '2026-03-18' THEN NULL ELSE retirement_date END,
                    promotion_role, promotion_ready_date, seniority_rank, category, pf_no, hrms, crew_id,
                    dob, doa, do_report, status, working_at, gradation, grading_due, cli, cli_id, pme_due,
                    technical_due, transportation_due
                FROM employee_old;
                """
            ))
            conn.execute(text("DROP TABLE employee_old;"))
        ssts_table_exists = conn.execute(
            text("SELECT 1 FROM sqlite_master WHERE type='table' AND name='sstsdevicesnapshot';")
        ).fetchone()
        if ssts_table_exists:
            conn.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS ix_sstsdevicesnapshot_run_id_device_id "
                    "ON sstsdevicesnapshot (run_id, device_id);"
                )
            )
            conn.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS ix_sstsdevicesnapshot_observed_day_offline "
                    "ON sstsdevicesnapshot (observed_day, offline_minutes);"
                )
            )
            ssts_cols = conn.execute(text("PRAGMA table_info(sstsdevicesnapshot);")).fetchall()
            ssts_names = {c[1] for c in ssts_cols}
            if ssts_cols and "remark" not in ssts_names:
                conn.execute(text("ALTER TABLE sstsdevicesnapshot ADD COLUMN remark TEXT;"))
        pf_history_table_exists = conn.execute(
            text("SELECT 1 FROM sqlite_master WHERE type='table' AND name='sstspfcounsellinghistory';")
        ).fetchone()
        if pf_history_table_exists:
            conn.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS ix_sstspfhistory_date_speed "
                    "ON sstspfcounsellinghistory (report_date, pf_enter_speed);"
                )
            )
            conn.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS ix_sstspfhistory_crew_date "
                    "ON sstspfcounsellinghistory (crew_id, crew_name, report_date);"
                )
            )


def get_session() -> Session:
    with Session(engine) as session:
        yield session
