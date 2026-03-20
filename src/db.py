from pathlib import Path
from sqlalchemy import text
from sqlmodel import SQLModel, Session, create_engine

DB_PATH = Path("data/hr.db")
DATABASE_URL = f"sqlite:///{DB_PATH.as_posix()}"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    SQLModel.metadata.create_all(engine)
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
            ("dob", "DATE"),
            ("doa", "DATE"),
            ("do_report", "DATE"),
            ("status", "TEXT"),
            ("working_at", "TEXT"),
            ("gradation", "TEXT"),
            ("cli", "TEXT"),
            ("pme_due", "DATE"),
            ("technical_due", "DATE"),
            ("transportation_due", "DATE"),
        ]:
            if col not in names:
                conn.execute(text(f"ALTER TABLE employee ADD COLUMN {col} {ddl};"))
        conn.execute(text("UPDATE employee SET status = 'ACTIVE' WHERE status IS NULL OR TRIM(status) = '';"))
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
                    dob DATE,
                    doa DATE,
                    do_report DATE,
                    status TEXT,
                    working_at TEXT,
                    gradation TEXT,
                    cli TEXT,
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
                    seniority_rank, category, pf_no, hrms, dob, doa, do_report, status, working_at,
                    gradation, cli, pme_due, technical_due, transportation_due
                )
                SELECT
                    id, name, role, hire_date,
                    CASE WHEN retirement_date = '2026-03-18' THEN NULL ELSE retirement_date END,
                    promotion_role, promotion_ready_date, seniority_rank, category, pf_no, hrms,
                    dob, doa, do_report, status, working_at, gradation, cli, pme_due,
                    technical_due, transportation_due
                FROM employee_old;
                """
            ))
            conn.execute(text("DROP TABLE employee_old;"))


def get_session() -> Session:
    with Session(engine) as session:
        yield session
