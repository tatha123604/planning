from pathlib import Path
from sqlalchemy import text
from sqlmodel import SQLModel, Session, create_engine

DB_PATH = Path("data/hr.db")
DATABASE_URL = f"sqlite:///{DB_PATH.as_posix()}"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    SQLModel.metadata.create_all(engine)
    # lightweight migration: add seniority_rank if missing
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


def get_session() -> Session:
    with Session(engine) as session:
        yield session
