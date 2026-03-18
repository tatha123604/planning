import csv
import json
from datetime import date
from pathlib import Path
from typing import Iterable

from sqlmodel import Session, select

from .models import Employee, Requirement
from .logic import normalize_role

CONFIG_REQUIREMENTS = Path("config/requirements.json")
DATA_EMPLOYEES = Path("data/employees.csv")


def _parse_date(value: str) -> date:
    return date.fromisoformat(value) if value else None


def seed_requirements(session: Session) -> None:
    if session.exec(select(Requirement)).first():
        return
    if CONFIG_REQUIREMENTS.exists():
        data = json.loads(CONFIG_REQUIREMENTS.read_text())
        for role, needed in data.items():
            role = normalize_role(role)
            session.add(Requirement(role=role, needed=needed))
    session.commit()


def seed_employees(session: Session) -> None:
    if session.exec(select(Employee)).first():
        return
    if DATA_EMPLOYEES.exists():
        with DATA_EMPLOYEES.open() as fh:
            reader: Iterable[dict] = csv.DictReader(fh)
            for row in reader:
                role = normalize_role(row["role"])
                promo_role = normalize_role(row.get("promotion_role") or None)
                session.add(
                    Employee(
                        name=row["name"],
                        role=role,
                        hire_date=_parse_date(row["hire_date"]),
                        retirement_date=_parse_date(row["retirement_date"]),
                        promotion_role=promo_role,
                        promotion_ready_date=_parse_date(row.get("promotion_ready_date", "")),
                        category=row.get("category") or None,
                        pf_no=row.get("pf_no") or None,
                        hrms=row.get("hrms") or None,
                        dob=_parse_date(row.get("dob", "")),
                        doa=_parse_date(row.get("doa", "")),
                        do_report=_parse_date(row.get("do_report", "")),
                        status=row.get("status") or None,
                        working_at=row.get("working_at") or None,
                    )
                )
    session.commit()


def seed_all(session: Session) -> None:
    seed_requirements(session)
    seed_employees(session)
