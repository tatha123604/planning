import csv
import json
import sqlite3
from datetime import date
from pathlib import Path
from typing import Iterable

from sqlmodel import Session, select

from .models import Employee, Requirement
from .logic import normalize_role

CONFIG_REQUIREMENTS = Path("config/requirements.json")
DATA_EMPLOYEES = Path("data/employees.csv")
BOOTSTRAP_DB = Path("/bootstrap/hr.db")


def _parse_date(value: str) -> date:
    return date.fromisoformat(value) if value else None


def _parse_optional_date(value: object) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


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


def sync_bootstrap(session: Session) -> None:
    """Top up a fresh or undersized Railway volume from the tracked baseline DB."""
    if not BOOTSTRAP_DB.exists():
        return

    employees = session.exec(select(Employee)).all()
    live_count = len(employees)

    bootstrap = sqlite3.connect(BOOTSTRAP_DB)
    bootstrap.row_factory = sqlite3.Row
    try:
        cur = bootstrap.cursor()
        cur.execute("SELECT COUNT(*) FROM employee")
        bootstrap_count = cur.fetchone()[0]
        if live_count >= bootstrap_count:
            return

        by_pf = {e.pf_no: e for e in employees if e.pf_no}
        by_hrms = {e.hrms: e for e in employees if e.hrms}
        by_name_role = {(e.name, e.role): e for e in employees}

        employee_cols = [
            "name",
            "role",
            "hire_date",
            "retirement_date",
            "promotion_role",
            "promotion_ready_date",
            "seniority_rank",
            "category",
            "pf_no",
            "hrms",
            "dob",
            "doa",
            "do_report",
            "status",
            "working_at",
            "gradation",
            "cli",
            "pme_due",
            "technical_due",
            "transportation_due",
        ]

        rows = cur.execute(
            "SELECT name, role, hire_date, retirement_date, promotion_role, promotion_ready_date, "
            "seniority_rank, category, pf_no, hrms, dob, doa, do_report, status, working_at, "
            "gradation, cli, pme_due, technical_due, transportation_due FROM employee"
        ).fetchall()

        for row in rows:
            payload = {
                "name": row["name"],
                "role": row["role"],
                "hire_date": _parse_optional_date(row["hire_date"]),
                "retirement_date": _parse_optional_date(row["retirement_date"]),
                "promotion_role": row["promotion_role"],
                "promotion_ready_date": _parse_optional_date(row["promotion_ready_date"]),
                "seniority_rank": row["seniority_rank"],
                "category": row["category"],
                "pf_no": row["pf_no"],
                "hrms": row["hrms"],
                "dob": _parse_optional_date(row["dob"]),
                "doa": _parse_optional_date(row["doa"]),
                "do_report": _parse_optional_date(row["do_report"]),
                "status": row["status"],
                "working_at": row["working_at"],
                "gradation": row["gradation"],
                "cli": row["cli"],
                "pme_due": _parse_optional_date(row["pme_due"]),
                "technical_due": _parse_optional_date(row["technical_due"]),
                "transportation_due": _parse_optional_date(row["transportation_due"]),
            }

            existing = None
            if payload["pf_no"]:
                existing = by_pf.get(payload["pf_no"])
            if existing is None and payload["hrms"]:
                existing = by_hrms.get(payload["hrms"])
            if existing is None:
                existing = by_name_role.get((payload["name"], payload["role"]))

            if existing is None:
                employee = Employee(**payload)
                session.add(employee)
                if employee.pf_no:
                    by_pf[employee.pf_no] = employee
                if employee.hrms:
                    by_hrms[employee.hrms] = employee
                by_name_role[(employee.name, employee.role)] = employee
                continue

            for col in employee_cols:
                value = payload[col]
                if value in (None, ""):
                    continue
                current = getattr(existing, col)
                if current in (None, ""):
                    setattr(existing, col, value)

        existing_requirements = {r.role: r for r in session.exec(select(Requirement)).all()}
        req_rows = cur.execute("SELECT role, needed FROM requirement").fetchall()
        for row in req_rows:
            requirement = existing_requirements.get(row["role"])
            if requirement is None:
                session.add(Requirement(role=row["role"], needed=row["needed"]))
            elif not requirement.needed and row["needed"]:
                requirement.needed = row["needed"]

        session.commit()
    finally:
        bootstrap.close()


def seed_all(session: Session) -> None:
    seed_requirements(session)
    seed_employees(session)
    sync_bootstrap(session)
