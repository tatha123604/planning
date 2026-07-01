from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, Iterable, List

from sqlmodel import Session, select

from .models import Employee, Requirement

ROLE_ORDER = ["LPM", "Motorman", "LPP", "LPG", "LPS/SHT", "ALP"]


def role_sort_key(role: str) -> int:
    try:
        return ROLE_ORDER.index(role)
    except ValueError:
        return len(ROLE_ORDER)


def normalize_role(role: str | None) -> str | None:
    if role is None:
        return None
    role_clean = role.strip()
    upper = role_clean.upper()
    canonical = {
        "SHUNTER": "LPS/SHT",
        "SHT": "LPS/SHT",
        "LPS(SHUNTER)": "LPS/SHT",
        "LPS/SHT": "LPS/SHT",
        "MOTORMAN": "Motorman",
        "MTM": "Motorman",
        "MOTOR MAN": "Motorman",
        "M/MAN": "Motorman",
        "MMAN": "Motorman",
        "LPM": "LPM",
        "LPP": "LPP",
        "LPP(LOCO)": "LPP",
        "LPP/LOCO": "LPP",
        "LPS": "LPS/SHT",
        "ALP": "ALP",
        "SR.ALP": "ALP",
        "LPG": "LPG",
    }
    return canonical.get(upper, role_clean)


def is_employee_active(employee: Employee, as_of: date) -> bool:
    return employee.retirement_date is None or employee.retirement_date > as_of


def is_superior(current_role: str | None, target_role: str | None) -> bool:
    if not current_role or not target_role:
        return False
    return role_sort_key(target_role) < role_sort_key(current_role)


def fetch_active_employees(session: Session, as_of: date) -> List[Employee]:
    employees = session.exec(select(Employee)).all()
    return [e for e in employees if is_employee_active(e, as_of)]


def purge_retired_employees(session: Session, as_of: date) -> int:
    retired = [e for e in session.exec(select(Employee)).all() if not is_employee_active(e, as_of)]
    for employee in retired:
        session.delete(employee)
    if retired:
        session.commit()
    return len(retired)


def apply_promotions(employees: Iterable[Employee], as_of: date) -> List[Employee]:
    promoted: List[Employee] = []
    for e in employees:
        if (
            e.promotion_role
            and e.promotion_ready_date
            and e.promotion_ready_date <= as_of
            and is_superior(e.role, e.promotion_role)
        ):
            clone = Employee(
                id=e.id,
                name=e.name,
                role=e.promotion_role,
                hire_date=e.hire_date,
                retirement_date=e.retirement_date,
                promotion_role=None,
                promotion_ready_date=None,
                seniority_rank=e.seniority_rank,
                category=e.category,
                pf_no=e.pf_no,
                hrms=e.hrms,
                crew_id=e.crew_id,
                dob=e.dob,
                doa=e.doa,
                do_report=e.do_report,
                status=e.status,
                working_at=e.working_at,
            )
            promoted.append(clone)
        else:
            promoted.append(e)
    return promoted


def headcount_by_role(employees: Iterable[Employee]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for e in employees:
        counts[e.role] = counts.get(e.role, 0) + 1
    return counts


def project_retirements(employees: Iterable[Employee], start: date, horizon_months: int) -> Dict[str, List[Employee]]:
    horizon_end = start + timedelta(days=30 * horizon_months)
    retiring: Dict[str, List[Employee]] = {}
    for e in employees:
        if e.retirement_date and start <= e.retirement_date <= horizon_end:
            retiring.setdefault(e.role, []).append(e)
    return retiring


def project_retirements_window(employees: Iterable[Employee], start: date, end: date) -> Dict[str, List[Employee]]:
    """Retirements between start and end dates (inclusive)."""
    if end < start:
        start, end = end, start
    retiring: Dict[str, List[Employee]] = {}
    for e in employees:
        if e.retirement_date and start <= e.retirement_date <= end:
            retiring.setdefault(e.role, []).append(e)
    return retiring


def build_recruit_plan(requirements: Dict[str, int], employees: List[Employee], as_of: date, horizon_months: int, lead_time_days: int) -> Dict[str, List[str]]:
    plan: Dict[str, List[str]] = {}
    current_counts = headcount_by_role(employees)
    retiring = project_retirements(employees, as_of, horizon_months)
    fmt = lambda d: d.strftime("%d-%m-%Y")

    for role, required in sorted(requirements.items(), key=lambda kv: role_sort_key(kv[0])):
        current = current_counts.get(role, 0)
        shortage_now = max(required - current, 0)
        steps: List[str] = []
        if shortage_now:
            steps.append(f"Hire {shortage_now} now to meet current requirement {required} (have {current}).")

        for r in retiring.get(role, []):
            target_date = r.retirement_date - timedelta(days=lead_time_days)
            steps.append(
                f"Hire 1 by {fmt(target_date)} to backfill {r.name} retiring on {fmt(r.retirement_date)}."
            )

        if steps:
            plan[role] = steps
    return plan


def build_promotion_plan(employees: Iterable[Employee], as_of: date, horizon_months: int) -> List[str]:
    horizon_end = as_of + timedelta(days=30 * horizon_months)
    moves: List[str] = []
    fmt = lambda d: d.strftime("%d-%m-%Y")
    def sort_key(e: Employee):
        rank = e.seniority_rank if e.seniority_rank is not None else 10**9
        ready = e.promotion_ready_date if e.promotion_ready_date else date.max
        return (role_sort_key(e.role), rank, ready, e.name)

    for e in sorted(employees, key=sort_key):
        if (
            e.promotion_role
            and e.promotion_ready_date
            and is_superior(e.role, e.promotion_role)
            and as_of < e.promotion_ready_date <= horizon_end
        ):
            rank_txt = f" (rank {e.seniority_rank})" if e.seniority_rank is not None else ""
            moves.append(
                f"Promote {e.name} ({e.role}{rank_txt}) to {e.promotion_role} on {fmt(e.promotion_ready_date)}"
            )
    return moves


def load_requirements_map(session: Session) -> Dict[str, int]:
    reqs = session.exec(select(Requirement)).all()
    return {r.role: r.needed for r in reqs}
