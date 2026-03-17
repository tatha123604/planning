from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, Iterable, List

from sqlmodel import Session, select

from .models import Employee, Requirement

ROLE_ORDER = ["LPM", "Motorman", "LPP", "LPG", "LPS", "ALP"]


def role_sort_key(role: str) -> int:
    try:
        return ROLE_ORDER.index(role)
    except ValueError:
        return len(ROLE_ORDER)


def normalize_role(role: str | None) -> str | None:
    if role is None:
        return None
    role = role.strip()
    if role in {"Shunter", "LPS(Shunter)"}:
        return "LPS"
    return role


def is_superior(current_role: str | None, target_role: str | None) -> bool:
    if not current_role or not target_role:
        return False
    return role_sort_key(target_role) < role_sort_key(current_role)


def fetch_active_employees(session: Session, as_of: date) -> List[Employee]:
    employees = session.exec(select(Employee)).all()
    return [e for e in employees if e.retirement_date > as_of]


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
        if start < e.retirement_date <= horizon_end:
            retiring.setdefault(e.role, []).append(e)
    return retiring


def build_recruit_plan(requirements: Dict[str, int], employees: List[Employee], as_of: date, horizon_months: int, lead_time_days: int) -> Dict[str, List[str]]:
    plan: Dict[str, List[str]] = {}
    current_counts = headcount_by_role(employees)
    retiring = project_retirements(employees, as_of, horizon_months)

    for role, required in sorted(requirements.items(), key=lambda kv: role_sort_key(kv[0])):
        current = current_counts.get(role, 0)
        shortage_now = max(required - current, 0)
        steps: List[str] = []
        if shortage_now:
            steps.append(f"Hire {shortage_now} now to meet current requirement {required} (have {current}).")

        for r in retiring.get(role, []):
            target_date = r.retirement_date - timedelta(days=lead_time_days)
            steps.append(
                f"Hire 1 by {target_date.isoformat()} to backfill {r.name} retiring on {r.retirement_date.isoformat()}."
            )

        if steps:
            plan[role] = steps
    return plan


def build_promotion_plan(employees: Iterable[Employee], as_of: date, horizon_months: int) -> List[str]:
    horizon_end = as_of + timedelta(days=30 * horizon_months)
    moves: List[str] = []
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
                f"Promote {e.name} ({e.role}{rank_txt}) to {e.promotion_role} on {e.promotion_ready_date.isoformat()}"
            )
    return moves


def load_requirements_map(session: Session) -> Dict[str, int]:
    reqs = session.exec(select(Requirement)).all()
    return {r.role: r.needed for r in reqs}