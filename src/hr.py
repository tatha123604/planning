import argparse
import csv
import json
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional


@dataclass
class Employee:
    emp_id: str
    name: str
    role: str
    hire_date: date
    retirement_date: date
    promotion_role: Optional[str]
    promotion_ready_date: Optional[date]

    @staticmethod
    def from_row(row: Dict[str, str]) -> "Employee":
        def parse_d(d: str) -> date:
            return date.fromisoformat(d) if d else None

        return Employee(
            emp_id=row["id"],
            name=row["name"],
            role=row["role"],
            hire_date=date.fromisoformat(row["hire_date"]),
            retirement_date=date.fromisoformat(row["retirement_date"]),
            promotion_role=row.get("promotion_role") or None,
            promotion_ready_date=parse_d(row.get("promotion_ready_date", "")),
        )


def load_requirements(path: Path) -> Dict[str, int]:
    with path.open() as fh:
        return json.load(fh)


def load_employees(path: Path) -> List[Employee]:
    with path.open() as fh:
        return [Employee.from_row(row) for row in csv.DictReader(fh)]


def apply_promotions(employees: List[Employee], as_of: date) -> List[Employee]:
    updated: List[Employee] = []
    for e in employees:
        if e.promotion_role and e.promotion_ready_date and e.promotion_ready_date <= as_of:
            updated.append(
                Employee(
                    emp_id=e.emp_id,
                    name=e.name,
                    role=e.promotion_role,
                    hire_date=e.hire_date,
                    retirement_date=e.retirement_date,
                    promotion_role=None,
                    promotion_ready_date=None,
                )
            )
        else:
            updated.append(e)
    return updated


def headcount_by_role(employees: List[Employee]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for e in employees:
        counts[e.role] = counts.get(e.role, 0) + 1
    return counts


def project_retirements(employees: List[Employee], start: date, horizon_months: int) -> Dict[str, List[Employee]]:
    horizon_end = start + timedelta(days=30 * horizon_months)
    retiring: Dict[str, List[Employee]] = {}
    for e in employees:
        if start < e.retirement_date <= horizon_end:
            retiring.setdefault(e.role, []).append(e)
    return retiring


def build_recruit_plan(
    requirements: Dict[str, int],
    employees: List[Employee],
    as_of: date,
    horizon_months: int,
    lead_time_days: int,
) -> Dict[str, List[str]]:
    plan: Dict[str, List[str]] = {}
    current_counts = headcount_by_role(employees)
    retiring = project_retirements(employees, as_of, horizon_months)

    for role, required in requirements.items():
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


def build_promotion_plan(employees: List[Employee], as_of: date, horizon_months: int) -> List[str]:
    horizon_end = as_of + timedelta(days=30 * horizon_months)
    moves: List[str] = []
    for e in employees:
        if e.promotion_role and e.promotion_ready_date:
            if as_of < e.promotion_ready_date <= horizon_end:
                moves.append(
                    f"Promote {e.name} ({e.role}) to {e.promotion_role} on {e.promotion_ready_date.isoformat()}"
                )
    return moves


def retire_past(employees: List[Employee], as_of: date) -> List[Employee]:
    return [e for e in employees if e.retirement_date > as_of]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="HR headcount & recruiting planner")
    parser.add_argument("--as-of", default=date.today().isoformat(), help="Date to plan from (YYYY-MM-DD)")
    parser.add_argument("--horizon-months", type=int, default=12, help="Planning window in months")
    parser.add_argument("--lead-time-days", type=int, default=90, help="Recruit lead time in days")
    parser.add_argument("--requirements", type=Path, default=Path("config/requirements.json"))
    parser.add_argument("--employees", type=Path, default=Path("data/employees.csv"))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    as_of = date.fromisoformat(args.as_of)

    requirements = load_requirements(args.requirements)
    employees = load_employees(args.employees)

    # remove already retired and apply promotions effective as_of
    active = retire_past(employees, as_of)
    active = apply_promotions(active, as_of)

    counts = headcount_by_role(active)
    recruit_plan = build_recruit_plan(requirements, active, as_of, args.horizon_months, args.lead_time_days)
    promotion_plan = build_promotion_plan(active, as_of, args.horizon_months)

    print(f"Planning as of {as_of.isoformat()} with horizon {args.horizon_months} months\n")
    print("Current headcount vs requirement:")
    for role, required in requirements.items():
        have = counts.get(role, 0)
        delta = have - required
        status = "OK" if delta >= 0 else "SHORT" if delta < 0 else "EXCESS"
        print(f"- {role}: have {have}, need {required} ({status}, delta {delta})")

    print("\nRecruiting actions:")
    if not recruit_plan:
        print("- None; already staffed within horizon")
    else:
        for role, steps in recruit_plan.items():
            print(f"- {role}:")
            for step in steps:
                print(f"  * {step}")

    print("\nPromotions to schedule:")
    if not promotion_plan:
        print("- None in window")
    else:
        for move in promotion_plan:
            print(f"- {move}")


if __name__ == "__main__":
    main()