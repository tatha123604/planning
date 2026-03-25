from __future__ import annotations

from datetime import date, datetime, timedelta
import json
import math
import os
from io import BytesIO
from pathlib import Path
from typing import Optional
import re
from urllib.parse import quote

from fastapi import Depends, FastAPI, Form, Request, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook, Workbook
import pandas as pd
from sqlmodel import Session, select
from starlette.middleware.base import BaseHTTPMiddleware

from .db import DB_PATH, get_session, init_db
from .logic import (
    ROLE_ORDER,
    apply_promotions,
    build_promotion_plan,
    build_recruit_plan,
    fetch_active_employees,
    headcount_by_role,
    load_requirements_map,
    project_retirements,
    project_retirements_window,
    role_sort_key,
    normalize_role,
)
from .models import (
    CliMatrixOverdueSnapshot,
    CliMatrixSummarySnapshot,
    Employee,
    NonContinuousSignOffSnapshot,
    NonContinuousSignOnSnapshot,
    Requirement,
    SubNonContinuousSignOffSnapshot,
    SubNonContinuousSignOnSnapshot,
)
from non_continuous_duty import (
    build_non_continuous_workbook,
    parse_non_continuous_source,
)
from .seed import seed_all
from processor import (
    build_output_workbook,
    build_sheet2_df,
    build_summary_df,
    coerce_report_date,
    infer_report_date,
    report_date_iso,
)

BASE_PATH = Path(__file__).resolve().parent.parent
TEMPLATE_STORE_DIR = DB_PATH.parent / "saved_templates"
CLI_MATRIX_2026_03_24_CLEANUP_SENTINEL = DB_PATH.parent / ".cli_matrix_cleanup_2026_03_24.done"
GOOGLE_SHEETS_READONLY_SCOPE = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
NON_CONTINUOUS_VARIANTS = {
    "non_sub": {
        "active_page": "non_continuous_duty",
        "page_title": "NON SUB NON CONTINUOUS DUTY",
        "heading_title": "NON SUB NON CONTINUOUS DUTY SIGN_ON/SIGN_OFF",
        "nav_label": "NON SUB NON CONT DUTY",
        "route_base": "/non-continuous-duty",
        "feature_name": "NON SUB NON CONTINUOUS DUTY",
        "sign_on_label": "NON SUB NON CONTINUOUS DUTY SIGN_ON",
        "sign_off_label": "NON SUB NON CONTINUOUS DUTY SIGN_OFF",
        "sheet_title": "NON SUB NON CONT. DUTY",
        "sign_on_model": NonContinuousSignOnSnapshot,
        "sign_off_model": NonContinuousSignOffSnapshot,
    },
    "sub": {
        "active_page": "sub_non_continuous_duty",
        "page_title": "SUB NON CONTINUOUS DUTY",
        "heading_title": "SUB NON CONTINUOUS DUTY SIGN_ON/SIGN_OFF",
        "nav_label": "SUB NON CONT DUTY",
        "route_base": "/sub-non-continuous-duty",
        "feature_name": "SUB NON CONTINUOUS DUTY",
        "sign_on_label": "SUB NON CONTINUOUS DUTY SIGN_ON",
        "sign_off_label": "SUB NON CONTINUOUS DUTY SIGN_OFF",
        "sheet_title": "SUB NON CONT. DUTY",
        "sign_on_model": SubNonContinuousSignOnSnapshot,
        "sign_off_model": SubNonContinuousSignOffSnapshot,
    },
}
templates = Jinja2Templates(directory=str(BASE_PATH / "templates"))
# Jinja filter for dd-mm-yyyy display
def format_dmy(value):
    if not value:
        return ""
    try:
        return value.strftime("%d-%m-%Y")
    except Exception:
        return str(value)
templates.env.filters["dmy"] = format_dmy


def filter_hire_by(value, days: int = 30):
    if not value:
        return None
    try:
        return value - timedelta(days=days)
    except Exception:
        return None
templates.env.filters["hire_by"] = filter_hire_by

ADMIN_USER = "admin"
ADMIN_PASS = "sdah1234"
_AUTH_COOKIE = "session"
_ALLOWED_PATHS = {"/login", "/logout", "/health"}
_ALLOWED_PREFIXES = ("/static", "/openapi.json", "/docs", "/redoc")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        path = request.url.path
        if path in _ALLOWED_PATHS or any(path.startswith(pref) for pref in _ALLOWED_PREFIXES):
            return await call_next(request)
        if request.cookies.get(_AUTH_COOKIE) == "ok":
            return await call_next(request)
        return RedirectResponse(url="/login", status_code=302)

def _parse_as_of(request: Request, as_of: Optional[str]) -> date:
    """Resolve as_of date from query or cookie; fallback to today."""
    if as_of:
        try:
            return date.fromisoformat(as_of)
        except ValueError:
            pass
    # fallback to reports end-date cookie if present (keep pages in sync)
    rep_end = request.cookies.get("reports_end_date")
    if rep_end:
        try:
            return date.fromisoformat(rep_end)
        except ValueError:
            pass
    cookie_val = request.cookies.get("as_of")
    if cookie_val:
        try:
            return date.fromisoformat(cookie_val)
        except ValueError:
            pass
    return date.today()


def _parse_date_cookie(request: Request, key: str, param: Optional[str]) -> date:
    """Resolve date from query param or cookie name=key; fallback to today."""
    if param:
        try:
            return date.fromisoformat(param)
        except ValueError:
            pass
    cookie_val = request.cookies.get(key)
    if cookie_val:
        try:
            return date.fromisoformat(cookie_val)
        except ValueError:
            pass
    return date.today()

app = FastAPI(title="HR Planner")
app.mount("/static", StaticFiles(directory=str(BASE_PATH / "static")), name="static")
app.add_middleware(AuthMiddleware)

# Always serve fresh pages (avoid browser caching dashboards/reports)
@app.middleware("http")
async def add_no_cache_headers(request, call_next):
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store"
    return response


def _extract_date(text: str) -> date:
    """Pull YYYY-MM-DD from a string; return date.max if missing so undated items stay last."""
    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if match:
        try:
            return date.fromisoformat(match.group(0))
        except ValueError:
            pass
    return date.max


def build_simple_recruit_plan(retiring: dict[str, list[Employee]], lead_days: int = 30) -> dict[str, list[str]]:
    """Create backfill steps 1 month before each retirement."""
    plan: dict[str, list[str]] = {}
    for role, people in retiring.items():
        for e in people:
            if not e.retirement_date:
                continue
            hire_by = e.retirement_date - timedelta(days=lead_days)
            step = (
                f"Hire 1 by {hire_by.strftime('%d-%m-%Y')} "
                f"to backfill {e.name} retiring on {e.retirement_date.strftime('%d-%m-%Y')}."
            )
            plan.setdefault(role, []).append(step)
    return plan


def build_cli_distribution(employees: list[Employee]) -> list[dict[str, int | str]]:
    """Aggregate gradation counts per CLI (case-insensitive)."""
    dist: dict[str, dict[str, int | str]] = {}
    for e in employees:
        cli_raw = (e.cli or "").strip()
        cli_key = cli_raw.lower() if cli_raw else "unassigned"
        label = cli_raw or "Unassigned"
        grad = (e.gradation or "").strip().upper()
        grad_key = grad[0] if grad else ""
        if cli_key not in dist:
            dist[cli_key] = {"cli": label, "A": 0, "B": 0, "C": 0, "total": 0}
        # keep the first non-empty label we see for this key
        if not dist[cli_key]["cli"] and cli_raw:
            dist[cli_key]["cli"] = cli_raw
        if grad_key in ("A", "B", "C"):
            dist[cli_key][grad_key] += 1  # type: ignore[index]
        dist[cli_key]["total"] += 1  # type: ignore[index]
    return [
        {"cli": counts["cli"], "A": counts["A"], "B": counts["B"], "C": counts["C"], "total": counts["total"]}
        for _, counts in sorted(dist.items(), key=lambda item: item[0])
    ]


@app.on_event("startup")
def on_startup() -> None:
    init_db()
    session = next(get_session())
    try:
        seed_all(session)
        _run_one_time_cli_matrix_cleanup(session)
    finally:
        session.close()


@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/login")
def login_form(request: Request, error: str | None = None):
    return templates.TemplateResponse(
        "login.html",
        {"request": request, "error": error},
    )


@app.post("/login")
async def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == ADMIN_USER and password == ADMIN_PASS:
        response = RedirectResponse(url="/", status_code=302)
        response.set_cookie(_AUTH_COOKIE, "ok", httponly=True, max_age=86400)
        return response
    return templates.TemplateResponse(
        "login.html",
        {"request": request, "error": "Invalid credentials"},
        status_code=401,
    )


@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie(_AUTH_COOKIE)
    return response


@app.get("/")
def index(
    request: Request,
    as_of: Optional[str] = None,
    horizon_months: int = 12,
    lead_time_days: int = 90,
    session: Session = Depends(get_session),
):
    plan_date = _parse_as_of(request, as_of)
    today = date.today()
    horizon_days = 0
    horizon_months = 0
    lead_time_days = 0

    employees_now = fetch_active_employees(session, today)
    employees_now = apply_promotions(employees_now, today)
    requirements_map = load_requirements_map(session)
    counts = headcount_by_role(employees_now)
    retiring_raw = project_retirements_window(employees_now, today, plan_date)
    recruit_plan = build_simple_recruit_plan(retiring_raw, lead_days=30)
    promotion_plan = build_promotion_plan(employees_now, today, horizon_months)
    target_date = plan_date
    retire_counts = {role: len(peeps) for role, peeps in retiring_raw.items()}

    recruit_plan = {
        role: sorted(steps, key=_extract_date)
        for role, steps in recruit_plan.items()
    }
    retiring = {
        role: sorted(people, key=lambda e: e.retirement_date or date.max)
        for role, people in retiring_raw.items()
    }

    requirements = sorted(session.exec(select(Requirement)).all(), key=lambda r: role_sort_key(r.role))
    employees = sorted(employees_now, key=lambda e: (role_sort_key(e.role), e.name))

    response = templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "as_of": plan_date,
            "horizon_months": horizon_months,
            "lead_time_days": lead_time_days,
            "counts": counts,
            "requirements": requirements,
            "recruit_plan": recruit_plan,
            "promotion_plan": promotion_plan,
            "employees": employees,
            "retiring": {k: retiring[k] for k in sorted(retiring, key=role_sort_key)},
            "retire_counts": retire_counts,
            "target_date": target_date,
            "role_order": ROLE_ORDER,
            "active_page": "dashboard",
        },
    )
    response.set_cookie("as_of", plan_date.isoformat())
    response.set_cookie("reports_end_date", plan_date.isoformat())
    return response


@app.get("/employees")
def employees_page(
    request: Request,
    q: Optional[str] = None,
    role: Optional[str] = None,
    working_at: Optional[str] = None,
    cli: Optional[str] = None,
    gradation: Optional[str] = None,
    sort: str = "role",
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_gradation: Optional[str] = None,
    sync_notice: Optional[str] = None,
    sync_error: Optional[str] = None,
    session: Session = Depends(get_session),
):
    roster_filter_active = any([roster_name, roster_cli, roster_gradation])
    employees_open = not roster_filter_active
    roster_open = roster_filter_active

    employees_all = session.exec(select(Employee)).all()
    raw_working = {e.working_at for e in employees_all if e.working_at}
    working_opts_filtered = {wa for wa in raw_working if wa.upper().startswith("CC(")}
    working_opts = sorted(working_opts_filtered if working_opts_filtered else raw_working)
    cli_opts_map: dict[str, str] = {}
    for val in [e.cli for e in employees_all if e.cli]:
        key = val.strip().lower()
        if key not in cli_opts_map:
            cli_opts_map[key] = val.strip()
    cli_opts = [v for _, v in sorted(cli_opts_map.items(), key=lambda item: item[0])]
    gradation_opts = sorted({e.gradation for e in employees_all if e.gradation})
    employees = list(employees_all)

    if q:
        q_lower = q.lower()
        employees = [
            e
            for e in employees
            if q_lower in e.name.lower()
            or q_lower in e.role.lower()
            or (e.cli and q_lower in e.cli.lower())
            or (e.gradation and q_lower in e.gradation.lower())
        ]
    if role:
        employees = [e for e in employees if e.role == role]
    if working_at:
        wa_lower = working_at.lower()
        employees = [e for e in employees if e.working_at and wa_lower in e.working_at.lower()]
    if cli:
        cli_lower = cli.strip().lower()
        employees = [e for e in employees if e.cli and cli_lower in e.cli.strip().lower()]
    if gradation:
        grad_lower = gradation.lower()
        employees = [e for e in employees if e.gradation and grad_lower in e.gradation.lower()]

    def sort_key(e: Employee):
        if sort == "name":
            return (e.name.lower(),)
        if sort == "retirement":
            return (e.retirement_date or date.max, e.name)
        if sort == "hire":
            return (e.hire_date, e.name)
        if sort == "cli":
            return ((e.cli or "").strip().lower(), e.name)
        if sort == "working_at":
            return ((e.working_at or "").lower(), e.name)
        return (role_sort_key(e.role), e.name)

    employees = sorted(employees, key=sort_key)

    cli_roster = [e for e in employees_all if e.cli]
    if roster_name:
        name_lower = roster_name.lower()
        cli_roster = [e for e in cli_roster if name_lower in e.name.lower()]
    if roster_cli:
        roster_cli_lower = roster_cli.strip().lower()
        cli_roster = [e for e in cli_roster if e.cli and roster_cli_lower in e.cli.strip().lower()]
    if roster_gradation:
        grad_lower = roster_gradation.lower()
        cli_roster = [e for e in cli_roster if e.gradation and grad_lower in e.gradation.lower()]
    cli_roster = sorted(cli_roster, key=lambda e: ((e.cli or "").strip().lower(), e.name))

    return templates.TemplateResponse(
        "employees.html",
        {
            "request": request,
            "employees": employees,
            "role_order": ROLE_ORDER,
            "active_page": "employees",
            "query": q or "",
            "filter_role": role or "",
            "filter_gradation": gradation or "",
            "sort": sort,
            "working_opts": working_opts,
            "cli_opts": cli_opts,
            "gradation_opts": gradation_opts,
            "cli_roster": cli_roster,
            "roster_name": roster_name or "",
            "roster_cli": roster_cli or "",
            "roster_gradation": roster_gradation or "",
            "employees_open": employees_open,
            "roster_open": roster_open,
            "sync_notice": sync_notice or "",
            "sync_error": sync_error or "",
            "google_sync_ready": _google_sheet_sync_ready(),
            "google_sync_range": os.getenv("GOOGLE_SHEETS_EMPLOYEE_RANGE", "").strip() or "Employees!A:ZZ",
        },
    )


@app.post("/employees/sync-google")
def sync_employees_from_google_sheet(request: Request, session: Session = Depends(get_session)):
    wants_json = request.headers.get("x-requested-with", "").lower() == "fetch"
    try:
        rows, sheet_range = _fetch_google_employee_rows()
        added, updated = _import_employee_rows(session, rows, source_label=f"Google Sheet ({sheet_range})")
        message_text = f"Google Sheet sync complete: {added} added, {updated} updated."
        if wants_json:
            return JSONResponse({"ok": True, "message": message_text})
        message = quote(message_text)
        return RedirectResponse(url=f"/employees?sync_notice={message}#google-sync-card", status_code=303)
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Google Sheet sync failed."
        if wants_json:
            return JSONResponse({"ok": False, "message": detail}, status_code=exc.status_code)
        return RedirectResponse(url=f"/employees?sync_error={quote(detail)}#google-sync-card", status_code=303)
    except Exception as exc:
        if wants_json:
            return JSONResponse({"ok": False, "message": str(exc)}, status_code=500)
        return RedirectResponse(url=f"/employees?sync_error={quote(str(exc))}#google-sync-card", status_code=303)


@app.get("/employees/{emp_id}")
def edit_employee_page(emp_id: int, request: Request, session: Session = Depends(get_session)):
    employee = session.get(Employee, emp_id)
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    return templates.TemplateResponse(
        "employees_edit.html",
        {
            "request": request,
            "employee": employee,
            "role_order": ROLE_ORDER,
            "active_page": "employees",
        },
    )


@app.post("/employees/{emp_id}")
def update_employee(
    emp_id: int,
    name: str = Form(...),
    role: str = Form(...),
    retirement_date: str = Form(...),
    promotion_role: Optional[str] = Form(None),
    promotion_ready_date: Optional[str] = Form(None),
    category: Optional[str] = Form(None),
    pf_no: Optional[str] = Form(None),
    hrms: Optional[str] = Form(None),
    dob: Optional[str] = Form(None),
    doa: Optional[str] = Form(None),
    do_report: Optional[str] = Form(None),
    seniority_rank: Optional[str] = Form(None),
    status: Optional[str] = Form(None),
    working_at: Optional[str] = Form(None),
    gradation: Optional[str] = Form(None),
    cli: Optional[str] = Form(None),
    pme_due: Optional[str] = Form(None),
    technical_due: Optional[str] = Form(None),
    transportation_due: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    def to_date(val: Optional[str]) -> Optional[date]:
        return date.fromisoformat(val) if val else None
    def to_int(val: Optional[str]) -> Optional[int]:
        return int(val) if val not in (None, "", "None") else None

    employee = session.get(Employee, emp_id)
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")

    employee.name = name.strip()
    employee.role = role.strip()
    employee.retirement_date = to_date(retirement_date)
    employee.promotion_role = promotion_role.strip() if promotion_role else None
    employee.promotion_ready_date = to_date(promotion_ready_date)
    employee.category = category.strip() if category else None
    employee.pf_no = pf_no.strip() if pf_no else None
    employee.hrms = hrms.strip() if hrms else None
    employee.dob = to_date(dob)
    employee.doa = to_date(doa)
    employee.do_report = to_date(do_report)
    employee.seniority_rank = to_int(seniority_rank)
    employee.status = status.strip() if status else None
    employee.working_at = working_at.strip() if working_at else None
    employee.gradation = gradation.strip() if gradation else None
    employee.cli = cli.strip() if cli else None
    employee.pme_due = to_date(pme_due)
    employee.technical_due = to_date(technical_due)
    employee.transportation_due = to_date(transportation_due)

    session.add(employee)
    session.commit()
    return RedirectResponse("/", status_code=303)


@app.post("/employees/{emp_id}/delete")
def delete_employee(emp_id: int, session: Session = Depends(get_session)):
    employee = session.get(Employee, emp_id)
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    session.delete(employee)
    session.commit()
    return RedirectResponse("/", status_code=303)


@app.get("/uploads")
def uploads_page(request: Request):
    return templates.TemplateResponse(
        "uploads.html",
        {
            "request": request,
            "active_page": "uploads",
            "role_order": ROLE_ORDER,
        },
    )


def _cli_matrix_context(
    request: Request,
    error: Optional[str] = None,
    report_date: str = "",
    summary_rows: Optional[list[dict]] = None,
    overdue_rows: Optional[list[dict]] = None,
    saved_notice: str = "",
    cached_template_name: str = "",
):
    return {
        "request": request,
        "active_page": "cli_matrix",
        "role_order": ROLE_ORDER,
        "error": error,
        "report_date": report_date,
        "summary_rows": summary_rows or [],
        "overdue_rows": overdue_rows or [],
        "saved_notice": saved_notice,
        "cached_template_name": cached_template_name,
    }


def _template_store_paths(key: str) -> tuple[Path, Path]:
    safe_key = re.sub(r"[^a-zA-Z0-9_.-]+", "_", key.strip().lower()) or "template"
    return (
        TEMPLATE_STORE_DIR / f"{safe_key}.bin",
        TEMPLATE_STORE_DIR / f"{safe_key}.name",
    )


def _save_persistent_template(key: str, filename: str, payload: bytes) -> str:
    TEMPLATE_STORE_DIR.mkdir(parents=True, exist_ok=True)
    data_path, name_path = _template_store_paths(key)
    data_path.write_bytes(payload)
    stored_name = filename or "template.xlsx"
    name_path.write_text(stored_name, encoding="utf-8")
    return stored_name


def _load_persistent_template(key: str) -> tuple[bytes | None, str]:
    data_path, name_path = _template_store_paths(key)
    if not data_path.exists():
        return None, ""
    stored_name = name_path.read_text(encoding="utf-8").strip() if name_path.exists() else "template.xlsx"
    return data_path.read_bytes(), stored_name


EMPLOYEE_ALIAS_MAP = {
    "name": "name",
    "sl": "name",
    "slno": "name",
    "n": "name",
    "slname": "name",
    "degn": "role",
    "designation": "role",
    "design": "role",
    "role": "role",
    "hiredate": "hire_date",
    "dateofapptt": "hire_date",
    "dateofappt": "hire_date",
    "dateofappointment": "hire_date",
    "doa": "doa",
    "retirementdate": "retirement_date",
    "dor": "retirement_date",
    "promotionrole": "promotion_role",
    "promotionreadydate": "promotion_ready_date",
    "category": "category",
    "pf": "pf_no",
    "pfno": "pf_no",
    "pfnolen": "pf_no",
    "hrms": "hrms",
    "hrmsid": "hrms",
    "dob": "dob",
    "doareport": "do_report",
    "doreport": "do_report",
    "status": "status",
    "workingat": "working_at",
    "lobby": "working_at",
    "workingplace": "working_at",
    "gradation": "gradation",
    "cli": "cli",
    "pme": "pme_due",
    "pmedue": "pme_due",
    "pme_due": "pme_due",
    "technical": "technical_due",
    "technicaldue": "technical_due",
    "technical_due": "technical_due",
    "transportation": "transportation_due",
    "transportationdue": "transportation_due",
    "transportation_due": "transportation_due",
}


def _employee_norm(value: object | None) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum()) if value is not None else ""


def _derive_hire_date(dob_val: date | None, retirement_val: date | None) -> date | None:
    if dob_val:
        try:
            return dob_val.replace(year=dob_val.year + 25)
        except ValueError:
            return dob_val.replace(month=2, day=28, year=dob_val.year + 25)
    if retirement_val:
        return retirement_val - timedelta(days=35 * 365)
    return None


def _import_employee_rows(session: Session, rows: list[tuple | list], source_label: str = "sheet") -> tuple[int, int]:
    if not rows:
        raise HTTPException(status_code=400, detail=f"{source_label} is empty.")

    header_raw = None
    for r in rows:
        if any(cell not in (None, "", " ") for cell in r):
            header_raw = r
            break
    if header_raw is None:
        raise HTTPException(status_code=400, detail=f"{source_label} appears empty (no header row).")

    header_norm = [_employee_norm(h) for h in header_raw]
    mapped_cols = [EMPLOYEE_ALIAS_MAP.get(h, "") for h in header_norm]

    col_index: dict[str, int] = {}
    for idx, canonical in enumerate(mapped_cols):
        if canonical and canonical not in col_index:
            col_index[canonical] = idx

    required_cols = {"name", "role", "retirement_date"}
    missing_required = required_cols - set(col_index)

    if missing_required:
        first_row = rows[rows.index(header_raw)]
        if isinstance(first_row[0], (int, float)) and isinstance(first_row[1], str) and len(first_row) >= 14:
            positional_map = {
                "name": 1,
                "role": 2,
                "pf_no": 3,
                "hrms": 4,
                "category": 6,
                "gradation": 7,
                "working_at": 10,
                "cli": 11,
                "dob": 12,
                "retirement_date": 13,
                "pme_due": 14,
                "technical_due": 15,
                "transportation_due": 16,
            }
            for key, idx in positional_map.items():
                if key not in col_index and idx < len(first_row):
                    col_index[key] = idx
            missing_required = required_cols - set(col_index)

    if missing_required:
        raise HTTPException(status_code=400, detail=f"Missing columns in {source_label}: {', '.join(sorted(missing_required))}")

    added = 0
    updated = 0

    for row in rows[rows.index(header_raw) + 1 :]:
        def get(col: str) -> object | None:
            idx = col_index.get(col)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        name = get("name")
        role_raw = get("role")
        if name in (None, "") or role_raw in (None, ""):
            continue

        try:
            hire_date = _excel_to_date(get("hire_date"))
            retirement_date = _excel_to_date(get("retirement_date"))
            promo_ready = _excel_to_date(get("promotion_ready_date")) if "promotion_ready_date" in col_index else None
            dob = _excel_to_date(get("dob")) if "dob" in col_index else None
            doa = _excel_to_date(get("doa")) if "doa" in col_index else None
            do_report = _excel_to_date(get("do_report")) if "do_report" in col_index else None
            pme_due = _excel_to_date(get("pme_due")) if "pme_due" in col_index else None
            technical_due = _excel_to_date(get("technical_due")) if "technical_due" in col_index else None
            transportation_due = _excel_to_date(get("transportation_due")) if "transportation_due" in col_index else None
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Date parse error in {source_label}: {exc}") from exc

        if retirement_date is None:
            raise HTTPException(status_code=400, detail=f"retirement_date is required in {source_label}.")
        if hire_date is None:
            hire_date = _derive_hire_date(dob, retirement_date)
        if hire_date is None:
            raise HTTPException(status_code=400, detail=f"hire_date missing in {source_label} and could not be derived.")

        role = normalize_role(str(role_raw))
        promo_role = normalize_role(str(get("promotion_role"))) if "promotion_role" in col_index else None
        category = str(get("category")).strip() if "category" in col_index and get("category") else None
        pf_no = str(get("pf_no")).strip() if "pf_no" in col_index and get("pf_no") else None
        hrms = str(get("hrms")).strip() if "hrms" in col_index and get("hrms") else None
        status_val = str(get("status")).strip() if "status" in col_index and get("status") else None
        working_at = str(get("working_at")).strip() if "working_at" in col_index and get("working_at") else None

        existing = None
        if pf_no:
            existing = session.exec(select(Employee).where(Employee.pf_no == pf_no)).first()
        if existing is None and hrms:
            existing = session.exec(select(Employee).where(Employee.hrms == hrms)).first()
        if existing is None:
            existing = session.exec(
                select(Employee).where(Employee.name == str(name).strip(), Employee.role == role)
            ).first()

        if existing:
            existing.hire_date = hire_date
            existing.retirement_date = retirement_date
            existing.promotion_role = promo_role
            existing.promotion_ready_date = promo_ready
            existing.category = category
            existing.pf_no = pf_no
            existing.hrms = hrms
            existing.dob = dob
            existing.doa = doa
            existing.do_report = do_report
            existing.status = status_val
            existing.working_at = working_at
            existing.gradation = str(get("gradation")).strip() if "gradation" in col_index and get("gradation") else existing.gradation
            existing.cli = str(get("cli")).strip() if "cli" in col_index and get("cli") else existing.cli
            existing.pme_due = pme_due if pme_due else existing.pme_due
            existing.technical_due = technical_due if technical_due else existing.technical_due
            existing.transportation_due = transportation_due if transportation_due else existing.transportation_due
            updated += 1
        else:
            session.add(
                Employee(
                    name=str(name).strip(),
                    role=role,
                    hire_date=hire_date,
                    retirement_date=retirement_date,
                    promotion_role=promo_role,
                    promotion_ready_date=promo_ready,
                    category=category,
                    pf_no=pf_no,
                    hrms=hrms,
                    dob=dob,
                    doa=doa,
                    do_report=do_report,
                    status=status_val,
                    working_at=working_at,
                    gradation=str(get("gradation")).strip() if "gradation" in col_index and get("gradation") else None,
                    cli=str(get("cli")).strip() if "cli" in col_index and get("cli") else None,
                    pme_due=pme_due,
                    technical_due=technical_due,
                    transportation_due=transportation_due,
                )
            )
            added += 1

    session.commit()
    if added == 0 and updated == 0:
        raise HTTPException(status_code=400, detail=f"No rows imported from {source_label}. Check the sheet data or headers.")
    return added, updated


def _google_sheet_sync_ready() -> bool:
    return bool(
        os.getenv("GOOGLE_SHEETS_EMPLOYEE_SPREADSHEET_ID", "").strip()
        and (
            os.getenv("GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON", "").strip()
            or os.getenv("GOOGLE_SHEETS_SERVICE_ACCOUNT_FILE", "").strip()
        )
    )


def _normalize_google_sheet_range(sheet_range: str) -> str:
    raw = (sheet_range or "").strip()
    if "!" not in raw:
        return raw
    sheet_name, cell_range = raw.split("!", 1)
    sheet_name = sheet_name.strip()
    if not sheet_name:
        return raw
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        return f"{sheet_name}!{cell_range}"
    if any(ch.isspace() for ch in sheet_name):
        escaped = sheet_name.replace("'", "''")
        return f"'{escaped}'!{cell_range}"
    return f"{sheet_name}!{cell_range}"


def _fetch_google_employee_rows() -> tuple[list[list[str]], str]:
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_EMPLOYEE_SPREADSHEET_ID", "").strip()
    sheet_range = _normalize_google_sheet_range(
        os.getenv("GOOGLE_SHEETS_EMPLOYEE_RANGE", "").strip() or "Employees!A:ZZ"
    )
    if not spreadsheet_id:
        raise HTTPException(status_code=400, detail="Google Sheet sync is not configured: missing GOOGLE_SHEETS_EMPLOYEE_SPREADSHEET_ID.")

    service_account_json = os.getenv("GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = os.getenv("GOOGLE_SHEETS_SERVICE_ACCOUNT_FILE", "").strip()
    if not service_account_json and not service_account_file:
        raise HTTPException(status_code=400, detail="Google Sheet sync is not configured: missing service account credentials.")

    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except Exception as exc:
        raise HTTPException(status_code=500, detail="Google Sheets client libraries are not installed on the server.") from exc

    try:
        if service_account_json:
            info = json.loads(service_account_json)
            credentials = service_account.Credentials.from_service_account_info(
                info,
                scopes=GOOGLE_SHEETS_READONLY_SCOPE,
            )
        else:
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file,
                scopes=GOOGLE_SHEETS_READONLY_SCOPE,
            )
        service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
        ).execute()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read Google Sheet: {exc}") from exc

    rows = result.get("values", [])
    if not rows:
        raise HTTPException(status_code=400, detail="Google Sheet returned no rows.")
    return rows, sheet_range


def _run_one_time_cli_matrix_cleanup(session: Session) -> None:
    if CLI_MATRIX_2026_03_24_CLEANUP_SENTINEL.exists():
        return
    target_date = date(2026, 3, 24)
    for model in (CliMatrixSummarySnapshot, CliMatrixOverdueSnapshot):
        rows = session.exec(select(model).where(model.report_date == target_date)).all()
        for row in rows:
            session.delete(row)
    session.commit()
    CLI_MATRIX_2026_03_24_CLEANUP_SENTINEL.write_text("done", encoding="utf-8")


def _cli_matrix_record_date(value) -> date | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if hasattr(value, "to_pydatetime"):
        try:
            return value.to_pydatetime().date()
        except Exception:
            return None
    return None


def _save_cli_matrix_snapshots(
    session: Session,
    report_date_value: date,
    summary_df,
    overdue_df,
) -> None:
    _cleanup_cli_matrix_snapshots(session)

    existing_summary = session.exec(
        select(CliMatrixSummarySnapshot).where(
            CliMatrixSummarySnapshot.report_date == report_date_value
        )
    ).all()
    for row in existing_summary:
        session.delete(row)

    existing_overdue = session.exec(
        select(CliMatrixOverdueSnapshot).where(
            CliMatrixOverdueSnapshot.report_date == report_date_value
        )
    ).all()
    for row in existing_overdue:
        session.delete(row)

    for record in summary_df.to_dict(orient="records"):
        session.add(
            CliMatrixSummarySnapshot(
                report_date=report_date_value,
                row_no=int(record["S.No."]),
                cli_id=str(record["CLI ID"]),
                cli_name=str(record["CLI Name"]),
                alloted_desig=str(record["Alloted Desig."]),
                fp_over_due=int(record["FP Over Due"]),
                oldest_fp_overdue_date=_cli_matrix_record_date(
                    record["Oldest FP OverDue Date"]
                ),
            )
        )

    for record in overdue_df.to_dict(orient="records"):
        session.add(
            CliMatrixOverdueSnapshot(
                report_date=report_date_value,
                row_no=int(record["S.No."]),
                cli_id=str(record["CLI ID"]),
                cli_name=str(record["CLI Name"]),
                alloted_desig=str(record["Alloted Desig."]),
                fp_over_due=int(record["FP Over Due"]),
                oldest_fp_overdue_date=_cli_matrix_record_date(
                    record["Oldest FP OverDue Date"]
                ),
                counsel_over_due=int(record["Counsel Over Due"]),
                oldest_counsel_overdue_date=_cli_matrix_record_date(
                    record["Oldest Counsel OverDue Date"]
                ),
                grading_overdue=int(record["Grading OverDue"]),
                oldest_grading_overdue_date=_cli_matrix_record_date(
                    record["Oldest Grading OverDue"]
                ),
                total_over_due_cases=int(record["Total Over Due Cases"]),
            )
        )

    session.commit()


def _cleanup_cli_matrix_snapshots(session: Session) -> None:
    cutoff_date = date.today() - timedelta(days=30)
    old_summary = session.exec(
        select(CliMatrixSummarySnapshot).where(
            CliMatrixSummarySnapshot.report_date < cutoff_date
        )
    ).all()
    for row in old_summary:
        session.delete(row)

    old_overdue = session.exec(
        select(CliMatrixOverdueSnapshot).where(
            CliMatrixOverdueSnapshot.report_date < cutoff_date
        )
    ).all()
    for row in old_overdue:
        session.delete(row)

    session.commit()


def _load_cli_matrix_snapshots(session: Session, report_date_value: date) -> tuple[list[dict], list[dict]]:
    summary_rows = session.exec(
        select(CliMatrixSummarySnapshot)
        .where(CliMatrixSummarySnapshot.report_date == report_date_value)
        .order_by(CliMatrixSummarySnapshot.row_no)
    ).all()
    overdue_rows = session.exec(
        select(CliMatrixOverdueSnapshot)
        .where(CliMatrixOverdueSnapshot.report_date == report_date_value)
        .order_by(CliMatrixOverdueSnapshot.row_no)
    ).all()

    return (
        [
            {
                "S.No.": row.row_no,
                "CLI ID": row.cli_id,
                "CLI Name": row.cli_name,
                "Alloted Desig.": row.alloted_desig,
                "FP Over Due": row.fp_over_due,
                "Oldest FP OverDue Date": row.oldest_fp_overdue_date,
            }
            for row in summary_rows
        ],
        [
            {
                "S.No.": row.row_no,
                "CLI ID": row.cli_id,
                "CLI Name": row.cli_name,
                "Alloted Desig.": row.alloted_desig,
                "FP Over Due": row.fp_over_due,
                "Oldest FP OverDue Date": row.oldest_fp_overdue_date,
                "Counsel Over Due": row.counsel_over_due,
                "Oldest Counsel OverDue Date": row.oldest_counsel_overdue_date,
                "Grading OverDue": row.grading_overdue,
                "Oldest Grading OverDue": row.oldest_grading_overdue_date,
                "Total Over Due Cases": row.total_over_due_cases,
            }
            for row in overdue_rows
        ],
    )


def _delete_cli_matrix_snapshots_for_date(session: Session, report_date_value: date) -> None:
    for model in (CliMatrixSummarySnapshot, CliMatrixOverdueSnapshot):
        rows = session.exec(select(model).where(model.report_date == report_date_value)).all()
        for row in rows:
            session.delete(row)
    session.commit()


def _non_continuous_context(
    request: Request,
    variant_key: str,
    error: Optional[str] = None,
    report_date: str = "",
    sign_on_rows: Optional[list[dict]] = None,
    sign_off_rows: Optional[list[dict]] = None,
    saved_notice: str = "",
    template_token: str = "",
    source_name: str = "",
    cached_template_name: str = "",
):
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    return {
        "request": request,
        "active_page": config["active_page"],
        "role_order": ROLE_ORDER,
        "error": error,
        "report_date": report_date,
        "sign_on_rows": sign_on_rows or [],
        "sign_off_rows": sign_off_rows or [],
        "saved_notice": saved_notice,
        "template_token": template_token,
        "source_name": source_name,
        "cached_template_name": cached_template_name,
        "page_title": config["page_title"],
        "heading_title": config["heading_title"],
        "route_base": config["route_base"],
        "feature_name": config["feature_name"],
        "sign_on_label": config["sign_on_label"],
        "sign_off_label": config["sign_off_label"],
        "allow_reason_edit": variant_key == "non_sub",
    }


def _non_continuous_template_key(variant_key: str) -> str:
    return f"{variant_key}_non_continuous_template"


def _cache_non_continuous_template(variant_key: str, filename: str, payload: bytes) -> tuple[str, str]:
    stored_name = _save_persistent_template(_non_continuous_template_key(variant_key), filename, payload)
    return "saved", stored_name


def _load_non_continuous_template(variant_key: str) -> tuple[bytes | None, str]:
    return _load_persistent_template(_non_continuous_template_key(variant_key))


def _cleanup_non_continuous_snapshots(session: Session, variant_key: str) -> None:
    cutoff_date = date.today() - timedelta(days=30)
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    for model in (config["sign_on_model"], config["sign_off_model"]):
        rows = session.exec(select(model).where(model.report_date < cutoff_date)).all()
        for row in rows:
            session.delete(row)
    session.commit()


def _replace_non_continuous_section(
    session: Session,
    model,
    report_date_value: date,
    rows: list[dict],
) -> None:
    existing = session.exec(
        select(model).where(model.report_date == report_date_value)
    ).all()
    for row in existing:
        session.delete(row)

    for record in rows:
        session.add(
            model(
                report_date=report_date_value,
                row_no=int(record["SNO."]),
                crew_id=record["CREW ID"] or None,
                crew_name=record["CREW NAME"] or None,
                desig=record["DESIG."] or None,
                station=record["STATION"] or None,
                event_time=record["EVENT TIME"] or None,
                sup_id=record["SUP ID"] or None,
                entry_point=record["ENTRY POINT"] or None,
                train_no=record["TRAIN NO."] or None,
                loco_no=record["LOCO NO."] or None,
                duty_type=record["DUTY TYPE"] or None,
                route_stn=record["ROUTE STN"] or None,
                reason=record["REASON"] or None,
            )
        )


def _save_non_continuous_snapshot(
    session: Session,
    variant_key: str,
    report_date_value: date,
    section: str,
    rows: list[dict],
) -> None:
    _cleanup_non_continuous_snapshots(session, variant_key)
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    model = config["sign_on_model"] if section == "sign_on" else config["sign_off_model"]
    _replace_non_continuous_section(session, model, report_date_value, rows)
    session.commit()


def _load_non_continuous_snapshot(
    session: Session,
    variant_key: str,
    report_date_value: date,
) -> tuple[list[dict], list[dict]]:
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    sign_on_rows = session.exec(
        select(config["sign_on_model"])
        .where(config["sign_on_model"].report_date == report_date_value)
        .order_by(config["sign_on_model"].row_no)
    ).all()
    sign_off_rows = session.exec(
        select(config["sign_off_model"])
        .where(config["sign_off_model"].report_date == report_date_value)
        .order_by(config["sign_off_model"].row_no)
    ).all()

    def serialize(rows):
        return [
            {
                "SNO.": row.row_no,
                "CREW ID": row.crew_id or "",
                "CREW NAME": row.crew_name or "",
                "DESIG.": row.desig or "",
                "STATION": row.station or "",
                "EVENT TIME": row.event_time or "",
                "SUP ID": row.sup_id or "",
                "ENTRY POINT": row.entry_point or "",
                "TRAIN NO.": row.train_no or "",
                "LOCO NO.": row.loco_no or "",
                "DUTY TYPE": row.duty_type or "",
                "ROUTE STN": row.route_stn or "",
                "REASON": row.reason or "",
            }
            for row in rows
        ]

    return serialize(sign_on_rows), serialize(sign_off_rows)


def _delete_non_continuous_snapshots_for_date(
    session: Session,
    variant_key: str,
    report_date_value: date,
) -> None:
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    for model in (config["sign_on_model"], config["sign_off_model"]):
        rows = session.exec(select(model).where(model.report_date == report_date_value)).all()
        for row in rows:
            session.delete(row)
    session.commit()


def _update_non_continuous_reason(
    session: Session,
    variant_key: str,
    report_date_value: date,
    section: str,
    row_no: int,
    reason: str,
):
    config = NON_CONTINUOUS_VARIANTS[variant_key]
    model = config["sign_on_model"] if section == "sign_on" else config["sign_off_model"]
    row = session.exec(
        select(model)
        .where(model.report_date == report_date_value)
        .where(model.row_no == row_no)
    ).first()
    if not row:
        raise HTTPException(status_code=404, detail="Row not found.")
    row.reason = reason.strip() or None
    session.add(row)
    session.commit()
    return row.reason or ""


@app.get("/cli-matrix")
def cli_matrix_page(
    request: Request,
    error: Optional[str] = None,
    report_date: Optional[str] = None,
    session: Session = Depends(get_session),
):
    _cleanup_cli_matrix_snapshots(session)
    selected_date = coerce_report_date(report_date) or date.today()
    summary_rows, overdue_rows = _load_cli_matrix_snapshots(session, selected_date)
    _, cached_template_name = _load_persistent_template("cli_matrix")
    saved_notice = ""
    if report_date and not summary_rows and not overdue_rows:
        saved_notice = "No saved CLI Matrix snapshot found for the selected date."
    return templates.TemplateResponse(
        "cli_matrix.html",
        _cli_matrix_context(
            request,
            error=error,
            report_date=selected_date.isoformat(),
            summary_rows=summary_rows,
            overdue_rows=overdue_rows,
            saved_notice=saved_notice,
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/cli-matrix/preview")
async def preview_cli_matrix(
    request: Request,
    source_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    source_name = source_file.filename or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(source_name)
    selected_date = report_date_iso(inferred_date)
    if not source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "cli_matrix.html",
            _cli_matrix_context(
                request,
                error="Latest CLI Matrix must be an .xlsx file.",
                report_date=selected_date,
            ),
        )

    try:
        source_bytes = await source_file.read()
        summary_df = build_summary_df(source_bytes)
        overdue_df = build_sheet2_df(source_bytes)
        cached_template_name = ""
        if template_file and template_file.filename:
            if not template_file.filename.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Template workbook must be an .xlsx file.")
            cached_template_name = _save_persistent_template(
                "cli_matrix",
                template_file.filename,
                await template_file.read(),
            )
        else:
            _, cached_template_name = _load_persistent_template("cli_matrix")
        if inferred_date:
            _save_cli_matrix_snapshots(session, inferred_date, summary_df, overdue_df)
    except Exception as exc:
        return templates.TemplateResponse(
            "cli_matrix.html",
            _cli_matrix_context(
                request,
                error=f"CLI Matrix preview failed: {exc}",
                report_date=selected_date,
                cached_template_name=cached_template_name if "cached_template_name" in locals() else "",
            ),
        )

    return templates.TemplateResponse(
        "cli_matrix.html",
        _cli_matrix_context(
            request,
            report_date=selected_date,
            summary_rows=summary_df.to_dict(orient="records"),
            overdue_rows=overdue_df.to_dict(orient="records"),
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/cli-matrix/generate")
async def generate_cli_matrix(
    request: Request,
    source_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    source_name = source_file.filename or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(source_name)
    selected_date = report_date_iso(inferred_date)
    if not source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "cli_matrix.html",
            _cli_matrix_context(
                request,
                error="Latest CLI Matrix must be an .xlsx file.",
                report_date=selected_date,
                cached_template_name=_load_persistent_template("cli_matrix")[1],
            ),
        )

    try:
        source_bytes = await source_file.read()
        if template_file and template_file.filename:
            if not template_file.filename.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Template workbook must be an .xlsx file.")
            template_bytes = await template_file.read()
            cached_template_name = _save_persistent_template(
                "cli_matrix",
                template_file.filename,
                template_bytes,
            )
        else:
            template_bytes, cached_template_name = _load_persistent_template("cli_matrix")
            if not template_bytes:
                raise ValueError("Please upload the template workbook once before downloading.")
        summary_df = build_summary_df(source_bytes)
        overdue_df = build_sheet2_df(source_bytes)
        output = build_output_workbook(
            source_bytes,
            template_bytes,
            source_name,
            inferred_date,
        )
        if inferred_date:
            _save_cli_matrix_snapshots(session, inferred_date, summary_df, overdue_df)
    except Exception as exc:
        return templates.TemplateResponse(
            "cli_matrix.html",
            _cli_matrix_context(
                request,
                error=f"CLI Matrix generation failed: {exc}",
                report_date=selected_date,
                cached_template_name=cached_template_name if "cached_template_name" in locals() else _load_persistent_template("cli_matrix")[1],
            ),
        )

    base_name = source_name.rsplit(".", 1)[0] if "." in source_name else "CLI_Matrix"
    filename = f"{base_name}_updated.xlsx"
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/cli-matrix/reset")
async def reset_cli_matrix_data(
    request: Request,
    report_date: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    selected_date = coerce_report_date(report_date) or date.today()
    _delete_cli_matrix_snapshots_for_date(session, selected_date)
    _, cached_template_name = _load_persistent_template("cli_matrix")
    return templates.TemplateResponse(
        "cli_matrix.html",
        _cli_matrix_context(
            request,
            report_date=selected_date.isoformat(),
            saved_notice=f"Saved CLI Matrix data for {selected_date.strftime('%d-%m-%Y')} has been deleted.",
            cached_template_name=cached_template_name,
        ),
    )


@app.get("/non-continuous-duty")
def non_continuous_duty_page(
    request: Request,
    error: Optional[str] = None,
    report_date: Optional[str] = None,
    source_name: Optional[str] = None,
    template_token: Optional[str] = None,
    session: Session = Depends(get_session),
):
    variant_key = "non_sub"
    _cleanup_non_continuous_snapshots(session, variant_key)
    selected_date = coerce_report_date(report_date) or date.today()
    sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(session, variant_key, selected_date)
    _, cached_template_name = _load_non_continuous_template(variant_key)
    saved_notice = ""
    if report_date and not sign_on_rows and not sign_off_rows:
        saved_notice = "No saved NON SUB NON CONTINUOUS DUTY snapshot found for the selected date."
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            error=error,
            report_date=selected_date.isoformat(),
            sign_on_rows=sign_on_rows,
            sign_off_rows=sign_off_rows,
            saved_notice=saved_notice,
            template_token="saved" if cached_template_name else "",
            source_name=source_name or "",
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/non-continuous-duty/preview")
async def preview_non_continuous_duty(
    request: Request,
    source_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    template_token: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "non_sub"
    source_name = source_file.filename or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(source_name)
    selected_date = report_date_iso(inferred_date)
    if not source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Source workbook must be an .xlsx file.",
                report_date=selected_date,
            ),
        )

    try:
        source_bytes = await source_file.read()
        section, rows = parse_non_continuous_source(source_bytes)
        if inferred_date:
            _save_non_continuous_snapshot(session, variant_key, inferred_date, section, rows)
        cached_template_name = ""
        if template_file and template_file.filename:
            if not template_file.filename.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Formal / template workbook must be an .xlsx file.")
            template_token, cached_template_name = _cache_non_continuous_template(
                variant_key,
                template_file.filename,
                await template_file.read(),
            )
        else:
            _, cached_template_name = _load_non_continuous_template(variant_key)
            template_token = "saved" if cached_template_name else ""
        sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(
            session, variant_key, inferred_date or date.today()
        )
    except Exception as exc:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error=f"NON CONTINUOUS DUTY preview failed: {exc}",
                report_date=selected_date,
                template_token="saved" if ("cached_template_name" in locals() and cached_template_name) else "",
                source_name=source_name,
                cached_template_name=cached_template_name if "cached_template_name" in locals() else "",
            ),
        )

    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            report_date=selected_date,
            sign_on_rows=sign_on_rows,
            sign_off_rows=sign_off_rows,
            template_token=template_token or "",
            source_name=source_name,
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/non-continuous-duty/generate")
async def generate_non_continuous_duty(
    request: Request,
    source_file: Optional[UploadFile] = File(None),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    template_token: Optional[str] = Form(None),
    source_name: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "non_sub"
    variant_config = NON_CONTINUOUS_VARIANTS[variant_key]
    uploaded_source_name = source_file.filename if source_file and source_file.filename else ""
    template_name = template_file.filename if template_file else ""
    display_source_name = uploaded_source_name or source_name or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(display_source_name)
    selected_date = report_date_iso(inferred_date)

    if uploaded_source_name and not uploaded_source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Source workbook must be an .xlsx file.",
                report_date=selected_date,
                template_token=template_token or "",
                source_name=display_source_name,
            ),
        )
    if not uploaded_source_name and not display_source_name:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Please click Generate first or choose a source workbook before downloading.",
                report_date=selected_date,
                template_token=template_token or "",
            ),
        )

    try:
        if uploaded_source_name:
            source_bytes = await source_file.read()
            section, rows = parse_non_continuous_source(source_bytes)
            if inferred_date:
                _save_non_continuous_snapshot(session, variant_key, inferred_date, section, rows)
        if template_file and template_name:
            if not template_name.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Formal / template workbook must be an .xlsx file.")
            template_bytes = await template_file.read()
            template_token, cached_template_name = _cache_non_continuous_template(variant_key, template_name, template_bytes)
        else:
            template_bytes, cached_template_name = _load_non_continuous_template(variant_key)
            template_token = "saved" if cached_template_name else ""
            if not template_bytes:
                raise ValueError("Please choose the formal / template workbook once before downloading.")
        sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(
            session, variant_key, inferred_date or date.today()
        )
        if not sign_on_rows and not sign_off_rows:
            raise ValueError("No saved NON SUB NON CONTINUOUS DUTY data found for the selected date. Please click Generate first.")
        output = build_non_continuous_workbook(
            sign_on_rows,
            sign_off_rows,
            template_bytes,
            inferred_date,
            sheet_title=variant_config["sheet_title"],
            output_sign_on_title=variant_config["sign_on_label"],
            output_sign_off_title=variant_config["sign_off_label"],
        )
    except Exception as exc:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error=f"NON CONTINUOUS DUTY generation failed: {exc}",
                report_date=selected_date,
                template_token=template_token or "",
                source_name=display_source_name,
                cached_template_name=cached_template_name if 'cached_template_name' in locals() else "",
            ),
        )

    base_name = display_source_name.rsplit(".", 1)[0] if "." in display_source_name else "NON_CONTINUOUS_DUTY"
    filename = f"{base_name}_updated.xlsx"
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/non-continuous-duty/reset")
async def reset_non_continuous_duty_data(
    request: Request,
    report_date: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "non_sub"
    selected_date = coerce_report_date(report_date) or date.today()
    _delete_non_continuous_snapshots_for_date(session, variant_key, selected_date)
    _, cached_template_name = _load_non_continuous_template(variant_key)
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            report_date=selected_date.isoformat(),
            saved_notice=f"Saved NON SUB data for {selected_date.strftime('%d-%m-%Y')} has been deleted.",
            template_token="saved" if cached_template_name else "",
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/non-continuous-duty/reason")
async def update_non_continuous_duty_reason(
    request: Request,
    session: Session = Depends(get_session),
):
    try:
        payload = await request.json()
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Invalid request payload.") from exc

    report_date_raw = str(payload.get("report_date") or "").strip()
    section = str(payload.get("section") or "").strip().lower()
    reason = str(payload.get("reason") or "")

    if section not in {"sign_on", "sign_off"}:
        raise HTTPException(status_code=400, detail="Invalid section.")

    try:
        report_date_value = date.fromisoformat(report_date_raw)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Invalid report date.") from exc

    try:
        row_no = int(payload.get("row_no"))
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="Invalid row number.") from exc

    saved_reason = _update_non_continuous_reason(
        session,
        "non_sub",
        report_date_value,
        section,
        row_no,
        reason,
    )
    return JSONResponse({"ok": True, "reason": saved_reason})


@app.get("/sub-non-continuous-duty")
def sub_non_continuous_duty_page(
    request: Request,
    error: Optional[str] = None,
    report_date: Optional[str] = None,
    source_name: Optional[str] = None,
    template_token: Optional[str] = None,
    session: Session = Depends(get_session),
):
    variant_key = "sub"
    _cleanup_non_continuous_snapshots(session, variant_key)
    selected_date = coerce_report_date(report_date) or date.today()
    sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(session, variant_key, selected_date)
    _, cached_template_name = _load_non_continuous_template(variant_key)
    saved_notice = ""
    if report_date and not sign_on_rows and not sign_off_rows:
        saved_notice = "No saved SUB NON CONTINUOUS DUTY snapshot found for the selected date."
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            error=error,
            report_date=selected_date.isoformat(),
            sign_on_rows=sign_on_rows,
            sign_off_rows=sign_off_rows,
            saved_notice=saved_notice,
            template_token="saved" if cached_template_name else "",
            source_name=source_name or "",
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/sub-non-continuous-duty/preview")
async def preview_sub_non_continuous_duty(
    request: Request,
    source_file: UploadFile = File(...),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    template_token: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "sub"
    source_name = source_file.filename or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(source_name)
    selected_date = report_date_iso(inferred_date)
    if not source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Source workbook must be an .xlsx file.",
                report_date=selected_date,
            ),
        )

    try:
        source_bytes = await source_file.read()
        section, rows = parse_non_continuous_source(source_bytes)
        if inferred_date:
            _save_non_continuous_snapshot(session, variant_key, inferred_date, section, rows)
        cached_template_name = ""
        if template_file and template_file.filename:
            if not template_file.filename.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Formal / template workbook must be an .xlsx file.")
            template_token, cached_template_name = _cache_non_continuous_template(
                variant_key,
                template_file.filename,
                await template_file.read(),
            )
        else:
            _, cached_template_name = _load_non_continuous_template(variant_key)
            template_token = "saved" if cached_template_name else ""
        sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(
            session, variant_key, inferred_date or date.today()
        )
    except Exception as exc:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error=f"SUB NON CONTINUOUS DUTY preview failed: {exc}",
                report_date=selected_date,
                template_token="saved" if ("cached_template_name" in locals() and cached_template_name) else "",
                source_name=source_name,
                cached_template_name=cached_template_name if "cached_template_name" in locals() else "",
            ),
        )

    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            report_date=selected_date,
            sign_on_rows=sign_on_rows,
            sign_off_rows=sign_off_rows,
            template_token=template_token or "",
            source_name=source_name,
            cached_template_name=cached_template_name,
        ),
    )


@app.post("/sub-non-continuous-duty/generate")
async def generate_sub_non_continuous_duty(
    request: Request,
    source_file: Optional[UploadFile] = File(None),
    template_file: Optional[UploadFile] = File(None),
    report_date: Optional[str] = Form(None),
    template_token: Optional[str] = Form(None),
    source_name: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "sub"
    variant_config = NON_CONTINUOUS_VARIANTS[variant_key]
    uploaded_source_name = source_file.filename if source_file and source_file.filename else ""
    template_name = template_file.filename if template_file else ""
    display_source_name = uploaded_source_name or source_name or ""
    inferred_date = coerce_report_date(report_date) or infer_report_date(display_source_name)
    selected_date = report_date_iso(inferred_date)

    if uploaded_source_name and not uploaded_source_name.lower().endswith((".xlsx", ".xlsm")):
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Source workbook must be an .xlsx file.",
                report_date=selected_date,
                template_token=template_token or "",
                source_name=display_source_name,
            ),
        )
    if not uploaded_source_name and not display_source_name:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error="Please click Generate first or choose a source workbook before downloading.",
                report_date=selected_date,
                template_token=template_token or "",
            ),
        )

    try:
        if uploaded_source_name:
            source_bytes = await source_file.read()
            section, rows = parse_non_continuous_source(source_bytes)
            if inferred_date:
                _save_non_continuous_snapshot(session, variant_key, inferred_date, section, rows)
        if template_file and template_name:
            if not template_name.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Formal / template workbook must be an .xlsx file.")
            template_bytes = await template_file.read()
            template_token, cached_template_name = _cache_non_continuous_template(variant_key, template_name, template_bytes)
        else:
            template_bytes, cached_template_name = _load_non_continuous_template(variant_key)
            template_token = "saved" if cached_template_name else ""
            if not template_bytes:
                raise ValueError("Please choose the formal / template workbook once before downloading.")
        sign_on_rows, sign_off_rows = _load_non_continuous_snapshot(
            session, variant_key, inferred_date or date.today()
        )
        if not sign_on_rows and not sign_off_rows:
            raise ValueError("No saved SUB NON CONTINUOUS DUTY data found for the selected date. Please click Generate first.")
        output = build_non_continuous_workbook(
            sign_on_rows,
            sign_off_rows,
            template_bytes,
            inferred_date,
            sheet_title=variant_config["sheet_title"],
            output_sign_on_title=variant_config["sign_on_label"],
            output_sign_off_title=variant_config["sign_off_label"],
        )
    except Exception as exc:
        return templates.TemplateResponse(
            "non_continuous_duty.html",
            _non_continuous_context(
                request,
                variant_key,
                error=f"SUB NON CONTINUOUS DUTY generation failed: {exc}",
                report_date=selected_date,
                template_token=template_token or "",
                source_name=display_source_name,
                cached_template_name=cached_template_name if 'cached_template_name' in locals() else "",
            ),
        )

    base_name = display_source_name.rsplit(".", 1)[0] if "." in display_source_name else "SUB_NON_CONTINUOUS_DUTY"
    filename = f"{base_name}_updated.xlsx"
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/sub-non-continuous-duty/reset")
async def reset_sub_non_continuous_duty_data(
    request: Request,
    report_date: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    variant_key = "sub"
    selected_date = coerce_report_date(report_date) or date.today()
    _delete_non_continuous_snapshots_for_date(session, variant_key, selected_date)
    _, cached_template_name = _load_non_continuous_template(variant_key)
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_context(
            request,
            variant_key,
            report_date=selected_date.isoformat(),
            saved_notice=f"Saved SUB data for {selected_date.strftime('%d-%m-%Y')} has been deleted.",
            template_token="saved" if cached_template_name else "",
            cached_template_name=cached_template_name,
        ),
    )


@app.get("/requirements")
def requirements_page(
    request: Request,
    as_of: Optional[str] = None,
    horizon_months: int = 12,
    lead_time_days: int = 90,
    session: Session = Depends(get_session),
):
    plan_date = _parse_as_of(request, as_of)
    today = date.today()
    horizon_days = 0
    horizon_months = 0
    lead_time_days = 0
    employees_now = fetch_active_employees(session, today)
    employees_now = apply_promotions(employees_now, today)
    requirements_map = load_requirements_map(session)
    counts = headcount_by_role(employees_now)
    retiring_raw = project_retirements_window(employees_now, today, plan_date)
    recruit_plan_simple = build_simple_recruit_plan(retiring_raw, lead_days=30)
    promotion_plan = build_promotion_plan(employees_now, today, horizon_months)
    requirements = sorted(session.exec(select(Requirement)).all(), key=lambda r: role_sort_key(r.role))
    target_date = plan_date
    retire_counts = {role: len(peeps) for role, peeps in retiring_raw.items()}

    recruit_plan = {
        role: sorted(steps, key=_extract_date)
        for role, steps in recruit_plan_simple.items()
    }
    retiring = {
        role: sorted(people, key=lambda e: e.retirement_date)
        for role, people in retiring_raw.items()
    }

    response = templates.TemplateResponse(
        "requirements.html",
        {
            "request": request,
            "active_page": "requirements",
            "role_order": ROLE_ORDER,
            "as_of": plan_date,
            "horizon_months": horizon_months,
            "lead_time_days": lead_time_days,
            "counts": counts,
            "requirements": requirements,
            "recruit_plan": recruit_plan,
            "promotion_plan": promotion_plan,
            "retiring": {k: retiring[k] for k in sorted(retiring, key=role_sort_key)},
            "retire_counts": retire_counts,
            "target_date": target_date,
        },
    )
    response.set_cookie("as_of", plan_date.isoformat())
    return response


@app.get("/reports")
def reports_page(
    request: Request,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    role: Optional[str] = None,
    session: Session = Depends(get_session),
):
    start = _parse_date_cookie(request, "reports_start_date", start_date)
    end = _parse_date_cookie(request, "reports_end_date", end_date)
    if end < start:
        start, end = end, start
    horizon_months = 0
    employees = session.exec(select(Employee)).all()
    cli_distribution = build_cli_distribution(employees)
    dynamic_roles = sorted({e.role for e in employees if e.role not in ROLE_ORDER})
    role_headers = ROLE_ORDER + [r for r in dynamic_roles if r not in ROLE_ORDER]
    working_summary = []
    working_map: dict[str, dict[str, int]] = {}
    allowed_working = [
        "CC(R) BT",
        "CC(R) DDJ",
        "CC(R) NH",
        "CC(R) NORTH",
        "CC(R) RHA",
        "CC(R) KOAA",
        "CC(R) SOUTH",
    ]
    allowed_norm = {loc.upper(): loc for loc in allowed_working}
    for e in employees:
        loc_raw = (e.working_at or "").strip()
        loc_key = loc_raw.upper()
        if loc_key not in allowed_norm:
            continue  # skip non-CCR entries
        loc = allowed_norm[loc_key]  # use canonical casing
        role_key = e.role
        working_map.setdefault(loc, {}).setdefault(role_key, 0)
        working_map[loc][role_key] += 1
    for loc in sorted(working_map.keys(), key=lambda x: x.lower()):
        counts = {r: working_map[loc].get(r, 0) for r in role_headers}
        working_summary.append(
            {
                "working_at": loc,
                "counts": counts,
                "total": sum(counts.values()),
            }
        )
    retirements: dict[str, int] = {}
    retiring_list = []
    for e in employees:
        if e.retirement_date and start <= e.retirement_date <= end:
            retirements[e.role] = retirements.get(e.role, 0) + 1
            retiring_list.append(e)
    retirements = {k: retirements.get(k, 0) for k in ROLE_ORDER if k in retirements} | {
        k: v for k, v in retirements.items() if k not in ROLE_ORDER
    }
    if role:
        retiring_list = [e for e in retiring_list if e.role == role]
    retiring_list = sorted(retiring_list, key=lambda e: (e.retirement_date, role_sort_key(e.role), e.name))
    response = templates.TemplateResponse(
        "reports.html",
        {
            "request": request,
            "active_page": "reports",
        "start_date": start,
        "end_date": end,
        "retirements": retirements,
        "role_filter": role or "",
        "retiring_list": retiring_list,
        "dashboard_link": f"/?as_of={start.isoformat()}&horizon_months={horizon_months}",
        "cli_distribution": cli_distribution,
        "role_headers": role_headers,
        "working_summary": working_summary,
    },
)
    response.set_cookie("as_of", end.isoformat())
    response.set_cookie("reports_start_date", start.isoformat())
    response.set_cookie("reports_end_date", end.isoformat())
    return response


@app.get("/reports/cli-distribution.xlsx")
def download_cli_distribution(session: Session = Depends(get_session)):
    employees = session.exec(select(Employee)).all()
    cli_distribution = build_cli_distribution(employees)

    wb = Workbook()
    ws = wb.active
    ws.title = "CLI Distribution"
    ws.append(["CLI", "Gradation A", "Gradation B", "Gradation C", "Total"])
    for row in cli_distribution:
        ws.append([row["cli"], row["A"], row["B"], row["C"], row["total"]])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"cli_distribution_{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/requirements")
def upsert_requirement(
    role: str = Form(...),
    needed: int = Form(...),
    session: Session = Depends(get_session),
):
    role = role.strip()
    record = session.get(Requirement, role)
    if record:
        record.needed = needed
    else:
        record = Requirement(role=role, needed=needed)
        session.add(record)
    session.commit()
    return RedirectResponse("/", status_code=303)


@app.post("/employees")
def add_employee(
    name: str = Form(...),
    role: str = Form(...),
    hire_date: str = Form(...),
    retirement_date: str = Form(...),
    promotion_role: Optional[str] = Form(None),
    promotion_ready_date: Optional[str] = Form(None),
    category: Optional[str] = Form(None),
    pf_no: Optional[str] = Form(None),
    hrms: Optional[str] = Form(None),
    dob: Optional[str] = Form(None),
    doa: Optional[str] = Form(None),
    do_report: Optional[str] = Form(None),
    seniority_rank: Optional[str] = Form(None),
    status: Optional[str] = Form(None),
    working_at: Optional[str] = Form(None),
    gradation: Optional[str] = Form(None),
    cli: Optional[str] = Form(None),
    pme_due: Optional[str] = Form(None),
    technical_due: Optional[str] = Form(None),
    transportation_due: Optional[str] = Form(None),
    session: Session = Depends(get_session),
):
    def to_date(val: Optional[str]) -> Optional[date]:
        return date.fromisoformat(val) if val else None
    def to_int(val: Optional[str]) -> Optional[int]:
        return int(val) if val not in (None, "", "None") else None

    role_norm = role.strip()
    existing = session.exec(
        select(Employee).where(Employee.name == name.strip(), Employee.role == role_norm)
    ).first()

    if existing:
        existing.hire_date = to_date(hire_date)
        existing.retirement_date = to_date(retirement_date)
        existing.promotion_role = promotion_role.strip() if promotion_role else None
        existing.promotion_ready_date = to_date(promotion_ready_date)
        existing.category = category.strip() if category else None
        existing.pf_no = pf_no.strip() if pf_no else None
        existing.hrms = hrms.strip() if hrms else None
        existing.dob = to_date(dob)
        existing.doa = to_date(doa)
        existing.do_report = to_date(do_report)
        existing.seniority_rank = to_int(seniority_rank)
        existing.status = status.strip() if status else None
        existing.working_at = working_at.strip() if working_at else None
        existing.gradation = gradation.strip() if gradation else None
        existing.cli = cli.strip() if cli else None
        existing.pme_due = to_date(pme_due)
        existing.technical_due = to_date(technical_due)
        existing.transportation_due = to_date(transportation_due)
    else:
        employee = Employee(
            name=name.strip(),
            role=role_norm,
            hire_date=to_date(hire_date),
            retirement_date=to_date(retirement_date),
        promotion_role=promotion_role.strip() if promotion_role else None,
        promotion_ready_date=to_date(promotion_ready_date),
        category=category.strip() if category else None,
        pf_no=pf_no.strip() if pf_no else None,
        hrms=hrms.strip() if hrms else None,
        dob=to_date(dob),
        doa=to_date(doa),
        do_report=to_date(do_report),
        seniority_rank=to_int(seniority_rank),
        status=status.strip() if status else None,
        working_at=working_at.strip() if working_at else None,
        gradation=gradation.strip() if gradation else None,
        cli=cli.strip() if cli else None,
        pme_due=to_date(pme_due),
        technical_due=to_date(technical_due),
        transportation_due=to_date(transportation_due),
    )
        session.add(employee)
    session.commit()
    return RedirectResponse("/", status_code=303)


@app.get("/api/plan")
def api_plan(
    as_of: Optional[str] = None,
    horizon_months: int = 12,
    lead_time_days: int = 90,
    session: Session = Depends(get_session),
):
    plan_date = date.fromisoformat(as_of) if as_of else date.today()
    employees = fetch_active_employees(session, plan_date)
    employees = apply_promotions(employees, plan_date)
    requirements_map = load_requirements_map(session)
    counts = headcount_by_role(employees)
    recruit_plan = build_recruit_plan(requirements_map, employees, plan_date, horizon_months, lead_time_days)
    promotion_plan = build_promotion_plan(employees, plan_date, horizon_months)
    return JSONResponse(
        {
            "as_of": plan_date.isoformat(),
            "horizon_months": horizon_months,
            "lead_time_days": lead_time_days,
            "counts": counts,
            "recruit_plan": recruit_plan,
            "promotion_plan": promotion_plan,
        }
    )


def _excel_to_date(val: object) -> date | None:
    if val is None or val == "":
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, (int, float)):
        origin = date(1899, 12, 30)
        return origin + timedelta(days=int(val))
    if isinstance(val, str):
        s = val.strip()
        if not s or set(s) <= set(".-/"):
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        raise ValueError(f"Unrecognized date format: {val!r}")
    raise ValueError(f"Unsupported date value: {val!r}")


def _to_int(val: object | None) -> int | None:
    if val is None or val == "":
        return None
    try:
        return int(val)
    except (TypeError, ValueError):
        try:
            return int(float(val))
        except Exception as exc:
            raise ValueError(f"Invalid integer value: {val!r}") from exc


@app.post("/upload")
async def upload_employees(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=400,
            detail="Upload an .xlsx file with columns: name, role, hire_date, retirement_date. Optional: promotion_role, promotion_ready_date, category, pf_no, hrms, dob, doa, do_report, status, working_at.",
        )

    content = await file.read()
    wb = load_workbook(filename=BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    _import_employee_rows(session, rows, source_label="uploaded workbook")
    return RedirectResponse("/", status_code=303)


@app.post("/upload-seniority")
async def upload_seniority(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Upload an .xlsx file with columns: name, role, seniority_rank (or seniority). Optional: promotion_role, promotion_ready_date.")

    content = await file.read()
    wb = load_workbook(filename=BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(status_code=400, detail="Workbook is empty.")

    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    rank_col = "seniority_rank" if "seniority_rank" in header else ("seniority" if "seniority" in header else None)
    if rank_col is None:
        raise HTTPException(status_code=400, detail="Missing column: seniority_rank")
    required_cols = {"name", "role", rank_col}
    if not required_cols.issubset(set(header)):
        missing = required_cols - set(header)
        raise HTTPException(status_code=400, detail=f"Missing columns: {', '.join(missing)}")

    col_index = {col: header.index(col) for col in header if col}
    updated = 0

    for row in rows[1:]:
        def get(col: str) -> object | None:
            idx = col_index.get(col)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        name = get("name")
        role_raw = get("role")
        if name in (None, "") or role_raw in (None, ""):
            continue

        try:
            rank = _to_int(get(rank_col))
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Invalid seniority_rank: {exc}") from exc

        promo_role = normalize_role(str(get("promotion_role"))) if "promotion_role" in col_index else None
        try:
            promo_ready = _excel_to_date(get("promotion_ready_date")) if "promotion_ready_date" in col_index else None
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Date parse error: {exc}") from exc

        role = normalize_role(str(role_raw))
        employee = session.exec(
            select(Employee).where(Employee.name == str(name).strip(), Employee.role == role)
        ).first()
        if not employee:
            continue

        employee.seniority_rank = rank
        if promo_role:
            employee.promotion_role = promo_role
        if promo_ready:
            employee.promotion_ready_date = promo_ready
        updated += 1

    session.commit()
    if updated == 0:
        raise HTTPException(status_code=400, detail="No matching employees updated. Ensure names/roles match the roster.")
    return RedirectResponse("/", status_code=303)
