from __future__ import annotations

from collections import Counter
import csv
from datetime import date, datetime, timedelta
import json
import math
import os
from io import BytesIO, StringIO
from pathlib import Path
from typing import Optional
import re
from urllib.parse import quote

from fastapi import Depends, FastAPI, Form, Request, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, StreamingResponse
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
GOOGLE_EMPLOYEE_STATION_TABS = ["North", "South", "KOAA", "DDJ", "RHA", "NH", "BT"]
TEMPLATE_STORE_DIR = DB_PATH.parent / "saved_templates"
CLI_MATRIX_2026_03_24_CLEANUP_SENTINEL = DB_PATH.parent / ".cli_matrix_cleanup_2026_03_24.done"
EMPLOYEE_MASTER_SMART_CLEANUP_SENTINEL = DB_PATH.parent / ".employee_master_smart_cleanup_2026_03_27.done"
EMPLOYEE_MASTER_KEEP_BOTH_FILE = DB_PATH.parent / "employee_master_keep_both.json"
EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE = DB_PATH.parent / "employee_master_source_snapshot.json"
EMPLOYEE_MASTER_EXTRA_REVIEW_KEEP_FILE = DB_PATH.parent / "employee_master_extra_review_keep.json"
LI_GRADING_METADATA_FILE = DB_PATH.parent / "li_grading_metadata.json"
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


def _sensitive_action_password() -> str:
    return os.getenv("SENSITIVE_ACTION_PASSWORD") or "11111"


def _validate_sensitive_action_password(password: str | None) -> None:
    if (password or "") != _sensitive_action_password():
        raise HTTPException(status_code=403, detail="Code validation failed. Enter the current action code to continue.")

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
        _run_one_time_employee_master_cleanup(session)
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


def _cli_page_context(
    request: Request,
    session: Session,
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_gradation: Optional[str] = None,
    grading_update_notice: str = "",
    grading_update_warning: str = "",
    grading_update_error: Optional[str] = None,
    grading_update_details: Optional[list[str]] = None,
    grading_warning_details: Optional[list[str]] = None,
) -> dict[str, object]:
    employees_all = session.exec(select(Employee)).all()
    cli_distribution = build_cli_distribution(employees_all)
    grading_meta = _load_li_grading_metadata()
    grading_report_date = coerce_report_date(grading_meta.get("report_date"))
    grading_saved_at = ""
    saved_at_raw = grading_meta.get("saved_at", "")
    if saved_at_raw:
        try:
            grading_saved_at = datetime.fromisoformat(saved_at_raw).strftime("%d/%m/%Y %I:%M %p")
        except ValueError:
            grading_saved_at = saved_at_raw
    cli_opts_map: dict[str, str] = {}
    for val in [e.cli for e in employees_all if e.cli]:
        key = val.strip().lower()
        if key not in cli_opts_map:
            cli_opts_map[key] = val.strip()
    cli_opts = [v for _, v in sorted(cli_opts_map.items(), key=lambda item: item[0])]
    gradation_opts = sorted({e.gradation for e in employees_all if e.gradation})

    roster_filter_active = any([roster_name, roster_cli, roster_gradation])
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

    return {
        "request": request,
        "active_page": "cli",
        "cli_distribution": cli_distribution,
        "cli_roster": cli_roster,
        "cli_opts": cli_opts,
        "gradation_opts": gradation_opts,
        "roster_name": roster_name or "",
        "roster_cli": roster_cli or "",
        "roster_gradation": roster_gradation or "",
        "roster_open": roster_filter_active,
        "grading_source_name": grading_meta.get("filename", ""),
        "grading_report_date": grading_report_date.strftime("%d-%m-%Y") if grading_report_date else "",
        "grading_saved_at": grading_saved_at,
        "grading_update_notice": grading_update_notice,
        "grading_update_warning": grading_update_warning,
        "grading_update_error": grading_update_error or "",
        "grading_update_details": grading_update_details or [],
        "grading_warning_details": grading_warning_details or [],
    }


@app.get("/cli")
def cli_page(
    request: Request,
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_gradation: Optional[str] = None,
    session: Session = Depends(get_session),
):
    return templates.TemplateResponse(
        "cli.html",
        _cli_page_context(
            request,
            session,
            roster_name=roster_name,
            roster_cli=roster_cli,
            roster_gradation=roster_gradation,
        ),
    )


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
    sync_warning: Optional[str] = None,
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
            "sync_warning": sync_warning or "",
            "sync_error": sync_error or "",
            "google_sync_ready": _google_sheet_sync_ready(),
            "google_sync_range": ", ".join(GOOGLE_EMPLOYEE_STATION_TABS),
        },
    )


@app.post("/employees/sync-google")
def sync_employees_from_google_sheet(
    request: Request,
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    wants_json = request.headers.get("x-requested-with", "").lower() == "fetch"
    try:
        _validate_sensitive_action_password(action_password)
        sources, source_label = _fetch_google_employee_rows()
        added = 0
        updated = 0
        warnings: list[str] = []
        sync_details: list[str] = []
        sync_stats: dict[str, int] = {"unchanged": 0}
        global_pf_counts: Counter[str] = Counter()
        global_hrms_counts: Counter[str] = Counter()

        for rows, _, _ in sources:
            if not rows:
                continue
            header_raw = next((r for r in rows if any(cell not in (None, "", " ") for cell in r)), None)
            if header_raw is None:
                continue
            header_norm = [_employee_norm(h) for h in header_raw]
            mapped_cols = [EMPLOYEE_ALIAS_MAP.get(h, "") for h in header_norm]
            col_index: dict[str, int] = {}
            for idx, canonical in enumerate(mapped_cols):
                if canonical and canonical not in col_index:
                    col_index[canonical] = idx

            for row in rows[rows.index(header_raw) + 1 :]:
                if "pf_no" in col_index:
                    idx = col_index["pf_no"]
                    if idx < len(row) and row[idx] not in (None, ""):
                        global_pf_counts[str(row[idx]).strip()] += 1
                if "hrms" in col_index:
                    idx = col_index["hrms"]
                    if idx < len(row) and row[idx] not in (None, ""):
                        global_hrms_counts[str(row[idx]).strip()] += 1

        for rows, sheet_name, working_at in sources:
            a, u = _import_employee_rows(
                session,
                rows,
                source_label=f"Google Sheet ({sheet_name})",
                working_at_override=working_at,
                warnings=warnings,
                sync_details=sync_details,
                sync_stats=sync_stats,
                global_pf_counts=global_pf_counts,
                global_hrms_counts=global_hrms_counts,
            )
            added += a
            updated += u
        unchanged = sync_stats.get("unchanged", 0)
        skipped = sync_stats.get("skipped", 0)
        if added == 0 and updated == 0 and skipped == 0:
            message_text = "No change found"
        else:
            message_text = f"Google Sheet sync complete: {added} added, {updated} updated, {unchanged} unchanged, {skipped} skipped from {source_label}."
        warning_text = ""
        if warnings:
            preview = "; ".join(warnings[:3])
            if len(warnings) > 3:
                preview += f"; and {len(warnings) - 3} more"
            warning_text = f"Auto-corrected {len(warnings)} date value(s): {preview}"
        if wants_json:
            return JSONResponse(
                {
                    "ok": True,
                    "message": message_text,
                    "warning_message": warning_text,
                    "warning_details": warnings,
                    "sync_details": sync_details,
                }
            )
        message = quote(message_text)
        redirect_url = f"/employees?sync_notice={message}#google-sync-card"
        if warning_text:
            redirect_url = f"/employees?sync_notice={message}&sync_warning={quote(warning_text)}#google-sync-card"
        return RedirectResponse(url=redirect_url, status_code=303)
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
    crew_id: Optional[str] = Form(None),
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
    employee.crew_id = crew_id.strip() if crew_id else None
    employee.dob = to_date(dob)
    employee.doa = to_date(doa)
    employee.do_report = to_date(do_report)
    employee.seniority_rank = to_int(seniority_rank)
    employee.status = status.strip() if status else employee.status
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


def _uploads_context(
    request: Request,
    update_error: Optional[str] = None,
    update_notice: str = "",
    update_warning: str = "",
    update_details: Optional[list[str]] = None,
    warning_details: Optional[list[str]] = None,
    grading_update_error: Optional[str] = None,
    grading_update_notice: str = "",
    grading_update_warning: str = "",
    grading_update_details: Optional[list[str]] = None,
    grading_warning_details: Optional[list[str]] = None,
    cleanup_notice: str = "",
    cleanup_error: Optional[str] = None,
    cleanup_summary: Optional[dict[str, int]] = None,
    cleanup_plan: Optional[list[dict[str, object]]] = None,
    cleanup_conflicts: Optional[list[dict[str, object]]] = None,
    cleanup_details: Optional[list[str]] = None,
    extra_review_notice: str = "",
    extra_review_error: Optional[str] = None,
    extra_review_summary: Optional[dict[str, object]] = None,
    extra_review_groups: Optional[list[dict[str, object]]] = None,
    extra_review_details: Optional[list[str]] = None,
):
    snapshot_ready = EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.exists()
    snapshot_saved_at = ""
    if snapshot_ready:
        try:
            snapshot_saved_at = datetime.fromtimestamp(
                EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.stat().st_mtime
            ).strftime("%d/%m/%Y %I:%M %p")
        except Exception:
            snapshot_saved_at = ""
    cleanup_groups = _group_cleanup_items(cleanup_plan or [], cleanup_conflicts or [])
    return {
        "request": request,
        "active_page": "uploads",
        "role_order": ROLE_ORDER,
        "update_error": update_error,
        "update_notice": update_notice,
        "update_warning": update_warning,
        "update_details": update_details or [],
        "warning_details": warning_details or [],
        "grading_update_error": grading_update_error,
        "grading_update_notice": grading_update_notice,
        "grading_update_warning": grading_update_warning,
        "grading_update_details": grading_update_details or [],
        "grading_warning_details": grading_warning_details or [],
        "cleanup_notice": cleanup_notice,
        "cleanup_error": cleanup_error,
        "cleanup_summary": cleanup_summary or {},
        "cleanup_plan": cleanup_plan or [],
        "cleanup_conflicts": cleanup_conflicts or [],
        "cleanup_groups": cleanup_groups,
        "cleanup_details": cleanup_details or [],
        "extra_review_notice": extra_review_notice,
        "extra_review_error": extra_review_error,
        "extra_review_summary": extra_review_summary or {},
        "extra_review_groups": extra_review_groups or [],
        "extra_review_details": extra_review_details or [],
        "source_snapshot_ready": snapshot_ready,
        "source_snapshot_saved_at": snapshot_saved_at,
    }


@app.get("/uploads")
def uploads_page(request: Request):
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(request),
    )


@app.post("/uploads/employee-master-cleanup-preview")
def preview_employee_master_cleanup(request: Request, session: Session = Depends(get_session)):
    plan, conflicts, summary = _build_combined_cleanup_view(session)
    notice = "No cleanup candidate found" if not plan and not conflicts else ""
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
        ),
    )


@app.post("/uploads/employee-master-cleanup-apply")
def apply_employee_master_cleanup(
    request: Request,
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        plan, conflicts, summary = _build_combined_cleanup_view(session)
        if not plan:
            notice = "No cleanup candidate found"
            return templates.TemplateResponse(
                "uploads.html",
                _uploads_context(
                    request,
                    cleanup_notice=notice,
                    cleanup_summary=summary,
                    cleanup_conflicts=conflicts,
                ),
                status_code=200,
            )

        cleanup_details: list[str] = []
        removed = _apply_duplicate_cleanup_plan(session, plan, cleanup_details)
        notice = f"Smart cleanup complete: {removed} duplicate row(s) deleted."
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(
                request,
                cleanup_notice=notice,
                cleanup_summary={
                    "merge_groups": len(plan),
                    "rows_to_delete": removed,
                    "conflict_groups": len(conflicts),
                },
                cleanup_conflicts=conflicts,
                cleanup_details=cleanup_details,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Smart cleanup failed."
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, cleanup_error=detail),
            status_code=exc.status_code,
        )


@app.post("/uploads/employee-master-cleanup-merge")
def merge_employee_master_conflict(
    request: Request,
    conflict_reason: str = Form(...),
    conflict_row_ids: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        row_ids = [int(value) for value in conflict_row_ids.split(",") if value.strip()]
    except ValueError:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, cleanup_error="Invalid conflict row selection."),
            status_code=400,
        )

    cleanup_details: list[str] = []
    removed = _merge_conflict_rows(
        session,
        reason=conflict_reason,
        row_ids=row_ids,
        details=cleanup_details,
    )
    plan, conflicts, summary = _build_combined_cleanup_view(session)
    notice = (
        f"Manual merge complete: {removed} duplicate row(s) deleted."
        if removed
        else "Manual merge could not be applied."
    )
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
            cleanup_details=cleanup_details,
        ),
    )


@app.post("/uploads/employee-master-cleanup-keep-both")
def keep_both_employee_master_conflict(
    request: Request,
    conflict_reason: str = Form(...),
    conflict_row_ids: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        row_ids = sorted(int(value) for value in conflict_row_ids.split(",") if value.strip())
    except ValueError:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, cleanup_error="Invalid conflict row selection."),
            status_code=400,
        )

    employees = [session.get(Employee, row_id) for row_id in row_ids]
    rows = [employee for employee in employees if employee is not None]
    keep_both_keys = _load_keep_both_decisions()
    if len(rows) >= 2:
        keep_both_keys.add(_cleanup_conflict_key(conflict_reason, rows))
        _save_keep_both_decisions(keep_both_keys)

    plan, conflicts, summary = _build_combined_cleanup_view(session)
    notice = "Conflict marked as keep both."
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
        ),
    )


@app.post("/uploads/employee-master-extra-preview")
def preview_employee_master_extra_rows(request: Request, session: Session = Depends(get_session)):
    groups, summary = _build_employee_master_extra_review(session)
    notice = ""
    if not _load_employee_master_source_snapshot():
        notice = "Upload the latest Service Particulars + CMS files once to review possible extra rows."
    elif not groups:
        notice = "No possible extra rows found."
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary={
                "merge_groups": 0,
                "rows_to_delete": 0,
                "conflict_groups": summary.get("groups", 0),
            },
            cleanup_conflicts=[_extra_group_to_conflict_item(group) for group in groups],
        ),
    )


@app.post("/uploads/employee-master-extra-merge")
def merge_employee_master_extra_rows(
    request: Request,
    review_reason: str = Form(...),
    keep_id: int = Form(...),
    review_row_ids: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        row_ids = [int(value) for value in review_row_ids.split(",") if value.strip()]
    except ValueError:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, extra_review_error="Invalid extra-row selection."),
            status_code=400,
        )

    details: list[str] = []
    removed = _merge_employee_rows(
        session,
        reason=review_reason,
        keep_id=keep_id,
        remove_ids=row_ids,
        details=details,
    )
    plan, conflicts, summary = _build_combined_cleanup_view(session)
    notice = (
        f"Extra row merge complete: {removed} row(s) deleted."
        if removed
        else "Extra row merge could not be applied."
    )
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
            cleanup_details=details,
        ),
    )


@app.post("/uploads/employee-master-extra-delete")
def delete_employee_master_extra_rows(
    request: Request,
    review_reason: str = Form(...),
    review_row_ids: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        row_ids = [int(value) for value in review_row_ids.split(",") if value.strip()]
    except ValueError:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, extra_review_error="Invalid extra-row selection."),
            status_code=400,
        )

    details: list[str] = []
    removed = _delete_employee_rows(
        session,
        reason=review_reason,
        row_ids=row_ids,
        details=details,
    )
    plan, conflicts, summary = _build_combined_cleanup_view(session)
    notice = (
        f"Deleted {removed} extra row(s) from the current table."
        if removed
        else "No extra rows were deleted."
    )
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice=notice,
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
            cleanup_details=details,
        ),
    )


@app.post("/uploads/employee-master-extra-keep")
def keep_employee_master_extra_rows(
    request: Request,
    review_reason: str = Form(...),
    review_row_ids: str = Form(...),
    keep_id: Optional[int] = Form(None),
    session: Session = Depends(get_session),
):
    try:
        row_ids = sorted(int(value) for value in review_row_ids.split(",") if value.strip())
    except ValueError:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, extra_review_error="Invalid extra-row selection."),
            status_code=400,
        )

    keep_keys = _load_string_set(EMPLOYEE_MASTER_EXTRA_REVIEW_KEEP_FILE)
    keep_keys.add(_review_group_key(review_reason, keep_id, row_ids))
    _save_string_set(EMPLOYEE_MASTER_EXTRA_REVIEW_KEEP_FILE, keep_keys)

    plan, conflicts, summary = _build_combined_cleanup_view(session)
    return templates.TemplateResponse(
        "uploads.html",
        _uploads_context(
            request,
            cleanup_notice="Review group marked as keep.",
            cleanup_summary=summary,
            cleanup_plan=plan,
            cleanup_conflicts=conflicts,
        ),
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


def _save_li_grading_metadata(filename: str | None) -> None:
    report_date = infer_report_date(filename or "")
    payload = {
        "filename": filename or "",
        "report_date": report_date.isoformat() if report_date else "",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    LI_GRADING_METADATA_FILE.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_li_grading_metadata() -> dict[str, str]:
    if not LI_GRADING_METADATA_FILE.exists():
        return {"filename": "", "report_date": "", "saved_at": ""}
    try:
        raw = json.loads(LI_GRADING_METADATA_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"filename": "", "report_date": "", "saved_at": ""}
    if not isinstance(raw, dict):
        return {"filename": "", "report_date": "", "saved_at": ""}
    return {
        "filename": str(raw.get("filename") or ""),
        "report_date": str(raw.get("report_date") or ""),
        "saved_at": str(raw.get("saved_at") or ""),
    }


EMPLOYEE_ALIAS_MAP = {
    "name": "name",
    "employeename": "name",
    "empname": "name",
    "staffname": "name",
    "crewname": "name",
    "personname": "name",
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
    "crewid": "crew_id",
    "crewidno": "crew_id",
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


def _attempt_date_string_fix(value: str) -> tuple[date | None, str | None]:
    s = (value or "").strip()
    if not s:
        return None, None
    match = re.match(r"^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{5})\s*$", s)
    if not match:
        return None, None
    day, month, year = match.groups()
    candidates: list[tuple[int, str]] = []
    seen: set[str] = set()
    for idx in range(len(year)):
        trimmed = year[:idx] + year[idx + 1 :]
        if len(trimmed) != 4 or trimmed in seen:
            continue
        seen.add(trimmed)
        try:
            parsed_year = int(trimmed)
        except ValueError:
            continue
        if not (1900 <= parsed_year <= 2100):
            continue
        candidates.append((abs(parsed_year - date.today().year), trimmed))
    candidates.sort(key=lambda item: item[0])
    for _, trimmed in candidates:
        candidate = f"{int(day):02d}/{int(month):02d}/{trimmed}"
        try:
            return datetime.strptime(candidate, "%d/%m/%Y").date(), candidate
        except ValueError:
            continue
    return None, None


def _excel_to_date_with_correction(
    val: object,
    warnings: Optional[list[str]],
    source_label: str,
    row_hint: str,
    field_name: str,
) -> date | None:
    try:
        return _excel_to_date(val)
    except ValueError:
        if isinstance(val, str):
            fixed_date, fixed_text = _attempt_date_string_fix(val)
            if fixed_date is not None:
                if warnings is not None:
                    warnings.append(f"{source_label} {row_hint}: {field_name} {val!r} -> {fixed_text}")
                return fixed_date
        raise


def _format_sync_value(value: object | None) -> str:
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if value is None:
        return "blank"
    text = str(value).strip()
    return text if text else "blank"


def _import_employee_rows(
    session: Session,
    rows: list[tuple | list],
    source_label: str = "sheet",
    working_at_override: Optional[str] = None,
    warnings: Optional[list[str]] = None,
    sync_details: Optional[list[str]] = None,
    sync_stats: Optional[dict[str, int]] = None,
    global_pf_counts: Optional[Counter[str]] = None,
    global_hrms_counts: Optional[Counter[str]] = None,
) -> tuple[int, int]:
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

    required_cols = {"name", "role"}
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

    if global_pf_counts is None:
        global_pf_counts = Counter()
    if global_hrms_counts is None:
        global_hrms_counts = Counter()

    added = 0
    updated = 0
    data_rows = rows[rows.index(header_raw) + 1 :]
    source_pf_counts: Counter[str] = Counter()
    source_hrms_counts: Counter[str] = Counter()
    for row in data_rows:
        if "pf_no" in col_index:
            idx = col_index["pf_no"]
            if idx < len(row) and row[idx] not in (None, ""):
                source_pf_counts[str(row[idx]).strip()] += 1
        if "hrms" in col_index:
            idx = col_index["hrms"]
            if idx < len(row) and row[idx] not in (None, ""):
                source_hrms_counts[str(row[idx]).strip()] += 1

    for row in data_rows:
        def get(col: str) -> object | None:
            idx = col_index.get(col)
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        def has_col(col: str) -> bool:
            return col in col_index

        name = get("name")
        role_raw = get("role")
        if name in (None, "") or role_raw in (None, ""):
            continue
        row_hint = str(name).strip()
        raw_hrms = get("hrms")
        if raw_hrms not in (None, ""):
            row_hint = f"{row_hint} ({str(raw_hrms).strip()})"
        raw_crew_id = get("crew_id")
        if raw_crew_id not in (None, "") and raw_hrms in (None, ""):
            row_hint = f"{row_hint} ({str(raw_crew_id).strip()})"

        try:
            hire_date = _excel_to_date_with_correction(get("hire_date"), warnings, source_label, row_hint, "hire_date") if has_col("hire_date") else None
            retirement_date = _excel_to_date_with_correction(get("retirement_date"), warnings, source_label, row_hint, "retirement_date") if has_col("retirement_date") else None
            promo_ready = _excel_to_date_with_correction(get("promotion_ready_date"), warnings, source_label, row_hint, "promotion_ready_date") if has_col("promotion_ready_date") else None
            dob = _excel_to_date_with_correction(get("dob"), warnings, source_label, row_hint, "dob") if has_col("dob") else None
            doa = _excel_to_date_with_correction(get("doa"), warnings, source_label, row_hint, "doa") if has_col("doa") else None
            do_report = _excel_to_date_with_correction(get("do_report"), warnings, source_label, row_hint, "do_report") if has_col("do_report") else None
            pme_due = _excel_to_date_with_correction(get("pme_due"), warnings, source_label, row_hint, "pme_due") if has_col("pme_due") else None
            technical_due = _excel_to_date_with_correction(get("technical_due"), warnings, source_label, row_hint, "technical_due") if has_col("technical_due") else None
            transportation_due = _excel_to_date_with_correction(get("transportation_due"), warnings, source_label, row_hint, "transportation_due") if has_col("transportation_due") else None
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Date parse error in {source_label}: {exc}") from exc

        role = normalize_role(str(role_raw))
        promo_role = normalize_role(str(get("promotion_role"))) if has_col("promotion_role") and get("promotion_role") else None
        category = str(get("category")).strip() if has_col("category") and get("category") else None
        pf_no = str(get("pf_no")).strip() if has_col("pf_no") and get("pf_no") else None
        hrms = str(get("hrms")).strip() if has_col("hrms") and get("hrms") else None
        crew_id = str(get("crew_id")).strip() if has_col("crew_id") and get("crew_id") else None
        status_val = str(get("status")).strip() if has_col("status") and get("status") else None
        working_at = str(get("working_at")).strip() if has_col("working_at") and get("working_at") else None
        if working_at:
            working_at = " ".join(working_at.split())
        elif working_at_override:
            working_at = working_at_override

        existing = None
        if pf_no and (source_pf_counts.get(pf_no, 0) > 1 or global_pf_counts.get(pf_no, 0) > 1):
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(
                    f"{source_label} {row_hint}: skipped because PF No {pf_no} appears multiple times in the Google Sheet."
                )
            continue
        if hrms and (source_hrms_counts.get(hrms, 0) > 1 or global_hrms_counts.get(hrms, 0) > 1):
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(
                    f"{source_label} {row_hint}: skipped because HRMS {hrms} appears multiple times in the Google Sheet."
                )
            continue
        pf_matches = session.exec(select(Employee).where(Employee.pf_no == pf_no)).all() if pf_no else []
        hrms_matches = session.exec(select(Employee).where(Employee.hrms == hrms)).all() if hrms else []
        if len(pf_matches) > 1:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(
                    f"{source_label} {row_hint}: skipped because PF No {pf_no} matches multiple employees in the current database."
                )
            continue
        if len(hrms_matches) > 1:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(
                    f"{source_label} {row_hint}: skipped because HRMS {hrms} matches multiple employees in the current database."
                )
            continue
        if pf_matches and hrms_matches and pf_matches[0].id != hrms_matches[0].id:
            if sync_stats is not None:
                sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
            if warnings is not None:
                warnings.append(
                    f"{source_label} {row_hint}: skipped because PF No {pf_no} and HRMS {hrms} point to different employees."
                )
            continue
        both_matches = [employee for employee in pf_matches if hrms and employee.hrms == hrms] if pf_matches and hrms else []
        if len(both_matches) == 1:
            existing = both_matches[0]
        elif pf_matches:
            existing = pf_matches[0]
        elif hrms_matches:
            existing = hrms_matches[0]
        else:
            exact_matches = session.exec(
                select(Employee).where(Employee.name == str(name).strip(), Employee.role == role)
            ).all()
            if len(exact_matches) == 1 and not exact_matches[0].pf_no and not exact_matches[0].hrms:
                existing = exact_matches[0]
            elif pf_no is None and hrms is None:
                if sync_stats is not None:
                    sync_stats["skipped"] = sync_stats.get("skipped", 0) + 1
                if warnings is not None:
                    warnings.append(
                        f"{source_label} {row_hint}: skipped because both PF No and HRMS are blank and strict sync requires an identifier."
                    )
                continue

        if hire_date is None and (has_col("hire_date") or has_col("doa") or has_col("dob") or has_col("retirement_date")):
            hire_date = doa or _derive_hire_date(dob, retirement_date)
        if hire_date is None and existing is not None:
            hire_date = existing.hire_date
        if hire_date is None:
            raise HTTPException(status_code=400, detail=f"hire_date missing in {source_label} and could not be derived.")

        if existing:
            retirement_target = retirement_date if has_col("retirement_date") else existing.retirement_date
            promo_role_target = promo_role if has_col("promotion_role") else existing.promotion_role
            promo_ready_target = promo_ready if has_col("promotion_ready_date") else existing.promotion_ready_date
            category_target = category if has_col("category") else existing.category
            pf_no_target = pf_no if has_col("pf_no") else existing.pf_no
            hrms_target = hrms if has_col("hrms") else existing.hrms
            crew_id_target = crew_id if has_col("crew_id") else existing.crew_id
            dob_target = dob if has_col("dob") else existing.dob
            doa_target = doa if has_col("doa") else existing.doa
            do_report_target = do_report if has_col("do_report") else existing.do_report
            status_target = status_val if has_col("status") else existing.status
            working_at_target = working_at if (has_col("working_at") or working_at_override is not None) else existing.working_at
            new_gradation = str(get("gradation")).strip() if has_col("gradation") and get("gradation") else (None if has_col("gradation") else existing.gradation)
            new_cli = str(get("cli")).strip() if has_col("cli") and get("cli") else (None if has_col("cli") else existing.cli)
            pme_due_target = pme_due if has_col("pme_due") else existing.pme_due
            technical_due_target = technical_due if has_col("technical_due") else existing.technical_due
            transportation_due_target = transportation_due if has_col("transportation_due") else existing.transportation_due
            field_updates = [
                ("Name", existing.name, str(name).strip()),
                ("Designation", existing.role, role),
                ("Hire Date", existing.hire_date, hire_date),
                ("Retirement Date", existing.retirement_date, retirement_target),
                ("Promotion Designation", existing.promotion_role, promo_role_target),
                ("Promotion Ready Date", existing.promotion_ready_date, promo_ready_target),
                ("Category", existing.category, category_target),
                ("PF No", existing.pf_no, pf_no_target),
                ("HRMS ID", existing.hrms, hrms_target),
                ("CREW ID", existing.crew_id, crew_id_target),
                ("DOB", existing.dob, dob_target),
                ("DOA", existing.doa, doa_target),
                ("DO Report", existing.do_report, do_report_target),
                ("Status", existing.status, status_target),
                ("Working At", existing.working_at, working_at_target),
                ("Gradation", existing.gradation, new_gradation),
                ("CLI", existing.cli, new_cli),
                ("PME Due", existing.pme_due, pme_due_target),
                ("Technical Due", existing.technical_due, technical_due_target),
                ("Transportation Due", existing.transportation_due, transportation_due_target),
            ]
            changed_fields = [
                f"{label}: {_format_sync_value(old_value)} -> {_format_sync_value(new_value)}"
                for label, old_value, new_value in field_updates
                if old_value != new_value
            ]

            existing.name = str(name).strip()
            existing.role = role
            existing.hire_date = hire_date
            existing.retirement_date = retirement_target
            existing.promotion_role = promo_role_target
            existing.promotion_ready_date = promo_ready_target
            existing.category = category_target
            existing.pf_no = pf_no_target
            existing.hrms = hrms_target
            existing.crew_id = crew_id_target
            existing.dob = dob_target
            existing.doa = doa_target
            existing.do_report = do_report_target
            existing.status = status_target
            existing.working_at = working_at_target
            existing.gradation = new_gradation
            existing.cli = new_cli
            existing.pme_due = pme_due_target
            existing.technical_due = technical_due_target
            existing.transportation_due = transportation_due_target
            if changed_fields:
                updated += 1
                if sync_details is not None:
                    sync_details.append(f"Updated {row_hint}: {'; '.join(changed_fields)}")
            elif sync_stats is not None:
                sync_stats["unchanged"] = sync_stats.get("unchanged", 0) + 1
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
                    crew_id=crew_id,
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
            if sync_details is not None:
                sync_details.append(
                    f"Added {row_hint}: Designation {_format_sync_value(role)}; Working At {_format_sync_value(working_at)}"
                )

    session.commit()
    unchanged = sync_stats.get("unchanged", 0) if sync_stats is not None else 0
    skipped = sync_stats.get("skipped", 0) if sync_stats is not None else 0
    if added == 0 and updated == 0 and unchanged == 0 and skipped == 0:
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


def _sheet_name_key(value: str) -> str:
    value = (value or "").strip()
    if value.startswith("'") and value.endswith("'"):
        value = value[1:-1].replace("''", "'")
    return " ".join(value.split()).casefold()


def _resolve_google_sheet_range(service, spreadsheet_id: str, requested_range: str) -> str:
    raw = (requested_range or "").strip() or "Employees!A:ZZ"
    if "!" in raw:
        requested_name, cell_range = raw.split("!", 1)
    else:
        requested_name, cell_range = raw, "A:ZZ"
    requested_name = requested_name.strip()
    cell_range = cell_range.strip() or "A:ZZ"

    metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [
        sheet.get("properties", {}).get("title", "").strip()
        for sheet in metadata.get("sheets", [])
        if sheet.get("properties", {}).get("title")
    ]
    if not titles:
        raise HTTPException(status_code=400, detail="Google Sheet has no visible tabs.")

    wanted_key = _sheet_name_key(requested_name)
    actual_title = next((title for title in titles if _sheet_name_key(title) == wanted_key), None)
    if not actual_title:
        available = ", ".join(titles)
        raise HTTPException(
            status_code=400,
            detail=f"Google Sheet tab '{requested_name}' was not found. Available tabs: {available}",
        )

    return _normalize_google_sheet_range(f"{actual_title}!{cell_range}")


def _list_google_sheet_titles(service, spreadsheet_id: str) -> list[str]:
    metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return [
        sheet.get("properties", {}).get("title", "").strip()
        for sheet in metadata.get("sheets", [])
        if sheet.get("properties", {}).get("title")
    ]


def _fetch_google_employee_rows() -> tuple[list[tuple[list[list[str]], str, Optional[str]]], str]:
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_EMPLOYEE_SPREADSHEET_ID", "").strip()
    requested_range = os.getenv("GOOGLE_SHEETS_EMPLOYEE_RANGE", "").strip() or "Employees!A:ZZ"
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
        titles = _list_google_sheet_titles(service, spreadsheet_id)
        if not titles:
            raise HTTPException(status_code=400, detail="Google Sheet has no visible tabs.")

        title_map = {_sheet_name_key(title): title for title in titles}
        sources: list[tuple[list[list[str]], str, Optional[str]]] = []

        for station_name in GOOGLE_EMPLOYEE_STATION_TABS:
            actual_title = title_map.get(_sheet_name_key(station_name))
            if not actual_title:
                continue
            sheet_range = _normalize_google_sheet_range(f"{actual_title}!A:ZZ")
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_range,
            ).execute()
            rows = result.get("values", [])
            if rows and len(rows) > 1:
                sources.append((rows, actual_title, station_name))

        if sources:
            return sources, ", ".join(station for _, _, station in sources)

        sheet_range = _resolve_google_sheet_range(service, spreadsheet_id, requested_range)
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
    return [(rows, sheet_range, None)], sheet_range


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


def _run_one_time_employee_master_cleanup(session: Session) -> None:
    if EMPLOYEE_MASTER_SMART_CLEANUP_SENTINEL.exists():
        return
    plan, _, summary = _build_duplicate_cleanup_plan(session)
    details: list[str] = []
    removed = _apply_duplicate_cleanup_plan(session, plan, details) if plan else 0
    EMPLOYEE_MASTER_SMART_CLEANUP_SENTINEL.write_text(
        json.dumps(
            {
                "merge_groups": summary.get("merge_groups", 0),
                "rows_to_delete": summary.get("rows_to_delete", 0),
                "removed": removed,
                "ran_on": datetime.now().isoformat(timespec="seconds"),
            }
        ),
        encoding="utf-8",
    )


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
    grading_meta = _load_li_grading_metadata()
    report_date = coerce_report_date(grading_meta.get("report_date")) or date.today()

    wb = Workbook()
    ws = wb.active
    ws.title = "CLI Distribution"
    ws.append(["CLI", "Gradation A", "Gradation B", "Gradation C", "Total"])
    for row in cli_distribution:
        ws.append([row["cli"], row["A"], row["B"], row["C"], row["total"]])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"cli_distribution_{report_date.isoformat()}.xlsx"
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
    crew_id: Optional[str] = Form(None),
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
        existing.crew_id = crew_id.strip() if crew_id else None
        existing.dob = to_date(dob)
        existing.doa = to_date(doa)
        existing.do_report = to_date(do_report)
        existing.seniority_rank = to_int(seniority_rank)
        existing.status = status.strip() if status else existing.status
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
        crew_id=crew_id.strip() if crew_id else None,
        dob=to_date(dob),
        doa=to_date(doa),
        do_report=to_date(do_report),
        seniority_rank=to_int(seniority_rank),
        status=status.strip() if status else "ACTIVE",
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


def _clean_import_text(value: object | None, *, blank_na: bool = False) -> str | None:
    if value is None:
        return None
    text = " ".join(str(value).strip().split())
    if not text:
        return None
    if text in {"-", "--"}:
        return None
    if blank_na and text.upper() in {"NA", "N/A"}:
        return None
    return text


def _decode_uploaded_text(content: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode("utf-8", errors="replace")


def _normalize_import_name(value: object | None) -> str | None:
    text = _clean_import_text(value)
    if text is None:
        return None
    text = text.upper()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b", " ", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split()) or None


def _emp_no_last5(value: object | None) -> str | None:
    text = _clean_import_text(value)
    if text is None:
        return None
    text = re.sub(r"[^A-Z0-9]", "", text.upper())
    if not text:
        return None
    return text[-5:] if len(text) >= 5 else text


def _find_employee_master_merge_candidate(
    employees: list[Employee],
    *,
    emp_no: str | None,
    name: str | None,
    role: str | None,
    dob: date | None,
) -> Employee | None:
    target_name = _normalize_import_name(name)
    target_last5 = _emp_no_last5(emp_no)
    target_role = normalize_role(role) if role else None

    if target_name and dob:
        candidates = [
            employee
            for employee in employees
            if _normalize_import_name(employee.name) == target_name and employee.dob == dob
        ]
        if len(candidates) == 1:
            candidate = candidates[0]
            candidate_pf = _clean_import_text(candidate.pf_no)
            if candidate_pf is None or emp_no is None:
                return candidate
            if target_last5 and _emp_no_last5(candidate_pf) == target_last5:
                return candidate
            return None

    if target_name and target_role:
        candidates = [
            employee
            for employee in employees
            if _normalize_import_name(employee.name) == target_name
            and normalize_role(employee.role) == target_role
        ]
        if len(candidates) == 1:
            candidate = candidates[0]
            candidate_pf = _clean_import_text(candidate.pf_no)
            if candidate_pf and target_last5 and _emp_no_last5(candidate_pf) == target_last5:
                return candidate

    return None


def _employee_has_value(value: object | None) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def _employee_completeness(employee: Employee) -> int:
    fields = (
        "name",
        "role",
        "hire_date",
        "retirement_date",
        "promotion_ready_date",
        "category",
        "pf_no",
        "hrms",
        "crew_id",
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
    )
    return sum(1 for field in fields if _employee_has_value(getattr(employee, field)))


def _employee_cleanup_sort_key(employee: Employee) -> tuple[int, int, int]:
    has_crew_id = 1 if _employee_has_value(employee.crew_id) else 0
    return (-has_crew_id, -_employee_completeness(employee), employee.id or 0)


def _dedupe_uploaded_employee_rows(session: Session, sync_details: list[str]) -> int:
    groups: dict[tuple[str, str, str], list[Employee]] = {}
    for employee in session.exec(select(Employee)).all():
        name_key = _normalize_import_name(employee.name)
        pf_key = _clean_import_text(employee.pf_no)
        dob_key = employee.dob.isoformat() if employee.dob else None
        if not name_key or not pf_key or not dob_key:
            continue
        groups.setdefault((name_key, dob_key, pf_key), []).append(employee)

    removed = 0
    merge_fields = (
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
        "promotion_ready_date",
        "category",
        "hrms",
        "crew_id",
        "doa",
        "do_report",
        "status",
        "gradation",
        "cli",
        "pme_due",
        "technical_due",
        "transportation_due",
    )

    for employees in groups.values():
        if len(employees) < 2:
            continue

        by_working_at: dict[str, list[Employee]] = {}
        for employee in employees:
            working_at_key = (_clean_import_text(employee.working_at) or "").upper()
            by_working_at.setdefault(working_at_key, []).append(employee)

        for working_at_group in by_working_at.values():
            if len(working_at_group) < 2:
                continue

            ordered = sorted(working_at_group, key=_employee_cleanup_sort_key)
            keeper = ordered[0]
            merged_count = 0

            for duplicate in ordered[1:]:
                for field_name in merge_fields:
                    if not _employee_has_value(getattr(keeper, field_name)) and _employee_has_value(
                        getattr(duplicate, field_name)
                    ):
                        setattr(keeper, field_name, getattr(duplicate, field_name))
                session.delete(duplicate)
                removed += 1
                merged_count += 1

            if merged_count:
                location = _clean_import_text(keeper.working_at) or "blank working_at"
                sync_details.append(
                    f"Deduplicated {keeper.name}: kept 1 row for EMP NO {_format_sync_value(keeper.pf_no)} at {location}; removed {merged_count} duplicate row(s)."
                )

    return removed


def _working_at_key(value: object | None) -> str:
    return (_clean_import_text(value) or "").upper()


def _one_working_at_blank(first: object | None, second: object | None) -> bool:
    first_key = _working_at_key(first)
    second_key = _working_at_key(second)
    return (not first_key and bool(second_key)) or (bool(first_key) and not second_key)


def _cleanup_row_payload(employee: Employee) -> dict[str, object]:
    return {
        "id": employee.id,
        "name": employee.name,
        "designation": employee.role,
        "dob": employee.dob.strftime("%d/%m/%Y") if employee.dob else "",
        "hire_date": employee.hire_date.strftime("%d/%m/%Y") if employee.hire_date else "",
        "retirement_date": employee.retirement_date.strftime("%d/%m/%Y") if employee.retirement_date else "",
        "emp_no": employee.pf_no or "",
        "working_at": employee.working_at or "",
        "crew_id": employee.crew_id or "",
        "hrms": employee.hrms or "",
        "category": employee.category or "",
        "gradation": employee.gradation or "",
        "cli": employee.cli or "",
    }


def _cleanup_conflict_key(reason: str, rows: list[Employee]) -> str:
    row_ids = ",".join(str(employee.id or 0) for employee in sorted(rows, key=lambda item: item.id or 0))
    return f"{reason}|{row_ids}"


def _load_string_set(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return set()
    if not isinstance(raw, list):
        return set()
    return {str(item) for item in raw if item}


def _save_string_set(path: Path, keys: set[str]) -> None:
    path.write_text(
        json.dumps(sorted(keys), ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_keep_both_decisions() -> set[str]:
    return _load_string_set(EMPLOYEE_MASTER_KEEP_BOTH_FILE)


def _save_keep_both_decisions(keys: set[str]) -> None:
    _save_string_set(EMPLOYEE_MASTER_KEEP_BOTH_FILE, keys)


def _review_group_key(reason: str, keep_id: int | None, review_ids: list[int]) -> str:
    review_string = ",".join(str(row_id) for row_id in sorted(review_ids))
    return f"{reason}|{keep_id or 0}|{review_string}"


def _serialize_employee_master_snapshot(records: dict[str, dict[str, object]]) -> list[dict[str, object]]:
    payload: list[dict[str, object]] = []
    for emp_no, record in sorted(records.items()):
        payload.append(
            {
                "row_hint": str(record.get("row_hint") or ""),
                "name": _clean_import_text(record.get("name")) or "",
                "role": _clean_import_text(record.get("role")) or "",
                "pf_no": emp_no,
                "crew_id": _clean_import_text(record.get("crew_id")) or "",
                "dob": record.get("dob").isoformat() if isinstance(record.get("dob"), date) else "",
                "category": _clean_import_text(record.get("category"), blank_na=True) or "",
            }
        )
    return payload


def _save_employee_master_source_snapshot(records: dict[str, dict[str, object]]) -> None:
    EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.write_text(
        json.dumps(_serialize_employee_master_snapshot(records), ensure_ascii=True, indent=2),
        encoding="utf-8",
    )


def _load_employee_master_source_snapshot() -> list[dict[str, object]]:
    if not EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.exists():
        return []
    try:
        raw = json.loads(EMPLOYEE_MASTER_SOURCE_SNAPSHOT_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []
    if not isinstance(raw, list):
        return []
    return [item for item in raw if isinstance(item, dict)]


def _build_duplicate_cleanup_plan(session: Session) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, int]]:
    employees = session.exec(select(Employee)).all()
    plan: list[dict[str, object]] = []
    conflicts: list[dict[str, object]] = []
    used_ids: set[int] = set()
    seen_conflicts: set[tuple[int, ...]] = set()
    keep_both_keys = _load_keep_both_decisions()

    def register_plan(reason: str, rows: list[Employee]) -> None:
        if len(rows) < 2:
            return
        ordered = sorted(rows, key=_employee_cleanup_sort_key)
        keeper = ordered[0]
        remove_rows = ordered[1:]
        used_ids.update(employee.id for employee in ordered if employee.id is not None)
        plan.append(
            {
                "reason": reason,
                "keep": _cleanup_row_payload(keeper),
                "remove": [_cleanup_row_payload(employee) for employee in remove_rows],
            }
        )

    def register_conflict(reason: str, rows: list[Employee]) -> None:
        if len(rows) < 2:
            return
        row_ids = tuple(sorted(employee.id or 0 for employee in rows))
        if row_ids in seen_conflicts:
            return
        seen_conflicts.add(row_ids)
        key_string = _cleanup_conflict_key(reason, rows)
        if key_string in keep_both_keys:
            return
        ordered_rows = sorted(rows, key=_employee_cleanup_sort_key)
        conflicts.append(
            {
                "reason": reason,
                "conflict_key": key_string,
                "row_ids": [employee.id for employee in sorted(rows, key=lambda item: item.id or 0)],
                "suggested_keep_id": ordered_rows[0].id if ordered_rows else None,
                "rows": [_cleanup_row_payload(employee) for employee in sorted(rows, key=lambda item: item.id or 0)],
            }
        )

    def has_dob_mismatch(rows: list[Employee]) -> bool:
        dob_keys = {employee.dob.isoformat() for employee in rows if employee.dob}
        return len(dob_keys) > 1

    by_name_dob: dict[tuple[str, str], list[Employee]] = {}
    for employee in employees:
        name_key = _normalize_import_name(employee.name)
        dob_key = employee.dob.isoformat() if employee.dob else None
        if not name_key or not dob_key:
            continue
        by_name_dob.setdefault((name_key, dob_key), []).append(employee)

    for rows in by_name_dob.values():
        if len(rows) < 2:
            continue
        register_plan("Same Name + DOB", rows)

    by_name_crew: dict[tuple[str, str], list[Employee]] = {}
    for employee in employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        name_key = _normalize_import_name(employee.name)
        crew_key = _clean_import_text(employee.crew_id)
        if not name_key or not crew_key:
            continue
        by_name_crew.setdefault((name_key, crew_key), []).append(employee)

    for rows in by_name_crew.values():
        if len(rows) < 2:
            continue
        if has_dob_mismatch(rows):
            register_conflict("Same Name + same CREW ID but DOB differs", rows)
            continue
        register_plan("Same Name + same CREW ID", rows)

    by_dob_last5: dict[tuple[str, str], list[Employee]] = {}
    for employee in employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        dob_key = employee.dob.isoformat() if employee.dob else None
        last5_key = _emp_no_last5(employee.pf_no)
        if not dob_key or not last5_key:
            continue
        by_dob_last5.setdefault((dob_key, last5_key), []).append(employee)

    for rows in by_dob_last5.values():
        if len(rows) < 2:
            continue
        register_plan("Same DOB + EMP NO last 5 match", rows)

    by_name_last5: dict[tuple[str, str], list[Employee]] = {}
    for employee in employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        name_key = _normalize_import_name(employee.name)
        last5_key = _emp_no_last5(employee.pf_no)
        if not name_key or not last5_key:
            continue
        by_name_last5.setdefault((name_key, last5_key), []).append(employee)

    for rows in by_name_last5.values():
        if len(rows) < 2:
            continue

        working_groups: dict[str, list[Employee]] = {}
        for employee in rows:
            working_groups.setdefault(_working_at_key(employee.working_at), []).append(employee)

        blank_group = working_groups.get("", [])
        filled_groups = [group for key, group in working_groups.items() if key]
        if not blank_group:
            continue

        if has_dob_mismatch(rows):
            register_conflict("Same Name + EMP NO last 5 match but DOB differs", rows)
            continue

        if len(filled_groups) == 1:
            register_plan("Same Name + EMP NO last 5 match and one Working At is blank", rows)
        else:
            register_conflict("Same Name + EMP NO last 5 match but Working At differs", rows)

    by_name_role: dict[tuple[str, str], list[Employee]] = {}
    for employee in employees:
        if employee.id is not None and employee.id in used_ids:
            continue
        name_key = _normalize_import_name(employee.name)
        role_key = normalize_role(employee.role) if employee.role else None
        if not name_key or not role_key:
            continue
        by_name_role.setdefault((name_key, role_key), []).append(employee)

    for rows in by_name_role.values():
        if len(rows) < 2:
            continue

        by_working_last5: dict[tuple[str, str], list[Employee]] = {}
        for employee in rows:
            last5_key = _emp_no_last5(employee.pf_no)
            if not last5_key:
                continue
            by_working_last5.setdefault((_working_at_key(employee.working_at), last5_key), []).append(employee)

        by_last5_all_working: dict[str, set[str]] = {}
        for working_key, last5_key in by_working_last5.keys():
            by_last5_all_working.setdefault(last5_key, set()).add(working_key)

        for last5_key, working_keys in by_last5_all_working.items():
            if len(working_keys) > 1:
                conflict_rows = [
                    employee
                    for employee in rows
                    if _emp_no_last5(employee.pf_no) == last5_key
                ]
                register_conflict("Same Name + Designation + EMP NO last 5 match but Working At differs", conflict_rows)

        for group_rows in by_working_last5.values():
            if len(group_rows) > 1:
                if has_dob_mismatch(group_rows):
                    register_conflict("Same Name + Designation + EMP NO last 5 match but DOB differs", group_rows)
                else:
                    register_plan("Same Name + Designation + EMP NO last 5 match", group_rows)

    conflicts = [
        item
        for item in conflicts
        if not any(row_id and row_id in used_ids for row_id in item["row_ids"])
    ]

    summary = {
        "merge_groups": len(plan),
        "rows_to_delete": sum(len(item["remove"]) for item in plan),
        "conflict_groups": len(conflicts),
    }
    return plan, conflicts, summary


def _apply_duplicate_cleanup_plan(
    session: Session,
    plan: list[dict[str, object]],
    details: list[str],
) -> int:
    merge_fields = (
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
        "promotion_ready_date",
        "category",
        "hrms",
        "crew_id",
        "doa",
        "do_report",
        "status",
        "working_at",
        "gradation",
        "cli",
        "pme_due",
        "technical_due",
        "transportation_due",
    )
    removed = 0

    for item in plan:
        keep_id = item["keep"]["id"]
        remove_ids = [row["id"] for row in item["remove"]]
        keeper = session.get(Employee, keep_id) if keep_id is not None else None
        if keeper is None:
            continue

        merged_count = 0
        for duplicate_id in remove_ids:
            duplicate = session.get(Employee, duplicate_id) if duplicate_id is not None else None
            if duplicate is None:
                continue
            for field_name in merge_fields:
                if not _employee_has_value(getattr(keeper, field_name)) and _employee_has_value(getattr(duplicate, field_name)):
                    setattr(keeper, field_name, getattr(duplicate, field_name))
            session.delete(duplicate)
            removed += 1
            merged_count += 1

        if merged_count:
            details.append(
                f"{item['reason']}: kept {keeper.name} ({_format_sync_value(keeper.pf_no)}), removed {merged_count} duplicate row(s)."
            )

    session.commit()
    return removed


def _merge_conflict_rows(
    session: Session,
    *,
    reason: str,
    row_ids: list[int],
    details: list[str],
) -> int:
    rows = [session.get(Employee, row_id) for row_id in row_ids]
    employees = [row for row in rows if row is not None]
    if len(employees) < 2:
        return 0

    ordered = sorted(employees, key=_employee_cleanup_sort_key)
    keeper = ordered[0]
    removed = 0
    merge_fields = (
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
        "promotion_ready_date",
        "category",
        "hrms",
        "crew_id",
        "doa",
        "do_report",
        "status",
        "working_at",
        "gradation",
        "cli",
        "pme_due",
        "technical_due",
        "transportation_due",
    )

    for duplicate in ordered[1:]:
        for field_name in merge_fields:
            if not _employee_has_value(getattr(keeper, field_name)) and _employee_has_value(getattr(duplicate, field_name)):
                setattr(keeper, field_name, getattr(duplicate, field_name))
        session.delete(duplicate)
        removed += 1

    if removed:
        details.append(
            f"Manual merge applied for {reason}: kept {keeper.name} ({_format_sync_value(keeper.pf_no)}), removed {removed} duplicate row(s)."
        )
    session.commit()
    return removed


def _merge_employee_rows(
    session: Session,
    *,
    reason: str,
    keep_id: int,
    remove_ids: list[int],
    details: list[str],
) -> int:
    keeper = session.get(Employee, keep_id)
    if keeper is None:
        return 0

    merge_fields = (
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
        "promotion_ready_date",
        "category",
        "hrms",
        "crew_id",
        "doa",
        "do_report",
        "status",
        "working_at",
        "gradation",
        "cli",
        "pme_due",
        "technical_due",
        "transportation_due",
    )
    removed = 0

    for duplicate_id in remove_ids:
        duplicate = session.get(Employee, duplicate_id)
        if duplicate is None or duplicate is keeper:
            continue
        for field_name in merge_fields:
            if not _employee_has_value(getattr(keeper, field_name)) and _employee_has_value(getattr(duplicate, field_name)):
                setattr(keeper, field_name, getattr(duplicate, field_name))
        session.delete(duplicate)
        removed += 1

    if removed:
        details.append(
            f"{reason}: kept {keeper.name} ({_format_sync_value(keeper.pf_no)}), removed {removed} extra row(s)."
        )
    session.commit()
    return removed


def _delete_employee_rows(
    session: Session,
    *,
    reason: str,
    row_ids: list[int],
    details: list[str],
) -> int:
    removed = 0
    for row_id in row_ids:
        employee = session.get(Employee, row_id)
        if employee is None:
            continue
        details.append(
            f"{reason}: deleted {employee.name} ({_format_sync_value(employee.pf_no)}) from the current table."
        )
        session.delete(employee)
        removed += 1
    session.commit()
    return removed


def _build_employee_master_extra_review(session: Session) -> tuple[list[dict[str, object]], dict[str, object]]:
    snapshot = _load_employee_master_source_snapshot()
    if not snapshot:
        return [], {
            "groups": 0,
            "review_rows": 0,
            "mergeable_groups": 0,
            "db_only_groups": 0,
            "reason_counts": [],
        }

    employees = session.exec(select(Employee)).all()
    keep_keys = _load_string_set(EMPLOYEE_MASTER_EXTRA_REVIEW_KEEP_FILE)

    source_by_pf: dict[str, dict[str, object]] = {}
    source_by_crew: dict[str, dict[str, object]] = {}
    for row in snapshot:
        pf_value = _clean_import_text(row.get("pf_no"))
        crew_value = _clean_import_text(row.get("crew_id"))
        if pf_value and pf_value not in source_by_pf:
            source_by_pf[pf_value] = row
        if crew_value and crew_value not in source_by_crew:
            source_by_crew[crew_value] = row

    represented_ids: set[int] = set()
    represented_rows: list[Employee] = []
    for employee in employees:
        pf_value = _clean_import_text(employee.pf_no)
        crew_value = _clean_import_text(employee.crew_id)
        if (pf_value and pf_value in source_by_pf) or (crew_value and crew_value in source_by_crew):
            if employee.id is not None:
                represented_ids.add(employee.id)
            represented_rows.append(employee)

    by_name_crew_keep: dict[tuple[str, str], list[Employee]] = {}
    by_name_last5_keep: dict[tuple[str, str], list[Employee]] = {}
    by_dob_last5_keep: dict[tuple[str, str], list[Employee]] = {}
    for employee in represented_rows:
        name_key = _normalize_import_name(employee.name)
        crew_key = _clean_import_text(employee.crew_id)
        last5_key = _emp_no_last5(employee.pf_no)
        if name_key and crew_key:
            by_name_crew_keep.setdefault((name_key, crew_key), []).append(employee)
        if name_key and last5_key:
            by_name_last5_keep.setdefault((name_key, last5_key), []).append(employee)
        if employee.dob and last5_key:
            by_dob_last5_keep.setdefault((employee.dob.isoformat(), last5_key), []).append(employee)

    groups: list[dict[str, object]] = []

    def add_group(reason: str, keep_row: Employee | None, review_rows: list[Employee]) -> None:
        if not review_rows:
            return
        ordered_review = sorted(review_rows, key=_employee_cleanup_sort_key)
        review_ids = [employee.id for employee in ordered_review if employee.id is not None]
        if not review_ids:
            return
        keep_id = keep_row.id if keep_row is not None else None
        group_key = _review_group_key(reason, keep_id, review_ids)
        if group_key in keep_keys:
            return
        groups.append(
            {
                "reason": reason,
                "group_key": group_key,
                "keep": _cleanup_row_payload(keep_row) if keep_row is not None else None,
                "review_rows": [_cleanup_row_payload(employee) for employee in ordered_review],
                "row_ids": review_ids,
                "can_merge": keep_row is not None,
            }
        )

    for employee in employees:
        if employee.id is None or employee.id in represented_ids:
            continue
        name_key = _normalize_import_name(employee.name)
        crew_key = _clean_import_text(employee.crew_id)
        last5_key = _emp_no_last5(employee.pf_no)
        matched = False

        if name_key and crew_key:
            keep_matches = by_name_crew_keep.get((name_key, crew_key), [])
            if len(keep_matches) == 1:
                keep_row = keep_matches[0]
                reason = "Possible extra row: same Name + same CREW ID"
                if employee.dob and keep_row.dob and employee.dob != keep_row.dob:
                    reason = "Possible extra row: same Name + same CREW ID but DOB differs"
                add_group(reason, keep_row, [employee])
                matched = True

        if matched:
            continue

        if employee.dob and last5_key:
            keep_matches = by_dob_last5_keep.get((employee.dob.isoformat(), last5_key), [])
            if len(keep_matches) == 1:
                keep_row = keep_matches[0]
                add_group("Possible extra row: same DOB + EMP NO last 5 match", keep_row, [employee])
                matched = True

        if matched:
            continue

        if name_key and last5_key:
            keep_matches = by_name_last5_keep.get((name_key, last5_key), [])
            if len(keep_matches) == 1:
                keep_row = keep_matches[0]
                if _one_working_at_blank(employee.working_at, keep_row.working_at):
                    reason = "Possible extra row: same Name + EMP NO last 5 match and one Working At is blank"
                    if employee.dob and keep_row.dob and employee.dob != keep_row.dob:
                        reason = "Possible extra row: same Name + EMP NO last 5 match but DOB differs"
                    add_group(reason, keep_row, [employee])
                    matched = True

        if matched:
            continue

        add_group("Only in current DB, no latest source match", None, [employee])

    reason_counts = Counter(group["reason"] for group in groups)
    summary = {
        "groups": len(groups),
        "review_rows": sum(len(group["review_rows"]) for group in groups),
        "mergeable_groups": sum(1 for group in groups if group["can_merge"]),
        "db_only_groups": sum(1 for group in groups if not group["can_merge"]),
        "reason_counts": [{"reason": reason, "count": count} for reason, count in reason_counts.most_common()],
    }
    return groups, summary


def _extra_group_to_conflict_item(group: dict[str, object]) -> dict[str, object]:
    keep_row = group.get("keep")
    review_rows = list(group.get("review_rows") or [])
    rows = []
    suggested_keep_id = None
    if isinstance(keep_row, dict):
        rows.append(keep_row)
        suggested_keep_id = keep_row.get("id")
    rows.extend(review_rows)
    return {
        "reason": group.get("reason", "Possible extra row"),
        "row_ids": list(group.get("row_ids") or []),
        "suggested_keep_id": suggested_keep_id,
        "rows": rows,
        "merge_action": "/uploads/employee-master-extra-merge" if suggested_keep_id else "",
        "delete_action": "/uploads/employee-master-extra-delete",
        "keep_action": "/uploads/employee-master-extra-keep",
        "keep_button_label": "Keep",
        "keep_id": suggested_keep_id,
        "allow_merge": bool(suggested_keep_id),
        "allow_delete": True,
    }


def _build_combined_cleanup_view(session: Session) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, int]]:
    base_plan, base_conflicts, _ = _build_duplicate_cleanup_plan(session)

    plan: list[dict[str, object]] = list(base_plan)
    conflicts: list[dict[str, object]] = []
    covered_review_ids: set[int] = set()

    for item in base_plan:
        if isinstance(item.get("keep"), dict) and item["keep"].get("id") is not None:
            covered_review_ids.add(int(item["keep"]["id"]))
        for row in item.get("remove", []):
            if isinstance(row, dict) and row.get("id") is not None:
                covered_review_ids.add(int(row["id"]))

    for item in base_conflicts:
        row_ids = [int(row_id) for row_id in item.get("row_ids", []) if row_id is not None]
        covered_review_ids.update(row_ids)
        conflicts.append(
            {
                **item,
                "merge_action": "/uploads/employee-master-cleanup-merge",
                "delete_action": "",
                "keep_action": "/uploads/employee-master-cleanup-keep-both",
                "keep_button_label": "Keep Both",
                "keep_id": item.get("suggested_keep_id"),
                "allow_merge": True,
                "allow_delete": False,
            }
        )

    extra_groups, _ = _build_employee_master_extra_review(session)
    for group in extra_groups:
        review_ids = [int(row_id) for row_id in group.get("row_ids", []) if row_id is not None]
        if any(row_id in covered_review_ids for row_id in review_ids):
            continue
        covered_review_ids.update(review_ids)
        keep_row = group.get("keep")
        reason = str(group.get("reason") or "")
        if keep_row and "DOB differs" not in reason:
            plan.append(
                {
                    "reason": reason,
                    "keep": keep_row,
                    "remove": list(group.get("review_rows") or []),
                }
            )
            continue
        conflicts.append(_extra_group_to_conflict_item(group))

    summary = {
        "merge_groups": len(plan),
        "rows_to_delete": sum(len(item.get("remove", [])) for item in plan),
        "conflict_groups": len(conflicts),
    }
    return plan, conflicts, summary


def _cleanup_item_summary(item: dict[str, object], kind: str) -> str:
    if kind == "plan":
        keep = item.get("keep") or {}
        keep_name = str(keep.get("name") or "Unknown")
        remove_count = len(item.get("remove", []))
        return f"{keep_name} - keep 1, delete {remove_count}"
    rows = list(item.get("rows") or [])
    names = [str(row.get("name") or "Unknown") for row in rows[:3]]
    more = max(len(rows) - len(names), 0)
    suffix = f" +{more} more" if more else ""
    return f"{', '.join(names)}{suffix}"


def _cleanup_group_names(items: list[dict[str, object]], kind: str) -> list[str]:
    names: list[str] = []
    for item in items:
        if kind == "plan":
            keep = item.get("keep") or {}
            remove = list(item.get("remove") or [])
            row_names = [str(keep.get("name") or "").strip()] + [str(row.get("name") or "").strip() for row in remove]
        else:
            row_names = [str(row.get("name") or "").strip() for row in list(item.get("rows") or [])]
        for name in row_names:
            if name and name not in names:
                names.append(name)
    return names


def _group_cleanup_items(plan: list[dict[str, object]], conflicts: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[str, dict[str, object]] = {}

    def add_item(reason: str, kind: str, item: dict[str, object]) -> None:
        group = grouped.setdefault(
            reason,
            {
                "reason": reason,
                "plan_items": [],
                "conflict_items": [],
            },
        )
        key = "plan_items" if kind == "plan" else "conflict_items"
        item_copy = dict(item)
        item_copy["summary"] = _cleanup_item_summary(item, kind)
        search_names = _cleanup_group_names([item], kind)
        item_copy["search_text"] = " ".join(search_names).lower()
        group[key].append(item_copy)

    for item in plan:
        add_item(str(item.get("reason") or "Auto merge"), "plan", item)
    for item in conflicts:
        add_item(str(item.get("reason") or "Conflict"), "conflict", item)

    output: list[dict[str, object]] = []
    for reason, group in grouped.items():
        all_items = list(group["plan_items"]) + list(group["conflict_items"])
        names = _cleanup_group_names(group["plan_items"], "plan") + [
            name for name in _cleanup_group_names(group["conflict_items"], "conflict")
            if name not in _cleanup_group_names(group["plan_items"], "plan")
        ]
        output.append(
            {
                "reason": reason,
                "item_count": len(all_items),
                "names": names,
                "items": all_items,
            }
        )
    output.sort(key=lambda item: (item["reason"].lower(), item["item_count"]))
    return output


def _cleanup_employee_master_duplicates_for_record(
    employees: list[Employee],
    target: Employee,
    *,
    emp_no: str | None,
    name: str | None,
    role: str | None,
    dob: date | None,
    sync_details: list[str],
    session: Session,
) -> int:
    target_name = _normalize_import_name(name)
    target_role = normalize_role(role) if role else None
    target_last5 = _emp_no_last5(emp_no)
    target_working_at = _working_at_key(target.working_at)
    removed = 0

    merge_fields = (
        "role",
        "hire_date",
        "retirement_date",
        "promotion_role",
        "promotion_ready_date",
        "category",
        "hrms",
        "crew_id",
        "doa",
        "do_report",
        "status",
        "working_at",
        "gradation",
        "cli",
        "pme_due",
        "technical_due",
        "transportation_due",
    )

    duplicates: list[tuple[Employee, str]] = []
    for employee in list(employees):
        if employee is target:
            continue

        candidate_pf = _clean_import_text(employee.pf_no)
        candidate_name = _normalize_import_name(employee.name)
        same_working_at = _working_at_key(employee.working_at) == target_working_at
        blank_vs_value_working_at = _one_working_at_blank(employee.working_at, target.working_at)
        candidate_last5 = _emp_no_last5(candidate_pf)

        if dob and target_last5 and employee.dob == dob and candidate_last5 == target_last5:
            duplicates.append((employee, "Same DOB + EMP NO last 5"))
            continue

        if target_name and dob and candidate_name == target_name and employee.dob == dob:
            if same_working_at and (
                candidate_pf is None or emp_no is None or (target_last5 and candidate_last5 == target_last5)
            ):
                duplicates.append((employee, "Same Name + DOB"))
                continue
            if blank_vs_value_working_at and target_last5 and candidate_pf and candidate_last5 == target_last5:
                duplicates.append((employee, "Same Name + DOB and one Working At is blank"))
                continue

        if not same_working_at:
            continue

        if target_name and target_role and candidate_name == target_name and normalize_role(employee.role) == target_role:
            if candidate_pf and target_last5 and candidate_last5 == target_last5:
                duplicates.append((employee, "Same Name + Designation"))

    for duplicate, reason in duplicates:
        for field_name in merge_fields:
            if not _employee_has_value(getattr(target, field_name)) and _employee_has_value(getattr(duplicate, field_name)):
                setattr(target, field_name, getattr(duplicate, field_name))
        session.delete(duplicate)
        if duplicate in employees:
            employees.remove(duplicate)
        removed += 1
        sync_details.append(
            f"Deduplicated {target.name}: removed duplicate row by {reason} at {_clean_import_text(target.working_at) or 'blank working_at'}."
        )

    return removed


def _build_service_particular_records(
    content: bytes,
    warnings: list[str],
) -> tuple[dict[str, dict[str, object]], dict[str, str]]:
    wb = load_workbook(filename=BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(status_code=400, detail="Service Particulars workbook is empty.")

    header_row = next((row for row in rows if any(cell not in (None, "", " ") for cell in row)), None)
    if header_row is None:
        raise HTTPException(status_code=400, detail="Service Particulars workbook has no header row.")

    header = {_employee_norm(cell): idx for idx, cell in enumerate(header_row) if cell not in (None, "")}
    required = {
        "crewname": "CREW NAME",
        "crewdesg": "CREW DESG",
        "crewid": "CREW ID",
        "empno": "EMP NO",
        "birthdate": "BIRTH DATE",
        "appointdate": "APPOINT DATE",
        "retirementdate": "RETIREMENT DATE",
    }
    missing = [label for key, label in required.items() if key not in header]
    if missing:
        raise HTTPException(status_code=400, detail=f"Service Particulars is missing columns: {', '.join(missing)}")

    records: dict[str, dict[str, object]] = {}
    crew_to_emp: dict[str, str] = {}
    duplicate_emp: set[str] = set()
    duplicate_crew: set[str] = set()

    for row in rows[rows.index(header_row) + 1 :]:
        if not any(cell not in (None, "", " ") for cell in row):
            continue

        emp_no = _clean_import_text(row[header["empno"]])
        crew_id = _clean_import_text(row[header["crewid"]])
        name = _clean_import_text(row[header["crewname"]])
        role_raw = _clean_import_text(row[header["crewdesg"]])
        row_hint = name or crew_id or emp_no or "Unknown row"

        if not emp_no:
            warnings.append(f"Service Particulars {row_hint}: skipped because EMP NO is blank.")
            continue
        if emp_no in records:
            duplicate_emp.add(emp_no)
            continue
        if crew_id and crew_id in crew_to_emp:
            duplicate_crew.add(crew_id)
            continue

        try:
            dob = _excel_to_date_with_correction(row[header["birthdate"]], warnings, "Service Particulars", row_hint, "birth_date")
            hire_date = _excel_to_date_with_correction(row[header["appointdate"]], warnings, "Service Particulars", row_hint, "appoint_date")
            retirement_date = _excel_to_date_with_correction(row[header["retirementdate"]], warnings, "Service Particulars", row_hint, "retirement_date")
            promotion_ready = None
            if "promotiondate" in header:
                promotion_ready = _excel_to_date_with_correction(row[header["promotiondate"]], warnings, "Service Particulars", row_hint, "promotion_date")
        except Exception as exc:
            warnings.append(f"Service Particulars {row_hint}: skipped because {exc}.")
            continue

        role = normalize_role(role_raw or "")
        if not name or not role or not hire_date:
            warnings.append(f"Service Particulars {row_hint}: skipped because name, designation, or appoint date is missing.")
            continue

        records[emp_no] = {
            "row_hint": row_hint,
            "name": name,
            "role": role,
            "pf_no": emp_no,
            "crew_id": crew_id,
            "dob": dob,
            "hire_date": hire_date,
            "doa": hire_date,
            "retirement_date": retirement_date,
            "promotion_ready_date": promotion_ready,
            "present_fields": {
                "name",
                "role",
                "pf_no",
                "crew_id",
                "dob",
                "hire_date",
                "doa",
                "retirement_date",
                "promotion_ready_date",
            },
        }
        if crew_id:
            crew_to_emp[crew_id] = emp_no

    for emp_no in sorted(duplicate_emp):
        warnings.append(f"Service Particulars duplicate EMP NO skipped: {emp_no}")
        records.pop(emp_no, None)
    for crew_id in sorted(duplicate_crew):
        emp_no = crew_to_emp.get(crew_id)
        if emp_no:
            records.pop(emp_no, None)
        warnings.append(f"Service Particulars duplicate CREW ID skipped: {crew_id}")

    if not records:
        raise HTTPException(status_code=400, detail="Service Particulars did not produce any usable employee rows.")
    return records, crew_to_emp


def _merge_cms_other_bio(
    records: dict[str, dict[str, object]],
    crew_to_emp: dict[str, str],
    content: bytes,
    warnings: list[str],
) -> None:
    text = _decode_uploaded_text(content)
    reader = csv.DictReader(StringIO(text))
    if not reader.fieldnames:
        raise HTTPException(status_code=400, detail="CMS other bio data file has no header row.")

    normalized_header = {_employee_norm(name): name for name in reader.fieldnames if name}
    if "crewid" not in normalized_header:
        raise HTTPException(status_code=400, detail="CMS other bio data is missing column: CREWID")

    seen_hrms: set[str] = set()
    for row in reader:
        hrms_key = normalized_header["crewid"]
        hrms = _clean_import_text(row.get(hrms_key))
        row_hint = _clean_import_text(row.get(normalized_header.get("crewname", hrms_key))) or hrms or "Unknown row"
        if not hrms:
            warnings.append("CMS other bio data row skipped because CREWID is blank.")
            continue
        if hrms in seen_hrms:
            warnings.append(f"CMS other bio data duplicate CREWID skipped: {hrms}")
            continue
        seen_hrms.add(hrms)

        emp_no = crew_to_emp.get(hrms)
        if not emp_no or emp_no not in records:
            warnings.append(f"CMS other bio data {row_hint} ({hrms}): no matching Service Particulars row found.")
            continue

        record = records[emp_no]
        present_fields: set[str] = record["present_fields"]  # type: ignore[assignment]
        name = _clean_import_text(row.get(normalized_header.get("crewname", "")))
        if name:
            record["name"] = name
            record["row_hint"] = f"{name} ({hrms})"
            present_fields.add("name")

        role_raw = _clean_import_text(row.get(normalized_header.get("desig", "")))
        if role_raw:
            record["role"] = normalize_role(role_raw)
            present_fields.add("role")

        category = _clean_import_text(row.get(normalized_header.get("category", "")), blank_na=True)
        if "category" in normalized_header:
            record["category"] = category
            present_fields.add("category")

        try:
            retirement_raw = row.get(normalized_header["retirementdate"]) if "retirementdate" in normalized_header else None
            if _clean_import_text(retirement_raw) is not None:
                record["retirement_date"] = _excel_to_date_with_correction(
                    retirement_raw,
                    warnings,
                    "CMS other bio data",
                    str(record["row_hint"]),
                    "retirement_date",
                )
                present_fields.add("retirement_date")
            appoint_raw = row.get(normalized_header["appointmentdate"]) if "appointmentdate" in normalized_header else None
            if _clean_import_text(appoint_raw) is not None:
                appoint_date = _excel_to_date_with_correction(
                    appoint_raw,
                    warnings,
                    "CMS other bio data",
                    str(record["row_hint"]),
                    "appointment_date",
                )
                record["hire_date"] = appoint_date
                record["doa"] = appoint_date
                present_fields.update({"hire_date", "doa"})
            pme_raw = row.get(normalized_header["pmedue"]) if "pmedue" in normalized_header else None
            if _clean_import_text(pme_raw) is not None:
                record["pme_due"] = _excel_to_date_with_correction(
                    pme_raw,
                    warnings,
                    "CMS other bio data",
                    str(record["row_hint"]),
                    "pme_due",
                )
                present_fields.add("pme_due")
        except Exception as exc:
            warnings.append(f"CMS other bio data {record['row_hint']}: {exc}")


def _upsert_employee_master_records(
    session: Session,
    records: dict[str, dict[str, object]],
    warnings: list[str],
    sync_details: list[str],
) -> tuple[int, int, int, int, int]:
    added = 0
    updated = 0
    unchanged = 0
    skipped = 0
    deduplicated = 0
    employees = session.exec(select(Employee)).all()

    by_pf: dict[str, list[Employee]] = {}
    by_crew_id: dict[str, list[Employee]] = {}

    def rebuild_exact_indexes() -> None:
        by_pf.clear()
        by_crew_id.clear()
        for employee in employees:
            pf_value = _clean_import_text(employee.pf_no)
            if pf_value:
                by_pf.setdefault(pf_value, []).append(employee)
            crew_value = _clean_import_text(employee.crew_id)
            if crew_value:
                by_crew_id.setdefault(crew_value, []).append(employee)

    rebuild_exact_indexes()

    for emp_no, record in records.items():
        row_hint = str(record.get("row_hint") or record.get("name") or emp_no)
        name = _clean_import_text(record.get("name"))
        role = _clean_import_text(record.get("role"))
        hire_date = record.get("hire_date")
        present_fields = set(record.get("present_fields") or set())

        if not name or not role or not isinstance(hire_date, date):
            warnings.append(f"{row_hint}: skipped because name, designation, or appoint date is missing after merge.")
            skipped += 1
            continue

        pf_matches = by_pf.get(emp_no, [])
        if len(pf_matches) > 1:
            warnings.append(f"{row_hint}: skipped because EMP NO {emp_no} matches multiple employees in the current database.")
            skipped += 1
            continue

        existing = pf_matches[0] if pf_matches else None
        crew_id = _clean_import_text(record.get("crew_id"))
        if existing is None and crew_id:
            crew_matches = by_crew_id.get(crew_id, [])
            if len(crew_matches) > 1:
                warnings.append(f"{row_hint}: skipped because CREW ID {crew_id} matches multiple employees in the current database.")
                skipped += 1
                continue
            if len(crew_matches) == 1:
                existing = crew_matches[0]
                if existing.pf_no and existing.pf_no != emp_no:
                    warnings.append(f"{row_hint}: skipped because EMP NO {emp_no} conflicts with existing employee EMP NO {existing.pf_no}.")
                    skipped += 1
                    continue

        if existing is None:
            existing = _find_employee_master_merge_candidate(
                employees,
                emp_no=emp_no,
                name=name,
                role=role,
                dob=record.get("dob") if isinstance(record.get("dob"), date) else None,
            )

        if existing and crew_id:
            crew_conflicts = [employee for employee in by_crew_id.get(crew_id, []) if employee is not existing]
            if crew_conflicts:
                warnings.append(f"{row_hint}: skipped because CREW ID {crew_id} already belongs to another employee.")
                skipped += 1
                continue

        record_values = {
            "name": name,
            "role": normalize_role(role),
            "hire_date": hire_date,
            "doa": record.get("doa"),
            "retirement_date": record.get("retirement_date"),
            "promotion_ready_date": record.get("promotion_ready_date"),
            "category": record.get("category"),
            "pf_no": emp_no,
            "crew_id": crew_id,
            "dob": record.get("dob"),
            "pme_due": record.get("pme_due"),
        }

        if existing:
            field_labels = {
                "name": "Name",
                "role": "Designation",
                "hire_date": "Hire Date",
                "doa": "DOA",
                "retirement_date": "Retirement Date",
                "promotion_ready_date": "Promotion Date",
                "category": "Category",
                "pf_no": "EMP NO",
                "crew_id": "CREW ID",
                "dob": "DOB",
                "pme_due": "PME Due",
            }
            changed_fields: list[str] = []
            for field_name, label in field_labels.items():
                if field_name not in present_fields and field_name != "pf_no":
                    continue
                old_value = getattr(existing, field_name)
                new_value = record_values[field_name]
                if old_value != new_value:
                    changed_fields.append(
                        f"{label}: {_format_sync_value(old_value)} -> {_format_sync_value(new_value)}"
                    )
                    setattr(existing, field_name, new_value)
            if changed_fields:
                updated += 1
                sync_details.append(f"Updated {row_hint}: {'; '.join(changed_fields)}")
            else:
                unchanged += 1
            deduplicated += _cleanup_employee_master_duplicates_for_record(
                employees,
                existing,
                emp_no=emp_no,
                name=name,
                role=role,
                dob=record.get("dob") if isinstance(record.get("dob"), date) else None,
                sync_details=sync_details,
                session=session,
            )
            rebuild_exact_indexes()
        else:
            employee = Employee(
                name=record_values["name"],
                role=record_values["role"],
                hire_date=record_values["hire_date"],
                retirement_date=record_values["retirement_date"],
                promotion_ready_date=record_values["promotion_ready_date"],
                category=record_values["category"],
                pf_no=record_values["pf_no"],
                crew_id=record_values["crew_id"],
                dob=record_values["dob"],
                doa=record_values["doa"],
                pme_due=record_values["pme_due"],
                status="ACTIVE",
            )
            session.add(employee)
            employees.append(employee)
            added += 1
            sync_details.append(
                f"Added {row_hint}: EMP NO {_format_sync_value(emp_no)}; CREW ID {_format_sync_value(crew_id)}"
            )
            deduplicated += _cleanup_employee_master_duplicates_for_record(
                employees,
                employee,
                emp_no=emp_no,
                name=name,
                role=role,
                dob=record.get("dob") if isinstance(record.get("dob"), date) else None,
                sync_details=sync_details,
                session=session,
            )
            rebuild_exact_indexes()

    deduplicated += _dedupe_uploaded_employee_rows(session, sync_details)
    session.commit()
    return added, updated, unchanged, skipped, deduplicated


@app.post("/uploads/employee-master-sync")
async def upload_employee_master_sync(
    request: Request,
    service_file: UploadFile = File(...),
    cms_file: UploadFile = File(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        service_name = service_file.filename or ""
        cms_name = cms_file.filename or ""
        if not service_name.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Service Particulars file must be an .xlsx workbook.")
        if not cms_name.lower().endswith(".csv"):
            raise HTTPException(status_code=400, detail="CMS other bio data file must be a .csv file.")

        service_content = await service_file.read()
        cms_content = await cms_file.read()
        warnings: list[str] = []
        sync_details: list[str] = []

        records, hrms_to_emp = _build_service_particular_records(service_content, warnings)
        _merge_cms_other_bio(records, hrms_to_emp, cms_content, warnings)
        _save_employee_master_source_snapshot(records)
        added, updated, unchanged, skipped, deduplicated = _upsert_employee_master_records(
            session,
            records,
            warnings,
            sync_details,
        )

        if added == 0 and updated == 0 and skipped == 0 and deduplicated == 0:
            notice = "No change found"
        else:
            notice = (
                "Employee table update complete: "
                f"{added} added, {updated} updated, {unchanged} unchanged, {skipped} skipped, {deduplicated} deduplicated."
            )
        warning_text = ""
        if warnings:
            warning_text = f"Mismatch / auto-fixed records: {len(warnings)}"

        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(
                request,
                update_notice=notice,
                update_warning=warning_text,
                update_details=sync_details,
                warning_details=warnings,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "Employee table update failed."
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, update_error=detail),
            status_code=exc.status_code,
        )
    except Exception as exc:
        return templates.TemplateResponse(
            "uploads.html",
            _uploads_context(request, update_error=str(exc)),
            status_code=500,
        )


def _normalize_li_grading_header(value: object | None) -> str:
    text = _clean_import_text(value)
    if text is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "", text.upper())


def _parse_li_grading_workbook(content: bytes) -> tuple[list[dict[str, object]], list[str]]:
    workbook = load_workbook(filename=BytesIO(content), data_only=True)
    worksheet = workbook.active
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(status_code=400, detail="LI Grading workbook is empty.")

    header_row_index: int | None = None
    crew_idx: int | None = None
    name_idx: int | None = None
    role_idx: int | None = None
    current_grade_idx: int | None = None
    due_date_idx: int | None = None

    for idx, row in enumerate(rows):
        normalized = [_normalize_li_grading_header(cell) for cell in row]
        if "CREWID" not in normalized or "NAME" not in normalized or "CURRENTGRADE" not in normalized:
            continue
        role_idx = next((i for i, value in enumerate(normalized) if value in {"DESIG", "DESIGNATION", "ROLE"}), None)
        if role_idx is None:
            continue
        current_grade_idx = normalized.index("CURRENTGRADE")
        due_date_idx = next((i for i in range(current_grade_idx + 1, len(normalized)) if normalized[i] == "DUEDATE"), None)
        if due_date_idx is None:
            due_date_idx = next((i for i, value in enumerate(normalized) if value == "DUEDATE"), None)
        if due_date_idx is None:
            continue
        crew_idx = normalized.index("CREWID")
        name_idx = normalized.index("NAME")
        header_row_index = idx
        break

    if header_row_index is None or None in {crew_idx, name_idx, role_idx, current_grade_idx, due_date_idx}:
        raise HTTPException(
            status_code=400,
            detail="Could not find the LI Grading columns. Required columns: CREW ID, NAME, DESIG., CURRENT GRADE, DUE DATE.",
        )

    warnings: list[str] = []
    records: list[dict[str, object]] = []

    for row_number, row in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        def get(column_index: int | None) -> object | None:
            if column_index is None or column_index >= len(row):
                return None
            return row[column_index]

        crew_id = _clean_import_text(get(crew_idx))
        name = _clean_import_text(get(name_idx))
        role_raw = _clean_import_text(get(role_idx))
        current_grade = _clean_import_text(get(current_grade_idx))
        due_raw = get(due_date_idx)

        if not any([crew_id, name, role_raw, current_grade, due_raw]):
            continue

        row_hint = name or crew_id or f"row {row_number}"
        if not name:
            warnings.append(f"LI Grading row {row_number}: skipped because NAME is blank.")
            continue
        if not role_raw:
            warnings.append(f"LI Grading {row_hint}: skipped because DESIG. is blank.")
            continue
        if not current_grade:
            warnings.append(f"LI Grading {row_hint}: skipped because CURRENT GRADE is blank.")
            continue

        try:
            due_date = _excel_to_date_with_correction(due_raw, warnings, "LI Grading", row_hint, "due_date") if due_raw not in (None, "") else None
        except ValueError as exc:
            warnings.append(f"LI Grading {row_hint}: skipped because DUE DATE is invalid ({exc}).")
            continue

        records.append(
            {
                "crew_id": crew_id,
                "name": name,
                "role": normalize_role(role_raw),
                "gradation": current_grade.upper(),
                "grading_due": due_date,
                "row_hint": row_hint,
            }
        )

    if not records:
        raise HTTPException(status_code=400, detail="LI Grading workbook did not produce any usable rows.")
    return records, warnings


@app.post("/upload-li-grading")
async def upload_li_grading(
    request: Request,
    file: UploadFile = File(...),
    action_password: str = Form(...),
    session: Session = Depends(get_session),
):
    try:
        _validate_sensitive_action_password(action_password)
        filename = file.filename or ""
        if not filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Upload the LI Grading .xlsx workbook.")

        records, warnings = _parse_li_grading_workbook(await file.read())
        employees = session.exec(select(Employee)).all()

        by_crew: dict[str, list[Employee]] = {}
        by_name_role: dict[tuple[str, str], list[Employee]] = {}
        for employee in employees:
            crew_key = (_clean_import_text(employee.crew_id) or "").upper()
            if crew_key:
                by_crew.setdefault(crew_key, []).append(employee)
            name_key = _normalize_import_name(employee.name)
            role_key = normalize_role(employee.role)
            if name_key and role_key:
                by_name_role.setdefault((name_key, role_key), []).append(employee)

        updated = 0
        unchanged = 0
        skipped = 0
        details: list[str] = []
        touched_ids: set[int] = set()

        for record in records:
            row_hint = str(record["row_hint"])
            crew_key = str(record.get("crew_id") or "").upper()
            name_key = _normalize_import_name(record.get("name"))
            role_key = str(record.get("role") or "")
            target: Employee | None = None

            if crew_key:
                crew_matches = by_crew.get(crew_key, [])
                filtered_matches = [
                    employee
                    for employee in crew_matches
                    if _normalize_import_name(employee.name) == name_key and normalize_role(employee.role) == role_key
                ]
                if len(filtered_matches) == 1:
                    target = filtered_matches[0]
                elif len(filtered_matches) > 1:
                    warnings.append(f"LI Grading {row_hint}: skipped because CREW ID, NAME, and DESIGNATION matched multiple roster rows.")
                    skipped += 1
                    continue
                elif crew_matches:
                    warnings.append(f"LI Grading {row_hint}: skipped because CREW ID {crew_key} matched the roster but NAME / DESIGNATION did not match.")
                    skipped += 1
                    continue

            if target is None:
                if not name_key or not role_key:
                    warnings.append(f"LI Grading {row_hint}: no matching CLI Roster row found.")
                    skipped += 1
                    continue
                fallback_matches = by_name_role.get((name_key, role_key), [])
                if len(fallback_matches) == 1:
                    target = fallback_matches[0]
                elif len(fallback_matches) > 1:
                    warnings.append(f"LI Grading {row_hint}: skipped because NAME + DESIGNATION matched multiple CLI Roster rows.")
                    skipped += 1
                    continue
                else:
                    warnings.append(f"LI Grading {row_hint}: no matching CLI Roster row found.")
                    skipped += 1
                    continue

            if target.id is not None and target.id in touched_ids:
                warnings.append(f"LI Grading {row_hint}: skipped because that roster row already received a grading update from another row in this workbook.")
                skipped += 1
                continue

            new_grade = _clean_import_text(record.get("gradation"))
            new_due = record.get("grading_due")
            old_grade = _clean_import_text(target.gradation)
            old_due = target.grading_due

            if old_grade == new_grade and old_due == new_due:
                unchanged += 1
                if target.id is not None:
                    touched_ids.add(target.id)
                continue

            changes: list[str] = []
            if old_grade != new_grade:
                changes.append(f"Gradation: {_format_sync_value(old_grade)} -> {_format_sync_value(new_grade)}")
            if old_due != new_due:
                changes.append(f"Grading Due: {_format_sync_value(old_due)} -> {_format_sync_value(new_due)}")

            target.gradation = new_grade
            target.grading_due = new_due
            updated += 1
            if target.id is not None:
                touched_ids.add(target.id)
            details.append(
                f"Updated {target.name} ({target.crew_id or target.hrms or target.id}): " + "; ".join(changes)
            )

        session.commit()
        _save_li_grading_metadata(filename)
        if updated == 0 and unchanged > 0 and skipped == 0:
            notice = "No change found in LI grading file."
        else:
            notice_parts = []
            if updated:
                notice_parts.append(f"{updated} updated")
            if skipped:
                notice_parts.append(f"{skipped} skipped")
            if not notice_parts:
                notice_parts.append("No change found")
            notice = "LI grading update complete: " + ", ".join(notice_parts) + "."
        warning_message = f"Mismatch / auto-fixed records: {len(warnings)}" if warnings else ""
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_notice=notice,
                grading_update_warning=warning_message,
                grading_update_details=details,
                grading_warning_details=warnings,
            ),
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else "LI grading update failed."
        return templates.TemplateResponse(
            "cli.html",
            _cli_page_context(
                request,
                session,
                grading_update_error=detail,
            ),
            status_code=exc.status_code,
        )


@app.post("/upload")
async def upload_employees(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=400,
            detail="Upload an .xlsx file with columns: name, designation (role), hire_date, retirement_date. Optional: promotion designation (promotion_role), promotion_ready_date, category, pf_no, hrms, dob, doa, do_report, working_at.",
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
        raise HTTPException(status_code=400, detail="Upload an .xlsx file with columns: name, designation (role), seniority_rank (or seniority). Optional: promotion designation (promotion_role), promotion_ready_date.")

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
        raise HTTPException(status_code=400, detail="No matching employees updated. Ensure names/designations match the roster.")
    return RedirectResponse("/", status_code=303)
