from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
import math
from io import BytesIO
import json
import os
from pathlib import Path
from typing import Optional
import re
from urllib import error as urlerror
from urllib import request as urlrequest

from fastapi import Depends, FastAPI, Form, Request, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook, Workbook
from sqlmodel import Session, select
from starlette.middleware.base import BaseHTTPMiddleware

from .db import get_session, init_db
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
from .models import Employee, Requirement, SstsDeviceSnapshot, SstsSnapshotRun
from .seed import seed_all

BASE_PATH = Path(__file__).resolve().parent.parent
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
ASSET_VER = "v20260321c"
templates.env.globals["asset_ver"] = ASSET_VER
ASSET_VER = "v20260321b"
templates.env.globals["asset_ver"] = ASSET_VER


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
SSTS_WEB_URL = "http://164.52.197.129/devices"
SSTS_API_BASE_URL = "http://164.52.197.129:3000"
SSTS_API_LOGIN_URL = f"{SSTS_API_BASE_URL}/auth/login/"
SSTS_API_DEVICE_URL = f"{SSTS_API_BASE_URL}/device"
SSTS_API_USER = os.getenv("SSTS_API_USER", "srdeeopsdah@gmail.com")
SSTS_API_PASSWORD = os.getenv("SSTS_API_PASSWORD", "sdah1234")
SSTS_OFFLINE_THRESHOLD_MINUTES = 120
SSTS_RECENTLY_ONLINE_THRESHOLD_MINUTES = 5
SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES = 300
SSTS_RECENT_OFFLINE_MAX_MINUTES = 24 * 60
SSTS_REFRESH_INTERVAL_MINUTES = 30


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


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _ensure_utc(dt: datetime | None) -> datetime | None:
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def _parse_ssts_timestamp(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(timezone.utc)
    except ValueError:
        return None


def _minutes_since(now_utc: datetime, lastupdate: datetime | None) -> int | None:
    if not lastupdate:
        return None
    diff = now_utc - _ensure_utc(lastupdate)
    return max(0, int(diff.total_seconds() // 60))


def _ssts_is_offline(snapshot: SstsDeviceSnapshot) -> bool:
    return (snapshot.offline_minutes or 0) > SSTS_OFFLINE_THRESHOLD_MINUTES


def _ssts_is_recently_offline(snapshot: SstsDeviceSnapshot) -> bool:
    minutes = snapshot.offline_minutes or 0
    return SSTS_OFFLINE_THRESHOLD_MINUTES < minutes < SSTS_RECENT_OFFLINE_MAX_MINUTES


def _ssts_is_online_now(snapshot: SstsDeviceSnapshot) -> bool:
    minutes = snapshot.offline_minutes
    if minutes is None:
        return False
    return minutes <= SSTS_RECENTLY_ONLINE_THRESHOLD_MINUTES


def _snapshot_to_row(snapshot: SstsDeviceSnapshot) -> dict[str, object]:
    return {
        "device_id": snapshot.device_id,
        "name": snapshot.name,
        "uniqueid": snapshot.uniqueid or "",
        "phone": snapshot.phone or "",
        "contact": snapshot.contact or "",
        "lastupdate": snapshot.lastupdate,
        "lastupdate_label": _ensure_utc(snapshot.lastupdate).strftime("%d-%m-%Y %H:%M UTC")
        if _ensure_utc(snapshot.lastupdate)
        else "No signal",
        "offline_minutes": snapshot.offline_minutes,
        "offline_hours": round((snapshot.offline_minutes or 0) / 60, 1) if snapshot.offline_minutes is not None else None,
    }


def _ssts_sort_key(snapshot: SstsDeviceSnapshot) -> tuple[int, int, str]:
    return (
        -(snapshot.offline_minutes or -1),
        snapshot.device_id,
        snapshot.name.lower(),
    )


def _ssts_post_json(url: str, payload: dict[str, object], headers: dict[str, str] | None = None) -> dict[str, object]:
    body = json.dumps(payload).encode("utf-8")
    request_headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    if headers:
        request_headers.update(headers)
    req = urlrequest.Request(url, data=body, headers=request_headers)
    with urlrequest.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8", "replace"))


def _ssts_get_json(url: str, headers: dict[str, str] | None = None) -> object:
    request_headers = {"User-Agent": "Mozilla/5.0"}
    if headers:
        request_headers.update(headers)
    req = urlrequest.Request(url, headers=request_headers)
    with urlrequest.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8", "replace"))


def fetch_ssts_devices() -> list[dict[str, object]]:
    login_payload = {"username": SSTS_API_USER, "password": SSTS_API_PASSWORD}
    login_data = _ssts_post_json(SSTS_API_LOGIN_URL, login_payload)
    token = str(login_data.get("token") or "").strip()
    if not token:
        raise RuntimeError("SSTS login succeeded but token was missing.")
    devices = _ssts_get_json(SSTS_API_DEVICE_URL, headers={"Authorization": token})
    if not isinstance(devices, list):
        raise RuntimeError("Unexpected SSTS device response format.")
    return [item for item in devices if isinstance(item, dict)]


def refresh_ssts_snapshot(session: Session, force: bool = False) -> dict[str, object]:
    now_utc = _utc_now()
    latest_run = session.exec(select(SstsSnapshotRun).order_by(SstsSnapshotRun.observed_at.desc())).first()
    if latest_run and not force:
        latest_run_time = _ensure_utc(latest_run.observed_at)
        age_minutes = int((now_utc - latest_run_time).total_seconds() // 60)
        if age_minutes < SSTS_REFRESH_INTERVAL_MINUTES and latest_run.fetch_status == "ok":
            return {
                "status": "cached",
                "observed_at": latest_run.observed_at,
                "source_count": latest_run.source_count,
                "message": f"Using last sync from {latest_run_time.strftime('%d-%m-%Y %H:%M UTC')}.",
            }
    observed_at = now_utc.replace(second=0, microsecond=0)
    try:
        devices = fetch_ssts_devices()
        run = SstsSnapshotRun(
            observed_at=observed_at,
            source_count=len(devices),
            fetch_status="ok",
        )
        session.add(run)
        session.commit()
        session.refresh(run)
        snapshots: list[SstsDeviceSnapshot] = []
        for device in devices:
            lastupdate = _parse_ssts_timestamp(str(device.get("lastupdate") or ""))
            snapshots.append(
                SstsDeviceSnapshot(
                    run_id=run.id or 0,
                    observed_at=observed_at,
                    observed_day=observed_at.date(),
                    device_id=int(device.get("id") or 0),
                    name=str(device.get("name") or "Unknown"),
                    uniqueid=str(device.get("uniqueid") or "") or None,
                    phone=str(device.get("phone") or "") or None,
                    contact=str(device.get("contact") or "") or None,
                    lastupdate=lastupdate,
                    offline_minutes=_minutes_since(observed_at, lastupdate),
                    attributes=str(device.get("attributes") or "") or None,
                )
            )
        session.add_all(snapshots)
        session.commit()
        return {
            "status": "fetched",
            "observed_at": observed_at,
            "source_count": len(snapshots),
            "message": f"Fetched {len(snapshots)} rakes from SSTS.",
        }
    except (urlerror.URLError, HTTPException, RuntimeError, ValueError, json.JSONDecodeError) as exc:
        run = SstsSnapshotRun(
            observed_at=observed_at,
            source_count=0,
            fetch_status="error",
            fetch_error=str(exc),
        )
        session.add(run)
        session.commit()
        return {
            "status": "error",
            "observed_at": observed_at,
            "source_count": 0,
            "message": f"SSTS sync failed: {exc}",
        }


def _distinct_ssts_runs(session: Session) -> list[SstsSnapshotRun]:
    return list(session.exec(select(SstsSnapshotRun).order_by(SstsSnapshotRun.observed_at.desc())).all())


def _snapshots_for_run(session: Session, run_id: int) -> list[SstsDeviceSnapshot]:
    snapshots = session.exec(
        select(SstsDeviceSnapshot)
        .where(SstsDeviceSnapshot.run_id == run_id)
        .order_by(SstsDeviceSnapshot.name)
    ).all()
    return list(snapshots)


def build_ssts_report_context(session: Session) -> dict[str, object]:
    runs = _distinct_ssts_runs(session)
    latest_run = next((run for run in runs if run.fetch_status == "ok"), None)
    if not latest_run:
        return {
            "latest_run": None,
            "latest_rows": [],
            "current_offline": [],
            "current_recently_offline": [],
            "previous_day_offline": [],
            "previous_day_recently_offline": [],
            "recently_online": [],
            "daily_summary": [],
        }

    latest_snapshots = _snapshots_for_run(session, latest_run.id or 0)
    latest_map = {row.device_id: row for row in latest_snapshots}
    latest_run_time = _ensure_utc(latest_run.observed_at)
    previous_day_cutoff = latest_run_time - timedelta(days=1)
    previous_day_run = next(
        (run for run in runs if run.fetch_status == "ok" and _ensure_utc(run.observed_at) <= previous_day_cutoff),
        None,
    )
    previous_day_snapshots = _snapshots_for_run(session, previous_day_run.id or 0) if previous_day_run else []

    current_offline = [_snapshot_to_row(row) for row in sorted(latest_snapshots, key=_ssts_sort_key) if _ssts_is_offline(row)]
    current_recently_offline = [
        _snapshot_to_row(row)
        for row in sorted(latest_snapshots, key=_ssts_sort_key)
        if _ssts_is_recently_offline(row)
    ]
    previous_day_offline = [
        _snapshot_to_row(row)
        for row in sorted(previous_day_snapshots, key=_ssts_sort_key)
        if _ssts_is_offline(row)
    ]
    previous_day_recently_offline = [
        _snapshot_to_row(row)
        for row in sorted(previous_day_snapshots, key=_ssts_sort_key)
        if _ssts_is_recently_offline(row)
    ]

    latest_by_day: dict[date, SstsSnapshotRun] = {}
    for run in runs:
        if run.fetch_status != "ok":
            continue
        latest_by_day.setdefault(run.observed_at.date(), run)
    daily_summary = []
    for day, run in sorted(latest_by_day.items(), key=lambda item: item[0], reverse=True)[:7]:
        rows = _snapshots_for_run(session, run.id or 0)
        offline_count = sum(1 for row in rows if _ssts_is_offline(row))
        recent_offline_count = sum(1 for row in rows if _ssts_is_recently_offline(row))
        daily_summary.append(
            {
                "day": day.strftime("%d-%m-%Y"),
                "observed_at": _ensure_utc(run.observed_at).strftime("%d-%m-%Y %H:%M UTC"),
                "total_rakes": len(rows),
                "offline_count": offline_count,
                "recent_offline_count": recent_offline_count,
            }
        )

    recovery_runs = [run for run in runs if run.fetch_status == "ok" and _ensure_utc(run.observed_at) <= latest_run_time]
    snapshots_by_run = {run.id: _snapshots_for_run(session, run.id or 0) for run in recovery_runs}
    history_by_device: dict[int, list[SstsDeviceSnapshot]] = {}
    for run in sorted(recovery_runs, key=lambda item: item.observed_at):
        for row in snapshots_by_run.get(run.id, []):
            history_by_device.setdefault(row.device_id, []).append(row)

    recently_online = []
    for device_id, latest_row in latest_map.items():
        if not _ssts_is_online_now(latest_row):
            continue
        history = history_by_device.get(device_id, [])
        if len(history) < 2:
            continue
        previous_row = history[-2]
        if (previous_row.offline_minutes or 0) > SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES:
            row = _snapshot_to_row(latest_row)
            row["previous_offline_hours"] = round((previous_row.offline_minutes or 0) / 60, 1)
            row["previous_seen"] = _ensure_utc(previous_row.observed_at).strftime("%d-%m-%Y %H:%M UTC")
            recently_online.append(row)
    recently_online.sort(key=lambda row: (-float(row.get("previous_offline_hours") or 0), str(row.get("name") or "").lower()))

    return {
        "latest_run": latest_run,
        "latest_rows": [_snapshot_to_row(row) for row in sorted(latest_snapshots, key=_ssts_sort_key)],
        "current_offline": current_offline,
        "current_recently_offline": current_recently_offline,
        "previous_day_run": previous_day_run,
        "previous_day_offline": previous_day_offline,
        "previous_day_recently_offline": previous_day_recently_offline,
        "recently_online": recently_online,
        "daily_summary": daily_summary,
    }


@app.on_event("startup")
def on_startup() -> None:
    init_db()
    session = next(get_session())
    try:
        seed_all(session)
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
        },
    )


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


@app.get("/ssts-report")
def ssts_report_page(
    request: Request,
    refresh: int = 0,
    session: Session = Depends(get_session),
):
    sync_state = refresh_ssts_snapshot(session, force=bool(refresh))
    context = build_ssts_report_context(session)
    latest_run = context.get("latest_run")
    latest_summary = {
        "total_rakes": len(context.get("latest_rows", [])),
        "offline_count": len(context.get("current_offline", [])),
        "recently_offline_count": len(context.get("current_recently_offline", [])),
        "recently_online_count": len(context.get("recently_online", [])),
    }
    return templates.TemplateResponse(
        "ssts_report.html",
        {
            "request": request,
            "active_page": "ssts_report",
            "ssts_web_url": SSTS_WEB_URL,
            "sync_state": sync_state,
            "latest_run": latest_run,
            "latest_summary": latest_summary,
            **context,
        },
    )


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
    if not rows:
        raise HTTPException(status_code=400, detail="Workbook is empty.")

    # Normalize headers and map common aliases (to support varied spreadsheets)
    def norm(val: object | None) -> str:
        return "".join(ch for ch in str(val).lower() if ch.isalnum()) if val is not None else ""

    # find first non-empty row to use as header
    header_raw = None
    for r in rows:
        if any(cell not in (None, "", " ") for cell in r):
            header_raw = r
            break
    if header_raw is None:
        raise HTTPException(status_code=400, detail="Workbook appears empty (no header row).")

    header_norm = [norm(h) for h in header_raw]
    alias_map = {
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

    mapped_cols: list[str] = []
    for h in header_norm:
        mapped_cols.append(alias_map.get(h, ""))

    col_index: dict[str, int] = {}
    for idx, canonical in enumerate(mapped_cols):
        if canonical and canonical not in col_index:
            col_index[canonical] = idx

    required_cols = {"name", "role", "retirement_date"}
    missing_required = required_cols - set(col_index)

    # Fallback: known North sheet positional layout when header row is missing but data present
    if missing_required:
        first_row = rows[rows.index(header_raw)]
        if isinstance(first_row[0], (int, float)) and isinstance(first_row[1], str) and len(first_row) >= 14:
            # assume order: SL, Name, Degn, PF, HRMS, Gender, Category, Gradation, mob, WhatsApp, LOBBY, CLI, DOB, DOR, PME, Technical, Transportation, Tr10...
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
        raise HTTPException(status_code=400, detail=f"Missing columns: {', '.join(sorted(missing_required))}")
    added = 0
    updated = 0
    def derive_hire_date(dob_val: date | None, retirement_val: date | None) -> date | None:
        if dob_val:
            try:
                return dob_val.replace(year=dob_val.year + 25)
            except ValueError:
                # Feb 29 safety
                return dob_val.replace(month=2, day=28, year=dob_val.year + 25)
        if retirement_val:
            return retirement_val - timedelta(days=35 * 365)
        return None

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
            raise HTTPException(status_code=400, detail=f"Date parse error: {exc}") from exc

        if retirement_date is None:
            raise HTTPException(status_code=400, detail="retirement_date is required in the sheet.")
        if hire_date is None:
            hire_date = derive_hire_date(dob, retirement_date)
        if hire_date is None:
            raise HTTPException(status_code=400, detail="hire_date missing and could not be derived (need hire_date or dob).")

        role = normalize_role(str(role_raw))
        promo_role = normalize_role(str(get("promotion_role"))) if "promotion_role" in col_index else None
        category = str(get("category")).strip() if "category" in col_index and get("category") else None
        pf_no = str(get("pf_no")).strip() if "pf_no" in col_index and get("pf_no") else None
        hrms = str(get("hrms")).strip() if "hrms" in col_index and get("hrms") else None
        status_val = str(get("status")).strip() if "status" in col_index and get("status") else None
        working_at = str(get("working_at")).strip() if "working_at" in col_index and get("working_at") else None

        # upsert by (name, role) to prevent duplicates
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
        raise HTTPException(status_code=400, detail="No rows imported. Check the sheet data or headers.")
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
