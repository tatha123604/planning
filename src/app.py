from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
import math
from io import BytesIO
import json
import os
from pathlib import Path
from typing import Optional
import re
import threading
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest
from uuid import uuid4

from fastapi import Depends, FastAPI, Form, Request, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, PlainTextResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
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
ASSET_VER = "v20260512b"
templates.env.globals["asset_ver"] = ASSET_VER


def format_cli_label(cli_name: str | None, cli_id: str | None = None) -> str:
    name = (cli_name or "").strip()
    cli_id_clean = (cli_id or "").strip()
    if name and cli_id_clean:
        return f"{name} ({cli_id_clean})"
    return name or cli_id_clean


templates.env.filters["cli_label"] = format_cli_label


def filter_hire_by(value, days: int = 30):
    if not value:
        return None
    try:
        return value - timedelta(days=days)
    except Exception:
        return None
templates.env.filters["hire_by"] = filter_hire_by


def _sanitize_export_title(value: object | None) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    return text or "Table Export"


def _sanitize_excel_sheet_title(value: object | None) -> str:
    text = _sanitize_export_title(value)
    text = re.sub(r'[\[\]\*:/\\?]', " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:31] or "Table Export"


def _sanitize_export_filename(value: object | None, suffix: str) -> str:
    title = _sanitize_export_title(value)
    slug = re.sub(r"[^A-Za-z0-9]+", "_", title).strip("_").lower() or "table_export"
    return f"{slug}.{suffix}"


def _normalize_export_text(value: object | None) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _parse_export_report_date(value: object | None) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    for parser in (
        lambda raw: date.fromisoformat(raw),
        lambda raw: datetime.strptime(raw, "%d-%m-%Y").date(),
        lambda raw: datetime.strptime(raw, "%d/%m/%Y").date(),
    ):
        try:
            return parser(text).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return ""


def _coerce_export_table_payload(payload: object) -> tuple[str, list[str], list[list[str]], str]:
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Export payload must be an object.")

    headers_payload = payload.get("headers")
    rows_payload = payload.get("rows")
    if not isinstance(headers_payload, list) or not headers_payload:
        raise HTTPException(status_code=400, detail="Export requires at least one column.")
    if rows_payload is not None and not isinstance(rows_payload, list):
        raise HTTPException(status_code=400, detail="Export rows payload is invalid.")

    title = _sanitize_export_title(payload.get("title"))
    report_date_label = _parse_export_report_date(payload.get("report_date"))
    headers = [
        _normalize_export_text(header) or f"Column {index + 1}"
        for index, header in enumerate(headers_payload)
    ]
    width = len(headers)
    rows: list[list[str]] = []
    for raw_row in rows_payload or []:
        if not isinstance(raw_row, list):
            continue
        normalized = [_normalize_export_text(cell) for cell in raw_row[:width]]
        if len(normalized) < width:
            normalized.extend([""] * (width - len(normalized)))
        rows.append(normalized)
    return title, headers, rows, report_date_label


def _pdf_escape_text(value: object | None) -> str:
    text = _normalize_export_text(value).encode("latin-1", "replace").decode("latin-1")
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _estimate_pdf_text_width(value: object | None, font_size: float, *, bold: bool = False) -> float:
    text = _normalize_export_text(value)
    if not text:
        return 0.0

    total_units = 0.0
    for char in text:
        if char in " .,:;|!'`":
            total_units += 0.28
        elif char in "[](){}frtIjl":
            total_units += 0.38
        elif char in "mwMW@#%&":
            total_units += 0.92
        elif char.isdigit():
            total_units += 0.56
        elif char.isupper():
            total_units += 0.67
        else:
            total_units += 0.57
    return total_units * font_size * (1.05 if bold else 1.0)


def _truncate_pdf_text_to_width(value: object | None, max_width: float, font_size: float) -> str:
    text = _normalize_export_text(value)
    if not text or _estimate_pdf_text_width(text, font_size) <= max_width:
        return text

    suffix = "..."
    available = max_width - _estimate_pdf_text_width(suffix, font_size)
    if available <= font_size * 0.8:
        return suffix

    trimmed = ""
    for char in text:
        if _estimate_pdf_text_width(trimmed + char, font_size) > available:
            break
        trimmed += char
    return (trimmed.rstrip() or text[:1]) + suffix


def _split_pdf_token_to_width(token: str, max_width: float, font_size: float) -> list[str]:
    if not token:
        return [""]

    parts: list[str] = []
    current = ""
    for char in token:
        candidate = current + char
        if current and _estimate_pdf_text_width(candidate, font_size) > max_width:
            parts.append(current)
            current = char
        else:
            current = candidate
    if current:
        parts.append(current)
    return parts or [token]


def _wrap_pdf_text_to_width(
    value: object | None,
    max_width: float,
    font_size: float,
    *,
    max_lines: int,
) -> list[str]:
    text = _normalize_export_text(value)
    if not text:
        return [""]

    words = text.split(" ")
    lines: list[str] = []
    current = ""

    def push(segment: str) -> None:
        nonlocal current, lines
        if not current:
            current = segment
            return
        candidate = f"{current} {segment}"
        if _estimate_pdf_text_width(candidate, font_size) <= max_width:
            current = candidate
            return
        lines.append(current)
        current = segment

    for word in words:
        pieces = (
            _split_pdf_token_to_width(word, max_width, font_size)
            if _estimate_pdf_text_width(word, font_size) > max_width
            else [word]
        )
        for piece in pieces:
            push(piece)

    if current:
        lines.append(current)

    if len(lines) > max_lines:
        overflow = " ".join(lines[max_lines - 1 :])
        lines = lines[: max_lines - 1] + [
            _truncate_pdf_text_to_width(overflow, max_width, font_size)
        ]
    return lines or [""]


def _fit_pdf_column_widths(
    headers: list[str],
    rows: list[list[str]],
    table_width: float,
    body_font_size: float,
    header_font_size: float,
) -> list[float]:
    column_count = len(headers)
    if column_count <= 1:
        return [table_width]

    min_width = max(28.0, min(76.0, table_width / max(column_count * 1.9, 1)))
    if min_width * column_count > table_width:
        return [table_width / column_count] * column_count

    max_width = max(min_width + 12.0, table_width * (0.34 if column_count <= 4 else 0.24 if column_count <= 8 else 0.18))
    sample_rows = rows[: min(len(rows), 250)]
    preferred: list[float] = []

    for index, header in enumerate(headers):
        values = [header] + [row[index] if index < len(row) else "" for row in sample_rows]
        measured = sorted(
            _estimate_pdf_text_width(cell, body_font_size)
            for cell in values[1:]
            if _normalize_export_text(cell)
        )
        if measured:
            pivot = max(0, math.ceil(len(measured) * 0.9) - 1)
            content_width = measured[pivot]
        else:
            content_width = 0.0
        header_width = _estimate_pdf_text_width(header, header_font_size, bold=True)
        preferred.append(min(max(max(header_width, content_width) + 14.0, min_width), max_width))

    total = sum(preferred)
    if total < table_width:
        scale = table_width / total if total else 1.0
        preferred = [width * scale for width in preferred]
    elif total > table_width:
        overflow = total - table_width
        adjusted = preferred[:]
        while overflow > 0.5:
            shrinkable = [max(0.0, width - min_width) for width in adjusted]
            total_shrinkable = sum(shrinkable)
            if total_shrinkable <= 0:
                adjusted = [table_width / column_count] * column_count
                break
            next_widths: list[float] = []
            for width, room in zip(adjusted, shrinkable):
                if room <= 0:
                    next_widths.append(width)
                    continue
                reduction = min(room, overflow * (room / total_shrinkable))
                next_widths.append(width - reduction)
            adjusted = next_widths
            overflow = sum(adjusted) - table_width
        preferred = adjusted

    width_delta = table_width - sum(preferred)
    preferred[-1] += width_delta
    return preferred


def _build_pdf_row_layout(
    values: list[str],
    widths: list[float],
    font_size: float,
    *,
    max_lines: int,
) -> dict[str, object]:
    horizontal_padding = 5.0
    vertical_padding = 4.0
    line_height = font_size + 2.2
    cells: list[list[str]] = []

    for index, width in enumerate(widths):
        text = values[index] if index < len(values) else ""
        inner_width = max(width - (horizontal_padding * 2), font_size * 1.8)
        cells.append(
            _wrap_pdf_text_to_width(
                text,
                inner_width,
                font_size,
                max_lines=max_lines,
            )
        )

    row_line_count = max((len(cell_lines) for cell_lines in cells), default=1)
    row_height = max(18.0, (row_line_count * line_height) + (vertical_padding * 2))
    return {
        "cells": cells,
        "height": row_height,
        "line_height": line_height,
        "padding_x": horizontal_padding,
        "padding_y": vertical_padding,
        "font_size": font_size,
    }


def _build_table_pdf_bytes(
    title: str,
    headers: list[str],
    rows: list[list[str]],
    report_date_label: str,
) -> bytes:
    page_width = 842.0
    page_height = 595.0
    margin_left = 26.0
    margin_right = 26.0
    margin_bottom = 24.0
    table_top = 516.0
    table_width = page_width - margin_left - margin_right
    column_count = max(1, len(headers))
    body_font_size = 8.8 if column_count <= 5 else 8.1 if column_count <= 8 else 7.4 if column_count <= 11 else 6.8
    header_font_size = min(body_font_size + 0.6, 9.2)

    widths = _fit_pdf_column_widths(headers, rows, table_width, body_font_size, header_font_size)
    header_layout = _build_pdf_row_layout(headers, widths, header_font_size, max_lines=3)
    body_rows = rows or [["No rows available."] + [""] * max(0, len(headers) - 1)]
    row_layouts = [_build_pdf_row_layout(row, widths, body_font_size, max_lines=5) for row in body_rows]

    usable_height = table_top - margin_bottom
    pages: list[list[dict[str, object]]] = []
    current_page: list[dict[str, object]] = []
    current_height = float(header_layout["height"])

    for layout in row_layouts:
        layout_height = float(layout["height"])
        if current_page and current_height + layout_height > usable_height:
            pages.append(current_page)
            current_page = []
            current_height = float(header_layout["height"])
        current_page.append(layout)
        current_height += layout_height
    pages.append(current_page)

    def add_text(
        commands: list[str],
        font_name: str,
        font_size: float,
        x: float,
        y: float,
        text: str,
        color: tuple[float, float, float],
    ) -> None:
        commands.append(f"{color[0]:.3f} {color[1]:.3f} {color[2]:.3f} rg")
        commands.append(
            f"BT /{font_name} {font_size:.2f} Tf 1 0 0 1 {x:.2f} {y:.2f} Tm ({_pdf_escape_text(text)}) Tj ET"
        )

    def draw_row(
        commands: list[str],
        row_layout: dict[str, object],
        top_y: float,
        *,
        fill_color: tuple[float, float, float],
        border_color: tuple[float, float, float],
        text_color: tuple[float, float, float],
        font_name: str,
    ) -> float:
        row_height = float(row_layout["height"])
        bottom_y = top_y - row_height
        commands.append(f"{fill_color[0]:.3f} {fill_color[1]:.3f} {fill_color[2]:.3f} rg")
        commands.append(f"{margin_left:.2f} {bottom_y:.2f} {table_width:.2f} {row_height:.2f} re f")
        commands.append("0.55 w")
        commands.append(f"{border_color[0]:.3f} {border_color[1]:.3f} {border_color[2]:.3f} RG")
        commands.append(f"{margin_left:.2f} {bottom_y:.2f} {table_width:.2f} {row_height:.2f} re S")

        x = margin_left
        cells: list[list[str]] = row_layout["cells"]  # type: ignore[assignment]
        line_height = float(row_layout["line_height"])
        padding_x = float(row_layout["padding_x"])
        padding_y = float(row_layout["padding_y"])
        font_size = float(row_layout["font_size"])

        for width, cell_lines in zip(widths, cells):
            commands.append(f"{x + width:.2f} {bottom_y:.2f} m {x + width:.2f} {top_y:.2f} l S")
            text_x = x + padding_x
            text_y = top_y - padding_y - font_size
            for line in cell_lines:
                add_text(commands, font_name, font_size, text_x, text_y, line, text_color)
                text_y -= line_height
            x += width
        return bottom_y

    objects: list[str] = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
    ]
    page_refs: list[int] = []

    total_pages = len(pages)
    for page_index, page_rows in enumerate(pages, start=1):
        commands: list[str] = []
        add_text(commands, "F2", 15.5, margin_left, 560.0, title, (0.086, 0.192, 0.298))
        if report_date_label:
            add_text(
                commands,
                "F1",
                9.4,
                margin_left,
                548.0,
                f"Updated on: {report_date_label}",
                (0.306, 0.427, 0.529),
            )
        add_text(
            commands,
            "F1",
            9.0,
            margin_left,
            536.0 if report_date_label else 544.0,
            f"Rows: {len(rows)}   Page: {page_index}/{total_pages}",
            (0.306, 0.427, 0.529),
        )

        current_y = table_top
        current_y = draw_row(
            commands,
            header_layout,
            current_y,
            fill_color=(0.102, 0.286, 0.467),
            border_color=(0.620, 0.753, 0.886),
            text_color=(0.965, 0.984, 1.000),
            font_name="F2",
        )

        for row_index, row_layout in enumerate(page_rows):
            current_y = draw_row(
                commands,
                row_layout,
                current_y,
                fill_color=(0.968, 0.980, 0.992) if row_index % 2 == 0 else (1.000, 1.000, 1.000),
                border_color=(0.792, 0.867, 0.925),
                text_color=(0.122, 0.180, 0.239),
                font_name="F1",
            )

        stream_body = "\n".join(commands).encode("latin-1", "replace")
        content_object = (
            f"<< /Length {len(stream_body)} >>\nstream\n".encode("latin-1")
            + stream_body
            + b"\nendstream"
        ).decode("latin-1")
        objects.append(content_object)
        content_ref = len(objects)
        page_object = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_width} {page_height}] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_ref} 0 R >>"
        )
        objects.append(page_object)
        page_refs.append(len(objects))

    objects[1] = f"<< /Type /Pages /Count {len(page_refs)} /Kids [{' '.join(f'{ref} 0 R' for ref in page_refs)}] >>"

    pdf_parts = [b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"]
    offsets = [0]
    running_offset = len(pdf_parts[0])
    for index, obj in enumerate(objects, start=1):
        offsets.append(running_offset)
        chunk = f"{index} 0 obj\n{obj}\nendobj\n".encode("latin-1")
        pdf_parts.append(chunk)
        running_offset += len(chunk)

    xref_offset = running_offset
    xref_entries = ["0000000000 65535 f \n"] + [f"{offset:010d} 00000 n \n" for offset in offsets[1:]]
    pdf_parts.append(f"xref\n0 {len(objects) + 1}\n".encode("latin-1"))
    pdf_parts.append("".join(xref_entries).encode("latin-1"))
    pdf_parts.append(
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF".encode("latin-1")
    )
    return b"".join(pdf_parts)

ADMIN_USER = "admin"
ADMIN_PASS = "sdah1234"
_AUTH_COOKIE = "session"
_ALLOWED_PATHS = {"/login", "/logout", "/health"}
_ALLOWED_PREFIXES = ("/static", "/openapi.json", "/docs", "/redoc")
SSTS_WEB_URL = "http://164.52.197.129/devices"
SSTS_API_BASE_URL = "http://164.52.197.129:3000"
SSTS_API_LOGIN_URL = f"{SSTS_API_BASE_URL}/auth/login/"
SSTS_API_DEVICE_URL = f"{SSTS_API_BASE_URL}/device"
SSTS_API_TRAINS_REPORT_URL = f"{SSTS_API_BASE_URL}/train/tr/reportforperiod"
SSTS_API_PUNCT_URL = f"{SSTS_API_BASE_URL}/timetable/tc/punct"
SSTS_API_USER = os.getenv("SSTS_API_USER", "srdeeopsdah@gmail.com")
SSTS_API_PASSWORD = os.getenv("SSTS_API_PASSWORD", "sdah1234")
SSTS_OFFLINE_THRESHOLD_MINUTES = 120
SSTS_RECENTLY_ONLINE_THRESHOLD_MINUTES = 5
SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES = 300
SSTS_RECENT_OFFLINE_MAX_MINUTES = 24 * 60
SSTS_REFRESH_INTERVAL_MINUTES = 5
SSTS_PF_REPORT_CACHE_TTL_MINUTES = 20
SSTS_PF_ANALYSIS_TASK_TTL_MINUTES = 180
SSTS_EXCLUDED_RAKE_NAMES = {"TEST1", "TEST2"}
IST = timezone(timedelta(hours=5, minutes=30))
_SSTS_PF_REPORT_CACHE: dict[str, tuple[datetime, dict[str, object]]] = {}
_SSTS_PF_ANALYSIS_TASKS: dict[str, dict[str, object]] = {}
_SSTS_PF_ANALYSIS_LOCK = threading.Lock()


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        path = request.url.path
        if path in _ALLOWED_PATHS or any(path.startswith(pref) for pref in _ALLOWED_PREFIXES):
            return await call_next(request)
        if request.cookies.get(_AUTH_COOKIE) == "ok":
            return await call_next(request)
        if path.startswith("/exports/") or request.headers.get("x-requested-with", "").lower() == "fetch":
            return PlainTextResponse("Authentication required. Please sign in again and retry the export.", status_code=401)
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


def _format_ist(dt: datetime | None, *, include_seconds: bool = False) -> str:
    dt_utc = _ensure_utc(dt)
    if dt_utc is None:
        return "No signal"
    fmt = "%d-%m-%Y %H:%M:%S IST" if include_seconds else "%d-%m-%Y %H:%M IST"
    return dt_utc.astimezone(IST).strftime(fmt)


def _format_ist_time(dt: datetime | None) -> str:
    dt_utc = _ensure_utc(dt)
    if dt_utc is None:
        return ""
    return dt_utc.astimezone(IST).strftime("%H:%M")


def _format_duration(minutes: int | None) -> str | None:
    if minutes is None:
        return None
    total_minutes = max(0, int(minutes))
    hours = total_minutes // 60
    mins = total_minutes % 60
    if hours > 0:
        return f"{hours}h {mins}m"
    return f"{mins}m"


def _parse_ssts_timestamp(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        # The SSTS API sends timestamps with a trailing "Z", but the original
        # dashboard treats them as local server wall time. Preserve that wall
        # time in IST so our report matches the source system exactly.
        normalized = value.strip().replace("T", " ")
        if normalized.endswith("Z"):
            normalized = normalized[:-1]
        parsed = datetime.fromisoformat(normalized)
        if parsed.tzinfo is not None:
            parsed = parsed.replace(tzinfo=None)
        return parsed.replace(tzinfo=IST).astimezone(timezone.utc)
    except ValueError:
        return None


def _parse_report_date(value: str | None) -> date | None:
    if not value:
        return None
    normalized = value.strip()
    if not normalized:
        return None
    for separator in ("/", "-"):
        parts = normalized.split(separator)
        if len(parts) == 3 and len(parts[0]) == 2 and len(parts[1]) == 2 and len(parts[2]) == 4:
            try:
                day_value = int(parts[0])
                month_value = int(parts[1])
                year_value = int(parts[2])
                return date(year_value, month_value, day_value)
            except ValueError:
                return None
    try:
        return date.fromisoformat(normalized)
    except ValueError:
        return None


def _minutes_since(now_utc: datetime, lastupdate: datetime | None) -> int | None:
    if not lastupdate:
        return None
    diff = now_utc - _ensure_utc(lastupdate)
    return max(0, int(diff.total_seconds() // 60))


def _snapshot_offline_minutes(
    snapshot: SstsDeviceSnapshot,
    reference_time: datetime | None = None,
) -> int | None:
    reference_utc = _ensure_utc(reference_time or snapshot.observed_at)
    if reference_utc is None:
        return snapshot.offline_minutes
    return _minutes_since(reference_utc, snapshot.lastupdate)


def _ssts_is_offline(snapshot: SstsDeviceSnapshot, reference_time: datetime | None = None) -> bool:
    return (_snapshot_offline_minutes(snapshot, reference_time) or 0) > SSTS_OFFLINE_THRESHOLD_MINUTES


def _ssts_is_recently_offline(snapshot: SstsDeviceSnapshot, reference_time: datetime | None = None) -> bool:
    minutes = _snapshot_offline_minutes(snapshot, reference_time) or 0
    return SSTS_OFFLINE_THRESHOLD_MINUTES < minutes < SSTS_RECENT_OFFLINE_MAX_MINUTES


def _ssts_is_online_now(snapshot: SstsDeviceSnapshot, reference_time: datetime | None = None) -> bool:
    minutes = _snapshot_offline_minutes(snapshot, reference_time)
    if minutes is None:
        return False
    return minutes <= SSTS_RECENTLY_ONLINE_THRESHOLD_MINUTES


def _ssts_normalize_rake_name(value: object | None) -> str:
    return str(value or "").strip().upper()


def _ssts_is_excluded_rake_name(value: object | None) -> bool:
    return _ssts_normalize_rake_name(value) in SSTS_EXCLUDED_RAKE_NAMES


def _snapshot_to_row(snapshot: SstsDeviceSnapshot, reference_time: datetime | None = None) -> dict[str, object]:
    offline_minutes = _snapshot_offline_minutes(snapshot, reference_time)
    return {
        "device_id": snapshot.device_id,
        "name": snapshot.name,
        "uniqueid": snapshot.uniqueid or "",
        "phone": snapshot.phone or "",
        "contact": snapshot.contact or "",
        "lastupdate": snapshot.lastupdate,
        "lastupdate_label": _format_ist(snapshot.lastupdate, include_seconds=True),
        "offline_minutes": offline_minutes,
        "offline_hours": round((offline_minutes or 0) / 60, 1) if offline_minutes is not None else None,
        "offline_duration": _format_duration(offline_minutes),
    }


def _ssts_sort_key(snapshot: SstsDeviceSnapshot, reference_time: datetime | None = None) -> tuple[int, int, str]:
    offline_minutes = _snapshot_offline_minutes(snapshot, reference_time)
    return (
        -(offline_minutes or -1),
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


def _ssts_get_json_with_params(
    url: str,
    params: dict[str, object],
    headers: dict[str, str] | None = None,
) -> object:
    encoded = urlparse.urlencode(
        {key: value for key, value in params.items() if value not in (None, "")},
        doseq=True,
    )
    full_url = f"{url}?{encoded}" if encoded else url
    return _ssts_get_json(full_url, headers=headers)


def fetch_ssts_token() -> str:
    login_payload = {"username": SSTS_API_USER, "password": SSTS_API_PASSWORD}
    login_data = _ssts_post_json(SSTS_API_LOGIN_URL, login_payload)
    token = str(login_data.get("token") or "").strip()
    if not token:
        raise RuntimeError("SSTS login succeeded but token was missing.")
    return token


def fetch_ssts_devices() -> list[dict[str, object]]:
    token = fetch_ssts_token()
    devices = _ssts_get_json(SSTS_API_DEVICE_URL, headers={"Authorization": token})
    if not isinstance(devices, list):
        raise RuntimeError("Unexpected SSTS device response format.")
    return [
        item
        for item in devices
        if isinstance(item, dict) and not _ssts_is_excluded_rake_name(item.get("name"))
    ]


def fetch_ssts_trains_report(report_day: date, token: str) -> list[dict[str, object]]:
    response = _ssts_get_json_with_params(
        SSTS_API_TRAINS_REPORT_URL,
        {"train_date": report_day.isoformat()},
        headers={"Authorization": token},
    )
    if not isinstance(response, list):
        raise RuntimeError("Unexpected SSTS trains report format.")
    return [
        item
        for item in response
        if isinstance(item, dict) and not _ssts_is_excluded_rake_name(item.get("device_name"))
    ]


def _format_time_value(value: object) -> str:
    if value in (None, ""):
        return ""
    try:
        text = str(value)
        if "T" in text:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.astimezone(IST)
            return parsed.strftime("%H:%M:%S")
    except ValueError:
        pass
    return str(value)


def _build_pf_report_rows_for_train(
    train: dict[str, object],
    report_day: date,
    token: str,
) -> list[dict[str, object]]:
    base_row = {
        "report_date": report_day.strftime("%d-%m-%Y"),
        "train_no": str(train.get("train_no") or ""),
        "rake_no": str(train.get("device_name") or ""),
        "device_id": train.get("device_id"),
        "org": str(train.get("org") or ""),
        "dest": str(train.get("dest") or ""),
        "crew_name": str(train.get("crew_name") or ""),
    }
    params = {
        "train_date": report_day.isoformat(),
        "train_no": train.get("train_no"),
        "device_id": train.get("device_id"),
        "org": train.get("org"),
        "dep": train.get("dep"),
        "dest": train.get("dest"),
        "arr": train.get("arr"),
        "recalc": "true",
    }
    try:
        response = _ssts_get_json_with_params(
            SSTS_API_PUNCT_URL,
            params,
            headers={"Authorization": token},
        )
    except (urlerror.URLError, RuntimeError, ValueError, json.JSONDecodeError):
        response = []
    if not isinstance(response, list) or not response:
        return [
            {
                **base_row,
                "station": "",
                "srl_no": "",
                "sch_arr": "",
                "sch_dep": _format_time_value(train.get("dep")),
                "act_arr": _format_time_value(train.get("act_arr")),
                "act_dep": _format_time_value(train.get("act_dep")),
                "stop_time": "",
                "geofence_enter_speed": "",
                "pf_enter_speed": "",
                "pf_distance": "",
                "remarks": "",
                "status_message": "Data not found or Device might be Offline",
            }
        ]
    detail_rows: list[dict[str, object]] = []
    for item in response:
        if not isinstance(item, dict):
            continue
        detail_rows.append(
            {
                **base_row,
                "station": str(item.get("stn_code") or ""),
                "srl_no": item.get("srl_no") or "",
                "sch_arr": _format_time_value(item.get("sch_arr")),
                "sch_dep": _format_time_value(item.get("sch_dep")),
                "act_arr": _format_time_value(item.get("act_arr")),
                "act_dep": _format_time_value(item.get("act_dep")),
                "crew_name": str(item.get("crew_name") or base_row.get("crew_name") or ""),
                "stop_time": _format_time_value(item.get("stop_time")),
                "geofence_enter_speed": item.get("geofence_enter_speed")
                if item.get("geofence_enter_speed") is not None
                else "",
                "pf_enter_speed": item.get("pf_enter_speed") if item.get("pf_enter_speed") is not None else "",
                "pf_distance": item.get("pf_distance") if item.get("pf_distance") is not None else "",
                "remarks": str(item.get("remarks") or ""),
                "status_message": "",
            }
        )
    return detail_rows or [
        {
            **base_row,
            "station": "",
            "srl_no": "",
            "sch_arr": "",
            "sch_dep": _format_time_value(train.get("dep")),
            "act_arr": _format_time_value(train.get("act_arr")),
            "act_dep": _format_time_value(train.get("act_dep")),
            "stop_time": "",
            "geofence_enter_speed": "",
            "pf_enter_speed": "",
            "pf_distance": "",
            "remarks": "",
            "status_message": "Data not found or Device might be Offline",
        }
    ]


def build_ssts_pf_entering_context(report_day: date) -> dict[str, object]:
    cache_key = report_day.isoformat()
    cached_entry = _SSTS_PF_REPORT_CACHE.get(cache_key)
    now_utc = _utc_now()
    if cached_entry:
        cached_at, cached_payload = cached_entry
        if (now_utc - cached_at) < timedelta(minutes=SSTS_PF_REPORT_CACHE_TTL_MINUTES):
            return dict(cached_payload)

    token = fetch_ssts_token()
    trains = fetch_ssts_trains_report(report_day, token)
    rows: list[dict[str, object]] = []
    missing_count = 0
    if trains:
        with ThreadPoolExecutor(max_workers=6) as executor:
            future_map = {
                executor.submit(_build_pf_report_rows_for_train, train, report_day, token): train
                for train in trains
            }
            for future in as_completed(future_map):
                train_rows = future.result()
                rows.extend(train_rows)
                if any(row.get("status_message") for row in train_rows):
                    missing_count += 1
    rows.sort(
        key=lambda row: (
            str(row.get("train_no") or ""),
            999999 if row.get("srl_no") in ("", None) else int(row.get("srl_no") or 0),
            str(row.get("station") or ""),
        )
    )
    payload = {
        "pf_report_day": report_day.isoformat(),
        "pf_report_day_label": report_day.strftime("%d-%m-%Y"),
        "pf_report_rows": rows,
        "pf_report_total_trains": len(trains),
        "pf_report_total_rows": len(rows),
        "pf_report_missing_count": missing_count,
    }
    _SSTS_PF_REPORT_CACHE[cache_key] = (now_utc, payload)
    stale_keys = [
        key
        for key, (cached_at, _) in _SSTS_PF_REPORT_CACHE.items()
        if (now_utc - cached_at) >= timedelta(minutes=SSTS_PF_REPORT_CACHE_TTL_MINUTES)
    ]
    for stale_key in stale_keys:
        _SSTS_PF_REPORT_CACHE.pop(stale_key, None)
    return dict(payload)


def _pf_speed_value(value: object) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _cleanup_ssts_pf_analysis_tasks() -> None:
    now_utc = _utc_now()
    stale_ids: list[str] = []
    with _SSTS_PF_ANALYSIS_LOCK:
        for task_id, payload in _SSTS_PF_ANALYSIS_TASKS.items():
            updated_at = payload.get("updated_at")
            if not isinstance(updated_at, datetime):
                stale_ids.append(task_id)
                continue
            if (now_utc - updated_at) >= timedelta(minutes=SSTS_PF_ANALYSIS_TASK_TTL_MINUTES):
                stale_ids.append(task_id)
        for task_id in stale_ids:
            _SSTS_PF_ANALYSIS_TASKS.pop(task_id, None)


def _set_ssts_pf_analysis_task(task_id: str, **values: object) -> None:
    with _SSTS_PF_ANALYSIS_LOCK:
        payload = _SSTS_PF_ANALYSIS_TASKS.setdefault(task_id, {})
        payload.update(values)
        payload["updated_at"] = _utc_now()


def _get_ssts_pf_analysis_task(task_id: str | None) -> dict[str, object] | None:
    if not task_id:
        return None
    _cleanup_ssts_pf_analysis_tasks()
    with _SSTS_PF_ANALYSIS_LOCK:
        payload = _SSTS_PF_ANALYSIS_TASKS.get(task_id)
        if not payload:
            return None
        return dict(payload)


def _build_ssts_pf_speed_analysis_result(report_day: date) -> dict[str, object]:
    raw_context = build_ssts_pf_entering_context(report_day)
    all_rows_by_train: dict[str, list[dict[str, object]]] = {}
    for row in raw_context.get("pf_report_rows", []):
        if not isinstance(row, dict):
            continue
        train_no = str(row.get("train_no") or "").strip()
        if train_no:
            all_rows_by_train.setdefault(train_no, []).append(dict(row))
    for rows in all_rows_by_train.values():
        rows.sort(
            key=lambda row: (
                999999 if row.get("srl_no") in ("", None) else int(row.get("srl_no") or 0),
                str(row.get("station") or ""),
            )
        )

    filtered_rows: list[dict[str, object]] = []
    for row in raw_context.get("pf_report_rows", []):
        if not isinstance(row, dict):
            continue
        geofence_speed = _pf_speed_value(row.get("geofence_enter_speed"))
        pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
        if (geofence_speed is not None and geofence_speed > 40) or (pf_speed is not None and pf_speed > 40):
            filtered_rows.append(dict(row))

    filtered_rows.sort(
        key=lambda row: (
            str(row.get("train_no") or ""),
            999999 if row.get("srl_no") in ("", None) else int(row.get("srl_no") or 0),
            str(row.get("station") or ""),
        )
    )

    summary_by_train: dict[str, dict[str, object]] = {}
    detail_rows_by_train: dict[str, list[dict[str, object]]] = {}
    for row in filtered_rows:
        train_no = str(row.get("train_no") or "").strip()
        if not train_no:
            continue
        detail_rows_by_train[train_no] = all_rows_by_train.get(train_no, [row])
        summary = summary_by_train.setdefault(
            train_no,
            {
                "report_date": str(row.get("report_date") or report_day.strftime("%d-%m-%Y")),
                "train_no": train_no,
                "rake_no": str(row.get("rake_no") or ""),
                "device_id": row.get("device_id") or "",
                "org": str(row.get("org") or ""),
                "dest": str(row.get("dest") or ""),
                "crew_name": str(row.get("crew_name") or ""),
                "occurrence_count": 0,
                "max_geofence_enter_speed": "",
                "max_pf_enter_speed": "",
            },
        )
        if not summary.get("crew_name") and row.get("crew_name"):
            summary["crew_name"] = str(row.get("crew_name") or "")
        geofence_speed = _pf_speed_value(row.get("geofence_enter_speed"))
        pf_speed = _pf_speed_value(row.get("pf_enter_speed"))
        if pf_speed is not None and pf_speed > 40:
            summary["occurrence_count"] = int(summary.get("occurrence_count") or 0) + 1
        if geofence_speed is not None:
            current_max = _pf_speed_value(summary.get("max_geofence_enter_speed"))
            if current_max is None or geofence_speed > current_max:
                summary["max_geofence_enter_speed"] = int(geofence_speed) if geofence_speed.is_integer() else geofence_speed
        if pf_speed is not None:
            current_max = _pf_speed_value(summary.get("max_pf_enter_speed"))
            if current_max is None or pf_speed > current_max:
                summary["max_pf_enter_speed"] = int(pf_speed) if pf_speed.is_integer() else pf_speed

    summary_rows = sorted(
        summary_by_train.values(),
        key=lambda row: (-int(row.get("occurrence_count") or 0), str(row.get("train_no") or "")),
    )

    return {
        "pf_report_day": report_day.isoformat(),
        "pf_report_day_label": report_day.strftime("%d-%m-%Y"),
        "pf_analysis_summary_rows": summary_rows,
        "pf_analysis_detail_rows_by_train": detail_rows_by_train,
        "pf_analysis_total_trains": len(summary_rows),
        "pf_analysis_total_rows": len(filtered_rows),
        "pf_analysis_source_total_trains": int(raw_context.get("pf_report_total_trains") or 0),
        "pf_analysis_source_total_rows": int(raw_context.get("pf_report_total_rows") or 0),
        "pf_analysis_missing_count": int(raw_context.get("pf_report_missing_count") or 0),
    }


def _run_ssts_pf_analysis_task(task_id: str, report_day: date) -> None:
    try:
        _set_ssts_pf_analysis_task(
            task_id,
            status="running",
            progress=8,
            message="Preparing PF analysis...",
        )
        _set_ssts_pf_analysis_task(
            task_id,
            progress=24,
            message="Fetching train-wise PF data...",
        )
        result = _build_ssts_pf_speed_analysis_result(report_day)
        _set_ssts_pf_analysis_task(
            task_id,
            progress=88,
            message="Building 40+ speed summary...",
        )
        _set_ssts_pf_analysis_task(
            task_id,
            status="completed",
            progress=100,
            message="Analysis complete.",
            result=result,
        )
    except Exception as exc:
        _set_ssts_pf_analysis_task(
            task_id,
            status="error",
            progress=100,
            message=f"Analysis failed: {exc}",
            error=str(exc),
        )


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
                "message": f"Using last sync from {_format_ist(latest_run_time)}.",
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
    return [snapshot for snapshot in snapshots if not _ssts_is_excluded_rake_name(snapshot.name)]


def build_ssts_report_context(
    session: Session,
    selected_day: date | None = None,
    analysis_day: date | None = None,
) -> dict[str, object]:
    runs = _distinct_ssts_runs(session)
    latest_run = next((run for run in runs if run.fetch_status == "ok"), None)
    if not latest_run:
        return {
            "latest_run": None,
            "latest_rows": [],
            "online_now": [],
            "current_not_online": [],
            "current_offline": [],
            "current_recently_offline": [],
            "recently_online": [],
            "daily_summary": [],
            "analysis_day_options": [],
            "selected_analysis_day": None,
            "selected_analysis_day_label": None,
            "selected_analysis_rows": [],
            "selected_day": None,
            "selected_day_label": None,
            "selected_day_run": None,
            "selected_day_recent_offline_rows": [],
            "selected_day_recently_online_rows": [],
        }

    latest_snapshots = _snapshots_for_run(session, latest_run.id or 0)
    latest_map = {row.device_id: row for row in latest_snapshots}
    latest_run_time = _ensure_utc(latest_run.observed_at)
    current_reference_time = _utc_now()

    online_now = [
        _snapshot_to_row(row, reference_time=current_reference_time)
        for row in sorted(latest_snapshots, key=lambda item: (item.name.lower(), item.device_id))
        if _ssts_is_online_now(row, reference_time=current_reference_time)
    ]
    current_not_online = [
        _snapshot_to_row(row, reference_time=current_reference_time)
        for row in sorted(latest_snapshots, key=lambda item: _ssts_sort_key(item, current_reference_time))
        if not _ssts_is_online_now(row, reference_time=current_reference_time)
    ]
    current_offline = [
        _snapshot_to_row(row, reference_time=current_reference_time)
        for row in sorted(latest_snapshots, key=lambda item: _ssts_sort_key(item, current_reference_time))
        if _ssts_is_offline(row, reference_time=current_reference_time)
    ]
    current_recently_offline = [
        _snapshot_to_row(row, reference_time=current_reference_time)
        for row in sorted(latest_snapshots, key=lambda item: _ssts_sort_key(item, current_reference_time))
        if _ssts_is_recently_offline(row, reference_time=current_reference_time)
    ]
    recovery_runs = [run for run in runs if run.fetch_status == "ok" and _ensure_utc(run.observed_at) <= latest_run_time]
    snapshots_by_run = {run.id: _snapshots_for_run(session, run.id or 0) for run in recovery_runs}
    history_by_device: dict[int, list[SstsDeviceSnapshot]] = {}
    recently_online_by_run_id: dict[int, int] = {}
    recently_online_info_by_run_device: dict[tuple[int, int], dict[str, object]] = {}
    for run in sorted(recovery_runs, key=lambda item: item.observed_at):
        current_rows = snapshots_by_run.get(run.id, [])
        recovered_count = 0
        for row in current_rows:
            history = history_by_device.setdefault(row.device_id, [])
            if _ssts_is_online_now(row) and history:
                previous_row = history[-1]
                if (previous_row.offline_minutes or 0) > SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES:
                    recovered_count += 1
                    recently_online_info_by_run_device[(run.id or 0, row.device_id)] = {
                        "previous_offline_hours": round((previous_row.offline_minutes or 0) / 60, 1),
                        "previous_offline_duration": _format_duration(previous_row.offline_minutes),
                        "previous_seen": _format_ist(previous_row.observed_at),
                    }
            history.append(row)
        recently_online_by_run_id[run.id or 0] = recovered_count

    latest_by_day: dict[date, SstsSnapshotRun] = {}
    for run in runs:
        if run.fetch_status != "ok":
            continue
        observed_day_ist = _ensure_utc(run.observed_at).astimezone(IST).date()
        latest_by_day.setdefault(observed_day_ist, run)
    daily_summary = []
    for day, run in sorted(latest_by_day.items(), key=lambda item: item[0], reverse=True)[:7]:
        rows = _snapshots_for_run(session, run.id or 0)
        offline_count = sum(1 for row in rows if _ssts_is_offline(row))
        recent_offline_count = sum(1 for row in rows if _ssts_is_recently_offline(row))
        daily_summary.append(
            {
                "day": day.strftime("%d-%m-%Y"),
                "day_iso": day.isoformat(),
                "observed_at": _format_ist(run.observed_at),
                "total_rakes": len(rows),
                "offline_count": offline_count,
                "recent_offline_count": recent_offline_count,
                "recently_online_count": recently_online_by_run_id.get(run.id or 0, 0),
            }
        )

    analysis_day_options = [
        {"day": row["day"], "day_iso": row["day_iso"]}
        for row in daily_summary
    ]
    analysis_day_value = analysis_day if analysis_day in latest_by_day else None
    if analysis_day_value is None:
        analysis_day_value = selected_day
    if analysis_day_value is None and analysis_day_options:
        analysis_day_value = date.fromisoformat(str(analysis_day_options[0]["day_iso"]))

    selected_analysis_rows: list[dict[str, object]] = []
    selected_analysis_summary = {
        "day_label": analysis_day_value.strftime("%d-%m-%Y") if analysis_day_value else "",
        "continuous_offline_count": 0,
        "mixed_online_offline_count": 0,
        "total_rakes": 0,
    }
    if analysis_day_value is not None:
        selected_day_runs = [
            run
            for run in runs
            if run.fetch_status == "ok" and _ensure_utc(run.observed_at).astimezone(IST).date() == analysis_day_value
        ]
        selected_day_runs.sort(key=lambda item: item.observed_at)
        day_rows_by_run = {
            run.id or 0: _snapshots_for_run(session, run.id or 0)
            for run in selected_day_runs
        }
        rake_points: dict[int, dict[str, object]] = {}
        for run in selected_day_runs:
            run_time = _ensure_utc(run.observed_at)
            for row in day_rows_by_run.get(run.id or 0, []):
                state = "offline" if _ssts_is_offline(row) else "online"
                rake = rake_points.setdefault(
                    row.device_id,
                    {
                        "device_id": row.device_id,
                        "name": row.name,
                        "uniqueid": row.uniqueid or "",
                        "remark": row.remark or "",
                        "lastupdate": row.lastupdate,
                        "lastupdate_label": _format_ist(row.lastupdate, include_seconds=True),
                        "points": [],
                    },
                )
                rake["name"] = row.name
                rake["uniqueid"] = row.uniqueid or ""
                rake["lastupdate"] = row.lastupdate
                rake["lastupdate_label"] = _format_ist(row.lastupdate, include_seconds=True)
                if row.remark:
                    rake["remark"] = row.remark
                rake["points"].append({"time": run_time, "state": state})

        day_start = datetime.combine(analysis_day_value, datetime.min.time(), tzinfo=IST).astimezone(timezone.utc)
        day_end = day_start + timedelta(days=1)
        for rake in rake_points.values():
            points = sorted(
                [point for point in rake.get("points", []) if point.get("time") is not None],
                key=lambda point: point["time"],
            )
            if not points:
                continue

            device_history = history_by_device.get(int(rake["device_id"]), [])
            prior_snapshot = None
            for snapshot in reversed(device_history):
                snapshot_time = _ensure_utc(snapshot.observed_at)
                if snapshot_time is not None and snapshot_time < day_start:
                    prior_snapshot = snapshot
                    break

            seed_state = None
            if prior_snapshot is not None:
                seed_state = "offline" if _ssts_is_offline(prior_snapshot, reference_time=day_start) else "online"
            elif points[0]["time"] > day_start:
                seed_state = str(points[0]["state"])

            if seed_state is not None and points[0]["time"] > day_start:
                points.insert(0, {"time": day_start, "state": seed_state})

            segments: list[dict[str, object]] = []

            def build_segment(
                state: str,
                start_time: datetime,
                end_time: datetime,
                *,
                count_for_periods: bool = False,
            ) -> dict[str, object]:
                duration_minutes = max(0, int((end_time - start_time).total_seconds() // 60))
                display_end = end_time - timedelta(minutes=1)
                return {
                    "state": state,
                    "start_label": _format_ist_time(start_time),
                    "end_label": _format_ist_time(display_end),
                    "duration_label": _format_duration(duration_minutes) or "",
                    "width_percent": round((duration_minutes / (24 * 60)) * 100, 2),
                    "summary_label": f"{_format_ist_time(start_time)} to {_format_ist_time(display_end)} {'offline' if state == 'offline' else 'online'}",
                    "count_for_periods": count_for_periods,
                }

            current_state = str(points[0]["state"])
            segment_start = max(day_start, points[0]["time"])
            for point in points[1:]:
                point_time = point["time"]
                point_state = str(point["state"])
                if point_state == current_state:
                    continue
                duration_minutes = max(0, int((point_time - segment_start).total_seconds() // 60))
                if duration_minutes > 0:
                    segments.append(
                        build_segment(
                            current_state,
                            segment_start,
                            point_time,
                            count_for_periods=duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                        )
                    )
                current_state = point_state
                segment_start = point_time

            lastupdate_time = _ensure_utc(rake.get("lastupdate"))
            if analysis_day_value == current_reference_time.astimezone(IST).date():
                reference_end = min(day_end, current_reference_time)
            else:
                reference_end = min(day_end, points[-1]["time"] + timedelta(minutes=SSTS_REFRESH_INTERVAL_MINUTES))

            current_segment_start = segment_start
            current_segment_end = reference_end
            followup_offline_start: datetime | None = None

            if lastupdate_time is not None:
                bounded_lastupdate = min(reference_end, max(day_start, lastupdate_time))
                if current_state == "offline":
                    current_segment_start = min(current_segment_start, bounded_lastupdate)
                elif current_state == "online" and current_segment_start < bounded_lastupdate < reference_end:
                    current_segment_end = bounded_lastupdate
                    followup_offline_start = bounded_lastupdate

            duration_minutes = max(0, int((current_segment_end - current_segment_start).total_seconds() // 60))
            if duration_minutes > 0:
                display_end_time = current_segment_end
                if current_state == "online" and lastupdate_time is not None:
                    display_end_time = min(current_segment_end, max(current_segment_start, lastupdate_time))
                segments.append(
                    build_segment(
                        current_state,
                        current_segment_start,
                        display_end_time,
                        count_for_periods=duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                    )
                )

            if followup_offline_start is not None and followup_offline_start < reference_end:
                offline_duration_minutes = max(0, int((reference_end - followup_offline_start).total_seconds() // 60))
                if offline_duration_minutes > 0:
                    segments.append(
                        build_segment(
                            "offline",
                            followup_offline_start,
                            reference_end,
                            count_for_periods=offline_duration_minutes >= SSTS_OFFLINE_THRESHOLD_MINUTES,
                        )
                    )

            offline_periods = sum(
                1 for segment in segments if segment["state"] == "offline" and segment.get("count_for_periods")
            )
            online_periods = sum(
                1 for segment in segments if segment["state"] == "online" and segment.get("count_for_periods")
            )
            selected_analysis_rows.append(
                {
                    "device_id": rake["device_id"],
                    "name": rake["name"],
                    "uniqueid": rake["uniqueid"],
                    "lastupdate_label": rake["lastupdate_label"],
                    "remark": rake["remark"],
                    "segments": segments,
                    "offline_periods": offline_periods,
                    "online_periods": online_periods,
                    "timeline_summary": ", ".join(str(segment["summary_label"]) for segment in segments),
                }
            )

        selected_analysis_summary["total_rakes"] = len(selected_analysis_rows)
        selected_analysis_summary["continuous_offline_count"] = sum(
            1
            for row in selected_analysis_rows
            if row.get("segments")
            and {str(segment.get("state")) for segment in row.get("segments", [])} == {"offline"}
        )
        selected_analysis_summary["mixed_online_offline_count"] = sum(
            1
            for row in selected_analysis_rows
            if {"online", "offline"}.issubset({str(segment.get("state")) for segment in row.get("segments", [])})
        )

        selected_analysis_rows.sort(
            key=lambda item: (
                -int(item.get("offline_periods") or 0),
                -int(item.get("online_periods") or 0),
                str(item.get("name") or "").lower(),
            )
        )

    recently_online = []
    for device_id, latest_row in latest_map.items():
        if not _ssts_is_online_now(latest_row, reference_time=current_reference_time):
            continue
        history = history_by_device.get(device_id, [])
        if len(history) < 2:
            continue
        recovery_row = history[-1]
        previous_row = history[-2]
        if recovery_row.run_id != (latest_run.id or 0):
            continue
        if (previous_row.offline_minutes or 0) > SSTS_PREVIOUSLY_OFFLINE_THRESHOLD_MINUTES:
            row = _snapshot_to_row(latest_row, reference_time=current_reference_time)
            row["recovery_seen"] = _format_ist(recovery_row.observed_at)
            row["recovery_lastupdate_label"] = _format_ist(recovery_row.lastupdate)
            row["recovery_offline_duration"] = _format_duration(recovery_row.offline_minutes)
            row["previous_offline_minutes"] = previous_row.offline_minutes or 0
            row["previous_offline_duration"] = _format_duration(previous_row.offline_minutes)
            row["previous_seen"] = _format_ist(previous_row.observed_at)
            recently_online.append(row)
    recently_online.sort(
        key=lambda row: (-int(row.get("previous_offline_minutes") or 0), str(row.get("name") or "").lower())
    )

    available_days = [item["day_iso"] for item in daily_summary]
    selected_day_value = selected_day
    if selected_day_value is None and available_days:
        selected_day_value = date.fromisoformat(str(available_days[0]))
    selected_day_run = latest_by_day.get(selected_day_value) if selected_day_value else None
    selected_day_recent_offline_rows = []
    selected_day_recently_online_rows = []
    if selected_day_run:
        selected_run_id = selected_day_run.id or 0
        for row in sorted(snapshots_by_run.get(selected_run_id, []), key=_ssts_sort_key):
            detail = _snapshot_to_row(row)
            if _ssts_is_recently_offline(row):
                detail["status"] = "Offline 2 Hours to 1 Day"
                selected_day_recent_offline_rows.append(detail)
                continue
            if _ssts_is_online_now(row) and (selected_run_id, row.device_id) in recently_online_info_by_run_device:
                detail["status"] = "Recently Back Online"
                detail.update(recently_online_info_by_run_device[(selected_run_id, row.device_id)])
                selected_day_recently_online_rows.append(detail)

    return {
        "latest_run": latest_run,
        "latest_rows": [
            _snapshot_to_row(row, reference_time=current_reference_time)
            for row in sorted(latest_snapshots, key=lambda item: _ssts_sort_key(item, current_reference_time))
        ],
        "online_now": online_now,
        "current_not_online": current_not_online,
        "current_offline": current_offline,
        "current_recently_offline": current_recently_offline,
        "recently_online": recently_online,
        "daily_summary": daily_summary,
        "analysis_day_options": analysis_day_options,
        "selected_analysis_day": analysis_day_value.isoformat() if analysis_day_value else None,
        "selected_analysis_day_label": analysis_day_value.strftime("%d-%m-%Y") if analysis_day_value else None,
        "selected_analysis_rows": selected_analysis_rows,
        "selected_analysis_summary": selected_analysis_summary,
        "selected_day": selected_day_value.isoformat() if selected_day_value else None,
        "selected_day_label": selected_day_value.strftime("%d-%m-%Y") if selected_day_value else None,
        "selected_day_run": selected_day_run,
        "selected_day_recent_offline_rows": selected_day_recent_offline_rows,
        "selected_day_recently_online_rows": selected_day_recently_online_rows,
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
    category: Optional[str] = None,
    gradation: Optional[str] = None,
    cli_status: Optional[str] = None,
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
    category_opts = sorted({e.category for e in employees_all if e.category})
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
    if category:
        category_lower = category.lower()
        employees = [e for e in employees if e.category and category_lower in e.category.lower()]
    if gradation:
        grad_lower = gradation.lower()
        employees = [e for e in employees if e.gradation and grad_lower in e.gradation.lower()]
    if cli_status == "assigned":
        employees = [e for e in employees if e.cli or e.cli_id]
    elif cli_status == "unassigned":
        employees = [e for e in employees if not e.cli and not e.cli_id]

    def sort_key(e: Employee):
        if sort == "name":
            return (e.name.lower(),)
        if sort == "retirement":
            return (e.retirement_date or date.max, e.name)
        if sort == "hire":
            return (e.hire_date, e.name)
        if sort == "category":
            return ((e.category or "").lower(), e.name)
        if sort == "gradation":
            return ((e.gradation or "").lower(), e.name)
        if sort == "cli":
            return ((e.cli or "").strip().lower(), e.name)
        if sort == "working_at":
            return ((e.working_at or "").lower(), e.name)
        return (role_sort_key(e.role), e.name)

    employees = sorted(employees, key=sort_key)
    total_count = len(employees)
    page = 1
    total_pages = 1
    page_start = 0
    employee_return_to = str(request.url)
    employee_return_to_query = urlparse.quote(employee_return_to, safe="")

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
            "filter_category": category or "",
            "filter_gradation": gradation or "",
            "filter_cli_status": cli_status or "",
            "sort": sort,
            "working_opts": working_opts,
            "cli_opts": cli_opts,
            "category_opts": category_opts,
            "gradation_opts": gradation_opts,
            "cli_roster": cli_roster,
            "roster_name": roster_name or "",
            "roster_cli": roster_cli or "",
            "roster_gradation": roster_gradation or "",
            "employees_open": employees_open,
            "roster_open": roster_open,
            "total_count": total_count,
            "page": page,
            "total_pages": total_pages,
            "page_start": page_start,
            "employee_return_to": employee_return_to,
            "employee_return_to_query": employee_return_to_query,
            "sync_error": "",
            "sync_notice": "",
            "sync_warning": "",
            "sync_backup": "",
            "sync_backup_label": "",
            "google_sync_ready": False,
            "google_sync_range": "",
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
    unassigned_name: Optional[str] = None,
    unassigned_role: Optional[str] = None,
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
    unassigned_cli_staff = [e for e in employees if not e.cli and not e.cli_id]
    if unassigned_name:
        name_lower = unassigned_name.lower()
        unassigned_cli_staff = [e for e in unassigned_cli_staff if name_lower in e.name.lower()]
    if unassigned_role:
        unassigned_cli_staff = [e for e in unassigned_cli_staff if e.role == unassigned_role]
    unassigned_cli_staff = sorted(unassigned_cli_staff, key=lambda e: (role_sort_key(e.role), e.name))
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
            "unassigned_name": unassigned_name or "",
            "unassigned_role": unassigned_role or "",
            "unassigned_cli_staff": unassigned_cli_staff,
        },
    )
    response.set_cookie("as_of", end.isoformat())
    response.set_cookie("reports_start_date", start.isoformat())
    response.set_cookie("reports_end_date", end.isoformat())
    return response


def _cli_page_context(
    request: Request,
    session: Session,
    *,
    roster_name: str | None = None,
    roster_cli: str | None = None,
    roster_role: str | None = None,
    roster_gradation: str | None = None,
    roster_cli_status: str | None = None,
) -> dict[str, object]:
    employees = session.exec(select(Employee)).all()
    cli_opts = sorted({(e.cli or "").strip() for e in employees if (e.cli or "").strip()})
    role_opts = ROLE_ORDER + sorted({e.role for e in employees if e.role not in ROLE_ORDER})
    gradation_opts = sorted({e.gradation for e in employees if e.gradation})
    cli_distribution = build_cli_distribution(employees)
    for row in cli_distribution:
        row["detail_href"] = f"/cli?roster_cli={urlparse.quote(str(row.get('cli') or ''))}#cli-distribution-detail"
        row["selected"] = bool(roster_cli and str(row.get("cli") or "").strip().lower() == roster_cli.strip().lower())

    totals_all = {
        "A": sum(int(row.get("A") or 0) for row in cli_distribution),
        "B": sum(int(row.get("B") or 0) for row in cli_distribution),
        "C": sum(int(row.get("C") or 0) for row in cli_distribution),
        "total": sum(int(row.get("total") or 0) for row in cli_distribution),
        "total_staff": sum(int(row.get("total_staff") or 0) for row in cli_distribution),
    }

    cli_roster = list(employees)
    if roster_cli_status in (None, ""):
        cli_roster = [e for e in cli_roster if e.cli or e.cli_id]
    elif roster_cli_status == "assigned":
        cli_roster = [e for e in cli_roster if e.cli or e.cli_id]
    elif roster_cli_status == "unassigned":
        cli_roster = [e for e in cli_roster if not e.cli and not e.cli_id]
    if roster_name:
        name_lower = roster_name.lower()
        cli_roster = [e for e in cli_roster if name_lower in e.name.lower()]
    if roster_cli:
        cli_lower = roster_cli.strip().lower()
        cli_roster = [e for e in cli_roster if e.cli and cli_lower in e.cli.strip().lower()]
    if roster_role:
        cli_roster = [e for e in cli_roster if e.role == roster_role]
    if roster_gradation:
        grad_lower = roster_gradation.lower()
        cli_roster = [e for e in cli_roster if e.gradation and grad_lower in e.gradation.lower()]
    cli_roster = sorted(cli_roster, key=lambda e: ((e.cli or "").strip().lower(), role_sort_key(e.role), e.name))

    return {
        "request": request,
        "active_page": "cli",
        "cli_distribution": cli_distribution,
        "cli_distribution_totals_all": totals_all,
        "cli_distribution_breakdown": [],
        "cli_distribution_totals": {"A": 0, "B": 0, "C": 0, "total": 0},
        "cli_distribution_detail_label": roster_cli or "",
        "cli_bio_reference_rows": [],
        "cli_roster": cli_roster,
        "cli_opts": cli_opts,
        "role_opts": role_opts,
        "gradation_opts": gradation_opts,
        "roster_name": roster_name or "",
        "roster_cli": roster_cli or "",
        "roster_role": roster_role or "",
        "roster_gradation": roster_gradation or "",
        "roster_cli_status": roster_cli_status or "",
        "roster_open": True,
        "grading_update_error": "",
        "grading_update_notice": "",
        "grading_update_warning": "",
        "grading_update_details": [],
        "grading_warning_details": [],
        "grading_source_name": "",
        "grading_report_date": "",
        "grading_saved_at": "",
    }


@app.get("/cli")
def cli_page(
    request: Request,
    roster_name: Optional[str] = None,
    roster_cli: Optional[str] = None,
    roster_role: Optional[str] = None,
    roster_gradation: Optional[str] = None,
    roster_cli_status: Optional[str] = None,
    session: Session = Depends(get_session),
):
    return templates.TemplateResponse(
        "cli.html",
        _cli_page_context(
            request,
            session,
            roster_name=roster_name,
            roster_cli=roster_cli,
            roster_role=roster_role,
            roster_gradation=roster_gradation,
            roster_cli_status=roster_cli_status,
        ),
    )


@app.get("/cli-distribution-planner")
def cli_distribution_planner_page(request: Request, session: Session = Depends(get_session)):
    employees = session.exec(select(Employee)).all()
    cli_opts = sorted({(e.cli or "").strip() for e in employees if (e.cli or "").strip()})
    return templates.TemplateResponse(
        "cli_distribution_planner.html",
        {
            "request": request,
            "active_page": "cli_distribution_planner",
            "cli_opts": cli_opts,
            "role_opts": ROLE_ORDER + sorted({e.role for e in employees if e.role not in ROLE_ORDER}),
            "cli_manual_targets": [],
            "cli_plan_staff_opts": sorted(employees, key=lambda e: (role_sort_key(e.role), e.name)),
            "cli_plan_notice": "",
            "cli_plan_error": "",
            "cli_plan_selected_exclude_cli": [],
            "cli_plan_selected_retiring_cli": [],
            "cli_plan_selected_exclude_staff_ids": [],
            "cli_plan_created_at": "",
            "cli_plan_summary": [],
            "cli_plan_assignments": [],
            "cli_plan_current_cli_opts": [],
            "cli_plan_proposed_cli_opts": [],
        },
    )


@app.get("/top-performer")
def top_performer_page(request: Request):
    return templates.TemplateResponse(
        "top_performer.html",
        {
            "request": request,
            "active_page": "top_performer",
            "saved_at_label": "",
            "warnings": [],
            "minimum_runs": 3,
            "summary": None,
            "results": [],
            "comparison": None,
        },
    )


@app.get("/cli-matrix")
def cli_matrix_page(request: Request, report_date: Optional[str] = None):
    return templates.TemplateResponse(
        "cli_matrix.html",
        {
            "request": request,
            "active_page": "cli_matrix",
            "error": "",
            "saved_notice": "",
            "cached_template_name": "",
            "report_date": report_date or "",
            "source_report_date": "",
            "source_report_date_label": "",
            "summary_rows": [],
            "overdue_rows": [],
        },
    )


def _non_continuous_page_context(
    request: Request,
    *,
    active_page: str,
    route_base: str,
    page_title: str,
    heading_title: str,
    report_date: str | None = None,
) -> dict[str, object]:
    return {
        "request": request,
        "active_page": active_page,
        "route_base": route_base,
        "page_title": page_title,
        "heading_title": heading_title,
        "error": "",
        "saved_notice": "",
        "template_token": "",
        "source_name": "",
        "source_report_date": "",
        "source_report_date_label": "",
        "cached_template_name": "",
        "report_date": report_date or "",
        "sign_on_label": "SIGN_ON",
        "sign_off_label": "SIGN_OFF",
        "sign_on_rows": [],
        "sign_off_rows": [],
        "allow_reason_edit": False,
    }


@app.get("/non-continuous-duty")
def non_continuous_duty_page(request: Request, report_date: Optional[str] = None):
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_page_context(
            request,
            active_page="non_continuous_duty",
            route_base="/non-continuous-duty",
            page_title="Non Continuous Duty",
            heading_title="NON SUB NON CONT DUTY",
            report_date=report_date,
        ),
    )


@app.get("/sub-non-continuous-duty")
def sub_non_continuous_duty_page(request: Request, report_date: Optional[str] = None):
    return templates.TemplateResponse(
        "non_continuous_duty.html",
        _non_continuous_page_context(
            request,
            active_page="sub_non_continuous_duty",
            route_base="/sub-non-continuous-duty",
            page_title="Sub Non Continuous Duty",
            heading_title="SUB NON CONT DUTY",
            report_date=report_date,
        ),
    )


@app.get("/ssts-report")
def ssts_report_page(
    request: Request,
    force: int = 0,
    report_tab: str = "online_offline",
    selected_day: str | None = None,
    analysis_day: str | None = None,
    detail_view: str | None = None,
    pf_day: str | None = None,
    pf_task_id: str | None = None,
    pf_train: str | None = None,
    session: Session = Depends(get_session),
):
    sync_result = refresh_ssts_snapshot(session, force=bool(force))
    active_report_tab = report_tab if report_tab in {"online_offline", "pf_entering"} else "online_offline"
    selected_day_value = _parse_report_date(selected_day)
    analysis_day_value = _parse_report_date(analysis_day)
    pf_day_value = selected_day_value or date.today()
    parsed_pf_day = _parse_report_date(pf_day)
    if parsed_pf_day is not None:
        pf_day_value = parsed_pf_day
    context = build_ssts_report_context(session, selected_day=selected_day_value, analysis_day=analysis_day_value)
    pf_context = {
        "pf_report_day": pf_day_value.isoformat(),
        "pf_report_day_label": pf_day_value.strftime("%d-%m-%Y"),
        "pf_report_rows": [],
        "pf_analysis_summary_rows": [],
        "pf_analysis_selected_rows": [],
        "pf_analysis_selected_train": "",
        "pf_analysis_total_trains": 0,
        "pf_analysis_total_rows": 0,
        "pf_analysis_source_total_trains": 0,
        "pf_analysis_source_total_rows": 0,
        "pf_analysis_missing_count": 0,
        "pf_analysis_status": "idle",
        "pf_analysis_task_id": pf_task_id or "",
        "pf_analysis_message": "",
        "pf_report_error": None,
    }
    if active_report_tab == "pf_entering":
        task_payload = _get_ssts_pf_analysis_task(pf_task_id)
        if task_payload:
            pf_context["pf_analysis_status"] = str(task_payload.get("status") or "idle")
            pf_context["pf_analysis_message"] = str(task_payload.get("message") or "")
            pf_context["pf_analysis_task_id"] = pf_task_id or ""
            if task_payload.get("report_day"):
                pf_context["pf_report_day"] = str(task_payload.get("report_day"))
                try:
                    pf_day_value = date.fromisoformat(str(task_payload.get("report_day")))
                    pf_context["pf_report_day_label"] = pf_day_value.strftime("%d-%m-%Y")
                except ValueError:
                    pass
            if task_payload.get("status") == "completed":
                result = task_payload.get("result")
                if isinstance(result, dict):
                    pf_context.update({key: value for key, value in result.items() if key != "pf_analysis_detail_rows_by_train"})
                    detail_rows_by_train = result.get("pf_analysis_detail_rows_by_train")
                    if isinstance(detail_rows_by_train, dict):
                        selected_train_value = pf_train or (
                            str(pf_context["pf_analysis_summary_rows"][0].get("train_no") or "")
                            if pf_context["pf_analysis_summary_rows"]
                            else ""
                        )
                        pf_context["pf_analysis_selected_train"] = selected_train_value
                        selected_rows = detail_rows_by_train.get(selected_train_value, [])
                        pf_context["pf_analysis_selected_rows"] = selected_rows if isinstance(selected_rows, list) else []
            elif task_payload.get("status") == "error":
                pf_context["pf_report_error"] = str(task_payload.get("message") or "PF analysis failed.")
    latest_run = context.get("latest_run")
    latest_summary = {
        "total_rakes": len(context.get("latest_rows", [])),
        "online_now_count": len(context.get("online_now", [])),
        "offline_count": len(context.get("current_not_online", [])),
        "recent_offline_count": len(context.get("current_recently_offline", [])),
        "recently_offline_count": len(context.get("current_recently_offline", [])),
        "recently_online_count": len(context.get("recently_online", [])),
    }
    return templates.TemplateResponse(
        "ssts_report.html",
        {
            "request": request,
            "active_page": "ssts_report",
            "active_report_tab": active_report_tab,
            "ssts_web_url": SSTS_WEB_URL,
            "latest_run": latest_run,
            "latest_summary": latest_summary,
            "ssts_sync_status": sync_result.get("status"),
            "ssts_sync_message": sync_result.get("message"),
            "ssts_sync_observed_at": sync_result.get("observed_at"),
            "IST": IST,
            "active_detail_view": detail_view if detail_view in {"recent_offline", "recently_online"} else None,
            **context,
            **pf_context,
        },
    )


@app.post("/ssts-report/pf-analysis/start")
async def start_ssts_pf_analysis(pf_day: str = Form(...)):
    report_day = _parse_report_date(pf_day)
    if report_day is None:
        raise HTTPException(status_code=400, detail="Invalid PF analysis date.")

    task_id = uuid4().hex
    _set_ssts_pf_analysis_task(
        task_id,
        status="pending",
        progress=2,
        message="Queued for analysis...",
        report_day=report_day.isoformat(),
        result=None,
    )
    worker = threading.Thread(target=_run_ssts_pf_analysis_task, args=(task_id, report_day), daemon=True)
    worker.start()
    return JSONResponse(
        {
            "task_id": task_id,
            "status": "pending",
            "status_url": f"/ssts-report/pf-analysis/status?task_id={task_id}",
            "result_url": f"/ssts-report?report_tab=pf_entering&pf_task_id={task_id}",
        }
    )


@app.get("/ssts-report/pf-analysis/status")
def ssts_pf_analysis_status(task_id: str):
    task_payload = _get_ssts_pf_analysis_task(task_id)
    if not task_payload:
        raise HTTPException(status_code=404, detail="PF analysis task not found.")
    return JSONResponse(
        {
            "task_id": task_id,
            "status": str(task_payload.get("status") or "idle"),
            "progress": int(task_payload.get("progress") or 0),
            "message": str(task_payload.get("message") or ""),
            "result_url": f"/ssts-report?report_tab=pf_entering&pf_task_id={task_id}",
        }
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


@app.post("/exports/table.xlsx")
async def export_table_xlsx(request: Request):
    try:
        payload = await request.json()
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid export payload.") from exc
    title, headers, rows, report_date_label = _coerce_export_table_payload(payload)

    wb = Workbook()
    ws = wb.active
    ws.title = _sanitize_excel_sheet_title(title)
    column_count = max(1, len(headers))
    if report_date_label:
        ws.append([title])
        ws.append([f"Updated on: {report_date_label}"])
        header_row_index = 3
    else:
        ws.append([title])
        header_row_index = 2
    ws.append(headers)
    for row in rows:
        ws.append(row)

    header_fill = PatternFill(fill_type="solid", fgColor="16314C")
    header_font = Font(bold=True, color="F3FBFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws[header_row_index]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    ws.freeze_panes = f"A{header_row_index + 1}"
    last_row = header_row_index + max(len(rows), 1)
    ws.auto_filter.ref = f"A{header_row_index}:{get_column_letter(column_count)}{last_row}"

    title_alignment = Alignment(horizontal="left", vertical="center")
    title_font = Font(bold=True, color="16314C")
    for row_index in range(1, header_row_index):
        cell = ws.cell(row=row_index, column=1)
        cell.font = title_font
        cell.alignment = title_alignment
        ws.merge_cells(
            start_row=row_index,
            start_column=1,
            end_row=row_index,
            end_column=column_count,
        )

    for column_index, header in enumerate(headers, start=1):
        max_length = len(header)
        for row in rows:
            max_length = max(max_length, len(row[column_index - 1]))
        ws.column_dimensions[get_column_letter(column_index)].width = min(max(max_length + 2, 10), 42)

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = _sanitize_export_filename(title, "xlsx")
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{urlparse.quote(filename)}"},
    )


@app.post("/exports/table.pdf")
async def export_table_pdf(request: Request):
    try:
        payload = await request.json()
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid export payload.") from exc

    title, headers, rows, report_date_label = _coerce_export_table_payload(payload)
    pdf_bytes = _build_table_pdf_bytes(title, headers, rows, report_date_label)
    filename = _sanitize_export_filename(title, "pdf")
    return StreamingResponse(
        BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{urlparse.quote(filename)}"},
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
